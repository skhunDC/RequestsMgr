// Code.gs - simplified supplies request system

const SHEET_ORDERS = 'Orders';
const SHEET_CATALOG = 'Catalog';
const SS_ID_PROP = 'SS_ID';

const STOCK_LIST = {
  Office: [
    'Copy Paper 8.5×11 (case)',
    'Ballpoint Pens (box)',
    'Sharpie Markers (pack)',
    'Hanging File Folders (box)',
    'Thermal Receipt Paper (case)',
    'Shipping Labels 4×6 (roll)',
    'Packing Tape (6-pack)',
    'Envelopes #10 (box)'
  ],
  Cleaning: [
    'Nitrile Gloves (box)',
    'Paper Towels (case)',
    'Trash Liners 33gal (case)',
    'Disinfectant Spray (case)',
    'Glass Cleaner (1 gal)',
    'Floor Cleaner Concentrate (1 gal)',
    'Lint Rollers (12-pack)'
  ],
  Operations: [
    'Poly Garment Bags (roll)',
    'Wire Hangers 18" (case)',
    'Suit Hangers w/ Bar (case)',
    'Garment Tags (roll)',
    'Spotting Agent – Protein (qt)',
    'Spotting Agent – Tannin (qt)',
    'Detergent – Laundry (5 gal)',
    'Laundry Nets (each)',
    'Sizing/Finishing Spray (case)',
    'Laundry Bags – Customer (pack)',
    'Twine/Hook Ties (roll)'
  ]
};

function getSs_() {
  const props = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActive();
  if (!ss) {
    const id = props.getProperty(SS_ID_PROP);
    if (id) {
      ss = SpreadsheetApp.openById(id);
    } else {
      ss = SpreadsheetApp.create('SuppliesTracking');
      props.setProperty(SS_ID_PROP, ss.getId());
    }
  }
  return ss;
}

function getSession() {
  init_();
  const email = Session.getActiveUser().getEmail();
  return { email };
}

function getCatalog(req) {
  init_();
  const includeArchived = req && req.includeArchived;
  const sheet = getSs_().getSheetByName(SHEET_CATALOG);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  return rows
    .map(r => Object.fromEntries(r.map((v, i) => [header[i], v])))
    .filter(r => includeArchived || r.archived !== true);
}

function addCatalogItem(req) {
  return withLock_(() => {
    const { description, category } = req;
    const sheet = getSs_().getSheetByName(SHEET_CATALOG);
    const sku = uuid_();
    sheet.appendRow([sku, description, category, false]);
    return { sku, description, category, archived: false };
  });
}

function setCatalogArchived(req) {
  return withLock_(() => {
    const { sku, archived } = req;
    const sheet = getSs_().getSheetByName(SHEET_CATALOG);
    const values = sheet.getDataRange().getValues();
    const header = values.shift();
    const skuIdx = header.indexOf('sku');
    const archIdx = header.indexOf('archived');
    const row = values.findIndex(r => r[skuIdx] === sku);
    if (row >= 0) {
      sheet.getRange(row + 2, archIdx + 1).setValue(archived);
    }
    return 'OK';
  });
}

function submitOrder(payload) {
  init_();
  const session = getSession();
  const sheet = getSs_().getSheetByName(SHEET_ORDERS);
  const lines = payload && Array.isArray(payload.lines) ? payload.lines : [];
  if (lines.length === 0) return [];
  const nowIso = nowIso_();
  return withLock_(() => {
    const orders = [];
    lines.forEach(line => {
      const order = {
        id: uuid_(),
        ts: nowIso,
        requester: session.email,
        description: line.description,
        qty: Number(line.qty),
        status: 'PENDING',
        approver: ''
      };
      sheet.appendRow([
        order.id,
        order.ts,
        order.requester,
        order.description,
        order.qty,
        order.status,
        order.approver
      ]);
      orders.push(order);
    });
    SpreadsheetApp.flush();
    return orders;
  });
}

function listMyOrders(req) {
  init_();
  const email = (req && req.email) || getSession().email;
  const sheet = getSs_().getSheetByName(SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  const keys = header.map(h => String(h).toLowerCase());
  const reqIdx = keys.indexOf('requester');
  return rows
    .filter(r => reqIdx >= 0 && r[reqIdx] === email)
    .map(r => Object.fromEntries(keys.map((k, i) => [k, r[i]])));
}

function listPendingApprovals() {
  init_();
  const sheet = getSs_().getSheetByName(SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  const keys = header.map(h => String(h).toLowerCase());
  return rows
    .filter(r => r[keys.indexOf('status')] === 'PENDING')
    .map(r => Object.fromEntries(keys.map((k, i) => [k, r[i]])));
}

function decideOrder(req) {
  const session = getSession();
  return withLock_(() => {
    const { id, decision } = req;
    const sheet = getSs_().getSheetByName(SHEET_ORDERS);
    const values = sheet.getDataRange().getValues();
    const header = values.shift();
    const keys = header.map(h => String(h).toLowerCase());
    const idIdx = keys.indexOf('id');
    const statusIdx = keys.indexOf('status');
    const approverIdx = keys.indexOf('approver');
    const requesterIdx = keys.indexOf('requester');
    const descIdx = keys.indexOf('description');
    const row = values.findIndex(r => r[idIdx] === id);
    if (row >= 0) {
      const r = row + 2;
      sheet.getRange(r, statusIdx + 1).setValue(decision);
      sheet.getRange(r, approverIdx + 1).setValue(session.email);
      const requester = values[row][requesterIdx];
      const desc = values[row][descIdx];
      GmailApp.sendEmail(requester, 'Supply Request ' + decision, '', {
        htmlBody: `<p>Your request for ${desc} was ${decision}.</p>`
      });
    }
    return 'OK';
  });
}

function uuid_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

function withLock_(fn) {
  // Standalone scripts don't have a document context, so `getDocumentLock`
  // can return `null`. Use a script lock instead to avoid null dereference
  // errors when submitting orders.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function init_() {
  const ss = getSs_();
  let sheet = ss.getSheetByName(SHEET_ORDERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ORDERS);
    sheet.appendRow(['id', 'ts', 'requester', 'description', 'qty', 'status', 'approver']);
  }
  sheet = ss.getSheetByName(SHEET_CATALOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CATALOG);
    sheet.appendRow(['sku', 'description', 'category', 'archived']);
  }
  seedCatalogIfEmpty_();
}

function seedCatalogIfEmpty_() {
  const sheet = getSs_().getSheetByName(SHEET_CATALOG);
  if (sheet.getLastRow() > 1) return;
  Object.keys(STOCK_LIST).forEach(cat => {
    STOCK_LIST[cat].forEach(desc => {
      sheet.appendRow([uuid_(), desc, cat, false]);
    });
  });
}

function doGet() {
  init_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Supplies Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

