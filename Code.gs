// Code.gs - simplified supplies request system

const SHEET_ORDERS = 'Orders';
const SHEET_CATALOG = 'Catalog';

const LT_EMAILS = [
  'skhun@dublincleaners.com',
  'ss.sku@protonmail.com',
  'brianmbutler77@gmail.com',
  'brianbutler@dublincleaners.com',
  'rbrown5940@gmail.com',
  'rbrown@dublincleaners.com',
  'davepdublincleaners@gmail.com',
  'lisamabr@yahoo.com',
  'dddale40@gmail.com',
  'nismosil85@gmail.com',
  'mlackey@dublincleaners.com',
  'china99@mail.com'
];

const STATIC_ADMINS = ['skhun@dublincleaners.com', 'ss.sku@protonmail.com'];
const ADMIN_PROP = 'ADMINS';
const SS_ID_PROP = 'SS_ID';

const APPROVER_BY_CATEGORY = {
  Office: 'skhun@dublincleaners.com',
  Cleaning: 'ss.sku@protonmail.com',
  Operations: 'skhun@dublincleaners.com'
};

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
  const isLt = LT_EMAILS.includes(email);
  const isAdmin = getAdmins_().includes(email);
  return { email, isLt, isAdmin };
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

function submitOrder(req) {
  const session = getSession();
  if (!session.isLt) throw new Error('Forbidden');
  return withLock_(() => {
    const sheet = getSs_().getSheetByName(SHEET_ORDERS);
    req.lines.forEach(line => {
      const approver = resolveApprover_(line);
      sheet.appendRow([
        uuid_(),
        nowIso_(),
        session.email,
        line.description,
        line.qty,
        'PENDING',
        approver
      ]);
      const html = `<p>${session.email} requested ${line.qty} × ${line.description}.</p>`;
      GmailApp.sendEmail(approver, 'Supply Request', '', { htmlBody: html });
    });
    return 'OK';
  });
}

function listMyOrders(req) {
  init_();
  const email = Session.getActiveUser().getEmail();
  const sheet = getSs_().getSheetByName(SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  return rows
    .filter(r => r[header.indexOf('requester')] === email)
    .map(r => Object.fromEntries(r.map((v, i) => [header[i], v])));
}

function listPendingApprovals() {
  const session = getSession();
  if (!session.isAdmin) throw new Error('Forbidden');
  const sheet = getSs_().getSheetByName(SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  return rows
    .filter(r => r[header.indexOf('status')] === 'PENDING')
    .map(r => Object.fromEntries(r.map((v, i) => [header[i], v])));
}

function decideOrder(req) {
  const session = getSession();
  if (!session.isAdmin) throw new Error('Forbidden');
  return withLock_(() => {
    const { id, decision } = req;
    const sheet = getSs_().getSheetByName(SHEET_ORDERS);
    const values = sheet.getDataRange().getValues();
    const header = values.shift();
    const idIdx = header.indexOf('id');
    const statusIdx = header.indexOf('status');
    const approverIdx = header.indexOf('approver');
    const row = values.findIndex(r => r[idIdx] === id);
    if (row >= 0) {
      const r = row + 2;
      sheet.getRange(r, statusIdx + 1).setValue(decision);
      sheet.getRange(r, approverIdx + 1).setValue(session.email);
      const requester = values[row][header.indexOf('requester')];
      const desc = values[row][header.indexOf('description')];
      GmailApp.sendEmail(requester, 'Supply Request ' + decision, '', {
        htmlBody: `<p>Your request for ${desc} was ${decision}.</p>`
      });
    }
    return 'OK';
  });
}

function resolveApprover_(line) {
  const catalog = getCatalog({ includeArchived: true });
  const item = catalog.find(it => it.description === line.description);
  const cat = item ? item.category : null;
  return (cat && APPROVER_BY_CATEGORY[cat]) || STATIC_ADMINS[0];
}

function uuid_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getAdmins_() {
  const props = PropertiesService.getScriptProperties();
  const extra = props.getProperty(ADMIN_PROP);
  return STATIC_ADMINS.concat(extra ? JSON.parse(extra) : []);
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

