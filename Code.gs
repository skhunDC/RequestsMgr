// Code.gs - simplified supplies request system

const SHEET_ORDERS = 'Orders';
const SHEET_CATALOG = 'Catalog';
const SHEET_ARCHIVE = 'OrdersArchive';
const SS_ID_PROP = 'SS_ID';
const ORDERS_HEADERS = [
  'id',
  'ts',
  'requester',
  'item',
  'qty',
  'est_cost',
  'status',
  'approver',
  'decision_ts',
  'override?',
  'justification',
  'cost_center',
  'gl_code'
];

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
  const ids = [];
  const nowIso = nowIso_();
  withLock_(() => {
    payload.lines.forEach(line => {
      const id = uuid_();
      sheet.appendRow([
        id,
        nowIso,
        session.email,
        line.description,
        Number(line.qty),
        '',
        'PENDING',
        '',
        '',
        '',
        '',
        '',
        ''
      ]);
      ids.push(id);
    });
  });
  return ids;
}

function listRecentOrders() {
  init_();
  const sheet = getSs_().getSheetByName(SHEET_ORDERS);
  const archive = getOrCreateSheet(SHEET_ARCHIVE, ORDERS_HEADERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  const tsIdx = header.indexOf('ts');
  const now = new Date();
  const cutoff = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
  const recent = [];
  const old = [];
  rows.forEach(r => {
    const ts = new Date(r[tsIdx]);
    if (ts >= cutoff) {
      recent.push(r);
    } else {
      old.push(r);
    }
  });
  if (old.length) {
    withLock_(() => {
      old.forEach(r => archive.appendRow(r));
      for (let i = rows.length - 1; i >= 0; i--) {
        const ts = new Date(rows[i][tsIdx]);
        if (ts < cutoff) sheet.deleteRow(i + 2);
      }
    });
  }
  recent.sort((a, b) => new Date(b[tsIdx]) - new Date(a[tsIdx]));
  return recent.map(r => Object.fromEntries(header.map((h, i) => [h, r[i]])));
}

function listPendingApprovals() {
  init_();
  const sheet = getSs_().getSheetByName(SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  return rows
    .filter(r => r[header.indexOf('status')] === 'PENDING')
    .map(r => Object.fromEntries(r.map((v, i) => [header[i], v])));
}

function decideOrder(req) {
  const session = getSession();
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

function uuid_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
}

function getOrCreateSheet(name, headers) {
  const ss = getSs_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const match = headers.every((h, i) => h === current[i]);
  if (!match) {
    sheet.clear();
    sheet.appendRow(headers);
  }
  return sheet;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  const success = lock.tryLock(5000);
  if (!success) {
    throw new Error('Could not obtain lock');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function createOrder_(payload) {
  const sheet = getOrCreateSheet(SHEET_ORDERS, ORDERS_HEADERS);
  const id = Utilities.getUuid();
  const ts = new Date().toISOString();
  const requester = Session.getActiveUser().getEmail() || 'unknown@local';
  const qty = Number(payload.qty);
  const estCost = Number(payload.est_cost);
  const override = String(payload.override || 'NO');
  const row = [
    id,
    ts,
    requester,
    String(payload.item || ''),
    qty,
    estCost,
    'PENDING',
    '',
    '',
    override,
    String(payload.justification || ''),
    String(payload.cost_center || ''),
    String(payload.gl_code || '')
  ];
  withLock_(() => sheet.appendRow(row));
  return { ok: true, id, status: 'PENDING' };
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';
    const payload = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    if (action === 'createOrder') {
      return jsonResponse_(createOrder_(payload));
    }
    if (action === 'getOrders') {
      return jsonResponse_({ ok: true, rows: listRecentOrders() });
    }
    return jsonResponse_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message });
  }
}

function init_() {
  const ss = getSs_();
  getOrCreateSheet(SHEET_ORDERS, ORDERS_HEADERS);
  let sheet = ss.getSheetByName(SHEET_CATALOG);
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

