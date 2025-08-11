// Refactored authentication & access control with role-based model

const SHEET_USERS = 'Users';
const SHEET_ORDERS = 'Orders';
const SHEET_CATALOG = 'Catalog';
const SHEET_AUDIT = 'Audit';

const ORDER_HEADER = ['id', 'ts', 'requester', 'description', 'qty', 'status', 'approver'];
const DEV_ALLOWED = ['skhun@dublincleaners.com', 'ss.sku@protonmail.com'];

// ----- Utils -----
function uuid() {
  return Utilities.getUuid();
}

function nowIso() {
  return new Date().toISOString();
}

function safeLower(s) {
  return String(s || '').trim().toLowerCase();
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('Another change is in progress. Please try again in a few seconds.');
  }
  try {
    return fn();
  } finally {
    try {
      lock.releaseLock();
    } catch (err) {
      // ignore
    }
  }
}

// ----- Sheet Helpers -----
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActive();
  if (!ss) {
    const id = props.getProperty('SS_ID');
    if (id) {
      ss = SpreadsheetApp.openById(id);
    } else {
      ss = SpreadsheetApp.create('SuppliesTracking');
      props.setProperty('SS_ID', ss.getId());
    }
  }
  return ss;
}

function getOrCreateSheet(name) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureHeaders(sheet, headers) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  const current = range.getValues()[0];
  if (!current[0]) {
    range.setValues([headers]);
  }
}

// ----- Bootstrapping -----
function ensureSeedUsers() {
  withLock(() => {
    const sheet = getOrCreateSheet(SHEET_USERS);
    ensureHeaders(sheet, ['email', 'role', 'active', 'added_ts', 'added_by']);
    const rows = sheet.getDataRange().getValues();
    const header = rows.shift();
    const emailIdx = header.indexOf('email');
    const roleIdx = header.indexOf('role');
    const activeIdx = header.indexOf('active');
    const seeds = [
      { email: 'skhun@dublincleaners.com', role: 'super_admin' },
      { email: 'ss.sku@protonmail.com', role: 'developer' },
    ];
    seeds.forEach(seed => {
      const row = rows.findIndex(r => safeLower(r[emailIdx]) === seed.email);
      if (row >= 0) {
        const rNum = row + 2;
        sheet.getRange(rNum, roleIdx + 1).setValue(seed.role);
        sheet.getRange(rNum, activeIdx + 1).setValue(true);
      } else {
        sheet.appendRow([seed.email, seed.role, true, nowIso(), 'seed']);
      }
    });
  });
}

function onOpen() {
  if (typeof setUpTriggers === 'function') setUpTriggers();
}

function setUpTriggers() {}

// ----- Identity & Roles -----
function getActiveEmail() {
  return safeLower(Session.getActiveUser().getEmail());
}

function getUserRecord(email) {
  const em = safeLower(email);
  if (!em) return null;
  const sheet = getOrCreateSheet(SHEET_USERS);
  ensureHeaders(sheet, ['email', 'role', 'active', 'added_ts', 'added_by']);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  const emailIdx = header.indexOf('email');
  const roleIdx = header.indexOf('role');
  const activeIdx = header.indexOf('active');
  const rec = rows.find(r => safeLower(r[emailIdx]) === em);
  if (!rec) return { email: em, role: 'requester', active: true };
  return {
    email: em,
    role: rec[roleIdx],
    active: rec[activeIdx] === true || String(rec[activeIdx]).toUpperCase() === 'TRUE',
  };
}

function requireLoggedIn() {
  const email = getActiveEmail();
  if (!email) throw new Error('Login required');
  return email;
}

function requireRole(allowed) {
  const email = requireLoggedIn();
  const user = getUserRecord(email);
  if (!user || !user.active) throw new Error('Access denied');
  if (!allowed || allowed.length === 0) return email;
  if (allowed.includes(user.role)) return email;
  throw new Error('Access denied');
}

function isDevConsoleAllowed(email) {
  const em = safeLower(email);
  if (!em) return false;
  if (DEV_ALLOWED.includes(em)) return true;
  const rec = getUserRecord(em);
  return !!(rec && rec.active && (rec.role === 'developer' || rec.role === 'super_admin'));
}

// ----- CSRF -----
function createCsrfToken(email) {
  const token = uuid();
  CacheService.getUserCache().put('csrf', token, 21600);
  return token;
}

function validateCsrf(email, token) {
  const cached = CacheService.getUserCache().get('csrf');
  if (!cached || cached !== token) throw new Error('Login required');
}

// getSession helper removed; use 'session.get' action instead

// ----- HTTP Entrypoints -----
function doGet(e) {
  ensureSeedUsers();
  init_();
  const email = getActiveEmail();
  const rec = getUserRecord(email);
  const bootstrap = {
    email: email || '',
    role: rec ? rec.role : null,
    isLoggedIn: !!email,
    devConsoleAllowed: isDevConsoleAllowed(email),
    csrf: email ? createCsrfToken(email) : '',
  };
  const t = HtmlService.createTemplateFromFile('index');
  t.BOOTSTRAP = bootstrap;
  return t
    .evaluate()
    .setTitle('Supplies Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const email = getActiveEmail();
    const body = JSON.parse(e.postData.contents);
    validateCsrf(email, body.csrf);
    const data = handleAction_(email, body.action, body.payload || {});
    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, data })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const res = mapError_(err);
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(
      ContentService.MimeType.JSON
    );
  }
}

function handleAction_(email, action, payload) {
  switch (action) {
    case 'session.get': {
      requireLoggedIn();
      const rec = getUserRecord(email);
      return { email, role: rec ? rec.role : null, devConsoleAllowed: isDevConsoleAllowed(email) };
    }
    case 'catalog.list':
      requireLoggedIn();
      return getCatalog(payload);
    case 'catalog.add':
      requireRole(['developer', 'super_admin']);
      return withLock(() => addCatalogItem(payload));
    case 'catalog.archive':
      requireRole(['developer', 'super_admin']);
      return withLock(() => setCatalogArchived(payload));
    case 'orders.submit':
      requireRole(['requester', 'approver', 'developer', 'super_admin']);
      return withLock(() => submitOrder(payload));
    case 'orders.mine':
      requireLoggedIn();
      return listMyOrders(payload);
    case 'orders.pending':
      requireRole(['approver', 'developer', 'super_admin']);
      return listPendingApprovals();
    case 'orders.decide':
      requireRole(['approver', 'developer', 'super_admin']);
      return withLock(() => decideOrder(payload));
    default:
      throw new Error('UNKNOWN');
  }
}

function mapError_(err) {
  const message = (err && err.message) || String(err);
  return { ok: false, error: message };
}

// ----- Catalog & Orders -----
const APPROVER_BY_CATEGORY = {
  Office: 'skhun@dublincleaners.com',
  Cleaning: 'ss.sku@protonmail.com',
  Operations: 'skhun@dublincleaners.com',
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
    'Envelopes #10 (box)',
  ],
  Cleaning: [
    'Nitrile Gloves (box)',
    'Paper Towels (case)',
    'Trash Liners 33gal (case)',
    'Disinfectant Spray (case)',
    'Glass Cleaner (1 gal)',
    'Floor Cleaner Concentrate (1 gal)',
    'Lint Rollers (12-pack)',
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
    'Twine/Hook Ties (roll)',
  ],
};

function getCatalog(req) {
  init_();
  const includeArchived = req && req.includeArchived;
  const sheet = getOrCreateSheet(SHEET_CATALOG);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  return rows
    .map(r => Object.fromEntries(r.map((v, i) => [header[i], v])))
    .filter(r => includeArchived || r.archived !== true);
}

function addCatalogItem(req) {
  const { description, category } = req;
  if (!description) throw new Error('VALIDATION');
  const sheet = getOrCreateSheet(SHEET_CATALOG);
  const sku = uuid();
  sheet.appendRow([sku, description, category, false]);
  const record = { sku, description, category, archived: false };
  appendAudit('catalog.add', record);
  return record;
}

function setCatalogArchived(req) {
  const { sku, archived } = req;
  const sheet = getOrCreateSheet(SHEET_CATALOG);
  const values = sheet.getDataRange().getValues();
  const header = values.shift();
  const skuIdx = header.indexOf('sku');
  const archIdx = header.indexOf('archived');
  const row = values.findIndex(r => r[skuIdx] === sku);
  if (row >= 0) {
    sheet.getRange(row + 2, archIdx + 1).setValue(archived);
    appendAudit('catalog.archive', { sku, archived });
  }
  return 'OK';
}

function submitOrder(payload) {
  const sheet = getOrCreateSheet(SHEET_ORDERS);
  const ts = nowIso();
  const email = getActiveEmail();
  const records = [];
  payload.lines.forEach(line => {
    const record = {
      id: uuid(),
      ts,
      requester: email,
      description: line.description,
      qty: Number(line.qty),
      status: 'PENDING',
      approver: resolveApprover_(line),
    };
    sheet.appendRow(ORDER_HEADER.map(h => record[h]));
    records.push(record);
    const html = `<p>${email} requested ${record.qty} × ${record.description}.</p>`;
    GmailApp.sendEmail(record.approver, 'Supply Request', '', { htmlBody: html });
  });
  appendAudit('orders.submit', { count: records.length });
  return records;
}

function listMyOrders(req) {
  init_();
  const email = getActiveEmail();
  const sheet = getOrCreateSheet(SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  return rows
    .filter(r => r[header.indexOf('requester')] === email)
    .map(r => Object.fromEntries(r.map((v, i) => [header[i], v])))
    .sort((a, b) => b.ts.localeCompare(a.ts));
}

function listPendingApprovals() {
  const sheet = getOrCreateSheet(SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  return rows
    .filter(r => r[header.indexOf('status')] === 'PENDING')
    .map(r => Object.fromEntries(r.map((v, i) => [header[i], v])))
    .sort((a, b) => b.ts.localeCompare(a.ts));
}

function decideOrder(req) {
  const { id, decision } = req;
  const sheet = getOrCreateSheet(SHEET_ORDERS);
  const values = sheet.getDataRange().getValues();
  const header = values.shift();
  const idIdx = header.indexOf('id');
  const statusIdx = header.indexOf('status');
  const approverIdx = header.indexOf('approver');
  const row = values.findIndex(r => r[idIdx] === id);
  if (row >= 0) {
    const r = row + 2;
    sheet.getRange(r, statusIdx + 1).setValue(decision);
    sheet.getRange(r, approverIdx + 1).setValue(getActiveEmail());
    const requester = values[row][header.indexOf('requester')];
    const desc = values[row][header.indexOf('description')];
    GmailApp.sendEmail(requester, 'Supply Request ' + decision, '', {
      htmlBody: `<p>Your request for ${desc} was ${decision}.</p>`,
    });
    appendAudit('orders.decide', { id, decision });
  }
  return 'OK';
}

function resolveApprover_(line) {
  const catalog = getCatalog({ includeArchived: true });
  const item = catalog.find(it => it.description === line.description);
  const cat = item ? item.category : null;
  return (cat && APPROVER_BY_CATEGORY[cat]) || DEV_ALLOWED[0];
}

// ----- Audit & Utils -----
function appendAudit(action, data) {
  const sheet = getOrCreateSheet(SHEET_AUDIT);
  ensureHeaders(sheet, ['ts', 'email', 'action', 'data']);
  sheet.appendRow([nowIso(), getActiveEmail(), action, JSON.stringify(data)]);
}

function init_() {
  const orders = getOrCreateSheet(SHEET_ORDERS);
  ensureHeaders(orders, ORDER_HEADER);
  const catalog = getOrCreateSheet(SHEET_CATALOG);
  ensureHeaders(catalog, ['sku', 'description', 'category', 'archived']);
  seedCatalogIfEmpty_();
}

function seedCatalogIfEmpty_() {
  const sheet = getOrCreateSheet(SHEET_CATALOG);
  if (sheet.getLastRow() > 1) return;
  Object.keys(STOCK_LIST).forEach(cat => {
    STOCK_LIST[cat].forEach(desc => {
      sheet.appendRow([uuid(), desc, cat, false]);
    });
  });
}

