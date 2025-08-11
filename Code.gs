// Refactored authentication & access control with role-based model

const SHEET_USERS = 'Users';
const SHEET_SYSTEM = 'System';
const SHEET_ORDERS = 'Orders';
const SHEET_CATALOG = 'Catalog';
const SHEET_AUDIT = 'Audit';

const ORDER_HEADER = ['id', 'ts', 'requester', 'description', 'qty', 'status', 'approver'];
const DEV_SEED_KEY = 'DEV_EMAILS_SEED';
const DEV_EMAILS = ['skhun@dublincleaners.com', 'ss.sku@protonmail.com'];

// ----- Locking -----
function withLock(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('System busy. Please retry.');
  try {
    return fn();
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
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
function seedDevUsers() {
  withLock(() => {
    const sysSheet = getOrCreateSheet(SHEET_SYSTEM);
    ensureHeaders(sysSheet, ['key', 'value']);
    const data = sysSheet.getDataRange().getValues();
    let row = data.findIndex(r => r[0] === DEV_SEED_KEY);
    if (row < 0) {
      sysSheet.appendRow([DEV_SEED_KEY, JSON.stringify(DEV_EMAILS)]);
      row = sysSheet.getLastRow() - 1; // zero-indexed data w/out header
    }
    const emails = JSON.parse(sysSheet.getRange(row + 1, 2).getValue() || '[]');

    const userSheet = getOrCreateSheet(SHEET_USERS);
    ensureHeaders(userSheet, ['email', 'roles', 'active']);
    const uRows = userSheet.getDataRange().getValues();
    const header = uRows.shift();
    const emailIdx = header.indexOf('email');
    const rolesIdx = header.indexOf('roles');
    const activeIdx = header.indexOf('active');
    emails.forEach(em => {
      const email = String(em).toLowerCase();
      const r = uRows.findIndex(row => String(row[emailIdx]).toLowerCase() === email);
      if (r >= 0) {
        userSheet.getRange(r + 2, 1, 1, 3).setValues([[email, 'developer,super_admin', true]]);
      } else {
        userSheet.appendRow([email, 'developer,super_admin', true]);
      }
    });
  });
}

function onOpen() {
  if (typeof setUpTriggers === 'function') setUpTriggers();
}

function setUpTriggers() {}

// ----- Identity & Roles -----
function getActiveUserEmail_() {
  return (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
}

function getUserRecord_(email) {
  if (!email) return null;
  const sheet = getOrCreateSheet(SHEET_USERS);
  ensureHeaders(sheet, ['email', 'roles', 'active']);
  const rows = sheet.getDataRange().getValues();
  rows.shift();
  const rec = rows.find(r => String(r[0]).toLowerCase() === email);
  if (!rec) return null;
  return {
    email,
    roles: String(rec[1] || '')
      .split(',')
      .map(r => r.trim())
      .filter(Boolean),
    active: rec[2] === true,
  };
}

function getRolesForEmail_(email) {
  const rec = getUserRecord_(email);
  return rec ? rec.roles : [];
}

function hasRole_(email, role) {
  return getRolesForEmail_(email).includes(role);
}

function requireRole(required) {
  const email = getActiveUserEmail_();
  if (!email) throw new Error('NOT_AUTHENTICATED');
  const rec = getUserRecord_(email);
  if (!rec || !rec.active) throw new Error('NOT_AUTHORIZED');
  const needed = Array.isArray(required) ? required : [required];
  if (!required || needed.length === 0) return email;
  if (needed.some(r => rec.roles.includes(r))) return email;
  throw new Error('NOT_AUTHORIZED');
}

function isDeveloperOrSuper_(email) {
  const roles = getRolesForEmail_(email);
  return roles.includes('developer') || roles.includes('super_admin');
}

// ----- CSRF -----
function getCsrfToken_(email) {
  const token = Utilities.getUuid();
  CacheService.getUserCache().put(token, email, 21600);
  return token;
}

function validateCsrf_(email, token) {
  const cached = CacheService.getUserCache().get(token);
  if (cached !== email) throw new Error('NOT_AUTHENTICATED');
}

function getSession() {
  const email = getActiveUserEmail_();
  const csrf = getCsrfToken_(email);
  return { csrf };
}

// ----- HTTP Entrypoints -----
function doGet(e) {
  seedDevUsers();
  init_();
  const email = getActiveUserEmail_();
  const t = HtmlService.createTemplateFromFile('index');
  t.googleEmail = email;
  t.userRoles = getRolesForEmail_(email).join(',');
  t.csrfToken = email ? getCsrfToken_(email) : '';
  t.appUrl = ScriptApp.getService().getUrl();
  return t
    .evaluate()
    .setTitle('Supplies Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function api(req) {
  const email = getActiveUserEmail_();
  validateCsrf_(email, req && req.csrf);
  return handleAction_(email, req.action, req.payload || {});
}

function doPost(e) {
  try {
    const email = getActiveUserEmail_();
    const body = JSON.parse(e.postData.contents);
    validateCsrf_(email, body.csrf);
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
    case 'users.list':
      requireRole(['developer', 'super_admin']);
      return listUsers_();
    case 'users.upsert':
      requireRole(['developer', 'super_admin']);
      return withLock(() => upsertUser_(payload));
    case 'users.remove':
      requireRole(['developer', 'super_admin']);
      return withLock(() => removeUser_(payload));
    case 'catalog.list':
      requireRole([]);
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
      requireRole([]);
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
  const msg = (err && err.message) || String(err);
  let code = msg;
  if (!['NOT_AUTHENTICATED', 'NOT_AUTHORIZED', 'VALIDATION', 'BUSY'].includes(msg)) {
    if (msg.startsWith('VALIDATION')) code = 'VALIDATION';
    else if (msg === 'System busy. Please retry.') code = 'BUSY';
    else code = 'UNKNOWN';
  }
  return { ok: false, code, message: msg };
}

// ----- Users API -----
function listUsers_() {
  const sheet = getOrCreateSheet(SHEET_USERS);
  ensureHeaders(sheet, ['email', 'roles', 'active']);
  const rows = sheet.getDataRange().getValues();
  rows.shift();
  return rows
    .filter(r => r[0])
    .map(r => ({
      email: String(r[0]).toLowerCase(),
      roles: String(r[1] || '')
        .split(',')
        .map(v => v.trim())
        .filter(Boolean),
      active: r[2] === true,
    }));
}

function upsertUser_(user) {
  const email = (user.email || '').toLowerCase().trim();
  if (!email) throw new Error('VALIDATION');
  const roles = (user.roles || []).join(',');
  const active = user.active !== false;
  const sheet = getOrCreateSheet(SHEET_USERS);
  ensureHeaders(sheet, ['email', 'roles', 'active']);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  const emailIdx = header.indexOf('email');
  const row = rows.findIndex(r => String(r[emailIdx]).toLowerCase() === email);
  const rowVals = [email, roles, active];
  if (row >= 0) {
    sheet.getRange(row + 2, 1, 1, 3).setValues([rowVals]);
  } else {
    sheet.appendRow(rowVals);
  }
  appendAudit('users.upsert', { email, roles: user.roles, active });
  return { email, roles: user.roles, active };
}

function removeUser_(payload) {
  const email = (payload.email || '').toLowerCase().trim();
  if (!email) throw new Error('VALIDATION');
  const sheet = getOrCreateSheet(SHEET_USERS);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  const emailIdx = header.indexOf('email');
  const activeIdx = header.indexOf('active');
  const row = rows.findIndex(r => String(r[emailIdx]).toLowerCase() === email);
  if (row >= 0) {
    sheet.getRange(row + 2, activeIdx + 1).setValue(false);
    appendAudit('users.remove', { email });
  }
  return 'OK';
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
  const sku = uuid_();
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
  const nowIso = nowIso_();
  const email = getActiveUserEmail_();
  const records = [];
  payload.lines.forEach(line => {
    const record = {
      id: uuid_(),
      ts: nowIso,
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
  const email = getActiveUserEmail_();
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
    sheet.getRange(r, approverIdx + 1).setValue(getActiveUserEmail_());
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
  return (cat && APPROVER_BY_CATEGORY[cat]) || DEV_EMAILS[0];
}

// ----- Audit & Utils -----
function appendAudit(action, data) {
  const sheet = getOrCreateSheet(SHEET_AUDIT);
  ensureHeaders(sheet, ['ts', 'email', 'action', 'data']);
  sheet.appendRow([nowIso_(), getActiveUserEmail_(), action, JSON.stringify(data)]);
}

function uuid_() {
  return Utilities.getUuid();
}

function nowIso_() {
  return new Date().toISOString();
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
      sheet.appendRow([uuid_(), desc, cat, false]);
    });
  });
}

