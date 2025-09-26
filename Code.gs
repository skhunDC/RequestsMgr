/* eslint-env googleappsscript */

// Core constants
const SHEETS = {
  ORDERS: 'Orders',
  CATALOG: 'Catalog',
  BUDGETS: 'Budgets',
  AUDIT: 'Audit',
  ROLES: 'Roles',
  LT_DEVS: 'LT_Devs',
  LT_AUTH: 'LT_Auth'
};
const SS_ID_PROP = 'SS_ID';
const DEV_EMAILS = ['skhun@dublincleaners.com', 'ss.sku@protonmail.com'];
const DEV_EMAILS_LOWER = DEV_EMAILS.map(email => normalizeEmail_(email));
const UPLOAD_FOLDER_PROP = 'UPLOAD_FOLDER_ID';
const DRIVE_VIEW_PREFIX = 'https://drive.google.com/uc?export=view&id=';

let CURRENT_SESSION_EMAIL = '';

const ORDER_HEADERS = ['id', 'ts', 'requester', 'item', 'qty', 'est_cost', 'status', 'approver', 'decision_ts', 'override?', 'justification', 'eta_details', 'proof_image'];
const CATALOG_HEADERS = ['sku', 'description', 'category', 'vendor', 'price', 'override_required', 'threshold', 'gl_code', 'cost_center', 'active', 'image_url'];

// Seed catalog items grouped by category
const STOCK_LIST = {
  Office: [
    'Copy Paper 8.5\u00d711 (case)',
    'Ballpoint Pens (box)',
    'Sharpie Markers (pack)',
    'Hanging File Folders (box)',
    'Thermal Receipt Paper (case)',
    'Shipping Labels 4\u00d76 (roll)',
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

// ---------- Initialization ----------
function getSs_() {
  const props = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActive();
  if (!ss) {
    const id = props.getProperty(SS_ID_PROP);
    ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('SuppliesTracking');
    if (!id) props.setProperty(SS_ID_PROP, ss.getId());
  }
  return ss;
}

function getOrCreateSheet_(name, headers) {
  const ss = getSs_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    return sh;
  }
  const existing = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  const missing = headers.filter(h => existing.indexOf(h) === -1);
  if (missing.length) {
    const startCol = existing.length ? existing.length + 1 : 1;
    missing.forEach((header, i) => {
      sh.getRange(1, startCol + i).setValue(header);
    });
  }
  return sh;
}

function init_() {
  const month = new Date().toISOString().slice(0, 7);
  // Orders
  getOrCreateSheet_(SHEETS.ORDERS, ORDER_HEADERS);
  // Catalog
  seedCatalogIfEmpty_();
  // Budgets
  const budgets = getOrCreateSheet_(SHEETS.BUDGETS, ['cost_center', 'month', 'budget', 'spent_to_date']);
  if (budgets.getLastRow() === 1) {
    const rows = [
      ['ADMIN', month, 200, 0],
      ['OPS', month, 300, 0]
    ];
    budgets.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  // Audit
  getOrCreateSheet_(SHEETS.AUDIT, ['ts', 'actor', 'entity', 'entity_id', 'action', 'diff_json']);
  // Roles
  const roles = getOrCreateSheet_(SHEETS.ROLES, ['email', 'role']);

  const email = getActiveUserNormalizedEmail_();

  const roleRows = readAll_(roles);
  const existing = roleRows.map(r => normalizeEmail_(r.email));
  if (email && existing.indexOf(email) === -1) roles.appendRow([email, 'requester']);
  DEV_EMAILS.forEach(dev => {
    if (existing.indexOf(normalizeEmail_(dev)) === -1) roles.appendRow([normalizeEmail_(dev), 'developer']);
  });
  readAll_(roles).forEach(row => ensureAuthRow_(row.email));
  // LT_Devs
  const lt = getOrCreateSheet_(SHEETS.LT_DEVS, ['email', 'salt', 'hash']);
  const ltRows = readAll_(lt);
  DEV_EMAILS.forEach(dev => {
    const normalized = normalizeEmail_(dev);
    if (!ltRows.some(r => normalizeEmail_(r.email) === normalized)) {
      writeRow_(lt, { email: normalized, salt: '', hash: '' });
    }
  });
  DEV_EMAILS.forEach(dev => ensureAuthRow_(dev));
}

function seedCatalogIfEmpty_() {
  const sheet = getOrCreateSheet_(SHEETS.CATALOG, CATALOG_HEADERS);
  if (sheet.getLastRow() > 1) return;
  const rows = [];
  Object.keys(STOCK_LIST).forEach(cat => {
    STOCK_LIST[cat].forEach(description => {
      rows.push([uuid_(), description, cat, '', 0, false, 0, '', '', true]);
    });
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// ---------- Helpers ----------
function indexHeaders_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[h] = i; });
  return map;
}

function readAll_(sheet) {
  const values = sheet.getDataRange().getValues();
  const header = values.shift();
  return values.map(r => Object.fromEntries(header.map((h, i) => [h, r[i]])));
}

function writeRow_(sheet, obj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : '');
  sheet.appendRow(row);
}

function nowIso_() {
  return new Date().toISOString();
}

function uuid_() {
  return Utilities.getUuid();
}

function ensureFolderShare_(folder) {
  if (!folder) return;
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    try {
      folder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err2) {
      // ignore
    }
  }
}

function ensureFilePublic_(file) {
  if (!file) return;
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (err2) {
      // ignore
    }
  }
}

function getUploadFolder_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(UPLOAD_FOLDER_PROP);
  let folder = null;
  if (existingId) {
    try {
      folder = DriveApp.getFolderById(existingId);
    } catch (err) {
      folder = null;
    }
  }
  if (!folder) {
    folder = DriveApp.createFolder('SuppliesTracker Uploads');
    props.setProperty(UPLOAD_FOLDER_PROP, folder.getId());
  }
  ensureFolderShare_(folder);
  return folder;
}

function hexDigest_(bytes) {
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function generateSalt_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 48);
}

function computeSaltedHash_(password, salt) {
  const material = (salt || '') + '::' + password;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, material, Utilities.Charset.UTF_8);
  return hexDigest_(digest);
}

function devSheet_() {
  return getOrCreateSheet_(SHEETS.LT_DEVS, ['email', 'salt', 'hash']);
}

function ensureDevRows_() {
  const sheet = devSheet_();
  const rows = readAll_(sheet);
  DEV_EMAILS.forEach(email => {
    const normalized = normalizeEmail_(email);
    if (!rows.some(r => normalizeEmail_(r.email) === normalized)) {
      writeRow_(sheet, { email: normalized, salt: '', hash: '' });
    }
  });
}

function getDevRow_(email) {
  const sheet = devSheet_();
  const normalized = normalizeEmail_(email);
  return readAll_(sheet).find(r => normalizeEmail_(r.email) === normalized) || null;
}

function upsertDevRow_(email, updates) {
  const sheet = devSheet_();
  const headers = indexHeaders_(sheet);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);
  const normalized = normalizeEmail_(email);
  const rowIdx = rows.findIndex(r => normalizeEmail_(r[headers.email]) === normalized);
  if (rowIdx === -1) {
    const headerRow = Object.keys(headers).sort((a, b) => headers[a] - headers[b]);
    const row = headerRow.map(key => {
      if (key === 'email') return normalized;
      if (Object.prototype.hasOwnProperty.call(updates, key)) return updates[key];
      return '';
    });
    sheet.appendRow(row);
  } else {
    const rowNumber = rowIdx + 2;
    Object.keys(updates).forEach(key => {
      if (typeof headers[key] === 'undefined') return;
      sheet.getRange(rowNumber, headers[key] + 1).setValue(updates[key]);
    });
  }
}

function authSheet_() {
  return getOrCreateSheet_(SHEETS.LT_AUTH, ['email', 'salt', 'hash', 'updated_ts']);
}

function ensureAuthRow_(email) {
  const sheet = authSheet_();
  const normalized = normalizeEmail_(email);
  if (!normalized) return;
  const rows = readAll_(sheet);
  if (!rows.some(r => normalizeEmail_(r.email) === normalized)) {
    writeRow_(sheet, { email: normalized, salt: '', hash: '', updated_ts: '' });
  }
}

function getAuthRow_(email) {
  const sheet = authSheet_();
  const normalized = normalizeEmail_(email);
  if (!normalized) return null;
  return readAll_(sheet).find(r => normalizeEmail_(r.email) === normalized) || null;
}

function upsertAuthRow_(email, updates) {
  const sheet = authSheet_();
  const headers = indexHeaders_(sheet);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1);
  const normalized = normalizeEmail_(email);
  if (!normalized) return;
  const rowIdx = rows.findIndex(r => normalizeEmail_(r[headers.email]) === normalized);
  if (rowIdx === -1) {
    const headerRow = Object.keys(headers).sort((a, b) => headers[a] - headers[b]);
    const row = headerRow.map(key => {
      if (key === 'email') return normalized;
      if (Object.prototype.hasOwnProperty.call(updates, key)) return updates[key];
      return '';
    });
    sheet.appendRow(row);
  } else {
    const rowNumber = rowIdx + 2;
    Object.keys(updates).forEach(key => {
      if (typeof headers[key] === 'undefined') return;
      sheet.getRange(rowNumber, headers[key] + 1).setValue(updates[key]);
    });
  }
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function setUserPassword_(email, password) {
  const normalized = normalizeEmail_(email);
  if (!normalized) throw new Error('Valid email required');
  const pwd = String(password || '').trim();
  if (!pwd || pwd.length < 12) throw new Error('Password must be at least 12 characters');
  const salt = generateSalt_();
  const hash = computeSaltedHash_(pwd, salt);
  withLock_(() => {
    upsertAuthRow_(normalized, { salt, hash, updated_ts: nowIso_() });
  });
  appendAudit_('Auth', normalized, 'SET_PASSWORD', '{}');
}

function verifyUserPassword_(email, password) {
  const normalized = normalizeEmail_(email);
  if (!normalized) return false;
  const row = getAuthRow_(normalized);
  if (!row || !row.salt || !row.hash) return false;
  const hashed = computeSaltedHash_(password, row.salt);
  return hashed === row.hash;
}


function getActiveUserEmail_() {
  let email = '';
  if (CURRENT_SESSION_EMAIL) {
    return CURRENT_SESSION_EMAIL;
  }
  try {
    const user = Session.getActiveUser ? Session.getActiveUser() : null;
    if (user) {
      if (typeof user.getEmail === 'function') {
        try {
          email = user.getEmail();
        } catch (err) {
          email = '';
        }
      }
      if (!email && typeof user.getUserLoginId === 'function') {
        try {
          email = user.getUserLoginId();
        } catch (err2) {
          email = '';
        }
      }
    }
  } catch (err3) {
    email = '';
  }
  email = String(email || '').trim();
  if (!email) {
    try {
      const effective = Session.getEffectiveUser ? Session.getEffectiveUser() : null;
      if (effective && typeof effective.getEmail === 'function') {
        const effEmail = String(effective.getEmail() || '').trim();
        const normalized = normalizeEmail_(effEmail);
        if (normalized && DEV_EMAILS_LOWER.indexOf(normalized) !== -1) {
          email = effEmail;
        }
      }
    } catch (err4) {
      email = email || '';
    }
  }
  return email;

}

function getActiveUserNormalizedEmail_() {
  return normalizeEmail_(getActiveUserEmail_());
}

function requireDevEmail_() {
  const email = getActiveUserNormalizedEmail_();
  if (!email) throw new Error('Forbidden');
  const role = getUserRole_(email);
  if (role === 'developer' || role === 'super_admin') return email;
  if (DEV_EMAILS_LOWER.indexOf(email) !== -1) return email;
  throw new Error('Forbidden');
}

function getDevSessionToken_() {
  return CacheService.getUserCache().get('devSession') || '';
}

function refreshDevSessionToken_(token) {
  if (!token) return;
  CacheService.getUserCache().put('devSession', token, 1800);
}

function clearDevSessionToken_() {
  CacheService.getUserCache().remove('devSession');
}

function createDevSessionToken_() {
  const token = Utilities.getUuid();
  refreshDevSessionToken_(token);
  return token;
}

function requireDevSession_(token) {
  if (!token) throw new Error('Developer session required');
  const cached = getDevSessionToken_();
  if (!cached || cached !== token) throw new Error('Developer session expired');
  refreshDevSessionToken_(token);
}

function verifyDevPassword_(row, password) {
  if (!row || !row.salt || !row.hash) return false;
  const hashed = computeSaltedHash_(password, row.salt);
  return hashed === row.hash;
}

function siteSessionCache_() {
  return CacheService.getScriptCache();
}

function siteSessionKey_(token) {
  return 'siteSession:' + token;
}

function getSiteSession_(token) {
  const key = siteSessionKey_(token);
  const raw = token ? siteSessionCache_().get(key) : null;
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function storeSiteSession_(token, data) {
  if (!token || !data) return;
  siteSessionCache_().put(siteSessionKey_(token), JSON.stringify(data), 1800);
}

function createSiteSessionToken_(email) {
  const token = Utilities.getUuid();
  storeSiteSession_(token, { email: normalizeEmail_(email), ts: Date.now() });
  return token;
}

function refreshSiteSession_(token) {
  const session = getSiteSession_(token);
  if (!session) return null;
  storeSiteSession_(token, session);
  return session;
}

function clearSiteSession_(token) {
  if (!token) return;
  siteSessionCache_().remove(siteSessionKey_(token));
}

function requireSiteSession_(token) {
  if (!token) throw new Error('Login required');
  const session = refreshSiteSession_(token);
  if (!session || !session.email) throw new Error('Login expired');
  return session;
}

function parseDataUrl_(dataUrl) {
  if (!dataUrl) throw new Error('Missing image data');
  const match = dataUrl.match(/^data:(.+?);base64,(.+)$/);
  if (!match) throw new Error('Invalid image data');
  const contentType = match[1] || 'application/octet-stream';
  const bytes = Utilities.base64Decode(match[2]);
  return { contentType, bytes };
}

function buildUploadFilename_(nameHint, filename, contentType) {
  const fallbackName = filename ? filename.replace(/\.[^/.]+$/, '') : '';
  const base = (nameHint || fallbackName || 'image').toLowerCase();
  let sanitized = base.replace(/[^a-z0-9]+/g, '-');
  sanitized = sanitized.replace(/-+/g, '-').replace(/^-|-$/g, '');
  sanitized = sanitized.slice(0, 40);
  if (!sanitized) sanitized = 'image';
  let ext = '';
  if (filename && filename.indexOf('.') !== -1) {
    ext = filename.split('.').pop().toLowerCase();
  }
  if (!ext && contentType && contentType.indexOf('/') !== -1) {
    ext = contentType.split('/')[1];
  }
  if (!ext) ext = 'png';
  const stamp = new Date().toISOString().replace(/[-:TZ.]/g, '').slice(0, 14);
  return `${sanitized}-${stamp}.${ext}`;
}

  function withLock_(fn) {
    let lock;
    try {
      lock = LockService.getDocumentLock();
    } catch (err) {
      lock = null;
    }
    if (!lock || typeof lock.tryLock !== 'function') {
      lock = LockService.getScriptLock();
    }
    if (!lock.tryLock(5000)) throw new Error('System busy, please retry.');
    try {
      return fn();
    } finally {
      try {
        if (lock && typeof lock.releaseLock === 'function') {
          lock.releaseLock();
        }
      } catch (err) { /* ignore */ }
    }
  }

function appendAudit_(entity, entity_id, action, diffJson) {
  const sheet = getOrCreateSheet_(SHEETS.AUDIT, ['ts', 'actor', 'entity', 'entity_id', 'action', 'diff_json']);
  writeRow_(sheet, {
    ts: nowIso_(),
    actor: getActiveUserEmail_(),
    entity,
    entity_id,
    action,
    diff_json: diffJson || ''
  });
}

function getUserRole_(email) {
  const sheet = getOrCreateSheet_(SHEETS.ROLES, ['email', 'role']);
  const normalized = normalizeEmail_(email);
  if (!normalized) return 'viewer';
  const row = readAll_(sheet).find(r => normalizeEmail_(r.email) === normalized);
  return row ? (String(row.role || '').trim().toLowerCase() || 'viewer') : 'viewer';
}

function requireRole_(allowed) {

  const email = getActiveUserNormalizedEmail_();

  const role = getUserRole_(email);
  if (allowed.indexOf(role) === -1) throw new Error('Forbidden');
  return role;
}

function getBudgetSnapshot_(cost_center, month) {
  const sheet = getOrCreateSheet_(SHEETS.BUDGETS, ['cost_center', 'month', 'budget', 'spent_to_date']);
  const row = readAll_(sheet).find(r => r.cost_center === cost_center && r.month === month) || {};
  const budget = Number(row.budget) || 0;
  const spent = Number(row.spent_to_date) || 0;
  return { budget, spent_to_date: spent, pct: budget ? spent / budget : 0 };
}

function willExceedBudget_(cc, month, addAmount) {
  const snap = getBudgetSnapshot_(cc, month);
  const pctAfter = snap.budget ? (snap.spent_to_date + addAmount) / snap.budget : 0;
  return { pctAfter, warns: pctAfter >= 0.8 && pctAfter <= 1, blocks: pctAfter > 1 };
}

function getSession_() {
  init_();

  const email = getActiveUserNormalizedEmail_();

  const role = getUserRole_(email);
  const cache = CacheService.getUserCache();
  let csrf = cache.get('csrf');
  if (!csrf) {
    csrf = uuid_();
    cache.put('csrf', csrf, 21600);
  }
  return { email, role, csrf, devEmails: DEV_EMAILS };
}

function checkCsrf_(token) {
  const cache = CacheService.getUserCache();
  const csrf = cache.get('csrf');
  if (!csrf || csrf !== token) throw new Error('Bad CSRF');
}

// ---------- APIs ----------
const ROUTER_HANDLERS = {
  getsession: () => getSession_(),
  listcatalog: () => readAll_(getOrCreateSheet_(SHEETS.CATALOG, CATALOG_HEADERS))
    .filter(r => String(r.active) !== 'false'),
  listorders: req => apiListOrders_(req.filter || {}),
  createorder: req => apiCreateOrder_(req.payload || {}),
  setorderstatus: req => apiSetOrderStatus_(req.id, req.decision),
  updateorderproof: req => apiUpdateOrderProof_(req.id, req.eta || '', req.image || ''),
  listbudgets: () => readAll_(getOrCreateSheet_(SHEETS.BUDGETS, ['cost_center', 'month', 'budget', 'spent_to_date'])),
  updatecatalogimage: req => apiUpdateCatalogImage_(req.sku, req.image || ''),
  uploadimage: req => apiUploadImage_(req || {}),
  devstatus: req => apiDevStatus_(req || {}),
  devlogin: req => apiDevLogin_(req.password || ''),
  devsetpassword: req => apiDevSetPassword_(req || {}),
  devadduser: req => apiDevAddUser_(req || {}),
  devlistroles: req => apiDevListRoles_(req || {}),
  devlogout: req => apiDevLogout_(req.token || ''),
  sitestatus: req => apiSiteStatus_(req || {}),
  sitelogin: req => apiSiteLogin_(req || {}),
  sitelogout: req => apiSiteLogout_(req || {})
};

function router(req) {
  if (typeof req === 'string') {
    req = { action: req };
  }
  req = req || {};
  const rawAction = typeof req.action === 'string' ? req.action : '';
  const action = rawAction.trim();
  if (!action) throw new Error('Unknown action');
  const normalized = action.toLowerCase();
  const skipCsrf = ['getsession', 'sitelogin', 'sitestatus'].indexOf(normalized) !== -1;
  if (!skipCsrf) {
    checkCsrf_(req.csrf);
  }
  const handler = ROUTER_HANDLERS[normalized];
  if (!handler) {
    const keys = Object.keys(req).sort();
    appendAudit_('Router', '-', 'UNKNOWN_ACTION', JSON.stringify({ action, normalized, keys }));
    throw new Error('Unknown action: ' + action);
  }
  let sessionContext = null;
  if (['getsession', 'sitelogin', 'sitestatus'].indexOf(normalized) === -1) {
    sessionContext = requireSiteSession_(String(req.siteToken || ''));
    CURRENT_SESSION_EMAIL = sessionContext && sessionContext.email ? sessionContext.email : '';
  }
  try {
    return handler(req);
  } finally {
    CURRENT_SESSION_EMAIL = '';
  }
}

function apiListOrders_(filter) {
  const sheet = getOrCreateSheet_(SHEETS.ORDERS, ORDER_HEADERS);
  const rows = readAll_(sheet).map(r => ({
    id: r.id,
    ts: r.ts,
    requester: r.requester,
    item: r.item,
    qty: Number(r.qty) || 0,
    est_cost: Number(r.est_cost) || 0,
    status: r.status,
    approver: r.approver,
    decision_ts: r.decision_ts,
    non_catalog: String(r['override?']) === 'true',
    details: r.justification,
    eta_details: r.eta_details || '',
    proof_image: r.proof_image || '',
    statusChip: r.status
  }));
  const email = getActiveUserEmail_();
  const normalizedEmail = normalizeEmail_(email);
  let res = rows;
  if (filter.mineOnly) res = res.filter(r => normalizeEmail_(r.requester) === normalizedEmail);
  if (filter.status && filter.status.length) res = res.filter(r => filter.status.indexOf(r.status) !== -1);
  if (filter.search) {
    const s = String(filter.search).toLowerCase();
    res = res.filter(r => (r.item || '').toLowerCase().includes(s) || (r.requester || '').toLowerCase().includes(s));
  }
  if (filter.sinceTs) res = res.filter(r => r.ts >= filter.sinceTs);
  res.sort((a, b) => b.ts.localeCompare(a.ts));
  return res;
}

function apiCreateOrder_(payload) {
  const email = getActiveUserEmail_();
  if (!payload.item) throw new Error('Missing item');
  const qty = Number(payload.qty);
  if (!qty || qty < 1) throw new Error('Missing qty');
  const estCost = Number(payload.est_cost);
  const nonCatalog = String(payload.non_catalog) === 'true';
  const details = payload.description ? String(payload.description) : '';
  const order = {
    id: uuid_(),
    ts: nowIso_(),
    requester: email,
    item: payload.item,
    qty,
    est_cost: estCost || 0,
    status: 'PENDING',
    approver: '',
    decision_ts: '',
    'override?': nonCatalog,
    justification: details,
    eta_details: '',
    proof_image: ''
  };
  withLock_(() => {
    writeRow_(getOrCreateSheet_(SHEETS.ORDERS, ORDER_HEADERS), order);
  });
  appendAudit_('Orders', order.id, 'CREATE', JSON.stringify(order));
  sendGmailHtml_(email, 'Order Submitted', '<p>Your order was submitted.</p>');
  postToChatWebhook_('Order ' + order.id + ' created');
  return order;
}

function apiSetOrderStatus_(id, decision) {
  if (!id) throw new Error('Missing id');
  const normalizedDecision = String(decision || '').trim().toUpperCase();
  const allowed = ['APPROVED', 'DENIED', 'ON-HOLD'];
  if (allowed.indexOf(normalizedDecision) === -1) throw new Error('Invalid decision');
  requireRole_(['approver', 'developer', 'super_admin']);
  const email = getActiveUserEmail_();
  const stamp = nowIso_();
  const sheet = getOrCreateSheet_(SHEETS.ORDERS, ORDER_HEADERS);
  const headers = indexHeaders_(sheet);
  const idIdx = headers.id;
  const statusIdx = headers.status;
  const approverIdx = headers.approver;
  const decisionIdx = headers.decision_ts;
  if ([idIdx, statusIdx, approverIdx, decisionIdx].some(idx => typeof idx === 'undefined')) {
    throw new Error('Orders sheet missing columns');
  }
  withLock_(() => {
    const data = sheet.getDataRange().getValues();
    data.shift();
    const rowIdx = data.findIndex(row => row[idIdx] === id);
    if (rowIdx === -1) throw new Error('Order not found');
    const rowNumber = rowIdx + 2;
    sheet.getRange(rowNumber, statusIdx + 1).setValue(normalizedDecision);
    sheet.getRange(rowNumber, approverIdx + 1).setValue(email);
    sheet.getRange(rowNumber, decisionIdx + 1).setValue(stamp);
  });
  appendAudit_('Orders', id, 'DECISION', JSON.stringify({ decision: normalizedDecision }));
  postToChatWebhook_('Order ' + id + ' marked ' + normalizedDecision);
  return { id, status: normalizedDecision, approver: email, decision_ts: stamp };
}

function apiUpdateOrderProof_(id, eta, image) {
  if (!id) throw new Error('Missing id');
  requireRole_(['approver', 'developer', 'super_admin']);
  const sheet = getOrCreateSheet_(SHEETS.ORDERS, ORDER_HEADERS);
  const headers = indexHeaders_(sheet);
  const idIdx = headers.id;
  const etaIdx = headers.eta_details;
  const proofIdx = headers.proof_image;
  if (typeof idIdx === 'undefined' || typeof etaIdx === 'undefined' || typeof proofIdx === 'undefined') {
    throw new Error('Orders sheet missing columns');
  }
  withLock_(() => {
    const data = sheet.getDataRange().getValues();
    data.shift();
    const rowIdx = data.findIndex(r => r[idIdx] === id);
    if (rowIdx === -1) throw new Error('Order not found');
    const rowNumber = rowIdx + 2;
    sheet.getRange(rowNumber, etaIdx + 1).setValue(eta);
    sheet.getRange(rowNumber, proofIdx + 1).setValue(image);
  });
  appendAudit_('Orders', id, 'UPDATE_PROOF', JSON.stringify({ eta_details: eta, proof_image: image ? 'set' : '' }));
  return { id, eta_details: eta, proof_image: image };
}

function apiUpdateCatalogImage_(sku, image) {
  if (!sku) throw new Error('Missing sku');
  requireRole_(['developer', 'super_admin']);
  const sheet = getOrCreateSheet_(SHEETS.CATALOG, CATALOG_HEADERS);
  const headers = indexHeaders_(sheet);
  const skuIdx = headers.sku;
  const imageIdx = headers.image_url;
  if (typeof skuIdx === 'undefined' || typeof imageIdx === 'undefined') {
    throw new Error('Catalog sheet missing columns');
  }
  withLock_(() => {
    const data = sheet.getDataRange().getValues();
    data.shift();
    const rowIdx = data.findIndex(r => r[skuIdx] === sku);
    if (rowIdx === -1) throw new Error('Catalog item not found');
    sheet.getRange(rowIdx + 2, imageIdx + 1).setValue(image);
  });
  appendAudit_('Catalog', sku, 'UPDATE_IMAGE', JSON.stringify({ image_url: image ? 'set' : '' }));
  return { sku, image_url: image };
}

function apiUploadImage_(payload) {
  requireRole_(['developer', 'super_admin']);
  const data = payload && payload.data;
  if (!data) throw new Error('Missing image data');
  const parsed = parseDataUrl_(data);
  const folder = getUploadFolder_();
  ensureFolderShare_(folder);
  const fileName = buildUploadFilename_(payload && payload.name, payload && payload.filename, parsed.contentType);
  const blob = Utilities.newBlob(parsed.bytes, parsed.contentType, fileName);
  const file = folder.createFile(blob);
  ensureFilePublic_(file);
  appendAudit_('Uploads', file.getId(), 'CREATE', JSON.stringify({ name: fileName, contentType: parsed.contentType }));
  return { url: DRIVE_VIEW_PREFIX + file.getId() };
}

function apiDevStatus_() {

  const email = getActiveUserNormalizedEmail_();

  const role = getUserRole_(email);
  const allowed = !!email && (role === 'developer' || role === 'super_admin' || DEV_EMAILS_LOWER.indexOf(email) !== -1);
  if (!allowed) {
    clearDevSessionToken_();
    return { allowed: false, hasPassword: false, sessionActive: false, token: '' };
  }
  ensureDevRows_();
  const row = getDevRow_(email);
  const token = getDevSessionToken_();
  if (token) refreshDevSessionToken_(token);
  return {
    allowed: true,
    hasPassword: !!(row && row.hash),
    sessionActive: !!token,
    token: token || ''
  };
}

function apiDevLogin_(password) {
  if (!password) throw new Error('Missing password');
  const email = requireDevEmail_();
  ensureDevRows_();
  const row = getDevRow_(email);
  if (!row || !row.hash) throw new Error('Developer password not set');
  if (!verifyDevPassword_(row, password)) throw new Error('Invalid developer credentials');
  const token = createDevSessionToken_();
  appendAudit_('DevAuth', email, 'LOGIN', '{}');
  return { token };
}

function apiDevSetPassword_(req) {
  const email = requireDevEmail_();
  ensureDevRows_();
  const row = getDevRow_(email);
  const newPassword = String(req.newPassword || '').trim();
  if (!newPassword || newPassword.length < 12) {
    throw new Error('Password must be at least 12 characters');
  }
  const hasExisting = row && row.hash;
  if (hasExisting) {
    let authed = false;
    if (req.token) {
      try {
        requireDevSession_(req.token);
        authed = true;
      } catch (err) {
        authed = false;
      }
    }
    if (!authed && req.currentPassword) {
      authed = verifyDevPassword_(row, req.currentPassword);
    }
    if (!authed) throw new Error('Current password required');
  }
  const salt = generateSalt_();
  const hash = computeSaltedHash_(newPassword, salt);
  upsertDevRow_(email, { salt, hash });
  const token = createDevSessionToken_();
  appendAudit_('DevAuth', email, 'SET_PASSWORD', '{}');
  return { token };
}

function apiDevAddUser_(req) {
  const email = requireDevEmail_();
  requireDevSession_(req.token || '');
  const payload = req.payload || {};
  const targetEmail = normalizeEmail_(payload.email);
  const role = String(payload.role || '').trim();
  const password = String(payload.password || '').trim();
  if (!targetEmail || targetEmail.indexOf('@') === -1) throw new Error('Valid email required');
  const allowedRoles = ['requester', 'approver', 'developer', 'super_admin'];
  if (allowedRoles.indexOf(role) === -1) throw new Error('Invalid role');
  const sheet = getOrCreateSheet_(SHEETS.ROLES, ['email', 'role']);
  const headers = indexHeaders_(sheet);
  withLock_(() => {
    const data = sheet.getDataRange().getValues();
    const header = data[0] || Object.keys(headers).sort((a, b) => headers[a] - headers[b]);
    const rows = data.slice(1);
    const emailIdx = headers.email;
    const roleIdx = headers.role;
    const existingIdx = rows.findIndex(r => String(r[emailIdx]).toLowerCase() === targetEmail);
    if (existingIdx === -1) {
      const row = header.map((col, i) => {
        if (col === 'email') return targetEmail;
        if (col === 'role') return role;
        return '';
      });
      sheet.appendRow(row);
    } else {
      sheet.getRange(existingIdx + 2, roleIdx + 1).setValue(role);
    }
  });
  appendAudit_('Roles', targetEmail, 'UPSERT', JSON.stringify({ role }));
  ensureAuthRow_(targetEmail);
  if (password) {
    setUserPassword_(targetEmail, password);
  }
  return { email: targetEmail, role };
}

function apiDevListRoles_(req) {
  requireDevEmail_();
  requireDevSession_(req.token || '');
  return readAll_(getOrCreateSheet_(SHEETS.ROLES, ['email', 'role']));
}

function apiDevLogout_(token) {
  requireDevEmail_();
  requireDevSession_(token);
  clearDevSessionToken_();
  return { success: true };
}

function apiSiteStatus_(req) {
  const token = String(req.token || '').trim();
  if (!token) {
    return { authed: false };
  }
  const session = refreshSiteSession_(token);
  if (!session || !session.email) {
    return { authed: false };
  }
  const email = normalizeEmail_(session.email);
  const role = getUserRole_(email);
  if (!role || role === 'viewer') {
    clearSiteSession_(token);
    return { authed: false };
  }
  storeSiteSession_(token, { email, ts: Date.now() });
  return { authed: true, email, role, token };
}

function apiSiteLogin_(req) {
  const email = normalizeEmail_(req.email);
  const password = String(req.password || '').trim();
  if (!email || email.indexOf('@') === -1) throw new Error('Enter a valid email address');
  if (!password) throw new Error('Password required');
  ensureAuthRow_(email);
  const role = getUserRole_(email);
  if (!role || role === 'viewer') throw new Error('Account not authorized');
  let valid = verifyUserPassword_(email, password);
  if (!valid) {
    const authRow = getAuthRow_(email);
    const missingPassword = !authRow || !authRow.hash;
    if (missingPassword && DEV_EMAILS_LOWER.indexOf(email) !== -1) {
      setUserPassword_(email, password);
      valid = true;
    } else if (!valid) {
      const devRow = getDevRow_(email);
      if (devRow && verifyDevPassword_(devRow, password)) {
        setUserPassword_(email, password);
        valid = true;
      }
    }
  }
  if (!valid) throw new Error('Invalid email or password');
  const token = createSiteSessionToken_(email);
  appendAudit_('Auth', email, 'LOGIN', '{}');
  return { token, email, role };
}

function apiSiteLogout_(req) {
  const token = String(req.token || '').trim();
  if (!token) throw new Error('Login required');
  const session = requireSiteSession_(token);
  clearSiteSession_(token);
  appendAudit_('Auth', session.email, 'LOGOUT', '{}');
  return { success: true };
}

// ---------- Notification placeholders ----------
function sendGmailHtml_(to, subject, html) {
  appendAudit_('Notifications', '-', 'EMAIL_PLACEHOLDER', JSON.stringify({ to, subject }));
  // TODO: replace with GmailApp.sendEmail when integrating email notifications
}

function postToChatWebhook_(message) {
  appendAudit_('Notifications', '-', 'CHAT_PLACEHOLDER', JSON.stringify({ message }));
  // TODO: replace with actual chat webhook POST request
}

// ---------- Triggers ----------
function dailyDigest_() {
  appendAudit_('System', '-', 'DAILY_DIGEST_PLACEHOLDER', '{}');
  // TODO: replace with real digest logic
}

function setUpTriggers() {
  ScriptApp.newTrigger('dailyDigest_').timeBased().everyDays(1).create();
}

// ---------- Entry ----------
function doGet() {
  const session = getSession_();
  const tpl = HtmlService.createTemplateFromFile('index');
  tpl.session = session;
  return tpl.evaluate().setTitle('Supplies Tracker').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
