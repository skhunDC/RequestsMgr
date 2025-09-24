/* eslint-env googleappsscript */

// Core constants
const SHEETS = {
  ORDERS: 'Orders',
  CATALOG: 'Catalog',
  BUDGETS: 'Budgets',
  AUDIT: 'Audit',
  ROLES: 'Roles',
  LT_DEVS: 'LT_Devs'
};
const SS_ID_PROP = 'SS_ID';
const DEV_EMAILS = ['skhun@dublincleaners.com', 'ss.sku@protonmail.com'];
const UPLOAD_FOLDER_PROP = 'UPLOAD_FOLDER_ID';
const DRIVE_VIEW_PREFIX = 'https://drive.google.com/uc?export=view&id=';

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
  const email = Session.getActiveUser().getEmail();
  const existing = readAll_(roles).map(r => r.email);
  if (email && existing.indexOf(email) === -1) roles.appendRow([email, 'requester']);
  DEV_EMAILS.forEach(dev => {
    if (existing.indexOf(dev) === -1) roles.appendRow([dev, 'developer']);
  });
  // LT_Devs
  const lt = getOrCreateSheet_(SHEETS.LT_DEVS, ['email']);
  if (lt.getLastRow() === 1) DEV_EMAILS.forEach(dev => lt.appendRow([dev]));
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
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) throw new Error('System busy, please retry.');
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (err) { /* ignore */ }
  }
}

function appendAudit_(entity, entity_id, action, diffJson) {
  const sheet = getOrCreateSheet_(SHEETS.AUDIT, ['ts', 'actor', 'entity', 'entity_id', 'action', 'diff_json']);
  writeRow_(sheet, {
    ts: nowIso_(),
    actor: Session.getActiveUser().getEmail(),
    entity,
    entity_id,
    action,
    diff_json: diffJson || ''
  });
}

function getUserRole_(email) {
  const sheet = getOrCreateSheet_(SHEETS.ROLES, ['email', 'role']);
  const row = readAll_(sheet).find(r => r.email === email);
  return row ? row.role : 'viewer';
}

function requireRole_(allowed) {
  const email = Session.getActiveUser().getEmail();
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
  const email = Session.getActiveUser().getEmail();
  const role = getUserRole_(email);
  const cache = CacheService.getUserCache();
  let csrf = cache.get('csrf');
  if (!csrf) {
    csrf = uuid_();
    cache.put('csrf', csrf, 21600);
  }
  return { email, role, csrf };
}

function checkCsrf_(token) {
  const cache = CacheService.getUserCache();
  const csrf = cache.get('csrf');
  if (!csrf || csrf !== token) throw new Error('Bad CSRF');
}

// ---------- APIs ----------
function router(req) {
  if (typeof req === 'string') {
    req = { action: req };
  }
  req = req || {};
  const rawAction = typeof req.action === 'string' ? req.action : '';
  const action = rawAction.trim();
  if (!action) throw new Error('Unknown action');
  const normalized = action.toLowerCase();
  if (['getsession', 'listcatalog', 'listorders', 'listbudgets'].indexOf(normalized) === -1) {
    checkCsrf_(req.csrf);
  }
  switch (normalized) {
    case 'getsession':
      return getSession_();
    case 'listcatalog':
      return readAll_(getOrCreateSheet_(SHEETS.CATALOG, CATALOG_HEADERS))
        .filter(r => String(r.active) !== 'false');
    case 'listorders':
      return apiListOrders_(req.filter || {});
    case 'createorder':
      return apiCreateOrder_(req.payload || {});
    case 'bulkdecision':
      return apiBulkDecision_(req.ids || [], req.decision, req.comment || '');
    case 'updateorderproof':
      return apiUpdateOrderProof_(req.id, req.eta || '', req.image || '');
    case 'listbudgets':
      return readAll_(getOrCreateSheet_(SHEETS.BUDGETS, ['cost_center', 'month', 'budget', 'spent_to_date']));
    case 'updatecatalogimage':
      return apiUpdateCatalogImage_(req.sku, req.image || '');
    case 'uploadimage':
      return apiUploadImage_(req || {});
    default:
      throw new Error('Unknown action: ' + action);
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
    override: String(r['override?']) === 'true',
    justification: r.justification,
    eta_details: r.eta_details || '',
    proof_image: r.proof_image || '',
    statusChip: r.status
  }));
  const email = Session.getActiveUser().getEmail();
  let res = rows;
  if (filter.mineOnly) res = res.filter(r => r.requester === email);
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
  const email = Session.getActiveUser().getEmail();
  if (!payload.item) throw new Error('Missing item');
  const qty = Number(payload.qty);
  if (!qty || qty < 1) throw new Error('Missing qty');
  const estCost = Number(payload.est_cost);
  const catalog = readAll_(getOrCreateSheet_(SHEETS.CATALOG, CATALOG_HEADERS));
  const catRow = catalog.find(r => r.sku === payload.sku);
  if (catRow && String(catRow.override_required) === 'true') {
    if (!(payload.override === true && payload.justification && payload.justification.length >= 40)) {
      throw new Error('Override justification required');
    }
  }
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
    'override?': payload.override === true,
    justification: payload.justification || '',
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

function apiBulkDecision_(ids, decision, comment) {
  if (!decision) throw new Error('Missing decision');
  requireRole_(['approver', 'developer', 'super_admin']);
  const email = Session.getActiveUser().getEmail();
  const sheet = getOrCreateSheet_(SHEETS.ORDERS, ORDER_HEADERS);
  const headers = indexHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const idIdx = headers.id;
  const updates = [];
  withLock_(() => {
    ids.forEach(id => {
      const rowIdx = data.findIndex(r => r[idIdx] === id);
      if (rowIdx === -1) return;
      const row = data[rowIdx];
      const statusIdx = headers.status;
      const current = row[statusIdx];
      if (current === 'PENDING' && ['APPROVED', 'DENIED', 'ON-HOLD'].indexOf(decision) === -1) return;
      if (current === 'ON-HOLD' && ['APPROVED', 'DENIED'].indexOf(decision) === -1) return;
      const est = Number(row[headers.est_cost]) || 0;
      const month = String(row[headers.ts]).slice(0, 7);
      const hasCostCenter = typeof headers.cost_center !== 'undefined';
      const cc = hasCostCenter ? row[headers.cost_center] : '';
      if (decision === 'APPROVED' && hasCostCenter && cc) {
        const { warns, blocks } = willExceedBudget_(cc, month, est);
        if (blocks && !(['developer', 'super_admin'].indexOf(getUserRole_(email)) !== -1 && comment)) {
          throw new Error('Budget exceeded');
        }
        if (warns) updates.push({ type: 'warn', id });
        const budgetSheet = getOrCreateSheet_(SHEETS.BUDGETS, ['cost_center', 'month', 'budget', 'spent_to_date']);
        const rows = readAll_(budgetSheet);
        const bRow = rows.find(r => r.cost_center === cc && r.month === month);
        if (bRow) {
          const spent = Number(bRow.spent_to_date) + est;
          const rIdx = rows.indexOf(bRow) + 2;
          budgetSheet.getRange(rIdx, 4).setValue(spent);
        }
      }
      const r = rowIdx + 2;
      sheet.getRange(r, headers.status + 1).setValue(decision);
      sheet.getRange(r, headers.approver + 1).setValue(email);
      sheet.getRange(r, headers.decision_ts + 1).setValue(nowIso_());
      appendAudit_('Orders', id, 'DECISION', JSON.stringify({ decision, comment }));
      updates.push({ type: 'ok', id });
    });
  });
  if (updates.length) postToChatWebhook_('Bulk decision: ' + decision);
  return { updates };
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
