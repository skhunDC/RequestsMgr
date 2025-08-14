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
  }
  return sh;
}

function init_() {
  const month = new Date().toISOString().slice(0, 7);
  // Orders
  getOrCreateSheet_(SHEETS.ORDERS, ['id', 'ts', 'requester', 'item', 'qty', 'est_cost', 'status', 'approver', 'decision_ts', 'override?', 'justification', 'cost_center', 'gl_code']);
  // Catalog
  const catalog = getOrCreateSheet_(SHEETS.CATALOG, ['sku', 'desc', 'category', 'vendor', 'price', 'override_required', 'threshold', 'gl_code', 'cost_center', 'active']);
  if (catalog.getLastRow() === 1) {
    const seed = [
      ['PAPER', 'Copy Paper 8.5x11', 'Office', 'OfficeMax', 30, false, 0, '6000', 'ADMIN', true],
      ['GLOVES', 'Nitrile Gloves', 'Cleaning', 'SafetyCo', 20, false, 0, '6100', 'OPS', true],
      ['SOLVENT', 'Special Solvent', 'Operations', 'ChemCorp', 50, true, 40, '6200', 'OPS', true]
    ];
    catalog.getRange(2, 1, seed.length, seed[0].length).setValues(seed);
  }
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
  req = req || {};
  const action = req.action;
  if (action !== 'getSession' && action !== 'listCatalog' && action !== 'listOrders' && action !== 'listBudgets') {
    checkCsrf_(req.csrf);
  }
  switch (action) {
    case 'getSession':
      return getSession_();
    case 'listCatalog':
      return readAll_(getOrCreateSheet_(SHEETS.CATALOG, ['sku', 'desc', 'category', 'vendor', 'price', 'override_required', 'threshold', 'gl_code', 'cost_center', 'active']))
        .filter(r => String(r.active) !== 'false');
    case 'listOrders':
      return apiListOrders_(req.filter || {});
    case 'createOrder':
      return apiCreateOrder_(req.payload || {});
    case 'bulkDecision':
      return apiBulkDecision_(req.ids || [], req.decision, req.comment || '');
    case 'listBudgets':
      return readAll_(getOrCreateSheet_(SHEETS.BUDGETS, ['cost_center', 'month', 'budget', 'spent_to_date']));
    default:
      throw new Error('Unknown action');
  }
}

function apiListOrders_(filter) {
  const sheet = getOrCreateSheet_(SHEETS.ORDERS, ['id', 'ts', 'requester', 'item', 'qty', 'est_cost', 'status', 'approver', 'decision_ts', 'override?', 'justification', 'cost_center', 'gl_code']);
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
    cost_center: r.cost_center,
    gl_code: r.gl_code,
    statusChip: r.status
  }));
  const email = Session.getActiveUser().getEmail();
  let res = rows;
  if (filter.mineOnly) res = res.filter(r => r.requester === email);
  if (filter.status && filter.status.length) res = res.filter(r => filter.status.indexOf(r.status) !== -1);
  if (filter.search) {
    const s = String(filter.search).toLowerCase();
    res = res.filter(r => (r.item || '').toLowerCase().includes(s) || (r.requester || '').toLowerCase().includes(s) || (r.gl_code || '').toLowerCase().includes(s));
  }
  if (filter.costCenter) res = res.filter(r => r.cost_center === filter.costCenter);
  if (filter.sinceTs) res = res.filter(r => r.ts >= filter.sinceTs);
  res.sort((a, b) => b.ts.localeCompare(a.ts));
  return res;
}

function apiCreateOrder_(payload) {
  const email = Session.getActiveUser().getEmail();
  ['item', 'qty', 'est_cost', 'cost_center', 'gl_code'].forEach(k => {
    if (!payload[k]) throw new Error('Missing ' + k);
  });
  const catalog = readAll_(getOrCreateSheet_(SHEETS.CATALOG, ['sku', 'desc', 'category', 'vendor', 'price', 'override_required', 'threshold', 'gl_code', 'cost_center', 'active']));
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
    qty: Number(payload.qty),
    est_cost: Number(payload.est_cost),
    status: 'PENDING',
    approver: '',
    decision_ts: '',
    'override?': payload.override === true,
    justification: payload.justification || '',
    cost_center: payload.cost_center,
    gl_code: payload.gl_code
  };
  withLock_(() => {
    writeRow_(getOrCreateSheet_(SHEETS.ORDERS, ['id', 'ts', 'requester', 'item', 'qty', 'est_cost', 'status', 'approver', 'decision_ts', 'override?', 'justification', 'cost_center', 'gl_code']), order);
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
  const sheet = getOrCreateSheet_(SHEETS.ORDERS, ['id', 'ts', 'requester', 'item', 'qty', 'est_cost', 'status', 'approver', 'decision_ts', 'override?', 'justification', 'cost_center', 'gl_code']);
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
      const cc = row[headers.cost_center];
      const month = String(row[headers.ts]).slice(0, 7);
      if (decision === 'APPROVED') {
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
