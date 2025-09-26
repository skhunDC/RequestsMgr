/* eslint-env googleappsscript */

const SHEETS = {
  ORDERS: 'Orders',
  CATALOG: 'Catalog',
  LOGS: 'Logs'
};

const ORDER_HEADERS = ['id', 'ts', 'requester', 'description', 'qty', 'status', 'approver'];
const CATALOG_HEADERS = ['sku', 'description', 'category', 'archived'];
const LOG_HEADERS = ['ts', 'actor', 'fn', 'cid', 'message', 'stack', 'context'];

const CACHE_KEYS = {
  CATALOG: 'catalog:v1',
  ORDERS_PREFIX: 'orders',
  RID_PREFIX: 'rid'
};

const CACHE_TTLS = {
  CATALOG: 300,
  ORDERS: 180,
  RID: 300
};

const SCRIPT_PROP_SHEET_ID = 'SUPPLIES_TRACKING_SHEET_ID';
const MAX_PAGE_SIZE = 50;

function doGet() {
  ensureSetup_();
  const template = HtmlService.createTemplateFromFile('index');
  template.session = {
    email: getActiveUserEmail_()
  };
  return template.evaluate().setTitle('Supplies Tracker');
}

function listCatalog(request) {
  return withErrorHandling_('listCatalog', request && request.cid, request, () => {
    ensureSetup_();
    const pageSize = clamp_(Number(request && request.pageSize) || 20, 1, MAX_PAGE_SIZE);
    const startIndex = Number(request && request.nextToken) || 0;

    const cache = CacheService.getScriptCache();
    let items = [];
    const cached = cache.get(CACHE_KEYS.CATALOG);
    if (cached) {
      items = JSON.parse(cached);
    } else {
      const sheet = getSheet_(SHEETS.CATALOG, CATALOG_HEADERS);
      items = readTable_(sheet, CATALOG_HEADERS)
        .filter(row => !row.archived)
        .map(row => ({
          sku: row.sku,
          description: row.description,
          category: row.category
        }));
      cache.put(CACHE_KEYS.CATALOG, JSON.stringify(items), CACHE_TTLS.CATALOG);
    }

    const slice = items.slice(startIndex, startIndex + pageSize);
    const nextToken = startIndex + slice.length < items.length ? String(startIndex + slice.length) : '';
    return {
      ok: true,
      items: slice,
      nextToken
    };
  });
}

function listOrders(request) {
  return withErrorHandling_('listOrders', request && request.cid, request, () => {
    ensureSetup_();
    const email = normalizeEmail_(getActiveUserEmail_());
    const scope = request && request.scope === 'all' ? 'all' : 'mine';
    const tokenKey = scope === 'all' ? 'all' : email;
    const pageSize = clamp_(Number(request && request.pageSize) || 15, 1, MAX_PAGE_SIZE);
    const startIndex = Number(request && request.nextToken) || 0;

    const cacheKey = [CACHE_KEYS.ORDERS_PREFIX, tokenKey].join(':');
    const cache = CacheService.getScriptCache();
    let orders = [];
    const cached = cache.get(cacheKey);
    if (cached) {
      orders = JSON.parse(cached);
    } else {
      const sheet = getSheet_(SHEETS.ORDERS, ORDER_HEADERS);
      const rows = readTable_(sheet, ORDER_HEADERS);
      const filtered = scope === 'all'
        ? rows
        : rows.filter(row => normalizeEmail_(row.requester) === email);
      orders = filtered
        .map(row => ({
          id: row.id,
          ts: row.ts,
          requester: row.requester,
          description: row.description,
          qty: Number(row.qty) || 0,
          status: row.status || 'pending',
          approver: row.approver || ''
        }))
        .sort((a, b) => (b.ts || '').localeCompare(a.ts || ''));
      cache.put(cacheKey, JSON.stringify(orders), CACHE_TTLS.ORDERS);
    }

    const slice = orders.slice(startIndex, startIndex + pageSize);
    const nextToken = startIndex + slice.length < orders.length ? String(startIndex + slice.length) : '';
    return {
      ok: true,
      orders: slice,
      nextToken,
      scope
    };
  });
}

function createOrder(request) {
  return withErrorHandling_('createOrder', request && request.cid, request, () => {
    ensureSetup_();
    const rid = String(request && request.clientRequestId || '').trim();
    if (!rid) {
      throw new Error('clientRequestId is required.');
    }
    const cache = CacheService.getScriptCache();
    const ridKey = [CACHE_KEYS.RID_PREFIX, rid].join(':');
    const existing = cache.get(ridKey);
    if (existing) {
      return {
        ok: true,
        order: JSON.parse(existing)
      };
    }

    const description = sanitizeString_(request && request.description);
    if (!description) {
      throw new Error('Description is required.');
    }
    const qty = parsePositiveInteger_(request && request.qty);
    if (!qty) {
      throw new Error('Quantity must be at least 1.');
    }

    const email = normalizeEmail_(getActiveUserEmail_());
    const now = new Date();
    const order = {
      id: uuid_(),
      ts: toIsoString_(now),
      requester: email,
      description,
      qty,
      status: 'pending',
      approver: ''
    };

    const sheet = getSheet_(SHEETS.ORDERS, ORDER_HEADERS);
    withLock_(() => {
      sheet.appendRow([
        order.id,
        order.ts,
        order.requester,
        order.description,
        order.qty,
        order.status,
        order.approver
      ]);
    });

    cache.put(ridKey, JSON.stringify(order), CACHE_TTLS.RID);
    invalidateOrdersCache_(email);
    invalidateOrdersCache_('all');

    return {
      ok: true,
      order
    };
  });
}

function updateOrderStatus(request) {
  return withErrorHandling_('updateOrderStatus', request && request.cid, request, () => {
    ensureSetup_();
    const rid = String(request && request.clientRequestId || '').trim();
    if (!rid) {
      throw new Error('clientRequestId is required.');
    }
    const orderId = String(request && request.orderId || '').trim();
    if (!orderId) {
      throw new Error('orderId is required.');
    }
    const status = normalizeStatus_(request && request.status);
    const cache = CacheService.getScriptCache();
    const ridKey = [CACHE_KEYS.RID_PREFIX, rid].join(':');
    const cached = cache.get(ridKey);
    if (cached) {
      return {
        ok: true,
        order: JSON.parse(cached)
      };
    }

    const sheet = getSheet_(SHEETS.ORDERS, ORDER_HEADERS);
    const headers = sheet.getRange(1, 1, 1, ORDER_HEADERS.length).getValues()[0];
    const idIdx = headers.indexOf('id');
    if (idIdx === -1) {
      throw new Error('Orders sheet is misconfigured.');
    }

    let updatedOrder = null;
    const headerMap = mapHeaders_(headers);
    const statusIdx = headerMap.status;
    const approverIdx = headerMap.approver;
    const requesterIdx = headerMap.requester;
    const qtyIdx = headerMap.qty;
    const descIdx = headerMap.description;
    const tsIdx = headerMap.ts;
    if ([statusIdx, approverIdx, requesterIdx, qtyIdx, descIdx, tsIdx].some(idx => typeof idx !== 'number')) {
      throw new Error('Orders sheet is misconfigured.');
    }
    const approverEmail = getActiveUserEmail_();
    withLock_(() => {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return;
      }
      const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
      const data = dataRange.getValues();
      for (let r = 0; r < data.length; r++) {
        if (String(data[r][idIdx]).trim() === orderId) {
          data[r][statusIdx] = status;
          data[r][approverIdx] = approverEmail;
          sheet.getRange(r + 2, statusIdx + 1).setValue(status);
          sheet.getRange(r + 2, approverIdx + 1).setValue(approverEmail);
          updatedOrder = {
            id: orderId,
            ts: String(data[r][tsIdx] || ''),
            requester: String(data[r][requesterIdx] || ''),
            description: String(data[r][descIdx] || ''),
            qty: Number(data[r][qtyIdx]) || 0,
            status,
            approver: approverEmail
          };
          break;
        }
      }
    });

    if (!updatedOrder) {
      throw new Error('Order not found.');
    }

    cache.put(ridKey, JSON.stringify(updatedOrder), CACHE_TTLS.RID);
    invalidateOrdersCache_(normalizeEmail_(updatedOrder.requester));
    invalidateOrdersCache_('all');

    return {
      ok: true,
      order: updatedOrder
    };
  });
}

function logClientError(request) {
  return withErrorHandling_('logClientError', request && request.cid, request, () => {
    ensureSetup_();
    const sheet = getSheet_(SHEETS.LOGS, LOG_HEADERS);
    const entry = [
      toIsoString_(new Date()),
      normalizeEmail_(getActiveUserEmail_()),
      String(request && request.context || ''),
      String(request && request.cid || ''),
      String(request && request.message || ''),
      String(request && request.stack || ''),
      String(request && request.payload ? JSON.stringify(request.payload) : '')
    ];
    withLock_(() => {
      sheet.appendRow(entry);
    });
    return { ok: true };
  });
}

function withErrorHandling_(fnName, cid, context, fn) {
  try {
    return fn();
  } catch (err) {
    logServerError_(fnName, cid, err, context);
    return {
      ok: false,
      code: 'SERVER_ERROR',
      message: err && err.message ? err.message : 'Unexpected error.'
    };
  }
}

function ensureSetup_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Unable to acquire initialization lock.');
  }
  try {
    const ss = getSpreadsheet_();
    const orders = ss.getSheetByName(SHEETS.ORDERS) || ss.insertSheet(SHEETS.ORDERS);
    ensureHeaders_(orders, ORDER_HEADERS);

    const catalog = ss.getSheetByName(SHEETS.CATALOG) || ss.insertSheet(SHEETS.CATALOG);
    ensureHeaders_(catalog, CATALOG_HEADERS);
    if (catalog.getLastRow() <= 1) {
      const defaults = [
        ['SKU-001', 'Copy Paper 8.5x11 (case)', 'Office', false],
        ['SKU-014', 'Nitrile Gloves (box)', 'Cleaning', false],
        ['SKU-027', 'Poly Garment Bags (roll)', 'Operations', false]
      ];
      catalog.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
    }

    const logs = ss.getSheetByName(SHEETS.LOGS) || ss.insertSheet(SHEETS.LOGS);
    ensureHeaders_(logs, LOG_HEADERS);
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty(SCRIPT_PROP_SHEET_ID);
  let ss = null;
  if (storedId) {
    try {
      ss = SpreadsheetApp.openById(storedId);
    } catch (err) {
      ss = null;
    }
  }
  if (!ss) {
    ss = SpreadsheetApp.getActive();
  }
  if (!ss) {
    ss = SpreadsheetApp.create('SuppliesTracking');
  }
  if (ss && ss.getId() !== storedId) {
    props.setProperty(SCRIPT_PROP_SHEET_ID, ss.getId());
  }
  return ss;
}

function getSheet_(name, headers) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureHeaders_(sheet, headers);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.appendRow(headers);
    return;
  }
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let updated = false;
  headers.forEach((header, idx) => {
    if (existing[idx] !== header) {
      sheet.getRange(1, idx + 1).setValue(header);
      updated = true;
    }
  });
  if (updated && sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

function readTable_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  const range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  const values = range.getValues();
  return values.map(row => {
    const record = {};
    headers.forEach((header, idx) => {
      record[header] = row[idx];
    });
    if (record.archived !== undefined) {
      record.archived = record.archived === true || String(record.archived).toLowerCase() === 'true';
    }
    return record;
  });
}

function mapHeaders_(headers) {
  const map = {};
  headers.forEach((header, idx) => {
    map[header] = idx;
  });
  return map;
}

function normalizeStatus_(status) {
  const value = String(status || '').trim().toLowerCase();
  if (!value) {
    throw new Error('status is required.');
  }
  const allowed = ['pending', 'approved', 'declined'];
  if (allowed.indexOf(value) === -1) {
    throw new Error('Unsupported status.');
  }
  return value;
}

function invalidateOrdersCache_(key) {
  const cache = CacheService.getScriptCache();
  const cacheKey = [CACHE_KEYS.ORDERS_PREFIX, normalizeEmail_(key)].join(':');
  cache.remove(cacheKey);
}

function parsePositiveInteger_(value) {
  const num = Math.floor(Number(value));
  return num > 0 ? num : 0;
}

function sanitizeString_(value) {
  return String(value || '').trim();
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function toIsoString_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'UTC', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function clamp_(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function uuid_() {
  return Utilities.getUuid();
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Could not obtain lock.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getActiveUserEmail_() {
  return normalizeEmail_(Session.getActiveUser().getEmail());
}

function logServerError_(fnName, cid, err, context) {
  const payload = {
    ts: toIsoString_(new Date()),
    fn: fnName,
    cid: cid || '',
    message: err && err.message ? err.message : String(err),
    stack: err && err.stack ? err.stack : '',
    context: context ? JSON.stringify(context) : ''
  };
  Logger.log(JSON.stringify(payload));
  try {
    const sheet = getSheet_(SHEETS.LOGS, LOG_HEADERS);
    withLock_(() => {
      sheet.appendRow([
        payload.ts,
        normalizeEmail_(getActiveUserEmail_()),
        payload.fn,
        payload.cid,
        payload.message,
        payload.stack,
        payload.context
      ]);
    });
  } catch (logErr) {
    Logger.log('Failed to log to sheet: ' + logErr);
  }
}
