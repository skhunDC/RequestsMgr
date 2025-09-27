/* eslint-env googleappsscript */

const SCRIPT_PROP_SHEET_ID = 'SUPPLIES_TRACKING_SHEET_ID';
const MAX_PAGE_SIZE = 50;

const SHEETS = {
  CATALOG: 'Catalog',
  LOGS: 'Logs',
  STATUS_LOG: 'StatusLog'
};

const LOCATION_OPTIONS = ['Plant', 'Short North', 'South Dublin', 'Muirfield', 'Morse Rd.', 'Granville', 'Newark'];

const REQUEST_TYPES = {
  supplies: {
    sheetName: 'SuppliesRequests',
    headers: ['id', 'ts', 'requester', 'description', 'qty', 'location', 'notes', 'status', 'approver', 'eta'],
    normalize(request) {
      const location = normalizeLocation_(request && request.location);
      const description = sanitizeString_(request && request.description);
      if (!description) {
        throw new Error('Description is required.');
      }
      const qty = parsePositiveInteger_(request && request.qty);
      if (!qty) {
        throw new Error('Quantity must be at least 1.');
      }
      const notes = sanitizeString_(request && request.notes);
      return { description, qty, location, notes };
    },
    buildSummary(fields) {
      return fields.description || 'Supplies request';
    },
    buildDetails(fields) {
      const details = [];
      if (fields.location) {
        details.push(`Location: ${fields.location}`);
      }
      if (fields.qty) {
        details.push(`Quantity: ${fields.qty}`);
      }
      if (fields.notes) {
        details.push(`Notes: ${fields.notes}`);
      }
      if (fields.eta) {
        details.push(`ETA: ${fields.eta}`);
      }
      return details;
    }
  },
  it: {
    sheetName: 'ITRequests',
    headers: ['id', 'ts', 'requester', 'issue', 'device', 'urgency', 'details', 'status', 'approver', 'location'],
    normalize(request) {
      const location = normalizeLocation_(request && request.location);
      const issue = sanitizeString_(request && request.issue);
      if (!issue) {
        throw new Error('Issue summary is required.');
      }
      const device = sanitizeString_(request && request.device);
      const urgencyRaw = sanitizeString_(request && request.urgency).toLowerCase();
      const allowed = ['low', 'normal', 'high', 'critical'];
      const urgency = allowed.indexOf(urgencyRaw) === -1 ? 'normal' : urgencyRaw;
      const details = sanitizeString_(request && request.details);
      return { location, issue, device, urgency, details };
    },
    buildSummary(fields) {
      return fields.issue || 'IT request';
    },
    buildDetails(fields) {
      const details = [];
      if (fields.location) {
        details.push(`Location: ${fields.location}`);
      }
      if (fields.device) {
        details.push(`Device/System: ${fields.device}`);
      }
      if (fields.urgency) {
        details.push(`Urgency: ${capitalize_(fields.urgency)}`);
      }
      if (fields.details) {
        details.push(`Details: ${fields.details}`);
      }
      return details;
    }
  },
  maintenance: {
    sheetName: 'MaintenanceRequests',
    headers: ['id', 'ts', 'requester', 'location', 'issue', 'urgency', 'accessNotes', 'status', 'approver'],
    normalize(request) {
      const location = normalizeLocation_(request && request.location);
      const issue = sanitizeString_(request && request.issue);
      if (!issue) {
        throw new Error('Issue description is required.');
      }
      const urgencyRaw = sanitizeString_(request && request.urgency).toLowerCase();
      const allowed = ['low', 'normal', 'high', 'critical'];
      const urgency = allowed.indexOf(urgencyRaw) === -1 ? 'normal' : urgencyRaw;
      const accessNotes = sanitizeString_(request && request.accessNotes);
      return { location, issue, urgency, accessNotes };
    },
    buildSummary(fields) {
      return fields.issue || 'Maintenance request';
    },
    buildDetails(fields) {
      const details = [];
      if (fields.location) {
        details.push(`Location: ${fields.location}`);
      }
      if (fields.urgency) {
        details.push(`Urgency: ${capitalize_(fields.urgency)}`);
      }
      if (fields.accessNotes) {
        details.push(`Access notes: ${fields.accessNotes}`);
      }
      return details;
    }
  }
};

const LOG_HEADERS = ['ts', 'actor', 'fn', 'cid', 'message', 'stack', 'context'];
const STATUS_LOG_HEADERS = ['ts', 'type', 'requestId', 'actor', 'status'];

const CACHE_KEYS = {
  CATALOG: 'catalog:v1',
  REQUESTS_PREFIX: 'requests',
  RID_PREFIX: 'rid'
};

const CACHE_TTLS = {
  CATALOG: 300,
  REQUESTS: 180,
  RID: 300
};

function getRequiredSheetDefinitions_() {
  const definitions = {};
  Object.keys(REQUEST_TYPES).forEach(type => {
    const def = REQUEST_TYPES[type];
    definitions[def.sheetName] = def.headers.slice();
  });
  definitions[SHEETS.CATALOG] = ['sku', 'description', 'category', 'archived'];
  definitions[SHEETS.LOGS] = LOG_HEADERS.slice();
  definitions[SHEETS.STATUS_LOG] = STATUS_LOG_HEADERS.slice();
  return definitions;
}

function doGet() {
  ensureSetup_();
  const template = HtmlService.createTemplateFromFile('index');
  template.session = {
    email: getActiveUserEmail_()
  };
  return template.evaluate().setTitle('Request Manager');
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
      const sheet = getSheet_(SHEETS.CATALOG, ['sku', 'description', 'category', 'archived']);
      items = readTable_(sheet, ['sku', 'description', 'category', 'archived'])
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

function listRequests(request) {
  return withErrorHandling_('listRequests', request && request.cid, request, () => {
    ensureSetup_();
    const type = normalizeType_(request && request.type);
    const def = REQUEST_TYPES[type];
    const email = normalizeEmail_(getActiveUserEmail_());
    const scope = request && request.scope === 'all' ? 'all' : 'mine';
    const scopeKey = scope === 'all' ? 'all' : email;
    const pageSize = clamp_(Number(request && request.pageSize) || 15, 1, MAX_PAGE_SIZE);
    const startIndex = Number(request && request.nextToken) || 0;

    const cacheKey = [CACHE_KEYS.REQUESTS_PREFIX, type, scopeKey].join(':');
    const cache = CacheService.getScriptCache();
    let records = [];
    const cached = cache.get(cacheKey);
    if (cached) {
      records = JSON.parse(cached);
    } else {
      const sheet = getSheet_(def.sheetName, def.headers);
      const rows = readTable_(sheet, def.headers);
      const filtered = scope === 'all'
        ? rows
        : rows.filter(row => normalizeEmail_(row.requester) === email);
      records = filtered
        .map(row => buildClientRequest_(type, row))
        .sort((a, b) => (b.ts || '').localeCompare(a.ts || ''));
      cache.put(cacheKey, JSON.stringify(records), CACHE_TTLS.REQUESTS);
    }

    const slice = records.slice(startIndex, startIndex + pageSize);
    const nextToken = startIndex + slice.length < records.length ? String(startIndex + slice.length) : '';
    return {
      ok: true,
      type,
      requests: slice,
      nextToken,
      scope
    };
  });
}

function createRequest(request) {
  return withErrorHandling_('createRequest', request && request.cid, request, () => {
    ensureSetup_();
    const rid = String(request && request.clientRequestId || '').trim();
    if (!rid) {
      throw new Error('clientRequestId is required.');
    }
    const type = normalizeType_(request && request.type);
    const def = REQUEST_TYPES[type];

    const cache = CacheService.getScriptCache();
    const ridKey = [CACHE_KEYS.RID_PREFIX, rid].join(':');
    const existing = cache.get(ridKey);
    if (existing) {
      return {
        ok: true,
        request: JSON.parse(existing)
      };
    }

    const fields = def.normalize(request);
    const email = normalizeEmail_(getActiveUserEmail_());
    const now = new Date();
    const record = {
      id: uuid_(),
      ts: toIsoString_(now),
      requester: email,
      status: 'pending',
      approver: '',
      type,
      fields
    };

    const rowValues = def.headers.map(header => {
      switch (header) {
        case 'id':
          return record.id;
        case 'ts':
          return record.ts;
        case 'requester':
          return record.requester;
        case 'status':
          return record.status;
        case 'approver':
          return record.approver;
        default:
          return Object.prototype.hasOwnProperty.call(fields, header) ? fields[header] : '';
      }
    });

    const sheet = getSheet_(def.sheetName, def.headers);
    withLock_(() => {
      sheet.appendRow(rowValues);
    });

    const rowObject = Object.assign({}, fields, {
      id: record.id,
      ts: record.ts,
      requester: record.requester,
      status: record.status,
      approver: record.approver
    });
    const clientRecord = buildClientRequest_(type, rowObject);

    cache.put(ridKey, JSON.stringify(clientRecord), CACHE_TTLS.RID);
    invalidateRequestCache_(type, email);
    invalidateRequestCache_(type, 'all');

    return {
      ok: true,
      request: clientRecord
    };
  });
}

function updateRequestStatus(request) {
  return withErrorHandling_('updateRequestStatus', request && request.cid, request, () => {
    ensureSetup_();
    const rid = String(request && request.clientRequestId || '').trim();
    if (!rid) {
      throw new Error('clientRequestId is required.');
    }
    const type = normalizeType_(request && request.type);
    const def = REQUEST_TYPES[type];
    const requestId = String(request && request.requestId || '').trim();
    if (!requestId) {
      throw new Error('requestId is required.');
    }
    const status = normalizeStatus_(request && request.status);
    const hasEta = request && Object.prototype.hasOwnProperty.call(request, 'eta');
    const etaValue = hasEta ? sanitizeString_(request.eta) : '';
    if (type === 'supplies' && status === 'ordered' && !etaValue) {
      throw new Error('ETA is required when marking a supplies request as ordered.');
    }

    const cache = CacheService.getScriptCache();
    const ridKey = [CACHE_KEYS.RID_PREFIX, rid].join(':');
    const cached = cache.get(ridKey);
    if (cached) {
      return {
        ok: true,
        request: JSON.parse(cached)
      };
    }

    const sheet = getSheet_(def.sheetName, def.headers);
    const headers = sheet.getRange(1, 1, 1, def.headers.length).getValues()[0];
    const idIdx = headers.indexOf('id');
    if (idIdx === -1) {
      throw new Error('Request sheet is misconfigured.');
    }

    let updatedRecord = null;
    const headerMap = mapHeaders_(headers);
    const approverEmail = getActiveUserEmail_();

    withLock_(() => {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return;
      }
      const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
      const data = dataRange.getValues();
      for (let r = 0; r < data.length; r++) {
        if (String(data[r][idIdx]).trim() === requestId) {
          data[r][headerMap.status] = status;
          data[r][headerMap.approver] = approverEmail;
          sheet.getRange(r + 2, headerMap.status + 1).setValue(status);
          sheet.getRange(r + 2, headerMap.approver + 1).setValue(approverEmail);
          if (headerMap.eta !== undefined && hasEta) {
            data[r][headerMap.eta] = etaValue;
            sheet.getRange(r + 2, headerMap.eta + 1).setValue(etaValue);
          }
          const rowObject = {};
          headers.forEach((header, idx) => {
            rowObject[header] = data[r][idx];
          });
          rowObject.status = status;
          rowObject.approver = approverEmail;
          updatedRecord = buildClientRequest_(type, rowObject);
          break;
        }
      }
    });

    if (!updatedRecord) {
      throw new Error('Request not found.');
    }

    cache.put(ridKey, JSON.stringify(updatedRecord), CACHE_TTLS.RID);
    invalidateRequestCache_(type, normalizeEmail_(updatedRecord.requester));
    invalidateRequestCache_(type, 'all');

    recordStatusAction_(type, requestId, status, approverEmail);

    return {
      ok: true,
      request: updatedRecord
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
    const requiredSheets = getRequiredSheetDefinitions_();

    Object.keys(requiredSheets).forEach(name => {
      const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
      normalizeSheetStructure_(sheet, requiredSheets[name]);
    });

    ss.getSheets().forEach(sheet => {
      const name = sheet.getName();
      if (!Object.prototype.hasOwnProperty.call(requiredSheets, name)) {
        ss.deleteSheet(sheet);
      }
    });

    const catalog = ss.getSheetByName(SHEETS.CATALOG);
    if (catalog.getLastRow() <= 1) {
      const defaults = [
        ['SKU-001', 'Copy Paper 8.5x11 (case)', 'Office', false],
        ['SKU-014', 'Nitrile Gloves (box)', 'Cleaning', false],
        ['SKU-027', 'Poly Garment Bags (roll)', 'Operations', false]
      ];
      catalog.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
    }
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
    ss = SpreadsheetApp.create('RequestManager');
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

function normalizeSheetStructure_(sheet, headers) {
  const totalRows = sheet.getLastRow();
  const totalColumns = sheet.getLastColumn();
  let rows = [];
  if (totalRows > 1 && totalColumns > 0) {
    const headerRow = sheet.getRange(1, 1, 1, totalColumns).getValues()[0];
    const headerMap = {};
    headerRow.forEach((header, idx) => {
      const key = String(header || '').trim().toLowerCase();
      if (key && headerMap[key] === undefined) {
        headerMap[key] = idx;
      }
    });
    const dataRange = sheet.getRange(2, 1, totalRows - 1, totalColumns);
    const data = dataRange.getValues();
    rows = data.map(row => headers.map(header => {
      const idx = headerMap[String(header).toLowerCase()];
      return idx === undefined ? '' : row[idx];
    }));
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  const maxColumns = sheet.getMaxColumns();
  if (maxColumns > headers.length) {
    sheet.deleteColumns(headers.length + 1, maxColumns - headers.length);
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
  const aliasMap = {
    complete: 'completed',
    completed: 'completed',
    'in progress': 'in_progress',
    'in-progress': 'in_progress'
  };
  const normalized = aliasMap[value] || value;
  const allowed = ['pending', 'completed', 'in_progress', 'declined', 'approved', 'denied', 'ordered'];
  if (allowed.indexOf(normalized) === -1) {
    throw new Error('Unsupported status.');
  }
  return normalized;
}

function normalizeType_(type) {
  const value = String(type || '').trim().toLowerCase();
  if (!value || !REQUEST_TYPES[value]) {
    throw new Error('Unsupported request type.');
  }
  return value;
}

function invalidateRequestCache_(type, key) {
  const cache = CacheService.getScriptCache();
  const normalizedKey = key === 'all' ? 'all' : normalizeEmail_(key);
  const cacheKey = [CACHE_KEYS.REQUESTS_PREFIX, type, normalizedKey].join(':');
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

function normalizeLocation_(value) {
  const name = sanitizeString_(value);
  if (!name) {
    throw new Error('Location is required.');
  }
  const match = LOCATION_OPTIONS.find(option => option.toLowerCase() === name.toLowerCase());
  if (!match) {
    throw new Error('Unsupported location.');
  }
  return match;
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

function recordStatusAction_(type, requestId, status, actor) {
  const sheet = getSheet_(SHEETS.STATUS_LOG, STATUS_LOG_HEADERS);
  const entry = [
    toIsoString_(new Date()),
    String(type || ''),
    String(requestId || ''),
    normalizeEmail_(actor),
    String(status || '')
  ];
  withLock_(() => {
    sheet.appendRow(entry);
  });
}

function getFieldNames_(headers) {
  const base = ['id', 'ts', 'requester', 'status', 'approver'];
  return headers.filter(header => base.indexOf(header) === -1);
}

function buildClientRequest_(type, row) {
  const def = REQUEST_TYPES[type];
  const fieldNames = getFieldNames_(def.headers);
  const fields = {};
  fieldNames.forEach(name => {
    fields[name] = row[name] !== undefined ? row[name] : '';
  });
  const record = {
    id: String(row.id || ''),
    ts: String(row.ts || ''),
    requester: String(row.requester || ''),
    status: String(row.status || 'pending').toLowerCase() || 'pending',
    approver: String(row.approver || ''),
    type,
    fields
  };
  record.summary = def.buildSummary(fields);
  record.details = def.buildDetails(fields);
  return record;
}

function capitalize_(value) {
  const text = String(value || '');
  if (!text) {
    return '';
  }
  return text.charAt(0).toUpperCase() + text.slice(1);
}
