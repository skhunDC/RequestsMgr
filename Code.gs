/* eslint-env googleappsscript */

const SCRIPT_PROP_SHEET_ID = 'SUPPLIES_TRACKING_SHEET_ID';
const SCRIPT_PROP_SETUP_VERSION = 'SUPPLIES_TRACKING_SETUP_VERSION';
const SCRIPT_PROP_STATUS_EMAILS = 'SUPPLIES_TRACKING_STATUS_EMAILS';
const CURRENT_SETUP_VERSION = '4';
const MAX_PAGE_SIZE = 50;

const SHEETS = {
  CATALOG: 'Catalog',
  LOGS: 'Logs',
  STATUS_LOG: 'StatusLog',
  REQUEST_NOTES: 'RequestNotes'
};

const LOCATION_OPTIONS = ['Plant', 'Short North', 'South Dublin', 'Muirfield', 'Morse Rd.', 'Granville', 'Newark'];

const EMAIL_TIMEZONE = 'America/New_York';
const PRIMARY_NOTIFICATION_EMAIL = 'skhun@dublincleaners.com';
const EMAIL_SENDER_NAME = 'Request Manager';
const REQUEST_MANAGER_APP_URL = 'https://script.google.com/macros/s/AKfycbxf6fr9FKGjQCPE31Li-woofA6k8H7SqNcO09HayFdKfJBeSiQJXIfOd_bJ4MVfynoJag/exec';

const REQUEST_TYPES = {
  supplies: {
    sheetName: 'SuppliesRequests',
    headers: ['id', 'ts', 'requester', 'description', 'qty', 'location', 'notes', 'eta', 'status', 'approver'],
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
      if (fields.supplier) {
        details.push(`Supplier: ${fields.supplier}`);
      }
      if (fields.estimatedCost) {
        const estimatedCostDetail = buildSuppliesEstimatedCostDetail_(fields);
        if (estimatedCostDetail) {
          details.push(estimatedCostDetail);
        }
      }
      if (fields.notes) {
        details.push(`Notes: ${fields.notes}`);
      }
      if (fields.eta) {
        const formatted = formatDateForDisplay_(fields.eta);
        if (formatted) {
          details.push(`ETA: ${formatted}`);
        }
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
      const urgency = normalizeUrgencyValue_(request && request.urgency);
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
        const urgency = normalizeUrgencyValue_(fields.urgency);
        details.push(`Urgency: ${capitalize_(urgency)}`);
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
      const urgency = normalizeUrgencyValue_(request && request.urgency);
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
        const urgency = normalizeUrgencyValue_(fields.urgency);
        details.push(`Urgency: ${capitalize_(urgency)}`);
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
const REQUEST_NOTE_HEADERS = ['ts', 'type', 'requestId', 'actor', 'note'];

const CACHE_KEYS = {
  CATALOG: 'catalog:v3',
  CATALOG_USAGE: 'catalog-usage:v1',
  CATALOG_DESC_INDEX: 'catalog-by-desc:v1',
  REQUESTS_PREFIX: 'requests',
  RID_PREFIX: 'rid',
  STATUS_EMAILS: 'status-emails:v1'
};

const DEVICE_RATE_LIMIT = Object.freeze({
  MAX_REQUESTS: 12,
  WINDOW_MS: 24 * 60 * 60 * 1000,
  PROP_PREFIX: 'device-limit'
});

const DEFAULT_STATUS_APPROVER_EMAILS = Object.freeze([
  'skhun@dublincleaners.com',
  'ss.sku@protonmail.com',
  'rbown@dublincleaners.com',
  'bbutler@dublincleaners.com',
  'mlackey@dublincleaners.com',
  'rbrown5940@gmail.com',
  'brianmbutler77@gmail.com'
].map(normalizeEmail_));

const CACHE_TTLS = {
  CATALOG: 300,
  REQUESTS: 180,
  RID: 300,
  STATUS_EMAILS: 300
};

let runtimeCatalogItems_ = null;
let runtimeCatalogDescriptionIndex_ = null;

function supportsRequestNotes_(type) {
  return type === 'it' || type === 'maintenance';
}

function getRequestNotesMap_(type) {
  if (!supportsRequestNotes_(type)) {
    return {};
  }
  const sheet = getSheet_(SHEETS.REQUEST_NOTES, REQUEST_NOTE_HEADERS);
  const rows = readTable_(sheet, REQUEST_NOTE_HEADERS);
  const map = {};
  rows.forEach(entry => {
    const entryType = String(entry && entry.type || '').trim().toLowerCase();
    if (entryType !== type) {
      return;
    }
    const requestId = String(entry && entry.requestId || '').trim();
    const noteText = sanitizeString_(entry && entry.note);
    if (!requestId || !noteText) {
      return;
    }
    let tsValue = '';
    const rawTs = entry && entry.ts;
    if (rawTs instanceof Date && !isNaN(rawTs.getTime())) {
      tsValue = toIsoString_(rawTs);
    } else {
      tsValue = sanitizeString_(rawTs);
    }
    const rawActor = sanitizeString_(entry && entry.actor);
    const actor = normalizeEmail_(rawActor) || rawActor;
    if (!Array.isArray(map[requestId])) {
      map[requestId] = [];
    }
    map[requestId].push({
      ts: tsValue,
      actor,
      note: noteText
    });
  });
  Object.keys(map).forEach(requestId => {
    map[requestId].sort((a, b) => String(b.ts || '').localeCompare(String(a.ts || '')));
  });
  return map;
}

function getRequiredSheetDefinitions_() {
  const definitions = {};
  Object.keys(REQUEST_TYPES).forEach(type => {
    const def = REQUEST_TYPES[type];
    definitions[def.sheetName] = def.headers.slice();
  });
  definitions[SHEETS.CATALOG] = ['sku', 'description', 'category', 'estimatedCost', 'supplier', 'archived'];
  definitions[SHEETS.LOGS] = LOG_HEADERS.slice();
  definitions[SHEETS.STATUS_LOG] = STATUS_LOG_HEADERS.slice();
  definitions[SHEETS.REQUEST_NOTES] = REQUEST_NOTE_HEADERS.slice();
  return definitions;
}

function doGet() {
  ensureSetup_();
  const template = HtmlService.createTemplateFromFile('index');
  const auth = getStatusAuthContext_();
  template.session = {
    email: auth.email,
    canManageStatuses: auth.authorized,
    statusAuth: {
      email: auth.email,
      authorized: auth.authorized,
      reason: auth.reason,
      allowlistSource: auth.allowlistSource,
      allowlistSize: auth.allowlistSize
    }
  };
  return template.evaluate().setTitle('Request Manager');
}

function listCatalog(request) {
  return withErrorHandling_('listCatalog', request && request.cid, request, () => {
    ensureSetup_();
    const fetchAll = Boolean(request && request.fetchAll);
    const pageSize = clamp_(Number(request && request.pageSize) || 20, 1, MAX_PAGE_SIZE);
    const startIndex = fetchAll ? 0 : Number(request && request.nextToken) || 0;

    const items = getCatalogItems_();
    const slice = fetchAll ? items : items.slice(startIndex, startIndex + pageSize);
    const nextToken = fetchAll || startIndex + slice.length >= items.length ? '' : String(startIndex + slice.length);
    return {
      ok: true,
      items: slice,
      nextToken
    };
  });
}

function getCatalogItems_() {
  if (Array.isArray(runtimeCatalogItems_)) {
    return runtimeCatalogItems_;
  }
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEYS.CATALOG);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed)) {
        runtimeCatalogItems_ = parsed;
        return parsed;
      }
      cache.remove(CACHE_KEYS.CATALOG);
    } catch (err) {
      cache.remove(CACHE_KEYS.CATALOG);
    }
  }
  const items = buildCatalogItemsFromSheet_();
  cache.put(CACHE_KEYS.CATALOG, JSON.stringify(items), CACHE_TTLS.CATALOG);
  runtimeCatalogItems_ = items;
  return items;
}

function buildCatalogItemsFromSheet_() {
  const sheet = getSheet_(SHEETS.CATALOG, ['sku', 'description', 'category', 'estimatedCost', 'supplier', 'archived']);
  const usageCounts = getCatalogUsageCounts_();
  return readTable_(sheet, ['sku', 'description', 'category', 'estimatedCost', 'supplier', 'archived'])
    .filter(row => !row.archived)
    .map(row => {
      const description = sanitizeString_(row.description);
      const usageKey = description.toLowerCase();
      const usageCount = usageCounts[usageKey] || 0;
      return {
        sku: sanitizeString_(row.sku),
        description,
        category: sanitizeString_(row.category),
        estimatedCost: sanitizeString_(row.estimatedCost),
        supplier: sanitizeString_(row.supplier),
        usageCount
      };
    })
    .sort((a, b) => {
      if (b.usageCount !== a.usageCount) {
        return b.usageCount - a.usageCount;
      }
      const categoryCompare = String(a.category || '').localeCompare(String(b.category || ''), undefined, { sensitivity: 'base' });
      if (categoryCompare !== 0) {
        return categoryCompare;
      }
      return String(a.description || '').localeCompare(String(b.description || ''), undefined, { sensitivity: 'base' });
    });
}

function getCatalogUsageCounts_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEYS.CATALOG_USAGE);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      // ignore and rebuild
    }
  }

  const def = REQUEST_TYPES.supplies;
  const sheet = getSheet_(def.sheetName, def.headers);
  const rows = readTable_(sheet, def.headers);
  const counts = rows.reduce((acc, row) => {
    const description = sanitizeString_(row.description);
    if (!description) {
      return acc;
    }
    const key = description.toLowerCase();
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  cache.put(CACHE_KEYS.CATALOG_USAGE, JSON.stringify(counts), CACHE_TTLS.CATALOG);
  return counts;
}

function getCatalogDescriptionIndex_() {
  if (runtimeCatalogDescriptionIndex_ && typeof runtimeCatalogDescriptionIndex_ === 'object') {
    return runtimeCatalogDescriptionIndex_;
  }
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEYS.CATALOG_DESC_INDEX);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed && typeof parsed === 'object') {
        runtimeCatalogDescriptionIndex_ = parsed;
        return parsed;
      }
      cache.remove(CACHE_KEYS.CATALOG_DESC_INDEX);
    } catch (err) {
      cache.remove(CACHE_KEYS.CATALOG_DESC_INDEX);
    }
  }
  const items = getCatalogItems_();
  const index = items.reduce((acc, item) => {
    const key = sanitizeString_(item && item.description).toLowerCase();
    if (!key || acc[key]) {
      return acc;
    }
    acc[key] = {
      supplier: sanitizeString_(item && item.supplier),
      estimatedCost: sanitizeString_(item && item.estimatedCost),
      sku: sanitizeString_(item && item.sku),
      category: sanitizeString_(item && item.category)
    };
    return acc;
  }, {});
  cache.put(CACHE_KEYS.CATALOG_DESC_INDEX, JSON.stringify(index), CACHE_TTLS.CATALOG);
  runtimeCatalogDescriptionIndex_ = index;
  return index;
}

function getAllRequestsForType_(type) {
  const def = REQUEST_TYPES[type];
  if (!def) {
    throw new Error('Unsupported request type.');
  }
  const cache = CacheService.getScriptCache();
  const cacheKey = [CACHE_KEYS.REQUESTS_PREFIX, type, 'all'].join(':');
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed)) {
        return parsed;
      }
      cache.remove(cacheKey);
    } catch (err) {
      cache.remove(cacheKey);
    }
  }
  const sheet = getSheet_(def.sheetName, def.headers);
  const rows = readTable_(sheet, def.headers);
  const notesMap = getRequestNotesMap_(type);
  const records = rows
    .map(row => {
      const record = buildClientRequest_(type, row);
      record.notes = Array.isArray(notesMap[record.id]) ? notesMap[record.id] : [];
      return record;
    })
    .sort((a, b) => (b.ts || '').localeCompare(a.ts || ''));
  cache.put(cacheKey, JSON.stringify(records), CACHE_TTLS.REQUESTS);
  return records;
}

function listRequests(request) {
  return withErrorHandling_('listRequests', request && request.cid, request, () => {
    ensureSetup_();
    const type = normalizeType_(request && request.type);
    const scope = normalizeScope_(request && request.scope);
    const pageSize = clamp_(Number(request && request.pageSize) || 15, 1, MAX_PAGE_SIZE);
    const startIndex = Number(request && request.nextToken) || 0;

    const cache = CacheService.getScriptCache();
    const records = getAllRequestsForType_(type);

    const userEmail = normalizeEmail_(getActiveUserEmail_());
    let scopedRecords = records;
    if (scope === 'mine') {
      const mineCacheKey = [CACHE_KEYS.REQUESTS_PREFIX, type, userEmail || ''].join(':');
      const cachedMine = cache.get(mineCacheKey);
      if (cachedMine) {
        try {
          const parsedMine = JSON.parse(cachedMine);
          if (Array.isArray(parsedMine)) {
            scopedRecords = parsedMine;
          } else {
            cache.remove(mineCacheKey);
            scopedRecords = records.filter(record => normalizeEmail_(record.requester) === userEmail);
          }
        } catch (err) {
          cache.remove(mineCacheKey);
          scopedRecords = records.filter(record => normalizeEmail_(record.requester) === userEmail);
        }
      } else {
        scopedRecords = records.filter(record => normalizeEmail_(record.requester) === userEmail);
      }
      cache.put(mineCacheKey, JSON.stringify(scopedRecords), CACHE_TTLS.REQUESTS);
    }

    const slice = scopedRecords.slice(startIndex, startIndex + pageSize);
    const nextToken = startIndex + slice.length < scopedRecords.length ? String(startIndex + slice.length) : '';
    return {
      ok: true,
      type,
      scope,
      requests: slice,
      nextToken
    };
  });
}

function getDashboardMetrics(request) {
  return withErrorHandling_('getDashboardMetrics', request && request.cid, request, () => {
    ensureSetup_();
    const metrics = {};
    let totalRequests = 0;
    let outstandingRequests = 0;
    Object.keys(REQUEST_TYPES).forEach(type => {
      const records = getAllRequestsForType_(type);
      const total = records.length;
      const outstanding = records.reduce((count, record) => {
        const status = String(record && record.status || '').trim().toLowerCase();
        if (status === 'approved' || status === 'completed') {
          return count;
        }
        return count + 1;
      }, 0);
      metrics[type] = { total, outstanding };
      totalRequests += total;
      outstandingRequests += outstanding;
    });
    return {
      ok: true,
      metrics,
      totals: {
        totalRequests,
        outstandingRequests
      },
      generatedAt: toIsoString_(new Date())
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

    const deviceId = normalizeDeviceId_(request && request.deviceId);
    if (!deviceId) {
      throw new Error('Device identifier is required.');
    }

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
    const requesterName = sanitizeString_(request && request.requesterName);
    if (!email && !requesterName) {
      throw new Error('Your name is required to submit requests.');
    }
    const requesterIdentity = email || requesterName;
    const now = new Date();
    const nowMs = now.getTime();
    const record = {
      id: uuid_(),
      ts: toIsoString_(now),
      requester: requesterIdentity,
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
    const props = PropertiesService.getScriptProperties();
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
      throw new Error('Could not obtain lock.');
    }
    try {
      const limitState = evaluateDeviceRateLimit_(deviceId, nowMs, props);
      if (!limitState.allowed) {
        return limitState.response;
      }
      sheet.appendRow(rowValues);
      commitDeviceRateLimitUsage_(limitState, props, nowMs);
    } finally {
      lock.releaseLock();
    }

    const rowObject = Object.assign({}, fields, {
      id: record.id,
      ts: record.ts,
      requester: record.requester,
      status: record.status,
      approver: record.approver
    });
    const clientRecord = buildClientRequest_(type, rowObject);

    cache.put(ridKey, JSON.stringify(clientRecord), CACHE_TTLS.RID);
    if (type === 'supplies') {
      invalidateCatalogCache_();
    }
    invalidateRequestCache_(type, email);
    if (!email && requesterName) {
      invalidateRequestCache_(type, requesterName);
    }
    invalidateRequestCache_(type, 'all');

    if (type === 'it' || type === 'maintenance') {
      sendNewRequestNotification_(type, clientRecord);
    }

    return {
      ok: true,
      request: clientRecord
    };
  });
}

function sendWeeklySuppliesSummary() {
  const result = withErrorHandling_('sendWeeklySuppliesSummary', '', {}, () => {
    ensureSetup_();
    const records = getAllRequestsForType_('supplies');
    const outstanding = records.filter(record => {
      const statusKey = toStatusKey_(record.status);
      return statusKey !== 'approved' && statusKey !== 'denied';
    });
    sendSuppliesSummaryEmail_(outstanding);
    return { ok: true };
  });
  if (!result || result.ok !== true) {
    throw new Error(result && result.message ? result.message : 'Failed to send weekly supplies summary.');
  }
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
    const hasEta = Object.prototype.hasOwnProperty.call(request || {}, 'eta');
    const etaValue = hasEta ? normalizeDateOnly_(request && request.eta) : '';

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
    const statusCol = headerMap.status;
    if (statusCol === undefined) {
      throw new Error('Request sheet is missing a status column.');
    }
    const approverCol = headerMap.approver;
    if (approverCol === undefined) {
      throw new Error('Request sheet is missing an approver column.');
    }
    const etaCol = headerMap.eta;
    const statusAuth = assertAuthorizedStatusActor_();
    const approverEmail = statusAuth.email;

    withLock_(() => {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return;
      }
      const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
      const data = dataRange.getValues();
      for (let r = 0; r < data.length; r++) {
        if (String(data[r][idIdx]).trim() === requestId) {
          const currentStatus = toStatusKey_(data[r][statusCol]);
          const nextStatus = status;
          if (hasEta && !canEditEtaStatus_(currentStatus, nextStatus)) {
            throw new Error('ETA can only be set when the request has been approved.');
          }
          data[r][statusCol] = nextStatus;
          data[r][approverCol] = approverEmail;
          sheet.getRange(r + 2, statusCol + 1).setValue(nextStatus);
          sheet.getRange(r + 2, approverCol + 1).setValue(approverEmail);
          if (etaCol !== undefined && hasEta) {
            data[r][etaCol] = etaValue;
            sheet.getRange(r + 2, etaCol + 1).setValue(etaValue);
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

function addRequestNote(request) {
  return withErrorHandling_('addRequestNote', request && request.cid, request, () => {
    ensureSetup_();
    const rid = String(request && request.clientRequestId || '').trim();
    if (!rid) {
      throw new Error('clientRequestId is required.');
    }
    const type = normalizeType_(request && request.type);
    if (!supportsRequestNotes_(type)) {
      throw new Error('Notes are supported for IT and maintenance requests only.');
    }
    const requestId = String(request && request.requestId || '').trim();
    if (!requestId) {
      throw new Error('requestId is required.');
    }
    const noteText = sanitizeString_(request && request.note);
    if (!noteText) {
      throw new Error('Note text is required.');
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

    const statusAuth = assertAuthorizedStatusActor_();
    const actorEmail = statusAuth.email;

    const existingRecords = getAllRequestsForType_(type);
    const existing = existingRecords.find(entry => entry.id === requestId);
    if (!existing) {
      throw new Error('Request not found.');
    }

    const notesSheet = getSheet_(SHEETS.REQUEST_NOTES, REQUEST_NOTE_HEADERS);
    const entry = [
      toIsoString_(new Date()),
      type,
      requestId,
      actorEmail,
      noteText
    ];
    withLock_(() => {
      notesSheet.appendRow(entry);
    });

    invalidateRequestCache_(type, 'all');
    invalidateRequestCache_(type, normalizeEmail_(existing.requester));

    const updatedRecords = getAllRequestsForType_(type);
    const updated = updatedRecords.find(entry => entry.id === requestId);
    if (!updated) {
      throw new Error('Request not found.');
    }

    cache.put(ridKey, JSON.stringify(updated), CACHE_TTLS.RID);

    return {
      ok: true,
      request: updated
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

function submitAppFeedback(request) {
  return withErrorHandling_('submitAppFeedback', request && request.cid, request, () => {
    ensureSetup_();
    const feedback = request && request.feedback ? request.feedback : request;
    const typeValue = sanitizeString_(feedback && feedback.type);
    const summary = sanitizeString_(feedback && feedback.summary);
    const details = sanitizeString_(feedback && feedback.details);
    const wish = sanitizeString_(feedback && feedback.wish);
    const providedName = sanitizeString_(feedback && feedback.name);
    const fromEmail = normalizeEmail_(feedback && feedback.fromEmail);
    if (!typeValue) {
      throw new Error('Select what kind of idea you are sharing.');
    }
    if (!summary) {
      throw new Error('Add a short headline so we know where to start.');
    }
    if (!details) {
      throw new Error('Tell us what you experienced.');
    }
    if (!wish) {
      throw new Error('Share what would make it better.');
    }
    const typeKey = typeValue.toLowerCase();
    const typeLabel = typeKey === 'bug' ? 'Bug' : typeKey === 'improvement' ? 'Improvement' : 'Idea';
    const subjectSummary = summary.length > 70 ? `${summary.slice(0, 67)}…` : summary;
    const subject = `[Request Manager] ${typeLabel} feedback – ${subjectSummary}`;
    const friendlyName = providedName || deriveDisplayNameFromEmail_(fromEmail) || 'Request Manager teammate';
    const htmlBody = buildFeedbackEmailBody_({
      type: typeLabel,
      summary,
      details,
      wish,
      name: friendlyName,
      fromEmail,
      submittedAt: formatTimestampForEmail_(new Date())
    });
    MailApp.sendEmail({
      to: PRIMARY_NOTIFICATION_EMAIL,
      subject,
      htmlBody,
      name: EMAIL_SENDER_NAME
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
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(SCRIPT_PROP_SETUP_VERSION) === CURRENT_SETUP_VERSION) {
    return;
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    throw new Error('Initialization is in progress. Please try again in a few seconds.');
  }

  try {
    if (props.getProperty(SCRIPT_PROP_SETUP_VERSION) === CURRENT_SETUP_VERSION) {
      return;
    }

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
        ['SKU-001', 'Copy Paper 8.5x11 (case)', 'Office', '', '', false],
        ['SKU-014', 'Nitrile Gloves (box)', 'Cleaning', '', '', false],
        ['SKU-027', 'Poly Garment Bags (roll)', 'Operations', '', '', false]
      ];
      catalog.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
    }

    props.setProperty(SCRIPT_PROP_SETUP_VERSION, CURRENT_SETUP_VERSION);
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
    const rawKey = String(header || '').trim();
    if (!rawKey) {
      return;
    }
    if (map[rawKey] === undefined) {
      map[rawKey] = idx;
    }
    const lowerKey = rawKey.toLowerCase();
    if (map[lowerKey] === undefined) {
      map[lowerKey] = idx;
    }
    const normalizedKey = lowerKey.replace(/\s+/g, '_');
    if (map[normalizedKey] === undefined) {
      map[normalizedKey] = idx;
    }
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

function toStatusKey_(status) {
  return String(status || '').trim().toLowerCase().replace(/\s+/g, '_');
}

function canEditEtaStatus_(currentStatus, nextStatus) {
  const allowed = ['approved', 'ordered', 'completed'];
  const next = toStatusKey_(nextStatus);
  if (next) {
    return allowed.indexOf(next) !== -1;
  }
  const current = toStatusKey_(currentStatus);
  return allowed.indexOf(current) !== -1;
}

function normalizeType_(type) {
  const value = String(type || '').trim().toLowerCase();
  if (!value || !REQUEST_TYPES[value]) {
    throw new Error('Unsupported request type.');
  }
  return value;
}

function normalizeScope_(scope) {
  const value = String(scope || '').trim().toLowerCase();
  return value === 'mine' ? 'mine' : 'all';
}

function invalidateCatalogCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_KEYS.CATALOG);
  cache.remove(CACHE_KEYS.CATALOG_USAGE);
  cache.remove(CACHE_KEYS.CATALOG_DESC_INDEX);
  runtimeCatalogItems_ = null;
  runtimeCatalogDescriptionIndex_ = null;
}

function invalidateRequestCache_(type, key) {
  const cache = CacheService.getScriptCache();
  const normalizedKey = key === 'all' ? 'all' : normalizeEmail_(key);
  const cacheKey = [CACHE_KEYS.REQUESTS_PREFIX, type, normalizedKey].join(':');
  cache.remove(cacheKey);
}

function getAuthorizedStatusEmails_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEYS.STATUS_EMAILS);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed && Array.isArray(parsed.emails)) {
        return parsed;
      }
    } catch (err) {
      cache.remove(CACHE_KEYS.STATUS_EMAILS);
    }
  }

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(SCRIPT_PROP_STATUS_EMAILS);
  let emails = [];
  let allowlistSource = 'default';

  if (raw) {
    const parsedEmails = parseEmailList_(raw);
    if (parsedEmails.length) {
      emails = parsedEmails.slice();
      allowlistSource = 'script_property';
    }
  }

  if (!emails.length) {
    emails = DEFAULT_STATUS_APPROVER_EMAILS.slice();
  } else {
    DEFAULT_STATUS_APPROVER_EMAILS.forEach(email => {
      if (emails.indexOf(email) === -1) {
        emails.push(email);
      }
    });
  }

  const ownerEmail = normalizeEmail_(Session.getEffectiveUser().getEmail());
  if (ownerEmail && emails.indexOf(ownerEmail) === -1) {
    emails.push(ownerEmail);
  }

  const normalized = emails.map(normalizeEmail_).filter(Boolean);
  const uniqueEmails = Array.from(new Set(normalized));
  const payload = { emails: uniqueEmails, allowlistSource };
  cache.put(CACHE_KEYS.STATUS_EMAILS, JSON.stringify(payload), CACHE_TTLS.STATUS_EMAILS);
  return payload;
}

function parseEmailList_(value) {
  if (!value) {
    return [];
  }
  const trimmed = String(value).trim();
  if (!trimmed) {
    return [];
  }
  let entries = [];
  try {
    const parsed = JSON.parse(trimmed);
    if (Array.isArray(parsed)) {
      entries = parsed.slice();
    } else if (typeof parsed === 'string') {
      entries = [parsed];
    }
  } catch (err) {
    entries = trimmed.split(/[\n,;]+/);
  }
  return entries.map(normalizeEmail_).filter(Boolean);
}

function getStatusAuthContext_() {
  const allowlist = getAuthorizedStatusEmails_();
  const email = getActiveUserEmail_();
  const authorized = Boolean(email) && allowlist.emails.indexOf(email) !== -1;
  let reason = 'authorized';
  if (!email) {
    reason = 'missing_email';
  } else if (!authorized) {
    reason = 'not_listed';
  }
  return {
    email,
    authorized,
    reason,
    allowlistSource: allowlist.allowlistSource,
    allowlistSize: allowlist.emails.length
  };
}

function assertAuthorizedStatusActor_() {
  const context = getStatusAuthContext_();
  if (!context.email) {
    throw new Error('We could not confirm your Google Account email. Sign in with an authorized account or ask an administrator to configure SUPPLIES_TRACKING_STATUS_EMAILS.');
  }
  if (!context.authorized) {
    const account = context.email || 'This account';
    throw new Error(`${account} is not authorized to update requests. Ask an administrator to update the approver allowlist (SUPPLIES_TRACKING_STATUS_EMAILS).`);
  }
  return context;
}

function isAuthorizedStatusActor_(email) {
  const normalized = normalizeEmail_(email);
  if (!normalized) {
    return false;
  }
  const allowlist = getAuthorizedStatusEmails_();
  return allowlist.emails.indexOf(normalized) !== -1;
}

function parsePositiveInteger_(value) {
  const num = Math.floor(Number(value));
  return num > 0 ? num : 0;
}

function parseCurrencyText_(value) {
  const text = sanitizeString_(value);
  if (!text) {
    return null;
  }
  const numericPart = text.replace(/[^0-9.,-]/g, '');
  if (!numericPart) {
    return null;
  }
  const normalized = numericPart.replace(/,/g, '');
  const amount = parseFloat(normalized);
  if (isNaN(amount)) {
    return null;
  }
  const decimalMatch = normalized.match(/\.(\d+)/);
  const decimals = decimalMatch ? decimalMatch[1].length : 0;
  const prefixMatch = text.match(/^[^0-9-]*/);
  const prefix = prefixMatch ? prefixMatch[0] : '';
  const suffixMatch = text.match(/[^0-9.,-]*$/);
  const suffix = suffixMatch ? suffixMatch[0] : '';
  return { amount, decimals, prefix, suffix };
}

function formatAmountWithGrouping_(amount, decimals) {
  const safeDecimals = Math.max(0, Math.min(4, Number(decimals) || 0));
  try {
    return amount.toLocaleString('en-US', {
      minimumFractionDigits: safeDecimals,
      maximumFractionDigits: safeDecimals
    });
  } catch (err) {
    return amount.toFixed(safeDecimals);
  }
}

function formatCurrencyForDisplay_(value, decimalsHint) {
  const text = sanitizeString_(value);
  if (!text) {
    return '';
  }
  const parsed = parseCurrencyText_(text);
  if (!parsed) {
    if (/^\$/i.test(text)) {
      return text;
    }
    if (/[A-Za-z]/.test(text) || /[^0-9.,\s-]/.test(text)) {
      return text;
    }
    return `$${text}`;
  }
  const decimals = parsed.decimals > 0
    ? parsed.decimals
    : (typeof decimalsHint === 'number' && decimalsHint >= 0 ? decimalsHint : 2);
  const formattedAmount = formatAmountWithGrouping_(parsed.amount, decimals);
  const prefix = parsed.prefix || '$';
  const suffix = parsed.suffix || '';
  return `${prefix}${formattedAmount}${suffix}`.trim();
}

function sanitizeString_(value) {
  return String(value || '').trim();
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function deriveDisplayNameFromEmail_(email) {
  const normalized = normalizeEmail_(email);
  if (!normalized) {
    return '';
  }
  const local = normalized.split('@')[0] || '';
  if (!local) {
    return '';
  }
  const parts = local
    .replace(/[._-]+/g, ' ')
    .split(' ')
    .filter(Boolean)
    .map(segment => segment.charAt(0).toUpperCase() + segment.slice(1));
  return parts.join(' ');
}

function normalizeDateOnly_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) {
      throw new Error('Invalid ETA date.');
    }
    const tz = Session.getScriptTimeZone() || 'UTC';
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  if (!text) {
    return '';
  }
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    throw new Error('ETA must use the YYYY-MM-DD format.');
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    throw new Error('Invalid ETA date.');
  }
  const paddedMonth = month.toString().padStart(2, '0');
  const paddedDay = day.toString().padStart(2, '0');
  return `${year}-${paddedMonth}-${paddedDay}`;
}

function formatDateForDisplay_(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) {
      return '';
    }
    const tz = Session.getScriptTimeZone() || 'UTC';
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  if (typeof value === 'number' && !isNaN(value)) {
    const tz = Session.getScriptTimeZone() || 'UTC';
    return Utilities.formatDate(new Date(value), tz, 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  if (!text) {
    return '';
  }
  try {
    return normalizeDateOnly_(text);
  } catch (err) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
      return text;
    }
    const parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      const tz = Session.getScriptTimeZone() || 'UTC';
      return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
    }
    return text;
  }
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

function normalizeDeviceId_(value) {
  const text = sanitizeString_(value);
  if (!text) {
    return '';
  }
  const normalized = text.replace(/[^0-9A-Za-z_-]/g, '').slice(0, 80);
  return normalized;
}

function evaluateDeviceRateLimit_(deviceId, nowMs, props) {
  const key = [DEVICE_RATE_LIMIT.PROP_PREFIX, deviceId].join(':');
  const raw = props.getProperty(key);
  let count = 0;
  let resetAt = nowMs + DEVICE_RATE_LIMIT.WINDOW_MS;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      const parsedResetAt = parsed && parsed.resetAt !== undefined ? Number(parsed.resetAt) : NaN;
      const parsedCount = parsed && parsed.count !== undefined ? Number(parsed.count) : NaN;
      if (!isNaN(parsedResetAt) && parsedResetAt > nowMs) {
        resetAt = parsedResetAt;
        if (!isNaN(parsedCount) && parsedCount >= 0) {
          count = parsedCount;
        }
      }
    } catch (err) {
      count = 0;
      resetAt = nowMs + DEVICE_RATE_LIMIT.WINDOW_MS;
    }
  }
  if (count >= DEVICE_RATE_LIMIT.MAX_REQUESTS) {
    return {
      allowed: false,
      response: {
        ok: false,
        code: 'RATE_LIMIT_EXCEEDED',
        message: 'This device has reached the maximum of 12 requests in 24 hours. Please try again later.'
      }
    };
  }
  return {
    allowed: true,
    key,
    count,
    resetAt
  };
}

function commitDeviceRateLimitUsage_(state, props, nowMs) {
  const baseCount = Number(state && state.count);
  const nextCount = !isNaN(baseCount) && baseCount >= 0 ? baseCount + 1 : 1;
  let resetAt = Number(state && state.resetAt);
  if (isNaN(resetAt) || resetAt <= nowMs) {
    resetAt = nowMs + DEVICE_RATE_LIMIT.WINDOW_MS;
  }
  props.setProperty(state.key, JSON.stringify({
    count: nextCount,
    resetAt
  }));
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

function sendNewRequestNotification_(type, record) {
  if (!record || !record.id) {
    return;
  }
  try {
    const requestTypeLabel = type === 'it' ? 'IT' : 'Maintenance';
    const summary = sanitizeString_(record.summary) || `${requestTypeLabel} request`;
    const subject = `[Request Manager] New ${requestTypeLabel} Request – ${summary}`;
    const htmlBody = buildNewRequestEmailBody_(type, record);
    MailApp.sendEmail({
      to: PRIMARY_NOTIFICATION_EMAIL,
      subject,
      htmlBody,
      name: EMAIL_SENDER_NAME
    });
  } catch (err) {
    logServerError_('sendNewRequestNotification', record && record.id, err, {
      type,
      requestId: record && record.id
    });
  }
}

function buildNewRequestEmailBody_(type, record) {
  const requestTypeLabel = type === 'it' ? 'IT' : 'Maintenance';
  const title = `New ${requestTypeLabel} Request`;
  const submittedAt = formatTimestampForEmail_(record && record.ts);
  const summary = sanitizeString_(record && record.summary);
  const generalRows = [
    ['Request ID', record && record.id],
    ['Submitted', submittedAt],
    ['Requester', record && record.requester],
    ['Current Status', formatStatusForEmail_(record && record.status)]
  ];
  const detailPairs = getRequestFieldPairsForEmail_(type, record && record.fields);
  const detailRowsHtml = detailPairs
    .filter(([label, value]) => sanitizeString_(value))
    .map(([label, value]) =>
      `<tr><th style="text-align:left;padding:6px 12px;background:#f5f7fa;width:160px;">${escapeHtml_(label)}</th>` +
      `<td style="padding:6px 12px;">${escapeHtml_(value)}</td></tr>`
    )
    .join('');
  const generalRowsHtml = generalRows
    .filter(([label, value]) => sanitizeString_(value))
    .map(([label, value]) =>
      `<tr><th style="text-align:left;padding:6px 12px;background:#f5f7fa;width:160px;">${escapeHtml_(label)}</th>` +
      `<td style="padding:6px 12px;">${escapeHtml_(value)}</td></tr>`
    )
    .join('');
  const trackerUrl = getSpreadsheetUrlSafe_();
  const appUrl = sanitizeString_(REQUEST_MANAGER_APP_URL);
  const appLinkHtml = appUrl
    ? `<p style="margin:20px 0 0;">Open this request in the <a style="color:#1d72b8;" href="${escapeHtml_(appUrl)}" target="_blank" rel="noopener">Request Manager app</a>.</p>`
    : '';
  const detailsSection = detailRowsHtml
    ? `<table style="border-collapse:collapse;width:100%;margin-top:12px;border:1px solid #d2d6dc;">${detailRowsHtml}</table>`
    : '<p style="margin:12px 0 0;color:#52606d;">No additional details were provided.</p>';
  const summarySection = summary
    ? `<p style="margin:0 0 16px;font-size:16px;color:#243b53;"><strong>Summary:</strong> ${escapeHtml_(summary)}</p>`
    : '<p style="margin:0 0 16px;font-size:16px;color:#243b53;">A new request has been submitted. Details are below.</p>';
  return [
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#102a43;line-height:1.6;">',
    `<h2 style="margin:0 0 12px;font-size:20px;">${escapeHtml_(title)}</h2>`,
    summarySection,
    `<table style="border-collapse:collapse;width:100%;border:1px solid #d2d6dc;">${generalRowsHtml}</table>`,
    '<h3 style="margin:20px 0 8px;font-size:16px;color:#243b53;">Request Details</h3>',
    detailsSection,
    trackerUrl
      ? `<p style="margin:20px 0 0;">Review the full request list in <a style="color:#1d72b8;" href="${escapeHtml_(trackerUrl)}" target="_blank" rel="noopener">Google Sheets</a>.</p>`
      : '',
    appLinkHtml,
    '</div>'
  ].join('');
}

function buildFeedbackEmailBody_(entry) {
  const typeLabel = sanitizeString_(entry && entry.type) || 'Idea';
  const summary = sanitizeString_(entry && entry.summary);
  const details = sanitizeString_(entry && entry.details);
  const wish = sanitizeString_(entry && entry.wish);
  const submittedAt = sanitizeString_(entry && entry.submittedAt);
  const name = sanitizeString_(entry && entry.name) || 'Request Manager teammate';
  const fromEmail = normalizeEmail_(entry && entry.fromEmail);
  const reachBackHtml = fromEmail
    ? `<p style="margin:8px 0 0;font-size:14px;color:#627d98;">Reply to: <a style="color:#0b57d0;" href="mailto:${escapeHtml_(fromEmail)}">${escapeHtml_(fromEmail)}</a></p>`
    : '';
  const summaryHtml = summary
    ? `<p style="margin:16px 0 0;font-size:16px;color:#243b53;"><strong>${escapeHtml_(summary)}</strong></p>`
    : '';
  const detailsHtml = details
    ? `<p style="margin:16px 0 0;line-height:1.6;color:#334e68;"><strong>What they noticed:</strong><br>${formatMultilineForEmail_(details)}</p>`
    : '';
  const wishHtml = wish
    ? `<p style="margin:16px 0 0;line-height:1.6;color:#334e68;"><strong>What they'd love to see:</strong><br>${formatMultilineForEmail_(wish)}</p>`
    : '';
  return [
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#102a43;line-height:1.6;">',
    `<h2 style="margin:0;font-size:20px;">${escapeHtml_(typeLabel)} insight just arrived</h2>`,
    `<p style="margin:12px 0 0;font-size:15px;color:#334e68;">Shared by <strong>${escapeHtml_(name)}</strong>${submittedAt ? ` on ${escapeHtml_(submittedAt)}` : ''}.</p>`,
    reachBackHtml,
    summaryHtml,
    detailsHtml,
    wishHtml,
    '</div>'
  ].join('');
}

function getRequestFieldPairsForEmail_(type, fields) {
  const safeFields = fields || {};
  if (type === 'it') {
    return [
      ['Issue', safeFields.issue],
      ['Location', safeFields.location],
      ['Device/System', safeFields.device],
      ['Urgency', safeFields.urgency ? capitalize_(safeFields.urgency) : ''],
      ['Additional Details', safeFields.details]
    ];
  }
  if (type === 'maintenance') {
    return [
      ['Issue', safeFields.issue],
      ['Location', safeFields.location],
      ['Urgency', safeFields.urgency ? capitalize_(safeFields.urgency) : ''],
      ['Access Notes', safeFields.accessNotes]
    ];
  }
  return [];
}

function sendSuppliesSummaryEmail_(records) {
  const count = Array.isArray(records) ? records.length : 0;
  const subject = count
    ? `[Request Manager] Weekly Supplies Summary – ${count} awaiting review`
    : '[Request Manager] Weekly Supplies Summary – No pending requests';
  const htmlBody = buildSuppliesSummaryEmailBody_(Array.isArray(records) ? records : []);
  MailApp.sendEmail({
    to: PRIMARY_NOTIFICATION_EMAIL,
    subject,
    htmlBody,
    name: EMAIL_SENDER_NAME
  });
}

function buildSuppliesSummaryEmailBody_(records) {
  const trackerUrl = getSpreadsheetUrlSafe_();
  const count = records.length;
  const introText = count
    ? `${count} supplies ${count === 1 ? 'request requires' : 'requests require'} approval or denial.`
    : 'No supplies requests are waiting for approval or denial this week.';
  let tableHtml = '';
  if (count) {
    const rowsHtml = records
      .slice()
      .sort((a, b) => String(a && a.ts || '').localeCompare(String(b && b.ts || '')))
      .map(record => {
        const fields = record && record.fields ? record.fields : {};
        const submitted = formatTimestampForEmail_(record && record.ts);
        const status = formatStatusForEmail_(record && record.status);
        const description = fields.description || record.summary || 'Supplies request';
        const qty = fields.qty !== undefined && fields.qty !== null && fields.qty !== '' ? String(fields.qty) : '—';
        const location = fields.location || '—';
        const requester = record && record.requester ? record.requester : '—';
        const notes = sanitizeString_(fields.notes);
        const supplier = sanitizeString_(fields.supplier);
        const eta = sanitizeString_(fields.eta);
        const estimatedCost = sanitizeString_(fields.estimatedCost);
        const detailSegments = [];
        if (supplier) {
          detailSegments.push(`<strong>Supplier:</strong> ${escapeHtml_(supplier)}`);
        }
        if (estimatedCost) {
          detailSegments.push(`<strong>Est. Cost:</strong> ${escapeHtml_(estimatedCost)}`);
        }
        if (eta) {
          detailSegments.push(`<strong>ETA:</strong> ${escapeHtml_(formatDateOnlyForEmail_(eta))}`);
        }
        if (notes) {
          detailSegments.push(`<strong>Notes:</strong> ${escapeHtml_(notes)}`);
        }
        const extraDetails = detailSegments.length
          ? `<div style="margin-top:6px;font-size:12px;color:#52606d;">${detailSegments.join('<br>')}</div>`
          : '';
        return [
          '<tr>',
          `<td style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">${escapeHtml_(submitted)}</td>`,
          `<td style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">${escapeHtml_(requester)}</td>`,
          `<td style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">${escapeHtml_(description)}${extraDetails}</td>`,
          `<td style="padding:10px 12px;border-bottom:1px solid #d2d6dc;text-align:center;">${escapeHtml_(qty)}</td>`,
          `<td style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">${escapeHtml_(location)}</td>`,
          `<td style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">${escapeHtml_(status)}</td>`,
          '</tr>'
        ].join('');
      })
      .join('');
    tableHtml = [
      '<table style="border-collapse:collapse;width:100%;margin-top:16px;border:1px solid #d2d6dc;">',
      '<thead>',
      '<tr style="background:#f5f7fa;color:#243b53;text-align:left;">',
      '<th style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">Submitted</th>',
      '<th style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">Requester</th>',
      '<th style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">Item</th>',
      '<th style="padding:10px 12px;border-bottom:1px solid #d2d6dc;text-align:center;width:70px;">Qty</th>',
      '<th style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">Location</th>',
      '<th style="padding:10px 12px;border-bottom:1px solid #d2d6dc;">Status</th>',
      '</tr>',
      '</thead>',
      `<tbody>${rowsHtml}</tbody>`,
      '</table>'
    ].join('');
  } else {
    tableHtml = '<p style="margin:16px 0 0;color:#52606d;">No pending supplies requests were found.</p>';
  }

  return [
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#102a43;line-height:1.6;">',
    '<h2 style="margin:0 0 12px;font-size:20px;">Weekly Supplies Summary</h2>',
    `<p style="margin:0 0 12px;">${escapeHtml_(introText)}</p>`,
    tableHtml,
    trackerUrl
      ? `<p style="margin:20px 0 0;">Review all requests: <a style="color:#1d72b8;" href="${escapeHtml_(trackerUrl)}" target="_blank" rel="noopener">Open Request Tracker</a></p>`
      : '',
    '</div>'
  ].join('');
}

function formatStatusForEmail_(status) {
  const key = toStatusKey_(status);
  if (!key) {
    return 'Pending';
  }
  return key
    .split('_')
    .filter(Boolean)
    .map(part => part.charAt(0).toUpperCase() + part.slice(1))
    .join(' ');
}

function formatTimestampForEmail_(value) {
  if (!value) {
    return '—';
  }
  let date;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    date = value;
  } else {
    date = new Date(String(value));
  }
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return String(value);
  }
  return Utilities.formatDate(date, EMAIL_TIMEZONE, "MMM d, yyyy h:mm a 'ET'");
}

function formatDateOnlyForEmail_(value) {
  if (!value) {
    return '';
  }
  const text = String(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const date = new Date(`${text}T00:00:00`);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, EMAIL_TIMEZONE, 'MMM d, yyyy');
    }
  }
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, EMAIL_TIMEZONE, 'MMM d, yyyy');
  }
  return text;
}

function getSpreadsheetUrlSafe_() {
  try {
    const ss = getSpreadsheet_();
    return ss && typeof ss.getUrl === 'function' ? ss.getUrl() : '';
  } catch (err) {
    return '';
  }
}

function escapeHtml_(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatMultilineForEmail_(value) {
  const text = sanitizeString_(value);
  if (!text) {
    return '';
  }
  return escapeHtml_(text).replace(/\r?\n/g, '<br>');
}

function getFieldNames_(headers) {
  const base = ['id', 'ts', 'requester', 'status', 'approver'];
  return headers.filter(header => base.indexOf(header) === -1);
}

function enrichSuppliesFieldsFromCatalog_(fields) {
  if (!fields) {
    return;
  }
  const descriptionKey = sanitizeString_(fields.description).toLowerCase();
  if (!descriptionKey) {
    return;
  }
  const index = getCatalogDescriptionIndex_();
  const match = index[descriptionKey];
  if (!match) {
    return;
  }
  if (match.supplier && !fields.supplier) {
    fields.supplier = match.supplier;
  }
  if (match.estimatedCost && !fields.estimatedCost) {
    fields.estimatedCost = match.estimatedCost;
  }
  if (match.sku && !fields.catalogSku) {
    fields.catalogSku = match.sku;
  }
  if (match.category && !fields.category) {
    fields.category = match.category;
  }
}

function buildSuppliesEstimatedCostDetail_(fields) {
  const costText = sanitizeString_(fields && fields.estimatedCost);
  if (!costText) {
    return '';
  }
  const displayCost = formatCurrencyForDisplay_(costText) || costText;
  if (fields) {
    fields.estimatedCost = displayCost;
  }
  const qtyValue = fields && fields.qty;
  const qty = Number(qtyValue);
  const parsed = parseCurrencyText_(costText);
  if (!qty || !isFinite(qty) || qty <= 0 || !parsed) {
    return displayCost ? `Estimated cost: ${displayCost}` : '';
  }
  const decimals = parsed.decimals > 0 ? parsed.decimals : 2;
  const totalAmount = parsed.amount * qty;
  const prefix = parsed.prefix || '$';
  const suffix = parsed.suffix || '';
  const formattedTotal = formatAmountWithGrouping_(totalAmount, decimals);
  const totalLabel = `${prefix}${formattedTotal}${suffix}`.trim();
  if (fields) {
    fields.estimatedCostTotal = totalLabel;
  }
  return `Estimated cost: ${totalLabel}`;
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
    fields,
    notes: []
  };
  if (Object.prototype.hasOwnProperty.call(fields, 'eta')) {
    fields.eta = formatDateForDisplay_(fields.eta);
  }
  if (Object.prototype.hasOwnProperty.call(fields, 'urgency')) {
    fields.urgency = normalizeUrgencyValue_(fields.urgency);
  }
  if (type === 'supplies') {
    enrichSuppliesFieldsFromCatalog_(fields);
  }
  record.summary = def.buildSummary(fields);
  record.details = def.buildDetails(fields);
  return record;
}

function normalizeUrgencyValue_(value) {
  const text = sanitizeString_(value).toLowerCase();
  if (!text) {
    return 'normal';
  }
  if (text === 'high') {
    return 'critical';
  }
  if (text === 'medium') {
    return 'normal';
  }
  const allowed = ['low', 'normal', 'critical'];
  return allowed.indexOf(text) === -1 ? 'normal' : text;
}

function capitalize_(value) {
  const text = String(value || '');
  if (!text) {
    return '';
  }
  return text.charAt(0).toUpperCase() + text.slice(1);
}
