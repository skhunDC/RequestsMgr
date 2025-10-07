/* eslint-env googleappsscript */

const APP_CONFIG = Object.freeze({
  sheetIdPropertyKey: "REQUESTS_APP_SHEET_ID",
  sheetName: "Orders",
  logSheetName: "Logs",
  cachePrefix: "requests-app",
  clientRequestTtlSeconds: 300,
  listCacheTtlSeconds: 180,
  maxPageSize: 50,
  statuses: ["New", "In Progress", "On Hold", "Fulfilled", "Cancelled"],
  brand: {
    title: "Dublin Cleaners Request Manager",
    logoUrl: "https://www.dublincleaners.com/wp-content/uploads/2024/12/Dublin-Logos-stacked.png"
  }
});

const ROLES = Object.freeze({
  DEVELOPER: "developer",
  MANAGER: "manager",
  REQUESTER: "requester"
});

const ROLE_PROPERTY_KEYS = Object.freeze({
  [ROLES.DEVELOPER]: "REQUESTS_APP_DEVELOPERS",
  [ROLES.MANAGER]: "REQUESTS_APP_MANAGERS",
  [ROLES.REQUESTER]: "REQUESTS_APP_REQUESTERS"
});

const DEFAULT_DEVELOPERS = Object.freeze(
  [
    "skhun@dublincleaners.com",
    "ss.sku@protonmail.com",
    "rbown@dublincleaners.com",
    "bbutler@dublincleaners.com",
    "mlackey@dublincleaners.com",
    "rbrown5940@gmail.com",
    "brianmbutler77@gmail.com"
  ].map(normalizeEmail_)
);

const REQUEST_HEADERS = Object.freeze([
  "id",
  "ts",
  "requester",
  "description",
  "qty",
  "status",
  "approver",
  "location",
  "notes"
]);

function doGet(e) {
  const cid = buildCorrelationId_(e && e.parameter && e.parameter.cid);
  const email = getSessionUserEmail_();
  if (!email) {
    return renderAccessDeniedPage_(
      cid,
      "We could not confirm your Google Workspace session. Sign in with your Dublin Cleaners account."
    );
  }
  const roles = resolveUserRoles_(email);
  if (!roles.length) {
    logSecurityEvent_({ reason: "unauthorized-doGet", email, cid });
    return renderAccessDeniedPage_(
      cid,
      "Your account is not authorized for this tool. Contact operations to request access."
    );
  }
  const template = HtmlService.createTemplateFromFile("index");
  template.bootstrap = buildBootstrapPayload_(email, roles, cid);
  const output = template.evaluate();
  output.setTitle(APP_CONFIG.brand.title);
  output.addMetaTag("theme-color", "#0b57d0");
  return output;
}

function getBootstrap(payload) {
  return handleServerCall_("getBootstrap", [ROLES.REQUESTER], payload, (context) => ({
    ok: true,
    bootstrap: buildBootstrapPayload_(context.email, context.roles, context.cid)
  }));
}

function listRequests(payload) {
  return handleServerCall_("listRequests", [ROLES.REQUESTER], payload, (context) => {
    const pageSize = clamp_(Number(payload && payload.pageSize) || 20, 1, APP_CONFIG.maxPageSize);
    const cursor = payload && payload.cursor ? String(payload.cursor) : "";
    const rows = fetchRequestsForUser_(context.email, context.roles, {
      limit: pageSize,
      cursor,
      cid: context.cid
    });
    return {
      ok: true,
      requests: rows.items,
      nextCursor: rows.nextCursor || ""
    };
  });
}

function createRequest(payload) {
  return handleServerCall_("createRequest", [ROLES.REQUESTER], payload, (context) => {
    const validated = validateRequestInput_(payload && payload.request, context.email);
    const clientRequestId = requireClientRequestId_(payload && payload.clientRequestId);
    const processedId = getProcessedRequestId_(clientRequestId);
    if (processedId) {
      const existing = findRequestById_(processedId, context.email, context.roles);
      return { ok: true, duplicate: true, request: existing };
    }
    const record = persistRequest_(validated, context);
    markClientRequestProcessed_(clientRequestId, record.id);
    invalidateRequestCacheForUser_(context.email);
    return { ok: true, request: record };
  });
}

function updateRequestStatus(payload) {
  return handleServerCall_(
    "updateRequestStatus",
    [ROLES.MANAGER, ROLES.DEVELOPER],
    payload,
    (context) => {
      const requestId = String((payload && payload.requestId) || "").trim();
      const status = normalizeStatus_(payload && payload.status);
      if (!requestId) {
        throw createUserFacingError_("INVALID_REQUEST", "A valid request id is required.");
      }
      if (!status) {
        throw createUserFacingError_("INVALID_STATUS", "Choose a status before updating.");
      }
      const sheet = ensureRequestsSheet_();
      const lock = LockService.getScriptLock();
      lock.waitLock(3000);
      try {
        const range = sheet.getDataRange();
        const values = range.getValues();
        const header = values.shift();
        const idIndex = header.indexOf("id");
        const statusIndex = header.indexOf("status");
        const approverIndex = header.indexOf("approver");
        const tsIndex = header.indexOf("ts");
        if (idIndex === -1 || statusIndex === -1) {
          throw createUserFacingError_(
            "CONFIG_ERROR",
            "Request sheet is missing required columns."
          );
        }
        let updated = null;
        const nowIso = toIsoString_(new Date());
        values.some((row, idx) => {
          if (String(row[idIndex]).trim() === requestId) {
            row[statusIndex] = status;
            if (approverIndex > -1) {
              row[approverIndex] = context.email;
            }
            if (tsIndex > -1) {
              row[tsIndex] = nowIso;
            }
            sheet.getRange(idx + 2, 1, 1, header.length).setValues([row]);
            updated = mapRowToRequest_(header, row);
            return true;
          }
          return false;
        });
        if (!updated) {
          throw createUserFacingError_("NOT_FOUND", "Request not found or you do not have access.");
        }
        invalidateRequestCacheForUser_(context.email);
        return { ok: true, request: updated };
      } finally {
        lock.releaseLock();
      }
    }
  );
}

function logClientEvent(payload) {
  return handleServerCall_("logClientEvent", [ROLES.REQUESTER], payload, (context) => {
    const event = payload && payload.event ? sanitizeString_(payload.event, 500) : "";
    if (event) {
      logSecurityEvent_({
        reason: "client-event",
        email: context.email,
        cid: context.cid,
        message: event
      });
    }
    return { ok: true };
  });
}

function handleServerCall_(fnName, requiredRoles, payload, handler) {
  const cid = buildCorrelationId_(payload && payload.cid);
  try {
    const context = authorizeUser_(fnName, requiredRoles, cid);
    context.cid = cid;
    return handler(context);
  } catch (err) {
    const sanitized = sanitizeError_(err);
    Logger.log(
      JSON.stringify({
        ts: new Date().toISOString(),
        fn: fnName,
        cid,
        code: sanitized.code,
        message: sanitized.message,
        stack: sanitized.stack || ""
      })
    );
    return { ok: false, code: sanitized.code, message: sanitized.message };
  }
}

function authorizeUser_(fnName, requiredRoles, cid) {
  const email = getSessionUserEmail_();
  if (!email) {
    logSecurityEvent_({ reason: "no-session", email: "", fnName, cid });
    throw createUserFacingError_(
      "UNAUTHENTICATED",
      "Please sign in with your Dublin Cleaners Google account."
    );
  }
  const roles = resolveUserRoles_(email);
  if (!roles.length) {
    logSecurityEvent_({ reason: "no-roles", email, fnName, cid });
    throw createUserFacingError_("FORBIDDEN", "Your account is not authorized for this tool.");
  }
  if (!hasRole_(roles, requiredRoles)) {
    logSecurityEvent_({ reason: "role-mismatch", email, fnName, cid, requiredRoles });
    throw createUserFacingError_("FORBIDDEN", "You are not allowed to perform that action.");
  }
  return { email, roles };
}

function buildBootstrapPayload_(email, roles, cid) {
  ensureRequestsSheet_();
  const requests = fetchRequestsForUser_(email, roles, { limit: 20, cid });
  return {
    cid,
    session: { email, roles },
    brand: APP_CONFIG.brand,
    statusOptions: APP_CONFIG.statuses,
    requests: requests.items,
    nextCursor: requests.nextCursor || "",
    requestSchema: {
      fields: [
        { id: "description", label: "What do you need?", required: true, maxLength: 280 },
        { id: "qty", label: "Quantity", required: true, type: "number", min: 1, max: 9999 },
        { id: "location", label: "Location", required: false, maxLength: 120 },
        { id: "notes", label: "Notes for the team", required: false, maxLength: 500 }
      ]
    }
  };
}

function fetchRequestsForUser_(email, roles, options) {
  const limit = clamp_(Number(options && options.limit) || 20, 1, APP_CONFIG.maxPageSize);
  const cursor = options && options.cursor ? String(options.cursor) : "";
  const scopeKey = hasRole_(roles, [ROLES.MANAGER, ROLES.DEVELOPER]) ? "all" : `user:${email}`;
  const cacheKey = `${APP_CONFIG.cachePrefix}:list:${scopeKey}`;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      return sliceFromCursor_(parsed, cursor, limit);
    } catch (err) {
      cache.remove(cacheKey);
    }
  }
  const sheet = ensureRequestsSheet_();
  const values = sheet.getDataRange().getValues();
  const header = values.shift();
  const items = values
    .map((row) => mapRowToRequest_(header, row))
    .filter((record) => scopeKey === "all" || record.requester === email);
  items.sort((a, b) => (a.ts < b.ts ? 1 : -1));
  cache.put(cacheKey, JSON.stringify(items), APP_CONFIG.listCacheTtlSeconds);
  return sliceFromCursor_(items, cursor, limit);
}

function sliceFromCursor_(items, cursor, limit) {
  let start = 0;
  if (cursor) {
    const index = items.findIndex((item) => item.id === cursor);
    if (index >= 0) {
      start = index + 1;
    }
  }
  const slice = items.slice(start, start + limit);
  const nextCursor =
    slice.length === limit && start + limit < items.length ? slice[slice.length - 1].id : "";
  return { items: slice, nextCursor };
}

function persistRequest_(request, context) {
  const sheet = ensureRequestsSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(3000);
  try {
    const id = Utilities.getUuid();
    const nowIso = toIsoString_(new Date());
    const row = [
      id,
      nowIso,
      context.email,
      request.description,
      request.qty,
      "New",
      "",
      request.location || "",
      request.notes || ""
    ];
    sheet.appendRow(row);
    return mapRowToRequest_(REQUEST_HEADERS, row);
  } finally {
    lock.releaseLock();
  }
}

function mapRowToRequest_(header, row) {
  const record = {};
  header.forEach((column, index) => {
    record[column] = row[index] || "";
  });
  record.ts = typeof record.ts === "string" ? record.ts : toIsoString_(record.ts);
  record.qty = Number(record.qty || 0);
  record.status = record.status || "New";
  return record;
}

function validateRequestInput_(input, email) {
  if (!input || typeof input !== "object") {
    throw createUserFacingError_("INVALID_REQUEST", "Provide request details before submitting.");
  }
  const description = sanitizeString_(input.description, 280);
  if (!description) {
    throw createUserFacingError_("INVALID_DESCRIPTION", "Describe what you need.");
  }
  const qty = parsePositiveInteger_(input.qty);
  if (!qty) {
    throw createUserFacingError_("INVALID_QTY", "Quantity must be at least 1.");
  }
  const location = sanitizeString_(input.location, 120);
  const notes = sanitizeString_(input.notes, 500);
  return {
    description,
    qty,
    location,
    notes,
    requester: email
  };
}

function requireClientRequestId_(value) {
  const id = String(value || "").trim();
  if (!id) {
    throw createUserFacingError_(
      "MISSING_CLIENT_REQUEST_ID",
      "Client request id missing. Refresh and try again."
    );
  }
  if (id.length > 64) {
    throw createUserFacingError_("INVALID_CLIENT_REQUEST_ID", "Client request id is too long.");
  }
  return id;
}

function getProcessedRequestId_(clientRequestId) {
  const cache = CacheService.getScriptCache();
  return cache.get(`${APP_CONFIG.cachePrefix}:rid:${clientRequestId}`) || "";
}

function markClientRequestProcessed_(clientRequestId, requestId) {
  const cache = CacheService.getScriptCache();
  cache.put(
    `${APP_CONFIG.cachePrefix}:rid:${clientRequestId}`,
    requestId,
    APP_CONFIG.clientRequestTtlSeconds
  );
}

function findRequestById_(requestId, email, roles) {
  if (!requestId) {
    return null;
  }
  const scopeKey = hasRole_(roles, [ROLES.MANAGER, ROLES.DEVELOPER]) ? "all" : `user:${email}`;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(`${APP_CONFIG.cachePrefix}:list:${scopeKey}`);
  if (!cached) {
    return null;
  }
  try {
    const items = JSON.parse(cached);
    return items.find((item) => item.id === requestId) || null;
  } catch (err) {
    return null;
  }
}

function invalidateRequestCacheForUser_(email) {
  const cache = CacheService.getScriptCache();
  cache.remove(`${APP_CONFIG.cachePrefix}:list:all`);
  cache.remove(`${APP_CONFIG.cachePrefix}:list:user:${email}`);
}

function ensureRequestsSheet_() {
  const scriptProps = PropertiesService.getScriptProperties();
  let sheetId = scriptProps.getProperty(APP_CONFIG.sheetIdPropertyKey);
  let spreadsheet = null;
  if (sheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(sheetId);
    } catch (err) {
      spreadsheet = null;
    }
  }
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create("Requests Manager Data");
    scriptProps.setProperty(APP_CONFIG.sheetIdPropertyKey, spreadsheet.getId());
  }
  let sheet = spreadsheet.getSheetByName(APP_CONFIG.sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(APP_CONFIG.sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, REQUEST_HEADERS.length).setValues([REQUEST_HEADERS]);
    return sheet;
  }
  const headerWidth = Math.max(sheet.getLastColumn(), REQUEST_HEADERS.length);
  const header = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  const aligned = REQUEST_HEADERS.every((column, index) => header[index] === column);
  if (!aligned) {
    sheet.clear();
    sheet.getRange(1, 1, 1, REQUEST_HEADERS.length).setValues([REQUEST_HEADERS]);
  }
  return sheet;
}

function ensureLogSheet_() {
  const spreadsheet = ensureRequestsSheet_().getParent();
  let logSheet = spreadsheet.getSheetByName(APP_CONFIG.logSheetName);
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(APP_CONFIG.logSheetName);
    logSheet
      .getRange(1, 1, 1, 7)
      .setValues([["ts", "email", "fn", "reason", "cid", "message", "requiredRoles"]]);
  }
  return logSheet;
}

function resolveUserRoles_(email) {
  const normalized = normalizeEmail_(email);
  if (!normalized) {
    return [];
  }
  const roleSet = {};
  Object.keys(ROLE_PROPERTY_KEYS).forEach((role) => {
    const emails = getEmailsForRole_(role);
    if (emails.indexOf(normalized) !== -1) {
      roleSet[role] = true;
    }
  });
  if (DEFAULT_DEVELOPERS.indexOf(normalized) !== -1) {
    roleSet[ROLES.DEVELOPER] = true;
    roleSet[ROLES.MANAGER] = true;
    roleSet[ROLES.REQUESTER] = true;
  }
  return Object.keys(roleSet);
}

function getEmailsForRole_(role) {
  const key = ROLE_PROPERTY_KEYS[role];
  if (!key) {
    return [];
  }
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) {
    if (role === ROLES.DEVELOPER || role === ROLES.MANAGER || role === ROLES.REQUESTER) {
      return DEFAULT_DEVELOPERS;
    }
    return [];
  }
  return raw.split(",").map(normalizeEmail_).filter(Boolean);
}

function hasRole_(roles, required) {
  if (!required || !required.length) {
    return true;
  }
  return required.some((role) => roles.indexOf(role) !== -1);
}

function normalizeEmail_(email) {
  if (!email) {
    return "";
  }
  return String(email).trim().toLowerCase();
}

function getSessionUserEmail_() {
  const user = Session.getActiveUser();
  return normalizeEmail_(user && typeof user.getEmail === "function" ? user.getEmail() : "");
}

function sanitizeString_(value, maxLength) {
  if (value === null || value === undefined) {
    return "";
  }
  const str = String(value).trim();
  if (!str) {
    return "";
  }
  if (maxLength && str.length > maxLength) {
    return str.substring(0, maxLength);
  }
  return str;
}

function parsePositiveInteger_(value) {
  const num = Number(value);
  if (!isFinite(num) || num < 1) {
    return 0;
  }
  return Math.floor(num);
}

function clamp_(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function buildCorrelationId_(candidate) {
  const value = String(candidate || "").trim();
  if (value && value.length <= 64) {
    return value;
  }
  return Utilities.getUuid();
}

function normalizeStatus_(value) {
  const status = sanitizeString_(value, 40);
  if (!status) {
    return "";
  }
  const match = APP_CONFIG.statuses.find((item) => item.toLowerCase() === status.toLowerCase());
  return match || "";
}

function toIsoString_(value) {
  if (!value) {
    return "";
  }
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return value.toISOString();
  }
  if (typeof value === "string") {
    return value;
  }
  try {
    return new Date(value).toISOString();
  } catch (err) {
    return "";
  }
}

function renderAccessDeniedPage_(cid, message) {
  const html = `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${
    APP_CONFIG.brand.title
  }</title><style>body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;margin:0;background:#f3f4f6;color:#1f2933;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:24px;}main{max-width:420px;background:#fff;border-radius:12px;box-shadow:0 22px 45px rgba(15,23,42,0.12);padding:32px;text-align:center;}img{width:120px;margin-bottom:16px;}h1{margin:0 0 8px;font-size:1.4rem;}p{margin:0;color:#4b5563;}small{display:block;margin-top:18px;color:#9aa5b1;}</style></head><body><main><img src="${
    APP_CONFIG.brand.logoUrl
  }" alt="Dublin Cleaners logo"><h1>Access restricted</h1><p>${message}</p><small>Correlation ID: ${cid}</small></main></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle(APP_CONFIG.brand.title);
}

function sanitizeError_(err) {
  if (!err) {
    return { code: "ERROR", message: "Something went wrong. Try again." };
  }
  if (err.isUserFacing) {
    return { code: err.code || "ERROR", message: err.message || "Something went wrong." };
  }
  return {
    code: err.code || "ERROR",
    message: "Unexpected error. Please try again later.",
    stack: err && err.stack ? String(err.stack) : ""
  };
}

function createUserFacingError_(code, message) {
  const error = new Error(message);
  error.code = code || "ERROR";
  error.isUserFacing = true;
  return error;
}

function logSecurityEvent_(entry) {
  try {
    const sheet = ensureLogSheet_();
    sheet.appendRow([
      toIsoString_(new Date()),
      entry && entry.email ? entry.email : "",
      entry && entry.fnName ? entry.fnName : "",
      entry && entry.reason ? entry.reason : "",
      entry && entry.cid ? entry.cid : "",
      entry && entry.message ? entry.message : "",
      entry && entry.requiredRoles ? JSON.stringify(entry.requiredRoles) : ""
    ]);
  } catch (err) {
    Logger.log("Failed to log security event: " + (err && err.message));
  }
}
