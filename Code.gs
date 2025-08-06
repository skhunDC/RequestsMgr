// Code.gs - Google Apps Script server-side for Centralized Supplies Ordering & Tracking System
// Provides routes, sheet access, auth checks, notifications, and developer console operations.

/**
 * Configuration
 */
const SHEET_ORDERS = 'Orders';
const SHEET_CATALOG = 'Catalog';
const SHEET_AUDIT = 'Audit';

// Initial developer emails; additional addresses stored in ScriptProperties under key DEV_LIST.
const STATIC_DEVS = ['skhun@dublincleaners.com', 'ss.sku@protonmail.com'];
const DEV_PROP_KEY = 'DEV_LIST';

// Allow only Leadership Team emails (domain specific or explicit list)
const LT_DOMAIN = 'dublincleaners.com';

// Google Chat webhook for notifications
const CHAT_WEBHOOK = 'https://chat.googleapis.com/v1/spaces/...'; // replace with real webhook

/** Utility helpers */

/** Returns the active user's email. */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/** Check if user is part of Leadership Team. */
function isLtUser(email) {
  return email && email.toLowerCase().endsWith('@' + LT_DOMAIN);
}

/** Get developer emails including dynamic list from script properties. */
function getDeveloperEmails() {
  const props = PropertiesService.getScriptProperties();
  const dynamic = props.getProperty(DEV_PROP_KEY);
  return STATIC_DEVS.concat(dynamic ? JSON.parse(dynamic) : []);
}

/** Check if user is developer. */
function isDeveloper(email) {
  return getDeveloperEmails().includes(email);
}

/** Ensure active user authorized; throw otherwise. */
function assertAuthorized() {
  const email = getUserEmail();
  if (!isLtUser(email)) {
    throw new Error('Unauthorized');
  }
}

/** Append JSON diff to Audit sheet */
function logAudit(entry) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_AUDIT);
  sheet.appendRow([new Date(), JSON.stringify(entry)]);
}

/** Budget guardrail: warn >80%, block >=100% unless super-admin override. */
function checkBudget(costCenter, newAmount, override) {
  // Placeholder budget logic; assumes budgets stored in named range "BUDGET_" + costCenter
  const ss = SpreadsheetApp.getActive();
  const budget = Number(ss.getRangeByName('BUDGET_' + costCenter).getValue());
  const spent = Number(ss.getRangeByName('SPENT_' + costCenter).getValue());
  const future = spent + newAmount;
  const pct = future / budget;
  if (pct >= 1 && !override) {
    throw new Error('Budget exceeded');
  }
  return pct >= 0.8;
}

/** Gmail notification */
function sendMail(to, subject, html) {
  GmailApp.sendEmail(to, subject, '', {htmlBody: html});
}

/** Google Chat notification */
function sendChat(text) {
  UrlFetchApp.fetch(CHAT_WEBHOOK, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({text: text})
  });
}

/** Fetch catalog items */
function getCatalog() {
  assertAuthorized();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_CATALOG);
  const values = sheet.getDataRange().getValues();
  const header = values.shift();
  return values.map(row => Object.fromEntries(row.map((v,i) => [header[i], v])));
}

/** Create a new order */
function createOrder(order) {
  assertAuthorized();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const email = getUserEmail();
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
    const id = Utilities.getUuid();
    const row = [id, new Date(), email, order.item, order.qty, order.est_cost, 'PENDING', '', '', order.override, order.justification];
    sheet.appendRow(row);

    const warn = checkBudget(order.cost_center, order.est_cost * order.qty, false);
    if (warn) {
      sendMail(email, 'Budget nearing limit', '<p>Budget has reached 80%.</p>');
    }

    sendMail(email, 'Request Submitted', '<p>Your supply request has been submitted.</p>');
    sendChat('New supply request from ' + email);

    logAudit({action: 'createOrder', id: id, order: order});
    return {id: id};
  } finally {
    lock.releaseLock();
  }
}

/** List orders for current user */
function listMyOrders() {
  assertAuthorized();
  const email = getUserEmail();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  const values = sheet.getDataRange().getValues();
  const header = values.shift();
  return values.filter(r => r[2] === email).map(r => Object.fromEntries(r.map((v,i)=>[header[i], v])));
}

/** Approver: list pending orders */
function listPendingOrders() {
  assertAuthorized();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  const values = sheet.getDataRange().getValues();
  const header = values.shift();
  return values.filter(r => r[6] === 'PENDING').map(r => Object.fromEntries(r.map((v,i)=>[header[i], v])));
}

/** Bulk approve or deny orders */
function decideOrders(ids, decision, comment) {
  assertAuthorized();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
    const values = sheet.getDataRange().getValues();
    const header = values.shift();
    const idIndex = header.indexOf('id');
    const statusIndex = header.indexOf('status');
    const approverIndex = header.indexOf('approver');
    const decisionTsIndex = header.indexOf('decision_ts');

    ids.forEach(id => {
      const rowNum = values.findIndex(r => r[idIndex] === id);
      if (rowNum >= 0) {
        const sheetRow = rowNum + 2; // offset for header
        sheet.getRange(sheetRow, statusIndex+1).setValue(decision);
        sheet.getRange(sheetRow, approverIndex+1).setValue(getUserEmail());
        sheet.getRange(sheetRow, decisionTsIndex+1).setValue(new Date());
        sendMail(values[rowNum][2], 'Request ' + decision, '<p>Your request was ' + decision.toLowerCase() + '.</p>');
        logAudit({action: 'decide', id: id, decision: decision, comment: comment});
      }
    });
  } finally {
    lock.releaseLock();
  }
  return 'OK';
}

/** Spend analytics */
function getSpendAnalytics() {
  assertAuthorized();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  const values = sheet.getDataRange().getValues();
  const header = values.shift();
  const estCostIndex = header.indexOf('est_cost');
  const tsIndex = header.indexOf('ts');
  const costCenterIndex = header.indexOf('cost_center');
  const monthly = {};
  const byCenter = {};
  values.forEach(r => {
    const month = Utilities.formatDate(r[tsIndex], Session.getScriptTimeZone(), 'yyyy-MM');
    monthly[month] = (monthly[month] || 0) + r[estCostIndex];
    const cc = r[costCenterIndex];
    byCenter[cc] = (byCenter[cc] || 0) + r[estCostIndex];
  });
  return {monthly, byCenter};
}

/** Developer console: add developer email */
function addDeveloper(email) {
  if (!isDeveloper(getUserEmail())) throw new Error('Forbidden');
  const props = PropertiesService.getScriptProperties();
  const list = getDeveloperEmails().filter(e => !STATIC_DEVS.includes(e));
  if (!list.includes(email)) list.push(email);
  props.setProperty(DEV_PROP_KEY, JSON.stringify(list));
  return list;
}

/** Developer console: remove developer email */
function removeDeveloper(email) {
  if (!isDeveloper(getUserEmail())) throw new Error('Forbidden');
  const props = PropertiesService.getScriptProperties();
  let list = getDeveloperEmails().filter(e => !STATIC_DEVS.includes(e));
  list = list.filter(e => e !== email);
  props.setProperty(DEV_PROP_KEY, JSON.stringify(list));
  return list;
}

/** Daily trigger digest */
function sendDailyDigest() {
  const pending = listPendingOrders();
  if (!pending.length) return;
  const emails = getDeveloperEmails();
  const html = pending.map(p => `<li>${p.item} - ${p.requester}</li>`).join('');
  sendMail(emails.join(','), 'Pending Approvals', `<ul>${html}</ul>`);
}

/** doGet - renders index.html */
function doGet() {
  assertAuthorized();
  const template = HtmlService.createTemplateFromFile('index');
  template.userEmail = getUserEmail();
  return template.evaluate()
    .setTitle('Supplies Ordering')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Enable HTML includes
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

