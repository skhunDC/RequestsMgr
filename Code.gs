// Code.gs - Centralized Supplies Ordering & Tracking System

const SHEET_ORDERS = 'Orders';
const SHEET_CATALOG = 'Catalog';
const SHEET_BUDGETS = 'Budgets';
const SHEET_AUDIT = 'Audit';

const ORDERS_HEADERS = ['id','ts','requester','item','qty','est_cost','status','approver','decision_ts','override?','justification','cost_center','gl_code'];
const CATALOG_HEADERS = ['sku','desc','category','vendor','price','override_required','threshold','gl_code','cost_center','active'];
const BUDGET_HEADERS = ['cost_center','month','budget','spent_to_date'];
const AUDIT_HEADERS = ['ts','actor','entity','entity_id','action','diff_json'];

function doGet(e){
  return HtmlService.createHtmlOutputFromFile('index');
}

function doPost(e){
  return ContentService.createTextOutput(JSON.stringify({ok:false,error:'POST not used'})).setMimeType(ContentService.MimeType.JSON);
}

function router_(action, payloadJson){
  const email = getActiveUserEmail_();
  try {
    const payload = payloadJson ? JSON.parse(payloadJson) : {};
    switch(action){
      case 'submitOrder': return jsonOk_(submitOrder_(email, payload));
      case 'listMyRequests': return jsonOk_(listMyRequests_(email));
      case 'listApprovals': return jsonOk_(listApprovals_(email));
      case 'bulkDecision': return jsonOk_(bulkDecision_(email, payload));
      default: return jsonErr_('Unknown action: '+action);
    }
  } catch(err){
    return jsonErr_(String(err));
  }
}

function submitOrder_(email, p){
  requireRole_(email, ['requester']);
  const sheet = getOrCreateSheet_(SHEET_ORDERS, ORDERS_HEADERS);
  const map = getHeaderMap_(sheet, ORDERS_HEADERS);
  const lock = LockService.getDocumentLock();
  let locked = false;
  try {
    locked = lock.tryLock(30000);
    if(!locked) throw new Error('Could not obtain lock');
    const order = {
      id: Utilities.getUuid(),
      ts: new Date(),
      requester: email,
      item: p.item || '',
      qty: Number(p.qty) || 0,
      est_cost: Number(p.est_cost) || 0,
      status: 'PENDING',
      approver: '',
      decision_ts: '',
      'override?': !!p.override,
      justification: p.justification || '',
      cost_center: p.cost_center || '',
      gl_code: p.gl_code || ''
    };
    const row = new Array(ORDERS_HEADERS.length).fill('');
    ORDERS_HEADERS.forEach(h => { row[map[h]-1] = order[h]; });
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    appendAudit_(email, SHEET_ORDERS, order.id, 'create', order);
    return order;
  } finally {
    if(locked) lock.releaseLock();
  }
}

function listMyRequests_(email){
  const sheet = getOrCreateSheet_(SHEET_ORDERS, ORDERS_HEADERS);
  const map = getHeaderMap_(sheet, ORDERS_HEADERS);
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return [];
  const values = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();
  const orders = values.map(r => {
    const obj = {};
    ORDERS_HEADERS.forEach(h => { obj[h] = r[map[h]-1]; });
    return obj;
  });
  return orders.filter(o => o.requester === email).sort((a,b)=> new Date(b.ts) - new Date(a.ts));
}

function listApprovals_(email){
  requireRole_(email, ['approver']);
  const sheet = getOrCreateSheet_(SHEET_ORDERS, ORDERS_HEADERS);
  const map = getHeaderMap_(sheet, ORDERS_HEADERS);
  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return [];
  const values = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();
  const orders = values.map(r => {
    const obj = {};
    ORDERS_HEADERS.forEach(h => { obj[h] = r[map[h]-1]; });
    return obj;
  });
  return orders.filter(o => o.status === 'PENDING').sort((a,b)=> new Date(b.ts) - new Date(a.ts));
}

function bulkDecision_(email, p){
  requireRole_(email, ['approver']);
  if(!p || !Array.isArray(p.ids) || !p.decision) throw new Error('Invalid payload');
  const decision = p.decision;
  const sheet = getOrCreateSheet_(SHEET_ORDERS, ORDERS_HEADERS);
  const map = getHeaderMap_(sheet, ORDERS_HEADERS);
  const lock = LockService.getDocumentLock();
  let locked = false;
  const updated = [];
  try {
    locked = lock.tryLock(30000);
    if(!locked) throw new Error('Could not obtain lock');
    const lastRow = sheet.getLastRow();
    if(lastRow < 2) return [];
    const range = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn());
    const values = range.getValues();
    const idIdx = map.id - 1;
    const statusIdx = map.status - 1;
    const approverIdx = map.approver - 1;
    const decisionTsIdx = map.decision_ts - 1;
    const now = new Date();
    p.ids.forEach(id => {
      const rowIndex = values.findIndex(r => r[idIdx] === id);
      if(rowIndex === -1) return;
      const current = values[rowIndex][statusIdx];
      const allowed = current === 'PENDING' ? ['APPROVED','DENIED','ON-HOLD'] : current === 'ON-HOLD' ? ['APPROVED','DENIED'] : [];
      if(allowed.indexOf(decision) === -1) return;
      values[rowIndex][statusIdx] = decision;
      values[rowIndex][approverIdx] = email;
      values[rowIndex][decisionTsIdx] = now;
      const obj = {};
      ORDERS_HEADERS.forEach(h => { obj[h] = values[rowIndex][map[h]-1]; });
      updated.push(obj);
      appendAudit_(email, SHEET_ORDERS, id, 'decision', {status: decision, comment: p.comment || ''});
    });
    range.setValues(values);
    SpreadsheetApp.flush();
    return updated;
  } finally {
    if(locked) lock.releaseLock();
  }
}

function getOrCreateSheet_(name, headers){
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if(!sheet){
    sheet = ss.insertSheet(name);
  }
  if(sheet.getLastRow() === 0){
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sheet;
}

function getHeaderMap_(sheet, headers){
  const existing = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(h => {
    const idx = existing.indexOf(h);
    if(idx === -1) throw new Error('Missing header "'+h+'" in sheet "'+sheet.getName()+'"');
    map[h] = idx+1;
  });
  return map;
}

function appendAudit_(actor, entity, entityId, action, diff){
  const sheet = getOrCreateSheet_(SHEET_AUDIT, AUDIT_HEADERS);
  const map = getHeaderMap_(sheet, AUDIT_HEADERS);
  const row = new Array(AUDIT_HEADERS.length).fill('');
  row[map.ts-1] = new Date();
  row[map.actor-1] = actor;
  row[map.entity-1] = entity;
  row[map.entity_id-1] = entityId;
  row[map.action-1] = action;
  row[map.diff_json-1] = diff ? JSON.stringify(diff) : '';
  sheet.appendRow(row);
}

function requireRole_(email, allowed){
  // Placeholder for future role enforcement
  return true;
}

function jsonOk_(data){ return { ok:true, data:data, error:null }; }
function jsonErr_(msg){ return { ok:false, data:null, error:msg }; }
function getActiveUserEmail_(){ return Session.getActiveUser().getEmail() || 'anonymous@unknown'; }
