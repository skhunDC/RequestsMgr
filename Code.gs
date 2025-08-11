/* Code.gs - supplies request system */

const SHEET_ORDERS = 'Orders';
const SHEET_CATALOG = 'Catalog';
const SHEET_USERS = 'Users';

const DEV_CONSOLE_SEEDS = ['skhun@dublincleaners.com','ss.sku@protonmail.com'];
const ALL_ROLES = ['viewer','requester','approver','developer','super_admin'];
const SS_ID_PROP = 'SS_ID';

const STOCK_LIST = {
  Office: [
    'Copy Paper 8.5×11 (case)',
    'Ballpoint Pens (box)',
    'Sharpie Markers (pack)',
    'Hanging File Folders (box)',
    'Thermal Receipt Paper (case)',
    'Shipping Labels 4×6 (roll)',
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

function getSs(){
  const props = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActive();
  if(!ss){
    const id = props.getProperty(SS_ID_PROP);
    if(id){
      ss = SpreadsheetApp.openById(id);
    } else {
      ss = SpreadsheetApp.create('SuppliesTracking');
      props.setProperty(SS_ID_PROP, ss.getId());
    }
  }
  return ss;
}

function getOrCreateSheet(name){
  const ss = getSs();
  let sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); }
  return sh;
}

function uuid(){ return Utilities.getUuid(); }
function nowIso(){ return new Date().toISOString(); }
function safeLower(s){ return (s || '').toString().trim().toLowerCase(); }

function withLock(fn){
  const lock = LockService.getScriptLock();
  let acquired = false;
  try{
    acquired = lock.tryLock(20000);
    if(!acquired) throw new Error('Another change is in progress. Please try again in a few seconds.');
    return fn();
  } finally {
    if(acquired){
      try{ lock.releaseLock(); }catch(e){ /* ignore */ }
    }
  }
}

function getActiveEmail(){
  return safeLower(Session.getActiveUser().getEmail());
}

function getUserRecord(email){
  email = safeLower(email);
  const sheet = getOrCreateSheet(SHEET_USERS);
  const values = sheet.getDataRange().getValues();
  if(!values.length) return null;
  const header = values.shift();
  const emailIdx = header.indexOf('email');
  const roleIdx = header.indexOf('role');
  const activeIdx = header.indexOf('active');
  const row = values.find(r => safeLower(r[emailIdx]) === email);
  if(!row) return null;
  return {email: safeLower(row[emailIdx]), role: row[roleIdx], active: row[activeIdx] === true || row[activeIdx] === 'TRUE'};
}

function ensureSeedUsers(){
  const sheet = getOrCreateSheet(SHEET_USERS);
  if(sheet.getLastRow() === 0){
    sheet.appendRow(['email','role','active','added_ts','added_by']);
  }
  const existing = sheet.getDataRange().getValues().slice(1).map(r => safeLower(r[0]));
  const seeds = [
    {email:'skhun@dublincleaners.com', role:'super_admin'},
    {email:'ss.sku@protonmail.com', role:'developer'}
  ];
  seeds.forEach(s => {
    if(!existing.includes(s.email)){
      sheet.appendRow([s.email, s.role, true, nowIso(), 'seed']);
    }
  });
}

function requireLoggedIn(){
  const email = getActiveEmail();
  if(!email) throw new Error('Login required');
  return email;
}

function requireRole(allowed){
  const email = requireLoggedIn();
  const rec = getUserRecord(email);
  if(!rec || !rec.active || allowed.indexOf(rec.role) === -1){
    throw new Error('Access denied');
  }
  return rec;
}

function isDevConsoleAllowed(email){
  email = safeLower(email);
  if(DEV_CONSOLE_SEEDS.includes(email)) return true;
  const rec = getUserRecord(email);
  return rec && rec.active && (rec.role === 'developer' || rec.role === 'super_admin');
}

function init(){
  const ss = getSs();
  let sh = ss.getSheetByName(SHEET_ORDERS);
  if(!sh){
    sh = ss.insertSheet(SHEET_ORDERS);
    sh.appendRow(['id','ts','requester','description','qty','est_cost','status','approver','decision_ts','override?','justification']);
  }
  sh = ss.getSheetByName(SHEET_CATALOG);
  if(!sh){
    sh = ss.insertSheet(SHEET_CATALOG);
    sh.appendRow(['sku','description','category','archived']);
  }
  seedCatalogIfEmpty();
}

function seedCatalogIfEmpty(){
  const sh = getOrCreateSheet(SHEET_CATALOG);
  if(sh.getLastRow() > 1) return;
  Object.keys(STOCK_LIST).forEach(cat => {
    STOCK_LIST[cat].forEach(desc => {
      sh.appendRow([uuid(), desc, cat, false]);
    });
  });
}

function getCatalog(req){
  requireRole(ALL_ROLES);
  init();
  const includeArchived = req && req.includeArchived;
  const sheet = getOrCreateSheet(SHEET_CATALOG);
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  return rows.map(r => Object.fromEntries(r.map((v,i)=>[header[i], v]))).filter(r => includeArchived || r.archived !== true);
}

function addCatalogItem(req){
  return withLock(() => {
    requireRole(['developer','super_admin']);
    const {description, category} = req;
    const sh = getOrCreateSheet(SHEET_CATALOG);
    const sku = uuid();
    sh.appendRow([sku, description, category, false]);
    return {sku, description, category, archived:false};
  });
}

function setCatalogArchived(req){
  return withLock(() => {
    requireRole(['developer','super_admin']);
    const {sku, archived} = req;
    const sh = getOrCreateSheet(SHEET_CATALOG);
    const values = sh.getDataRange().getValues();
    const header = values.shift();
    const skuIdx = header.indexOf('sku');
    const archIdx = header.indexOf('archived');
    const row = values.findIndex(r => r[skuIdx] === sku);
    if(row >= 0){
      sh.getRange(row+2, archIdx+1).setValue(archived);
    }
    return 'OK';
  });
}

function submitRequest(payload){
  requireRole(['requester','approver','developer','super_admin']);
  if(!payload || !payload.requester || !payload.items || !payload.items.length){
    throw new Error('Invalid payload: requester and at least one item are required.');
  }
  return withLock(() => {
    init();
    const sh = getOrCreateSheet(SHEET_ORDERS);
    const now = new Date();
    const rows = [];
    const status = 'PENDING';
    const approver = payload.approver || '';
    const estCost = '';
    payload.items.forEach(it => {
      if(!it || !it.desc || !it.qty) return;
      rows.push([uuid(), now, payload.requester, it.desc, Number(it.qty)||0, estCost, status, approver, '', '', '']);
    });
    if(!rows.length) throw new Error('No valid items to submit.');
    sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
    return {ok:true, id:rows[0][0], message:'Request submitted: '+rows.length+' item(s).'};
  });
}

function listMyOrders(req){
  requireRole(ALL_ROLES);
  const email = getActiveEmail();
  const sh = getOrCreateSheet(SHEET_ORDERS);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift();
  return rows.filter(r => r[header.indexOf('requester')] === email).map(r => Object.fromEntries(r.map((v,i)=>[header[i], v])));
}

function listPendingApprovals(){
  requireRole(['approver','developer','super_admin']);
  const sh = getOrCreateSheet(SHEET_ORDERS);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift();
  return rows.filter(r => r[header.indexOf('status')] === 'PENDING').map(r => Object.fromEntries(r.map((v,i)=>[header[i], v])));
}

function decideOrder(req){
  requireRole(['approver','developer','super_admin']);
  return withLock(() => {
    const {id, decision} = req;
    const sh = getOrCreateSheet(SHEET_ORDERS);
    const values = sh.getDataRange().getValues();
    const header = values.shift();
    const idIdx = header.indexOf('id');
    const statusIdx = header.indexOf('status');
    const approverIdx = header.indexOf('approver');
    const row = values.findIndex(r => r[idIdx] === id);
    if(row >= 0){
      const r = row + 2;
      sh.getRange(r, statusIdx+1).setValue(decision);
      sh.getRange(r, approverIdx+1).setValue(getActiveEmail());
      const requester = values[row][header.indexOf('requester')];
      const desc = values[row][header.indexOf('description')];
      GmailApp.sendEmail(requester, 'Supply Request '+decision, '', {htmlBody:`<p>Your request for ${desc} was ${decision}.</p>`});
    }
    return 'OK';
  });
}

function getCsrfToken(email){
  const cache = CacheService.getUserCache();
  let token = cache.get(email+'_csrf');
  if(!token){
    token = uuid();
    cache.put(email+'_csrf', token, 21600);
  }
  return token;
}

function jsonResponse(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  init();
  ensureSeedUsers();
  const email = getActiveEmail();
  const isLoggedIn = !!email;
  let role = null;
  let devConsoleAllowed = false;
  let csrf = '';
  if(isLoggedIn){
    const user = getUserRecord(email);
    role = user ? user.role : null;
    devConsoleAllowed = isDevConsoleAllowed(email);
    csrf = getCsrfToken(email);
  }
  const t = HtmlService.createTemplateFromFile('index');
  t.BOOTSTRAP = {email, role, isLoggedIn, devConsoleAllowed, csrf};
  return t.evaluate().setTitle('Supplies Tracker').addMetaTag('viewport','width=device-width, initial-scale=1');
}

function doPost(e){
  ensureSeedUsers();
  let body;
  try{
    body = JSON.parse(e.postData.contents || '{}');
  }catch(err){
    return jsonResponse({ok:false, error:'Invalid JSON'});
  }
  const email = getActiveEmail();
  if(!email) return jsonResponse({ok:false, error:'Login required'});
  const cache = CacheService.getUserCache();
  const token = cache.get(email+'_csrf');
  if(!token || token !== body.csrf){
    return jsonResponse({ok:false, error:'Invalid CSRF token'});
  }
  try{
    let data;
    const action = body.action;
    const payload = body.payload || {};
    switch(action){
      case 'session.get': {
        requireLoggedIn();
        const rec = getUserRecord(email);
        data = {email, role: rec? rec.role : null, devConsoleAllowed: isDevConsoleAllowed(email)};
        break;
      }
      case 'users.list': {
        requireRole(['developer','super_admin']);
        const sh = getOrCreateSheet(SHEET_USERS);
        const values = sh.getDataRange().getValues();
        const header = values.shift();
        data = values.map(r => Object.fromEntries(r.map((v,i)=>[header[i], v])));
        break;
      }
      case 'users.upsert': {
        requireRole(['developer','super_admin']);
        data = withLock(() => {
          const target = safeLower(payload.email);
          const role = payload.role;
          const active = payload.active;
          const sheet = getOrCreateSheet(SHEET_USERS);
          const rows = sheet.getDataRange().getValues();
          const header = rows.shift();
          const emailIdx = header.indexOf('email');
          const roleIdx = header.indexOf('role');
          const activeIdx = header.indexOf('active');
          const rowNum = rows.findIndex(r => safeLower(r[emailIdx]) === target);
          if(rowNum >= 0){
            const r = rowNum+2;
            sheet.getRange(r, roleIdx+1).setValue(role);
            sheet.getRange(r, activeIdx+1).setValue(active);
          } else {
            sheet.appendRow([target, role, active, nowIso(), email]);
          }
          return 'OK';
        });
        break;
      }
      case 'users.remove': {
        requireRole(['developer','super_admin']);
        data = withLock(() => {
          const target = safeLower(payload.email);
          const sheet = getOrCreateSheet(SHEET_USERS);
          const values2 = sheet.getDataRange().getValues();
          const header2 = values2.shift();
          const emailIdx2 = header2.indexOf('email');
          const activeIdx2 = header2.indexOf('active');
          const rowNum2 = values2.findIndex(r => safeLower(r[emailIdx2]) === target);
          if(rowNum2 >= 0){
            sheet.getRange(rowNum2+2, activeIdx2+1).setValue(false);
          }
          return 'OK';
        });
        break;
      }
      case 'role.me': {
        requireLoggedIn();
        const rrec = getUserRecord(email);
        data = rrec ? rrec.role : null;
        break;
      }
      default:
        return jsonResponse({ok:false, error:'Unknown action'});
    }
    return jsonResponse({ok:true, data});
  }catch(err){
    return jsonResponse({ok:false, error: err.message});
  }
}

function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

