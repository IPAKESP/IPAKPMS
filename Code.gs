/**
 * IPAK PMS 2025 - Google Sheet backend (email + PIN sign-in).
 * Bound to a Google Sheet with tabs: Users, Employees, Forms, Audit (run setup() once to create them).
 *
 * Users tab columns: email | role | person | pin | note
 *   role   : admin / hr  -> see and edit everything, manage employees, import
 *            lm          -> sees employees whose Direct Supervisor equals "person"
 *            hod         -> sees employees whose Head of Department equals "person"
 *   person : the name exactly as written in the Employees tab (sup / hod columns). Only for lm and hod.
 *   pin    : 4 or more characters. Share it with the person privately.
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const EMP_COLS = ['id', 'name', 'desig', 'loc', 'geo', 'dept', 'section', 'grade', 'doj', 'sup', 'hod', 'perf24'];
const MAX_TRIES = 5, LOCK_SECONDS = 600;

/** Run once from the editor. Creates missing tabs and asks for permissions. */
function setup() {
  const tabs = {
    Users: ['email', 'role', 'person', 'pin', 'note'],
    Employees: EMP_COLS,
    Forms: ['id', 'json', 'updatedBy', 'updatedAt'],
    Audit: ['time', 'user', 'action', 'empId']
  };
  Object.keys(tabs).forEach(function (n) {
    const s = SS.getSheetByName(n) || SS.insertSheet(n);
    if (s.getLastRow() === 0) s.appendRow(tabs[n]);
    s.getRange(1, 1, 1, tabs[n].length).setFontWeight('bold');
    s.setFrozenRows(1);
  });
  const u = SS.getSheetByName('Users');
  if (u.getLastRow() < 2) {
    u.appendRow(['', 'admin', '', '', 'Apna email aur PIN likhein. Role: admin / hr / lm / hod']);
    u.appendRow(['', 'hr', '', '', 'HR ka email aur PIN']);
    u.appendRow(['', 'lm', 'Muhammad Abubakar', '', 'Line manager. person = naam bilkul waisa jaisa Employees tab ke sup column mein hai']);
    u.appendRow(['', 'hod', 'Muhammad Kamran Khan', '', 'HOD. person = naam bilkul waisa jaisa Employees tab ke hod column mein hai']);
  }
  const first = SS.getSheets()[0];
  if (first && first.getName() !== 'Users' && first.getLastRow() === 0) { try { SS.deleteSheet(first); } catch (e) {} }
  Logger.log('Setup OK. Now add emails and PINs in the Users tab.');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('IPAK PMS 2025')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- authentication ---------- */
function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase(); }

function me_(cred) {
  const email = norm_(cred && cred.email), pin = String((cred && cred.pin) == null ? '' : cred.pin).trim();
  if (!email || !pin) throw new Error('Enter your email and PIN.');
  const cache = CacheService.getScriptCache(), key = 'fail:' + email;
  if (Number(cache.get(key) || 0) >= MAX_TRIES) throw new Error('Too many wrong attempts. Try again in 10 minutes.');
  const rows = SS.getSheetByName('Users').getDataRange().getValues().slice(1);
  const r = rows.find(function (x) { return norm_(x[0]) === email; });
  const ok = r && String(r[3] == null ? '' : r[3]).trim() !== '' && String(r[3]).trim() === pin;
  if (!ok) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), LOCK_SECONDS);
    throw new Error('Wrong email or PIN.');
  }
  cache.remove(key);
  const role = norm_(r[1]);
  if (['admin', 'hr', 'lm', 'hod'].indexOf(role) < 0) throw new Error('Your role is not set correctly in the Users tab.');
  return { email: email, role: role, person: String(r[2] || '').trim() };
}

/* ---------- helpers ---------- */
function rows_(name) {
  const s = SS.getSheetByName(name);
  if (!s || s.getLastRow() < 2) return [];
  return s.getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== '' && r[0] != null; });
}
function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SS.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v == null ? '' : v).slice(0, 10);
}
function readEmps_() {
  return rows_('Employees').map(function (r) {
    const e = {};
    EMP_COLS.forEach(function (c, i) { e[c] = r[i]; });
    e.id = Number(e.id);
    e.doj = dateStr_(e.doj);
    e.perf24 = (e.perf24 === '' || e.perf24 == null || isNaN(Number(e.perf24))) ? null : Number(e.perf24);
    ['name', 'desig', 'loc', 'geo', 'dept', 'section', 'grade', 'sup', 'hod'].forEach(function (k) { e[k] = String(e[k] == null ? '' : e[k]).trim(); });
    return e;
  });
}
function empRow_(emp) {
  return EMP_COLS.map(function (c) {
    if (c === 'id') return Number(emp.id);
    if (c === 'perf24') return (emp.perf24 === null || emp.perf24 === '' || emp.perf24 == null) ? '' : Number(emp.perf24);
    return String(emp[c] == null ? '' : emp[c]);
  });
}
function canSee_(u, e) {
  if (u.role === 'admin' || u.role === 'hr') return true;
  if (u.role === 'lm') return norm_(e.sup) === norm_(u.person);
  if (u.role === 'hod') return norm_(e.hod) === norm_(u.person);
  return false;
}
function isAdmin_(u) { return u.role === 'admin' || u.role === 'hr'; }
function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const v = sheet.getRange(1, 1, last, 1).getValues();
  for (let i = 1; i < v.length; i++) if (Number(v[i][0]) === Number(id)) return i + 1;
  return 0;
}
function upsert_(name, id, cols) {
  const s = SS.getSheetByName(name), r = findRow_(s, id);
  if (r) s.getRange(r, 1, 1, cols.length).setValues([cols]); else s.appendRow(cols);
}
function log_(u, action, id) { SS.getSheetByName('Audit').appendRow([new Date(), u.email, action, id]); }
function validForm_(f) { return f && Array.isArray(f.objs) && f.objs.length <= 7 && Array.isArray(f.bpi) && f.bpi.length === 6; }

/* ---------- API called from the web page (every call carries the sign-in) ---------- */
function getState(cred) {
  const u = me_(cred);
  const emps = readEmps_().filter(function (e) { return canSee_(u, e); });
  const ids = {};
  emps.forEach(function (e) { ids[e.id] = true; });
  const forms = {};
  rows_('Forms').forEach(function (r) {
    if (ids[Number(r[0])]) { try { forms[Number(r[0])] = JSON.parse(r[1]); } catch (err) {} }
  });
  return { user: u, emps: emps, forms: forms };
}

function saveForm(cred, id, form) {
  const u = me_(cred);
  id = Number(id);
  if (!id || !validForm_(form)) throw new Error('Invalid form data.');
  const json = JSON.stringify(form);
  if (json.length > 45000) throw new Error('This form is too large to save.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const emp = readEmps_().filter(function (e) { return e.id === id; })[0];
    if (!emp || !canSee_(u, emp)) throw new Error('You are not allowed to edit this form.');
    upsert_('Forms', id, [id, json, u.email, new Date()]);
    log_(u, 'saveForm', id);
  } finally { lock.releaseLock(); }
}

function saveEmp(cred, emp) {
  const u = me_(cred);
  if (!isAdmin_(u)) throw new Error('Only Admin or HR can change the employee register.');
  const id = Number(emp && emp.id);
  if (!id || !emp.name) throw new Error('Employee ID and name are required.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { upsert_('Employees', id, empRow_(emp)); log_(u, 'saveEmp', id); } finally { lock.releaseLock(); }
}

function deleteEmp(cred, id) {
  const u = me_(cred);
  if (!isAdmin_(u)) throw new Error('Only Admin or HR can remove employees.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ['Employees', 'Forms'].forEach(function (n) {
      const s = SS.getSheetByName(n), r = findRow_(s, id);
      if (r) s.deleteRow(r);
    });
    log_(u, 'deleteEmp', id);
  } finally { lock.releaseLock(); }
}

/** Replaces ALL employees and forms with the content of a backup file (Admin/HR only). */
function importAll(cred, state) {
  const u = me_(cred);
  if (!isAdmin_(u)) throw new Error('Only Admin or HR can import data.');
  if (!state || !Array.isArray(state.emps) || !state.forms) throw new Error('This is not a PMS backup file.');
  const seen = {};
  state.emps.forEach(function (e) {
    if (!e || !Number(e.id) || !e.name) throw new Error('A row in the file has no employee ID or name.');
    if (seen[e.id]) throw new Error('Employee ID ' + e.id + ' appears twice in the file.');
    seen[e.id] = true;
  });
  const formRows = [];
  state.emps.forEach(function (e) {
    const f = state.forms[e.id];
    if (!f) return;
    if (!validForm_(f)) throw new Error('The form for employee ' + e.id + ' is not valid.');
    const json = JSON.stringify(f);
    if (json.length > 45000) throw new Error('The form for employee ' + e.id + ' is too large.');
    formRows.push([Number(e.id), json, u.email, new Date()]);
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const es = SS.getSheetByName('Employees'), fs = SS.getSheetByName('Forms');
    if (es.getLastRow() > 1) es.getRange(2, 1, es.getLastRow() - 1, EMP_COLS.length).clearContent();
    if (fs.getLastRow() > 1) fs.getRange(2, 1, fs.getLastRow() - 1, 4).clearContent();
    if (state.emps.length) es.getRange(2, 1, state.emps.length, EMP_COLS.length).setValues(state.emps.map(empRow_));
    if (formRows.length) fs.getRange(2, 1, formRows.length, 4).setValues(formRows);
    log_(u, 'importAll', state.emps.length + ' employees');
  } finally { lock.releaseLock(); }
}
