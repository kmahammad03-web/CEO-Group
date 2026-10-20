/**
 * CEO Group — Operations backend (Google Apps Script)          v4 · Phase 1
 * --------------------------------------------------------------------------
 * Bound to the CEO Operations sheet (Extensions → Apps Script).
 * Deploy → Manage deployments → Edit → New version   (keeps the same Web App URL)
 *
 * Phase 1:
 *  • Students live in a SEPARATE sheet (CEO Students DB). Its link goes in Settings!B8.
 *    syncFromDb() copies it into the "Students" tab every 10 minutes (installed by setup).
 *  • The landing page looks the phone up (action=lookup) → shows the student's name,
 *    the module running now in their group and the doctor.
 *  • Unregistered numbers: the student types a name; only attendance is recorded.
 *
 * Form Responses 1:
 *  A Timestamp | B Phone | C Accreditation | D Module | E Page Opened At | F Submission ID | G Source
 *  H..O formulas (date, day, time, entry, quality, phone key, student name, group) | P Typed Name (unregistered)
 */

const CODE_VERSION = 'v5.2';

const SETTINGS = {
  RESPONSES_SHEET: 'Form Responses 1',
  CATALOG_SHEET: 'Module Catalog',
  TIMEZONE: 'Africa/Cairo',
  RAW_COLS: 7,          // A..G
  FORMULA_COL: 8,       // H
  FORMULA_COUNT: 8,     // H..O
  TYPED_NAME_COL: 16,   // P
  ACC_ORDER: [
    'EIU MBA', 'EIU MBA Healthcare', 'EIU BBA 1 Year', 'EIU DBA',
    'AITU MBA', 'AITU DBA', 'QUALIFI level 7 leadership', 'QUALIFI level 7 Innovation'
  ]
};

const T = {
  STUDENTS: 'Students', TRACKER: 'Tracker', SCHEDULE: 'Group Schedule', DASH: 'Dashboard', SETTINGS: 'Settings',
  DB_TAB: 'Students',                 // tab name inside the Students DB sheet
  DASH_FIRST_ROW: 9,
  FR: "'Form Responses 1'", GS: "'Group Schedule'"
};
// Operations › Students columns
const SC = { KEYMATCH: 1, KEY: 2, PHONE: 3, NAME: 4, ACC: 5, GROUP: 6, STATUS: 7, START_MOD: 8, START_DATE: 9, F_FIRST: 10, F_COUNT: 7, NOTE: 17, SOURCE: 18, DONE: 19, PLAN_NOTE: 20, WIDTH: 20 };
// Students DB columns
const DBC = { PHONE: 1, NAME: 2, ACC: 3, GROUP: 4, STATUS: 5, START_MOD: 6, START_DATE: 7, NOTES: 8, SOURCE: 9, DONE: 10, WIDTH: 10 };
const ACTIVE_STATUS = 'Still studying';
const AUTO_SOURCE = /^(EIU|AITU) · /;   // rows written by importSources_ (manual rows have another / empty source)

/* =========================== Web endpoints =========================== */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'catalog';
  if (action === 'catalog') return json_({ ok: true, catalog: getCatalog_() });
  if (action === 'lookup') return json_(lookup_(p.phone));
  if (action === 'lecture') return json_(lectureInfo_(p.phone));
  if (action === 'ping') return json_({ ok: true, time: new Date().toISOString() });
  if (action === 'version') return json_({ ok: true, version: CODE_VERSION, timezone: SpreadsheetApp.getActive().getSpreadsheetTimeZone() });
  if (action === 'check') return json_(checkSubmission_(String(p.id || '')));
  return json_({ ok: false, error: 'Unknown action' });
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Invalid request' }); }
  if (data.website) return json_({ ok: true });   // honeypot

  const phone = normalizePhone_(data.phone);
  if (!/^\+?\d{8,15}$/.test(phone)) return json_({ ok: false, error: 'Invalid phone number' });

  // Who is this? (server decides – the page cannot fake module/accreditation)
  const found = findStudent_(phone);
  let accreditation = '', module = '', typedName = '', masked = '';
  if (found.length) {
    const wanted = String(data.accreditation || '').trim();
    const en = found.filter(function (x) { return x.acc === wanted; })[0] || found.filter(function (x) { return x.active; })[0] || found[0];
    accreditation = en.acc; module = en.module || ''; masked = en.name;
  } else {
    typedName = String(data.name || '').replace(/\s+/g, ' ').trim().slice(0, 80);
    if (typedName.length < 3) {
      // old page (v1.x) still sends accreditation + module: accept them
      const catalog = getCatalog_(), acc = String(data.accreditation || '').trim(), mod = String(data.module || '').trim();
      if (catalog[acc] && catalog[acc].indexOf(mod) !== -1) { accreditation = acc; module = mod; }
      else return json_({ ok: false, error: 'Name is required' });
    }
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return json_({ ok: false, error: 'Server busy, please try again' }); }
  try {
    const sh = sheet_(SETTINGS.RESPONSES_SHEET);
    const now = new Date();
    const submissionId = String(data.submissionId || Utilities.getUuid()).slice(0, 64);
    const row = firstEmptyRow_(sh);
    if (row > 2) {
      const ids = sh.getRange(2, 6, row - 2, 1).getValues();
      for (let i = ids.length - 1; i >= Math.max(0, ids.length - 300); i--) {
        if (ids[i][0] === submissionId) return json_({ ok: true, timestamp: now.toISOString(), resent: true, known: found.length > 0 });
      }
    }
    const duplicate = isDuplicateToday_(sh, row, phone, module, now);
    sh.getRange(row, 2).setNumberFormat('@');
    sh.getRange(row, 1, 1, SETTINGS.RAW_COLS).setValues([[
      now, phone, accreditation, module, parseDate_(data.pageOpenedAt) || '', submissionId, String(data.source || 'Landing Page')
    ]]);
    sh.getRange(row, SETTINGS.TYPED_NAME_COL).setValue(typedName);
    writeResponseFormulas_(sh, row, 1);
    return json_({ ok: true, timestamp: now.toISOString(), duplicate: duplicate, row: row,
                   known: found.length > 0, name: masked || typedName, module: module, accreditation: accreditation });
  } finally {
    lock.releaseLock();
  }
}

/* =========================== Phone lookup =========================== */
/** Public lookup used by the page. Returns a masked name only (privacy). */
function lookup_(rawPhone) {
  const phone = normalizePhone_(rawPhone);
  if (!/^\+?\d{8,15}$/.test(phone)) return { ok: false, error: 'Invalid phone number' };
  let found = findStudent_(phone);
  if (!found.length) return { ok: true, found: false };
  if (found.some(function (x) { return x.active; })) found = found.filter(function (x) { return x.active; });   // hide old / inactive programmes
  return { ok: true, found: true, name: found[0].name,
           enrollments: found.map(function (x) { return { accreditation: x.acc, group: x.group, module: x.module, doctor: x.doctor, moduleEnds: x.moduleEnds, active: x.active }; }) };
}

/* ---------- WhatsApp follow-up numbers (Settings!A21:C40, created by setup – editable) ---------- */
const WA_HEAD_ROW = 20;
const WA_DEFAULT = [
  ['EIU MBA', '201044015001', 'EIU MBA · Group 1, EIU MBA · Group 2, EIU MBA · Group 3 New, EIU MBA'],
  ['EIU MBA – Golden', '201158882585', 'EIU MBA · Golden'],
  ['EIU MBA Healthcare', '201070988446', 'EIU MBA Healthcare'],
  ['EIU BBA', '201070988446', 'EIU BBA · Top Up, EIU BBA 1 Year, EIU BBA 3 Years'],
  ['AITU MBA', '201070988446', 'AITU MBA'],
  ['DBA (EIU / AITU)', '201158882585', 'DBA (EIU / AITU), EIU DBA, AITU DBA'],
  ['Pre DBA', '201158882585', ''],
  ['UK programmes', '201158882585', ''],
  ['Qualifi', '201158882585', 'QUALIFI level 7 leadership, QUALIFI level 7 Innovation']
];
function readContacts_() {
  const set = sheet_(T.SETTINGS);
  let rows = WA_DEFAULT;
  if (set && String(set.getRange(WA_HEAD_ROW, 1).getValue()).trim() === 'WhatsApp follow-up') {
    const v = set.getRange(WA_HEAD_ROW + 1, 1, 20, 3).getValues().filter(function (r) { return String(r[0]).trim() && String(r[1]).replace(/\D/g, ''); });
    if (v.length) rows = v;
  }
  return rows.map(function (r) {
    return { label: String(r[0]).trim(), whatsapp: String(r[1]).replace(/\D/g, ''),
             keys: String(r[2] || '').split(',').map(function (k) { return k.trim(); }).filter(String) };
  });
}
function contactFor_(contacts, group, acc) {
  const byGroup = contacts.filter(function (c) { return c.keys.indexOf(group) !== -1; })[0];
  const byAcc = contacts.filter(function (c) { return c.keys.indexOf(acc) !== -1; })[0];
  return (byGroup || byAcc || contacts[0] || { whatsapp: '' }).whatsapp;
}
function writeContactsTable_() {
  const set = sheet_(T.SETTINGS);
  if (!set || String(set.getRange(WA_HEAD_ROW, 1).getValue()).trim() === 'WhatsApp follow-up') return;
  set.getRange(WA_HEAD_ROW, 1, 1, 3).setValues([['WhatsApp follow-up', 'WhatsApp number', 'Groups / accreditations it follows (comma separated)']])
     .setFontWeight('bold').setBackground('#0c2536').setFontColor('#ffffff');
  set.getRange(WA_HEAD_ROW + 1, 2, WA_DEFAULT.length, 1).setNumberFormat('@');
  set.getRange(WA_HEAD_ROW + 1, 1, WA_DEFAULT.length, 3).setValues(WA_DEFAULT);
  set.getRange(WA_HEAD_ROW + 1 + WA_DEFAULT.length, 1).setValue('Column A is what an unregistered student chooses on the landing page. Add rows here for new programmes.');
}

/** "My next lecture": for a registered phone → next lectures of each group (date, time, doctor, Zoom link)
 *  + the WhatsApp number that follows that group. Unknown phone → found:false + the list of follow-up numbers. */
function lectureInfo_(rawPhone) {
  const contacts = readContacts_();
  const base = lookup_(rawPhone);
  if (!base.ok) return base;
  if (!base.found) return { ok: true, found: false, contacts: contacts.map(function (c) { return { label: c.label, whatsapp: c.whatsapp }; }) };
  const lw = sheet_(LEC.SHEET);
  const last = lw ? lastDataRow_(lw, 1) : 1;
  const rows = last >= 2 ? lw.getRange(2, 1, last - 1, 16).getValues() : [];
  const tz = SETTINGS.TIMEZONE, dk = function (d) { return Number(Utilities.formatDate(d, tz, 'yyyyMMdd')); };
  const today = dk(new Date());
  const lastZoom = {};                                   // most recent Zoom link per group (fallback)
  rows.forEach(function (r) {
    const z = String(r[8] || '').trim();
    if (r[0] instanceof Date && /^https?:/.test(z) && (!lastZoom[r[3]] || r[0] >= lastZoom[r[3]].d)) lastZoom[r[3]] = { d: r[0], z: z };
  });
  const enrollments = base.enrollments.map(function (e) {
    const list = [];
    rows.forEach(function (r) {
      if (!(r[0] instanceof Date) || String(r[3]) !== e.group) return;
      const status = String(r[10] || '').trim(), moved = r[12] instanceof Date ? r[12] : null;
      let d = r[0], rescheduled = false;
      if (status === 'لم يحضر') { if (!moved) return; d = moved; rescheduled = true; }
      if (dk(d) < today) return;
      const z = String(r[8] || '').trim();
      list.push({ date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), time: String(r[2] || ''), module: String(r[4] || e.module || ''),
                  n: r[5] || '', doctor: String(r[6] || '') || e.doctor || '', zoom: /^https?:/.test(z) ? z : (lastZoom[e.group] ? lastZoom[e.group].z : ''),
                  today: dk(d) === today, rescheduled: rescheduled, _k: dk(d) });
    });
    list.sort(function (a, b) { return a._k - b._k; });
    return { accreditation: e.accreditation, group: e.group, module: e.module, doctor: e.doctor, whatsapp: contactFor_(contacts, e.group, e.accreditation),
             lectures: list.slice(0, 3).map(function (x) { delete x._k; return x; }) };
  });
  return { ok: true, found: true, name: base.name, enrollments: enrollments };
}

/** All enrollments for a phone, with the module running now + doctor. */
function findStudent_(phone) {
  const key = phoneKey_(phone);
  if (!key) return [];
  const st = sheet_(T.STUDENTS);
  const last = lastDataRow_(st, SC.PHONE);
  if (last < 2) return [];
  const rows = st.getRange(2, 1, last - 1, SC.STATUS).getValues().filter(function (r) { return String(r[SC.KEY - 1]) === key || phoneKey_(r[SC.PHONE - 1]) === key; });
  if (!rows.length) return [];
  const today = dayStart_(new Date());
  const sched = readSchedule_();
  const byId = {}; sched.forEach(function (s) { byId[s.id] = s; });
  // running plan rows from Tracker
  const tr = sheet_(T.TRACKER), trLast = lastDataRow_(tr, 1);
  const trVals = trLast >= 2 ? tr.getRange(2, 1, trLast - 1, 8).getValues() : [];
  return rows.map(function (r) {
    const acc = String(r[SC.ACC - 1]), group = String(r[SC.GROUP - 1]), mk = key + '|' + acc;
    let s = null;
    for (let i = 0; i < trVals.length; i++) {
      if (trVals[i][0] !== mk) continue;
      const x = byId[trVals[i][6]];
      if (x && x.start <= today && today <= x.end) { s = x; break; }
    }
    const active = String(r[SC.STATUS - 1]) === ACTIVE_STATUS;
    if (!s && active) s = sched.filter(function (x) { return x.group === group && x.accs.indexOf(acc) !== -1 && x.start <= today && today <= x.end; })[0] || null;
    return { acc: acc, group: group, name: maskName_(r[SC.NAME - 1]), active: String(r[SC.STATUS - 1]) === ACTIVE_STATUS,
             module: s ? s.module : '', doctor: s ? s.doctor : '', moduleEnds: s ? dayKeyToIso_(s.end) : '' };
  });
}

/** "Alghamdi, Faten Sami S" → "Faten A." · "Ahmed Mohsen Hassouna" → "Ahmed M." */
function maskName_(name) {
  let s = String(name || '').replace(/\s+/g, ' ').trim();
  if (s.indexOf(',') > 0) { const p = s.split(','); s = (p[1] + ' ' + p[0]).trim(); }
  const t = s.split(' ').filter(String);
  if (!t.length) return '';
  return t[0].charAt(0).toUpperCase() + t[0].slice(1) + (t[1] ? ' ' + t[1].charAt(0).toUpperCase() + '.' : '');
}

/* =========================== Students DB sync =========================== */
/** Trigger target (every 10 min) + menu. Copies the Students DB into the Students tab. */
function syncFromDb() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return; }
  try { return syncFromDb_(); } finally { lock.releaseLock(); }
}

function syncFromDb_() {
  const set = sheet_(T.SETTINGS);
  const url = String(set.getRange('B8').getValue() || '').trim();
  if (!url) throw new Error('Paste the Students DB link in Settings!B8 first.');
  const dbSheet = SpreadsheetApp.openByUrl(url).getSheetByName(T.DB_TAB);
  if (!dbSheet) throw new Error('Tab "' + T.DB_TAB + '" not found in the Students DB.');

  const dbLast = lastDataRow_(dbSheet, DBC.PHONE);
  const dbVals = dbLast >= 2 ? dbSheet.getRange(2, 1, dbLast - 1, DBC.WIDTH).getValues() : [];
  const st = sheet_(T.STUDENTS);
  const stLast = Math.max(lastDataRow_(st, SC.PHONE), lastDataRow_(st, SC.KEYMATCH));
  const stVals = stLast >= 2 ? st.getRange(2, 1, stLast - 1, SC.WIDTH).getValues() : [];
  const opsRow = {}; stVals.forEach(function (v, i) { if (v[0]) opsRow[v[0]] = i + 2; });

  const seen = {}, appended = [];
  let updated = 0, duplicates = 0;
  dbVals.forEach(function (d) {
    const key = phoneKey_(d[DBC.PHONE - 1]), acc = String(d[DBC.ACC - 1] || '').trim();
    if (!key || !acc) return;
    const mk = key + '|' + acc;
    if (seen[mk]) { duplicates++; return; }
    seen[mk] = true;
    const want = [mk, key, String(d[DBC.PHONE - 1]).replace(/\.0+$/, ''), String(d[DBC.NAME - 1] || '').trim(), acc,
                  String(d[DBC.GROUP - 1] || '').trim(), String(d[DBC.STATUS - 1] || '').trim(),
                  String(d[DBC.START_MOD - 1] || '').trim(), d[DBC.START_DATE - 1] instanceof Date ? d[DBC.START_DATE - 1] : ''];
    const note = String(d[DBC.NOTES - 1] || ''), source = String(d[DBC.SOURCE - 1] || '') || 'Students DB', done = String(d[DBC.DONE - 1] || '');
    const r = opsRow[mk];
    if (r) {
      const cur = stVals[r - 2];
      const same = want.every(function (v, i) { return sameVal_(v, cur[i]); }) && sameVal_(note, cur[SC.NOTE - 1]) && sameVal_(source, cur[SC.SOURCE - 1]) && sameVal_(done, cur[SC.DONE - 1]);
      if (!same) {
        st.getRange(r, SC.KEY, 1, 2).setNumberFormat('@');
        st.getRange(r, 1, 1, 9).setValues([want]);
        st.getRange(r, SC.NOTE, 1, 3).setValues([[note, source, done]]);
        updated++;
      }
    } else {
      appended.push({ want: want, note: note, source: source, done: done });
    }
  });
  // rows that disappeared from the DB
  let removed = 0;
  stVals.forEach(function (v, i) {
    if (v[0] && !seen[v[0]] && v[SC.STATUS - 1] !== 'Removed from DB') { st.getRange(i + 2, SC.STATUS).setValue('Removed from DB'); removed++; }
  });
  // append new students
  if (appended.length) {
    const start = stLast + 1;
    st.getRange(start, SC.KEY, appended.length, 2).setNumberFormat('@');
    st.getRange(start, 1, appended.length, 9).setValues(appended.map(function (a) { return a.want; }));
    st.getRange(start, SC.NOTE, appended.length, 3).setValues(appended.map(function (a) { return [a.note, a.source, a.done]; }));
    // live formulas only for active students (keeps the sheet fast); syncStudents_ adds them if a student becomes active later
    st.getRange(start, SC.F_FIRST, appended.length, SC.F_COUNT).setFormulas(appended.map(function (a, i) {
      return a.want[SC.STATUS - 1] === ACTIVE_STATUS ? studentFormulas_(start + i) : ['', '', '', '', '', '', ''];
    }));
    if (start > 2) st.getRange(2, 1, 1, SC.WIDTH).copyTo(st.getRange(start, 1, appended.length, SC.WIDTH), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  const planned = syncStudents_();
  if (planned || appended.length) refreshDashboard_();
  const summary = 'added ' + appended.length + ' · updated ' + updated + ' · removed ' + removed + ' · plans ' + planned + (duplicates ? ' · duplicates skipped ' + duplicates : '');
  set.getRange('B9').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
  set.getRange('C9').setValue(summary);
  return summary;
}

function sameVal_(a, b) {
  // dates are compared by calendar day (Cairo) – the same date can come back from Sheets a few hours/seconds apart
  if (a instanceof Date || b instanceof Date) {
    const ka = a instanceof Date ? Utilities.formatDate(a, SETTINGS.TIMEZONE, 'yyyyMMdd') : String(a == null ? '' : a).trim();
    const kb = b instanceof Date ? Utilities.formatDate(b, SETTINGS.TIMEZONE, 'yyyyMMdd') : String(b == null ? '' : b).trim();
    return ka === kb;
  }
  return String(a == null ? '' : a).trim() === String(b == null ? '' : b).trim();
}

/** Creates Tracker plans for active students that are not imported and have no plan yet. */
function syncStudents_() {
  const st = sheet_(T.STUDENTS), tr = sheet_(T.TRACKER);
  const last = lastDataRow_(st, SC.PHONE);
  if (last < 2) return 0;
  const vals = st.getRange(2, 1, last - 1, SC.WIDTH).getValues();
  const fJ = st.getRange(2, SC.F_FIRST, last - 1, 1).getFormulas();
  const trLast = lastDataRow_(tr, 1);
  const tracked = {};
  if (trLast >= 2) tr.getRange(2, 1, trLast - 1, 1).getValues().forEach(function (x) { tracked[x[0]] = true; });
  const sched = readSchedule_(), catalog = getCatalog_();
  const newTracker = [];
  let added = 0;
  vals.forEach(function (v, i) {
    const r = i + 2, mk = v[0], key = String(v[SC.KEY - 1]), acc = String(v[SC.ACC - 1]), group = String(v[SC.GROUP - 1]);
    if (!mk || !key || !acc) return;
    if (v[SC.STATUS - 1] !== ACTIVE_STATUS) return;
    if (!fJ[i][0]) st.getRange(r, SC.F_FIRST, 1, SC.F_COUNT).setFormulas([studentFormulas_(r)]);
    if (tracked[mk]) return;
    const startMod = String(v[SC.START_MOD - 1] || '').trim(), startDate = v[SC.START_DATE - 1];
    let plan, note = '';
    if (!sched.some(function (s) { return s.group === group; })) { plan = []; note = 'No schedule for group "' + group + '" – attendance only'; }
    else if (startMod && startDate instanceof Date) {
      const done = String(v[SC.DONE - 1] || '').split(';').map(function (x) { return x.trim(); }).filter(String);
      plan = buildPlanFromHistory_(sched, group, startMod, startDate, done, goLiveDay_());
      if (!plan.length) note = 'No remaining modules (all modules already dated)';
    } else {
      plan = buildPlan_(sched, group, acc, startMod, new Date(), catalog[acc]);
      if (!plan.length) note = 'No current / upcoming module in the schedule';
    }
    if (!plan.length) { if (String(v[SC.PLAN_NOTE - 1]) !== note) st.getRange(r, SC.PLAN_NOTE).setValue(note); return; }
    if (groupsFor_(sched, acc).indexOf(group) === -1) st.getRange(r, SC.PLAN_NOTE).setValue('⚠ Group does not match accreditation');
    else if (v[SC.PLAN_NOTE - 1]) st.getRange(r, SC.PLAN_NOTE).setValue('');
    plan.forEach(function (p, k) { newTracker.push([mk, key, v[SC.NAME - 1], acc, group, k + 1, p.id, p.module]); });
    tracked[mk] = true; added++;
  });
  if (newTracker.length) appendTracker_(newTracker);
  return added;
}

/** Same rule as the initial import: last dated module (if it matches the schedule) + every module not yet taken. */
function buildPlanFromHistory_(sched, group, lastMod, lastDate, doneMods, goLive) {
  const rows = sched.filter(function (s) { return s.group === group; });
  const ld = dayStart_(lastDate), plan = [], seen = {};
  const done = {}; doneMods.forEach(function (m) { done[m] = true; }); done[lastMod] = true;
  const ldMs = new Date(lastDate).getTime();
  const match = rows.filter(function (s) { return s.module === lastMod && Math.abs(dayKeyToDate_(s.start).getTime() - ldMs) <= 21 * 864e5; })[0];
  let curStart = ld;
  if (match) { plan.push(match); seen[match.module] = true; curStart = match.start; }
  rows.forEach(function (s) { if (s.start > curStart && !done[s.module] && !seen[s.module]) { plan.push(s); seen[s.module] = true; } });
  return plan.filter(function (p) { return p.end >= goLive; });
}
function dayKeyToDate_(n) { const s = String(n); return new Date(Number(s.slice(0, 4)), Number(s.slice(4, 6)) - 1, Number(s.slice(6, 8))); }
function goLiveDay_() { const v = sheet_(T.SETTINGS).getRange('B2').getValue(); return v instanceof Date ? dayStart_(v) : dayStart_(new Date()); }

/* =========================== Import from EIU / AITU Data Input sheets =========================== */
const SRC = [
  { key: 'EIU',  url: 'B11', tab: 'MBA Groups',       acc: 'EIU MBA',            status: 'Studying Status', group: { col: 'Study Group' }, modules: true },
  { key: 'EIU',  url: 'B11', tab: 'Healthcare Group', acc: 'EIU MBA Healthcare', status: 'Status',          group: 'EIU MBA Healthcare', modules: true },
  { key: 'EIU',  url: 'B11', tab: 'Top Up Group',     acc: 'EIU BBA 1 Year',     status: 'Studying Status', group: 'EIU BBA · Top Up', modules: true },
  { key: 'EIU',  url: 'B11', tab: 'DBA Group',        acc: 'EIU DBA',            status: 'Studying Status', group: 'DBA (EIU / AITU)', modules: false },
  { key: 'AITU', url: 'B12', tab: 'MBA Groups',       acc: 'AITU MBA',           status: 'Studying Status', group: 'AITU MBA', modules: true },
  { key: 'AITU', url: 'B12', tab: 'DBA Group',        acc: 'AITU DBA',           status: 'Studying Status', group: 'DBA (EIU / AITU)', modules: false }
];
const EIU_GROUPS = { 'Group 1': 'EIU MBA · Group 1', 'Group 2': 'EIU MBA · Group 2', 'Group3 New': 'EIU MBA · Group 3 New', 'Golden Package': 'EIU MBA · Golden' };

function importSources() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(60000); } catch (e) { return; }
  try { return importSources_(); } finally { lock.releaseLock(); }
}

/** Reads every student (any status) from the source sheets and upserts them into the Students DB. */
function importSources_() {
  const set = sheet_(T.SETTINGS), catalog = getCatalog_();
  const books = {};
  const students = [], seen = {};
  SRC.forEach(function (src) {
    const url = String(set.getRange(src.url).getValue() || '').trim();
    if (!url) return;
    if (!books[src.url]) books[src.url] = SpreadsheetApp.openByUrl(url);
    const sh = books[src.url].getSheetByName(src.tab);
    if (!sh) throw new Error('Tab "' + src.tab + '" not found in ' + src.key + ' sheet');
    const vals = sh.getDataRange().getValues();
    const head = vals[0].map(function (h) { return String(h).replace(/\s+/g, ' ').trim(); });
    const pc = head.indexOf('Phone'), nc = head.indexOf('Name'), sc = head.indexOf(src.status);
    if (pc < 0 || nc < 0 || sc < 0) throw new Error('Columns Phone / Name / ' + src.status + ' not found in ' + src.key + ' › ' + src.tab);
    const gc = src.group.col ? head.indexOf(src.group.col) : -1;
    const mcols = {};
    if (src.modules) {
      const used = {};
      head.forEach(function (h, i) { const m = catalogMatch_(catalog[src.acc] || [], h); if (m && !used[m]) { mcols[i] = m; used[m] = true; } });
    }
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const key = phoneKey_(row[pc]), name = String(row[nc] || '').replace(/\s+/g, ' ').trim();
      if (!key || !name || name.charAt(0) === '#') continue;   // skip broken rows (#N/A, #REF! …)
      const mk = key + '|' + src.acc;
      if (seen[mk]) continue;
      seen[mk] = true;
      let group = src.group;
      if (gc >= 0) { const g = String(row[gc] || '').replace(/\s+/g, ' ').trim(); group = EIU_GROUPS[g] || ('EIU MBA · ' + (g || '(no group)')); }
      const dated = {};
      Object.keys(mcols).forEach(function (c) { const d = toDate_(row[c]); if (d) dated[mcols[c]] = d; });
      let lastMod = '', lastDate = '';
      Object.keys(dated).forEach(function (m) { if (!lastDate || dated[m] > lastDate) { lastDate = dated[m]; lastMod = m; } });
      students.push([phoneDigits_(row[pc]), name, src.acc, group, String(row[sc] || '').replace(/\s+/g, ' ').trim(),
                     lastMod, lastDate, src.key + ' · ' + src.tab, Object.keys(dated).sort().join('; ')]);
    }
  });
  if (!students.length) return 'no source links in Settings!B11 / B12';

  // upsert into the Students DB
  const db = SpreadsheetApp.openByUrl(String(set.getRange('B8').getValue()).trim()).getSheetByName(T.DB_TAB);
  const last = lastDataRow_(db, DBC.PHONE);
  const cur = last >= 2 ? db.getRange(2, 1, last - 1, DBC.WIDTH).getValues() : [];
  const at = {}; cur.forEach(function (v, i) { at[phoneKey_(v[0]) + '|' + String(v[DBC.ACC - 1]).trim()] = i; });
  let added = 0, updated = 0, gone = 0;
  const inSource = {};
  students.forEach(function (s) {
    const mk = phoneKey_(s[0]) + '|' + s[2]; inSource[mk] = true;
    const i = at[mk];
    if (i === undefined) {
      cur.push([s[0], s[1], s[2], s[3], s[4], s[5], s[6], '', s[7], s[8]]); added++;
      return;
    }
    const row = cur[i];
    if (!AUTO_SOURCE.test(String(row[DBC.SOURCE - 1] || '')) && String(row[DBC.SOURCE - 1] || '') !== '') return;   // manual row: CS owns it
    const next = [s[0], s[1], s[2], s[3], s[4], s[5], s[6], row[DBC.NOTES - 1], s[7], s[8]];
    const diff = next.findIndex(function (v, k) { return !sameVal_(v, row[k]); });
    if (diff !== -1) {
      if (updated < 3) Logger.log('import update ' + mk + ' · column ' + (diff + 1) + ': "' + row[diff] + '" → "' + next[diff] + '"');
      cur[i] = next; updated++;
    }
  });
  cur.forEach(function (row, i) {
    const mk = phoneKey_(row[0]) + '|' + String(row[DBC.ACC - 1]).trim();
    if (AUTO_SOURCE.test(String(row[DBC.SOURCE - 1] || '')) && !inSource[mk] && row[DBC.STATUS - 1] !== 'Not in source') { cur[i][DBC.STATUS - 1] = 'Not in source'; gone++; }
  });
  if (added || updated || gone) {
    db.getRange(2, DBC.PHONE, cur.length, 1).setNumberFormat('@');
    db.getRange(2, 1, cur.length, DBC.WIDTH).setValues(cur);
  }
  const summary = 'sources ' + students.length + ' · new ' + added + ' · updated ' + updated + ' · not in source ' + gone;
  set.getRange('B13').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
  set.getRange('C13').setValue(summary);
  return summary + ' | Students DB → ' + syncFromDb_();
}

function catalogMatch_(mods, header) {
  const s = String(header || '').replace(/\s+/g, ' ').trim(), low = s.toLowerCase();
  if (!s) return null;
  const code = s.match(/([A-Za-z]{2,5})\s*(\d{3})/);
  if (code) {
    const c = (code[1] + code[2]).toUpperCase();
    for (let i = 0; i < mods.length; i++) if (mods[i].toUpperCase().replace(/\s/g, '').indexOf(c) === 0) return mods[i];
  }
  for (let i = 0; i < mods.length; i++) {
    const ml = mods[i].toLowerCase();
    if (low === ml || low.indexOf(ml) === 0 || low.indexOf(ml) !== -1) return mods[i];
  }
  for (let i = 0; i < mods.length; i++) {
    const name = mods[i].split(' - ').slice(1).join(' - ').toLowerCase();
    if (name && (low.indexOf(name.slice(0, 25)) !== -1 || name.indexOf(low.slice(0, 25)) !== -1)) return mods[i];
  }
  return null;
}
function toDate_(v) {
  let d = null;
  if (v instanceof Date) d = v;
  else if (typeof v === 'string') { const m = v.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])); }
  if (!d || isNaN(d.getTime())) return null;
  const y = d.getFullYear();
  return (y >= 2015 && y <= 2029) ? d : null;
}
function phoneDigits_(v) { const m = String(v == null ? '' : v).replace(/\.0+$/, '').match(/\d{8,}/); return m ? m[0] : ''; }

function installTriggers_(withImport) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['syncFromDb', 'importSources'].indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFromDb').timeBased().everyMinutes(10).create();
  if (withImport) ScriptApp.newTrigger('importSources').timeBased().everyHours(1).create();
}

/* =========================== Setup + menu =========================== */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(SETTINGS.TIMEZONE);
  if (!sheet_(SETTINGS.RESPONSES_SHEET)) throw new Error('Sheet "' + SETTINGS.RESPONSES_SHEET + '" not found');
  sheet_(SETTINGS.RESPONSES_SHEET).getRange('B:B').setNumberFormat('@');
  const set = sheet_(T.SETTINGS);
  const backup = String(set.getRange('B15').getValue() || '').trim();
  if (backup && !String(set.getRange('B16').getValue() || '').trim()) Logger.log(restoreFromBackup_(backup));
  repairResponses_();
  writeCancelledList_();
  writeContactsTable_();
  const url = String(set.getRange('B8').getValue() || '').trim();
  if (url) {
    const hasSources = !!String(set.getRange('B11').getValue() || set.getRange('B12').getValue() || '').trim();
    Logger.log(hasSources ? 'Import from EIU / AITU: ' + importSources_() : 'Students DB sync: ' + syncFromDb_());
    installTriggers_(hasSources);
    Logger.log('Auto-sync every 10 minutes: ON' + (hasSources ? ' · import from EIU / AITU every hour: ON' : ''));
  } else {
    Logger.log('⚠ Settings!B8 is empty – paste the Students DB link and run setup again.');
  }
  refreshDashboard_();
  Logger.log('Setup OK · ' + CODE_VERSION + ' · timezone ' + ss.getSpreadsheetTimeZone());
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('CEO Tools')
    .addItem('Import students from EIU / AITU sheets now', 'menuImport')
    .addItem('Sync students from Students DB now', 'menuSync')
    .addItem('Rebuild dashboard module list', 'menuDash')
    .addItem('Repair response formulas', 'menuRepair')
    .addToUi();
}
function menuImport() { toast_(importSources_()); }
function menuSync() { toast_('Students DB → ' + syncFromDb_()); }
function menuDash() { refreshDashboard_(); toast_('Dashboard rebuilt'); }
function menuRepair() { toast_('Formulas restored on ' + repairResponses_() + ' rows'); }

/* =========================== Assignments + Lectures automation =========================== */
// Assignments: A Account Manager | B Session Date | C Module | D Case | E Dr. Name | F Program | G Acc. | H Link | I Students
//              J Action | K Action time | L CS Action | M CS time | N Dr. Fees | O Fees time | P Note | Q Sent to doctor | R..V formulas
const AS = { SHEET: 'Assignments', AGENT: 1, DATE: 2, MODULE: 3, CASE: 4, DOCTOR: 5, LINK: 8, ACTION: 10, ACTION_AT: 11,
             CS: 12, CS_AT: 13, FEES: 14, FEES_AT: 15, NOTE: 16, SENT: 17, WIDTH: 17 };
// Lectures: K Doctor attended? | L Reason | M New date | O Recorded at
const LEC = { SHEET: 'Lectures', STATUS: 11, REASON: 12, NEWDATE: 13, AT: 15 };

/** Simple trigger: runs on every manual edit. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet(), name = sh.getName();
    const r1 = Math.max(2, e.range.getRow()), r2 = e.range.getRow() + e.range.getNumRows() - 1;
    if (r2 < 2) return;
    const c1 = e.range.getColumn(), c2 = c1 + e.range.getNumColumns() - 1;
    const hit = function (c) { return c1 <= c && c2 >= c; };
    if (name === AS.SHEET) processAssignments_(r1, r2, { action: hit(AS.ACTION), cs: hit(AS.CS), fees: hit(AS.FEES) });
    else if (name === LEC.SHEET && (hit(LEC.STATUS) || hit(LEC.REASON) || hit(LEC.NEWDATE))) stampLectures_(sh, r1, r2);
  } catch (err) { console.error(err); toast_('Error: ' + err.message); }
}

/** Stamps the date & time of every Action / CS Action / Dr. Fees change; fills Session date for new rows. */
function processAssignments_(r1, r2, edited) {
  const sh = sheet_(AS.SHEET);
  if (!sh || r2 < r1) return 0;
  edited = edited || {};
  const vals = sh.getRange(r1, 1, r2 - r1 + 1, AS.WIDTH).getValues();
  const now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  let changed = 0;
  vals.forEach(function (v, i) {
    const r = r1 + i, upd = {};
    const started = [AS.AGENT, AS.MODULE, AS.CASE, AS.DOCTOR, AS.LINK, AS.ACTION].some(function (c) { return v[c - 1] !== '' && v[c - 1] !== null; });
    if (!started) return;
    if (!v[AS.DATE - 1]) upd[AS.DATE] = today;
    const action = String(v[AS.ACTION - 1] || '').trim();
    if (edited.action) {
      upd[AS.ACTION_AT] = action ? now : '';
      if (/^pending$/i.test(action) && !v[AS.SENT - 1]) upd[AS.SENT] = now;
    }
    if (edited.cs) upd[AS.CS_AT] = String(v[AS.CS - 1] || '').trim() ? now : '';
    if (edited.fees) upd[AS.FEES_AT] = String(v[AS.FEES - 1] || '').trim() ? now : '';
    const keys = Object.keys(upd);
    if (!keys.length) return;
    keys.forEach(function (c) { sh.getRange(r, Number(c)).setValue(upd[c]); });
    changed++;
  });
  return changed;
}

function stampLectures_(sh, r1, r2) {
  const v = sh.getRange(r1, 1, r2 - r1 + 1, LEC.AT).getValues(), now = new Date();
  v.forEach(function (row, i) {
    if (!row[0]) return;
    const filled = row[LEC.STATUS - 1] !== '' || row[LEC.REASON - 1] !== '' || row[LEC.NEWDATE - 1] !== '';
    sh.getRange(r1 + i, LEC.AT).setValue(filled ? now : '');
  });
}

/* =========================== One-time restore from a backup copy =========================== */
/** Copies Settings links, check-ins and open assignments from the backup copy (Settings!B15) into this sheet. */
function restoreFromBackup_(url) {
  const bk = SpreadsheetApp.openByUrl(url), set = sheet_(T.SETTINGS), out = [];
  // 1) Settings links
  const bset = bk.getSheetByName(T.SETTINGS);
  if (bset) ['B8', 'B10', 'B11', 'B12'].forEach(function (a) {
    const v = bset.getRange(a).getValue();
    if (v !== '' && v !== null) set.getRange(a).setValue(v);
  });
  // 2) check-ins (A..G raw + P typed name)
  const bfr = bk.getSheetByName(SETTINGS.RESPONSES_SHEET), fr = sheet_(SETTINGS.RESPONSES_SHEET);
  if (bfr) {
    const n = lastDataRow_(bfr, 1) - 1;
    if (n > 0) {
      const raw = bfr.getRange(2, 1, n, SETTINGS.RAW_COLS).getValues();
      const typed = bfr.getRange(2, SETTINGS.TYPED_NAME_COL, n, 1).getValues();
      const cur = lastDataRow_(fr, 1) - 1;
      if (cur > 0) { fr.getRange(2, 1, cur, SETTINGS.RAW_COLS).clearContent(); fr.getRange(2, SETTINGS.TYPED_NAME_COL, cur, 1).clearContent(); }
      fr.getRange(2, 2, n, 1).setNumberFormat('@');
      fr.getRange(2, 1, n, SETTINGS.RAW_COLS).setValues(raw);
      fr.getRange(2, SETTINGS.TYPED_NAME_COL, n, 1).setValues(typed);
      repairResponses_();
    }
    out.push('check-ins ' + Math.max(n, 0));
  }
  // 3) assignments: keep only open ones (Pending / no action yet); Done and canceled are dropped
  const bas = bk.getSheetByName(AS.SHEET), as = sheet_(AS.SHEET);
  if (bas) {
    const head = bas.getRange(1, 1, 1, bas.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    const col = function (name) { return head.indexOf(name); };
    const n = lastDataRow_(bas, 1) - 1;
    const rows = n > 0 ? bas.getRange(2, 1, n, head.length).getValues() : [];
    const isNew = col('CS Action') !== -1;
    const keep = [];
    rows.forEach(function (v) {
      if (!v.slice(0, 9).some(function (x) { return x !== '' && x !== null; })) return;
      const g = function (name) { const i = col(name); return i === -1 ? '' : v[i]; };
      const action = String(g('Action') || '').trim();
      if (/^(done|canceled|cancelled)$/i.test(action)) return;
      const act = /^pending$/i.test(action) ? 'Pending' : '';
      const id = String(g('Request ID') || '').match(/^AS-(\d+)$/);
      if (!act && id && Number(id[1]) <= 2364) return;      // old CS Tasks history that is not Pending
      let note = String(g('Note') || '').trim();
      let fees = String(g('Dr. Fees') || '').trim();
      if (fees && ['Pending', 'Done'].indexOf(fees) === -1) { note = [note, 'Dr. Fees: ' + fees].filter(String).join(' · '); fees = ''; }
      const sent = isNew ? g('Sent to doctor') : (g('Sent date') || (act ? g('Session Date Request') : ''));
      keep.push([g('Account Manager'), g('Session Date Request'), g('Moudule'), g('Case'), g('Dr. Name'), g('Program'), g('Acc.'),
                 g('Zoom link // Assignment link'), g('عدد الطلاب تصحيح'), act,
                 isNew ? g('Action date & time') : (act ? sent : ''), isNew ? g('CS Action') : '', isNew ? g('CS date & time') : '',
                 fees, isNew ? g('Fees date & time') : '', note, act ? sent : '']);
    });
    const cur = Math.max(lastDataRow_(as, 1), lastDataRow_(as, AS.ACTION)) - 1;
    if (cur > 0) as.getRange(2, 1, cur, AS.WIDTH).clearContent();
    if (keep.length) as.getRange(2, 1, keep.length, AS.WIDTH).setValues(keep);
    out.push('open assignments ' + keep.length + ' of ' + rows.length);
  }
  // 4) lecture attendance already recorded in the backup (same date + group + module + lecture #)
  const blec = bk.getSheetByName(LEC.SHEET), lec = sheet_(LEC.SHEET);
  if (blec && lec) {
    const key = function (r) { return Utilities.formatDate(new Date(r[0]), SETTINGS.TIMEZONE, 'yyyyMMdd') + '|' + r[3] + '|' + r[4] + '|' + r[5]; };
    const bn = lastDataRow_(blec, 1) - 1, n = lastDataRow_(lec, 1) - 1;
    if (bn > 0 && n > 0) {
      const map = {};
      blec.getRange(2, 1, bn, LEC.AT).getValues().forEach(function (r) { if (r[0] && (r[LEC.STATUS - 1] || r[LEC.REASON - 1])) map[key(r)] = r; });
      const cur = lec.getRange(2, 1, n, LEC.AT).getValues();
      let k = 0;
      cur.forEach(function (r, i) {
        const b = r[0] && map[key(r)];
        if (b) { lec.getRange(i + 2, LEC.STATUS, 1, 5).setValues([b.slice(LEC.STATUS - 1, LEC.AT)]); k++; }
      });
      out.push('lecture records ' + k);
    }
  }
  const msg = 'Restored ' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'dd/MM/yyyy HH:mm') + ' · ' + out.join(' · ');
  set.getRange('B16').setValue(msg);
  return msg;
}

/** Writes the live list of cancelled lectures into Reports (FILTER is written by the script so Sheets parses it natively). */
function writeCancelledList_() {
  const rp = sheet_('Reports'), row = Number(sheet_(T.SETTINGS).getRange('B17').getValue());
  if (!rp || !row || !sheet_(LEC.SHEET)) return;
  rp.getRange(row, 1).setFormula(
    '=IFERROR(SORT(FILTER({Lectures!A2:A,Lectures!D2:D,Lectures!E2:E,Lectures!F2:F,Lectures!G2:G,Lectures!L2:L,Lectures!M2:M,Lectures!N2:N},' +
    'Lectures!P2:P="Cancelled",Lectures!A2:A>=$B$3,Lectures!A2:A<=$G$5,(Lectures!D2:D=$B$5)+($B$5="All")>0),1,FALSE),"No cancelled lectures in this period")');
  rp.getRange(row, 1, 40, 1).setNumberFormat('dd/mm/yyyy');
  rp.getRange(row, 7, 40, 1).setNumberFormat('dd/mm/yyyy');
}

function toast_(msg) { try { SpreadsheetApp.getActive().toast(msg, 'CEO Tools', 6); } catch (e) {} }

/* Editor tests */
function testLookup() { Logger.log(JSON.stringify(lookup_('966531146380'))); }
function testSubmit() {
  Logger.log(doPost({ postData: { contents: JSON.stringify({ phone: '01000000000', name: 'Test Student', submissionId: 'TEST-' + Date.now(), source: 'Editor test' }) } }).getContent());
}

/* =========================== Formulas =========================== */
function responseFormulas_(r) {
  return [
    '=IF($A' + r + '="","",INT($A' + r + '))',
    '=IF($A' + r + '="","",TEXT($A' + r + ',"dddd"))',
    '=IF($A' + r + '="","",MOD($A' + r + ',1))',
    '=IF($A' + r + '="","",IF(COUNTIFS($M$2:$M' + r + ',$M' + r + ',$D$2:$D' + r + ',$D' + r + ',$H$2:$H' + r + ',$H' + r + ')>1,"Duplicate","First"))',
    '=IF($A' + r + '="","",IF($B' + r + '="","Missing phone",IF($C' + r + '="",IF($P' + r + '<>"","Unregistered","Missing accreditation"),IF($D' + r + '="","No module running",IF($K' + r + '="Duplicate","Duplicate","Complete")))))',
    '=IF($A' + r + '="","",RIGHT(SUBSTITUTE($B' + r + ',"+",""),9))',
    '=IF($A' + r + '="","",IFERROR(INDEX(Students!$D:$D,MATCH($M' + r + '&"|"&$C' + r + ',Students!$A:$A,0)),IF($P' + r + '<>"",$P' + r + '&" (unregistered)","Not in student list")))',
    '=IF($A' + r + '="","",IFERROR(INDEX(Students!$F:$F,MATCH($M' + r + '&"|"&$C' + r + ',Students!$A:$A,0)),""))'
  ];
}
function writeResponseFormulas_(sh, fromRow, n) {
  const f = [];
  for (let i = 0; i < n; i++) f.push(responseFormulas_(fromRow + i));
  sh.getRange(fromRow, SETTINGS.FORMULA_COL, n, SETTINGS.FORMULA_COUNT).setFormulas(f);
  sh.getRange(fromRow, 8, n, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(fromRow, 10, n, 1).setNumberFormat('hh:mm AM/PM');
}
function repairResponses_() {
  const sh = sheet_(SETTINGS.RESPONSES_SHEET);
  const n = Math.max(firstEmptyRow_(sh) - 2 + 200, 999);
  writeResponseFormulas_(sh, 2, n);
  return n;
}
function studentFormulas_(r) {
  const FR = T.FR, look = 'MATCH($A' + r + '&"|1",Tracker!$N:$N,0)';
  return [
    '=IFERROR(INDEX(Tracker!$H:$H,' + look + '),"—")',
    '=IFERROR(INDEX(Tracker!$M:$M,' + look + '),"")',
    '=IFERROR(INDEX(Tracker!$K:$K,' + look + '),"")',
    '=IF(COUNTIFS(' + FR + '!$M:$M,$B' + r + ',' + FR + '!$C:$C,$E' + r + ',' + FR + '!$H:$H,TODAY())>0,"✔ Today","")',
    '=COUNTIFS(' + FR + '!$M:$M,$B' + r + ',' + FR + '!$C:$C,$E' + r + ',' + FR + '!$K:$K,"First")',
    '=IF($N' + r + '=0,"",MAXIFS(' + FR + '!$A:$A,' + FR + '!$M:$M,$B' + r + ',' + FR + '!$C:$C,$E' + r + '))',
    '=COUNTIFS(Tracker!$A:$A,$A' + r + ',Tracker!$M:$M,"Missing")'
  ];
}
function trackerFormulas_(r) {
  const FR = T.FR, GS = T.GS;
  const crit = FR + '!$M:$M,$B' + r + ',' + FR + '!$C:$C,$D' + r + ',' + FR + '!$D:$D,$H' + r;
  const win = FR + '!$H:$H,">="&$I' + r + ',' + FR + '!$H:$H,"<="&$J' + r;
  return [
    '=INDEX(' + GS + '!$E:$E,MATCH($G' + r + ',' + GS + '!$A:$A,0))',
    '=INDEX(' + GS + '!$F:$F,MATCH($G' + r + ',' + GS + '!$A:$A,0))',
    '=COUNTIFS(' + crit + ',' + win + ',' + FR + '!$K:$K,"First")',
    '=IF($K' + r + '=0,"",MAXIFS(' + FR + '!$A:$A,' + crit + ',' + win + '))',
    '=IF($K' + r + '>0,"Attended",IF(IFERROR(INDEX(Students!$G:$G,MATCH($A' + r + ',Students!$A:$A,0)),"")<>"Still studying","Inactive",IF($J' + r + '<Settings!$B$2,"Before tracking",IF(TODAY()<$I' + r + ',"Upcoming",IF(TODAY()<=$J' + r + ',"In progress","Missing")))))',
    '=$A' + r + '&"|"&IF(AND($I' + r + '<=TODAY(),$J' + r + '>=TODAY()),1,0)'
  ];
}
function dashFormulas_(r) {
  const GS = T.GS, m = 'MATCH($B' + r + ',' + GS + '!$A:$A,0)', tc = 'Tracker!$G:$G,$B' + r;
  return [
    '=INDEX(' + GS + '!$B:$B,' + m + ')', '=INDEX(' + GS + '!$D:$D,' + m + ')',
    '=INDEX(' + GS + '!$E:$E,' + m + ')', '=INDEX(' + GS + '!$F:$F,' + m + ')',
    '=IF(TODAY()<$E' + r + ',"Upcoming",IF(TODAY()<=$F' + r + ',"Running","Finished"))',
    '=COUNTIF(Tracker!$G:$G,$B' + r + ')', '=COUNTIFS(' + tc + ',Tracker!$M:$M,"Attended")',
    '=IF($H' + r + '=0,"",$I' + r + '/$H' + r + ')', '=SUMIFS(Tracker!$K:$K,' + tc + ')',
    '=COUNTIFS(' + tc + ',Tracker!$M:$M,"Missing")', '=COUNTIFS(' + tc + ',Tracker!$M:$M,"In progress")'
  ];
}

/* =========================== Tracker + dashboard =========================== */
function appendTracker_(rows) {
  const tr = sheet_(T.TRACKER);
  const start = lastDataRow_(tr, 1) + 1;
  tr.getRange(start, 2, rows.length, 1).setNumberFormat('@');
  tr.getRange(start, 1, rows.length, 8).setValues(rows);
  tr.getRange(start, 3, rows.length, 1).setFormulas(rows.map(function (_, i) {
    return ['=IFERROR(INDEX(Students!$D:$D,MATCH($A' + (start + i) + ',Students!$A:$A,0)),"")'];
  }));
  tr.getRange(start, 9, rows.length, 6).setFormulas(rows.map(function (_, i) { return trackerFormulas_(start + i); }));
  if (start > 2) tr.getRange(2, 1, 1, 14).copyTo(tr.getRange(start, 1, rows.length, 14), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function refreshDashboard_() {
  const db = sheet_(T.DASH), tr = sheet_(T.TRACKER);
  if (!db || !tr) return;
  const trLast = lastDataRow_(tr, 1);
  const used = {};
  if (trLast >= 2) tr.getRange(2, 7, trLast - 1, 1).getValues().forEach(function (x) { if (x[0]) used[x[0]] = true; });
  const ids = readSchedule_().filter(function (s) { return used[s.id]; }).map(function (s) { return s.id; });
  const first = T.DASH_FIRST_ROW;
  const oldLast = Math.max(lastDataRow_(db, 2), first - 1);
  if (ids.length) {
    db.getRange(first, 2, ids.length, 1).setValues(ids.map(function (id) { return [id]; }));
    db.getRange(first, 3, ids.length, 11).setFormulas(ids.map(function (_, i) { return dashFormulas_(first + i); }));
    const newLast = first + ids.length - 1;
    if (newLast > oldLast) db.getRange(first, 2, 1, 12).copyTo(db.getRange(oldLast + 1, 2, newLast - oldLast, 12), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }
  const clearFrom = first + ids.length;
  if (oldLast >= clearFrom) db.getRange(clearFrom, 2, oldLast - clearFrom + 1, 12).clearContent();
}

/* =========================== Shared helpers =========================== */
function readSchedule_() {
  const sh = sheet_(T.SCHEDULE), last = lastDataRow_(sh, 1);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 9).getValues()
    .filter(function (r) { return r[0] && r[1] && r[3] && r[4] instanceof Date; })
    .map(function (r) {
      return { id: String(r[0]), group: String(r[1]), accs: String(r[2]).split('/').map(function (a) { return a.trim(); }),
               module: String(r[3]).trim(), start: dayStart_(r[4]), end: r[5] instanceof Date ? dayStart_(r[5]) : dayStart_(r[4]),
               doctor: String(r[8] || '').trim() };
    })
    .sort(function (a, b) { return a.start - b.start; });
}
function groupsFor_(sched, acc) {
  const out = [];
  sched.forEach(function (s) { if (s.accs.indexOf(acc) !== -1 && out.indexOf(s.group) === -1) out.push(s.group); });
  return out;
}
function buildPlan_(sched, group, acc, startModule, refDate, catalogMods) {
  const rows = sched.filter(function (s) { return s.group === group; });
  if (!rows.length) return [];
  const ref = dayStart_(refDate || new Date());
  let occ = null;
  if (startModule) {
    occ = rows.filter(function (s) { return s.module === startModule && s.start <= ref && ref <= s.end; })[0] ||
          rows.filter(function (s) { return s.module === startModule && s.start >= ref; })[0] || null;
  }
  if (!occ) occ = rows.filter(function (s) { return s.start <= ref && ref <= s.end; })[0] ||
                  rows.filter(function (s) { return s.start > ref; })[0] || null;
  if (!occ) return [];
  const plan = [occ], seen = {}; seen[occ.module] = true;
  const need = (catalogMods && catalogMods.length) ? catalogMods.length : 99;
  rows.forEach(function (s) {
    if (plan.length >= need) return;
    if (s.start > occ.start && !seen[s.module]) { plan.push(s); seen[s.module] = true; }
  });
  return plan;
}
function checkSubmission_(id) {
  if (!id) return { ok: false, error: 'Missing id' };
  const sh = sheet_(SETTINGS.RESPONSES_SHEET);
  const row = firstEmptyRow_(sh);
  if (row <= 2) return { ok: true, found: false };
  const start = Math.max(2, row - 500);
  const vals = sh.getRange(start, 1, row - start, 11).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][5]) === id) {
      return { ok: true, found: true, timestamp: vals[i][0] instanceof Date ? vals[i][0].toISOString() : '', duplicate: vals[i][10] === 'Duplicate' };
    }
  }
  return { ok: true, found: false };
}
function getCatalog_() {
  const sh = sheet_(SETTINGS.CATALOG_SHEET);
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 4).getValues();
  const map = {};
  values.forEach(function (r) {
    const acc = String(r[0]).trim(), mod = String(r[1]).replace(/\s+/g, ' ').trim(), active = String(r[3]).trim().toLowerCase();
    if (!acc || !mod || active === 'no') return;
    if (!map[acc]) map[acc] = [];
    if (map[acc].indexOf(mod) === -1) map[acc].push(mod);
  });
  const ordered = {};
  SETTINGS.ACC_ORDER.forEach(function (a) { if (map[a]) ordered[a] = map[a]; });
  Object.keys(map).forEach(function (a) { if (!ordered[a]) ordered[a] = map[a]; });
  return ordered;
}
function firstEmptyRow_(sh) {
  const last = Math.max(sh.getLastRow(), 2);
  const colA = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) if (colA[i][0] !== '' && colA[i][0] !== null) return i + 3;
  return 2;
}
function isDuplicateToday_(sh, row, phone, module, now) {
  if (row <= 2) return false;
  const start = Math.max(2, row - 3000);
  const vals = sh.getRange(start, 1, row - start, 4).getValues();
  const today = Utilities.formatDate(now, SETTINGS.TIMEZONE, 'yyyy-MM-dd'), key = phoneKey_(phone);
  for (let i = vals.length - 1; i >= 0; i--) {
    const ts = vals[i][0];
    if (!(ts instanceof Date)) continue;
    if (Utilities.formatDate(ts, SETTINGS.TIMEZONE, 'yyyy-MM-dd') !== today) continue;
    if (phoneKey_(vals[i][1]) === key && String(vals[i][3]) === module) return true;
  }
  return false;
}
function phoneKey_(p) {
  const d = String(p == null ? '' : p).replace(/\.0+$/, '').replace(/\D/g, '');
  return d.length >= 8 ? d.slice(-9) : '';
}
function normalizePhone_(p) {
  let s = String(p || '')
    .replace(/[٠-٩]/g, function (d) { return d.charCodeAt(0) - 0x0660; })
    .replace(/[۰-۹]/g, function (d) { return d.charCodeAt(0) - 0x06F0; })
    .replace(/[\s\-().]/g, '');
  if (s.indexOf('00') === 0) s = '+' + s.slice(2);
  return s;
}
function sheet_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function lastDataRow_(sh, col) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const v = sh.getRange(2, col, n - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '' && v[i][0] !== null) return i + 2;
  return 1;
}
/** Calendar day in Cairo as a sortable number (20261019) – immune to the script's own time zone. */
function dayStart_(d) { return Number(Utilities.formatDate(new Date(d), SETTINGS.TIMEZONE, 'yyyyMMdd')); }
function dayKeyToIso_(n) { const s = String(n); return s.slice(0, 4) + '-' + s.slice(4, 6) + '-' + s.slice(6, 8); }
function parseDate_(iso) { if (!iso) return null; const d = new Date(iso); return isNaN(d.getTime()) ? null : d; }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
