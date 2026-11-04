/*
 * Audience Coverage - trial build.
 *
 * Everything in this file runs inside the visitor's browser tab. A file the
 * office loads is read into memory, passed through the import gate (which
 * removes names, dates of birth, addresses and contact details), counted into
 * group totals, and shown. Nothing from the file is written to browser storage
 * or sent anywhere. The only thing saved is the "office profile": rules about
 * what tag values mean, which contain no personal information.
 */
'use strict';

/* ------------------------------------------------------------------ constants */
const PROFILE_KEY = 'audienceCoverage.officeProfile.v1';
const MIN_CELL = 10; // groups smaller than this are never shown
const AGE_BRACKETS = ['18-24', '25-34', '35-44', '45-54', '55-64', '65-74', '75+'];
const PERIODS = [
  { days: 7, short: '1 week', long: 'Last week' },
  { days: 30, short: '1 month', long: 'Last month' },
  { days: 91, short: '3 months', long: 'Last 3 months' },
  { days: 182, short: '6 months', long: 'Last 6 months' },
  { days: 365, short: '1 year', long: 'Last year' },
];
const DIMS = [
  { key: 'suburb', label: 'Suburb' }, { key: 'postcode', label: 'Postcode' },
  { key: 'age', label: 'Age' }, { key: 'gender', label: 'Gender' },
  { key: 'community', label: 'Community' }, { key: 'religion', label: 'Religion' },
  { key: 'housing', label: 'Housing' }, { key: 'issues', label: 'Issues' },
  { key: 'newElector', label: 'New elector' }, { key: 'reach', label: 'Reachability' },
];
const DIM_LABEL = Object.fromEntries(DIMS.map(d => [d.key, d.label]));
const CHANNEL_COLOURS = { Email: '#2E6FE0', Phone: '#1F7549', SMS: '#2A9560', Letter: '#EC9308', 'In person': '#7E5BD0', Doorknock: '#C24E15', Other: '#94A3B8' };
const UNTAGGED = 'Untagged';

/* Column recognition. Headers are compared lower-case with everything except
   letters and digits removed, so "Date of Birth", "DOB" and "date_of_birth" match. */
const COL = {
  name: ['firstname', 'lastname', 'surname', 'givenname', 'givennames', 'othernames', 'knownas', 'knownfirstfriend', 'name', 'fullname', 'preferredname', 'middlename', 'title', 'salutation'],
  dob: ['dob', 'dateofbirth', 'birthdate', 'birthday'],
  address: ['unit', 'number', 'street', 'streetname', 'streetnumber', 'address', 'address1', 'address2', 'address3', 'habitat', 'postal1', 'postal2', 'postal3', 'postaladdress', 'postalcode', 'residentialaddress', 'lot', 'housenumber'],
  email: ['email', 'emailaddress', 'email1', 'email2', 'hasemail'],
  mobile: ['mobile', 'mobilephone', 'mobileno', 'mobilenumber', 'cell', 'hasmobile'],
  phone: ['phone', 'homephone', 'phonenumber', 'landline', 'workphone', 'hasphone', 'telephone'],
  enrol: ['enroldate', 'enrolmentdate', 'enrolleddate', 'enrolldate'],
  id: ['id', 'entityid', 'ccid', 'electorid', 'personid', 'recordid'],
  gender: ['gender', 'sex'],
  community: ['ethnicity', 'community', 'culturalbackground'],
  religion: ['religion'],
  housing: ['housing', 'tenure', 'housingtenure'],
  suburb: ['suburb', 'subtown', 'locality', 'town', 'city'],
  postcode: ['postcode', 'pcode'],
  age: ['agebracket', 'agegroup', 'agerange', 'age', 'agecode'],
  sa1: ['sa1', 'sa1code'],
  issues: ['issues', 'interests', 'issue', 'interest'],
  ccTopics: ['activitytypes', 'interactionalltopics', 'interactionalltopicsmulti', 'tags'],
  knownDrop: ['chknumuq', 'chknumsh', 'datasource', 'volunteer', 'elecstate', 'eleclocal', 'elecfed'],
};
const COL_LOOKUP = {};
for (const [k, list] of Object.entries(COL)) for (const h of list) COL_LOOKUP[h] = k;

/* The office profile. Rules are tried in order against each tag value; the
   office's own answers (profile.values) take priority over the rules. */
const DEFAULT_RULES = [
  { re: '^ethnic\\s*[/:-]\\s*(.+)$', cls: 'community', cap: 1, note: 'Ethnic/… tags set community' },
  { re: '^religion\\s*[/:-]\\s*(.+)$', cls: 'religion', cap: 1, note: 'Religion/… tags set religion' },
  { re: '^housing\\s*[/:-]\\s*(.+)$', cls: 'housing', cap: 1, note: 'Housing/… tags set housing' },
  { re: '^contact details (added|changed|updated)', cls: 'admin', note: 'Contact detail changes are admin' },
  { re: 'office administration', cls: 'admin', note: 'Office administration is admin' },
  { re: 'record adjustment', cls: 'admin', note: 'Record adjustments are admin' },
  { re: 'persuasion|voting intention|target(ing)? ?(list|tag)?\\b|swing|^vi\\s*[:\\d]', cls: 'ignore', note: 'Persuasion and targeting tags are never used' },
  { re: '^pfn\\s*:', cls: 'contact', note: 'PFN: topics are real contact' },
];
const CHANNEL_CODES = {
  'INCOMING LETTER': ['Letter', 'in'], 'OUTGOING LETTER': ['Letter', 'out'],
  IWU: ['In person', 'in'], IE: ['Email', 'in'], OE: ['Email', 'out'], IC: ['Phone', 'in'], OC: ['Phone', 'out'],
  VM: ['Phone', 'in'], IS: ['SMS', 'in'], OS: ['SMS', 'out'], IL: ['Letter', 'in'], OL: ['Letter', 'out'],
  DK: ['Doorknock', 'out'], IM: ['In person', 'in'], OM: ['In person', 'out'],
};
const CODE_RE = new RegExp('\\b(' + Object.keys(CHANNEL_CODES).map(c => c.replace(' ', '\\s+')).join('|') + ')\\b');
const CLASSES = [
  ['contact', 'Counts as contact'], ['admin', 'Admin (not contact)'], ['community', 'Community'],
  ['religion', 'Religion'], ['housing', 'Housing'], ['issue', 'Issue or interest'], ['ignore', 'Ignore'],
];

/* ------------------------------------------------------------------ state */
const state = {
  data: null,          // processed import (in memory only)
  period: 2,           // index into PERIODS
  dim: 'suburb',
  sel: null,           // selected group name
  hmRow: 'suburb', hmCol: 'age',
  profile: loadProfile(),
};

/* ------------------------------------------------------------------ helpers */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmt = n => Math.round(n).toLocaleString('en-AU');
const pct = (a, b) => b ? Math.round(a / b * 100) : 0;
const norm = h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, '');
const DAY = 86400000;
const dayNum = d => Math.floor(d.getTime() / DAY);
const fromDay = n => new Date(n * DAY);
const fmtDate = d => d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });

function loadProfile() {
  try {
    const p = JSON.parse(localStorage.getItem(PROFILE_KEY) || 'null');
    if (p && p.values && typeof p.version === 'number') return p;
  } catch (e) { /* storage blocked or unreadable: fall back to defaults */ }
  return { version: 1, values: {} };
}
function saveProfile() {
  // Only tag-value rules are saved. Never rows, counts or file contents.
  const clean = { version: state.profile.version, officeName: String(state.profile.officeName || '').slice(0, 60), values: {} };
  for (const [k, v] of Object.entries(state.profile.values)) clean.values[k] = { cls: v.cls, value: v.value || '' };
  try { localStorage.setItem(PROFILE_KEY, JSON.stringify(clean)); return true; } catch (e) { return false; }
}

function toast(msg) {
  const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg;
  document.body.appendChild(t); setTimeout(() => t.remove(), 3200);
}

/* Excel serial numbers, dd/mm/yyyy and yyyy-mm-dd all become Dates. */
function parseDate(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number' && v > 1000 && v < 80000) return new Date(Date.UTC(1899, 11, 30) + v * DAY);
  const s = String(v).trim();
  if (/^\d{4,5}(\.\d+)?$/.test(s)) { const n = +s; if (n > 1000 && n < 80000) return new Date(Date.UTC(1899, 11, 30) + n * DAY); }
  let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (m) { let y = +m[3]; if (y < 100) y += y > 30 ? 1900 : 2000; const d = new Date(Date.UTC(y, +m[2] - 1, +m[1])); return isNaN(d) ? null : d; }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) { const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])); return isNaN(d) ? null : d; }
  return null;
}
function ageBracket(years) {
  if (years == null || isNaN(years) || years < 16 || years > 120) return null;
  if (years < 25) return '18-24';
  if (years >= 75) return '75+';
  const lo = Math.floor((years - 25) / 10) * 10 + 25; return `${lo}-${lo + 9}`;
}
function normAge(v) {
  const s = String(v ?? '').trim(); if (!s) return null;
  const m = s.replace(/\s/g, '').match(/^(\d{2})\s*[-–to]+\s*(\d{2})$/);
  if (m) { const b = ageBracket(+m[1]); return b; }
  if (/^\d{2}\+$/.test(s.replace(/\s/g, ''))) return ageBracket(+s.slice(0, 2));
  if (/^\d{1,3}$/.test(s)) return ageBracket(+s);
  return s; // office-specific code: kept as written
}
const truthy = v => { const s = String(v ?? '').trim().toLowerCase(); return !!s && !['n', 'no', 'false', '0', 'none', '-'].includes(s); };
/* Campaign Central's multi-value fields can arrive run together with no
   separator ("Phone BankingPFN: Ethnic/Korean"). Split where a lower-case
   letter, digit or bracket runs straight into a capital letter. */
const splitRunTogether = v => splitMulti(String(v ?? '').replace(/([a-z0-9)])(?=[A-Z])/g, '$1;'));
const splitMulti = v => String(v ?? '').split(/[;,|]/).map(x => x.trim()).filter(x => x && x.toLowerCase() !== 'none');

/* ------------------------------------------------------------------ file reading */
function parseCSV(text) {
  const rows = []; let row = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); field = ''; if (row.length > 1 || row[0] !== '') rows.push(row); row = [];
    } else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}
function loadXlsxLib() {
  if (window.XLSX) return Promise.resolve();
  return new Promise((res, rej) => {
    const s = document.createElement('script'); s.src = 'vendor/xlsx.full.min.js';
    s.onload = res; s.onerror = () => rej(new Error('Could not load the spreadsheet reader')); document.head.appendChild(s);
  });
}
async function readFile(file) {
  const name = file.name || 'file';
  if (/\.xlsx?$/i.test(name)) {
    await loadXlsxLib();
    const wb = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false, cellFormula: false, cellHTML: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '', blankrows: false });
  }
  let text = await file.text();
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  return parseCSV(text);
}

/* ------------------------------------------------------------------ office profile */
function classify(value) {
  const v = String(value || '').trim(); if (!v) return null;
  const own = state.profile.values[v];
  if (own) return { cls: own.cls, value: own.value || v, src: 'office' };
  for (const r of DEFAULT_RULES) {
    const m = v.match(new RegExp(r.re, 'i'));
    if (m) return { cls: r.cls, value: r.cap ? m[r.cap].trim() : v, src: 'rule' };
  }
  return null; // unclassified
}

/* ------------------------------------------------------------------ the import gate */
/* Persuasion and targeting tags are removed outright, even from memory. */
function dropIgnored(v) { const c = classify(v); return c && c.cls === 'ignore' && c.src === 'rule' ? '' : String(v ?? '').trim(); }
function runGate(rows, fileName) {
  const report = { fileName, rowCount: Math.max(0, rows.length - 1), kept: [], converted: [], dropped: [], strippedNames: [], strippedDob: [], strippedAddress: [], warnings: [], scanDropped: [] };
  if (rows.length < 2) throw new Error('The file has no data rows.');
  const head = rows[0].map(h => String(h ?? '').trim());
  const roles = head.map(h => {
    const n = norm(h);
    let m;
    if ((m = n.match(/^topic(\d+)$/))) return { role: 'topic', set: +m[1] };
    if ((m = n.match(/^tags?(\d+)$/))) return { role: 'tag', set: +m[1] };
    if ((m = n.match(/^interaction(?:details?)?(?:set)?(\d+)$/)) || (m = n.match(/^(?:interaction)?notes?(\d+)$/))) return { role: 'note', set: +m[1] };
    const k = COL_LOOKUP[n];
    return k ? { role: k } : { role: 'unknown' };
  });

  // Content scan backstop: a column we'd keep is dropped if its values look like
  // emails, phone numbers, dates or personal names.
  const sample = rows.slice(1, 1 + Math.min(rows.length - 1, 3000));
  const keepRoles = new Set(['gender', 'community', 'religion', 'housing', 'suburb', 'postcode', 'age', 'issues', 'topic', 'tag', 'ccTopics', 'sa1']);
  head.forEach((h, i) => {
    const r = roles[i]; if (!keepRoles.has(r.role)) return;
    let email = 0, phone = 0, date = 0, filled = 0; const uniq = new Set();
    for (const row of sample) {
      const v = String(row[i] ?? '').trim(); if (!v) continue; filled++; uniq.add(v);
      if (/[^\s@]+@[^\s@]+\.[a-z]{2,}/i.test(v)) email++;
      if (/(\+?61|0)[\s-]?\d([\s-]?\d){7,}/.test(v)) phone++;
      if (r.role !== 'topic' && r.role !== 'tag' && /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(v)) date++;
    }
    const nameLike = ['gender', 'community', 'religion', 'housing', 'suburb'].includes(r.role) && filled > 200 && uniq.size / filled > 0.6;
    if (filled && (email / filled > 0.02 || phone / filled > 0.02 || date / filled > 0.2 || nameLike)) {
      report.scanDropped.push({ col: h, why: email ? 'email addresses' : phone ? 'phone numbers' : date ? 'dates' : 'too many unique values to be a category (possible names)' });
      roles[i] = { role: 'scanDropped' };
    }
  });

  const idx = role => roles.map((r, i) => r.role === role ? i : -1).filter(i => i >= 0);
  head.forEach((h, i) => {
    const r = roles[i].role;
    if (r === 'name') report.strippedNames.push(h);
    else if (r === 'dob') report.strippedDob.push(h);
    else if (r === 'address') report.strippedAddress.push(h);
    else if (['email', 'mobile', 'phone'].includes(r)) report.converted.push({ col: h, to: 'yes/no flag (value discarded)' });
    else if (r === 'enrol') report.converted.push({ col: h, to: 'new-elector flag (date discarded)' });
    else if (r === 'note') report.converted.push({ col: h, to: 'date and channel code only (text discarded)' });
    else if (r === 'id') report.kept.push({ col: h, as: 'ID, kept in this tab only for target lists' });
    else if (['unknown', 'knownDrop', 'scanDropped'].includes(r)) { if (r !== 'scanDropped') report.dropped.push(h); }
    else if (r === 'sa1') report.dropped.push(h); // kept for Census matching in phase 5
    else report.kept.push({ col: h, as: r === 'topic' || r === 'tag' || r === 'ccTopics' ? 'tag values, read by the office profile' : DIM_LABEL[r] || r });
  });
  if (report.strippedDob.length) report.converted.push({ col: report.strippedDob.join(', '), to: 'age bracket (date of birth discarded)' });

  const sets = {};
  roles.forEach((r, i) => { if (r.set) { (sets[r.set] = sets[r.set] || {})[r.role] = i; } });
  const setNums = Object.keys(sets).map(Number).sort((a, b) => a - b);
  const I = Object.fromEntries(['id', 'gender', 'community', 'religion', 'housing', 'suburb', 'postcode', 'age', 'issues', 'email', 'mobile', 'phone', 'enrol', 'ccTopics'].map(k => [k, idx(k)]));
  const dobCols = idx('dob');
  const hasReach = I.email.length + I.mobile.length + I.phone.length > 0;
  const hasEnrol = I.enrol.length > 0;
  const today = new Date();

  /* Build people. Values that could identify someone are never copied out of
     the row: we read what we need and keep only categories, flags and dates of
     contact. */
  const people = []; const byId = new Map(); let dupes = 0;
  const unclassified = new Map(); // value -> {cols:Set, people:count}
  let datedContacts = 0, undatedContacts = 0, adminSkipped = 0, maxDay = -Infinity;
  let persuasionDropped = 0;
  const noteNoCode = { n: 0 };

  const vals = (row, cols) => cols.map(c => String(row[c] ?? '').trim()).filter(Boolean);
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const id = I.id.length ? String(row[I.id[0]] ?? '').trim() : '';
    const attrs = { suburb: [], postcode: [], age: [], gender: [], community: [], religion: [], housing: [], issues: [], newElector: [], reach: [] };
    for (const k of ['suburb', 'postcode', 'gender', 'community', 'religion', 'housing']) for (const v of vals(row, I[k])) attrs[k].push(v);
    for (const v of vals(row, I.age)) { const a = normAge(v); if (a) attrs.age.push(a); }
    if (!attrs.age.length && dobCols.length) {
      const d = parseDate(row[dobCols[0]]);
      if (d) { const yrs = (today - d) / (365.25 * DAY); const a = ageBracket(yrs); if (a) attrs.age.push(a); }
    }
    for (const c of I.issues) for (const v of splitMulti(row[c])) attrs.issues.push(v);
    if (hasReach) {
      const e = I.email.some(c => truthy(row[c])), m = I.mobile.some(c => truthy(row[c])), p = I.phone.some(c => truthy(row[c]));
      if (e) attrs.reach.push('Has email'); if (m) attrs.reach.push('Has mobile'); if (p) attrs.reach.push('Has phone');
      if (!e && !m && !p) attrs.reach.push('No contact details');
    }
    if (hasEnrol) {
      const d = parseDate(row[I.enrol[0]]);
      attrs.newElector.push(d && (today - d) / DAY <= 365 ? 'New elector (12 mo)' : 'Enrolled 12+ months');
    }

    const contacts = []; const seenTags = new Set();
    const applyTag = (v, col) => {
      if (!v) return null;
      const c = classify(v);
      if (!c) {
        if (v.length <= 60) { // longer values could be free text, never show them
          const u = unclassified.get(v) || { cols: new Set(), people: 0 }; u.cols.add(col);
          if (!seenTags.has(v)) { u.people++; seenTags.add(v); } unclassified.set(v, u);
        }
        return null;
      }
      if (c.cls === 'ignore' && c.src === 'rule') persuasionDropped++;
      if (c.cls === 'community' || c.cls === 'religion' || c.cls === 'housing') { if (!attrs[c.cls].includes(c.value)) attrs[c.cls].push(c.value); }
      if (c.cls === 'issue') { if (!attrs.issues.includes(c.value)) attrs.issues.push(c.value); }
      return c.cls;
    };
    for (const n of setNums) {
      const s = sets[n];
      const topic = s.topic != null ? String(row[s.topic] ?? '').trim() : '';
      const tag = s.tag != null ? String(row[s.tag] ?? '').trim() : '';
      const note = s.note != null ? String(row[s.note] ?? '') : '';
      const ct = applyTag(topic, head[s.topic] || 'TOPIC'), cg = applyTag(tag, head[s.tag] || 'Tag');
      if (!topic && !tag && !note.trim()) continue;
      if (ct === 'admin' || cg === 'admin') { if (note.trim()) adminSkipped++; continue; }
      // From the note we keep only the date at the start and the first channel code.
      const dm = note.match(/^\s*(\d{1,2}\/\d{1,2}\/\d{4})/);
      const d = dm ? parseDate(dm[1]) : null;
      const cm = note.toUpperCase().match(CODE_RE);
      const code = cm ? cm[1].replace(/\s+/g, ' ') : null;
      const isContact = code || ct === 'contact' || cg === 'contact';
      if (!isContact) continue;
      if (note.trim() && !code) noteNoCode.n++;
      const [ch, dir] = code ? CHANNEL_CODES[code] : ['Other', null];
      const t = d ? dayNum(d) : null;
      if (t != null && t <= dayNum(today)) { maxDay = Math.max(maxDay, t); datedContacts++; } else undatedContacts++;
      contacts.push({ t: t != null && t <= dayNum(today) ? t : null, ch, dir });
    }
    for (const c of I.ccTopics) for (const v of splitRunTogether(row[c])) {
      const cls = applyTag(v, head[c]);
      if (cls === 'contact') { contacts.push({ t: null, ch: 'Other', dir: null }); undatedContacts++; }
    }

    for (const k of ['community', 'religion', 'housing', 'issues']) if (!attrs[k].length) attrs[k].push(UNTAGGED);
    for (const k of ['suburb', 'postcode', 'age', 'gender']) if (!attrs[k].length) attrs[k].push('Unknown');

    const person = { id, attrs, contacts };
    if (id && byId.has(id)) { // duplicate record: merge contacts and tags
      const p = byId.get(id); dupes++;
      p.contacts.push(...contacts);
      for (const k in attrs) for (const v of attrs[k]) if (v !== UNTAGGED && v !== 'Unknown' && !p.attrs[k].includes(v)) { p.attrs[k] = p.attrs[k].filter(x => x !== UNTAGGED && x !== 'Unknown'); p.attrs[k].push(v); }
      continue;
    }
    if (id) byId.set(id, person);
    people.push(person);
  }

  /* A cleaned copy of the file, kept only in memory so that office-profile
     answers can be re-applied without asking for the file again. It holds no
     names, dates of birth, addresses, contact details or note text. */
  const sHead = [], sCols = [];
  head.forEach((h, i) => {
    const r = roles[i];
    if (r.role === 'topic') { sHead.push(`TOPIC (${r.set})`); sCols.push(row => dropIgnored(row[i])); }
    else if (r.role === 'tag') { sHead.push(`Tag (${r.set})`); sCols.push(row => dropIgnored(row[i])); }
    else if (r.role === 'ccTopics') { sHead.push(h); sCols.push(row => splitRunTogether(row[i]).filter(v => dropIgnored(v)).join('; ')); }
    else if (r.role === 'note') {
      sHead.push(`Interaction Details Set ${r.set}`);
      sCols.push(row => { const n = String(row[i] ?? ''); const dm = n.match(/^\s*(\d{1,2}\/\d{1,2}\/\d{4})/); const cm = n.toUpperCase().match(CODE_RE); return [dm ? dm[1] : '', cm ? cm[1].replace(/\s+/g, ' ') : ''].filter(Boolean).join(' - '); });
    }
    else if (['id', 'gender', 'community', 'religion', 'housing', 'suburb', 'postcode', 'age', 'issues'].includes(r.role)) { sHead.push(h); sCols.push(row => row[i]); }
    else if (['email', 'mobile', 'phone'].includes(r.role)) { sHead.push('Has ' + r.role); sCols.push(row => truthy(row[i]) ? 'Y' : ''); }
    else if (r.role === 'enrol') { sHead.push('EnrolDate'); sCols.push(row => { const d = parseDate(row[i]); return d && (today - d) / DAY <= 365 ? `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()}` : (d ? '01/01/1900' : ''); }); }
  });
  if (dobCols.length && !I.age.length) { sHead.push('Age Bracket'); sCols.push(row => { const d = parseDate(row[dobCols[0]]); return d ? (ageBracket((today - d) / (365.25 * DAY)) || '') : ''; }); }
  const sanitized = [sHead, ...rows.slice(1).map(row => sCols.map(f => f(row)))];

  const mode = datedContacts > 0 ? 'dated' : 'undated';
  const asOfDay = isFinite(maxDay) ? maxDay : dayNum(today);
  const dims = DIMS.filter(d => {
    if (d.key === 'reach') return hasReach;
    if (d.key === 'newElector') return hasEnrol;
    return people.some(p => p.attrs[d.key].some(v => v !== UNTAGGED && v !== 'Unknown'));
  }).map(d => d.key);

  if (report.strippedNames.length) report.warnings.push({ lvl: 'bad', title: `Names found and removed: ${report.strippedNames.join(', ')}`, body: 'These columns were dropped as the file was read and were never shown, counted or stored. Before using real data, the office export should remove names before the file leaves Campaign Central.' });
  if (report.strippedDob.length) report.warnings.push({ lvl: 'bad', title: `Dates of birth found: ${report.strippedDob.join(', ')}`, body: 'Each date of birth was turned into an age bracket in memory and then discarded. Supplying an age bracket column instead avoids this.' });
  if (report.strippedAddress.length) report.warnings.push({ lvl: 'bad', title: `Address details removed: ${report.strippedAddress.join(', ')}`, body: 'Street and postal address fields are not needed and were dropped.' });
  for (const s of report.scanDropped) report.warnings.push({ lvl: 'bad', title: `Column "${s.col}" dropped by the content check`, body: `It looked like it contained ${s.why}, so none of it was used.` });
  if (dupes) report.warnings.push({ lvl: 'warn', title: `${fmt(dupes)} duplicate records merged`, body: 'Rows sharing an ID were combined into one person.' });
  if (mode === 'dated' && undatedContacts) report.warnings.push({ lvl: 'warn', title: `${fmt(undatedContacts)} contacts have no date`, body: 'They count towards "ever contacted" but cannot be placed in a time period.' });
  if (!I.id.length) report.warnings.push({ lvl: 'warn', title: 'No ID column', body: 'Target-list downloads need a Campaign Central ID column. Everything else works.' });
  if (!hasReach) report.warnings.push({ lvl: 'info', title: 'No email, mobile or phone columns', body: 'Reachability is not available for this file.' });
  if (persuasionDropped) report.warnings.push({ lvl: 'bad', title: `${fmt(persuasionDropped)} persuasion or targeting tags removed`, body: 'Voting-intention and targeting tags are never read, counted or kept, even in memory.' });
  if (adminSkipped) report.warnings.push({ lvl: 'info', title: `${fmt(adminSkipped)} admin records excluded`, body: 'Record adjustments and contact-detail changes are not counted as contact.' });

  return {
    people, report, mode, asOfDay, dims, hasIds: I.id.length > 0, hasReach, hasEnrol,
    unclassified: [...unclassified.entries()].map(([value, u]) => ({ value, cols: [...u.cols], people: u.people })).sort((a, b) => b.people - a.people),
    datedContacts, undatedContacts, importedAt: new Date(), fileName, sanitized,
  };
}

/* ------------------------------------------------------------------ counting */
function windowFor(periodIdx, shift = 0) {
  const d = state.data; const days = PERIODS[periodIdx].days;
  const end = d.asOfDay - shift * days; return [end - days + 1, end];
}
function inWin(c, w) { return w === null ? true : c.t != null && c.t >= w[0] && c.t <= w[1]; }

function aggregate(dimKey, w) {
  const d = state.data; const groups = new Map();
  let C = 0, R = 0, N = d.people.length;
  const tot = { office: 0, constituent: 0, ch: {} };
  for (const p of d.people) {
    let n = 0, off = 0, con = 0, lastOff = null; const ch = {};
    for (const c of p.contacts) {
      if (!inWin(c, w)) continue;
      n++; ch[c.ch] = (ch[c.ch] || 0) + 1;
      if (c.dir === 'out') { off++; if (c.t != null && (lastOff == null || c.t > lastOff)) lastOff = c.t; } else if (c.dir === 'in') con++;
    }
    C += n; if (n) R++;
    tot.office += off; tot.constituent += con; for (const k in ch) tot.ch[k] = (tot.ch[k] || 0) + ch[k];
    const reachable = p.attrs.reach.length ? !p.attrs.reach.includes('No contact details') : null;
    for (const g of p.attrs[dimKey]) {
      let G = groups.get(g);
      if (!G) { G = { name: g, dim: dimKey, e: 0, c: 0, r: 0, reachable: 0, email: 0, noDetails: 0, office: 0, constituent: 0, ch: {}, lastOff: null }; groups.set(g, G); }
      G.e++; G.c += n; if (n) G.r++; G.office += off; G.constituent += con;
      if (reachable) G.reachable++; if (p.attrs.reach.includes('Has email')) G.email++; if (reachable === false) G.noDetails++;
      for (const k in ch) G.ch[k] = (G.ch[k] || 0) + ch[k];
      if (lastOff != null && (G.lastOff == null || lastOff > G.lastOff)) G.lastOff = lastOff;
    }
  }
  const list = [...groups.values()];
  const eSum = list.reduce((s, g) => s + g.e, 0), cSum = list.reduce((s, g) => s + g.c, 0);
  const overallReach = N ? R / N : 0;
  for (const g of list) {
    g.small = g.e < MIN_CELL;
    g.eShare = g.e / eSum; g.cShare = cSum ? g.c / cSum : 0;
    g.fs = cSum ? g.cShare / g.eShare : null;
    g.reachPct = g.e ? g.r / g.e : 0;
    g.status = g.fs == null ? 'n' : g.fs < 0.8 ? 'u' : g.fs > 1.25 ? 'o' : 'b';
    g.lowReach = !g.small && g.fs != null && g.fs >= 0.8 && overallReach > 0 && g.reachPct < overallReach * 0.5;
    g.score = g.fs == null ? 0 : Math.round(100 * Math.min(1, g.fs));
  }
  // Balance score: 100 means contact is spread across groups in exact proportion to electors.
  const visible = list.filter(g => !g.small);
  const balance = cSum ? Math.round(100 * (1 - 0.5 * list.reduce((s, g) => s + Math.abs(g.cShare - g.eShare), 0))) : 0;
  return { list, visible, N, C, R, tot, balance, overallReach, smallCount: list.length - visible.length };
}
const periodWindow = () => state.data.mode === 'dated' ? windowFor(state.period) : null;
const prevWindow = () => state.data.mode === 'dated' ? windowFor(state.period, 1) : null;

/* ------------------------------------------------------------------ colours */
function fsColour(fs) {
  if (fs == null) return ['#EEF1F5', false];
  if (fs < 0.5) return ['#E85D2F', true]; if (fs < 0.65) return ['#F08B5B', true]; if (fs < 0.8) return ['#F8C9AC', false];
  if (fs <= 1.25) return ['#EEF1F5', false];
  if (fs <= 1.5) return ['#DBE7FA', false]; if (fs <= 1.8) return ['#8FB1EE', false]; return ['#2E6FE0', true];
}
const STATUS_TXT = { u: 'Underserved', o: 'Overserved', b: 'Balanced', n: 'No contact' };
const fsTxt = fs => fs == null ? '–' : (fs >= 10 ? fs.toFixed(0) : fs.toFixed(1)) + '×';

/* ------------------------------------------------------------------ treemap (squarified) */
function squarify(items, x, y, w, h) {
  const out = []; const total = items.reduce((s, i) => s + i.v, 0); if (!total) return out;
  const scale = w * h / total; let rest = items.map(i => ({ ...i, a: i.v * scale }));
  while (rest.length) {
    const short = Math.min(w, h); let row = [rest[0]], best = worst(row, short), k = 1;
    while (k < rest.length) { const next = [...row, rest[k]]; const wr = worst(next, short); if (wr > best) break; row = next; best = wr; k++; }
    const sum = row.reduce((s, r) => s + r.a, 0);
    if (w >= h) { const cw = sum / h; let cy = y; for (const r of row) { const rh = r.a / cw; out.push({ ...r, x, y: cy, w: cw, h: rh }); cy += rh; } x += cw; w -= cw; }
    else { const ch = sum / w; let cx = x; for (const r of row) { const rw = r.a / ch; out.push({ ...r, x: cx, y, w: rw, h: ch }); cx += rw; } y += ch; h -= ch; }
    rest = rest.slice(k);
  }
  return out;
  function worst(row, s) { const sum = row.reduce((t, r) => t + r.a, 0); const mx = Math.max(...row.map(r => r.a)), mn = Math.min(...row.map(r => r.a)); return Math.max(s * s * mx / (sum * sum), sum * sum / (s * s * mn)); }
}

/* ------------------------------------------------------------------ rendering */
function render() {
  const main = $('#main');
  const p = state.profile;
  $('#profVer').textContent = 'v' + p.version;
  const uq = state.data ? state.data.unclassified.length : 0;
  $('#uqBadge').textContent = uq; $('#uqBadge').classList.toggle('hidden', !uq);
  if (!state.data) { main.innerHTML = renderEmpty(); bindEmpty(); return; }
  main.innerHTML = renderDashboard();
  bindDashboard();
}

function renderEmpty() {
  $('#officeName').textContent = 'No file loaded'; $('#officeAv').textContent = '—';
  return `
  <div class="topbar"><span class="chip grey"><i></i>No file open</span><span class="meta">Load an export to see your coverage</span></div>
  <div class="wrap"><div class="empty">
    <div class="kicker">Audience coverage · trial build</div>
    <h1>Who are we actually reaching?</h1>
    <p class="lede">Load a de-identified Campaign Central export to see which groups the office has spoken to enough, which it hasn't, and where the next doorknock, phone bank or email campaign should go.</p>
    <div class="drop" id="drop" style="margin-top:22px">
      <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="#2E6FE0" stroke-width="1.8"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
      <h2>Drop your export here</h2>
      <p>CSV or Excel. The file is read in this tab and never uploaded.</p>
      <button class="btn primary" id="pick">Choose a file</button>
    </div>
    <div class="opts">
      <div class="card opt" id="useSample"><h3>Try the sample electorate</h3><p>18,000 fictional electors with a year and a half of dated contact history.</p></div>
      <div class="card opt" id="usePoison"><h3>Test the privacy gate</h3><p>Loads a small fictional file full of names, dates of birth, addresses and phone numbers, so you can watch them get stripped.</p></div>
      <a class="card opt" href="sample/demo-electorate-export.csv" download style="color:inherit;text-decoration:none"><h3>Download the sample file</h3><p>See the column layout the tool expects. Every row is invented.</p></a>
    </div>
    <div class="card how">
      <h2 style="font-size:16px;margin:0">Preparing an export</h2>
      <ol>
        <li>Export the whole electorate from Campaign Central, including people never contacted, with their interaction history.</li>
        <li>Delete first name, last name and other name columns, and the street and postal address columns.</li>
        <li>Replace date of birth with an age bracket (18-24, 25-34, 35-44 and so on).</li>
        <li>Keep the Campaign Central ID if you want to download target lists. IDs stay in this tab.</li>
      </ol>
      <p class="note" style="margin-top:10px">If anything is missed, the import gate removes it before anything is counted and tells you what it found.</p>
    </div>
  </div></div>`;
}
function bindEmpty() {
  const drop = $('#drop');
  $('#pick').onclick = () => $('#fileInput').click();
  drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files[0]) startImport(e.dataTransfer.files[0]); };
  $('#useSample').onclick = loadSample;
  $('#usePoison').onclick = () => startImport(makePoisonFile());
}

function renderDashboard() {
  const d = state.data; const w = periodWindow(); const per = PERIODS[state.period];
  if (!d.dims.includes(state.dim)) state.dim = d.dims[0];
  const A = aggregate(state.dim, w);
  const prev = d.mode === 'dated' ? aggregate(state.dim, prevWindow()) : null;
  const vis = A.visible.filter(g => g.name !== 'Unknown');
  const counts = { u: 0, o: 0, b: 0, l: 0 };
  // status counts across every dimension, like the mockup's summary pills
  const allGroups = [];
  for (const k of d.dims) { const a = k === state.dim ? A : aggregate(k, w); for (const g of a.visible) if (g.name !== 'Unknown' && g.name !== UNTAGGED) allGroups.push(g); }
  for (const g of allGroups) { counts[g.status] = (counts[g.status] || 0) + 1; if (g.lowReach) counts.l++; }
  if (!state.sel || !A.visible.find(g => g.name === state.sel)) {
    const cand = vis.filter(g => g.status === 'u').sort((a, b) => b.e - a.e)[0] || vis.sort((a, b) => b.e - a.e)[0];
    state.sel = cand ? cand.name : null;
  }
  const tagDims = ['community', 'religion', 'housing', 'issues'].filter(k => d.dims.includes(k));
  const completeness = tagDims.length ? Math.round(tagDims.reduce((s, k) => s + d.people.filter(p => !p.attrs[k].includes(UNTAGGED)).length / d.people.length, 0) / tagDims.length * 100) : null;
  const delta = prev && prev.C ? A.balance - prev.balance : null;
  const officeShare = A.tot.office + A.tot.constituent ? A.tot.office / (A.tot.office + A.tot.constituent) : 0;
  const asOf = fromDay(d.asOfDay);
  const officeTitle = d.officeName;
  $('#officeName').textContent = officeTitle; $('#officeAv').textContent = officeTitle.replace(/[^A-Za-z\s]/g, '').split(/\s+/).filter(Boolean).map(s => s[0]).join('').slice(0, 2).toUpperCase();

  return `
  <div class="topbar">
    <span class="chip blue"><i></i>${d.mode === 'dated' ? 'Dated mode' : 'Undated mode'}</span>
    <span class="meta" title="Loaded ${esc(d.importedAt.toLocaleString('en-AU'))}">${fmt(d.people.length)} electors · as at ${esc(fmtDate(asOf))}</span>
    <span class="meta mono" style="font-size:12px">profile v${state.profile.version}</span>
    <span class="ok"><i style="width:7px;height:7px;border-radius:50%;background:var(--green-m);display:inline-block"></i>File open in this tab${d.hasIds ? ' · target lists available' : ''}</span>
    <button class="btn" id="newImport"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>New import</button>
    <button class="btn" id="closeFile" title="Clear the file from this tab">Close</button>
  </div>
  <div class="wrap">
    <div class="hero">
      <div>
        <div class="kicker">Audience coverage · ${esc(officeTitle)} · ${esc(d.mode === 'dated' ? per.long : 'All contact on record')}</div>
        <h1>Who are we actually reaching?</h1>
        <p class="lede">Which groups the office has spoken to enough, which it hasn't, and where the next doorknock, phone bank or email campaign should go.</p>
        <div class="pills">
          <span class="pill u"><i></i>${counts.u} underserved</span><span class="pill o"><i></i>${counts.o} overserved</span>
          <span class="pill b"><i></i>${counts.b} balanced</span><span class="pill l"><i></i>${counts.l} low reach</span>
        </div>
      </div>
      <div class="stats">
        ${completeness != null ? `<div class="stat"><div class="k">Tag completeness</div><div class="v">${completeness}<small>%</small></div><div class="s">${esc(tagDims.map(k => DIM_LABEL[k].toLowerCase()).join(', '))}</div></div><div class="vsep"></div>` : ''}
        <div class="stat big" title="100 means the office's contact is spread across ${esc(DIM_LABEL[state.dim].toLowerCase())} groups in exact proportion to the number of electors in each.">
          <div class="k">Coverage balance</div><div class="v">${A.balance}<small>/100</small></div>
          <div class="s">across ${esc(DIM_LABEL[state.dim].toLowerCase())} · ${delta == null ? 'no earlier period' : (delta >= 0 ? '+' : '−') + Math.abs(delta) + ' vs previous period'}</div>
          <div class="s" style="max-width:250px;margin-left:auto;margin-top:4px;color:var(--faint)">100 = contact spread in exact proportion to the number of electors in each group</div>
        </div>
      </div>
    </div>

    <div class="card period">
      <div class="lbl"><div class="kicker" style="font-size:11px">Time period</div><b>${esc(d.mode === 'dated' ? per.long : 'All contact on record')}</b></div>
      <div class="slider ${d.mode === 'dated' ? '' : 'off'}">
        <div class="track"><div class="fill" style="width:${state.period / (PERIODS.length - 1) * 100}%"></div></div>
        <div class="stops">${PERIODS.map((p, i) => `<button class="stop ${i < state.period ? 'past' : ''} ${i === state.period ? 'on' : ''}" data-period="${i}"><span class="dot"></span>${p.short}</button>`).join('')}</div>
      </div>
      <div class="mini"><div class="k">CONTACTS</div><b>${fmt(A.C)}</b><span>${fmt(A.R)} people reached (${pct(A.R, A.N)}%)</span></div>
      ${A.tot.office + A.tot.constituent === 0 ? `<div class="mini" style="text-align:left"><div class="k">WHO STARTED IT</div><span>Not recorded in this file</span></div>` : `
      <div class="mini" style="text-align:left"><div class="k">WHO STARTED IT</div>
        <div class="split"><i style="width:${officeShare * 100}%;background:var(--blue)"></i><i style="width:${(1 - officeShare) * 100}%;background:var(--purple)"></i></div>
        <div class="splitlbl"><span><b style="font-size:12px;display:inline;color:var(--blue-d)">${Math.round(officeShare * 100)}%</b> office</span><span><b style="font-size:12px;display:inline;color:var(--purple)">${100 - Math.round(officeShare * 100)}%</b> constituent</span></div>
      </div>`}
    </div>

    <div class="grid2">
      <div class="card">
        <div class="cardhead"><div><div class="kicker" style="font-size:11px">Who you're reaching</div><h2>Attention by group</h2><div class="sub">Box size = electors · colour = share of contacts vs share of electors</div></div><div class="note">From the roll · every elector counted</div></div>
        <div class="tabs">${d.dims.map(k => `<button class="tab ${k === state.dim ? 'on' : ''}" data-dim="${k}">${DIM_LABEL[k]}</button>`).join('')}</div>
        ${renderTreemap(A)}
        <div class="legend">
          <span style="color:var(--orange);font-weight:600">Underserved</span>
          <span class="sw">${['#E85D2F', '#F08B5B', '#F8C9AC', '#EEF1F5', '#DBE7FA', '#8FB1EE', '#2E6FE0'].map(c => `<i style="background:${c}"></i>`).join('')}</span>
          <span style="color:var(--blue-d);font-weight:600">Overserved</span>
          <span style="display:inline-flex;align-items:center;gap:5px;color:var(--amber);font-weight:600"><i style="width:8px;height:8px;border-radius:50%;background:var(--amber-m);display:inline-block"></i>Low reach</span>
          <span class="note">People count in every group they're tagged with. Untagged people form their own group.${A.smallCount ? ` ${A.smallCount} group${A.smallCount > 1 ? 's' : ''} under ${MIN_CELL} electors hidden.` : ''}</span>
        </div>
      </div>
      ${renderSegment(A)}
    </div>

    ${d.mode === 'dated' ? renderTrend() : ''}

    <div class="grid2b">
      ${renderGap(allGroups)}
      ${renderHeatmap(w)}
    </div>

    ${renderMoves(allGroups, A)}
  </div>`;
}

function renderTreemap(A) {
  const items = A.visible.slice().sort((a, b) => b.e - a.e);
  const top = items.slice(0, 18);
  const rest = items.slice(18);
  const data = top.map(g => ({ v: g.e, g }));
  if (rest.length) data.push({ v: rest.reduce((s, g) => s + g.e, 0), g: { name: `${rest.length} smaller groups`, e: rest.reduce((s, g) => s + g.e, 0), other: true } });
  const W = 1000, H = 440;
  const rects = squarify(data, 0, 0, W, H);
  return `<div class="tm" id="tm">${rects.map(r => {
    const g = r.g; const [bg, dark] = g.other ? ['#F4F6F9', false] : fsColour(g.fs);
    const big = r.w > 120 && r.h > 90, med = r.w > 70 && r.h > 44;
    return `<button class="tile ${dark ? 'dark' : ''} ${g.name === state.sel && !g.other ? 'sel' : ''}" ${g.other ? '' : `data-group="${esc(g.name)}"`} style="left:${r.x / W * 100}%;top:${r.y / H * 100}%;width:${r.w / W * 100}%;height:${r.h / H * 100}%;background:${bg}" title="${esc(g.name)}: ${fmt(g.e)} electors${g.other ? '' : ` · ${Math.round(g.reachPct * 100)}% reached · ${fsTxt(g.fs)} fair share`}">
      ${med ? `<b>${esc(g.name)}</b><div class="e">${fmt(g.e)} electors</div>` : ''}
      ${big && !g.other ? `<div class="d">${Math.round(g.reachPct * 100)}% reached · ${fsTxt(g.fs)} fair share</div><span class="st">${STATUS_TXT[g.status].toUpperCase()}</span>` : ''}
      ${g.lowReach ? '<span class="lr" title="Low reach"></span>' : ''}
    </button>`;
  }).join('')}</div>`;
}

function renderSegment(A) {
  const d = state.data; const g = A.list.find(x => x.name === state.sel);
  if (!g) return '<div class="card"><div class="sub">Select a group.</div></div>';
  const chTotal = Object.values(g.ch).reduce((s, v) => s + v, 0);
  const chs = Object.entries(g.ch).sort((a, b) => b[1] - a[1]);
  const offPct = g.office + g.constituent ? g.office / (g.office + g.constituent) : 0;
  const low = g.reachPct < A.overallReach * 0.75;
  const notReached = g.e - g.r;
  return `<div class="card seg">
    <div class="cardhead"><div class="kicker" style="font-size:11px">Segment detail</div><span class="status st-${g.status}">${STATUS_TXT[g.status]}</span></div>
    <h2 style="margin-top:8px">${esc(g.name)}</h2>
    <div class="sub">${esc(DIM_LABEL[g.dim])} · ${fmt(g.e)} electors (${(g.eShare * 100).toFixed(1)}%)</div>
    <div class="scorebox"><span class="n" style="color:${g.status === 'u' ? 'var(--orange)' : g.status === 'o' ? 'var(--blue-d)' : 'var(--green)'}">${g.score}</span><span class="note">/100 coverage</span><span class="fs" style="color:${g.status === 'u' ? 'var(--orange)' : 'var(--ink2)'}">${fsTxt(g.fs)} fair share</span></div>
    <div class="note">${(g.cShare * 100).toFixed(1)}% of contacts · ${(g.eShare * 100).toFixed(1)}% of electors</div>
    <div class="reachline ${low ? 'low' : 'ok'}"><i></i>Reached ${Math.round(g.reachPct * 100)}% of electors in this group <span style="font-weight:500;opacity:.8">(electorate ${Math.round(A.overallReach * 100)}%)</span></div>
    <div class="sec">Their funnel</div>
    ${funnelRow('Electors', g.e, g.e)}
    ${d.hasReach ? funnelRow('Reachable', g.reachable, g.e) : ''}
    ${funnelRow('Reached', g.r, g.e)}
    <div class="sec">Channel mix</div>
    ${chTotal ? `<div class="mix">${chs.map(([k, v]) => `<i style="width:${v / chTotal * 100}%;background:${CHANNEL_COLOURS[k] || '#94A3B8'}"></i>`).join('')}</div>
    <div class="mixl">${chs.map(([k, v]) => `<span style="--c:${CHANNEL_COLOURS[k] || '#94A3B8'}">${esc(k)} ${pct(v, chTotal)}%</span>`).join('')}</div>` : '<div class="note">No contact in this period.</div>'}
    <div class="sec">Who started the contact</div>
    ${g.office + g.constituent === 0 ? '<div class="note">Not recorded in this file for this period.</div>' : `<div class="mix"><i style="width:${offPct * 100}%;background:var(--blue)"></i><i style="width:${(1 - offPct) * 100}%;background:var(--purple)"></i></div>
    <div class="mixl" style="justify-content:space-between"><span style="--c:var(--blue)">Office reached out ${Math.round(offPct * 100)}%</span><span style="--c:var(--purple)">Constituent came to us ${100 - Math.round(offPct * 100)}%</span></div>`}
    <div class="note" style="margin-top:10px">${g.lastOff != null ? `Last office-initiated contact · ${esc(fmtDate(fromDay(g.lastOff)))}${d.asOfDay - g.lastOff > 0 ? ` (${d.asOfDay - g.lastOff} day${d.asOfDay - g.lastOff === 1 ? '' : 's'} before the latest record)` : ''}` : 'No office-initiated contact in this period'}</div>
    <div style="margin-top:16px"><button class="btn primary block" data-target="${esc(g.dim)}|${esc(g.name)}" ${d.hasIds ? '' : 'disabled title="This file has no ID column"'}><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>Download target list</button></div>
    <div class="note" style="text-align:center;margin-top:6px">${d.hasIds ? `${fmt(notReached)} IDs not reached ${d.mode === 'dated' ? 'in this period' : 'yet'} · available while the file is open` : 'Needs a Campaign Central ID column'}</div>
  </div>`;
}
function funnelRow(label, n, of) {
  return `<div class="fun"><div class="r"><em style="font-style:normal">${label}</em><em style="font-style:normal"><b>${fmt(n)}</b><span>${pct(n, of)}%</span></em></div><div class="bar"><i style="width:${of ? n / of * 100 : 0}%"></i></div></div>`;
}

function renderTrend() {
  const d = state.data; const months = [];
  const end = fromDay(d.asOfDay);
  for (let i = 11; i >= 0; i--) {
    const s = new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth() - i, 1));
    const e = new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth() - i + 1, 0));
    months.push({ label: s.toLocaleDateString('en-AU', { month: 'short' }), w: [dayNum(s), Math.min(dayNum(e), d.asOfDay)], partial: i === 0 && dayNum(e) > d.asOfDay });
  }
  const pts = months.map(m => { let c = 0, r = 0, off = 0; for (const p of d.people) { let n = 0; for (const x of p.contacts) if (inWin(x, m.w)) { n++; if (x.dir === 'out') off++; } c += n; if (n) r++; } return { ...m, c, r, off }; });
  const max = Math.max(1, ...pts.map(p => p.c));
  const W = 1000, H = 150, pad = 26, bw = (W - pad * 2) / pts.length;
  const bars = pts.map((p, i) => {
    const h = p.c / max * (H - 40); const ho = p.off / max * (H - 40); const x = pad + i * bw + 6;
    return `<g><title>${p.label}: ${fmt(p.c)} contacts, ${fmt(p.r)} people reached${p.partial ? ' (month to date)' : ''}</title>
      <rect x="${x}" y="${H - 22 - h}" width="${bw - 12}" height="${h}" rx="3" fill="#C9B8F0" ${p.partial ? 'opacity=".55"' : ''}/>
      <rect x="${x}" y="${H - 22 - ho}" width="${bw - 12}" height="${ho}" rx="3" fill="#2E6FE0" ${p.partial ? 'opacity=".55"' : ''}/>
      <text x="${x + (bw - 12) / 2}" y="${H - 6}" text-anchor="middle" font-size="12" fill="#64748B">${p.label}</text>
      <text x="${x + (bw - 12) / 2}" y="${H - 26 - h}" text-anchor="middle" font-size="11" fill="#475569">${fmt(p.c)}</text></g>`;
  }).join('');
  return `<div class="card trend">
    <div class="cardhead"><div><div class="kicker" style="font-size:11px">Over time</div><h2>Contact by month</h2><div class="sub">Built from the dated contact history in this one file. Nothing is stored between imports.</div></div>
    <div class="legend" style="margin:0"><span style="display:inline-flex;gap:6px;align-items:center"><i style="width:10px;height:10px;border-radius:2px;background:#2E6FE0;display:inline-block"></i>Office reached out</span><span style="display:inline-flex;gap:6px;align-items:center"><i style="width:10px;height:10px;border-radius:2px;background:#C9B8F0;display:inline-block"></i>Constituent came to us</span></div></div>
    <svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" role="img" aria-label="Contacts per month for the last 12 months">${bars}</svg>
  </div>`;
}

function renderGap(allGroups) {
  const minE = Math.max(MIN_CELL, Math.round(state.data.people.length * 0.01));
  const pool = allGroups.filter(g => g.e >= minE && g.fs != null && g.name !== UNTAGGED && g.name !== 'Unknown');
  const lo = pool.slice().sort((a, b) => a.fs - b.fs).slice(0, 6);
  const hi = pool.slice().sort((a, b) => b.fs - a.fs).slice(0, 6).reverse();
  const rows = [...lo, ...hi.filter(g => !lo.includes(g))];
  const pos = fs => { const v = Math.log2(Math.max(0.125, Math.min(8, fs))); return v; }; // -3..3
  return `<div class="card gap">
    <div class="kicker" style="font-size:11px">Attention gap</div><h2>Overserved vs underserved</h2>
    <div class="sub">Share of contacts vs share of electors, across every dimension. 1.0× is a fair share.</div>
    <div class="gaphead"><div></div><div><span>← less than fair share</span><span>more →</span></div><div></div></div>
    ${rows.map(g => {
      const v = pos(g.fs); const half = 50; const len = Math.abs(v) / 3 * half;
      const left = v < 0 ? half - len : half;
      return `<div class="row" data-jump="${esc(g.dim)}|${esc(g.name)}"><div class="nm"><b>${esc(g.name)}</b><span>${esc(DIM_LABEL[g.dim])} · ${fmt(g.e)} electors</span></div>
      <div class="axis"><i style="left:${left}%;width:${Math.max(1, len)}%;background:${v < 0 ? 'var(--orange-m)' : 'var(--blue)'}"></i></div>
      <div class="v" style="color:${v < 0 ? 'var(--orange)' : 'var(--blue-d)'}">${fsTxt(g.fs)}</div></div>`;
    }).join('')}
    <div class="note" style="margin-top:8px">Groups under ${fmt(minE)} electors are left out of this list.</div>
  </div>`;
}

function renderHeatmap(w) {
  const d = state.data;
  if (!d.dims.includes(state.hmRow)) state.hmRow = d.dims[0];
  if (!d.dims.includes(state.hmCol) || state.hmCol === state.hmRow) state.hmCol = d.dims.find(k => k !== state.hmRow) || state.hmRow;
  const R = aggregate(state.hmRow, w).visible.filter(g => g.name !== 'Unknown').sort((a, b) => b.e - a.e).slice(0, 12).map(g => g.name);
  let Cg = aggregate(state.hmCol, w).visible.filter(g => g.name !== 'Unknown').sort((a, b) => b.e - a.e).slice(0, 8).map(g => g.name);
  if (state.hmCol === 'age') Cg = AGE_BRACKETS.filter(a => Cg.includes(a));
  // pair counts
  const cell = {}; let N = 0, C = 0;
  for (const p of d.people) {
    let n = 0; for (const c of p.contacts) if (inWin(c, w)) n++;
    N++; C += n;
    for (const r of p.attrs[state.hmRow]) for (const c of p.attrs[state.hmCol]) { const k = r + '\u0000' + c; const x = cell[k] || (cell[k] = { e: 0, c: 0 }); x.e++; x.c += n; }
  }
  const opt = sel => d.dims.map(k => `<option value="${k}" ${k === sel ? 'selected' : ''}>${DIM_LABEL[k]}</option>`).join('');
  return `<div class="card hm">
    <div class="kicker" style="font-size:11px">Where groups overlap</div><h2>Coverage across intersections</h2>
    <div class="sub">Fair share for each combination. Hover a cell for its numbers.</div>
    <div style="margin:12px 0"><select id="hmRow">${opt(state.hmRow)}</select>×<select id="hmCol" style="margin-left:6px">${opt(state.hmCol)}</select></div>
    <table><thead><tr><th></th>${Cg.map(c => `<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>
    ${R.map(r => `<tr><th class="rh" title="${esc(r)}">${esc(r)}</th>${Cg.map(c => {
      const x = cell[r + '\u0000' + c];
      if (!x || x.e < MIN_CELL) return '<td class="blank" title="Fewer than 10 electors">·</td>';
      const fs = C ? (x.c / C) / (x.e / N) : null; const [bg, dark] = fsColour(fs);
      return `<td style="background:${bg};color:${dark ? '#fff' : 'var(--ink)'}" title="${esc(r)} × ${esc(c)}: ${fmt(x.e)} electors, ${fmt(x.c)} contacts">${fs == null ? '–' : fs.toFixed(1)}</td>`;
    }).join('')}</tr>`).join('')}
    </tbody></table>
    <div class="note" style="margin-top:10px">Cells under ${MIN_CELL} electors are blanked so small groups can't be singled out.</div>
  </div>`;
}

function renderMoves(allGroups, A) {
  const d = state.data; const per = d.mode === 'dated' ? PERIODS[state.period].long.toLowerCase() : 'on record';
  const minE = Math.max(50, Math.round(d.people.length * 0.02));
  const pool = allGroups.filter(g => g.e >= minE && g.name !== UNTAGGED && g.name !== 'Unknown' && g.dim !== 'reach');
  const moves = []; const used = new Set();
  const lr = pool.filter(g => g.lowReach).sort((a, b) => b.c - a.c)[0];
  if (lr) { moves.push({ g: lr, kind: 'Low reach', icon: 'people', title: `${lr.name}: lots of contact, but only ${Math.round(lr.reachPct * 100)}% of people reached. Widen the net.`, body: `${fsTxt(lr.fs)} fair share of contacts ${per}, landing on the same ${fmt(lr.r)} people. ${fmt(lr.e - lr.r)} electors haven't heard from the office.` }); used.add(lr.dim); }
  const under = pool.filter(g => g.status === 'u').sort((a, b) => (b.e * (1 - b.fs)) - (a.e * (1 - a.fs)));
  for (const g of under) {
    if (moves.length >= 3) break; if (used.has(g.dim)) continue; used.add(g.dim);
    const emailShare = g.e ? g.email / g.e : 0, noDet = g.e ? g.noDetails / g.e : 0;
    const how = !d.hasReach ? ['Outreach', 'Plan outreach'] : noDet > 0.4 ? ['Doorknock', 'Doorknock'] : emailShare >= 0.5 ? ['Email campaign', 'Email campaign'] : ['Phone bank', 'Phone bank'];
    const offPct = g.office + g.constituent ? Math.round(g.constituent / (g.office + g.constituent) * 100) : null;
    moves.push({ g, kind: how[0], icon: how[0] === 'Doorknock' ? 'pin' : how[0] === 'Email campaign' ? 'mail' : 'phone',
      title: `${how[1]}: ${g.name} is getting ${fsTxt(g.fs)} its fair share.`,
      body: `${fmt(g.e)} electors, ${Math.round(g.reachPct * 100)}% reached ${per}.${g.lastOff != null ? ` Last office-initiated contact ${d.asOfDay - g.lastOff ? (d.asOfDay - g.lastOff) + (d.asOfDay - g.lastOff === 1 ? ' day' : ' days') + ' before the latest record' : 'on the latest record date'}` : ' No office-initiated contact in this period'}${offPct != null ? `; ${offPct}% of recent contact was them coming to us.` : '.'}${d.hasReach ? ` ${Math.round(emailShare * 100)}% have an email address on file.` : ''}` });
  }
  if (!moves.length) return '';
  const ICON = {
    people: '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
    pin: '<path d="M20 10c0 6-8 12-8 12S4 16 4 10a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/>',
    mail: '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-10 6L2 7"/>',
    phone: '<path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1.9.4 1.8.7 2.7a2 2 0 0 1-.5 2.1L8 9.8a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.7.7a2 2 0 0 1 1.7 2z"/>',
  };
  return `<div style="margin-top:26px"><div class="kicker" style="font-size:11px">Close the gap</div><h2 style="font-size:19px;margin:2px 0 2px">Recommended next moves</h2>
  <div class="sub" style="color:var(--muted);font-size:13px">Where the next doorknock, phone bank or email campaign should go. Download the list, then pull it in Campaign Central.</div>
  <div class="moves">${moves.map(m => `<div class="card move">
    <div class="tag"><span class="ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">${ICON[m.icon]}</svg></span>${esc(m.kind)} · ${esc(DIM_LABEL[m.g.dim])}</div>
    <h3>${esc(m.title)}</h3><p>${esc(m.body)}</p>
    <div class="acts"><button class="btn primary" data-target="${esc(m.g.dim)}|${esc(m.g.name)}" ${d.hasIds ? '' : 'disabled title="This file has no ID column"'}><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>Target list</button><button class="btn" data-jump="${esc(m.g.dim)}|${esc(m.g.name)}">Inspect</button></div>
  </div>`).join('')}</div></div>`;
}

function bindDashboard() {
  const main = $('#main');
  main.querySelectorAll('[data-period]').forEach(b => b.onclick = () => { state.period = +b.dataset.period; render(); });
  main.querySelectorAll('[data-dim]').forEach(b => b.onclick = () => { state.dim = b.dataset.dim; state.sel = null; render(); });
  main.querySelectorAll('[data-group]').forEach(b => b.onclick = () => { state.sel = b.dataset.group; render(); });
  main.querySelectorAll('[data-jump]').forEach(b => b.onclick = () => { const [dim, name] = b.dataset.jump.split('|'); state.dim = dim; state.sel = name; render(); $('#tm').scrollIntoView({ behavior: 'smooth', block: 'center' }); });
  main.querySelectorAll('[data-target]').forEach(b => b.onclick = () => { const i = b.dataset.target.indexOf('|'); downloadTargets(b.dataset.target.slice(0, i), b.dataset.target.slice(i + 1)); });
  $('#hmRow').onchange = e => { state.hmRow = e.target.value; render(); };
  $('#hmCol').onchange = e => { state.hmCol = e.target.value; render(); };
  $('#newImport').onclick = () => $('#fileInput').click();
  $('#closeFile').onclick = () => { state.data = null; state.sel = null; render(); toast('File cleared from this tab.'); };
}

/* Target list: Campaign Central IDs only, for people in the group not reached in the period. */
function downloadTargets(dim, name) {
  const d = state.data; if (!d.hasIds) return;
  const w = periodWindow();
  const ids = d.people.filter(p => p.id && p.attrs[dim].includes(name) && !p.contacts.some(c => inWin(c, w))).map(p => p.id);
  const blob = new Blob(['ID\r\n' + ids.map(i => /[",\r\n]/.test(i) ? `"${i.replace(/"/g, '""')}"` : i).join('\r\n') + '\r\n'], { type: 'text/csv' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = `target-list-${DIM_LABEL[dim].toLowerCase().replace(/\s+/g, '-')}-${name.toLowerCase().replace(/[^a-z0-9]+/g, '-')}.csv`;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  toast(`${fmt(ids.length)} IDs downloaded. Pull them in Campaign Central to build the list.`);
}

/* ------------------------------------------------------------------ import flow */
function modal(html) { $('#modalRoot').innerHTML = `<div class="modal-bg"><div class="modal">${html}</div></div>`; const x = $('#modalRoot .x'); if (x) x.onclick = closeModal; }
function closeModal() { $('#modalRoot').innerHTML = ''; }
let pending = null;

async function startImport(file, opts = {}) {
  modal(`<header><div><div class="kicker">New import · step 1 of 3</div><h2>Reading ${esc(file.name)}</h2></div><button class="x">×</button></header>
    <div class="body"><p class="note" id="pgTxt">Reading the file in this tab. Nothing is uploaded.</p><div class="prog"><i id="pg" style="width:15%"></i></div><p class="note" style="margin-bottom:0">A large spreadsheet can take 10-20 seconds.</p></div>`);
  try {
    const data = await readAndGate(file);
    data.officeName = opts.officeName || state.profile.officeName || 'Your office';
    pending = data;
    showGate();
  } catch (e) {
    modal(`<header><div><div class="kicker">New import</div><h2>That file couldn't be read</h2></div><button class="x">×</button></header>
      <div class="body"><p>${esc(e.message || e)}</p><p class="note">Try saving it as CSV or .xlsx from Campaign Central or Excel.</p></div>`);
  }
}

/* Reading a large spreadsheet takes several seconds, so it happens in a
   background worker and the page stays responsive. If workers aren't
   available, it falls back to reading here. */
async function readAndGate(file) {
  const stage = (txt, w) => { const el = $('#pgTxt'); if (el) el.textContent = txt; const b = $('#pg'); if (b) b.style.width = w; };
  if (typeof Worker !== 'undefined') {
    try {
      const buf = await file.arrayBuffer();
      return await new Promise((resolve, reject) => {
        const w = new Worker('worker.js');
        w.onmessage = e => {
          const m = e.data;
          if (m.stage) stage(m.stage, m.w);
          else { w.terminate(); m.error ? reject(new Error(m.error)) : resolve(m.data); }
        };
        w.onerror = e => { w.terminate(); reject(new Error('worker-failed')); e.preventDefault(); };
        w.postMessage({ buf, name: file.name, profile: state.profile }, [buf]);
      });
    } catch (e) { if (e.message !== 'worker-failed') throw e; }
  }
  stage('Reading the file…', '40%'); await new Promise(r => setTimeout(r, 30));
  const rows = await readFile(file);
  stage('Running the import gate…', '75%'); await new Promise(r => setTimeout(r, 30));
  return runGate(rows, file.name);
}

function collapse(items, key) {
  // "TOPIC (1)", "TOPIC (2)"... shown once as "TOPIC (1-6)"
  const groups = new Map(); const out = [];
  for (const it of items) {
    const m = it.col.match(/^(.*?)(\d+)(\)?)\s*$/);
    if (m && /topic|tag|interaction|note/i.test(m[1])) { const k = m[1] + '#' + m[3] + '#' + (it[key] || ''); const g = groups.get(k); if (g) g.nums.push(+m[2]); else { const ng = { ...it, base: m[1], tail: m[3], nums: [+m[2]] }; groups.set(k, ng); out.push(ng); } }
    else out.push(it);
  }
  return out.map(it => it.nums ? { ...it, col: `${it.base}${it.nums.length > 1 ? Math.min(...it.nums) + '-' + Math.max(...it.nums) : it.nums[0]}${it.tail}` } : it);
}
function showGate() {
  const d = pending, r = d.report;
  const bad = r.warnings.filter(w => w.lvl === 'bad');
  const g = (cls, icon, title, body) => `<div class="g ${cls}"><span class="i">${icon}</span><div><b>${title}</b><span>${body}</span></div></div>`;
  modal(`<header><div><div class="kicker">New import · step 2 of 3</div><h2>Import gate</h2></div><button class="x">×</button></header>
  <div class="body"><div class="gate">
    ${g('info', 'i', `${esc(r.fileName)} · ${fmt(r.rowCount)} rows · ${fmt(d.people.length)} people`, `${d.mode === 'dated' ? `Dated mode: ${fmt(d.datedContacts)} dated contacts, so trends come from this file alone.` : 'Undated mode: this file has no contact dates, so the dashboard shows all contact on record.'}`)}
    ${bad.map(w => g('bad', '!', esc(w.title), esc(w.body))).join('')}
    ${g('ok', '✓', 'Kept for counting', `<div class="cols">${collapse(r.kept, 'as').map(k => `<code title="${esc(k.as)}">${esc(k.col)}</code>`).join('')}</div>`)}
    ${r.converted.length ? g('ok', '✓', 'Converted, original values discarded', `<div class="cols">${collapse(r.converted, 'to').map(k => `<code title="${esc(k.to)}">${esc(k.col)} → ${esc(k.to.split(' (')[0])}</code>`).join('')}</div>`) : ''}
    ${r.strippedNames.length + r.strippedAddress.length + r.dropped.length + r.scanDropped.length ? g('info', '–', 'Dropped, never read into the dashboard', `<div class="cols">${[...r.strippedNames, ...r.strippedAddress, ...r.scanDropped.map(s => s.col)].map(c => `<code class="s">${esc(c)}</code>`).join('')}${r.dropped.map(c => `<code>${esc(c)}</code>`).join('')}</div>`) : ''}
    ${r.warnings.filter(w => w.lvl !== 'bad').map(w => g(w.lvl === 'warn' ? 'warn' : 'info', w.lvl === 'warn' ? '!' : 'i', esc(w.title), esc(w.body))).join('')}
  </div>
  <p class="note" style="margin-top:14px">Rows stay in this tab's memory and disappear when you close the file or the tab. Nothing is saved or sent.</p></div>
  <footer><span class="note">${d.unclassified.length ? `${d.unclassified.length} tag value${d.unclassified.length > 1 ? 's need' : ' needs'} classifying next` : 'All tag values recognised'}</span>
  <div style="display:flex;gap:8px"><button class="btn" id="gCancel">Cancel</button><button class="btn primary" id="gNext">${d.unclassified.length ? 'Next' : 'Show dashboard'}</button></div></footer>`);
  $('#gCancel').onclick = () => { pending = null; closeModal(); };
  $('#gNext').onclick = () => d.unclassified.length ? showClassify(false) : finishImport();
}

function showClassify(fromNav) {
  const d = fromNav ? state.data : pending;
  const list = d ? d.unclassified : [];
  modal(`<header><div><div class="kicker">${fromNav ? 'Office profile' : 'New import · step 3 of 3'}</div><h2>Unclassified tags</h2></div><button class="x">×</button></header>
  <div class="body">
    ${list.length ? `<p class="note" style="margin-top:0">These tag values aren't in office profile v${state.profile.version} yet. Tell us what each one means; the answer is saved to the profile on this computer for future imports. Only short tag values are listed; free text is never shown.</p>
    ${list.map((u, i) => `<div class="uq"><div><b>${esc(u.value)}</b><span>found in ${esc(u.cols.join(', '))} · ${fmt(u.people)} electors</span></div>
      <select data-uq="${i}"><option value="">Leave unclassified</option>${CLASSES.map(([k, l]) => `<option value="${k}">${l}</option>`).join('')}</select>
      <input data-uqv="${i}" placeholder="Group name" class="hidden" value="${esc(u.value)}"></div>`).join('')}` : '<p>Nothing to classify. Every tag value in the open file is recognised.</p>'}
  </div>
  <footer><span class="note">Saved answers contain no personal information.</span><div style="display:flex;gap:8px">${fromNav ? '' : '<button class="btn" id="cSkip">Skip for now</button>'}<button class="btn primary" id="cSave">${fromNav ? 'Save' : 'Save and show dashboard'}</button></div></footer>`);
  document.querySelectorAll('[data-uq]').forEach(s => s.onchange = () => { const inp = document.querySelector(`[data-uqv="${s.dataset.uq}"]`); inp.classList.toggle('hidden', !['community', 'religion', 'housing', 'issue'].includes(s.value)); });
  const skip = $('#cSkip'); if (skip) skip.onclick = finishImport;
  $('#cSave').onclick = () => {
    let n = 0;
    document.querySelectorAll('[data-uq]').forEach(s => {
      if (!s.value) return; const u = list[+s.dataset.uq];
      const v = document.querySelector(`[data-uqv="${s.dataset.uq}"]`).value.trim().slice(0, 60);
      state.profile.values[u.value] = { cls: s.value, value: v || u.value }; n++;
    });
    if (n) { state.profile.version++; const ok = saveProfile(); toast(ok ? `Office profile updated to v${state.profile.version}.` : `Profile updated for this session (this browser blocks saving).`); }
    if (fromNav) { if (n && state.data) reprocess(); closeModal(); render(); }
    else { if (n) reprocessPending(); finishImport(); }
  };
}

/* The profile changed, so re-count from the cleaned copy held in memory. The
   original gate report (what was stripped from the real file) is kept. */
function rerun(old) {
  const fresh = runGate(old.sanitized, old.fileName);
  return { ...fresh, report: old.report, officeName: old.officeName, importedAt: old.importedAt };
}
function reprocessPending() { if (pending) pending = rerun(pending); }
function reprocess() { if (state.data) state.data = rerun(state.data); }

function finishImport() {
  state.data = pending; pending = null; state.sel = null;
  if (state.data.mode !== 'dated') state.period = PERIODS.length - 1;
  closeModal(); render(); window.scrollTo(0, 0);
}

async function loadSample() {
  try {
    const res = await fetch('sample/demo-electorate-export.csv', { cache: 'no-store' });
    if (!res.ok) throw new Error('Sample file not found');
    const file = new File([await res.blob()], 'demo-electorate-export.csv', { type: 'text/csv' });
    startImport(file, { officeName: 'Riverbend (sample)' });
  } catch (e) { toast('Could not load the sample file.'); }
}

/* A fictional file that deliberately breaks the rules, to show the gate working. */
function makePoisonFile() {
  const first = ['Olivia', 'Liam', 'Mia', 'Noah', 'Ava', 'Jack', 'Zara', 'Omar', 'Mei', 'Arjun'];
  const last = ['Testperson', 'Samplerson', 'Fakely', 'Madeup', 'Notreal'];
  const subs = ['Riverbend', 'Kellbrook', 'Ashdale', 'Fennvale', 'Bellwood'];
  const lines = ['First Name,Last Name,DOB,Street,Email,Mobile,Gender,Ethnicity,Suburb,Postcode,TOPIC (1),Tag (1),Interaction Details Set 1,Notes Field'];
  for (let i = 0; i < 400; i++) {
    const f = first[i % 10], l = last[i % 5]; const y = 1940 + (i * 7) % 66; const m = 1 + i % 12, dd = 1 + i % 28;
    const d = new Date(Date.UTC(2026, 9, 28) - (i * 13 % 300) * DAY);
    const note = i % 3 ? `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()} - ${i % 2 ? 'IE' : 'OC'} ${f} ${l} called about a private matter at ${10 + i} Example St` : '';
    lines.push([f, l, `${String(dd).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`, `${10 + i} Example St`, `${f.toLowerCase()}.${l.toLowerCase()}@example.com`, `04${String(10000000 + i * 7919).slice(0, 8)}`, i % 2 ? 'F' : 'M', ['Australian', 'Chinese', 'Indian', 'Lebanese'][i % 4], subs[i % 5], '2991', i % 3 ? 'PFN: COMMUNITY INTERACTION' : '', '', `"${note}"`, `"${f} ${l} prefers calls after 5pm"`].join(','));
  }
  return new File([lines.join('\n')], 'privacy-gate-test (fictional names).csv', { type: 'text/csv' });
}

/* ------------------------------------------------------------------ office profile view */
function showProfile() {
  const vals = Object.entries(state.profile.values);
  modal(`<header><div><div class="kicker">Office profile v${state.profile.version}</div><h2>How tag values are read</h2></div><button class="x">×</button></header>
  <div class="body">
    <p class="note" style="margin-top:0">The profile holds rules only. It contains no names, records or counts, and it is the only thing this page saves on your computer.</p>
    <div class="sec">Office name</div>
    <input id="pName" maxlength="60" placeholder="e.g. Parramatta" value="${esc(state.profile.officeName || '')}" style="border:1px solid var(--line);border-radius:8px;padding:7px 10px;width:100%">
    <div class="note" style="margin-top:4px">Shown on the dashboard. Saved on this computer only.</div>
    <div class="sec">Built-in rules</div>
    <table class="plain"><tbody>${DEFAULT_RULES.map(r => `<tr><td>${esc(r.note)}</td><td class="mono" style="font-size:11px;color:var(--muted)">${esc(r.re)}</td></tr>`).join('')}</tbody></table>
    <div class="sec">Channel codes read from interaction notes</div>
    <div class="cols">${Object.entries(CHANNEL_CODES).map(([k, [ch, dir]]) => `<code>${esc(k)} = ${esc(ch)}, ${dir === 'in' ? 'constituent' : 'office'}</code>`).join('')}</div>
    <div class="sec">This office's answers</div>
    ${vals.length ? `<table class="plain"><thead><tr><th>Tag value</th><th>Means</th></tr></thead><tbody>${vals.map(([k, v]) => `<tr><td>${esc(k)}</td><td>${esc(CLASSES.find(c => c[0] === v.cls)?.[1] || v.cls)}${['community', 'religion', 'housing', 'issue'].includes(v.cls) ? ': ' + esc(v.value) : ''}</td></tr>`).join('')}</tbody></table>` : '<p class="note">None yet. Answers from the "Unclassified tags" step appear here.</p>'}
  </div>
  <footer><button class="btn" id="pReset" ${vals.length ? '' : 'disabled'}>Reset to built-in rules</button><button class="btn primary" id="pClose">Done</button></footer>`);
  $('#pClose').onclick = () => {
    const nm = $('#pName').value.trim().slice(0, 60);
    if (nm !== (state.profile.officeName || '')) { state.profile.officeName = nm; saveProfile(); if (state.data) state.data.officeName = nm || 'Your office'; render(); }
    closeModal();
  };
  $('#pReset').onclick = () => { state.profile = { version: state.profile.version + 1, officeName: state.profile.officeName, values: {} }; saveProfile(); if (state.data) reprocess(); closeModal(); render(); toast('Office profile reset.'); };
}

/* ------------------------------------------------------------------ wiring */
if (typeof document !== 'undefined') {
$('#fileInput').onchange = e => { const f = e.target.files[0]; e.target.value = ''; if (f) startImport(f); };
document.querySelectorAll('[data-nav]').forEach(b => b.onclick = () => {
  const n = b.dataset.nav;
  if (n === 'dashboard') { closeModal(); window.scrollTo(0, 0); }
  if (n === 'import') $('#fileInput').click();
  if (n === 'profile') showProfile();
  if (n === 'unclassified') showClassify(true);
});
render();
}
