/** 
 * Quote Tool — Apps Script backend (CSV-aware)
 * - Uses your existing Quotes sheet headers from the provided CSV
 * - Append-only saves (audit log)
 * - Versions in Column B: FIRST4 + DDMMYYYY + #N
 * - Self-led activities written to the 'Self Led Activities' (or 'Self Led Activites') column
 * - Also writes Group -> 🪄!H1 and Version -> 🪄!H2 on save and on load
 */
const CFG = {
  accomSheet: 'Accom',
  accomProviderRange: 'A3:A',
  accomBookingRange:  'B3:B',
  accomTypeRange:     'C3:C',

  activitiesSheet: 'Activities',
  activitiesRange: 'A3:A',

  settingsSheet: 'Settings',
  timeRange: 'A4:A',

  quotesSheet: 'Quotes',

  fmtCurrency: '£#,##0.00',
  fmtPercent: '0.00%'.toString(),
  fmtDate: 'dd/mm/yyyy',
};

// === Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Quote Tool')
    .addItem('Open Sidebar', 'showGroupSidebar')
    .addToUi();
}
function showGroupSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('QuoteSideBar').setTitle('Quote Tool');
  SpreadsheetApp.getUi().showSidebar(html);
}

// === Utilities
function readColumn_(sheetName, a1) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return [];
  const rng = sh.getRange(a1);
  const values = rng.getValues().flat();
  const seen = new Set();
  const out = [];
  for (const v of values) {
    const s = (v === null || v === undefined) ? '' : String(v).trim();
    if (!s || seen.has(s)) continue;
    seen.add(s); out.push(v);
  }
  return out;
}
function toHM_(d) { const hh = String(d.getHours()).padStart(2,'0'); const mm = String(d.getMinutes()).padStart(2,'0'); return `${hh}:${mm}`; }
function fmtTimeString_(v){
  if (v instanceof Date) return toHM_(v);
  const s = (v === null || v === undefined) ? '' : String(v).trim();
  if (!s) return '';
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return `${m[1].padStart(2,'0')}:${m[2]}`;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return toHM_(d);
  return s;
}
function toISO_(d){ if (!(d instanceof Date)) return ''; const y=d.getFullYear(); const m=String(d.getMonth()+1).padStart(2,'0'); const day=String(d.getDate()).padStart(2,'0'); return `${y}-${m}-${day}`; }
function toDate_(iso){ if (!iso) return ''; const p = String(iso).split('-').map(Number); if (p.length!==3) return ''; return new Date(p[0], p[1]-1, p[2]); }
function parseMoney_(s) { if (s===null||s===undefined) return ''; const n=parseFloat(String(s).replace(/[^0-9.]/g,'')); return isNaN(n)?'':n; }
function parsePercent_(s) { if (s===null||s===undefined) return ''; const n=parseFloat(String(s).replace(/[^0-9.\-]/g,'')); if (isNaN(n)) return ''; return n>1?n/100:n; }

function getQuotesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.quotesSheet);
  if (!sh) {
    sh = ss.insertSheet(CFG.quotesSheet);
    sh.getRange(1,1,1,27).setValues([quoteHeadersFromCsv_()]);
  }
  return sh;
}
function quoteHeadersFromCsv_() { return ["Group Name", "Version ", "Participants", "Leaders", "Accommodation Provider", "Booking Method", "Type", "Board", "Activity Transport", "Return Travel", "Other Charges", "Other Charges Description", "Admin Charge", "Free Places", "Arrival", "Departure", "TOMS Check", "Discount %", "Discount £", "Sub Groups", "Charge Type", "Activities", "Self Led Activites", "Start Time", "Show Meals?", "Arrival Time", "Departure Time"]; }
function headerIndexMap_(sh){
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h,i)=>{ map[(h||'').toString().trim().toLowerCase()] = i+1; });
  return map;
}

// === Version helpers
function versionBase_(groupName, arrivalISO){
  const first4 = (groupName||'').replace(/[^A-Za-z]/g,'').toUpperCase().slice(0,4);
  const d = toDate_(arrivalISO);
  const dd = String(d.getDate()).padStart(2,'0');
  const mm = String(d.getMonth()+1).padStart(2,'0');
  const yyyy = d.getFullYear();
  return `${first4}${dd}${mm}${yyyy}`;
}
function nextVersionNumber_(base){
  const sh = getQuotesSheet_();
  const rows = Math.max(0, sh.getLastRow()-1);
  if (!rows) return 1;
  const vers = sh.getRange(2, 2, rows, 1).getValues().flat(); // Column B
  let maxN = 0;
  vers.forEach(v=>{ const s=(v||'').toString(); if(s.startsWith(base+'#')){ const m=s.match(/#(\d+)$/); if(m) maxN=Math.max(maxN,parseInt(m[1],10)); } });
  return maxN+1;
}

function updateMagic_(groupName, version){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('🪄');
  if (!sh) sh = ss.insertSheet('🪄');
  sh.getRange('H1').setValue(groupName||'');
  sh.getRange('H2').setValue(version||'');
}

// === Public API for Sidebar
function getSidebarInit(forceLists) {
  const providers = readColumn_(''+CFG.accomSheet, CFG.accomProviderRange).map(v=>String(v).trim());
  const bookingMethods = readColumn_(''+CFG.accomSheet, CFG.accomBookingRange).map(v=>String(v).trim());
  const types = readColumn_(''+CFG.accomSheet, CFG.accomTypeRange).map(v=>String(v).trim());
  const activities = readColumn_('' + CFG.activitiesSheet, CFG.activitiesRange).map(v => String(v).trim()).sort((a, b) => a.localeCompare(b));
  const timesRaw = readColumn_(''+CFG.settingsSheet, CFG.timeRange);
  const times = Array.from(new Set(timesRaw.map(fmtTimeString_).filter(Boolean)));

  const groupNames = getAllGroupNames_();
  const boardOptions = ['Room Only','None','Full Board','Half Board','Bed & Breakfast'];
  const lists = { providers, bookingMethods, types, activities, times, boardOptions, chargeTypes: ['Peak','Off Peak'], groupNames };

  const defaults = {
    groupName: '',
    participants: '',
    leaders: '',
    freePlaces: '',
    subGroups: 1,
    accommodationProvider: 'Celtic Camping',
    bookingMethod: 'TYF to Book',
    type: 'Bunkhouse',
    board: 'Full Board',
    activityTransport: '',
    returnTravel: '',
    otherCharges: '',
    otherChargesDesc: '',
    adminChargePctDisplay: '7%',
    discountPctDisplay: '',
    discountGBP: '',
    chargeType: 'Peak',
    breakfastTime: '08:00',
    arrival: '2026-05-01',
    arrivalTime: '12:30',
    departure: '2026-05-01',
    departureTime: '14:00',
    activities: [],
    selfLedActivities: []
  };
  return { lists, defaults };
}

function getAllGroupNames_(){
  const sh = getQuotesSheet_();
  const map = headerIndexMap_(sh);
  const col = map['group name'] || 1;
  const rows = Math.max(0, sh.getLastRow()-1);
  if (!rows) return [];
  const vals = sh.getRange(2, col, rows, 1).getValues().flat();
  const set = new Set();
  vals.forEach(v=>{ const s=(v||'').toString().trim(); if(s) set.add(s); });
  return Array.from(set).sort((a,b)=>a.localeCompare(b));
}

function getVersionsForGroup(groupName){
  const sh = getQuotesSheet_();
  const map = headerIndexMap_(sh);
  const colGroup = map['group name'] || 1;
  const colTs = 1; // ✅ Column A is Timestamp
  const rows = Math.max(0, sh.getLastRow()-1);
  if (!rows) return [];
  const data = sh.getRange(2, 1, rows, sh.getLastColumn()).getValues();

  const out = [];
  const wanted = String(groupName).trim();
  data.forEach(r=>{
    const g = (r[colGroup-1]||'').toString().trim();
    if (g === wanted) {
      const ts = r[colTs-1];
      if (ts instanceof Date) out.push(ts.toISOString());       // store ISO value
      else if (ts) out.push(String(ts));                        // fallback
    }
  });
  out.sort((a,b)=> new Date(b) - new Date(a)); // newest first
  return out;
}



function loadQuoteByVersion(timestampISO){
  const sh = getQuotesSheet_();
  const map = headerIndexMap_(sh);
  const colTs = 1; // ✅ Column A is Timestamp
  const rows = Math.max(0, sh.getLastRow()-1);
  if (!rows) return { ok:false, message:'No data' };

  const data = sh.getRange(2, 1, rows, sh.getLastColumn()).getValues();
  let row = null;

  const wanted = String(timestampISO).trim();
  for (const r of data){
    const ts = r[colTs-1];
    const tsIso = ts instanceof Date ? ts.toISOString() : String(ts).trim();
    if (tsIso === wanted) { row = r; break; }
  }
  if(!row) return { ok:false, message:'Timestamp not found' };

  const get = (name) => row[(map[name]-1)];
  const has = (name) => Object.prototype.hasOwnProperty.call(map, name);
  const getAny = (names) => { for (const n of names) if (has(n)) return get(n); return ''; };

  const v = { };
  v.groupName = has('group name') ? (get('group name')||'') : '';
  v.participants = has('participants') ? Number(get('participants')||0) : 0;
  v.leaders = has('leaders') ? Number(get('leaders')||0) : 0;
  v.freePlaces = has('free places') ? Number(get('free places')||0) : 0;
  v.subGroups = has('sub groups') ? Number(get('sub groups')||0) : 1;

  v.accommodationProvider = has('accommodation provider') ? (get('accommodation provider')||'') : '';
  v.bookingMethod = has('booking method') ? (get('booking method')||'') : '';
  v.type = has('type') ? (get('type')||'') : '';
  v.board = has('board') ? (get('board') || '') : '';

  v.activityTransport = has('activity transport') ? Number(get('activity transport')||0) : 0;
  v.returnTravel = has('return travel') ? (get('return travel')||'') : '';

  v.otherCharges = has('other charges') ? (get('other charges')||'') : '';
  v.otherChargesDesc = has('other charges description') ? (get('other charges description')||'') : '';
  v.adminChargePct = has('admin charge') ? (get('admin charge')||'') : '';
  v.discountPct = has('discount %') ? (get('discount %')||'') : '';
  v.discountGBP = has('discount £') ? (get('discount £')||'') : '';
  v.chargeType = has('charge type') ? (get('charge type')||'') : '';

  const arrDate = has('arrival') ? get('arrival') : '';
  v.arrival = arrDate instanceof Date ? toISO_(arrDate) : (arrDate||'');
  v.arrivalTime = has('arrival time') ? fmtTimeString_(get('arrival time')) : '';

  const depDate = has('departure') ? get('departure') : '';
  v.departure = depDate instanceof Date ? toISO_(depDate) : (depDate||'');
  v.departureTime = has('departure time') ? fmtTimeString_(get('departure time')) : '';

  v.breakfastTime = has('breakfast time') ? fmtTimeString_(get('breakfast time')) : '';

  const acts = (has('activities') ? (get('activities')||'') : '').toString().split(',').map(s=>s.trim()).filter(Boolean);
  const sl = (getAny(['self led activities','self led activites'])||'').toString().split(',').map(s=>s.trim()).filter(Boolean);
  const selfLedSet = new Set(sl);
  const items = acts.map(name => ({ name, selfLed: selfLedSet.has(name) }));

  updateMagic_(v.groupName, wanted); // H2 holds the ISO timestamp
  return { ok:true, values:v, items };
}

function saveGroupData(payload){
  try {
    const p = payload||{}; const v = p.values||{};
    const sh = getQuotesSheet_();
    const map = headerIndexMap_(sh);
    const width = sh.getLastColumn();
    const row = new Array(width).fill('');

    const base = versionBase_(v.groupName, v.arrival);
    const vn = nextVersionNumber_(base);
    const version = `${base}#${vn}`;

    const combo = Array.isArray(v.activitiesCombined) ? v.activitiesCombined : [];
    const allActs = combo.map(o => (o && o.name) ? o.name : (typeof o === 'string' ? o : '')).filter(Boolean);
    const selfLedActs = combo.filter(o => o && o.selfLed).map(o => o.name).filter(Boolean);

    function put(name, value){ if (Object.prototype.hasOwnProperty.call(map,name)) row[map[name]-1] = value; }
    function putAny(names, value){ for (const n of names) if (Object.prototype.hasOwnProperty.call(map,n)) { row[map[n]-1] = value; return; } }

    put('timestamp', new Date());

    put('group name', v.groupName||'');
    put('participants', Number(v.participants)||0);
    put('leaders', Number(v.leaders)||0);
    put('free places', Number(v.freePlaces)||0);
    put('sub groups', Number(v.subGroups)||0);

    put('accommodation provider', v.accommodationProvider||'');
    put('booking method', v.bookingMethod||'');
    put('type', v.type||'');
    put('board', v.board || '');


    put('activity transport', Number(v.activityTransport)||0);
    put('return travel', parseMoney_(v.returnTravel));

    put('other charges', parseMoney_(v.otherCharges));
    put('other charges description', v.otherChargesDesc||'');
    put('admin charge', parsePercent_(v.adminChargePct));
    put('discount %', parsePercent_(v.discountPct));
    put('discount £', parseMoney_(v.discountGBP));
    put('charge type', v.chargeType||'');

    put('breakfast time', v.breakfastTime||'');
    put('arrival', toDate_(v.arrival));
    put('arrival time', v.arrivalTime||'');

    putAny(['self led activities','self led activites'], selfLedActs.join(', '));

    put('departure', toDate_(v.departure));
    put('departure time', v.departureTime||'');

    put('activities', allActs.join(', '));

    sh.appendRow(row);

    function fmt(name, fmt){ if(Object.prototype.hasOwnProperty.call(map,name)) sh.getRange(sh.getLastRow(), map[name]).setNumberFormat(fmt); }
    fmt('return travel', CFG.fmtCurrency);
    fmt('other charges', CFG.fmtCurrency);
    fmt('discount £', CFG.fmtCurrency);
    fmt('admin charge', CFG.fmtPercent);
    fmt('discount %', CFG.fmtPercent);
    fmt('arrival', CFG.fmtDate);
    fmt('departure', CFG.fmtDate);

    updateMagic_(v.groupName, version);

    return { ok:true, version };
  } catch(err) {
    return { ok:false, message: err && err.message ? err.message : 'Unknown error' };
  }
}
