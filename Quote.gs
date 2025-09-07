/** 
 * Quote Tool — Apps Script backend (CSV-aware)
 * - Uses your existing Quotes sheet headers from the provided CSV
 * - Append-only saves (audit log)
 * - Versions in Column B: FIRST4 + DDMMYYYY + #N
 * - Self-led activities written to the 'Self Led Activities' (or 'Self Led Activites') column
 * - Also writes Group -> 🪄!L1 and Version -> 🪄!L2 on save and on load
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

function updateMagic_(groupName, when){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('🪄');
  if (!sh) sh = ss.insertSheet('🪄');

  // L1: group name (text)
  sh.getRange('L1').setValue(groupName || '');

  // L2: timestamp as a real Date, formatted for display
  const cell = sh.getRange('L2');
  const d = (when instanceof Date) ? when : new Date(when);
  if (isNaN(d.getTime())) {            // safety: bad date -> clear cell
    cell.setValue('');
    return;
  }
  cell.setValue(d);                    // write as Date
  cell.setNumberFormat('dd/mm/yyyy hh:mm:ss'); // display like your Quotes sheet
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
    selfLedActivities: [],
    showMeals: false
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
  const colTs = 1; // Column A = Timestamp
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
  const splitList = (s) => String(s ?? '')
    .split(',')
    .map(x => x.trim())
    .filter(x => x.length > 0);

  const v = {};
  v.groupName       = has('group name') ? (get('group name')||'') : '';
  v.participants    = has('participants') ? Number(get('participants')||0) : 0;
  v.leaders         = has('leaders') ? Number(get('leaders')||0) : 0;
  v.freePlaces      = has('free places') ? Number(get('free places')||0) : 0;
  v.subGroups       = has('sub groups') ? Number(get('sub groups')||0) : 1;

  v.accommodationProvider = has('accommodation provider') ? (get('accommodation provider')||'') : '';
  v.bookingMethod         = has('booking method') ? (get('booking method')||'') : '';
  v.type                  = has('type') ? (get('type')||'') : '';
  v.board                 = has('board') ? (get('board') || '') : '';

  v.activityTransport = has('activity transport') ? Number(get('activity transport')||0) : 0;
  v.returnTravel      = has('return travel') ? (get('return travel')||'') : '';

  v.otherCharges     = has('other charges') ? (get('other charges')||'') : '';
  v.otherChargesDesc = has('other charges description') ? (get('other charges description')||'') : '';
  v.adminChargePct   = has('admin charge') ? (get('admin charge')||'') : '';
  v.discountPct      = has('discount %') ? (get('discount %')||'') : '';
  v.discountGBP      = has('discount £') ? (get('discount £')||'') : '';
  v.chargeType       = has('charge type') ? (get('charge type')||'') : '';

  const arrDate   = has('arrival') ? get('arrival') : '';
  v.arrival       = arrDate instanceof Date ? toISO_(arrDate) : (arrDate||'');
  v.arrivalTime   = has('arrival time') ? fmtTimeString_(get('arrival time')) : '';

  const depDate   = has('departure') ? get('departure') : '';
  v.departure     = depDate instanceof Date ? toISO_(depDate) : (depDate||'');
  v.departureTime = has('departure time') ? fmtTimeString_(get('departure time')) : '';

  v.breakfastTime = has('breakfast time') ? fmtTimeString_(get('breakfast time')) : '';
  v.showMeals = has('show meals?') ? Boolean(get('show meals?')) : false;


  // --- Activities: build from BOTH columns, preserving duplicates ---
  const actsStr = has('activities') ? (get('activities') || '') : '';
  const slStr   = getAny(['self led activities','self led activites']) || '';

  const items = [
    ...splitList(actsStr).map(name => ({ name, selfLed: false })),
    ...splitList(slStr).map(name   => ({ name, selfLed: true  }))
  ];

  updateMagic_(v.groupName, wanted); // H2 holds the ISO timestamp
  return { ok:true, values:v, items };
}


function saveGroupData(payload){
  try {
    const p = payload||{}; 
    const v = p.values||{};
    const sh = getQuotesSheet_();
    const map = headerIndexMap_(sh);
    const width = sh.getLastColumn();
    const row = new Array(width).fill('');
    const ts = new Date();

    const combo = Array.isArray(v.activitiesCombined) ? v.activitiesCombined : [];

    // ✅ split by self-led; keep duplicates; preserve order in each list
    const actsNonSelf = combo.filter(o => o && !o.selfLed).map(o => o.name).filter(Boolean);
    const actsSelf    = combo.filter(o => o &&  o.selfLed).map(o => o.name).filter(Boolean);

    function put(name, value){ 
      if (Object.prototype.hasOwnProperty.call(map,name)) {
        row[map[name]-1] = value; 
        Logger.log('Put "%s" (col %s) = %s', name, map[name], value);
      }
    }
    function putAny(names, value){ 
      for (const n of names) {
        if (Object.prototype.hasOwnProperty.call(map,n)) { 
          row[map[n]-1] = value; 
          Logger.log('PutAny "%s" (col %s) = %s', n, map[n], value);
          return; 
        } 
      } 
    }

    put('timestamp', ts);
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

    // 👇 This is the key one
    put('other charges', parseMoney_(v.otherCharges));
    put('other charges description', v.otherChargesDesc||'');
    put('admin charge', parsePercent_(v.adminChargePct));
    put('discount %', parsePercent_(v.discountPct));
    put('discount £', parseMoney_(v.discountGBP));
    put('charge type', v.chargeType||'');

    put('breakfast time', v.breakfastTime||'');
    put('show meals?', !!v.showMeals);
    put('arrival', toDate_(v.arrival));
    put('arrival time', v.arrivalTime||'');

    put('departure', toDate_(v.departure));
    put('departure time', v.departureTime||'');

    put('activities', actsNonSelf.join(', '));
    putAny(['self led activities','self led activites'], actsSelf.join(', '));

    Logger.log('Row array about to append: %s', JSON.stringify(row));

    sh.appendRow(row);

    function fmt(name, fmt){ 
      if(Object.prototype.hasOwnProperty.call(map,name)) {
        sh.getRange(sh.getLastRow(), map[name]).setNumberFormat(fmt); 
        Logger.log('Formatted "%s" (col %s) as %s', name, map[name], fmt);
      }
    }
    fmt('return travel', CFG.fmtCurrency);
    fmt('other charges', CFG.fmtCurrency);
    fmt('discount £', CFG.fmtCurrency);
    fmt('admin charge', CFG.fmtPercent);
    fmt('discount %', CFG.fmtPercent);
    fmt('arrival', CFG.fmtDate);
    fmt('departure', CFG.fmtDate);

    Logger.log('=== SAVE DEBUG END ===');

    updateMagic_(v.groupName, ts);               // write a real Date to 🪄!H2

    return { ok:true, version: ts.toISOString() };// keep ISO for the UI chip
  } catch(err) {
    Logger.log('SaveGroupData ERROR: %s', err);
    return { ok:false, message: err && err.message ? err.message : 'Unknown error' };
  }
}

// === Convert current 🪄 + Info write
function colIndex_(letters) {
  const s = String(letters || '').toUpperCase().replace(/[^A-Z]/g,'');
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n || 1;
}

/**
 * Convert the latest 🪄 row to Logistics (B:J), and append a mapped row in Info.
 * @param {{values: object}} payload - same shape as saveGroupData input
 */
function convertQuote(payload) {
  try {
    const v = (payload && payload.values) || {};
    const ss = SpreadsheetApp.getActive();

    // Sheets
    const magic = ss.getSheetByName('🪄');
    if (!magic) return { ok: false, message: '🪄 sheet not found' };

    const logistics = ss.getSheetByName('Logistics') || ss.insertSheet('Logistics');
    const info = ss.getSheetByName('Info') || ss.insertSheet('Info');

    // ---------- Prepare rows from 🪄 A2:I (trim trailing blanks) ----------
    const lastRowAny = magic.getLastRow();
    if (lastRowAny < 2) return { ok: false, message: 'No data on 🪄 to convert' };

    const n = lastRowAny - 1; // rows from 2..last
    const block = magic.getRange(2, /*A*/1, n, /*A:I*/9).getValues();

    const isEmptyRow = (r) => r.every(c => c === '' || c === null);
    let lastNonEmpty = block.length - 1;
    while (lastNonEmpty >= 0 && isEmptyRow(block[lastNonEmpty])) lastNonEmpty--;
    if (lastNonEmpty < 0) return { ok:false, message: 'No A:I data to copy' };

    const rowsToCopy = block.slice(0, lastNonEmpty + 1).filter(r => !isEmptyRow(r));

    // ---------- Step 1: Append mapped row in Info FIRST ----------
    const rowInfo = info.getLastRow() + 1;
    const A  = colIndex_('A'),  D  = colIndex_('D'),  E  = colIndex_('E');
    const K  = colIndex_('K'),  L  = colIndex_('L'),  M  = colIndex_('M'), N = colIndex_('N');
    const U  = colIndex_('U'),  BV = colIndex_('BV'), BW = colIndex_('BW');
    const CA = colIndex_('CA'), CB = colIndex_('CB'), CC = colIndex_('CC');
    const AL = colIndex_('AL'), BB = colIndex_('BB'), AM = colIndex_('AM');

    // Base fields
    info.getRange(rowInfo, A ).setValue(v.groupName || '');
    info.getRange(rowInfo, D ).setValue(toDate_(v.arrival));
    info.getRange(rowInfo, E ).setValue(toDate_(v.departure));
    info.getRange(rowInfo, K ).setValue(v.accommodationProvider || '');
    info.getRange(rowInfo, L ).setValue(v.bookingMethod || '');
    info.getRange(rowInfo, M ).setValue(v.type || '');
    info.getRange(rowInfo, N ).setValue(v.board || '');
    info.getRange(rowInfo, U ).setValue(v.groupName || '');
    info.getRange(rowInfo, BV).setValue(parsePercent_(v.adminChargePct));
    info.getRange(rowInfo, BW).setValue(v.chargeType || '');
    info.getRange(rowInfo, CA).setValue(parseMoney_(v.discountGBP));
    info.getRange(rowInfo, CB).setValue(parsePercent_(v.discountPct));
    info.getRange(rowInfo, CC).setValue(Number(v.freePlaces) || 0);

    // Formats
    info.getRange(rowInfo, D ).setNumberFormat(CFG.fmtDate);
    info.getRange(rowInfo, E ).setNumberFormat(CFG.fmtDate);
    info.getRange(rowInfo, BV).setNumberFormat(CFG.fmtPercent);
    info.getRange(rowInfo, CB).setNumberFormat(CFG.fmtPercent);
    info.getRange(rowInfo, CA).setNumberFormat(CFG.fmtCurrency);

    // Transposed arrays from 🪄 (AL -> Info!AL, AM -> Info!BB)
    const rowsMagic = Math.max(0, magic.getLastRow() - 1);
    if (rowsMagic > 0) {
      const alVals = magic.getRange(/*row*/2, AL, rowsMagic, 1).getValues()
        .flat().filter(x => x !== '' && x !== null);
      if (alVals.length) info.getRange(rowInfo, AL, 1, alVals.length).setValues([alVals]);

      const amVals = magic.getRange(/*row*/2, AM, rowsMagic, 1).getValues()
        .flat().filter(x => x !== '' && x !== null);
      if (amVals.length) info.getRange(rowInfo, BB, 1, amVals.length).setValues([amVals]);
    }

    // ---------- Step 2: Copy 🪄 A2:I -> Logistics B:J with EXACT DV MATCH for time columns ----------
    const DataValidationCriteria = SpreadsheetApp.DataValidationCriteria;

    // Normalize a value to "HH:mm" string (no date part)
    function hhmm(val) {
      if (val === '' || val === null) return '';
      if (val instanceof Date) {
        const h = String(val.getHours()).padStart(2,'0');
        const m = String(val.getMinutes()).padStart(2,'0');
        return `${h}:${m}`;
      }
      if (typeof val === 'number') {
        const total = Math.round(val * 1440);
        const h = String(Math.floor(total / 60) % 24).padStart(2,'0');
        const m = String(total % 60).padStart(2,'0');
        return `${h}:${m}`;
      }
      const s = String(val).trim();
      const m = s.match(/^(\d{1,2}):(\d{2})/);
      return m ? `${m[1].padStart(2,'0')}:${m[2]}` : s;
    }

    // From a DV range, decide if it’s a TIME list, and build a map "HH:mm" -> exact allowed value
    function buildTimeMapFromDv(dv) {
      const out = { isTime: false, map: new Map() };
      if (!dv) return out;

      const type = dv.getCriteriaType();
      const vals = dv.getCriteriaValues();
      let rng = null, items = null;

      if (type === DataValidationCriteria.VALUE_IN_RANGE) {
        rng = vals && vals[0];
      } else if (type === DataValidationCriteria.VALUE_IN_LIST) {
        items = vals && vals[0];
      } else {
        return out;
      }

      if (rng) {
        const nums = rng.getValues();
        const disps = rng.getDisplayValues();
        let timeLike = 0, total = 0;
        for (let i = 0; i < nums.length; i++) {
          for (let j = 0; j < nums[i].length; j++) {
            const disp = (disps[i][j] || '').toString();
            const looksTime = /^\d{1,2}:\d{2}(?::\d{2})?$/.test(disp);
            const num = nums[i][j];
            if (disp !== '' && (looksTime || (num instanceof Date) || (typeof num === 'number' && num >= 0 && num < 1))) {
              total++;
              if (looksTime || (num instanceof Date) || (typeof num === 'number' && num >= 0 && num < 1)) timeLike++;
              if (disp !== '') {
                out.map.set(hhmm(disp), (num instanceof Date) ? num : num);
              }
            }
          }
        }
        out.isTime = timeLike > 0 && timeLike >= Math.max(3, Math.floor(total * 0.6));
        return out;
      }

      if (Array.isArray(items)) {
        let timeLike = 0, total = 0;
        items.forEach(s => {
          total++;
          const k = hhmm(s);
          if (/^\d{1,2}:\d{2}$/.test(k)) {
            timeLike++;
            const parts = k.split(':').map(Number);
            out.map.set(k, (parts[0] * 60 + parts[1]) / 1440);
          }
        });
        out.isTime = timeLike > 0 && timeLike >= Math.max(3, Math.floor(total * 0.6));
      }
      return out;
    }

    // Build per-column DV maps (for columns B..J -> j = 0..8)
    const dvInfoPerCol = [];
    for (let j = 0; j < 9; j++) {
      const cell = logistics.getRange(2, 2 + j);
      dvInfoPerCol[j] = buildTimeMapFromDv(cell.getDataValidation());
    }

    // Coerce each time-like column to the exact DV value
    const normalized = rowsToCopy.map(r => {
      const out = r.slice();
      for (let j = 0; j < 9; j++) {
        const dvInf = dvInfoPerCol[j];
        if (!dvInf.isTime) continue;           // only adjust DV time columns
        const key = hhmm(out[j]);
        if (!key) continue;
        if (dvInf.map.has(key)) {
          out[j] = dvInf.map.get(key);         // exact item from Settings range
        } else {
          // fallback to a clean numeric time fraction
          const m = key.match(/^(\d{1,2}):(\d{2})$/);
          if (m) out[j] = (Number(m[1]) * 60 + Number(m[2])) / 1440;
        }
      }
      return out;
    });

    // Paste
    const destRowLog = logistics.getLastRow() + 1;
    logistics.getRange(destRowLog, /*B*/2, normalized.length, /*B:J*/9).setValues(normalized);

    // >>> resize slicers to include the new rows <<<
    SpreadsheetApp.flush();
    try {
      resizeAllRangeSlicersAuto(); // or: resizeSlicersBatchTwoPhase('Logistics');
    } catch (e) {
      Logger.log('resizeAllRangeSlicersAuto failed: %s', e);
    }

    return {
      ok: true,
      message: `Converted: Logistics rows ${destRowLog}..${destRowLog + normalized.length - 1}, Info row ${rowInfo}`
    };
  } catch (err) {
    Logger.log('convertQuote ERROR: %s', err && err.stack || err);
    return { ok: false, message: err && err.message ? err.message : 'Unknown error' };
  }
}






