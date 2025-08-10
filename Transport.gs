/***** =========================
 *  Transport Logistics – Code.gs (fast, no expand)
 *  ========================= *****/

/** CONFIG **/
const MAIN_SHEET_NAME      = 'Logistics';       // active row + meta
const TRANSPORT_SHEET_NAME = '🚌';     // storage (A=ID, B:… = fields)
const LOCATION_SHEET_NAME  = 'Location';        // From/To list A2:A
const SETTINGS_SHEET_NAME  = 'Settings';        // Times list A4:A

// B:… headers in this order (A is ID) — Charge removed
const TRANSPORT_HEADERS = [
  '🚌Status','From','To','Outbound','Return',
  'TYF PAX','Transport Notes','RB PAX','RB Transport Notes',
  'Richards Bros Reference','Cost'
];

const STATUS_OPTIONS = ['Booked','Requested','Needs Booking','Not Required','On Xero'];

// Caches (use new keys to avoid collisions with older versions)
const DROPDOWN_CACHE_TTL   = 1200; // 20 min
const INDEX_CACHE_KEY_V3   = 'transport_index_v3';
const LOC_CACHE_KEY_V2     = 'locations_v2';
const TIMES_CACHE_KEY_V2   = 'times_v2';


/** MENU / ENTRY POINTS **/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Transport Tools')
    .addItem('Open Sidebar', 'openTransportSidebar')
    .addSeparator()
    .addItem('Warm caches (faster first load)', 'warmTransportCaches')
    .addItem('Rebuild transport index', 'rebuildTransportIndex')
    .addToUi();

  // Pre-warm silently so the first sidebar open is snappy
  try { warmTransportCaches(); } catch (e) {}
}

function openTransportSidebar() {
  const t = HtmlService.createTemplateFromFile('TransportSidebar');
  SpreadsheetApp.getUi().showSidebar(
    t.evaluate().setTitle('Transport Logistics')
  );
  return true;
}


/** UTILITIES **/
function getActiveActivityId_() {
  // Activity ID is in Logistics!AB:AB
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sh) throw new Error(`Sheet '${MAIN_SHEET_NAME}' not found.`);
  const row = sh.getActiveRange().getRow();
  const ACTIVITY_ID_COL = sh.getRange('AB1').getColumn(); // AB
  const activityId = String(sh.getRange(row, ACTIVITY_ID_COL).getValue() || '').trim();
  return { row, activityId };
}

function getDocCache_() { try { return CacheService.getDocumentCache(); } catch(e){ return null; } }
function getProps_()     { return PropertiesService.getDocumentProperties(); }

function getCached_(key, computeFn, ttl = DROPDOWN_CACHE_TTL, force = false) {
  const cache = getDocCache_();
  if (!cache || force) return computeFn();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const val = computeFn();
  cache.put(key, JSON.stringify(val), ttl);
  return val;
}

function uniq_(arr){ const s=new Set, out=[]; for (const v of arr) if (v && !s.has(v)) { s.add(v); out.push(v); } return out; }
function coerceInt_(v){ const n=parseInt(String(v).replace(/[^\d-]/g,''),10); return isNaN(n)?0:n; }
function coerceFloat_(v){ const n=parseFloat(String(v).replace(/[^0-9.\-]/g,'')); return isNaN(n)?0:n; }

function normalizeTime_(x){
  if (x instanceof Date) return Utilities.formatDate(x, Session.getScriptTimeZone(), 'HH:mm');
  const s = String(x||'').trim();
  const m = s.match(/^(\d{1,2}):([0-5]\d)$/); if (m) return ('0'+m[1]).slice(-2)+':'+m[2];
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
}

function fmtDatePretty_(v) {
  return (v instanceof Date)
    ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'EEE d MMM yyyy')
    : String(v || '');
}


/** ID→ROW INDEX (fast, cached) **/
function buildTransportIndex_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TRANSPORT_SHEET_NAME);
  if (!sh) throw new Error(`Sheet '${TRANSPORT_SHEET_NAME}' not found.`);
  const last = sh.getLastRow();
  if (last < 2) return {}; // header only

  // Read only A2:A<lastRow>
  const height = last - 1;
  const ids = sh.getRange(2, 1, height, 1).getValues();
  const map = {};
  for (let i = 0; i < height; i++) {
    const id = String(ids[i][0] || '').trim();
    if (id) map[id] = i + 2; // +2 because data starts at row 2
  }
  return map;
}

function getTransportIndex_(force) {
  const cache = getDocCache_();
  const props = getProps_();

  if (!force) {
    if (cache) {
      const hit = cache.get(INDEX_CACHE_KEY_V3);
      if (hit) return JSON.parse(hit);
    }
    const p = props.getProperty(INDEX_CACHE_KEY_V3);
    if (p) {
      const obj = JSON.parse(p);
      if (cache) cache.put(INDEX_CACHE_KEY_V3, p, 3600); // 60 min
      return obj;
    }
  }
  const idx = buildTransportIndex_();
  const json = JSON.stringify(idx);
  props.setProperty(INDEX_CACHE_KEY_V3, json);
  const cache2 = getDocCache_();
  if (cache2) cache2.put(INDEX_CACHE_KEY_V3, json, 3600);
  return idx;
}

function updateTransportIndex_(activityId, row) {
  const props = getProps_();
  let obj = {};
  const existing = props.getProperty(INDEX_CACHE_KEY_V3);
  if (existing) { try { obj = JSON.parse(existing); } catch(e){} }
  if (row === -1) delete obj[activityId]; else obj[activityId] = row;

  const json = JSON.stringify(obj);
  props.setProperty(INDEX_CACHE_KEY_V3, json);
  const cache = getDocCache_();
  if (cache) cache.put(INDEX_CACHE_KEY_V3, json, 3600);
}

function findRowByActivityId_(activityId) {
  const idx = getTransportIndex_(false);
  if (idx && activityId in idx) return idx[activityId];
  const idx2 = getTransportIndex_(true); // one rebuild on miss
  return (idx2 && activityId in idx2) ? idx2[activityId] : -1;
}


/** DROPDOWN SOURCES (cached to lastRow) **/
function getLocations_(force=false){
  return getCached_(LOC_CACHE_KEY_V2, () => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOCATION_SHEET_NAME);
    if (!sh) return [];
    const last = sh.getLastRow();
    if (last < 2) return [];
    const vals = sh.getRange(2, 1, last - 1, 1).getValues(); // A2:A<last>
    const out = [];
    for (let i = 0; i < vals.length; i++) {
      const v = String(vals[i][0] || '').trim();
      if (v) out.push(v);
    }
    return uniq_(out);
  }, DROPDOWN_CACHE_TTL, force);
}

function getTimes_(force=false){
  return getCached_(TIMES_CACHE_KEY_V2, () => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    if (!sh) return [];
    const last = sh.getLastRow();
    if (last < 4) return [];
    const height = last - 3;
    const raw = sh.getRange(4, 1, height, 1).getValues(); // A4:A<last>
    const out = [];
    for (let i = 0; i < raw.length; i++) {
      const n = normalizeTime_(raw[i][0]);
      if (n) out.push(n);
    }
    return uniq_(out);
  }, DROPDOWN_CACHE_TTL, force);
}


/** DATA APIs used by HTML **/

// One fast call for initial load (lists + row)
function getInitData(forceLists){
  const lists = {
    statusOptions: STATUS_OPTIONS,
    locations: getLocations_(!!forceLists),
    times: getTimes_(!!forceLists)
  };

  const row = getTransportRow(); // includes meta + values

  return {
    statusOptions: lists.statusOptions,
    locations: lists.locations,
    times: lists.times,
    activityId: row.activityId,
    isNew: row.isNew,
    meta: row.meta,
    values: row.values
  };
}

// Lists only (used by “Sync”)
function getTransportLists(force){
  return {
    statusOptions: STATUS_OPTIONS,
    locations: getLocations_(!!force),
    times: getTimes_(!!force)
  };
}

// Row values + header meta (B,G,H,J,K) and ID from AB
function getTransportRow() {
  const { row, activityId } = getActiveActivityId_();

  // Batch-read the first 11 columns in one go (A..K)
  const main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  let meta = { subgroup:'', date:'', time:'', activity:'', location:'' };
  if (main && row) {
    const vals = main.getRange(row, 1, 1, Math.max(11, main.getLastColumn())).getValues()[0];
    meta.subgroup = String(vals[1] || '');      // B
    meta.date     = fmtDatePretty_(vals[6]);    // G
    meta.time     = normalizeTime_(vals[7]);    // H
    meta.activity = String(vals[9] || '');      // J
    meta.location = String(vals[10]|| '');      // K
  }

  if (!activityId) {
    return {
      error: `No Activity ID in AB for row ${row}.`,
      activityId: '',
      isNew: true,
      meta,
      values: {
        status:'', from:'', to:'', outbound:'', ret:'',
        tyfPax:0, notes:'', rbPax:0, rbNotes:'', rbRef:'', cost:0
      }
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TRANSPORT_SHEET_NAME);
  if (!sh) throw new Error(`Sheet '${TRANSPORT_SHEET_NAME}' not found.`);
  if (sh.getLastRow() === 0) sh.appendRow(['ID'].concat(TRANSPORT_HEADERS));

  const matchRow = findRowByActivityId_(activityId);
  if (matchRow === -1) {
    return {
      activityId, isNew: true, meta,
      values: {
        status:'', from:'', to:'', outbound:'', ret:'',
        tyfPax:0, notes:'', rbPax:0, rbNotes:'', rbRef:'', cost:0
      }
    };
  }

  // Read actual header row to handle old sheets that still have 🚌Charge
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 2, 1, Math.max(0, lastCol - 1)).getValues()[0] || [];
  const hasCharge = header.indexOf('🚌Charge') !== -1;

  const rowVals = sh.getRange(matchRow, 2, 1, header.length).getValues()[0];
  const idx = (base) => base + (hasCharge && base >= 1 ? 1 : 0); // shift by 1 after Status if Charge exists

  return {
    activityId, isNew:false, meta,
    values: {
      status:   rowVals[idx(0)],
      from:     rowVals[idx(1)],
      to:       rowVals[idx(2)],
      outbound: normalizeTime_(rowVals[idx(3)]),
      ret:      normalizeTime_(rowVals[idx(4)]),
      tyfPax:   coerceInt_(rowVals[idx(5)]),
      notes:    rowVals[idx(6)],
      rbPax:    coerceInt_(rowVals[idx(7)]),
      rbNotes:  rowVals[idx(8)],
      rbRef:    rowVals[idx(9)],
      cost:     coerceFloat_(rowVals[idx(10)])
    }
  };
}


/** SAVE (update-or-create; single fast write) **/
function saveTransportData(data) {
  try {
    const activityId = String(data.activityId || '').trim();
    if (!activityId) return { ok:false, message:'Missing Activity ID.' };

    const allowedTimes = new Set(getTimes_(false));
    const outbound = data.values.outbound ? normalizeTime_(data.values.outbound) : '';
    const ret      = data.values.ret      ? normalizeTime_(data.values.ret)      : '';
    if (allowedTimes.size) {
      if (outbound && !allowedTimes.has(outbound)) return { ok:false, message:`Outbound must match Settings times (got "${outbound}")` };
      if (ret && !allowedTimes.has(ret))           return { ok:false, message:`Return must match Settings times (got "${ret}")` };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(TRANSPORT_SHEET_NAME);
    if (!sh) throw new Error(`Sheet '${TRANSPORT_SHEET_NAME}' not found.`);
    if (sh.getLastRow() === 0) sh.appendRow(['ID'].concat(TRANSPORT_HEADERS));

    let row = findRowByActivityId_(activityId);
    if (row === -1) { row = sh.getLastRow() + 1; sh.getRange(row, 1).setValue(activityId); }

    // Work with the actual header row to tolerate legacy sheets (with 🚌Charge)
    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 2, 1, Math.max(0, lastCol - 1)).getValues()[0] || [];
    const hasCharge = header.indexOf('🚌Charge') !== -1;
    const idx = (base) => base + (hasCharge && base >= 1 ? 1 : 0);

    // Start from existing row to preserve any unknown columns (incl. legacy Charge)
    let rowVals;
    if (row > 1 && row <= sh.getLastRow()) {
      rowVals = sh.getRange(row, 2, 1, header.length).getValues()[0];
    } else {
      rowVals = new Array(header.length).fill('');
    }

    // Assign our fields
    rowVals[idx(0)]  = data.values.status || '';
    rowVals[idx(1)]  = data.values.from   || '';
    rowVals[idx(2)]  = data.values.to     || '';
    rowVals[idx(3)]  = outbound;
    rowVals[idx(4)]  = ret;
    rowVals[idx(5)]  = coerceInt_(data.values.tyfPax);
    rowVals[idx(6)]  = data.values.notes  || '';
    rowVals[idx(7)]  = coerceInt_(data.values.rbPax);
    rowVals[idx(8)]  = data.values.rbNotes || '';
    rowVals[idx(9)]  = String(data.values.rbRef || '');
    rowVals[idx(10)] = coerceFloat_(data.values.cost);

    // Write back in one go
    sh.getRange(row, 2, 1, header.length).setValues([rowVals]);

    // Format Cost as £ using the live header row
    const costIdx = header.indexOf('Cost'); // 0-based from column B
    if (costIdx >= 0) {
      const col = 2 + costIdx; // absolute column number
      sh.getRange(row, col).setNumberFormat('£#,##0.00');
    }

    // Keep index hot
    updateTransportIndex_(activityId, row);

    return { ok:true, message:'Saved' };
  } catch (e) {
    return { ok:false, message:e.message || 'Save failed' };
  }
}


/** OPTIONAL: warmers for speed **/
function warmTransportCaches() {
  getTransportIndex_(true);  // rebuild
  getLocations_(true);       // refresh cache
  getTimes_(true);           // refresh cache
}

function rebuildTransportIndex() {
  getTransportIndex_(true);
  SpreadsheetApp.getUi().alert('Transport index rebuilt.');
}
