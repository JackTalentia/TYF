/* =====================================================
 * Info Sidebar — Code.gs (row-map loader + full UI support)
 * ===================================================== */
const INFO_SHEET_NAME      = 'Info';
const LOGISTICS_SHEET_NAME = 'Logistics';
const CONTACTS_SHEET_NAME  = '👤';
const DUE_SHEET_NAME       = '💰';
const PAYMENTS_SHEET_NAME  = '💵';

/* ---- Test switches (only affect Info row resolution) ---- */
const TEST_MODE = false;     // set true to force TEST_ROW for Info sheet reads/writes
const TEST_ROW  = 14;        // A1 row number in Info

/* Menu */
function onOpen(){
  SpreadsheetApp.getUi().createMenu('Info')
    .addItem('Open Info','openInfoSidebar')
    .addToUi();
}

function ping(){ return 'pong'; }

/* Sidebar launcher */
function openInfoSidebar(){
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getActiveSheet();
  if (!sh || sh.getName() !== LOGISTICS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert(`Run from '${LOGISTICS_SHEET_NAME}' (active row column A = Group).`);
    return;
  }
  const row = sh.getActiveRange().getRow();
  const groupName = String(sh.getRange(row, 1).getDisplayValue() || '').trim();
  const t = HtmlService.createTemplateFromFile('InfoSidebar');
  t.groupName = groupName;
  t.logRow    = row;
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('Info'));
}

/* ---------- Utilities ---------- */
function _ensureSheet_(name, headerRowIndex, headerValues) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headerRowIndex && headerValues && headerValues.length) {
    sh.getRange(headerRowIndex, 1, 1, headerValues.length)
      .setValues([headerValues]).setFontWeight('bold');
  }
  return sh;
}
function _A1rowOfGroupInInfo_(group) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheetByName(INFO_SHEET_NAME);
  if (!info) throw new Error(`Sheet '${INFO_SHEET_NAME}' not found.`);
  const last = info.getLastRow();
  if (last < 2) return -1;
  const keys = info.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const idx0 = keys.indexOf(String(group || '').trim());
  return (idx0 === -1) ? -1 : (idx0 + 2);
}

function getInfoGroups(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INFO_SHEET_NAME);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];

  // Read, clean, de-dup, sort case-insensitively
  const vals = sh.getRange(2, 1, last - 1, 1).getDisplayValues()
                 .map(r => String(r[0] || '').trim())
                 .filter(Boolean);
  const unique = Array.from(new Set(vals));
  unique.sort((a,b) => a.localeCompare(b, undefined, {sensitivity:'base'}));
  return unique;
}

/** Resolve Info row directly from a group name */
function getInfoRowForGroup(groupName){
  return _A1rowOfGroupInInfo_(groupName);
}


function _isoFromAny_(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const y=v.getFullYear(), m=('0'+(v.getMonth()+1)).slice(-2), d=('0'+v.getDate()).slice(-2);
    return `${y}-${m}-${d}`;
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const y = m[3].length === 2 ? ('20'+m[3]) : m[3];
    return `${y}-${('0'+m[2]).slice(-2)}-${('0'+m[1]).slice(-2)}`;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? '' : _isoFromAny_(d);
}
function _dmyFromIso_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${m[3]}/${m[2]}/${m[1].slice(-2)}` : String(s || '');
}
function _parseDateMs_(v) {
  if (v instanceof Date) return v.getTime();
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const y = m[3].length===2 ? (2000 + (+m[3])) : (+m[3]);
    const d = new Date(y, (+m[2])-1, +m[1]);
    return isNaN(d.getTime()) ? NaN : d.getTime();
  }
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? NaN : d2.getTime();
}
function _colLetter_(n){ let s=''; while(n>0){ const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26);} return s; }
function _num_(x){ const m=String(x||'').match(/-?\d+(\.\d+)?/); return m?+m[0]:''; }

/* ---------- Resolve Info row via Logistics!AA (27) or fallback by group ---------- */
function getInfoRowFromLogistics(logisticsRow, groupName){
  if (TEST_MODE) return TEST_ROW;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOGISTICS_SHEET_NAME);
  if (!log) throw new Error(`Sheet '${LOGISTICS_SHEET_NAME}' not found.`);
  const aa = Number(log.getRange(logisticsRow, 27).getValue()); // AA
  if (!isNaN(aa) && aa >= 2) return Math.floor(aa);
  const key = String(groupName || log.getRange(logisticsRow,1).getDisplayValue() || '').trim();
  return _A1rowOfGroupInInfo_(key);
}

/* ---------- Read whole Info row then map ---------- */
function infoRowToMap(row){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INFO_SHEET_NAME);
  if (!sh) throw new Error(`Sheet '${INFO_SHEET_NAME}' not found.`);
  const lastCol = Math.max(1, sh.getLastColumn());
  if (row < 2 || row > sh.getLastRow()) return { __empty:true, row, lastCol };

  const headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(h => String(h||'').trim());
  const values  = sh.getRange(row,1,1,lastCol).getDisplayValues()[0];

  const byHeader = {};
  headers.forEach((h,i)=>{ if (h) byHeader[h] = values[i]; });

  const byIndex = {};
  const byCol   = {};
  for (let i=0;i<lastCol;i++){
    byIndex[i+1] = values[i];
    byCol[_colLetter_(i+1)] = values[i];
  }
  return { row, lastCol, headers, byHeader, byIndex, byCol };
}

function buildInfoObjFromMap(map){
  const v = (idx) => map.byIndex[idx] ?? '';

  const obj = {
    'Group'        : v(1),

    // Dates -> ISO for <input type="date">
    'Arrival'      : _isoFromAny_(v(4)),
    'Departure'    : _isoFromAny_(v(5)),

    'Deal Status'  : v(6),
    'Type'         : v(7),
    'Source'       : v(8),
    'Hubspot Link' : v(9),

    'Accommodation' : v(11),
    'Booking Method': v(12),

    // FIX: Board is column 14 (not 13)
    'Board'         : v(14),
    // Optional: expose Accom Type (col 13) if desired by UI
    'Accommodation Type': v(13),

    'Extra Information'    : v(16),
    'Accommodation Status' : v(17),

    // More dates/percents
    'Deposit Date'         : _isoFromAny_(v(18)),
    'Deposit %'            : _num_(v(19)),
    'Balance Date'         : _isoFromAny_(v(20)),
    'Balance %'            : _num_(v(21)),

    'Prefix'               : v(22),

    // Finance
    'Admin Charge %'             : _num_(v(75)),
    'Charge Type'                : v(76),
    'Additional Activity Charge' : _num_(v(77)),
    'Other Charges Amount'       : _num_(v(78)),
    'Other Charges Description'  : v(79),
    'Discount £ (per person)'    : _num_(v(80)),
    'Discount % (per booking)'   : _num_(v(81)),
    'Free Places'                : _num_(v(82)),
    'Xero Invoice'               : v(83),

    '__row' : map.row,
    '__lastCol' : map.lastCol
  };

  // P1..P16 (39..54), L1..L16 (55..70)
  for (let i=1;i<=16;i++) obj['P'+i] = v(38+i);
  for (let i=1;i<=16;i++) obj['L'+i] = v(54+i);

  return obj;
}

/** One-call: Info object by row */
function getInfoObjByRow(row){
  const m = infoRowToMap(row);
  if (m.__empty) return m;
  return buildInfoObjFromMap(m);
}

/* ---------- Lists ---------- */
function getAccomList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Accom');
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 3) return [];
  return sh.getRange(3, 1, last - 2, 1).getValues()
    .map(r => String(r[0] || '').trim()).filter(Boolean);
}

/* ---------- Contacts / Payments / TotalDue remain by GROUP NAME ---------- */
function getContacts(groupName) {
  const key = String(groupName || '').trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONTACTS_SHEET_NAME);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 5) return [];
  const rows = sh.getRange(5, 1, last - 4, 6).getValues(); // A..F
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() !== key) continue;
    out.push({ name:rows[i][1]||'', title:rows[i][2]||'', mobile:rows[i][3]||'', work:rows[i][4]||'', email:rows[i][5]||'' });
  }
  return out;
}
function saveContacts(groupName, contacts) {
  const name = String(groupName || '').trim();
  if (!name) return { ok:false, message:'Group required for contacts.' };
  const sh = _ensureSheet_(CONTACTS_SHEET_NAME, 4, ['Group','Name','Job Title','Mobile','Work Phone','Email']);
  const last = sh.getLastRow();
  if (last >= 5) {
    const data = sh.getRange(5, 1, last - 4, 6).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0] || '').trim() === name) sh.deleteRow(5 + i);
    }
  }
  const rows = (contacts || [])
    .map(c => [name, c.name||'', c.title||'', c.mobile||'', c.work||'', c.email||''])
    .filter(r => r.slice(1).some(x => String(x||'').trim() !== ''));
  if (rows.length) {
    const start = Math.max(sh.getLastRow()+1, 5);
    sh.getRange(start, 1, rows.length, 6).setValues(rows);
  }
  return { ok:true };
}

function getTotalDue(groupName) {
  const key = String(groupName || '').trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DUE_SHEET_NAME);
  if (!sh) return '';
  const last = sh.getLastRow();
  if (last < 5) return '';
  const data = sh.getRange(5, 1, last - 4, Math.max(6, sh.getLastColumn())).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === key) return data[i][5] || '';
  }
  return '';
}
function saveTotalDue(groupName, totalDue) {
  const name = String(groupName || '').trim();
  if (!name) return { ok:false, message:'Group required for total due.' };
  const sh = _ensureSheet_(DUE_SHEET_NAME, 4, ['Group','','','','','Total Due']);
  const last = sh.getLastRow();
  if (last >= 5) {
    const data = sh.getRange(5,1,last-4,Math.max(6,sh.getLastColumn())).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === name) {
        sh.getRange(5 + i, 6).setValue(totalDue);
        return { ok:true };
      }
    }
  }
  const start = Math.max(sh.getLastRow()+1, 5);
  sh.getRange(start, 1, 1, 6).setValues([[name,'','','','', totalDue]]);
  return { ok:true };
}

function getPayments(groupName) {
  const key = String(groupName || '').trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(PAYMENTS_SHEET_NAME);
  if (!sh) return { lines: [] };
  const last = sh.getLastRow();
  if (last < 5) return { lines: [] };
  const data = sh.getRange(5, 1, last - 4, Math.max(4, sh.getLastColumn())).getValues();
  const lines = [];
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() !== key) continue;
    lines.push({
      amount:  data[i][1],
      dateDue: _isoFromAny_(data[i][2]),
      datePaid:_isoFromAny_(data[i][3])
    });
  }
  lines.sort((a,b)=>{
    const am=_parseDateMs_(a.dateDue), bm=_parseDateMs_(b.dateDue);
    if (isNaN(am) && isNaN(bm)) return 0;
    if (isNaN(am)) return 1;
    if (isNaN(bm)) return -1;
    return am - bm;
  });
  return { lines };
}
function savePayments(groupName, lines) {
  const name = String(groupName || '').trim();
  if (!name) return { ok:false, message:'Group required for payments.' };
  const sh = _ensureSheet_(PAYMENTS_SHEET_NAME, 4, ['Group','Amount','Date Due','Date Paid']);
  const last = sh.getLastRow();
  if (last >= 5) {
    const data = sh.getRange(5,1,last-4,Math.max(4,sh.getLastColumn())).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0] || '').trim() === name) sh.deleteRow(5 + i);
    }
  }
  const out = (lines || [])
    .map(x => [ name,
                (x.amount===''||x.amount==null)?'':(+x.amount),
                _dmyFromIso_(x.dateDue || ''),
                _dmyFromIso_(x.datePaid || '') ])
    .filter(r => r[1]!=='' || r[2] || r[3]);
  if (out.length) {
    out.sort((a,b)=>{
      const am=_parseDateMs_(a[2]), bm=_parseDateMs_(b[2]);
      if (isNaN(am) && isNaN(bm)) return 0;
      if (isNaN(am)) return 1;
      if (isNaN(bm)) return -1;
      return am - bm;
    });
    const start = Math.max(sh.getLastRow()+1, 5);
    sh.getRange(start, 1, out.length, 4).setValues(out);
  }
  return { ok:true, message:'Payments saved' };
}

/* ---------- Save back to Info (writes by fixed columns) ---------- */
function saveInfoData(originalName, updated) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const info = ss.getSheetByName(INFO_SHEET_NAME);
  if (!info) throw new Error(`Sheet '${INFO_SHEET_NAME}' not found.`);

  const newName = String((updated && updated['Group']) || originalName || '').trim();
  if (!newName) return { ok:false, message:'Group name is required.' };

  let row = TEST_MODE ? TEST_ROW : _A1rowOfGroupInInfo_(originalName || newName);
  if (!TEST_MODE && row === -1) {
    row = info.getLastRow() + 1;
    if (row < 2) row = 2;
  }
  function setC(col, val) { if (val !== undefined) info.getRange(row, col).setValue(val); }

  setC(1, newName);
  setC(4, _dmyFromIso_(updated['Arrival'] || ''));
  setC(5, _dmyFromIso_(updated['Departure'] || ''));
  setC(6, updated['Deal Status'] || updated['Status'] || '');
  setC(7, updated['Type'] || '');
  setC(8, updated['Source'] || '');
  setC(9, updated['Hubspot Link'] || '');

  setC(11, updated['Accommodation'] || '');
  setC(12, updated['Booking Method'] || '');
  setC(13, updated['Accommodation Type'] || '');
  setC(14, updated['Board'] || '');
  setC(16, updated['Extra Information'] || '');
  setC(17, updated['Accommodation Status'] || '');
  setC(18, _dmyFromIso_(updated['Deposit Date'] || ''));
  setC(19, updated['Deposit %'] || '');
  setC(20, _dmyFromIso_(updated['Balance Date'] || ''));
  setC(21, updated['Balance %'] || '');

  const prefix = String(updated['Prefix'] || '').trim();
  setC(22, prefix);
  for (let i = 0; i < 16; i++) setC(23 + i, prefix ? `${prefix} #${i+1}` : '');
  for (let i = 0; i < 16; i++) setC(39 + i, updated['P'+(i+1)] || '');
  for (let i = 0; i < 16; i++) setC(55 + i, updated['L'+(i+1)] || '');

  setC(75, updated['Admin Charge %'] || updated['Admin %'] || '');
  setC(76, updated['Charge Type'] || '');
  setC(77, updated['Additional Activity Charge'] || updated['Additional Activity Charge (pp)'] || '');
  setC(78, updated['Other Charges Amount'] || updated['Other Charges (Amount)'] || '');
  setC(79, updated['Other Charges Description'] || updated['Other Charges (Description)'] || '');
  setC(80, updated['Discount £ (per person)'] || updated['Discount £ (pp)'] || '');
  setC(81, updated['Discount % (per booking)'] || updated['Discount %'] || '');
  setC(82, updated['Free Places'] || '');
  setC(83, updated['Xero Invoice'] || updated['Xero Invoice Number'] || '');

  return { ok:true, message:'Saved' };
}
