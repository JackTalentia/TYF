/***** =========================
 *  Info Sidebar – Code.gs
 *  ========================= *****/

const INFO_SHEET_NAME      = 'Info';
const LOGISTICS_SHEET_NAME = 'Logistics';
const PAYMENTS_SHEET_NAME  = '💵';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Info Tools')
    .addItem('Open Info Sidebar', 'openInfoSidebar')
    .addToUi();
}

// Open from Logistics (Group in col A of active row)
function openInfoSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== LOGISTICS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert(
      `Please run from the '${LOGISTICS_SHEET_NAME}' sheet (active row column A = Group).`
    );
    return;
  }
  const row = sh.getActiveRange().getRow();
  const groupName = String(sh.getRange(row, 1).getValue() || '').trim();

  const t = HtmlService.createTemplateFromFile('InfoSidebar');
  t.groupName = groupName; // may be empty (new)
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle('Info'));
}

/* ---------- Utilities ---------- */
function _headers_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
}
function _findHeaderIndex_(headers, names) {
  for (const n of names) {
    const i = headers.indexOf(n);
    if (i !== -1) return i;
  }
  return -1;
}
function _parseDateMs_(v) {
  if (v instanceof Date) return v.getTime();
  const s = String(v || '').trim(); if (!s) return NaN;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const d = +m[1], mo = (+m[2]) - 1, y = m[3].length === 2 ? 2000 + (+m[3]) : (+m[3]);
    const dt = new Date(y, mo, d);
    return isNaN(dt.getTime()) ? NaN : dt.getTime();
  }
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? NaN : d2.getTime();
}
function _dmyFromIso_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${m[3]}/${m[2]}/${m[1].slice(-2)}` : String(s || '');
}

/* ---------- Info load/save ---------- */
function getInfoData(groupName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(INFO_SHEET_NAME);
  if (!sh) throw new Error(`Sheet '${INFO_SHEET_NAME}' not found.`);

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return null;

  const values  = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const gi = headers.indexOf('Group');
  if (gi === -1) throw new Error(`'Group' column not found in '${INFO_SHEET_NAME}'.`);

  if (!String(groupName||'').trim()) return null;

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][gi] || '').trim() === groupName) {
      const obj = {};
      for (let c = 0; c < headers.length; c++) if (headers[c]) obj[headers[c]] = values[r][c];
      return obj;
    }
  }
  return null;
}

function saveInfoData(originalName, updated) {
  try {
    const newName = String((updated && updated['Group']) || originalName || '').trim();
    if (!newName) return { ok:false, message:'Group name is required.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(INFO_SHEET_NAME);
    if (!sh) throw new Error(`Sheet '${INFO_SHEET_NAME}' not found.`);

    const headers = _headers_(sh);
    if (headers.length === 0) return { ok:false, message:'Info sheet appears empty.' };

    const gi = headers.indexOf('Group');
    if (gi === -1) return { ok:false, message:`'Group' column not found in '${INFO_SHEET_NAME}'.` };

    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    const data = sh.getRange(1, 1, Math.max(lastRow, 1), Math.max(lastCol, 1)).getValues();

    // find row by original or new name
    let rowIdx = -1;
    const seek = String(originalName || newName);
    if (seek) {
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][gi] || '').trim() === seek) { rowIdx = r; break; }
      }
    }
    if (rowIdx === -1 && newName) {
      for (let r = 1; r < data.length; r++) {
        if (String(data[r][gi] || '').trim() === newName) { rowIdx = r; break; }
      }
    }

    let row;
    if (rowIdx === -1) {
      rowIdx = sh.getLastRow() + 1;
      row = new Array(headers.length).fill('');
    } else {
      row = sh.getRange(rowIdx + 1, 1, 1, headers.length).getValues()[0];
    }

    row[gi] = newName;
    Object.keys(updated || {}).forEach(k => {
      const i = headers.indexOf(k);
      if (i >= 0) row[i] = updated[k];
    });

    sh.getRange(rowIdx + 1, 1, 1, headers.length).setValues([row]);
    return { ok:true, message:(seek ? 'Saved' : 'Created') };
  } catch (e) {
    return { ok:false, message:e.message || 'Save failed' };
  }
}

/* ---------- Supplier list from Accom ---------- */
function getAccomList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Accom');
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 3) return [];
  return sh.getRange(3, 1, last - 2, 1).getValues()
    .map(r => String(r[0] || '').trim()).filter(Boolean);
}

/* ---------- Payments (💵) load/save ---------- */
function _ensurePaymentsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PAYMENTS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PAYMENTS_SHEET_NAME);
    sh.getRange(1,1,1,5).setValues([['Group','Percent','Amount','Date Due','Date Paid']]);
    sh.getRange(1,1,1,5).setFontWeight('bold');
  } else {
    const head = _headers_(sh);
    const need = ['Group','Percent','Amount','Date Due','Date Paid'];
    for (const h of need) {
      if (head.indexOf(h) === -1) {
        sh.insertColumnAfter(sh.getLastColumn());
        sh.getRange(1, sh.getLastColumn(), 1, 1).setValue(h).setFontWeight('bold');
      }
    }
  }
  return sh;
}

function getPayments(groupName) {
  const sh = _ensurePaymentsSheet_();
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return { lines: [] };

  const values = sh.getRange(1,1,lastRow,lastCol).getValues();
  const head = values[0].map(x => String(x||'').trim());
  const gi  = _findHeaderIndex_(head, ['Group']);
  const pi  = _findHeaderIndex_(head, ['Percent','%']);
  const ai  = _findHeaderIndex_(head, ['Amount','Amount £','£']);
  const ddi = _findHeaderIndex_(head, ['Date Due','Due Date','Due']);
  const dpi = _findHeaderIndex_(head, ['Date Paid','Paid Date','Paid']);
  if (gi === -1 || ai === -1 || ddi === -1 || dpi === -1) return { lines: [] };

  const lines = [];
  for (let r=1; r<values.length; r++) {
    if (String(values[r][gi] || '').trim() !== String(groupName||'').trim()) continue;
    lines.push({
      percent: values[r][pi],
      amount:  values[r][ai],
      dateDue: values[r][ddi],
      datePaid:values[r][dpi],
    });
  }
  lines.sort((a,b)=>{
    const am = _parseDateMs_(a.dateDue), bm = _parseDateMs_(b.dateDue);
    if (isNaN(am) && isNaN(bm)) return 0;
    if (isNaN(am)) return 1;
    if (isNaN(bm)) return -1;
    return am - bm;
  });
  return { lines };
}

function savePayments(groupName, lines) {
  try {
    const name = String(groupName||'').trim();
    if (!name) return { ok:false, message:'Group name required for payments.' };

    const sh = _ensurePaymentsSheet_();
    const head = _headers_(sh);
    const gi  = _findHeaderIndex_(head, ['Group']);
    const pi  = _findHeaderIndex_(head, ['Percent','%']);
    const ai  = _findHeaderIndex_(head, ['Amount','Amount £','£']);
    const ddi = _findHeaderIndex_(head, ['Date Due','Due Date','Due']);
    const dpi = _findHeaderIndex_(head, ['Date Paid','Paid Date','Paid']);
    if (gi === -1 || ai === -1 || ddi === -1 || dpi === -1)
      return { ok:false, message:'Payments sheet missing required columns.' };

    // delete existing rows for group (bottom-up)
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow >= 2) {
      const data = sh.getRange(2,1,lastRow-1,lastCol).getValues();
      for (let i=data.length-1; i>=0; i--) {
        if (String(data[i][gi]||'').trim() === name) sh.deleteRow(i+2);
      }
    }

    // normalise + sort new lines
    const toWrite = (lines||[])
      .map(x => ({
        percent: x.percent === '' || x.percent == null ? '' : +x.percent,
        amount:  x.amount  === '' || x.amount  == null ? '' : +x.amount,
        dateDue: _dmyFromIso_(x.dateDue || ''),
        datePaid:_dmyFromIso_(x.datePaid || '')
      }))
      .filter(x => x.percent!=='' || x.amount!=='' || x.dateDue || x.datePaid);

    toWrite.sort((a,b)=>{
      const am = _parseDateMs_(a.dateDue), bm = _parseDateMs_(b.dateDue);
      if (isNaN(am) && isNaN(bm)) return 0;
      if (isNaN(am)) return 1;
      if (isNaN(bm)) return -1;
      return am - bm;
    });

    if (toWrite.length) {
      const rows = toWrite.map(x => {
        const row = new Array(sh.getLastColumn()).fill('');
        row[gi]  = name;
        if (pi !== -1) row[pi] = x.percent;
        row[ai]  = x.amount;
        row[ddi] = x.dateDue;
        row[dpi] = x.datePaid;
        return row;
      });
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, rows.length, sh.getLastColumn()).setValues(rows);
    }
    return { ok:true, message:'Payments saved' };
  } catch (e) {
    return { ok:false, message:e.message || 'Payments save failed' };
  }
}
