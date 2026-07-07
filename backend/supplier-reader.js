'use strict';

const ExcelJS = require('exceljs');

// Required fields that must come from PO Header (or equivalent single-sheet columns)
const REQUIRED_HEADER_COLS = [
  'PO_Number',
  'Cargo_Ready_Planned_Collection_Date', 'Carrier_Booking_Request_Date',
  'Traffic_Mode', 'Booking_Group',
  'No_of_Cartons', 'Unit_Weight_KG', 'Carton_Type'
];

// Required fields that must come from PO Lines (used only by legacy single-sheet fallback)
const REQUIRED_SKU_COLS = [
  'PO_Number', 'SKU', 'Booking_Qty', 'No_of_Cartons', 'Unit_Weight_KG'
];

// ── Helpers ───────────────────────────────────────────────────────────────────

function cellVal(cell) {
  let val = cell.value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object' && 'result' in val) val = val.result;
  if (val instanceof Date) return val;
  if (typeof val === 'object' && 'richText' in val) return (val.richText || []).map(r => r.text || '').join('');
  if (typeof val === 'object' && 'formula' in val) return val.result ?? '';
  return val;
}

// Maps user-friendly column display labels (as written in the Excel header) back to
// the internal field names used throughout the codebase.
const COLUMN_ALIASES = {
  'Total no. of Cartons of booking': 'No_of_Cartons',
  'Total items weight of booking':   'Unit_Weight_KG',
};

/**
 * Extract rows from a sheet into an array of plain objects.
 * The header row is auto-detected as the first row containing `anchorCol`.
 * Column labels are normalised (whitespace / parenthetical annotations stripped).
 */
function readSheet(sheet, anchorCol) {
  // Detect header row
  let headerRowNum = 1;
  let found = false;
  sheet.eachRow((row, rowNum) => {
    if (found) return;
    row.eachCell(cell => {
      if (String(cell.value || '').replace(/\s*\(.*?\)/, '').trim() === anchorCol) {
        headerRowNum = rowNum;
        found = true;
      }
    });
  });

  const headers = [];
  sheet.getRow(headerRowNum).eachCell((cell, colNum) => {
    const raw = String(cell.value || '').replace(/\s*\(.*?\)/, '').trim();
    headers[colNum] = COLUMN_ALIASES[raw] || raw;
  });

  const rows = [];
  sheet.eachRow((row, rowNum) => {
    if (rowNum <= headerRowNum) return;
    const obj = {};
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const key = headers[colNum];
      if (key) obj[key] = cellVal(cell);
    });
    if (!Object.values(obj).some(v => v !== '')) return; // skip blank rows
    rows.push({ _rowNum: rowNum, ...obj });
  });

  return { rows, headerRowNum };
}

// ── Legacy single-sheet parser (SUPPLIER_INPUT) ───────────────────────────────

const LEGACY_REQUIRED_COLS = [
  'PO_Number', 'SKU', 'No_of_Cartons', 'Unit_Weight_KG', 'Booking_Qty',
  'Cargo_Ready_Planned_Collection_Date', 'Carrier_Booking_Request_Date',
  'Traffic_Mode',
  'Booking_Group'
];

function parseSingleSheet(sheet) {
  const { rows: rawRows, headerRowNum } = readSheet(sheet, 'PO_Number');

  const rows = [];
  const validationErrors = [];

  for (const obj of rawRows) {
    const po  = String(obj.PO_Number || '').trim();
    const sku = String(obj.SKU       || '').trim();
    if (!po && !sku) continue;

    const missing = LEGACY_REQUIRED_COLS.filter(c => {
      if (String(obj.Factory_ID || '').trim() === '9999' &&
          ['Factory_Name','Factory_Street1','Factory_City','Factory_PostalCd','Factory_CountryCd'].includes(c)) return false;
      return !obj[c] || String(obj[c]).trim() === '';
    });
    if (missing.length > 0) {
      validationErrors.push(`Row ${obj._rowNum}: missing required fields: ${missing.join(', ')}`);
    }
    if (String(obj.Collection_Type || '').trim() === 'Collection' &&
        (!obj.Collection_Time || String(obj.Collection_Time).trim() === '')) {
      validationErrors.push(`Row ${obj._rowNum}: Collection_Time is required when Collection_Type is "Collection"`);
    }

    rows.push(obj);
  }

  // Fill-down Booking_Group within each PO group
  const poGroupMap = {};
  for (const row of rows) {
    const po = String(row.PO_Number || '').trim();
    const bg = String(row.Booking_Group || '').trim();
    if (po && bg && !poGroupMap[po]) poGroupMap[po] = bg;
  }
  for (const row of rows) {
    const po = String(row.PO_Number || '').trim();
    if (po && (!row.Booking_Group || String(row.Booking_Group).trim() === '') && poGroupMap[po]) {
      row.Booking_Group = poGroupMap[po];
    }
  }

  return { rows, validationErrors, sheetName: sheet.name, headerRowNum,
    headerPoRefs: [...new Set(rows.map(r => String(r.PO_Number || '').trim()).filter(Boolean))] };
}

// ── PO Header-only parser ─────────────────────────────────────────────────────

/**
 * Parse a workbook that has a 'PO Header' (or 'BOOKING_HEADER') sheet only.
 * Each row becomes a header-only placeholder (_headerOnly: true).
 * SKU lines are NOT expected — they will be auto-booked from the carrier ASN feed.
 */
function parseHeaderOnly(workbook) {
  const wsHdr = workbook.getWorksheet('PO Header') || workbook.getWorksheet('BOOKING_HEADER');
  const { rows: rawRows, headerRowNum } = readSheet(wsHdr, 'PO_Number');

  const rows = [];
  const validationErrors = [];

  for (const obj of rawRows) {
    const po = String(obj.PO_Number || '').trim();
    if (!po) continue;

    const missing = REQUIRED_HEADER_COLS.filter(c => !obj[c] || String(obj[c]).trim() === '');
    if (missing.length > 0) {
      validationErrors.push(`Row ${obj._rowNum}: missing required fields: ${missing.join(', ')}`);
    }
    if (String(obj.Collection_Type || '').trim() === 'Collection' &&
        (!obj.Collection_Time || String(obj.Collection_Time).trim() === '')) {
      validationErrors.push(`Row ${obj._rowNum}: Collection_Time is required when Collection_Type is "Collection"`);
    }

    rows.push({ ...obj, _headerOnly: true });
  }

  return {
    rows,
    validationErrors,
    sheetName: wsHdr.name,
    headerRowNum,
    headerPoRefs: [...new Set(rows.map(r => String(r.PO_Number || '').trim()).filter(Boolean))]
  };
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Parse a supplier Excel buffer.
 * Current format: PO Header sheet only — SKUs auto-booked from ASN feed.
 * Legacy fallback: SUPPLIER_INPUT single-sheet format.
 */
async function parse(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  // PO Header sheet present — use header-only parser (PO Lines not used)
  const wsHdr = workbook.getWorksheet('PO Header') || workbook.getWorksheet('BOOKING_HEADER');
  if (wsHdr) return parseHeaderOnly(workbook);

  // Legacy fallback: SUPPLIER_INPUT or first visible sheet
  let sheet = workbook.getWorksheet('SUPPLIER_INPUT');
  if (!sheet) {
    sheet = workbook.worksheets.find(ws => ws.state !== 'veryHidden' && ws.state !== 'hidden');
  }
  if (!sheet) sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Supplier Excel has no worksheets');

  return parseSingleSheet(sheet);
}

module.exports = { parse };
