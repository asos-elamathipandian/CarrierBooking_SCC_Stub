'use strict';

const ExcelJS = require('exceljs');

// Required fields that must come from PO Header (or equivalent single-sheet columns)
const REQUIRED_HEADER_COLS = [
  'PO_Number',
  'Cargo_Ready_Planned_Collection_Date', 'Carrier_Booking_Request_Date',
  'Traffic_Mode', 'Mode_Of_Transport', 'Booking_Group',
  'Factory_ID', 'Factory_Name', 'Factory_Street1', 'Factory_City', 'Factory_PostalCd', 'Factory_CountryCd'
];

// Required fields that must come from PO Lines (or equivalent single-sheet columns)
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
    headers[colNum] = String(cell.value || '').replace(/\s*\(.*?\)/, '').trim();
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

// ── Two-sheet parser (PO Header + PO Lines) ────────────────────────────────────

function parseTwoSheet(workbook) {
  const wsHdr = workbook.getWorksheet('PO Header');
  const wsSku = workbook.getWorksheet('PO Lines');

  const { rows: hdrRows } = readSheet(wsHdr, 'PO_Number');
  const { rows: skuRows } = readSheet(wsSku, 'PO_Number');

  // Build a map of PO → header row
  const hdrMap = new Map();
  for (const h of hdrRows) {
    const po = String(h.PO_Number || '').trim();
    if (po && !hdrMap.has(po)) hdrMap.set(po, h);
  }

  const rows = [];
  const validationErrors = [];

  for (const skuRow of skuRows) {
    const po  = String(skuRow.PO_Number || '').trim();
    const sku = String(skuRow.SKU       || '').trim();
    if (!po && !sku) continue;

    const hdr = hdrMap.get(po) || {};

    // Merge: header fields + SKU fields (SKU fields take precedence on overlap)
    const merged = { ...hdr, ...skuRow };

    // Validate required header fields
    // Factory_ID 9999 = Dummy Factory: address fields are not required
    const isDummyFactory = String(merged.Factory_ID || '').trim() === '9999';
    const FACTORY_ADDRESS_COLS = ['Factory_Name', 'Factory_Street1', 'Factory_City', 'Factory_PostalCd', 'Factory_CountryCd'];
    const missingHdr = REQUIRED_HEADER_COLS.filter(c => {
      if (isDummyFactory && FACTORY_ADDRESS_COLS.includes(c)) return false;
      return !merged[c] || String(merged[c]).trim() === '';
    });
    if (missingHdr.length > 0) {
      validationErrors.push(
        `PO Lines row ${skuRow._rowNum} (PO ${po}): missing header fields: ${missingHdr.join(', ')}`
      );
    }

    // Validate required SKU fields
    const missingSku = REQUIRED_SKU_COLS.filter(c => !merged[c] || String(merged[c]).trim() === '');
    if (missingSku.length > 0) {
      validationErrors.push(
        `PO Lines row ${skuRow._rowNum}: missing required fields: ${missingSku.join(', ')}`
      );
    }

    // Collection_Time mandatory when Collection_Type = 'Collection'
    if (String(merged.Collection_Type || '').trim() === 'Collection' &&
        (!merged.Collection_Time || String(merged.Collection_Time).trim() === '')) {
      validationErrors.push(
        `PO Lines row ${skuRow._rowNum}: Collection_Time is required when Collection_Type is "Collection"`
      );
    }

    rows.push(merged);
  }

  return { rows, validationErrors, sheetName: 'PO Header+PO Lines', headerRowNum: 3,
    headerPoRefs: [...hdrMap.keys()] };
}

// ── Legacy single-sheet parser (SUPPLIER_INPUT) ───────────────────────────────

const LEGACY_REQUIRED_COLS = [
  'PO_Number', 'SKU', 'No_of_Cartons', 'Unit_Weight_KG', 'Booking_Qty',
  'Cargo_Ready_Planned_Collection_Date', 'Carrier_Booking_Request_Date',
  'Traffic_Mode', 'Mode_Of_Transport',
  'Factory_Name', 'Factory_ID', 'Factory_Street1', 'Factory_City', 'Factory_PostalCd', 'Factory_CountryCd',
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

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Parse a supplier Excel buffer.
 * Supports the two-sheet format (PO Header + PO Lines) and the legacy
 * single-sheet format (SUPPLIER_INPUT) for backward compatibility.
 */
async function parse(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  // Two-sheet format takes priority (new names first, legacy names as fallback)
  const wsHdr = workbook.getWorksheet('PO Header') || workbook.getWorksheet('BOOKING_HEADER');
  const wsSku = workbook.getWorksheet('PO Lines') || workbook.getWorksheet('SKU_LINES');
  if (wsHdr && wsSku) {
    return parseTwoSheet(workbook);
  }

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
