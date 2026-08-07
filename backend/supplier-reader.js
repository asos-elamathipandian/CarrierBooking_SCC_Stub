'use strict';

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Required fields that must come from PO Header (or equivalent single-sheet columns)
const REQUIRED_HEADER_COLS = [
  'PO_Number',
  'Cargo_Ready_Planned_Collection_Date', 'Carrier_Booking_Request_Date',
  'Booking_Group',
  'No_of_Cartons', 'Unit_Weight_KG'
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

function normalizeHeaderName(name) {
  return String(name || '')
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s*\(.*?\)/g, '')
    .replace(/[\r\n\t]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

// Maps user-friendly column display labels (as written in the Excel header)
// back to the internal field names used throughout the codebase.
const DEFAULT_COLUMN_ALIASES = {
  'Total no. of Cartons of booking':   'No_of_Cartons',
  'Total items weight of booking':     'Unit_Weight_KG',
  'Header_Booking_Qty (total units in booking)': 'Header_Booking_Qty',
  'Total booked units of a booking':              'Header_Booking_Qty',
  'ASN Number': 'ASN_Number',
  'ASN No': 'ASN_Number',
  'Box Qty': 'No_of_Cartons',
  'Unit Qty': 'Header_Booking_Qty',
  'KOLI TOTAL KG': 'Unit_Weight_KG',
  'KOLI TOTAL KG ': 'Unit_Weight_KG',
  'Koli Total Kg': 'Unit_Weight_KG',
};

const IDEATEKS_REQUIRED_COLS = [
  'ASN_Number',
  'No_of_Cartons',
  'Header_Booking_Qty',
  'Unit_Weight_KG'
];

function isoTodayPlus(daysToAdd) {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + daysToAdd);
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function applyIdeateksDefaults(obj) {
  if (!obj.Cargo_Ready_Planned_Collection_Date) {
    obj.Cargo_Ready_Planned_Collection_Date = isoTodayPlus(0);
  }
  if (!obj.Carrier_Booking_Request_Date) {
    obj.Carrier_Booking_Request_Date = isoTodayPlus(1);
  }
  obj.Booking_Group = 'Single Booking';
  if (!obj.Traffic_Mode) obj.Traffic_Mode = 'CFS';
  if (!obj.Carton_Type) obj.Carton_Type = 'BDCM1';
  if (!obj.Pack_Type) obj.Pack_Type = 'Bulk Flat';
  if (!obj.Collection_Type) obj.Collection_Type = 'Delivery';
  if (!obj.Hazardous) obj.Hazardous = 'N/A';
}

function loadColumnAliases() {
  const mappingPath = process.env.SUPPLIER_COLUMN_MAPPING_FILE
    ? path.resolve(process.cwd(), process.env.SUPPLIER_COLUMN_MAPPING_FILE)
    : path.join(__dirname, '..', 'config', 'supplier-column-mapping.json');

  const aliasMap = {};

  // Seed with built-in aliases.
  for (const [alias, canonical] of Object.entries(DEFAULT_COLUMN_ALIASES)) {
    aliasMap[normalizeHeaderName(alias)] = canonical;
  }

  if (!fs.existsSync(mappingPath)) {
    return aliasMap;
  }

  try {
    const raw = JSON.parse(fs.readFileSync(mappingPath, 'utf8'));

    // Two supported formats:
    // 1) { "Alias Name": "Canonical_Name" }
    // 2) { "Canonical_Name": ["Alias 1", "Alias 2"] }
    for (const [k, v] of Object.entries(raw || {})) {
      if (Array.isArray(v)) {
        for (const alias of v) {
          aliasMap[normalizeHeaderName(alias)] = k;
        }
      } else if (typeof v === 'string') {
        aliasMap[normalizeHeaderName(k)] = v;
      }
    }
  } catch (err) {
    // Ignore invalid mapping file and continue with built-in aliases.
    console.warn('[supplier-reader] Invalid supplier-column-mapping.json:', err.message);
  }

  return aliasMap;
}

const COLUMN_ALIASES = loadColumnAliases();

function mapHeader(rawHeader) {
  const normalized = normalizeHeaderName(rawHeader);
  return COLUMN_ALIASES[normalized] || String(rawHeader || '').replace(/\s*\(.*?\)/g, '').trim();
}

function hasRequiredHeaders(headers, required) {
  const set = new Set((headers || []).filter(Boolean));
  return required.every(h => set.has(h));
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
      if (mapHeader(cell.value) === anchorCol) {
        headerRowNum = rowNum;
        found = true;
      }
    });
  });

  if (!found) {
    return { rows: [], headerRowNum: 1, headerFound: false, headers: [] };
  }

  const headers = [];
  sheet.getRow(headerRowNum).eachCell((cell, colNum) => {
    headers[colNum] = mapHeader(cell.value);
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

  return { rows, headerRowNum, headerFound: true, headers };
}

// ── Legacy single-sheet parser (SUPPLIER_INPUT) ───────────────────────────────

const LEGACY_REQUIRED_COLS = [
  'PO_Number', 'SKU', 'No_of_Cartons', 'Unit_Weight_KG', 'Booking_Qty',
  'Cargo_Ready_Planned_Collection_Date', 'Carrier_Booking_Request_Date',
  'Traffic_Mode',
  'Booking_Group'
];

function parseSingleSheet(sheet) {
  const { rows: rawRows, headerRowNum, headerFound, headers } = readSheet(sheet, 'PO_Number');
  if (!headerFound) {
    throw new Error(`Could not find PO_Number header in sheet "${sheet.name}"`);
  }

  if (!hasRequiredHeaders(headers, ['PO_Number'])) {
    throw new Error(`Sheet "${sheet.name}" is missing required header: PO_Number`);
  }

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

    // Enforce whole number for booking units
    const rawBkq = parseFloat(obj.Header_Booking_Qty);
    if (!isNaN(rawBkq) && rawBkq !== Math.floor(rawBkq)) {
      const rounded = Math.round(rawBkq);
      validationErrors.push(`Row ${obj._rowNum}: Total booked units (${rawBkq}) must be a whole number — rounded to ${rounded}`);
      obj.Header_Booking_Qty = rounded;
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
function parseHeaderOnlySheet(wsHdr) {
  const poRead  = readSheet(wsHdr, 'PO_Number');
  const asnRead = poRead.headerFound ? null : readSheet(wsHdr, 'ASN_Number');
  const source  = poRead.headerFound ? poRead : asnRead;

  if (!source || !source.headerFound) {
    throw new Error(`Could not find PO_Number or ASN_Number header in sheet "${wsHdr.name}"`);
  }

  const { rows: rawRows, headerRowNum, headers } = source;
  const usesAsnAnchor = hasRequiredHeaders(headers, ['ASN_Number']) && !hasRequiredHeaders(headers, ['PO_Number']);

  const requiredCols = usesAsnAnchor ? IDEATEKS_REQUIRED_COLS : REQUIRED_HEADER_COLS;
  if (!hasRequiredHeaders(headers, requiredCols)) {
    const existing = new Set(headers.filter(Boolean));
    const missing = requiredCols.filter(c => !existing.has(c));
    throw new Error(`Sheet "${wsHdr.name}" is missing required headers: ${missing.join(', ')}`);
  }

  const rows = [];
  const validationErrors = [];
  const headerAsnRefs = [];

  for (const obj of rawRows) {
    if (usesAsnAnchor) {
      applyIdeateksDefaults(obj);
      const asn = String(obj.ASN_Number || '').trim();
      if (!asn) continue;
      headerAsnRefs.push(asn);
    } else {
      const po = String(obj.PO_Number || '').trim();
      if (!po) continue;
    }

    const requiredForRow = usesAsnAnchor ? IDEATEKS_REQUIRED_COLS : REQUIRED_HEADER_COLS;
    const missing = requiredForRow.filter(c => !obj[c] || String(obj[c]).trim() === '');
    if (missing.length > 0) {
      validationErrors.push(`Row ${obj._rowNum}: missing required fields: ${missing.join(', ')}`);
    }
    if (String(obj.Collection_Type || '').trim() === 'Collection' &&
        (!obj.Collection_Time || String(obj.Collection_Time).trim() === '')) {
      validationErrors.push(`Row ${obj._rowNum}: Collection_Time is required when Collection_Type is "Collection"`);
    }

    // Enforce whole number for booking units
    const rawBkq = parseFloat(obj.Header_Booking_Qty);
    if (!isNaN(rawBkq) && rawBkq !== Math.floor(rawBkq)) {
      const rounded = Math.round(rawBkq);
      validationErrors.push(`Row ${obj._rowNum}: Total booked units (${rawBkq}) must be a whole number — rounded to ${rounded}`);
      obj.Header_Booking_Qty = rounded;
    }

    rows.push({ ...obj, _headerOnly: true });
  }

  // For Ideateks-style files anchored by ASN, enforce Single Booking per row.
  if (usesAsnAnchor) {
    for (const row of rows) row.Booking_Group = 'Single Booking';
  }

  return {
    rows,
    validationErrors,
    sheetName: wsHdr.name,
    headerRowNum,
    headerPoRefs: [...new Set(rows.map(r => String(r.PO_Number || '').trim()).filter(Boolean))],
    headerAsnRefs: [...new Set(headerAsnRefs.map(a => String(a).trim()).filter(Boolean))]
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

  // Preferred sheet names first, then fall back to any visible sheet that matches headers.
  const preferred = [
    workbook.getWorksheet('PO Header'),
    workbook.getWorksheet('BOOKING_HEADER')
  ].filter(Boolean);
  const visibleSheets = workbook.worksheets.filter(ws => ws.state !== 'veryHidden' && ws.state !== 'hidden');
  const orderedSheets = [...new Set([...preferred, ...visibleSheets])];

  for (const ws of orderedSheets) {
    try {
      return parseHeaderOnlySheet(ws);
    } catch (_) {
      // Try next sheet.
    }
  }

  // Legacy fallback: SUPPLIER_INPUT or first visible sheet
  let sheet = workbook.getWorksheet('SUPPLIER_INPUT');
  if (!sheet) {
    sheet = visibleSheets[0];
  }
  if (!sheet) sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Supplier Excel has no worksheets');

  return parseSingleSheet(sheet);
}

module.exports = { parse };
