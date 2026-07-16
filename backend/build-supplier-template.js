'use strict';

/**
 * Generates samples/Supplier PO sheet-DDMMYYYY.xlsx
 * Single-sheet design:
 *   PO Header — one row per PO (mandatory booking fields + carton totals + defaults)
 *   No PO Lines sheet — SKUs and quantities are auto-booked from Databricks ASN feed.
 * Run: node backend/build-supplier-template.js
 */

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const OUT_DIR  = path.join(__dirname, '..', 'samples');
const OUT_FILE = path.join(OUT_DIR, 'Supplier PO sheet-DDMMYYYY.xlsx');

const CARTON_TYPES = [
  ['BDCM1',                1.40, 60.00, 30.00, 40.00],
  ['BDCM3',                1.00, 45.00, 29.50, 18.80],
  ['C5',                   1.00, 60.00, 30.00, 20.00],
  ['Cartons',              1.00, 45.00, 60.00, 40.00],
  ['A1',                   1.00, 59.50, 28.50, 37.50],
  ['A2',                   1.00, 59.50, 28.50, 32.50],
  ['A3',                   1.00, 59.50, 28.50, 26.00],
  ['A4',                   1.00, 59.50, 28.50, 19.00],
  ['B1',                   1.00, 52.00, 25.50, 37.50],
  ['B2',                   1.00, 52.00, 25.50, 32.50],
  ['B3',                   1.00, 52.00, 25.50, 26.00],
  ['B4',                   1.00, 52.00, 25.50, 19.00],
  ['C1',                   1.00, 45.00, 28.50, 37.50],
  ['C2',                   1.00, 45.00, 28.50, 32.50],
  ['C3',                   1.00, 45.00, 28.50, 26.00],
  ['C4',                   1.00, 45.00, 28.50, 19.00],
  ['Hanging',              0.70, 213.00, 94.00, 60.00],
  ['Hanging Non-Standard', null, null, null, null],
  ['Non-Standard',         null, null, null, null]
];

const BOOKING_GROUP_OPTIONS = [
  'Single Booking',
  ...Array.from({ length: 25 }, (_, i) => 'Multiple POs-BK' + String(i + 1).padStart(3, '0')),
  'Multiple'
];

// Pastel fills used for both column headers (data rows) and the legend.
// Darker shade = header row;  pastel shade = data cell background.
const PALETTE = {
  mandatory: { header: 'FFD9534F', pastel: 'FFFCE8E8', text: 'FF7B1F1F' },
  default:   { header: 'FF5FAD56', pastel: 'FFE8F5E9', text: 'FF1B5E20' },
  auto:      { header: 'FF4A90C4', pastel: 'FFE3F2FD', text: 'FF0D47A1' },
  optional:  { header: 'FF9E9E9E', pastel: 'FFF5F5F5', text: 'FF424242' },
};

/**
 * Writes a README block starting at row 1.
 * Each bullet is its own merged row so the text is fully readable.
 * Returns the row number after the last README row (i.e. the legend row).
 */
function addReadme(ws, sheetTitle, bullets, lastColLetter) {
  let rowNum = 1;

  // Title row
  ws.mergeCells('A' + rowNum + ':' + lastColLetter + rowNum);
  const titleCell = ws.getCell('A' + rowNum);
  titleCell.value = '📋  ' + sheetTitle;
  titleCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  titleCell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  ws.getRow(rowNum).height = 14;
  rowNum++;

  // One row per bullet
  bullets.forEach(bullet => {
    ws.mergeCells('A' + rowNum + ':' + lastColLetter + rowNum);
    const cell = ws.getCell('A' + rowNum);
    cell.value = '  •  ' + bullet;
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEF3FA' } };
    cell.font  = { color: { argb: 'FF1A1A2E' }, size: 9 };
    cell.alignment = { wrapText: false, vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(rowNum).height = 13;
    rowNum++;
  });

  return rowNum; // caller uses this as the legend row number
}

/**
 * Writes the colour-key legend row at `legendRow`.
 * Order: Mandatory → Defaulted → Pre-filled (auto) → Optional
 */
function addLegend(ws, legendRow) {
  const items = [
    { label: '  Mandatory — must be filled  ',  ...PALETTE.mandatory },
    { label: '  Defaulted — editable preset  ',  ...PALETTE.default   },
    { label: '  Pre-filled — auto from lookup  ', ...PALETTE.auto      },
    { label: '  Optional — leave blank if n/a  ', ...PALETTE.optional  },
  ];
  items.forEach(({ label, header }, i) => {
    const cell = ws.getCell(legendRow, i + 1);
    cell.value = label;
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: header } };
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } } };
  });
  ws.getRow(legendRow).height = 14;
  return legendRow + 1; // header row comes next
}

/**
 * Writes column headers at `headerRow`.
 * Order: mandatory → default → auto → optional  (enforced by caller via column array order).
 */
function applyHeaderRow(ws, columns, headerRow) {
  const hr = ws.getRow(headerRow);
  columns.forEach((col, i) => {
    const cell = hr.getCell(i + 1);
    cell.value = col.label;
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'medium', color: { argb: 'FF888888' } } };
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: (PALETTE[col.type] || PALETTE.optional).header } };
  });
  hr.height = 52;  // tall enough for 3–4 lines of wrapped header text
  columns.forEach((col, i) => { ws.getColumn(i + 1).width = col.width; });
  return headerRow + 1; // first data row
}

/**
 * Writes a group-label row that spans consecutive columns of the same type.
 * e.g.  [←─ MANDATORY (8 cols) ─→][←─ DEFAULTED (3) ─→][←─ AUTO-FILLED (4) ─→][←─ OPTIONAL (2) ─→]
 */
function addGroupRow(ws, columns, rowNum) {
  const LABELS = {
    mandatory: '◀  MANDATORY — must be completed  ▶',
    default:   '◀  DEFAULTED — pre-filled, editable  ▶',
    auto:      '◀  AUTO-FILLED — from Carton_Type lookup  ▶',
    optional:  '◀  OPTIONAL — leave blank if not needed  ▶',
  };
  // Build consecutive groups
  const groups = [];
  let cur = null;
  columns.forEach((col, i) => {
    if (!cur || cur.type !== col.type) {
      cur = { type: col.type, start: i + 1, end: i + 1 };
      groups.push(cur);
    } else {
      cur.end = i + 1;
    }
  });
  for (const g of groups) {
    const startLetter = ws.getColumn(g.start).letter;
    const endLetter   = ws.getColumn(g.end).letter;
    if (g.start !== g.end) ws.mergeCells(`${startLetter}${rowNum}:${endLetter}${rowNum}`);
    const cell = ws.getCell(rowNum, g.start);
    cell.value = LABELS[g.type] || g.type.toUpperCase();
    const pal  = PALETTE[g.type] || PALETTE.optional;
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: pal.header } };
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      right:  { style: 'medium', color: { argb: 'FFFFFFFF' } },
      bottom: { style: 'thin',   color: { argb: 'FFCCCCCC' } },
    };
  }
  ws.getRow(rowNum).height = 14;
}

function styleAuto(cell) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALETTE.auto.pastel } };
  cell.font = { color: { argb: PALETTE.auto.text }, size: 10 };
  cell.protection = { locked: true };
}

async function build() {
  if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CarrierBookingStub';
  wb.created = new Date();

  // ── HOW TO USE (first visible tab) ───────────────────────────────────────────
  const wsI = wb.addWorksheet('Instructions to use');
  wsI.properties.tabColor = { argb: 'FFFFA500' };
  const iLastCol = 'H';

  const iTitle = wsI.getRow(1);
  wsI.mergeCells('A1:H1');
  const iTitleCell = wsI.getCell('A1');
  iTitleCell.value = '📋  Carrier Booking — VBKREQ Simulator  |  Instructions to Use';
  iTitleCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  iTitleCell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
  iTitleCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  iTitle.height = 22;

  const iLines = [
    // Overview
    { text: '1.  Fill in one row per PO in the PO Header sheet and email the file to the ASOS L&D Inbound team: InboundService@asos.com', bold: true, indent: 1, bg: 'FFD6E4F0' },
    { text: '', bg: 'FFFFFFFF' },
    // Mandatory fields
    { text: '2.  MANDATORY fields (pink/red columns) — must be filled for every PO row:', bold: true, indent: 1, bg: 'FFFDE8E8' },
    { text: '       •  PO_Number', bold: false, indent: 2, bg: 'FFFEF4F4' },
    { text: '       •  Booking_Group                           (see Booking_Group rules below)', bold: false, indent: 2, bg: 'FFFEF4F4' },
    { text: '       •  Cargo_Ready_Planned_Collection_Date   (DD/MM/YYYY)', bold: false, indent: 2, bg: 'FFFEF4F4' },
    { text: '       •  Carrier_Booking_Request_Date            (DD/MM/YYYY)', bold: false, indent: 2, bg: 'FFFEF4F4' },
    { text: '       •  Total booked units of a booking   (total units for the booking — maps to VBKREQ header BKQ measure)', bold: false, indent: 2, bg: 'FFFEF4F4' },
    { text: '       •  Total no. of Cartons of booking    (whole number > 0  |  default: 1)', bold: false, indent: 2, bg: 'FFFEF4F4' },
    { text: '       •  Total items weight of booking      (kg per individual unit/garment  |  default: 0.21)', bold: false, indent: 2, bg: 'FFFEF4F4' },
    { text: '', bg: 'FFFFFFFF' },
    // Defaulted fields
    { text: '3.  DEFAULTED fields (green columns) — pre-filled with sensible values; update only if different for your shipment:', bold: true, indent: 1, bg: 'FFD6E4F0' },
    { text: '       •  Carton_Type = BDCM1                  (select from dropdown if different)', bold: false, indent: 2, bg: 'FFE8F5E9' },
    { text: '       •  Pack_Type = Bulk Flat', bold: false, indent: 2, bg: 'FFE8F5E9' },
    { text: '       •  Collection_Type = Delivery           (change to "Collection" if carrier collects from factory)', bold: false, indent: 2, bg: 'FFE8F5E9' },
    { text: '       •  Hazardous = N/A', bold: false, indent: 2, bg: 'FFE8F5E9' },
    { text: '       •  Traffic_Mode = CFS                   (change to CY if full container load)', bold: false, indent: 2, bg: 'FFE8F5E9' },
    { text: '', bg: 'FFFFFFFF' },
    // Auto-filled fields
    { text: '4.  AUTO-FILLED fields (blue columns) — calculated automatically from Carton_Type — do NOT edit:', bold: true, indent: 1, bg: 'FFD6E4F0' },
    { text: '       •  Carton_Length_cm   •  Carton_Width_cm   •  Carton_Height_cm   •  Carton_Weight_KG', bold: false, indent: 2, bg: 'FFE3F2FD' },
    { text: '', bg: 'FFFFFFFF' },
    // Optional fields
    { text: '5.  OPTIONAL fields (grey columns) — leave blank if not applicable:', bold: true, indent: 1, bg: 'FFF5F5F5' },
    { text: '       •  Collection_Time (HH:MM)   •  Remarks', bold: false, indent: 2, bg: 'FFF5F5F5' },
    { text: '', bg: 'FFFFFFFF' },
    // Booking_Group rules
    { text: '6.  Booking_Group rules:', bold: true, indent: 1, bg: 'FFE9F3FB' },
    { text: '       •  "Single Booking"                        →  one carrier booking request per PO', bold: false, indent: 2, bg: 'FFE9F3FB' },
    { text: '       •  "Multiple POs-BK001" … "Multiple POs-BK025"  →  POs sharing the same code are merged into one carrier booking request', bold: false, indent: 2, bg: 'FFE9F3FB' },
    { text: '       •  "Multiple"                              →  all POs in the file combined into a single carrier booking request', bold: false, indent: 2, bg: 'FFE9F3FB' },
    { text: '', bg: 'FFFFFFFF' },
    // Notes
    { text: '7.  "Total booked units" maps directly to the VBKREQ header BKQ measure. A report note will flag if it differs from the ASN line-level sum.', bold: false, indent: 1, bg: 'FFFEF3CD' },
    { text: '8.  No_of_Cartons / Unit_Weight_KG / Carton_Type drive the booking-level carton totals (QUR / G / N / VOL) in the carrier booking request.', bold: false, indent: 1, bg: 'FFFEF3CD' },
    { text: '9.  Supplier ID, Factory, Mode of Transport, Ship Date and Expected Delivery are sourced automatically from ASOS systems — do not add them here.', bold: false, indent: 1, bg: 'FFFEF3CD' },
    { text: '', bg: 'FFFFFFFF' },
    // Footer warning
    { text: '⚠  Do NOT rename, move, or delete column headers in the PO Header sheet.', bold: true, indent: 1, bg: 'FFFFF3CD' },
  ];

  iLines.forEach((line, idx) => {
    const rNum = idx + 2;
    wsI.mergeCells(`A${rNum}:H${rNum}`);
    const cell = wsI.getCell(`A${rNum}`);
    cell.value = line.text;
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: line.bg } };
    cell.font  = { bold: line.bold || false, size: 10, color: { argb: 'FF1A1A2E' } };
    cell.alignment = { wrapText: false, vertical: 'middle', horizontal: 'left', indent: line.indent || 1 };
    wsI.getRow(rNum).height = line.text === '' ? 6 : 15;
  });

  // Column widths for the instruction sheet
  ['A','B','C','D','E','F','G','H'].forEach((col, i) => {
    wsI.getColumn(i + 1).width = i === 0 ? 120 : 10;
  });
  wsI.views = [{ showGridLines: false }];

  // CARTON_LOOKUP (hidden)
  const wsC = wb.addWorksheet('CARTON_LOOKUP');
  wsC.state = 'veryHidden';
  wsC.addRow(['Carton_Type', 'Weight_KG', 'Length_cm', 'Width_cm', 'Height_cm']);
  CARTON_TYPES.forEach(ct => wsC.addRow(ct));

  // LISTS_LOOKUP (hidden) — Booking_Group options A2:A28
  const wsL = wb.addWorksheet('LISTS_LOOKUP');
  wsL.state = 'veryHidden';
  wsL.addRow(['Booking_Group']);
  BOOKING_GROUP_OPTIONS.forEach(opt => wsL.addRow([opt]));

  // ── PO Header ─────────────────────────────────────────────────────────────────
  const wsH = wb.addWorksheet('PO Header');
  wsH.properties.tabColor = { argb: 'FF1F4E79' };
  // Columns ordered: Mandatory → Auto-filled → Defaulted → Optional
  const hCols = [
    // Mandatory
    { key: 'PO_Number',                           label: 'PO_Number',                           width: 16, type: 'mandatory' },
    { key: 'Booking_Group',                       label: 'Booking_Group',                       width: 16, type: 'mandatory' },
    { key: 'Cargo_Ready_Planned_Collection_Date', label: 'Cargo_Ready_Planned_Collection_Date', width: 14, type: 'mandatory' },
    { key: 'Carrier_Booking_Request_Date',        label: 'Carrier_Booking_Request_Date',        width: 14, type: 'mandatory' },
    { key: 'Header_Booking_Qty',                  label: 'Total booked units of a booking',     width: 14, type: 'mandatory' },
    { key: 'No_of_Cartons',                       label: 'Total no. of Cartons of booking',     width: 14, type: 'mandatory' },
    { key: 'Unit_Weight_KG',                      label: 'Total items weight of booking',       width: 14, type: 'mandatory' },
    // Auto-fill (from Carton_Type via CARTON_LOOKUP)
    { key: 'Carton_Length_cm', label: 'Carton_Length_cm', width: 12, type: 'auto' },
    { key: 'Carton_Width_cm',  label: 'Carton_Width_cm',  width: 12, type: 'auto' },
    { key: 'Carton_Height_cm', label: 'Carton_Height_cm', width: 12, type: 'auto' },
    { key: 'Carton_Weight_KG', label: 'Carton_Weight_KG', width: 12, type: 'auto' },
    // Defaulted
    { key: 'Carton_Type',      label: 'Carton_Type',      width: 12, type: 'default' },
    { key: 'Pack_Type',        label: 'Pack_Type',        width: 12, type: 'default' },
    { key: 'Collection_Type',  label: 'Collection_Type',  width: 14, type: 'default' },
    { key: 'Hazardous',        label: 'Hazardous',        width: 10, type: 'default' },
    { key: 'Traffic_Mode',     label: 'Traffic_Mode',     width: 12, type: 'default' },
    // Optional
    { key: 'Collection_Time',  label: 'Collection_Time (HH:MM)', width: 14, type: 'optional' },
    { key: 'Remarks',          label: 'Remarks',                 width: 18, type: 'optional' },
  ];
  const hLastCol = wsH.getColumn(hCols.length).letter;
  const hLegendRow  = 1;
  const hGroupRow   = addLegend(wsH, hLegendRow);   // row 2 = group label bar
  addGroupRow(wsH, hCols, hGroupRow);               // row 2 = ◀ MANDATORY ▶ | ◀ DEFAULTED ▶ | ...
  const hHeaderRow  = hGroupRow + 1;                // row 3 = column headers
  const hFirstDataR = applyHeaderRow(wsH, hCols, hHeaderRow); // row 4+ = data
  // Centre-align all data columns; include wrapText so header cells retain their wrapping
  hCols.forEach((_, i) => {
    wsH.getColumn(i + 1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  // Re-assert wrapText on header cells explicitly (column alignment can override cell level in some ExcelJS versions)
  const hHdrRow = wsH.getRow(hHeaderRow);
  hCols.forEach((_, i) => {
    const cell = hHdrRow.getCell(i + 1);
    cell.alignment = { ...cell.alignment, horizontal: 'center', vertical: 'middle', wrapText: true };
  });
  const hIdx = {};
  hCols.forEach((c, i) => { hIdx[c.key] = i + 1; });
  const hLet = n => wsH.getColumn(n).letter;
  wsH.getColumn(hIdx['PO_Number']).numFmt = '@'; // prevent scientific notation
  for (let r = hFirstDataR; r <= hFirstDataR + 499; r++) {
    const row = wsH.getRow(r);
    row.getCell(hIdx['Traffic_Mode']).value      = 'CFS';
    row.getCell(hIdx['Booking_Group']).value     = 'Single Booking';
    row.getCell(hIdx['Pack_Type']).value         = 'Bulk Flat';
    row.getCell(hIdx['Collection_Type']).value   = 'Delivery';
    row.getCell(hIdx['Hazardous']).value         = 'N/A';
    row.getCell(hIdx['No_of_Cartons']).value     = 1;
    row.getCell(hIdx['Unit_Weight_KG']).value    = 0.21;
    row.getCell(hIdx['Carton_Type']).value       = 'BDCM1';
    const hCt = hLet(hIdx['Carton_Type']) + r;
    row.getCell(hIdx['Carton_Length_cm']).value  = { formula: 'IFERROR(VLOOKUP(' + hCt + ',CARTON_LOOKUP!$A:$E,3,0),"")' };
    row.getCell(hIdx['Carton_Width_cm']).value   = { formula: 'IFERROR(VLOOKUP(' + hCt + ',CARTON_LOOKUP!$A:$E,4,0),"")' };
    row.getCell(hIdx['Carton_Height_cm']).value  = { formula: 'IFERROR(VLOOKUP(' + hCt + ',CARTON_LOOKUP!$A:$E,5,0),"")' };
    row.getCell(hIdx['Carton_Weight_KG']).value  = { formula: 'IFERROR(VLOOKUP(' + hCt + ',CARTON_LOOKUP!$A:$E,2,0),"")' };
    ['Carton_Length_cm','Carton_Width_cm','Carton_Height_cm','Carton_Weight_KG'].forEach(k => styleAuto(row.getCell(hIdx[k])));
    ['Carton_Length_cm','Carton_Width_cm','Carton_Height_cm'].forEach(k => (row.getCell(hIdx[k]).numFmt = '0.00'));
    row.getCell(hIdx['Carton_Weight_KG']).numFmt = '0.0000';
    row.getCell(hIdx['Unit_Weight_KG']).numFmt   = '0.0000';
    ['Cargo_Ready_Planned_Collection_Date','Carrier_Booking_Request_Date'].forEach(k => {
      row.getCell(hIdx[k]).numFmt = 'DD/MM/YYYY';
    });
    row.getCell(hIdx['PO_Number']).numFmt = '@'; // cell-level: forces Excel to respect text format on paste
    row.commit();
  }
  for (let r = hFirstDataR; r <= hFirstDataR + 499; r++) {
    wsH.getCell(r, hIdx['PO_Number']).dataValidation = {
      type: 'textLength', operator: 'greaterThan', formulae: [0], allowBlank: false,
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Required', error: 'PO_Number cannot be blank'
    };
    ['Cargo_Ready_Planned_Collection_Date','Carrier_Booking_Request_Date'].forEach(k => {
      wsH.getCell(r, hIdx[k]).dataValidation = {
        type: 'date', operator: 'greaterThan', formulae: [new Date(2020, 0, 1)],
        showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid date', error: 'Enter a valid date (DD/MM/YYYY)',
        showInputMessage: true, promptTitle: 'Date format', prompt: 'DD/MM/YYYY'
      };
    });
    wsH.getCell(r, hIdx['Traffic_Mode']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['"CFS,CY"'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Traffic Mode', error: 'Select CFS or CY'
    };
    wsH.getCell(r, hIdx['Booking_Group']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['LISTS_LOOKUP!$A$2:$A$28'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Booking Group', error: 'Select a Booking Group from the list'
    };
    wsH.getCell(r, hIdx['Pack_Type']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['"Flat,Bulk Flat,Hanging"'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Pack Type', error: 'Select Flat, Bulk Flat or Hanging'
    };
    wsH.getCell(r, hIdx['Collection_Type']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['"Collection,Delivery"'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Collection Type', error: 'Select Collection or Delivery'
    };
    wsH.getCell(r, hIdx['Hazardous']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['"Flammable,Glass - Hazardous,Hazardous,N/A"'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Hazardous', error: 'Select a hazardous classification'
    };
    wsH.getCell(r, hIdx['No_of_Cartons']).dataValidation = {
      type: 'whole', operator: 'greaterThan', formulae: [0], allowBlank: false,
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'No of Cartons required',
      error: 'Enter a whole number greater than 0'
    };
    wsH.getCell(r, hIdx['Unit_Weight_KG']).dataValidation = {
      type: 'decimal', operator: 'greaterThan', formulae: [0], allowBlank: false,
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Unit Weight required',
      error: 'Enter a positive weight in KG (e.g. 0.21)'
    };
    wsH.getCell(r, hIdx['Carton_Type']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['CARTON_LOOKUP!$A$2:$A$20'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Carton Type', error: 'Select a carton type from the list'
    };
    const tc = hLet(hIdx['Collection_Time']) + r;
    wsH.getCell(r, hIdx['Collection_Time']).dataValidation = {
      type: 'custom',
      formulae: ['OR(' + tc + '="",AND(LEN(' + tc + ')=5,MID(' + tc + ',3,1)=":",ISNUMBER(VALUE(LEFT(' + tc + ',2))),ISNUMBER(VALUE(RIGHT(' + tc + ',2))),VALUE(LEFT(' + tc + ',2))<=23,VALUE(RIGHT(' + tc + ',2))<=59))'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Time', error: 'Enter time as HH:MM (e.g. 09:30) or leave blank'
    };
  }
  wsH.views = [{ state: 'frozen', ySplit: hFirstDataR - 1, xSplit: 1, showGridLines: true }];

  await wb.xlsx.writeFile(OUT_FILE);
  console.log('✅  SupplierInput_template.xlsx written to:\n    ' + OUT_FILE);
  console.log('\nSheets:  PO Header only (500 rows) — no PO Lines sheet');
  console.log('\nPO Header mandatory:', hCols.filter(c => c.type === 'mandatory').map(c => c.label).join(', '));
  console.log('PO Header defaults :', hCols.filter(c => c.type === 'default').map(c => c.label).join(', '));
  console.log('PO Header auto-fill:', hCols.filter(c => c.type === 'auto').map(c => c.label).join(', '));
  console.log('PO Header optional :', hCols.filter(c => c.type === 'optional').map(c => c.label).join(', '));
}

build().catch(err => {
  console.error('FAILED:', err.message);
  process.exit(1);
});
