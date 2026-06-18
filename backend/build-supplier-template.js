'use strict';

/**
 * Generates samples/SupplierInput_template.xlsx
 * Two-sheet design:
 *   PO Header — one row per PO (header-level fields + factory auto-fill)
 *   PO Lines  — one row per SKU (qty / weight / carton fields)
 * Run: node backend/build-supplier-template.js
 */

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const OUT_DIR  = path.join(__dirname, '..', 'samples');
const OUT_FILE = path.join(OUT_DIR, 'SupplierInput_template.xlsx');

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
  hr.height = 22;
  columns.forEach((col, i) => { ws.getColumn(i + 1).width = col.width; });
  return headerRow + 1; // first data row
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
    // Point 1
    { text: '1.  The supplier must complete two sheets — PO Header and PO Lines —', bold: true, indent: 1, bg: 'FFD6E4F0' },
    { text: '     with the PO Number used as the primary key to link both sheets.', bold: false, indent: 2, bg: 'FFE9F3FB' },
    { text: '', bg: 'FFFFFFFF' },
    // Point 2
    { text: '2.  All mandatory fields must be completed; otherwise, booking creation will fail.', bold: true, indent: 1, bg: 'FFFDE8E8' },
    { text: '     If a field is unclear or you prefer not to provide a value, leave the default pre-filled value where applicable.', bold: false, indent: 2, bg: 'FFFEF4F4' },
    { text: '', bg: 'FFFFFFFF' },
    // Point 3
    { text: '3.  Enter one row per PO in the PO Header sheet. Ensure all mandatory (pink) columns are filled for each PO.', bold: true, indent: 1, bg: 'FFD6E4F0' },
    { text: '', bg: 'FFFFFFFF' },
    // Booking_Group rules
    { text: '     Booking_Group rules:', bold: true, indent: 2, bg: 'FFE9F3FB' },
    { text: '       "Single Booking"                          →  one PO per booking', bold: false, indent: 3, bg: 'FFE9F3FB' },
    { text: '       "Multiple POs-BK001" … "Multiple POs-BK025"  →  POs with the same code (BK001, BK002 …) are combined into one booking', bold: false, indent: 3, bg: 'FFE9F3FB' },
    { text: '       "Multiple"                                →  all POs are combined into a single booking', bold: false, indent: 3, bg: 'FFE9F3FB' },
    { text: '', bg: 'FFFFFFFF' },
    // Point 4
    { text: '4.  Dates must be in DD/MM/YYYY format.', bold: true, indent: 1, bg: 'FFFDE8E8' },
    { text: '', bg: 'FFFFFFFF' },
    // Point 5 (Factory)
    { text: '5.  Factory_ID: select from the dropdown — it is based on the FACTORY_LOOKUP tab (last tab). Factory_Name and address auto-fill when you pick an ID.', bold: false, indent: 1, bg: 'FFE9F3FB' },
    { text: '     Default Factory_ID is 9999 (Dummy Factory) — no address details are required for this value.', bold: false, indent: 2, bg: 'FFE9F3FB' },
    { text: '', bg: 'FFFFFFFF' },
    // Point 6
    { text: '6.  In the PO Lines sheet, use multiple rows for POs that contain multiple SKU lines.', bold: true, indent: 1, bg: 'FFD6E4F0' },
    { text: '     Each row must have the same PO_Number as the corresponding header row in PO Header.', bold: false, indent: 2, bg: 'FFE9F3FB' },
    { text: '', bg: 'FFFFFFFF' },
    // Colour guide
    { text: '🎨  Colour guide:', bold: true, indent: 1, bg: 'FFF5F5F5' },
    { text: '  🔴  Pink / Red  — MANDATORY field (must be filled)', bold: false, indent: 2, bg: 'FFFCE8E8' },
    { text: '  🟢  Green       — DEFAULTED field (pre-set, editable)', bold: false, indent: 2, bg: 'FFE8F5E9' },
    { text: '  🔵  Blue        — AUTO-FILLED field (formula / lookup — do not edit)', bold: false, indent: 2, bg: 'FFE3F2FD' },
    { text: '  ⚪  Grey        — OPTIONAL field (leave blank if not applicable)', bold: false, indent: 2, bg: 'FFF5F5F5' },
    { text: '', bg: 'FFFFFFFF' },
    // Footer note
    { text: '⚠  Do NOT modify or delete column headers on any sheet.', bold: true, indent: 1, bg: 'FFFFF3CD' },
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
  // Columns ordered: Mandatory → Defaulted → Pre-filled (auto) → Optional
  const hCols = [
    // Mandatory
    { key: 'PO_Number',                           label: 'PO_Number',                           width: 22, type: 'mandatory' },
    { key: 'Cargo_Ready_Planned_Collection_Date', label: 'Cargo_Ready_Planned_Collection_Date', width: 34, type: 'mandatory' },
    { key: 'Carrier_Booking_Request_Date',        label: 'Carrier_Booking_Request_Date',        width: 28, type: 'mandatory' },
    { key: 'Traffic_Mode',                        label: 'Traffic_Mode',                        width: 14, type: 'mandatory' },
    { key: 'Mode_Of_Transport',                   label: 'Mode_Of_Transport',                   width: 18, type: 'mandatory' },
    { key: 'Booking_Group',                       label: 'Booking_Group',                       width: 30, type: 'mandatory' },
    { key: 'Factory_ID',                          label: 'Factory_ID',                          width: 18, type: 'mandatory' },
    // Defaulted
    { key: 'Pack_Type',                           label: 'Pack_Type',                           width: 14, type: 'default' },
    { key: 'Collection_Type',                     label: 'Collection_Type',                     width: 18, type: 'default' },
    { key: 'Hazardous',                           label: 'Hazardous',                           width: 20, type: 'default' },
    // Pre-filled (auto via VLOOKUP)
    { key: 'Factory_Name',                        label: 'Factory_Name',                        width: 28, type: 'auto' },
    { key: 'Factory_Street1',                     label: 'Factory_Street1',                     width: 30, type: 'auto' },
    { key: 'Factory_Street2',                     label: 'Factory_Street2',                     width: 24, type: 'auto' },
    { key: 'Factory_City',                        label: 'Factory_City',                        width: 20, type: 'auto' },
    { key: 'Factory_PostalCd',                    label: 'Factory_PostalCd',                    width: 16, type: 'auto' },
    { key: 'Factory_CountryCd',                   label: 'Factory_CountryCd',                   width: 16, type: 'auto' },
    // Optional
    { key: 'Collection_Time',                     label: 'Collection_Time (HH:MM)',             width: 24, type: 'optional' },
    { key: 'Remarks',                             label: 'Remarks',                             width: 30, type: 'optional' },
  ];
  const hLastCol = wsH.getColumn(hCols.length).letter;
  const hLegendRow  = 1;
  const hHeaderRow  = addLegend(wsH, hLegendRow);
  const hFirstDataR = applyHeaderRow(wsH, hCols, hHeaderRow);
  const hIdx = {};
  hCols.forEach((c, i) => { hIdx[c.key] = i + 1; });
  const hLet = n => wsH.getColumn(n).letter;
  for (let r = hFirstDataR; r <= hFirstDataR + 49; r++) {
    const row = wsH.getRow(r);
    row.getCell(hIdx['Traffic_Mode']).value      = 'CFS';
    row.getCell(hIdx['Mode_Of_Transport']).value = 'Road';
    row.getCell(hIdx['Booking_Group']).value     = 'Single Booking';
    row.getCell(hIdx['Pack_Type']).value         = 'Bulk Flat';
    row.getCell(hIdx['Collection_Type']).value   = 'Delivery';
    row.getCell(hIdx['Hazardous']).value         = 'N/A';
    row.getCell(hIdx['Factory_ID']).value        = '9999';
    const fid = '$' + hLet(hIdx['Factory_ID']) + r;
    // Factory_Name: 9999 → "Dummy Factory", else VLOOKUP
    row.getCell(hIdx['Factory_Name']).value = {
      formula: 'IF(' + fid + '="9999","Dummy Factory",IFERROR(VLOOKUP(' + fid + ',FACTORY_LOOKUP!$A:$G,2,0),""))'
    };
    styleAuto(row.getCell(hIdx['Factory_Name']));
    // Address fields: 9999 → blank, else VLOOKUP
    [['Factory_Street1',3],['Factory_Street2',4],
     ['Factory_City',5],['Factory_PostalCd',6],['Factory_CountryCd',7]].forEach(([k, col]) => {
      row.getCell(hIdx[k]).value = {
        formula: 'IF(' + fid + '="9999","",IFERROR(VLOOKUP(' + fid + ',FACTORY_LOOKUP!$A:$G,' + col + ',0),""))'
      };
      styleAuto(row.getCell(hIdx[k]));
    });
    ['Cargo_Ready_Planned_Collection_Date','Carrier_Booking_Request_Date'].forEach(k => {
      row.getCell(hIdx[k]).numFmt = 'DD/MM/YYYY';
    });
    row.commit();
  }
  for (let r = hFirstDataR; r <= hFirstDataR + 49; r++) {
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
    wsH.getCell(r, hIdx['Mode_Of_Transport']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['"Sea,Air,Road,Rail,Eco"'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Mode of Transport', error: 'Select Sea, Air, Road, Rail or Eco'
    };
    wsH.getCell(r, hIdx['Booking_Group']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['LISTS_LOOKUP!$A$2:$A$28'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Booking Group', error: 'Select a Booking Group from the list'
    };
    wsH.getCell(r, hIdx['Factory_ID']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['FACTORY_LOOKUP!$A$3:$A$100'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Factory ID', error: 'Select a Factory ID from the FACTORY_LOOKUP tab'
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
    const tc = hLet(hIdx['Collection_Time']) + r;
    wsH.getCell(r, hIdx['Collection_Time']).dataValidation = {
      type: 'custom',
      formulae: ['OR(' + tc + '="",AND(LEN(' + tc + ')=5,MID(' + tc + ',3,1)=":",ISNUMBER(VALUE(LEFT(' + tc + ',2))),ISNUMBER(VALUE(RIGHT(' + tc + ',2))),VALUE(LEFT(' + tc + ',2))<=23,VALUE(RIGHT(' + tc + ',2))<=59))'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Time', error: 'Enter time as HH:MM (e.g. 09:30) or leave blank'
    };
  }
  // Cross-sheet warning: PO_Number in PO Header has no matching rows in PO Lines → orange
  const hPoCol = hLet(hIdx['PO_Number']);
  wsH.addConditionalFormatting({
    ref: `${hPoCol}${hFirstDataR}:${hPoCol}${hFirstDataR + 49}`,
    rules: [{
      type: 'expression',
      formulae: [`AND($${hPoCol}${hFirstDataR}<>"",COUNTIF('PO Lines'!$${hPoCol}:$${hPoCol},$${hPoCol}${hFirstDataR})=0)`],
      style: {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFD966' } },
        font: { color: { argb: 'FF7F6000' }, bold: true }
      }
    }]
  });
  wsH.views = [{ state: 'frozen', ySplit: hFirstDataR - 1, xSplit: 1, showGridLines: true }];

  // ── PO Lines ──────────────────────────────────────────────────────────────────
  const wsS = wb.addWorksheet('PO Lines');
  wsS.properties.tabColor = { argb: 'FF7B2C2C' };
  // Columns ordered: Mandatory → Defaulted → Pre-filled (auto) → Optional
  const sCols = [
    // Mandatory
    { key: 'PO_Number',        label: 'PO_Number',        width: 22, type: 'mandatory' },
    { key: 'SKU',              label: 'SKU',              width: 22, type: 'mandatory' },
    { key: 'Booking_Qty',      label: 'Booking_Qty',      width: 14, type: 'mandatory' },
    { key: 'No_of_Cartons',    label: 'No_of_Cartons',    width: 16, type: 'mandatory' },
    { key: 'Unit_Weight_KG',   label: 'Unit_Weight_KG',   width: 16, type: 'mandatory' },
    // Defaulted
    { key: 'Carton_Type',      label: 'Carton_Type',      width: 22, type: 'default' },
    // Pre-filled (auto-calc)
    { key: 'Carton_Length_cm', label: 'Carton_Length_cm', width: 18, type: 'auto' },
    { key: 'Carton_Width_cm',  label: 'Carton_Width_cm',  width: 17, type: 'auto' },
    { key: 'Carton_Height_cm', label: 'Carton_Height_cm', width: 17, type: 'auto' },
    { key: 'Carton_Weight_KG', label: 'Carton_Weight_KG', width: 18, type: 'auto' },
    { key: 'Gross_Weight_KG',  label: 'Gross_Weight_KG',  width: 17, type: 'auto' },
    { key: 'Net_Weight_KG',    label: 'Net_Weight_KG',    width: 16, type: 'auto' },
    { key: 'Volume_M3',        label: 'Volume_M3',        width: 13, type: 'auto' },
    // Optional
    { key: 'EAN_Barcode',      label: 'EAN_Barcode',      width: 18, type: 'optional' },
    { key: 'Colour_Code',      label: 'Colour_Code',      width: 14, type: 'optional' },
    { key: 'Size_Code',        label: 'Size_Code',        width: 12, type: 'optional' },
  ];
  const sLastCol = wsS.getColumn(sCols.length).letter;
  const sLegendRow  = 1;
  const sHeaderRow  = addLegend(wsS, sLegendRow);
  const sFirstDataR = applyHeaderRow(wsS, sCols, sHeaderRow);
  const sIdx = {};
  sCols.forEach((c, i) => { sIdx[c.key] = i + 1; });
  const sLet = n => wsS.getColumn(n).letter;
  for (let r = sFirstDataR; r <= sFirstDataR + 199; r++) {
    const row = wsS.getRow(r);
    row.getCell(sIdx['Carton_Type']).value = 'BDCM1';
    const ct = sLet(sIdx['Carton_Type']) + r;
    row.getCell(sIdx['Carton_Length_cm']).value = { formula: 'IFERROR(VLOOKUP(' + ct + ',CARTON_LOOKUP!$A:$E,3,0),"")' };
    row.getCell(sIdx['Carton_Width_cm']).value  = { formula: 'IFERROR(VLOOKUP(' + ct + ',CARTON_LOOKUP!$A:$E,4,0),"")' };
    row.getCell(sIdx['Carton_Height_cm']).value = { formula: 'IFERROR(VLOOKUP(' + ct + ',CARTON_LOOKUP!$A:$E,5,0),"")' };
    row.getCell(sIdx['Carton_Weight_KG']).value = { formula: 'IFERROR(VLOOKUP(' + ct + ',CARTON_LOOKUP!$A:$E,2,0),"")' };
    const noC  = sLet(sIdx['No_of_Cartons'])  + r;
    const unitW= sLet(sIdx['Unit_Weight_KG']) + r;
    const bkq  = sLet(sIdx['Booking_Qty'])    + r;
    const cL   = sLet(sIdx['Carton_Length_cm'])  + r;
    const cW   = sLet(sIdx['Carton_Width_cm'])   + r;
    const cH   = sLet(sIdx['Carton_Height_cm'])  + r;
    const cWt  = sLet(sIdx['Carton_Weight_KG'])  + r;
    row.getCell(sIdx['Gross_Weight_KG']).value = { formula: 'IFERROR(' + cWt + '*' + noC + ',0)' };
    row.getCell(sIdx['Net_Weight_KG']).value   = { formula: 'IFERROR(' + unitW + '*' + bkq + ',0)' };
    row.getCell(sIdx['Volume_M3']).value        = { formula: 'IFERROR((' + cL + '*' + cW + '*' + cH + '/1000000)*' + noC + ',0)' };
    ['Carton_Length_cm','Carton_Width_cm','Carton_Height_cm','Carton_Weight_KG',
     'Gross_Weight_KG','Net_Weight_KG','Volume_M3'].forEach(k => styleAuto(row.getCell(sIdx[k])));
    ['Carton_Length_cm','Carton_Width_cm','Carton_Height_cm'].forEach(k => (row.getCell(sIdx[k]).numFmt = '0.00'));
    ['Carton_Weight_KG','Gross_Weight_KG','Net_Weight_KG','Unit_Weight_KG'].forEach(k => (row.getCell(sIdx[k]).numFmt = '0.0000'));
    row.getCell(sIdx['Volume_M3']).numFmt = '0.0000';
    row.commit();
  }
  for (let r = sFirstDataR; r <= sFirstDataR + 199; r++) {
    wsS.getCell(r, sIdx['PO_Number']).dataValidation = {
      type: 'textLength', operator: 'greaterThan', formulae: [0], allowBlank: false,
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Required', error: 'PO_Number cannot be blank'
    };
    wsS.getCell(r, sIdx['SKU']).dataValidation = {
      type: 'textLength', operator: 'greaterThan', formulae: [0], allowBlank: false,
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Required', error: 'SKU cannot be blank'
    };
    wsS.getCell(r, sIdx['Booking_Qty']).dataValidation = {
      type: 'whole', operator: 'greaterThan', formulae: [0],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Booking Qty', error: 'Booking_Qty must be a whole number greater than 0'
    };
    wsS.getCell(r, sIdx['No_of_Cartons']).dataValidation = {
      type: 'whole', operator: 'greaterThan', formulae: [0],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid No of Cartons', error: 'Enter a whole number greater than 0'
    };
    wsS.getCell(r, sIdx['Unit_Weight_KG']).dataValidation = {
      type: 'decimal', operator: 'greaterThan', formulae: [0],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Weight', error: 'Enter a positive weight in KG'
    };
    wsS.getCell(r, sIdx['Carton_Type']).dataValidation = {
      type: 'list', allowBlank: false, formulae: ['CARTON_LOOKUP!$A$2:$A$20'],
      showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Invalid Carton Type', error: 'Select a carton type from the list'
    };
  }
  // Cross-sheet warning: PO_Number in PO Lines not found in PO Header → red
  const sPoCol = sLet(sIdx['PO_Number']);
  wsS.addConditionalFormatting({
    ref: `${sPoCol}${sFirstDataR}:${sPoCol}${sFirstDataR + 199}`,
    rules: [{
      type: 'expression',
      formulae: [`AND($${sPoCol}${sFirstDataR}<>"",COUNTIF('PO Header'!$${sPoCol}:$${sPoCol},$${sPoCol}${sFirstDataR})=0)`],
      style: {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } },
        font: { color: { argb: 'FF9C0006' }, bold: true }
      }
    }]
  });
  wsS.views = [{ state: 'frozen', ySplit: sFirstDataR - 1, xSplit: 1, showGridLines: true }];

  // ── FACTORY_LOOKUP (last tab — visible) ───────────────────────────────────────
  const wsF = wb.addWorksheet('FACTORY_LOOKUP');
  wsF.properties.tabColor = { argb: 'FF1F6F5F' };
  wsF.mergeCells('A1:G1');
  const fb = wsF.getCell('A1');
  fb.value = 'FACTORY LOOKUP TABLE — Add one row per factory. Factory_ID appears as a dropdown on PO Header. Name and address auto-fill when you select an ID. Do NOT change column headers.';
  fb.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAD3' } };
  fb.font  = { bold: true, color: { argb: 'FF0B4F2A' }, size: 11 };
  fb.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
  wsF.getRow(1).height = 36;
  const fHdrs = ['Factory_ID','Factory_Name','Factory_Street1','Factory_Street2','Factory_City','Factory_PostalCd','Factory_CountryCd'];
  const fWidths = [18,30,32,28,20,16,16];
  const fhr = wsF.getRow(2);
  fHdrs.forEach((h, i) => {
    const c = fhr.getCell(i + 1);
    c.value = h;
    c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF274E13' } };
    c.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    wsF.getColumn(i + 1).width = fWidths[i];
  });
  fhr.height = 28;
  wsF.views = [{ state: 'frozen', ySplit: 2, showGridLines: true }];
  // Pre-seed: 9999 = Dummy Factory (no address required)
  const dummyRow = wsF.addRow(['9999', 'Dummy Factory', '', '', '', '', '']);
  dummyRow.getCell(1).font = { bold: true, color: { argb: 'FF7F7F7F' }, italic: true };
  dummyRow.getCell(2).font = { bold: true, color: { argb: 'FF7F7F7F' }, italic: true };

  await wb.xlsx.writeFile(OUT_FILE);
  console.log('✅  SupplierInput_template.xlsx written to:\n    ' + OUT_FILE);
  console.log('\nSheets:  PO Header (50 rows)  |  PO Lines (200 rows)  |  FACTORY_LOOKUP');
  console.log('\nPO Header mandatory:', hCols.filter(c => c.type === 'mandatory').map(c => c.label).join(', '));
  console.log('PO Header defaults :', hCols.filter(c => c.type === 'default').map(c => c.label).join(', '));
  console.log('PO Header auto-fill:', hCols.filter(c => c.type === 'auto').map(c => c.label).join(', '));
  console.log('\nPO Lines mandatory:', sCols.filter(c => c.type === 'mandatory').map(c => c.label).join(', '));
  console.log('PO Lines defaults :', sCols.filter(c => c.type === 'default').map(c => c.label).join(', '));
  console.log('PO Lines auto-calc:', sCols.filter(c => c.type === 'auto').map(c => c.label).join(', '));
  console.log('PO Lines optional :', sCols.filter(c => c.type === 'optional').map(c => c.label).join(', '));
}

build().catch(err => {
  console.error('FAILED:', err.message);
  process.exit(1);
});
