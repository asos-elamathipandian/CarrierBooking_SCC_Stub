'use strict';

/**
 * Generates samples/SupplierInput_template.xlsx
 * — the file suppliers fill in and upload via Step 1.
 * Run: node backend/build-supplier-template.js
 */

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const OUT_DIR  = path.join(__dirname, '..', 'samples');
const OUT_FILE = path.join(OUT_DIR, 'SupplierInput_template.xlsx');

// All 19 carton types (name, weight, L, W, H)
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
  ['Hanging Non-Standard', null, null,  null,  null],
  ['Non-Standard',         null, null,  null,  null]
];

async function build() {
  if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

  const wb = new ExcelJS.Workbook();
  wb.creator = 'CarrierBookingStub';
  wb.created = new Date();

  // ── Hidden sheet: CARTON_LOOKUP (data source for dropdowns / VLOOKUPs) ──────
  const wsLookup = wb.addWorksheet('CARTON_LOOKUP');
  wsLookup.state = 'veryHidden';
  wsLookup.addRow(['Carton_Type','Weight_KG','Length_cm','Width_cm','Height_cm']);
  CARTON_TYPES.forEach(ct => wsLookup.addRow(ct));

  // ── Main sheet: SUPPLIER_INPUT ───────────────────────────────────────────────
  const ws = wb.addWorksheet('SUPPLIER_INPUT');

  // ── Row 1: Instructions banner ───────────────────────────────────────────────
  ws.mergeCells('A1:AB1');
  const banner = ws.getCell('A1');
  banner.value =
    '⚠  INSTRUCTIONS: One row per SKU — fill in all RED columns for every SKU line. ' +
    'Multiple rows can share the same PO_Number/ASN_Ref. ' +
    'No_of_Cartons and Unit_Weight_KG must be filled per SKU row. ' +
    'Carton dimensions auto-fill from Carton_Type. Dates in DD/MM/YYYY format. Do NOT modify column headers.';
  banner.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  banner.font   = { bold: true, color: { argb: 'FF7B3F00' }, size: 11 };
  banner.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
  ws.getRow(1).height = 32;

  // ── Row 2: Colour-key legend (one cell each, no merging) ─────────────────────
  const legendItems = [
    { col: 1, label: ' RED = Mandatory ',           argb: 'FFC0392B' },
    { col: 2, label: ' GREEN = Has default ',        argb: 'FF27AE60' },
    { col: 3, label: ' BLUE = Auto-calculated ',     argb: 'FF2E75B6' },
    { col: 4, label: ' GREY = Optional ',            argb: 'FF555555' },
  ];
  legendItems.forEach(({ col, label, argb }) => {
    const cell = ws.getCell(2, col);
    cell.value = label;
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
  });
  ws.getRow(2).height = 22;

  // ── Row 3: Column headers ────────────────────────────────────────────────────
  const columns = [
    // mandatory (red)
    { key: 'PO_Number',                           label: 'PO_Number',                           width: 22, type: 'mandatory' },
    { key: 'ASN_Ref',                             label: 'ASN_Ref',                             width: 22, type: 'mandatory' },
    { key: 'SKU',                                 label: 'SKU',                                 width: 22, type: 'mandatory' },
    { key: 'No_of_Cartons',                       label: 'No_of_Cartons',                       width: 16, type: 'mandatory' },
    { key: 'Unit_Weight_KG',                      label: 'Unit_Weight_KG',                      width: 16, type: 'mandatory' },
    { key: 'Cargo_Ready_Planned_Collection_Date', label: 'Cargo_Ready_Planned_Collection_Date', width: 34, type: 'mandatory' },
    { key: 'Carrier_Booking_Request_Date',        label: 'Carrier_Booking_Request_Date',        width: 28, type: 'mandatory' },
    { key: 'Traffic_Mode',                        label: 'Traffic_Mode',                        width: 14, type: 'mandatory' },
    // optional supplier fills
    { key: 'EAN_Barcode',    label: 'EAN_Barcode',    width: 18, type: 'optional' },
    { key: 'Colour_Code',    label: 'Colour_Code',    width: 14, type: 'optional' },
    { key: 'Size_Code',      label: 'Size_Code',      width: 12, type: 'optional' },
    { key: 'Booking_Qty',    label: 'Booking_Qty',    width: 14, type: 'mandatory' },
    // carton (default + dropdown)
    { key: 'Carton_Type',    label: 'Carton_Type',    width: 22, type: 'default', default: 'BDCM1' },
    // auto dims (VLOOKUP)
    { key: 'Carton_Length_cm', label: 'Carton_Length_cm', width: 18, type: 'auto' },
    { key: 'Carton_Width_cm',  label: 'Carton_Width_cm',  width: 17, type: 'auto' },
    { key: 'Carton_Height_cm', label: 'Carton_Height_cm', width: 17, type: 'auto' },
    { key: 'Carton_Weight_KG', label: 'Carton_Weight_KG', width: 18, type: 'auto' },
    // auto-calc
    { key: 'Gross_Weight_KG',  label: 'Gross_Weight_KG',  width: 17, type: 'auto' },
    { key: 'Net_Weight_KG',    label: 'Net_Weight_KG',    width: 16, type: 'auto' },
    { key: 'Volume_M3',        label: 'Volume_M3',        width: 13, type: 'auto' },
    // dropdowns with defaults
    { key: 'Pack_Type',        label: 'Pack_Type',        width: 14, type: 'default', default: 'Flat' },
    { key: 'Collection_Type',  label: 'Collection_Type',  width: 18, type: 'default', default: 'Delivery' },
    // Collection_Time becomes mandatory when Collection_Type = 'Collection' (HH:MM format)
    { key: 'Collection_Time',  label: 'Collection_Time (HH:MM)', width: 24, type: 'optional' },
    { key: 'Hazardous',        label: 'Hazardous',        width: 20, type: 'default', default: 'N/A' },
    // optional dates
    { key: 'Expected_Delivery_Date', label: 'Expected_Delivery_Date', width: 24, type: 'optional' },
    { key: 'ASN_Delivery_Date',      label: 'ASN_Delivery_Date',      width: 20, type: 'optional' },
    // misc
    { key: 'ASOS_Intake_Week', label: 'ASOS_Intake_Week', width: 18, type: 'optional' },
    { key: 'Var_Unit', label: 'Var_Unit', width: 11, type: 'default', default: 0 },
    { key: 'Var_Pct',  label: 'Var_Pct',  width: 11, type: 'default', default: 0 },
    { key: 'Remarks',  label: 'Remarks',  width: 30, type: 'optional' }
  ];

  const headerRow = ws.getRow(3);
  columns.forEach((col, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = col.label;
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { bottom: { style: 'medium', color: { argb: 'FFAAAAAA' } } };

    if (col.type === 'mandatory') {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0392B' } };
    } else if (col.type === 'auto') {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
    } else if (col.type === 'default') {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF27AE60' } };
    } else {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF555555' } };
    }
  });
  headerRow.height = 36;

  // ── Column widths ─────────────────────────────────────────────────────────────
  columns.forEach((col, i) => { ws.getColumn(i + 1).width = col.width; });

  // ── Data rows 4–53 (50 rows) ─────────────────────────────────────────────────
  const colIdx = {};
  columns.forEach((col, i) => { colIdx[col.key] = i + 1; });

  const colLetter = n => ws.getColumn(n).letter;

  for (let r = 4; r <= 53; r++) {
    const row = ws.getRow(r);

    // Static defaults
    row.getCell(colIdx['Carton_Type']).value   = 'BDCM1';
    row.getCell(colIdx['Pack_Type']).value      = 'Flat';
    row.getCell(colIdx['Collection_Type']).value= 'Delivery';
    row.getCell(colIdx['Hazardous']).value      = 'N/A';
    row.getCell(colIdx['Var_Unit']).value       = 0;
    row.getCell(colIdx['Var_Pct']).value        = 0;

    // VLOOKUP formulas for carton dimensions from hidden sheet
    const ctRef = `${colLetter(colIdx['Carton_Type'])}${r}`;
    row.getCell(colIdx['Carton_Length_cm']).value = { formula: `IFERROR(VLOOKUP(${ctRef},CARTON_LOOKUP!$A:$E,3,0),"")` };
    row.getCell(colIdx['Carton_Width_cm']).value  = { formula: `IFERROR(VLOOKUP(${ctRef},CARTON_LOOKUP!$A:$E,4,0),"")` };
    row.getCell(colIdx['Carton_Height_cm']).value = { formula: `IFERROR(VLOOKUP(${ctRef},CARTON_LOOKUP!$A:$E,5,0),"")` };
    row.getCell(colIdx['Carton_Weight_KG']).value = { formula: `IFERROR(VLOOKUP(${ctRef},CARTON_LOOKUP!$A:$E,2,0),"")` };

    // Auto-calc formulas
    const noC  = `${colLetter(colIdx['No_of_Cartons'])}${r}`;
    const unitW= `${colLetter(colIdx['Unit_Weight_KG'])}${r}`;
    const bkq  = `${colLetter(colIdx['Booking_Qty'])}${r}`;
    const cL   = `${colLetter(colIdx['Carton_Length_cm'])}${r}`;
    const cW   = `${colLetter(colIdx['Carton_Width_cm'])}${r}`;
    const cH   = `${colLetter(colIdx['Carton_Height_cm'])}${r}`;
    const cWt  = `${colLetter(colIdx['Carton_Weight_KG'])}${r}`;

    row.getCell(colIdx['Gross_Weight_KG']).value = { formula: `IFERROR(${cWt}*${noC},0)` };
    row.getCell(colIdx['Net_Weight_KG']).value   = { formula: `IFERROR(${unitW}*${bkq},0)` };
    row.getCell(colIdx['Volume_M3']).value        = { formula: `IFERROR((${cL}*${cW}*${cH}/1000000)*${noC},0)` };

    // Alternating row shading
    const rowFill = r % 2 === 0
      ? { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9F9F9' } }
      : { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

    // Style auto cells blue-tint, locked
    [
      'Carton_Length_cm','Carton_Width_cm','Carton_Height_cm','Carton_Weight_KG',
      'Gross_Weight_KG','Net_Weight_KG','Volume_M3'
    ].forEach(k => {
      const cell = row.getCell(colIdx[k]);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6E8F8' } };
      cell.font = { color: { argb: 'FF1A5276' }, size: 10 };
      cell.protection = { locked: true };
    });

    // Number formats
    ['Carton_Length_cm','Carton_Width_cm','Carton_Height_cm'].forEach(k =>
      (row.getCell(colIdx[k]).numFmt = '0.00')
    );
    ['Carton_Weight_KG','Gross_Weight_KG','Net_Weight_KG','Unit_Weight_KG'].forEach(k =>
      (row.getCell(colIdx[k]).numFmt = '0.0000')
    );
    row.getCell(colIdx['Volume_M3']).numFmt = '0.0000';

    // Date format
    ['Cargo_Ready_Planned_Collection_Date','Carrier_Booking_Request_Date',
     'Expected_Delivery_Date','ASN_Delivery_Date'].forEach(k =>
      (row.getCell(colIdx[k]).numFmt = 'DD/MM/YYYY')
    );

    row.commit();
  }

  // ── Data Validations (rows 4–53) ─────────────────────────────────────────────
  for (let r = 4; r <= 53; r++) {
    // Traffic_Mode
    ws.getCell(r, colIdx['Traffic_Mode']).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['"CFS,CY"'],
      showErrorMessage: true, errorTitle: 'Invalid value',
      error: 'Please select CFS or CY'
    };
    // Carton_Type — from hidden lookup sheet
    ws.getCell(r, colIdx['Carton_Type']).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['CARTON_LOOKUP!$A$2:$A$20'],
      showErrorMessage: true
    };
    // Pack_Type
    ws.getCell(r, colIdx['Pack_Type']).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['"Flat,Bulk Flat,Hanging"']
    };
    // Collection_Type
    ws.getCell(r, colIdx['Collection_Type']).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['"Collection,Delivery"']
    };
    // Hazardous
    ws.getCell(r, colIdx['Hazardous']).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['"Flammable,Glass - Hazardous,Hazardous,N/A"']
    };
    // No_of_Cartons — whole number > 0
    ws.getCell(r, colIdx['No_of_Cartons']).dataValidation = {
      type: 'whole', operator: 'greaterThan',
      formulae: [0],
      showErrorMessage: true, errorTitle: 'Invalid', error: 'Enter a whole number greater than 0'
    };
    // Unit_Weight_KG — decimal > 0
    ws.getCell(r, colIdx['Unit_Weight_KG']).dataValidation = {
      type: 'decimal', operator: 'greaterThan',
      formulae: [0],
      showErrorMessage: true, errorTitle: 'Invalid', error: 'Enter a positive weight in KG'
    };
    // Date fields
    ['Cargo_Ready_Planned_Collection_Date','Carrier_Booking_Request_Date',
     'Expected_Delivery_Date','ASN_Delivery_Date'].forEach(k => {
      ws.getCell(r, colIdx[k]).dataValidation = {
        type: 'date', operator: 'greaterThan',
        formulae: [new Date(2020, 0, 1)],
        showErrorMessage: true, errorTitle: 'Invalid date',
        error: 'Enter a valid date (DD/MM/YYYY format)'
      };
    });
  }

  // ── Freeze rows 1-3 (banner + header + legend) ────────────────────────────────
  ws.views = [{ state: 'frozen', ySplit: 3, xSplit: 0, showGridLines: true }];

  // ── Sheet tab colour ──────────────────────────────────────────────────────────
  ws.properties.tabColor = { argb: 'FF1F4E79' };

  await wb.xlsx.writeFile(OUT_FILE);
  console.log(`✅  SupplierInput_template.xlsx created at:\n    ${OUT_FILE}`);
  console.log('\nColumn summary:');
  console.log('  🔴 Mandatory (7) :', columns.filter(c => c.type === 'mandatory').map(c => c.label).join(', '));
  console.log('  🟢 Defaulted (6) :', columns.filter(c => c.type === 'default').map(c => c.label).join(', '));
  console.log('  🔵 Auto-calc (7) :', columns.filter(c => c.type === 'auto').map(c => c.label).join(', '));
  console.log('  ⚫ Optional (7)  :', columns.filter(c => c.type === 'optional').map(c => c.label).join(', '));
}

build().catch(err => {
  console.error('❌ Failed:', err.message);
  process.exit(1);
});
