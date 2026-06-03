'use strict';

/**
 * Generates the SupplierBible_template.xlsx with all 6 sheets,
 * dropdowns, VLOOKUP formulas, and pre-set defaults.
 * Run: node backend/build-bible-template.js
 */

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const OUT_DIR  = path.join(__dirname, '..', 'bible');
const OUT_FILE = path.join(OUT_DIR, 'SupplierBible_template.xlsx');

const CARTON_TYPES = [
  { name: 'BDCM1',               weight: 1.40, L: 60.00, W: 30.00, H: 40.00 },
  { name: 'BDCM3',               weight: 1.00, L: 45.00, W: 29.50, H: 18.80 },
  { name: 'C5',                  weight: 1.00, L: 60.00, W: 30.00, H: 20.00 },
  { name: 'Cartons',             weight: 1.00, L: 45.00, W: 60.00, H: 40.00 },
  { name: 'A1',                  weight: 1.00, L: 59.50, W: 28.50, H: 37.50 },
  { name: 'A2',                  weight: 1.00, L: 59.50, W: 28.50, H: 32.50 },
  { name: 'A3',                  weight: 1.00, L: 59.50, W: 28.50, H: 26.00 },
  { name: 'A4',                  weight: 1.00, L: 59.50, W: 28.50, H: 19.00 },
  { name: 'B1',                  weight: 1.00, L: 52.00, W: 25.50, H: 37.50 },
  { name: 'B2',                  weight: 1.00, L: 52.00, W: 25.50, H: 32.50 },
  { name: 'B3',                  weight: 1.00, L: 52.00, W: 25.50, H: 26.00 },
  { name: 'B4',                  weight: 1.00, L: 52.00, W: 25.50, H: 19.00 },
  { name: 'C1',                  weight: 1.00, L: 45.00, W: 28.50, H: 37.50 },
  { name: 'C2',                  weight: 1.00, L: 45.00, W: 28.50, H: 32.50 },
  { name: 'C3',                  weight: 1.00, L: 45.00, W: 28.50, H: 26.00 },
  { name: 'C4',                  weight: 1.00, L: 45.00, W: 28.50, H: 19.00 },
  { name: 'Hanging',             weight: 0.70, L: 213.00, W: 94.00, H: 60.00 },
  { name: 'Hanging Non-Standard',weight: null, L: null,  W: null,  H: null  },
  { name: 'Non-Standard',        weight: null, L: null,  W: null,  H: null  }
];

const HEADER_FILL  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
const HEADER_FONT  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
const AUTO_FILL    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
const REQUIRED_FILL= { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
const DEFAULT_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F8E8' } };

function styleHeader(ws) {
  ws.getRow(1).height = 22;
  ws.getRow(1).eachCell(cell => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFAED6F1' } } };
  });
}

async function buildTemplate() {
  if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

  const wb = new ExcelJS.Workbook();
  wb.creator = 'CarrierBookingStub';
  wb.created = new Date();

  // ── Sheet 5: CARTON_TYPES (build first so VLOOKUP range is ready) ────────────
  const wsCarton = wb.addWorksheet('CARTON_TYPES');
  wsCarton.columns = [
    { header: 'Carton_Type', key: 'name',   width: 22 },
    { header: 'Weight_KG',   key: 'weight', width: 12 },
    { header: 'Length_cm',   key: 'L',      width: 12 },
    { header: 'Width_cm',    key: 'W',      width: 12 },
    { header: 'Height_cm',   key: 'H',      width: 12 }
  ];
  styleHeader(wsCarton);
  CARTON_TYPES.forEach(ct => wsCarton.addRow([ct.name, ct.weight, ct.L, ct.W, ct.H]));
  wsCarton.getColumn(1).font = { bold: true };

  // ── Sheet 1: SUPPLIER_INPUT ───────────────────────────────────────────────────
  const ws1 = wb.addWorksheet('SUPPLIER_INPUT');

  const cols = [
    // Required — supplier must fill
    { header: 'PO_Number',                          key: 'PO_Number',                          width: 20, type: 'required' },
    { header: 'ASN_Ref',                            key: 'ASN_Ref',                            width: 20, type: 'required' },
    { header: 'No_of_Cartons',                      key: 'No_of_Cartons',                      width: 16, type: 'required' },
    { header: 'Unit_Weight_KG',                     key: 'Unit_Weight_KG',                     width: 16, type: 'required' },
    { header: 'Cargo_Ready_Planned_Collection_Date',key: 'Cargo_Ready_Planned_Collection_Date',width: 34, type: 'required' },
    { header: 'Carrier_Booking_Request_Date',       key: 'Carrier_Booking_Request_Date',       width: 28, type: 'required' },
    { header: 'Traffic_Mode',                       key: 'Traffic_Mode',                       width: 16, type: 'required' },
    // Supplier fills
    { header: 'EAN_Barcode',   key: 'EAN_Barcode',  width: 18, type: 'supplier' },
    { header: 'Colour_Code',   key: 'Colour_Code',  width: 14, type: 'supplier' },
    { header: 'Size_Code',     key: 'Size_Code',    width: 12, type: 'supplier' },
    // Carton type (default BDCM1, dropdown)
    { header: 'Carton_Type',   key: 'Carton_Type',  width: 20, type: 'dropdown' },
    // Auto from carton type VLOOKUP
    { header: 'Carton_Length_cm', key: 'Carton_Length_cm', width: 18, type: 'auto' },
    { header: 'Carton_Width_cm',  key: 'Carton_Width_cm',  width: 17, type: 'auto' },
    { header: 'Carton_Height_cm', key: 'Carton_Height_cm', width: 17, type: 'auto' },
    { header: 'Carton_Weight_KG', key: 'Carton_Weight_KG', width: 18, type: 'auto' },
    // Auto-calc
    { header: 'Gross_Weight_KG', key: 'Gross_Weight_KG', width: 17, type: 'auto' },
    { header: 'Net_Weight_KG',   key: 'Net_Weight_KG',   width: 15, type: 'auto' },
    { header: 'Volume_M3',       key: 'Volume_M3',        width: 13, type: 'auto' },
    // Booking qty (default from ASN, overridable)
    { header: 'Booking_Qty',     key: 'Booking_Qty',      width: 14, type: 'supplier' },
    // Dropdowns with defaults
    { header: 'Pack_Type',       key: 'Pack_Type',        width: 14, type: 'dropdown', default: 'Flat' },
    { header: 'Collection_Type', key: 'Collection_Type',  width: 18, type: 'dropdown', default: 'Delivery' },
    { header: 'Hazardous',       key: 'Hazardous',        width: 20, type: 'dropdown', default: 'N/A' },
    // Optional dates
    { header: 'Expected_Delivery_Date', key: 'Expected_Delivery_Date', width: 24, type: 'supplier' },
    { header: 'ASN_Delivery_Date',      key: 'ASN_Delivery_Date',      width: 20, type: 'supplier' },
    // SCC-only fields
    { header: 'Var_Unit', key: 'Var_Unit', width: 12, type: 'default', default: 0 },
    { header: 'Var_Pct',  key: 'Var_Pct',  width: 12, type: 'default', default: 0 },
    { header: 'Remarks',  key: 'Remarks',  width: 30, type: 'supplier' }
  ];

  ws1.columns = cols.map(c => ({ header: c.header, key: c.key, width: c.width }));
  styleHeader(ws1);

  // Style header cells by column type
  cols.forEach((col, idx) => {
    const cell = ws1.getRow(1).getCell(idx + 1);
    if (col.type === 'required') {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    } else if (col.type === 'auto') {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    }
  });

  // Add 50 pre-formatted data rows
  const cartonTypeColIdx = cols.findIndex(c => c.key === 'Carton_Type') + 1; // 1-based
  const noCartonsColIdx  = cols.findIndex(c => c.key === 'No_of_Cartons') + 1;
  const unitWtColIdx     = cols.findIndex(c => c.key === 'Unit_Weight_KG') + 1;
  const bkqColIdx        = cols.findIndex(c => c.key === 'Booking_Qty') + 1;
  const ctLIdx = cols.findIndex(c => c.key === 'Carton_Length_cm') + 1;
  const ctWIdx = cols.findIndex(c => c.key === 'Carton_Width_cm') + 1;
  const ctHIdx = cols.findIndex(c => c.key === 'Carton_Height_cm') + 1;
  const ctWtIdx= cols.findIndex(c => c.key === 'Carton_Weight_KG') + 1;
  const grossIdx= cols.findIndex(c => c.key === 'Gross_Weight_KG') + 1;
  const netIdx  = cols.findIndex(c => c.key === 'Net_Weight_KG') + 1;
  const volIdx  = cols.findIndex(c => c.key === 'Volume_M3') + 1;

  const colLetter = n => ws1.getColumn(n).letter;

  for (let r = 2; r <= 51; r++) {
    const row = ws1.getRow(r);

    // Defaults
    const ctCell = row.getCell(cartonTypeColIdx);
    ctCell.value = 'BDCM1';

    row.getCell(cols.findIndex(c => c.key === 'Pack_Type')       + 1).value = 'Flat';
    row.getCell(cols.findIndex(c => c.key === 'Collection_Type') + 1).value = 'Delivery';
    row.getCell(cols.findIndex(c => c.key === 'Hazardous')       + 1).value = 'N/A';
    row.getCell(cols.findIndex(c => c.key === 'Var_Unit')        + 1).value = 0;
    row.getCell(cols.findIndex(c => c.key === 'Var_Pct')         + 1).value = 0;

    // VLOOKUP formulas for carton dims from CARTON_TYPES sheet
    const ctRef = `${colLetter(cartonTypeColIdx)}${r}`;
    row.getCell(ctLIdx).value  = { formula: `IFERROR(VLOOKUP(${ctRef},CARTON_TYPES!$A:$E,3,0),"")` };
    row.getCell(ctWIdx).value  = { formula: `IFERROR(VLOOKUP(${ctRef},CARTON_TYPES!$A:$E,4,0),"")` };
    row.getCell(ctHIdx).value  = { formula: `IFERROR(VLOOKUP(${ctRef},CARTON_TYPES!$A:$E,5,0),"")` };
    row.getCell(ctWtIdx).value = { formula: `IFERROR(VLOOKUP(${ctRef},CARTON_TYPES!$A:$E,2,0),"")` };

    // Auto-calc formulas
    const noC  = `${colLetter(noCartonsColIdx)}${r}`;
    const unitW= `${colLetter(unitWtColIdx)}${r}`;
    const bkq  = `${colLetter(bkqColIdx)}${r}`;
    const cL   = `${colLetter(ctLIdx)}${r}`;
    const cW   = `${colLetter(ctWIdx)}${r}`;
    const cH   = `${colLetter(ctHIdx)}${r}`;
    const cWt  = `${colLetter(ctWtIdx)}${r}`;

    row.getCell(grossIdx).value = { formula: `IFERROR(${cWt}*${noC},0)` };
    row.getCell(netIdx).value   = { formula: `IFERROR(${unitW}*${bkq},0)` };
    row.getCell(volIdx).value   = { formula: `IFERROR((${cL}*${cW}*${cH}/1000000)*${noC},0)` };

    // Style auto cells
    [ctLIdx, ctWIdx, ctHIdx, ctWtIdx, grossIdx, netIdx, volIdx].forEach(ci => {
      row.getCell(ci).fill = AUTO_FILL;
      row.getCell(ci).protection = { locked: true };
    });

    row.commit();
  }

  // Data validation — dropdowns
  const cartonNames = CARTON_TYPES.map(c => c.name).join(',');

  for (let r = 2; r <= 51; r++) {
    // Traffic_Mode
    ws1.getCell(r, cols.findIndex(c => c.key === 'Traffic_Mode') + 1).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['"CFS,CY"'],
      showErrorMessage: true, errorTitle: 'Invalid', error: 'Please select CFS or CY'
    };
    // Carton_Type
    ws1.getCell(r, cartonTypeColIdx).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['CARTON_TYPES!$A$2:$A$20'],
      showErrorMessage: true
    };
    // Pack_Type
    ws1.getCell(r, cols.findIndex(c => c.key === 'Pack_Type') + 1).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['"Flat,Bulk Flat,Hanging"']
    };
    // Collection_Type
    ws1.getCell(r, cols.findIndex(c => c.key === 'Collection_Type') + 1).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['"Collection,Delivery"']
    };
    // Hazardous
    ws1.getCell(r, cols.findIndex(c => c.key === 'Hazardous') + 1).dataValidation = {
      type: 'list', allowBlank: false,
      formulae: ['"Flammable,Glass - Hazardous,Hazardous,N/A"']
    };
    // Dates
    [
      'Cargo_Ready_Planned_Collection_Date',
      'Carrier_Booking_Request_Date',
      'Expected_Delivery_Date',
      'ASN_Delivery_Date'
    ].forEach(k => {
      const ci = cols.findIndex(c => c.key === k) + 1;
      ws1.getCell(r, ci).dataValidation = {
        type: 'date', operator: 'greaterThan',
        formulae: [new Date(2020, 0, 1)],
        showErrorMessage: true, errorTitle: 'Invalid date',
        error: 'Please enter a valid date (DD/MM/YYYY)'
      };
      ws1.getCell(r, ci).numFmt = 'DD/MM/YYYY';
    });
  }

  // Freeze header row
  ws1.views = [{ state: 'frozen', ySplit: 1 }];

  // Legend row at top (row 1 note)
  ws1.getRow(1).getCell(1).note = {
    texts: [
      { font: { bold: true }, text: 'RED headers = Mandatory supplier input\n' },
      { font: { color: { argb: 'FF2E75B6' } }, text: 'BLUE headers = Auto-calculated\n' },
      { text: 'WHITE headers = Optional/defaulted' }
    ]
  };

  // ── Sheet 2: PO_FEED_EXTRACT ────────────────────────────────────────────────
  const ws2 = wb.addWorksheet('PO_FEED_EXTRACT');
  ws2.columns = [
    { header: 'orderId',       width: 20 }, { header: 'supplierName',  width: 30 },
    { header: 'supplierId',    width: 16 }, { header: 'factoryName',   width: 30 },
    { header: 'factoryId',     width: 14 }, { header: 'factoryCity',   width: 16 },
    { header: 'factoryCountry',width: 16 }, { header: 'fcName',        width: 20 },
    { header: 'fcId',          width: 10 }, { header: 'carrierId',     width: 12 },
    { header: 'loadingPortId', width: 16 }, { header: 'incoterms',     width: 12 }
  ];
  styleHeader(ws2);
  ws2.getRow(2).getCell(1).value = '(Auto-populated by backend when feeds are fetched from Azure Blob)';
  ws2.views = [{ state: 'frozen', ySplit: 1 }];

  // ── Sheet 3: ASN_FEED_EXTRACT ───────────────────────────────────────────────
  const ws3 = wb.addWorksheet('ASN_FEED_EXTRACT');
  ws3.columns = [
    { header: 'documentId',   width: 24 }, { header: 'fcId',         width: 10 },
    { header: 'receivedDate', width: 20 }, { header: 'orderId',      width: 18 },
    { header: 'sku',          width: 16 }, { header: 'receivedQty',  width: 14 },
    { header: 'shipmentRef',  width: 24 }
  ];
  styleHeader(ws3);
  ws3.getRow(2).getCell(1).value = '(Auto-populated by backend when feeds are fetched from Azure Blob)';
  ws3.views = [{ state: 'frozen', ySplit: 1 }];

  // ── Sheet 4: MASTER ─────────────────────────────────────────────────────────
  const ws4 = wb.addWorksheet('MASTER');
  const masterHeaders = [
    'Booking_Ref','PO_Number','ASN_Ref','Supplier_Name','Supplier_ID',
    'Factory_Name','Factory_ID','Factory_Street1','Factory_Street2','Factory_Street3',
    'Factory_City','Factory_PostalCd','Factory_CountryCd',
    'FC_Name','FC_ID','FC_Street1','FC_Street2','FC_Street3',
    'FC_City','FC_StateProvinceCd','FC_PostalCd','FC_CountryCd',
    'Carrier_ID','Carrier_Name','Loading_Port_LOCODE',
    'SKU','EAN_Barcode','Colour_Code','Size_Code','Product_Style','Description',
    'Carton_Type','Carton_Length_cm','Carton_Width_cm','Carton_Height_cm','Carton_Weight_KG',
    'No_of_Cartons','Unit_Weight_KG','Booking_Qty',
    'Gross_Weight_KG','Net_Weight_KG','Volume_M3',
    'Pack_Type','Collection_Type','Hazardous','Traffic_Mode',
    'Cargo_Ready_Planned_Collection_Date','Carrier_Booking_Request_Date',
    'Expected_Delivery_Date','ASN_Delivery_Date',
    'Var_Unit','Var_Pct','Remarks'
  ];
  ws4.addRow(masterHeaders);
  styleHeader(ws4);
  ws4.columns = masterHeaders.map(h => ({ header: h, width: 22 }));
  ws4.getRow(2).getCell(1).value = '(Auto-populated by backend — Build Bible step merges all sources here)';
  ws4.views = [{ state: 'frozen', ySplit: 1 }];

  // ── Sheet 6: GENERATION_LOG ─────────────────────────────────────────────────
  const ws6 = wb.addWorksheet('GENERATION_LOG');
  ws6.columns = [
    { header: 'Timestamp',   width: 22 }, { header: 'Booking_Ref', width: 18 },
    { header: 'PO_Numbers',  width: 30 }, { header: 'Filename',    width: 50 },
    { header: 'CtrlNumber',  width: 14 }, { header: 'SFTP_Status', width: 14 }
  ];
  styleHeader(ws6);
  ws6.views = [{ state: 'frozen', ySplit: 1 }];

  await wb.xlsx.writeFile(OUT_FILE);
  console.log(`✅ SupplierBible_template.xlsx created at: ${OUT_FILE}`);
}

buildTemplate().catch(err => {
  console.error('❌ Failed to build Bible template:', err.message);
  process.exit(1);
});
