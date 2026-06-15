'use strict';

const ExcelJS = require('exceljs');

const REQUIRED_COLS = [
  'PO_Number', 'SKU', 'No_of_Cartons', 'Unit_Weight_KG', 'Booking_Qty',
  'Cargo_Ready_Planned_Collection_Date', 'Carrier_Booking_Request_Date',
  'Traffic_Mode', 'Mode_Of_Transport',
  'Factory_Name', 'Factory_ID', 'Factory_Street1', 'Factory_City', 'Factory_PostalCd', 'Factory_CountryCd',
  'Booking_Group'
];

/**
 * Parse supplier-provided Excel file.
 * Reads the first sheet and extracts all rows as objects keyed by header row.
 */
async function parse(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  // Find SUPPLIER_INPUT sheet by name; fall back to first visible sheet
  let sheet = workbook.getWorksheet('SUPPLIER_INPUT');
  if (!sheet) {
    sheet = workbook.worksheets.find(ws => ws.state !== 'veryHidden' && ws.state !== 'hidden');
  }
  if (!sheet) sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Supplier Excel has no worksheets');

  // Auto-detect the header row — find the first row that contains 'PO_Number'
  let headerRowNum = 1;
  sheet.eachRow((row, rowNum) => {
    if (headerRowNum !== 1) return; // already found
    row.eachCell(cell => {
      if (String(cell.value || '').trim() === 'PO_Number') headerRowNum = rowNum;
    });
  });

  const headers = [];
  sheet.getRow(headerRowNum).eachCell((cell, colNum) => {
    // Strip the "(HH:MM)" annotation added to Collection_Time label
    headers[colNum] = String(cell.value || '').replace(/\s*\(.*?\)/, '').trim();
  });

  const rows = [];
  const validationErrors = [];

  sheet.eachRow((row, rowNum) => {
    if (rowNum <= headerRowNum) return; // skip banner/legend/header rows
    const obj = {};
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const key = headers[colNum];
      if (!key) return;
      // Formula cells → use result; date objects → keep as-is; else string
      let val = cell.value;
      if (val !== null && val !== undefined && typeof val === 'object' && 'result' in val) {
        val = val.result; // unwrap formula result
      }
      obj[key] = (val !== null && val !== undefined) ? val : '';
    });
    // Skip entirely empty rows
    if (!Object.values(obj).some(v => v !== '')) return;
    // Skip pre-filled template rows that have defaults but no booking identity
    if ((!obj.PO_Number || String(obj.PO_Number).trim() === '') &&
        (!obj.SKU       || String(obj.SKU).trim()        === '')) return;

    // Validate required fields
    const missing = REQUIRED_COLS.filter(c => !obj[c] || String(obj[c]).trim() === '');
    if (missing.length > 0) {
      validationErrors.push(`Row ${rowNum}: missing required fields: ${missing.join(', ')}`);
    }
    // Collection_Time mandatory when Collection_Type = 'Collection'
    if (String(obj.Collection_Type || '').trim() === 'Collection' &&
        (!obj.Collection_Time || String(obj.Collection_Time).trim() === '')) {
      validationErrors.push(`Row ${rowNum}: Collection_Time is required when Collection_Type is "Collection"`);
    }

    rows.push(obj);
  });

  return { rows, validationErrors, sheetName: sheet.name, headerRowNum };
}

module.exports = { parse };
