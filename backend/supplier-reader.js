'use strict';

const ExcelJS = require('exceljs');

const REQUIRED_COLS = [
  'PO_Number', 'ASN_Ref', 'No_of_Cartons', 'Unit_Weight_KG',
  'Cargo_Ready_Planned_Collection_Date', 'Carrier_Booking_Request_Date', 'Traffic_Mode'
];

/**
 * Parse supplier-provided Excel file.
 * Reads the first sheet and extracts all rows as objects keyed by header row.
 */
async function parse(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Supplier Excel has no worksheets');

  const headers = [];
  sheet.getRow(1).eachCell((cell, colNum) => {
    headers[colNum] = String(cell.value || '').trim();
  });

  const rows = [];
  const validationErrors = [];

  sheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return; // skip header
    const obj = {};
    row.eachCell((cell, colNum) => {
      const key = headers[colNum];
      if (key) obj[key] = cell.value !== null && cell.value !== undefined ? cell.value : '';
    });
    // Skip entirely empty rows
    if (!Object.values(obj).some(v => v !== '')) return;

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

  return { rows, validationErrors, sheetName: sheet.name };
}

module.exports = { parse };
