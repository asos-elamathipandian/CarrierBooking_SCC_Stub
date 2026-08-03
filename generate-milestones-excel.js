const ExcelJS = require('exceljs');
const path = require('path');

const milestones = [
  // Common events
  { event: 'Closed',                                  maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'ASN/IWT Consignment Create Date',         maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Shipment Date',                           maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Manifest Date',                           maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'ASN Handover Date',                       maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'ASN Carrier Query',                       maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'ASN Carrier Query Complete',              maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Date Carrier Booking Confirmed',          maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Carrier Booking Request Date',            maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Carrier Booking Date',                    maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Supplier Booking Transmission Date',      maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'FC Booking Requested',                    maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'FC Arrived Date',                         maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'FC Booking Confirmed',                    maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Receipted First Destination',             maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Receipted Final Destination',             maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Arrival at Dest Port Date',               maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'ETA for Schedule Calculation',            maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Expected Handover Date',                  maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Expected Shipment Date',                  maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },
  { event: 'Estimated Delivery Date First',           maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'Yes' },

  // Maersk Air + Ocean but NOT Advanced
  { event: 'Original Supplier Booking Creation Date',      maerskAir: 'Yes', maerskOcean: 'No',  advanced: 'No' },
  { event: 'Supplier Booking Creation Date',               maerskAir: 'Yes', maerskOcean: 'No',  advanced: 'No' },
  { event: 'Original Supplier Booking Transmission Date',  maerskAir: 'Yes', maerskOcean: 'No',  advanced: 'No' },
  { event: 'Cargo Ready Date',                             maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'No' },
  { event: 'Import Customs Cleared',                       maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'No' },
  { event: 'Carrier Collected from Dest Port',             maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'No' },
  { event: 'ASN Receiving Closed',                         maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'No' },

  // Maersk Air only
  { event: 'Invoice Date',                            maerskAir: 'Yes', maerskOcean: 'No',  advanced: 'No' },
  { event: 'PO Expected Delivery Date First',         maerskAir: 'Yes', maerskOcean: 'Yes', advanced: 'No' },

  // Maersk Ocean only
  { event: 'QCC Start Date',                          maerskAir: 'No',  maerskOcean: 'Yes', advanced: 'No' },
  { event: 'QCC Complete Date',                       maerskAir: 'No',  maerskOcean: 'Yes', advanced: 'No' },
  { event: 'Estimated Delivery Date Final',           maerskAir: 'No',  maerskOcean: 'Yes', advanced: 'No' },
  { event: 'Latest Planned Shipment Date',            maerskAir: 'No',  maerskOcean: 'Yes', advanced: 'No' },
  { event: 'Transhipment Arrival Date',               maerskAir: 'No',  maerskOcean: 'Yes', advanced: 'No' },

  // Advanced only
  { event: 'Reprocessing Start Date',                 maerskAir: 'No',  maerskOcean: 'No',  advanced: 'Yes' },
  { event: 'Reprocessing Complete Date',              maerskAir: 'No',  maerskOcean: 'No',  advanced: 'Yes' },
  { event: 'Arrived at Advanced',                     maerskAir: 'No',  maerskOcean: 'No',  advanced: 'Yes' },
  { event: 'Despatched from Advanced',                maerskAir: 'No',  maerskOcean: 'No',  advanced: 'Yes' },
];

async function generateExcel() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Milestone Comparison');

  // Column widths
  sheet.columns = [
    { header: 'Milestone / Event',           key: 'event',       width: 42 },
    { header: 'Maersk Air',                  key: 'maerskAir',   width: 14 },
    { header: 'Maersk Ocean',                key: 'maerskOcean', width: 14 },
    { header: 'Advanced ASN',                key: 'advanced',    width: 14 },
  ];

  // Header styling
  const headerRow = sheet.getRow(1);
  headerRow.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin' }, bottom: { style: 'thin' },
      left: { style: 'thin' }, right: { style: 'thin' }
    };
  });
  headerRow.height = 24;

  // Data rows
  milestones.forEach((m, idx) => {
    const row = sheet.addRow([m.event, m.maerskAir, m.maerskOcean, m.advanced]);
    row.height = 18;

    const rowBg = idx % 2 === 0 ? 'FFEAF0FB' : 'FFFFFFFF';

    row.eachCell((cell, colNumber) => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        left: { style: 'thin', color: { argb: 'FFD0D0D0' } },
        right: { style: 'thin', color: { argb: 'FFD0D0D0' } }
      };
      cell.alignment = { vertical: 'middle', horizontal: colNumber === 1 ? 'left' : 'center' };

      if (colNumber === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
      } else {
        const val = cell.value;
        if (val === 'Yes') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
          cell.font  = { color: { argb: 'FF276221' }, bold: true };
        } else {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
          cell.font  = { color: { argb: 'FF9C0006' }, bold: true };
        }
      }
    });
  });

  // Freeze header
  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  // Auto-filter
  sheet.autoFilter = { from: 'A1', to: 'D1' };

  const outPath = path.join(__dirname, 'output', 'Maersk_vs_Advanced_Milestones.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log(`Excel saved to: ${outPath}`);
}

generateExcel().catch(console.error);
