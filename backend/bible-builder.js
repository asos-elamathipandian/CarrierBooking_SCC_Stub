'use strict';

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const BIBLE_DIR = path.join(__dirname, '..', 'bible');
const BIBLE_FILE = path.join(BIBLE_DIR, 'SupplierBible_working.xlsx');
const LOG_FILE   = path.join(BIBLE_DIR, 'generation-log.json');

const CARTON_TYPES = {
  'BDCM1':   { weight: 1.40, L: 60.00, W: 30.00, H: 40.00 },
  'BDCM3':   { weight: 1.00, L: 45.00, W: 29.50, H: 18.80 },
  'C5':      { weight: 1.00, L: 60.00, W: 30.00, H: 20.00 },
  'Cartons': { weight: 1.00, L: 45.00, W: 60.00, H: 40.00 },
  'A1':      { weight: 1.00, L: 59.50, W: 28.50, H: 37.50 },
  'A2':      { weight: 1.00, L: 59.50, W: 28.50, H: 32.50 },
  'A3':      { weight: 1.00, L: 59.50, W: 28.50, H: 26.00 },
  'A4':      { weight: 1.00, L: 59.50, W: 28.50, H: 19.00 },
  'B1':      { weight: 1.00, L: 52.00, W: 25.50, H: 37.50 },
  'B2':      { weight: 1.00, L: 52.00, W: 25.50, H: 32.50 },
  'B3':      { weight: 1.00, L: 52.00, W: 25.50, H: 26.00 },
  'B4':      { weight: 1.00, L: 52.00, W: 25.50, H: 19.00 },
  'C1':      { weight: 1.00, L: 45.00, W: 28.50, H: 37.50 },
  'C2':      { weight: 1.00, L: 45.00, W: 28.50, H: 32.50 },
  'C3':      { weight: 1.00, L: 45.00, W: 28.50, H: 26.00 },
  'C4':      { weight: 1.00, L: 45.00, W: 28.50, H: 19.00 },
  'Hanging': { weight: 0.70, L: 213.00, W: 94.00, H: 60.00 },
  'Hanging Non-Standard': { weight: null, L: null, W: null, H: null },
  'Non-Standard':         { weight: null, L: null, W: null, H: null }
};

/**
 * Merge supplier data + PO feeds + ASN feeds into MASTER rows.
 * Write to working Excel with all 6 sheets.
 */
async function build(supplierData, feedData) {
  if (!fs.existsSync(BIBLE_DIR)) fs.mkdirSync(BIBLE_DIR, { recursive: true });

  const { rows: supplierRows } = supplierData;
  const { poFeeds, asnFeeds } = feedData;

  // Index feeds by their key fields
  const poByOrderId = {};
  const poByLinesku = {}; // key: `${orderId}_${sku}`
  for (const po of poFeeds) {
    poByOrderId[po.orderId] = po;
    for (const li of (po.lineItems || [])) {
      poByLinesku[`${po.orderId}_${li.sku}`] = { po, line: li };
    }
  }

  const asnByDocId = {};
  for (const asn of asnFeeds) {
    asnByDocId[asn.documentId] = asn;
  }

  // Build MASTER rows — one row per (PO, ASN, SKU)
  const masterRows = [];
  for (const sRow of supplierRows) {
    const poNum  = String(sRow.PO_Number || '').trim();
    const asnRef = String(sRow.ASN_Ref   || '').trim();

    const po  = poByOrderId[poNum];
    const asn = asnByDocId[asnRef];

    // Carton type / dimensions from supplier row (or carton master default)
    const cartonType = String(sRow.Carton_Type || 'BDCM1').trim();
    const ct = CARTON_TYPES[cartonType] || CARTON_TYPES['BDCM1'];
    const noCartons  = parseFloat(sRow.No_of_Cartons) || 0;
    const unitWeight = parseFloat(sRow.Unit_Weight_KG) || 0;
    const cL = parseFloat(sRow.Carton_Length_cm) || ct.L || 0;
    const cW = parseFloat(sRow.Carton_Width_cm)  || ct.W || 0;
    const cH = parseFloat(sRow.Carton_Height_cm) || ct.H || 0;
    const cWt= parseFloat(sRow.Carton_Weight_KG) || ct.weight || 0;

    // ── One row per SKU: use SKU from supplier row directly ──────────────────
    const sku = String(sRow.SKU || '').trim();

    // Look up this SKU in the PO and ASN feeds
    const poLine  = poByLinesku[`${poNum}_${sku}`];
    const asnLinesForPO = asn
      ? asn.lines.filter(l => String(l.orderId) === poNum)
      : [];
    const asnLine = asnLinesForPO.find(l => l.sku === sku);

      // Booking qty: mandatory from supplier row
      const skuQty     = parseFloat(sRow.Booking_Qty) || 0;
    const skuCartons = noCartons; // cartons are per-SKU row as filled by supplier

    masterRows.push({
        // Booking identity
        Booking_Ref: sRow.Booking_Ref || '',
        PO_Number:   poNum,
        ASN_Ref:     asnRef,

        // From PO feed
        Supplier_Name:     po?.supplierName   || '',
        Supplier_ID:       po?.supplierId     || '',
        Factory_Name:      po?.factoryName    || '',
        Factory_ID:        po?.factoryId      || '',
        Factory_Street1:   po?.factoryStreet1 || '',
        Factory_Street2:   po?.factoryStreet2 || '',
        Factory_Street3:   po?.factoryStreet3 || '',
        Factory_City:      po?.factoryCity    || '',
        Factory_PostalCd:  po?.factoryPostal  || '',
        Factory_CountryCd: po?.factoryCountry || '',
        FC_Name:           po?.fcName         || '',
        FC_ID:             po?.fcId           || sRow.FC_ID || 'FC01',
        FC_Street1:        po?.fcStreet1      || '',
        FC_Street2:        po?.fcStreet2      || '',
        FC_Street3:        po?.fcStreet3      || '',
        FC_City:           po?.fcCity         || '',
        FC_StateProvinceCd: po?.fcState       || '',
        FC_PostalCd:       po?.fcPostal       || '',
        FC_CountryCd:      po?.fcCountry      || 'GB',
        Carrier_ID:        po?.carrierId      || '',
        Carrier_Name:      po?.carrierName    || '',
        Loading_Port_LOCODE: po?.loadingPortId || '',
        F1_ID:             po?.f1Id           || '',

        // SKU — from supplier row; enrich description/style from PO feed
        SKU:           sku,
        Product_Style: poLine?.line?.productStyle || '',
        Description:   poLine?.line?.description  || '',

        // Supplier-provided per SKU row
        EAN_Barcode:   sRow.EAN_Barcode || '',
        Colour_Code:   sRow.Colour_Code || '',
        Size_Code:     sRow.Size_Code   || '',

        // Carton data — filled per SKU row by supplier
        Carton_Type:     cartonType,
        Carton_Length_cm: cL,
        Carton_Width_cm:  cW,
        Carton_Height_cm: cH,
        Carton_Weight_KG: cWt,
        No_of_Cartons:   skuCartons,
        Unit_Weight_KG:  unitWeight,
        Booking_Qty:     skuQty,

        // Calculated
        Gross_Weight_KG: parseFloat((cWt * skuCartons).toFixed(4)),
        Net_Weight_KG:   parseFloat((unitWeight * skuQty).toFixed(4)),
        Volume_M3:       parseFloat(((cL * cW * cH / 1000000) * skuCartons).toFixed(4)),

        // Other booking fields from supplier row
        Pack_Type:       sRow.Pack_Type        || 'Flat',
        Collection_Type: sRow.Collection_Type  || 'Delivery',
        Hazardous:       sRow.Hazardous        || 'N/A',
        Traffic_Mode:    sRow.Traffic_Mode     || '',
        Cargo_Ready_Planned_Collection_Date: sRow.Cargo_Ready_Planned_Collection_Date || '',
        Carrier_Booking_Request_Date:        sRow.Carrier_Booking_Request_Date        || '',
        Expected_Delivery_Date:              sRow.Expected_Delivery_Date              || '',
        ASN_Delivery_Date:                   sRow.ASN_Delivery_Date                  || '',
        Var_Unit: sRow.Var_Unit ?? 0,
        Var_Pct:  sRow.Var_Pct  ?? 0,
        Remarks:  sRow.Remarks  || '',
        ASOS_Intake_Week:    sRow.ASOS_Intake_Week || '',
        Collection_Time:     sRow.Collection_Time  || '',
        Incoterms:           po?.incoterms          || '',
        Transport_Mode_Code: poLine?.line?.mode || po?.lineItems?.[0]?.mode || '30'
      });
  }

  // Write Excel
  await writeExcel(masterRows, supplierRows, poFeeds, asnFeeds);

  return { masterRows, filePath: BIBLE_FILE };
}

async function writeExcel(masterRows, supplierRows, poFeeds, asnFeeds) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'CarrierBookingStub';
  wb.created = new Date();

  const headerStyle = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } }, alignment: { horizontal: 'center' } };
  const autoFillStyle = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } } };

  function sanitize(v) {
    if (v === null || v === undefined) return '';
    // Unwrap ExcelJS formula/shared-formula objects
    if (typeof v === 'object' && !(v instanceof Date)) {
      if ('result' in v) v = v.result;
      else if ('formula' in v || 'sharedFormula' in v) return '';
      else if ('richText' in v) return v.richText.map(r => r.text || '').join('');
      else return '';
    }
    // Re-check after unwrap (result could itself be an object)
    if (v !== null && typeof v === 'object' && !(v instanceof Date)) return '';
    return v ?? '';
  }

  function addSheet(name, headers, rows) {
    const ws = wb.addWorksheet(name);
    ws.addRow(headers);
    ws.getRow(1).eachCell(cell => Object.assign(cell, headerStyle));
    ws.getRow(1).height = 20;
    for (const row of rows) {
      ws.addRow(headers.map(h => sanitize(row[h])));
    }
    ws.columns.forEach(col => { col.width = 22; });
    return ws;
  }

  // Sheet 1: SUPPLIER_INPUT (from raw supplier rows)
  {
    const hdrs = [
      'PO_Number','ASN_Ref','SKU','No_of_Cartons','Unit_Weight_KG',
      'Cargo_Ready_Planned_Collection_Date','Carrier_Booking_Request_Date',
      'Traffic_Mode','EAN_Barcode','Colour_Code','Size_Code','Carton_Type',
      'Carton_Length_cm','Carton_Width_cm','Carton_Height_cm','Carton_Weight_KG',
      'Gross_Weight_KG','Net_Weight_KG','Volume_M3','Booking_Qty','Pack_Type',
      'Collection_Type','Collection_Time','Hazardous','Expected_Delivery_Date',
      'ASN_Delivery_Date','ASOS_Intake_Week','Var_Unit','Var_Pct','Remarks'
    ];
    addSheet('SUPPLIER_INPUT', hdrs, supplierRows);
  }

  // Sheet 2: PO_FEED_EXTRACT
  {
    const hdrs = ['orderId','supplierName','supplierId','factoryName','factoryId',
      'factoryCity','factoryCountry','fcName','fcId','carrierId','loadingPortId','incoterms'];
    addSheet('PO_FEED_EXTRACT', hdrs, poFeeds);
  }

  // Sheet 3: ASN_FEED_EXTRACT
  {
    const asnFlat = [];
    for (const asn of asnFeeds) {
      for (const li of (asn.lines || [])) {
        asnFlat.push({ documentId: asn.documentId, fcId: asn.fcId, receivedDate: asn.receivedDate, ...li });
      }
    }
    const hdrs = ['documentId','fcId','receivedDate','orderId','sku','receivedQty','shipmentRef'];
    addSheet('ASN_FEED_EXTRACT', hdrs, asnFlat);
  }

  // Sheet 4: MASTER
  {
    const hdrs = Object.keys(masterRows[0] || {});
    const ws = addSheet('MASTER', hdrs, masterRows);
    // Highlight auto-calc columns
    const autoCols = ['Gross_Weight_KG','Net_Weight_KG','Volume_M3',
                      'Carton_Length_cm','Carton_Width_cm','Carton_Height_cm','Carton_Weight_KG'];
    ws.eachRow((row, rn) => {
      if (rn === 1) return;
      row.eachCell((cell, cn) => {
        const hdr = hdrs[cn - 1];
        if (autoCols.includes(hdr)) Object.assign(cell, { style: autoFillStyle });
      });
    });
  }

  // Sheet 5: CARTON_TYPES
  {
    const ws = wb.addWorksheet('CARTON_TYPES');
    const hdrs = ['Carton_Type','Weight_KG','Length_cm','Width_cm','Height_cm'];
    ws.addRow(hdrs);
    ws.getRow(1).eachCell(cell => Object.assign(cell, headerStyle));
    for (const [name, dims] of Object.entries(CARTON_TYPES)) {
      ws.addRow([name, dims.weight, dims.L, dims.W, dims.H]);
    }
    ws.columns.forEach(col => { col.width = 18; });
  }

  // Sheet 6: GENERATION_LOG (from JSON sidecar — no Excel re-read needed)
  {
    const hdrs = ['Timestamp','Booking_Ref','PO_Numbers','Filename','CtrlNumber','SFTP_Status'];
    const existing = readLogFromJson();
    addSheet('GENERATION_LOG', hdrs, existing);
  }

  await wb.xlsx.writeFile(BIBLE_FILE);
}

// ── JSON-based log (avoids ExcelJS re-read / shared-formula crash) ────────────
function readLogFromJson() {
  if (!fs.existsSync(LOG_FILE)) return [];
  try { return JSON.parse(fs.readFileSync(LOG_FILE, 'utf8')); }
  catch (_) { return []; }
}

function writeLogToJson(entries) {
  try { fs.writeFileSync(LOG_FILE, JSON.stringify(entries, null, 2), 'utf8'); }
  catch (err) { console.error('writeLogToJson error:', err.message); }
}

function appendGenerationLog(entry) {
  const entries = readLogFromJson();
  entries.push(entry);
  writeLogToJson(entries);
}

function getGenerationLog() {
  return readLogFromJson();
}

module.exports = { build, appendGenerationLog, getGenerationLog };
