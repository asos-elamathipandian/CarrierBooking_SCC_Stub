'use strict';

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const BIBLE_DIR = path.join(__dirname, '..', 'bible');
const BIBLE_FILE = path.join(BIBLE_DIR, 'SupplierBible_working.xlsx');
const LOG_FILE   = path.join(BIBLE_DIR, 'generation-log.json');

// ── FC Master — keyed by POFC / FC_ID from carrier feed ───────────────────────
// Add / update entries here when new FCs come online.
const FC_MASTER = {
  'FC01': { name: 'ASOS Barnsley',           street1: 'Langthwaite Grange Industrial Estate', street2: '', city: 'South Kirkby',   postal: 'WF9 3NR',  country: 'GB' },
  'FC02': { name: 'ASOS Lichfield',           street1: 'Swan Island',                          street2: '', city: 'Lichfield',      postal: 'WS13 7DD', country: 'GB' },
  'FC03': { name: 'ASOS Hemel Hempstead',     street1: 'Maylands Avenue',                      street2: '', city: 'Hemel Hempstead',postal: 'HP2 7DE',  country: 'GB' },
  'FC04': { name: 'FC04 Berlin',             street1: 'An der Anhalter Bahn 6',               street2: '', city: 'Berlin',         postal: '',         country: 'DE' },
  'FC05': { name: 'ASOS US Hub',              street1: '5025 Bleigh Avenue',                   street2: '', city: 'Philadelphia',   postal: 'PA 19136', country: 'US' },
};

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
 * Write to working Excel with all sheets.
 */
async function build(supplierData, feedData) {
  if (!fs.existsSync(BIBLE_DIR)) fs.mkdirSync(BIBLE_DIR, { recursive: true });

  const { rows: supplierRows } = supplierData;
  const { poFeeds = [], asnFeeds = [], carrierAsnFiles = [] } = feedData;

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

  // ── Carrier ASN index ──────────────────────────────────────────────────────
  // carrierAsnIndex[poId][sku] = { asnId, qty, ean, description, size, colour, style, packFormat, country }
  // A PO may appear in multiple carrier files (e.g. split shipments) — merge all.
  const carrierAsnIndex = {}; // poId -> sku -> carrier line data
  const carrierPoMeta   = {}; // poId -> { supplier, supplierCode, shippingPoint, shippingTerms, fcId }
  // Sort files oldest-first so later files overwrite earlier ones (latest wins)
  const sortedCarrierFiles = [...carrierAsnFiles].sort((a, b) => {
    const ta = a.lastModified ? new Date(a.lastModified).getTime() : 0;
    const tb = b.lastModified ? new Date(b.lastModified).getTime() : 0;
    return ta - tb;
  });
  for (const file of sortedCarrierFiles) {
    for (const asnGroup of (file.parsed || [])) {
      const poId = asnGroup.poId;
      if (!carrierAsnIndex[poId]) carrierAsnIndex[poId] = {};
      // Always overwrite meta with latest file's data
      carrierPoMeta[poId] = {
        supplier:         asnGroup.supplier         || '',
        supplierCode:     asnGroup.supplierCode     || '',
        shippingPoint:    asnGroup.shippingPoint    || '',
        shippingTerms:    asnGroup.shippingTerms    || '',
        pofc:             asnGroup.pofc             || '',
        finalDestination: asnGroup.finalDestination || '',
        mode:             asnGroup.mode             || '',
        carrier:          asnGroup.carrier          || '',
        factoryCode:      asnGroup.factoryCode       || '',
        factoryName:      asnGroup.factoryName       || '',
        factoryID:        asnGroup.factoryID         || '',
        factoryStreet1:   asnGroup.factoryStreet1    || '',
        factoryCity:      asnGroup.factoryCity       || '',
        factoryPostal:    asnGroup.factoryPostal     || '',
        factoryCountry:   asnGroup.factoryCountry    || '',
        supplierStreet1:  asnGroup.supplierStreet1   || '',
        supplierCity:     asnGroup.supplierCity      || '',
        supplierPostal:   asnGroup.supplierPostal    || '',
        supplierCountry:  asnGroup.supplierCountry   || ''
      };
      for (const line of asnGroup.lines) {
        // Files sorted oldest→newest so every write overwrites with a later file's data
        carrierAsnIndex[poId][line.sku] = {
          asnId:                asnGroup.asnId,
          pofc:                 asnGroup.pofc,
          finalDestination:     asnGroup.finalDestination,
          shipDate:             asnGroup.shipDate,
          ean:                  line.ean,
          description:          line.description,
          size:                 line.size,
          colour:               line.colour,
          style:                line.style,
          packFormat:           line.packFormat,
          country:              line.country,
          quantity:             line.quantity,
          expectedDeliveryDate: line.expectedDeliveryDate || ''
        };
      }
    }
  }
  const hasCarrierData = Object.keys(carrierAsnIndex).length > 0;

  // Build MASTER rows — one row per (PO, ASN, SKU)
  // When carrier ASN data is available, only include SKUs present in the carrier feed.
  // SKUs in supplier template but NOT in carrier feed are excluded and surfaced as warnings.
  const masterRows = [];
  const skuWarnings = []; // { poNum, sku } excluded because not on carrier feed

  for (const sRow of supplierRows) {
    const poNum  = String(sRow.PO_Number || '').trim();
    const sku    = String(sRow.SKU       || '').trim();

    // If carrier data exists for this PO, skip SKUs not in the carrier feed
    if (hasCarrierData && carrierAsnIndex[poNum]) {
      if (!carrierAsnIndex[poNum][sku]) {
        skuWarnings.push(`PO ${poNum} — SKU ${sku}: found in supplier template but NOT in carrier feed — excluded from VBKREQ`);
        continue;
      }
    }

    const carrierLine = carrierAsnIndex[poNum]?.[sku];
    // ASN ref: use carrier ASNID if available, otherwise fall back to supplier row
    const asnRef = carrierLine?.asnId || String(sRow.ASN_Ref || '').trim();

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

    // Look up this SKU in the PO and ASN feeds
    const poLine  = poByLinesku[`${poNum}_${sku}`];
    const asnLinesForPO = asn
      ? asn.lines.filter(l => String(l.orderId) === poNum)
      : [];
    const asnLine = asnLinesForPO.find(l => l.sku === sku);

    // Booking qty: use supplier row qty — carrier feed is outbound to carrier (planned),
    // not a shipment confirmation, so its Quantity is not used for capping.
    const skuQty     = parseFloat(sRow.Booking_Qty) || 0;
    const skuCartons = noCartons;

    masterRows.push({
        // Booking identity
        Booking_Ref:   sRow.Booking_Ref   || '',
        Booking_Group: sRow.Booking_Group || '',
        PO_Number:     poNum,
        ASN_Ref:       asnRef,  // carrier ASNID when available

        // Supplier — carrier feed takes priority; fall back to supplier template columns
        // When using Databricks source, supplier is the ID code (name not available in ADE table)
        Supplier_Name:      carrierPoMeta[poNum]?.supplier     || sRow.Supplier_Name || carrierPoMeta[poNum]?.supplierCode || sRow.Supplier_ID || '',
        Supplier_ID:        carrierPoMeta[poNum]?.supplierCode || sRow.Supplier_ID   || '',
        Supplier_Street1:   carrierPoMeta[poNum]?.supplierStreet1 || '',
        Supplier_City:      carrierPoMeta[poNum]?.supplierCity    || '',
        Supplier_PostalCd:  carrierPoMeta[poNum]?.supplierPostal  || '',
        Supplier_CountryCd: carrierPoMeta[poNum]?.supplierCountry || '',

        // Factory — Databricks primary, supplier template fallback
        Factory_ID:        carrierPoMeta[poNum]?.factoryID    || sRow.Factory_ID        || '',
        Factory_Name:      (() => {
          const fId   = carrierPoMeta[poNum]?.factoryID    || sRow.Factory_ID    || '';
          const fName = carrierPoMeta[poNum]?.factoryName  || sRow.Factory_Name  || '';
          return String(fId).trim() === '9999' ? (fName || 'Dummy Factory') : fName;
        })(),
        Factory_Street1:   carrierPoMeta[poNum]?.factoryStreet1 || sRow.Factory_Street1   || '',
        Factory_Street2:   sRow.Factory_Street2   || '',
        Factory_Street3:   sRow.Factory_Street3   || '',
        Factory_City:      carrierPoMeta[poNum]?.factoryCity    || sRow.Factory_City      || '',
        Factory_PostalCd:  carrierPoMeta[poNum]?.factoryPostal  || sRow.Factory_PostalCd  || '',
        Factory_CountryCd: carrierPoMeta[poNum]?.factoryCountry || sRow.Factory_CountryCd || '',
        Country_Of_Origin:  carrierLine?.country   || '',

        // FC — resolved from FC_MASTER by FinalDestination from carrier feed (TradePartner FD)
        ...(() => {
          const fcId = carrierPoMeta[poNum]?.finalDestination || 'FC01';
          const fc   = FC_MASTER[fcId] || {};
          return {
            FC_ID:              fcId,
            FC_Name:            fc.name    || '',
            FC_Street1:         fc.street1 || '',
            FC_Street2:         fc.street2 || '',
            FC_Street3:         '',
            FC_City:            fc.city    || '',
            FC_StateProvinceCd: '',
            FC_PostalCd:        fc.postal  || '',
            FC_CountryCd:       fc.country || 'GB',
          };
        })(),

        Carrier_ID:          '3',
        Carrier_Name:        '',
        Loading_Port_LOCODE: carrierPoMeta[poNum]?.shippingPoint || '',
        POFC_ID:             carrierPoMeta[poNum]?.pofc           || '',
        F1_ID:               carrierPoMeta[poNum]?.pofc           || '',

        // SKU — carrier feed takes priority for enrichment, then PO feed
        SKU:           sku,
        Product_Style: poLine?.line?.productStyle || carrierLine?.style       || '',
        Description:   poLine?.line?.description  || carrierLine?.description || '',

        // Supplier-provided per SKU row — enrich from carrier feed if available
        EAN_Barcode:   sRow.EAN_Barcode || carrierLine?.ean    || '',
        Colour_Code:   sRow.Colour_Code || carrierLine?.colour || '',
        Size_Code:     sRow.Size_Code   || carrierLine?.size   || '',

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
        Pack_Type:       sRow.Pack_Type        || 'Bulk Flat',
        Collection_Type: sRow.Collection_Type  || 'Delivery',
        Hazardous:       sRow.Hazardous        || 'N/A',
        Traffic_Mode:    sRow.Traffic_Mode     || '',
        Mode_Of_Transport: carrierPoMeta[poNum]?.mode || sRow.Mode_Of_Transport || 'Sea',
        Cargo_Ready_Planned_Collection_Date: sRow.Cargo_Ready_Planned_Collection_Date || '',
        Carrier_Booking_Request_Date:        sRow.Carrier_Booking_Request_Date        || '',
        Ship_Date:                           carrierLine?.shipDate              || '',
        Expected_Delivery_Date:              carrierLine?.expectedDeliveryDate || sRow.Expected_Delivery_Date || '',
        ASN_Delivery_Date:                   sRow.ASN_Delivery_Date                  || '',
        Var_Unit: sRow.Var_Unit ?? 0,
        Var_Pct:  sRow.Var_Pct  ?? 0,
        Remarks:  sRow.Remarks  || '',
        ASOS_Intake_Week:    sRow.ASOS_Intake_Week || '',
        Collection_Time:     sRow.Collection_Time  || '',
        Incoterms:           carrierPoMeta[poNum]?.shippingTerms || '',
        Transport_Mode_Code: poLine?.line?.mode || po?.lineItems?.[0]?.mode || carrierPoMeta[poNum]?.mode || '30'
      });
  }

  // ── Second pass: carrier SKUs missing from supplier template ────────────────
  // Track which (poNum, sku) combos were covered by supplier rows
  const coveredKeys = new Set(masterRows.map(r => `${r.PO_Number}_${r.SKU}`));

  // Build a PO-level header lookup from already-built master rows so that
  // carrier-only SKUs inherit the same Booking_Group, dates, and factory data
  // as the other rows for the same PO (fixes mis-grouping of BK002, BK008 etc.)
  const poHeaderByPO = {};
  for (const r of masterRows) {
    if (!poHeaderByPO[r.PO_Number]) poHeaderByPO[r.PO_Number] = r;
  }

  if (hasCarrierData) {
    for (const [poNum, skuMap] of Object.entries(carrierAsnIndex)) {
      const po = poByOrderId[poNum];
      const poHdr = poHeaderByPO[poNum] || {};
      for (const [sku, carrierLine] of Object.entries(skuMap)) {
        if (coveredKeys.has(`${poNum}_${sku}`)) continue; // already in master rows
        // SKU is on carrier feed but supplier didn't include it — add with Booking_Qty=0
        const poLine = poByLinesku[`${poNum}_${sku}`];
        const ct = CARTON_TYPES['BDCM1'];
        masterRows.push({
          Booking_Ref:   poHdr.Booking_Ref   || '',
          Booking_Group: poHdr.Booking_Group || '',
          PO_Number:     poNum,
          ASN_Ref:       carrierLine.asnId,
          _missingFromSupplier: true,   // flag for Excel highlighting

          Supplier_Name:      carrierPoMeta[poNum]?.supplier      || '',
          Supplier_ID:        carrierPoMeta[poNum]?.supplierCode  || '',
          Supplier_Street1:   carrierPoMeta[poNum]?.supplierStreet1 || '',
          Supplier_City:      carrierPoMeta[poNum]?.supplierCity    || '',
          Supplier_PostalCd:  carrierPoMeta[poNum]?.supplierPostal  || '',
          Supplier_CountryCd: carrierPoMeta[poNum]?.supplierCountry || '',

          // Factory — inherit from supplier rows for this PO if available
          Factory_Name:       poHdr.Factory_Name      || '',
          Factory_ID:         poHdr.Factory_ID        || '',
          Factory_Street1:    poHdr.Factory_Street1   || '',
          Factory_Street2:    poHdr.Factory_Street2   || '',
          Factory_Street3:    poHdr.Factory_Street3   || '',
          Factory_City:       poHdr.Factory_City      || '',
          Factory_PostalCd:   poHdr.Factory_PostalCd  || '',
          Factory_CountryCd:  poHdr.Factory_CountryCd || '',
          Country_Of_Origin:   carrierLine.country || '',

          // FC — resolved from FC_MASTER by FinalDestination (TradePartner FD)
          ...(() => {
            const fcId = carrierPoMeta[poNum]?.finalDestination || carrierLine.finalDestination || 'FC01';
            const fc   = FC_MASTER[fcId] || {};
            return {
              FC_ID:              fcId,
              FC_Name:            fc.name    || '',
              FC_Street1:         fc.street1 || '',
              FC_Street2:         fc.street2 || '',
              FC_Street3:         '',
              FC_City:            fc.city    || '',
              FC_StateProvinceCd: '',
              FC_PostalCd:        fc.postal  || '',
              FC_CountryCd:       fc.country || 'GB',
            };
          })(),

          Carrier_ID:          '3',
          Carrier_Name:        '',
          Loading_Port_LOCODE: carrierPoMeta[poNum]?.shippingPoint || '',
          POFC_ID:             carrierPoMeta[poNum]?.pofc           || '',
          F1_ID:               carrierPoMeta[poNum]?.pofc           || '',

          SKU:           sku,
          Product_Style: poLine?.line?.productStyle || carrierLine.style       || '',
          Description:   poLine?.line?.description  || carrierLine.description || '',

          EAN_Barcode:   carrierLine.ean    || '',
          Colour_Code:   carrierLine.colour || '',
          Size_Code:     carrierLine.size   || '',

          // No carton/weight data — supplier hasn't provided it
          Carton_Type:      'BDCM1',
          Carton_Length_cm:  ct.L,
          Carton_Width_cm:   ct.W,
          Carton_Height_cm:  ct.H,
          Carton_Weight_KG:  ct.weight,
          No_of_Cartons:     0,
          Unit_Weight_KG:    0,
          Booking_Qty:       0,

          Gross_Weight_KG:   0,
          Net_Weight_KG:     0,
          Volume_M3:         0,

          Pack_Type:       carrierLine.packFormat === 'H' ? 'Hanging' : (poHdr.Pack_Type || 'Bulk Flat'),
          Collection_Type: poHdr.Collection_Type || 'Delivery',
          Hazardous:       poHdr.Hazardous       || 'N/A',
          Traffic_Mode:    poHdr.Traffic_Mode || po?.lineItems?.[0]?.mode || '',
          Mode_Of_Transport: carrierPoMeta[poNum]?.mode || poHdr.Mode_Of_Transport || 'Sea',
          Cargo_Ready_Planned_Collection_Date: poHdr.Cargo_Ready_Planned_Collection_Date || '',
          Carrier_Booking_Request_Date:        poHdr.Carrier_Booking_Request_Date        || '',
          Ship_Date:                           carrierLine.shipDate || '',
          Expected_Delivery_Date:              carrierLine.expectedDeliveryDate || '',
          ASN_Delivery_Date:                   '',
          Var_Unit: 0,
          Var_Pct:  0,
          Remarks:  'SKU from carrier feed — not found in supplier template',
          ASOS_Intake_Week:    '',
          Collection_Time:     '',
          Incoterms:           carrierPoMeta[poNum]?.shippingTerms || '',
          Transport_Mode_Code: poLine?.line?.mode || po?.lineItems?.[0]?.mode || carrierPoMeta[poNum]?.mode || '30'
        });
      }
    }
  }

  // ── Extra SKUs on carrier feed but not in supplier template ─────────────────
  const extraSkuWarnings = [];
  if (hasCarrierData) {
    for (const [poNum, skuMap] of Object.entries(carrierAsnIndex)) {
      for (const sku of Object.keys(skuMap)) {
        if (!coveredKeys.has(`${poNum}_${sku}`)) {
          const c = carrierAsnIndex[poNum][sku];
          extraSkuWarnings.push(`PO ${poNum} — SKU ${sku} is on the PO/ASN but was NOT in the supplier template — included in VBKREQ with Booking_Qty=0`);
        }
      }
    }
  }

  // Write Excel
  await writeExcel(masterRows, supplierRows, carrierAsnFiles);

  return { masterRows, filePath: BIBLE_FILE, warnings: skuWarnings, extraSkuWarnings };
}

async function writeExcel(masterRows, supplierRows, carrierAsnFiles) {
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
      const excelRow = ws.addRow(headers.map(h => {
        const v = sanitize(row[h]);
        // Store PO_Number as string to prevent Excel scientific notation
        return h === 'PO_Number' && v !== '' ? String(v) : v;
      }));
      // Highlight rows where carrier sent a SKU the supplier didn't include
      if (row._missingFromSupplier) {
        excelRow.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } }; // amber
        });
      }
    }
    ws.columns.forEach(col => { col.width = 22; });
    // Set PO_Number column to text format to prevent scientific notation
    const poColIdx = headers.indexOf('PO_Number') + 1;
    if (poColIdx > 0) ws.getColumn(poColIdx).numFmt = '@';
    return ws;
  }

  // Sheet 1: SUPPLIER_INPUT (from raw supplier rows)
  {
    const hdrs = [
      'PO_Number','ASN_Ref','SKU','No_of_Cartons','Unit_Weight_KG',
      'Cargo_Ready_Planned_Collection_Date','Carrier_Booking_Request_Date',
      'Traffic_Mode','Mode_Of_Transport',
      'Factory_Name','Factory_ID','Factory_Street1','Factory_Street2',
      'Factory_City','Factory_PostalCd','Factory_CountryCd',
      'EAN_Barcode','Colour_Code','Size_Code','Carton_Type',
      'Carton_Length_cm','Carton_Width_cm','Carton_Height_cm','Carton_Weight_KG',
      'Gross_Weight_KG','Net_Weight_KG','Volume_M3','Booking_Qty','Pack_Type',
      'Collection_Type','Collection_Time','Hazardous','Expected_Delivery_Date',
      'ASN_Delivery_Date','ASOS_Intake_Week','Var_Unit','Var_Pct','Remarks'
    ];
    addSheet('SUPPLIER_INPUT', hdrs, supplierRows);
  }

  // Sheet 2: CARRIER_FEED_EXTRACT — flattened from carrier ASN files
  {
    const carrierFlat = [];
    for (const file of (carrierAsnFiles || [])) {
      for (const asnGroup of (file.parsed || [])) {
        for (const line of (asnGroup.lines || [])) {
          carrierFlat.push({
            filename:      file.filename,
            blobPath:      file.blobPath || '',
            asnId:         asnGroup.asnId,
            poId:          asnGroup.poId,
            fcId:          asnGroup.fcId,
            shipDate:      asnGroup.shipDate,
            supplier:      asnGroup.supplier,
            supplierCode:  asnGroup.supplierCode,
            shippingPoint: asnGroup.shippingPoint,
            shippingTerms: asnGroup.shippingTerms,
            sku:           line.sku,
            ean:           line.ean,
            description:   line.description,
            size:          line.size,
            colour:        line.colour,
            style:         line.style,
            packFormat:    line.packFormat,
            country:       line.country,
            quantity:      line.quantity
          });
        }
      }
    }
    const hdrs = ['filename','blobPath','asnId','poId','fcId','shipDate','supplier',
                  'supplierCode','shippingPoint','shippingTerms',
                  'sku','ean','description','size','colour','style','packFormat','country','quantity'];
    addSheet('CARRIER_FEED_EXTRACT', hdrs, carrierFlat);
  }

  // Sheet 3: MASTER
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

  // Sheet 4: CARTON_TYPES
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

  // Sheet 5: GENERATION_LOG (from JSON sidecar — no Excel re-read needed)
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

function updateGenerationLog(filename, ctrlNumber, patch) {
  const entries = readLogFromJson();
  const idx = entries.slice().reverse().findIndex(e => e.filename === filename && e.ctrlNumber === ctrlNumber);
  if (idx !== -1) {
    const realIdx = entries.length - 1 - idx;
    entries[realIdx] = { ...entries[realIdx], ...patch };
    writeLogToJson(entries);
  }
}

function getGenerationLog() {
  return readLogFromJson();
}

module.exports = { build, appendGenerationLog, updateGenerationLog, getGenerationLog };
