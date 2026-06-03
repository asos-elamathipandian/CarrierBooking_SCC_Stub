'use strict';

const { create } = require('xmlbuilder2');
const fs = require('fs');
const path = require('path');

const COUNTER_FILE = path.join(__dirname, 'ctrl-counter.json');

// ─── Carton type master ───────────────────────────────────────────────────────
const CARTON_TYPES = {
  'BDCM1':               { weight: 1.40, L: 60.00, W: 30.00, H: 40.00 },
  'BDCM3':               { weight: 1.00, L: 45.00, W: 29.50, H: 18.80 },
  'C5':                  { weight: 1.00, L: 60.00, W: 30.00, H: 20.00 },
  'Cartons':             { weight: 1.00, L: 45.00, W: 60.00, H: 40.00 },
  'A1':                  { weight: 1.00, L: 59.50, W: 28.50, H: 37.50 },
  'A2':                  { weight: 1.00, L: 59.50, W: 28.50, H: 32.50 },
  'A3':                  { weight: 1.00, L: 59.50, W: 28.50, H: 26.00 },
  'A4':                  { weight: 1.00, L: 59.50, W: 28.50, H: 19.00 },
  'B1':                  { weight: 1.00, L: 52.00, W: 25.50, H: 37.50 },
  'B2':                  { weight: 1.00, L: 52.00, W: 25.50, H: 32.50 },
  'B3':                  { weight: 1.00, L: 52.00, W: 25.50, H: 26.00 },
  'B4':                  { weight: 1.00, L: 52.00, W: 25.50, H: 19.00 },
  'C1':                  { weight: 1.00, L: 45.00, W: 28.50, H: 37.50 },
  'C2':                  { weight: 1.00, L: 45.00, W: 28.50, H: 32.50 },
  'C3':                  { weight: 1.00, L: 45.00, W: 28.50, H: 26.00 },
  'C4':                  { weight: 1.00, L: 45.00, W: 28.50, H: 19.00 },
  'Hanging':             { weight: 0.70, L: 213.00, W: 94.00, H: 60.00 }
};

// FC → Destination port LOCODE lookup
const FC_LOCODE = {
  'FC01': 'GBBSY',
  'FC02': 'GBBSY',
  'P005': 'GBBSY',
  'POROP': 'POROP'
};

function getCtrlNumber() {
  let data = { counter: 91800256 };
  if (fs.existsSync(COUNTER_FILE)) {
    try { data = JSON.parse(fs.readFileSync(COUNTER_FILE, 'utf8')); } catch (_) {}
  }
  const current = data.counter;
  data.counter = current + 1;
  fs.writeFileSync(COUNTER_FILE, JSON.stringify(data, null, 2));
  return String(current);
}

function formatDateYMD(dateVal) {
  if (!dateVal) return '';
  // Accept DD/MM/YYYY or YYYY-MM-DD or Date object
  if (dateVal instanceof Date) {
    return dateVal.toISOString().slice(0, 10).replace(/-/g, '');
  }
  const s = String(dateVal).trim();
  if (/^\d{8}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10).replace(/-/g, '');
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    const [d, m, y] = s.split('/');
    return `${y}${m}${d}`;
  }
  return s.replace(/[-/]/g, '').slice(0, 8);
}

function nowDateTimeStr() {
  const n = new Date();
  const pad = v => String(v).padStart(2, '0');
  return `${n.getFullYear()}${pad(n.getMonth()+1)}${pad(n.getDate())} ${pad(n.getHours())}${pad(n.getMinutes())}${pad(n.getSeconds())}`;
}

function nowFilenameStr() {
  return nowDateTimeStr().replace(' ', '').replace(/:/g, '');
}

function hazardousCode(val) {
  if (!val || val === 'N/A') return 'N';
  return 'Y';
}

/**
 * Build VBKREQ XML from masterRows array.
 * Each row represents one PO line item.
 * Rows from the same booking are grouped by Booking_Ref.
 */
async function build(masterRows) {
  if (!masterRows || masterRows.length === 0) throw new Error('No master rows to generate VBKREQ');

  const ctrlNumber = getCtrlNumber();
  const now = nowDateTimeStr();
  const filenameTs = nowFilenameStr();
  const filename = `DAVIESTN_E2ASOS_VBKREQ_1.0_${filenameTs}${ctrlNumber}.xml`;

  // Use first row for booking-level fields
  const first = masterRows[0];
  const bookingRef = first.Booking_Ref || `VB-STUB-${ctrlNumber}`;
  const trafficMode = first.Traffic_Mode || 'CFS';
  const originCountry = first.Factory_CountryCd || 'XX';
  const hazCode = hazardousCode(first.Hazardous);
  const collectionType = first.Collection_Type || 'Delivery';
  const cargoReadyDate = formatDateYMD(first.Cargo_Ready_Planned_Collection_Date);
  const bookingReqDate = formatDateYMD(first.Carrier_Booking_Request_Date);
  const asnDeliveryDate = formatDateYMD(first.ASN_Delivery_Date);
  const expectedDeliveryDate = formatDateYMD(first.Expected_Delivery_Date);
  const loadingPortLocode = first.Loading_Port_LOCODE || first.Loading_Port_ID || 'XXXX';
  const fcId = first.FC_ID || 'FC01';
  const destLocode = FC_LOCODE[fcId] || 'GBBSY';

  // Totals
  let totalNet = 0, totalGross = 0, totalVol = 0, totalCartons = 0, totalBkq = 0;
  for (const row of masterRows) {
    const ct = CARTON_TYPES[row.Carton_Type] || {};
    const noCartons = parseFloat(row.No_of_Cartons) || 0;
    const unitWt = parseFloat(row.Unit_Weight_KG) || 0;
    const bkq = parseFloat(row.Booking_Qty) || 0;
    const cL = parseFloat(row.Carton_Length_cm) || ct.L || 0;
    const cW = parseFloat(row.Carton_Width_cm) || ct.W || 0;
    const cH = parseFloat(row.Carton_Height_cm) || ct.H || 0;
    const cWt = parseFloat(row.Carton_Weight_KG) || ct.weight || 0;

    row._gross = parseFloat((cWt * noCartons).toFixed(4));
    row._net   = parseFloat((unitWt * bkq).toFixed(4));
    row._vol   = parseFloat(((cL * cW * cH / 1000000) * noCartons).toFixed(4));
    row._cartons = noCartons;
    row._bkq   = bkq;

    totalGross   += row._gross;
    totalNet     += row._net;
    totalVol     += row._vol;
    totalCartons += row._cartons;
    totalBkq     += row._bkq;
  }

  // Group lines by PO
  const poGroups = {};
  for (const row of masterRows) {
    const po = row.PO_Number;
    if (!poGroups[po]) poGroups[po] = [];
    poGroups[po].push(row);
  }

  // Build XML
  const root = create({ version: '1.0', encoding: 'UTF-8' })
    .ele('XMLBundle');

  const transmission = root.ele('XMLTransmission', {
    CtrlNumber: ctrlNumber,
    Receiver: 'E2ASOS',
    Sender: 'DAVIESTN',
    Timestamp: now
  });

  const group = transmission.ele('XMLGroup', {
    CtrlNumber: ctrlNumber,
    GroupType: 'BP',
    IncludedMessages: '1'
  });

  const transaction = group.ele('XMLTransaction', {
    CtrlNumber: ctrlNumber,
    TransactionType: 'BPM-VBKREQ'
  });

  const bpMsg = transaction.ele('BpMessage', {
    MessageType: 'VBKREQ',
    PurposeCd: '15'
  });

  // Header
  bpMsg.ele('Mode').txt('30');
  bpMsg.ele('Reference', { RefTypeCd: 'QY', SourceRefTypeCd: '128' }).txt(trafficMode);
  bpMsg.ele('Reference', { RefTypeCd: '4B', SourceRefTypeCd: '128' }).txt(originCountry);
  bpMsg.ele('Reference', { RefTypeCd: 'BH', SourceRefTypeCd: '128' }).txt(hazCode);
  bpMsg.ele('Reference', { RefTypeCd: 'SFZ', SourceRefTypeCd: '128' }).txt('13');
  bpMsg.ele('Reference', { RefTypeCd: 'CC', SourceRefTypeCd: '128' }).txt('Green');
  bpMsg.ele('Reference', { RefTypeCd: 'CD', SourceRefTypeCd: '128' }).txt(collectionType);

  // TradePartner: SU
  const tpSU = bpMsg.ele('TradePartner', { RoleCd: 'SU' });
  tpSU.ele('TradePartnerName').txt(first.Supplier_Name || '');
  tpSU.ele('TradePartnerID', { Qualifier: '93' }).txt(first.Supplier_ID || '');

  // TradePartner: FA (Factory — from address)
  const tpFA = bpMsg.ele('TradePartner', { RoleCd: 'FA' });
  tpFA.ele('TradePartnerName').txt(first.Factory_Name || '');
  tpFA.ele('TradePartnerID', { Qualifier: '93' }).txt(first.Factory_ID || '');
  const addrFA = tpFA.ele('TradePartnerAddress');
  const streets = [first.Factory_Street1, first.Factory_Street2, first.Factory_Street3].filter(Boolean);
  streets.forEach(s => addrFA.ele('Street').txt(s));
  if (first.Factory_City) addrFA.ele('City').txt(first.Factory_City);
  if (first.Factory_PostalCd) addrFA.ele('PostalCd').txt(first.Factory_PostalCd);
  if (first.Factory_CountryCd) addrFA.ele('CountryCd').txt(first.Factory_CountryCd);

  // TradePartner: FD with address
  const tpFD = bpMsg.ele('TradePartner', { RoleCd: 'FD' });
  tpFD.ele('TradePartnerName').txt(first.FC_Name || '');
  tpFD.ele('TradePartnerID', { Qualifier: '93' }).txt(fcId);
  const addrFD = tpFD.ele('TradePartnerAddress');
  const fcStreets = [first.FC_Street1, first.FC_Street2, first.FC_Street3].filter(Boolean);
  fcStreets.forEach(s => addrFD.ele('Street').txt(s));
  if (first.FC_City) addrFD.ele('City').txt(first.FC_City);
  if (first.FC_StateProvinceCd) addrFD.ele('StateProvinceCd').txt(first.FC_StateProvinceCd);
  if (first.FC_PostalCd) addrFD.ele('PostalCd').txt(first.FC_PostalCd);
  addrFD.ele('CountryCd').txt(first.FC_CountryCd || 'GB');

  // TradePartner: CA
  const tpCA = bpMsg.ele('TradePartner', { RoleCd: 'CA' });
  tpCA.ele('TradePartnerID', { Qualifier: '93' }).txt(first.Carrier_ID || '3');

  // TradePartner: SL
  const tpSL = bpMsg.ele('TradePartner', { RoleCd: 'SL' });
  tpSL.ele('TradePartnerID', { Qualifier: '93' }).txt(loadingPortLocode);

  // TradePartner: FD without address (second occurrence)
  const tpFD2 = bpMsg.ele('TradePartner', { RoleCd: 'FD' });
  tpFD2.ele('TradePartnerID', { Qualifier: '93' }).txt(`${fcId} ${first.FC_Name || ''}`.trim());

  // Status elements
  bpMsg.ele('Status').ele('Date', { DateTypeCd: '018', TimeZone: 'LT' }).txt(cargoReadyDate);
  bpMsg.ele('Status').ele('Date', { DateTypeCd: '081', TimeZone: 'LT' }).txt(cargoReadyDate);

  const stL = bpMsg.ele('Status');
  stL.ele('Location', { LocTypeCd: 'L' }).ele('LocationID', { Qualifier: 'UN' }).txt(loadingPortLocode);

  const stE = bpMsg.ele('Status');
  stE.ele('Location', { LocTypeCd: 'E' }).ele('LocationID', { Qualifier: 'UN' }).txt(destLocode);

  const stD = bpMsg.ele('Status');
  stD.ele('Location', { LocTypeCd: 'D' }).ele('LocationID', { Qualifier: 'UN' }).txt(destLocode);

  const bookingReqDateTime = `${bookingReqDate} ${now.split(' ')[1]}`;
  bpMsg.ele('Status').ele('Date', { DateTypeCd: '211', TimeZone: 'LT' }).txt(bookingReqDateTime);
  bpMsg.ele('Status').ele('Date', { DateTypeCd: 'OSBT', TimeZone: 'LT' }).txt(bookingReqDateTime);

  if (asnDeliveryDate) bpMsg.ele('Status').ele('Date', { DateTypeCd: '238' }).txt(asnDeliveryDate);
  if (expectedDeliveryDate) bpMsg.ele('Status').ele('Date', { DateTypeCd: '065' }).txt(expectedDeliveryDate);

  bpMsg.ele('Status').ele('Date', { DateTypeCd: 'OSBK' }).txt(now);
  bpMsg.ele('Status').ele('Date', { DateTypeCd: 'SBK' }).txt(now);

  // Document
  const doc = bpMsg.ele('Document', { DocType: 'BOOK', Key: bookingRef });
  doc.ele('Reference', { RefTypeCd: 'ACE', SourceRefTypeCd: '128' }).txt(bookingRef);
  doc.ele('Reference', { RefTypeCd: 'V0',  SourceRefTypeCd: '128' }).txt('2.0');
  doc.ele('Measure', { Qualifier: 'N',   SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'KG' }).txt(totalNet.toFixed(4));
  doc.ele('Measure', { Qualifier: 'G',   SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'KG' }).txt(totalGross.toFixed(4));
  doc.ele('Measure', { Qualifier: 'VOL', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'M3' }).txt(totalVol.toFixed(4));
  doc.ele('Measure', { Qualifier: 'QUR', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'CT' }).txt(totalCartons.toFixed(4));
  doc.ele('Measure', { Qualifier: 'BKQ', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'UN' }).txt(totalBkq.toFixed(6));

  // Orders
  for (const [poNum, lines] of Object.entries(poGroups)) {
    const order = doc.ele('Order', { Key: poNum, OrderType: 'PO' });
    order.ele('OrderID').txt(poNum);

    for (const row of lines) {
      const ct = CARTON_TYPES[row.Carton_Type] || {};
      const cL = parseFloat(row.Carton_Length_cm) || ct.L || 0;
      const cW = parseFloat(row.Carton_Width_cm)  || ct.W || 0;
      const cH = parseFloat(row.Carton_Height_cm) || ct.H || 0;
      const sku  = row.SKU || '';
      const ean  = row.EAN_Barcode || '';
      const lineKey = `${poNum}_${sku}_${ean}`;
      const description = row.Description || row.Product_Style || `SKU ${sku}`;
      const cartonType = row.Carton_Type || 'BDCM1';
      const packType = row.Pack_Type || 'Bulk Flat';
      const productStyle = row.Product_Style || '';
      const hazRef = row.Hazardous || 'N/A';
      const lineFC = row.FC_ID || fcId;

      const li = order.ele('LineItem', { Key: lineKey });
      li.ele('LineItemDescription').txt(description);
      li.ele('Attribute', { AttributeTypeCd: 'SI' }).txt(ean);
      li.ele('Attribute', { AttributeTypeCd: 'SK' }).txt(sku);
      li.ele('Attribute', { AttributeTypeCd: 'CL' }).txt(row.Colour_Code || '');
      li.ele('Attribute', { AttributeTypeCd: 'CM' }).txt('2003');
      li.ele('Attribute', { AttributeTypeCd: 'IZ' }).txt(row.Size_Code || '');
      li.ele('Reference', { RefTypeCd: 'PAC', SourceRefTypeCd: '128' }).txt(packType);
      li.ele('Reference', { RefTypeCd: 'PT',  SourceRefTypeCd: '128' }).txt(productStyle);
      li.ele('Reference', { RefTypeCd: 'HZ',  SourceRefTypeCd: '128' }).txt(hazRef);
      li.ele('Reference', { RefTypeCd: 'DSC', SourceRefTypeCd: '128' }).txt(description);
      li.ele('Reference', { RefTypeCd: '98',  SourceRefTypeCd: '128' }).txt(cartonType);
      li.ele('Reference', { RefTypeCd: 'LN',  SourceRefTypeCd: '128' }).txt(cL.toFixed(4));
      li.ele('Reference', { RefTypeCd: 'WD',  SourceRefTypeCd: '128' }).txt(cW.toFixed(4));
      li.ele('Reference', { RefTypeCd: 'HT',  SourceRefTypeCd: '128' }).txt(cH.toFixed(4));
      li.ele('Measure', { Qualifier: 'BKQ', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'UN' }).txt(row._bkq.toFixed(6));
      li.ele('Measure', { Qualifier: 'G',   SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'KG' }).txt(row._gross.toFixed(4));
      li.ele('Measure', { Qualifier: 'N',   SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'KG' }).txt(row._net.toFixed(4));
      li.ele('Measure', { Qualifier: 'VOL', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'M3' }).txt(row._vol.toFixed(4));
      li.ele('Measure', { Qualifier: 'QUR', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'CT' }).txt(row._cartons.toFixed(4));
      li.ele('TradePartner', { RoleCd: 'FS' }).ele('TradePartnerID', { Qualifier: '93' }).txt(lineFC);
    }
  }

  const xml = root.end({ prettyPrint: true });
  return { xml, filename, ctrlNumber };
}

module.exports = { build, CARTON_TYPES };
