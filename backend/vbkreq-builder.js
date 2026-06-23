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

// Transport mode string → code mapping
const MODE_MAP = {
  '10': '10', '30': '30', '40': '40', '50': '50', '60': '60', '70': '70',
  'SEA': '10', 'OCEAN': '10', 'FCL': '10', 'LCL': '10',
  'ROAD': '30', 'TRUCK': '30',
  'AIR': '40', 'AIR-ASOS': '40', 'AIR ASOS': '40',
  'RAIL': '50',
  'AIR-SUPPLIER': '60', 'AIR SUPPLIER': '60',
  'ECO': '70', 'ECO SEA': '70', 'ECO AIR': '70', 'ECO SEA/AIR': '70'
};

function resolveMode(modeVal) {
  if (!modeVal) return '30';
  const key = String(modeVal).toUpperCase().trim();
  return MODE_MAP[key] || String(modeVal).trim() || '30';
}

function getCtrlNumber() {
  let data = { counter: 91800256, bookingCounter: 1000000001, bookingVersions: {} };
  if (fs.existsSync(COUNTER_FILE)) {
    try { data = JSON.parse(fs.readFileSync(COUNTER_FILE, 'utf8')); } catch (_) {}
  }
  if (!data.bookingVersions) data.bookingVersions = {};
  if (!data.bookingCounter)  data.bookingCounter  = 1000000001;
  const current = data.counter;
  data.counter = current + 1;
  fs.writeFileSync(COUNTER_FILE, JSON.stringify(data, null, 2));
  return String(current);
}

function getBookingRef() {
  let data = { counter: 91800256, bookingCounter: 1000000001, bookingVersions: {} };
  if (fs.existsSync(COUNTER_FILE)) {
    try { data = JSON.parse(fs.readFileSync(COUNTER_FILE, 'utf8')); } catch (_) {}
  }
  if (!data.bookingCounter) data.bookingCounter = 1000000001;
  const ref = `VB-${data.bookingCounter}`;
  data.bookingCounter += 1;
  fs.writeFileSync(COUNTER_FILE, JSON.stringify(data, null, 2));
  return ref;
}

/**
 * Get/manage version for a booking ref.
 * purposeCd 13 (new)  → version 1.0, stores 1
 * purposeCd 15 (resub) → increments stored version
 * purposeCd 01 (cancel) → returns existing version unchanged
 */
function getBookingVersion(bookingRef, purposeCd) {
  let data = { counter: 91800256, bookingCounter: 1000000001, bookingVersions: {} };
  if (fs.existsSync(COUNTER_FILE)) {
    try { data = JSON.parse(fs.readFileSync(COUNTER_FILE, 'utf8')); } catch (_) {}
  }
  if (!data.bookingVersions) data.bookingVersions = {};
  if (!data.bookingCounter)  data.bookingCounter  = 1000000001;

  let version;
  if (purposeCd === '15') {
    const current = data.bookingVersions[bookingRef] || 1;
    version = current + 1;
    data.bookingVersions[bookingRef] = version;
  } else if (purposeCd === '01') {
    version = data.bookingVersions[bookingRef] || 1;
  } else {
    // 13 — new request: start at 1, but preserve any existing version so a
    // re-run of the same Booking_Ref doesn't silently reset the counter.
    version = data.bookingVersions[bookingRef] || 1;
    data.bookingVersions[bookingRef] = version;
  }

  fs.writeFileSync(COUNTER_FILE, JSON.stringify(data, null, 2));
  return `${version}.0`;
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
 * purposeCd: '13' = Request (default), '15' = Re-Submission, '01' = Cancellation
 */
async function build(masterRows, purposeCd) {
  if (!masterRows || masterRows.length === 0) throw new Error('No master rows to generate VBKREQ');

  const pcd = purposeCd || '13';
  const ctrlNumber = getCtrlNumber();
  const now = nowDateTimeStr();
  const filenameTs = nowFilenameStr();
  const filename = `DAVIESTN_E2ASOS_VBKREQ_1.0_${filenameTs}${ctrlNumber}.xml`;

  // Use first row for booking-level fields
  const first = masterRows[0];
  const bookingRef = first.Booking_Ref || getBookingRef();
  const version = getBookingVersion(bookingRef, pcd);
  const trafficMode = first.Traffic_Mode || 'CFS';
  const originCountry = first.Country_Of_Origin || first.Factory_CountryCd || 'XX';
  const hazCode = hazardousCode(first.Hazardous);
  const collectionType = first.Collection_Type || 'Delivery';
  const collectionTime = first.Collection_Time || '';
  const intakeWeek = first.ASOS_Intake_Week || '';
  const remarks = first.Remarks || '';
  const cargoReadyDate = formatDateYMD(first.Cargo_Ready_Planned_Collection_Date);
  const bookingReqDate = formatDateYMD(first.Carrier_Booking_Request_Date);
  const shipDate = formatDateYMD(first.Ship_Date);
  const asnDeliveryDate = formatDateYMD(first.ASN_Delivery_Date);
  const expectedDeliveryDate = formatDateYMD(first.Expected_Delivery_Date);
  const loadingPortLocode = first.Loading_Port_LOCODE || first.Loading_Port_ID || 'XXXX';
  const fcId = first.FC_ID || 'FC01';
  const destLocode = FC_LOCODE[fcId] || 'GBBSY';
  // Transport mode: from Mode_Of_Transport column (Sea/Air/Road etc), fallback to 30 (Road)
  const transportModeCode = resolveMode(first.Mode_Of_Transport || '30');

  // Totals — pre-compute per-row values used in Document sections
  for (const row of masterRows) {
    const ct = CARTON_TYPES[row.Carton_Type] || {};
    const noCartons = parseFloat(row.No_of_Cartons) || 0;
    const unitWt = parseFloat(row.Unit_Weight_KG) || 0;
    const bkq = parseFloat(row.Booking_Qty) || 0;
    const cL = parseFloat(row.Carton_Length_cm) || ct.L || 0;
    const cW = parseFloat(row.Carton_Width_cm) || ct.W || 0;
    const cH = parseFloat(row.Carton_Height_cm) || ct.H || 0;
    const cWt = parseFloat(row.Carton_Weight_KG) || ct.weight || 0;

    row._gross   = parseFloat((cWt * noCartons).toFixed(4));
    row._net     = parseFloat(unitWt.toFixed(4));   // N = Unit_Weight_KG per line
    row._vol     = parseFloat(((cL * cW * cH / 1000000) * noCartons).toFixed(4));
    row._cartons = noCartons;
    row._bkq     = bkq;
  }

  const root = create({ version: '1.0', encoding: 'UTF-8' })
    .ele('XMLBundle');

  const transmission = root.ele('XMLTransmission', {
    CtrlNumber: ctrlNumber,
    Receiver: 'DAVIESTN',
    Sender: 'E2ASOS',
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
    PurposeCd: pcd
  });

  // Header
  bpMsg.ele('Mode').txt(transportModeCode);
  bpMsg.ele('Reference', { RefTypeCd: 'QY',  SourceRefTypeCd: '128' }).txt(trafficMode);
  bpMsg.ele('Reference', { RefTypeCd: '4B',  SourceRefTypeCd: '128' }).txt(originCountry);
  bpMsg.ele('Reference', { RefTypeCd: 'BH',  SourceRefTypeCd: '128' }).txt(hazCode);
  bpMsg.ele('Reference', { RefTypeCd: 'CC',  SourceRefTypeCd: '128' }).txt('Green');
  bpMsg.ele('Reference', { RefTypeCd: 'CD',  SourceRefTypeCd: '128' }).txt(collectionType);
  // CT (Collection Time) — mandatory when collection type is Collection
  if (collectionType === 'Collection' || collectionTime) {
    bpMsg.ele('Reference', { RefTypeCd: 'CT', SourceRefTypeCd: '128' }).txt(collectionTime || '09:00');
  }
  // WK (ASOS Intake Week) — include if provided
  if (intakeWeek) {
    bpMsg.ele('Reference', { RefTypeCd: 'WK', SourceRefTypeCd: '128' }).txt(intakeWeek);
  }
  // Booking Request Comments remark
  if (remarks) {
    bpMsg.ele('Remark', { Qualifier: 'BRC' }).txt(remarks);
  }

  // TradePartner: SU
  const tpSU = bpMsg.ele('TradePartner', { RoleCd: 'SU' });
  tpSU.ele('TradePartnerName').txt(first.Supplier_Name || '');
  tpSU.ele('TradePartnerID', { Qualifier: '93' }).txt(first.Supplier_ID || '');

  // TradePartner: FA (Factory — from address)
  const tpFA = bpMsg.ele('TradePartner', { RoleCd: 'FA' });
  tpFA.ele('TradePartnerName').txt(first.Factory_Name || '');
  tpFA.ele('TradePartnerID', { Qualifier: '93' }).txt(first.Factory_ID || '');
  const streets = [first.Factory_Street1, first.Factory_Street2, first.Factory_Street3].filter(Boolean);
  const hasFactoryAddress = streets.length > 0 || first.Factory_City || first.Factory_PostalCd || first.Factory_CountryCd;
  if (hasFactoryAddress) {
    const addrFA = tpFA.ele('TradePartnerAddress');
    streets.forEach(s => addrFA.ele('Street').txt(s));
    if (first.Factory_City) addrFA.ele('City').txt(first.Factory_City);
    if (first.Factory_PostalCd) addrFA.ele('PostalCd').txt(first.Factory_PostalCd);
    if (first.Factory_CountryCd) addrFA.ele('CountryCd').txt(first.Factory_CountryCd);
  }

  // TradePartner: FD — Final Destination (from <FinalDestination> in carrier feed, FC address lookup)
  const FC_ADDRESS_LOOKUP = {
    'FC01': {
      name: 'FC01 Barnsley',
      streets: ['Greater London House', 'Hampstead Road -', 'London'],
      city: 'London',
      stateProvinceCd: 'YorkShire',
      postalCd: 'NW1 7FB',
      countryCd: 'GB'
    }
  };
  const tpFD = bpMsg.ele('TradePartner', { RoleCd: 'FD' });
  const fcLookup = FC_ADDRESS_LOOKUP[fcId];
  tpFD.ele('TradePartnerName').txt(fcLookup ? fcLookup.name : (first.FC_Name || ''));
  tpFD.ele('TradePartnerID', { Qualifier: '93' }).txt(fcId);
  const addrFD = tpFD.ele('TradePartnerAddress');
  if (fcLookup) {
    fcLookup.streets.forEach(s => addrFD.ele('Street').txt(s));
    addrFD.ele('City').txt(fcLookup.city);
    addrFD.ele('StateProvinceCd').txt(fcLookup.stateProvinceCd);
    addrFD.ele('PostalCd').txt(fcLookup.postalCd);
    addrFD.ele('CountryCd').txt(fcLookup.countryCd);
  } else {
    const fcStreets = [first.FC_Street1, first.FC_Street2, first.FC_Street3].filter(Boolean);
    fcStreets.forEach(s => addrFD.ele('Street').txt(s));
    if (first.FC_City) addrFD.ele('City').txt(first.FC_City);
    if (first.FC_StateProvinceCd) addrFD.ele('StateProvinceCd').txt(first.FC_StateProvinceCd);
    if (first.FC_PostalCd) addrFD.ele('PostalCd').txt(first.FC_PostalCd);
    addrFD.ele('CountryCd').txt(first.FC_CountryCd || 'GB');
  }

  // TradePartner: CA
  const tpCA = bpMsg.ele('TradePartner', { RoleCd: 'CA' });
  tpCA.ele('TradePartnerID', { Qualifier: '93' }).txt(first.Carrier_ID || '3');

  // TradePartner: SL
  const tpSL = bpMsg.ele('TradePartner', { RoleCd: 'SL' });
  tpSL.ele('TradePartnerID', { Qualifier: '93' }).txt(loadingPortLocode);

  // Status elements
  bpMsg.ele('Status').ele('Date', { DateTypeCd: '018', TimeZone: 'LT' }).txt(cargoReadyDate);
  bpMsg.ele('Status').ele('Date', { DateTypeCd: '081', TimeZone: 'LT' }).txt(bookingReqDate);

  const stL = bpMsg.ele('Status');
  stL.ele('Location', { LocTypeCd: 'L' }).ele('LocationID', { Qualifier: 'UN' }).txt(loadingPortLocode);

  const stE = bpMsg.ele('Status');
  stE.ele('Location', { LocTypeCd: 'E' }).ele('LocationID', { Qualifier: 'UN' }).txt(destLocode);

  const stD = bpMsg.ele('Status');
  stD.ele('Location', { LocTypeCd: 'D' }).ele('LocationID', { Qualifier: 'UN' }).txt(destLocode);

  bpMsg.ele('Status').ele('Date', { DateTypeCd: '211', TimeZone: 'LT' }).txt(now);
  bpMsg.ele('Status').ele('Date', { DateTypeCd: 'OSBT', TimeZone: 'LT' }).txt(now);

  if (shipDate) bpMsg.ele('Status').ele('Date', { DateTypeCd: '238' }).txt(shipDate);
  if (expectedDeliveryDate) bpMsg.ele('Status').ele('Date', { DateTypeCd: '065' }).txt(expectedDeliveryDate);

  bpMsg.ele('Status').ele('Date', { DateTypeCd: 'OSBK' }).txt(now);
  bpMsg.ele('Status').ele('Date', { DateTypeCd: 'SBK' }).txt(now);

  // Cancellation date — only for PurposeCd 01
  if (pcd === '01') {
    bpMsg.ele('Status').ele('Date', { DateTypeCd: '177', TimeZone: 'LT' }).txt(now);
  }

  // One Document per VBKREQ — aggregate totals across all rows
  const totalNet     = parseFloat(masterRows.reduce((s, r) => s + r._net,     0).toFixed(4));
  const totalGross   = parseFloat(masterRows.reduce((s, r) => s + r._gross,   0).toFixed(4));
  const totalVol     = parseFloat(masterRows.reduce((s, r) => s + r._vol,     0).toFixed(4));
  const totalCartons = parseFloat(masterRows.reduce((s, r) => s + r._cartons, 0).toFixed(4));
  const totalBkq     = parseFloat(masterRows.reduce((s, r) => s + r._bkq,     0).toFixed(6));

  const doc = bpMsg.ele('Document', { DocType: 'BOOK', Key: bookingRef });
  doc.ele('Reference', { RefTypeCd: 'ACE', SourceRefTypeCd: '128' }).txt(bookingRef);
  doc.ele('Reference', { RefTypeCd: 'V0',  SourceRefTypeCd: '128' }).txt(version);
  doc.ele('Measure', { Qualifier: 'N',   SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'KG' }).txt(totalNet.toFixed(4));
  doc.ele('Measure', { Qualifier: 'G',   SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'KG' }).txt(totalGross.toFixed(4));
  doc.ele('Measure', { Qualifier: 'VOL', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'M3' }).txt(totalVol.toFixed(4));
  doc.ele('Measure', { Qualifier: 'QUR', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'CT' }).txt(totalCartons.toFixed(4));
  doc.ele('Measure', { Qualifier: 'BKQ', SourceQualifier: '738', SourceUOMCd: '355', UOMCd: 'UN' }).txt(totalBkq.toFixed(6));

  // Group by PO — one Order per PO, line items carry the ASN ref
  const poGroups = {};
  for (const row of masterRows) {
    const poKey = String(row.PO_Number || '');
    if (!poGroups[poKey]) poGroups[poKey] = [];
    poGroups[poKey].push(row);
  }

  for (const [poNum, lines] of Object.entries(poGroups)) {
    const order = doc.ele('Order', { Key: poNum, OrderType: 'PO' });
    order.ele('OrderID').txt(poNum);

    for (const row of lines) {
      const asnRef = String(row.ASN_Ref || '');
      const ct = CARTON_TYPES[row.Carton_Type] || {};
      const cL = parseFloat(row.Carton_Length_cm) || ct.L || 0;
      const cW = parseFloat(row.Carton_Width_cm)  || ct.W || 0;
      const cH = parseFloat(row.Carton_Height_cm) || ct.H || 0;
      const sku          = row.SKU || '';
      const lineKey      = `${poNum}_${sku}_${asnRef}`;
      const description  = row.Description || row.Product_Style || `SKU ${sku}`;
      const cartonType   = row.Carton_Type || 'BDCM1';
      const packType     = row.Pack_Type || 'Bulk Flat';
      const productStyle = row.Product_Style || '';
      const hazRef2      = row.Hazardous || 'N/A';
      const lineFC       = row.FC_ID || fcId;

      const li = order.ele('LineItem', { Key: lineKey });
      li.ele('LineItemDescription').txt(description);
      li.ele('Attribute', { AttributeTypeCd: 'SI' }).txt(asnRef);
      li.ele('Attribute', { AttributeTypeCd: 'SK' }).txt(sku);
      if (row.Colour_Code)   li.ele('Attribute', { AttributeTypeCd: 'CL' }).txt(row.Colour_Code);
      if (row.Size_Code)     li.ele('Attribute', { AttributeTypeCd: 'IZ' }).txt(row.Size_Code);
      li.ele('Reference', { RefTypeCd: 'PAC', SourceRefTypeCd: '128' }).txt(packType);
      li.ele('Reference', { RefTypeCd: 'PT',  SourceRefTypeCd: '128' }).txt(productStyle);
      li.ele('Reference', { RefTypeCd: 'HZ',  SourceRefTypeCd: '128' }).txt(hazRef2);
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
  return { xml, filename, ctrlNumber, version, bookingRef };
}

module.exports = { build, CARTON_TYPES };
