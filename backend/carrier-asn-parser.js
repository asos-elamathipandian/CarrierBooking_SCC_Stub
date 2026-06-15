'use strict';

const xml2js = require('xml2js');

/**
 * Parse a DavisTurner carrier ASN XML (ROOT/Data/PurchaseOrder structure).
 *
 * Returns an array of ASN groups — one entry per unique ASNID found in the file
 * (a single file can in theory cover multiple POs / ASNIDs).
 *
 * Each entry:
 * {
 *   asnId:            string,       // <ASNID> from PO header
 *   poId:             string,       // <PurchaseOrder_ID>
 *   pofc:             string,       // <POFC> — used as TradePartner FS
 *   finalDestination: string,       // <FinalDestination> — used as TradePartner FD (FC address lookup)
 *   shipDate:         string,       // <ShipDate> DDMMYYYY -> YYYY-MM-DD normalised
 *   supplier:      string,          // <SupplierName>
 *   supplierCode:  string,          // <SupplierCode> — used as Supplier_ID
 *   shippingPoint: string,          // <ShippingPoint> — used as Loading Port
 *   shippingTerms: string,          // <ShippingTerms> — e.g. FOB
 *   lines: [{
 *     sku:                  string,  // <SKUItemID>
 *     ean:                  string,  // <PrimaryEAN>
 *     description:          string,  // <Description>
 *     size:                 string,  // <Size>
 *     colour:               string,  // <Colour>
 *     style:                string,  // <LegacyStyle> (OptionID)
 *     packFormat:           string,  // <PackingFormat> F=Flat, H=Hanging
 *     country:              string,  // <CountryOfOrigin>
 *     quantity:             number,  // <Quantity> — shipped qty
 *     expectedDeliveryDate: string,  // <ExpectedDeliveryDate> DDMMYYYY -> YYYY-MM-DD
 *   }]
 * }
 */
async function parse(xmlString) {
  const parser = new xml2js.Parser({ explicitArray: true, mergeAttrs: false });
  const result = await parser.parseStringPromise(xmlString);

  const dataEl = result?.ROOT?.Data?.[0];
  if (!dataEl) throw new Error('Invalid carrier ASN XML — missing ROOT/Data');

  const poElements = dataEl.PurchaseOrder || [];
  if (poElements.length === 0) throw new Error('No PurchaseOrder elements in carrier ASN XML');

  const asnMap = {}; // asnId -> entry

  for (const po of poElements) {
    const poId             = po.PurchaseOrder_ID?.[0] || '';
    const pofc             = po.POFC?.[0]             || '';
    const finalDestination = po.FinalDestination?.[0] || '';
    const supplier         = po.SupplierName?.[0]     || '';
    const supplierCode  = po.SupplierCode?.[0]  || '';
    const shippingPoint = po.ShippingPoint?.[0] || '';
    const shippingTerms = po.ShippingTerms?.[0] || '';
    const rawDate       = po.ShipDate?.[0] || '';

    // Normalise DDMMYYYY -> YYYY-MM-DD
    const shipDate = normDate(rawDate);

    for (const line of (po.LineCollection?.[0]?.Line || [])) {
      const asnId      = line.ASNID?.[0] || po.ASNID?.[0] || '';
      const sku        = line.SKUItemID?.[0] || '';
      const ean        = line.PrimaryEAN?.[0] || '';
      const description= line.Description?.[0] || '';
      const size       = line.Size?.[0] || '';
      const colour     = line.Colour?.[0] || '';
      const style      = line.LegacyStyle?.[0] || line.OptionID?.[0] || '';
      const packFormat           = line.PackingFormat?.[0] || 'F';
      const country              = line.CountryOfOrigin?.[0] || '';
      const quantity             = parseFloat(line.Quantity?.[0]) || 0;
      const expectedDeliveryDate = normDate(line.ExpectedDeliveryDate?.[0] || '');

      if (!sku) continue;

      if (!asnMap[asnId]) {
        asnMap[asnId] = { asnId, poId, pofc, finalDestination, shipDate, supplier, supplierCode, shippingPoint, shippingTerms, lines: [] };
      }
      asnMap[asnId].lines.push({ sku, ean, description, size, colour, style, packFormat, country, quantity, expectedDeliveryDate });
    }
  }

  return Object.values(asnMap);
}

/**
 * Normalise date string to YYYY-MM-DD.
 * Accepts DDMMYYYY (e.g. 26052026) or already YYYY-MM-DD.
 */
function normDate(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{8}$/.test(s)) {
    // Could be DDMMYYYY or YYYYMMDD — detect by year range
    const first4 = parseInt(s.slice(0, 4), 10);
    if (first4 >= 2000 && first4 <= 2100) {
      // YYYYMMDD
      return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
    }
    // DDMMYYYY
    return `${s.slice(4,8)}-${s.slice(2,4)}-${s.slice(0,2)}`;
  }
  return s;
}

module.exports = { parse };
