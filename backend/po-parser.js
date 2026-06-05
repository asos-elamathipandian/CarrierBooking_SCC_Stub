'use strict';

const xml2js = require('xml2js');

/**
 * Parse ASOS PO feed XML (BPM-850) into structured object.
 */
async function parse(xmlString) {
  const parser = new xml2js.Parser({ explicitArray: true, mergeAttrs: false });
  const result = await parser.parseStringPromise(xmlString);

  const bpMsg = result?.XMLBundle?.XMLTransmission?.[0]
    ?.XMLGroup?.[0]?.XMLTransaction?.[0]?.BpMessage?.[0];
  if (!bpMsg) throw new Error('Invalid PO feed XML structure');

  const order = bpMsg?.Order?.[0];
  if (!order) throw new Error('No Order element found in PO feed');

  const orderId = order?.OrderID?.[0] || order?.$?.Key;

  // Extract trade partners
  const partners = {};
  for (const tp of (bpMsg.TradePartner || [])) {
    const role = tp?.$?.RoleCd;
    if (role) partners[role] = tp;
  }

  // Also check order-level trade partners
  for (const tp of (order.TradePartner || [])) {
    const role = tp?.$?.RoleCd;
    if (role && !partners[role]) partners[role] = tp;
  }

  const getPartnerName = (p) => p?.TradePartnerName?.[0] || '';
  const getPartnerId   = (p) => p?.TradePartnerID?.[0]?._ || p?.TradePartnerID?.[0] || '';
  const getAddr        = (p) => {
    const a = p?.TradePartnerAddress?.[0];
    if (!a) return {};
    return {
      streets:  (a.Street || []),
      city:     a.City?.[0] || '',
      state:    a.StateProvinceCd?.[0] || '',
      postal:   a.PostalCd?.[0] || '',
      country:  a.CountryCd?.[0] || ''
    };
  };

  const faAddr = getAddr(partners['FA']);
  const fdAddr = getAddr(partners['FD']);

  // Parse reference map
  const refs = {};
  for (const ref of (order.Reference || [])) {
    refs[ref?.$?.RefTypeCd] = ref?._ || ref;
  }

  // Parse dates
  const dates = {};
  for (const d of (order.Date || [])) {
    dates[d?.$?.DateTypeCd] = d?._ || d;
  }

  // FOB / Incoterms
  const incoterms = order?.FOBInstructions?.[0]?.TransTermsCd?.[0]?._ ||
                    order?.FOBInstructions?.[0]?.TransTermsCd?.[0] || '';

  // Line items
  const lineItems = (order.LineItem || []).map(li => {
    const liRefs = {};
    for (const r of (li.Reference || [])) liRefs[r?.$?.RefTypeCd] = r?._ || r;

    const liAttrs = {};
    for (const a of (li.Attribute || [])) liAttrs[a?.$?.AttributeTypeCd] = a?._ || a;

    const liMeasures = {};
    for (const m of (li.Measure || [])) liMeasures[m?.$?.Qualifier] = m?._ || m;

    return {
      key:          li?.$?.Key || '',
      sku:          liAttrs['SK'] || '',
      productStyle: liRefs['PT'] || '',
      description:  liRefs['VP'] || '',
      poQty:        parseFloat(liMeasures['PO102']) || 0,
      mode:         li?.Mode?.[0] || '30'
    };
  });

  return {
    orderId,
    supplierName:     getPartnerName(partners['SU']),
    supplierId:       getPartnerId(partners['SU']),
    factoryName:      getPartnerName(partners['FA']),
    factoryId:        getPartnerId(partners['FA']),
    factoryStreet1:   faAddr.streets?.[0] || '',
    factoryStreet2:   faAddr.streets?.[1] || '',
    factoryStreet3:   faAddr.streets?.[2] || '',
    factoryCity:      faAddr.city,
    factoryPostal:    faAddr.postal,
    factoryCountry:   faAddr.country,
    fcName:           getPartnerName(partners['FD']),
    fcId:             getPartnerId(partners['FD']),
    fcStreet1:        fdAddr.streets?.[0] || '',
    fcStreet2:        fdAddr.streets?.[1] || '',
    fcStreet3:        fdAddr.streets?.[2] || '',
    fcCity:           fdAddr.city,
    fcState:          fdAddr.state,
    fcPostal:         fdAddr.postal,
    fcCountry:        fdAddr.country || 'GB',
    carrierId:        getPartnerId(partners['CA']),
    carrierName:      getPartnerName(partners['CA']),
    loadingPortId:    getPartnerId(partners['SL']),
    f1Id:             getPartnerId(partners['F1']),
    incoterms,
    poType:           refs['8X'] || '',
    lineItems
  };
}

module.exports = { parse };
