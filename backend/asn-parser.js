'use strict';

const xml2js = require('xml2js');

/**
 * Parse ASOS ASN feed XML (BPM-856) into structured object.
 */
async function parse(xmlString) {
  const parser = new xml2js.Parser({ explicitArray: true, mergeAttrs: false });
  const result = await parser.parseStringPromise(xmlString);

  const bpMsg = result?.XMLBundle?.XMLTransmission?.[0]
    ?.XMLGroup?.[0]?.XMLTransaction?.[0]?.BpMessage?.[0];
  if (!bpMsg) throw new Error('Invalid ASN feed XML structure');

  const doc = bpMsg?.Document?.[0];
  const documentId = doc?.DocumentID?.[0] || doc?.$?.Key || '';

  // Parse dates
  const receivedDate = bpMsg?.Date?.find(d => d?.$?.DateTypeCd === 'PDX_RCVD')?._ || '';

  // Trade partners at message level
  const partners = {};
  for (const tp of (bpMsg.TradePartner || [])) {
    const role = tp?.$?.RoleCd;
    if (role) partners[role] = tp;
  }
  const fcId = partners['FD']?.TradePartnerID?.[0]?._ || partners['FD']?.TradePartnerID?.[0] || '';

  // Extract line items from all orders in the document
  const lines = [];
  for (const order of (doc?.Order || [])) {
    const orderId = order?.OrderID?.[0] || order?.$?.Key || '';
    for (const li of (order?.LineItem || [])) {
      const liAttrs = {};
      for (const a of (li.Attribute || [])) liAttrs[a?.$?.AttributeTypeCd] = a?._ || a;

      const liMeasures = {};
      for (const m of (li.Measure || [])) liMeasures[m?.$?.Qualifier] = m?._ || m;

      const liRefs = {};
      for (const r of (li.Reference || [])) liRefs[r?.$?.RefTypeCd] = r?._ || r;

      lines.push({
        orderId,
        sku:          liAttrs['SK'] || '',
        receivedQty:  parseFloat(liMeasures['RCV']) || 0,
        shipmentRef:  liRefs['BAF'] || ''
      });
    }
  }

  return {
    documentId,
    receivedDate,
    fcId,
    lines
  };
}

module.exports = { parse };
