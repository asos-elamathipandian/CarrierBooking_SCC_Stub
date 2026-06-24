'use strict';

/**
 * Databricks replacement for the blob-based carrier ASN feed.
 *
 * Queries:
 *   supplychain.conformed.aim_shipment_detail_v1        — shipment / ASN data
 *   sourcingandbuying.conformed.bam033j_purchase_order_v1 — supplier name, factory,
 *                                                           freight terms, EAN, size,
 *                                                           colour, description per SKU
 *
 * Enable by setting  ASN_SOURCE=databricks  in .env.
 */

const db = require('./databricks-client');

/** Normalise an ISO timestamp / date string to YYYY-MM-DD. */
function toDateStr(val) {
  if (!val) return '';
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (isNaN(d.getTime())) return s.slice(0, 10);
  return d.toISOString().slice(0, 10);
}

/**
 * Fetch carrier ASN + PO enrichment data from Databricks for the given PO refs.
 *
 * Returns { carrierAsnFiles, errors } matching the structure of
 * blobClient.fetchCarrierFeedsOnly() so the rest of the pipeline is unchanged.
 */
async function fetchAsnsByPoRefs(poRefs) {
  if (!poRefs || poRefs.length === 0) {
    return { poFeeds: [], asnFeeds: [], carrierAsnFiles: [], errors: [] };
  }

  // Only numeric PO refs — prevents SQL injection from template data
  const safePOs = poRefs.map(p => String(p).trim()).filter(p => /^\d+$/.test(p));
  if (safePOs.length === 0) {
    return { poFeeds: [], asnFeeds: [], carrierAsnFiles: [], errors: ['No valid numeric PO references provided'] };
  }

  const poList = safePOs.map(p => `'${p}'`).join(', ');

  // ── Query 1: shipment lines from aim_shipment_detail_v1 ───────────────────
  const shipmentSql = `
    WITH latest AS (
      SELECT *,
        ROW_NUMBER() OVER (PARTITION BY asnId ORDER BY _IngestedDate DESC) AS _rn
      FROM supplychain.conformed.aim_shipment_detail_v1
      WHERE size(filter(orderLineItems, ol -> array_contains(array(${poList}), ol.orderId))) > 0
    )
    SELECT
      s.asnId,
      s.mode,
      s.countryOfManufacture,
      s.portOfLoad,
      s.finalDestination,
      s.firstDestination,
      s.carrier,
      s.supplier                         AS supplierCode,
      s.asnEstimatedShipmentDate,
      s.latestPlannedShipmentDate,
      s.asnDeliveryDateFinalDest,
      s.asnLoadingType,
      ol.orderId                         AS poId,
      ol.sku,
      TRY_CAST(ol.bookedQty AS DOUBLE)   AS bookedQty
    FROM latest s
    LATERAL VIEW EXPLODE(orderLineItems) AS ol
    WHERE s._rn = 1
      AND ol.orderId IN (${poList})
    ORDER BY s.asnId, ol.orderId, ol.sku
  `;

  let rows;
  try {
    rows = await db.query(shipmentSql);
  } catch (err) {
    return {
      poFeeds: [], asnFeeds: [], carrierAsnFiles: [],
      errors: [`Databricks shipment query failed: ${err.message}`]
    };
  }

  if (!rows || rows.length === 0) {
    return {
      poFeeds: [], asnFeeds: [], carrierAsnFiles: [],
      errors: safePOs.map(p => `Databricks: no shipment record found for PO ${p}`)
    };
  }

  // ── Query 2: PO enrichment from bam033j_purchase_order_v1 ────────────────
  // Separate query — if it fails we still have shipment data above
  const poEnrichMap = {}; // OrderNo -> { SupplierID, SupplierName, LadingPort, Incoterms }
  try {
    const poSql = `
      SELECT
        OrderNo,
        CAST(SupplierID AS STRING)        AS SupplierID,
        SupplierName,
        LadingPort,
        FreightTermsDescription           AS Incoterms,
        PODtl[0].OriginCountryID          AS FirstCountry,
        SupplierAddress1                  AS SupplierStreet1,
        SupplierCity,
        SupplierPostCode,
        SupplierCountryCode,
        FactoryID,
        FactoryName,
        FactoryAddress1                   AS FactoryStreet1,
        FactoryCity,
        FactoryPostCode,
        FactoryCountryCode,
        ROW_NUMBER() OVER (PARTITION BY OrderNo ORDER BY _IngestedDate DESC) AS _rn
      FROM sourcingandbuying.conformed.bam033j_purchase_order_v1
      WHERE OrderNo IN (${poList})
    `;
    const poRows = await db.query(poSql);
    for (const r of (poRows || [])) {
      if (r._rn === 1 || !poEnrichMap[r.OrderNo]) {
        poEnrichMap[r.OrderNo] = {
          supplierID:      r.SupplierID          || '',
          supplierName:    r.SupplierName        || '',
          ladingPort:      r.LadingPort          || '',
          incoterms:       r.Incoterms           || '',
          country:         r.FirstCountry        || '',
          supplierStreet1: r.SupplierStreet1     || '',
          supplierCity:    r.SupplierCity        || '',
          supplierPostal:  r.SupplierPostCode    || '',
          supplierCountry: r.SupplierCountryCode || '',
          factoryID:       r.FactoryID           ? String(r.FactoryID) : '',
          factoryName:     r.FactoryName         || '',
          factoryStreet1:  r.FactoryStreet1      || '',
          factoryCity:     r.FactoryCity         || '',
          factoryPostal:   r.FactoryPostCode     || '',
          factoryCountry:  r.FactoryCountryCode  || ''
        };
      }
    }
    console.log(`[Databricks PO] enrichment fetched for ${Object.keys(poEnrichMap).length} PO(s)`);
  } catch (err) {
    console.warn(`[Databricks PO] enrichment query failed (continuing without it): ${err.message}`);
  }

  // ── Group shipment rows by (asnId, poId) ──────────────────────────────────
  const asnPoMap = {};
  for (const row of rows) {
    const key = `${row.asnId}::${row.poId}`;
    const enrich = poEnrichMap[row.poId] || {};
    if (!asnPoMap[key]) {
      asnPoMap[key] = {
        asnId:            row.asnId             || '',
        poId:             row.poId              || '',
        pofc:             row.firstDestination  || row.finalDestination || '',
        finalDestination: row.finalDestination  || '',
        shipDate:         toDateStr(row.asnEstimatedShipmentDate || row.latestPlannedShipmentDate),
        supplier:         enrich.supplierName   || '',
        supplierCode:     enrich.supplierID     || row.supplierCode || '',
        shippingPoint:    enrich.ladingPort     || row.portOfLoad   || '',
        shippingTerms:    enrich.incoterms      || '',
        supplierStreet1:  enrich.supplierStreet1 || '',
        supplierCity:     enrich.supplierCity    || '',
        supplierPostal:   enrich.supplierPostal  || '',
        supplierCountry:  enrich.supplierCountry || '',
        factoryID:        enrich.factoryID       || '',
        factoryName:      enrich.factoryName     || '',
        factoryStreet1:   enrich.factoryStreet1  || '',
        factoryCity:      enrich.factoryCity     || '',
        factoryPostal:    enrich.factoryPostal   || '',
        factoryCountry:   enrich.factoryCountry  || '',
        mode:             row.mode              || '',
        carrier:          row.carrier           || '',
        expectedDeliveryDate: toDateStr(row.asnDeliveryDateFinalDest),
        lines: []
      };
    }
    asnPoMap[key].lines.push({
      sku:         row.sku                  || '',
      ean:         '',
      description: '',
      size:        '',
      colour:      '',
      style:       '',
      packFormat:  row.asnLoadingType === 'H' ? 'H' : 'F',
      country:     row.countryOfManufacture || enrich.country || '',
      quantity:    row.bookedQty            || 0,
      expectedDeliveryDate: toDateStr(row.asnDeliveryDateFinalDest)
    });
  }

  const foundPOs = new Set(Object.values(asnPoMap).map(g => g.poId));
  const errors   = safePOs
    .filter(p => !foundPOs.has(p))
    .map(p => `Databricks: no shipment record found for PO ${p}`);

  const parsed = Object.values(asnPoMap);
  console.log(`[Databricks ASN] ${parsed.length} ASN group(s) fetched for ${safePOs.length} PO ref(s)`);

  return {
    poFeeds: [], asnFeeds: [],
    carrierAsnFiles: [{
      filename:     'DATABRICKS_shipments.json',
      xml:          null,
      poRef:        safePOs[0],
      lastModified: new Date(),
      parsed
    }],
    errors
  };
}

module.exports = { fetchAsnsByPoRefs };

