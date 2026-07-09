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
 * bam033j dates are stored 1 calendar day ahead of the intended business date
 * (dates ingested as midnight UTC of the following day). Subtract 1 day to correct.
 */
function bam033jDate(val) {
  if (!val) return '';
  const s = toDateStr(val);
  if (!s || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() - 1);
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
  console.log(`[Databricks ASN] fetchAsnsByPoRefs called with: ${JSON.stringify(poRefs)}`);
  const safePOs = poRefs.map(p => String(p).trim()).filter(p => /^\d+$/.test(p));
  console.log(`[Databricks ASN] safePOs after filter: ${JSON.stringify(safePOs)}`);
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
      s.bookingRequested,
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
        CAST(SupplierID AS STRING)                                         AS SupplierID,
        SupplierName,
        LadingPort,
        FreightTermsDescription                                            AS Incoterms,
        CASE WHEN size(PODtl) > 0 THEN PODtl[0].OriginCountryID ELSE NULL END        AS FirstCountry,
        CASE WHEN size(PODtl) > 0 THEN CAST(PODtl[0].OptionItemID AS STRING) ELSE NULL END AS OptionItemID,
        CAST(Factory AS STRING)                                            AS FactoryID,
        FactoryDesc                                                        AS FactoryName,
        Status                                                             AS POStatus,
        ExFactoryDate,
        ExpectedShipmentDate,
        ExpectedHandoverDate,
        HandoverWindowStartDate,
        HandoverWindowEndDate,
        ExpectedDeliveryDateFirstLocation,
        ROW_NUMBER() OVER (PARTITION BY OrderNo ORDER BY _IngestedDate DESC) AS _rn
      FROM sourcingandbuying.conformed.bam033j_purchase_order_v1
      WHERE OrderNo IN (${poList})
    `;
    console.log(`[Databricks PO] querying bam033j for POs: ${poList}`);
    const poRows = await db.query(poSql);
    console.log(`[Databricks PO] bam033j returned ${(poRows || []).length} row(s)`);
    if (poRows && poRows.length > 0) {
      console.log(`[Databricks PO] sample OrderNo values: ${[...new Set(poRows.slice(0,5).map(r => r.OrderNo))].join(', ')}`);
      console.log(`[Databricks PO] sample SupplierName: ${poRows[0].SupplierName}, LadingPort: ${poRows[0].LadingPort}, FactoryName: ${poRows[0].FactoryName}`);
      console.log(`[Databricks PO] dates — ExFactoryDate: ${poRows[0].ExFactoryDate}, ExpectedShipmentDate: ${poRows[0].ExpectedShipmentDate}, ExpectedHandoverDate: ${poRows[0].ExpectedHandoverDate}, HandoverWindowStartDate: ${poRows[0].HandoverWindowStartDate}, ExpectedDeliveryDateFirstLocation: ${poRows[0].ExpectedDeliveryDateFirstLocation}`);
    }
    for (const r of (poRows || [])) {
      if (Number(r._rn) === 1 || !poEnrichMap[r.OrderNo]) {
        poEnrichMap[r.OrderNo] = {
          poStatus:             r.POStatus            || '',
          supplierID:           r.SupplierID          || '',
          supplierName:         r.SupplierName        || '',
          ladingPort:           r.LadingPort          || '',
          incoterms:            r.Incoterms           || '',
          country:              r.FirstCountry        || '',
          optionID:             r.OptionItemID        || '',
          factoryID:            r.FactoryID           || '',
          factoryName:          r.FactoryName         || '',
          expectedShipmentDate: bam033jDate(r.ExFactoryDate || r.ExpectedShipmentDate || ''),
          expectedDeliveryDate: bam033jDate(r.ExpectedDeliveryDateFirstLocation || r.ExpectedHandoverDate || ''),
          exFactoryDate:        bam033jDate(r.ExFactoryDate        || ''),
          expShipmentDate:      bam033jDate(r.ExpectedShipmentDate || ''),
          expHandoverDate:      bam033jDate(r.ExpectedHandoverDate || ''),
          expDeliveryDate:      bam033jDate(r.ExpectedDeliveryDateFirstLocation || ''),
          supplierStreet1: '',
          supplierCity:    '',
          supplierPostal:  '',
          supplierCountry: '',
          factoryStreet1:  '',
          factoryCity:     '',
          factoryPostal:   '',
          factoryCountry:  ''
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
        shipDate:         toDateStr(row.asnEstimatedShipmentDate || row.latestPlannedShipmentDate || enrich.expectedShipmentDate),
        supplier:         enrich.supplierName   || '',
        supplierCode:     enrich.supplierID     || row.supplierCode || '',
        shippingPoint:    enrich.ladingPort     || row.portOfLoad   || '',
        shippingTerms:    enrich.incoterms      || '',
        supplierStreet1:  '',
        supplierCity:     '',
        supplierPostal:   '',
        supplierCountry:  '',
        factoryID:        enrich.factoryID      || '',
        factoryName:      enrich.factoryName    || '',
        factoryStreet1:   '',
        factoryCity:      '',
        factoryPostal:    '',
        factoryCountry:   '',
        mode:             row.mode              || '',
        carrier:          row.carrier           || '',
        bookingRequested: row.bookingRequested  || null,
        expectedDeliveryDate: toDateStr(row.asnDeliveryDateFinalDest || enrich.expectedDeliveryDate),
        lines: []
      };
    }
    asnPoMap[key].lines.push({
      sku:         row.sku                  || '',
      ean:         '',
      description: '',
      size:        '',
      colour:      '',
      style:       enrich.optionID || '',
      packFormat:  row.asnLoadingType === 'H' ? 'H' : 'F',
      country:     row.countryOfManufacture || enrich.country || '',
      quantity:    row.bookedQty            || 0,
      expectedDeliveryDate: toDateStr(row.asnDeliveryDateFinalDest || enrich.expectedDeliveryDate)
    });
  }

  // ── Query 3: ASN cancellations from bam036e_asn_v1 ─────────────────────
  // _notification_type values: 'C' = Created (initial), 'M' = Modified, 'D' = Deleted/Cancelled.
  // Only 'D' (delete) means the ASN has been cancelled — 'C' is the creation event.
  // Non-fatal — if this query fails we proceed without the cancellation check.
  const cancelledAsnIds = new Set();
  try {
    const uniqueAsnIds = [...new Set(Object.values(asnPoMap).map(g => g.asnId).filter(Boolean))];
    if (uniqueAsnIds.length > 0) {
      const asnList = uniqueAsnIds.map(a => `'${a}'`).join(', ');
      const asnCancelSql = `
        WITH latest AS (
          SELECT _asn_nbr, _notification_type,
            ROW_NUMBER() OVER (PARTITION BY _asn_nbr ORDER BY _last_update_timestamp DESC, _IngestedDate DESC) AS _rn
          FROM supplychain.conformed.bam036e_asn_v1
          WHERE _asn_nbr IN (${asnList})
        )
        SELECT _asn_nbr
        FROM latest
        WHERE _rn = 1 AND _notification_type = 'D'
      `;
      console.log(`[Databricks bam036e] checking cancellation status for ${uniqueAsnIds.length} ASN(s)`);
      const cancelRows = await db.query(asnCancelSql);
      for (const r of (cancelRows || [])) {
        if (r._asn_nbr) cancelledAsnIds.add(String(r._asn_nbr));
      }
      console.log(`[Databricks bam036e] ${cancelledAsnIds.size} cancelled ASN(s) found`);
    }
  } catch (err) {
    console.warn(`[Databricks bam036e] cancellation check failed (continuing without it): ${err.message}`);
  }

  // ── Separate active vs cancelled ASN/PO groups ───────────────────────────
  const cancelledItems = [];
  const allGroups      = Object.values(asnPoMap);

  const parsed = allGroups.filter(group => {
    const enrich = poEnrichMap[group.poId] || {};

    if (cancelledAsnIds.has(String(group.asnId))) {
      console.warn(`[Databricks ASN] SKIPPED — ASN ${group.asnId} / PO ${group.poId} is cancelled (bam036e _notification_type=D)`);
      cancelledItems.push({
        type:   'ASN',
        asnId:  group.asnId,
        poId:   group.poId,
        reason: `ASN ${group.asnId} (PO ${group.poId}) is cancelled — latest notification type is D (deleted)`
      });
      return false;
    }

    if (group.bookingRequested) {
      const when = toDateStr(group.bookingRequested);
      console.warn(`[Databricks ASN] SKIPPED — ASN ${group.asnId} / PO ${group.poId} already has a booking request (bookingRequested=${when})`);
      cancelledItems.push({
        type:   'ALREADY_BOOKED',
        asnId:  group.asnId,
        poId:   group.poId,
        reason: `ASN ${group.asnId} (PO ${group.poId}) already has a carrier booking request — submitted on ${when}`
      });
      return false;
    }

    if (enrich.poStatus === 'C') {
      console.warn(`[Databricks ASN] SKIPPED — PO ${group.poId} has Status=C (cancelled)`);
      cancelledItems.push({
        type:  'PO',
        asnId: group.asnId,
        poId:  group.poId,
        reason: `PO ${group.poId} is cancelled (Status=C) — booking creation skipped`
      });
      return false;
    }

    return true;
  });

  const foundPOs = new Set(allGroups.map(g => g.poId));
  const errors   = safePOs
    .filter(p => !foundPOs.has(p))
    .map(p => `Databricks: no shipment record found for PO ${p}`);

  console.log(`[Databricks ASN] ${parsed.length} active ASN group(s), ${cancelledItems.length} cancelled, for ${safePOs.length} PO ref(s)`);
  if (parsed.length > 0) {
    const p = parsed[0];
    console.log(`[Databricks ASN] shipDate="${p.shipDate}", expectedDeliveryDate="${p.expectedDeliveryDate}", mode="${p.mode}"`);
    console.log(`[Databricks ASN] raw asnEstimatedShipmentDate="${rows[0]?.asnEstimatedShipmentDate}", latestPlannedShipmentDate="${rows[0]?.latestPlannedShipmentDate}", asnDeliveryDateFinalDest="${rows[0]?.asnDeliveryDateFinalDest}"`);
  }

  return {
    poFeeds: [], asnFeeds: [],
    carrierAsnFiles: [{
      filename:     'DATABRICKS_shipments.json',
      xml:          null,
      poRef:        safePOs[0],
      lastModified: new Date(),
      parsed
    }],
    cancelledItems,
    errors
  };
}

module.exports = { fetchAsnsByPoRefs };

