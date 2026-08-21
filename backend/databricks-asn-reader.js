'use strict';

const fs   = require('fs');
const path = require('path');
const db   = require('./databricks-client');

const GEN_LOG_PATH = path.join(__dirname, '..', 'bible', 'generation-log.json');

function getToolCancelledAsnIds() {
  try {
    if (!fs.existsSync(GEN_LOG_PATH)) return new Set();
    const log     = JSON.parse(fs.readFileSync(GEN_LOG_PATH, 'utf8'));
    const entries = Array.isArray(log) ? log : (log.entries || []);
    const latestByAsn = {};
    for (const e of entries) {
      for (const asn of (e.asnRefs || [])) {
        const key = String(asn);
        if (!latestByAsn[key] || e.timestamp > latestByAsn[key].timestamp)
          latestByAsn[key] = e;
      }
    }
    return new Set(
      Object.entries(latestByAsn)
        .filter(([, e]) => e.purposeCd === '01')
        .map(([asn]) => asn)
    );
  } catch (_) { return new Set(); }
}

/** Normalise a date value to YYYY-MM-DD; returns '' for nulls and 1900-01-01 sentinels. */
function toDateStr(val) {
  if (!val) return '';
  const s = String(val).trim();
  if (s === '1900-01-01') return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s === '1900-01-01' ? '' : s;
  const d = new Date(s);
  if (isNaN(d.getTime())) return s.slice(0, 10);
  const iso = d.toISOString().slice(0, 10);
  return iso === '1900-01-01' ? '' : iso;
}

/**
 * Fetch carrier ASN + PO enrichment data from Databricks Serve layer for the given PO refs.
 *
 * Single star-schema query across:
 *   fact_purchase_order_commitment_v1  (grain: PO + ASN + SKU, snapshot table)
 *   dim_advanced_shipment_notice_v1    (ASN status, carrier, lading port)
 *   dim_purchase_order_v1              (incoterms)
 *   dim_supplier_v1                    (supplier ID/name/country)
 *   dim_factory_v1                     (factory code/name/country)
 *
 * Returns { carrierAsnFiles, errors } matching the structure expected by the pipeline.
 */
async function fetchAsnsByPoRefs(poRefs) {
  if (!poRefs || poRefs.length === 0) {
    return { poFeeds: [], asnFeeds: [], carrierAsnFiles: [], errors: [] };
  }

  console.log(`[Databricks ASN] fetchAsnsByPoRefs called with: ${JSON.stringify(poRefs)}`);
  const safePOs = poRefs.map(p => String(p).trim()).filter(p => /^\d+$/.test(p));
  console.log(`[Databricks ASN] safePOs after filter: ${JSON.stringify(safePOs)}`);
  if (safePOs.length === 0) {
    return { poFeeds: [], asnFeeds: [], carrierAsnFiles: [], errors: ['No valid numeric PO references provided'] };
  }

  const poList = safePOs.map(p => `'${p}'`).join(', ');

  // Single star-schema query — QUALIFY picks the latest daily snapshot per PO+ASN+SKU
  const sql = `
    WITH latest_facts AS (
      SELECT
        f.dim_purchase_order_sk                                                    AS poId,
        f.dim_advanced_shipment_notice_sk                                          AS asnId,
        f.dim_product_sk                                                           AS sku,
        CAST(f.dim_first_warehouse_sk AS STRING)                                   AS firstDestination,
        CAST(f.dim_final_warehouse_sk  AS STRING)                                  AS finalDestination,
        f.dim_purchase_order_status_sk                                             AS poStatus,
        f.is_booked_by_carrier,
        f.quantity                                                                 AS bookedQty,
        CAST(f.dim_expected_factory_date_sk                                AS STRING) AS exFactoryDate,
        CAST(f.dim_expected_shipment_date_sk                               AS STRING) AS expectedShipmentDate,
        CAST(f.dim_first_warehouse_current_expected_delivery_date_sk       AS STRING) AS expectedDeliveryDate,
        f.dim_factory_sk,
        f.dim_supplier_sk
      FROM sourcingandbuying.serve.fact_purchase_order_commitment_v1 f
      WHERE f.dim_purchase_order_sk IN (${poList})
        AND f.dim_advanced_shipment_notice_sk != 'Unknown'
      QUALIFY ROW_NUMBER() OVER (
        PARTITION BY f.dim_purchase_order_sk, f.dim_advanced_shipment_notice_sk, f.dim_product_sk
        ORDER BY f.dim_date_sk DESC
      ) = 1
    )
    SELECT
      lf.poId,
      lf.asnId,
      lf.sku,
      lf.firstDestination,
      lf.finalDestination,
      lf.poStatus,
      lf.is_booked_by_carrier    AS isBookedByCarrier,
      lf.bookedQty,
      lf.exFactoryDate,
      lf.expectedShipmentDate,
      lf.expectedDeliveryDate,
      asn.asn_id,
      asn.asn_status_code,
      asn.carrier_code,
      asn.lading_port_code,
      sup.supplier_id,
      sup.supplier               AS supplierName,
      sup.primary_country_code   AS supplierCountry,
      fac.factory_code           AS factoryID,
      fac.factory                AS factoryName,
      fac.factory_country_code,
      po.inco_terms,
      DATE_FORMAT(po.dim_original_purchase_order_shipment_date_sk,           'yyyy-MM-dd') AS poShipDate,
      DATE_FORMAT(po.dim_current_requested_intake_first_destination_date_sk, 'yyyy-MM-dd') AS poDeliveryDate
    FROM latest_facts lf
    LEFT JOIN supplychain.serve.dim_advanced_shipment_notice_v1 asn
           ON lf.asnId = asn.dim_advanced_shipment_notice_sk
    LEFT JOIN sourcingandbuying.serve.dim_supplier_v1 sup
           ON lf.dim_supplier_sk = sup.dim_supplier_sk
    LEFT JOIN sourcingandbuying.serve.dim_factory_v1 fac
           ON lf.dim_factory_sk = fac.dim_factory_sk
    LEFT JOIN sourcingandbuying.serve.dim_purchase_order_v1 po
           ON lf.poId = po.dim_purchase_order_sk
    WHERE asn.asn_id IS NOT NULL
    ORDER BY lf.asnId, lf.poId, lf.sku
  `;

  let rows;
  try {
    console.log(`[Databricks Serve] querying ${safePOs.length} PO(s) via fact_purchase_order_commitment_v1`);
    rows = await db.query(sql);
    console.log(`[Databricks Serve] ${(rows || []).length} row(s) returned`);
  } catch (err) {
    return {
      poFeeds: [], asnFeeds: [], carrierAsnFiles: [],
      errors: [`Databricks serve query failed: ${err.message}`]
    };
  }

  if (!rows || rows.length === 0) {
    return {
      poFeeds: [], asnFeeds: [], carrierAsnFiles: [],
      errors: safePOs.map(p => `Databricks: no shipment record found for PO ${p}`)
    };
  }

  if (rows.length > 0) {
    const s = rows[0];
    console.log(`[Databricks Serve] sample: ASN=${s.asn_id}, PO=${s.poId}, sku=${s.sku}, exFactory=${s.exFactoryDate}, shipDate=${s.expectedShipmentDate}, delivery=${s.expectedDeliveryDate}, lading=${s.lading_port_code}, supplier=${s.supplierName}, factory=${s.factoryName}, status=${s.asn_status_code}, booked=${s.isBookedByCarrier}`);
  }

  // Cancellation check pass — asn_status_code 'D' = deleted/cancelled (same as bam036e _notification_type)
  const cancelledAsnIds = new Set();
  for (const row of rows) {
    if (row.asn_status_code === 'D') cancelledAsnIds.add(String(row.asn_id || row.asnId));
  }

  // Group rows by ASN+PO, collecting per-SKU lines
  const asnPoMap = {};
  for (const row of rows) {
    const key = `${row.asnId}::${row.poId}`;
    if (!asnPoMap[key]) {
      asnPoMap[key] = {
        asnId:            String(row.asn_id   || row.asnId  || ''),
        poId:             String(row.poId                   || ''),
        pofc:             row.firstDestination              || '',
        finalDestination: row.finalDestination              || '',
        poStatus:         row.poStatus                      || '',
        isBookedByCarrier: (row.isBookedByCarrier || 'No'),
        // bookingRequested mirrors isBookedByCarrier for downstream compatibility
        bookingRequested: row.isBookedByCarrier === 'Yes' ? 'booked' : null,
        bookingCanceled:  null,
        shipDate:         row.poShipDate       || toDateStr(row.expectedShipmentDate || row.exFactoryDate),
        supplier:         row.supplierName                  || '',
        supplierCode:     row.supplier_id                   || '',
        shippingPoint:    row.lading_port_code              || '',
        shippingTerms:    row.inco_terms                    || '',
        supplierStreet1:  '',
        supplierCity:     '',
        supplierPostal:   '',
        supplierCountry:  '',
        factoryID:        row.factoryID                     || '',
        factoryName:      row.factoryName                   || '',
        factoryStreet1:   '',
        factoryCity:      '',
        factoryPostal:    '',
        factoryCountry:   '',
        mode:             row.carrier_code                  || '',
        carrier:          row.carrier_code                  || '',
        expectedDeliveryDate: row.poDeliveryDate || toDateStr(row.expectedDeliveryDate),
        lines: []
      };
    }
    asnPoMap[key].lines.push({
      sku:         String(row.sku        || ''),
      ean:         '',
      description: '',
      size:        '',
      colour:      '',
      style:       '',
      packFormat:  'F',
      country:     row.factory_country_code || row.supplierCountry || '',
      quantity:    row.bookedQty ?? 0,
      expectedDeliveryDate: row.poDeliveryDate || toDateStr(row.expectedDeliveryDate)
    });
  }

  // Separate active vs cancelled ASN/PO groups
  const toolCancelledAsnIds = getToolCancelledAsnIds();
  const cancelledItems = [];
  const allGroups      = Object.values(asnPoMap);

  const parsed = allGroups.filter(group => {
    if (cancelledAsnIds.has(String(group.asnId))) {
      console.warn(`[Databricks Serve] SKIPPED — ASN ${group.asnId} / PO ${group.poId} is cancelled (asn_status_code=D)`);
      cancelledItems.push({
        type:   'ASN',
        asnId:  group.asnId,
        poId:   group.poId,
        reason: `ASN ${group.asnId} (PO ${group.poId}) is cancelled`
      });
      return false;
    }

    if (group.isBookedByCarrier === 'Yes') {
      const toolCancelled = toolCancelledAsnIds.has(String(group.asnId));
      if (!toolCancelled) {
        console.warn(`[Databricks Serve] SKIPPED — ASN ${group.asnId} / PO ${group.poId} already booked by carrier`);
        cancelledItems.push({
          type:   'ALREADY_BOOKED',
          asnId:  group.asnId,
          poId:   group.poId,
          reason: `ASN ${group.asnId} (PO ${group.poId}) already has a carrier booking`
        });
        return false;
      }
      console.log(`[Databricks Serve] ASN ${group.asnId} booked but tool-cancelled — allowing re-book`);
    }

    if (group.poStatus === 'C') {
      console.warn(`[Databricks Serve] SKIPPED — PO ${group.poId} has Status=C (cancelled)`);
      cancelledItems.push({
        type:  'PO',
        asnId: group.asnId,
        poId:  group.poId,
        reason: `PO ${group.poId} is cancelled (Status=C)`
      });
      return false;
    }

    return true;
  });

  const foundPOs = new Set(allGroups.map(g => g.poId));
  const errors   = safePOs
    .filter(p => !foundPOs.has(p))
    .map(p => `Databricks: no shipment record found for PO ${p}`);

  console.log(`[Databricks Serve] ${parsed.length} active group(s), ${cancelledItems.length} cancelled, for ${safePOs.length} PO ref(s)`);
  if (parsed.length > 0) {
    const p = parsed[0];
    console.log(`[Databricks Serve] shipDate="${p.shipDate}", expectedDeliveryDate="${p.expectedDeliveryDate}", mode="${p.mode}", shippingPoint="${p.shippingPoint}"`);
  }

  // One carrierAsnFile entry per PO that has active parsed data — so the
  // frontend correctly marks each PO as "found" or "not found" individually.
  const parsedByPO = {};
  for (const group of parsed) {
    if (!parsedByPO[group.poId]) parsedByPO[group.poId] = [];
    parsedByPO[group.poId].push(group);
  }
  const carrierAsnFiles = Object.entries(parsedByPO).map(([poId, poGroups]) => ({
    filename:     'DATABRICKS_shipments.json',
    xml:          null,
    poRef:        poId,
    lastModified: new Date(),
    parsed:       poGroups
  }));

  return {
    poFeeds: [], asnFeeds: [],
    carrierAsnFiles,
    cancelledItems,
    errors
  };
}

/**
 * Resolve PO references from ASN references using the serve layer fact table.
 * Returns { asnToPoMap, poRefs, errors } where asnToPoMap[asn] = [po1, po2, ...].
 */
async function resolvePoRefsByAsnRefs(asnRefs) {
  const cleanAsns = [...new Set((asnRefs || []).map(a => String(a || '').trim()).filter(Boolean))];
  if (!cleanAsns.length) return { asnToPoMap: {}, poRefs: [], errors: [] };

  const safeAsns = cleanAsns.filter(a => /^[A-Za-z0-9_./-]+$/.test(a));
  const rejected = cleanAsns.filter(a => !safeAsns.includes(a));

  if (!safeAsns.length) {
    return {
      asnToPoMap: {},
      poRefs: [],
      errors: rejected.map(a => `ASN ${a}: invalid format`)
    };
  }

  const asnList = safeAsns.map(a => `'${a}'`).join(', ');
  const sql = `
    SELECT DISTINCT
      f.dim_advanced_shipment_notice_sk AS asnId,
      f.dim_purchase_order_sk           AS poId
    FROM sourcingandbuying.serve.fact_purchase_order_commitment_v1 f
    WHERE f.dim_advanced_shipment_notice_sk IN (${asnList})
      AND f.dim_advanced_shipment_notice_sk != 'Unknown'
      AND f.dim_purchase_order_sk IS NOT NULL
  `;

  let rows;
  try {
    rows = await db.query(sql);
  } catch (err) {
    return {
      asnToPoMap: {},
      poRefs: [],
      errors: [`Databricks ASN->PO resolution failed: ${err.message}`]
    };
  }

  const asnToPoSet = {};
  for (const asn of safeAsns) asnToPoSet[asn] = new Set();
  for (const r of (rows || [])) {
    const asn = String(r.asnId || '').trim();
    const po  = String(r.poId  || '').trim();
    if (!asn || !po) continue;
    if (!asnToPoSet[asn]) asnToPoSet[asn] = new Set();
    asnToPoSet[asn].add(po);
  }

  const asnToPoMap = {};
  const poSet = new Set();
  const errors = [];

  for (const asn of cleanAsns) {
    const poList = [...(asnToPoSet[asn] || new Set())];
    if (!poList.length) {
      errors.push(`ASN ${asn}: no PO found in Databricks`);
      continue;
    }
    asnToPoMap[asn] = poList;
    poList.forEach(po => poSet.add(po));
  }

  for (const asn of rejected) {
    errors.push(`ASN ${asn}: invalid format`);
  }

  return { asnToPoMap, poRefs: [...poSet], errors };
}

module.exports = { fetchAsnsByPoRefs, resolvePoRefsByAsnRefs };

