'use strict';
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const db = require('./databricks-client');

db.query(`
  WITH latest_facts AS (
    SELECT f.dim_purchase_order_sk AS poId, f.dim_advanced_shipment_notice_sk AS asnId,
           DATE_FORMAT(f.dim_expected_factory_date_sk, 'yyyy-MM-dd') AS exFactory
    FROM sourcingandbuying.serve.fact_purchase_order_commitment_v1 f
    WHERE f.dim_purchase_order_sk = '500037820158'
    QUALIFY ROW_NUMBER() OVER (PARTITION BY f.dim_purchase_order_sk, f.dim_advanced_shipment_notice_sk, f.dim_product_sk ORDER BY f.dim_date_sk DESC) = 1
  )
  SELECT lf.poId, lf.asnId, lf.exFactory,
    DATE_FORMAT(po.dim_original_purchase_order_shipment_date_sk,           'yyyy-MM-dd') AS poShipDate,
    DATE_FORMAT(po.dim_current_requested_intake_first_destination_date_sk, 'yyyy-MM-dd') AS poDeliveryDate
  FROM latest_facts lf
  LEFT JOIN sourcingandbuying.serve.dim_purchase_order_v1 po ON lf.poId = po.dim_purchase_order_sk
  LIMIT 1
`).then(rows => {
  if (!rows || !rows[0]) { console.log('no rows'); return; }
  Object.entries(rows[0]).forEach(([k, v]) => console.log(k.padEnd(20), '=', v));
}).catch(e => console.error(e.message));
