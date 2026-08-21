'use strict';
// Explores schema + sample data for Serve layer tables before migrating queries.
// Run: node backend/explore-serve-tables.js [po_number]
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const db = require('./databricks-client');

const SERVE_TABLES = [
  'supplychain.serve.dim_advanced_shipment_notice_v1',
  'sourcingandbuying.serve.dim_purchase_order_v1',
  'sourcingandbuying.serve.dim_supplier_v1',
  'sourcingandbuying.serve.dim_factory_v1',
  'sourcingandbuying.serve.fact_purchase_order_commitment_v1',
];

async function describeTable(table) {
  console.log(`\n${'─'.repeat(70)}`);
  console.log(`TABLE: ${table}`);
  console.log('─'.repeat(70));
  try {
    const rows = await db.query(`DESCRIBE TABLE ${table}`);
    for (const r of rows) {
      const col     = r.col_name || r.column_name || Object.values(r)[0] || '';
      const dtype   = r.data_type || r.type || Object.values(r)[1] || '';
      const comment = r.comment   || '';
      if (col && !col.startsWith('#')) {
        console.log(`  ${col.padEnd(55)} ${dtype.padEnd(20)} ${comment}`);
      }
    }
  } catch (err) {
    console.log(`  ERROR: ${err.message}`);
  }
}

async function sampleFact(poRef) {
  console.log(`\n${'─'.repeat(70)}`);
  console.log(`SAMPLE: fact_purchase_order_commitment_v1 WHERE purchase_order_id = '${poRef}'`);
  console.log('─'.repeat(70));
  try {
    const rows = await db.query(`
      SELECT f.*, po.purchase_order_id, asn.asn_id, asn.asn_status_code
      FROM sourcingandbuying.serve.fact_purchase_order_commitment_v1 f
      JOIN sourcingandbuying.serve.dim_purchase_order_v1  po  ON f.dim_purchase_order_sk  = po.dim_purchase_order_sk
      JOIN supplychain.serve.dim_advanced_shipment_notice_v1 asn ON f.dim_advanced_shipment_notice_sk = asn.dim_advanced_shipment_notice_sk
      WHERE po.purchase_order_id = '${poRef}'
      LIMIT 3
    `);
    if (!rows || rows.length === 0) {
      console.log('  No rows found for this PO.');
    } else {
      for (const [i, r] of rows.entries()) {
        console.log(`\n  Row ${i + 1}:`);
        for (const [k, v] of Object.entries(r)) {
          if (v !== null && v !== undefined && v !== '') {
            console.log(`    ${k.padEnd(50)} = ${v}`);
          }
        }
      }
    }
  } catch (err) {
    console.log(`  ERROR: ${err.message}`);
  }
}

async function sampleAsn(poRef) {
  console.log(`\n${'─'.repeat(70)}`);
  console.log(`SAMPLE: dim_advanced_shipment_notice_v1 for PO '${poRef}'`);
  console.log('─'.repeat(70));
  try {
    const rows = await db.query(`
      SELECT asn.*
      FROM supplychain.serve.dim_advanced_shipment_notice_v1 asn
      JOIN sourcingandbuying.serve.fact_purchase_order_commitment_v1 f ON f.dim_advanced_shipment_notice_sk = asn.dim_advanced_shipment_notice_sk
      JOIN sourcingandbuying.serve.dim_purchase_order_v1 po ON f.dim_purchase_order_sk = po.dim_purchase_order_sk
      WHERE po.purchase_order_id = '${poRef}'
      LIMIT 2
    `);
    if (!rows || rows.length === 0) {
      console.log('  No rows found.');
    } else {
      for (const [i, r] of rows.entries()) {
        console.log(`\n  Row ${i + 1}:`);
        for (const [k, v] of Object.entries(r)) {
          if (v !== null && v !== undefined && v !== '') {
            console.log(`    ${k.padEnd(50)} = ${v}`);
          }
        }
      }
    }
  } catch (err) {
    console.log(`  ERROR: ${err.message}`);
  }
}

async function main() {
  const poRef = process.argv[2];

  // Describe all serve tables
  for (const t of SERVE_TABLES) {
    await describeTable(t);
  }

  // Sample data if a PO ref was provided
  if (poRef) {
    console.log(`\n${'='.repeat(70)}`);
    console.log(`SAMPLE DATA for PO: ${poRef}`);
    console.log('='.repeat(70));
    await sampleFact(poRef);
    await sampleAsn(poRef);
  } else {
    console.log('\nTip: pass a PO number to also see sample data:');
    console.log('  node backend/explore-serve-tables.js 500036995371');
  }
}

main().catch(err => { console.error(err.message); process.exit(1); });
