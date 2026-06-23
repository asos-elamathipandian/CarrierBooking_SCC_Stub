'use strict';

/**
 * Run this once to verify Databricks connectivity and print column names
 * for the tables we plan to use as the carrier feed replacement.
 *
 * Usage:
 *   node backend/test-databricks.js
 */

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const db = require('./databricks-client');

const TABLES = [
  'aim_shipment_detail_v1',
];

async function main() {
  const cfg = db.getConfig();
  if (!cfg) {
    console.error('❌  Databricks not configured.\n   Set DATABRICKS_HOST, DATABRICKS_HTTP_PATH and DATABRICKS_TOKEN in .env');
    process.exit(1);
  }

  console.log(`\nConnecting to ${cfg.host} …\n`);

  for (const table of TABLES) {
    console.log(`── ${table} ──────────────────────────────`);
    try {
      // Schema
      const schema = await db.query(`DESCRIBE TABLE supplychain.conformed.${table}`);
      console.log('Columns:');
      for (const col of schema) {
        console.log(`  ${String(col.col_name || col.column_name || JSON.stringify(col)).padEnd(45)} ${col.data_type || col.type || ''}`);
      }

      // Sample row
      const sample = await db.query(`SELECT * FROM supplychain.conformed.${table} LIMIT 1`);
      if (sample.length) {
        console.log('\nSample row:');
        console.log(JSON.stringify(sample[0], null, 2));
      }
    } catch (err) {
      console.error(`  ❌  ${err.message}`);
    }
    console.log();
  }

  console.log('Done.');
}

main().catch(err => { console.error(err.message); process.exit(1); });
