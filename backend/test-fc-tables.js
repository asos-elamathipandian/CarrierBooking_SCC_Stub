'use strict';
/**
 * Discovery script — find FC/warehouse/location tables in Databricks.
 * Run: node backend/test-fc-tables.js
 */
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const db = require('./databricks-client');

async function main() {
  console.log('Searching for FC/warehouse/location tables in Databricks...\n');

  // 1. Search for tables with FC/warehouse/location in name across known catalogs
  const catalogs = ['supplychain', 'sourcingandbuying'];
  for (const cat of catalogs) {
    try {
      const tables = await db.query(
        `SHOW TABLES IN ${cat}.conformed LIKE '*fc*'`
      );
      if (tables.length) console.log(`[${cat}.conformed] tables matching *fc*:`, tables.map(r => r.tableName || r.table_name || JSON.stringify(r)));
      else console.log(`[${cat}.conformed] no tables matching *fc*`);

      const tables2 = await db.query(
        `SHOW TABLES IN ${cat}.conformed LIKE '*warehouse*'`
      );
      if (tables2.length) console.log(`[${cat}.conformed] tables matching *warehouse*:`, tables2.map(r => r.tableName || r.table_name || JSON.stringify(r)));
      else console.log(`[${cat}.conformed] no tables matching *warehouse*`);

      const tables3 = await db.query(
        `SHOW TABLES IN ${cat}.conformed LIKE '*location*'`
      );
      if (tables3.length) console.log(`[${cat}.conformed] tables matching *location*:`, tables3.map(r => r.tableName || r.table_name || JSON.stringify(r)));
      else console.log(`[${cat}.conformed] no tables matching *location*`);

      const tables4 = await db.query(
        `SHOW TABLES IN ${cat}.conformed LIKE '*fulfil*'`
      );
      if (tables4.length) console.log(`[${cat}.conformed] tables matching *fulfil*:`, tables4.map(r => r.tableName || r.table_name || JSON.stringify(r)));
      else console.log(`[${cat}.conformed] no tables matching *fulfil*`);

    } catch (e) {
      console.log(`[${cat}] error:`, e.message);
    }
  }

  // 2. Also check what finalDestination values look like in shipment data
  try {
    const sample = await db.query(`
      SELECT DISTINCT finalDestination, firstDestination
      FROM supplychain.conformed.aim_shipment_detail_v1
      WHERE finalDestination IS NOT NULL
      LIMIT 20
    `);
    console.log('\nSample finalDestination values from aim_shipment_detail_v1:');
    sample.forEach(r => console.log(' ', JSON.stringify(r)));
  } catch (e) {
    console.log('Could not query aim_shipment_detail_v1:', e.message);
  }
}

main().catch(err => { console.error('FAILED:', err.message); process.exit(1); });
