'use strict';
// Verifies carrier-cancellation detection end to end against real POs.
// Run: node backend/test-carrier-cancel-detection.js [po...]
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const reader = require('./databricks-asn-reader');

const POS = process.argv.slice(2).length ? process.argv.slice(2) : ['500038036055', '500036576514'];

(async () => {
  const res = await reader.fetchAsnsByPoRefs(POS);
  console.log('\nerrors:', res.errors);
  console.log('cancelledItems:', JSON.stringify(res.cancelledItems, null, 2));
  console.log('carrierCancelledRefs:', JSON.stringify(res.carrierCancelledRefs, null, 2));
  for (const f of res.carrierAsnFiles || []) {
    for (const g of f.parsed || []) {
      console.log(`PO ${g.poId} / ASN ${g.asnId} — booked=${g.isBookedByCarrier}, lastBooked=${g.lastCarrierBookedDate || '-'}, carrierCancelled=${!!g.carrierCancelled}, lines=${g.lines.length}`);
    }
  }
  process.exit(0);
})().catch(err => { console.error('FAILED:', err.message); process.exit(1); });
