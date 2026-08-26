'use strict';
// One-off script: send a sample report email using the last N log entries.
// Usage: node backend/test-report-send.js [count=3]
require('dotenv').config();
const path          = require('path');
const fs            = require('fs');
const reportSender  = require('./report-sender');

const LOG_PATH = path.join(__dirname, '..', 'bible', 'generation-log.json');
const count    = parseInt(process.argv[2] || '3', 10);

async function main() {
  if (!reportSender.isConfigured()) {
    console.error('REPORT_TO not configured — check .env');
    process.exit(1);
  }

  const allEntries = JSON.parse(fs.readFileSync(LOG_PATH, 'utf8'));
  const sample     = allEntries.slice(-count);

  if (!sample.length) {
    console.error('No entries in generation log.');
    process.exit(1);
  }

  console.log(`Sending sample report with ${sample.length} entr${sample.length === 1 ? 'y' : 'ies'}…`);

  // Temporarily patch readLog/readState so sendScheduledReport picks up our sample entries
  const bibleBuilder = require('./bible-builder');
  const origGet = bibleBuilder.getGenerationLog;
  bibleBuilder.getGenerationLog = () => sample;

  // Reset lastReportTime to epoch so all sample entries pass the filter
  const STATE_PATH = path.join(__dirname, '..', 'bible', 'report-state.json');
  const origState  = fs.existsSync(STATE_PATH) ? fs.readFileSync(STATE_PATH, 'utf8') : null;
  fs.writeFileSync(STATE_PATH, JSON.stringify({ lastReportTime: new Date(0).toISOString() }));

  try {
    await reportSender.sendScheduledReport({
      supplierBuffers:      [],
      lastGenerations:      sample,
      supplierHeaderPoRefs: [...new Set(sample.flatMap(e => e.poNumbers || []))],
      skippedGroups:        [],
      cancelledItems:       []
    });
    console.log('Done — check your inbox.');
  } finally {
    bibleBuilder.getGenerationLog = origGet;
    // Restore original state so real report tracking is unaffected
    if (origState) fs.writeFileSync(STATE_PATH, origState);
    else fs.unlinkSync(STATE_PATH);
  }
}

main().catch(err => { console.error(err.message); process.exit(1); });
