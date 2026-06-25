'use strict';
/**
 * sharepoint-scheduler.js
 *
 * Runs a cron job at times specified in SP_SCHEDULE (e.g. "09:00,13:00").
 * Each run:
 *   1. Lists all Excel files in the configured SharePoint folder.
 *   2. Downloads each one and saves it to bible/sharepoint-sync/.
 *   3. Auto-parses them using supplier-reader and updates sessionState
 *      (identical to what happens when a user uploads files manually).
 *   4. Writes a status file (bible/sp-sync-status.json) so the UI can
 *      show the last-sync time and outcome.
 *
 * Schedule is driven by node-cron.
 * The sessionState object is passed in from server.js (shared reference).
 */

const cron          = require('node-cron');
const path          = require('path');
const fs            = require('fs');
const sp            = require('./sharepoint-client');
const supplierReader = require('./supplier-reader');

const SYNC_DIR        = path.join(__dirname, '..', 'bible', 'sharepoint-sync');
const STATUS_FILE     = path.join(__dirname, '..', 'bible', 'sp-sync-status.json');

// ── Status helpers ────────────────────────────────────────────────────────────

function readStatus() {
  try {
    if (fs.existsSync(STATUS_FILE)) return JSON.parse(fs.readFileSync(STATUS_FILE, 'utf8'));
  } catch (_) {}
  return { lastSync: null, files: [], error: null, running: false };
}

function writeStatus(patch) {
  const current = readStatus();
  const next = { ...current, ...patch };
  fs.mkdirSync(path.dirname(STATUS_FILE), { recursive: true });
  fs.writeFileSync(STATUS_FILE, JSON.stringify(next, null, 2));
}

// ── Core sync logic ───────────────────────────────────────────────────────────

async function runSync(sessionState) {
  if (!sp.isConfigured()) {
    writeStatus({ error: 'SharePoint not configured in .env', running: false });
    return;
  }

  console.log('[SP Scheduler] Starting sync…');
  const now = new Date();
  writeStatus({ running: true, error: null });
  fs.mkdirSync(SYNC_DIR, { recursive: true });

  let files;
  try {
    files = await sp.listTemplateFiles();
  } catch (err) {
    console.error('[SP Scheduler] List failed:', err.message);
    writeStatus({ running: false, error: `List failed: ${err.message}`, lastSync: new Date().toISOString() });
    return;
  }

  if (!files.length) {
    console.warn('[SP Scheduler] No Excel files found in SharePoint folder.');
    writeStatus({ running: false, files: [], lastSync: new Date().toISOString(), error: null });
    return;
  }

  // Pick the most recently modified Excel file — no filename restrictions
  files.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
  const latest = files[0];
  console.log(`[SP Scheduler] Latest file: "${latest.name}" (modified ${latest.lastModifiedDateTime})`);

  // Skip only if this exact file version was already processed last time
  const prevStatus = readStatus();
  const alreadyProcessed =
    prevStatus.lastProcessedId === latest.id &&
    prevStatus.lastProcessedModified === latest.lastModifiedDateTime;

  if (alreadyProcessed) {
    console.log(`[SP Scheduler] "${latest.name}" unchanged since last sync — skipping.`);
    writeStatus({ running: false, lastSync: now.toISOString(), skipped: true });
    return;
  }

  const downloaded = [];
  const buffers    = [];

  try {
    const buf       = await sp.downloadFile(latest.id);
    const localPath = path.join(SYNC_DIR, latest.name);
    fs.writeFileSync(localPath, buf);
    downloaded.push({ name: latest.name, size: buf.length, lastModified: latest.lastModifiedDateTime, localPath });
    buffers.push({ name: latest.name, buffer: buf });
    console.log(`[SP Scheduler] Downloaded latest: ${latest.name} (${Math.round(buf.length / 1024)} KB)`);
  } catch (err) {
    console.error(`[SP Scheduler] Download failed for ${latest.name}:`, err.message);
    writeStatus({ running: false, lastSync: now.toISOString(), error: `Download failed: ${err.message}` });
    return;
  }

  // Parse all downloaded files and merge into sessionState (same as parse-supplier endpoint)
  let allRows = [];
  let allValidationErrors = [];
  let allHeaderPoRefs = [];

  for (const f of buffers) {
    try {
      const parsed = await supplierReader.parse(f.buffer);
      allRows               = allRows.concat(parsed.rows);
      allValidationErrors   = allValidationErrors.concat(
        (parsed.validationErrors || []).map(e => `[${f.name}] ${e}`)
      );
      allHeaderPoRefs.push(...(parsed.headerPoRefs || []));
      console.log(`[SP Scheduler] Parsed ${f.name}: ${parsed.rows.length} row(s)`);
    } catch (err) {
      console.error(`[SP Scheduler] Parse failed for ${f.name}:`, err.message);
    }
  }

  // Update session state (the same object server.js holds)
  sessionState.supplierData         = { rows: allRows, validationErrors: allValidationErrors };
  sessionState.supplierHeaderPoRefs = allHeaderPoRefs;
  sessionState.supplierBuffers      = buffers;
  sessionState.feedData             = null;
  sessionState.masterData           = null;
  sessionState.lastXml              = null;
  sessionState.lastFilename         = null;

  const poRefs = [...new Set(allHeaderPoRefs.map(p => String(p).trim()).filter(Boolean))];
  console.log(`[SP Scheduler] Sync complete — ${buffers.length} file(s), ${allRows.length} row(s), ${poRefs.length} PO(s)`);

  writeStatus({
    running:               false,
    lastSync:              now.toISOString(),
    skipped:               false,
    error:                 null,
    lastProcessedId:       latest.id,
    lastProcessedModified: latest.lastModifiedDateTime,
    files:    downloaded.map(f => ({ name: f.name, size: f.size, lastModified: f.lastModified })),
    poRefs,
    rowCount: allRows.length
  });
}

// ── Schedule builder ──────────────────────────────────────────────────────────

/**
 * Parse "09:00,13:00" into cron expressions ["0 9 * * *", "0 13 * * *"]
 */
function timesToCron(timesStr) {
  return (timesStr || '').split(',').map(t => {
    const [h, m] = t.trim().split(':').map(Number);
    if (isNaN(h) || isNaN(m)) return null;
    return `${m} ${h} * * *`;
  }).filter(Boolean);
}

function start(sessionState) {
  const timesStr = process.env.SP_SCHEDULE || '';
  const cronExprs = timesToCron(timesStr);

  if (!cronExprs.length) {
    console.log('[SP Scheduler] SP_SCHEDULE not set — scheduler disabled.');
    return;
  }

  for (const expr of cronExprs) {
    cron.schedule(expr, () => {
      runSync(sessionState).catch(err =>
        console.error('[SP Scheduler] Unexpected error:', err.message)
      );
    });
    console.log(`[SP Scheduler] Scheduled at cron "${expr}" (from SP_SCHEDULE=${timesStr})`);
  }
}

module.exports = { start, runSync, readStatus };
