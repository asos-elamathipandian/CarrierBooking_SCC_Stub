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

const cron           = require('node-cron');
const path           = require('path');
const fs             = require('fs');
const sp             = require('./sharepoint-client');
const supplierReader  = require('./supplier-reader');
const emailIngestor  = require('./email-ingestor');

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

/**
 * Append one entry to syncHistory (max 10, newest first).
 * entry: { timestamp, outcome, rowCount, poCount, files, error }
 */
function appendHistory(entry) {
  const current = readStatus();
  const history = Array.isArray(current.syncHistory) ? current.syncHistory : [];
  history.unshift(entry);
  writeStatus({ syncHistory: history.slice(0, 10) });
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

  // ── Pre-step: pull any new emails from the dedicated ASOS mailbox ───────────
  // Runs before listing SharePoint so just-uploaded files are included in this cycle.
  if (emailIngestor.isConfigured()) {
    try {
      const ingestResult = await emailIngestor.ingest();
      if (ingestResult.uploaded > 0) {
        console.log(`[SP Scheduler] Email ingest: ${ingestResult.uploaded} file(s) uploaded to SharePoint.`);
      }
      if (ingestResult.errors.length) {
        console.warn(`[SP Scheduler] Email ingest errors: ${ingestResult.errors.join('; ')}`);
      }
    } catch (err) {
      // Non-fatal — log and continue with the SP sync
      console.error('[SP Scheduler] Email ingest failed (continuing):', err.message);
    }
  }
  // ── End pre-step ─────────────────────────────────────────────────────────────

  let files;
  try {
    files = await sp.listTemplateFiles();
  } catch (err) {
    console.error('[SP Scheduler] List failed:', err.message);
    writeStatus({ running: false, error: `List failed: ${err.message}`, lastSync: now.toISOString() });
    appendHistory({ timestamp: now.toISOString(), outcome: 'error', error: `List failed: ${err.message}`, rowCount: 0, poCount: 0, files: [] });
    return;
  }

  if (!files.length) {
    console.warn('[SP Scheduler] No Excel files found in SharePoint folder or subfolders.');
    writeStatus({ running: false, files: [], lastSync: now.toISOString(), error: null });
    appendHistory({ timestamp: now.toISOString(), outcome: 'no_files', rowCount: 0, poCount: 0, files: [], error: null });
    return;
  }

  // Load the set of already-processed file versions: { [id]: lastModifiedDateTime }
  const prevStatus = readStatus();
  const processedMap = prevStatus.processedMap || {};

  // Group files by supplier folder, pick the latest per folder, filter to unprocessed
  const byFolder = {};
  for (const f of files) {
    const key = f.supplierFolder || '__root__';
    if (!byFolder[key]) byFolder[key] = [];
    byFolder[key].push(f);
  }

  const toProcess = [];
  for (const [folder, folderFiles] of Object.entries(byFolder)) {
    folderFiles.sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));
    const latest = folderFiles[0];
    const alreadyProcessed = processedMap[latest.id] === latest.lastModifiedDateTime;
    if (alreadyProcessed) {
      console.log(`[SP Scheduler] "${latest.name}" (${folder}) — unchanged, skipping.`);
    } else {
      console.log(`[SP Scheduler] New/updated: "${latest.name}" (${folder})`);
      toProcess.push(latest);
    }
  }

  if (!toProcess.length) {
    console.log('[SP Scheduler] All supplier files already processed — nothing to do.');
    writeStatus({ running: false, lastSync: now.toISOString(), skipped: true });
    appendHistory({ timestamp: now.toISOString(), outcome: 'skipped', rowCount: 0, poCount: 0, files: [], error: null });
    return;
  }

  const downloaded = [];
  const buffers    = [];

  for (const file of toProcess) {
    try {
      const buf       = await sp.downloadFile(file.id);
      const localName = file.supplierFolder ? `${file.supplierFolder}-${file.name}` : file.name;
      const localPath = path.join(SYNC_DIR, localName);
      fs.writeFileSync(localPath, buf);
      downloaded.push({ name: file.name, supplierFolder: file.supplierFolder, size: buf.length, lastModified: file.lastModifiedDateTime });
      buffers.push({ name: file.name, buffer: buf, id: file.id, lastModifiedDateTime: file.lastModifiedDateTime });
      console.log(`[SP Scheduler] Downloaded: ${file.name} from "${file.supplierFolder || 'root'}" (${Math.round(buf.length / 1024)} KB)`);
    } catch (err) {
      console.error(`[SP Scheduler] Download failed for ${file.name}:`, err.message);
    }
  }

  if (!buffers.length) {
    writeStatus({ running: false, lastSync: now.toISOString(), error: 'All downloads failed' });
    appendHistory({ timestamp: now.toISOString(), outcome: 'error', error: 'All downloads failed', rowCount: 0, poCount: 0, files: [], error: null });
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
      // Mark as processed
      processedMap[f.id] = f.lastModifiedDateTime;
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

  const syncedFiles = downloaded.map(f => ({ name: f.name, supplierFolder: f.supplierFolder, size: f.size, lastModified: f.lastModified }));
  writeStatus({
    running:      false,
    lastSync:     now.toISOString(),
    skipped:      false,
    error:        null,
    processedMap,
    files:    syncedFiles,
    poRefs,
    rowCount: allRows.length
  });
  appendHistory({
    timestamp: now.toISOString(),
    outcome:   'synced',
    rowCount:  allRows.length,
    poCount:   poRefs.length,
    files:     syncedFiles.map(f => f.supplierFolder ? `${f.name} (${f.supplierFolder})` : f.name),
    error:     null
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

  // ── Catch-up sync on startup ──────────────────────────────────────────────
  // If the server restarted after a scheduled slot passed today and no sync
  // has run today yet, fire one immediately so we don't wait until the next slot.
  const status = readStatus();
  const lastSync = status.lastSync ? new Date(status.lastSync) : null;
  const today = new Date().toDateString();
  const syncedToday = lastSync && lastSync.toDateString() === today;

  if (!syncedToday && sp.isConfigured()) {
    console.log('[SP Scheduler] No sync yet today — running catch-up sync on startup…');
    setTimeout(() => {
      runSync(sessionState).catch(err =>
        console.error('[SP Scheduler] Catch-up sync error:', err.message)
      );
    }, 3000); // 3s delay to let the server finish booting
  }
}

module.exports = { start, runSync, readStatus, writeStatus, appendHistory };
