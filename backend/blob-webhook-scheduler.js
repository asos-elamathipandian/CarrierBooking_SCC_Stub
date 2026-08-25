'use strict';

/**
 * blob-webhook-scheduler.js
 *
 * Polls the Azure Blob Storage container (fed by Power Automate) for new
 * supplier Excel files, parses them, and merges into sessionState —
 * mirrors sharepoint-scheduler.js's runSync() but sourced from Blob
 * Storage instead of a direct Graph SharePoint read.
 *
 * Primary inbound path (Power Automate → Blob). Report emails are sent
 * from here via Mail.Send. SharePoint scheduler remains as fallback.
 *
 * Status is tracked in bible/blob-sync-status.json so the UI can show
 * last-sync time/outcome, same shape as sp-sync-status.json.
 */

const path           = require('path');
const fs             = require('fs');
const cron           = require('node-cron');
const blob           = require('./blob-webhook-client');
const supplierReader = require('./supplier-reader');
const reportSender   = require('./report-sender');
const pipeline       = require('./pipeline-runner');

const STATUS_FILE = path.join(__dirname, '..', 'bible', 'blob-sync-status.json');

function readStatus() {
  try {
    if (fs.existsSync(STATUS_FILE)) return JSON.parse(fs.readFileSync(STATUS_FILE, 'utf8'));
  } catch (_) {}
  return { lastSync: null, files: [], error: null, running: false, processedMap: {} };
}

function writeStatus(patch) {
  const current = readStatus();
  const next = { ...current, ...patch };
  fs.mkdirSync(path.dirname(STATUS_FILE), { recursive: true });
  fs.writeFileSync(STATUS_FILE, JSON.stringify(next, null, 2));
}

/**
 * List new/updated blobs, parse them, and merge rows into sessionState.
 * Successfully processed blobs are deleted so the container only ever
 * holds unprocessed files.
 */
async function runSync(sessionState) {
  if (!blob.isConfigured()) {
    writeStatus({ error: 'Blob webhook storage not configured in .env', running: false });
    return;
  }

  console.log('[Blob Sync] Starting sync…');
  const now = new Date();
  writeStatus({ running: true, error: null });

  let files;
  try {
    files = await blob.listBlobs();
  } catch (err) {
    console.error('[Blob Sync] List failed:', err.message);
    writeStatus({ running: false, error: `List failed: ${err.message}`, lastSync: now.toISOString() });
    return;
  }

  if (!files.length) {
    console.log('[Blob Sync] No new supplier files in container.');
    writeStatus({ running: false, files: [], lastSync: now.toISOString(), error: null });
    return;
  }

  let allRows = [];
  let allValidationErrors = [];
  let allHeaderPoRefs = [];
  const processed = [];

  for (const file of files) {
    try {
      const buffer = await blob.downloadBlob(file.name);
      const parsed = await supplierReader.parse(buffer);
      allRows = allRows.concat(parsed.rows);
      allValidationErrors = allValidationErrors.concat(
        (parsed.validationErrors || []).map(e => `[${file.name}] ${e}`)
      );
      allHeaderPoRefs.push(...(parsed.headerPoRefs || []));
      processed.push({ name: file.name, size: file.size, lastModified: file.lastModified });
      await blob.archiveBlob(file.name);
      console.log(`[Blob Sync] Ingested and archived "${file.name}" — ${parsed.rows.length} row(s)`);
    } catch (err) {
      console.error(`[Blob Sync] Failed to process "${file.name}":`, err.message);
    }
  }

  if (!processed.length) {
    writeStatus({ running: false, lastSync: now.toISOString(), error: 'All files failed to process' });
    return;
  }

  sessionState.supplierData         = { rows: allRows, validationErrors: allValidationErrors };
  sessionState.supplierHeaderPoRefs = allHeaderPoRefs;
  sessionState.feedData             = null;
  sessionState.masterData           = null;
  sessionState.lastXml              = null;
  sessionState.lastFilename         = null;

  const poRefs = [...new Set(allHeaderPoRefs.map(p => String(p).trim()).filter(Boolean))];
  console.log(`[Blob Sync] Sync complete — ${processed.length} file(s), ${allRows.length} row(s), ${poRefs.length} PO(s)`);

  writeStatus({
    running:  false,
    lastSync: now.toISOString(),
    error:    null,
    files:    processed,
    poRefs,
    rowCount: allRows.length
  });

  // When BLOB_PIPELINE_SCHEDULE is set, hold here — the cron will fire the pipeline at the scheduled slot
  if (process.env.BLOB_PIPELINE_SCHEDULE) {
    console.log('[Blob Sync] BLOB_PIPELINE_SCHEDULE set — deferring pipeline run to next scheduled slot.');
    return;
  }

  // Run the full pipeline: Databricks → bible → VBKREQs → SFTP → report
  const result = await pipeline.run(sessionState);
  if (result.error) {
    console.error(`[Blob Sync] Pipeline error: ${result.error}`);
    writeStatus({ pipelineError: result.error });
  } else {
    writeStatus({
      pipelineError:    null,
      lastPipelineRun:  now.toISOString(),
      generationCount:  result.generations.length,
      skippedCount:     result.skippedGroups.length,
    });
  }
}

// ── Scheduled pipeline run (BLOB_PIPELINE_SCHEDULE mode) ─────────────────────

async function runScheduledPipeline(sessionState) {
  if (!(sessionState.supplierData?.rows?.length)) {
    console.log('[Blob Pipeline] No supplier data in session — nothing to run.');
    return;
  }
  console.log('[Blob Pipeline] Scheduled pipeline run starting…');
  const now = new Date();
  const result = await pipeline.run(sessionState);
  if (result.error) {
    console.error('[Blob Pipeline] Pipeline error:', result.error);
    writeStatus({ pipelineError: result.error });
  } else {
    writeStatus({
      pipelineError:   null,
      lastPipelineRun: now.toISOString(),
      generationCount: result.generations.length,
      skippedCount:    result.skippedGroups.length,
    });
    console.log(`[Blob Pipeline] Done — ${result.generations.length} generated, ${result.skippedGroups.length} skipped.`);
  }
}

// ── Schedule builder ─────────────────────────────────────────────────────────

function timesToCron(timesStr) {
  return (timesStr || '').split(',').map(t => {
    const [h, m] = t.trim().split(':').map(Number);
    if (isNaN(h) || isNaN(m)) return null;
    return `${m} ${h} * * *`;
  }).filter(Boolean);
}

function start(sessionState) {
  // BLOB_PIPELINE_SCHEDULE: pipeline-only cron (blob ingest is instant, pipeline runs at these times)
  const pipelineCrons = timesToCron(process.env.BLOB_PIPELINE_SCHEDULE || '');
  if (pipelineCrons.length) {
    for (const expr of pipelineCrons) {
      cron.schedule(expr, () => {
        runScheduledPipeline(sessionState).catch(err =>
          console.error('[Blob Pipeline] Scheduled run error:', err.message)
        );
      });
      console.log(`[Blob Pipeline] Scheduled at cron "${expr}" (from BLOB_PIPELINE_SCHEDULE)`);
    }
  }

  // SP_SCHEDULE: legacy full blob sync + pipeline cron (instant mode only)
  const cronExprs = timesToCron(process.env.SP_SCHEDULE || '');
  if (!cronExprs.length && !pipelineCrons.length) {
    console.log('[Blob Sync] No schedule set — blob scheduler disabled.');
    return;
  }
  for (const expr of cronExprs) {
    cron.schedule(expr, () => {
      runSync(sessionState).catch(err =>
        console.error('[Blob Sync] Scheduled sync error:', err.message)
      );
    });
    console.log(`[Blob Sync] Scheduled at cron "${expr}" (from SP_SCHEDULE)`);
  }
}

module.exports = { runSync, start, readStatus, writeStatus };
