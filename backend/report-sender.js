'use strict';

/**
 * report-sender.js
 * Sends an HTML email report after each scheduler run listing all carrier
 * booking requests generated since the last report.
 *
 * Uses the same App Registration as SharePoint (client-credentials flow).
 * Required .env vars:
 *   SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET  — existing Graph auth
 *   REPORT_TO      — comma-separated recipient address(es)
 *   REPORT_FROM    — sender mailbox (defaults to EMAIL_INGEST_MAILBOX)
 *                    The app registration needs Mail.Send on this mailbox.
 */

const path    = require('path');
const fs      = require('fs');
const ExcelJS = require('exceljs');
const { ClientSecretCredential } = require('@azure/identity');

const LOG_PATH   = path.join(__dirname, '..', 'bible', 'generation-log.json');
const STATE_PATH = path.join(__dirname, '..', 'bible', 'report-state.json');

const GRAPH_BASE  = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

// ── Auth ──────────────────────────────────────────────────────────────────────

function isConfigured() {
  return !!(
    process.env.SP_TENANT_ID &&
    process.env.SP_CLIENT_ID &&
    process.env.SP_CLIENT_SECRET &&
    process.env.REPORT_TO
  );
}

let _credential = null;
function getCredential() {
  if (!_credential) {
    _credential = new ClientSecretCredential(
      process.env.SP_TENANT_ID,
      process.env.SP_CLIENT_ID,
      process.env.SP_CLIENT_SECRET
    );
  }
  return _credential;
}

async function getAccessToken() {
  const token = await getCredential().getToken(GRAPH_SCOPE);
  return token.token;
}

async function graphPost(apiPath, body) {
  const token = await getAccessToken();
  const res = await fetch(`${GRAPH_BASE}${apiPath}`, {
    method:  'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body:    JSON.stringify(body)
  });
  if (res.status !== 202 && !res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${text}`);
  }
}

// ── State (tracks last report time) ──────────────────────────────────────────

function readLog() {
  try { return JSON.parse(fs.readFileSync(LOG_PATH, 'utf8')); } catch { return []; }
}
function readState() {
  try { return JSON.parse(fs.readFileSync(STATE_PATH, 'utf8')); } catch { return {}; }
}
function writeState(state) {
  fs.mkdirSync(path.dirname(STATE_PATH), { recursive: true });
  fs.writeFileSync(STATE_PATH, JSON.stringify(state, null, 2));
}

// ── HTML builder ──────────────────────────────────────────────────────────────

function buildSummaryHtml(entries, runTime, sessionCtx) {
  const dateStr = new Date(runTime).toLocaleString('en-GB', { dateStyle: 'full', timeStyle: 'short' });

  const newBookings   = entries.filter(e => !e.purposeCd || e.purposeCd === '13').length;
  const resubmitted   = entries.filter(e => e.purposeCd === '15').length;
  const cancellations = entries.filter(e => e.purposeCd === '01').length;

  const totalSubmitted   = (sessionCtx.supplierHeaderPoRefs || []).length;
  const skippedCount     = (sessionCtx.skippedGroups || []).length;
  const alreadyBooked    = (sessionCtx.cancelledItems || []).filter(c => c.type === 'ALREADY_BOOKED').length;
  const asnCancelled     = (sessionCtx.cancelledItems || []).filter(c => c.type !== 'ALREADY_BOOKED').length;

  const suppliers = [...new Set(entries.map(e => e.supplier).filter(Boolean))];

  const row = (label, value, color = '#222') =>
    `<tr><td style="padding:7px 16px 7px 0;color:#555;white-space:nowrap">${label}</td>` +
    `<td style="padding:7px 0;font-weight:bold;color:${color}">${value}</td></tr>`;

  const hasAttachment = !!(sessionCtx.supplierBuffers?.length && sessionCtx.lastGenerations?.length);

  return `<!DOCTYPE html>
<html><head><style>
  body  { font-family: Calibri, Arial, sans-serif; font-size: 13px; color: #222; margin: 24px; }
  h2    { color: #1F4E79; margin-bottom: 4px; }
  .card { background:#f5f8fc; border:1px solid #d0dce8; border-radius:6px; padding:16px 20px; display:inline-block; margin-top:12px; }
  .note { margin-top:16px; color:#444; font-size:12px; }
  .footer { margin-top:24px; color:#aaa; font-size:11px; border-top:1px solid #e0e0e0; padding-top:8px; }
</style></head><body>
  <h2>&#128666; Carrier Booking Request &mdash; Run Report</h2>
  <p style="color:#555">${dateStr}</p>
  <div class="card">
    <table style="border-collapse:collapse">
      ${suppliers.length ? row('Supplier',                    suppliers.join(', '))                    : ''}
      ${totalSubmitted  ? row('Total POs submitted',          totalSubmitted)                          : ''}
      ${row('New bookings generated',      newBookings   || 0, newBookings   ? '#1e7e34' : '#888')}
      ${resubmitted   ? row('Re-submissions (data changed)',  resubmitted,   '#d97706') : ''}
      ${cancellations ? row('Cancellations',                  cancellations, '#c0392b') : ''}
      ${skippedCount  ? row('Skipped (no changes)',           skippedCount,  '#888')    : ''}
      ${alreadyBooked ? row('Already booked externally',      alreadyBooked, '#888')    : ''}
      ${asnCancelled  ? row('ASN cancelled / no ASN',         asnCancelled,  '#888')    : ''}
    </table>
  </div>
  ${hasAttachment
    ? `<p class="note">&#128206; The tagged supplier template with booking details is attached.</p>`
    : `<p class="note" style="color:#888">No supplier template available to attach for this run.</p>`}
  <div class="footer">Generated by ASOS Carrier Booking Tool &mdash; Azure hosted</div>
</body></html>`;
}

// ── Supplier Excel tagging ────────────────────────────────────────────────────

/**
 * Write the VB Ref and booking detail columns into each supplier Excel buffer.
 * Non-fatal: returns [] if tagging fails.
 */
async function buildTaggedSupplierAttachments(supplierBuffers, generations, logEntries = []) {
  if (!supplierBuffers?.length || !generations?.length) return [];

  // Build per-PO detail map from log entries (richer than generations array)
  const poToDetail = {};
  for (const e of logEntries) {
    const detail = {
      ref:           e.bookingRef || '',
      filename:      e.filename   || '',
      asnRefs:       (e.asnRefs   || []).join(', '),
      noOfCartons:   e.noOfCartons  != null ? String(e.noOfCartons)  : '',
      cargoReadyDate: e.cargoReadyDate || '',
      purposeCd:     e.purposeCd || '13',
      sftp:          e.sftp || ''
    };
    for (const po of (e.poNumbers || [])) poToDetail[String(po).trim()] = detail;
  }
  // Fallback: fill any POs only in generations (no log entry yet)
  const poToRef = {};
  for (const gen of generations) {
    const ref = gen.bookingRef || gen.ctrlNumber || gen.filename || '';
    for (const po of (gen.poNumbers || [])) {
      poToRef[String(po).trim()] = ref;
      if (!poToDetail[String(po).trim()]) {
        poToDetail[String(po).trim()] = { ref, filename: gen.filename || '', asnRefs: (gen.asnRefs || []).join(', '), noOfCartons: '', cargoReadyDate: '', purposeCd: '13', sftp: '' };
      }
    }
  }

  const purposeLabel = cd => cd === '15' ? 'Re-submit' : cd === '01' ? 'Cancel' : 'New';
  const sftpLabel    = s  => s === 'uploaded' ? 'Uploaded' : s === 'local' ? 'Local' : s === 'error' ? 'Error' : '';

  // Extra columns appended after VBKREQ_Ref
  const EXTRA_COLS = [
    { header: 'VBKREQ_File',    key: 'filename',       width: 55 },
    { header: 'ASN_Refs',       key: 'asnRefs',        width: 30 },
    { header: 'Cartons',        key: 'noOfCartons',    width: 10 },
    { header: 'Cargo_Ready',    key: 'cargoReadyDate', width: 16 },
    { header: 'Booking_Type',   key: 'purposeCd',      width: 12, transform: purposeLabel },
    { header: 'SFTP_Status',    key: 'sftp',           width: 12, transform: sftpLabel  },
  ];

  const attachments = [];
  for (const file of supplierBuffers) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(file.buffer);

      // Auto-detect PO Header sheet (same logic as /api/tagged-supplier/:idx)
      const wsH = wb.getWorksheet('PO Header') || wb.getWorksheet('BOOKING_HEADER') ||
        wb.worksheets.find(ws => {
          let found = false;
          ws.eachRow((row) => { row.eachCell(c => { if (String(c.value||'').trim()==='PO_Number') found=true; }); });
          return found;
        });
      if (!wsH) { console.warn(`[Report] No PO Header sheet in ${file.name} — skipping attachment`); continue; }

      let headerRowNum = 1, poColIdx = 1;
      wsH.eachRow((row, rowNum) => {
        row.eachCell((cell, colNum) => {
          const v = String(cell.value || '').replace(/\s*\(.*?\)/, '').trim();
          if (v === 'PO_Number') { headerRowNum = rowNum; poColIdx = colNum; }
        });
      });

      const baseColIdx = wsH.columnCount + 1;
      const allNewCols  = ['VBKREQ_Ref', ...EXTRA_COLS.map(c => c.header)];
      const colWidths   = [52, ...EXTRA_COLS.map(c => c.width)];

      // Write header cells for all new columns
      allNewCols.forEach((hdr, i) => {
        const ci   = baseColIdx + i;
        const cell = wsH.getRow(headerRowNum).getCell(ci);
        cell.value = hdr;
        cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
        cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        wsH.getColumn(ci).width = colWidths[i];
      });

      wsH.eachRow((row, rowNum) => {
        if (rowNum <= headerRowNum) return;
        const po     = String(row.getCell(poColIdx).value || '').trim();
        if (!po) return;
        const detail = poToDetail[po];
        const refCell = row.getCell(baseColIdx);
        if (detail?.ref) {
          refCell.value = detail.ref;
          refCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
          refCell.font  = { color: { argb: 'FF1B5E20' }, bold: true, size: 10 };
        } else {
          refCell.value = 'Not generated';
          refCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE8E8' } };
          refCell.font  = { color: { argb: 'FF7B1F1F' }, size: 10 };
        }
        // Write extra detail columns
        EXTRA_COLS.forEach((col, i) => {
          const ci   = baseColIdx + 1 + i;
          const cell = row.getCell(ci);
          const raw  = detail ? (detail[col.key] ?? '') : '';
          cell.value = col.transform ? col.transform(raw) : raw;
          cell.font  = { size: 10 };
        });
        row.commit();
      });

      // Strip conditional formatting to avoid ExcelJS serialization errors
      for (const ws of wb.worksheets) ws.conditionalFormattings = [];

      const buf      = await wb.xlsx.writeBuffer();
      const baseName = file.name.replace(/\.xlsx?$/i, '');
      attachments.push({
        '@odata.type':  '#microsoft.graph.fileAttachment',
        name:           `${baseName}_VBRef.xlsx`,
        contentType:    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        contentBytes:   Buffer.from(buf).toString('base64')
      });
    } catch (err) {
      console.warn(`[Report] Failed to tag ${file.name} for attachment:`, err.message);
    }
  }
  return attachments;
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Send a report email covering all generation-log entries since the last report.
 * Pass sessionCtx = { supplierBuffers, lastGenerations } to attach tagged supplier Excel files.
 */
async function sendScheduledReport(sessionCtx = {}) {
  if (!isConfigured()) {
    console.log('[Report] REPORT_TO not configured — skipping report.');
    return;
  }

  const state = readState();
  const lastReportTime = state.lastReportTime ? new Date(state.lastReportTime) : new Date(0);
  const now = new Date();

  const allEntries = readLog();
  const newEntries = allEntries.filter(e => new Date(e.timestamp) > lastReportTime);

  if (newEntries.length === 0) {
    console.log('[Report] No new entries since last report — skipping.');
    return;
  }

  const fromMailbox = process.env.REPORT_FROM || process.env.EMAIL_INGEST_MAILBOX;
  const toList = (process.env.REPORT_TO || '').split(',').map(s => s.trim()).filter(Boolean);

  if (!fromMailbox) {
    console.warn('[Report] No sender mailbox (set REPORT_FROM or EMAIL_INGEST_MAILBOX) — skipping.');
    return;
  }

  const nowGb   = now.toLocaleString('en-GB', { dateStyle: 'short', timeStyle: 'short' });
  const subject = `Carrier Booking Report — ${nowGb} (${newEntries.length} booking${newEntries.length !== 1 ? 's' : ''})`;
  const html    = buildSummaryHtml(newEntries, now.toISOString(), sessionCtx);

  const attachments = await buildTaggedSupplierAttachments(
    sessionCtx.supplierBuffers,
    sessionCtx.lastGenerations,
    newEntries
  );

  try {
    await graphPost(`/users/${encodeURIComponent(fromMailbox)}/sendMail`, {
      message: {
        subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: toList.map(addr => ({ emailAddress: { address: addr } })),
        ...(attachments.length ? { attachments } : {})
      },
      saveToSentItems: false
    });
    console.log(`[Report] Sent to ${toList.join(', ')} — ${newEntries.length} booking(s) reported${attachments.length ? `, ${attachments.length} attachment(s)` : ''}.`);
    writeState({ lastReportTime: now.toISOString() });
  } catch (err) {
    console.error('[Report] Failed to send:', err.message);
    // Non-fatal — do not throw; scheduler continues regardless
  }
}

module.exports = { isConfigured, sendScheduledReport };
