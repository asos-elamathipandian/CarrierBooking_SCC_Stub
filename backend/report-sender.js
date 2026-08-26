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

function fmtDate(val) {
  if (!val) return '';
  const s = String(val);
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return new Date(s).toLocaleDateString('en-GB');
  return s;
}

function buildHtml(entries, runTime) {
  const dateStr = new Date(runTime).toLocaleString('en-GB', { dateStyle: 'full', timeStyle: 'short' });

  const rows = entries.map(e => `
    <tr>
      <td>${e.supplier || '—'}</td>
      <td>${(e.poNumbers || []).join('<br>')}</td>
      <td><strong>${e.bookingRef || ''}</strong></td>
      <td>${(e.asnRefs || []).join('<br>') || '—'}</td>
      <td style="font-size:11px;color:#555">${e.filename || ''}</td>
      <td style="font-size:11px;color:#555">${e.supplierTemplate || '—'}</td>
      <td>${e.bookingGroup || ''}</td>
      <td style="${e.purposeCd==='01'?'color:#c0392b':e.purposeCd==='15'?'color:#d97706':'color:#1e7e34'};font-weight:bold">${
        e.purposeCd === '01' ? 'Cancellation'
        : e.purposeCd === '15'
          ? 'Re-submission' + (e.resubmissionReason ? '<br/><span style="font-size:10px;font-weight:normal">' + e.resubmissionReason + '</span>' : '')
          : 'New'
      }</td>
      <td>${fmtDate(e.cargoReadyDate)}</td>
      <td style="text-align:center">${e.noOfCartons != null ? e.noOfCartons : '—'}</td>
      <td style="text-align:center">${e.totalWeight != null ? e.totalWeight : '—'}</td>
      <td style="${e.bkqDiscrepancy ? 'color:#c0392b;font-weight:bold;background:#fff5f5' : e.headerBkq != null ? 'color:#1e7e34' : 'color:#555'}">${
        e.bkqDiscrepancy
          ? `&#9888; Mismatch — Header: ${e.headerBkq} units | ASN lines: ${e.lineBkqSum} units`
          : e.headerBkq != null
            ? `&#10004; ${e.headerBkq} units (header matches ASN lines: ${e.lineBkqSum} units)`
            : `${e.lineBkqSum != null ? e.lineBkqSum + ' units (ASN lines only — no header override)' : '\u2014'}`
      }</td>      <td style="color:${e.sftp === 'uploaded' ? '#1e7e34' : '#c0392b'};font-weight:bold">
        ${e.sftp === 'uploaded' ? '&#10004; Uploaded' : e.sftp === 'error' ? `&#10008; Error${e.sftpError ? ': ' + e.sftpError : ''}` : e.sftp || 'Pending'}
      </td>
      <td style="color:#555;white-space:nowrap">${new Date(e.timestamp).toLocaleString('en-GB')}</td>
    </tr>`).join('');

  const tableOrMsg = entries.length === 0
    ? `<p style="color:#888;font-style:italic">No new carrier booking requests since the last report.</p>`
    : `<table>
        <thead><tr>
          <th>Supplier</th>
          <th>PO Number(s)</th>
          <th>VB Ref</th>
          <th>ASN Ref(s)</th>
          <th>Filename</th>
          <th>Supplier Template</th>
          <th>Booking Group</th>
          <th>Purpose</th>
          <th>Cargo Ready Date</th>
          <th>No. of Cartons</th>
          <th>Total Weight&nbsp;(KG)</th>
          <th>BKQ (Header vs Lines)</th>
          <th>SFTP Status</th>
          <th>Generated At</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;

  return `<!DOCTYPE html>
<html><head><style>
  body  { font-family: Calibri, Arial, sans-serif; font-size: 13px; color: #222; margin: 24px; }
  h2    { color: #1F4E79; margin-bottom: 4px; }
  p     { margin: 4px 0 12px; }
  table { border-collapse: collapse; width: 100%; margin-top: 12px; font-size: 12px; }
  th    { background: #1F4E79; color: #fff; padding: 7px 10px; text-align: left; white-space: nowrap; }
  td    { padding: 6px 10px; border-bottom: 1px solid #e0e0e0; vertical-align: top; }
  tr:nth-child(even) td { background: #f5f8fc; }
  .footer { margin-top: 24px; color: #aaa; font-size: 11px; border-top: 1px solid #e0e0e0; padding-top: 8px; }
</style></head><body>
  <h2>&#128666; Carrier Booking Request &mdash; Run Report</h2>
  <p>${dateStr} &nbsp;|&nbsp; <strong>${entries.length}</strong> booking request${entries.length !== 1 ? 's' : ''} in this report</p>
  ${tableOrMsg}
  <div class="footer">Generated by ASOS Carrier Booking Tool &mdash; Azure hosted</div>
</body></html>`;
}

// ── Supplier Excel tagging ────────────────────────────────────────────────────

/**
 * Write the VB Ref into each supplier Excel buffer and return named attachment objects.
 * Non-fatal: returns [] if tagging fails.
 */
async function buildTaggedSupplierAttachments(supplierBuffers, generations) {
  if (!supplierBuffers?.length || !generations?.length) return [];

  const poToRef = {};
  for (const gen of generations) {
    const ref = gen.bookingRef || gen.ctrlNumber || gen.filename || '';
    for (const po of (gen.poNumbers || [])) poToRef[String(po).trim()] = ref;
  }

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

      const newColIdx = wsH.columnCount + 1;
      const hdrCell   = wsH.getRow(headerRowNum).getCell(newColIdx);
      hdrCell.value = 'VBKREQ_Ref';
      hdrCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
      hdrCell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      hdrCell.alignment = { horizontal: 'center', vertical: 'middle' };
      wsH.getColumn(newColIdx).width = 52;

      wsH.eachRow((row, rowNum) => {
        if (rowNum <= headerRowNum) return;
        const po  = String(row.getCell(poColIdx).value || '').trim();
        if (!po) return;
        const ref  = poToRef[po];
        const cell = row.getCell(newColIdx);
        if (ref) {
          cell.value = ref;
          cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
          cell.font  = { color: { argb: 'FF1B5E20' }, bold: true, size: 10 };
        } else {
          cell.value = 'Not generated';
          cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE8E8' } };
          cell.font  = { color: { argb: 'FF7B1F1F' }, size: 10 };
        }
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
  const html    = buildHtml(newEntries, now.toISOString());

  const attachments = await buildTaggedSupplierAttachments(
    sessionCtx.supplierBuffers,
    sessionCtx.lastGenerations
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
