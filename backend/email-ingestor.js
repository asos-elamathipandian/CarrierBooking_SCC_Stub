'use strict';

/**
 * email-ingestor.js
 *
 * OPTIONAL INGESTION PATH — an alternative to the Power Automate flow
 * (Outlook "When a new email arrives" → SharePoint "Create file").
 * Only active if EMAIL_INGEST_MAILBOX is set; leave that env var unset to
 * keep Power Automate as the sole ingestion path.
 *
 * Reads emails from a dedicated ASOS mailbox and deposits any Excel
 * attachments (.xlsx / .xlsm) into the correct SharePoint supplier subfolder
 * so the SP scheduler can pick them up at its 9 AM & 1 PM scheduled runs.
 *
 * Uses the SAME App Registration as sharepoint-client.js.
 *
 * PERMISSION MODES
 *   Read-only (default) — needs only the Mail.Read *application* permission.
 *     Messages are never modified. De-duplication is handled locally via a
 *     receivedDateTime watermark plus a set of processed message IDs stored
 *     in bible/email-ingest-status.json.
 *   Read-write (EMAIL_READONLY_MODE=false) — needs Mail.ReadWrite. Uses the
 *     mailbox itself as the queue: unread messages are processed, then marked
 *     as read and optionally moved to EMAIL_PROCESSED_FOLDER.
 *
 * Required .env vars:
 *   EMAIL_INGEST_MAILBOX  — InboundService@asos.com
 *   SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET  (shared with SharePoint)
 *
 * Optional .env vars:
 *   EMAIL_READONLY_MODE     — "false" to use the legacy Mail.ReadWrite behaviour.
 *                             Defaults to true (Mail.Read only).
 *   EMAIL_LOOKBACK_DAYS     — how far back to scan on the very first run,
 *                             before a watermark exists. Default 7.
 *   EMAIL_SUPPLIER_MAP      — JSON: { "domain.com": "SP_FOLDER", "email@x.com": "SP_FOLDER2" }
 *                             Exact email address is checked first, then sender domain.
 *                             Unrecognised senders get a folder derived from their display name.
 *   EMAIL_PROCESSED_FOLDER  — mailbox subfolder name to move processed emails into
 *                             (e.g. "Processed"). Read-write mode only; ignored
 *                             in read-only mode.
 *
 * Note: Graph API returns attachment content inline (base64) for files up to 3 MB.
 * Standard supplier Excel templates are well within this limit.
 */

const path = require('path');
const fs   = require('fs');
const { ClientSecretCredential } = require('@azure/identity');
const sp = require('./sharepoint-client');

const GRAPH_BASE  = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

const STATE_FILE      = path.join(__dirname, '..', 'bible', 'email-ingest-status.json');
const MAX_TRACKED_IDS = 1000;

// ── Auth ──────────────────────────────────────────────────────────────────────

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

async function getToken() {
  const token = await getCredential().getToken(GRAPH_SCOPE);
  return token.token;
}

// ── Generic Graph request ─────────────────────────────────────────────────────

async function graphRequest(method, path, body) {
  const token = await getToken();
  const url   = path.startsWith('http') ? path : `${GRAPH_BASE}${path}`;
  const opts  = {
    method,
    headers: {
      Authorization:  `Bearer ${token}`,
      Accept:         'application/json',
      'Content-Type': 'application/json'
    }
  };
  if (body !== undefined) opts.body = JSON.stringify(body);

  const res = await fetch(url, opts);

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph ${method} ${path} → ${res.status}: ${text}`);
  }
  // 204 No Content (e.g. PATCH mark-as-read) returns no body
  if (res.status === 204) return null;
  const ct = res.headers.get('content-type') || '';
  if (!ct.includes('application/json')) return null;
  return res.json();
}

// ── Config check ──────────────────────────────────────────────────────────────

function isConfigured() {
  const { EMAIL_INGEST_MAILBOX, SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET } = process.env;
  return !!(EMAIL_INGEST_MAILBOX && SP_TENANT_ID && SP_CLIENT_ID && SP_CLIENT_SECRET)
    && !SP_CLIENT_ID.startsWith('REPLACE')
    && !SP_TENANT_ID.startsWith('REPLACE');
}

function isReadOnly() {
  return String(process.env.EMAIL_READONLY_MODE || 'true').toLowerCase() !== 'false';
}

// ── Ingest state (read-only mode de-duplication) ──────────────────────────────

function readState() {
  try {
    if (fs.existsSync(STATE_FILE)) {
      const s = JSON.parse(fs.readFileSync(STATE_FILE, 'utf8'));
      return {
        watermark:    s.watermark || null,
        processedIds: Array.isArray(s.processedIds) ? s.processedIds : []
      };
    }
  } catch (_) {}
  return { watermark: null, processedIds: [] };
}

function writeState(state) {
  try {
    fs.mkdirSync(path.dirname(STATE_FILE), { recursive: true });
    fs.writeFileSync(STATE_FILE, JSON.stringify({
      watermark:    state.watermark,
      processedIds: state.processedIds.slice(-MAX_TRACKED_IDS),
      lastRun:      new Date().toISOString()
    }, null, 2));
  } catch (err) {
    console.warn(`[Email Ingestor] Could not persist ingest state: ${err.message}`);
  }
}

/**
 * Start of the scan window: the stored watermark, or EMAIL_LOOKBACK_DAYS ago
 * on first run. The watermark is inclusive (ge) so a message received in the
 * same second as the previous run is not missed; processedIds prevents the
 * resulting overlap from being ingested twice.
 */
function scanFrom(state) {
  if (state.watermark) return state.watermark;
  const days = Number(process.env.EMAIL_LOOKBACK_DAYS) || 7;
  return new Date(Date.now() - days * 86400000).toISOString();
}

// ── Supplier folder resolution ────────────────────────────────────────────────

/**
 * Map a sender email address to a SharePoint supplier folder name.
 *
 * Resolution order:
 *   1. Exact sender email address in EMAIL_SUPPLIER_MAP
 *   2. Sender email domain in EMAIL_SUPPLIER_MAP
 *   3. Fallback: sanitised sender display name (or email username)
 */
function resolveSupplierFolder(senderAddress, senderName) {
  let map = {};
  try { map = JSON.parse(process.env.EMAIL_SUPPLIER_MAP || '{}'); } catch (_) {}

  const addr   = (senderAddress || '').toLowerCase().trim();
  const domain = addr.split('@')[1] || '';

  if (map[addr])   return map[addr];
  if (map[domain]) return map[domain];

  // Fallback: use display name or email username, sanitised to safe folder chars
  const raw = senderName || addr.split('@')[0] || 'UNKNOWN';
  return raw.replace(/[^A-Za-z0-9_\-]/g, '_').toUpperCase().slice(0, 50);
}

// ── Mailbox helpers ───────────────────────────────────────────────────────────

/**
 * List candidate messages that have at least one attachment.
 *
 * Read-only mode selects by receivedDateTime watermark (no mailbox writes);
 * read-write mode uses the unread flag as the queue.
 */
async function listCandidateMessages(mailbox, state) {
  const criteria = isReadOnly()
    ? `receivedDateTime ge ${scanFrom(state)} and hasAttachments eq true`
    : 'isRead eq false and hasAttachments eq true';
  const filter = encodeURIComponent(criteria);
  const select = encodeURIComponent("id,subject,from,receivedDateTime,hasAttachments");
  const path   = `/users/${encodeURIComponent(mailbox)}/messages?$filter=${filter}&$select=${select}&$top=50&$orderby=${encodeURIComponent('receivedDateTime asc')}`;
  const data   = await graphRequest('GET', path);
  return (data && Array.isArray(data.value)) ? data.value : [];
}

/**
 * Fetch file attachments for a message and return only .xlsx / .xlsm ones.
 * If EMAIL_ATTACHMENT_MATCH is set, only attachments whose filename contains
 * at least one configured token (case-insensitive) are accepted.
 *
 * EMAIL_ATTACHMENT_MATCH supports either:
 *   - a single token: "Supplier PO sheet"
 *   - a comma-separated list: "Supplier PO sheet,IDEATEKS IRSALIYE"
 * Graph API returns contentBytes (base64) inline for attachments ≤ 3 MB.
 */
async function getExcelAttachments(mailbox, messageId) {
  const path = `/users/${encodeURIComponent(mailbox)}/messages/${messageId}/attachments`;
  const data = await graphRequest('GET', path);
  const all  = (data && Array.isArray(data.value)) ? data.value : [];
  const matchTokens = String(process.env.EMAIL_ATTACHMENT_MATCH || '')
    .split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean);
  return all.filter(a => {
    if (a['@odata.type'] !== '#microsoft.graph.fileAttachment') return false;
    if (!a.contentBytes) return false;
    const name = (a.name || '').toLowerCase();
    if (!/\.(xlsx|xlsm)$/i.test(name)) return false;
    if (matchTokens.length && !matchTokens.some(token => name.includes(token))) return false;
    return true;
  });
}

/**
 * Mark a message as read.
 */
async function markAsRead(mailbox, messageId) {
  await graphRequest('PATCH',
    `/users/${encodeURIComponent(mailbox)}/messages/${messageId}`,
    { isRead: true }
  );
}

/**
 * Move a message into a named subfolder (creates the folder if it doesn't exist).
 */
async function moveToFolder(mailbox, messageId, folderName) {
  const encodedMailbox = encodeURIComponent(mailbox);

  // Find or create the target mail folder
  const filterQ = encodeURIComponent(`displayName eq '${folderName}'`);
  const existing = await graphRequest('GET',
    `/users/${encodedMailbox}/mailFolders?$filter=${filterQ}&$top=5`
  );

  let folderId;
  if (existing && existing.value && existing.value.length > 0) {
    folderId = existing.value[0].id;
  } else {
    const created = await graphRequest('POST',
      `/users/${encodedMailbox}/mailFolders`,
      { displayName: folderName }
    );
    folderId = created.id;
  }

  await graphRequest('POST',
    `/users/${encodedMailbox}/messages/${messageId}/move`,
    { destinationId: folderId }
  );
}

// ── Main entry point ──────────────────────────────────────────────────────────

/**
 * Ingest supplier template emails from the dedicated ASOS mailbox.
 *
 * For each candidate email with an Excel attachment:
 *   1. Resolves the supplier folder from the sender address / EMAIL_SUPPLIER_MAP
 *   2. Uploads the attachment to SharePoint under SP_FOLDER_PATH/{supplierFolder}/
 *   3. Records it as processed — locally in read-only mode, or by marking the
 *      email as read (and optionally moving it) in read-write mode
 *
 * Returns { processed: number, uploaded: number, errors: string[] }
 */
async function ingest() {
  if (!isConfigured()) {
    return { processed: 0, uploaded: 0, errors: ['EMAIL_INGEST_MAILBOX or auth env vars not configured'] };
  }
  if (!sp.isConfigured()) {
    return { processed: 0, uploaded: 0, errors: ['SharePoint not configured — cannot upload attachments'] };
  }

  const mailbox  = process.env.EMAIL_INGEST_MAILBOX;
  const readOnly = isReadOnly();
  const state    = readState();
  const seen     = new Set(state.processedIds);
  console.log(`[Email Ingestor] Checking ${mailbox} for new supplier templates… (${readOnly ? 'read-only' : 'read-write'} mode)`);

  let messages;
  try {
    messages = await listCandidateMessages(mailbox, state);
  } catch (err) {
    return { processed: 0, uploaded: 0, errors: [`Failed to list messages: ${err.message}`] };
  }

  if (readOnly) messages = messages.filter(m => !seen.has(m.id));

  if (!messages.length) {
    console.log('[Email Ingestor] No new emails with attachments.');
    return { processed: 0, uploaded: 0, errors: [] };
  }

  console.log(`[Email Ingestor] Found ${messages.length} email(s) to process.`);

  let uploaded = 0;
  const errors = [];

  // Marks a message done: locally in read-only mode, in the mailbox otherwise.
  const markDone = async (msg) => {
    if (readOnly) {
      state.processedIds.push(msg.id);
      seen.add(msg.id);
      if (msg.receivedDateTime && (!state.watermark || msg.receivedDateTime > state.watermark)) {
        state.watermark = msg.receivedDateTime;
      }
      return;
    }
    await markAsRead(mailbox, msg.id);
    const processedFolder = process.env.EMAIL_PROCESSED_FOLDER;
    if (processedFolder) await moveToFolder(mailbox, msg.id, processedFolder);
  };

  for (const msg of messages) {
    const senderAddress  = msg.from?.emailAddress?.address || '';
    const senderName     = msg.from?.emailAddress?.name    || '';
    const subject        = msg.subject || '(no subject)';
    const supplierFolder = resolveSupplierFolder(senderAddress, senderName);

    let attachments;
    try {
      attachments = await getExcelAttachments(mailbox, msg.id);
    } catch (err) {
      errors.push(`[${senderAddress}] Failed to fetch attachments for "${subject}": ${err.message}`);
      continue;
    }

    if (!attachments.length) {
      // Has attachments but none are Excel — record it so it isn't reprocessed
      console.log(`[Email Ingestor] "${subject}" from ${senderAddress} — no Excel attachments, skipping.`);
      await markDone(msg).catch(e =>
        console.warn(`[Email Ingestor] Could not mark as processed: ${e.message}`)
      );
      continue;
    }

    let anyUploaded = false;
    for (const att of attachments) {
      try {
        const buffer = Buffer.from(att.contentBytes, 'base64');
        await sp.uploadToSupplierFolder(supplierFolder, att.name, buffer);
        console.log(`[Email Ingestor] Uploaded "${att.name}" → SharePoint/${supplierFolder}/ (from: ${senderAddress})`);
        uploaded++;
        anyUploaded = true;
      } catch (err) {
        const detail = `[${senderAddress}] Failed to upload "${att.name}": ${err.message}`;
        console.error(`[Email Ingestor] ${detail}`);
        errors.push(detail);
      }
    }

    // Only record as processed once at least one attachment landed in SharePoint
    if (anyUploaded) {
      try {
        await markDone(msg);
      } catch (err) {
        console.warn(`[Email Ingestor] Could not mark message as processed: ${err.message}`);
      }
    }
  }

  if (readOnly) writeState(state);

  console.log(`[Email Ingestor] Done — ${uploaded} file(s) uploaded to SharePoint, ${errors.length} error(s).`);
  return { processed: messages.length, uploaded, errors };
}

module.exports = { isConfigured, isReadOnly, ingest };
