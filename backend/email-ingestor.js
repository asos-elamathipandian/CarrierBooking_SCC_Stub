'use strict';

/**
 * email-ingestor.js
 *
 * Reads unread emails from a dedicated ASOS mailbox and deposits any Excel
 * attachments (.xlsx / .xlsm) into the correct SharePoint supplier subfolder
 * so the SP scheduler can pick them up at its 9 AM & 1 PM scheduled runs.
 *
 * Uses the SAME App Registration as sharepoint-client.js.
 * The registration must have the Mail.ReadWrite *application* permission
 * admin-consented for the target mailbox (no extra credentials needed).
 *
 * Required .env vars:
 *   EMAIL_INGEST_MAILBOX  — e.g. carrier-templates@asos.com
 *   SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET  (shared with SharePoint)
 *
 * Optional .env vars:
 *   EMAIL_SUPPLIER_MAP      — JSON: { "domain.com": "SP_FOLDER", "email@x.com": "SP_FOLDER2" }
 *                             Exact email address is checked first, then sender domain.
 *                             Unrecognised senders get a folder derived from their display name.
 *   EMAIL_PROCESSED_FOLDER  — mailbox subfolder name to move processed emails into
 *                             (e.g. "Processed"). If not set, emails are only marked as read.
 *
 * Note: Graph API returns attachment content inline (base64) for files up to 3 MB.
 * Standard supplier Excel templates are well within this limit.
 */

const { ClientSecretCredential } = require('@azure/identity');
const sp = require('./sharepoint-client');

const GRAPH_BASE  = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

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
 * List unread messages that have at least one attachment.
 */
async function listUnreadMessages(mailbox) {
  const filter = encodeURIComponent("isRead eq false and hasAttachments eq true");
  const select = encodeURIComponent("id,subject,from,receivedDateTime,hasAttachments");
  const path   = `/users/${encodeURIComponent(mailbox)}/messages?$filter=${filter}&$select=${select}&$top=50&$orderby=${encodeURIComponent('receivedDateTime asc')}`;
  const data   = await graphRequest('GET', path);
  return (data && Array.isArray(data.value)) ? data.value : [];
}

/**
 * Fetch file attachments for a message and return only .xlsx / .xlsm ones.
 * Graph API returns contentBytes (base64) inline for attachments ≤ 3 MB.
 */
async function getExcelAttachments(mailbox, messageId) {
  const path = `/users/${encodeURIComponent(mailbox)}/messages/${messageId}/attachments`;
  const data = await graphRequest('GET', path);
  const all  = (data && Array.isArray(data.value)) ? data.value : [];
  return all.filter(a =>
    a['@odata.type'] === '#microsoft.graph.fileAttachment' &&
    /\.(xlsx|xlsm)$/i.test(a.name || '') &&
    a.contentBytes   // must have inline content
  );
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
 * For each unread email with an Excel attachment:
 *   1. Resolves the supplier folder from the sender address / EMAIL_SUPPLIER_MAP
 *   2. Uploads the attachment to SharePoint under SP_FOLDER_PATH/{supplierFolder}/
 *   3. Marks the email as read (and optionally moves it to EMAIL_PROCESSED_FOLDER)
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

  const mailbox = process.env.EMAIL_INGEST_MAILBOX;
  console.log(`[Email Ingestor] Checking ${mailbox} for new supplier templates…`);

  let messages;
  try {
    messages = await listUnreadMessages(mailbox);
  } catch (err) {
    return { processed: 0, uploaded: 0, errors: [`Failed to list messages: ${err.message}`] };
  }

  if (!messages.length) {
    console.log('[Email Ingestor] No unread emails with attachments.');
    return { processed: 0, uploaded: 0, errors: [] };
  }

  console.log(`[Email Ingestor] Found ${messages.length} unread email(s) to process.`);

  let uploaded = 0;
  const errors = [];

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
      // Has attachments but none are Excel — mark as read to avoid reprocessing
      console.log(`[Email Ingestor] "${subject}" from ${senderAddress} — no Excel attachments, skipping.`);
      await markAsRead(mailbox, msg.id).catch(e =>
        console.warn(`[Email Ingestor] Could not mark as read: ${e.message}`)
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

    // Mark as read once at least one attachment was successfully uploaded
    if (anyUploaded) {
      try {
        await markAsRead(mailbox, msg.id);
        const processedFolder = process.env.EMAIL_PROCESSED_FOLDER;
        if (processedFolder) {
          await moveToFolder(mailbox, msg.id, processedFolder);
        }
      } catch (err) {
        console.warn(`[Email Ingestor] Could not mark message as read/move: ${err.message}`);
      }
    }
  }

  console.log(`[Email Ingestor] Done — ${uploaded} file(s) uploaded to SharePoint, ${errors.length} error(s).`);
  return { processed: messages.length, uploaded, errors };
}

module.exports = { isConfigured, ingest };
