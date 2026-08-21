'use strict';
// Quick local test for Mail.Read — lists emails + attachments without uploading to SharePoint.
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const { ClientSecretCredential } = require('@azure/identity');

const GRAPH_BASE  = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

async function getToken() {
  const cred = new ClientSecretCredential(
    process.env.SP_TENANT_ID,
    process.env.SP_CLIENT_ID,
    process.env.SP_CLIENT_SECRET
  );
  return (await cred.getToken(GRAPH_SCOPE)).token;
}

async function graph(path) {
  const token = await getToken();
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`Graph ${res.status}: ${text}`);
  return JSON.parse(text);
}

async function main() {
  const mailbox = process.argv[2] || process.env.EMAIL_INGEST_MAILBOX;
  if (!mailbox) { console.error('Usage: node test-email-ingest.js [mailbox]'); process.exit(1); }

  const lookbackDays = Number(process.env.EMAIL_LOOKBACK_DAYS) || 7;
  const since = new Date(Date.now() - lookbackDays * 86400000).toISOString();

  console.log(`\nMailbox : ${mailbox}`);
  console.log(`Scanning: last ${lookbackDays} days (since ${since})\n`);

  const filter  = encodeURIComponent(`receivedDateTime ge ${since} and hasAttachments eq true`);
  const select  = encodeURIComponent('id,subject,from,receivedDateTime');
  const orderby = encodeURIComponent('receivedDateTime desc');
  const data    = await graph(
    `/users/${encodeURIComponent(mailbox)}/messages?$filter=${filter}&$select=${select}&$top=20&$orderby=${orderby}`
  );

  const messages = data.value || [];
  console.log(`Found ${messages.length} email(s) with attachments:\n`);

  const matchTokens = String(process.env.EMAIL_ATTACHMENT_MATCH || '')
    .split(',').map(s => s.trim().toLowerCase()).filter(Boolean);

  for (const msg of messages) {
    const sender = msg.from?.emailAddress?.address || 'unknown';
    console.log(`  [${msg.receivedDateTime?.slice(0,10)}] "${msg.subject}" — from: ${sender}`);

    const attData  = await graph(
      `/users/${encodeURIComponent(mailbox)}/messages/${msg.id}/attachments`
    );
    const all = attData.value || [];
    const excel = all.filter(a => {
      if (a['@odata.type'] !== '#microsoft.graph.fileAttachment') return false;
      if (!a.contentBytes) return false;
      const name = (a.name || '').toLowerCase();
      if (!/\.(xlsx|xlsm)$/i.test(name)) return false;
      if (matchTokens.length && !matchTokens.some(t => name.includes(t))) return false;
      return true;
    });

    if (!excel.length) {
      console.log(`     No matching Excel attachments (${all.length} total attachment(s))`);
    } else {
      for (const a of excel) {
        const kb = Math.round(Buffer.from(a.contentBytes, 'base64').length / 1024);
        console.log(`     ✓ ${a.name}  (${kb} KB)`);
      }
    }
  }

  console.log('\nMail.Read test complete.');
}

main().catch(err => { console.error(err.message); process.exit(1); });
