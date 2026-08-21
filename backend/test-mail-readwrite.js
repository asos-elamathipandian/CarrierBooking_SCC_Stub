'use strict';
// Quick local test for Mail.ReadWrite — patches a message (toggles isRead) then restores it.
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

async function graph(method, path, body) {
  const token = await getToken();
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    method,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    ...(body ? { body: JSON.stringify(body) } : {})
  });
  if (res.status === 204) return null;
  const text = await res.text();
  if (!res.ok) throw new Error(`Graph ${method} ${res.status}: ${text}`);
  return JSON.parse(text);
}

async function main() {
  const mailbox = process.env.EMAIL_INGEST_MAILBOX;
  if (!mailbox) { console.error('EMAIL_INGEST_MAILBOX not set'); process.exit(1); }

  console.log(`\nMailbox: ${mailbox}\n`);

  // Fetch one message to test with
  const data = await graph('GET',
    `/users/${encodeURIComponent(mailbox)}/messages?$select=id,subject,isRead&$top=1`
  );
  const msg = data?.value?.[0];
  if (!msg) { console.log('No messages in mailbox to test with.'); return; }

  console.log(`Test message: "${msg.subject || '(no subject)'}" — isRead: ${msg.isRead}`);

  // Flip isRead then restore — net zero change to the mailbox
  const flipped = !msg.isRead;
  await graph('PATCH', `/users/${encodeURIComponent(mailbox)}/messages/${msg.id}`, { isRead: flipped });
  console.log(`✓ PATCH succeeded (isRead → ${flipped})`);

  await graph('PATCH', `/users/${encodeURIComponent(mailbox)}/messages/${msg.id}`, { isRead: msg.isRead });
  console.log(`✓ Restored (isRead → ${msg.isRead})`);

  console.log('\nMail.ReadWrite test complete — permission is working.');
}

main().catch(err => { console.error('✗', err.message); process.exit(1); });
