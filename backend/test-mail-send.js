'use strict';
// Quick local test for Mail.Send — sends a test email via Graph API using the SP.
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

async function main() {
  const from = process.env.REPORT_FROM || process.env.EMAIL_INGEST_MAILBOX;
  const to   = process.argv[2] || process.env.REPORT_TO?.split(',')[0]?.trim();

  if (!from) { console.error('Set REPORT_FROM or EMAIL_INGEST_MAILBOX in .env'); process.exit(1); }
  if (!to)   { console.error('Usage: node test-mail-send.js [recipient]'); process.exit(1); }

  console.log(`\nSending test email...`);
  console.log(`  From : ${from}`);
  console.log(`  To   : ${to}\n`);

  const token = await getToken();
  const res = await fetch(`${GRAPH_BASE}/users/${encodeURIComponent(from)}/sendMail`, {
    method:  'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: {
        subject: '[Test] Carrier Booking Tool — Mail.Send test',
        body: {
          contentType: 'HTML',
          content: `<p>This is a test email sent via the Graph API (Mail.Send) from the <strong>carrier-booking-sharepoint-write</strong> service principal.</p>
                    <p>Sent at: ${new Date().toLocaleString('en-GB')}</p>
                    <p>If you received this, Mail.Send is working correctly.</p>`
        },
        toRecipients: [{ emailAddress: { address: to } }]
      },
      saveToSentItems: false
    })
  });

  if (res.status === 202) {
    console.log('✓ Mail.Send success — email sent (202 Accepted)');
  } else {
    const text = await res.text();
    console.error(`✗ Graph ${res.status}: ${text}`);
    process.exit(1);
  }
}

main().catch(err => { console.error(err.message); process.exit(1); });
