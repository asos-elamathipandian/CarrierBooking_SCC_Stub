'use strict';
// Grants the carrier-booking-sharepoint-write SP write access to the AIMproject site.
// Uses your personal az login identity — run `az login` first if needed.
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const { AzureCliCredential } = require('@azure/identity');

const GRAPH_BASE  = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

const SP_APP_ID   = process.env.SP_CLIENT_ID;          // carrier-booking-sharepoint-write
const SP_APP_NAME = 'carrier-booking-sharepoint-write';
const SITE_URL    = process.env.SP_SITE_URL;           // https://asos1.sharepoint.com/sites/AIMproject

async function main() {
  if (!SP_APP_ID || !SITE_URL) {
    console.error('SP_CLIENT_ID and SP_SITE_URL must be set in .env'); process.exit(1);
  }

  console.log(`\nSite : ${SITE_URL}`);
  console.log(`App  : ${SP_APP_NAME} (${SP_APP_ID})\n`);

  // Use personal az login identity — must have site owner/admin rights
  const cred = new AzureCliCredential();
  const token = (await cred.getToken(GRAPH_SCOPE)).token;

  // Resolve site URL to Graph site path: hostname:/sitepath
  const match = SITE_URL.match(/^https?:\/\/([^/]+)(\/.*)?$/);
  const host  = match[1];
  const path  = (match[2] || '/').replace(/^\//, '');
  const siteGraphPath = `${host}:/${path}`;

  // Check existing permissions first
  console.log('Checking existing site permissions...');
  const existing = await fetch(`${GRAPH_BASE}/sites/${siteGraphPath}:/permissions`, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }
  });
  if (existing.ok) {
    const data = await existing.json();
    const perms = (data.value || []);
    const already = perms.find(p =>
      p.grantedToIdentities?.some(g => g.application?.id === SP_APP_ID) ||
      p.grantedTo?.application?.id === SP_APP_ID
    );
    if (already) {
      console.log(`✓ Permission already exists: roles=[${already.roles?.join(',')}]`);
      console.log('\nRun node backend/test-sharepoint.js to verify.');
      return;
    }
    console.log(`No existing permission found for this app. Proceeding to grant...\n`);
  } else {
    const txt = await existing.text();
    console.warn(`Could not list permissions (${existing.status}): ${txt}`);
    console.log('Attempting grant anyway...\n');
  }

  // POST the write grant
  const res = await fetch(`${GRAPH_BASE}/sites/${siteGraphPath}:/permissions`, {
    method:  'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      roles: ['write'],
      grantedToIdentities: [{
        application: { id: SP_APP_ID, displayName: SP_APP_NAME }
      }]
    })
  });

  const body = await res.text();
  if (res.ok) {
    const data = JSON.parse(body);
    console.log(`✓ Permission granted — id: ${data.id}, roles: [${data.roles?.join(',')}]`);
    console.log('\nNow run:  node backend/test-sharepoint.js');
  } else {
    console.error(`✗ Failed (${res.status}): ${body}`);
    if (res.status === 403) {
      console.error('\nYour account does not have site owner/admin rights on this site.');
      console.error('Ask the AIMproject site owner or a SharePoint admin to run this script.');
    }
    process.exit(1);
  }
}

main().catch(err => { console.error(err.message); process.exit(1); });
