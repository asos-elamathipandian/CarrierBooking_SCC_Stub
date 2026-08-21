'use strict';
/**
 * test-sharepoint.js — quick connectivity test for SharePoint / Graph API.
 * Run:  node backend/test-sharepoint.js
 *
 * What it does:
 *   1. Obtains an access token using client credentials
 *   2. Resolves the SharePoint site ID
 *   3. Lists Excel files in SP_FOLDER_PATH
 *   4. Downloads the first file found and reports its size
 */

require('dotenv').config();
const sp = require('./sharepoint-client');

async function main() {
  if (!sp.isConfigured()) {
    console.error('❌ SharePoint not configured — fill in SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET, SP_SITE_URL in .env');
    process.exit(1);
  }

  console.log('🔑 Fetching access token…');
  // listTemplateFiles triggers token + site + drive resolution internally
  console.log(`📂 Listing Excel files in SP_FOLDER_PATH="${process.env.SP_FOLDER_PATH || '/'}"…`);

  let files;
  try {
    files = await sp.listTemplateFiles();
  } catch (err) {
    console.error('❌ Failed to list files:', err.message);
    process.exit(1);
  }

  if (!files.length) {
    console.warn('⚠️  No .xlsx / .xlsm files found in the configured folder.');
    process.exit(0);
  }

  console.log(`✅ Found ${files.length} file(s):`);
  for (const f of files) {
    const kb = Math.round((f.size || 0) / 1024);
    const modified = f.lastModifiedDateTime
      ? new Date(f.lastModifiedDateTime).toLocaleString('en-GB')
      : '?';
    console.log(`   • ${f.name}  (${kb} KB, modified ${modified})`);
  }

  // Try downloading the first file
  const first = files[0];
  console.log(`\n⬇  Downloading "${first.name}" (id: ${first.id})…`);
  try {
    const buf = await sp.downloadFile(first.id);
    console.log(`✅ Download OK — ${buf.length} bytes received.`);
  } catch (err) {
    console.error('❌ Download failed:', err.message);
    process.exit(1);
  }

  console.log('\n🎉 SharePoint read test passed!');

  // ── WRITE test (upload a tiny file then delete it) ──────────────────────────
  const testFileName = `_write-test-${Date.now()}.txt`;
  const testContent  = Buffer.from(`SP write test — ${new Date().toISOString()}`);
  console.log(`\n⬆  Write test — uploading "${testFileName}"…`);

  let uploadedItem;
  try {
    uploadedItem = await sp.uploadToSupplierFolder('', testFileName, testContent);
    console.log(`✅ Upload OK — id: ${uploadedItem.id}`);
  } catch (err) {
    console.error('❌ Upload failed (Sites.Selected write grant not in place):', err.message);
    process.exit(1);
  }

  // Delete the test file to leave SharePoint clean
  const driveId = uploadedItem.parentReference?.driveId;
  if (driveId) {
    const { ClientSecretCredential } = require('@azure/identity');
    const cred  = new ClientSecretCredential(process.env.SP_TENANT_ID, process.env.SP_CLIENT_ID, process.env.SP_CLIENT_SECRET);
    const token = (await cred.getToken('https://graph.microsoft.com/.default')).token;
    const delRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${uploadedItem.id}`, {
      method: 'DELETE', headers: { Authorization: `Bearer ${token}` }
    });
    if (delRes.status === 204) console.log('🗑  Test file deleted — SharePoint left clean.');
    else console.warn(`⚠  Could not delete test file (status ${delRes.status}) — remove manually from SharePoint.`);
  }

  console.log('\n🎉 SharePoint read + write test passed!');
}

main();
