'use strict';

/**
 * sharepoint-client.js
 * Read supplier template Excel files from a SharePoint document library
 * using Microsoft Graph API with an App Registration (client credentials).
 *
 * Required .env vars:
 *   SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET
 *   SP_SITE_URL      — e.g. https://asos.sharepoint.com/sites/YourSite
 *   SP_FOLDER_PATH   — folder path in the default drive, e.g. /Carrier Booking/Templates
 */

const { ClientSecretCredential } = require('@azure/identity');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

// ── Auth ──────────────────────────────────────────────────────────────────────

function isConfigured() {
  const { SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET, SP_SITE_URL } = process.env;
  return !!(SP_TENANT_ID && SP_CLIENT_ID && SP_CLIENT_SECRET && SP_SITE_URL)
    && !SP_TENANT_ID.startsWith('REPLACE')
    && !SP_CLIENT_ID.startsWith('REPLACE');
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

async function graphGet(path) {
  const token = await getAccessToken();
  const url = path.startsWith('http') ? path : `${GRAPH_BASE}${path}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API ${res.status}: ${body}`);
  }
  return res.json();
}

async function graphGetBytes(path) {
  const token = await getAccessToken();
  const url = path.startsWith('http') ? path : `${GRAPH_BASE}${path}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API ${res.status}: ${body}`);
  }
  return Buffer.from(await res.arrayBuffer());
}

// ── Site resolution ───────────────────────────────────────────────────────────

let _siteId = null;
async function getSiteId() {
  if (_siteId) return _siteId;
  const siteUrl = process.env.SP_SITE_URL || '';
  // e.g. https://asos.sharepoint.com/sites/YourSite
  const match = siteUrl.match(/^https?:\/\/([^/]+)(\/.*)?$/);
  if (!match) throw new Error(`Invalid SP_SITE_URL: ${siteUrl}`);
  const hostname = match[1];
  const sitePath = (match[2] || '/').replace(/^\//, '');
  const data = await graphGet(`/sites/${hostname}:/${sitePath}`);
  _siteId = data.id;
  return _siteId;
}

// ── Drive resolution ──────────────────────────────────────────────────────────

let _driveId = null;
async function getDriveId() {
  if (_driveId) return _driveId;
  const siteId = await getSiteId();
  // Use the default document library (first drive)
  const data = await graphGet(`/sites/${siteId}/drives`);
  const defaultDrive = (data.value || [])[0];
  if (!defaultDrive) throw new Error('No drives found on site');
  _driveId = defaultDrive.id;
  return _driveId;
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * List Excel (.xlsx / .xlsm) files in SP_FOLDER_PATH.
 * Returns array of { id, name, size, lastModifiedDateTime, webUrl }
 */
async function listTemplateFiles() {
  const siteId  = await getSiteId();
  const driveId = await getDriveId();
  const folder  = (process.env.SP_FOLDER_PATH || '/').replace(/\/$/, '') || '/';

  let endpoint;
  if (folder === '/') {
    endpoint = `/sites/${siteId}/drives/${driveId}/root/children`;
  } else {
    endpoint = `/sites/${siteId}/drives/${driveId}/root:${folder}:/children`;
  }

  const data = await graphGet(endpoint);
  return (data.value || [])
    .filter(f => !f.folder && /\.(xlsx|xlsm)$/i.test(f.name))
    .map(f => ({
      id:                   f.id,
      name:                 f.name,
      size:                 f.size,
      lastModifiedDateTime: f.lastModifiedDateTime,
      webUrl:               f.webUrl
    }));
}

/**
 * Download a file by its Graph item ID and return a Buffer.
 */
async function downloadFile(itemId) {
  const siteId  = await getSiteId();
  const driveId = await getDriveId();
  return graphGetBytes(`/sites/${siteId}/drives/${driveId}/items/${itemId}/content`);
}

module.exports = { isConfigured, listTemplateFiles, downloadFile };
