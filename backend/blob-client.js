'use strict';

const { BlobServiceClient } = require('@azure/storage-blob');
const poParser  = require('./po-parser');
const asnParser = require('./asn-parser');
const fs   = require('fs');
const path = require('path');

const LOCAL_FEEDS_DIR = path.join(__dirname, '..', 'samples', 'feeds');

function isLocalMode() {
  return !process.env.AZURE_STORAGE_CONNECTION_STRING || !process.env.AZURE_BLOB_CONTAINER_NAME;
}

function getContainerClient() {
  const connStr = process.env.AZURE_STORAGE_CONNECTION_STRING;
  const container = process.env.AZURE_BLOB_CONTAINER_NAME;
  const svcClient = BlobServiceClient.fromConnectionString(connStr);
  return svcClient.getContainerClient(container);
}

// PO feed source: storage account asbamintstgeunendtoend01
// Blob path: aimpurchaseorder/{year}/{month}/{day}/ASOS_E2ASOS_PO_PO_{PONumber}_{timestamp}.xml
const PO_FEED_CONTAINER  = 'bam033v-aimpurchaseorder-endtoend';
const PO_FEED_BASE       = 'aimpurchaseorder/';

function hasPOFeedBlob() {
  return !!process.env.AZURE_PO_FEED_CONNECTION_STRING;
}

function parseSasConnectionString(connStr) {
  const parts = {};
  // Split on semicolons but only at key=value boundaries
  connStr.split(/;(?=[A-Za-z])/).forEach(p => {
    const i = p.indexOf('=');
    if (i > 0) parts[p.slice(0, i)] = p.slice(i + 1);
  });
  return parts;
}

function getPOFeedContainerClient() {
  const connStr = process.env.AZURE_PO_FEED_CONNECTION_STRING;
  let svcClient;
  if (connStr.startsWith('SharedAccessSignature=') || connStr.includes(';BlobEndpoint=')) {
    // SAS-format connection string — build URL manually
    const parts = parseSasConnectionString(connStr);
    const endpoint = (parts['BlobEndpoint'] || '').replace(/\/$/, '');
    const sas      = parts['SharedAccessSignature'] || '';
    svcClient = new BlobServiceClient(`${endpoint}?${sas}`);
  } else {
    svcClient = BlobServiceClient.fromConnectionString(connStr);
  }
  return svcClient.getContainerClient(PO_FEED_CONTAINER);
}

/**
 * LOCAL MODE: read feed XML from samples/feeds/PO_{ref}.xml or ASN_{ref}.xml
 */
function readLocalFeed(type, ref) {
  // Try exact filename first, then case-insensitive scan
  const exact = path.join(LOCAL_FEEDS_DIR, `${type}_${ref}.xml`);
  if (fs.existsSync(exact)) return fs.readFileSync(exact, 'utf8');
  if (!fs.existsSync(LOCAL_FEEDS_DIR)) return null;
  const files = fs.readdirSync(LOCAL_FEEDS_DIR);
  const match = files.find(f => f.toLowerCase() === `${type.toLowerCase()}_${ref.toLowerCase()}.xml`);
  return match ? fs.readFileSync(path.join(LOCAL_FEEDS_DIR, match), 'utf8') : null;
}

/**
 * Find blob by prefix pattern and return its content as string.
 * Blob naming: ASOS_E2ASOS_PO_PO_{PONumber}_*.xml
 */
async function findAndReadBlob(containerClient, prefix) {
  for await (const blob of containerClient.listBlobsFlat({ prefix })) {
    const blobClient = containerClient.getBlobClient(blob.name);
    const download = await blobClient.download(0);
    const chunks = [];
    for await (const chunk of download.readableStreamBody) {
      chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
    }
    return Buffer.concat(chunks).toString('utf8');
  }
  return null;
}

/**
 * Search for a PO blob by PO number across date subfolders.
 * Blob path: aimpurchaseorder/{year}/{MM}/{DD}/ASOS_E2ASOS_PO_PO_{poRef}_{timestamp}.xml
 * Scans current year first, then previous year as fallback.
 * Returns content of the most recently modified matching blob.
 */
async function findPOBlob(containerClient, poRef) {
  const years = [new Date().getFullYear(), new Date().getFullYear() - 1];
  const segment = `_PO_${poRef}_`;

  for (const year of years) {
    const prefix = `${PO_FEED_BASE}${year}/`;
    let bestBlob = null;

    for await (const blob of containerClient.listBlobsFlat({ prefix })) {
      if (!blob.name.includes(segment)) continue;
      if (!bestBlob || blob.properties.lastModified > bestBlob.properties.lastModified) {
        bestBlob = blob;
      }
    }

    if (bestBlob) {
      const blobClient = containerClient.getBlobClient(bestBlob.name);
      const download = await blobClient.download(0);
      const chunks = [];
      for await (const chunk of download.readableStreamBody) {
        chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
      }
      console.log(`[PO feed] Found: ${bestBlob.name}`);
      return Buffer.concat(chunks).toString('utf8');
    }
  }
  return null;
}

/**
 * Fetch PO and ASN feeds — uses Azure Blob when env vars are set,
 * otherwise falls back to local files in samples/feeds/.
 * Local filenames: PO_{PONumber}.xml  and  ASN_{ASNRef}.xml
 */
async function fetchFeeds(poRefs, asnRefs) {
  const local = isLocalMode();
  const containerClient = local ? null : getContainerClient();
  const poFeeds = [];
  const asnFeeds = [];
  const errors = [];

  // Reusable PO feed container client (initialised once if env var is set)
  const poFeedClient = (!local && hasPOFeedBlob()) ? getPOFeedContainerClient() : null;

  for (const poRef of poRefs) {
    try {
      let xmlStr;
      if (local) {
        xmlStr = readLocalFeed('PO', poRef);
        if (!xmlStr) { errors.push(`[LOCAL] PO feed file not found: samples/feeds/PO_${poRef}.xml`); continue; }
      } else if (poFeedClient) {
        // Primary: fetch from asbamintstgeunendtoend01 / bam033v-aimpurchaseorder-endtoend
        // Searches aimpurchaseorder/{year}/{MM}/{DD}/ASOS_E2ASOS_PO_PO_{poRef}_{ts}.xml
        xmlStr = await findPOBlob(poFeedClient, poRef);
        if (!xmlStr) {
          // Fallback: try legacy container if configured
          if (containerClient) {
            xmlStr = await findAndReadBlob(containerClient, `ASOS_E2ASOS_PO_PO_${poRef}_`);
          }
          if (!xmlStr) { errors.push(`PO feed not found in blob for ref: ${poRef}`); continue; }
        }
      } else {
        xmlStr = await findAndReadBlob(containerClient, `ASOS_E2ASOS_PO_PO_${poRef}_`);
        if (!xmlStr) { errors.push(`PO feed not found in blob for ref: ${poRef}`); continue; }
      }
      poFeeds.push(await poParser.parse(xmlStr));
    } catch (err) {
      errors.push(`Error fetching PO ${poRef}: ${err.message}`);
    }
  }

  for (const asnRef of asnRefs) {
    try {
      let xmlStr;
      if (local) {
        xmlStr = readLocalFeed('ASN', asnRef);
        if (!xmlStr) { errors.push(`[LOCAL] ASN feed file not found: samples/feeds/ASN_${asnRef}.xml`); continue; }
      } else {
        xmlStr = await findAndReadBlob(containerClient, `ASOS_E2ASOS_ASN_ASN_${asnRef}_`);
        if (!xmlStr) { errors.push(`ASN feed not found in blob for ref: ${asnRef}`); continue; }
      }
      asnFeeds.push(await asnParser.parse(xmlStr));
    } catch (err) {
      errors.push(`Error fetching ASN ${asnRef}: ${err.message}`);
    }
  }

  return { poFeeds, asnFeeds, errors, localMode: local };
}

/**
 * Search blobs by name prefix (or substring in local mode).
 * Returns up to 50 results with name, size, lastModified.
 */
async function searchBlobs(query) {
  if (isLocalMode()) {
    if (!fs.existsSync(LOCAL_FEEDS_DIR)) return [];
    return fs.readdirSync(LOCAL_FEEDS_DIR)
      .filter(f => !query || f.toLowerCase().includes(query.toLowerCase()))
      .map(f => {
        const stat = fs.statSync(path.join(LOCAL_FEEDS_DIR, f));
        return { name: f, size: stat.size, lastModified: stat.mtime };
      });
  }

  const containerClient = getContainerClient();
  const results = [];
  for await (const blob of containerClient.listBlobsFlat({ prefix: query })) {
    results.push({
      name: blob.name,
      size: blob.properties.contentLength,
      lastModified: blob.properties.lastModified
    });
    if (results.length >= 50) break;
  }
  return results;
}

module.exports = { fetchFeeds, searchBlobs };
