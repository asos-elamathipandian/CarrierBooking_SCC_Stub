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

  for (const poRef of poRefs) {
    try {
      let xmlStr;
      if (local) {
        xmlStr = readLocalFeed('PO', poRef);
        if (!xmlStr) { errors.push(`[LOCAL] PO feed file not found: samples/feeds/PO_${poRef}.xml`); continue; }
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

module.exports = { fetchFeeds };
