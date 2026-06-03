'use strict';

const { BlobServiceClient } = require('@azure/storage-blob');
const poParser  = require('./po-parser');
const asnParser = require('./asn-parser');

function getContainerClient() {
  const connStr = process.env.AZURE_STORAGE_CONNECTION_STRING;
  const container = process.env.AZURE_BLOB_CONTAINER_NAME;
  if (!connStr || !container) {
    throw new Error(
      'Azure Blob Storage not configured. Set AZURE_STORAGE_CONNECTION_STRING and AZURE_BLOB_CONTAINER_NAME environment variables.'
    );
  }
  const svcClient = BlobServiceClient.fromConnectionString(connStr);
  return svcClient.getContainerClient(container);
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
 * Fetch PO and ASN feeds from Azure Blob by ref arrays.
 */
async function fetchFeeds(poRefs, asnRefs) {
  const containerClient = getContainerClient();
  const poFeeds = [];
  const asnFeeds = [];
  const errors = [];

  for (const poRef of poRefs) {
    const prefix = `ASOS_E2ASOS_PO_PO_${poRef}_`;
    try {
      const xmlStr = await findAndReadBlob(containerClient, prefix);
      if (xmlStr) {
        const parsed = await poParser.parse(xmlStr);
        poFeeds.push(parsed);
      } else {
        errors.push(`PO feed not found in blob for ref: ${poRef}`);
      }
    } catch (err) {
      errors.push(`Error fetching PO ${poRef}: ${err.message}`);
    }
  }

  for (const asnRef of asnRefs) {
    const prefix = `ASOS_E2ASOS_ASN_ASN_${asnRef}_`;
    try {
      const xmlStr = await findAndReadBlob(containerClient, prefix);
      if (xmlStr) {
        const parsed = await asnParser.parse(xmlStr);
        asnFeeds.push(parsed);
      } else {
        errors.push(`ASN feed not found in blob for ref: ${asnRef}`);
      }
    } catch (err) {
      errors.push(`Error fetching ASN ${asnRef}: ${err.message}`);
    }
  }

  return { poFeeds, asnFeeds, errors };
}

module.exports = { fetchFeeds };
