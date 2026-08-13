'use strict';

/**
 * blob-webhook-client.js
 *
 * Bridge for supplier Excel files pushed by Power Automate into a
 * dedicated Azure Storage container, used because the Graph
 * Sites.Read.All permission (for pulling directly from SharePoint) and
 * Power Automate's premium HTTP connector are both blocked in this tenant.
 *
 * Flow: SharePoint file created → Power Automate → Azure Blob Storage
 * "Create blob" (Standard connector, no premium/DLP conflict) → this
 * module lists/downloads new blobs → parsed into sessionState.
 *
 * Required .env vars:
 *   AZURE_WEBHOOK_STORAGE_CONNECTION_STRING — connection string for the
 *     storage account (Account Key, from Azure Portal → Access keys)
 *   AZURE_WEBHOOK_CONTAINER_NAME            — container name, e.g. "supplier-uploads"
 */

const { BlobServiceClient } = require('@azure/storage-blob');

function isConfigured() {
  const { AZURE_WEBHOOK_STORAGE_CONNECTION_STRING, AZURE_WEBHOOK_CONTAINER_NAME } = process.env;
  return !!(AZURE_WEBHOOK_STORAGE_CONNECTION_STRING && AZURE_WEBHOOK_CONTAINER_NAME);
}

function getContainerClient() {
  const connStr   = process.env.AZURE_WEBHOOK_STORAGE_CONNECTION_STRING;
  const container = process.env.AZURE_WEBHOOK_CONTAINER_NAME;
  const svcClient = BlobServiceClient.fromConnectionString(connStr);
  return svcClient.getContainerClient(container);
}

/**
 * List Excel blobs in the container.
 * Returns [{ name, lastModified, size }]
 */
async function listBlobs() {
  const containerClient = getContainerClient();
  const results = [];
  for await (const blob of containerClient.listBlobsFlat()) {
    if (/\.(xlsx|xlsm)$/i.test(blob.name)) {
      results.push({
        name:         blob.name,
        lastModified: blob.properties.lastModified,
        size:         blob.properties.contentLength
      });
    }
  }
  return results;
}

/**
 * Download a blob's content as a Buffer.
 */
async function downloadBlob(name) {
  const containerClient = getContainerClient();
  const blobClient = containerClient.getBlobClient(name);
  const download = await blobClient.download();
  const chunks = [];
  for await (const chunk of download.readableStreamBody) {
    chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
  }
  return Buffer.concat(chunks);
}

/**
 * Delete a blob after it has been successfully ingested.
 */
async function deleteBlob(name) {
  const containerClient = getContainerClient();
  await containerClient.getBlobClient(name).deleteIfExists();
}

module.exports = { isConfigured, listBlobs, downloadBlob, deleteBlob };
