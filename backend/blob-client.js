'use strict';

const { BlobServiceClient } = require('@azure/storage-blob');
const poParser  = require('./po-parser');
const asnParser = require('./asn-parser');
const carrierAsnParser = require('./carrier-asn-parser');
const fs   = require('fs');
const path = require('path');

const LOCAL_FEEDS_DIR = path.join(__dirname, '..', 'samples', 'feeds');

function isLocalMode() {
  // Local mode only if NEITHER the legacy storage vars NOR the PO feed var are set
  return (!process.env.AZURE_STORAGE_CONNECTION_STRING || !process.env.AZURE_BLOB_CONTAINER_NAME)
    && !process.env.AZURE_PO_FEED_CONNECTION_STRING;
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

// Carrier ASN feed source: same storage account
// Container: bam036-asnin-endtoend
// Blob path: ASNInDavisTurner/{year}/{MM}/{DD}/ASOS_{date}_{time}_DavisTurner_{seq}
const ASN_CARRIER_CONTAINER = 'bam036-asnin-endtoend';
const ASN_CARRIER_FOLDER    = 'ASNInDavisTurner';

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
 * Container client for carrier ASN feed — same storage account, different container.
 * Container: bam036-asnin-endtoend
 */
function getCarrierASNContainerClient() {
  const connStr = process.env.AZURE_PO_FEED_CONNECTION_STRING;
  let svcClient;
  if (connStr.startsWith('SharedAccessSignature=') || connStr.includes(';BlobEndpoint=')) {
    const parts = parseSasConnectionString(connStr);
    const endpoint = (parts['BlobEndpoint'] || '').replace(/\/$/, '');
    const sas      = parts['SharedAccessSignature'] || '';
    svcClient = new BlobServiceClient(`${endpoint}?${sas}`);
  } else {
    svcClient = BlobServiceClient.fromConnectionString(connStr);
  }
  return svcClient.getContainerClient(ASN_CARRIER_CONTAINER);
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
 * Search for a PO blob by PO number scanning month-by-month backwards from today.
 * Blob path: aimpurchaseorder/{year}/{MM}/{DD}/ASOS_E2ASOS_PO_PO_{poRef}_{timestamp}.xml
 * Stops as soon as the first match is found.
 * Looks back up to 48 months (~4 years).
 */
async function findPOBlob(containerClient, poRef) {
  const segment = `_PO_${poRef}_`;
  const now = new Date();

  for (let offset = 0; offset < 12; offset++) {
    const d = new Date(now.getFullYear(), now.getMonth() - offset, 1);
    const year  = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const prefix = `${PO_FEED_BASE}${year}/${month}/`;

    let blobCount = 0;
    let bestBlob = null;
    for await (const blob of containerClient.listBlobsFlat({ prefix })) {
      blobCount++;
      if (!blob.name.includes(segment)) continue;
      if (!bestBlob || blob.properties.lastModified > bestBlob.properties.lastModified) {
        bestBlob = blob;
      }
    }
    console.log(`[PO feed] Scanned ${prefix} — ${blobCount} blob(s)${bestBlob ? ` — MATCH: ${bestBlob.name}` : ''}`);

    if (bestBlob) {
      // Extract YYYY-MM-DD from blob path: aimpurchaseorder/2026/05/26/filename
      const parts = bestBlob.name.split('/');
      // parts: ['aimpurchaseorder','2026','05','26','filename']
      const blobDate = (parts.length >= 4)
        ? `${parts[1]}-${parts[2]}-${parts[3]}`
        : null;

      const blobClient = containerClient.getBlobClient(bestBlob.name);
      const download = await blobClient.download(0);
      const chunks = [];
      for await (const chunk of download.readableStreamBody) {
        chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
      }
      return { xml: Buffer.concat(chunks).toString('utf8'), blobDate, filename: bestBlob.name.split('/').pop() };
    }
  }
  console.log(`[PO feed] Not found after 48-month scan: ${poRef}`);
  return null;
}

/**
 * LOCAL MODE: scan samples/feeds/ for carrier ASN files containing the PO ref.
 * Scans any XML file that is NOT a PO_* or ASN_* feed file and checks content
 * for <PurchaseOrder_ID>{poRef}</PurchaseOrder_ID>.
 */
function readLocalCarrierASNs(poRef) {
  if (!fs.existsSync(LOCAL_FEEDS_DIR)) return [];
  const needle = `<PurchaseOrder_ID>${poRef}</PurchaseOrder_ID>`;
  return fs.readdirSync(LOCAL_FEEDS_DIR)
    .filter(f => /\.xml$/i.test(f) && !/^(PO|ASN)_/i.test(f))
    .reduce((acc, f) => {
      try {
        const content = fs.readFileSync(path.join(LOCAL_FEEDS_DIR, f), 'utf8');
        if (content.includes(needle)) {
          let parsed = [];
          try { parsed = carrierAsnParser.parse(content); } catch (_) {}
          acc.push({ filename: f, xml: content, poRef, parsed });
        }
      } catch (_) { /* skip unreadable files */ }
      return acc;
    }, []);
}

/**
 * Scan ASNInDavisTurner/{year}/{MM}/{DD}/ in bam036-asnin-endtoend for blobs whose
 * content contains <PurchaseOrder_ID>{poRef}</PurchaseOrder_ID>.
 * Uses the PO's shipDate (DateTypeCd 017) to derive the date folder; falls back
 * to cancelDate (037) if shipDate is absent.
 */
async function fetchCarrierASNsForPO(containerClient, poRef, shipDate, cancelDate) {
  const rawDate = shipDate || cancelDate;
  if (!rawDate) {
    console.log(`[Carrier ASN] No date available for PO ${poRef} — skipping scan`);
    return [];
  }

  // Accept YYYY-MM-DD or YYYYMMDD
  const d = new Date(rawDate.length === 8
    ? `${rawDate.slice(0, 4)}-${rawDate.slice(4, 6)}-${rawDate.slice(6, 8)}`
    : rawDate);

  if (isNaN(d.getTime())) {
    console.log(`[Carrier ASN] Invalid date "${rawDate}" for PO ${poRef} — skipping scan`);
    return [];
  }

  const year  = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day   = String(d.getDate()).padStart(2, '0');
  const prefix = `${ASN_CARRIER_FOLDER}/${year}/${month}/${day}/`;
  const needle  = `<PurchaseOrder_ID>${poRef}</PurchaseOrder_ID>`;

  console.log(`[Carrier ASN] Scanning ${ASN_CARRIER_CONTAINER}/${prefix} for PO ${poRef}`);

  const matched = [];
  let scanned = 0;
  for await (const blob of containerClient.listBlobsFlat({ prefix })) {
    scanned++;
    const blobClientRef = containerClient.getBlobClient(blob.name);
    const download = await blobClientRef.download(0);
    const chunks = [];
    for await (const chunk of download.readableStreamBody) {
      chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
    }
    const content = Buffer.concat(chunks).toString('utf8');
    if (content.includes(needle)) {
      const filename = blob.name.split('/').pop();
      const lastModified = blob.properties.lastModified || new Date(0);
      console.log(`[Carrier ASN] MATCH: ${blob.name}`);
      let parsed = [];
      try {
        const all = await carrierAsnParser.parse(content);
        // Filter to only ASN groups belonging to the target PO
        parsed = all.filter(g => g.poId === poRef);
      } catch (e) {
        console.warn(`[Carrier ASN] Parse warning for ${filename}: ${e.message}`);
      }
      matched.push({ filename, blobPath: blob.name, xml: content, poRef, parsed, lastModified });
    }
  }
  console.log(`[Carrier ASN] Scanned ${scanned} blob(s) in ${prefix} — ${matched.length} match(es) for PO ${poRef}`);
  return matched;
}

/**
 * Fetch PO and ASN feeds — uses Azure Blob when env vars are set,
 * otherwise falls back to local files in samples/feeds/.
 * Local filenames: PO_{PONumber}.xml  and  ASN_{ASNRef}.xml
 */
async function fetchFeeds(poRefs, asnRefs) {
  const local = isLocalMode();
  // Only create the legacy container client if its env vars are actually present
  const hasLegacyBlob = !!(process.env.AZURE_STORAGE_CONNECTION_STRING && process.env.AZURE_BLOB_CONTAINER_NAME);
  const containerClient = hasLegacyBlob ? getContainerClient() : null;
  const poFeeds = [];
  const poFeedXmls = {}; // orderId -> raw XML string (for preview)
  const asnFeeds = [];
  const errors = [];

  // Reusable PO feed container client (initialised once if env var is set)
  const poFeedClient = hasPOFeedBlob() ? getPOFeedContainerClient() : null;

  // orderId -> blobDate (YYYY-MM-DD) from the matched PO blob path
  const poBlobDates    = {};
  const poFilenames    = {}; // orderId -> original blob/local filename

  for (const poRef of poRefs) {
    try {
      let xmlStr;
      let blobDate = null;
      let blobFilename = `PO_${poRef}.xml`;
      if (poFeedClient) {
        // Fetch from asbamintstgeunendtoend01 / bam033v-aimpurchaseorder-endtoend
        const result = await findPOBlob(poFeedClient, poRef);
        if (!result) {
          // Fallback: try legacy container if configured
          if (containerClient) xmlStr = await findAndReadBlob(containerClient, `ASOS_E2ASOS_PO_PO_${poRef}_`);
          if (!xmlStr) { errors.push(`PO feed not found in blob for ref: ${poRef}`); continue; }
        } else {
          xmlStr = result.xml;
          blobDate = result.blobDate;
          if (result.filename) blobFilename = result.filename;
        }
      } else if (local) {
        xmlStr = readLocalFeed('PO', poRef);
        if (!xmlStr) { errors.push(`[LOCAL] PO feed file not found: samples/feeds/PO_${poRef}.xml`); continue; }
      } else {
        xmlStr = await findAndReadBlob(containerClient, `ASOS_E2ASOS_PO_PO_${poRef}_`);
        if (!xmlStr) { errors.push(`PO feed not found in blob for ref: ${poRef}`); continue; }
      }
      const parsed = await poParser.parse(xmlStr);
      poFeeds.push(parsed);
      poFeedXmls[parsed.orderId] = xmlStr;
      if (blobDate) poBlobDates[parsed.orderId] = blobDate;
      poFilenames[parsed.orderId] = blobFilename;
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

  // ── Carrier ASN feeds (bam036-asnin-endtoend / ASNInDavisTurner) ──────────
  // Derived from PO ship dates — no separate asnRefs needed.
  const carrierAsnFiles = [];
  const carrierAsnClient = hasPOFeedBlob() ? getCarrierASNContainerClient() : null;

  for (const po of poFeeds) {
    try {
      let files;
      if (local) {
        files = readLocalCarrierASNs(po.orderId);
        if (files.length === 0) {
          errors.push(`[LOCAL] No carrier ASN files found in samples/feeds/ for PO ${po.orderId}`);
        }
      } else if (carrierAsnClient) {
        // Prefer blob path date, fall back to XML dates
        const dateToUse = poBlobDates[po.orderId] || po.shipDate || po.cancelDate;
        files = await fetchCarrierASNsForPO(carrierAsnClient, po.orderId, dateToUse, null);
        if (files.length === 0) {
          errors.push(`Carrier ASN not found in blob for PO ${po.orderId} (date: ${dateToUse || 'unknown'})`);
        }
      } else {
        errors.push(`Carrier ASN skipped for PO ${po.orderId} — AZURE_PO_FEED_CONNECTION_STRING not set`);
        continue;
      }
      carrierAsnFiles.push(...files);
    } catch (err) {
      errors.push(`Error fetching carrier ASN for PO ${po.orderId}: ${err.message}`);
    }
  }

  return { poFeeds, poFeedXmls, poFilenames, asnFeeds, carrierAsnFiles, errors, localMode: local };
}

/**
 * Scan ASNInDavisTurner date folders backwards from today (up to daysBack days)
 * looking for carrier files containing <PurchaseOrder_ID>{poRef}</PurchaseOrder_ID>.
 */
async function fetchCarrierASNsForPOByDateRange(containerClient, poRef, daysBack = 60) {
  const needle = `<PurchaseOrder_ID>${poRef}</PurchaseOrder_ID>`;
  const matched = [];
  const today = new Date();

  for (let d = 0; d < daysBack; d++) {
    const dt = new Date(today);
    dt.setDate(dt.getDate() - d);
    const year  = dt.getFullYear();
    const month = String(dt.getMonth() + 1).padStart(2, '0');
    const day   = String(dt.getDate()).padStart(2, '0');
    const prefix = `${ASN_CARRIER_FOLDER}/${year}/${month}/${day}/`;

    for await (const blob of containerClient.listBlobsFlat({ prefix })) {
      const blobRef = containerClient.getBlobClient(blob.name);
      const download = await blobRef.download(0);
      const chunks = [];
      for await (const chunk of download.readableStreamBody) {
        chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
      }
      const content = Buffer.concat(chunks).toString('utf8');
      if (content.includes(needle)) {
        const filename = blob.name.split('/').pop();
        const lastModified = blob.properties.lastModified || new Date(0);
        let parsed = [];
        try {
          const all = await carrierAsnParser.parse(content);
          // Filter to only ASN groups belonging to the target PO
          parsed = all.filter(g => g.poId === poRef);
        } catch (_) {}
        matched.push({ filename, blobPath: blob.name, xml: content, poRef, parsed, lastModified });
      }
    }
    if (matched.length > 0) break; // found files for this PO — stop scanning
  }
  return matched;
}

/**
 * Fetch carrier ASN feeds directly from blob without needing PO feeds.
 * Scans ASNInDavisTurner date folders for each PO ref.
 * Falls back to local samples/feeds/ when env vars are not set.
 */
async function fetchCarrierFeedsOnly(poRefs) {
  const local = isLocalMode() || !hasPOFeedBlob();
  const carrierAsnFiles = [];
  const errors = [];

  for (const poRef of poRefs) {
    try {
      let files;
      if (local) {
        files = readLocalCarrierASNs(poRef);
        if (files.length === 0) {
          errors.push(`[LOCAL] No carrier ASN files found in samples/feeds/ for PO ${poRef}`);
        }
      } else {
        const carrierClient = getCarrierASNContainerClient();
        files = await fetchCarrierASNsForPOByDateRange(carrierClient, poRef, 60);
        if (files.length === 0) {
          errors.push(`Carrier ASN not found in blob for PO ${poRef} (scanned last 60 days)`);
        }
      }
      carrierAsnFiles.push(...files);
    } catch (err) {
      errors.push(`Error fetching carrier ASN for PO ${poRef}: ${err.message}`);
    }
  }

  return { poFeeds: [], asnFeeds: [], carrierAsnFiles, errors, localMode: local };
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

module.exports = { fetchFeeds, fetchCarrierFeedsOnly, searchBlobs };
