'use strict';

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const supplierReader = require('./supplier-reader');
const blobClient = require('./blob-client');
const bibleBuilder = require('./bible-builder');
const vbkreqBuilder = require('./vbkreq-builder');
const sftpUploader = require('./sftp-uploader');
const poParser  = require('./po-parser');
const asnParser = require('./asn-parser');

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'frontend')));

// Serve generated bible downloads
const bibleDir = path.join(__dirname, '..', 'bible');
if (!fs.existsSync(bibleDir)) fs.mkdirSync(bibleDir, { recursive: true });
app.use('/bible', express.static(bibleDir));

// Multer: store uploads in memory (max 10MB) — Excel files
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['.xlsx', '.xls'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(ext)) cb(null, true);
    else cb(new Error('Only Excel files (.xlsx, .xls) are accepted'));
  }
});

// Multer: XML feed files (max 50MB)
const uploadXml = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.xml') cb(null, true);
    else cb(new Error('Only XML files are accepted'));
  }
});

// In-memory session state (per server instance)
let sessionState = {
  supplierData: null,
  feedData: null,
  masterData: null,
  lastXml: null,
  lastFilename: null
};

// ─────────────────────────────────────────────
// POST /api/parse-supplier
// Accept supplier Excel upload, extract PO/ASN refs
// ─────────────────────────────────────────────
app.post('/api/parse-supplier', upload.single('supplierFile'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const parsed = await supplierReader.parse(req.file.buffer);
    // Debug: log header keys and first row
    if (parsed.rows.length > 0) {
      console.log('[parse-supplier] headers found:', Object.keys(parsed.rows[0]));
      console.log('[parse-supplier] first row sample:', JSON.stringify(parsed.rows[0]).slice(0, 300));
    } else {
      console.log('[parse-supplier] no rows parsed — headerRow:', parsed.headerRowNum);
    }
    sessionState.supplierData = parsed;
    sessionState.feedData = null;
    sessionState.masterData = null;
    sessionState.lastXml = null;
    sessionState.lastFilename = null;

    // Sanitize rows for JSON: convert ExcelJS Date/RichText/formula objects to primitives
    const sanitizeVal = v => {
      if (v === null || v === undefined) return '';
      if (v instanceof Date) return v.toLocaleDateString('en-GB');
      if (typeof v === 'object' && 'result' in v) return String(v.result ?? '');
      if (typeof v === 'object' && 'richText' in v) return (v.richText || []).map(r => r.text || '').join('');
      if (typeof v === 'object' && 'formula' in v) return '';
      if (typeof v === 'object') return String(v);
      return v;
    };
    const safeRows = parsed.rows.map(r =>
      Object.fromEntries(Object.entries(r).map(([k, v]) => [k, sanitizeVal(v)]))
    );

    res.json({
      success: true,
      rowCount: safeRows.length,
      poRefs: [...new Set(safeRows.map(r => String(r.PO_Number || '').trim()).filter(Boolean))],
      rows: safeRows,
      preview: safeRows.slice(0, 5),
      validationErrors: parsed.validationErrors || []
    });
  } catch (err) {
    console.error('parse-supplier error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/fetch-feeds
// Fetch PO + ASN XMLs from Azure Blob by refs
// ─────────────────────────────────────────────
app.post('/api/fetch-feeds', async (req, res) => {
  try {
    const { poRefs, asnRefs } = req.body;
    if (!poRefs) return res.status(400).json({ error: 'poRefs required' });

    const feedData = await blobClient.fetchFeeds(poRefs, asnRefs || []);
    sessionState.feedData = feedData;

    res.json({
      success: true,
      poFeedCount: feedData.poFeeds.length,
      asnFeedCount: feedData.asnFeeds.length,
      carrierAsnCount: (feedData.carrierAsnFiles || []).length,
      localMode: feedData.localMode || false,
      errors: feedData.errors || [],
      summary: {
        posFound: feedData.poFeeds.map(p => p.orderId),
        asnsFound: feedData.asnFeeds.map(a => a.documentId),
        carrierAsnsFound: (feedData.carrierAsnFiles || []).map(f => f.filename)
      },
      feedsSummary: feedData.poFeeds.map(p => ({
        orderId:      p.orderId,
        filename:     (feedData.poFilenames || {})[p.orderId] || `PO_${p.orderId}.xml`,
        supplierName: p.supplierName,
        factoryName:  p.factoryName,
        shipDate:     p.shipDate,
        cancelDate:   p.cancelDate,
        incoterms:    p.incoterms,
        lineCount:    p.lineItems.length,
        lineItems:    p.lineItems.map(li => ({
          sku:          li.sku,
          productStyle: li.productStyle,
          description:  li.description,
          poQty:        li.poQty,
          mode:         li.mode
        }))
      })),
      carrierAsnFiles: (feedData.carrierAsnFiles || []).map(f => ({
        filename:  f.filename,
        poRef:     f.poRef,
        blobPath:  f.blobPath || null,
        asnGroups: (f.parsed || []).map(g => ({
          asnId:    g.asnId,
          fcId:     g.fcId,
          shipDate: g.shipDate,
          supplier: g.supplier,
          lines:    (g.lines || []).map(l => ({
            sku:         l.sku,
            ean:         l.ean,
            description: l.description,
            size:        l.size,
            colour:      l.colour,
            quantity:    l.quantity,
            packFormat:  l.packFormat
          }))
        }))
      }))
    });
  } catch (err) {
    console.error('fetch-feeds error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/feed-raw?type=po&ref={orderId}
//            OR  ?type=carrier&ref={filename}
// Return raw XML for preview from session
// ─────────────────────────────────────────────
app.get('/api/feed-raw', (req, res) => {
  if (!sessionState.feedData) {
    return res.status(404).json({ error: 'No feed data in session. Fetch feeds first.' });
  }
  const { type, ref } = req.query;
  if (!ref) return res.status(400).json({ error: 'ref parameter required' });

  const download = req.query.download === '1';

  if (type === 'po') {
    const xml = sessionState.feedData.poFeedXmls?.[ref];
    if (!xml) return res.status(404).json({ error: `PO feed not found for orderId: ${ref}` });
    if (download) res.set('Content-Disposition', `attachment; filename="PO_${ref}.xml"`);
    res.set('Content-Type', 'application/xml; charset=utf-8').send(xml);
  } else if (type === 'carrier') {
    const file = (sessionState.feedData.carrierAsnFiles || []).find(f => f.filename === ref);
    if (!file) return res.status(404).json({ error: `Carrier ASN not found: ${ref}` });
    if (download) res.set('Content-Disposition', `attachment; filename="${ref}"`);
    res.set('Content-Type', 'application/xml; charset=utf-8').send(file.xml);
  } else {
    res.status(400).json({ error: 'type must be "po" or "carrier"' });
  }
});

// ─────────────────────────────────────────────
// POST /api/upload-feeds
// Accept uploaded PO and/or ASN XML files and parse them
// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// POST /api/upload-feeds
// Accept uploaded PO and/or ASN XML files and parse them
// ─────────────────────────────────────────────
app.post('/api/upload-feeds', (req, res, next) => {
  uploadXml.fields([
    { name: 'poFeedFile',  maxCount: 1 },
    { name: 'asnFeedFile', maxCount: 1 }
  ])(req, res, (err) => {
    if (err) return res.status(400).json({ error: err.message });
    next();
  });
}, async (req, res) => {
  try {
    const poFeeds  = [];
    const asnFeeds = [];
    const errors   = [];

    if (req.files?.poFeedFile?.[0]) {
      const xmlStr = req.files.poFeedFile[0].buffer.toString('utf8');
      try {
        poFeeds.push(await poParser.parse(xmlStr));
      } catch (e) {
        errors.push(`PO parse error: ${e.message}`);
      }
    }

    if (req.files?.asnFeedFile?.[0]) {
      const xmlStr = req.files.asnFeedFile[0].buffer.toString('utf8');
      try {
        asnFeeds.push(await asnParser.parse(xmlStr));
      } catch (e) {
        errors.push(`ASN parse error: ${e.message}`);
      }
    }

    if (poFeeds.length === 0 && asnFeeds.length === 0 && errors.length === 0) {
      return res.status(400).json({ error: 'No PO or ASN XML file uploaded.' });
    }

    const feedData = { poFeeds, asnFeeds, errors, localMode: false };
    sessionState.feedData = feedData;

    res.json({
      success: true,
      poFeedCount:  poFeeds.length,
      asnFeedCount: asnFeeds.length,
      errors,
      summary: {
        posFound:  poFeeds.map(p => p.orderId),
        asnsFound: asnFeeds.map(a => a.documentId)
      }
    });
  } catch (err) {
    console.error('upload-feeds error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/search-blob?query=
// List/search blobs by name prefix in Azure Blob Storage
// ─────────────────────────────────────────────
app.get('/api/search-blob', async (req, res) => {
  try {
    const query = (req.query.query || '').trim();
    const results = await blobClient.searchBlobs(query);
    res.json({ success: true, results });
  } catch (err) {
    console.error('search-blob error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/build-bible
// Merge supplier + feeds → write Excel
// ─────────────────────────────────────────────
app.post('/api/build-bible', async (req, res) => {
  try {
    if (!sessionState.supplierData) return res.status(400).json({ error: 'No supplier data. Run parse-supplier first.' });
    if (!sessionState.feedData) return res.status(400).json({ error: 'No feed data. Run fetch-feeds first.' });

    const { masterRows, filePath, warnings } = await bibleBuilder.build(
      sessionState.supplierData,
      sessionState.feedData
    );
    sessionState.masterData = masterRows;

    const filename = path.basename(filePath);
    res.json({
      success: true,
      masterRowCount: masterRows.length,
      downloadUrl: `/bible/${filename}`,
      warnings: warnings || []
    });
  } catch (err) {
    console.error('build-bible error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/generate-vbkreq
// Build VBKREQ XML from MASTER data
// ─────────────────────────────────────────────
app.post('/api/generate-vbkreq', async (req, res) => {
  try {
    if (!sessionState.masterData) return res.status(400).json({ error: 'No master data. Run build-bible first.' });

    const { xml, filename, ctrlNumber, version } = await vbkreqBuilder.build(sessionState.masterData, req.body.purposeCd || '13');
    sessionState.lastXml = xml;
    sessionState.lastFilename = filename;

    res.json({ success: true, xml, filename, ctrlNumber, version });
  } catch (err) {
    console.error('generate-vbkreq error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/upload-sftp
// Upload XML to E2open SFTP
// ─────────────────────────────────────────────
app.post('/api/upload-sftp', async (req, res) => {
  try {
    const { filename, xmlContent } = req.body;
    const fname = filename || sessionState.lastFilename;
    const xml = xmlContent || sessionState.lastXml;
    if (!fname || !xml) return res.status(400).json({ error: 'No XML to upload. Run generate-vbkreq first.' });

    const result = await sftpUploader.upload(fname, xml);
    res.json({ success: true, ...result });
  } catch (err) {
    console.error('upload-sftp error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/generation-log
// Return last 50 log entries
// ─────────────────────────────────────────────
app.get('/api/generation-log', async (req, res) => {
  try {
    const log = await bibleBuilder.getGenerationLog();
    res.json({ success: true, entries: log.slice(-50).reverse() });
  } catch (err) {
    console.error('generation-log error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// Serve frontend for all other routes
// ─────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'frontend', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`CarrierBookingStub running at http://localhost:${PORT}`);
});
