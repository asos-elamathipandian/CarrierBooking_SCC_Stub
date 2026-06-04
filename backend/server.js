'use strict';

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

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'frontend')));

// Serve generated bible downloads
const bibleDir = path.join(__dirname, '..', 'bible');
if (!fs.existsSync(bibleDir)) fs.mkdirSync(bibleDir, { recursive: true });
app.use('/bible', express.static(bibleDir));

// Multer: store uploads in memory (max 10MB)
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
    res.json({
      success: true,
      rowCount: parsed.rows.length,
      poRefs: [...new Set(parsed.rows.map(r => r.PO_Number).filter(Boolean))],
      asnRefs: [...new Set(parsed.rows.map(r => r.ASN_Ref).filter(Boolean))],
      preview: parsed.rows.slice(0, 5)
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
    if (!poRefs || !asnRefs) return res.status(400).json({ error: 'poRefs and asnRefs required' });

    const feedData = await blobClient.fetchFeeds(poRefs, asnRefs);
    sessionState.feedData = feedData;

    res.json({
      success: true,
      poFeedCount: feedData.poFeeds.length,
      asnFeedCount: feedData.asnFeeds.length,
      localMode: feedData.localMode || false,
      errors: feedData.errors || [],
      summary: {
        posFound: feedData.poFeeds.map(p => p.orderId),
        asnsFound: feedData.asnFeeds.map(a => a.documentId)
      }
    });
  } catch (err) {
    console.error('fetch-feeds error:', err);
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

    const { masterRows, filePath } = await bibleBuilder.build(
      sessionState.supplierData,
      sessionState.feedData
    );
    sessionState.masterData = masterRows;

    const filename = path.basename(filePath);
    res.json({
      success: true,
      masterRowCount: masterRows.length,
      downloadUrl: `/bible/${filename}`
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
