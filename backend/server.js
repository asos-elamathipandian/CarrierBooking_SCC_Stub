'use strict';

require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');

const supplierReader       = require('./supplier-reader');
const blobClient           = require('./blob-client');
const databricksAsnReader  = require('./databricks-asn-reader');
const bibleBuilder         = require('./bible-builder');
const vbkreqBuilder        = require('./vbkreq-builder');
const sftpUploader         = require('./sftp-uploader');
const poParser             = require('./po-parser');
const asnParser            = require('./asn-parser');
const carrierAsnParser     = require('./carrier-asn-parser');
const spClient             = require('./sharepoint-client');
const spScheduler          = require('./sharepoint-scheduler');

const app = express();

// Helper: find the original v1.0 booking timestamp for OSBT/OSBK preservation
function findOriginalTimestamp(bookingRef, logEntries) {
  const orig = (logEntries || [])
    .filter(e => e.bookingRef === bookingRef && (!e.purposeCd || e.purposeCd === '13'))
    .sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp))[0];
  if (!orig) return null;
  const d = new Date(orig.timestamp);
  const pad = v => String(v).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())} ${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}
const PORT = 3000;

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, '..', 'frontend')));

// Serve generated bible downloads
const bibleDir  = path.join(__dirname, '..', 'bible');
const outputDir = path.join(__dirname, '..', 'output');
if (!fs.existsSync(bibleDir))  fs.mkdirSync(bibleDir,  { recursive: true });
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
app.use('/bible',  express.static(bibleDir));
app.use('/output', express.static(outputDir));

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
  supplierHeaderPoRefs: [],
  feedData: null,
  masterData: null,
  lastXml: null,
  lastFilename: null,
  lastCtrlNumber: null,
  lastGenerations: []
};

// ─────────────────────────────────────────────
// POST /api/parse-supplier
// Accept supplier Excel upload, extract PO/ASN refs
// ─────────────────────────────────────────────
app.post('/api/parse-supplier', upload.array('supplierFiles', 20), async (req, res) => {
  try {
    const files = req.files;
    if (!files || files.length === 0) return res.status(400).json({ error: 'No files uploaded' });

    // Parse each file and merge rows + validation errors
    let allRows = [];
    let allValidationErrors = [];
    let allHeaderPoRefs = [];

    for (const file of files) {
      const parsed = await supplierReader.parse(file.buffer);
      if (parsed.rows.length > 0) {
        console.log(`[parse-supplier] ${file.originalname} — headers:`, Object.keys(parsed.rows[0]));
        console.log(`[parse-supplier] ${file.originalname} — first row:`, JSON.stringify(parsed.rows[0]).slice(0, 300));
      } else {
        console.log(`[parse-supplier] ${file.originalname} — no rows parsed — headerRow:`, parsed.headerRowNum);
      }
      allRows = allRows.concat(parsed.rows);
      allValidationErrors = allValidationErrors.concat(
        (parsed.validationErrors || []).map(e => `[${file.originalname}] ${e}`)
      );
      allHeaderPoRefs.push(...(parsed.headerPoRefs || []));
    }

    sessionState.supplierData = { rows: allRows, validationErrors: allValidationErrors };
    sessionState.supplierHeaderPoRefs = allHeaderPoRefs;
    sessionState.supplierBuffers = files.map(f => ({ name: f.originalname, buffer: f.buffer }));
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
    const safeRows = allRows.map(r =>
      Object.fromEntries(Object.entries(r).map(([k, v]) => [k, sanitizeVal(v)]))
    );

    // Compute distinct POs from BOOKING_HEADER (authoritative, even if SKU_LINES is empty)
    const poRefs = [...new Set(allHeaderPoRefs.map(p => String(p).trim()).filter(Boolean))];

    // Compute booking groups from SKU rows (falls back to header POs if no SKU rows)
    const sourceRows = safeRows.length > 0 ? safeRows
      : poRefs.map(po => ({ PO_Number: po, Booking_Group: 'Single Booking' }));
    const groupKeys = new Set();
    for (const row of sourceRows) {
      const bg = String(row.Booking_Group || '').trim();
      const po = String(row.PO_Number    || '').trim();
      if (!po) continue;
      if (bg === 'Multiple') { groupKeys.add('__ALL__'); }
      else {
        const m = bg.match(/^Multiple POs-(BK\d+)$/i);
        groupKeys.add(m ? m[1].toUpperCase() : 'PO__' + po);
      }
    }

    res.json({
      success: true,
      rowCount: allRows.filter(r => !r._headerOnly).length,
      fileCount: files.length,
      poCount: poRefs.length,
      bookingCount: groupKeys.size,
      poRefs,
      validationErrors: allValidationErrors
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

    const useDb = (process.env.ASN_SOURCE || '').toLowerCase() === 'databricks';
    const feedData = useDb
      ? await databricksAsnReader.fetchAsnsByPoRefs(poRefs)
      : await blobClient.fetchCarrierFeedsOnly(poRefs);
    sessionState.feedData = feedData;

    // Enrich ALREADY_BOOKED items with the VB Ref from our generation log
    const cancelledItems = feedData.cancelledItems || [];
    const genLog = bibleBuilder.getGenerationLog() || [];
    for (const item of cancelledItems) {
      if (item.type === 'ALREADY_BOOKED' && item.asnId) {
        const logEntry = genLog.find(e =>
          (e.asnRefs || []).map(String).includes(String(item.asnId)) ||
          (e.poNumbers || []).map(String).includes(String(item.poId))
        );
        if (logEntry) {
          item.vbRef  = logEntry.bookingRef || null;
          item.reason = `ASN ${item.asnId} (PO ${item.poId}) already has a carrier booking — ${logEntry.bookingRef ? `VB Ref: ${logEntry.bookingRef}` : 'submitted previously'}`;
        }
      }
    }

    res.json({
      success: true,
      carrierAsnCount: (feedData.carrierAsnFiles || []).length,
      localMode: feedData.localMode || false,
      errors: feedData.errors || [],
      cancelledItems,
      feedsSummary: [],
      carrierAsnFiles: (feedData.carrierAsnFiles || []).map(f => ({
        filename:  f.filename,
        poRef:     f.poRef,
        blobPath:  f.blobPath || null,
        asnGroups: (f.parsed || []).map(g => ({
          asnId:    g.asnId,
          fcId:     g.fcId,
          shipDate: g.shipDate,
          supplier: g.supplier,
          supplierCode:  g.supplierCode,
          shippingPoint: g.shippingPoint,
          shippingTerms: g.shippingTerms,
          lines:    (g.lines || []).map(l => ({
            sku:         l.sku,
            ean:         l.ean,
            description: l.description,
            size:        l.size,
            colour:      l.colour,
            quantity:    l.quantity,
            country:     l.country,
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
    { name: 'asnFeedFile', maxCount: 1 }
  ])(req, res, (err) => {
    if (err) return res.status(400).json({ error: err.message });
    next();
  });
}, async (req, res) => {
  try {
    const carrierAsnFiles = [];
    const errors = [];

    if (req.files?.asnFeedFile?.[0]) {
      const xmlStr   = req.files.asnFeedFile[0].buffer.toString('utf8');
      const filename = req.files.asnFeedFile[0].originalname;
      try {
        const parsed = await carrierAsnParser.parse(xmlStr);
        const poRef  = parsed[0]?.poId || '';
        carrierAsnFiles.push({ filename, xml: xmlStr, poRef, parsed });
      } catch (e) {
        errors.push(`Carrier feed parse error: ${e.message}`);
      }
    }

    if (carrierAsnFiles.length === 0 && errors.length === 0) {
      return res.status(400).json({ error: 'No carrier XML file uploaded.' });
    }

    const feedData = { poFeeds: [], asnFeeds: [], carrierAsnFiles, errors, localMode: false };
    sessionState.feedData = feedData;

    res.json({
      success: true,
      carrierAsnCount: carrierAsnFiles.length,
      errors,
      feedsSummary: [],
      carrierAsnFiles: carrierAsnFiles.map(f => ({
        filename:  f.filename,
        poRef:     f.poRef,
        asnGroups: (f.parsed || []).map(g => ({
          asnId:         g.asnId,
          fcId:          g.fcId,
          shipDate:      g.shipDate,
          supplier:      g.supplier,
          supplierCode:  g.supplierCode,
          shippingPoint: g.shippingPoint,
          shippingTerms: g.shippingTerms,
          lines: (g.lines || []).map(l => ({
            sku:         l.sku,
            ean:         l.ean,
            description: l.description,
            size:        l.size,
            colour:      l.colour,
            quantity:    l.quantity,
            country:     l.country,
            packFormat:  l.packFormat
          }))
        }))
      }))
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

    // feedData optional — default to empty if carrier feed not yet fetched
    const feedData = sessionState.feedData || { poFeeds: [], asnFeeds: [], carrierAsnFiles: [] };

    const { masterRows, filePath, warnings, extraSkuWarnings } = await bibleBuilder.build(
      sessionState.supplierData,
      feedData
    );
    sessionState.masterData = masterRows;

    const filename = path.basename(filePath);
    res.json({
      success: true,
      masterRowCount: masterRows.length,
      downloadUrl: `/bible/${filename}`,
      warnings: warnings || [],
      extraSkuWarnings: extraSkuWarnings || []
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

    const purposeCd = req.body.purposeCd || '13';
    const overrideCargoReady     = req.body.overrideCargoReady     || '';
    const overrideBookingReqDate = req.body.overrideBookingReqDate || '';
    // Optional map { [PO_Number]: bookingRef } sent by the UI when user selects a specific VB to re-submit
    const overrideBookingRefs    = req.body.overrideBookingRefs    || null;

    // Apply optional date overrides to a working copy of master rows
    const workingRows = sessionState.masterData.map(row => {
      if (!overrideCargoReady && !overrideBookingReqDate) return row;
      const r = { ...row };
      if (overrideCargoReady)     r.Cargo_Ready_Planned_Collection_Date = overrideCargoReady;
      if (overrideBookingReqDate) r.Carrier_Booking_Request_Date        = overrideBookingReqDate;
      return r;
    });

    // For cancellations AND re-submissions, reuse the correct bookingRef per PO from the generation log
    if (purposeCd === '01' || purposeCd === '15') {
      const logEntries = readGenerationLog();
      // Build most-recent map as fallback
      const poRefMap = {};
      for (const entry of logEntries) {
        for (const po of (entry.poNumbers || [])) {
          const key = String(po).trim();
          if (!poRefMap[key] || new Date(entry.timestamp) > new Date(poRefMap[key].timestamp)) {
            poRefMap[key] = { bookingRef: entry.bookingRef, timestamp: entry.timestamp };
          }
        }
      }
      for (const row of workingRows) {
        const poKey = String(row.PO_Number || '').trim();
        // UI-selected override takes priority; otherwise use the most-recent from the log
        const chosen = (overrideBookingRefs && overrideBookingRefs[poKey])
          ? overrideBookingRefs[poKey]
          : (poRefMap[poKey] ? poRefMap[poKey].bookingRef : null);
        if (chosen) row.Booking_Ref = chosen;
      }
    }

    // Derive a grouping key from Booking_Group:
    //   "Single Booking"      → one VBKREQ per PO  (key = PO_Number)
    //   "Multiple POs-BKxxx"  → all rows with same code in one VBKREQ (key = BKxxx)
    //   "Multiple"            → everything in one VBKREQ (key = __ALL__)
    //   blank / unknown       → treat as Single Booking for safety
    function resolveGroupKey(row) {
      const bg = String(row.Booking_Group || '').trim();
      if (bg === 'Multiple') return '__ALL__';
      const multiMatch = bg.match(/^Multiple POs-(BK\d+)$/i);
      if (multiMatch) return multiMatch[1].toUpperCase();
      // "Single Booking" or blank/legacy → one VBKREQ per PO
      return `PO__${String(row.PO_Number || '').trim()}`;
    }

    const groupMap = new Map();
    for (const row of workingRows) {
      const group = resolveGroupKey(row);
      if (!groupMap.has(group)) groupMap.set(group, []);
      groupMap.get(group).push(row);
    }

    const RESUB_FIELDS = [
      { key: 'Header_Booking_Qty',                  label: 'Total Units'          },
      { key: 'No_of_Cartons',                       label: 'Cartons'              },
      { key: 'Unit_Weight_KG',                      label: 'Unit Weight'          },
      { key: 'Cargo_Ready_Planned_Collection_Date', label: 'Cargo Ready Date'     },
      { key: 'Carrier_Booking_Request_Date',        label: 'Booking Request Date' },
      { key: 'Traffic_Mode',                        label: 'Traffic Mode'         },
      { key: 'Carton_Type',                         label: 'Carton Type'          },
    ];

    // Helper: find the original v1.0 booking timestamp for a given bookingRef (YYYYMMDD HHMMSS)
    // (defined at module scope — see top of file)


    const generations   = [];
    const skippedGroups = []; // same PO, no field changes — skip entirely
    for (const [group, groupRows] of groupMap) {
      const poNumbers = [...new Set(groupRows.map(r => r.PO_Number).filter(Boolean))];

      // ── Auto-upgrade to Cd 15 if this PO was previously submitted with different values ──
      let effectivePurposeCd = purposeCd;
      let autoResubmitReason = null;
      if (purposeCd === '13') {
        const logEntries = bibleBuilder.getGenerationLog();
        const prevEntry  = logEntries
          .filter(e => e.purposeCd !== '01' && (e.poNumbers || []).some(p => poNumbers.includes(String(p))))
          .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))[0];
        if (prevEntry?.masterRows?.length) {
          const prevFirst = prevEntry.masterRows[0];
          const newFirst  = groupRows[0] || {};
          const changes   = RESUB_FIELDS
            .filter(f => {
              const nv = String(newFirst[f.key]  || '').trim();
              const pv = String(prevFirst[f.key] || '').trim();
              return nv && pv && nv !== pv;
            })
            .map(f => `${f.label}: ${String(prevFirst[f.key]||'').trim()} -> ${String(newFirst[f.key]||'').trim()}`);
          if (changes.length > 0) {
            effectivePurposeCd = '15';
            autoResubmitReason = changes.join('; ');
            groupRows.forEach(r => { r.Booking_Ref = prevEntry.bookingRef; });
            console.log(`[Auto-resub] PO ${poNumbers.join(',')} — ${autoResubmitReason}`);
          } else {
            // Same PO, no changes — skip, do not generate a new VBKREQ
            const groupLabel = group === '__ALL__' ? 'Multiple' : group.startsWith('PO__') ? group.replace('PO__', '') : group;
            console.log(`[Skip] PO ${poNumbers.join(',')} already booked with no changes — skipping`);
            skippedGroups.push({ poNumbers, bookingRef: prevEntry.bookingRef, group: groupLabel });
            continue;
          }
        }
      }

      const { xml, filename, ctrlNumber, version, bookingRef: vbRef, headerBkq, lineBkqSum, bkqDiscrepancy } = await vbkreqBuilder.build(
        groupRows,
        effectivePurposeCd,
        { originalTimestamp: effectivePurposeCd !== '13' ? findOriginalTimestamp(groupRows[0]?.Booking_Ref, bibleBuilder.getGenerationLog()) : null }
      );
      const asnRefs   = [...new Set(groupRows.map(r => r.ASN_Ref).filter(Boolean))];
      const bookingRef = vbRef || groupRows[0]?.Booking_Ref || '';
      // Human-readable label: strip the internal PO__ prefix used for Single Booking keys
      const groupLabel = group === '__ALL__' ? 'Multiple' : group.startsWith('PO__') ? group.replace('PO__', '') : group;
      // Compute per-PO totals (PO_Header_* preferred; falls back to line-level fields)
      const _seenPOs = new Set();
      let _totalCartons = 0, _totalWeight = 0;
      for (const r of groupRows) {
        if (!_seenPOs.has(r.PO_Number)) {
          _seenPOs.add(r.PO_Number);
          _totalCartons += parseFloat(r.PO_Header_Cartons  || r.No_of_Cartons)  || 0;
          _totalWeight  += parseFloat(r.PO_Header_UnitWeight || r.Unit_Weight_KG) || 0;
        }
      }
      const _first = groupRows[0] || {};
      bibleBuilder.appendGenerationLog({
        timestamp:          new Date().toISOString(),
        bookingRef,
        poNumbers,
        asnRefs,
        filename,
        ctrlNumber,
        group:              groupLabel,
        purposeCd:          effectivePurposeCd,
        resubmissionReason: autoResubmitReason,
        sftp:               null,
        supplier:           _first.Supplier || _first.Supplier_Name || _first.supplierName || '',
        bookingGroup:       _first.Booking_Group || groupLabel,
        cargoReadyDate:     _first.Cargo_Ready_Planned_Collection_Date || _first.CargoReadyDate || '',
        noOfCartons:        _totalCartons || null,
        totalWeight:        _totalWeight  || null,
        headerBkq,
        lineBkqSum,
        bkqDiscrepancy,
        masterRows:         groupRows
      });
      generations.push({ group: groupLabel, xml, filename, ctrlNumber, version, poNumbers, asnRefs, bookingRef,
        autoResubmit: !!autoResubmitReason, resubmissionReason: autoResubmitReason });
    }

    sessionState.lastGenerations  = generations;
    sessionState.lastXml          = generations[0]?.xml          || null;
    sessionState.lastFilename     = generations[0]?.filename     || null;
    sessionState.lastCtrlNumber   = generations[0]?.ctrlNumber   || null;

    // ── Auto re-submission: already-booked POs with changed supplier data ────
    // Only trigger when processing fresh new submissions (purposeCd='13').
    // If the supplier re-sends the same PO with different values (qty, dates, etc.)
    // and that PO was previously booked by this tool, automatically raise Cd 15.
    if (purposeCd === '13') {
      const alreadyBookedPOIds = [...new Set(
        (sessionState.feedData?.cancelledItems || [])
          .filter(c => c.type === 'ALREADY_BOOKED')
          .map(c => String(c.poId || '').trim())
          .filter(Boolean)
      )];

      if (alreadyBookedPOIds.length > 0) {
        const logEntries = bibleBuilder.getGenerationLog();
        const supplierRows = sessionState.supplierData?.rows || [];

        const RESUB_FIELDS = [
          { key: 'Header_Booking_Qty',                  label: 'Total Units'         },
          { key: 'No_of_Cartons',                       label: 'Cartons'             },
          { key: 'Unit_Weight_KG',                      label: 'Unit Weight'         },
          { key: 'Cargo_Ready_Planned_Collection_Date', label: 'Cargo Ready Date'    },
          { key: 'Carrier_Booking_Request_Date',        label: 'Booking Request Date'},
          { key: 'Traffic_Mode',                        label: 'Traffic Mode'        },
          { key: 'Carton_Type',                         label: 'Carton Type'         },
        ];

        // Group already-booked supplier rows by Booking_Group
        const abSupRows = supplierRows.filter(r => alreadyBookedPOIds.includes(String(r.PO_Number || '').trim()));
        const abGroupMap = new Map();
        for (const sRow of abSupRows) {
          const gk = resolveGroupKey(sRow);
          if (!abGroupMap.has(gk)) abGroupMap.set(gk, []);
          abGroupMap.get(gk).push(sRow);
        }

        for (const [abGroup, abRows] of abGroupMap) {
          const abPONums = [...new Set(abRows.map(r => String(r.PO_Number || '').trim()).filter(Boolean))];

          // Find most recent non-cancelled log entry for these POs
          const prevEntry = logEntries
            .filter(e => e.purposeCd !== '01' && (e.poNumbers || []).some(p => abPONums.includes(String(p))))
            .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))[0];

          if (!prevEntry?.masterRows?.length) {
            console.log(`[Auto-resub] PO ${abPONums.join(',')} booked outside this tool — skipping auto re-sub`);
            continue;
          }

          // Detect changed fields between new supplier row and stored master row
          const prevFirst = prevEntry.masterRows[0];
          const newFirst  = abRows[0];
          const changes = RESUB_FIELDS
            .filter(f => {
              const nv = String(newFirst[f.key] || '').trim();
              const pv = String(prevFirst[f.key] || '').trim();
              return nv && pv && nv !== pv;
            })
            .map(f => `${f.label}: ${String(prevFirst[f.key] || '').trim()} -> ${String(newFirst[f.key] || '').trim()}`);

          if (changes.length === 0) {
            console.log(`[Auto-resub] PO ${abPONums.join(',')} already booked — no changes detected, staying skipped`);
            continue;
          }

          console.log(`[Auto-resub] PO ${abPONums.join(',')} changes detected: ${changes.join('; ')} — raising Cd 15`);

          // Build re-sub rows: start from stored master rows, overlay new supplier values
          const resubRows = prevEntry.masterRows.map(r => ({
            ...r,
            Booking_Ref:                         prevEntry.bookingRef,
            Header_Booking_Qty:                  parseFloat(newFirst.Header_Booking_Qty)  || r.Header_Booking_Qty  || 0,
            No_of_Cartons:                       parseFloat(newFirst.No_of_Cartons)       || r.No_of_Cartons       || 0,
            PO_Header_Cartons:                   parseFloat(newFirst.No_of_Cartons)       || r.PO_Header_Cartons   || 0,
            Unit_Weight_KG:                      parseFloat(newFirst.Unit_Weight_KG)      || r.Unit_Weight_KG      || 0,
            PO_Header_UnitWeight:                parseFloat(newFirst.Unit_Weight_KG)      || r.PO_Header_UnitWeight|| 0,
            Cargo_Ready_Planned_Collection_Date: newFirst.Cargo_Ready_Planned_Collection_Date || r.Cargo_Ready_Planned_Collection_Date,
            Carrier_Booking_Request_Date:        newFirst.Carrier_Booking_Request_Date        || r.Carrier_Booking_Request_Date,
            Traffic_Mode:                        newFirst.Traffic_Mode || r.Traffic_Mode,
            Carton_Type:                         newFirst.Carton_Type  || r.Carton_Type,
          }));

          const abLabel = abGroup === '__ALL__' ? 'Multiple' : abGroup.startsWith('PO__') ? abGroup.replace('PO__', '') : abGroup;
          const { xml: abXml, filename: abFilename, ctrlNumber: abCtrl, version: abVer,
                  headerBkq: ab_hbkq, lineBkqSum: ab_lbkq, bkqDiscrepancy: ab_disc }
            = await vbkreqBuilder.build(resubRows, '15');

          // Save to local output dir (SFTP upload via normal upload-sftp flow)
          const abOutPath = path.join(outputDir, abFilename);
          fs.writeFileSync(abOutPath, abXml, 'utf8');

          const abSeenPOs = new Set(); let abCartons = 0, abWeight = 0;
          for (const r of resubRows) {
            if (!abSeenPOs.has(r.PO_Number)) {
              abSeenPOs.add(r.PO_Number);
              abCartons += parseFloat(r.PO_Header_Cartons || r.No_of_Cartons) || 0;
              abWeight  += parseFloat(r.PO_Header_UnitWeight || r.Unit_Weight_KG) || 0;
            }
          }
          const abFirst = resubRows[0] || {};
          bibleBuilder.appendGenerationLog({
            timestamp:          new Date().toISOString(),
            bookingRef:         prevEntry.bookingRef,
            poNumbers:          abPONums,
            asnRefs:            prevEntry.asnRefs || [],
            filename:           abFilename,
            ctrlNumber:         abCtrl,
            group:              abLabel,
            purposeCd:          '15',
            resubmissionReason: changes.join('; '),
            sftp:               null,
            supplier:           abFirst.Supplier_Name || prevEntry.supplier || '',
            bookingGroup:       abFirst.Booking_Group || abLabel,
            cargoReadyDate:     abFirst.Cargo_Ready_Planned_Collection_Date || '',
            noOfCartons:        abCartons || null,
            totalWeight:        abWeight  || null,
            headerBkq:          ab_hbkq,
            lineBkqSum:         ab_lbkq,
            bkqDiscrepancy:     ab_disc,
            masterRows:         resubRows,
          });

          generations.push({
            group:              abLabel,
            xml:                abXml,
            filename:           abFilename,
            ctrlNumber:         abCtrl,
            version:            abVer,
            poNumbers:          abPONums,
            asnRefs:            prevEntry.asnRefs || [],
            bookingRef:         prevEntry.bookingRef,
            autoResubmit:       true,
            resubmissionReason: changes.join('; '),
          });
        }
      }
    }

    res.json({ success: true, generations, skippedGroups });
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

    // Update log entry with SFTP outcome — find ctrlNumber for this filename
    const gen = (sessionState.lastGenerations || []).find(g => g.filename === fname);
    const ctrlNum = gen?.ctrlNumber || sessionState.lastCtrlNumber;
    bibleBuilder.updateGenerationLog(fname, ctrlNum, {
      sftp:       result.localMode ? 'local' : 'uploaded',
      sftpEnv:    result.sftpEnv || null,
      sftpPath:   result.remotePath || null,
      uploadedAt: result.uploadedAt || new Date().toISOString()
    });

    res.json({ success: true, ...result });
  } catch (err) {
    console.error('upload-sftp error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// POST /api/upload-sftp-batch
// Upload multiple XMLs over a single shared SFTP connection
// ─────────────────────────────────────────────
app.post('/api/upload-sftp-batch', async (req, res) => {
  try {
    const { files } = req.body; // [{ filename, xmlContent }]
    if (!Array.isArray(files) || files.length === 0) {
      return res.status(400).json({ error: 'files array is required' });
    }

    const results = await sftpUploader.uploadBatch(files);

    // Update generation log for each uploaded file
    for (const r of results) {
      const gen = (sessionState.lastGenerations || []).find(g => g.filename === r.filename);
      const ctrlNum = gen?.ctrlNumber || null;
      if (r.filename && ctrlNum !== undefined) {
        bibleBuilder.updateGenerationLog(r.filename, ctrlNum, {
          sftp:       r.ok ? (r.localMode ? 'local' : 'uploaded') : 'error',
          sftpEnv:    r.sftpEnv || null,
          sftpPath:   r.remotePath || null,
          uploadedAt: r.uploadedAt || new Date().toISOString()
        });
      }
    }

    res.json({ results });
  } catch (err) {
    console.error('upload-sftp-batch error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/tagged-template
// Build and return supplier Excel with VBKREQ_Ref column mapped to each PO
// ─────────────────────────────────────────────
app.get('/api/tagged-template', async (req, res) => {
  try {
    const generations = sessionState.lastGenerations || [];
    const headerPoRefs = sessionState.supplierHeaderPoRefs || [];
    if (!generations.length) return res.status(400).json({ error: 'No VBKREQs generated yet. Run the pipeline first.' });

    // Build PO → VBKREQ_Ref map
    const poToRef = {};
    for (const gen of generations) {
      const ref = gen.filename || gen.ctrlNumber || '';
      for (const po of (gen.poNumbers || [])) poToRef[String(po).trim()] = ref;
    }

    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    wb.creator = 'CarrierBookingStub';
    wb.created = new Date();
    const ws = wb.addWorksheet('PO_VBKREQ_MAP');
    ws.properties.tabColor = { argb: 'FF1F4E79' };

    // Header row
    const cols = [
      { header: 'PO_Number',   key: 'po',  width: 28 },
      { header: 'VBKREQ_Ref', key: 'ref', width: 60 },
      { header: 'Status',     key: 'st',  width: 18 }
    ];
    const hdr = ws.getRow(1);
    cols.forEach((c, i) => {
      const cell = hdr.getCell(i + 1);
      cell.value = c.header;
      cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
      cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getColumn(i + 1).width = c.width;
    });
    hdr.height = 24;

    // Data rows — one per PO from BOOKING_HEADER
    const allPOs = [...new Set(headerPoRefs.map(p => String(p).trim()).filter(Boolean))];
    allPOs.forEach((po, i) => {
      const ref = poToRef[po] || '';
      const row = ws.getRow(i + 2);
      row.getCell(1).value = po;
      row.getCell(2).value = ref;
      const matched = !!ref;
      row.getCell(3).value = matched ? 'Generated' : 'Not generated';
      row.getCell(3).fill = { type: 'pattern', pattern: 'solid',
        fgColor: { argb: matched ? 'FFE8F5E9' : 'FFFCE8E8' } };
      row.getCell(3).font = { color: { argb: matched ? 'FF1B5E20' : 'FF7B1F1F' }, bold: true, size: 10 };
      row.commit();
    });
    ws.views = [{ state: 'frozen', ySplit: 1, showGridLines: true }];

    const buf = await wb.xlsx.writeBuffer();
    const ts  = new Date().toISOString().replace(/[:.]/g, '').slice(0, 15);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="PO_VBKREQ_Map_${ts}.xlsx"`);
    res.send(buf);
  } catch (err) {
    console.error('tagged-template error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/tagged-supplier/:idx
// Return the original supplier Excel with VBKREQ_Ref written into PO Header
// ─────────────────────────────────────────────
app.get('/api/tagged-supplier/:idx', async (req, res) => {
  try {
    const idx      = parseInt(req.params.idx, 10);
    const buffers  = sessionState.supplierBuffers || [];
    const generations = sessionState.lastGenerations || [];

    if (!buffers[idx]) return res.status(404).json({ error: 'Supplier file not found. Re-upload the template.' });
    if (!generations.length) return res.status(400).json({ error: 'No VBKREQs generated yet. Run the pipeline first.' });

    // Build PO → VB ref map
    const poToRef = {};
    for (const gen of generations) {
      const ref = gen.bookingRef || gen.ctrlNumber || gen.filename || '';
      for (const po of (gen.poNumbers || [])) poToRef[String(po).trim()] = ref;
    }

    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffers[idx].buffer);

    const wsH = wb.getWorksheet('PO Header') || wb.getWorksheet('BOOKING_HEADER');
    if (!wsH) return res.status(400).json({ error: 'PO Header sheet not found in supplier file.' });

    // Locate header row and PO_Number column
    let headerRowNum = 1;
    let poColIdx     = 1;
    wsH.eachRow((row, rowNum) => {
      row.eachCell((cell, colNum) => {
        const v = String(cell.value || '').replace(/\s*\(.*?\)/, '').trim();
        if (v === 'PO_Number') { headerRowNum = rowNum; poColIdx = colNum; }
      });
    });

    // Append VBKREQ_Ref header after last used column
    const newColIdx = wsH.columnCount + 1;
    const hdrCell   = wsH.getRow(headerRowNum).getCell(newColIdx);
    hdrCell.value = 'VBKREQ_Ref';
    hdrCell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    hdrCell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    hdrCell.alignment = { horizontal: 'center', vertical: 'middle' };
    wsH.getColumn(newColIdx).width = 52;

    // Write VB ref into each data row
    wsH.eachRow((row, rowNum) => {
      if (rowNum <= headerRowNum) return;
      const po  = String(row.getCell(poColIdx).value || '').trim();
      if (!po) return;
      const ref = poToRef[po];
      const cell = row.getCell(newColIdx);
      if (ref) {
        cell.value = ref;
        cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
        cell.font  = { color: { argb: 'FF1B5E20' }, bold: true, size: 10 };
      } else {
        cell.value = 'Not generated';
        cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE8E8' } };
        cell.font  = { color: { argb: 'FF7B1F1F' }, size: 10 };
      }
      row.commit();
    });

    // Strip conditional formatting — cross-sheet CF rules cause ExcelJS serialization errors
    for (const ws of wb.worksheets) {
      ws.conditionalFormattings = [];
    }

    const buf      = await wb.xlsx.writeBuffer();
    const baseName = buffers[idx].name.replace(/\.xlsx?$/i, '');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}_VBRef.xlsx"`);
    res.send(buf);
  } catch (err) {
    console.error('tagged-supplier error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// ─────────────────────────────────────────────
// Helper: read full generation log from disk
// ─────────────────────────────────────────────
const LOG_PATH = path.join(__dirname, '..', 'bible', 'generation-log.json');
function readGenerationLog() {
  try {
    if (fs.existsSync(LOG_PATH)) return JSON.parse(fs.readFileSync(LOG_PATH, 'utf8'));
  } catch (_) {}
  return [];
}

// ─────────────────────────────────────────────
// POST /api/cancel-booking
// Generate + return cancellation VBKREQ(s) for given PO numbers
// ─────────────────────────────────────────────
app.post('/api/cancel-booking', async (req, res) => {
  try {
    const { inputs = [], purposeCd: reqPurpose = '01' } = req.body || {};
    const actionPurpose = reqPurpose === '15' ? '15' : '01'; // only 01 or 15 allowed here
    if (!inputs.length) return res.status(400).json({ error: 'No PO numbers or VB Refs provided.' });

    const logEntries = readGenerationLog();

    // Build a map: bookingRef → most-recent entry that has masterRows
    // Match inputs that look like VB refs (VB-...) or PO numbers
    const matchedByRef = new Map(); // bookingRef → entry
    for (const entry of logEntries) {
      if (!entry.masterRows?.length) continue;
      const ref = String(entry.bookingRef || '');
      const isNewer = !matchedByRef.has(ref) || new Date(entry.timestamp) > new Date(matchedByRef.get(ref).timestamp);
      if (!isNewer) continue;

      const matchesVbRef = inputs.some(i => /^VB-/i.test(i.trim()) && i.trim().toUpperCase() === ref.toUpperCase());
      const matchesPo    = inputs.some(i => !/^VB-/i.test(i.trim()) && (entry.poNumbers || []).map(String).includes(i.trim()));
      if (matchesVbRef || matchesPo) matchedByRef.set(ref, entry);
    }

    if (!matchedByRef.size) {
      return res.status(404).json({ error: 'No stored booking records found for the given PO(s) / VB Ref(s). Bookings must have been created with this tool to be cancelled here.' });
    }

    const generations = [];
    for (const [bookingRef, entry] of matchedByRef) {
      const workingRows = entry.masterRows.map(r => ({ ...r, Booking_Ref: bookingRef }));
      const { xml, filename, ctrlNumber, version, headerBkq: cancelHeaderBkq, lineBkqSum: cancelLineBkqSum, bkqDiscrepancy: cancelBkqDiscrepancy } = await vbkreqBuilder.build(
        workingRows,
        actionPurpose,
        { originalTimestamp: findOriginalTimestamp(bookingRef, bibleBuilder.getGenerationLog()) }
      );
      const poNums  = entry.poNumbers || [];
      const asnRefs = entry.asnRefs   || [];
      const groupLabel = entry.group || bookingRef;
      const _cancelFirst = workingRows[0] || {};
      const _cancelSeenPOs = new Set();
      let _cancelCartons = 0, _cancelWeight = 0;
      for (const r of workingRows) {
        if (!_cancelSeenPOs.has(r.PO_Number)) {
          _cancelSeenPOs.add(r.PO_Number);
          _cancelCartons += parseFloat(r.PO_Header_Cartons  || r.No_of_Cartons)  || 0;
          _cancelWeight  += parseFloat(r.PO_Header_UnitWeight || r.Unit_Weight_KG) || 0;
        }
      }
      bibleBuilder.appendGenerationLog({
        timestamp:      new Date().toISOString(),
        bookingRef,
        poNumbers:      poNums,
        asnRefs,
        filename,
        ctrlNumber,
        group:          groupLabel,
        purposeCd:      '01',
        sftp:           null,
        supplier:       _cancelFirst.Supplier || _cancelFirst.Supplier_Name || _cancelFirst.supplierName || '',
        bookingGroup:   _cancelFirst.Booking_Group || groupLabel,
        cargoReadyDate: _cancelFirst.Cargo_Ready_Planned_Collection_Date || _cancelFirst.CargoReadyDate || '',
        noOfCartons:    _cancelCartons || null,
        totalWeight:    _cancelWeight  || null,
        headerBkq:      cancelHeaderBkq,
        lineBkqSum:     cancelLineBkqSum,
        bkqDiscrepancy: cancelBkqDiscrepancy,
        masterRows:     workingRows
      });
      generations.push({ group: groupLabel, xml, filename, ctrlNumber, version, poNumbers: poNums, asnRefs, bookingRef });
    }
    sessionState.lastGenerations = generations;
    res.json({ success: true, generations });
  } catch (err) {
    console.error('cancel-booking error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/generation-log
// Return entries from the last 3 days, newest first
// ─────────────────────────────────────────────
app.get('/api/generation-log', async (req, res) => {
  try {
    const log      = await bibleBuilder.getGenerationLog();
    const cutoff   = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
    const filtered = log
      .filter(e => e.timestamp && new Date(e.timestamp) >= cutoff)
      .reverse();
    res.json({ success: true, entries: filtered });
  } catch (err) {
    console.error('generation-log error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/lookup-vbref?pos=PO1,PO2,...
// Return most-recent bookingRef per PO from log
// ─────────────────────────────────────────────
app.get('/api/lookup-vbref', (req, res) => {
  try {
    const inputList = String(req.query.pos || '').split(',').map(s => s.trim()).filter(Boolean);
    const entries   = readGenerationLog();
    const result    = {};
    for (const input of inputList) {
      const isVbRef = /^VB-/i.test(input);
      const matches = entries
        .filter(e => isVbRef
          ? String(e.bookingRef || '').toUpperCase() === input.toUpperCase()
          : (e.poNumbers || []).map(String).includes(String(input)))
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      if (matches.length) {
        if (isVbRef) {
          // Single VB ref lookup — return one entry
          result[input] = {
            bookingRef:    matches[0].bookingRef,
            poNumbers:     matches[0].poNumbers || [],
            timestamp:     matches[0].timestamp,
            filename:      matches[0].filename,
            hasMasterRows: !!(matches[0].masterRows?.length),
            allRefs:       null
          };
        } else {
          // PO lookup — return ALL unique VBs for this PO
          const seen = new Map();
          for (const e of matches) {
            const ref = String(e.bookingRef || '');
            if (!seen.has(ref)) seen.set(ref, e);
          }
          const allRefs = [...seen.values()].map(e => ({
            bookingRef:    e.bookingRef,
            timestamp:     e.timestamp,
            filename:      e.filename,
            hasMasterRows: !!(e.masterRows?.length)
          }));
          // Primary entry = most recent
          result[input] = {
            bookingRef:    allRefs[0].bookingRef,
            poNumbers:     matches[0].poNumbers || [],
            timestamp:     allRefs[0].timestamp,
            filename:      allRefs[0].filename,
            hasMasterRows: allRefs[0].hasMasterRows,
            allRefs
          };
        }
      }
    }
    res.json({ refs: result });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/sharepoint/status
// ─────────────────────────────────────────────
app.get('/api/sharepoint/status', (req, res) => {
  const status = spScheduler.readStatus();
  res.json({
    configured: spClient.isConfigured(),
    schedule:   process.env.SP_SCHEDULE || '',
    ...status
  });
});

// ─────────────────────────────────────────────
// POST /api/sharepoint/sync  — manual trigger
// ─────────────────────────────────────────────
app.post('/api/sharepoint/sync', async (req, res) => {
  if (!spClient.isConfigured()) {
    return res.status(400).json({ error: 'SharePoint not configured. Fill in SP_* vars in .env and restart.' });
  }
  // Run async — respond immediately so the UI can poll status
  spScheduler.runSync(sessionState).catch(err =>
    console.error('[SP sync] Error:', err.message)
  );
  res.json({ success: true, message: 'Sync started — poll /api/sharepoint/status for progress.' });
});

// ─────────────────────────────────────────────
// POST /api/sharepoint/dismiss-error  — clear stored error
// ─────────────────────────────────────────────
app.post('/api/sharepoint/dismiss-error', (req, res) => {
  spScheduler.writeStatus({ error: null });
  res.json({ success: true });
});

// ─────────────────────────────────────────────
// GET /api/sharepoint/files  — list current SP folder
// ─────────────────────────────────────────────
app.get('/api/sharepoint/files', async (req, res) => {
  if (!spClient.isConfigured()) {
    return res.status(400).json({ error: 'SharePoint not configured.' });
  }
  try {
    const files = await spClient.listTemplateFiles();
    res.json({ files });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─────────────────────────────────────────────
// GET /api/health — simple liveness check
// ─────────────────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({ ok: true });
});

// ─────────────────────────────────────────────
// GET /api/download-template
// Serve the blank supplier Excel template
// ─────────────────────────────────────────────
app.get('/api/download-template', (req, res) => {
  const today = new Date();
  const pad = n => String(n).padStart(2, '0');
  const dd = pad(today.getDate()), mm = pad(today.getMonth() + 1), yyyy = today.getFullYear();
  const slot = today.getHours() < 12 ? 'AM' : 'PM';
  const templateName = `Supplier PO sheet-${dd}${mm}${yyyy}-${slot}.xlsx`;
  const templatePath = path.join(__dirname, '..', 'samples', templateName);
  if (!fs.existsSync(templatePath)) {
    // Fallback: serve any matching Supplier PO sheet-*.xlsx in samples/
    const fallback = fs.readdirSync(path.join(__dirname, '..', 'samples'))
      .filter(f => /^Supplier PO sheet-.*\.xlsx$/i.test(f))
      .sort().reverse()[0];
    if (!fallback) return res.status(404).json({ error: 'Template not found.' });
    return res.download(path.join(__dirname, '..', 'samples', fallback), fallback);
  }
  res.download(templatePath, templateName);
});

// ─────────────────────────────────────────────
// Serve frontend for all other routes
// ─────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'frontend', 'index.html'));
});

// ─────────────────────────────────────────────
// Global error handler — always return JSON so
// the frontend never receives an HTML error page
// ─────────────────────────────────────────────
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  const status = err.status || err.statusCode || 500;
  res.status(status).json({ error: err.message || 'Internal server error' });
});

app.listen(PORT, () => {
  console.log(`CarrierBookingStub running at http://localhost:${PORT}`);
  // Clear any stale error from a previous server run so UI starts clean
  spScheduler.writeStatus({ error: null, running: false });
  spScheduler.start(sessionState);
});
