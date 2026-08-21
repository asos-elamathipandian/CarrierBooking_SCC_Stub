'use strict';

/**
 * pipeline-runner.js
 *
 * Runs the full automated booking pipeline without an HTTP request:
 *   fetch Databricks ASNs → build bible → generate VBKREQs → upload SFTP → report
 *
 * Called from blob-webhook-scheduler after supplier files are ingested.
 * The same underlying modules are used by the manual UI route handlers.
 */

const path = require('path');
const fs   = require('fs');

const databricksAsnReader = require('./databricks-asn-reader');
const bibleBuilder        = require('./bible-builder');
const vbkreqBuilder       = require('./vbkreq-builder');
const sftpUploader        = require('./sftp-uploader');
const reportSender        = require('./report-sender');

const OUTPUT_DIR = path.join(__dirname, '..', 'output');

// ── Shared helpers (mirrors server.js) ───────────────────────────────────────

const RESUB_FIELDS = [
  { key: 'Header_Booking_Qty',                  label: 'Total Units'          },
  { key: 'PO_Header_Cartons',                   label: 'Cartons'              },
  { key: 'PO_Header_UnitWeight',                label: 'Unit Weight'          },
  { key: 'Cargo_Ready_Planned_Collection_Date', label: 'Cargo Ready Date'     },
  { key: 'Carrier_Booking_Request_Date',        label: 'Booking Request Date' },
  { key: 'Traffic_Mode',                        label: 'Traffic Mode'         },
  { key: 'PO_Header_CartonType',                label: 'Carton Type'          },
];

const RESUB_FIELDS_AB = [
  { key: 'Header_Booking_Qty',                  label: 'Total Units'          },
  { key: 'No_of_Cartons',                       label: 'Cartons'              },
  { key: 'Unit_Weight_KG',                      label: 'Unit Weight'          },
  { key: 'Cargo_Ready_Planned_Collection_Date', label: 'Cargo Ready Date'     },
  { key: 'Carrier_Booking_Request_Date',        label: 'Booking Request Date' },
  { key: 'Traffic_Mode',                        label: 'Traffic Mode'         },
  { key: 'Carton_Type',                         label: 'Carton Type'          },
];

function normField(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  if (!s) return '';
  const dm = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (dm) return `${dm[3]}-${dm[2]}-${dm[1]}`;
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

function resolveGroupKey(row) {
  const bg = String(row.Booking_Group || '').trim();
  if (bg === 'Multiple') return '__ALL__';
  const m = bg.match(/^Multiple POs-(BK\d+)$/i);
  if (m) return m[1].toUpperCase();
  return `PO__${String(row.PO_Number || '').trim()}`;
}

function groupLabel(group) {
  if (group === '__ALL__') return 'Multiple';
  if (group.startsWith('PO__')) return group.replace('PO__', '');
  return group;
}

function findOriginalTimestamp(bookingRef, logEntries) {
  const orig = (logEntries || [])
    .filter(e => e.bookingRef === bookingRef && (!e.purposeCd || e.purposeCd === '13'))
    .sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp))[0];
  if (!orig) return null;
  const d   = new Date(orig.timestamp);
  const pad = v => String(v).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())} ${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function poTotals(rows) {
  const seen = new Set();
  let cartons = 0, weight = 0;
  for (const r of rows) {
    if (!seen.has(r.PO_Number)) {
      seen.add(r.PO_Number);
      cartons += parseFloat(r.PO_Header_Cartons  || r.No_of_Cartons)   || 0;
      weight  += parseFloat(r.PO_Header_UnitWeight || r.Unit_Weight_KG) || 0;
    }
  }
  return { cartons, weight };
}

// ── Core generate step (purposeCd='13', no manual overrides) ─────────────────

async function generateVbkreqs(sessionState) {
  const masterRows = sessionState.masterData;
  if (!masterRows?.length) throw new Error('No master data — build-bible must run first.');

  const logEntries = bibleBuilder.getGenerationLog();
  const groupMap   = new Map();
  for (const row of masterRows) {
    const k = resolveGroupKey(row);
    if (!groupMap.has(k)) groupMap.set(k, []);
    groupMap.get(k).push(row);
  }

  const generations   = [];
  const skippedGroups = [];

  for (const [group, groupRows] of groupMap) {
    const poNumbers = [...new Set(groupRows.map(r => r.PO_Number).filter(Boolean))];

    let effectivePurposeCd = '13';
    let autoResubmitReason = null;

    // Auto-upgrade to Cd 15 if previously submitted with different field values
    const prevEntry = logEntries
      .filter(e => e.purposeCd !== '01' && (e.poNumbers || []).some(p => poNumbers.includes(String(p))))
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))[0];

    if (prevEntry?.masterRows?.length) {
      const prevFirst = prevEntry.masterRows[0];
      const newFirst  = groupRows[0] || {};
      const changes   = RESUB_FIELDS
        .filter(f => {
          const nv = normField(newFirst[f.key]);
          const pv = normField(prevFirst[f.key]);
          return nv && pv && nv !== pv;
        })
        .map(f => `${f.label}: ${normField(prevFirst[f.key])} -> ${normField(newFirst[f.key])}`);

      if (changes.length > 0) {
        effectivePurposeCd = '15';
        autoResubmitReason = changes.join('; ');
        groupRows.forEach(r => { r.Booking_Ref = prevEntry.bookingRef; });
        console.log(`[Pipeline] Auto-resub PO ${poNumbers.join(',')} — ${autoResubmitReason}`);
      } else {
        const lbl = groupLabel(group);
        console.log(`[Pipeline] PO ${poNumbers.join(',')} already booked, no changes — skipping`);
        skippedGroups.push({ poNumbers, bookingRef: prevEntry.bookingRef, group: lbl });
        continue;
      }
    }

    const { xml, filename, ctrlNumber, version, bookingRef: vbRef,
            headerBkq, lineBkqSum, bkqDiscrepancy } = await vbkreqBuilder.build(
      groupRows,
      effectivePurposeCd,
      { originalTimestamp: effectivePurposeCd !== '13'
          ? findOriginalTimestamp(groupRows[0]?.Booking_Ref, logEntries)
          : null }
    );

    const asnRefs    = [...new Set(groupRows.map(r => r.ASN_Ref).filter(Boolean))];
    const bookingRef = vbRef || groupRows[0]?.Booking_Ref || '';
    const lbl        = groupLabel(group);
    const first      = groupRows[0] || {};
    const { cartons, weight } = poTotals(groupRows);

    bibleBuilder.appendGenerationLog({
      timestamp:          new Date().toISOString(),
      bookingRef,
      poNumbers,
      asnRefs,
      filename,
      ctrlNumber,
      group:              lbl,
      purposeCd:          effectivePurposeCd,
      resubmissionReason: autoResubmitReason,
      sftp:               null,
      supplier:           first.Supplier || first.Supplier_Name || first.supplierName || '',
      bookingGroup:       first.Booking_Group || lbl,
      cargoReadyDate:     first.Cargo_Ready_Planned_Collection_Date || first.CargoReadyDate || '',
      noOfCartons:        cartons || null,
      totalWeight:        weight  || null,
      headerBkq,
      lineBkqSum,
      bkqDiscrepancy,
      masterRows:         groupRows,
    });

    generations.push({ group: lbl, xml, filename, ctrlNumber, version,
                        poNumbers, asnRefs, bookingRef,
                        autoResubmit: !!autoResubmitReason,
                        resubmissionReason: autoResubmitReason });
  }

  // Auto Cd 15 for already-booked POs whose supplier data changed
  const alreadyBookedPOIds = [...new Set(
    (sessionState.feedData?.cancelledItems || [])
      .filter(c => c.type === 'ALREADY_BOOKED')
      .map(c => String(c.poId || '').trim())
      .filter(Boolean)
  )];

  if (alreadyBookedPOIds.length > 0) {
    const supplierRows = sessionState.supplierData?.rows || [];
    const abSupRows    = supplierRows.filter(r => alreadyBookedPOIds.includes(String(r.PO_Number || '').trim()));
    const abGroupMap   = new Map();
    for (const sRow of abSupRows) {
      const gk = resolveGroupKey(sRow);
      if (!abGroupMap.has(gk)) abGroupMap.set(gk, []);
      abGroupMap.get(gk).push(sRow);
    }

    for (const [abGroup, abRows] of abGroupMap) {
      const abPONums = [...new Set(abRows.map(r => String(r.PO_Number || '').trim()).filter(Boolean))];
      const prevEntry = logEntries
        .filter(e => e.purposeCd !== '01' && (e.poNumbers || []).some(p => abPONums.includes(String(p))))
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))[0];

      if (!prevEntry?.masterRows?.length) {
        console.log(`[Pipeline] PO ${abPONums.join(',')} booked outside tool — skipping auto re-sub`);
        continue;
      }

      const prevFirst = prevEntry.masterRows[0];
      const newFirst  = abRows[0];
      const changes   = RESUB_FIELDS_AB
        .filter(f => {
          const nv = String(newFirst[f.key] || '').trim();
          const pv = String(prevFirst[f.key] || '').trim();
          return nv && pv && nv !== pv;
        })
        .map(f => `${f.label}: ${String(prevFirst[f.key] || '').trim()} -> ${String(newFirst[f.key] || '').trim()}`);

      if (changes.length === 0) {
        console.log(`[Pipeline] PO ${abPONums.join(',')} already booked, no changes — staying skipped`);
        continue;
      }

      console.log(`[Pipeline] PO ${abPONums.join(',')} changes: ${changes.join('; ')} — raising Cd 15`);

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

      const abLbl = groupLabel(abGroup);
      const { xml: abXml, filename: abFilename, ctrlNumber: abCtrl, version: abVer,
              headerBkq: ab_hbkq, lineBkqSum: ab_lbkq, bkqDiscrepancy: ab_disc }
        = await vbkreqBuilder.build(resubRows, '15');

      fs.mkdirSync(OUTPUT_DIR, { recursive: true });
      fs.writeFileSync(path.join(OUTPUT_DIR, abFilename), abXml, 'utf8');

      const { cartons: abCartons, weight: abWeight } = poTotals(resubRows);
      const abFirst = resubRows[0] || {};
      bibleBuilder.appendGenerationLog({
        timestamp:          new Date().toISOString(),
        bookingRef:         prevEntry.bookingRef,
        poNumbers:          abPONums,
        asnRefs:            prevEntry.asnRefs || [],
        filename:           abFilename,
        ctrlNumber:         abCtrl,
        group:              abLbl,
        purposeCd:          '15',
        resubmissionReason: changes.join('; '),
        sftp:               null,
        supplier:           abFirst.Supplier_Name || prevEntry.supplier || '',
        bookingGroup:       abFirst.Booking_Group || abLbl,
        cargoReadyDate:     abFirst.Cargo_Ready_Planned_Collection_Date || '',
        noOfCartons:        abCartons || null,
        totalWeight:        abWeight  || null,
        headerBkq:          ab_hbkq,
        lineBkqSum:         ab_lbkq,
        bkqDiscrepancy:     ab_disc,
        masterRows:         resubRows,
      });

      generations.push({
        group: abLbl, xml: abXml, filename: abFilename,
        ctrlNumber: abCtrl, version: abVer,
        poNumbers: abPONums, asnRefs: prevEntry.asnRefs || [],
        bookingRef: prevEntry.bookingRef,
        autoResubmit: true, resubmissionReason: changes.join('; '),
      });
    }
  }

  return { generations, skippedGroups };
}

// ── Main pipeline ─────────────────────────────────────────────────────────────

/**
 * Run the full automated pipeline against the current sessionState.
 * sessionState must already have supplierData + supplierHeaderPoRefs populated
 * (done by blob-webhook-scheduler or sharepoint-scheduler before calling this).
 *
 * Returns a summary object: { poRefs, generations, skippedGroups, sftpResults, error }
 */
async function run(sessionState) {
  const poRefs = [...new Set(
    (sessionState.supplierHeaderPoRefs || []).map(p => String(p).trim()).filter(Boolean)
  )];

  if (!poRefs.length) {
    console.warn('[Pipeline] No PO refs in session — aborting.');
    return { poRefs: [], generations: [], skippedGroups: [], sftpResults: [], error: 'No PO refs' };
  }

  console.log(`[Pipeline] Starting automated run — ${poRefs.length} PO(s): ${poRefs.join(', ')}`);

  // ── Step 1: Fetch Databricks ASNs ────────────────────────────────────────
  let feedData;
  try {
    const useDb = (process.env.ASN_SOURCE || '').toLowerCase() === 'databricks';
    if (!useDb) {
      console.warn('[Pipeline] ASN_SOURCE != databricks — skipping Databricks fetch, using empty feed.');
      feedData = { carrierAsnFiles: [], cancelledItems: [], errors: [], localMode: true };
    } else {
      feedData = await databricksAsnReader.fetchAsnsByPoRefs(poRefs);
    }
    sessionState.feedData = feedData;
    console.log(`[Pipeline] ASN fetch complete — ${(feedData.carrierAsnFiles || []).length} file(s), ` +
      `${(feedData.cancelledItems || []).length} cancelled/booked item(s)`);
  } catch (err) {
    console.error('[Pipeline] ASN fetch failed:', err.message);
    return { poRefs, generations: [], skippedGroups: [], sftpResults: [], error: `ASN fetch: ${err.message}` };
  }

  // ── Step 2: Build bible ──────────────────────────────────────────────────
  let masterRows;
  try {
    const result = await bibleBuilder.build(sessionState.supplierData, feedData);
    masterRows = result.masterRows;
    sessionState.masterData = masterRows;
    if (result.warnings?.length) {
      console.warn('[Pipeline] Bible warnings:', result.warnings.join('; '));
    }
    console.log(`[Pipeline] Bible built — ${masterRows.length} master row(s)`);
  } catch (err) {
    console.error('[Pipeline] Bible build failed:', err.message);
    return { poRefs, generations: [], skippedGroups: [], sftpResults: [], error: `Bible build: ${err.message}` };
  }

  if (!masterRows.length) {
    console.warn('[Pipeline] No master rows after bible build — all ASNs may be cancelled/skipped.');
    return { poRefs, generations: [], skippedGroups: [], sftpResults: [], error: null };
  }

  // ── Step 3: Generate VBKREQs ─────────────────────────────────────────────
  let generations, skippedGroups;
  try {
    ({ generations, skippedGroups } = await generateVbkreqs(sessionState));
    sessionState.lastGenerations = generations;
    sessionState.lastXml         = generations[0]?.xml      || null;
    sessionState.lastFilename    = generations[0]?.filename  || null;
    sessionState.lastCtrlNumber  = generations[0]?.ctrlNumber || null;
    console.log(`[Pipeline] Generated ${generations.length} VBKREQ(s), skipped ${skippedGroups.length}`);
  } catch (err) {
    console.error('[Pipeline] VBKREQ generation failed:', err.message);
    return { poRefs, generations: [], skippedGroups: [], sftpResults: [], error: `Generate: ${err.message}` };
  }

  if (!generations.length) {
    console.log('[Pipeline] Nothing to upload — all groups skipped.');
    reportSender.sendScheduledReport().catch(e => console.error('[Pipeline] Report failed:', e.message));
    return { poRefs, generations: [], skippedGroups, sftpResults: [], error: null };
  }

  // ── Step 4: Upload to SFTP ───────────────────────────────────────────────
  let sftpResults = [];
  try {
    sftpResults = await sftpUploader.uploadBatch(
      generations.map(g => ({ filename: g.filename, xmlContent: g.xml }))
    );
    for (const r of sftpResults) {
      const gen = generations.find(g => g.filename === r.filename);
      if (gen) {
        bibleBuilder.updateGenerationLog(r.filename, gen.ctrlNumber, {
          sftp:       r.ok ? (r.localMode ? 'local' : 'uploaded') : 'error',
          sftpEnv:    r.sftpEnv  || null,
          sftpPath:   r.remotePath || null,
          uploadedAt: r.uploadedAt || new Date().toISOString(),
        });
      }
    }
    const ok  = sftpResults.filter(r => r.ok).length;
    const bad = sftpResults.filter(r => !r.ok).length;
    console.log(`[Pipeline] SFTP upload — ${ok} OK, ${bad} failed`);
  } catch (err) {
    console.error('[Pipeline] SFTP upload failed:', err.message);
    // Non-fatal — still send report so the operator knows what was generated
  }

  // ── Step 5: Send report ──────────────────────────────────────────────────
  reportSender.sendScheduledReport().catch(e => console.error('[Pipeline] Report failed:', e.message));

  return { poRefs, generations, skippedGroups, sftpResults, error: null };
}

module.exports = { run };
