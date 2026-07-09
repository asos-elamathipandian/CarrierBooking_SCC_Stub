'use strict';

const API = 'http://localhost:3000/api';

// ── Session state ──────────────────────────────────────────────────────────────
const state = {
  poRefs: [],
  asnRefs: [],
  supplierRows: [],
  supplierFileNames: [],
  feedsFetched: false,
  biblBuilt: false,
  lastXml: null,
  lastFilename: null,
  generations: []
};

// ── Element refs ───────────────────────────────────────────────────────────────
const supplierFileInput  = document.getElementById('supplierFileInput');
const dropZone           = document.getElementById('dropZone');
const dropZoneText       = document.getElementById('dropZoneText');
const btnParseSupplier   = document.getElementById('btnParseSupplier');
const btnRunPipeline     = document.getElementById('btnRunPipeline');
const refsPreview        = document.getElementById('refsPreview');

// ── Helpers ────────────────────────────────────────────────────────────────────
function setStatus(step, type, html) {
  const el = document.getElementById(`status${step}`);
  if (!el) return;
  el.className = `status-box ${type}`;
  el.innerHTML = html;
}

function setBadge(step, type) {
  const el = document.getElementById(`badge${step}`);
  if (!el) return;
  el.className = `step-badge ${type}`;
  if (type === 'done') el.textContent = '✓';
}

// Pipeline stage helpers
function psSetState(n, state) { // 'pending' | 'active' | 'done' | 'error'
  const el = document.getElementById('pstage' + n);
  if (!el) return;
  el.className = 'pipeline-stage ps-' + state;
  const num = document.getElementById('psNum' + n);
  if (num && state === 'done') num.textContent = '✓';
}
function psSetInline(n, html) {
  const el = document.getElementById('psInline' + n);
  if (el) el.innerHTML = html;
}
function psShowBody(n) {
  const el = document.getElementById('psBody' + n);
  if (el) el.style.display = '';
}
function psSetResult(id, html) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = html;
}
function psShowProceed(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = '';
}

function setLoading(btn, loading) {
  if (loading) {
    btn._origHTML = btn.innerHTML;
    btn.innerHTML = '<span class="spinner"></span> Working…';
    btn.disabled = true;
  } else {
    btn.innerHTML = btn._origHTML || btn.innerHTML;
    btn.disabled = false;
  }
}

// ── Drop zone & file input ────────────────────────────────────────────────────
function applyFilesToZone(files) {
  if (!files || files.length === 0) return;
  dropZone.classList.add('has-file');
  dropZone.classList.remove('drag-over');
  dropZoneText.textContent = files.length === 1
    ? '✓ ' + files[0].name
    : `✓ ${files.length} files selected`;
  const listEl = document.getElementById('supplierFileList');
  listEl.innerHTML = [...files].map(f => `<li>📄 ${f.name}</li>`).join('');
  listEl.style.display = 'block';
  btnParseSupplier.disabled = false;
}

// Click on zone triggers file browser
// Reset value first so selecting the same file again still fires the change event
dropZone.addEventListener('click', () => { supplierFileInput.value = ''; supplierFileInput.click(); });

// Drag events
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  const allDropped = [...e.dataTransfer.files];
  const valid   = allDropped.filter(f => /\.xlsx?$/i.test(f.name));
  const invalid = allDropped.filter(f => !/\.xlsx?$/i.test(f.name));
  if (invalid.length > 0) {
    dropZone.classList.remove('drag-over');
    setStatus(1, 'error', '❌ Only Excel files (.xlsx or .xls) are accepted');
    return;
  }
  if (valid.length === 0) return;
  const dt = new DataTransfer();
  valid.forEach(f => dt.items.add(f));
  supplierFileInput.files = dt.files;
  applyFilesToZone(supplierFileInput.files);
});

supplierFileInput.addEventListener('change', () => applyFilesToZone(supplierFileInput.files));

// ── Step 1: Parse Supplier ─────────────────────────────────────────────────────
btnParseSupplier.addEventListener('click', async () => {
  const files = supplierFileInput.files;
  if (!files || files.length === 0) return;

  setLoading(btnParseSupplier, true);
  const loadMsg = files.length === 1
    ? '⏳ Parsing supplier Excel…'
    : `⏳ Parsing ${files.length} supplier Excel files…`;
  setStatus(1, 'loading', loadMsg);
  refsPreview.innerHTML = '';

  try {
    const fd = new FormData();
    for (const file of files) fd.append('supplierFiles', file);
    const res  = await fetch(`${API}/parse-supplier`, { method: 'POST', body: fd });
    const data = await res.json();

    if (!res.ok) throw new Error(data.error || 'Unknown error');

    // Guard: no PO data found → likely a blank template
    if ((data.poCount || 0) === 0) {
      setStatus(1, 'error',
        '⚠️ No PO data found in the uploaded file. ' +
        'Please fill in the <strong>PO Header</strong> sheet with your PO details first, then upload again.'
      );
      setLoading(btnParseSupplier, false);
      btnParseSupplier.disabled = false;
      return;
    }

    // Header-only template (no PO Lines) — SKUs auto-booked from ASN feed; this is expected
    if ((data.rowCount || 0) === 0 && (data.poCount || 0) > 0) {
      setStatus(1, 'success',
        `✅ Found <strong>${data.poCount}</strong> PO${data.poCount !== 1 ? 's' : ''} in PO Header. ` +
        'SKUs will be auto-booked from the Databricks ASN feed — proceed to Step 2.'
      );
    }

    state.poRefs          = data.poRefs  || [];
    state.asnRefs         = [];
    state.supplierRows    = [];
    state.supplierFileNames = [...supplierFileInput.files].map(f => f.name);

    // Auto-populate PO refs in the blob fetch panel
    const blobPoRefTags = document.getElementById('blobPoRefTags');
    if (blobPoRefTags) {
      blobPoRefTags.innerHTML = state.poRefs.length
        ? state.poRefs.map(r => `<span class="tag">PO: ${r}</span>`).join('')
        : '<span style="color:#aaa;font-size:12px">No PO refs found in template.</span>';
    }

    const fileCountNote = data.fileCount > 1 ? ` from <strong>${data.fileCount}</strong> files` : '';
    let html = `✅ Parsed successfully${fileCountNote}.`;
    if (data.validationErrors && data.validationErrors.length) {
      html += `<br/>⚠️ Validation warnings:<br/>${data.validationErrors.join('<br/>')}`;
    }
    setStatus(1, 'success', html);
    setBadge(1, 'done');

    // Show compact summary
    refsPreview.innerHTML = '';
    renderSupplierSummary(data.poCount || 0, data.bookingCount || 0, data.rowCount || 0);

    // Unlock pipeline card and activate stage
    const pipelineCard = document.getElementById('pipelineCard');
    if (pipelineCard) pipelineCard.classList.remove('locked');
    const badge = document.getElementById('badgePipeline');
    if (badge) badge.className = 'step-badge active';
    if (btnRunPipeline) btnRunPipeline.disabled = false;
  } catch (err) {
    const msg = err.message === 'Failed to fetch'
      ? '❌ Cannot reach the server. Please ensure the server is running (<code>npm start</code>) then try again.'
      : `❌ ${err.message}`;
    setStatus(1, 'error', msg);
    // Re-enable button so user can retry
    setLoading(btnParseSupplier, false);
    btnParseSupplier.disabled = false;
    return;
  }
  setLoading(btnParseSupplier, false);
  btnParseSupplier.disabled = true;
});

// ── Pipeline: single Run button ───────────────────────────────────────────────
function progSet(n, state, text) {
  const el = document.getElementById('prog' + n);
  if (!el) return;
  const colours = { pending: '#94a3b8', active: '#1D4ED8', done: '#15803D', error: '#B91C1C', warn: '#B45309', skip: '#9CA3AF' };
  el.style.color      = colours[state] || colours.pending;
  el.style.fontWeight = state === 'done' || state === 'active' ? '700' : 'normal';
  if (text) el.textContent = text;
}

// ── Cancel / Re-Submit Booking card ──────────────────────────────────────────
const btnLookupCancel = document.getElementById('btnLookupCancel');
const btnRunCancel    = document.getElementById('btnRunCancel');
const btnRunResub     = document.getElementById('btnRunResub');
const cancelPoInput   = document.getElementById('cancelPoInput');

function updateAmendBtns() {
  const panel = document.getElementById('cancelLookupPanel');
  const checked = panel ? [...panel.querySelectorAll('input[type=checkbox][name=cancelVbSelect]:checked')] : [];
  const n = checked.length;
  if (btnRunCancel) {
    btnRunCancel.disabled = n === 0;
    btnRunCancel.textContent = n > 1 ? `✕ Cancel ${n} Selected VBs & Upload` : '✕ Cancel Selected VB & Upload';
  }
  if (btnRunResub) {
    btnRunResub.disabled = n === 0;
    btnRunResub.textContent = n > 1 ? `♻ Re-Submit ${n} Selected VBs & Upload` : '♻ Re-Submit Selected VB & Upload';
  }
}

async function renderCancelLookup(inputs) {
  const panel = document.getElementById('cancelLookupPanel');
  if (!panel) return;
  if (btnRunCancel) btnRunCancel.disabled = true;
  if (btnRunResub)  btnRunResub.disabled  = true;
  panel.style.display = '';
  panel.innerHTML = '<span style="font-size:12px;color:#888">Looking up bookings…</span>';
  try {
    const res  = await fetch(`${API}/lookup-vbref?pos=${encodeURIComponent(inputs.join(','))}`);
    const data = await res.json();
    const refs = data.refs || {};

    // Aggregate all VBs across all inputs
    const allRows = []; // { bookingRef, poDisplay, timestamp, hasMasterRows }
    let notFound = [];
    for (const input of inputs) {
      const found = refs[input];
      if (!found) { notFound.push(input); continue; }
      const poDisplay = (found.poNumbers || []).join(', ') || input;
      const refList = found.allRefs || [{ bookingRef: found.bookingRef, timestamp: found.timestamp, hasMasterRows: found.hasMasterRows }];
      for (const r of refList) allRows.push({ ...r, poDisplay });
    }

    if (!allRows.length) {
      panel.innerHTML = `<span style="font-size:12px;color:#B91C1C">⚠️ No booking records found for: ${notFound.map(escapeHtml).join(', ')}. Must have been created with this tool.</span>`;
      return;
    }

    // Group rows by PO
    const byPo = {}; // poDisplay -> [rows]
    for (const r of allRows) {
      if (!byPo[r.poDisplay]) byPo[r.poDisplay] = [];
      byPo[r.poDisplay].push(r);
    }

    const thS = 'background:#7F1D1D;color:#fff;padding:6px 10px;text-align:left;font-size:12px';
    const tdS = 'padding:8px 10px;border-bottom:1px solid #FEE2E2;font-size:12px;vertical-align:middle';
    const poHeaderS = 'background:#FEF2F2;padding:6px 10px;font-size:12px;font-weight:700;color:#7F1D1D;border-bottom:1px solid #FECACA';
    let globalIdx = 0;
    const tbodyRows = Object.entries(byPo).map(([poDisplay, rList]) => {
      const poHeader = `<tr><td colspan="4" style="${poHeaderS}">PO: ${escapeHtml(poDisplay)}</td></tr>`;
      const vbRows = rList.map(r => {
        const i = globalIdx++;
        const badge = r.hasMasterRows
          ? '<span style="color:#1B5E20;font-size:10px;font-weight:700">✅ Ready</span>'
          : '<span style="color:#B45309;font-size:10px">⚠️ No stored data</span>';
        const cb = r.hasMasterRows
          ? `<input type="checkbox" name="cancelVbSelect" value="${escapeHtml(r.bookingRef)}" data-po="${escapeHtml(r.poDisplay)}" id="cvb${i}" style="accent-color:#991B1B;cursor:pointer;width:15px;height:15px">`
          : `<input type="checkbox" disabled style="opacity:0.4;width:15px;height:15px">`;
        return `<tr>
          <td style="${tdS};text-align:center;width:36px">${cb}</td>
          <td style="${tdS}"><label for="cvb${i}" style="font-family:monospace;font-weight:700;color:#1F4E79;cursor:pointer">${escapeHtml(r.bookingRef)}</label></td>
          <td style="${tdS}">${new Date(r.timestamp).toLocaleDateString('en-GB')} ${new Date(r.timestamp).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}</td>
          <td style="${tdS}">${badge}</td>
        </tr>`;
      }).join('');
      return poHeader + vbRows;
    }).join('');

    const notFoundNote = notFound.length
      ? `<p style="font-size:11px;color:#B91C1C;margin:6px 0 0">⚠️ Not found: ${notFound.map(escapeHtml).join(', ')}</p>`
      : '';

    panel.innerHTML = `
      <p style="font-size:12px;color:#666;margin:0 0 8px">↓ Select one VB per PO — then click Re-Submit or Cancel:</p>
      <table style="width:100%;border-collapse:collapse">
        <thead><tr>
          <th style="${thS};width:36px"></th>
          <th style="${thS}">VB Ref</th>
          <th style="${thS}">Booked On</th>
          <th style="${thS}">Status</th>
        </tr></thead>
        <tbody>${tbodyRows}</tbody>
      </table>${notFoundNote}`;

    // Auto-check if only one ready option overall
    const readyRows = allRows.filter(r => r.hasMasterRows);
    if (readyRows.length === 1) {
      const cb = panel.querySelector('input[type=checkbox]');
      if (cb) { cb.checked = true; }
    }
    panel.querySelectorAll('input[type=checkbox][name=cancelVbSelect]').forEach(cb => {
      cb.addEventListener('change', () => {
        if (cb.checked) {
          // Uncheck all other checkboxes in the same PO group
          panel.querySelectorAll(`input[type=checkbox][name=cancelVbSelect][data-po="${cb.dataset.po}"]`).forEach(other => {
            if (other !== cb) other.checked = false;
          });
        }
        updateAmendBtns();
      });
    });
    updateAmendBtns();
  } catch (err) {
    panel.innerHTML = `<span style="font-size:12px;color:#922B21">❌ ${err.message}</span>`;
  }
}

if (btnLookupCancel) {
  btnLookupCancel.addEventListener('click', async () => {
    const inputs = (cancelPoInput?.value || '').split(/[\n,]+/).map(s => s.trim()).filter(Boolean);
    if (!inputs.length) return;
    await renderCancelLookup(inputs);
  });
}

if (btnRunCancel) {
  btnRunCancel.addEventListener('click', () => runAmendAction('01'));
}

if (btnRunResub) {
  btnRunResub.addEventListener('click', () => runAmendAction('15'));
}

async function runAmendAction(purposeCd) {
  const isResub = purposeCd === '15';
  const btn = isResub ? btnRunResub : btnRunCancel;
    const panel = document.getElementById('cancelLookupPanel');
    const selectedRefs = panel
      ? [...panel.querySelectorAll('input[type=checkbox][name=cancelVbSelect]:checked')].map(cb => cb.value)
      : [];
    if (!selectedRefs.length) return;
  setLoading(btn, true);
  const statusMsg = isResub ? '⏳ Generating re-submission VBKREQ…' : '⏳ Generating cancellation VBKREQ…';
  psSetResult('cancelStatus', `<span style="font-size:12px;color:#888">${statusMsg}</span>`);
    try {
      const genRes  = await fetch(`${API}/cancel-booking`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ inputs: selectedRefs, purposeCd })
      });
      const genData = await genRes.json();
    if (!genRes.ok) throw new Error(genData.error || (isResub ? 'Re-submission generation failed' : 'Cancel generation failed'));
      const gens = genData.generations || [];
      const uploadResults = [];
      for (const gen of gens) {
        try {
          const upRes  = await fetch(`${API}/upload-sftp`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filename: gen.filename, xmlContent: gen.xml })
          });
          const upData = await upRes.json();
          if (!upRes.ok) throw new Error(upData.error || 'Upload failed');
          uploadResults.push({ filename: gen.filename, ok: true, localMode: upData.localMode, sftpEnv: upData.sftpEnv });
        } catch (err) {
          uploadResults.push({ filename: gen.filename, ok: false, error: err.message });
        }
      }
    renderAmendResult(gens, uploadResults, purposeCd);
      loadHistory();
    } catch (err) {
      psSetResult('cancelStatus', `<div style="color:#922B21;font-size:13px">❌ ${err.message}</div>`);
    } finally {
    setLoading(btn, false);
    }
}

function renderAmendResult(generations, uploadResults, purposeCd) {
  const isResub = purposeCd === '15';
  const upMap = Object.fromEntries((uploadResults || []).map(r => [r.filename, r]));
  const thS = 'background:#7F1D1D;color:#fff;padding:7px 10px;text-align:left;font-size:12px;white-space:nowrap';
  const tdS = 'padding:7px 10px;border-bottom:1px solid #FEE2E2;font-size:12px';
  const rows = generations.map(gen => {
    const pos   = (gen.poNumbers || []).map(p => `<div>${escapeHtml(p)}</div>`).join('') || '—';
    const vbRef = escapeHtml(gen.bookingRef || '—');
    const up    = upMap[gen.filename];
    const upBadge = up
      ? (up.ok
          ? `<span style="color:#1B5E20;font-size:11px;font-weight:700">✅ ${up.localMode ? 'Saved locally' : `Uploaded to <strong style="color:${up.sftpEnv==='PROD'?'#991B1B':'#0369A1'}">${up.sftpEnv||'SFTP'}</strong>`}</span>`
          : `<span style="color:#922B21;font-size:11px">❌ ${escapeHtml(up.error)}</span>`)
      : '';
    const blob   = new Blob([gen.xml], { type: 'application/xml' });
    const dlHref = URL.createObjectURL(blob);
    return `<tr>
      <td style="${tdS}">${pos}</td>
      <td style="${tdS};font-family:monospace;font-weight:700;color:#7F1D1D">${vbRef}</td>
      <td style="${tdS}"><a class="download-link" href="${dlHref}" download="${escapeHtml(gen.filename)}" style="font-size:11px;display:block;margin-bottom:3px">⬇ ${escapeHtml(gen.filename)}</a>${upBadge}</td>
    </tr>`;
  }).join('');
  const vbHeader = isResub ? 'VB Ref Re-Submitted' : 'VB Ref Cancelled';
  psSetResult('cancelStatus', `
    <table style="width:100%;border-collapse:collapse">
      <thead><tr>
        <th style="${thS}">PO Number(s)</th>
        <th style="${thS}">${vbHeader}</th>
        <th style="${thS}">${isResub ? 'Re-Submission XML' : 'Cancellation XML'}</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>`);
}

const btnStopPipeline = document.getElementById('btnStopPipeline');
let _pipelineAbort = null;

if (btnStopPipeline) {
  btnStopPipeline.addEventListener('click', () => {
    if (_pipelineAbort) _pipelineAbort.abort();
  });
}

if (btnRunPipeline) {
  btnRunPipeline.addEventListener('click', async () => {
    _pipelineAbort = new AbortController();
    const signal = _pipelineAbort.signal;

    setLoading(btnRunPipeline, true);
    if (btnStopPipeline) btnStopPipeline.style.display = '';
    psSetResult('psFetchResult', '');
    psSetResult('psBuildWarnings', '');
    psSetResult('resultPanel', '');

    const progress = document.getElementById('pipelineProgress');
    if (progress) progress.style.display = '';
    progSet(1, 'pending', '📡 Fetch ASN');
    progSet(2, 'pending', '🗂 Build');
    progSet(3, 'pending', '⚡ Generate');
    progSet(4, 'pending', '🚀 Upload');

    const purposeCd = '13'; // pipeline always submits new bookings
    const pipelineStart = Date.now();
    try {
      // ── 1. Fetch ASN ──────────────────────────────────────────────────────
      progSet(1, 'active', '📡 Fetching ASN…');
      const fetchRes  = await fetch(`${API}/fetch-feeds`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ poRefs: state.poRefs, asnRefs: [] }),
        signal
      });
      const fetchData = await fetchRes.json();
      if (!fetchRes.ok) throw new Error(fetchData.error || 'Fetch failed');

      const carrierAsnFiles  = fetchData.carrierAsnFiles || [];
      const cancelledItems   = fetchData.cancelledItems  || [];
      const foundPOs         = new Set(carrierAsnFiles.map(f => f.poRef).filter(Boolean));
      const skippedPOIds     = new Set(cancelledItems.map(c => c.poId));
      const notFoundCount    = state.poRefs.filter(po => !foundPOs.has(po) && !skippedPOIds.has(po)).length;
      const alreadyBooked    = cancelledItems.filter(c => c.type === 'ALREADY_BOOKED');
      const cancelled        = cancelledItems.filter(c => c.type !== 'ALREADY_BOOKED');
      const asnGridRows      = state.poRefs.map(po => {
        const found   = foundPOs.has(po);
        const skipped = skippedPOIds.has(po);
        const item    = cancelledItems.find(c => c.poId === po);
        if (skipped && item?.type === 'ALREADY_BOOKED') {
          return `<div class="asn-status-row not-found">📋 <strong>PO ${po}</strong> — ASN ${item.asnId} already has a carrier booking${item.vbRef ? ` (<strong>${escapeHtml(item.vbRef)}</strong>)` : ''} — skipped</div>`;
        }
        if (skipped) {
          const label = item?.type === 'ASN' ? `ASN ${item.asnId} cancelled` : `PO cancelled (Status=C)`;
          return `<div class="asn-status-row not-found">🚫 <strong>PO ${po}</strong> — ${label} — booking skipped</div>`;
        }
        return `<div class="asn-status-row ${found ? 'found' : 'not-found'}">${found ? '✅' : '⚠️'} <strong>PO ${po}</strong> — ${found ? 'ASN found in Databricks' : 'No ASN found in Databricks'}</div>`;
      }).join('');
      const skippedSections = [
        alreadyBooked.length ? `<div style="margin-top:8px;padding:8px;background:#EFF6FF;border-left:3px solid #2563EB;border-radius:4px;font-size:12px">📋 <strong>${alreadyBooked.length} ASN(s) skipped — booking already exists:</strong><br/>${alreadyBooked.map(c => `&nbsp;• ${escapeHtml(c.reason)}`).join('<br/>')}</div>` : '',
        cancelled.length    ? `<div style="margin-top:8px;padding:8px;background:#FEF2F2;border-left:3px solid #DC2626;border-radius:4px;font-size:12px">🚫 <strong>${cancelled.length} ASN/PO(s) skipped — cancelled:</strong><br/>${cancelled.map(c => `&nbsp;• ${escapeHtml(c.reason)}`).join('<br/>')}</div>` : ''
      ].join('');
      const summaryExtras = [
        notFoundCount   ? ` &nbsp;⚠️ ${notFoundCount} PO(s) with no ASN in Databricks` : '',
        alreadyBooked.length ? ` &nbsp;📋 ${alreadyBooked.length} already booked` : '',
        cancelled.length     ? ` &nbsp;🚫 ${cancelled.length} cancelled` : ''
      ].join('');
      psSetResult('psFetchResult', `<details style="font-size:12px">
        <summary style="cursor:pointer;user-select:none;list-style:none;display:flex;align-items:center;gap:6px;padding:2px 0">
          <span style="font-size:10px;color:#888">▶</span>
          ✅ <strong>${carrierAsnFiles.length}</strong> Databricks ASN record(s) fetched${summaryExtras} <span style="font-size:11px;color:#888;font-style:italic">— click to expand</span>
        </summary>
        <div class="asn-status-grid" style="margin-top:8px">${asnGridRows}</div>
      </details>`);

      // Count truly active (non-cancelled) ASN groups across all returned files
      const totalActiveGroups = carrierAsnFiles.reduce((sum, f) => sum + (f.asnGroups || []).length, 0);

      if (carrierAsnFiles.length === 0 && cancelledItems.length === 0) {
        progSet(1, 'error', '📡 No ASN found');
        throw new Error(`No active ASN records found in Databricks for any of the submitted POs. Cannot proceed.`);
      }

      const doneExtras = [alreadyBooked.length ? `${alreadyBooked.length} already booked skipped` : '', cancelled.length ? `${cancelled.length} cancelled skipped` : ''].filter(Boolean).join(', ');
      const allSkipped = totalActiveGroups === 0;
      progSet(1, allSkipped ? 'warn' : 'done', `📡 ASN ${allSkipped ? '⚠️' : '✅'}${doneExtras ? ` (${doneExtras})` : ''}`);
      state.feedsFetched = true;

      // All POs were cancelled or already booked — skip remaining stages
      if (allSkipped) {
        progSet(2, 'skip', '🗂 Build — skipped');
        progSet(3, 'skip', '⚡ Generate — skipped');
        progSet(4, 'skip', '🚀 Upload — skipped');
        const badge = document.getElementById('badgePipeline');
        if (badge) { badge.className = 'step-badge'; badge.style.background = '#B45309'; badge.style.boxShadow = 'none'; badge.textContent = '⚠'; }
        return;
      }

      // ── 2. Build ──────────────────────────────────────────────────────────
      progSet(2, 'active', '🗂 Building…');
      const buildRes  = await fetch(`${API}/build-bible`, { method: 'POST', signal });
      const buildData = await buildRes.json();
      if (!buildRes.ok) throw new Error(buildData.error || 'Build failed');
      state.biblBuilt = true;
      let warningsHtml = '';
      if (buildData.warnings?.length) {
        warningsHtml += `<div style="margin-bottom:8px;padding:8px;background:#FEF9E7;border-left:3px solid #F39C12;border-radius:4px;font-size:12px">⚠️ <strong>${buildData.warnings.length} SKU(s) excluded</strong> — not on carrier ASN:<br/>${buildData.warnings.map(w => `&nbsp;• ${w}`).join('<br/>')}</div>`;
      }

      psSetResult('psBuildWarnings', warningsHtml);
      progSet(2, 'done', '🗂 Built ✅');

      // ── 3. Generate ───────────────────────────────────────────────────────
      progSet(3, 'active', '⚡ Generating…');
      const genRes  = await fetch(`${API}/generate-vbkreq`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ purposeCd }),
        signal
      });
      const genData = await genRes.json();
      if (!genRes.ok) throw new Error(genData.error || 'Generate failed');

      state.generations  = genData.generations || [];
      state.lastXml      = state.generations[0]?.xml      || null;
      state.lastFilename = state.generations[0]?.filename || null;

      progSet(3, 'done', '⚡ Generated ✅');

      // ── 4. Upload ─────────────────────────────────────────────────────────
      progSet(4, 'active', '🚀 Uploading…');
      const results = await Promise.all(
        state.generations.map(async gen => {
          try {
            const upRes  = await fetch(`${API}/upload-sftp`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ filename: gen.filename, xmlContent: gen.xml }),
              signal
            });
            const upData = await upRes.json();
            if (!upRes.ok) throw new Error(upData.error || 'Upload failed');
            return { filename: gen.filename, ok: true, remotePath: upData.remotePath, localMode: upData.localMode, sftpEnv: upData.sftpEnv };
          } catch (err) {
            if (err.name === 'AbortError') throw err; // bubble up so outer catch handles it
            return { filename: gen.filename, ok: false, error: err.message };
          }
        })
      );
      const ok   = results.filter(r => r.ok);
      const fail = results.filter(r => !r.ok);
      progSet(4, fail.length === 0 ? 'done' : 'error', fail.length === 0 ? '🚀 Uploaded ✅' : '🚀 Upload ⚠️');

      const badge = document.getElementById('badgePipeline');
      if (badge) { badge.className = 'step-badge ' + (fail.length === 0 ? 'done' : 'active'); if (fail.length === 0) badge.textContent = '✓'; }
      renderResult(state.generations, results, Date.now() - pipelineStart);
      loadHistory();

    } catch (err) {
      if (err.name === 'AbortError') {
        const stageNames = ['📡 Fetch ASN', '🗂 Build', '⚡ Generate', '🚀 Upload'];
        for (let i = 1; i <= 4; i++) {
          const el = document.getElementById('prog' + i);
          if (el && el.style.color === 'rgb(29, 78, 216)') progSet(i, 'error', stageNames[i-1] + ' ⏹');
        }
        psSetResult('psFetchResult', (document.getElementById('psFetchResult').innerHTML || '') +
          `<div style="color:#6B7280;font-size:13px;margin-top:6px">⏹ Pipeline stopped by user.</div>`);
      } else {
        psSetResult('psFetchResult', (document.getElementById('psFetchResult').innerHTML || '') +
          `<div style="color:#922B21;font-size:13px;margin-top:6px">❌ ${err.message}</div>`);
      }
    } finally {
      if (btnStopPipeline) btnStopPipeline.style.display = 'none';
      _pipelineAbort = null;
      setLoading(btnRunPipeline, false);
    }
  });
}

function renderSupplierSummary(poCount, bookingCount, skuRowCount) {
  const panel = document.getElementById('supplierParsePanel');
  if (!panel) return;
  panel.innerHTML = `
    <div style="display:flex;gap:16px;align-items:center;flex-wrap:wrap;padding:10px 4px">
      <div class="parse-stat-badge" style="font-size:14px;padding:8px 18px">📦 <strong>${poCount}</strong> PO${poCount !== 1 ? 's' : ''} parsed</div>
      <div class="parse-stat-badge" style="font-size:14px;padding:8px 18px">🚚 <strong>${bookingCount}</strong> carrier booking${bookingCount !== 1 ? 's' : ''} will be generated</div>
    </div>`;
  panel.style.display = 'block';
}

// ── Supplier parse summary ────────────────────────────────────────────────────
function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ── Pipeline result table ─────────────────────────────────────────────────────
function renderResult(generations, uploadResults, elapsedMs) {
  const panel = document.getElementById('resultPanel');
  if (!panel || !generations?.length) return;

  const thStyle = 'background:#1F4E79;color:#fff;padding:7px 10px;text-align:left;font-size:12px;white-space:nowrap';
  const tdStyle = 'padding:7px 10px;border-bottom:1px solid #e2e8f0;vertical-align:top;font-size:12px';
  const upMap   = Object.fromEntries((uploadResults || []).map(r => [r.filename, r]));

  const bookedPOs = new Set(generations.flatMap(g => g.poNumbers || []));
  const totalPOs  = (state.poRefs || []).length;
  const allBooked = totalPOs > 0 && bookedPOs.size >= totalPOs;
  const elapsed   = elapsedMs != null
    ? (elapsedMs < 60000
        ? `${(elapsedMs / 1000).toFixed(1)}s`
        : `${Math.floor(elapsedMs / 60000)}m ${Math.round((elapsedMs % 60000) / 1000)}s`)
    : null;
  const timeTag   = elapsed ? ` <span style="font-weight:400;opacity:.7">(${elapsed})</span>` : '';
  const summaryBanner = allBooked
    ? `<div style="margin-bottom:10px;padding:8px 14px;background:#ECFDF5;border-left:3px solid #059669;border-radius:4px;font-size:12px;color:#065F46">✅ All <strong>${totalPOs}</strong> PO${totalPOs !== 1 ? 's' : ''} are booked.${timeTag}</div>`
    : `<div style="margin-bottom:10px;padding:8px 14px;background:#FEF9E7;border-left:3px solid #D97706;border-radius:4px;font-size:12px;color:#92400E">⚠️ <strong>${bookedPOs.size}</strong> of <strong>${totalPOs}</strong> PO${totalPOs !== 1 ? 's' : ''} booked.${timeTag}</div>`;

  const rows = generations.map(gen => {
    const pos    = (gen.poNumbers || []).map(p => `<div>${escapeHtml(p)}</div>`).join('') || '—';
    const asns   = (gen.asnRefs   || []).map(a => `<div style="font-family:monospace">${escapeHtml(a)}</div>`).join('') || '<span style="color:#aaa">—</span>';
    const vbRef  = escapeHtml(gen.bookingRef || gen.ctrlNumber || '—');
    const up     = upMap[gen.filename];
    const upBadge = up
      ? (up.ok
          ? `<span style="color:#1B5E20;font-size:11px;font-weight:700">✅ ${up.localMode ? 'Saved locally' : `Uploaded to <strong style="color:${up.sftpEnv==='PROD'?'#991B1B':'#0369A1'}">${up.sftpEnv||'SFTP'}</strong>`}</span>`
          : `<span style="color:#922B21;font-size:11px">❌ ${escapeHtml(up.error)}</span>`)
      : '';
    const blob   = new Blob([gen.xml], { type: 'application/xml' });
    const dlHref = URL.createObjectURL(blob);
    return `<tr>
      <td style="${tdStyle}">${pos}</td>
      <td style="${tdStyle}">${asns}</td>
      <td style="${tdStyle};font-family:monospace;font-weight:700;color:#1F4E79">${vbRef}</td>
      <td style="${tdStyle}"><a class="download-link" href="${dlHref}" download="${escapeHtml(gen.filename)}" style="font-size:11px;display:block;margin-bottom:3px">⬇ ${escapeHtml(gen.filename)}</a>${upBadge}</td>
    </tr>`;
  }).join('');

  const fileNames = state.supplierFileNames || [];
  const taggedLinks = fileNames.length
    ? fileNames.map((name, idx) => {
        const taggedName = name.replace(/\.xlsx?$/i, '') + '_VBRef.xlsx';
        return `<a class="download-link" style="margin-right:12px;font-size:12px" href="${API}/tagged-supplier/${idx}" download="${escapeHtml(taggedName)}">⬇ ${escapeHtml(taggedName)}</a>`;
      }).join('')
    : '';

  const genCount = generations.length;
  panel.innerHTML = `
    <details>
      <summary style="cursor:pointer;user-select:none;list-style:none;display:flex;align-items:center;gap:6px;padding:2px 0;font-size:12px;margin-bottom:6px">
        <span style="font-size:10px;color:#888">▶</span>
        <strong>${genCount}</strong> carrier booking message${genCount !== 1 ? 's' : ''} generated
        <span style="font-size:11px;color:#888;font-style:italic">— click to expand</span>
      </summary>
      ${summaryBanner}
      <table style="width:100%;border-collapse:collapse;margin-top:8px;margin-bottom:${taggedLinks ? '10' : '0'}px">
        <thead><tr>
          <th style="${thStyle}">PO Number(s)</th>
          <th style="${thStyle}">ASN Ref(s)</th>
          <th style="${thStyle}">VB Ref</th>
          <th style="${thStyle}">VBKREQ File</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      ${taggedLinks ? `<div style="font-size:12px"><strong>Tagged supplier template(s) with VB Ref:</strong><div style="margin-top:6px">${taggedLinks}</div></div>` : ''}
    </details>`;
}

async function refreshLog() {} // no-op

// ── Recent Booking History ────────────────────────────────────────────────────
async function loadHistory() {
  const panel = document.getElementById('historyPanel');
  if (!panel) return;
  try {
    const res  = await fetch(`${API}/generation-log`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Failed to load history');

    const entries = data.entries || [];
    if (!entries.length) {
      panel.innerHTML = '<span style="font-size:12px;color:#aaa">No bookings generated in the last 3 days.</span>';
      return;
    }

    const thStyle = 'background:#374151;color:#fff;padding:6px 10px;text-align:left;font-size:12px;white-space:nowrap';
    const tdStyle = 'padding:6px 10px;border-bottom:1px solid #e2e8f0;vertical-align:top;font-size:12px';

    const rows = entries.map(e => {
      const ts      = e.timestamp ? new Date(e.timestamp).toLocaleString('en-GB') : '—';
      const vbRef   = escapeHtml(e.bookingRef || e.ctrlNumber || '—');
      const pos     = (e.poNumbers || []).map(p => `<div>${escapeHtml(p)}</div>`).join('') || '—';
      const asns    = (e.asnRefs   || []).map(a => `<div style="font-family:monospace">${escapeHtml(a)}</div>`).join('') || '<span style="color:#aaa">—</span>';
      const dlLink  = e.filename
        ? `<a class="download-link" href="/output/${encodeURIComponent(e.filename)}" download="${escapeHtml(e.filename)}" style="font-size:11px">${escapeHtml(e.filename)}</a>`
        : '—';
      const sftpBadge = e.sftp === 'uploaded'
        ? `<span style="color:#1B5E20;font-weight:700;font-size:11px">✅ Uploaded <span style="font-size:10px;padding:1px 5px;border-radius:8px;background:${e.sftpEnv==='PROD'?'#FEE2E2;color:#991B1B':'#DBEAFE;color:#1D4ED8'}">${e.sftpEnv||'SFTP'}</span></span>`
        : e.sftp === 'local'
          ? `<span style="color:#784212;font-size:11px">📁 Local</span>`
          : `<span style="color:#aaa;font-size:11px">—</span>`;
      const uploadedAt = e.uploadedAt ? new Date(e.uploadedAt).toLocaleString('en-GB') : '—';
      return `<tr>
        <td style="${tdStyle};white-space:nowrap">${ts}</td>
        <td style="${tdStyle};font-family:monospace;font-weight:700;color:#1F4E79">${vbRef}</td>
        <td style="${tdStyle}">${pos}</td>
        <td style="${tdStyle}">${asns}</td>
        <td style="${tdStyle}">${sftpBadge}<div style="font-size:10px;color:#888;margin-top:2px">${uploadedAt}</div></td>
        <td style="${tdStyle}">${dlLink}</td>
      </tr>`;
    }).join('');

    panel.innerHTML = `
      <table style="width:100%;border-collapse:collapse">
        <thead><tr>
          <th style="${thStyle}">Generated</th>
          <th style="${thStyle}">VB Ref</th>
          <th style="${thStyle}">PO(s)</th>
          <th style="${thStyle}">ASN(s)</th>
          <th style="${thStyle}">Upload Status</th>
          <th style="${thStyle}">Download XML</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;
  } catch (err) {
    if (panel) panel.innerHTML = `<span style="font-size:12px;color:#922B21">❌ ${err.message}</span>`;
  }
}

// Load history on page start and after each upload
loadHistory();

// ── SharePoint auto-sync status banner ───────────────────────────────────────
const spBanner    = document.getElementById('spSyncBanner');
const spStatus    = document.getElementById('spSyncStatus');
const spFiles     = document.getElementById('spSyncFiles');
const btnSpSync   = document.getElementById('btnSpSyncNow');
let _spPollTimer  = null;

async function loadSpStatus() {
  try {
    const res  = await fetch(`${API}/sharepoint/status`);
    const data = await res.json();

    if (!data.configured) {
      if (spBanner) spBanner.style.display = 'none';
      return;
    }

    if (spBanner) spBanner.style.display = '';

    if (data.running) {
      if (spStatus) spStatus.textContent = '⏳ Syncing…';
      // Poll every 3s while running
      if (!_spPollTimer) _spPollTimer = setInterval(loadSpStatus, 3000);
      return;
    }

    // Not running — clear poll timer
    if (_spPollTimer) { clearInterval(_spPollTimer); _spPollTimer = null; }

    const scheduleEl = document.getElementById('spScheduleLabel');
    if (scheduleEl) scheduleEl.textContent = data.schedule ? `⏰ ${data.schedule}` : '';
    const dismissBtn = document.getElementById('btnDismissSpError');
    if (data.error) {
      // Extract a short friendly message from the raw error
      const raw = data.error;
      let friendly = raw;
      if (raw.includes('401'))      friendly = 'SharePoint connection failed — not authorised (401). Permission may still be propagating.';
      else if (raw.includes('404')) friendly = 'SharePoint folder not found (404). Check SP_FOLDER_PATH in .env.';
      else if (raw.includes('403')) friendly = 'Access denied to SharePoint (403). Check app permissions.';
      else if (raw.includes('List failed')) friendly = `Could not list SharePoint files. ${raw.split(':')[0]}.`;
      if (spStatus) spStatus.innerHTML =
        `<span style="color:#B91C1C;font-weight:600">⚠ ${escapeHtml(friendly)}</span>` +
        `<details style="display:inline;margin-left:6px"><summary style="display:inline;cursor:pointer;font-size:11px;color:#9CA3AF">details</summary>` +
        `<div style="font-size:10px;color:#6B7280;margin-top:4px;word-break:break-all">${escapeHtml(raw)}</div></details>` +
        `<button onclick="dismissSpError()" style="display:inline-block;margin-left:10px;padding:2px 10px;font-size:11px;background:#FEE2E2;color:#B91C1C;border:1px solid #FECACA;border-radius:5px;cursor:pointer">✕ Dismiss</button>`;
    } else {
      if (dismissBtn) dismissBtn.style.display = 'none';
    }
    if (!data.error && data.lastSync) {
      const when = new Date(data.lastSync).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
      const date = new Date(data.lastSync).toLocaleDateString('en-GB');
      const skippedNote = data.skipped ? ' &nbsp;<span style="color:#92400E">(no changes — skipped)</span>' : '';
      if (spStatus) spStatus.innerHTML = `Last synced: <strong>${when} on ${date}</strong>${skippedNote} &nbsp;·&nbsp; ${data.rowCount || 0} row(s), ${(data.poRefs || []).length} PO(s)`;
    } else if (!data.error) {
      if (spStatus) spStatus.innerHTML = `Not yet synced`;
    }

    // Show file list
    if (spFiles && data.files && data.files.length) {
      spFiles.innerHTML = data.files.map(f => {
        const kb  = Math.round((f.size || 0) / 1024);
        const mod = f.lastModified ? new Date(f.lastModified).toLocaleDateString('en-GB') : '';
        return `<span style="display:inline-block;margin-right:10px;color:#0369A1">📄 ${escapeHtml(f.name)} <span style="color:#888">(${kb} KB${mod ? ', ' + mod : ''})</span></span>`;
      }).join('');
    } else if (spFiles) {
      spFiles.innerHTML = '';
    }

    // Show run history
    const histDetails = document.getElementById('spHistoryDetails');
    const histBody    = document.getElementById('spHistoryBody');
    if (histDetails && histBody && data.syncHistory && data.syncHistory.length) {
      histDetails.style.display = '';
      const outcomeLabel = { synced: '✅ Synced', skipped: '⏭ Skipped', no_files: '📭 No files', error: '❌ Error' };
      const outcomeColour = { synced: '#15803D', skipped: '#92400E', no_files: '#6B7280', error: '#B91C1C' };
      histBody.innerHTML = data.syncHistory.map((h, i) => {
        const ts    = new Date(h.timestamp);
        const date  = ts.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        const time  = ts.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
        const label = outcomeLabel[h.outcome] || h.outcome;
        const colour = outcomeColour[h.outcome] || '#555';
        const files = h.outcome === 'error'
          ? `<span style="color:#B91C1C">${escapeHtml(h.error || '')}</span>`
          : (h.files || []).map(f => escapeHtml(f)).join(', ') || '—';
        const bg = i % 2 === 0 ? '#F0F9FF' : '#fff';
        return `<tr style="background:${bg}">
          <td style="padding:4px 8px;white-space:nowrap;color:#374151">${date} ${time}</td>
          <td style="padding:4px 8px;font-weight:600;color:${colour}">${label}</td>
          <td style="padding:4px 8px;text-align:right;color:#374151">${h.outcome === 'synced' ? h.rowCount : '—'}</td>
          <td style="padding:4px 8px;text-align:right;color:#374151">${h.outcome === 'synced' ? h.poCount : '—'}</td>
          <td style="padding:4px 8px;color:#555">${files}</td>
        </tr>`;
      }).join('');
    } else if (histDetails) {
      histDetails.style.display = 'none';
    }

    // If a fresh sync just populated supplier data, update the PO tags
    if (data.lastSync && data.poRefs && data.poRefs.length && !state.poRefs.length) {
      state.poRefs = data.poRefs;
      renderPoTags(data.poRefs);
      document.getElementById('pipelineCard')?.classList.remove('locked');
      const badge = document.getElementById('badgePipeline');
      if (badge) badge.className = 'step-badge active';
    }
  } catch (_) {
    // silently ignore if server not up yet
  }
}

if (btnSpSync) {
  btnSpSync.addEventListener('click', async () => {
    setLoading(btnSpSync, true);
    try {
      await fetch(`${API}/sharepoint/sync`, { method: 'POST' });
      if (spStatus) spStatus.textContent = '⏳ Syncing…';
      if (!_spPollTimer) _spPollTimer = setInterval(loadSpStatus, 3000);
    } catch (err) {
      if (spStatus) spStatus.textContent = `❌ ${err.message}`;
    } finally {
      setLoading(btnSpSync, false);
    }
  });
}

async function dismissSpError() {
  try {
    await fetch(`${API}/sharepoint/dismiss-error`, { method: 'POST' });
    loadSpStatus();
  } catch (_) {}
}

// Initial SP status check
loadSpStatus();
