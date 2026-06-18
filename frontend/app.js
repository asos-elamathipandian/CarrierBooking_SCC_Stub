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
const xmlPreviewWrap     = document.getElementById('xmlPreviewWrap');
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
dropZone.addEventListener('click', () => supplierFileInput.click());

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

    // Warn if SKU_LINES is empty but header has POs
    if ((data.rowCount || 0) === 0 && (data.poCount || 0) > 0) {
      setStatus(1, 'info',
        `ℹ️ Found <strong>${data.poCount}</strong> PO(s) in PO Header, but <strong>PO Lines</strong> is empty. ` +
        'Fill in PO Lines before proceeding to fetch feeds and generate bookings.'
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

    // Show pipeline and activate stage 1
    document.getElementById('pipelineCard').style.display = '';
    const badge = document.getElementById('badgePipeline');
    if (badge) badge.className = 'step-badge active';
    if (btnRunPipeline) btnRunPipeline.disabled = false;
    wireDateClear('overrideCargoReady',    'clearCargoReady');
    wireDateClear('overrideBookingReqDate','clearBookingReqDate');
  } catch (err) {
    setStatus(1, 'error', `❌ ${err.message}`);
  } finally {
    setLoading(btnParseSupplier, false);
    btnParseSupplier.disabled = true;
  }
});

// ── Pipeline: single Run button ───────────────────────────────────────────────
function progSet(n, state, text) {
  const el = document.getElementById('prog' + n);
  if (!el) return;
  const colours = { pending: '#94a3b8', active: '#1D4ED8', done: '#15803D', error: '#B91C1C' };
  el.style.color      = colours[state] || colours.pending;
  el.style.fontWeight = state === 'done' || state === 'active' ? '700' : 'normal';
  if (text) el.textContent = text;
}

function wireDateClear(inputId, clearBtnId) {
  const input = document.getElementById(inputId);
  const btn   = document.getElementById(clearBtnId);
  if (!input || !btn) return;
  input.addEventListener('change', () => btn.classList.toggle('visible', !!input.value));
}
function clearDateField(inputId, clearBtnId) {
  const input = document.getElementById(inputId);
  const btn   = document.getElementById(clearBtnId);
  if (input) input.value = '';
  if (btn)   btn.classList.remove('visible');
}
window.clearDateField = clearDateField;

if (btnRunPipeline) {
  btnRunPipeline.addEventListener('click', async () => {
    setLoading(btnRunPipeline, true);
    psSetResult('psFetchResult', '');
    psSetResult('psBuildWarnings', '');
    psSetResult('psGenResult', '');
    psSetResult('psUploadResult', '');
    document.getElementById('generationsContainer').innerHTML = '';
    if (xmlPreviewWrap) xmlPreviewWrap.classList.remove('visible');

    const progress = document.getElementById('pipelineProgress');
    if (progress) progress.style.display = '';
    progSet(1, 'pending', '📡 Fetch ASN');
    progSet(2, 'pending', '🗂 Build');
    progSet(3, 'pending', '⚡ Generate');
    progSet(4, 'pending', '🚀 Upload');

    const purposeCd    = document.querySelector('input[name="purposeCd"]:checked')?.value || '13';
    const purposeLabel = { '13': 'Submission', '15': 'Re-Submission', '01': 'Cancellation' }[purposeCd] || purposeCd;
    const overrideCargoReady     = document.getElementById('overrideCargoReady')?.value    || '';
    const overrideBookingReqDate = document.getElementById('overrideBookingReqDate')?.value || '';

    try {
      // ── 1. Fetch ASN ──────────────────────────────────────────────────────
      progSet(1, 'active', '📡 Fetching ASN…');
      const fetchRes  = await fetch(`${API}/fetch-feeds`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ poRefs: state.poRefs, asnRefs: [] })
      });
      const fetchData = await fetchRes.json();
      if (!fetchRes.ok) throw new Error(fetchData.error || 'Fetch failed');

      const carrierAsnFiles = fetchData.carrierAsnFiles || [];
      const foundPOs        = new Set(carrierAsnFiles.map(f => f.poRef).filter(Boolean));
      const notFoundCount   = state.poRefs.filter(po => !foundPOs.has(po)).length;
      const asnGridRows     = state.poRefs.map(po => {
        const found = foundPOs.has(po);
        return `<div class="asn-status-row ${found ? 'found' : 'not-found'}">${found ? '✅' : '⚠️'} <strong>PO ${po}</strong> — ${found ? 'ASN found' : 'No ASN found'}</div>`;
      }).join('');
      psSetResult('psFetchResult', `<div style="font-size:12px;margin-bottom:6px">✅ <strong>${carrierAsnFiles.length}</strong> carrier ASN file(s) fetched${notFoundCount ? ` &nbsp;⚠️ ${notFoundCount} PO(s) with no ASN` : ''}.</div><div class="asn-status-grid">${asnGridRows}</div>`);

      if (carrierAsnFiles.length === 0) {
        progSet(1, 'error', '📡 No ASN found');
        throw new Error('No carrier ASN files found for any PO. Cannot proceed.');
      }
      progSet(1, 'done', '📡 ASN ✅');
      state.feedsFetched = true;

      // ── 2. Build ──────────────────────────────────────────────────────────
      progSet(2, 'active', '🗂 Building…');
      const buildRes  = await fetch(`${API}/build-bible`, { method: 'POST' });
      const buildData = await buildRes.json();
      if (!buildRes.ok) throw new Error(buildData.error || 'Build failed');
      state.biblBuilt = true;
      let warningsHtml = '';
      if (buildData.warnings?.length) {
        warningsHtml += `<div style="margin-bottom:8px;padding:8px;background:#FEF9E7;border-left:3px solid #F39C12;border-radius:4px;font-size:12px">⚠️ <strong>${buildData.warnings.length} SKU(s) excluded</strong> — not on carrier ASN:<br/>${buildData.warnings.map(w => `&nbsp;• ${w}`).join('<br/>')}</div>`;
      }
      if (buildData.extraSkuWarnings?.length) {
        warningsHtml += `<div style="margin-bottom:8px;padding:8px;background:#FEF9E7;border-left:3px solid #F39C12;border-radius:4px;font-size:12px">⚠️ <strong>${buildData.extraSkuWarnings.length} extra carrier ASN SKU(s)</strong> not in supplier template (Booking_Qty=0):<br/>${buildData.extraSkuWarnings.map(w => `&nbsp;• ${w}`).join('<br/>')}</div>`;
      }
      psSetResult('psBuildWarnings', warningsHtml);
      progSet(2, 'done', '🗂 Built ✅');

      // ── 3. Generate ───────────────────────────────────────────────────────
      progSet(3, 'active', '⚡ Generating…');
      const genRes  = await fetch(`${API}/generate-vbkreq`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ purposeCd, overrideCargoReady, overrideBookingReqDate })
      });
      const genData = await genRes.json();
      if (!genRes.ok) throw new Error(genData.error || 'Generate failed');

      state.generations  = genData.generations || [];
      state.lastXml      = state.generations[0]?.xml      || null;
      state.lastFilename = state.generations[0]?.filename || null;
      const count = state.generations.length;
      psSetResult('psGenResult', `<div style="font-size:13px;color:#1E8449">✅ <strong>${count}</strong> VBKREQ${count > 1 ? 's' : ''} generated — ${purposeLabel}</div>`);
      renderGenerations(state.generations);
      if (xmlPreviewWrap) xmlPreviewWrap.classList.add('visible');
      progSet(3, 'done', '⚡ Generated ✅');

      // ── 4. Upload ─────────────────────────────────────────────────────────
      progSet(4, 'active', '🚀 Uploading…');
      psSetResult('psUploadResult', `<span style="font-size:12px;color:#784212">⏳ Uploading ${count} VBKREQ(s)…</span>`);
      const results = [];
      for (let i = 0; i < state.generations.length; i++) {
        const gen = state.generations[i];
        try {
          const upRes  = await fetch(`${API}/upload-sftp`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ filename: gen.filename, xmlContent: gen.xml })
          });
          const upData = await upRes.json();
          if (!upRes.ok) throw new Error(upData.error || 'Upload failed');
          results.push({ filename: gen.filename, ok: true, remotePath: upData.remotePath, localMode: upData.localMode });
          const btn = document.querySelector(`.gen-upload-btn[data-idx="${i}"]`);
          if (btn) { btn.textContent = '✅ Uploaded'; btn.style.background = '#27AE60'; btn.disabled = true; }
        } catch (err) {
          results.push({ filename: gen.filename, ok: false, error: err.message });
        }
      }
      const ok   = results.filter(r => r.ok);
      const fail = results.filter(r => !r.ok);
      let upHtml = `<div style="font-size:13px">✅ <strong>${ok.length}</strong> of <strong>${results.length}</strong> uploaded to SFTP.</div>`;
      if (ok.length)   upHtml += '<div style="font-size:12px;margin-top:4px">' + ok.map(r => `&nbsp;• ${r.filename} → ${r.localMode ? 'local output/' : r.remotePath}`).join('<br/>') + '</div>';
      if (fail.length) upHtml += `<div style="font-size:12px;color:#922B21;margin-top:4px">❌ ${fail.length} failed:<br/>` + fail.map(r => `&nbsp;• ${r.filename}: ${r.error}`).join('<br/>') + '</div>';
      psSetResult('psUploadResult', upHtml);
      progSet(4, fail.length === 0 ? 'done' : 'error', fail.length === 0 ? '🚀 Uploaded ✅' : '🚀 Upload ⚠️');

      const badge = document.getElementById('badgePipeline');
      if (badge) { badge.className = 'step-badge ' + (fail.length === 0 ? 'done' : 'active'); if (fail.length === 0) badge.textContent = '✓'; }
      showTaggedTemplateDownload(state.generations || []);
      loadHistory();

    } catch (err) {
      psSetResult('psFetchResult', (document.getElementById('psFetchResult').innerHTML || '') +
        `<div style="color:#922B21;font-size:13px;margin-top:6px">❌ ${err.message}</div>`);
    } finally {
      setLoading(btnRunPipeline, false);
    }
  });
}

function renderSupplierSummary(poCount, bookingCount, skuRowCount) {
  const panel = document.getElementById('supplierParsePanel');
  if (!panel) return;
  const skuNote = skuRowCount === 0
    ? `<div style="font-size:12px;color:#784212;margin-top:6px">⚠️ PO Lines sheet is empty — fill it in before generating bookings.</div>`
    : `<div class="parse-stat-badge" style="font-size:14px;padding:8px 18px">📋 <strong>${skuRowCount}</strong> SKU row${skuRowCount !== 1 ? 's' : ''}</div>`;
  panel.innerHTML = `
    <div style="display:flex;gap:16px;align-items:center;flex-wrap:wrap;padding:10px 4px">
      <div class="parse-stat-badge" style="font-size:14px;padding:8px 18px">📦 <strong>${poCount}</strong> PO${poCount !== 1 ? 's' : ''} parsed</div>
      <div class="parse-stat-badge" style="font-size:14px;padding:8px 18px">🚚 <strong>${bookingCount}</strong> carrier booking${bookingCount !== 1 ? 's' : ''} will be generated</div>
      ${skuNote}
    </div>`;
  panel.style.display = 'block';
}

// ── Supplier parse summary ────────────────────────────────────────────────────
function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderGenerations(generations) {
  const container = document.getElementById('generationsContainer');
  container.innerHTML = '';
  generations.forEach((gen, i) => {
    const groupLabel = gen.group
      ? `Group: <strong>${gen.group}</strong>`
      : (generations.length > 1 ? `Booking ${i + 1}` : 'Generated VBKREQ');
    const poList = (gen.poNumbers || []).join(', ') || '—';
    const card = document.createElement('div');
    card.className = 'generation-card';
    card.innerHTML = `
      <div class="gen-header">
        <span class="gen-title">${groupLabel}</span>
        <span class="gen-meta">POs: ${poList} &nbsp;|&nbsp;Ctrl#: ${gen.ctrlNumber} &nbsp;|&nbsp;v${gen.version || 1}</span>
        <a class="download-link" href="#" download="${gen.filename}">⬇ ${gen.filename}</a>
        <button class="btn btn-danger btn-sm gen-upload-btn" data-idx="${i}">⬆ Upload SFTP</button>
      </div>
      <details class="gen-details">
        <summary>▶ Preview XML</summary>
        <pre class="xml-pre">${escapeHtml(gen.xml)}</pre>
      </details>
    `;
    const blob = new Blob([gen.xml], { type: 'application/xml' });
    card.querySelector('a').href = URL.createObjectURL(blob);
    card.querySelector('.gen-upload-btn').addEventListener('click', () => uploadGeneration(i));
    container.appendChild(card);
  });
}

async function uploadGeneration(idx) {
  const gen = state.generations[idx];
  if (!gen) return;
  const btn = document.querySelector(`.gen-upload-btn[data-idx="${idx}"]`);
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Uploading…'; }
  try {
    const res = await fetch(`${API}/upload-sftp`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: gen.filename, xmlContent: gen.xml })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Upload failed');
    if (btn) { btn.textContent = '✅ Uploaded'; btn.style.background = '#27AE60'; }
    showTaggedTemplateDownload(state.generations || []);
  } catch (err) {
    if (btn) { btn.disabled = false; btn.textContent = '⬆ Upload SFTP'; }
    setStatus(5, 'error', `❌ ${err.message}`);
  }
}

// ── PO → VBKREQ Map download ─────────────────────────────────────────────────
function showTaggedTemplateDownload(generations) {
  const panel = document.getElementById('taggedTemplatePanel');
  if (!panel) return;
  if (!generations || !generations.length) return;

  const thStyle = 'background:#1F4E79;color:#fff;padding:6px 10px;text-align:left;font-size:12px';
  const tdStyle = 'padding:6px 10px;border-bottom:1px solid #e2e8f0;vertical-align:top;font-size:12px';

  const rows = generations.map(gen => {
    const pos   = (gen.poNumbers || []).map(p => `<div>${escapeHtml(p)}</div>`).join('') || '—';
    const asns  = (gen.asnRefs   || []).map(a => `<div style="font-family:monospace">${escapeHtml(a)}</div>`).join('') || '<span style="color:#aaa">—</span>';
    const vbRef = escapeHtml(gen.bookingRef || gen.ctrlNumber || '—');
    return `<tr>
      <td style="${tdStyle}">${pos}</td>
      <td style="${tdStyle}">${asns}</td>
      <td style="${tdStyle};font-family:monospace;font-weight:700;color:#1F4E79">${vbRef}</td>
    </tr>`;
  }).join('');

  // Download links for each tagged supplier template
  const fileNames = state.supplierFileNames || [];
  const downloadLinks = fileNames.length
    ? fileNames.map((name, idx) => {
        const taggedName = name.replace(/\.xlsx?$/i, '') + '_VBRef.xlsx';
        return `<a class="download-link" style="margin-right:12px" href="${API}/tagged-supplier/${idx}" download="${escapeHtml(taggedName)}">\u2b07 ${escapeHtml(taggedName)}</a>`;
      }).join('')
    : '';

  panel.innerHTML = `
    <table style="width:100%;border-collapse:collapse;margin-bottom:14px">
      <thead><tr>
        <th style="${thStyle}">PO Number(s)</th>
        <th style="${thStyle}">ASN/s</th>
        <th style="${thStyle}">VB Ref</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    ${downloadLinks ? `<div style="margin-top:6px;font-size:12px"><strong>Download tagged supplier template(s) with VB Ref:</strong><br/><div style="margin-top:6px">${downloadLinks}</div></div>` : ''}`;
}

async function refreshLog() {} // kept as no-op — log replaced by tagged template

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
        ? `<span style="color:#1B5E20;font-weight:700;font-size:11px">✅ Uploaded</span>`
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
