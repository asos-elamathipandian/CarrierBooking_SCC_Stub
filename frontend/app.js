'use strict';

const API = 'http://localhost:3000/api';

// ── Session state ──────────────────────────────────────────────────────────────
const state = {
  poRefs: [],
  asnRefs: [],
  supplierRows: [],
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
const btnFetchFeeds      = document.getElementById('btnFetchFeeds');
const btnBuildBible      = document.getElementById('btnBuildBible');
const btnGenerateVbkreq  = document.getElementById('btnGenerateVbkreq');
const btnUploadSftp      = document.getElementById('btnUploadSftp');
const btnRefreshLog      = document.getElementById('btnRefreshLog');
const xmlPreviewWrap     = document.getElementById('xmlPreviewWrap');
const xmlPreview         = document.getElementById('xmlPreview');
const xmlDownloadLink    = document.getElementById('xmlDownloadLink');
const refsPreview        = document.getElementById('refsPreview');
const logBody            = document.getElementById('logBody');

// Step 2 — feed tab elements
const asnFeedInput     = document.getElementById('asnFeedInput');
const asnDropZone      = document.getElementById('asnDropZone');
const btnUploadFeeds   = document.getElementById('btnUploadFeeds');

// ── Helpers ────────────────────────────────────────────────────────────────────
function setStatus(step, type, html) {
  const el = document.getElementById(`status${step}`);
  el.className = `status-box ${type}`;
  el.innerHTML = html;
}

function setBadge(step, type) {
  const el = document.getElementById(`badge${step}`);
  el.className = `step-badge ${type}`;
  if (type === 'done') el.textContent = '✓';
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

    state.poRefs  = data.poRefs  || [];
    state.asnRefs = [];
    state.supplierRows = data.rows || data.preview || [];

    // Auto-populate PO refs in the blob fetch panel
    const blobPoRefTags = document.getElementById('blobPoRefTags');
    if (blobPoRefTags) {
      blobPoRefTags.innerHTML = state.poRefs.length
        ? state.poRefs.map(r => `<span class="tag">PO: ${r}</span>`).join('')
        : '<span style="color:#aaa;font-size:12px">No PO refs found in template.</span>';
    }

    const fileCountNote = data.fileCount > 1 ? ` from <strong>${data.fileCount}</strong> files` : '';
    let html = `✅ Parsed <strong>${data.rowCount}</strong> rows${fileCountNote} — found <strong>${state.poRefs.length}</strong> PO ref(s).`;
    if (data.validationErrors && data.validationErrors.length) {
      html += `<br/>⚠️ Validation warnings:<br/>${data.validationErrors.join('<br/>')}`;
    }
    setStatus(1, 'success', html);
    setBadge(1, 'done');

    // Show supplier rows table
    refsPreview.innerHTML = '';
    renderSupplierTable(data.rows || data.preview || []);

    btnFetchFeeds.disabled = false;
    setBadge(2, 'active');
  } catch (err) {
    setStatus(1, 'error', `❌ ${err.message}`);
  } finally {
    setLoading(btnParseSupplier, false);
    btnParseSupplier.disabled = true;
  }
});

// ── Step 2: Feed source tab switch ────────────────────────────────────────────
function switchFeedTab(tab) {
  document.getElementById('tabBtnUpload').classList.toggle('active', tab === 'upload');
  document.getElementById('tabBtnBlob').classList.toggle('active', tab === 'blob');
  document.getElementById('panelUploadFeeds').classList.toggle('active', tab === 'upload');
  document.getElementById('panelFetchBlob').classList.toggle('active', tab === 'blob');
  setStatus(2, 'info', '');
  document.getElementById('status2').style.display = 'none';
}
// expose globally so onclick= in HTML works
window.switchFeedTab = switchFeedTab;

// ── Step 2: PO / ASN XML file drop zones ─────────────────────────────────────
function applyFeedFile(zone, textElId, input, file) {
  if (!file) return;
  if (!/\.xml$/i.test(file.name)) {
    alert('Please select an XML file.');
    return;
  }
  const dt = new DataTransfer();
  dt.items.add(file);
  input.files = dt.files;
  zone.classList.add('has-file');
  document.getElementById(textElId).textContent = '✓ ' + file.name;
  updateUploadFeedsBtn();
}

function updateUploadFeedsBtn() {
  btnUploadFeeds.disabled = !asnFeedInput.files[0];
}

// Carrier ASN drop zone
asnDropZone.addEventListener('click', () => asnFeedInput.click());
asnDropZone.addEventListener('dragover', e => { e.preventDefault(); asnDropZone.classList.add('drag-over'); });
asnDropZone.addEventListener('dragleave', () => asnDropZone.classList.remove('drag-over'));
asnDropZone.addEventListener('drop', e => {
  e.preventDefault();
  applyFeedFile(asnDropZone, 'asnDropZoneText', asnFeedInput, e.dataTransfer.files[0]);
});
asnFeedInput.addEventListener('change', () => applyFeedFile(asnDropZone, 'asnDropZoneText', asnFeedInput, asnFeedInput.files[0]));

// ── Step 2: Upload XML Feeds ──────────────────────────────────────────────────
btnUploadFeeds.addEventListener('click', async () => {
  setLoading(btnUploadFeeds, true);
  setStatus(2, 'loading', '⏳ Uploading and parsing carrier feed XML…');

  try {
    const fd = new FormData();
    if (asnFeedInput.files[0]) fd.append('asnFeedFile', asnFeedInput.files[0]);

    const res  = await fetch(`${API}/upload-feeds`, { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    state.feedsFetched = true;
    let html = `✅ Carrier feed uploaded — <strong>${data.carrierAsnCount}</strong> file(s) parsed.`;
    if (data.errors?.length) html += `<br/>⚠️ ${data.errors.join('<br/>')}`;
    setStatus(2, 'success', html);
    setBadge(2, 'done');
    renderFeedPreview([], data.carrierAsnFiles || []);
    btnBuildBible.disabled = false;
    setBadge(3, 'active');
  } catch (err) {
    setStatus(2, 'error', `❌ ${err.message}`);
  } finally {
    setLoading(btnUploadFeeds, false);
  }
});

// ── Step 2: Fetch Feeds ────────────────────────────────────────────────────────
btnFetchFeeds.addEventListener('click', async () => {
  setLoading(btnFetchFeeds, true);
  setStatus(2, 'loading', '⏳ Fetching carrier feed from Azure Blob Storage…');
  const fetchStart = performance.now();

  try {
    const res  = await fetch(`${API}/fetch-feeds`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ poRefs: state.poRefs, asnRefs: state.asnRefs })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    const elapsed = ((performance.now() - fetchStart) / 1000).toFixed(2);
    state.feedsFetched = true;
    const modeTag = data.localMode ? ' <em style="color:#e67e22">[LOCAL MODE — reading from samples/feeds/]</em>' : '';
    let html = `✅ Fetched <strong>${data.carrierAsnCount || 0}</strong> carrier ASN file(s) in <strong>${elapsed}s</strong>.${modeTag}`;
    if (data.errors && data.errors.length) {
      html += `<br/>⚠️ ${data.errors.join('<br/>')}`;
    }
    setStatus(2, 'success', html);
    setBadge(2, 'done');
    renderFeedPreview([], data.carrierAsnFiles || []);
    btnBuildBible.disabled = false;
    setBadge(3, 'active');
  } catch (err) {
    setStatus(2, 'error', `❌ ${err.message}`);
  } finally {
    setLoading(btnFetchFeeds, false);
  }
});

// ── Blob Result: single unified table ────────────────────────────────────────────
function renderFeedPreview(feedsSummary, carrierAsnFiles) {
  const panel = document.getElementById('feedPreviewPanel');
  const wrap  = document.getElementById('blobResultTable');
  if (!panel || !wrap) return;
  panel.style.display = 'block';

  if (!carrierAsnFiles.length) {
    wrap.style.display = 'none';
    return;
  }

  let rowsHtml = '';
  for (const f of carrierAsnFiles) {
    const carrierViewUrl = `${API}/feed-raw?type=carrier&ref=${encodeURIComponent(f.filename)}`;
    const carrierDlUrl   = `${API}/feed-raw?type=carrier&ref=${encodeURIComponent(f.filename)}&download=1`;
    for (const g of (f.asnGroups || [])) {
      g.lines.forEach((l, i) => {
        rowsHtml += `<tr>
          <td>${i === 0 ? `<strong>${f.poRef}</strong>` : ''}</td>
          <td style="font-family:monospace;font-size:11px">${i === 0 ? (g.asnId || '—') : ''}</td>
          <td>${i === 0 ? (g.supplier || '—') : ''}</td>
          <td><strong>${l.sku}</strong></td>
          <td style="font-family:monospace;font-size:11px">${l.ean || '—'}</td>
          <td>${l.colour || '—'}</td>
          <td>${l.size   || '—'}</td>
          <td style="text-align:right">${l.quantity ?? '—'}</td>
          <td>${l.country || '—'}</td>
          <td style="font-size:11px">${i === 0 ? (g.shipDate || '—') : ''}</td>
          <td style="font-size:11px">${i === 0 ? `<a class="file-link" href="${carrierViewUrl}" target="_blank">${f.filename}</a> <a class="file-link-dl" href="${carrierDlUrl}" download>&#11015;</a>` : ''}</td>
        </tr>`;
      });
    }
  }

  // Check supplier template SKUs against carrier feed
  const carrierSkus = {}; // poRef -> Set<sku>
  for (const f of carrierAsnFiles) {
    if (!carrierSkus[f.poRef]) carrierSkus[f.poRef] = new Set();
    for (const g of (f.asnGroups || [])) {
      for (const l of (g.lines || [])) carrierSkus[f.poRef].add(l.sku);
    }
  }

  wrap.innerHTML = `
    <div class="blob-result-wrap">
      <div class="blob-result-scroll">
        <table class="blob-result-tbl">
          <thead><tr>
            <th>PO #</th><th>ASN ID</th><th>Supplier</th>
            <th>SKU</th><th>EAN</th><th>Colour</th><th>Size</th>
            <th>Qty</th><th>Country</th><th>Ship Date</th><th>Carrier File</th>
          </tr></thead>
          <tbody>${rowsHtml || '<tr><td colspan="11" style="text-align:center;color:#999">No data</td></tr>'}</tbody>
        </table>
      </div>
    </div>`;

  // Notice: supplier template SKUs not found in carrier feed
  const noticesDiv = document.createElement('div');
  noticesDiv.id = 'blobResultNotices';
  if (state.supplierRows && state.supplierRows.length) {
    const missing = [];
    for (const r of state.supplierRows) {
      const po  = String(r.PO_Number || '').trim();
      const sku = String(r.SKU       || '').trim();
      if (!po || !sku) continue;
      if (carrierSkus[po] && !carrierSkus[po].has(sku)) missing.push({ po, sku });
    }
    if (missing.length) {
      const list = missing.map(x => `<strong>${x.sku}</strong> (PO ${x.po})`).join(', ');
      const d = document.createElement('div');
      d.className = 'blob-missing-msg blob-missing-po';
      d.innerHTML = `⚠ ${missing.length} SKU(s) from your supplier template were <strong>not found in the carrier feed</strong>: ${list}`;
      noticesDiv.appendChild(d);
    }
  }
  wrap.appendChild(noticesDiv);
  wrap.style.display = 'block';
}



// ── Supplier parse table ──────────────────────────────────────────────────────────────────────────────
function renderSupplierTable(rows) {
  const panel = document.getElementById('supplierParsePanel');
  if (!panel || !rows || rows.length === 0) { if (panel) panel.style.display = 'none'; return; }

  const poGroups = {};
  for (const r of rows) {
    const po = String(r.PO_Number || '').trim() || '(no PO)';
    if (!poGroups[po]) poGroups[po] = [];
    poGroups[po].push(r);
  }

  const poCount      = Object.keys(poGroups).length;
  const totalQty     = rows.reduce((s, r) => s + (parseFloat(r.Booking_Qty)   || 0), 0);
  const totalCartons = rows.reduce((s, r) => s + (parseFloat(r.No_of_Cartons) || 0), 0);

  const fmtDate = v => {
    if (!v) return '\u2014';
    if (v instanceof Date || (typeof v === 'string' && v.includes('T'))) return new Date(v).toLocaleDateString('en-GB');
    return v;
  };

  let rowsHtml = '';
  for (const [po, poRows] of Object.entries(poGroups)) {
    poRows.forEach((r, i) => {
      rowsHtml += `<tr>
        <td>${i === 0 ? `<span class="tag" style="font-size:11px">${po}</span>` : ''}</td>
        <td><strong>${r.SKU || ''}</strong></td>
        <td style="font-family:monospace;font-size:11px">${r.EAN_Barcode || '\u2014'}</td>
        <td>${r.Colour_Code || '\u2014'}</td>
        <td>${r.Size_Code   || '\u2014'}</td>
        <td style="text-align:right">${r.Booking_Qty   ?? '\u2014'}</td>
        <td style="text-align:right">${r.No_of_Cartons ?? '\u2014'}</td>
        <td>${r.Carton_Type || '\u2014'}</td>
        <td>${r.Pack_Type   || '\u2014'}</td>
        <td style="font-size:11px">${fmtDate(r.Cargo_Ready_Planned_Collection_Date)}</td>
      </tr>`;
    });
  }

  panel.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:8px">
      <span style="font-size:12px;font-weight:700;color:#1F4E79">\ud83d\udccb Supplier rows parsed</span>
      <div class="parse-summary-stats">
        <span class="parse-stat-badge">\ud83d\udce6 ${poCount} PO ref${poCount !== 1 ? 's' : ''}</span>
        <span class="parse-stat-badge">\ud83c\udff7 ${rows.length} SKU row${rows.length !== 1 ? 's' : ''}</span>
        <span class="parse-stat-badge">\ud83d\udd22 ${totalQty.toLocaleString()} units</span>
        <span class="parse-stat-badge">\ud83d\udceb ${totalCartons.toLocaleString()} cartons</span>
      </div>
    </div>
    <div class="parse-table-wrap">
      <div class="parse-table-scroll">
        <table class="supplier-table">
          <thead><tr>
            <th>PO #</th><th>SKU</th><th>EAN Barcode</th>
            <th>Colour</th><th>Size</th><th>Booking Qty</th>
            <th>No. Cartons</th><th>Carton Type</th>
            <th>Pack Type</th><th>Cargo Ready Date</th>
          </tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
    </div>`;
  panel.style.display = 'block';
}

// ── Step 3: Build Bible ────────────────────────────────────────────────────────
btnBuildBible.addEventListener('click', async () => {
  setLoading(btnBuildBible, true);
  setStatus(3, 'loading', '⏳ Building Master Workbook — merging supplier data + carrier feed…');

  try {
    const res  = await fetch(`${API}/build-bible`, { method: 'POST' });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    state.biblBuilt = true;
    const dlLink = data.downloadUrl
      ? `<br/><a class="download-link" href="${data.downloadUrl}" download>⬇ Download Master Workbook</a>`
      : '';
    let html = `✅ Bible built with <strong>${data.masterRowCount}</strong> MASTER row(s).${dlLink}`;
    if (data.warnings && data.warnings.length) {
      html += `<br/><br/>⚠️ <strong>${data.warnings.length} SKU(s) excluded</strong> — in supplier template but NOT on carrier ASN:<br/>`
            + data.warnings.map(w => `&nbsp;• ${w}`).join('<br/>');
    }
    if (data.extraSkuWarnings && data.extraSkuWarnings.length) {
      html += `<br/><br/>⚠️ <strong>${data.extraSkuWarnings.length} extra SKU(s) from carrier ASN</strong> — not in your supplier template (included in VBKREQ with Booking_Qty=0 — please review):<br/>`
            + data.extraSkuWarnings.map(w => `&nbsp;• ${w}`).join('<br/>');
    }
    setStatus(3, 'success', html);
    setBadge(3, 'done');
    btnGenerateVbkreq.disabled = false;
    setBadge(4, 'active');
  } catch (err) {
    setStatus(3, 'error', `❌ ${err.message}`);
  } finally {
    setLoading(btnBuildBible, false);
  }
});

// ── Step 4: Generate VBKREQ ───────────────────────────────────────────────────
// Date picker — show/hide clear button
function wireDateClear(inputId, clearBtnId) {
  const input = document.getElementById(inputId);
  const btn   = document.getElementById(clearBtnId);
  if (!input || !btn) return;
  input.addEventListener('change', () => btn.classList.toggle('visible', !!input.value));
}
wireDateClear('overrideCargoReady',    'clearCargoReady');
wireDateClear('overrideBookingReqDate','clearBookingReqDate');

// Exposed to onclick= in HTML
function clearDateField(inputId, clearBtnId) {
  const input = document.getElementById(inputId);
  const btn   = document.getElementById(clearBtnId);
  if (input) input.value = '';
  if (btn)   btn.classList.remove('visible');
}
window.clearDateField = clearDateField;

btnGenerateVbkreq.addEventListener('click', async () => {
  setLoading(btnGenerateVbkreq, true);
  const purposeCd = document.querySelector('input[name="purposeCd"]:checked')?.value || '13';
  const purposeLabel = { '13': 'Request', '15': 'Re-Submission', '01': 'Cancellation' }[purposeCd] || purposeCd;
  setStatus(4, 'loading', `⏳ Generating VBKREQ XML — PurposeCd <strong>${purposeCd}</strong> (${purposeLabel})…`);
  xmlPreviewWrap.classList.remove('visible');

  const overrideCargoReady     = document.getElementById('overrideCargoReady')?.value    || '';
  const overrideBookingReqDate = document.getElementById('overrideBookingReqDate')?.value || '';

  try {
    const res  = await fetch(`${API}/generate-vbkreq`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ purposeCd, overrideCargoReady, overrideBookingReqDate })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    state.generations  = data.generations || [];
    state.lastXml      = state.generations[0]?.xml      || null;
    state.lastFilename = state.generations[0]?.filename || null;

    const count = state.generations.length;
    const dateNote = [];
    if (overrideCargoReady)     dateNote.push(`Cargo Ready: <strong>${overrideCargoReady}</strong>`);
    if (overrideBookingReqDate) dateNote.push(`Booking Req: <strong>${overrideBookingReqDate}</strong>`);
    const dateStr = dateNote.length ? `<br/>📅 Date overrides applied — ${dateNote.join(' &nbsp;|&nbsp; ')}` : '';
    setStatus(4, 'success',
      `✅ ${count} VBKREQ${count > 1 ? 's' : ''} generated — PurposeCd: ${purposeCd} (${purposeLabel})${dateStr}`
    );
    setBadge(4, 'done');
    renderGenerations(state.generations);
    xmlPreviewWrap.classList.add('visible');
    btnUploadSftp.disabled = false;
    setBadge(5, 'active');
    refreshLog();
  } catch (err) {
    setStatus(4, 'error', `❌ ${err.message}`);
  } finally {
    setLoading(btnGenerateVbkreq, false);
  }
});

// ── Step 5: Upload All SFTP ───────────────────────────────────────────────────
btnUploadSftp.addEventListener('click', async () => {
  if (!state.generations || state.generations.length === 0) {
    setStatus(5, 'error', '❌ No VBKREQs generated. Run step 4 first.');
    return;
  }
  setLoading(btnUploadSftp, true);
  setStatus(5, 'loading', `⏳ Uploading ${state.generations.length} VBKREQ(s) to E2open SFTP…`);

  const results = [];
  for (let i = 0; i < state.generations.length; i++) {
    const gen = state.generations[i];
    try {
      const res = await fetch(`${API}/upload-sftp`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename: gen.filename, xmlContent: gen.xml })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Upload failed');
      results.push({ filename: gen.filename, ok: true, remotePath: data.remotePath, localMode: data.localMode });
      const btn = document.querySelector(`.gen-upload-btn[data-idx="${i}"]`);
      if (btn) { btn.textContent = '✅ Uploaded'; btn.style.background = '#27AE60'; btn.disabled = true; }
    } catch (err) {
      results.push({ filename: gen.filename, ok: false, error: err.message });
      const btn = document.querySelector(`.gen-upload-btn[data-idx="${i}"]`);
      if (btn) { btn.disabled = false; btn.textContent = '⬆ Retry'; }
    }
  }

  const ok   = results.filter(r => r.ok);
  const fail = results.filter(r => !r.ok);
  let html = `✅ ${ok.length} of ${results.length} uploaded`;
  if (ok.length)   html += ':<br/>' + ok.map(r   => `&nbsp;• ${r.filename} → ${r.localMode ? 'local output/' : r.remotePath}`).join('<br/>');
  if (fail.length) html += `<br/>❌ ${fail.length} failed:<br/>` + fail.map(r => `&nbsp;• ${r.filename}: ${r.error}`).join('<br/>');
  setStatus(5, fail.length === 0 ? 'success' : 'error', html);
  setBadge(5, fail.length === 0 ? 'done' : 'active');
  refreshLog();
  setLoading(btnUploadSftp, false);
});

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
    setStatus(5, 'success',
      `✅ <strong>${gen.filename}</strong> uploaded — ${data.localMode ? 'saved locally (no SFTP configured)' : `remote: ${data.remotePath}`}`
    );
    setBadge(5, 'done');
    refreshLog();
  } catch (err) {
    if (btn) { btn.disabled = false; btn.textContent = '⬆ Upload SFTP'; }
    setStatus(5, 'error', `❌ ${err.message}`);
  }
}

// ── Generation Log ─────────────────────────────────────────────────────────────
async function refreshLog() {
  try {
    const res  = await fetch(`${API}/generation-log`);
    const data = await res.json();
    if (!res.ok || !data.entries) return;

    if (data.entries.length === 0) {
      logBody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:#999">No entries yet</td></tr>';
      return;
    }

    logBody.innerHTML = data.entries.map(e => `
      <tr>
        <td>${e.Timestamp || ''}</td>
        <td>${e.Booking_Ref || ''}</td>
        <td style="font-size:11px">${e.PO_Numbers || ''}</td>
        <td style="font-size:11px;word-break:break-all">${e.Filename || ''}</td>
        <td>${e.CtrlNumber || ''}</td>
        <td>${e.SFTP_Status
          ? `<span class="badge-ok">OK</span>`
          : `<span class="badge-fail">—</span>`}</td>
      </tr>`).join('');
  } catch (_) {}
}

btnRefreshLog.addEventListener('click', refreshLog);

// Load log on page load
refreshLog();
