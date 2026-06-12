'use strict';

const API = 'http://localhost:3000/api';

// ── Session state ──────────────────────────────────────────────────────────────
const state = {
  poRefs: [],
  asnRefs: [],
  supplierRows: [],   // raw rows from supplier Excel (keyed by PO_Number + SKU)
  feedsFetched: false,
  biblBuilt: false,
  lastXml: null,
  lastFilename: null
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
const poFeedInput      = document.getElementById('poFeedInput');
const asnFeedInput     = document.getElementById('asnFeedInput');
const poDropZone       = document.getElementById('poDropZone');
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
function applyFileToZone(file) {
  if (!file) return;
  dropZone.classList.add('has-file');
  dropZone.classList.remove('drag-over');
  dropZoneText.textContent = '✓ ' + file.name;
  btnParseSupplier.disabled = false;
}

// Click on zone triggers file browser
dropZone.addEventListener('click', () => supplierFileInput.click());

// Drag events
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  const file = e.dataTransfer.files[0];
  if (file && /\.xlsx?$/i.test(file.name)) {
    // Assign to the hidden input via DataTransfer so FormData works
    const dt = new DataTransfer();
    dt.items.add(file);
    supplierFileInput.files = dt.files;
    applyFileToZone(file);
  } else if (file) {
    dropZone.classList.remove('drag-over');
    setStatus(1, 'error', '❌ Please drop an Excel file (.xlsx or .xls)');
  }
});

supplierFileInput.addEventListener('change', () => applyFileToZone(supplierFileInput.files[0]));

// ── Step 1: Parse Supplier ─────────────────────────────────────────────────────
btnParseSupplier.addEventListener('click', async () => {
  const file = supplierFileInput.files[0];
  if (!file) return;

  setLoading(btnParseSupplier, true);
  setStatus(1, 'loading', '⏳ Parsing supplier Excel…');
  refsPreview.innerHTML = '';

  try {
    const fd = new FormData();
    fd.append('supplierFile', file);
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

    let html = `✅ Parsed <strong>${data.rowCount}</strong> rows — found <strong>${state.poRefs.length}</strong> PO ref(s).`;
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
  btnUploadFeeds.disabled = !(poFeedInput.files[0] || asnFeedInput.files[0]);
}

// PO drop zone
poDropZone.addEventListener('click', () => poFeedInput.click());
poDropZone.addEventListener('dragover', e => { e.preventDefault(); poDropZone.classList.add('drag-over'); });
poDropZone.addEventListener('dragleave', () => poDropZone.classList.remove('drag-over'));
poDropZone.addEventListener('drop', e => {
  e.preventDefault();
  applyFeedFile(poDropZone, 'poDropZoneText', poFeedInput, e.dataTransfer.files[0]);
});
poFeedInput.addEventListener('change', () => applyFeedFile(poDropZone, 'poDropZoneText', poFeedInput, poFeedInput.files[0]));

// ASN drop zone
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
  setStatus(2, 'loading', '⏳ Uploading and parsing PO & ASN XML feeds…');

  try {
    const fd = new FormData();
    if (poFeedInput.files[0])  fd.append('poFeedFile',  poFeedInput.files[0]);
    if (asnFeedInput.files[0]) fd.append('asnFeedFile', asnFeedInput.files[0]);

    const res  = await fetch(`${API}/upload-feeds`, { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    state.feedsFetched = true;
    let html = `✅ Uploaded — <strong>${data.poFeedCount}</strong> PO feed(s) and <strong>${data.asnFeedCount}</strong> ASN feed(s) parsed.`;
    if (data.summary?.posFound?.length)  html += `<br/>POs: ${data.summary.posFound.join(', ')}`;
    if (data.summary?.asnsFound?.length) html += `<br/>ASNs: ${data.summary.asnsFound.join(', ')}`;
    if (data.errors?.length) html += `<br/>⚠️ ${data.errors.join('<br/>')}`;
    setStatus(2, 'success', html);
    setBadge(2, 'done');
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
  setStatus(2, 'loading', '⏳ Fetching PO & ASN feeds from Azure Blob Storage…');
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
    let html = `✅ Fetched <strong>${data.poFeedCount}</strong> PO feed(s) and <strong>${data.carrierAsnCount || 0}</strong> carrier ASN file(s) in <strong>${elapsed}s</strong>.${modeTag}`;
    if (data.errors && data.errors.length) {
      html += `<br/>⚠️ ${data.errors.join('<br/>')}`;
    }
    setStatus(2, 'success', html);
    setBadge(2, 'done');
    renderFeedPreview(data.feedsSummary || [], data.carrierAsnFiles || []);
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

  if (!feedsSummary.length && !carrierAsnFiles.length) {
    wrap.style.display = 'none';
    return;
  }

  // Index carrier files: poRef -> file; poRef+sku -> { asnId, qty, filename }
  const carrierByPo  = {}; // poRef -> filename
  const asnIndex     = {}; // poRef -> sku -> { asnId, qty }
  for (const f of carrierAsnFiles) {
    const poRef = f.poRef;
    carrierByPo[poRef] = f.filename;
    if (!asnIndex[poRef]) asnIndex[poRef] = {};
    for (const g of (f.asnGroups || [])) {
      for (const l of (g.lines || [])) {
        asnIndex[poRef][l.sku] = { asnId: g.asnId, qty: l.quantity };
      }
    }
  }

  // Index PO files: orderId -> filename
  const poFileByOrder = {};
  for (const po of feedsSummary) {
    // filename heuristic: server returns po.filename if available, else build from orderId
    poFileByOrder[po.orderId] = po.filename || `PO_${po.orderId}.xml`;
  }

  let rowsHtml = '';
  const missingFromCarrier = [];
  const missingFromPO      = [];

  // Rows from PO feed (each PO SKU line)
  for (const po of feedsSummary) {
    const poFile     = poFileByOrder[po.orderId];
    const poViewUrl  = `${API}/feed-raw?type=po&ref=${encodeURIComponent(po.orderId)}`;
    const poDlUrl    = `${API}/feed-raw?type=po&ref=${encodeURIComponent(po.orderId)}&download=1`;
    const carrierFile = carrierByPo[po.orderId];
    const carrierViewUrl = carrierFile ? `${API}/feed-raw?type=carrier&ref=${encodeURIComponent(carrierFile)}` : null;
    const carrierDlUrl   = carrierFile ? `${API}/feed-raw?type=carrier&ref=${encodeURIComponent(carrierFile)}&download=1` : null;
    const skuMap = asnIndex[po.orderId] || {};

    const lines = po.lineItems || [];
    if (!lines.length) {
      // PO with no lines — still show a row
      rowsHtml += `<tr>
        <td><strong>${po.orderId}</strong></td>
        <td>—</td>
        <td>—</td>
        <td>${po.supplierName || '—'}</td>
        <td style="text-align:right">—</td>
        <td><a class="file-link" href="${poViewUrl}" target="_blank">${poFile}</a> <a class="file-link-dl" href="${poDlUrl}" download>&#11015;</a></td>
        <td>${carrierFile ? `<a class="file-link" href="${carrierViewUrl}" target="_blank">${carrierFile}</a> <a class="file-link-dl" href="${carrierDlUrl}" download>&#11015;</a>` : '<em style="color:#aaa">None matched</em>'}</td>
      </tr>`;
    }

    for (const l of lines) {
      const asnData = skuMap[l.sku];
      const hasPO   = true;
      const hasASN  = !!asnData;
      if (!hasASN) missingFromCarrier.push({ po: po.orderId, sku: l.sku });

      const rowClass = hasASN ? '' : 'row-missing-po';
      rowsHtml += `<tr class="${rowClass}">
        <td><strong>${po.orderId}</strong></td>
        <td><strong>${l.sku}</strong></td>
        <td>${hasASN ? `<span style="font-family:monospace">${asnData.asnId}</span>` : '<em style="color:#aaa">—</em>'}</td>
        <td>${po.supplierName || '—'}</td>
        <td style="text-align:right">${hasASN ? (asnData.qty ?? '—') : '<em style="color:#aaa">—</em>'}</td>
        <td><a class="file-link" href="${poViewUrl}" target="_blank">${poFile}</a> <a class="file-link-dl" href="${poDlUrl}" download>&#11015;</a></td>
        <td>${carrierFile ? `<a class="file-link" href="${carrierViewUrl}" target="_blank">${carrierFile}</a> <a class="file-link-dl" href="${carrierDlUrl}" download>&#11015;</a>` : '<em style="color:#aaa">None matched</em>'}</td>
      </tr>`;
    }
  }

  // Extra rows: carrier SKUs not in any PO feed
  for (const f of carrierAsnFiles) {
    const poRef = f.poRef;
    const poFeed = feedsSummary.find(p => p.orderId === poRef);
    const poSkus = new Set((poFeed?.lineItems || []).map(l => l.sku));
    const carrierFile     = f.filename;
    const carrierViewUrl  = `${API}/feed-raw?type=carrier&ref=${encodeURIComponent(carrierFile)}`;
    const carrierDlUrl    = `${API}/feed-raw?type=carrier&ref=${encodeURIComponent(carrierFile)}&download=1`;

    for (const g of (f.asnGroups || [])) {
      for (const l of (g.lines || [])) {
        if (!poSkus.has(l.sku)) {
          missingFromPO.push({ po: poRef, sku: l.sku });
          rowsHtml += `<tr class="row-missing-asn">
            <td><strong>${poRef}</strong></td>
            <td><strong>${l.sku}</strong></td>
            <td><span style="font-family:monospace">${g.asnId}</span></td>
            <td><em style="color:#aaa">—</em></td>
            <td style="text-align:right">${l.quantity ?? '—'}</td>
            <td><em style="color:#aaa">Not in PO feed</em></td>
            <td><a class="file-link" href="${carrierViewUrl}" target="_blank">${carrierFile}</a> <a class="file-link-dl" href="${carrierDlUrl}" download>&#11015;</a></td>
          </tr>`;
        }
      }
    }
  }

  wrap.innerHTML = `
    <div class="blob-result-wrap">
      <div class="blob-result-scroll">
        <table class="blob-result-tbl">
          <thead><tr>
            <th>PO #</th>
            <th>SKU</th>
            <th>ASN Ref</th>
            <th>Supplier</th>
            <th>Carrier Qty</th>
            <th>PO File</th>
            <th>Carrier ASN File</th>
          </tr></thead>
          <tbody>${rowsHtml || '<tr><td colspan="7" style="text-align:center;color:#999">No data</td></tr>'}</tbody>
        </table>
      </div>
    </div>`;

  // Append notices AFTER the table wrapper so they are never clipped by the scroll container
  const noticesDiv = document.createElement('div');
  noticesDiv.id = 'blobResultNotices';

  // 1. SKUs in supplier template but missing from E2open PO feed
  if (state.supplierRows && state.supplierRows.length) {
    // Build set of SKUs in PO feed per PO
    const poFeedSkus = {}; // orderId -> Set
    for (const po of feedsSummary) {
      poFeedSkus[po.orderId] = new Set((po.lineItems || []).map(l => String(l.sku)));
    }
    const notInPOFeed = [];
    for (const r of state.supplierRows) {
      const po  = String(r.PO_Number || '').trim();
      const sku = String(r.SKU       || '').trim();
      if (!po || !sku) continue;
      const poSkus = poFeedSkus[po];
      if (poSkus === undefined) continue; // PO not fetched
      if (!poSkus.has(sku)) notInPOFeed.push({ po, sku });
    }
    if (notInPOFeed.length) {
      const list = notInPOFeed.map(x => `<strong>${x.sku}</strong> (PO ${x.po})`).join(', ');
      const d = document.createElement('div');
      d.className = 'blob-missing-msg blob-missing-po';
      d.innerHTML = `⚠ ${notInPOFeed.length} SKU(s) from your supplier template were <strong>not found in the E2open PO feed</strong>: ${list}`;
      noticesDiv.appendChild(d);
    }
  }

  // 2. SKUs in PO feed but not in carrier ASN
  if (missingFromCarrier.length) {
    const list = missingFromCarrier.map(x => `<strong>${x.sku}</strong> (PO ${x.po})`).join(', ');
    const d = document.createElement('div');
    d.className = 'blob-missing-msg blob-missing-po';
    d.innerHTML = `⚠ ${missingFromCarrier.length} SKU(s) from the PO feed were <strong>not found in the carrier ASN</strong>: ${list}`;
    noticesDiv.appendChild(d);
  }

  // 3. SKUs in carrier ASN but not in PO feed
  if (missingFromPO.length) {
    const list = missingFromPO.map(x => `<strong>${x.sku}</strong> (PO ${x.po})`).join(', ');
    const d = document.createElement('div');
    d.className = 'blob-missing-msg blob-missing-asn';
    d.innerHTML = `❌ ${missingFromPO.length} SKU(s) in the carrier ASN were <strong>not found in the PO feed</strong>: ${list}`;
    noticesDiv.appendChild(d);
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
  setStatus(3, 'loading', '⏳ Building Bible Excel — merging supplier data + PO + ASN feeds…');

  try {
    const res  = await fetch(`${API}/build-bible`, { method: 'POST' });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    state.biblBuilt = true;
    const dlLink = data.downloadUrl
      ? `<br/><a class="download-link" href="${data.downloadUrl}" download>⬇ Download Bible Excel</a>`
      : '';
    let html = `✅ Bible built with <strong>${data.masterRowCount}</strong> MASTER row(s).${dlLink}`;
    if (data.warnings && data.warnings.length) {
      html += `<br/><br/>⚠️ <strong>${data.warnings.length} SKU(s) excluded</strong> — in supplier template but not on carrier feed:<br/>`
            + data.warnings.map(w => `&nbsp;• ${w}`).join('<br/>');
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
btnGenerateVbkreq.addEventListener('click', async () => {
  setLoading(btnGenerateVbkreq, true);
  const purposeCd = document.querySelector('input[name="purposeCd"]:checked')?.value || '13';
  const purposeLabel = { '13': 'Request', '15': 'Re-Submission', '01': 'Cancellation' }[purposeCd] || purposeCd;
  setStatus(4, 'loading', `⏳ Generating VBKREQ XML — PurposeCd <strong>${purposeCd}</strong> (${purposeLabel})…`);
  xmlPreviewWrap.classList.remove('visible');

  try {
    const res  = await fetch(`${API}/generate-vbkreq`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ purposeCd })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    state.lastXml      = data.xml;
    state.lastFilename = data.filename;

    setStatus(4, 'success', `✅ VBKREQ generated — <strong>${data.filename}</strong> (PurposeCd: ${purposeCd} | Version: ${data.version || ''} | CtrlNumber: ${data.ctrlNumber})`);
    setBadge(4, 'done');

    // Show XML preview
    xmlPreview.textContent = data.xml;
    xmlPreviewWrap.classList.add('visible');

    // Blob download
    const blob = new Blob([data.xml], { type: 'application/xml' });
    xmlDownloadLink.href = URL.createObjectURL(blob);
    xmlDownloadLink.download = data.filename;

    btnUploadSftp.disabled = false;
    setBadge(5, 'active');
    refreshLog();
  } catch (err) {
    setStatus(4, 'error', `❌ ${err.message}`);
  } finally {
    setLoading(btnGenerateVbkreq, false);
  }
});

// ── Step 5: Upload SFTP ────────────────────────────────────────────────────────
btnUploadSftp.addEventListener('click', async () => {
  if (!state.lastXml || !state.lastFilename) {
    setStatus(5, 'error', '❌ No XML generated. Run step 4 first.');
    return;
  }
  setLoading(btnUploadSftp, true);
  setStatus(5, 'loading', `⏳ Uploading <strong>${state.lastFilename}</strong> to E2open SFTP…`);

  try {
    const res  = await fetch(`${API}/upload-sftp`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: state.lastFilename, xmlContent: state.lastXml })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    setStatus(5, 'success',
      `✅ Uploaded successfully!<br/>
       Remote path: <strong>${data.remotePath}</strong><br/>
       Bytes sent: ${data.bytesSent}<br/>
       Uploaded at: ${data.uploadedAt}`
    );
    setBadge(5, 'done');
    refreshLog();
  } catch (err) {
    setStatus(5, 'error', `❌ Upload failed: ${err.message}`);
  } finally {
    setLoading(btnUploadSftp, false);
  }
});

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
