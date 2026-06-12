'use strict';

const API = 'http://localhost:3000/api';

// ── Session state ──────────────────────────────────────────────────────────────
const state = {
  poRefs: [],
  asnRefs: [],
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

// ── PO ↔ Carrier ASN Mapping table ────────────────────────────────────────────
function renderFeedMapping(feedsSummary, carrierAsnFiles) {
  const section = document.getElementById('feedMappingSection');
  const body    = document.getElementById('feedMappingBody');
  if (!section || !body) return;

  if (!feedsSummary.length && !carrierAsnFiles.length) {
    section.style.display = 'none';
    return;
  }

  // Build carrier ASN index: poId -> sku -> { asnId, fcId, shipDate, ean, size, colour, qty, pack }
  const asnIndex = {}; // poId -> sku -> line
  const asnSkusByPo = {}; // poId -> Set of skus that appear in carrier
  for (const f of carrierAsnFiles) {
    for (const g of (f.asnGroups || [])) {
      const poId = g.poId || f.poRef;
      if (!asnIndex[poId]) { asnIndex[poId] = {}; asnSkusByPo[poId] = new Set(); }
      for (const l of (g.lines || [])) {
        asnIndex[poId][l.sku] = { asnId: g.asnId, fcId: g.fcId, shipDate: g.shipDate, ean: l.ean, size: l.size, colour: l.colour, qty: l.quantity, pack: l.packFormat };
        asnSkusByPo[poId].add(l.sku);
      }
    }
  }

  let totalMatched = 0, totalPoOnly = 0, totalAsnOnly = 0;
  let allRowsHtml = '';

  for (const po of feedsSummary) {
    const poLines    = po.lineItems || [];
    const poSkus     = new Set(poLines.map(l => l.sku));
    const carrierMap = asnIndex[po.orderId] || {};
    const carrierSkus= asnSkusByPo[po.orderId] || new Set();

    // All unique SKUs across both sources
    const allSkus = new Set([...poSkus, ...carrierSkus]);

    for (const sku of allSkus) {
      const poLine  = poLines.find(l => l.sku === sku);
      const asnLine = carrierMap[sku];
      const inPO    = !!poLine;
      const inASN   = !!asnLine;

      let rowClass = '', badge = '';
      if (inPO && inASN) {
        rowClass = 'row-matched';
        badge = '<span class="badge-matched">✓ Matched</span>';
        totalMatched++;
      } else if (inPO) {
        rowClass = 'row-unmatched-po';
        badge = '<span class="badge-po-only">PO only</span>';
        totalPoOnly++;
      } else {
        rowClass = 'row-unmatched-asn';
        badge = '<span class="badge-asn-only">ASN only</span>';
        totalAsnOnly++;
      }

      allRowsHtml += `<tr class="${rowClass}">
        <td>${badge}</td>
        <td><strong>${po.orderId}</strong></td>
        <td><strong>${sku}</strong></td>
        <td>${inPO ? (poLine.description || '—') : '<em style="color:#aaa">—</em>'}</td>
        <td style="text-align:right">${inPO ? (poLine.poQty ?? '—') : '<em style="color:#aaa">—</em>'}</td>
        <td>${inPO ? (poLine.productStyle || '—') : '<em style="color:#aaa">—</em>'}</td>
        <td>${inASN ? (asnLine.asnId   || '—') : '<em style="color:#aaa">—</em>'}</td>
        <td>${inASN ? (asnLine.fcId    || '—') : '<em style="color:#aaa">—</em>'}</td>
        <td>${inASN ? (asnLine.shipDate|| '—') : '<em style="color:#aaa">—</em>'}</td>
        <td style="font-family:monospace;font-size:11px">${inASN ? (asnLine.ean || '—') : '<em style="color:#aaa">—</em>'}</td>
        <td>${inASN ? (asnLine.size    || '—') : '<em style="color:#aaa">—</em>'}</td>
        <td>${inASN ? (asnLine.colour  || '—') : '<em style="color:#aaa">—</em>'}</td>
        <td style="text-align:right">${inASN ? (asnLine.qty ?? '—') : '<em style="color:#aaa">—</em>'}</td>
        <td>${inASN ? (asnLine.pack === 'H' ? 'Hanging' : 'Flat') : '<em style="color:#aaa">—</em>'}</td>
      </tr>`;
    }
  }

  body.innerHTML = `
    <div class="mapping-summary">
      <span class="parse-stat-badge">Total SKUs: ${totalMatched + totalPoOnly + totalAsnOnly}</span>
      <span class="parse-stat-badge" style="background:#EAFAF1;border-color:#A9DFBF;color:#1E8449">✓ Matched: ${totalMatched}</span>
      ${totalPoOnly  > 0 ? `<span class="parse-stat-badge" style="background:#FEF9E7;border-color:#F9E79F;color:#784212">⚠ PO only (not in carrier): ${totalPoOnly}</span>` : ''}
      ${totalAsnOnly > 0 ? `<span class="parse-stat-badge" style="background:#FDEDEC;border-color:#F5B7B1;color:#922B21">✗ ASN only (not in PO): ${totalAsnOnly}</span>` : ''}
    </div>
    <div class="mapping-legend">
      <span style="color:#555">Row colours:</span>
      <span style="background:#EAFAF1;padding:2px 8px;border-radius:4px;font-size:11px">Green = matched both</span>
      <span style="background:#FEF9E7;padding:2px 8px;border-radius:4px;font-size:11px">Amber = PO only</span>
      <span style="background:#FDEDEC;padding:2px 8px;border-radius:4px;font-size:11px">Red = ASN only</span>
    </div>
    <div class="mapping-table-wrap">
      <div class="mapping-table-scroll">
        <table class="mapping-table">
          <thead>
            <tr>
              <th colspan="6" class="col-group-po">📦 E2open PO Feed</th>
              <th colspan="8" class="col-group-asn">🚛 Carrier ASN Feed</th>
            </tr>
            <tr>
              <th>Status</th>
              <th>PO #</th>
              <th>SKU</th>
              <th>Description</th>
              <th>PO Qty</th>
              <th>Style</th>
              <th class="col-asn">ASN ID</th>
              <th class="col-asn">FC</th>
              <th class="col-asn">Ship Date</th>
              <th class="col-asn">EAN</th>
              <th class="col-asn">Size</th>
              <th class="col-asn">Colour</th>
              <th class="col-asn">ASN Qty</th>
              <th class="col-asn">Pack</th>
            </tr>
          </thead>
          <tbody>${allRowsHtml}</tbody>
        </table>
      </div>
    </div>`;

  section.style.display = 'block';
}

// ── Feed preview rendering ─────────────────────────────────────────────────────
function renderFeedPreview(feedsSummary, carrierAsnFiles) {
  const panel        = document.getElementById('feedPreviewPanel');
  const poSection    = document.getElementById('poFeedsSection');
  const poTbody      = document.getElementById('poFeedsTableBody');
  const asnSection   = document.getElementById('carrierAsnSection');
  const asnList      = document.getElementById('carrierAsnList');

  panel.style.display = 'block';

  // ── PO feeds table ──
  poTbody.innerHTML = '';
  if (feedsSummary.length) {
    poSection.style.display = 'block';
    feedsSummary.forEach((po, i) => {
      const rowId   = `po-xml-row-${i}`;
      const preId   = `po-xml-pre-${i}`;
      const linesId = `po-lines-row-${i}`;
      const dlUrl   = `${API}/feed-raw?type=po&ref=${encodeURIComponent(po.orderId)}&download=1`;
      const lineItems = po.lineItems || [];
      const dataRow = document.createElement('tr');
      dataRow.innerHTML = `
        <td><strong>${po.orderId}</strong></td>
        <td>${po.supplierName || '\u2014'}</td>
        <td>${po.factoryName  || '\u2014'}</td>
        <td>${po.shipDate     || '\u2014'}</td>
        <td>${po.incoterms    || '\u2014'}</td>
        <td>${po.lineCount}</td>
        <td style="white-space:nowrap">
          ${lineItems.length > 0 ? `<button class="btn-view-xml" onclick="togglePoLines('${linesId}',this)" style="margin-right:4px">\ud83d\udccb Lines</button>` : ''}
          <button class="btn-view-xml" onclick="togglePoXml('${po.orderId}','${rowId}','${preId}',this)">View XML</button>
          <a class="btn-view-xml" href="${dlUrl}" download style="margin-left:4px;text-decoration:none">&#11015; Download</a>
        </td>`;
      poTbody.appendChild(dataRow);

      // Line items expand row
      const linesRow = document.createElement('tr');
      linesRow.id = linesId;
      linesRow.className = 'feed-xml-row';
      const lineItemsHtml = lineItems.length
        ? `<table class="lines-sub-table">
            <thead><tr><th>SKU</th><th>Product Style</th><th>Description</th><th>PO Qty</th><th>Mode</th></tr></thead>
            <tbody>${lineItems.map(l => `<tr>
              <td><strong>${l.sku}</strong></td>
              <td>${l.productStyle || '\u2014'}</td>
              <td>${l.description  || '\u2014'}</td>
              <td style="text-align:right">${l.poQty ?? '\u2014'}</td>
              <td>${l.mode || '\u2014'}</td>
            </tr>`).join('')}</tbody>
           </table>`
        : '<em style="color:#999;font-size:12px">No line items</em>';
      linesRow.innerHTML = `<td colspan="7" style="padding:0"><div class="lines-expand-wrap">${lineItemsHtml}</div></td>`;
      poTbody.appendChild(linesRow);

      // XML expand row
      const xmlRow = document.createElement('tr');
      xmlRow.id = rowId;
      xmlRow.className = 'feed-xml-row';
      xmlRow.innerHTML = `<td colspan="7"><pre class="feed-xml-pre" id="${preId}">Loading\u2026</pre></td>`;
      poTbody.appendChild(xmlRow);
    });
  } else {
    poSection.style.display = 'none';
  }

  renderFeedMapping(feedsSummary, carrierAsnFiles);

  // ── Carrier ASN list ──
  asnList.innerHTML = '';
  if (carrierAsnFiles.length) {
    asnSection.style.display = 'block';
    carrierAsnFiles.forEach((f, i) => {
      const xmlId   = `carrier-xml-${i}`;
      const linesId = `carrier-lines-${i}`;
      const dlUrl   = `${API}/feed-raw?type=carrier&ref=${encodeURIComponent(f.filename)}&download=1`;
      const groups  = f.asnGroups || [];
      const totalLines = groups.reduce((s, g) => s + (g.lines || []).length, 0);

      // Build per-group line tables
      let linesHtml = '';
      for (const g of groups) {
        const lineRows = (g.lines || []).map(l => `<tr>
          <td><strong>${l.sku}</strong></td>
          <td style="font-family:monospace;font-size:11px">${l.ean || '—'}</td>
          <td>${l.description || '—'}</td>
          <td>${l.size   || '—'}</td>
          <td>${l.colour || '—'}</td>
          <td style="text-align:right">${l.quantity ?? '—'}</td>
          <td>${l.packFormat === 'H' ? 'Hanging' : 'Flat'}</td>
        </tr>`).join('');
        linesHtml += `<div style="margin-bottom:${groups.length > 1 ? '10px' : '0'}">
          <div style="font-size:11px;color:#1a5276;font-weight:600;margin-bottom:4px">
            ASN: <strong>${g.asnId}</strong> &nbsp;·&nbsp; FC: ${g.fcId} &nbsp;·&nbsp; Ship: ${g.shipDate}${g.supplier ? ` &nbsp;·&nbsp; ${g.supplier}` : ''}
          </div>
          <table class="lines-sub-table">
            <thead><tr><th>SKU</th><th>EAN</th><th>Description</th><th>Size</th><th>Colour</th><th>Qty</th><th>Pack</th></tr></thead>
            <tbody>${lineRows}</tbody>
          </table>
        </div>`;
      }

      const item = document.createElement('div');
      item.className = 'carrier-asn-item';
      item.innerHTML = `
        <div class="carrier-asn-header">
          <div>
            <div class="carrier-asn-info">&#128196; ${f.filename}</div>
            <div class="carrier-asn-ref">PO: ${f.poRef}${f.blobPath ? ` &nbsp;&middot;&nbsp; <span style="color:#aaa">${f.blobPath}</span>` : ''}</div>
          </div>
          <div style="display:flex;gap:4px;align-items:center;flex-wrap:wrap">
            ${totalLines > 0 ? `<button class="btn-view-xml" onclick="toggleCarrierLines('${linesId}',this)">📋 ${totalLines} Lines</button>` : ''}
            <button class="btn-view-xml" onclick="toggleCarrierXml('${f.filename}','${xmlId}',this)">View XML</button>
            <a class="btn-view-xml" href="${dlUrl}" download style="text-decoration:none">&#11015; Download</a>
          </div>
        </div>
        <div class="carrier-asn-lines" id="${linesId}" style="display:none">
          <div class="lines-expand-wrap">${linesHtml || '<em style="color:#999;font-size:12px">No line items parsed</em>'}</div>
        </div>
        <div class="carrier-asn-xml" id="${xmlId}"><pre class="feed-xml-pre">Loading…</pre></div>`;
      asnList.appendChild(item);
    });
  } else {
    asnSection.style.display = 'none';
  }
}

async function togglePoXml(orderId, rowId, preId, btn) {
  const row = document.getElementById(rowId);
  const pre = document.getElementById(preId);
  if (row.classList.toggle('open')) {
    btn.textContent = 'Hide XML';
    if (pre.textContent === 'Loading…') {
      try {
        const r = await fetch(`${API}/feed-raw?type=po&ref=${encodeURIComponent(orderId)}`);
        pre.textContent = r.ok ? await r.text() : `Error: ${(await r.json()).error}`;
      } catch (e) { pre.textContent = `Error: ${e.message}`; }
    }
  } else {
    btn.textContent = 'View XML';
  }
}
window.togglePoXml = togglePoXml;

async function toggleCarrierXml(filename, xmlId, btn) {
  const wrap = document.getElementById(xmlId);
  const pre  = wrap.querySelector('pre');
  if (wrap.classList.toggle('open')) {
    btn.textContent = 'Hide XML';
    if (pre.textContent === 'Loading…') {
      try {
        const r = await fetch(`${API}/feed-raw?type=carrier&ref=${encodeURIComponent(filename)}`);
        pre.textContent = r.ok ? await r.text() : `Error: ${(await r.json()).error}`;
      } catch (e) { pre.textContent = `Error: ${e.message}`; }
    }
  } else {
    btn.textContent = 'View XML';
  }
}
window.toggleCarrierXml = toggleCarrierXml;

function togglePoLines(linesId, btn) {
  const row = document.getElementById(linesId);
  if (row.classList.toggle('open')) {
    btn._origText = btn.textContent;
    btn.textContent = '✕ Hide Lines';
  } else {
    btn.textContent = btn._origText || '📋 Lines';
  }
}
window.togglePoLines = togglePoLines;

function toggleCarrierLines(linesId, btn) {
  const wrap = document.getElementById(linesId);
  const open = wrap.style.display !== 'block';
  wrap.style.display = open ? 'block' : 'none';
  if (open) { btn._origText = btn.textContent; btn.textContent = '✕ Hide Lines'; }
  else { btn.textContent = btn._origText || btn.textContent; }
}
window.toggleCarrierLines = toggleCarrierLines;

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
