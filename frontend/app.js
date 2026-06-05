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
const blobSearchInput  = document.getElementById('blobSearchInput');
const btnBlobSearch    = document.getElementById('btnBlobSearch');
const blobResults      = document.getElementById('blobResults');

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
    state.asnRefs = data.asnRefs || [];

    let html = `✅ Parsed <strong>${data.rowCount}</strong> rows — found <strong>${state.poRefs.length}</strong> PO ref(s) and <strong>${state.asnRefs.length}</strong> ASN ref(s).`;
    if (data.validationErrors && data.validationErrors.length) {
      html += `<br/>⚠️ Validation warnings:<br/>${data.validationErrors.join('<br/>')}`;
    }
    setStatus(1, 'success', html);
    setBadge(1, 'done');

    // Show extracted refs
    refsPreview.innerHTML = `
      <div style="margin-top:10px">
        <strong style="font-size:12px;color:#555">PO Refs extracted:</strong>
        <div class="ref-tags">${state.poRefs.map(r => `<span class="tag">PO: ${r}</span>`).join('')}</div>
        <strong style="font-size:12px;color:#555;display:block;margin-top:8px">ASN Refs extracted:</strong>
        <div class="ref-tags">${state.asnRefs.map(r => `<span class="tag">ASN: ${r}</span>`).join('')}</div>
      </div>`;

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

// ── Step 2: Blob Search ───────────────────────────────────────────────────────
async function runBlobSearch() {
  const query = blobSearchInput.value.trim();
  blobResults.innerHTML = '<div style="padding:8px;color:#888;font-size:12px">Searching…</div>';
  try {
    const res  = await fetch(`${API}/search-blob?query=${encodeURIComponent(query)}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    if (!data.results || data.results.length === 0) {
      blobResults.innerHTML = '<div style="padding:8px;color:#888;font-size:12px">No blobs found.</div>';
      return;
    }
    blobResults.innerHTML = data.results.map(b => `
      <div class="blob-result-item">
        <span class="blob-result-name">${b.name}</span>
        <span class="blob-result-meta">${b.size != null ? (b.size / 1024).toFixed(1) + ' KB' : ''}${b.lastModified ? ' · ' + new Date(b.lastModified).toLocaleDateString() : ''}</span>
      </div>`).join('');
  } catch (err) {
    blobResults.innerHTML = `<div style="padding:8px;color:#C0392B;font-size:12px">❌ ${err.message}</div>`;
  }
}

btnBlobSearch.addEventListener('click', runBlobSearch);
blobSearchInput.addEventListener('keydown', e => { if (e.key === 'Enter') runBlobSearch(); });

// ── Step 2: Fetch Feeds ────────────────────────────────────────────────────────
btnFetchFeeds.addEventListener('click', async () => {
  setLoading(btnFetchFeeds, true);
  setStatus(2, 'loading', '⏳ Fetching PO & ASN feeds from Azure Blob Storage…');

  try {
    const res  = await fetch(`${API}/fetch-feeds`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ poRefs: state.poRefs, asnRefs: state.asnRefs })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Unknown error');

    state.feedsFetched = true;
    const modeTag = data.localMode ? ' <em style="color:#e67e22">[LOCAL MODE — reading from samples/feeds/]</em>' : '';
    let html = `✅ Fetched <strong>${data.poFeedCount}</strong> PO feed(s) and <strong>${data.asnFeedCount}</strong> ASN feed(s).${modeTag}`;
    if (data.summary) {
      if (data.summary.posFound.length) html += `<br/>POs found: ${data.summary.posFound.join(', ')}`;
      if (data.summary.asnsFound.length) html += `<br/>ASNs found: ${data.summary.asnsFound.join(', ')}`;
    }
    if (data.errors && data.errors.length) {
      html += `<br/>⚠️ ${data.errors.join('<br/>')}`;
    }
    setStatus(2, 'success', html);
    setBadge(2, 'done');
    btnBuildBible.disabled = false;
    setBadge(3, 'active');
  } catch (err) {
    setStatus(2, 'error', `❌ ${err.message}`);
  } finally {
    setLoading(btnFetchFeeds, false);
  }
});

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
    setStatus(3, 'success', `✅ Bible built with <strong>${data.masterRowCount}</strong> MASTER row(s).${dlLink}`);
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
