/* ── Drop Zone ──────────────────────────────────────────────── */
const dropZone   = document.getElementById('drop-zone');
const fileInput  = document.getElementById('file-input');
const processing = document.getElementById('processing');
const reviewCard = document.getElementById('review-card');
const uploadCard = document.getElementById('upload-card');

['dragenter','dragover'].forEach(evt =>
  dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.add('over'); })
);
['dragleave','drop'].forEach(evt =>
  dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.remove('over'); })
);
dropZone.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) handleFile(f); });
fileInput.addEventListener('change', () => { if (fileInput.files[0]) handleFile(fileInput.files[0]); });
dropZone.addEventListener('click', () => fileInput.click());

/* ── File Handling ──────────────────────────────────────────── */
async function handleFile(file) {
  if (!file.name.toLowerCase().endsWith('.pdf')) {
    showToast('Only PDF files are supported.', 'error');
    return;
  }
  dropZone.classList.add('hidden');
  processing.classList.remove('hidden');

  try {
    const form = new FormData();
    form.append('file', file);
    const res  = await fetch('/api/extract', { method: 'POST', body: form });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    populateForm(data);
    processing.classList.add('hidden');
    reviewCard.classList.remove('hidden');
  } catch (err) {
    processing.classList.add('hidden');
    dropZone.classList.remove('hidden');
    showToast('Extraction failed: ' + err.message, 'error');
  }
}

/* ── Populate Form ──────────────────────────────────────────── */
function populateForm(data) {
  document.getElementById('f-vendor').value  = data.vendor          ?? '';
  document.getElementById('f-invoice').value = data.invoice_number  ?? '';
  document.getElementById('f-wording').value = data.invoice_wording ?? '';
  document.getElementById('f-misc').value    = data.misc ?? '';

  // Date
  const dateVal = data.date ?? new Date().toISOString().slice(0, 10);
  document.getElementById('f-date').value = dateVal;

  // Lease/Rental (AI-determined from unit lookup)
  const leaseEl = document.getElementById('f-lease');
  leaseEl.value = (data.lease_or_rental || 'Lease');

  // Render line items
  renderLineItems(data.line_items || []);
}

/* ── Line Items ─────────────────────────────────────────────── */
function renderLineItems(items) {
  const container = document.getElementById('line-items-list');

  if (!items.length) {
    container.innerHTML = '<p style="color:var(--text-muted);padding:12px;">No line items extracted.</p>';
    return;
  }

  container.innerHTML = items.map((item, idx) => {
    const type     = (item.type     || 'repair').toLowerCase();
    const category = (item.category || 'parts').toLowerCase();
    const cost     = parseFloat(item.cost || 0).toFixed(2);
    const pmActive     = type === 'pm'     ? 'active-pm'     : '';
    const repairActive = type === 'repair' ? 'active-repair' : '';
    const rebillActive = type === 'rebill' ? 'active-rebill' : '';
    const partsActive  = category === 'parts'  ? 'active-parts'  : '';
    const laborActive  = category === 'labor'  ? 'active-labor'  : '';

    return `
      <div class="line-item-row" data-idx="${idx}">
        <span class="li-desc">${escHtml(item.description || '')}</span>
        <span class="li-cost">$${Number(cost).toLocaleString('en-US', {minimumFractionDigits:2})}</span>
        <div class="type-toggle">
          <button type="button" class="type-btn ${pmActive}" data-idx="${idx}" data-type="pm"
            onclick="setType(${idx}, 'pm')">Int. PMs</button>
          <button type="button" class="type-btn ${repairActive}" data-idx="${idx}" data-type="repair"
            onclick="setType(${idx}, 'repair')">Int. Repairs</button>
          <button type="button" class="type-btn ${rebillActive}" data-idx="${idx}" data-type="rebill"
            onclick="setType(${idx}, 'rebill')">Rebill</button>
        </div>
        <div class="category-toggle">
          <button type="button" class="cat-btn ${partsActive}" data-idx="${idx}" data-cat="parts"
            onclick="setCategory(${idx}, 'parts')">Parts</button>
          <button type="button" class="cat-btn ${laborActive}" data-idx="${idx}" data-cat="labor"
            onclick="setCategory(${idx}, 'labor')">Labor</button>
        </div>
      </div>`;
  }).join('');

  // Store items data for retrieval
  container.dataset.items = JSON.stringify(items);
}

function setType(idx, type) {
  const row = document.querySelector(`.line-item-row[data-idx="${idx}"]`);
  row.querySelectorAll('.type-btn').forEach(b => {
    b.className = 'type-btn';
    if (b.dataset.type === type) {
      if (type === 'pm')     b.classList.add('active-pm');
      if (type === 'repair') b.classList.add('active-repair');
      if (type === 'rebill') b.classList.add('active-rebill');
    }
  });
}

function setCategory(idx, cat) {
  const row = document.querySelector(`.line-item-row[data-idx="${idx}"]`);
  row.querySelectorAll('.cat-btn').forEach(b => {
    b.className = 'cat-btn';
    if (b.dataset.cat === cat) {
      b.classList.add(cat === 'parts' ? 'active-parts' : 'active-labor');
    }
  });
}

function getLineItems() {
  const container = document.getElementById('line-items-list');
  const base      = JSON.parse(container.dataset.items || '[]');
  return base.map((item, idx) => {
    const row       = document.querySelector(`.line-item-row[data-idx="${idx}"]`);
    const activeBtn = row?.querySelector('.type-btn.active-pm, .type-btn.active-repair, .type-btn.active-rebill');
    const activeCat = row?.querySelector('.cat-btn.active-parts, .cat-btn.active-labor');
    const type      = activeBtn?.dataset.type || 'repair';
    const category  = activeCat?.dataset.cat  || 'parts';
    return { ...item, type, category };
  });
}

/* ── Discard ────────────────────────────────────────────────── */
document.getElementById('btn-discard').addEventListener('click', () => {
  reviewCard.classList.add('hidden');
  dropZone.classList.remove('hidden');
  fileInput.value = '';
});

/* ── Generate / Download ─────────────────────────────────────── */
document.getElementById('btn-generate').addEventListener('click', async () => {
  const payload = {
    vendor:           document.getElementById('f-vendor').value,
    invoice_number:   document.getElementById('f-invoice').value,
    date:             document.getElementById('f-date').value,
    lease_or_rental:  document.getElementById('f-lease').value,
    taxable:          document.getElementById('f-taxable').value,
    customer:         document.getElementById('f-customer').value,
    invoice_wording:  document.getElementById('f-wording').value,
    line_items:       getLineItems(),
    misc:             parseFloat(document.getElementById('f-misc').value) || 0,
  };

  try {
    const res = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) throw new Error(await res.text());

    // Trigger download
    const blob     = await res.blob();
    const url      = URL.createObjectURL(blob);
    const a        = document.createElement('a');
    const disp     = res.headers.get('Content-Disposition') || '';
    const match    = disp.match(/filename=(.+)/);
    a.href         = url;
    a.download     = match ? match[1] : 'Rebill.xlsm';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('Rebill sheet downloaded!', 'success');
  } catch (err) {
    showToast('Generation failed: ' + err.message, 'error');
  }
});

/* ── Helpers ────────────────────────────────────────────────── */
function escHtml(str) {
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function showToast(msg, type = 'success') {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.className = `toast ${type}`;
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), 3500);
}
