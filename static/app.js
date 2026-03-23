const state = {
  uploads: [],
  initiatives: [],
  filtered: [],
  selectedUploadId: null,
  selectedItemId: null
};

const els = {
  headerMeta: document.getElementById("headerMeta"),
  fileInput: document.getElementById("fileInput"),
  uploadBtn: document.getElementById("uploadBtn"),
  uploadStatus: document.getElementById("uploadStatus"),
  uploadSelect: document.getElementById("uploadSelect"),
  uploadInfo: document.getElementById("uploadInfo"),
  reloadBtn: document.getElementById("reloadBtn"),
  searchInput: document.getElementById("searchInput"),
  sectionFilter: document.getElementById("sectionFilter"),
  themeFilter: document.getElementById("themeFilter"),
  ownerFilter: document.getElementById("ownerFilter"),
  impactFilter: document.getElementById("impactFilter"),
  relevanceFilter: document.getElementById("relevanceFilter"),
  fmrukOnlyFilter: document.getElementById("fmrukOnlyFilter"),
  tableBody: document.getElementById("tableBody"),
  detailPanel: document.getElementById("detailPanel"),
  kpiTotal: document.getElementById("kpiTotal"),
  kpiHighRelevance: document.getElementById("kpiHighRelevance"),
  kpiHighImpact: document.getElementById("kpiHighImpact"),
  kpiRelevant: document.getElementById("kpiRelevant"),
  topThemesList: document.getElementById("topThemesList"),
  topOwnersList: document.getElementById("topOwnersList")
};

document.addEventListener("DOMContentLoaded", () => {
  init().catch(err => {
    console.error(err);
    if (els.uploadStatus) {
      els.uploadStatus.textContent = "Front-end initialisation failed. Check /health and browser console.";
    }
  });
});

async function init() {
  bindEvents();
  await loadUploads();
}

function bindEvents() {
  els.uploadBtn?.addEventListener("click", uploadFile);
  els.uploadSelect?.addEventListener("change", onUploadChange);
  els.reloadBtn?.addEventListener("click", reloadCurrent);
  els.searchInput?.addEventListener("input", applyFilters);
  els.sectionFilter?.addEventListener("change", applyFilters);
  els.themeFilter?.addEventListener("change", applyFilters);
  els.ownerFilter?.addEventListener("change", applyFilters);
  els.impactFilter?.addEventListener("change", applyFilters);
  els.relevanceFilter?.addEventListener("change", applyFilters);
  els.fmrukOnlyFilter?.addEventListener("change", applyFilters);
}

async function uploadFile() {
  const file = els.fileInput.files[0];
  if (!file) {
    els.uploadStatus.textContent = "Please select a PDF or XLSX file first.";
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  els.uploadStatus.textContent = "Uploading and analysing...";

  try {
    const res = await fetch("/api/upload", {
      method: "POST",
      body: formData
    });

    const data = await res.json();

    if (!data.ok) {
      els.uploadStatus.textContent = data.error || "Upload failed.";
      return;
    }

    els.uploadStatus.textContent = `Upload complete. Parsed ${data.initiative_count} initiatives from ${data.filename}.`;
    await loadUploads(data.upload_id);
  } catch (err) {
    console.error(err);
    els.uploadStatus.textContent = "Upload failed due to a server error.";
  }
}

async function loadUploads(preferredUploadId = null) {
  const res = await fetch("/api/uploads", { cache: "no-store" });
  const uploads = await res.json();

  state.uploads = uploads;
  renderUploads();

  if (preferredUploadId) {
    state.selectedUploadId = preferredUploadId;
  } else if (!state.selectedUploadId && uploads.length) {
    state.selectedUploadId = uploads[0].id;
  }

  if (state.selectedUploadId) {
    els.uploadSelect.value = String(state.selectedUploadId);
    await loadInitiatives(state.selectedUploadId);
  } else {
    els.headerMeta.textContent = "No upload loaded";
    clearTableAndSummary();
  }
}

function renderUploads() {
  const current = state.selectedUploadId ? String(state.selectedUploadId) : "";

  els.uploadSelect.innerHTML = `<option value="">Select upload</option>`;
  state.uploads.forEach(u => {
    const option = document.createElement("option");
    option.value = String(u.id);
    option.textContent = `${u.filename} (${u.initiative_count} items)`;
    if (option.value === current) option.selected = true;
    els.uploadSelect.appendChild(option);
  });
}

async function onUploadChange() {
  const uploadId = Number(els.uploadSelect.value || 0);
  if (!uploadId) return;
  state.selectedUploadId = uploadId;
  await loadInitiatives(uploadId);
}

async function reloadCurrent() {
  if (!state.selectedUploadId) return;
  await loadInitiatives(state.selectedUploadId);
}

async function loadInitiatives(uploadId) {
  const [itemsRes, summaryRes] = await Promise.all([
    fetch(`/api/initiatives?upload_id=${uploadId}`, { cache: "no-store" }),
    fetch(`/api/summary?upload_id=${uploadId}`, { cache: "no-store" })
  ]);

  state.initiatives = await itemsRes.json();
  const summary = await summaryRes.json();

  const upload = state.uploads.find(x => x.id === uploadId);
  els.headerMeta.textContent = upload
    ? `${upload.filename} | ${upload.initiative_count} initiatives | uploaded ${formatDate(upload.uploaded_at)}`
    : "Upload loaded";

  els.uploadInfo.textContent = upload
    ? `File type: ${String(upload.file_type || "").toUpperCase()} | Stored upload ID: ${upload.id}`
    : "Choose a stored upload to inspect.";

  renderSummary(summary);
  populateFilters(state.initiatives);
  applyFilters();
}

function populateFilters(items) {
  fillSelect(els.sectionFilter, unique(items.map(x => x.section_name)));
  fillSelect(els.themeFilter, unique(items.map(x => x.theme)));
  fillSelect(els.ownerFilter, unique(items.map(x => x.primary_owner)));
}

function fillSelect(selectEl, values) {
  const current = selectEl.value;
  selectEl.innerHTML = `<option value="">All</option>`;
  values.forEach(v => {
    const option = document.createElement("option");
    option.value = v;
    option.textContent = v;
    if (v === current) option.selected = true;
    selectEl.appendChild(option);
  });
}

function unique(arr) {
  return [...new Set(arr.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function applyFilters() {
  const q = els.searchInput.value.trim().toLowerCase();
  const section = els.sectionFilter.value;
  const theme = els.themeFilter.value;
  const owner = els.ownerFilter.value;
  const impact = els.impactFilter.value;
  const minRelevance = Number(els.relevanceFilter.value || 0);
  const fmrukOnly = els.fmrukOnlyFilter.checked;

  let items = state.initiatives.filter(item => {
    const haystack = [
      item.section_name,
      item.subcategory,
      item.initiative_title,
      item.initiative_description,
      item.theme,
      item.internal_sub_theme,
      item.primary_owner,
      item.secondary_owner,
      item.potential_business_impact,
      item.rationale
    ].join(" ").toLowerCase();

    return (
      (!q || haystack.includes(q)) &&
      (!section || item.section_name === section) &&
      (!theme || item.theme === theme) &&
      (!owner || item.primary_owner === owner) &&
      (!impact || item.impact_level === impact) &&
      ((item.relevance_score || 0) >= minRelevance) &&
      (!fmrukOnly || item.is_fmruk_relevant === 1)
    );
  });

  items.sort((a, b) => (b.relevance_score || 0) - (a.relevance_score || 0));
  state.filtered = items;

  if (!items.find(x => x.id === state.selectedItemId)) {
    state.selectedItemId = items[0]?.id || null;
  }

  renderTable(items);
  const selected = items.find(x => x.id === state.selectedItemId);
  renderDetail(selected);
}

function renderSummary(summary) {
  els.kpiTotal.textContent = summary.total || 0;
  els.kpiHighRelevance.textContent = summary.high_relevance || 0;
  els.kpiHighImpact.textContent = summary.high_impact || 0;
  els.kpiRelevant.textContent = summary.fmruk_relevant || 0;

  renderList(els.topThemesList, summary.top_themes || []);
  renderList(els.topOwnersList, summary.top_owners || []);
}

function renderList(target, items) {
  target.innerHTML = "";
  if (!items.length) {
    const li = document.createElement("li");
    li.textContent = "None";
    target.appendChild(li);
    return;
  }

  items.forEach(item => {
    const li = document.createElement("li");
    li.textContent = `${item.name} (${item.count})`;
    target.appendChild(li);
  });
}

function renderTable(items) {
  els.tableBody.innerHTML = "";

  if (!items.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6">No initiatives match the current filters.</td>`;
    els.tableBody.appendChild(tr);
    return;
  }

  items.forEach(item => {
    const tr = document.createElement("tr");
    if (item.id === state.selectedItemId) tr.classList.add("selected");

    tr.innerHTML = `
      <td>${priorityBadge(item.relevance_score || 0)}</td>
      <td>${escapeHtml(item.section_name || "")}</td>
      <td>${escapeHtml(item.theme || "")}</td>
      <td>${impactBadge(item.impact_level || "Low")}</td>
      <td>${escapeHtml(item.primary_owner || "")}</td>
      <td>${escapeHtml(item.initiative_title || "")}</td>
    `;

    tr.addEventListener("click", () => {
      state.selectedItemId = item.id;
      renderTable(state.filtered);
      renderDetail(item);
    });

    els.tableBody.appendChild(tr);
  });
}

function renderDetail(item) {
  if (!item) {
    els.detailPanel.innerHTML = `Upload a file or choose a saved upload, then select an initiative.`;
    return;
  }

  els.detailPanel.innerHTML = `
    <h2 class="detail-title">${escapeHtml(item.initiative_title || "")}</h2>

    <div class="detail-grid">
      <div class="detail-box">
        <div class="detail-label">Section</div>
        <div class="detail-value">${escapeHtml(item.section_name || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Subcategory</div>
        <div class="detail-value">${escapeHtml(item.subcategory || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Lead</div>
        <div class="detail-value">${escapeHtml(item.lead_regulator || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Indicative Impact</div>
        <div class="detail-value">${escapeHtml(item.indicative_impact_on_firms || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Theme</div>
        <div class="detail-value">${escapeHtml(item.theme || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Sub-theme</div>
        <div class="detail-value">${escapeHtml(item.internal_sub_theme || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Primary Owner</div>
        <div class="detail-value">${escapeHtml(item.primary_owner || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Secondary Owner</div>
        <div class="detail-value">${escapeHtml(item.secondary_owner || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Relevance</div>
        <div class="detail-value">${escapeHtml(String(item.relevance_score || 0))}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Impact Level</div>
        <div class="detail-value">${escapeHtml(item.impact_level || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Timing Bucket</div>
        <div class="detail-value">${escapeHtml(item.timing_bucket || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">New</div>
        <div class="detail-value">${escapeHtml(item.is_new || "N/A")}</div>
      </div>
    </div>

    <div class="detail-block">
      <strong>Potential Business Impact:</strong><br />
      ${escapeHtml(item.potential_business_impact || "N/A")}
    </div>

    <div class="detail-block">
      <strong>Expected Key Milestones:</strong><br />
      ${escapeHtml(item.expected_key_milestones || "N/A")}
    </div>

    <div class="detail-block">
      <strong>Initiative Description:</strong><br />
      ${escapeHtml(item.initiative_description || "N/A")}
    </div>

    <div class="detail-block">
      <strong>Why it matters to FMRUK:</strong><br />
      ${escapeHtml(item.rationale || "N/A")}
    </div>

    <div class="detail-block">
      <strong>Suggested Action:</strong><br />
      ${escapeHtml(item.suggested_action || "N/A")}
    </div>
  `;
}

function priorityBadge(score) {
  if (score >= 80) return `<span class="badge badge-high">High</span>`;
  if (score >= 60) return `<span class="badge badge-medium">Medium</span>`;
  return `<span class="badge badge-low">Low</span>`;
}

function impactBadge(level) {
  const map = {
    High: "badge-high",
    Medium: "badge-medium",
    Low: "badge-low"
  };
  return `<span class="badge ${map[level] || "badge-low"}">${escapeHtml(level)}</span>`;
}

function clearTableAndSummary() {
  els.tableBody.innerHTML = `<tr><td colspan="6">No uploads available.</td></tr>`;
  els.detailPanel.innerHTML = `Upload a file or choose a saved upload, then select an initiative.`;
  els.kpiTotal.textContent = "0";
  els.kpiHighRelevance.textContent = "0";
  els.kpiHighImpact.textContent = "0";
  els.kpiRelevant.textContent = "0";
  els.topThemesList.innerHTML = "<li>None</li>";
  els.topOwnersList.innerHTML = "<li>None</li>";
}

function formatDate(value) {
  if (!value) return "Unknown";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return value;
  return d.toLocaleString("en-GB");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
