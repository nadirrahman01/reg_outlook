const state = {
  raw: [],
  filtered: [],
  selectedId: null,
  sortMode: "relevance"
};

const els = {
  lastUpdatedLabel: document.getElementById("lastUpdatedLabel"),
  searchInput: document.getElementById("searchInput"),
  authorityFilter: document.getElementById("authorityFilter"),
  sourceFilter: document.getElementById("sourceFilter"),
  themeFilter: document.getElementById("themeFilter"),
  subThemeFilter: document.getElementById("subThemeFilter"),
  ownerFilter: document.getElementById("ownerFilter"),
  statusFilter: document.getElementById("statusFilter"),
  relevanceFilter: document.getElementById("relevanceFilter"),
  openOnlyFilter: document.getElementById("openOnlyFilter"),
  fmrukOnlyFilter: document.getElementById("fmrukOnlyFilter"),
  highImpactOnlyFilter: document.getElementById("highImpactOnlyFilter"),
  tableBody: document.getElementById("tableBody"),
  detailPanel: document.getElementById("detailPanel"),
  kpiTotal: document.getElementById("kpiTotal"),
  kpiHigh: document.getElementById("kpiHigh"),
  kpiHighImpact: document.getElementById("kpiHighImpact"),
  kpiAuthorities: document.getElementById("kpiAuthorities"),
  topThemesList: document.getElementById("topThemesList"),
  topOwnersList: document.getElementById("topOwnersList"),
  clearFiltersBtn: document.getElementById("clearFiltersBtn"),
  sortByRelevanceBtn: document.getElementById("sortByRelevanceBtn"),
  sortByDateBtn: document.getElementById("sortByDateBtn"),
  refreshViewBtn: document.getElementById("refreshViewBtn"),
  downloadJsonBtn: document.getElementById("downloadJsonBtn")
};

async function init() {
  try {
    const res = await fetch("./data/regulatory_updates.json", { cache: "no-store" });
    const payload = await res.json();

    state.raw = Array.isArray(payload.items) ? payload.items : [];
    els.lastUpdatedLabel.textContent = `Last refresh: ${payload.generated_at || "Unknown"}`;

    populateFilters(state.raw);
    bindEvents();
    applyFilters();
  } catch (err) {
    console.error(err);
    els.detailPanel.innerHTML = `
      <div class="empty-state">
        Unable to load regulatory data. Check that <strong>data/regulatory_updates.json</strong> exists and is valid JSON.
      </div>
    `;
  }
}

function bindEvents() {
  [
    els.searchInput,
    els.authorityFilter,
    els.sourceFilter,
    els.themeFilter,
    els.subThemeFilter,
    els.ownerFilter,
    els.statusFilter,
    els.relevanceFilter,
    els.openOnlyFilter,
    els.fmrukOnlyFilter,
    els.highImpactOnlyFilter
  ].forEach(el => {
    el.addEventListener("input", applyFilters);
    el.addEventListener("change", applyFilters);
  });

  els.clearFiltersBtn.addEventListener("click", () => {
    els.searchInput.value = "";
    els.authorityFilter.value = "";
    els.sourceFilter.value = "";
    els.themeFilter.value = "";
    els.subThemeFilter.value = "";
    els.ownerFilter.value = "";
    els.statusFilter.value = "";
    els.relevanceFilter.value = "60";
    els.openOnlyFilter.checked = true;
    els.fmrukOnlyFilter.checked = true;
    els.highImpactOnlyFilter.checked = false;
    state.sortMode = "relevance";
    applyFilters();
  });

  els.sortByRelevanceBtn.addEventListener("click", () => {
    state.sortMode = "relevance";
    applyFilters();
  });

  els.sortByDateBtn.addEventListener("click", () => {
    state.sortMode = "date";
    applyFilters();
  });

  els.refreshViewBtn.addEventListener("click", applyFilters);

  els.downloadJsonBtn.addEventListener("click", () => {
    window.open("./data/regulatory_updates.json", "_blank");
  });
}

function populateFilters(items) {
  populateSelect(els.authorityFilter, uniqueSorted(items.map(x => x.authority)));
  populateSelect(els.sourceFilter, uniqueSorted(items.map(x => x.source)));
  populateSelect(els.themeFilter, uniqueSorted(items.map(x => x.theme)));
  populateSelect(els.subThemeFilter, uniqueSorted(items.map(x => x.sub_theme)));
  populateSelect(els.ownerFilter, uniqueSorted(items.map(x => x.primary_owner)));
  populateSelect(els.statusFilter, uniqueSorted(items.map(x => x.status)));
}

function populateSelect(selectEl, values) {
  values.forEach(value => {
    if (!value) return;
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  });
}

function uniqueSorted(arr) {
  return [...new Set(arr.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function applyFilters() {
  const query = els.searchInput.value.trim().toLowerCase();
  const authority = els.authorityFilter.value;
  const source = els.sourceFilter.value;
  const theme = els.themeFilter.value;
  const subTheme = els.subThemeFilter.value;
  const owner = els.ownerFilter.value;
  const status = els.statusFilter.value;
  const minRelevance = Number(els.relevanceFilter.value || 0);
  const openOnly = els.openOnlyFilter.checked;
  const fmrukOnly = els.fmrukOnlyFilter.checked;
  const highImpactOnly = els.highImpactOnlyFilter.checked;

  let items = [...state.raw].filter(item => {
    const haystack = [
      item.title,
      item.summary,
      item.theme,
      item.sub_theme,
      item.authority,
      item.type,
      item.rationale,
      item.potential_business_impact,
      item.primary_owner,
      item.secondary_owner
    ].join(" ").toLowerCase();

    const matchesQuery = !query || haystack.includes(query);
    const matchesAuthority = !authority || item.authority === authority;
    const matchesSource = !source || item.source === source;
    const matchesTheme = !theme || item.theme === theme;
    const matchesSubTheme = !subTheme || item.sub_theme === subTheme;
    const matchesOwner = !owner || item.primary_owner === owner;
    const matchesStatus = !status || item.status === status;
    const matchesRelevance = (item.relevance_score || 0) >= minRelevance;
    const matchesOpen =
      !openOnly || ["New", "Open Consultation", "Upcoming Implementation", "Supervisory Attention", "Final Rules / Policy"].includes(item.status);
    const matchesFmruk = !fmrukOnly || item.is_fmruk_relevant === true;
    const matchesHighImpact = !highImpactOnly || (item.impact_level || "") === "High";

    return (
      matchesQuery &&
      matchesAuthority &&
      matchesSource &&
      matchesTheme &&
      matchesSubTheme &&
      matchesOwner &&
      matchesStatus &&
      matchesRelevance &&
      matchesOpen &&
      matchesFmruk &&
      matchesHighImpact
    );
  });

  items.sort((a, b) => {
    if (state.sortMode === "date") {
      return getSortDate(b) - getSortDate(a);
    }
    return (b.relevance_score || 0) - (a.relevance_score || 0) || getSortDate(b) - getSortDate(a);
  });

  state.filtered = items;

  if (!items.find(x => x.id === state.selectedId)) {
    state.selectedId = items[0]?.id || null;
  }

  renderTable(items);
  renderKPIs(items);
  renderDetail(items.find(x => x.id === state.selectedId));
}

function getSortDate(item) {
  const d = item.due_date || item.published_date || "";
  const t = new Date(d).getTime();
  return Number.isNaN(t) ? 0 : t;
}

function renderTable(items) {
  els.tableBody.innerHTML = "";

  if (!items.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="8" class="muted">No items match the current filters.</td>`;
    els.tableBody.appendChild(tr);
    return;
  }

  items.forEach(item => {
    const tr = document.createElement("tr");
    if (item.id === state.selectedId) tr.classList.add("selected");

    tr.innerHTML = `
      <td>${priorityBadge(item.relevance_score || 0)}</td>
      <td>${escapeHtml(formatDate(item.due_date || item.published_date || ""))}</td>
      <td>${escapeHtml(item.authority || "")}</td>
      <td>${escapeHtml(item.theme || "")}</td>
      <td>${impactBadge(item.impact_level || "Low")}</td>
      <td>${escapeHtml(item.primary_owner || "")}</td>
      <td>${escapeHtml(item.title || "")}</td>
      <td>${statusBadge(item.status || "")}</td>
    `;

    tr.addEventListener("click", () => {
      state.selectedId = item.id;
      renderTable(state.filtered);
      renderDetail(item);
    });

    els.tableBody.appendChild(tr);
  });
}

function renderKPIs(items) {
  els.kpiTotal.textContent = items.length;
  els.kpiHigh.textContent = items.filter(x => (x.relevance_score || 0) >= 80).length;
  els.kpiHighImpact.textContent = items.filter(x => x.impact_level === "High").length;
  els.kpiAuthorities.textContent = new Set(items.map(x => x.authority).filter(Boolean)).size;

  renderFrequencyList(els.topThemesList, items.map(x => x.theme).filter(Boolean));
  renderFrequencyList(els.topOwnersList, items.map(x => x.primary_owner).filter(Boolean));
}

function renderFrequencyList(target, values) {
  const counts = {};
  values.forEach(v => counts[v] = (counts[v] || 0) + 1);

  target.innerHTML = "";

  Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .forEach(([label, count]) => {
      const li = document.createElement("li");
      li.textContent = `${label} (${count})`;
      target.appendChild(li);
    });

  if (!target.children.length) {
    const li = document.createElement("li");
    li.textContent = "None";
    target.appendChild(li);
  }
}

function renderDetail(item) {
  if (!item) {
    els.detailPanel.innerHTML = `<div class="empty-state">No item selected.</div>`;
    return;
  }

  els.detailPanel.innerHTML = `
    <h2 class="detail-title">${escapeHtml(item.title || "")}</h2>

    <div class="detail-meta-grid">
      <div class="detail-meta">
        <div class="detail-meta-label">Authority</div>
        <div class="detail-meta-value">${escapeHtml(item.authority || "N/A")}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Theme</div>
        <div class="detail-meta-value">${escapeHtml(item.theme || "N/A")}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Sub-theme</div>
        <div class="detail-meta-value">${escapeHtml(item.sub_theme || "N/A")}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Relevance</div>
        <div class="detail-meta-value">${escapeHtml(String(item.relevance_score || 0))}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Impact</div>
        <div class="detail-meta-value">${escapeHtml(item.impact_level || "N/A")}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Primary owner</div>
        <div class="detail-meta-value">${escapeHtml(item.primary_owner || "N/A")}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Secondary owner</div>
        <div class="detail-meta-value">${escapeHtml(item.secondary_owner || "N/A")}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Status</div>
        <div class="detail-meta-value">${escapeHtml(item.status || "N/A")}</div>
      </div>
    </div>

    <div class="detail-block">
      <strong>Potential business impact:</strong><br />
      ${escapeHtml(item.potential_business_impact || "No business impact generated.")}
    </div>

    <div class="detail-block">
      <strong>Summary:</strong><br />
      ${escapeHtml(item.summary || "No summary available.")}
    </div>

    <div class="detail-block">
      <strong>Why it matters to FMRUK:</strong><br />
      ${escapeHtml(item.rationale || "No rationale generated.")}
    </div>

    <div class="detail-block">
      <strong>Suggested internal response:</strong><br />
      ${escapeHtml(item.suggested_action || "Review and assess applicability.")}
    </div>

    <div class="detail-links">
      <strong>Source:</strong>
      ${
        item.url
          ? `<a href="${item.url}" target="_blank" rel="noopener noreferrer">Open original source</a>`
          : `<span class="muted">No source URL</span>`
      }
    </div>
  `;
}

function priorityBadge(score) {
  let label = "Low";
  let className = "priority-low";

  if (score >= 80) {
    label = "High";
    className = "priority-high";
  } else if (score >= 60) {
    label = "Medium";
    className = "priority-medium";
  }

  return `<span class="priority-badge ${className}">${label}</span>`;
}

function impactBadge(level) {
  const mapping = {
    High: "impact-high",
    Medium: "impact-medium",
    Low: "impact-low"
  };
  return `<span class="impact-badge ${mapping[level] || "impact-low"}">${escapeHtml(level)}</span>`;
}

function statusBadge(status) {
  const mapping = {
    "New": "status-new",
    "Open Consultation": "status-open",
    "Upcoming Implementation": "status-upcoming",
    "Supervisory Attention": "status-upcoming",
    "Final Rules / Policy": "status-final",
    "Closed / Historic": "status-closed"
  };

  return `<span class="status-badge ${mapping[status] || "status-upcoming"}">${escapeHtml(status || "N/A")}</span>`;
}

function formatDate(value) {
  if (!value) return "N/A";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString("en-GB", {
    year: "numeric",
    month: "short",
    day: "2-digit"
  });
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

init();
