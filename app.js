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
  statusFilter: document.getElementById("statusFilter"),
  relevanceFilter: document.getElementById("relevanceFilter"),
  upcomingOnlyFilter: document.getElementById("upcomingOnlyFilter"),
  fmrukOnlyFilter: document.getElementById("fmrukOnlyFilter"),
  tableBody: document.getElementById("tableBody"),
  detailPanel: document.getElementById("detailPanel"),
  kpiTotal: document.getElementById("kpiTotal"),
  kpiHigh: document.getElementById("kpiHigh"),
  kpiUpcoming: document.getElementById("kpiUpcoming"),
  kpiAuthorities: document.getElementById("kpiAuthorities"),
  topThemesList: document.getElementById("topThemesList"),
  topSourcesList: document.getElementById("topSourcesList"),
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
    els.statusFilter,
    els.relevanceFilter,
    els.upcomingOnlyFilter,
    els.fmrukOnlyFilter
  ].forEach(el => {
    el.addEventListener("input", applyFilters);
    el.addEventListener("change", applyFilters);
  });

  els.clearFiltersBtn.addEventListener("click", () => {
    els.searchInput.value = "";
    els.authorityFilter.value = "";
    els.sourceFilter.value = "";
    els.themeFilter.value = "";
    els.statusFilter.value = "";
    els.relevanceFilter.value = "60";
    els.upcomingOnlyFilter.checked = true;
    els.fmrukOnlyFilter.checked = true;
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

  els.refreshViewBtn.addEventListener("click", () => {
    applyFilters();
  });

  els.downloadJsonBtn.addEventListener("click", () => {
    window.open("./data/regulatory_updates.json", "_blank");
  });
}

function populateFilters(items) {
  populateSelect(els.authorityFilter, uniqueSorted(items.map(x => x.authority)));
  populateSelect(els.sourceFilter, uniqueSorted(items.map(x => x.source)));
  populateSelect(els.themeFilter, uniqueSorted(items.map(x => x.theme)));
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
  const status = els.statusFilter.value;
  const minRelevance = Number(els.relevanceFilter.value || 0);
  const upcomingOnly = els.upcomingOnlyFilter.checked;
  const fmrukOnly = els.fmrukOnlyFilter.checked;

  let items = [...state.raw].filter(item => {
    const haystack = [
      item.title,
      item.summary,
      item.theme,
      item.authority,
      item.type,
      item.rationale
    ]
      .join(" ")
      .toLowerCase();

    const matchesQuery = !query || haystack.includes(query);
    const matchesAuthority = !authority || item.authority === authority;
    const matchesSource = !source || item.source === source;
    const matchesTheme = !theme || item.theme === theme;
    const matchesStatus = !status || item.status === status;
    const matchesRelevance = (item.relevance_score || 0) >= minRelevance;
    const matchesUpcoming = !upcomingOnly || ["Open", "Upcoming", "Pipeline"].includes(item.status);
    const matchesFmruk = !fmrukOnly || item.is_fmruk_relevant === true;

    return (
      matchesQuery &&
      matchesAuthority &&
      matchesSource &&
      matchesTheme &&
      matchesStatus &&
      matchesRelevance &&
      matchesUpcoming &&
      matchesFmruk
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
  const d = item.due_date || item.published_date || item.generated_date;
  if (!d) return 0;
  const t = new Date(d).getTime();
  return Number.isNaN(t) ? 0 : t;
}

function renderTable(items) {
  els.tableBody.innerHTML = "";

  if (!items.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="7" class="muted">No items match the current filters.</td>`;
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
      <td>${escapeHtml(item.type || "")}</td>
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
  els.kpiUpcoming.textContent = items.filter(x => ["Open", "Upcoming", "Pipeline"].includes(x.status)).length;
  els.kpiAuthorities.textContent = new Set(items.map(x => x.authority).filter(Boolean)).size;

  renderFrequencyList(
    els.topThemesList,
    items.map(x => x.theme).filter(Boolean)
  );

  renderFrequencyList(
    els.topSourcesList,
    items.map(x => x.source).filter(Boolean)
  );
}

function renderFrequencyList(target, values) {
  const counts = {};
  values.forEach(v => (counts[v] = (counts[v] || 0) + 1));

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
    els.detailPanel.innerHTML = `
      <div class="empty-state">No item selected.</div>
    `;
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
        <div class="detail-meta-label">Relevance Score</div>
        <div class="detail-meta-value">${escapeHtml(String(item.relevance_score || 0))}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Status</div>
        <div class="detail-meta-value">${escapeHtml(item.status || "N/A")}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Type</div>
        <div class="detail-meta-value">${escapeHtml(item.type || "N/A")}</div>
      </div>
      <div class="detail-meta">
        <div class="detail-meta-label">Date</div>
        <div class="detail-meta-value">${escapeHtml(formatDate(item.due_date || item.published_date || ""))}</div>
      </div>
    </div>

    <div class="detail-summary">
      <strong>Summary:</strong><br />
      ${escapeHtml(item.summary || "No summary available.")}
    </div>

    <div class="detail-rationale">
      <strong>Why it matters to FMRUK:</strong><br />
      ${escapeHtml(item.rationale || "No rationale generated.")}
    </div>

    <div class="detail-rationale">
      <strong>Internal owner suggestion:</strong><br />
      ${escapeHtml(item.suggested_owner || "Compliance / Risk / Legal review")}
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

function statusBadge(status) {
  const map = {
    Open: "status-open",
    Upcoming: "status-upcoming",
    Pipeline: "status-upcoming",
    Closed: "status-closed"
  };

  const className = map[status] || "status-upcoming";
  return `<span class="status-badge ${className}">${escapeHtml(status || "N/A")}</span>`;
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
