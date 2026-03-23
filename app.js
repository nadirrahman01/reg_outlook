const STORAGE_KEY = "fmruk_reg_grid_v1";

const state = {
  raw: [],
  filtered: [],
  selectedItemId: null,
  datasetMeta: null,
};

const THEME_RULES = [
  {
    theme: "Market Conduct / Trading",
    subTheme: "Transaction Reporting",
    keywords: ["transaction reporting", "mifir", "regulatory reporting", "reporting fields", "approved reporting mechanism", "arm"],
    impact: "May require review of transaction reporting controls, reconciliations, exception monitoring and governance.",
    primaryOwner: "Compliance",
    secondaryOwner: "Operations"
  },
  {
    theme: "Market Conduct / Trading",
    subTheme: "Best Execution",
    keywords: ["best execution", "rts 28", "execution quality", "venue analysis", "execution policy"],
    impact: "Could affect best execution governance, monitoring, policy wording and oversight of execution arrangements.",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal"
  },
  {
    theme: "Market Conduct / Trading",
    subTheme: "Market Abuse / Surveillance",
    keywords: ["market abuse", "mar", "surveillance", "inside information", "personal account dealing"],
    impact: "May require updates to surveillance controls, market abuse risk assessments, training and escalation processes.",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal"
  },
  {
    theme: "Operational Resilience",
    subTheme: "Operational Resilience",
    keywords: ["operational resilience", "important business service", "impact tolerance", "resilience"],
    impact: "May require review of resilience mapping, testing, scenario analysis and governance over important business services.",
    primaryOwner: "Risk",
    secondaryOwner: "Compliance"
  },
  {
    theme: "Operational Resilience",
    subTheme: "Outsourcing / Third Party Risk",
    keywords: ["outsourcing", "third party", "critical third party", "service provider", "vendor risk"],
    impact: "Could require updates to outsourcing registers, due diligence, contractual controls and oversight of service providers.",
    primaryOwner: "Risk",
    secondaryOwner: "Technology"
  },
  {
    theme: "Prudential / MIFIDPRU",
    subTheme: "Capital / Liquidity",
    keywords: ["mifidpru", "capital", "liquidity", "own funds", "icara", "prudential", "concentration risk"],
    impact: "May affect prudential assessments, capital planning, ICARA assumptions, monitoring and governance.",
    primaryOwner: "Finance",
    secondaryOwner: "Risk"
  },
  {
    theme: "Prudential / MIFIDPRU",
    subTheme: "Remuneration",
    keywords: ["remuneration", "bonus", "pay", "compensation", "malus", "clawback", "incentive"],
    impact: "Could require updates to remuneration frameworks, governance, documentation and control testing.",
    primaryOwner: "HR",
    secondaryOwner: "Compliance"
  },
  {
    theme: "Consumer / Conduct",
    subTheme: "Consumer Duty",
    keywords: ["consumer duty", "fair value", "good outcomes", "vulnerable customers", "consumer"],
    impact: "May require review of product governance, distribution oversight, client communications and conduct monitoring.",
    primaryOwner: "Compliance",
    secondaryOwner: "Product"
  },
  {
    theme: "AML / Sanctions / Financial Crime",
    subTheme: "AML / KYC / Sanctions",
    keywords: ["aml", "anti-money laundering", "sanctions", "kyc", "financial crime", "fraud", "bribery"],
    impact: "May require review of AML controls, KYC processes, sanctions screening, escalation and training.",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal"
  },
  {
    theme: "Governance / SMCR",
    subTheme: "SMCR / Governance",
    keywords: ["smcr", "senior managers", "certification regime", "governance", "board", "committee", "dear ceo"],
    impact: "Could affect governance frameworks, accountabilities, committee reporting and senior management oversight.",
    primaryOwner: "Compliance",
    secondaryOwner: "HR"
  },
  {
    theme: "ESG / Sustainability",
    subTheme: "Sustainability / SDR",
    keywords: ["sustainability", "sdr", "greenwashing", "climate", "esg", "transition plan", "tnfd", "tcfd", "srs"],
    impact: "May require updates to sustainability disclosures, product communications, governance and control evidence.",
    primaryOwner: "Legal",
    secondaryOwner: "Compliance"
  },
  {
    theme: "Data / AI / Technology",
    subTheme: "AI / Data / Cyber",
    keywords: ["artificial intelligence", "ai", "machine learning", "cyber", "data", "ict", "technology", "data collections"],
    impact: "Could affect governance over AI use cases, data controls, cyber oversight and internal policy requirements.",
    primaryOwner: "Technology",
    secondaryOwner: "Compliance"
  }
];

const FMRUK_KEYWORDS = [
  "uk", "fca", "pra", "investment firm", "asset management", "investment management",
  "mifid", "mifidpru", "consumer duty", "transaction reporting", "market abuse",
  "best execution", "outsourcing", "operational resilience", "financial crime",
  "sanctions", "governance", "remuneration", "sustainability", "disclosure",
  "regulatory reporting", "firm", "firms", "data collections"
];

const els = {
  headerMeta: document.getElementById("headerMeta"),
  fileInput: document.getElementById("fileInput"),
  uploadBtn: document.getElementById("uploadBtn"),
  reloadBtn: document.getElementById("reloadBtn"),
  clearStorageBtn: document.getElementById("clearStorageBtn"),
  uploadStatus: document.getElementById("uploadStatus"),
  datasetInfo: document.getElementById("datasetInfo"),
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

document.addEventListener("DOMContentLoaded", init);

function init() {
  bindEvents();
  loadFromStorage();
  renderAll();
}

function bindEvents() {
  els.uploadBtn.addEventListener("click", handleUpload);
  els.reloadBtn.addEventListener("click", () => renderAll());
  els.clearStorageBtn.addEventListener("click", clearSavedData);
  els.searchInput.addEventListener("input", applyFilters);
  els.sectionFilter.addEventListener("change", applyFilters);
  els.themeFilter.addEventListener("change", applyFilters);
  els.ownerFilter.addEventListener("change", applyFilters);
  els.impactFilter.addEventListener("change", applyFilters);
  els.relevanceFilter.addEventListener("change", applyFilters);
  els.fmrukOnlyFilter.addEventListener("change", applyFilters);
}

function loadFromStorage() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;

  try {
    const payload = JSON.parse(raw);
    state.raw = Array.isArray(payload.items) ? payload.items : [];
    state.datasetMeta = payload.meta || null;
  } catch (err) {
    console.error(err);
  }
}

function saveToStorage() {
  const payload = {
    meta: state.datasetMeta,
    items: state.raw
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function clearSavedData() {
  localStorage.removeItem(STORAGE_KEY);
  state.raw = [];
  state.filtered = [];
  state.selectedItemId = null;
  state.datasetMeta = null;
  els.uploadStatus.textContent = "Saved browser data cleared.";
  renderAll();
}

async function handleUpload() {
  const file = els.fileInput.files[0];
  if (!file) {
    els.uploadStatus.textContent = "Please select an XLSX file first.";
    return;
  }

  els.uploadStatus.textContent = "Reading workbook...";

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });

    const parsedRows = workbook.SheetNames.flatMap(sheetName => {
      const ws = workbook.Sheets[sheetName];
      return XLSX.utils.sheet_to_json(ws, { defval: "" }).map(row => ({
        __sheet: sheetName,
        ...row
      }));
    });

    const normalised = normaliseRows(parsedRows, file.name);
    const analysed = analyseRows(normalised);

    state.raw = analysed;
    state.datasetMeta = {
      fileName: file.name,
      uploadedAt: new Date().toISOString(),
      sheetCount: workbook.SheetNames.length,
      rowCount: analysed.length
    };
    state.selectedItemId = analysed[0]?.id || null;

    saveToStorage();
    els.uploadStatus.textContent = `Loaded ${analysed.length} initiatives from ${file.name}.`;
    renderAll();
  } catch (err) {
    console.error(err);
    els.uploadStatus.textContent = "Upload failed. Check that the file is a valid FCA Grid XLSX.";
  }
}

function normaliseRows(rows, fileName) {
  return rows
    .map((row, index) => {
      const mapped = mapWorkbookRow(row);
      return {
        id: `${slugify(mapped.initiativeTitle || "item")}-${index}`,
        sourceFile: fileName,
        ...mapped
      };
    })
    .filter(item => item.initiativeTitle);
}

function mapWorkbookRow(row) {
  const entries = Object.entries(row).reduce((acc, [k, v]) => {
    acc[String(k).trim().toLowerCase()] = typeof v === "string" ? v.trim() : v;
    return acc;
  }, {});

  const get = (...patterns) => {
    for (const [key, value] of Object.entries(entries)) {
      for (const p of patterns) {
        if (key.includes(p)) return String(value || "").trim();
      }
    }
    return "";
  };

  const sectionName = get("sector", "section") || String(row.__sheet || "").trim();
  const subcategory = get("subcategory", "sub-category", "category");
  const leadRegulator = get("lead");
  const initiativeTitle = get("initiative", "title", "name");
  const initiativeDescription = get("description", "details", "detail", "summary");
  const expectedKeyMilestones = get("expected key milestones", "milestone");
  const indicativeImpactOnFirms = get("impact on firms", "indicative impact", "impact");
  const consumerInterest = get("consumer interest");
  const timingUpdated = get("timing updated", "change in timing");
  const isNew = get("new");
  const rawText = JSON.stringify(row);

  return {
    sectionName,
    subcategory,
    leadRegulator,
    initiativeTitle,
    initiativeDescription,
    expectedKeyMilestones,
    indicativeImpactOnFirms,
    consumerInterest,
    timingUpdated,
    isNew,
    timingBucket: inferTimingBucket(`${expectedKeyMilestones} ${initiativeDescription}`),
    rawText
  };
}

function analyseRows(items) {
  return items
    .map(item => {
      const classification = classifyItem(item);
      const impactLevel = determineImpactLevel(item);
      const relevanceScore = scoreRelevance(item, classification);
      const isFmrukRelevant = determineFmrukRelevance(item, classification);

      return {
        ...item,
        theme: classification.theme,
        internalSubTheme: classification.subTheme,
        potentialBusinessImpact: classification.impact,
        primaryOwner: classification.primaryOwner,
        secondaryOwner: classification.secondaryOwner,
        impactLevel,
        relevanceScore,
        isFmrukRelevant,
        rationale: buildRationale(item, classification),
        suggestedAction: buildSuggestedAction(impactLevel, classification, item.timingBucket)
      };
    })
    .sort((a, b) => (b.relevanceScore || 0) - (a.relevanceScore || 0));
}

function classifyItem(item) {
  const blob = `${item.initiativeTitle} ${item.initiativeDescription} ${item.sectionName} ${item.subcategory}`.toLowerCase();

  for (const rule of THEME_RULES) {
    if (rule.keywords.some(keyword => blob.includes(keyword))) {
      return rule;
    }
  }

  const sectionBlob = `${item.sectionName} ${item.subcategory}`.toLowerCase();

  if (sectionBlob.includes("investment management")) {
    return {
      theme: "Investment Management / Product",
      subTheme: "Investment Management",
      impact: "May require review of product governance, disclosures, oversight or firm implementation planning.",
      primaryOwner: "Compliance",
      secondaryOwner: "Legal"
    };
  }

  if (sectionBlob.includes("wholesale financial markets")) {
    return {
      theme: "Market Structure / Wholesale",
      subTheme: "Wholesale Markets",
      impact: "Could affect market-facing controls, policy interpretation and implementation planning for wholesale business lines.",
      primaryOwner: "Compliance",
      secondaryOwner: "Legal"
    };
  }

  return {
    theme: "General UK Regulatory Change",
    subTheme: "General",
    impact: "May require initial triage to determine whether any policy, control, governance or implementation response is needed.",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal"
  };
}

function determineImpactLevel(item) {
  const impactFlag = String(item.indicativeImpactOnFirms || "").trim().toUpperCase();
  const blob = `${item.initiativeTitle} ${item.initiativeDescription}`.toLowerCase();

  if (impactFlag === "H") return "High";
  if (impactFlag === "L") return "Low";

  const highTerms = [
    "consumer duty", "transaction reporting", "operational resilience", "outsourcing",
    "mifidpru", "capital", "liquidity", "market abuse", "sanctions", "data collections"
  ];
  const mediumTerms = ["consultation", "disclosure", "governance", "cyber", "ict", "sustainability"];

  if (highTerms.some(term => blob.includes(term))) return "High";
  if (mediumTerms.some(term => blob.includes(term))) return "Medium";
  return impactFlag === "U" ? "Medium" : "Low";
}

function scoreRelevance(item, classification) {
  const blob = `${item.initiativeTitle} ${item.initiativeDescription} ${item.sectionName} ${classification.theme} ${classification.subTheme}`.toLowerCase();
  let score = 25;

  for (const keyword of FMRUK_KEYWORDS) {
    if (blob.includes(keyword)) score += 4;
  }

  if (["Investment management", "Multi-sector", "Wholesale financial markets"].includes(item.sectionName)) {
    score += 10;
  }

  if ([
    "Transaction Reporting",
    "Operational Resilience",
    "Outsourcing / Third Party Risk",
    "Capital / Liquidity",
    "Consumer Duty",
    "AML / KYC / Sanctions",
    "AI / Data / Cyber"
  ].includes(classification.subTheme)) {
    score += 12;
  }

  return Math.min(score, 100);
}

function determineFmrukRelevance(item, classification) {
  const blob = `${item.initiativeTitle} ${item.initiativeDescription} ${item.sectionName} ${classification.theme}`.toLowerCase();

  if (["Investment management", "Multi-sector", "Wholesale financial markets"].includes(item.sectionName)) {
    return true;
  }

  return FMRUK_KEYWORDS.some(keyword => blob.includes(keyword));
}

function buildRationale(item, classification) {
  return `Classified from the uploaded Grid under ${item.sectionName || "Unknown section"}. Mapped to ${classification.theme} / ${classification.subTheme}. This is likely relevant to FMRUK where it could affect UK entity policy, controls, governance, reporting or implementation planning. Potential operational implication: ${classification.impact}`;
}

function buildSuggestedAction(impactLevel, classification, timingBucket) {
  if (impactLevel === "High") {
    return `Immediate triage by ${classification.primaryOwner}; engage ${classification.secondaryOwner}; assess whether policy, procedure, control or governance updates are required. Horizon: ${timingBucket}.`;
  }
  if (impactLevel === "Medium") {
    return `Review by ${classification.primaryOwner}; confirm applicability with ${classification.secondaryOwner}; track milestones and assign implementation owner if needed. Horizon: ${timingBucket}.`;
  }
  return `Monitor through ${classification.primaryOwner}; retain on watchlist and reassess if timing or scope changes. Horizon: ${timingBucket}.`;
}

function inferTimingBucket(text) {
  const blob = String(text || "").toLowerCase();
  if (blob.includes("q1 2026") || blob.includes("q2 2026") || blob.includes("january") || blob.includes("april")) {
    return "Near Term";
  }
  if (blob.includes("q3 2026") || blob.includes("q4 2026") || blob.includes("2026")) {
    return "Medium Term";
  }
  if (blob.includes("2027") || blob.includes("post july 2027")) {
    return "Longer Term";
  }
  return "To Be Confirmed";
}

function renderAll() {
  updateMeta();
  populateFilters();
  applyFilters();
}

function updateMeta() {
  if (!state.datasetMeta) {
    els.headerMeta.textContent = "No dataset loaded";
    els.datasetInfo.textContent = "No saved dataset found.";
    return;
  }

  els.headerMeta.textContent = `${state.datasetMeta.fileName} | ${state.datasetMeta.rowCount} initiatives | uploaded ${formatDate(state.datasetMeta.uploadedAt)}`;
  els.datasetInfo.textContent = `Stored in this browser only. Sheets: ${state.datasetMeta.sheetCount}. Source file: ${state.datasetMeta.fileName}`;
}

function populateFilters() {
  fillSelect(els.sectionFilter, unique(state.raw.map(x => x.sectionName)));
  fillSelect(els.themeFilter, unique(state.raw.map(x => x.theme)));
  fillSelect(els.ownerFilter, unique(state.raw.map(x => x.primaryOwner)));
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

function applyFilters() {
  const q = els.searchInput.value.trim().toLowerCase();
  const section = els.sectionFilter.value;
  const theme = els.themeFilter.value;
  const owner = els.ownerFilter.value;
  const impact = els.impactFilter.value;
  const minRelevance = Number(els.relevanceFilter.value || 0);
  const fmrukOnly = els.fmrukOnlyFilter.checked;

  let items = state.raw.filter(item => {
    const haystack = [
      item.sectionName,
      item.subcategory,
      item.initiativeTitle,
      item.initiativeDescription,
      item.theme,
      item.internalSubTheme,
      item.primaryOwner,
      item.secondaryOwner,
      item.potentialBusinessImpact,
      item.rationale
    ].join(" ").toLowerCase();

    return (
      (!q || haystack.includes(q)) &&
      (!section || item.sectionName === section) &&
      (!theme || item.theme === theme) &&
      (!owner || item.primaryOwner === owner) &&
      (!impact || item.impactLevel === impact) &&
      ((item.relevanceScore || 0) >= minRelevance) &&
      (!fmrukOnly || item.isFmrukRelevant === true)
    );
  });

  items.sort((a, b) => (b.relevanceScore || 0) - (a.relevanceScore || 0));
  state.filtered = items;

  if (!items.find(x => x.id === state.selectedItemId)) {
    state.selectedItemId = items[0]?.id || null;
  }

  renderSummary(items);
  renderTable(items);
  renderDetail(items.find(x => x.id === state.selectedItemId));
}

function renderSummary(items) {
  els.kpiTotal.textContent = items.length;
  els.kpiHighRelevance.textContent = items.filter(x => (x.relevanceScore || 0) >= 80).length;
  els.kpiHighImpact.textContent = items.filter(x => x.impactLevel === "High").length;
  els.kpiRelevant.textContent = items.filter(x => x.isFmrukRelevant === true).length;

  renderFrequencyList(els.topThemesList, items.map(x => x.theme));
  renderFrequencyList(els.topOwnersList, items.map(x => x.primaryOwner));
}

function renderFrequencyList(target, values) {
  target.innerHTML = "";
  const counts = {};
  values.filter(Boolean).forEach(v => counts[v] = (counts[v] || 0) + 1);

  const sorted = Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  if (!sorted.length) {
    target.innerHTML = "<li>None</li>";
    return;
  }

  sorted.forEach(([name, count]) => {
    const li = document.createElement("li");
    li.textContent = `${name} (${count})`;
    target.appendChild(li);
  });
}

function renderTable(items) {
  els.tableBody.innerHTML = "";

  if (!items.length) {
    els.tableBody.innerHTML = `<tr><td colspan="6">No initiatives match the current filters.</td></tr>`;
    return;
  }

  items.forEach(item => {
    const tr = document.createElement("tr");
    if (item.id === state.selectedItemId) tr.classList.add("selected");

    tr.innerHTML = `
      <td>${priorityBadge(item.relevanceScore || 0)}</td>
      <td>${escapeHtml(item.sectionName || "")}</td>
      <td>${escapeHtml(item.theme || "")}</td>
      <td>${impactBadge(item.impactLevel || "Low")}</td>
      <td>${escapeHtml(item.primaryOwner || "")}</td>
      <td>${escapeHtml(item.initiativeTitle || "")}</td>
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
    els.detailPanel.innerHTML = "Upload the FCA Grid XLSX, then select an initiative.";
    return;
  }

  els.detailPanel.innerHTML = `
    <h2 class="detail-title">${escapeHtml(item.initiativeTitle || "")}</h2>

    <div class="detail-grid">
      <div class="detail-box">
        <div class="detail-label">Section</div>
        <div class="detail-value">${escapeHtml(item.sectionName || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Subcategory</div>
        <div class="detail-value">${escapeHtml(item.subcategory || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Lead</div>
        <div class="detail-value">${escapeHtml(item.leadRegulator || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Indicative Impact</div>
        <div class="detail-value">${escapeHtml(item.indicativeImpactOnFirms || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Theme</div>
        <div class="detail-value">${escapeHtml(item.theme || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Sub-theme</div>
        <div class="detail-value">${escapeHtml(item.internalSubTheme || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Primary Owner</div>
        <div class="detail-value">${escapeHtml(item.primaryOwner || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Secondary Owner</div>
        <div class="detail-value">${escapeHtml(item.secondaryOwner || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Relevance</div>
        <div class="detail-value">${escapeHtml(String(item.relevanceScore || 0))}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Impact Level</div>
        <div class="detail-value">${escapeHtml(item.impactLevel || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">Timing Bucket</div>
        <div class="detail-value">${escapeHtml(item.timingBucket || "N/A")}</div>
      </div>
      <div class="detail-box">
        <div class="detail-label">New</div>
        <div class="detail-value">${escapeHtml(item.isNew || "N/A")}</div>
      </div>
    </div>

    <div class="detail-block">
      <strong>Potential Business Impact:</strong><br />
      ${escapeHtml(item.potentialBusinessImpact || "N/A")}
    </div>

    <div class="detail-block">
      <strong>Expected Key Milestones:</strong><br />
      ${escapeHtml(item.expectedKeyMilestones || "N/A")}
    </div>

    <div class="detail-block">
      <strong>Initiative Description:</strong><br />
      ${escapeHtml(item.initiativeDescription || "N/A")}
    </div>

    <div class="detail-block">
      <strong>Why it matters to FMRUK:</strong><br />
      ${escapeHtml(item.rationale || "N/A")}
    </div>

    <div class="detail-block">
      <strong>Suggested Action:</strong><br />
      ${escapeHtml(item.suggestedAction || "N/A")}
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

function unique(arr) {
  return [...new Set(arr.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function slugify(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60);
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
