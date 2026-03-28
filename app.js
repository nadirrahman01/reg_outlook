const STORAGE_KEY = "fmruk_reg_grid_v7_pdf_only";

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

const state = {
  raw: [],
  filtered: [],
  selectedItemId: null,
  datasetMeta: null
};

const KNOWN_SECTIONS = [
  "Multi-sector",
  "Banking, credit and lending",
  "Payments and cryptoassets",
  "Payment services and cryptoassets",
  "Insurance and reinsurance",
  "Investment management",
  "Pensions and retirement income",
  "Retail investments",
  "Wholesale financial markets",
  "Annex: initiatives completed/stopped"
];

const KNOWN_SUBCATEGORIES = [
  "Competition, innovation and other",
  "Conduct",
  "Cross-cutting/omnibus",
  "Sustainable finance",
  "Financial resilience",
  "Operational resilience",
  "Other single-sector initiatives",
  "Repeal and replacement of assimilated law under FSMA 2023"
];

const SECTION_LOOKUP = new Map(
  KNOWN_SECTIONS.map(value => [value.toLowerCase(), value])
);

const SUBCATEGORY_LOOKUP = new Map(
  KNOWN_SUBCATEGORIES.map(value => [value.toLowerCase(), value])
);

const HEADER_TEXT_PATTERNS = [
  /regulatory initiatives grid/i,
  /^lead$/i,
  /^initiative$/i,
  /expected key milestones/i,
  /engagement and key milestone/i,
  /indicative impact on firms/i,
  /consumer interest/i,
  /timing updated/i,
  /^new$/i
];

const ROW_NOISE_PATTERNS = [
  /^page \d+ of \d+$/i,
  /^oct-dec \d{4}$/i,
  /^jan-mar \d{4}$/i,
  /^apr-jun \d{4}$/i,
  /^jul-sept \d{4}$/i,
  /^jan-jun \d{4}$/i,
  /^post july \d{4}$/i,
  /^e\s*formal engagement planned/i
];

const REGULATOR_CODES = [
  "FCA",
  "PRA",
  "BOE",
  "HMT",
  "ICO",
  "TPR",
  "FRC",
  "PSR",
  "CMA",
  "FOS",
  "FSCS"
];

const DEFAULT_COLUMN_START_RATIOS = {
  lead: 0.02,
  initiative: 0.12,
  milestones: 0.45,
  impact: 0.77,
  consumer: 0.83,
  timing: 0.89,
  isNew: 0.95
};

const FMRUK_PROFILE = {
  summary:
    "UK entity with strongest exposure to investment management, MiFID / wholesale markets, governance, resilience, reporting and cross-cutting FCA change.",
  coreAreas: [
    "Investment management / funds",
    "MiFID / wholesale markets / market conduct",
    "Operational resilience / outsourcing / third parties",
    "Regulatory reporting / prudential / governance",
    "Financial crime, sanctions, data and sustainability change"
  ]
};

const GENERIC_STAGE_ACTIONS = {
  consultation: [
    "Review the consultation scope, identify impacted business lines and decide whether FMRUK should respond or feed into a group response.",
    "Capture policy interpretation points and implementation risks early so they can shape the eventual delivery plan."
  ],
  final_rules: [
    "Run a rule-to-control gap assessment against the final position and assign a delivery owner with dates.",
    "Confirm whether policy documents, procedures, governance packs, training or MI need to be updated before go-live."
  ],
  supervisory: [
    "Benchmark current controls against the supervisory expectation and gather evidence before any thematic or multi-firm engagement.",
    "Prepare management reporting on current-state compliance, open gaps and remediation ownership."
  ],
  reporting: [
    "Confirm the reporting perimeter, data lineage, control owners and sign-off model for any new or amended submission.",
    "Plan test cycles for data completeness, exceptions and reconciliations before the first live return or attestation."
  ],
  legislation: [
    "Assess whether the legislative or perimeter change affects permissions, legal entity scope or delegated arrangements.",
    "Coordinate legal interpretation, implementation planning and downstream policy updates before rules crystallise."
  ],
  monitoring: [
    "Retain the item on the regulatory watchlist and refresh the assessment when scope, timing or FCA messaging changes.",
    "Keep a light owner assigned so the business is not surprised if the initiative accelerates."
  ]
};

const CLASSIFICATION_RULES = [
  {
    id: "transaction-reporting",
    theme: "Market Conduct / Trading",
    subTheme: "Transaction Reporting",
    primaryOwner: "Compliance",
    secondaryOwner: "Operations",
    impactBias: 28,
    impactStatement:
      "Could affect transaction reporting scope, field mapping, reconciliations, exception handling, ARM oversight and governance.",
    signals: [
      {
        weight: 26,
        reason: "Touches transaction reporting obligations.",
        terms: ["transaction reporting", "rts 22", "order record keeping"]
      },
      {
        weight: 18,
        reason: "References MiFIR wholesale reporting mechanics.",
        terms: ["mifir", "arm", "approved reporting mechanism", "reporting fields"]
      }
    ],
    sectionBoosts: {
      "Wholesale financial markets": 8,
      "Investment management": 3
    },
    coreActions: [
      "Confirm whether the proposal affects in-scope instruments, desks, venues or reporting entities.",
      "Review field mapping, reconciliations, exception MI and any ARM or vendor dependency."
    ]
  },
  {
    id: "best-execution",
    theme: "Market Conduct / Trading",
    subTheme: "Best Execution",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 18,
    impactStatement:
      "Could require updates to execution governance, venue analysis, monitoring, policy wording and oversight of execution arrangements.",
    signals: [
      {
        weight: 26,
        reason: "Directly references best execution.",
        terms: ["best execution", "execution policy", "execution quality"]
      },
      {
        weight: 16,
        reason: "References RTS 28 or execution venue analysis.",
        terms: ["rts 28", "venue analysis"]
      }
    ],
    sectionBoosts: {
      "Wholesale financial markets": 7,
      "Investment management": 4
    },
    coreActions: [
      "Assess whether execution policy wording, venue oversight or client disclosures would need to change.",
      "Check whether execution monitoring MI, committee reporting or broker oversight need to be refreshed."
    ]
  },
  {
    id: "market-abuse",
    theme: "Market Conduct / Trading",
    subTheme: "Market Abuse / Surveillance",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 24,
    impactStatement:
      "May require updates to surveillance controls, market abuse risk assessments, PDMR or personal account dealing frameworks, training and escalation.",
    signals: [
      {
        weight: 24,
        reason: "Touches market abuse or MAR controls.",
        terms: ["market abuse", "mar", "inside information"]
      },
      {
        weight: 18,
        reason: "References surveillance or employee dealing controls.",
        terms: ["surveillance", "personal account dealing", "stors", "suspicious transaction"]
      }
    ],
    sectionBoosts: {
      "Wholesale financial markets": 8,
      "Investment management": 3
    },
    coreActions: [
      "Check surveillance scenarios, alert governance and escalation processes against the proposed change.",
      "Reconfirm training, insider list or employee dealing controls if the change broadens scope."
    ]
  },
  {
    id: "prudential-capital",
    theme: "Prudential / MIFIDPRU",
    subTheme: "Capital / Liquidity",
    primaryOwner: "Finance",
    secondaryOwner: "Risk",
    impactBias: 30,
    impactStatement:
      "May affect prudential assessments, capital planning, ICARA assumptions, liquidity monitoring and board governance.",
    signals: [
      {
        weight: 26,
        reason: "Directly references MIFIDPRU or prudential capital topics.",
        terms: ["mifidpru", "icara", "own funds", "prudential", "concentration risk"]
      },
      {
        weight: 18,
        reason: "Touches capital or liquidity framework requirements.",
        terms: ["capital", "liquidity", "k-factor", "fixed overhead"]
      }
    ],
    sectionBoosts: {
      "Investment management": 8,
      "Multi-sector": 3
    },
    coreActions: [
      "Assess ICARA assumptions, thresholds, governance and monitoring changes that may be required.",
      "Confirm whether prudential reporting, capital modelling or board packs need to be updated."
    ]
  },
  {
    id: "prudential-remuneration",
    theme: "Prudential / MIFIDPRU",
    subTheme: "Remuneration",
    primaryOwner: "HR",
    secondaryOwner: "Compliance",
    impactBias: 18,
    impactStatement:
      "Could require updates to remuneration frameworks, governance, documentation, MRT population, deferral and control testing.",
    signals: [
      {
        weight: 26,
        reason: "Directly references remuneration requirements.",
        terms: ["remuneration", "bonus", "compensation", "malus", "clawback", "mrt"]
      }
    ],
    sectionBoosts: {
      "Investment management": 5,
      "Multi-sector": 3
    },
    coreActions: [
      "Check whether remuneration policy wording, governance approvals or identified staff population need to change.",
      "Assess any impact on pay design, deferral mechanics or attestations."
    ]
  },
  {
    id: "operational-resilience",
    theme: "Operational Resilience",
    subTheme: "Operational Resilience",
    primaryOwner: "Risk",
    secondaryOwner: "Compliance",
    impactBias: 28,
    impactStatement:
      "May require review of resilience mapping, scenario testing, impact tolerances, important business services and governance over remediation.",
    signals: [
      {
        weight: 28,
        reason: "Directly references operational resilience.",
        terms: ["operational resilience", "impact tolerance", "important business service"]
      },
      {
        weight: 16,
        reason: "Touches resilience testing or service mapping.",
        terms: ["mapping", "scenario testing", "resilience"]
      }
    ],
    sectionBoosts: {
      "Multi-sector": 8,
      "Investment management": 4
    },
    coreActions: [
      "Review impacted business services, dependency mapping and scenario testing requirements.",
      "Check whether board reporting, self-assessment content or remediation governance would need updating."
    ]
  },
  {
    id: "outsourcing",
    theme: "Operational Resilience",
    subTheme: "Outsourcing / Third Party Risk",
    primaryOwner: "Risk",
    secondaryOwner: "Technology",
    impactBias: 26,
    impactStatement:
      "Could require updates to outsourcing registers, due diligence, contractual controls, concentration assessments and oversight of service providers.",
    signals: [
      {
        weight: 28,
        reason: "Directly references outsourcing or third-party risk.",
        terms: ["outsourcing", "third party", "third-party", "service provider", "vendor risk"]
      },
      {
        weight: 18,
        reason: "Touches critical third-party or supplier oversight.",
        terms: ["critical third party", "ctp", "supplier oversight"]
      }
    ],
    sectionBoosts: {
      "Multi-sector": 8,
      "Investment management": 4
    },
    coreActions: [
      "Confirm whether any critical or material outsourcing arrangement falls into scope.",
      "Review contracts, registers, due diligence and exit or concentration controls against the change."
    ]
  },
  {
    id: "consumer-duty",
    theme: "Consumer / Conduct",
    subTheme: "Consumer Duty",
    primaryOwner: "Compliance",
    secondaryOwner: "Product",
    impactBias: 20,
    impactStatement:
      "May require review of product governance, distribution oversight, fair value assessments, customer communications and conduct monitoring.",
    signals: [
      {
        weight: 26,
        reason: "Directly references Consumer Duty.",
        terms: ["consumer duty", "good outcomes", "fair value", "vulnerable customers"]
      },
      {
        weight: 14,
        reason: "Touches retail customer outcomes or product governance.",
        terms: ["consumer", "retail customer", "outcomes monitoring"]
      }
    ],
    sectionBoosts: {
      "Retail investments": 6,
      "Multi-sector": 3
    },
    coreActions: [
      "Check whether any FMRUK products, disclosures or distribution arrangements are in scope.",
      "Review fair value, target market, communications and monitoring evidence if retail exposure exists."
    ]
  },
  {
    id: "aml-sanctions",
    theme: "AML / Sanctions / Financial Crime",
    subTheme: "AML / KYC / Sanctions",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 24,
    impactStatement:
      "May require review of AML controls, client onboarding, KYC refresh, sanctions screening, investigations and training.",
    signals: [
      {
        weight: 26,
        reason: "Directly references AML, sanctions or financial crime controls.",
        terms: ["anti-money laundering", "aml", "sanctions", "financial crime", "kyc"]
      },
      {
        weight: 16,
        reason: "Touches fraud, screening or customer due diligence.",
        terms: ["screening", "fraud", "customer due diligence", "cdd", "edd"]
      }
    ],
    sectionBoosts: {
      "Multi-sector": 6,
      "Investment management": 3
    },
    coreActions: [
      "Assess whether onboarding, screening, investigations or monitoring scenarios need to change.",
      "Check whether policy documents, training or legal escalation paths require an update."
    ]
  },
  {
    id: "governance-smcr",
    theme: "Governance / SMCR",
    subTheme: "SMCR / Governance",
    primaryOwner: "Compliance",
    secondaryOwner: "HR",
    impactBias: 22,
    impactStatement:
      "Could affect governance frameworks, committee reporting, responsibilities maps, senior management accountability and certification processes.",
    signals: [
      {
        weight: 26,
        reason: "Directly references SMCR or governance change.",
        terms: ["smcr", "senior manager", "certification regime", "responsibilities map"]
      },
      {
        weight: 16,
        reason: "Touches board or committee governance expectations.",
        terms: ["governance", "board", "committee", "dear ceo"]
      }
    ],
    sectionBoosts: {
      "Multi-sector": 8,
      "Investment management": 3
    },
    coreActions: [
      "Check whether accountabilities, terms of reference, committee packs or certification processes need to change.",
      "Assess whether senior manager statements of responsibility or attestations are affected."
    ]
  },
  {
    id: "sustainability",
    theme: "ESG / Sustainability",
    subTheme: "Sustainability / SDR",
    primaryOwner: "Legal",
    secondaryOwner: "Compliance",
    impactBias: 20,
    impactStatement:
      "May require updates to sustainability disclosures, product communications, naming or labelling controls, governance and evidence of anti-greenwashing review.",
    signals: [
      {
        weight: 28,
        reason: "Directly references sustainability disclosures or anti-greenwashing.",
        terms: ["sustainability", "sdr", "greenwashing", "climate", "transition plan"]
      },
      {
        weight: 16,
        reason: "Touches ESG frameworks or sustainability ratings.",
        terms: ["esg", "tcfd", "tnfd", "esg ratings", "srs"]
      }
    ],
    sectionBoosts: {
      "Investment management": 7,
      "Multi-sector": 4
    },
    coreActions: [
      "Review product disclosures, marketing materials and approval controls for any sustainability claims in scope.",
      "Confirm whether governance, evidence packs or label eligibility assessments need to be refreshed."
    ]
  },
  {
    id: "data-ai-cyber",
    theme: "Data / AI / Technology",
    subTheme: "AI / Data / Cyber",
    primaryOwner: "Technology",
    secondaryOwner: "Compliance",
    impactBias: 20,
    impactStatement:
      "Could affect governance over AI use cases, cyber oversight, technology controls, data quality, inventories and policy requirements.",
    signals: [
      {
        weight: 26,
        reason: "Directly references AI, cyber or technology controls.",
        terms: ["artificial intelligence", "ai", "machine learning", "cyber", "ict"]
      },
      {
        weight: 16,
        reason: "Touches data quality, collections or technology governance.",
        terms: ["data", "data collection", "data collections", "technology", "digital"]
      }
    ],
    sectionBoosts: {
      "Multi-sector": 8,
      "Investment management": 3
    },
    coreActions: [
      "Assess whether internal AI, cyber or data governance frameworks need to change.",
      "Check whether inventories, controls, model governance, cyber testing or technology oversight are affected."
    ]
  },
  {
    id: "funds-investment-management",
    theme: "Investment Management / Product",
    subTheme: "Investment Management",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 20,
    impactStatement:
      "May require review of fund disclosures, product governance, delegated oversight, distribution arrangements or implementation planning for investment products.",
    signals: [
      {
        weight: 22,
        reason: "Directly references investment management or asset management.",
        terms: ["investment management", "asset management", "portfolio management", "fund"]
      },
      {
        weight: 16,
        reason: "Touches UCITS, AIFMD or fund distribution.",
        terms: ["ucits", "aifmd", "fund manager", "authorised fund"]
      }
    ],
    sectionBoosts: {
      "Investment management": 12,
      "Retail investments": 3
    },
    coreActions: [
      "Confirm which products, mandates or fund structures are in scope for FMRUK.",
      "Assess changes to disclosures, delegated oversight, product governance or fund operating model."
    ]
  },
  {
    id: "client-assets",
    theme: "Operations / Client Assets",
    subTheme: "CASS / Custody",
    primaryOwner: "Operations",
    secondaryOwner: "Compliance",
    impactBias: 24,
    impactStatement:
      "Could affect CASS controls, custody oversight, recordkeeping, reconciliations, acknowledgements and governance.",
    signals: [
      {
        weight: 26,
        reason: "Directly references client assets or custody controls.",
        terms: ["cass", "client assets", "custody", "safeguarding"]
      },
      {
        weight: 16,
        reason: "Touches reconciliations or asset recordkeeping.",
        terms: ["reconciliation", "acknowledgement letter", "client money", "asset record"]
      }
    ],
    sectionBoosts: {
      "Investment management": 5,
      "Multi-sector": 3
    },
    coreActions: [
      "Check whether custody model, CASS oversight, reconciliations or recordkeeping controls are impacted.",
      "Assess whether governance forums, attestations or third-party oversight need to be updated."
    ]
  },
  {
    id: "regulatory-reporting",
    theme: "Reporting / Disclosure",
    subTheme: "Regulatory Reporting / Data Collection",
    primaryOwner: "Operations",
    secondaryOwner: "Compliance",
    impactBias: 24,
    impactStatement:
      "May require review of reporting perimeter, data sourcing, control framework, sign-off and regulatory submission readiness.",
    signals: [
      {
        weight: 24,
        reason: "Directly references regulatory reporting or data collections.",
        terms: ["regulatory reporting", "data collection", "data collections", "return", "submission"]
      },
      {
        weight: 18,
        reason: "Touches templates, forms, returns or attestations.",
        terms: ["template", "reporting", "attestation", "notification"]
      }
    ],
    sectionBoosts: {
      "Multi-sector": 6,
      "Investment management": 3
    },
    coreActions: [
      "Confirm ownership for data, reconciliations, controls and submission sign-off.",
      "Assess whether system changes, data sourcing or testing are needed before live reporting."
    ]
  },
  {
    id: "payments-crypto",
    theme: "Payments / Cryptoassets",
    subTheme: "Payments / Crypto",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 10,
    impactStatement:
      "May affect perimeter analysis, client proposition assessment and watchlist monitoring if FMRUK activity expands into payments or crypto-related services.",
    signals: [
      {
        weight: 24,
        reason: "Directly references payments or cryptoassets.",
        terms: ["payment services", "payments", "cryptoasset", "stablecoin", "wallet"]
      }
    ],
    sectionBoosts: {
      "Payments and cryptoassets": 10,
      "Payment services and cryptoassets": 10
    },
    coreActions: [
      "Confirm whether the topic is genuinely in scope for FMRUK or should remain a watchlist item only.",
      "If relevant, assess perimeter, client proposition and outsourced provider impacts before deeper mobilisation."
    ]
  }
];

const SECTION_FALLBACKS = {
  "Investment management": {
    theme: "Investment Management / Product",
    subTheme: "Investment Management",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 16,
    impactStatement:
      "Likely to affect investment management products, disclosures, delegated oversight or implementation planning.",
    coreActions: [
      "Confirm which products or mandates are in scope.",
      "Assess whether disclosures, governance or operating model controls need an update."
    ]
  },
  "Wholesale financial markets": {
    theme: "Market Structure / Wholesale",
    subTheme: "Wholesale Markets",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 16,
    impactStatement:
      "Could affect wholesale market controls, policy interpretation, monitoring or implementation planning.",
    coreActions: [
      "Confirm which activities, desks or controls are in scope.",
      "Assess whether market-facing policies, surveillance or governance need to change."
    ]
  },
  "Multi-sector": {
    theme: "General UK Regulatory Change",
    subTheme: "Cross-cutting",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 14,
    impactStatement:
      "May require entity-wide triage across governance, controls, policy interpretation or implementation planning.",
    coreActions: [
      "Identify impacted control domains and coordinate cross-functional triage.",
      "Confirm whether the change is group-led, entity-led or shared."
    ]
  },
  default: {
    theme: "General UK Regulatory Change",
    subTheme: "General",
    primaryOwner: "Compliance",
    secondaryOwner: "Legal",
    impactBias: 10,
    impactStatement:
      "Requires initial triage to determine whether FMRUK policy, controls, governance or monitoring should change.",
    coreActions: [
      "Validate whether the topic is in scope for FMRUK.",
      "Keep it on the watchlist until scope and timing become clearer."
    ]
  }
};

const STAGE_RULES = [
  {
    id: "final_rules",
    label: "Final Rules / Implementation",
    signals: [
      { weight: 24, terms: ["policy statement", "final rules", "finalised guidance", "comes into force"] },
      { weight: 16, terms: ["implementation", "go-live", "effective date", "rules apply"] }
    ]
  },
  {
    id: "consultation",
    label: "Consultation / Policy Development",
    signals: [
      { weight: 24, terms: ["consultation paper", "consultation", "call for input", "discussion paper"] },
      { weight: 12, terms: ["feedback statement", "respond by", "consulting on"] }
    ]
  },
  {
    id: "supervisory",
    label: "Supervisory / Thematic Review",
    signals: [
      { weight: 24, terms: ["dear ceo", "thematic review", "multi-firm work", "supervisory"] },
      { weight: 14, terms: ["review", "survey", "benchmarking", "supervision"] }
    ]
  },
  {
    id: "reporting",
    label: "Reporting / Data Request",
    signals: [
      { weight: 24, terms: ["data collection", "data collections", "return", "submission", "attestation"] },
      { weight: 14, terms: ["reporting", "notification", "template"] }
    ]
  },
  {
    id: "legislation",
    label: "Legislative / Perimeter Change",
    signals: [
      { weight: 24, terms: ["hm treasury", "legislation", "statutory instrument", "repeal and replacement"] },
      { weight: 14, terms: ["fsma", "perimeter", "authorisation", "authorization", "regime"] }
    ]
  },
  {
    id: "monitoring",
    label: "Monitoring",
    signals: [
      { weight: 10, terms: ["roadmap", "work programme", "strategy", "monitoring"] }
    ]
  }
];

const RELEVANCE_SIGNAL_GROUPS = {
  positive: [
    {
      weight: 22,
      reason: "Touches investment management, funds or portfolio activity.",
      terms: ["investment management", "asset management", "fund", "portfolio management", "ucits", "aifmd"]
    },
    {
      weight: 22,
      reason: "Touches MiFID, prudential or wholesale market obligations.",
      terms: ["mifid", "mifidpru", "investment firm", "wholesale", "market abuse", "best execution", "transaction reporting"]
    },
    {
      weight: 18,
      reason: "Touches cross-cutting governance, resilience or outsourcing controls.",
      terms: ["operational resilience", "outsourcing", "third party", "governance", "smcr", "dear ceo"]
    },
    {
      weight: 16,
      reason: "Touches regulatory reporting, data or submissions that often land on the UK entity.",
      terms: ["regulatory reporting", "reporting", "data collection", "return", "attestation", "notification"]
    },
    {
      weight: 14,
      reason: "Touches financial crime, sanctions or client onboarding controls.",
      terms: ["aml", "anti-money laundering", "sanctions", "kyc", "financial crime"]
    },
    {
      weight: 12,
      reason: "Touches sustainability, disclosures or technology change with cross-cutting UK impact.",
      terms: ["sustainability", "sdr", "greenwashing", "cyber", "ai", "data"]
    }
  ],
  negative: [
    {
      weight: 24,
      reason: "Looks primarily insurance-specific.",
      terms: ["insurance", "reinsurance", "solvency", "policyholder"]
    },
    {
      weight: 22,
      reason: "Looks primarily pensions-specific.",
      terms: ["pensions", "retirement income", "defined benefit", "defined contribution"]
    },
    {
      weight: 18,
      reason: "Looks primarily retail banking or consumer credit-specific.",
      terms: ["mortgage", "consumer credit", "banking", "deposit", "current account"]
    },
    {
      weight: 14,
      reason: "Looks primarily payments or crypto-specific unless FMRUK scope is explicitly stated.",
      terms: ["payment services", "open banking", "cryptoasset", "stablecoin", "e-money"]
    }
  ],
  sectionAdjustments: {
    "Investment management": 18,
    "Wholesale financial markets": 16,
    "Multi-sector": 10,
    "Retail investments": 4,
    "Payments and cryptoassets": -10,
    "Payment services and cryptoassets": -10,
    "Insurance and reinsurance": -18,
    "Pensions and retirement income": -18,
    "Banking, credit and lending": -12,
    "Annex: initiatives completed/stopped": -20
  }
};

const els = {
  headerMeta: document.getElementById("headerMeta"),
  fileInput: document.getElementById("fileInput"),
  uploadBtn: document.getElementById("uploadBtn"),
  reloadBtn: document.getElementById("reloadBtn"),
  exportJsonBtn: document.getElementById("exportJsonBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  clearStorageBtn: document.getElementById("clearStorageBtn"),
  uploadStatus: document.getElementById("uploadStatus"),
  datasetInfo: document.getElementById("datasetInfo"),
  comparisonInfo: document.getElementById("comparisonInfo"),
  searchInput: document.getElementById("searchInput"),
  sectionFilter: document.getElementById("sectionFilter"),
  themeFilter: document.getElementById("themeFilter"),
  ownerFilter: document.getElementById("ownerFilter"),
  stageFilter: document.getElementById("stageFilter"),
  changeFilter: document.getElementById("changeFilter"),
  impactFilter: document.getElementById("impactFilter"),
  relevanceFilter: document.getElementById("relevanceFilter"),
  parseConfidenceFilter: document.getElementById("parseConfidenceFilter"),
  fmrukOnlyFilter: document.getElementById("fmrukOnlyFilter"),
  excludeAnnexFilter: document.getElementById("excludeAnnexFilter"),
  needsReviewFilter: document.getElementById("needsReviewFilter"),
  tableBody: document.getElementById("tableBody"),
  detailPanel: document.getElementById("detailPanel"),
  kpiTotal: document.getElementById("kpiTotal"),
  kpiImmediateAction: document.getElementById("kpiImmediateAction"),
  kpiHighRelevance: document.getElementById("kpiHighRelevance"),
  kpiHighImpact: document.getElementById("kpiHighImpact"),
  kpiNeedsReview: document.getElementById("kpiNeedsReview"),
  kpiChanges: document.getElementById("kpiChanges"),
  topThemesList: document.getElementById("topThemesList"),
  topOwnersList: document.getElementById("topOwnersList"),
  immediateActionsList: document.getElementById("immediateActionsList"),
  reviewList: document.getElementById("reviewList"),
  changesList: document.getElementById("changesList")
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
  els.exportJsonBtn.addEventListener("click", exportJson);
  els.exportCsvBtn.addEventListener("click", exportCsv);
  els.clearStorageBtn.addEventListener("click", clearSavedData);
  els.searchInput.addEventListener("input", applyFilters);
  els.sectionFilter.addEventListener("change", applyFilters);
  els.themeFilter.addEventListener("change", applyFilters);
  els.ownerFilter.addEventListener("change", applyFilters);
  els.stageFilter.addEventListener("change", applyFilters);
  els.changeFilter.addEventListener("change", applyFilters);
  els.impactFilter.addEventListener("change", applyFilters);
  els.relevanceFilter.addEventListener("change", applyFilters);
  els.parseConfidenceFilter.addEventListener("change", applyFilters);
  els.fmrukOnlyFilter.addEventListener("change", applyFilters);
  els.excludeAnnexFilter.addEventListener("change", applyFilters);
  els.needsReviewFilter.addEventListener("change", applyFilters);
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
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      meta: state.datasetMeta,
      items: state.raw
    })
  );
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
    els.uploadStatus.textContent = "Please select the FCA initiatives PDF first.";
    return;
  }

  const ext = file.name.split(".").pop().toLowerCase();
  if (ext !== "pdf") {
    els.uploadStatus.textContent = "PDF-only mode is enabled for now.";
    return;
  }

  if (!window.pdfjsLib) {
    els.uploadStatus.textContent = "PDF.js failed to load.";
    return;
  }

  const previousItems = state.raw.slice();
  const previousMeta = state.datasetMeta;

  try {
    els.uploadStatus.textContent = `Reading PDF: ${file.name}...`;
    const buffer = await file.arrayBuffer();
    const parsed = await parsePdfFile(buffer, file.name);
    const deduped = dedupeItems(parsed);

    if (!deduped.length) {
      els.uploadStatus.textContent =
        "No initiatives were detected in the PDF. Review the source formatting or parsing assumptions.";
      return;
    }

    els.uploadStatus.textContent = `Analysing ${deduped.length} extracted initiatives...`;
    const analysed = analyseRows(deduped);
    const comparison = compareWithPreviousDataset(previousItems, analysed);

    state.raw = comparison.items;
    state.datasetMeta = {
      fileName: file.name,
      uploadedAt: new Date().toISOString(),
      rowCount: comparison.items.length,
      fileType: "PDF",
      previousFileName: previousMeta?.fileName || "",
      comparisonSummary: comparison.summary,
      parserVersion: "v7"
    };
    state.selectedItemId = comparison.items[0]?.id || null;

    saveToStorage();
    els.uploadStatus.textContent = `Loaded ${comparison.items.length} initiatives from ${file.name}.`;
    renderAll();
  } catch (err) {
    console.error(err);
    els.uploadStatus.textContent = `Upload failed. ${err.message || "The PDF could not be parsed."}`;
  }
}

async function parsePdfFile(buffer, fileName) {
  const loadingTask = pdfjsLib.getDocument({ data: buffer });
  const pdf = await loadingTask.promise;
  const rows = [];
  let activeTemplate = null;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent();
    const pageRows = buildPdfRows(textContent.items, pageNumber);
    const headerTemplate = detectColumnTemplate(pageRows, viewport.width);
    activeTemplate = headerTemplate || activeTemplate || buildFallbackColumnTemplate(viewport.width);

    pageRows.forEach(row => {
      row.cells = assignRowCells(row, activeTemplate);
      row.pageWidth = viewport.width;
      rows.push(row);
    });
  }

  return extractPdfInitiatives(rows, fileName);
}

function buildPdfRows(items, pageNumber) {
  const tokens = items
    .map(item => ({
      text: normaliseWs(item.str),
      x: Number(item.transform?.[4] || 0),
      y: Number(item.transform?.[5] || 0),
      width: Number(item.width || 0),
      fontName: String(item.fontName || ""),
      pageNumber
    }))
    .filter(token => token.text);

  tokens.sort((a, b) => b.y - a.y || a.x - b.x);

  const rawRows = [];
  for (const token of tokens) {
    const current = rawRows[rawRows.length - 1];
    if (current && Math.abs(current.y - token.y) <= 3) {
      current.tokens.push(token);
      current.y = (current.y + token.y) / 2;
    } else {
      rawRows.push({
        y: token.y,
        tokens: [token],
        pageNumber
      });
    }
  }

  return rawRows.map(row => normaliseRow(row));
}

function normaliseRow(row) {
  const segments = mergeRowTokens(row.tokens.sort((a, b) => a.x - b.x));
  const text = normaliseWs(segments.map(segment => segment.text).join(" "));

  return {
    pageNumber: row.pageNumber,
    y: row.y,
    segments,
    text,
    hasBold: segments.some(segment => /bold/i.test(segment.fontName || "")),
    cells: {}
  };
}

function mergeRowTokens(tokens) {
  const segments = [];

  for (const token of tokens) {
    const previous = segments[segments.length - 1];
    if (!previous) {
      segments.push({ ...token });
      continue;
    }

    const previousEnd = previous.x + previous.width;
    const gap = token.x - previousEnd;

    if (gap <= 6) {
      previous.text = joinInline(previous.text, token.text);
      previous.width = Math.max(previous.width, token.x + token.width - previous.x);
      previous.fontName = previous.fontName || token.fontName;
    } else {
      segments.push({ ...token });
    }
  }

  return segments.map(segment => ({
    ...segment,
    text: normaliseWs(segment.text)
  }));
}

function joinInline(left, right) {
  if (!left) return right;
  if (!right) return left;
  if (/[-/()]$/.test(left) || /^[,.;:)]/.test(right)) {
    return `${left}${right}`;
  }
  return `${left} ${right}`;
}

function detectColumnTemplate(rows, pageWidth) {
  for (const row of rows) {
    const anchorMap = {};
    for (const segment of row.segments) {
      const value = segment.text.toLowerCase();
      if (!anchorMap.lead && /\blead\b/.test(value)) anchorMap.lead = segment.x;
      if (!anchorMap.initiative && /\binitiative\b/.test(value)) anchorMap.initiative = segment.x;
      if (!anchorMap.milestones && /(expected key|milestone)/.test(value)) anchorMap.milestones = segment.x;
      if (!anchorMap.impact && /\bimpact\b/.test(value)) anchorMap.impact = segment.x;
      if (!anchorMap.consumer && /\bconsumer\b/.test(value)) anchorMap.consumer = segment.x;
      if (!anchorMap.timing && /\btiming\b/.test(value)) anchorMap.timing = segment.x;
      if (!anchorMap.isNew && /^\s*new\s*$/.test(value)) anchorMap.isNew = segment.x;
    }

    const hits = Object.keys(anchorMap).length;
    if (hits >= 4 && anchorMap.initiative != null && anchorMap.milestones != null) {
      return buildColumnTemplate(anchorMap, pageWidth);
    }
  }

  return null;
}

function buildFallbackColumnTemplate(pageWidth) {
  const anchors = {};
  for (const [key, ratio] of Object.entries(DEFAULT_COLUMN_START_RATIOS)) {
    anchors[key] = pageWidth * ratio;
  }
  return buildColumnTemplate(anchors, pageWidth);
}

function buildColumnTemplate(anchorMap, pageWidth) {
  const keys = ["lead", "initiative", "milestones", "impact", "consumer", "timing", "isNew"];
  const starts = keys.map(key => ({
    key,
    start:
      anchorMap[key] != null
        ? anchorMap[key]
        : pageWidth * DEFAULT_COLUMN_START_RATIOS[key]
  }));

  const columns = [];
  for (let index = 0; index < starts.length; index += 1) {
    const current = starts[index];
    const previous = starts[index - 1];
    const next = starts[index + 1];
    columns.push({
      key: current.key,
      start: previous ? (previous.start + current.start) / 2 : -Infinity,
      end: next ? (current.start + next.start) / 2 : Infinity
    });
  }

  return columns;
}

function assignRowCells(row, template) {
  const cells = {
    lead: "",
    initiative: "",
    milestones: "",
    impact: "",
    consumer: "",
    timing: "",
    isNew: "",
    overflow: ""
  };

  for (const segment of row.segments) {
    const centre = segment.x + segment.width / 2;
    const column = template.find(item => centre >= item.start && centre < item.end);
    const key = column?.key || "overflow";
    cells[key] = normaliseWs(`${cells[key]} ${segment.text}`);
  }

  return cells;
}

function extractPdfInitiatives(rows, fileName) {
  const initiatives = [];
  let currentSection = "";
  let currentSubcategory = "";
  let current = null;

  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    const text = normaliseWs(row.text);
    if (!text || isNoiseRow(row)) continue;

    const knownSection = SECTION_LOOKUP.get(text.toLowerCase());
    if (knownSection) {
      currentSection = knownSection;
      currentSubcategory = "";
      continue;
    }

    const knownSubcategory = SUBCATEGORY_LOOKUP.get(text.toLowerCase());
    if (knownSubcategory) {
      currentSubcategory = knownSubcategory;
      continue;
    }

    const nextRow = rows[index + 1];
    const start = detectInitiativeStart(row, nextRow);

    if (start) {
      if (current?.initiativeTitle) {
        initiatives.push(finalisePdfInitiative(current, initiatives.length, fileName));
      }

      current = createInitiativeAccumulator(start, currentSection, currentSubcategory, row, fileName);
      ingestRowIntoAccumulator(current, row, { isStartRow: true });

      if (start.consumeNextRow && nextRow) {
        ingestRowIntoAccumulator(current, nextRow, {
          suppressInitiativeCell: true
        });
        index += 1;
      }

      continue;
    }

    if (current) {
      ingestRowIntoAccumulator(current, row, {});
    }
  }

  if (current?.initiativeTitle) {
    initiatives.push(finalisePdfInitiative(current, initiatives.length, fileName));
  }

  return initiatives.filter(item => item.initiativeTitle && item.initiativeTitle.length > 4);
}

function isNoiseRow(row) {
  const text = row.text;
  if (!text) return true;

  if (ROW_NOISE_PATTERNS.some(pattern => pattern.test(text))) return true;

  const headerHits = HEADER_TEXT_PATTERNS.reduce((count, pattern) => {
    return count + (pattern.test(text) ? 1 : 0);
  }, 0);

  return headerHits >= 3;
}

function detectInitiativeStart(row, nextRow) {
  const direct = extractLeadAndTitle(row);
  if (direct) return direct;

  const leadOnly = normaliseLeadToken(row.cells.lead || row.text);
  if (
    leadOnly &&
    nextRow &&
    !isNoiseRow(nextRow) &&
    !SECTION_LOOKUP.has(nextRow.text.toLowerCase()) &&
    !SUBCATEGORY_LOOKUP.has(nextRow.text.toLowerCase())
  ) {
    const nextTitle = cleanTitleCandidate(nextRow.cells.initiative || nextRow.text);
    if (isPlausibleTitle(nextTitle) && !extractLeadAndTitle(nextRow)) {
      return {
        lead: leadOnly,
        title: nextTitle,
        consumeNextRow: true
      };
    }
  }

  return null;
}

function extractLeadAndTitle(row) {
  const directLead = normaliseLeadToken(row.cells.lead);
  const directTitle = cleanTitleCandidate(row.cells.initiative);

  if (directLead && isPlausibleTitle(directTitle)) {
    return {
      lead: directLead,
      title: directTitle,
      consumeNextRow: false
    };
  }

  const split = splitLeadAndTitle(row.text);
  if (split && isPlausibleTitle(split.title)) {
    return {
      lead: split.lead,
      title: split.title,
      consumeNextRow: false
    };
  }

  if (row.cells.initiative) {
    const embedded = splitLeadAndTitle(row.cells.initiative);
    if (embedded && isPlausibleTitle(embedded.title)) {
      return {
        lead: embedded.lead,
        title: embedded.title,
        consumeNextRow: false
      };
    }
  }

  return null;
}

function createInitiativeAccumulator(start, sectionName, subcategory, row, fileName) {
  return {
    sourceFile: fileName,
    sectionName: sectionName || "",
    subcategory: subcategory || "",
    leadRegulator: start.lead || "",
    initiativeTitle: start.title || "",
    initiativeDescriptionParts: [],
    milestoneParts: [],
    impactParts: [],
    consumerParts: [],
    timingParts: [],
    isNewParts: [],
    rawParts: [row.text],
    pageNumbers: [row.pageNumber],
    parseWarnings: []
  };
}

function ingestRowIntoAccumulator(accumulator, row, options) {
  const opts = options || {};

  appendUnique(accumulator.rawParts, row.text);
  appendUnique(accumulator.pageNumbers, row.pageNumber);

  appendUnique(accumulator.milestoneParts, row.cells.milestones);
  appendUnique(accumulator.impactParts, row.cells.impact);
  appendUnique(accumulator.consumerParts, row.cells.consumer);
  appendUnique(accumulator.timingParts, row.cells.timing);
  appendUnique(accumulator.isNewParts, row.cells.isNew);

  if (opts.isStartRow) {
    const extraLeadText = stripLeadPrefix(row.cells.lead, accumulator.leadRegulator);
    appendUnique(accumulator.initiativeDescriptionParts, extraLeadText);

    const extraTitleText = stripRepeatedTitle(row.cells.initiative, accumulator.initiativeTitle);
    if (extraTitleText && !shouldTreatAsTitleContinuation(accumulator, extraTitleText, row)) {
      appendUnique(accumulator.initiativeDescriptionParts, extraTitleText);
    }
    return;
  }

  if (!opts.suppressInitiativeCell) {
    const continuation = stripRepeatedTitle(row.cells.initiative, accumulator.initiativeTitle);
    if (continuation) {
      if (shouldTreatAsTitleContinuation(accumulator, continuation, row)) {
        accumulator.initiativeTitle = normaliseWs(`${accumulator.initiativeTitle} ${continuation}`);
      } else {
        appendUnique(accumulator.initiativeDescriptionParts, continuation);
      }
    }
  }

  appendUnique(accumulator.initiativeDescriptionParts, row.cells.overflow);
}

function shouldTreatAsTitleContinuation(accumulator, value, row) {
  if (!value) return false;
  if (accumulator.initiativeTitle.length > 110) return false;
  if (row.cells.milestones || row.cells.impact || row.cells.consumer) return false;
  if (/[.:;]/.test(value)) return false;
  if (looksLikeMilestoneText(value)) return false;
  return value.length <= 90;
}

function finalisePdfInitiative(accumulator, index, fileName) {
  const raw = normaliseWs(accumulator.rawParts.join(" "));
  const milestones = dedupeStrings(
    [
      ...accumulator.milestoneParts,
      extractMilestones(raw)
    ].filter(Boolean)
  ).join(" | ");

  const description = dedupeStrings(accumulator.initiativeDescriptionParts).join(" ");
  const impactFlag = extractImpactFlag(accumulator.impactParts.join(" ") || raw);
  const consumerInterest = normaliseTrafficLight(accumulator.consumerParts.join(" "));
  const timingUpdated = normaliseYesNo(accumulator.timingParts.join(" "));
  const isNew = normaliseYesNo(accumulator.isNewParts.join(" ")) || (/\bnew\b/i.test(raw) ? "Yes" : "No");
  const parseAssessment = assessParseQuality(accumulator, {
    milestones,
    description,
    impactFlag
  });

  return {
    id: `${slugify(accumulator.initiativeTitle || "item")}-${index}`,
    canonicalKey: buildCanonicalKey(accumulator),
    sourceFile: fileName,
    sourcePages: accumulator.pageNumbers.slice().sort((a, b) => a - b),
    sectionName: accumulator.sectionName || "",
    subcategory: accumulator.subcategory || "",
    leadRegulator: accumulator.leadRegulator || "",
    initiativeTitle: accumulator.initiativeTitle || "",
    initiativeDescription: description || raw,
    expectedKeyMilestones: milestones,
    indicativeImpactOnFirms: impactFlag,
    consumerInterest,
    timingUpdated: timingUpdated || "",
    isNew,
    timingBucket: inferTimingBucket(`${milestones} ${description} ${raw}`),
    rawText: raw,
    parseConfidence: parseAssessment.score,
    parseConfidenceBand: parseAssessment.band,
    parseWarnings: parseAssessment.warnings
  };
}

function assessParseQuality(accumulator, context) {
  let score = 42;
  const warnings = [];

  if (accumulator.sectionName) score += 10;
  else warnings.push("Section was not cleanly extracted from the PDF.");

  if (accumulator.subcategory) score += 4;
  if (isLeadToken(accumulator.leadRegulator)) score += 12;
  else warnings.push("Lead regulator was not cleanly isolated.");

  if (accumulator.initiativeTitle && accumulator.initiativeTitle.length >= 12) score += 16;
  else warnings.push("Initiative title may be incomplete.");

  if (context.description && context.description.length >= 40) score += 10;
  else warnings.push("Description is thin and may need manual validation.");

  if (context.milestones) score += 8;
  else warnings.push("Milestones were not clearly separated from the PDF.");

  if (context.impactFlag) score += 6;
  if (accumulator.pageNumbers.length) score += 4;
  if (accumulator.rawParts.length > 1) score += 2;

  score = clamp(score, 0, 100);

  return {
    score,
    band: score >= 85 ? "High" : score >= 70 ? "Medium" : "Low",
    warnings: dedupeStrings(warnings)
  };
}

function dedupeItems(items) {
  const seen = new Set();
  const output = [];

  for (const item of items) {
    const key = item.canonicalKey || buildCanonicalKey(item);
    if (seen.has(key)) continue;
    seen.add(key);
    output.push(item);
  }

  return output;
}

function analyseRows(items) {
  return items
    .map(item => {
      const classification = classifyItem(item);
      const stage = determineStage(item);
      const impact = determineImpactLevel(item, classification, stage);
      const relevance = evaluateFmrukRelevance(item, classification, stage);
      const potentialBusinessImpact = buildPotentialBusinessImpact(classification, stage, item);
      const suggestedActions = buildSuggestedActions(item, classification, stage, impact, relevance);
      const review = determineReviewStatus(item, classification, stage, relevance);
      const rationale = buildRationale(item, classification, stage, relevance);

      return {
        ...item,
        theme: classification.theme,
        internalSubTheme: classification.subTheme,
        classificationConfidence: classification.confidence,
        classificationSignals: classification.signals,
        classificationAmbiguity: classification.ambiguity,
        potentialBusinessImpact,
        primaryOwner: classification.primaryOwner,
        secondaryOwner: classification.secondaryOwner,
        stage: stage.id,
        stageLabel: stage.label,
        stageConfidence: stage.confidence,
        stageSignals: stage.signals,
        impactLevel: impact.level,
        impactScore: impact.score,
        relevanceScore: relevance.score,
        relevanceBand: relevance.band,
        relevanceConfidence: relevance.confidence,
        relevanceSignals: relevance.positiveReasons,
        relevanceNegativeSignals: relevance.negativeReasons,
        isFmrukRelevant: relevance.isRelevant,
        rationale,
        suggestedActions,
        suggestedAction: suggestedActions.join(" "),
        immediateActionRequired: relevance.isRelevant && impact.level === "High" && stage.id !== "monitoring",
        needsReview: review.needsReview,
        reviewReasons: review.reasons
      };
    })
    .sort((a, b) => {
      if ((b.immediateActionRequired ? 1 : 0) !== (a.immediateActionRequired ? 1 : 0)) {
        return (b.immediateActionRequired ? 1 : 0) - (a.immediateActionRequired ? 1 : 0);
      }
      if ((b.relevanceScore || 0) !== (a.relevanceScore || 0)) {
        return (b.relevanceScore || 0) - (a.relevanceScore || 0);
      }
      return (b.impactScore || 0) - (a.impactScore || 0);
    });
}

function classifyItem(item) {
  const blob = buildBlob(item);
  const scored = CLASSIFICATION_RULES.map(rule => {
    const result = scoreSignalGroups(rule.signals, blob);
    const boost = rule.sectionBoosts?.[item.sectionName] || 0;
    return {
      rule,
      score: result.score + boost,
      signals: result.reasons
    };
  }).sort((a, b) => b.score - a.score);

  const best = scored[0];
  const runnerUp = scored[1];

  if (best && best.score >= 18) {
    return {
      theme: best.rule.theme,
      subTheme: best.rule.subTheme,
      primaryOwner: best.rule.primaryOwner,
      secondaryOwner: best.rule.secondaryOwner,
      impactBias: best.rule.impactBias,
      impactStatement: best.rule.impactStatement,
      coreActions: best.rule.coreActions,
      confidence: clamp(45 + best.score, 0, 100),
      signals: best.signals,
      ambiguity:
        runnerUp && best.score - runnerUp.score <= 6
          ? `Close call against ${runnerUp.rule.theme} / ${runnerUp.rule.subTheme}.`
          : ""
    };
  }

  const fallback = SECTION_FALLBACKS[item.sectionName] || SECTION_FALLBACKS.default;
  return {
    ...fallback,
    confidence: 52,
    signals: [`Fallback from section: ${item.sectionName || "Unknown section"}.`],
    ambiguity: ""
  };
}

function determineStage(item) {
  const blob = `${buildBlob(item)} ${item.expectedKeyMilestones}`.toLowerCase();
  const scored = STAGE_RULES.map(rule => {
    const result = scoreSignalGroups(rule.signals, blob);
    return {
      rule,
      score: result.score,
      reasons: result.reasons
    };
  }).sort((a, b) => b.score - a.score);

  const best = scored[0];
  if (best && best.score >= 12) {
    return {
      id: best.rule.id,
      label: best.rule.label,
      confidence: clamp(45 + best.score, 0, 100),
      signals: best.reasons
    };
  }

  return {
    id: "monitoring",
    label: "Monitoring",
    confidence: 50,
    signals: ["No strong stage markers were detected, so the item is being treated as watchlist monitoring."]
  };
}

function determineImpactLevel(item, classification, stage) {
  let score = 30;
  const impactFlag = String(item.indicativeImpactOnFirms || "").trim().toUpperCase();
  if (impactFlag === "H" || impactFlag === "HIGH") score += 40;
  if (impactFlag === "M" || impactFlag === "MEDIUM" || impactFlag === "U") score += 22;
  if (impactFlag === "L" || impactFlag === "LOW") score += 8;

  score += classification.impactBias || 0;

  if (stage.id === "final_rules") score += 12;
  if (stage.id === "reporting") score += 10;
  if (stage.id === "supervisory") score += 8;
  if (item.timingBucket === "Near Term") score += 8;
  if (item.timingBucket === "Longer Term") score -= 4;
  if (item.consumerInterest === "H") score += 4;

  score = clamp(score, 0, 100);

  return {
    score,
    level: score >= 78 ? "High" : score >= 55 ? "Medium" : "Low"
  };
}

function evaluateFmrukRelevance(item, classification, stage) {
  const blob = buildBlob(item);
  let score = 38;
  const positive = scoreSignalGroups(RELEVANCE_SIGNAL_GROUPS.positive, blob);
  const negative = scoreSignalGroups(RELEVANCE_SIGNAL_GROUPS.negative, blob);

  score += positive.score;
  score -= negative.score;
  score += RELEVANCE_SIGNAL_GROUPS.sectionAdjustments[item.sectionName] || 0;

  if (
    ["Transaction Reporting", "Best Execution", "Market Abuse / Surveillance", "Capital / Liquidity", "Operational Resilience", "Outsourcing / Third Party Risk", "SMCR / Governance", "Regulatory Reporting / Data Collection"].includes(
      classification.subTheme
    )
  ) {
    score += 10;
  }

  if (stage.id === "reporting" || stage.id === "final_rules") score += 4;
  if (item.leadRegulator.includes("FCA")) score += 3;
  if (item.sectionName === "Annex: initiatives completed/stopped") score -= 10;

  score = clamp(score, 0, 100);

  const positiveReasons = [...positive.reasons];
  const negativeReasons = [...negative.reasons];
  const sectionAdj = RELEVANCE_SIGNAL_GROUPS.sectionAdjustments[item.sectionName] || 0;
  if (sectionAdj > 0) positiveReasons.push(`Section boost: ${item.sectionName}.`);
  if (sectionAdj < 0) negativeReasons.push(`Section drag: ${item.sectionName}.`);

  const mixedSignals = positiveReasons.length > 0 && negativeReasons.length > 0;
  const isRelevant = score >= 58 || (score >= 50 && !negativeReasons.length);
  const band = score >= 80 ? "High" : score >= 58 ? "Medium" : "Low";
  const confidence = clamp(55 + Math.abs(score - 58), 0, 100);

  return {
    score,
    band,
    confidence,
    positiveReasons,
    negativeReasons,
    mixedSignals,
    isRelevant
  };
}

function determineReviewStatus(item, classification, stage, relevance) {
  const reasons = [];

  if (item.parseConfidence < 70) reasons.push("PDF extraction confidence is low.");
  if (classification.confidence < 60) reasons.push("Classification confidence is modest.");
  if (classification.ambiguity) reasons.push(classification.ambiguity);
  if (stage.confidence < 58) reasons.push("Initiative stage is uncertain.");
  if (relevance.mixedSignals) reasons.push("Relevance signals are mixed across FMRUK and non-FMRUK scope.");
  if (relevance.score >= 45 && relevance.score <= 65) reasons.push("Relevance score sits in the judgement zone and should be checked manually.");
  if (!item.expectedKeyMilestones) reasons.push("Milestones were not clearly extracted.");

  return {
    needsReview: reasons.length > 0,
    reasons: dedupeStrings(reasons)
  };
}

function buildPotentialBusinessImpact(classification, stage, item) {
  const timingNote =
    item.timingBucket && item.timingBucket !== "To Be Confirmed"
      ? ` Current horizon is ${item.timingBucket.toLowerCase()}.`
      : "";

  const stageNote = {
    consultation:
      " The main near-term need is understanding scope, shaping response and avoiding late delivery surprises.",
    final_rules:
      " Because this appears close to implementation, the focus should move quickly from interpretation into change delivery and control readiness.",
    supervisory:
      " The main exposure is evidencing current-state compliance and fixing weak controls before supervisory scrutiny increases.",
    reporting:
      " The practical risk sits in data quality, ownership, reconciliation and readiness to make accurate submissions.",
    legislation:
      " The key question is whether the perimeter or legal interpretation changes the entity obligations that sit underneath implementation.",
    monitoring:
      " This currently looks more like a watchlist item than an immediate delivery programme."
  }[stage.id] || "";

  return `${classification.impactStatement}${stageNote}${timingNote}`;
}

function buildSuggestedActions(item, classification, stage, impact, relevance) {
  const actions = [];

  if (relevance.isRelevant) {
    actions.push(
      `${classification.primaryOwner} should own initial triage with ${classification.secondaryOwner} support.`
    );
  } else {
    actions.push(
      `Keep this with ${classification.primaryOwner} as a watchlist item unless FMRUK scope becomes clearer.`
    );
  }

  const stageActions = GENERIC_STAGE_ACTIONS[stage.id] || GENERIC_STAGE_ACTIONS.monitoring;
  actions.push(stageActions[0]);
  actions.push(classification.coreActions[0]);

  if (impact.level === "High" || stage.id === "final_rules" || stage.id === "reporting") {
    actions.push(classification.coreActions[1]);
  }

  if (item.expectedKeyMilestones) {
    actions.push(`Track milestone text captured from the PDF: ${truncateText(item.expectedKeyMilestones, 150)}.`);
  }

  if (item.parseConfidence < 70 || relevance.mixedSignals) {
    actions.push("Validate the extracted row against the source PDF before relying on the detail for delivery planning.");
  }

  return dedupeStrings(actions).slice(0, 5);
}

function buildRationale(item, classification, stage, relevance) {
  const positives = relevance.positiveReasons.length
    ? `Positive relevance signals: ${relevance.positiveReasons.join(" ")}`
    : "No strong positive relevance signals were detected.";
  const negatives = relevance.negativeReasons.length
    ? ` Counter-signals: ${relevance.negativeReasons.join(" ")}`
    : "";

  return `Mapped to ${classification.theme} / ${classification.subTheme}. Treated as ${stage.label}. ${positives}${negatives} FMRUK profile assumption: ${FMRUK_PROFILE.summary}`;
}

function compareWithPreviousDataset(previousItems, newItems) {
  const previousMap = new Map(previousItems.map(item => [item.canonicalKey, item]));
  const nextMap = new Map(newItems.map(item => [item.canonicalKey, item]));
  let newCount = 0;
  let changedCount = 0;
  let existingCount = 0;

  const items = newItems.map(item => {
    const previous = previousMap.get(item.canonicalKey);
    let changeStatus = "Existing";

    if (!previous) {
      changeStatus = "New";
      newCount += 1;
    } else if (hasMeaningfulChanges(previous, item)) {
      changeStatus = "Changed";
      changedCount += 1;
    } else {
      existingCount += 1;
    }

    return {
      ...item,
      changeStatus
    };
  });

  let removedCount = 0;
  for (const previousKey of previousMap.keys()) {
    if (!nextMap.has(previousKey)) removedCount += 1;
  }

  return {
    items,
    summary: {
      newCount,
      changedCount,
      existingCount,
      removedCount
    }
  };
}

function hasMeaningfulChanges(previous, next) {
  const fields = [
    "initiativeDescription",
    "expectedKeyMilestones",
    "indicativeImpactOnFirms",
    "theme",
    "internalSubTheme",
    "stage",
    "impactLevel",
    "relevanceBand"
  ];

  return fields.some(field => normaliseWs(previous[field]) !== normaliseWs(next[field]));
}

function exportJson() {
  if (!state.filtered.length) {
    els.uploadStatus.textContent = "There is no filtered dataset to export.";
    return;
  }

  downloadTextFile(
    "fmruk-regulatory-dashboard.json",
    JSON.stringify(state.filtered, null, 2),
    "application/json"
  );
}

function exportCsv() {
  if (!state.filtered.length) {
    els.uploadStatus.textContent = "There is no filtered dataset to export.";
    return;
  }

  const csv = buildCsv(state.filtered);
  downloadTextFile("fmruk-regulatory-dashboard.csv", csv, "text/csv;charset=utf-8;");
}

function buildCsv(items) {
  const columns = [
    "changeStatus",
    "sectionName",
    "subcategory",
    "leadRegulator",
    "initiativeTitle",
    "theme",
    "internalSubTheme",
    "primaryOwner",
    "secondaryOwner",
    "stageLabel",
    "impactLevel",
    "relevanceScore",
    "parseConfidence",
    "timingBucket",
    "expectedKeyMilestones",
    "initiativeDescription",
    "potentialBusinessImpact",
    "rationale",
    "suggestedAction"
  ];

  const lines = [
    columns.join(","),
    ...items.map(item =>
      columns
        .map(column => csvEscape(item[column]))
        .join(",")
    )
  ];

  return lines.join("\n");
}

function csvEscape(value) {
  const stringValue = String(value ?? "");
  return `"${stringValue.replace(/"/g, '""')}"`;
}

function downloadTextFile(fileName, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
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
    els.comparisonInfo.textContent = "Upload a PDF to compare it with the last saved version.";
    return;
  }

  els.headerMeta.textContent = `${state.datasetMeta.fileName} | ${state.datasetMeta.rowCount} initiatives | uploaded ${formatDate(state.datasetMeta.uploadedAt)}`;
  els.datasetInfo.textContent = `Stored in this browser only. File type: ${state.datasetMeta.fileType}. Parser: ${state.datasetMeta.parserVersion}.`;

  const summary = state.datasetMeta.comparisonSummary;
  if (summary) {
    const previousText = state.datasetMeta.previousFileName
      ? ` vs ${state.datasetMeta.previousFileName}`
      : "";
    els.comparisonInfo.textContent =
      `Comparison${previousText}: ${summary.newCount} new, ${summary.changedCount} changed, ${summary.removedCount} removed.`;
  } else {
    els.comparisonInfo.textContent = "No comparison history is available yet.";
  }
}

function populateFilters() {
  fillSelect(els.sectionFilter, unique(state.raw.map(item => item.sectionName)));
  fillSelect(els.themeFilter, unique(state.raw.map(item => item.theme)));
  fillSelect(els.ownerFilter, unique(state.raw.map(item => item.primaryOwner)));
  fillSelect(els.stageFilter, unique(state.raw.map(item => item.stageLabel)));
}

function fillSelect(selectEl, values) {
  const current = selectEl.value;
  selectEl.innerHTML = `<option value="">All</option>`;
  values.forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    if (value === current) option.selected = true;
    selectEl.appendChild(option);
  });
}

function applyFilters() {
  const q = els.searchInput.value.trim().toLowerCase();
  const section = els.sectionFilter.value;
  const theme = els.themeFilter.value;
  const owner = els.ownerFilter.value;
  const stage = els.stageFilter.value;
  const change = els.changeFilter.value;
  const impact = els.impactFilter.value;
  const minRelevance = Number(els.relevanceFilter.value || 0);
  const minParseConfidence = Number(els.parseConfidenceFilter.value || 0);
  const fmrukOnly = els.fmrukOnlyFilter.checked;
  const excludeAnnex = els.excludeAnnexFilter.checked;
  const needsReviewOnly = els.needsReviewFilter.checked;

  const items = state.raw
    .filter(item => {
      const haystack = [
        item.sectionName,
        item.subcategory,
        item.leadRegulator,
        item.initiativeTitle,
        item.initiativeDescription,
        item.theme,
        item.internalSubTheme,
        item.primaryOwner,
        item.secondaryOwner,
        item.stageLabel,
        item.potentialBusinessImpact,
        item.rationale,
        ...(item.classificationSignals || []),
        ...(item.relevanceSignals || []),
        ...(item.relevanceNegativeSignals || []),
        ...(item.reviewReasons || [])
      ]
        .join(" ")
        .toLowerCase();

      return (
        (!q || haystack.includes(q)) &&
        (!section || item.sectionName === section) &&
        (!theme || item.theme === theme) &&
        (!owner || item.primaryOwner === owner) &&
        (!stage || item.stageLabel === stage) &&
        (!change || item.changeStatus === change) &&
        (!impact || item.impactLevel === impact) &&
        (item.relevanceScore || 0) >= minRelevance &&
        (item.parseConfidence || 0) >= minParseConfidence &&
        (!fmrukOnly || item.isFmrukRelevant === true) &&
        (!excludeAnnex || item.sectionName !== "Annex: initiatives completed/stopped") &&
        (!needsReviewOnly || item.needsReview === true)
      );
    })
    .sort((a, b) => {
      if ((b.immediateActionRequired ? 1 : 0) !== (a.immediateActionRequired ? 1 : 0)) {
        return (b.immediateActionRequired ? 1 : 0) - (a.immediateActionRequired ? 1 : 0);
      }
      return (b.relevanceScore || 0) - (a.relevanceScore || 0);
    });

  state.filtered = items;

  if (!items.find(item => item.id === state.selectedItemId)) {
    state.selectedItemId = items[0]?.id || null;
  }

  renderSummary(items);
  renderTable(items);
  renderDetail(items.find(item => item.id === state.selectedItemId));
}

function renderSummary(items) {
  els.kpiTotal.textContent = items.length;
  els.kpiImmediateAction.textContent = items.filter(item => item.immediateActionRequired).length;
  els.kpiHighRelevance.textContent = items.filter(item => (item.relevanceScore || 0) >= 80).length;
  els.kpiHighImpact.textContent = items.filter(item => item.impactLevel === "High").length;
  els.kpiNeedsReview.textContent = items.filter(item => item.needsReview).length;
  els.kpiChanges.textContent = items.filter(item => item.changeStatus !== "Existing").length;

  renderFrequencyList(els.topThemesList, items.map(item => item.theme));
  renderFrequencyList(els.topOwnersList, items.map(item => item.primaryOwner));
  renderItemList(
    els.immediateActionsList,
    items.filter(item => item.immediateActionRequired).slice(0, 5),
    item => `${item.initiativeTitle} (${item.primaryOwner})`
  );
  renderItemList(
    els.reviewList,
    items.filter(item => item.needsReview).slice(0, 5),
    item => `${item.initiativeTitle} (${truncateText(item.reviewReasons.join(" "), 80)})`
  );
  renderItemList(
    els.changesList,
    items.filter(item => item.changeStatus !== "Existing").slice(0, 5),
    item => `${item.changeStatus}: ${item.initiativeTitle}`
  );
}

function renderFrequencyList(target, values) {
  target.innerHTML = "";
  const counts = {};
  values.filter(Boolean).forEach(value => {
    counts[value] = (counts[value] || 0) + 1;
  });

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

function renderItemList(target, items, formatter) {
  target.innerHTML = "";
  if (!items.length) {
    target.innerHTML = "<li>None</li>";
    return;
  }

  items.forEach(item => {
    const li = document.createElement("li");
    li.textContent = formatter(item);
    target.appendChild(li);
  });
}

function renderTable(items) {
  els.tableBody.innerHTML = "";

  if (!items.length) {
    els.tableBody.innerHTML = `<tr><td colspan="8">No initiatives match the current filters.</td></tr>`;
    return;
  }

  items.forEach(item => {
    const tr = document.createElement("tr");
    if (item.id === state.selectedItemId) tr.classList.add("selected");

    tr.innerHTML = `
      <td>${priorityBadge(item.relevanceScore || 0)}</td>
      <td>${parseBadge(item.parseConfidence || 0)}</td>
      <td>${stageBadge(item.stageLabel || "Monitoring")}</td>
      <td>${changeBadge(item.changeStatus || "Existing")}</td>
      <td>${escapeHtml(item.sectionName || "")}</td>
      <td>${escapeHtml(item.theme || "")}</td>
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
    els.detailPanel.innerHTML = "Upload the FCA Grid PDF, then select an initiative.";
    return;
  }

  const sourcePages = item.sourcePages?.length ? item.sourcePages.join(", ") : "N/A";
  const parseWarnings = item.parseWarnings?.length
    ? renderListHtml(item.parseWarnings)
    : "<div class=\"detail-copy\">No parser warnings were generated.</div>";
  const reviewReasons = item.reviewReasons?.length
    ? renderListHtml(item.reviewReasons)
    : "<div class=\"detail-copy\">No manual review flags were generated.</div>";
  const actionList = renderListHtml(item.suggestedActions || []);
  const positiveSignals = renderSignalChips(item.relevanceSignals || [], "positive");
  const negativeSignals = renderSignalChips(item.relevanceNegativeSignals || [], "negative");
  const classificationSignals = renderSignalChips(item.classificationSignals || [], "neutral");

  els.detailPanel.innerHTML = `
    <h2 class="detail-title">${escapeHtml(item.initiativeTitle || "")}</h2>
    <p class="detail-subtitle">Lead: ${escapeHtml(item.leadRegulator || "N/A")} | Source pages: ${escapeHtml(sourcePages)}</p>

    <div class="detail-grid">
      <div class="detail-box"><div class="detail-label">Section</div><div class="detail-value">${escapeHtml(item.sectionName || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Subcategory</div><div class="detail-value">${escapeHtml(item.subcategory || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Theme</div><div class="detail-value">${escapeHtml(item.theme || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Sub-theme</div><div class="detail-value">${escapeHtml(item.internalSubTheme || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Primary Owner</div><div class="detail-value">${escapeHtml(item.primaryOwner || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Secondary Owner</div><div class="detail-value">${escapeHtml(item.secondaryOwner || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Stage</div><div class="detail-value">${escapeHtml(item.stageLabel || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Change Status</div><div class="detail-value">${escapeHtml(item.changeStatus || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Relevance Score</div><div class="detail-value">${escapeHtml(String(item.relevanceScore || 0))}</div></div>
      <div class="detail-box"><div class="detail-label">Impact Level</div><div class="detail-value">${escapeHtml(item.impactLevel || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Timing Bucket</div><div class="detail-value">${escapeHtml(item.timingBucket || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Parse Confidence</div><div class="detail-value">${escapeHtml(`${item.parseConfidence || 0} (${item.parseConfidenceBand || "N/A"})`)}</div></div>
      <div class="detail-box"><div class="detail-label">Indicative Impact</div><div class="detail-value">${escapeHtml(item.indicativeImpactOnFirms || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Consumer Interest</div><div class="detail-value">${escapeHtml(item.consumerInterest || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">Timing Updated</div><div class="detail-value">${escapeHtml(item.timingUpdated || "N/A")}</div></div>
      <div class="detail-box"><div class="detail-label">New Flag</div><div class="detail-value">${escapeHtml(item.isNew || "N/A")}</div></div>
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Potential Business Impact</div>
      <div class="detail-copy">${escapeHtml(item.potentialBusinessImpact || "N/A")}</div>
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Expected Key Milestones</div>
      <div class="detail-copy">${escapeHtml(item.expectedKeyMilestones || "N/A")}</div>
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Initiative Description</div>
      <div class="detail-copy">${escapeHtml(item.initiativeDescription || "N/A")}</div>
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Why It Matters To FMRUK</div>
      <div class="detail-copy">${escapeHtml(item.rationale || "N/A")}</div>
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Suggested Actions</div>
      ${actionList}
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Matched Classification Signals</div>
      <div class="signal-list">${classificationSignals}</div>
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Positive Relevance Signals</div>
      <div class="signal-list">${positiveSignals}</div>
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Counter-Signals</div>
      <div class="signal-list">${negativeSignals}</div>
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Parser Warnings</div>
      ${parseWarnings}
    </div>

    <div class="detail-block-panel">
      <div class="detail-block-title">Manual Review Flags</div>
      ${reviewReasons}
    </div>
  `;
}

function renderListHtml(values) {
  if (!values?.length) {
    return `<div class="detail-copy">None</div>`;
  }

  const items = values
    .map(value => `<li>${escapeHtml(value)}</li>`)
    .join("");
  return `<ul class="detail-list">${items}</ul>`;
}

function renderSignalChips(values, type) {
  if (!values?.length) {
    return `<span class="signal-chip signal-chip-${type}">None</span>`;
  }

  return values
    .map(
      value =>
        `<span class="signal-chip signal-chip-${type}">${escapeHtml(value)}</span>`
    )
    .join("");
}

function priorityBadge(score) {
  if (score >= 85) return `<span class="badge badge-high">High</span>`;
  if (score >= 58) return `<span class="badge badge-medium">Medium</span>`;
  return `<span class="badge badge-low">Low</span>`;
}

function parseBadge(score) {
  if (score >= 85) return `<span class="badge badge-high">High</span>`;
  if (score >= 70) return `<span class="badge badge-medium">Medium</span>`;
  return `<span class="badge badge-review">Review</span>`;
}

function stageBadge(label) {
  const lowKey = String(label || "").toLowerCase();
  if (lowKey.includes("final") || lowKey.includes("reporting")) {
    return `<span class="badge badge-high">${escapeHtml(label)}</span>`;
  }
  if (lowKey.includes("consultation") || lowKey.includes("supervisory")) {
    return `<span class="badge badge-medium">${escapeHtml(label)}</span>`;
  }
  return `<span class="badge badge-neutral">${escapeHtml(label)}</span>`;
}

function changeBadge(status) {
  if (status === "New") return `<span class="badge badge-high">New</span>`;
  if (status === "Changed") return `<span class="badge badge-medium">Changed</span>`;
  return `<span class="badge badge-neutral">Existing</span>`;
}

function buildBlob(item) {
  return [
    item.sectionName,
    item.subcategory,
    item.leadRegulator,
    item.initiativeTitle,
    item.initiativeDescription,
    item.expectedKeyMilestones,
    item.rawText
  ]
    .join(" ")
    .toLowerCase();
}

function scoreSignalGroups(groups, blob) {
  let score = 0;
  const reasons = [];

  groups.forEach(group => {
    const matchedTerms = group.terms.filter(term => blob.includes(term.toLowerCase()));
    if (matchedTerms.length) {
      score += group.weight;
      reasons.push(group.reason);
    }
  });

  return {
    score,
    reasons: dedupeStrings(reasons)
  };
}

function inferTimingBucket(text) {
  const blob = String(text || "").toLowerCase();
  if (
    /\bq1\b/.test(blob) ||
    /\bq2\b/.test(blob) ||
    /\bh1\b/.test(blob) ||
    /\bjan(uary)?\b/.test(blob) ||
    /\bfeb(ruary)?\b/.test(blob) ||
    /\bmar(ch)?\b/.test(blob) ||
    /\bapr(il)?\b/.test(blob) ||
    /\bmay\b/.test(blob) ||
    /\bjune?\b/.test(blob)
  ) {
    return "Near Term";
  }

  if (
    /\bq3\b/.test(blob) ||
    /\bq4\b/.test(blob) ||
    /\bh2\b/.test(blob) ||
    /\bjul(y)?\b/.test(blob) ||
    /\baug(ust)?\b/.test(blob) ||
    /\bsep(t|tember)?\b/.test(blob) ||
    /\boct(ober)?\b/.test(blob) ||
    /\bnov(ember)?\b/.test(blob) ||
    /\bdec(ember)?\b/.test(blob) ||
    /\b2026\b/.test(blob)
  ) {
    return "Medium Term";
  }

  if (/\b2027\b/.test(blob) || /post july 2027/.test(blob)) {
    return "Longer Term";
  }

  return "To Be Confirmed";
}

function extractMilestones(text) {
  const matches = [];
  const patterns = [
    /\b(?:q[1-4]|h[12])\s+\d{4}\b[^.]*\.?/gi,
    /\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s*[-/]\s*(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{4}\b/gi,
    /\b(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}\b[^.]*\.?/gi,
    /\b\d{1,2}\s+(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}\b[^.]*\.?/gi
  ];

  patterns.forEach(pattern => {
    const found = text.match(pattern);
    if (found) matches.push(...found);
  });

  return dedupeStrings(matches.map(normaliseWs)).slice(0, 5).join(" | ");
}

function looksLikeMilestoneText(text) {
  const value = String(text || "").toLowerCase();
  return /\b(?:q[1-4]|h[12]|january|february|march|april|may|june|july|august|september|october|november|december|202[5-9])\b/.test(value);
}

function splitLeadAndTitle(line) {
  const pattern = new RegExp(
    `^((?:${REGULATOR_CODES.join("|")})(?:\\/(?:${REGULATOR_CODES.join("|")}))*)\\s+(.+)$`,
    "i"
  );
  const match = normaliseWs(line).match(pattern);
  if (!match) return null;

  const lead = normaliseLeadToken(match[1]);
  const title = cleanTitleCandidate(match[2]);
  if (!lead || !title) return null;

  return { lead, title };
}

function isPlausibleTitle(value) {
  const title = normaliseWs(value);
  if (!title || title.length < 6) return false;
  if (isLeadToken(title)) return false;
  if (ROW_NOISE_PATTERNS.some(pattern => pattern.test(title))) return false;
  return !HEADER_TEXT_PATTERNS.some(pattern => pattern.test(title));
}

function cleanTitleCandidate(value) {
  let title = normaliseWs(value);
  const pattern = new RegExp(
    `^(?:${REGULATOR_CODES.join("|")})(?:\\/(?:${REGULATOR_CODES.join("|")}))*\\s+`,
    "i"
  );
  title = title.replace(pattern, "");
  return normaliseWs(title);
}

function stripLeadPrefix(value, lead) {
  const text = normaliseWs(value);
  if (!text || !lead) return text;
  if (text === lead) return "";
  if (text.startsWith(`${lead} `)) return normaliseWs(text.slice(lead.length));
  return text;
}

function stripRepeatedTitle(value, title) {
  const text = normaliseWs(value);
  const cleanTitle = normaliseWs(title);
  if (!text) return "";
  if (!cleanTitle) return text;
  if (text === cleanTitle) return "";
  if (text.startsWith(`${cleanTitle} `)) {
    return normaliseWs(text.slice(cleanTitle.length));
  }
  return text;
}

function normaliseLeadToken(value) {
  const text = normaliseWs(value).toUpperCase().replace(/\s+/g, "");
  if (!text) return "";
  const parts = text.split("/");
  if (!parts.every(part => REGULATOR_CODES.includes(part))) return "";
  return parts.join("/");
}

function isLeadToken(value) {
  return Boolean(normaliseLeadToken(value));
}

function extractImpactFlag(value) {
  const text = normaliseWs(value).toUpperCase();
  if (!text) return "";
  if (/\bHIGH\b/.test(text) || /\bH\b/.test(text)) return "H";
  if (/\bMEDIUM\b/.test(text) || /\bM\b/.test(text) || /\bU\b/.test(text)) return "U";
  if (/\bLOW\b/.test(text) || /\bL\b/.test(text)) return "L";
  return "";
}

function normaliseTrafficLight(value) {
  const text = normaliseWs(value).toUpperCase();
  if (!text) return "";
  if (/\b(?:H|HIGH)\b/.test(text)) return "H";
  if (/\b(?:M|MEDIUM)\b/.test(text)) return "M";
  if (/\b(?:L|LOW)\b/.test(text)) return "L";
  return "";
}

function normaliseYesNo(value) {
  const text = normaliseWs(value).toLowerCase();
  if (!text) return "";
  if (/\b(yes|y|new)\b/.test(text)) return "Yes";
  if (/\b(no|n)\b/.test(text)) return "No";
  return "";
}

function buildCanonicalKey(item) {
  return [
    normaliseWs(item.sectionName).toLowerCase(),
    normaliseWs(item.leadRegulator).toLowerCase(),
    normaliseWs(item.initiativeTitle).toLowerCase()
  ].join("|");
}

function appendUnique(target, value) {
  if (value == null || value === "") return;
  if (Array.isArray(target)) {
    if (!target.includes(value)) target.push(value);
  }
}

function dedupeStrings(values) {
  return [...new Set(values.map(value => normaliseWs(value)).filter(Boolean))];
}

function unique(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function truncateText(value, maxLength) {
  const text = String(value || "");
  if (text.length <= maxLength) return text;
  return `${text.slice(0, maxLength - 3)}...`;
}

function slugify(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 60);
}

function normaliseWs(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function formatDate(value) {
  if (!value) return "Unknown";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString("en-GB");
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
