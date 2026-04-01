const STORAGE_KEY = "fmruk_reg_intelligence_v10";
const PROFILE_KEY = "fmruk_profile_v10";
const FEEDBACK_KEY = "fmruk_feedback_v10";
const PDF_DB_NAME = "fmruk_pdf_workspace";
const PDF_STORE_NAME = "pdfs";
const PDF_RECORD_KEY = "current_pdf";
const APP_VERSION = "1.7.0";
const APP_UPDATED_AT = "01 April 2026";

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

const state = {
  currentPage: "workspace",
  currentWorkspaceView: "monitor",
  baseItems: [],
  raw: [],
  filtered: [],
  selectedItemId: null,
  datasetMeta: null,
  profile: null,
  feedback: {},
  pdfBuffer: null,
  pdfDocument: null,
  pdfRenderToken: 0,
  currentPdfHighlightKey: "",
  ask: {
    query: "",
    answer: "",
    results: []
  }
};

const PROFILE_FIELDS = [
  ["businessLines", "businessLinesProfile"],
  ["permissions", "permissionsProfile"],
  ["products", "productsProfile"],
  ["distributionModel", "distributionProfile"],
  ["outsourcedProviders", "outsourcingProfile"],
  ["reportingObligations", "reportingProfile"],
  ["committees", "committeesProfile"],
  ["regulatoryTouchpoints", "touchpointsProfile"]
];

const DEFAULT_PROFILE = {
  priorityMode: "balanced",
  effectMode: "balanced",
  ownershipMode: "balanced",
  businessLines: [
    "investment management",
    "fund governance",
    "wholesale markets",
    "portfolio management"
  ],
  permissions: [
    "mifid investment firm",
    "portfolio management",
    "arranging",
    "dealing"
  ],
  products: [
    "institutional mandates",
    "funds",
    "wholesale products",
    "investment services"
  ],
  distributionModel: [
    "institutional",
    "professional clients",
    "wholesale distribution",
    "limited retail"
  ],
  outsourcedProviders: [
    "fund administration",
    "market data vendors",
    "reporting vendors",
    "technology service providers"
  ],
  reportingObligations: [
    "transaction reporting",
    "mifidpru",
    "regulatory returns",
    "outsourcing register"
  ],
  committees: [
    "board",
    "risk committee",
    "compliance committee",
    "oversight committee"
  ],
  regulatoryTouchpoints: [
    "fca",
    "mifid",
    "mifidpru",
    "operational resilience",
    "consumer duty"
  ]
};

const FMRUK_EFFECT_MODES = {
  strict: {
    label: "Direct effect only",
    positiveMultiplier: 0.92,
    negativeMultiplier: 1.22,
    profileMultiplier: 1.08,
    threshold: 64,
    watchlistThreshold: 50,
    profileBias: 0
  },
  balanced: {
    label: "Balanced",
    positiveMultiplier: 1,
    negativeMultiplier: 1,
    profileMultiplier: 1,
    threshold: 58,
    watchlistThreshold: 46,
    profileBias: 3
  },
  broad: {
    label: "Watchlist inclusive",
    positiveMultiplier: 1.08,
    negativeMultiplier: 0.88,
    profileMultiplier: 0.96,
    threshold: 52,
    watchlistThreshold: 40,
    profileBias: 6
  }
};

const OWNERSHIP_MODES = {
  balanced: {
    label: "Balanced",
    biasOwners: []
  },
  controls: {
    label: "Control-led",
    biasOwners: ["Compliance", "Risk", "Legal"]
  },
  operations: {
    label: "Operations-led",
    biasOwners: ["Operations", "Technology", "Finance"]
  }
};

const PROFILE_PRIORITY_MODES = {
  balanced: {
    label: "Balanced",
    weights: {
      relevance: 0.38,
      impact: 0.28,
      urgency: 0.22,
      delivery: 0.12
    },
    ownerBoosts: [],
    obligationBoosts: [],
    themeBoosts: []
  },
  governance: {
    label: "Governance and policy",
    weights: {
      relevance: 0.36,
      impact: 0.26,
      urgency: 0.18,
      delivery: 0.20
    },
    ownerBoosts: ["Compliance", "Legal", "HR"],
    obligationBoosts: [
      "Governance & Accountability",
      "Policy & Procedures",
      "Training & Communications",
      "Disclosure & Communications"
    ],
    themeBoosts: ["Governance / SMCR", "AML / Sanctions / Financial Crime"]
  },
  operations: {
    label: "Operations and reporting",
    weights: {
      relevance: 0.32,
      impact: 0.28,
      urgency: 0.22,
      delivery: 0.18
    },
    ownerBoosts: ["Operations", "Technology", "Finance"],
    obligationBoosts: [
      "Reporting & MI",
      "Operational Resilience",
      "Outsourcing & Third Parties",
      "Operating Model"
    ],
    themeBoosts: [
      "Reporting / Disclosure",
      "Operational Resilience",
      "Operations / Post-Trade"
    ]
  },
  markets: {
    label: "Markets and post-trade",
    weights: {
      relevance: 0.34,
      impact: 0.28,
      urgency: 0.22,
      delivery: 0.16
    },
    ownerBoosts: ["Operations", "Compliance"],
    obligationBoosts: [
      "Market Conduct & Surveillance",
      "Reporting & MI",
      "Operating Model"
    ],
    themeBoosts: [
      "Market Conduct / Trading",
      "Market Structure / Wholesale",
      "Operations / Post-Trade"
    ]
  },
  executive: {
    label: "Executive escalation",
    weights: {
      relevance: 0.30,
      impact: 0.34,
      urgency: 0.24,
      delivery: 0.12
    },
    ownerBoosts: ["Compliance", "Operations", "Risk", "Legal"],
    obligationBoosts: [
      "Governance & Accountability",
      "Operational Resilience",
      "Reporting & MI",
      "Prudential & Capital"
    ],
    themeBoosts: []
  }
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

const EXCEL_QUARTER_COLUMNS = [
  { aliases: ["oct - dec 2025", "oct-dec 2025"], label: "Oct-Dec 2025", month: 9, year: 2025 },
  { aliases: ["jan-mar 2026", "jan - mar 2026"], label: "Jan-Mar 2026", month: 0, year: 2026 },
  { aliases: ["apr-jun 2026", "apr - jun 2026"], label: "Apr-Jun 2026", month: 3, year: 2026 },
  { aliases: ["jul- sep 2026", "jul-sep 2026", "jul - sep 2026"], label: "Jul-Sep 2026", month: 6, year: 2026 },
  { aliases: ["oct- dec 2026", "oct-dec 2026", "oct - dec 2026"], label: "Oct-Dec 2026", month: 9, year: 2026 },
  { aliases: ["jan-jun 2027", "jan - jun 2027"], label: "Jan-Jun 2027", month: 0, year: 2027 },
  { aliases: ["post july 2027"], label: "Post July 2027", month: 6, year: 2027 }
];

const SECTION_LOOKUP = new Map(
  KNOWN_SECTIONS.map(value => [value.toLowerCase(), value])
);

const SUBCATEGORY_LOOKUP = new Map(
  KNOWN_SUBCATEGORIES.map(value => [value.toLowerCase(), value])
);

const SECTION_ALIASES = new Map([
  ...KNOWN_SECTIONS.map(value => [normaliseLookupKey(value), value]),
  ["investment management", "Investment management"],
  ["retail investment", "Retail investments"],
  ["pensions", "Pensions and retirement income"],
  ["payment services and cryptoassets", "Payment services and cryptoassets"]
]);

const SUBCATEGORY_ALIASES = new Map([
  ...KNOWN_SUBCATEGORIES.map(value => [normaliseLookupKey(value), value]),
  ["competition innovation and other", "Competition, innovation and other"],
  ["cross-cutting / omnibus", "Cross-cutting/omnibus"],
  ["cross-cutting omnibus", "Cross-cutting/omnibus"],
  ["n/a", "N/A"],
  ["n/a ", "N/A"]
]);

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
  consumer: 0.84,
  timing: 0.9,
  isNew: 0.95
};

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
  /^jul-dec \d{4}$/i,
  /^jan-jun \d{4}$/i,
  /^post july \d{4}$/i,
  /^july to december \d{4}$/i,
  /^beyond (that )?horizon$/i,
  /^quarter-by-quarter$/i,
  /^e\s*formal engagement planned/i
];

const CLASSIFICATION_RULES = [
  {
    id: "settlement-cycle",
    theme: "Operations / Post-Trade",
    subTheme: "Settlement / T+1",
    primaryOwner: "Operations",
    secondaryOwner: "Compliance",
    impactBias: 28,
    impactStatement:
      "Could affect trade allocation, confirmation, affirmation, settlement instructions, fails management, liquidity planning and oversight of post-trade providers.",
    signals: [
      {
        weight: 28,
        reason: "Directly references T+1 or settlement-cycle reform.",
        terms: ["t+1", "t plus one", "settlement cycle", "shortened settlement cycle"]
      },
      {
        weight: 18,
        reason: "Touches post-trade operations or settlement processing.",
        terms: ["settlement", "affirmation", "allocation", "post-trade", "settlement fails"]
      }
    ],
    sectionBoosts: {
      "Wholesale financial markets": 8,
      "Investment management": 4
    },
    coreActions: [
      "Assess front-to-back readiness across trading, operations, custodians and outsourced post-trade providers.",
      "Review cut-off times, exception handling, client communication points and settlement fail monitoring."
    ]
  },
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
      "Confirm scope across desks, venues and reporting entities.",
      "Review field mapping, reconciliations, exception MI and vendor dependency."
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
      "Check execution policy, venue oversight and client disclosures.",
      "Review monitoring MI, committee reporting and broker oversight."
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
      "Check surveillance scenarios, alert governance and escalation.",
      "Reconfirm training, insider list and employee dealing controls."
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
      "Assess ICARA assumptions, thresholds and governance.",
      "Check prudential reporting, modelling and board packs."
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
      "Check remuneration policy, governance approvals and identified staff scope.",
      "Assess pay design, deferral mechanics and attestations."
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
      "Review impacted services, dependency mapping and scenario testing.",
      "Check board reporting, self-assessment content and remediation governance."
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
      "Confirm whether any material outsourcing arrangement is in scope.",
      "Review contracts, registers, due diligence and exit or concentration controls."
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
      "Check whether any FMRUK products, disclosures or distribution are in scope.",
      "Review fair value, target market, communications and monitoring evidence."
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
      "Assess onboarding, screening, investigations and monitoring scenarios.",
      "Check policy documents, training and escalation paths."
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
      "Check accountabilities, terms of reference and certification processes.",
      "Assess statements of responsibility, attestations and committee packs."
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
      "Review product disclosures, marketing materials and approval controls.",
      "Check governance, evidence packs and label eligibility assessments."
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
      "Assess whether internal AI, cyber or data governance needs to change.",
      "Check inventories, controls, testing and oversight."
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
      "Confirm which products, mandates or fund structures are in scope.",
      "Assess disclosures, delegated oversight and product governance."
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
      "Check whether custody model, CASS oversight or reconciliations are impacted.",
      "Assess governance forums, attestations and third-party oversight."
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
      "Confirm ownership for data, controls and sign-off.",
      "Assess system changes, sourcing and testing before live reporting."
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
      "Confirm whether the topic is in scope for FMRUK.",
      "If relevant, assess perimeter, client proposition and outsourced provider impacts."
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
      "Assess disclosures, governance and operating model controls."
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
      "Identify the control domains in scope.",
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
      "Keep it on the watchlist until scope and timing are clearer."
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
    signals: [{ weight: 10, terms: ["roadmap", "work programme", "strategy", "monitoring"] }]
  }
];

const OBLIGATION_RULES = [
  {
    name: "Policy & Procedures",
    signals: [
      { weight: 10, terms: ["policy", "handbook", "rule change", "guidance", "procedure"] },
      { weight: 8, terms: ["control framework", "process update", "implementation"] }
    ]
  },
  {
    name: "Governance & Accountability",
    signals: [
      { weight: 12, terms: ["governance", "board", "committee", "smcr", "accountability"] },
      { weight: 8, terms: ["oversight", "dear ceo", "attestation", "senior manager"] }
    ]
  },
  {
    name: "Training & Communications",
    signals: [
      { weight: 10, terms: ["training", "communication", "client communication", "employee dealing"] },
      { weight: 6, terms: ["disclosure", "consumer", "outcomes"] }
    ]
  },
  {
    name: "Reporting & MI",
    signals: [
      { weight: 12, terms: ["reporting", "data collection", "return", "submission", "template"] },
      { weight: 8, terms: ["mi", "attestation", "notification", "reconciliation"] }
    ]
  },
  {
    name: "Product & Distribution",
    signals: [
      { weight: 10, terms: ["product governance", "distribution", "fund", "target market"] },
      { weight: 8, terms: ["consumer duty", "fair value", "product"] }
    ]
  },
  {
    name: "Market Conduct & Surveillance",
    signals: [
      { weight: 12, terms: ["market abuse", "surveillance", "best execution", "transaction reporting"] },
      { weight: 8, terms: ["mifir", "stors", "inside information"] }
    ]
  },
  {
    name: "Outsourcing & Third Parties",
    signals: [
      { weight: 12, terms: ["outsourcing", "third party", "vendor", "service provider"] },
      { weight: 8, terms: ["critical third party", "supplier", "contractual"] }
    ]
  },
  {
    name: "Operational Resilience",
    signals: [
      { weight: 12, terms: ["operational resilience", "impact tolerance", "important business service"] },
      { weight: 8, terms: ["scenario testing", "mapping", "resilience"] }
    ]
  },
  {
    name: "Prudential & Capital",
    signals: [
      { weight: 12, terms: ["mifidpru", "icara", "capital", "liquidity", "own funds"] },
      { weight: 8, terms: ["prudential", "k-factor", "fixed overhead"] }
    ]
  },
  {
    name: "Disclosure & Communications",
    signals: [
      { weight: 10, terms: ["disclosure", "sustainability", "sdr", "greenwashing"] },
      { weight: 8, terms: ["consumer", "marketing", "naming", "label"] }
    ]
  },
  {
    name: "Legal Entity Perimeter",
    signals: [
      { weight: 12, terms: ["legislation", "fsma", "authorisation", "perimeter", "regime"] },
      { weight: 8, terms: ["hm treasury", "statutory instrument", "repeal and replacement"] }
    ]
  },
  {
    name: "Operating Model",
    signals: [
      { weight: 10, terms: ["operating model", "vendor", "technology", "data", "digital"] },
      { weight: 8, terms: ["outsourcing", "operational resilience", "implementation"] }
    ]
  }
];

const RELEVANCE_SIGNAL_GROUPS = {
  positive: [
    {
      weight: 18,
      reason: "Touches investment management, funds or portfolio activity.",
      terms: ["investment management", "asset management", "fund", "portfolio management", "ucits", "aifmd"]
    },
    {
      weight: 18,
      reason: "Touches MiFID, prudential or wholesale market obligations.",
      terms: ["mifid", "mifidpru", "investment firm", "wholesale", "market abuse", "best execution", "transaction reporting"]
    },
    {
      weight: 15,
      reason: "Touches governance, resilience or outsourcing controls.",
      terms: ["operational resilience", "outsourcing", "third party", "governance", "smcr", "dear ceo"]
    },
    {
      weight: 14,
      reason: "Touches reporting or submissions likely to sit with the UK entity.",
      terms: ["regulatory reporting", "reporting", "data collection", "return", "attestation", "notification"]
    },
    {
      weight: 12,
      reason: "Touches financial crime or onboarding controls.",
      terms: ["aml", "anti-money laundering", "sanctions", "kyc", "financial crime"]
    },
    {
      weight: 10,
      reason: "Touches sustainability, disclosures or technology change with UK impact.",
      terms: ["sustainability", "sdr", "greenwashing", "cyber", "ai", "data"]
    }
  ],
  negative: [
    {
      weight: 20,
      reason: "Looks primarily insurance-specific.",
      terms: ["insurance", "reinsurance", "solvency", "policyholder"]
    },
    {
      weight: 18,
      reason: "Looks primarily pensions-specific.",
      terms: ["pensions", "retirement income", "defined benefit", "defined contribution"]
    },
    {
      weight: 16,
      reason: "Looks primarily retail banking or consumer credit-specific.",
      terms: ["mortgage", "consumer credit", "banking", "deposit", "current account"]
    },
    {
      weight: 12,
      reason: "Looks primarily payments or crypto-specific unless FMRUK scope is explicit.",
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

const GENERIC_STAGE_ACTIONS = {
  consultation: [
    "Decide scope and whether FMRUK should respond or feed a group response.",
    "Capture policy points and delivery risks early."
  ],
  final_rules: [
    "Run a rule-to-control gap review and assign delivery dates.",
    "Confirm policy, procedure, training and MI updates before go-live."
  ],
  supervisory: [
    "Benchmark current controls against the supervisory expectation.",
    "Prepare management reporting on current state, gaps and remediation."
  ],
  reporting: [
    "Confirm perimeter, data lineage, control owners and sign-off.",
    "Plan testing for completeness, exceptions and reconciliations."
  ],
  legislation: [
    "Assess whether permissions, entity scope or delegated arrangements change.",
    "Coordinate legal interpretation and downstream policy updates."
  ],
  monitoring: [
    "Keep the item on the watchlist and refresh the assessment when timing or scope changes.",
    "Keep an owner assigned so the business is not surprised if the item accelerates."
  ]
};

const MONTH_LOOKUP = {
  january: 0,
  february: 1,
  march: 2,
  april: 3,
  may: 4,
  june: 5,
  july: 6,
  august: 7,
  september: 8,
  october: 9,
  november: 10,
  december: 11
};

const els = {};

document.addEventListener("DOMContentLoaded", () => {
  init().catch(err => {
    console.error(err);
  });
});

async function init() {
  mapEls();
  bindEvents();
  state.currentPage = getPageFromHash();
  state.profile = loadProfile();
  state.feedback = loadFeedback();
  loadFromStorage();
  renderProfileForm();
  updateReviewClock();
  window.setInterval(updateReviewClock, 30000);
  await hydratePdfFromStorage();
  renderPageState();
  renderWorkspaceView();
  renderAll();
}

function mapEls() {
  const ids = [
    "navWorkspaceBtn",
    "navSettingsBtn",
    "reviewPageMeta",
    "workspacePage",
    "settingsPage",
    "headerMeta",
    "portfolioNarrative",
    "reloadBtn",
    "exportJsonBtn",
    "exportCsvBtn",
    "exportBoardBriefBtn",
    "exportOwnerPackBtn",
    "clearStorageBtn",
    "excelInput",
    "fileInput",
    "uploadBtn",
    "uploadStatus",
    "sourceSummary",
    "searchInput",
    "workspaceModeTabs",
    "workspaceModeMeta",
    "monitorPane",
    "timelinePane",
    "sourcePane",
    "sectionFilter",
    "themeFilter",
    "ownerFilter",
    "stageFilter",
    "changeFilter",
    "clusterFilter",
    "relevanceFilter",
    "parseConfidenceFilter",
    "fmrukOnlyFilter",
    "excludeAnnexFilter",
    "needsReviewFilter",
    "saveProfileBtn",
    "resetProfileBtn",
    "clusterList",
    "deltaList",
    "initiativeList",
    "listMeta",
    "detailPanel",
    "detailMeta",
    "timelineList",
    "timelineMeta",
    "pdfPreviewStatus",
    "evidenceTrail",
    "pdfPreview",
    "askBtn",
    "askAnswer",
    "askResults",
    "profilePriorityMode",
    "profileEffectMode",
    "profileOwnershipMode",
    "settingsSummary",
    "settingsStatus"
  ];

  ids.forEach(id => {
    els[id] = document.getElementById(id);
  });

  PROFILE_FIELDS.forEach(([, domId]) => {
    els[domId] = document.getElementById(domId);
  });
}

function bindEvents() {
  els.navWorkspaceBtn?.addEventListener("click", () => navigateToPage("workspace"));
  els.navSettingsBtn?.addEventListener("click", () => navigateToPage("settings"));
  els.workspaceModeTabs?.addEventListener("click", handleWorkspaceModeClick);
  window.addEventListener("hashchange", syncPageFromHash);
  els.uploadBtn.addEventListener("click", handleUpload);
  els.reloadBtn.addEventListener("click", () => renderAll());
  els.exportJsonBtn.addEventListener("click", exportJson);
  els.exportCsvBtn.addEventListener("click", exportCsv);
  els.exportBoardBriefBtn.addEventListener("click", exportBoardBrief);
  els.exportOwnerPackBtn.addEventListener("click", exportOwnerPack);
  els.clearStorageBtn.addEventListener("click", clearSavedData);
  els.askBtn.addEventListener("click", runAskQuery);
  els.searchInput.addEventListener("keydown", event => {
    if (event.key === "Enter") {
      event.preventDefault();
      runAskQuery();
    }
  });
  els.saveProfileBtn.addEventListener("click", saveProfileFromForm);
  els.resetProfileBtn.addEventListener("click", resetProfile);

  [
    els.searchInput,
    els.sectionFilter,
    els.themeFilter,
    els.ownerFilter,
    els.stageFilter,
    els.changeFilter,
    els.clusterFilter,
    els.relevanceFilter,
    els.parseConfidenceFilter,
    els.fmrukOnlyFilter,
    els.excludeAnnexFilter,
    els.needsReviewFilter
  ].forEach(element => {
    if (!element) return;
    const eventName = element.tagName === "INPUT" && element.type === "text" ? "input" : "change";
    element.addEventListener(eventName, applyFilters);
  });
}

function getPageFromHash() {
  return window.location.hash === "#settings" ? "settings" : "workspace";
}

function syncPageFromHash() {
  state.currentPage = getPageFromHash();
  renderPageState();
}

function navigateToPage(page) {
  const nextPage = page === "settings" ? "settings" : "workspace";
  const nextHash = nextPage === "settings" ? "#settings" : "#workspace";
  if (window.location.hash !== nextHash) {
    window.location.hash = nextHash;
    return;
  }
  state.currentPage = nextPage;
  renderPageState();
}

function renderPageState() {
  document.body.dataset.page = state.currentPage;
  els.workspacePage?.classList.toggle("is-active", state.currentPage === "workspace");
  els.settingsPage?.classList.toggle("is-active", state.currentPage === "settings");
  els.navWorkspaceBtn?.classList.toggle("is-active", state.currentPage === "workspace");
  els.navSettingsBtn?.classList.toggle("is-active", state.currentPage === "settings");
}

function handleWorkspaceModeClick(event) {
  const button = event.target.closest("[data-workspace-view]");
  if (!button) return;
  state.currentWorkspaceView = button.dataset.workspaceView || "monitor";
  renderWorkspaceView();
}

function renderWorkspaceView() {
  const active = state.currentWorkspaceView || "monitor";
  els.monitorPane?.classList.toggle("is-active", active === "monitor");
  els.timelinePane?.classList.toggle("is-active", active === "timeline");
  els.sourcePane?.classList.toggle("is-active", active === "source");

  els.workspaceModeTabs?.querySelectorAll("[data-workspace-view]").forEach(button => {
    button.classList.toggle("is-active", button.dataset.workspaceView === active);
  });

  if (els.workspaceModeMeta) {
    els.workspaceModeMeta.textContent = {
      monitor: "Monitor view keeps the queue, detail and analyst review visible.",
      timeline: "Milestones view groups dated events into chronological buckets using the structured grid timeline.",
      source: "Source view prioritises PDF evidence and cross-checking against the currently selected initiative."
    }[active] || "Monitor view keeps the queue, detail and analyst review visible.";
  }
}

function setStatusMessage(message, scope = "workspace") {
  if (scope === "settings") {
    if (els.settingsStatus) els.settingsStatus.textContent = message;
    if (els.uploadStatus && !state.baseItems.length) {
      els.uploadStatus.textContent = message;
    }
    return;
  }

  if (els.uploadStatus) els.uploadStatus.textContent = message;
}

function updateReviewClock() {
  if (!els.reviewPageMeta) return;
  const formatter = new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
  els.reviewPageMeta.textContent = `${formatter.format(new Date())} · v${APP_VERSION}`;
}

function loadFromStorage() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return;

  try {
    const payload = JSON.parse(raw);
    state.baseItems = Array.isArray(payload.parsedItems) ? payload.parsedItems : [];
    state.raw = Array.isArray(payload.items) ? payload.items : [];
    state.datasetMeta = payload.meta || null;
    state.selectedItemId = state.raw[0]?.id || null;
  } catch (err) {
    console.error(err);
  }
}

function saveToStorage() {
  try {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        meta: state.datasetMeta,
        parsedItems: state.baseItems,
        items: state.raw
      })
    );
  } catch (err) {
    console.error(err);
    els.uploadStatus.textContent =
      "Storage limit reached. The current session still works, but the dataset may not persist.";
  }
}

function clearSavedData() {
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(FEEDBACK_KEY);
  localStorage.removeItem(PROFILE_KEY);
  state.baseItems = [];
  state.raw = [];
  state.filtered = [];
  state.selectedItemId = null;
  state.datasetMeta = null;
  state.feedback = {};
  state.profile = structuredClone(DEFAULT_PROFILE);
  state.ask = { query: "", answer: "", results: [] };
  renderProfileForm();
  clearStoredPdf()
    .catch(console.error)
    .finally(() => {
      state.pdfBuffer = null;
      state.pdfDocument = null;
      setStatusMessage("Saved data cleared.");
      renderAll();
    });
}

function loadProfile() {
  const raw = localStorage.getItem(PROFILE_KEY);
  if (!raw) return structuredClone(DEFAULT_PROFILE);

  try {
    const parsed = JSON.parse(raw);
    return mergeProfile(parsed);
  } catch (err) {
    console.error(err);
    return structuredClone(DEFAULT_PROFILE);
  }
}

function saveProfileFromForm() {
  const nextProfile = {};
  PROFILE_FIELDS.forEach(([field, domId]) => {
    nextProfile[field] = splitLines(els[domId].value);
  });
  nextProfile.priorityMode = els.profilePriorityMode?.value || DEFAULT_PROFILE.priorityMode;
  nextProfile.effectMode = els.profileEffectMode?.value || DEFAULT_PROFILE.effectMode;
  nextProfile.ownershipMode = els.profileOwnershipMode?.value || DEFAULT_PROFILE.ownershipMode;

  state.profile = mergeProfile(nextProfile);
  localStorage.setItem(PROFILE_KEY, JSON.stringify(state.profile));
  renderSettingsSummary();
  setStatusMessage("Settings saved. Re-scoring current dataset.", "settings");
  reanalyseCurrentDataset();
}

function resetProfile() {
  state.profile = structuredClone(DEFAULT_PROFILE);
  localStorage.setItem(PROFILE_KEY, JSON.stringify(state.profile));
  renderProfileForm();
  renderSettingsSummary();
  setStatusMessage("Settings reset to defaults.", "settings");
  reanalyseCurrentDataset();
}

function renderProfileForm() {
  PROFILE_FIELDS.forEach(([field, domId]) => {
    if (!els[domId]) return;
    els[domId].value = (state.profile?.[field] || []).join(", ");
  });

  if (els.profilePriorityMode) {
    els.profilePriorityMode.value =
      PROFILE_PRIORITY_MODES[state.profile?.priorityMode] ? state.profile.priorityMode : DEFAULT_PROFILE.priorityMode;
  }
  if (els.profileEffectMode) {
    els.profileEffectMode.value =
      FMRUK_EFFECT_MODES[state.profile?.effectMode] ? state.profile.effectMode : DEFAULT_PROFILE.effectMode;
  }
  if (els.profileOwnershipMode) {
    els.profileOwnershipMode.value =
      OWNERSHIP_MODES[state.profile?.ownershipMode] ? state.profile.ownershipMode : DEFAULT_PROFILE.ownershipMode;
  }
  renderSettingsSummary();
}

function mergeProfile(profile) {
  const merged = structuredClone(DEFAULT_PROFILE);
  PROFILE_FIELDS.forEach(([field]) => {
    merged[field] = Array.isArray(profile?.[field])
      ? profile[field].map(normaliseWs).filter(Boolean)
      : structuredClone(DEFAULT_PROFILE[field]);
  });
  merged.priorityMode = PROFILE_PRIORITY_MODES[profile?.priorityMode]
    ? profile.priorityMode
    : DEFAULT_PROFILE.priorityMode;
  merged.effectMode = FMRUK_EFFECT_MODES[profile?.effectMode]
    ? profile.effectMode
    : DEFAULT_PROFILE.effectMode;
  merged.ownershipMode = OWNERSHIP_MODES[profile?.ownershipMode]
    ? profile.ownershipMode
    : DEFAULT_PROFILE.ownershipMode;
  return merged;
}

function getPriorityModeConfig() {
  return PROFILE_PRIORITY_MODES[state.profile?.priorityMode] || PROFILE_PRIORITY_MODES.balanced;
}

function getEffectModeConfig() {
  return FMRUK_EFFECT_MODES[state.profile?.effectMode] || FMRUK_EFFECT_MODES.balanced;
}

function getOwnershipModeConfig() {
  return OWNERSHIP_MODES[state.profile?.ownershipMode] || OWNERSHIP_MODES.balanced;
}

function renderSettingsSummary() {
  if (!els.settingsSummary) return;
  const priority = getPriorityModeConfig().label;
  const effect = getEffectModeConfig().label;
  const ownership = getOwnershipModeConfig().label;
  els.settingsSummary.textContent = `Priority: ${priority} · FMRUK effect: ${effect} · Ownership bias: ${ownership}`;
  if (els.settingsStatus && !els.settingsStatus.textContent) {
    els.settingsStatus.textContent =
      "Use Settings/Admin to control weighting, FMRUK effect thresholds, routing bias and entity-profile defaults.";
  }
}

function loadFeedback() {
  const raw = localStorage.getItem(FEEDBACK_KEY);
  if (!raw) return {};
  try {
    return JSON.parse(raw) || {};
  } catch (err) {
    console.error(err);
    return {};
  }
}

function saveFeedback() {
  localStorage.setItem(FEEDBACK_KEY, JSON.stringify(state.feedback));
}

async function hydratePdfFromStorage() {
  if (!(state.datasetMeta?.sourceFiles?.pdfFileName || state.datasetMeta?.fileType === "PDF")) {
    return;
  }

  try {
    const buffer = await loadStoredPdf();
    if (buffer) {
      state.pdfBuffer = buffer;
      state.pdfDocument = await loadPdfDocumentFromBuffer(buffer);
    }
  } catch (err) {
    console.error(err);
  }
}

async function handleUpload() {
  const pdfFile = els.fileInput.files[0] || null;
  const excelFile = els.excelInput?.files[0] || null;

  if (!pdfFile && !excelFile) {
    els.uploadStatus.textContent = "Select an Excel file, a PDF, or both.";
    return;
  }

  const previousItems = state.raw.slice();
  const previousMeta = state.datasetMeta;

  try {
    let excelItems = [];
    let pdfItems = [];

    if (excelFile) {
      const ext = excelFile.name.split(".").pop().toLowerCase();
      if (!["xlsx", "xls"].includes(ext)) {
        els.uploadStatus.textContent = "Excel input must be .xlsx or .xls.";
        return;
      }
      if (!window.XLSX) {
        els.uploadStatus.textContent = "Excel library unavailable.";
        return;
      }

      els.uploadStatus.textContent = `Reading structured workbook: ${excelFile.name}...`;
      const excelBuffer = await excelFile.arrayBuffer();
      excelItems = parseExcelFile(excelBuffer, excelFile.name);
    }

    if (pdfFile) {
      const ext = pdfFile.name.split(".").pop().toLowerCase();
      if (ext !== "pdf") {
        els.uploadStatus.textContent = "PDF input must be a .pdf file.";
        return;
      }
      if (!window.pdfjsLib) {
        els.uploadStatus.textContent = "PDF library unavailable.";
        return;
      }

      els.uploadStatus.textContent = `Reading PDF evidence: ${pdfFile.name}...`;
      const buffer = await pdfFile.arrayBuffer();
      pdfItems = await parsePdfFile(buffer, pdfFile.name);
      await saveStoredPdf(buffer);
      state.pdfBuffer = buffer;
      state.pdfDocument = await loadPdfDocumentFromBuffer(buffer);
    } else {
      state.pdfBuffer = null;
      state.pdfDocument = null;
      state.currentPdfHighlightKey = "";
      clearStoredPdf().catch(console.error);
    }

    const ingestedItems = excelItems.length
      ? mergeStructuredAndPdfItems(excelItems, pdfItems)
      : pdfItems;
    const deduped = dedupeItems(ingestedItems);

    if (!deduped.length) {
      els.uploadStatus.textContent = "No initiatives detected. Check the uploaded sources.";
      return;
    }

    els.uploadStatus.textContent = `Scoring ${deduped.length} items...`;
    const analysed = analyseRows(deduped);
    const comparison = compareWithPreviousDataset(previousItems, analysed);
    const preferredSource = excelItems.length ? "Excel" : "PDF";
    const sourceLabel = buildSourceLabel(excelFile?.name, pdfFile?.name);

    state.baseItems = deduped;
    state.raw = sortItems(comparison.items);
    state.datasetMeta = {
      fileName: sourceLabel,
      uploadedAt: new Date().toISOString(),
      rowCount: state.raw.length,
      fileType: excelItems.length && pdfItems.length ? "Excel + PDF" : preferredSource,
      preferredSource,
      structuredSourceCount: excelItems.length,
      pdfSupportCount: pdfItems.length,
      sourceFiles: {
        excelFileName: excelFile?.name || "",
        pdfFileName: pdfFile?.name || ""
      },
      parserVersion: "v11",
      previousFileName: previousMeta?.fileName || "",
      comparisonSummary: comparison.summary
    };
    state.selectedItemId = state.raw[0]?.id || null;
    state.currentPdfHighlightKey = "";

    saveToStorage();
    els.uploadStatus.textContent = `Loaded ${state.raw.length} items using ${preferredSource}${pdfFile && excelFile ? " with PDF support" : ""}.`;
    renderAll();
  } catch (err) {
    console.error(err);
    els.uploadStatus.textContent = `Upload failed. ${err.message || "The selected sources could not be parsed."}`;
  }
}

function buildSourceLabel(excelFileName, pdfFileName) {
  if (excelFileName && pdfFileName) {
    return `${excelFileName} + ${pdfFileName}`;
  }
  return excelFileName || pdfFileName || "Current sources";
}

function parseExcelFile(buffer, fileName) {
  const workbook = window.XLSX.read(buffer, { type: "array" });
  const sheetName =
    workbook.SheetNames.find(name => normaliseLookupKey(name) === "live initiatives") ||
    workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("No worksheet was found in the Excel file.");
  }

  const sheet = workbook.Sheets[sheetName];
  const rows = window.XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: false
  });

  return rows
    .map((row, index) => parseExcelInitiativeRow(row, index + 2, fileName))
    .filter(item => item && item.initiativeTitle);
}

function parseExcelInitiativeRow(row, rowNumber, fileName) {
  const title = cleanTitleCandidate(getStructuredRowValue(row, ["initiative title"]));
  if (!title) return null;

  const sectionName = canonicaliseSectionName(getStructuredRowValue(row, ["sector"]));
  const subcategory = canonicaliseSubcategoryName(getStructuredRowValue(row, ["sub-category"]));
  const leadRegulator = normaliseRegulatorField(getStructuredRowValue(row, ["regulators involved"]));
  const description = cleanStructuredText(getStructuredRowValue(row, ["description of initiative"]));
  const expectedKeyMilestones = cleanStructuredText(getStructuredRowValue(row, ["expected milestones"]));
  const sourceLink = cleanStructuredText(getStructuredRowValue(row, ["link"]));
  const impactFlag = extractImpactFlag(
    getStructuredRowValue(row, ["indicative impact on firms to implement"])
  );
  const consumerInterest = consumerInterestFromStructured(
    getStructuredRowValue(row, [
      "is this initiative likely to be of significant interest to consumers and consumer organisations?"
    ])
  );
  const timingUpdated = normaliseYesNo(
    getStructuredRowValue(row, ["has the timing of this initiative changed since the previous grid?"])
  );
  const publicationStatus = normalisePublicationStatus(
    getStructuredRowValue(row, [
      "will this initiative be new to this publication (ie. this initiative was not included in the previous grid)?"
    ])
  );
  const isNew = publicationStatus === "New" ? "Yes" : "No";
  const quarterSignals = buildStructuredQuarterSignals(row);
  const milestoneEvents = mergeTimelineEvents(
    buildStructuredQuarterEvents(quarterSignals),
    extractTimelineMarkers(expectedKeyMilestones)
  );
  const warnings = [];
  if (!expectedKeyMilestones && !quarterSignals.length) {
    warnings.push("Structured milestones are limited.");
  }

  const rawText = normaliseWs(
    [
      sectionName,
      subcategory,
      leadRegulator,
      title,
      description,
      expectedKeyMilestones,
      quarterSignals.map(signal => `${signal.label}: ${signal.status}`).join(" "),
      sourceLink
    ].join(" ")
  );

  return {
    id: `${slugify(title || "item")}-${rowNumber}`,
    canonicalKey: buildCanonicalKey({
      sectionName,
      leadRegulator,
      initiativeTitle: title
    }),
    sourceFile: fileName,
    sourceType: "Excel",
    sourceRowNumber: rowNumber,
    sourcePages: [],
    sectionName,
    subcategory,
    leadRegulator,
    initiativeTitle: title,
    initiativeDescription: description,
    expectedKeyMilestones,
    indicativeImpactOnFirms: impactFlag,
    consumerInterest,
    timingUpdated,
    isNew,
    publicationStatus,
    timingBucket: inferTimingBucket(
      `${expectedKeyMilestones} ${quarterSignals.map(signal => signal.label).join(" ")} ${description}`
    ),
    rawText,
    parseConfidence: 98,
    parseConfidenceBand: "High",
    parseWarnings: warnings,
    evidence: emptyEvidenceSet(),
    sourceLink,
    structuredTimeline: {
      quarterSignals,
      events: milestoneEvents
    }
  };
}

function getStructuredRowValue(row, aliases) {
  const aliasSet = aliases.map(normaliseLookupKey);
  for (const [key, value] of Object.entries(row || {})) {
    if (aliasSet.includes(normaliseLookupKey(key))) {
      return normaliseStructuredCell(value);
    }
  }
  return "";
}

function normaliseStructuredCell(value) {
  return String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function buildStructuredQuarterSignals(row) {
  return EXCEL_QUARTER_COLUMNS.map(column => {
    const value = getStructuredRowValue(row, column.aliases);
    const status = parseQuarterSignal(value);
    if (!status) return null;
    return {
      label: column.label,
      status,
      dateValue: new Date(column.year, column.month, 1).getTime()
    };
  }).filter(Boolean);
}

function buildStructuredQuarterEvents(signals) {
  return signals.map(signal => ({
    label: signal.label,
    dateValue: signal.dateValue,
    detail:
      signal.status === "Engagement + milestone"
        ? `${signal.label}: engagement and milestone activity expected.`
        : `${signal.label}: ${signal.status.toLowerCase()} expected.`
  }));
}

function parseQuarterSignal(value) {
  const text = normaliseLookupKey(value);
  if (!text) return "";
  if (text.includes("both")) return "Engagement + milestone";
  if (text.includes("key milestone")) return "Key milestone";
  if (text.includes("engagement")) return "Engagement";
  return normaliseWs(value);
}

function normalisePublicationStatus(value) {
  const text = normaliseLookupKey(value);
  if (text.includes("new")) return "New";
  if (text.includes("previously published")) return "Previously published";
  return normaliseWs(value) || "Previously published";
}

function consumerInterestFromStructured(value) {
  const text = normaliseLookupKey(value);
  if (text === "yes") return "H";
  if (text === "no") return "L";
  return "";
}

function normaliseRegulatorField(value) {
  const parts = normaliseWs(value)
    .toUpperCase()
    .split(/[\/,;|]+/)
    .map(part => part.replace(/\s+/g, ""))
    .filter(Boolean);
  return parts.join("/");
}

function canonicaliseSectionName(value) {
  const key = normaliseLookupKey(value);
  return SECTION_ALIASES.get(key) || normaliseWs(value);
}

function canonicaliseSubcategoryName(value) {
  const key = normaliseLookupKey(value);
  return SUBCATEGORY_ALIASES.get(key) || normaliseWs(value);
}

function mergeStructuredAndPdfItems(excelItems, pdfItems) {
  if (!excelItems.length) return pdfItems;
  if (!pdfItems.length) return excelItems;

  const usedPdfKeys = new Set();
  return excelItems.map(item => {
    const match = findSupportingPdfMatch(item, pdfItems, usedPdfKeys);
    if (!match) return item;
    usedPdfKeys.add(`${match.canonicalKey}|${match.id}`);
    return mergeStructuredItemWithPdf(item, match);
  });
}

function findSupportingPdfMatch(structuredItem, pdfItems, usedPdfKeys) {
  let bestMatch = null;
  let bestScore = 0;

  pdfItems.forEach(pdfItem => {
    const usedKey = `${pdfItem.canonicalKey}|${pdfItem.id}`;
    if (usedPdfKeys.has(usedKey)) return;

    let score = 0;
    if (pdfItem.canonicalKey === structuredItem.canonicalKey) score += 100;
    if (looseTitleKey(pdfItem.initiativeTitle) === looseTitleKey(structuredItem.initiativeTitle)) score += 40;
    if (pdfItem.sectionName === structuredItem.sectionName) score += 12;
    if (pdfItem.leadRegulator === structuredItem.leadRegulator) score += 12;
    if (
      looseTitleKey(pdfItem.initiativeTitle).includes(looseTitleKey(structuredItem.initiativeTitle)) ||
      looseTitleKey(structuredItem.initiativeTitle).includes(looseTitleKey(pdfItem.initiativeTitle))
    ) {
      score += 18;
    }

    if (score > bestScore) {
      bestScore = score;
      bestMatch = pdfItem;
    }
  });

  return bestScore >= 62 ? bestMatch : null;
}

function mergeStructuredItemWithPdf(structuredItem, pdfItem) {
  return {
    ...structuredItem,
    sourceType: "Excel + PDF",
    sourcePages: pdfItem.sourcePages || [],
    evidence: flattenEvidence(pdfItem.evidence || {}).length ? pdfItem.evidence : structuredItem.evidence,
    parseWarnings: dedupeStrings([
      ...(structuredItem.parseWarnings || []),
      ...(pdfItem.parseWarnings || [])
    ]),
    rawText: normaliseWs(`${structuredItem.rawText || ""} ${pdfItem.rawText || ""}`),
    initiativeDescription: structuredItem.initiativeDescription || pdfItem.initiativeDescription,
    expectedKeyMilestones: structuredItem.expectedKeyMilestones || pdfItem.expectedKeyMilestones,
    structuredTimeline: {
      quarterSignals: structuredItem.structuredTimeline?.quarterSignals || [],
      events: mergeTimelineEvents(
        structuredItem.structuredTimeline?.events || [],
        extractTimelineMarkers(
          `${structuredItem.expectedKeyMilestones || ""} ${pdfItem.expectedKeyMilestones || ""} ${pdfItem.initiativeDescription || ""}`
        )
      )
    }
  };
}

function mergeTimelineEvents(primaryEvents, secondaryEvents) {
  const merged = new Map();
  [...(primaryEvents || []), ...(secondaryEvents || [])].forEach(event => {
    if (!event?.label || !Number.isFinite(event.dateValue)) return;
    const detail = normaliseWs(event.detail || "");
    const key = `${event.label}|${event.dateValue}`;
    const existing = merged.get(key);
    if (!existing) {
      merged.set(key, {
        label: event.label,
        dateValue: event.dateValue,
        detail
      });
      return;
    }

    if (detail.length > (existing.detail || "").length) {
      merged.set(key, {
        label: event.label,
        dateValue: event.dateValue,
        detail
      });
    }
  });
  return [...merged.values()].sort((a, b) => a.dateValue - b.dateValue);
}

function emptyEvidenceSet() {
  return {
    title: [],
    lead: [],
    description: [],
    milestones: [],
    impact: [],
    consumer: [],
    timing: [],
    isNew: [],
    general: []
  };
}

function looseTitleKey(value) {
  return normaliseWs(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\b(the|and|for|of|to|on|in|a)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

async function parsePdfFile(buffer, fileName) {
  const loadingTask = pdfjsLib.getDocument({ data: clonePdfBinary(buffer) });
  const pdf = await loadingTask.promise;
  const rows = [];
  let activeTemplate = null;

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent();
    const pageRows = buildPdfRows(textContent.items, pageNumber);
    const template = detectColumnTemplate(pageRows, viewport.width);
    activeTemplate = template || activeTemplate || buildFallbackColumnTemplate(viewport.width);

    pageRows.forEach((row, rowIndex) => {
      row.rowId = `p${pageNumber}-r${rowIndex + 1}`;
      row.pageWidth = viewport.width;
      row.cells = assignRowCells(row, activeTemplate);
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
        pageNumber,
        y: token.y,
        tokens: [token]
      });
    }
  }

  return rawRows.map(row => {
    const segments = mergeRowTokens(row.tokens.sort((a, b) => a.x - b.x));
    return {
      pageNumber: row.pageNumber,
      y: row.y,
      segments,
      text: normaliseWs(segments.map(segment => segment.text).join(" ")),
      hasBold: segments.some(segment => /bold/i.test(segment.fontName || ""))
    };
  });
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
  Object.entries(DEFAULT_COLUMN_START_RATIOS).forEach(([key, ratio]) => {
    anchors[key] = pageWidth * ratio;
  });
  return buildColumnTemplate(anchors, pageWidth);
}

function buildColumnTemplate(anchorMap, pageWidth) {
  const keys = ["lead", "initiative", "milestones", "impact", "consumer", "timing", "isNew"];
  const starts = keys.map(key => ({
    key,
    start: anchorMap[key] != null ? anchorMap[key] : pageWidth * DEFAULT_COLUMN_START_RATIOS[key]
  }));

  const columns = [];
  for (let index = 0; index < starts.length; index += 1) {
    const previous = starts[index - 1];
    const current = starts[index];
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
  const keys = ["lead", "initiative", "milestones", "impact", "consumer", "timing", "isNew", "overflow"];
  const cells = Object.fromEntries(
    keys.map(key => [
      key,
      {
        text: "",
        x1: null,
        x2: null
      }
    ])
  );

  row.segments.forEach(segment => {
    const centre = segment.x + segment.width / 2;
    const column = template.find(item => centre >= item.start && centre < item.end);
    const key = column?.key || "overflow";
    const cell = cells[key];
    cell.text = joinInline(cell.text, segment.text);
    cell.x1 = cell.x1 == null ? segment.x : Math.min(cell.x1, segment.x);
    cell.x2 = cell.x2 == null ? segment.x + segment.width : Math.max(cell.x2, segment.x + segment.width);
  });

  Object.values(cells).forEach(cell => {
    cell.text = normaliseWs(cell.text);
  });

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

      current = createAccumulator(start, currentSection, currentSubcategory, fileName, row);
      ingestRowIntoAccumulator(current, row, { isStartRow: true });

      if (start.consumeNextRow && nextRow) {
        ingestRowIntoAccumulator(current, nextRow, {
          consumeAsTitle: true
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
  if (isQuarterHeaderText(text)) return true;

  const headerHits = HEADER_TEXT_PATTERNS.reduce((count, pattern) => {
    return count + (pattern.test(text) ? 1 : 0);
  }, 0);

  return headerHits >= 3;
}

function detectInitiativeStart(row, nextRow) {
  const direct = extractLeadAndTitle(row);
  if (direct) return direct;

  const leadOnly = normaliseLeadToken(cellText(row, "lead") || row.text);
  if (
    leadOnly &&
    nextRow &&
    !isNoiseRow(nextRow) &&
    !SECTION_LOOKUP.has(nextRow.text.toLowerCase()) &&
    !SUBCATEGORY_LOOKUP.has(nextRow.text.toLowerCase())
  ) {
    const nextParts = extractInitiativeCellParts(nextRow);
    const nextTitle = cleanTitleCandidate(
      nextParts.boldText || cellText(nextRow, "initiative") || nextRow.text
    );
    if (
      isPlausibleTitle(nextTitle) &&
      !extractLeadAndTitle(nextRow) &&
      isLikelyTitleContinuationText(nextTitle, nextRow)
    ) {
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
  const directLead = normaliseLeadToken(cellText(row, "lead"));
  const parts = extractInitiativeCellParts(row);
  const directTitle = cleanTitleCandidate(parts.boldText || cellText(row, "initiative"));

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

  const embedded = splitLeadAndTitle(parts.boldText || cellText(row, "initiative"));
  if (embedded && isPlausibleTitle(embedded.title)) {
    return {
      lead: embedded.lead,
      title: embedded.title,
      consumeNextRow: false
    };
  }

  return null;
}

function createAccumulator(start, sectionName, subcategory, fileName, row) {
  const accumulator = {
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
    rawParts: [],
    pageNumbers: [],
    titleLineCount: 1,
    evidence: {
      title: [],
      lead: [],
      description: [],
      milestones: [],
      impact: [],
      consumer: [],
      timing: [],
      isNew: [],
      general: []
    }
  };

  recordEvidence(accumulator, "lead", row, "lead", start.lead);
  recordEvidence(accumulator, "title", row, "initiative", start.title);
  return accumulator;
}

function ingestRowIntoAccumulator(accumulator, row, options) {
  const opts = options || {};

  appendUnique(accumulator.pageNumbers, row.pageNumber);
  appendUnique(accumulator.rawParts, row.text);
  recordEvidence(accumulator, "general", row, "overflow", row.text);

  if (cellText(row, "milestones")) {
    appendUnique(accumulator.milestoneParts, cellText(row, "milestones"));
    recordEvidence(accumulator, "milestones", row, "milestones");
  }

  if (cellText(row, "impact")) {
    appendUnique(accumulator.impactParts, cellText(row, "impact"));
    recordEvidence(accumulator, "impact", row, "impact");
  }

  if (cellText(row, "consumer")) {
    appendUnique(accumulator.consumerParts, cellText(row, "consumer"));
    recordEvidence(accumulator, "consumer", row, "consumer");
  }

  if (cellText(row, "timing")) {
    appendUnique(accumulator.timingParts, cellText(row, "timing"));
    recordEvidence(accumulator, "timing", row, "timing");
  }

  if (cellText(row, "isNew")) {
    appendUnique(accumulator.isNewParts, cellText(row, "isNew"));
    recordEvidence(accumulator, "isNew", row, "isNew");
  }

  const initiativeParts = extractInitiativeCellParts(row);
  const initiativeCell = cleanStructuredText(
    stripRepeatedTitle(initiativeParts.fullText || cellText(row, "initiative"), accumulator.initiativeTitle)
  );
  const titleCandidate = cleanStructuredText(
    stripRepeatedTitle(initiativeParts.boldText, accumulator.initiativeTitle)
  );
  const initiativeBodyText = cleanStructuredText(
    stripRepeatedTitle(
      initiativeParts.bodyText || stripRepeatedTitle(initiativeCell, titleCandidate),
      accumulator.initiativeTitle
    )
  );
  const overflowText = cleanStructuredText(cellText(row, "overflow"));
  const leadResidue = cleanStructuredText(
    stripLeadPrefix(cellText(row, "lead"), accumulator.leadRegulator)
  );

  if (opts.isStartRow) {
    if (leadResidue) {
      appendUnique(accumulator.initiativeDescriptionParts, leadResidue);
      recordEvidence(accumulator, "description", row, "lead", leadResidue);
    }
    if (initiativeBodyText) {
      appendUnique(accumulator.initiativeDescriptionParts, initiativeBodyText);
      recordEvidence(accumulator, "description", row, "initiative", initiativeBodyText);
    } else if (initiativeCell && !shouldTreatAsTitleContinuation(accumulator, initiativeCell, row)) {
      appendUnique(accumulator.initiativeDescriptionParts, initiativeCell);
      recordEvidence(accumulator, "description", row, "initiative", initiativeCell);
    }
  } else if (opts.consumeAsTitle && (titleCandidate || initiativeCell)) {
    const titleText = titleCandidate || initiativeCell;
    accumulator.initiativeTitle = normaliseWs(`${accumulator.initiativeTitle} ${titleText}`);
    accumulator.titleLineCount += 1;
    recordEvidence(accumulator, "title", row, "initiative", titleText);
    if (initiativeBodyText) {
      appendUnique(accumulator.initiativeDescriptionParts, initiativeBodyText);
      recordEvidence(accumulator, "description", row, "initiative", initiativeBodyText);
    }
  } else if (titleCandidate || initiativeCell) {
    const continuationText = titleCandidate || (row.hasBold ? initiativeCell : "");
    if (continuationText && shouldTreatAsTitleContinuation(accumulator, continuationText, row)) {
      accumulator.initiativeTitle = normaliseWs(`${accumulator.initiativeTitle} ${continuationText}`);
      accumulator.titleLineCount += 1;
      recordEvidence(accumulator, "title", row, "initiative", continuationText);
      if (initiativeBodyText) {
        appendUnique(accumulator.initiativeDescriptionParts, initiativeBodyText);
        recordEvidence(accumulator, "description", row, "initiative", initiativeBodyText);
      }
    } else {
      const detailText = initiativeBodyText || initiativeCell;
      appendUnique(accumulator.initiativeDescriptionParts, detailText);
      recordEvidence(accumulator, "description", row, "initiative", detailText);
    }
  }

  if (overflowText) {
    appendUnique(accumulator.initiativeDescriptionParts, overflowText);
    recordEvidence(accumulator, "description", row, "overflow", overflowText);
  }
}

function recordEvidence(accumulator, field, row, cellKey, overrideText) {
  const text = normaliseWs(overrideText || cellText(row, cellKey) || row.text);
  if (!text) return;

  const source = row.cells?.[cellKey] || {};
  const x1 = source.x1 == null ? 0 : source.x1;
  const x2 = source.x2 == null ? row.pageWidth : source.x2;
  const entry = {
    key: `${field}:${row.rowId}:${slugify(text).slice(0, 24)}`,
    field,
    pageNumber: row.pageNumber,
    rowId: row.rowId,
    y: row.y,
    x1,
    x2,
    pageWidth: row.pageWidth,
    excerpt: text,
    rowText: row.text
  };

  pushUniqueEvidence(accumulator.evidence[field], entry);
}

function pushUniqueEvidence(target, entry) {
  if (!target.find(item => item.key === entry.key)) {
    target.push(entry);
  }
}

function finalisePdfInitiative(accumulator, index, fileName) {
  const raw = normaliseWs(accumulator.rawParts.join(" "));
  const title = cleanTitleCandidate(accumulator.initiativeTitle || "");
  const description = normaliseInitiativeDescription(
    dedupeStrings(accumulator.initiativeDescriptionParts).join(" "),
    title
  );
  const milestones = dedupeStrings(
    [
      ...accumulator.milestoneParts,
      extractMilestones(raw)
    ].filter(Boolean)
  )
    .slice(0, 5)
    .join(" | ");
  const impactFlag = extractImpactFlag(accumulator.impactParts.join(" ") || raw);
  const consumerInterest = normaliseTrafficLight(accumulator.consumerParts.join(" "));
  const timingUpdated = normaliseYesNo(accumulator.timingParts.join(" "));
  const isNew =
    normaliseYesNo(accumulator.isNewParts.join(" ")) || (/\bnew\b/i.test(raw) ? "Yes" : "No");
  const parseAssessment = assessParseQuality(accumulator, {
    description,
    milestones,
    impactFlag
  });

  return {
    id: `${slugify(title || "item")}-${index}`,
    canonicalKey: buildCanonicalKey({
      ...accumulator,
      initiativeTitle: title
    }),
    sourceFile: fileName,
    sourcePages: accumulator.pageNumbers.slice().sort((a, b) => a - b),
    sectionName: accumulator.sectionName || "",
    subcategory: accumulator.subcategory || "",
    leadRegulator: accumulator.leadRegulator || "",
    initiativeTitle: title || "",
    initiativeDescription: description || raw,
    expectedKeyMilestones: milestones,
    indicativeImpactOnFirms: impactFlag,
    consumerInterest,
    timingUpdated,
    isNew,
    timingBucket: inferTimingBucket(`${milestones} ${description} ${raw}`),
    rawText: raw,
    parseConfidence: parseAssessment.score,
    parseConfidenceBand: parseAssessment.band,
    parseWarnings: parseAssessment.warnings,
    evidence: accumulator.evidence
  };
}

function assessParseQuality(accumulator, context) {
  let score = 42;
  const warnings = [];

  if (accumulator.sectionName) score += 10;
  else warnings.push("Section not cleanly extracted.");

  if (accumulator.subcategory) score += 4;
  if (isLeadToken(accumulator.leadRegulator)) score += 12;
  else warnings.push("Lead regulator not cleanly extracted.");

  if (accumulator.initiativeTitle && accumulator.initiativeTitle.length >= 12) score += 16;
  else warnings.push("Title may be incomplete.");

  if (context.description && context.description.length >= 40) score += 10;
  else warnings.push("Description is thin.");

  if (context.milestones) score += 8;
  else warnings.push("Milestones not clearly separated.");

  if (context.impactFlag) score += 6;
  if (accumulator.pageNumbers.length) score += 4;

  score = clamp(score, 0, 100);

  return {
    score,
    band: score >= 85 ? "High" : score >= 70 ? "Medium" : "Low",
    warnings: dedupeStrings(warnings)
  };
}

function dedupeItems(items) {
  const deduped = new Map();

  items.forEach(item => {
    const key = item.canonicalKey || buildCanonicalKey(item);
    const existing = deduped.get(key);
    if (!existing) {
      deduped.set(key, item);
      return;
    }

    deduped.set(key, mergeDuplicateItems(existing, item));
  });

  return [...deduped.values()];
}

function mergeDuplicateItems(left, right) {
  const preferred = sourcePriority(right) > sourcePriority(left) ? right : left;
  const secondary = preferred === left ? right : left;

  return {
    ...preferred,
    sourceType:
      preferred.sourceType === secondary.sourceType
        ? preferred.sourceType
        : `${preferred.sourceType || "Source"} + ${secondary.sourceType || "Source"}`,
    initiativeDescription:
      preferred.initiativeDescription?.length >= secondary.initiativeDescription?.length
        ? preferred.initiativeDescription
        : secondary.initiativeDescription,
    expectedKeyMilestones:
      preferred.expectedKeyMilestones || secondary.expectedKeyMilestones,
    sourcePages: uniqueNumbers([...(preferred.sourcePages || []), ...(secondary.sourcePages || [])]),
    evidence:
      flattenEvidence(preferred.evidence || {}).length >= flattenEvidence(secondary.evidence || {}).length
        ? preferred.evidence
        : secondary.evidence,
    parseWarnings: dedupeStrings([
      ...(preferred.parseWarnings || []),
      ...(secondary.parseWarnings || [])
    ]),
    rawText: normaliseWs(`${preferred.rawText || ""} ${secondary.rawText || ""}`),
    structuredTimeline: {
      quarterSignals: [
        ...((preferred.structuredTimeline?.quarterSignals) || []),
        ...((secondary.structuredTimeline?.quarterSignals) || [])
      ],
      events: mergeTimelineEvents(
        preferred.structuredTimeline?.events || [],
        secondary.structuredTimeline?.events || []
      )
    }
  };
}

function sourcePriority(item) {
  if ((item.sourceType || "").includes("Excel")) return 3;
  if (item.sourceType === "PDF") return 1;
  return 0;
}

function uniqueNumbers(values) {
  return [...new Set((values || []).filter(value => Number.isFinite(value)))].sort((a, b) => a - b);
}

function analyseRows(items) {
  return items.map(item => {
    const classification = classifyItem(item);
    const stage = determineStage(item);
    const obligations = mapObligations(item, classification, stage);
    const timeline = buildTimelinePoint(item);
    const ownership = determineOwnership(item, classification, stage, obligations);
    const impact = determineImpactLevel(item, classification, stage, obligations);
    const relevance = evaluateFmrukRelevance(item, classification, stage, obligations);
    const uncertainty = determineUncertainty(item, classification, stage, relevance);
    const priority = determinePriority(
      item,
      classification,
      stage,
      obligations,
      ownership,
      impact,
      relevance,
      uncertainty,
      timeline
    );
    const suggestedActions = buildSuggestedActions(
      item,
      classification,
      stage,
      ownership,
      impact,
      relevance,
      obligations,
      uncertainty
    );
    const potentialBusinessImpact = buildPotentialBusinessImpact(
      item,
      classification,
      stage,
      obligations
    );
    const clusterLabel = determineCluster(item, classification, obligations);
    const fmrukReadThrough = buildFmrukReadThrough(
      item,
      classification,
      stage,
      obligations,
      ownership,
      relevance
    );

    let analysed = {
      ...item,
      theme: classification.theme,
      internalSubTheme: classification.subTheme,
      primaryOwner: ownership.primaryOwner,
      secondaryOwner: ownership.secondaryOwner,
      classificationConfidence: classification.confidence,
      classificationSignals: classification.signals,
      classificationAmbiguity: classification.ambiguity,
      stage: stage.id,
      stageLabel: stage.label,
      stageConfidence: stage.confidence,
      stageSignals: stage.signals,
      obligations,
      impactLevel: impact.level,
      impactScore: impact.score,
      relevanceScore: relevance.score,
      relevanceBand: relevance.band,
      isFmrukRelevant: relevance.isRelevant,
      fmrukEffectLevel: relevance.effectLevel,
      relevanceSignals: relevance.positiveReasons,
      relevanceNegativeSignals: relevance.negativeReasons,
      profileReasons: relevance.profileReasons,
      whyNotRelevant: relevance.whyNotRelevant,
      uncertaintyFlags: uncertainty.flags,
      needsReview: uncertainty.needsReview,
      priorityScoreBase: priority.score,
      priorityScore: priority.score,
      priorityBand: priority.band,
      priorityReasons: priority.reasons,
      urgencyScore: priority.urgencyScore,
      deliveryScore: priority.deliveryScore,
      potentialBusinessImpact,
      fmrukReadThrough,
      suggestedActions,
      suggestedAction: suggestedActions.join(" "),
      immediateActionRequired: priority.immediateActionRequired,
      rationale: buildRationale(item, classification, stage, relevance, obligations),
      timeline,
      clusterLabel
    };

    analysed = applyFeedbackOverrides(analysed);
    return analysed;
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
      confidence: clamp(46 + best.score, 0, 100),
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
      confidence: clamp(46 + best.score, 0, 100),
      signals: best.reasons
    };
  }

  return {
    id: "monitoring",
    label: "Monitoring",
    confidence: 50,
    signals: ["No strong stage marker detected."]
  };
}

function mapObligations(item, classification, stage) {
  const blob = buildBlob(item);
  const scored = OBLIGATION_RULES.map(rule => {
    const result = scoreSignalGroups(rule.signals, blob);
    let score = result.score;

    if (classification.subTheme === "Transaction Reporting" && rule.name === "Reporting & MI") {
      score += 10;
    }
    if (classification.subTheme === "Market Abuse / Surveillance" && rule.name === "Market Conduct & Surveillance") {
      score += 10;
    }
    if (classification.subTheme === "Outsourcing / Third Party Risk" && rule.name === "Outsourcing & Third Parties") {
      score += 10;
    }
    if (classification.subTheme === "Operational Resilience" && rule.name === "Operational Resilience") {
      score += 10;
    }
    if (classification.subTheme === "Capital / Liquidity" && rule.name === "Prudential & Capital") {
      score += 10;
    }
    if (classification.subTheme === "Settlement / T+1" && rule.name === "Operating Model") {
      score += 10;
    }
    if (classification.subTheme === "Settlement / T+1" && rule.name === "Reporting & MI") {
      score += 6;
    }
    if (stage.id === "reporting" && rule.name === "Reporting & MI") {
      score += 10;
    }
    if (stage.id === "legislation" && rule.name === "Legal Entity Perimeter") {
      score += 10;
    }
    if (stage.id === "final_rules" && rule.name === "Policy & Procedures") {
      score += 8;
    }

    return {
      name: rule.name,
      score,
      reasons: result.reasons
    };
  })
    .filter(itemScore => itemScore.score >= 10)
    .sort((a, b) => b.score - a.score)
    .slice(0, 5);

  if (!scored.length) {
    return [
      {
        name: "Policy & Procedures",
        reason: "Default control domain."
      }
    ];
  }

  return scored.map(entry => ({
    name: entry.name,
    reason: entry.reasons[0] || `${entry.name} is in scope.`
  }));
}

function determineImpactLevel(item, classification, stage, obligations) {
  let score = 30;
  const impactFlag = String(item.indicativeImpactOnFirms || "").trim().toUpperCase();
  if (impactFlag === "H" || impactFlag === "HIGH") score += 40;
  if (impactFlag === "U" || impactFlag === "M" || impactFlag === "MEDIUM") score += 22;
  if (impactFlag === "L" || impactFlag === "LOW") score += 8;

  score += classification.impactBias || 0;
  if (stage.id === "final_rules") score += 12;
  if (stage.id === "reporting") score += 10;
  if (stage.id === "supervisory") score += 8;
  if (item.timingBucket === "Near Term") score += 8;
  if (item.timingBucket === "Longer Term") score -= 4;
  if (item.consumerInterest === "H") score += 4;
  if (obligations.some(obligation => obligation.name === "Reporting & MI")) score += 5;
  if (obligations.some(obligation => obligation.name === "Operational Resilience")) score += 5;

  score = clamp(score, 0, 100);

  return {
    score,
    level: score >= 78 ? "High" : score >= 55 ? "Medium" : "Low"
  };
}

function evaluateFmrukRelevance(item, classification, stage, obligations) {
  const blob = buildBlob(item);
  const effectMode = getEffectModeConfig();
  let score = 28 + effectMode.profileBias;
  const positive = scoreSignalGroups(RELEVANCE_SIGNAL_GROUPS.positive, blob);
  const negative = scoreSignalGroups(RELEVANCE_SIGNAL_GROUPS.negative, blob);
  const profileMatch = scoreProfileMatch(item, state.profile);
  const mode = getPriorityModeConfig();

  score += positive.score * effectMode.positiveMultiplier;
  score -= negative.score * effectMode.negativeMultiplier;
  score += profileMatch.score * effectMode.profileMultiplier;
  score += RELEVANCE_SIGNAL_GROUPS.sectionAdjustments[item.sectionName] || 0;
  if (mode.themeBoosts.includes(classification.theme)) score += 6;
  if (obligations.some(obligation => mode.obligationBoosts.includes(obligation.name))) score += 4;

  if (
    [
      "Transaction Reporting",
      "Best Execution",
      "Market Abuse / Surveillance",
      "Capital / Liquidity",
      "Operational Resilience",
      "Outsourcing / Third Party Risk",
      "SMCR / Governance",
      "Regulatory Reporting / Data Collection",
      "Settlement / T+1"
    ].includes(classification.subTheme)
  ) {
    score += 8;
  }

  if (obligations.some(obligation => obligation.name === "Reporting & MI")) score += 4;
  if (obligations.some(obligation => obligation.name === "Governance & Accountability")) score += 4;
  if (stage.id === "reporting" || stage.id === "final_rules") score += 4;
  if (item.leadRegulator.includes("FCA")) score += 3;
  if ((item.leadRegulator.match(/\//g) || []).length >= 1) score += 2;
  if (
    ["Investment management", "Wholesale financial markets", "Multi-sector"].includes(item.sectionName)
  ) {
    score += 4;
  }
  if (
    ["Banking, credit and lending", "Insurance and reinsurance", "Pensions and retirement income"].includes(item.sectionName) &&
    profileMatch.score < 10
  ) {
    score -= 8;
  }
  if (item.indicativeImpactOnFirms === "H" && profileMatch.score >= 8) score += 4;
  if (item.sectionName === "Annex: initiatives completed/stopped") score -= 10;

  score = clamp(score, 0, 100);

  const positiveReasons = dedupeStrings([
    ...positive.reasons,
    ...profileMatch.reasons,
    ...(RELEVANCE_SIGNAL_GROUPS.sectionAdjustments[item.sectionName] > 0
      ? [`Section weighting: ${item.sectionName}.`]
      : [])
  ]);

  const negativeReasons = dedupeStrings([
    ...negative.reasons,
    ...(RELEVANCE_SIGNAL_GROUPS.sectionAdjustments[item.sectionName] < 0
      ? [`Section weighting reduces relevance: ${item.sectionName}.`]
      : [])
  ]);

  const isRelevant =
    score >= effectMode.threshold ||
    (
      score >= effectMode.watchlistThreshold &&
      (profileMatch.reasons.length > 0 || positiveReasons.length >= 2) &&
      negativeReasons.length < 2
    );
  const band = score >= 80 ? "High" : score >= 58 ? "Medium" : "Low";
  const effectLevel = !isRelevant
    ? score >= effectMode.watchlistThreshold
      ? "Watchlist"
      : "Out of scope"
    : score >= 82
      ? "Direct"
      : score >= 68
        ? "Material"
        : "Conditional";
  const whyNotRelevant =
    effectLevel === "Out of scope"
      ? `Current view: no material FMRUK effect. ${negativeReasons.join(" ") || "The initiative currently looks outside the UK entity scope."}`
      : effectLevel === "Watchlist"
        ? `Current view: keep on the watchlist until scope or milestones become clearer. ${negativeReasons.join(" ") || positiveReasons.join(" ")}`
        : "";

  return {
    score,
    band,
    isRelevant,
    effectLevel,
    positiveReasons,
    negativeReasons,
    profileReasons: profileMatch.reasons,
    whyNotRelevant
  };
}

function scoreProfileMatch(item, profile) {
  const blob = buildBlob(item);
  const categories = [
    ["businessLines", 16, "Matches business line"],
    ["permissions", 14, "Matches permission"],
    ["products", 12, "Matches product scope"],
    ["distributionModel", 10, "Matches distribution model"],
    ["outsourcedProviders", 10, "Matches outsourced provider exposure"],
    ["reportingObligations", 14, "Matches reporting obligation"],
    ["committees", 8, "Matches committee touchpoint"],
    ["regulatoryTouchpoints", 12, "Matches regulatory touchpoint"]
  ];

  let score = 0;
  const reasons = [];

  categories.forEach(([field, weight, label]) => {
    const entries = profile?.[field] || [];
    const matches = entries.filter(entry => blob.includes(entry.toLowerCase()));
    if (!matches.length) return;

    score += Math.min(weight, 8 + matches.length * 4);
    reasons.push(`${label}: ${matches.slice(0, 2).join(", ")}.`);
  });

  return {
    score,
    reasons
  };
}

function determineOwnership(item, classification, stage, obligations) {
  const blob = buildBlob(item);
  const mode = getPriorityModeConfig();
  const ownershipMode = getOwnershipModeConfig();
  const hasObligation = name =>
    obligations.some(obligation => obligation.name === name);
  let primaryOwner = classification.primaryOwner;
  let secondaryOwner = classification.secondaryOwner;

  if (classification.subTheme === "SMCR / Governance") {
    primaryOwner = "Compliance";
    secondaryOwner = "HR";
  } else if (item.leadRegulator.includes("HMT") && stage.id === "legislation") {
    primaryOwner = "Legal";
    secondaryOwner = "Compliance";
  } else if (classification.subTheme === "Settlement / T+1") {
    primaryOwner = "Operations";
    secondaryOwner = "Compliance";
  } else if (
    classification.subTheme === "Regulatory Reporting / Data Collection" ||
    stage.id === "reporting"
  ) {
    primaryOwner = "Operations";
    secondaryOwner = hasObligation("Governance & Accountability") ? "Compliance" : "Compliance";
  } else if (classification.subTheme === "Capital / Liquidity") {
    primaryOwner = "Finance";
    secondaryOwner = "Risk";
  } else if (classification.subTheme === "Outsourcing / Third Party Risk") {
    primaryOwner = "Risk";
    secondaryOwner = "Technology";
  } else if (/\bt\+1\b|\bt plus one\b|\bsettlement\b|\bpost-trade\b/.test(blob)) {
    primaryOwner = "Operations";
    secondaryOwner = "Compliance";
  } else if (
    hasObligation("Reporting & MI") &&
    (mode.ownerBoosts.includes("Operations") || mode.obligationBoosts.includes("Reporting & MI"))
  ) {
    primaryOwner = "Operations";
    secondaryOwner = primaryOwner === "Compliance" ? "Risk" : "Compliance";
  } else if (
    hasObligation("Governance & Accountability") &&
    mode.obligationBoosts.includes("Governance & Accountability")
  ) {
    primaryOwner = "Compliance";
    secondaryOwner = secondaryOwner === "Compliance" ? "Legal" : secondaryOwner;
  } else if (
    item.sectionName === "Wholesale financial markets" &&
    (hasObligation("Operating Model") || hasObligation("Reporting & MI"))
  ) {
    primaryOwner = "Operations";
    secondaryOwner = "Compliance";
  } else if (
    item.sectionName === "Investment management" &&
    classification.theme.includes("Investment Management")
  ) {
    primaryOwner = "Compliance";
    secondaryOwner = "Legal";
  }

  if (ownershipMode.biasOwners.includes("Compliance")) {
    if (
      hasObligation("Governance & Accountability") ||
      hasObligation("Policy & Procedures") ||
      classification.subTheme === "SMCR / Governance" ||
      classification.subTheme === "Consumer Duty"
    ) {
      primaryOwner = "Compliance";
      secondaryOwner = classification.subTheme === "SMCR / Governance" ? "HR" : (secondaryOwner || "Legal");
    }
  }

  if (ownershipMode.biasOwners.includes("Operations")) {
    if (
      hasObligation("Reporting & MI") ||
      hasObligation("Operating Model") ||
      classification.subTheme === "Settlement / T+1" ||
      classification.subTheme === "Regulatory Reporting / Data Collection"
    ) {
      primaryOwner = "Operations";
      secondaryOwner = hasObligation("Operating Model") ? "Technology" : "Compliance";
    }
  }

  if (primaryOwner === secondaryOwner) {
    secondaryOwner = classification.secondaryOwner !== primaryOwner
      ? classification.secondaryOwner
      : "Compliance";
  }

  return {
    primaryOwner,
    secondaryOwner
  };
}

function determinePriority(
  item,
  classification,
  stage,
  obligations,
  ownership,
  impact,
  relevance,
  uncertainty,
  timeline
) {
  const mode = getPriorityModeConfig();
  const weights = mode.weights;
  const urgencyScore = calculateUrgencyScore(item, stage, timeline);
  const deliveryScore = calculateDeliveryScore(classification, obligations, ownership, mode);

  let score =
    relevance.score * weights.relevance +
    impact.score * weights.impact +
    urgencyScore * weights.urgency +
    deliveryScore * weights.delivery;

  if (mode.themeBoosts.includes(classification.theme)) score += 6;
  if (mode.ownerBoosts.includes(ownership.primaryOwner)) score += 4;
  if ((obligations || []).some(obligation => mode.obligationBoosts.includes(obligation.name))) score += 4;
  if (relevance.effectLevel === "Direct" || relevance.effectLevel === "Material") score += 5;
  if (relevance.effectLevel === "Out of scope") score -= 18;
  if (relevance.effectLevel === "Watchlist") score -= 8;
  if ((item.structuredTimeline?.quarterSignals || []).length >= 3) score += 3;
  if ((item.leadRegulator.match(/\//g) || []).length >= 1) score += 2;
  if (item.timingUpdated === "Yes") score += 3;
  if (item.isNew === "Yes") score += 2;

  const uncertaintyPenalty = Math.min(10, (uncertainty.flags || []).length * 1.5);
  score -= uncertaintyPenalty;
  score = clamp(score, 0, 100);

  if (relevance.effectLevel === "Out of scope") {
    score = Math.min(score, 38);
  }

  const band = priorityBandFromScore(score);
  const reasons = [];
  if (relevance.score >= 80) reasons.push("High FMRUK relevance.");
  if (impact.level === "High") reasons.push("High firm impact.");
  if (urgencyScore >= 72) reasons.push("Near-term milestone or implementation date.");
  if (deliveryScore >= 68) reasons.push("Material delivery or control change.");
  if (relevance.effectLevel === "Out of scope") reasons.push("Currently out of FMRUK scope.");
  if (relevance.effectLevel === "Watchlist") reasons.push("Keep on the watchlist pending clearer scope.");
  if (mode.themeBoosts.includes(classification.theme) || mode.ownerBoosts.includes(ownership.primaryOwner)) {
    reasons.push(`Profile weighting: ${mode.label.toLowerCase()}.`);
  }
  if (!reasons.length) reasons.push("Included in the working review queue.");

  const immediateActionRequired = Boolean(
    relevance.isRelevant &&
      relevance.effectLevel !== "Conditional" &&
      stage.id !== "monitoring" &&
      (
        score >= 82 ||
        (impact.level === "High" && urgencyScore >= 72) ||
        (stage.id === "reporting" && relevance.score >= 65 && urgencyScore >= 68)
      )
  );

  return {
    score,
    band,
    reasons,
    urgencyScore,
    deliveryScore,
    immediateActionRequired
  };
}

function calculateUrgencyScore(item, stage, timeline) {
  let score = 28;
  if (stage.id === "final_rules") score += 24;
  if (stage.id === "reporting") score += 22;
  if (stage.id === "supervisory") score += 16;
  if (stage.id === "legislation") score += 14;
  if (stage.id === "consultation") score += 12;
  if (stage.id === "monitoring") score -= 10;

  if (item.timingBucket === "Near Term") score += 20;
  else if (item.timingBucket === "Medium Term") score += 12;
  else if (item.timingBucket === "Longer Term") score += 4;
  else score += 8;

  if (item.timingUpdated === "Yes") score += 8;
  if (item.isNew === "Yes") score += 6;
  if (item.consumerInterest === "H") score += 4;
  if ((item.structuredTimeline?.quarterSignals || []).length >= 3) score += 4;
  if ((timeline?.events || []).length >= 3) score += 6;
  if (/\b(goes live|implementation|effective|opens|final policy|policy statement)\b/i.test(item.expectedKeyMilestones || "")) {
    score += 6;
  }

  const now = Date.now();
  if (Number.isFinite(timeline?.sortValue)) {
    const days = Math.round((timeline.sortValue - now) / (24 * 60 * 60 * 1000));
    if (days <= 90) score += 10;
    else if (days <= 180) score += 6;
    else if (days >= 365) score -= 4;
  }

  return clamp(score, 0, 100);
}

function calculateDeliveryScore(classification, obligations, ownership, mode) {
  let score = 34;
  const obligationWeights = {
    "Reporting & MI": 18,
    "Operational Resilience": 18,
    "Governance & Accountability": 16,
    "Outsourcing & Third Parties": 14,
    "Prudential & Capital": 14,
    "Market Conduct & Surveillance": 12,
    "Policy & Procedures": 10,
    "Operating Model": 10,
    "Disclosure & Communications": 10
  };

  (obligations || []).slice(0, 3).forEach(obligation => {
    score += obligationWeights[obligation.name] || 8;
  });

  if (mode.ownerBoosts.includes(ownership.primaryOwner)) score += 8;
  if (mode.themeBoosts.includes(classification.theme)) score += 6;
  if ((ownership.primaryOwner || "") === "Operations" && classification.subTheme === "Settlement / T+1") score += 8;
  if ((ownership.primaryOwner || "") === "Compliance" && classification.subTheme === "SMCR / Governance") score += 6;

  return clamp(score, 0, 100);
}

function priorityBandFromScore(score) {
  if (score >= 86) return "Critical";
  if (score >= 72) return "High";
  if (score >= 58) return "Medium";
  return "Low";
}

function buildFmrukReadThrough(item, classification, stage, obligations, ownership, relevance) {
  if (relevance.effectLevel === "Out of scope") {
    return relevance.whyNotRelevant || "Current view: this initiative does not appear to have a material effect on FMRUK.";
  }

  if (relevance.effectLevel === "Watchlist") {
    return item.whyNotRelevant || "Current review suggests limited direct impact on FMRUK. Keep on the watchlist until the scope or timing becomes clearer.";
  }

  const domainText = obligations.length
    ? obligations.slice(0, 2).map(obligation => obligation.name.toLowerCase()).join(" and ")
    : "the current control environment";

  const ownerText = ownership.secondaryOwner
    ? `${ownership.primaryOwner} as lead, with ${ownership.secondaryOwner} support`
    : `${ownership.primaryOwner} as lead`;

  const stageText = {
    consultation: "The immediate requirement is to confirm scope early and decide whether a response or preparatory work is needed.",
    final_rules: "The item is close to implementation, so the focus should move into delivery readiness and control change.",
    supervisory: "The main read-through is evidencing current-state compliance and addressing weaker controls before scrutiny increases.",
    reporting: "The practical effect is likely to sit in data ownership, submissions, reconciliations and sign-off.",
    legislation: "The key issue is whether entity scope, permissions or legal interpretation change the obligations that apply to FMRUK.",
    monitoring: "The item should remain visible, but it does not yet look like an immediate delivery programme."
  }[stage.id] || "";

  const specificReadThroughs = {
    "Settlement / T+1":
      "For FMRUK, the practical effect is likely to sit in trade allocations, confirmations, affirmations, settlement cut-offs, exception handling and oversight of post-trade providers.",
    "Transaction Reporting":
      "For FMRUK, the main effect is likely to sit in reporting scope, data fields, reconciliations, exception management and control ownership across Operations and Compliance.",
    "SMCR / Governance":
      "For FMRUK, the main effect is likely to sit in accountability mapping, committee oversight, statements of responsibility and coordination with HR where certification or role scope changes arise.",
    "Regulatory Reporting / Data Collection":
      "For FMRUK, the main effect is likely to sit in reporting perimeter, data lineage, controls, sign-off and readiness to make accurate submissions.",
    "Outsourcing / Third Party Risk":
      "For FMRUK, the main effect is likely to sit in outsourced service oversight, registers, due diligence, contractual controls and dependency mapping.",
    "Operational Resilience":
      "For FMRUK, the main effect is likely to sit in important business services, scenario testing, dependency mapping and remediation governance."
  };

  const opening =
    specificReadThroughs[classification.subTheme] ||
    `For FMRUK, the main read-through is into ${domainText}. ${classification.impactStatement}`;

  const effectPrefix =
    relevance.effectLevel === "Direct"
      ? "This looks like a direct FMRUK item."
      : relevance.effectLevel === "Material"
        ? "This looks likely to affect FMRUK."
        : "This may affect FMRUK if the final scope lands on the UK entity.";

  return `${effectPrefix} ${opening} Ownership is likely to sit with ${ownerText}. ${stageText}`;
}

function determineUncertainty(item, classification, stage, relevance) {
  const flags = [];
  if ((item.parseConfidence || 0) < 70) flags.push("Low parse confidence.");
  if ((classification.confidence || 0) < 60) flags.push("Moderate classification confidence.");
  if (classification.ambiguity) flags.push(classification.ambiguity);
  if ((stage.confidence || 0) < 58) flags.push("Stage uncertain.");
  if (!item.expectedKeyMilestones) flags.push("No clear milestones extracted.");
  if (relevance.score >= 45 && relevance.score <= 65) {
    flags.push("Relevance needs judgement.");
  }
  if (relevance.positiveReasons.length && relevance.negativeReasons.length) {
    flags.push("Mixed relevance signals.");
  }
  return {
    flags: dedupeStrings([...flags, ...(item.parseWarnings || [])]),
    needsReview: flags.length > 0 || (item.parseWarnings || []).length > 0
  };
}

function buildPotentialBusinessImpact(item, classification, stage, obligations) {
  const obligationNames = obligations.map(obligation => obligation.name).slice(0, 3);
  const obligationText = obligationNames.length
    ? ` Control domains: ${obligationNames.join(", ")}.`
    : "";
  const stageNote = {
    consultation: " Main task: scope and response planning.",
    final_rules: " Main task: delivery and control readiness.",
    supervisory: " Main task: evidence current-state compliance.",
    reporting: " Main task: data, ownership and submission readiness.",
    legislation: " Main task: confirm perimeter and legal effect.",
    monitoring: " Main task: keep on watchlist."
  }[stage.id] || "";

  return `${classification.impactStatement}${obligationText}${stageNote}`;
}

function buildSuggestedActions(item, classification, stage, ownership, impact, relevance, obligations, uncertainty) {
  const actions = [];
  const stageActions = GENERIC_STAGE_ACTIONS[stage.id] || GENERIC_STAGE_ACTIONS.monitoring;

  if (relevance.isRelevant) {
    actions.push(`Confirm ${ownership.primaryOwner} as lead and ${ownership.secondaryOwner} as support.`);
  } else {
    actions.push(`Keep with ${ownership.primaryOwner} on the watchlist until scope is clearer.`);
  }

  actions.push(stageActions[0]);
  actions.push(classification.coreActions[0]);

  if (impact.level === "High" || stage.id === "final_rules" || stage.id === "reporting") {
    actions.push(classification.coreActions[1]);
  }

  if (obligations.length) {
    actions.push(`Start with ${obligations.slice(0, 2).map(itemObligation => itemObligation.name).join(" and ")}.`);
  }

  if (item.expectedKeyMilestones) {
    actions.push(`Track: ${truncateText(item.expectedKeyMilestones, 140)}.`);
  }

  if (uncertainty.needsReview) {
    actions.push("Check the extracted row against the PDF before acting.");
  }

  return dedupeStrings(actions).slice(0, 5);
}

function buildRationale(item, classification, stage, relevance, obligations) {
  const positiveText = relevance.positiveReasons.length
    ? `Positive: ${relevance.positiveReasons.join(" ")}`
    : "No strong positive signals.";
  const negativeText = relevance.negativeReasons.length
    ? ` Negative: ${relevance.negativeReasons.join(" ")}`
    : "";
  const obligationText = obligations.length
    ? ` Domains: ${obligations.map(obligation => obligation.name).join(", ")}.`
    : "";

  return `Theme: ${classification.theme} / ${classification.subTheme}. Stage: ${stage.label}. ${positiveText}${negativeText}${obligationText}`;
}

function buildTimelinePoint(item) {
  const events = mergeTimelineEvents(
    item.structuredTimeline?.events || [],
    extractTimelineMarkers(
      `${item.expectedKeyMilestones || ""} ${item.initiativeDescription || ""} ${item.rawText || ""}`
    )
  );
  const now = Date.now();
  const match = events.find(event => event.dateValue >= now - 45 * 24 * 60 * 60 * 1000) || events[0] || null;
  return {
    label: match?.label || item.timingBucket,
    sortValue: match?.dateValue || estimateTimingBucketDate(item.timingBucket),
    raw:
      item.expectedKeyMilestones ||
      (item.structuredTimeline?.quarterSignals || []).map(signal => `${signal.label}: ${signal.status}`).join(" | ") ||
      item.timingBucket ||
      "To Be Confirmed",
    events: events.slice(0, 6)
  };
}

function determineCluster(item, classification, obligations) {
  if (classification.subTheme === "Settlement / T+1") {
    return "Post-Trade and Settlement Change";
  }
  if (obligations.some(obligation => obligation.name === "Reporting & MI")) {
    return "Reporting and Data Change";
  }
  if (obligations.some(obligation => obligation.name === "Outsourcing & Third Parties")) {
    return "Resilience and Third-Party Oversight";
  }
  if (obligations.some(obligation => obligation.name === "Governance & Accountability")) {
    return "Governance and Accountability";
  }
  if (classification.theme.includes("Market")) {
    return "Market Conduct and Wholesale Change";
  }
  if (classification.theme.includes("Investment Management")) {
    return "Investment Management Programme";
  }
  if (classification.theme.includes("Prudential")) {
    return "Prudential and Capital Programme";
  }
  if (classification.theme.includes("Sustainability")) {
    return "Sustainability and Disclosure Change";
  }
  return classification.theme;
}

function applyFeedbackOverrides(item) {
  const feedback = state.feedback[item.canonicalKey];
  if (!feedback) return item;

  const next = { ...item };
  next.feedback = feedback;
  next.reviewStatus = feedback.reviewStatus || "Not reviewed";
  next.reviewer = feedback.reviewer || "";
  next.targetDate = feedback.targetDate || "";
  next.actionOwner = feedback.actionOwner || "";
  next.ownerAction = feedback.ownerAction || "";
  next.boardComment = feedback.boardComment || "";

  if (feedback.themeOverride) {
    next.theme = feedback.themeOverride;
    next.internalSubTheme = "Analyst override";
    next.classificationSignals = dedupeStrings([
      `Analyst override: theme set to ${feedback.themeOverride}.`,
      ...(next.classificationSignals || [])
    ]);
  }

  if (feedback.ownerOverride) {
    next.primaryOwner = feedback.ownerOverride;
  }

  if (feedback.markNotRelevant) {
    next.isFmrukRelevant = false;
    next.relevanceScore = Math.min(next.relevanceScore || 0, 35);
    next.relevanceBand = "Low";
    next.fmrukEffectLevel = "Out of scope";
    next.whyNotRelevant = "Marked out of scope for FMRUK on analyst review.";
  }

  if (feedback.markUrgent) {
    next.impactLevel = "High";
    next.impactScore = Math.max(next.impactScore || 0, 82);
    next.priorityScore = clamp(Math.max(next.priorityScore || 0, 86), 0, 100);
    next.priorityBand = priorityBandFromScore(next.priorityScore);
    next.immediateActionRequired = true;
  }

  if (feedback.markParseIssue) {
    next.parseConfidence = Math.max((next.parseConfidence || 70) - 18, 20);
    next.parseConfidenceBand = next.parseConfidence >= 85 ? "High" : next.parseConfidence >= 70 ? "Medium" : "Low";
    next.parseWarnings = dedupeStrings([
      ...(next.parseWarnings || []),
      "Analyst flagged a parsing issue."
    ]);
  }

  if (feedback.note) {
    next.uncertaintyFlags = dedupeStrings([
      ...(next.uncertaintyFlags || []),
      `Analyst note: ${feedback.note}`
    ]);
  }

  if (feedback.reviewStatus === "Escalate") {
    next.priorityScore = clamp((next.priorityScore || 0) + 8, 0, 100);
    next.priorityBand = priorityBandFromScore(next.priorityScore);
    next.immediateActionRequired = true;
  }

  if (feedback.reviewStatus === "Action defined") {
    next.priorityScore = clamp((next.priorityScore || 0) + 3, 0, 100);
    next.priorityBand = priorityBandFromScore(next.priorityScore);
  }

  next.needsReview = Boolean(
    next.needsReview ||
      feedback.markParseIssue ||
      feedback.note ||
      feedback.themeOverride ||
      feedback.ownerOverride ||
      !feedback.reviewStatus ||
      feedback.reviewStatus === "Not reviewed"
  );

  return next;
}

function compareWithPreviousDataset(previousItems, newItems) {
  const previousMap = new Map(previousItems.map(item => [item.canonicalKey, item]));
  const nextMap = new Map(newItems.map(item => [item.canonicalKey, item]));
  const removedItems = [];
  let newCount = 0;
  let changedCount = 0;
  let acceleratedCount = 0;
  let delayedCount = 0;

  const items = newItems.map(item => {
    const previous = previousMap.get(item.canonicalKey);
    if (!previous) {
      newCount += 1;
      return applyPriorityDeltaBoost({
        ...item,
        changeStatus: "New",
        changeNarrative: "New in the latest upload."
      }, "New");
    }

    const deltaType = determineDeltaType(previous, item);
    if (deltaType === "Accelerated") acceleratedCount += 1;
    else if (deltaType === "Delayed") delayedCount += 1;
    else if (deltaType === "Changed") changedCount += 1;

    return applyPriorityDeltaBoost({
      ...item,
      changeStatus: deltaType,
      changeNarrative: buildChangeNarrative(previous, item, deltaType)
    }, deltaType);
  });

  previousMap.forEach(previousItem => {
    if (!nextMap.has(previousItem.canonicalKey)) {
      removedItems.push(previousItem.initiativeTitle);
    }
  });

  return {
    items,
    summary: {
      newCount,
      changedCount,
      acceleratedCount,
      delayedCount,
      removedCount: removedItems.length,
      removedItems: removedItems.slice(0, 8)
    }
  };
}

function determineDeltaType(previous, next) {
  const previousDate = previous.timeline?.sortValue || estimateTimingBucketDate(previous.timingBucket);
  const nextDate = next.timeline?.sortValue || estimateTimingBucketDate(next.timingBucket);

  if (Number.isFinite(previousDate) && Number.isFinite(nextDate)) {
    if (nextDate < previousDate - 28 * 24 * 60 * 60 * 1000) return "Accelerated";
    if (nextDate > previousDate + 28 * 24 * 60 * 60 * 1000) return "Delayed";
  }

  const meaningfulFields = [
    "initiativeDescription",
    "expectedKeyMilestones",
    "indicativeImpactOnFirms",
    "theme",
    "internalSubTheme",
    "stageLabel",
    "impactLevel",
    "relevanceBand"
  ];

  const changed = meaningfulFields.some(field => normaliseWs(previous[field]) !== normaliseWs(next[field]));
  return changed ? "Changed" : "Existing";
}

function buildChangeNarrative(previous, next, deltaType) {
  if (deltaType === "New") return "New in the latest upload.";
  if (deltaType === "Accelerated") return "Timing is earlier than in the previous upload.";
  if (deltaType === "Delayed") return "Timing is later than in the previous upload.";
  if (deltaType === "Changed") return "Description, stage or triage outcome changed.";
  return "No material change.";
}

function applyPriorityDeltaBoost(item, deltaType) {
  const boost = {
    New: 5,
    Accelerated: 8,
    Changed: 3,
    Delayed: 1,
    Existing: 0
  }[deltaType] || 0;

  const score = clamp((item.priorityScoreBase || item.priorityScore || 0) + boost, 0, 100);
  const band = priorityBandFromScore(score);
  const immediateActionRequired = Boolean(
    item.immediateActionRequired ||
      (
        item.isFmrukRelevant &&
        item.stage !== "monitoring" &&
        (band === "Critical" || (deltaType === "Accelerated" && item.priorityScoreBase >= 74))
      )
  );

  return {
    ...item,
    priorityScore: score,
    priorityBand: band,
    immediateActionRequired
  };
}

function reanalyseCurrentDataset() {
  if (!state.baseItems.length) {
    renderAll();
    return;
  }

  const deltaMap = new Map(
    state.raw.map(item => [
      item.canonicalKey,
      {
        changeStatus: item.changeStatus || "Existing",
        changeNarrative: item.changeNarrative || "No comparison note."
      }
    ])
  );

  state.raw = sortItems(
    analyseRows(state.baseItems).map(item => {
      const delta = deltaMap.get(item.canonicalKey) || {
        changeStatus: "Existing",
        changeNarrative: "No comparison note."
      };
      return applyPriorityDeltaBoost(
        {
          ...item,
          ...delta
        },
        delta.changeStatus
      );
    })
  );

  saveToStorage();
  renderAll();
}

function renderAll() {
  renderPageState();
  renderWorkspaceView();
  renderSettingsSummary();
  updateMeta();
  populateFilters();
  applyFilters();
}

function updateMeta() {
  if (!state.datasetMeta) {
    els.headerMeta.textContent = "No dataset loaded";
    if (els.sourceSummary) {
      els.sourceSummary.textContent = "Excel acts as the structured source of truth. PDF adds evidence and context where available.";
    }
    return;
  }

  const summary = state.datasetMeta.comparisonSummary;
  const comparisonText = summary
    ? [
        `${summary.newCount} new`,
        `${summary.changedCount} changed`,
        `${summary.acceleratedCount} earlier`,
        `${summary.delayedCount} later`
      ].join(" · ")
    : "No baseline comparison";

  els.headerMeta.textContent =
    `${state.datasetMeta.preferredSource || state.datasetMeta.fileType} primary · ${state.datasetMeta.rowCount} items · ${formatDate(state.datasetMeta.uploadedAt)} · ${comparisonText}`;

  if (els.sourceSummary) {
    const excelName = state.datasetMeta.sourceFiles?.excelFileName || "";
    const pdfName = state.datasetMeta.sourceFiles?.pdfFileName || "";
    const sourceText = [
      excelName ? `Excel: ${excelName}` : "Excel: not loaded",
      pdfName ? `PDF: ${pdfName}` : "PDF: not loaded"
    ].join(" · ");
    els.sourceSummary.textContent = `${sourceText} · Preferred source: ${state.datasetMeta.preferredSource || state.datasetMeta.fileType}.`;
  }
}

function populateFilters() {
  fillSelect(els.sectionFilter, unique(state.raw.map(item => item.sectionName)));
  fillSelect(els.themeFilter, unique(state.raw.map(item => item.theme)));
  fillSelect(els.ownerFilter, unique(state.raw.map(item => item.primaryOwner)));
  fillSelect(els.stageFilter, unique(state.raw.map(item => item.stageLabel)));
  fillSelect(els.clusterFilter, unique(state.raw.map(item => item.clusterLabel)));
}

function fillSelect(selectEl, values) {
  if (!selectEl) return;
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
  const query = els.searchInput.value.trim().toLowerCase();
  const section = els.sectionFilter.value;
  const theme = els.themeFilter.value;
  const owner = els.ownerFilter.value;
  const stage = els.stageFilter.value;
  const change = els.changeFilter.value;
  const cluster = els.clusterFilter.value;
  const minRelevance = Number(els.relevanceFilter.value || 0);
  const minParseConfidence = Number(els.parseConfidenceFilter.value || 0);
  const fmrukOnly = els.fmrukOnlyFilter.checked;
  const excludeAnnex = els.excludeAnnexFilter.checked;
  const needsReviewOnly = els.needsReviewFilter.checked;

  const items = state.raw.filter(item => {
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
      item.clusterLabel,
      item.sourceType,
      item.publicationStatus,
      item.sourceLink,
      item.potentialBusinessImpact,
      item.rationale,
      item.fmrukReadThrough,
      item.whyNotRelevant,
      ...(item.classificationSignals || []),
      ...(item.relevanceSignals || []),
      ...(item.relevanceNegativeSignals || []),
      ...(item.profileReasons || []),
      ...(item.uncertaintyFlags || []),
      ...(item.obligations || []).map(obligation => obligation.name),
      ...(flattenEvidence(item.evidence) || []).map(entry => entry.excerpt)
    ]
      .join(" ")
      .toLowerCase();

    return (
      (!query || haystack.includes(query)) &&
      (!section || item.sectionName === section) &&
      (!theme || item.theme === theme) &&
      (!owner || item.primaryOwner === owner) &&
      (!stage || item.stageLabel === stage) &&
      (!change || item.changeStatus === change) &&
      (!cluster || item.clusterLabel === cluster) &&
      (item.relevanceScore || 0) >= minRelevance &&
      (item.parseConfidence || 0) >= minParseConfidence &&
      (!fmrukOnly || item.isFmrukRelevant === true) &&
      (!excludeAnnex || item.sectionName !== "Annex: initiatives completed/stopped") &&
      (!needsReviewOnly || item.needsReview === true)
    );
  });

  state.filtered = sortItems(items);

  if (!state.filtered.find(item => item.id === state.selectedItemId)) {
    state.selectedItemId = state.filtered[0]?.id || null;
  }

  renderSummary(state.filtered);
  renderInitiativeList(state.filtered);
  const selectedItem = state.filtered.find(item => item.id === state.selectedItemId);
  renderDetail(selectedItem);
  renderPdfEvidence(selectedItem);
  renderTimelineList(state.filtered);
  renderAskResults();
}

function sortItems(items) {
  return [...items].sort((a, b) => {
    if ((b.priorityScore || 0) !== (a.priorityScore || 0)) {
      return (b.priorityScore || 0) - (a.priorityScore || 0);
    }
    if ((b.immediateActionRequired ? 1 : 0) !== (a.immediateActionRequired ? 1 : 0)) {
      return (b.immediateActionRequired ? 1 : 0) - (a.immediateActionRequired ? 1 : 0);
    }
    const aDate = Number.isFinite(a.timeline?.sortValue) ? a.timeline.sortValue : Number.POSITIVE_INFINITY;
    const bDate = Number.isFinite(b.timeline?.sortValue) ? b.timeline.sortValue : Number.POSITIVE_INFINITY;
    if (aDate !== bDate) return aDate - bDate;
    if ((b.relevanceScore || 0) !== (a.relevanceScore || 0)) {
      return (b.relevanceScore || 0) - (a.relevanceScore || 0);
    }
    return (b.impactScore || 0) - (a.impactScore || 0);
  });
}

function renderSummary(items) {
  els.listMeta.textContent = `${items.length} items`;
  els.portfolioNarrative.textContent = buildPortfolioNarrative(items);
  renderClusterList(items);
  renderDeltaList(items);
}

function buildPortfolioNarrative(items) {
  if (!items.length) {
    return "No items in scope.";
  }

  const priorityCounts = topCounts(items.map(item => item.priorityBand), 4);
  const topThemes = topCounts(items.map(item => item.theme), 3)
    .map(entry => entry.name)
    .join(", ");
  const immediate = items.filter(item => item.immediateActionRequired).length;
  const reviewCount = items.filter(item => item.needsReview).length;
  const profileMode = getPriorityModeConfig().label;
  const effectMode = getEffectModeConfig().label;
  const ownershipMode = getOwnershipModeConfig().label;

  return `${items.length} items in scope. ${immediate} require immediate action. ${reviewCount} still need review. Priority weighting: ${profileMode}. FMRUK effect mode: ${effectMode}. Ownership bias: ${ownershipMode}. Priority split: ${priorityCounts.map(entry => `${entry.name} ${entry.count}`).join(", ") || "none"}. Themes: ${topThemes || "none"}.`;
}

function renderClusterList(items) {
  const clusters = topCounts(items.map(item => item.clusterLabel), 6);
  els.clusterList.innerHTML = "";

  if (!clusters.length) {
    els.clusterList.innerHTML = "<li>No cluster data.</li>";
    return;
  }

  clusters.forEach(cluster => {
    const li = document.createElement("li");
    li.textContent = `${cluster.name} · ${cluster.count}`;
    els.clusterList.appendChild(li);
  });
}

function renderDeltaList(items) {
  const summary = state.datasetMeta?.comparisonSummary;
  els.deltaList.innerHTML = "";

  if (!summary) {
    els.deltaList.innerHTML = "<li>No comparison.</li>";
    return;
  }

  const lines = [
    `${summary.newCount} new`,
    `${summary.changedCount} changed`,
    `${summary.acceleratedCount} earlier`,
    `${summary.delayedCount} later`,
    summary.removedItems?.length
      ? `Removed: ${summary.removedItems.join(", ")}`
      : `${summary.removedCount} removed`
  ];

  const topChanged = items
    .filter(item => item.changeStatus !== "Existing")
    .slice(0, 2)
    .map(item => `${item.changeStatus}: ${item.initiativeTitle}`);

  [...lines, ...topChanged].forEach(line => {
    const li = document.createElement("li");
    li.textContent = line;
    els.deltaList.appendChild(li);
  });
}

function renderTimelineList(items) {
  els.timelineList.innerHTML = "";

  const eventRows = items
    .flatMap(item => {
      const events = item.timeline?.events?.length
        ? item.timeline.events
        : [
            {
              label: item.timeline?.label || item.timingBucket || "Not set",
              dateValue: item.timeline?.sortValue,
              detail: item.expectedKeyMilestones || item.initiativeDescription || item.fmrukReadThrough || ""
            }
          ];

      return events
        .filter(event => Number.isFinite(event?.dateValue))
        .slice(0, 3)
        .map((event, index) => ({
          item,
          event,
          key: `${item.id}:${event.dateValue}:${index}`
        }));
    })
    .sort((a, b) => a.event.dateValue - b.event.dateValue)
    .slice(0, 30);

  if (!eventRows.length) {
    els.timelineMeta.textContent = "No milestones";
    els.timelineList.innerHTML = `<div class="timeline-empty">No milestones in scope.</div>`;
    return;
  }

  const buckets = [];
  eventRows.forEach(row => {
    const label = row.event.label || row.item.timeline?.label || row.item.timingBucket || "Not set";
    const sortValue = row.event.dateValue || row.item.timeline?.sortValue || 0;
    const key = `${label}|${sortValue}`;
    const existing = buckets[buckets.length - 1];
    if (existing && existing.key === key) {
      existing.items.push(row);
      return;
    }

    buckets.push({
      key,
      label,
      sortValue,
      items: [row]
    });
  });

  els.timelineMeta.textContent = `${eventRows.length} milestones · ${buckets.length} date groups`;

  const rail = document.createElement("div");
  rail.className = "timeline-rail";

  buckets.forEach(bucket => {
    const wrapper = document.createElement("section");
    wrapper.className = "timeline-bucket";

    const itemsHtml = bucket.items
      .map(row => {
        const { item, event } = row;
        const activeClass = item.id === state.selectedItemId ? " active" : "";
        return `
          <article class="timeline-entry${activeClass}" data-timeline-item="${escapeHtml(item.id)}">
            <div class="timeline-entry-title">${escapeHtml(item.initiativeTitle)}</div>
            <div class="timeline-entry-copy">${escapeHtml(truncateText(event.detail || item.expectedKeyMilestones || item.initiativeDescription, 120))}</div>
            <div class="timeline-entry-meta">${escapeHtml(`${item.stageLabel} · ${item.primaryOwner || "Owner"}`)}</div>
          </article>
        `;
      })
      .join("");

    wrapper.innerHTML = `
      <div class="timeline-bucket-head">
        <span class="timeline-date-pill">${escapeHtml(bucket.label)}</span>
        <span class="timeline-bucket-count">${bucket.items.length} item${bucket.items.length === 1 ? "" : "s"}</span>
      </div>
      <div class="timeline-bucket-items">${itemsHtml}</div>
    `;

    wrapper.querySelectorAll("[data-timeline-item]").forEach(entry => {
      entry.addEventListener("click", () => {
        const selected = state.filtered.find(item => item.id === entry.dataset.timelineItem);
        if (!selected) return;
        state.currentWorkspaceView = "monitor";
        renderWorkspaceView();
        state.selectedItemId = selected.id;
        renderInitiativeList(state.filtered);
        renderDetail(selected);
        renderPdfEvidence(selected);
        renderTimelineList(state.filtered);
      });
    });

    rail.appendChild(wrapper);
  });

  els.timelineList.appendChild(rail);
}

function renderInitiativeList(items) {
  els.initiativeList.innerHTML = "";

  if (!items.length) {
    els.initiativeList.innerHTML = `<div class="empty-list">No items match the current filters.</div>`;
    return;
  }

  items.forEach(item => {
    const card = document.createElement("article");
    card.className = "initiative-card";
    if (item.id === state.selectedItemId) card.classList.add("selected");
    const cardDescription = buildCardDescription(item);

    card.innerHTML = `
      <div class="initiative-card-header">
        <div class="initiative-headline">
          <h3 class="initiative-title" title="${escapeHtml(item.initiativeTitle)}">${escapeHtml(item.initiativeTitle)}</h3>
        </div>
        <div class="initiative-owner-wrap">
          <span class="owner-pill" title="${escapeHtml(item.primaryOwner || "Owner")}">${escapeHtml(item.primaryOwner || "Owner")}</span>
        </div>
      </div>
      <p class="initiative-copy">${escapeHtml(cardDescription)}</p>
      <div class="initiative-card-footer">
        <span class="priority-pill priority-${slugify(item.priorityBand || "low")}">${escapeHtml(item.priorityBand || "Low")} priority</span>
        <span class="stage-pill" title="${escapeHtml(item.stageLabel)}">${escapeHtml(item.stageLabel)}</span>
      </div>
    `;

    card.addEventListener("click", () => {
      state.selectedItemId = item.id;
      renderInitiativeList(state.filtered);
      renderDetail(item);
      renderPdfEvidence(item);
      renderTimelineList(state.filtered);
    });

    els.initiativeList.appendChild(card);
  });
}

function renderDetail(item) {
  if (!item) {
    els.detailMeta.textContent = "No selection";
    els.detailPanel.innerHTML = "Select an item.";
    return;
  }

  const sourcePages = item.sourcePages?.length ? item.sourcePages.join(", ") : "N/A";
  const feedback = state.feedback[item.canonicalKey] || {};
  const analystStatus = feedback.reviewStatus || "Not reviewed";
  const metaCards = [
    ["FMRUK effect", item.fmrukEffectLevel || "Not set"],
    ["Priority", `${item.priorityBand} (${Math.round(item.priorityScore || 0)})`],
    ["Stage", item.stageLabel || "N/A"],
    ["Owner", `${item.primaryOwner || "N/A"}${item.secondaryOwner ? ` / ${item.secondaryOwner}` : ""}`],
    ["Timing", item.timeline?.label || item.timingBucket || "N/A"],
    ["Source", `${item.sourceType || state.datasetMeta?.fileType || "N/A"}${sourcePages !== "N/A" ? ` · Pages ${sourcePages}` : ""}`],
    ["Review status", analystStatus],
    ["Due date", feedback.targetDate || "Not set"]
  ]
    .map(
      ([label, value]) => `
        <div class="meta-card">
          <div class="meta-label">${escapeHtml(label)}</div>
          <div class="meta-value">${escapeHtml(value)}</div>
        </div>
      `
    )
    .join("");

  const actionItems = (item.suggestedActions || [])
    .map(action => `<li>${escapeHtml(action)}</li>`)
    .join("");

  const reviewFlags = item.uncertaintyFlags?.length
    ? `Current review flags: ${item.uncertaintyFlags.join(" ")}`
    : "No current review flags.";

  els.detailMeta.textContent = `${item.primaryOwner} · ${item.priorityBand} priority · ${item.sourceType || state.datasetMeta?.fileType || "source"}`;
  els.detailPanel.innerHTML = `
    <section class="detail-hero">
      <div class="detail-kicker">${escapeHtml(item.stageLabel)} · ${escapeHtml(item.changeStatus || "Existing")}</div>
      <h3>${escapeHtml(item.initiativeTitle)}</h3>
      <p class="detail-summary">${escapeHtml(truncateText(item.initiativeDescription || item.changeNarrative, 260))}</p>
    </section>

    <section class="detail-grid">
      <section class="support-card">
        <h4>How does this initiative affect FMRUK?</h4>
        <p class="detail-body-copy">${escapeHtml(item.fmrukReadThrough || item.rationale)}</p>
      </section>
      <section class="support-card">
        <h4>Potential action points</h4>
        <ul class="action-list">${actionItems || "<li>No action points recorded.</li>"}</ul>
      </section>
    </section>

    <section class="detail-meta-grid">${metaCards}</section>

    <section class="feedback-panel">
      <h4>Analyst review</h4>
      <p class="feedback-note">Use this section to confirm review status, define actions and feed the board and owner packs.</p>
      <div class="feedback-actions">
        <button class="feedback-chip ${feedback.markUrgent ? "active" : ""}" id="feedbackUrgentBtn" type="button">Urgent</button>
        <button class="feedback-chip ${feedback.markNotRelevant ? "active" : ""}" id="feedbackNotRelevantBtn" type="button">Not Relevant</button>
        <button class="feedback-chip ${feedback.markParseIssue ? "active" : ""}" id="feedbackParseBtn" type="button">Parsing Issue</button>
        <button class="feedback-chip" id="clearFeedbackBtn" type="button">Clear</button>
      </div>
      <div class="feedback-grid">
        <label>
          <span>Review status</span>
          <select id="reviewStatusSelect">
            ${buildReviewStatusOptions(feedback.reviewStatus)}
          </select>
        </label>
        <label>
          <span>Due date</span>
          <input type="date" id="targetDateInput" value="${escapeHtml(feedback.targetDate || "")}" />
        </label>
        <label>
          <span>Theme override</span>
          <select id="themeOverrideSelect">${buildThemeOptions(item.theme, feedback.themeOverride)}</select>
        </label>
        <label>
          <span>Owner override</span>
          <select id="ownerOverrideSelect">${buildOwnerOptions(item.primaryOwner, feedback.ownerOverride)}</select>
        </label>
        <label>
          <span>Reviewer</span>
          <input type="text" id="reviewerInput" value="${escapeHtml(feedback.reviewer || "")}" placeholder="Reviewer name" />
        </label>
        <label>
          <span>Action owner</span>
          <input type="text" id="actionOwnerInput" value="${escapeHtml(feedback.actionOwner || "")}" placeholder="Named owner or team" />
        </label>
      </div>
      <label>
        <span>Owner report action</span>
        <textarea id="ownerActionInput" rows="3" placeholder="Action required for the owner pack">${escapeHtml(feedback.ownerAction || "")}</textarea>
      </label>
      <label>
        <span>Board note</span>
        <textarea id="boardCommentInput" rows="3" placeholder="Short governance note for the board pack">${escapeHtml(feedback.boardComment || "")}</textarea>
      </label>
      <label>
        <span>Review note</span>
        <textarea id="feedbackNoteInput" rows="3">${escapeHtml(feedback.note || "")}</textarea>
      </label>
      <div class="inline-meta">${escapeHtml(reviewFlags)}</div>
      <div class="inline-actions">
        <button class="secondary-btn" id="saveFeedbackBtn" type="button">Save review</button>
      </div>
    </section>
  `;

  attachDetailEvents(item);
}

function attachDetailEvents(item) {
  els.detailPanel.querySelectorAll("[data-evidence-key]").forEach(button => {
    button.addEventListener("click", () => {
      state.currentPdfHighlightKey = button.dataset.evidenceKey;
      renderPdfEvidence(item);
      const activeChip = els.evidenceTrail.querySelector(`[data-evidence-key="${button.dataset.evidenceKey}"]`);
      if (activeChip) activeChip.scrollIntoView({ behavior: "smooth", block: "nearest", inline: "nearest" });
    });
  });

  const urgentBtn = document.getElementById("feedbackUrgentBtn");
  const notRelevantBtn = document.getElementById("feedbackNotRelevantBtn");
  const parseBtn = document.getElementById("feedbackParseBtn");
  const clearBtn = document.getElementById("clearFeedbackBtn");
  const saveBtn = document.getElementById("saveFeedbackBtn");

  urgentBtn?.addEventListener("click", () => toggleFeedbackFlag(item, "markUrgent"));
  notRelevantBtn?.addEventListener("click", () => toggleFeedbackFlag(item, "markNotRelevant"));
  parseBtn?.addEventListener("click", () => toggleFeedbackFlag(item, "markParseIssue"));
  clearBtn?.addEventListener("click", () => clearFeedbackForItem(item));
  saveBtn?.addEventListener("click", () => saveDetailFeedback(item));
}

function toggleFeedbackFlag(item, field) {
  const current = state.feedback[item.canonicalKey] || {};
  state.feedback[item.canonicalKey] = {
    ...current,
    [field]: !current[field]
  };
  saveFeedback();
  reanalyseCurrentDataset();
  preserveSelection(item.canonicalKey);
}

function clearFeedbackForItem(item) {
  delete state.feedback[item.canonicalKey];
  saveFeedback();
  reanalyseCurrentDataset();
  preserveSelection(item.canonicalKey);
}

function saveDetailFeedback(item) {
  const current = state.feedback[item.canonicalKey] || {};
  const themeOverride = document.getElementById("themeOverrideSelect")?.value || "";
  const ownerOverride = document.getElementById("ownerOverrideSelect")?.value || "";
  const reviewStatus = document.getElementById("reviewStatusSelect")?.value || "Not reviewed";
  const reviewer = normaliseWs(document.getElementById("reviewerInput")?.value || "");
  const targetDate = normaliseWs(document.getElementById("targetDateInput")?.value || "");
  const actionOwner = normaliseWs(document.getElementById("actionOwnerInput")?.value || "");
  const ownerAction = normaliseWs(document.getElementById("ownerActionInput")?.value || "");
  const boardComment = normaliseWs(document.getElementById("boardCommentInput")?.value || "");
  const note = normaliseWs(document.getElementById("feedbackNoteInput")?.value || "");

  const next = {
    ...current,
    themeOverride: themeOverride || "",
    ownerOverride: ownerOverride || "",
    reviewStatus,
    reviewer,
    targetDate,
    actionOwner,
    ownerAction,
    boardComment,
    note
  };

  state.feedback[item.canonicalKey] = next;
  saveFeedback();
  els.uploadStatus.textContent = "Feedback saved.";
  reanalyseCurrentDataset();
  preserveSelection(item.canonicalKey);
}

function preserveSelection(canonicalKey) {
  const selected = state.raw.find(item => item.canonicalKey === canonicalKey);
  if (selected) {
    state.selectedItemId = selected.id;
  }
  renderAll();
}

function renderPdfEvidence(item) {
  els.evidenceTrail.innerHTML = "";
  els.pdfPreview.innerHTML = "";

  if (!item) {
    els.pdfPreviewStatus.textContent = "Select an item.";
    return;
  }

  if (!state.pdfDocument) {
    els.pdfPreviewStatus.textContent = "No PDF reference loaded. Excel is currently acting as the structured source.";
    return;
  }

  const evidence = getEvidenceForFields(item, ["title", "lead", "description", "milestones", "impact", "timing", "isNew"], 20);
  if (!evidence.length) {
    els.pdfPreviewStatus.textContent = "PDF loaded, but no evidence links were matched to this initiative.";
    return;
  }

  if (!state.currentPdfHighlightKey || !evidence.find(entry => entry.key === state.currentPdfHighlightKey)) {
    state.currentPdfHighlightKey = evidence[0].key;
  }

  evidence.forEach(entry => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "evidence-chip";
    if (entry.key === state.currentPdfHighlightKey) button.classList.add("active");
    button.dataset.evidenceKey = entry.key;
    button.textContent = `P${entry.pageNumber} ${fieldLabel(entry.field)}: ${truncateText(entry.excerpt, 48)}`;
    button.addEventListener("click", () => {
      state.currentPdfHighlightKey = entry.key;
      renderPdfEvidence(item);
    });
    els.evidenceTrail.appendChild(button);
  });

  renderPdfPreview(item, evidence).catch(err => {
    console.error(err);
    els.pdfPreviewStatus.textContent = "Preview could not be rendered.";
  });
}

async function renderPdfPreview(item, evidence) {
  els.pdfPreview.innerHTML = "";

  if (!state.pdfDocument) {
    els.pdfPreviewStatus.textContent = "No PDF reference loaded for this session.";
    return;
  }

  const pages = item.sourcePages?.length ? item.sourcePages.slice(0, 4) : [];
  if (!pages.length) {
    els.pdfPreviewStatus.textContent = "No source pages captured.";
    return;
  }

  els.pdfPreviewStatus.textContent = `Pages ${pages.join(", ")}`;
  const activeEvidence = evidence.find(entry => entry.key === state.currentPdfHighlightKey);
  const token = ++state.pdfRenderToken;

  for (const pageNumber of pages) {
    const page = await state.pdfDocument.getPage(pageNumber);
    if (token !== state.pdfRenderToken) return;

    const viewport = page.getViewport({ scale: 1.1 });
    const wrapper = document.createElement("div");
    wrapper.className = "pdf-page";
    const header = document.createElement("div");
    header.className = "pdf-page-header";
    header.textContent = `Page ${pageNumber}`;
    const canvasShell = document.createElement("div");
    canvasShell.className = "pdf-canvas-shell";

    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    const ratio = window.devicePixelRatio || 1;
    canvas.width = Math.floor(viewport.width * ratio);
    canvas.height = Math.floor(viewport.height * ratio);
    canvas.style.width = `${viewport.width}px`;
    canvas.style.height = `${viewport.height}px`;
    context.scale(ratio, ratio);

    await page.render({ canvasContext: context, viewport }).promise;
    if (token !== state.pdfRenderToken) return;

    const overlay = document.createElement("div");
    overlay.className = "pdf-overlay";
    overlay.style.width = `${viewport.width}px`;
    overlay.style.height = `${viewport.height}px`;

    evidence
      .filter(entry => entry.pageNumber === pageNumber)
      .forEach(entry => {
        const highlight = document.createElement("div");
        highlight.className = "pdf-highlight";
        if (entry.key === activeEvidence?.key) highlight.classList.add("active");

        const scale = viewport.width / entry.pageWidth;
        const x1 = Math.max(entry.x1 * scale, 12);
        const width = Math.max((entry.x2 - entry.x1) * scale, viewport.width * 0.28);
        const top = clamp(viewport.height - entry.y * scale - 12, 8, viewport.height - 24);

        highlight.style.left = `${x1}px`;
        highlight.style.top = `${top}px`;
        highlight.style.width = `${Math.min(width, viewport.width - x1 - 8)}px`;
        highlight.style.height = "20px";
        overlay.appendChild(highlight);
      });

    canvasShell.appendChild(canvas);
    canvasShell.appendChild(overlay);
    wrapper.appendChild(header);
    wrapper.appendChild(canvasShell);
    els.pdfPreview.appendChild(wrapper);
  }
}

function runAskQuery() {
  const query = normaliseWs(els.searchInput.value);
  state.ask.query = query;

  if (!query) {
    state.ask.answer = "No search run.";
    state.ask.results = [];
    renderAskResults();
    return;
  }

  const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
  const sourceItems = state.filtered.length ? state.filtered : state.raw;
  const scored = sourceItems
    .map(item => {
      const haystack = [
        item.initiativeTitle,
        item.theme,
        item.internalSubTheme,
        item.rationale,
        item.fmrukReadThrough,
        item.potentialBusinessImpact,
        item.whyNotRelevant,
        item.primaryOwner,
        item.stageLabel,
        item.sourceType,
        item.publicationStatus,
        ...(item.obligations || []).map(obligation => obligation.name),
        ...(flattenEvidence(item.evidence) || []).map(entry => entry.excerpt)
      ]
        .join(" ")
        .toLowerCase();

      let score = 0;
      terms.forEach(term => {
        if (haystack.includes(term)) score += 8;
      });
      if (haystack.includes(query.toLowerCase())) score += 12;
      if (item.isFmrukRelevant) score += 4;
      if (item.immediateActionRequired) score += 4;

      return { item, score };
    })
    .filter(result => result.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 6);

  state.ask.results = scored.map(result => result.item);

  if (!state.ask.results.length) {
    state.ask.answer = `No results for "${query}".`;
    renderAskResults();
    return;
  }

  const topThemes = topCounts(state.ask.results.map(item => item.theme), 2)
    .map(entry => entry.name)
    .join(", ");
  const topOwners = topCounts(state.ask.results.map(item => item.primaryOwner), 2)
    .map(entry => entry.name)
    .join(", ");

  state.ask.answer = `${state.ask.results.length} results. Themes: ${topThemes || "none"}. Owners: ${topOwners || "none"}.`;
  renderAskResults();
}

function renderAskResults() {
  els.askAnswer.textContent = state.ask.answer || "No search run.";
  els.askResults.innerHTML = "";

  if (!state.ask.query) {
    return;
  }

  if (!state.ask.results.length) {
    els.askResults.innerHTML = "<li>No results.</li>";
    return;
  }

  state.ask.results.forEach(item => {
    const li = document.createElement("li");
    li.innerHTML = `<strong>${escapeHtml(item.initiativeTitle)}</strong><br />${escapeHtml(
      truncateText(item.fmrukReadThrough || item.rationale, 160)
    )}`;
    li.addEventListener("click", () => {
      state.currentWorkspaceView = "monitor";
      renderWorkspaceView();
      state.selectedItemId = item.id;
      renderInitiativeList(state.filtered);
      renderDetail(item);
      renderPdfEvidence(item);
      renderTimelineList(state.filtered);
    });
    els.askResults.appendChild(li);
  });
}

function exportJson() {
  if (!state.filtered.length) {
    els.uploadStatus.textContent = "No filtered dataset to export.";
    return;
  }

  downloadTextFile(
    "fmruk-regulatory-intelligence.json",
    JSON.stringify(state.filtered, null, 2),
    "application/json"
  );
}

function exportCsv() {
  if (!state.filtered.length) {
    els.uploadStatus.textContent = "No filtered dataset to export.";
    return;
  }

  const columns = [
    "sourceType",
    "priorityBand",
    "priorityScore",
    "changeStatus",
    "clusterLabel",
    "sectionName",
    "subcategory",
    "leadRegulator",
    "initiativeTitle",
    "publicationStatus",
    "theme",
    "internalSubTheme",
    "primaryOwner",
    "secondaryOwner",
    "stageLabel",
    "impactLevel",
    "relevanceScore",
    "fmrukEffectLevel",
    "parseConfidence",
    "reviewStatus",
    "reviewer",
    "actionOwner",
    "targetDate",
    "timelineLabel",
    "timingBucket",
    "expectedKeyMilestones",
    "sourceLink",
    "fmrukReadThrough",
    "potentialBusinessImpact",
    "rationale",
    "suggestedAction",
    "ownerAction",
    "boardComment"
  ];

  const lines = [
    columns.join(","),
    ...state.filtered.map(item =>
      columns
        .map(column => {
          if (column === "timelineLabel") {
            return csvEscape(item.timeline?.label || "");
          }
          return csvEscape(item[column]);
        })
        .join(",")
    )
  ];

  downloadTextFile("fmruk-regulatory-intelligence.csv", lines.join("\n"), "text/csv;charset=utf-8;");
}

function exportBoardBrief() {
  const filtered = state.filtered.filter(item =>
    item.immediateActionRequired ||
    ["Critical", "High"].includes(item.priorityBand) ||
    item.reviewStatus === "Escalate" ||
    item.changeStatus !== "Existing"
  );
  const items = sortItems(filtered.length ? filtered : state.filtered).slice(0, 12);
  if (!items.length) {
    els.uploadStatus.textContent = "No items available for a board brief.";
    return;
  }

  const JsPDF = getPdfConstructor();
  if (!JsPDF) {
    els.uploadStatus.textContent = "PDF export library did not load.";
    return;
  }

  const doc = buildReportPdf({
    kind: "board",
    title: "FMRUK Regulatory Board Brief",
    subtitle: "Portfolio position, management read-through, priority actions and control points from the current FCA initiatives review.",
    audience: `${getPriorityModeConfig().label} weighting`,
    items,
    summaryText: buildPortfolioNarrative(items)
  });

  const stamp = new Date().toISOString().slice(0, 10);
  doc.save(`fmruk-board-brief-${stamp}.pdf`);
}

function exportOwnerPack() {
  const relevantItems = sortItems(
    state.filtered.filter(item =>
      item.isFmrukRelevant ||
      item.needsReview ||
      item.reviewStatus === "Action defined" ||
      item.reviewStatus === "Escalate"
    )
  );

  if (!relevantItems.length) {
    els.uploadStatus.textContent = "No items available for an owner pack.";
    return;
  }

  const JsPDF = getPdfConstructor();
  if (!JsPDF) {
    els.uploadStatus.textContent = "PDF export library did not load.";
    return;
  }

  const doc = buildReportPdf({
    kind: "owner",
    title: "FMRUK Owner Pack",
    subtitle: "Owner actions, timings and review controls.",
    audience: `${getPriorityModeConfig().label} weighting`,
    items: relevantItems,
    summaryText: buildPortfolioNarrative(relevantItems)
  });

  const stamp = new Date().toISOString().slice(0, 10);
  doc.save(`fmruk-owner-pack-${slugify(getPriorityModeConfig().label)}-${stamp}.pdf`);
}

function getPdfConstructor() {
  return window.jspdf?.jsPDF || null;
}

function buildReportPdf(config) {
  const JsPDF = getPdfConstructor();
  const doc = new JsPDF({
    unit: "pt",
    format: "a4",
    compress: true
  });

  const ctx = createPdfContext(doc, config);
  drawReportCover(ctx, config);
  startPdfPage(ctx, "Portfolio position");
  addOverviewSection(ctx, config);

  if (config.kind === "board") {
    addMovementSection(ctx, config);
    addMaterialInitiativeSection(ctx, config.items);
    addPortfolioAuditSection(ctx, config.items, "Board review points");
  } else {
    addOwnerSummarySection(ctx, config.items);
    addOwnerDetailSections(ctx, config.items);
    addPortfolioAuditSection(ctx, config.items, "Owner control points");
  }

  addMethodologySection(ctx, config);
  applyPdfPageNumbers(doc, config.title);
  return doc;
}

function createPdfContext(doc, config) {
  return {
    doc,
    pageWidth: doc.internal.pageSize.getWidth(),
    pageHeight: doc.internal.pageSize.getHeight(),
    margin: 46,
    contentWidth: doc.internal.pageSize.getWidth() - 92,
    y: 64,
    colors: {
      navy: [23, 58, 89],
      navyDeep: [16, 39, 61],
      teal: [42, 109, 123],
      gold: [141, 109, 38],
      red: [161, 69, 69],
      green: [47, 119, 80],
      text: [23, 38, 56],
      soft: [95, 111, 131],
      line: [215, 223, 231],
      lineStrong: [197, 208, 220],
      panel: [247, 249, 251],
      white: [255, 255, 255]
    },
    datasetLabel: state.datasetMeta?.fileName || "Current filtered dataset",
    audience: config.audience,
    activeHeader: config.title,
    preparedAt: formatDate(new Date().toISOString()),
    profileMode: getPriorityModeConfig().label,
    effectMode: getEffectModeConfig().label,
    ownershipMode: getOwnershipModeConfig().label
  };
}

function drawReportCover(ctx, config) {
  const { doc, pageWidth, pageHeight, colors, margin } = ctx;

  doc.setFillColor(...colors.navyDeep);
  doc.rect(0, 0, pageWidth, 110, "F");
  doc.setFillColor(...colors.teal);
  doc.rect(0, 110, pageWidth, 4, "F");

  doc.setTextColor(...colors.white);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("FIDELITY INVESTMENTS", margin, 36);
  doc.setFontSize(26);
  doc.text(config.title, margin, 68);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  const subtitleLines = doc.splitTextToSize(config.subtitle, pageWidth - margin * 2);
  doc.text(subtitleLines, margin, 90);

  const panelTop = 148;
  const panelHeight = 146;
  doc.setFillColor(...colors.white);
  doc.setDrawColor(...colors.lineStrong);
  doc.roundedRect(margin, panelTop, pageWidth - margin * 2, panelHeight, 10, 10, "FD");

  const metaRows = [
    ["Prepared", ctx.preparedAt],
    ["Source dataset", ctx.datasetLabel],
    ["Priority weighting", ctx.profileMode],
    ["FMRUK effect mode", ctx.effectMode],
    ["Ownership bias", ctx.ownershipMode],
    ["Items in scope", String(config.items.length)],
    ["Parser", state.datasetMeta?.parserVersion || "Current version"],
    ["Build", `v${APP_VERSION} · ${APP_UPDATED_AT}`]
  ];

  let metaY = panelTop + 26;
  metaRows.forEach((row, index) => {
    const x = margin + 18 + (index % 2) * ((pageWidth - margin * 2 - 36) / 2);
    if (index > 0 && index % 2 === 0) metaY += 28;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(...colors.soft);
    doc.text(row[0].toUpperCase(), x, metaY);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10.5);
    doc.setTextColor(...colors.text);
    doc.text(truncateText(row[1], 56), x, metaY + 13);
  });

  drawPdfMetrics(ctx, buildPdfMetrics(config.items), panelTop + panelHeight + 20, 94, 2);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.setTextColor(...colors.soft);
  doc.text("Prepared from the current filtered review set.", margin, pageHeight - 24);
}

function buildPdfMetrics(items) {
  return [
    { label: "Items in scope", value: String(items.length) },
    {
      label: "High / critical",
      value: String(items.filter(item => ["Critical", "High"].includes(item.priorityBand)).length)
    },
    {
      label: "Immediate action",
      value: String(items.filter(item => item.immediateActionRequired).length)
    },
    {
      label: "Reviewed",
      value: String(items.filter(item => getReportReviewStatus(item) !== "Not reviewed").length)
    }
  ];
}

function drawPdfMetrics(ctx, metrics, y, boxHeight, columns) {
  const { doc, pageWidth, margin, colors } = ctx;
  const gap = 14;
  const columnWidth = (pageWidth - margin * 2 - gap * (columns - 1)) / columns;
  metrics.forEach((metric, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const x = margin + column * (columnWidth + gap);
    const top = y + row * (boxHeight + gap);

    doc.setFillColor(...colors.panel);
    doc.setDrawColor(...colors.line);
    doc.roundedRect(x, top, columnWidth, boxHeight, 8, 8, "FD");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(...colors.soft);
    doc.text(metric.label.toUpperCase(), x + 14, top + 20);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(24);
    doc.setTextColor(...colors.navyDeep);
    doc.text(metric.value, x + 14, top + 58);
  });
}

function startPdfPage(ctx, label) {
  const { doc, pageWidth, margin, colors } = ctx;
  doc.addPage();
  ctx.activeHeader = label;
  ctx.y = 74;

  doc.setFillColor(...colors.panel);
  doc.rect(0, 0, pageWidth, 56, "F");
  doc.setDrawColor(...colors.line);
  doc.line(margin, 56, pageWidth - margin, 56);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.setTextColor(...colors.teal);
  doc.text(label.toUpperCase(), margin, 24);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.setTextColor(...colors.navyDeep);
  doc.text(ctx.datasetLabel, margin, 44);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(...colors.soft);
  doc.text(`${ctx.profileMode} · ${ctx.effectMode}`, pageWidth - margin, 24, { align: "right" });
}

function ensurePdfSpace(ctx, neededHeight) {
  if (ctx.y + neededHeight <= ctx.pageHeight - 54) return;
  startPdfPage(ctx, ctx.activeHeader);
}

function addPdfSectionTitle(ctx, kicker, title) {
  ensurePdfSpace(ctx, 48);
  const { doc, margin, colors } = ctx;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.setTextColor(...colors.teal);
  doc.text(kicker.toUpperCase(), margin, ctx.y);
  ctx.y += 16;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(20);
  doc.setTextColor(...colors.navyDeep);
  doc.text(title, margin, ctx.y);
  ctx.y += 22;
}

function addPdfParagraph(ctx, text, options = {}) {
  const { doc, margin, contentWidth, colors } = ctx;
  const width = options.width || contentWidth;
  const x = options.x || margin;
  const font = options.font || "helvetica";
  const style = options.style || "normal";
  const size = options.size || 11;
  const gap = options.gap || 14;

  doc.setFont(font, style);
  doc.setFontSize(size);
  doc.setTextColor(...(options.color || colors.soft));
  const lines = doc.splitTextToSize(text, width);
  ensurePdfSpace(ctx, lines.length * (size + 2) + gap);
  doc.text(lines, x, ctx.y);
  ctx.y += lines.length * (size + 2) + gap;
}

function addPdfBulletList(ctx, items, options = {}) {
  const { doc, margin, colors } = ctx;
  items.filter(Boolean).forEach(item => {
    const bullet = `${options.bullet || "-"} ${item}`;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(options.size || 10.5);
    doc.setTextColor(...(options.color || colors.text));
    const lines = doc.splitTextToSize(bullet, options.width || ctx.contentWidth);
    ensurePdfSpace(ctx, lines.length * 14 + 6);
    doc.text(lines, (options.x || margin) + 4, ctx.y);
    ctx.y += lines.length * 14 + 4;
  });
  ctx.y += 8;
}

function addOverviewSection(ctx, config) {
  addPdfSectionTitle(ctx, "Portfolio position", "Current review set");
  addPdfParagraph(ctx, config.summaryText, { color: ctx.colors.text, size: 11 });

  const queueRows = sortItems(config.items)
    .slice(0, 8)
    .map(item => [
      truncateText(item.initiativeTitle, 44),
      item.priorityBand || "Low",
      item.primaryOwner || "N/A",
      item.timeline?.label || item.timingBucket || "Not set",
      getReportReviewStatus(item)
    ]);

  addTableSection(
    ctx,
    "Current queue",
    ["Initiative", "Priority", "Owner", "Timing", "Review"],
    queueRows
  );

  const distributionRows = topCounts(config.items.map(item => item.theme), 6).map(entry => [
    entry.name,
    String(entry.count)
  ]);
  addTableSection(ctx, "Theme distribution", ["Theme", "Items"], distributionRows);
}

function addMovementSection(ctx, config) {
  addPdfSectionTitle(ctx, "Change", "Movement against prior upload");

  const summary = state.datasetMeta?.comparisonSummary;
  const bullets = summary
    ? [
        `${summary.newCount} items are new compared with the previous upload.`,
        `${summary.changedCount} items changed in description, stage or triage outcome.`,
        `${summary.acceleratedCount} items moved earlier and ${summary.delayedCount} moved later.`,
        summary.removedCount
          ? `${summary.removedCount} items no longer appear in the latest upload.`
          : "No items dropped out of the latest upload."
      ]
    : ["No prior upload comparison is available for this dataset."];

  addPdfBulletList(ctx, bullets, { color: ctx.colors.text });

  const movementRows = sortItems(config.items.filter(item => item.changeStatus !== "Existing"))
    .slice(0, 10)
    .map(item => [
      item.changeStatus,
      truncateText(item.initiativeTitle, 42),
      item.priorityBand || "Low",
      item.timeline?.label || item.timingBucket || "Not set",
      item.primaryOwner || "N/A"
    ]);

  addTableSection(
    ctx,
    "Material movements",
    ["Status", "Initiative", "Priority", "Timing", "Owner"],
    movementRows
  );
}

function addMaterialInitiativeSection(ctx, items) {
  addPdfSectionTitle(ctx, "Board review", "Priority initiative notes");
  addPdfParagraph(
    ctx,
    "The notes below summarise the FMRUK read-through, current action points, review status and source context for the items most likely to require management attention.",
    { color: ctx.colors.text, size: 11 }
  );

  items.forEach((item, index) => {
    addInitiativeCard(ctx, item, index + 1, "board");
  });
}

function addOwnerSummarySection(ctx, items) {
  addPdfSectionTitle(ctx, "Owner summary", "Action coverage by function");
  addPdfParagraph(
    ctx,
    "This pack groups the filtered queue by primary owner and carries through review status, due dates and named actions from the analyst review workspace.",
    { color: ctx.colors.text, size: 11 }
  );

  const grouped = groupBy(items, item => item.primaryOwner || "Unassigned");
  const rows = Object.entries(grouped)
    .sort((a, b) => b[1].length - a[1].length)
    .map(([owner, ownerItems]) => [
      owner,
      String(ownerItems.length),
      String(ownerItems.filter(item => ["Critical", "High"].includes(item.priorityBand)).length),
      String(ownerItems.filter(item => item.immediateActionRequired).length),
      String(ownerItems.filter(item => getReportReviewStatus(item) === "Not reviewed").length)
    ]);

  addTableSection(
    ctx,
    "Owner queue",
    ["Owner", "Items", "High / critical", "Immediate", "Not reviewed"],
    rows
  );

  const openRows = sortItems(items)
    .filter(item => item.needsReview || getReportReviewStatus(item) === "Not reviewed")
    .slice(0, 10)
    .map(item => [
      truncateText(item.initiativeTitle, 36),
      item.primaryOwner || "N/A",
      item.priorityBand || "Low",
      getReportReviewStatus(item),
      formatShortDate(item.targetDate) || "Not set"
    ]);

  addTableSection(
    ctx,
    "Open review items",
    ["Initiative", "Owner", "Priority", "Review", "Due"],
    openRows
  );
}

function addOwnerDetailSections(ctx, items) {
  const grouped = groupBy(items, item => item.primaryOwner || "Unassigned");
  Object.entries(grouped)
    .sort((a, b) => b[1].length - a[1].length)
    .forEach(([owner, ownerItems]) => {
      addPdfSectionTitle(ctx, owner, "Owner action pack");
      addPdfParagraph(
        ctx,
        `${owner} currently owns ${ownerItems.length} items in the filtered scope. The tracker below shows priority, timing, review status and any named action owner recorded in analyst review.`,
        { color: ctx.colors.text, size: 11 }
      );

      const trackerRows = ownerItems.slice(0, 12).map(item => [
        truncateText(item.initiativeTitle, 32),
        item.priorityBand || "Low",
        item.stageLabel,
        formatShortDate(item.targetDate) || "Not set",
        getReportReviewStatus(item),
        truncateText(item.actionOwner || item.primaryOwner || "Not set", 18)
      ]);

      addTableSection(
        ctx,
        `${owner} tracker`,
        ["Initiative", "Priority", "Stage", "Due", "Review", "Action owner"],
        trackerRows
      );

      ownerItems.forEach((item, index) => {
        addInitiativeCard(ctx, item, index + 1, "owner");
      });
    });
}

function addPortfolioAuditSection(ctx, items, title) {
  addPdfSectionTitle(ctx, "Validation", title);
  addPdfBulletList(ctx, buildPortfolioAuditPoints(items), {
    color: ctx.colors.text
  });

  const flaggedRows = sortItems(
    items.filter(item =>
      item.needsReview ||
      (item.parseConfidence || 0) < 70 ||
      getReportReviewStatus(item) === "Not reviewed"
    )
  )
    .slice(0, 10)
    .map(item => [
      truncateText(item.initiativeTitle, 36),
      item.priorityBand || "Low",
      item.primaryOwner || "N/A",
      getReportReviewStatus(item),
      truncateText((item.uncertaintyFlags || [])[0] || "Manual validation required.", 54)
    ]);

  addTableSection(
    ctx,
    "Items needing challenge",
    ["Initiative", "Priority", "Owner", "Review", "Primary issue"],
    flaggedRows
  );
}

function addMethodologySection(ctx, config) {
  addPdfSectionTitle(ctx, "Method", "How this report was assembled");

  const profileSummary = PROFILE_FIELDS.map(([field]) => {
    const label = field
      .replace(/([A-Z])/g, " $1")
      .replace(/^./, char => char.toUpperCase());
    const values = (state.profile?.[field] || []).slice(0, 3).join(", ");
    return `${label}: ${values || "Not set"}.`;
  });

  addPdfBulletList(
    ctx,
    [
      `Source: FCA initiatives grid PDF parsed into lead regulator, title, description, milestone, impact and timing fields.`,
      `Priority model: ${ctx.profileMode} weighting applied across relevance, impact, urgency and delivery change.`,
      `FMRUK effect mode: ${ctx.effectMode}. Ownership bias: ${ctx.ownershipMode}.`,
      "Relevance uses the saved FMRUK footprint plus effect-mode thresholds to determine whether an item is direct, material, conditional, watchlist or out of scope.",
      "Review status, due dates, named action owners, owner actions and board notes are taken from analyst review in the workspace.",
      "Lower parse confidence, missing milestones and mixed relevance signals are carried forward as review flags.",
      ...profileSummary.slice(0, 4)
    ],
    { color: ctx.colors.text }
  );

  addPdfParagraph(
    ctx,
    config.kind === "board"
      ? "This report is intended for internal governance review and challenge. It supports, but does not replace, legal and compliance judgement."
      : "This pack is intended for internal owner follow-up, control mobilisation and documented review.",
    { color: ctx.colors.soft, size: 10.5 }
  );
}

function addInitiativeCard(ctx, item, index, mode) {
  const { doc, margin, contentWidth, colors } = ctx;
  const title = `${index}. ${item.initiativeTitle}`;
  const meta = [
    `${item.priorityBand || "Low"} priority`,
    item.stageLabel,
    `${item.primaryOwner || "N/A"}${item.secondaryOwner ? ` / ${item.secondaryOwner}` : ""}`,
    item.timeline?.label || item.timingBucket || "Not set",
    item.changeStatus !== "Existing" ? item.changeStatus : null
  ]
    .filter(Boolean)
    .join(" · ");
  const sections = [
    ["FMRUK read-through", item.fmrukReadThrough || item.potentialBusinessImpact || item.rationale],
    ["Potential action points", buildReportActionText(item, mode)],
    [mode === "board" ? "Management review" : "Owner review", buildReportReviewText(item, mode)],
    ["Source and timing", buildReportSourceText(item)]
  ];

  if (item.needsReview) {
    sections.push([
      "Open points",
      (item.uncertaintyFlags || []).slice(0, 2).join(" ") || "Manual validation required."
    ]);
  }

  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  const titleLines = doc.splitTextToSize(title, contentWidth - 30);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  const metaLines = doc.splitTextToSize(meta, contentWidth - 30);
  const sectionLayouts = sections.map(([heading, copy]) => ({
    heading,
    lines: doc.splitTextToSize(copy, contentWidth - 30)
  }));
  const boxHeight =
    22 +
    titleLines.length * 16 +
    metaLines.length * 11 +
    sectionLayouts.reduce((sum, section) => sum + 16 + section.lines.length * 12 + 8, 0) +
    12;

  ensurePdfSpace(ctx, boxHeight + 12);

  doc.setFillColor(...colors.white);
  doc.setDrawColor(...colors.lineStrong);
  doc.roundedRect(margin, ctx.y, contentWidth, boxHeight, 8, 8, "FD");

  let cursorY = ctx.y + 18;
  doc.setTextColor(...colors.navyDeep);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(14);
  doc.text(titleLines, margin + 15, cursorY);
  cursorY += titleLines.length * 16;

  doc.setTextColor(...colors.soft);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9.5);
  doc.text(metaLines, margin + 15, cursorY);
  cursorY += metaLines.length * 11 + 10;

  sectionLayouts.forEach(section => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(9);
    doc.setTextColor(...colors.teal);
    doc.text(section.heading.toUpperCase(), margin + 15, cursorY);
    cursorY += 12;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(10.2);
    doc.setTextColor(...colors.text);
    doc.text(section.lines, margin + 15, cursorY);
    cursorY += section.lines.length * 12 + 10;
  });

  ctx.y += boxHeight + 12;
}

function buildReportActionText(item, mode) {
  const baseActions = (item.suggestedActions || []).slice(0, 3);
  const ownerAction = normaliseWs(item.ownerAction || "");
  if (mode === "owner" && ownerAction) {
    return `${ownerAction}${baseActions.length ? ` ${baseActions.join(" ")}` : ""}`;
  }
  if (baseActions.length) {
    return baseActions.join(" ");
  }
  return mode === "owner"
    ? "No owner action has been recorded."
    : "No management action points have been recorded.";
}

function buildReportReviewText(item, mode) {
  const parts = [`Status: ${getReportReviewStatus(item)}.`];
  if (item.reviewer) parts.push(`Reviewer: ${item.reviewer}.`);
  if (item.targetDate) parts.push(`Due: ${formatShortDate(item.targetDate)}.`);
  if (item.actionOwner) parts.push(`Action owner: ${item.actionOwner}.`);

  if (mode === "board") {
    if (item.boardComment) {
      parts.push(`Board note: ${item.boardComment}.`);
    } else if (item.immediateActionRequired) {
      parts.push("This item currently warrants management attention on timing, ownership and delivery readiness.");
    }
  } else if (item.ownerAction) {
    parts.push(`Recorded owner action: ${item.ownerAction}.`);
  }

  if (item.needsReview) {
    parts.push("Further challenge or source validation is still required.");
  }

  return parts.join(" ");
}

function buildReportSourceText(item) {
  const pages = item.sourcePages?.length ? item.sourcePages.join(", ") : "N/A";
  const milestone = item.expectedKeyMilestones || item.timeline?.raw || "No clear milestone extracted.";
  return `Lead regulator: ${item.leadRegulator || "Not isolated"}. Section: ${item.sectionName || "N/A"}. Timing reference: ${truncateText(milestone, 150)} Evidence pages: ${pages}.`;
}

function getReportReviewStatus(item) {
  return item.reviewStatus || "Not reviewed";
}

function formatShortDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  }).format(date);
}

function buildPortfolioAuditPoints(items) {
  const summary = state.datasetMeta?.comparisonSummary;
  const notReviewed = items.filter(item => getReportReviewStatus(item) === "Not reviewed").length;
  const actionDefined = items.filter(item => getReportReviewStatus(item) === "Action defined").length;
  const escalated = items.filter(item => getReportReviewStatus(item) === "Escalate").length;
  const lowParseCount = items.filter(item => (item.parseConfidence || 0) < 70).length;
  const missingMilestoneCount = items.filter(item => !item.expectedKeyMilestones).length;
  const urgentCount = items.filter(item => item.immediateActionRequired).length;

  return [
    `${urgentCount} items currently require immediate action under the active weighting model.`,
    `${notReviewed} items do not yet have a recorded analyst review status.`,
    `${actionDefined} items have a defined action recorded and ${escalated} are marked for escalation.`,
    `${lowParseCount} items carry lower parsing confidence and should be checked against the source PDF.`,
    `${missingMilestoneCount} items do not yet have a clean milestone extraction, which weakens timing confidence.`,
    summary?.removedCount
      ? `${summary.removedCount} items no longer appear in the latest upload and should be checked for completion, deferral or removal from the grid.`
      : "No items dropped out of the latest comparison set."
  ];
}

function addTableSection(ctx, title, head, body) {
  if (!body.length) return;

  addPdfParagraph(ctx, title, {
    font: "helvetica",
    style: "bold",
    size: 12,
    color: ctx.colors.navyDeep,
    gap: 10
  });

  const doc = ctx.doc;
  if (typeof doc.autoTable === "function") {
    ensurePdfSpace(ctx, 90);
    doc.autoTable({
      startY: ctx.y,
      head: [head],
      body,
      margin: {
        left: ctx.margin,
        right: ctx.margin
      },
      theme: "grid",
      tableWidth: "auto",
      styles: {
        font: "helvetica",
        fontSize: 9,
        cellPadding: 6,
        textColor: ctx.colors.text
      },
      headStyles: {
        fillColor: ctx.colors.navy,
        textColor: 255,
        fontStyle: "bold"
      },
      alternateRowStyles: {
        fillColor: [248, 250, 252]
      }
    });
    ctx.y = doc.lastAutoTable.finalY + 18;
    return;
  }

  body.forEach(row => {
    addPdfBulletList(ctx, [row.join(" | ")], { color: ctx.colors.text });
  });
}

function applyPdfPageNumbers(doc, title) {
  const totalPages = doc.getNumberOfPages();
  for (let page = 1; page <= totalPages; page += 1) {
    doc.setPage(page);
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    doc.setDrawColor(214, 223, 233);
    doc.line(46, pageHeight - 28, pageWidth - 46, pageHeight - 28);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(95, 111, 132);
    doc.text(title, 46, pageHeight - 12);
    doc.text(`Page ${page} of ${totalPages}`, pageWidth - 46, pageHeight - 12, { align: "right" });
  }
}

async function loadPdfDocumentFromBuffer(buffer) {
  const loadingTask = pdfjsLib.getDocument({ data: clonePdfBinary(buffer) });
  return loadingTask.promise;
}

function openPdfDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(PDF_DB_NAME, 1);
    request.onupgradeneeded = event => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains(PDF_STORE_NAME)) {
        db.createObjectStore(PDF_STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveStoredPdf(buffer) {
  const db = await openPdfDatabase();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(PDF_STORE_NAME, "readwrite");
    tx.objectStore(PDF_STORE_NAME).put(clonePdfBinary(buffer), PDF_RECORD_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

async function loadStoredPdf() {
  const db = await openPdfDatabase();
  const result = await new Promise((resolve, reject) => {
    const tx = db.transaction(PDF_STORE_NAME, "readonly");
    const request = tx.objectStore(PDF_STORE_NAME).get(PDF_RECORD_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return result ? clonePdfBinary(result) : null;
}

async function clearStoredPdf() {
  const db = await openPdfDatabase();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(PDF_STORE_NAME, "readwrite");
    tx.objectStore(PDF_STORE_NAME).delete(PDF_RECORD_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

function getEvidenceForFields(item, fields, limit) {
  const seen = new Set();
  const entries = [];

  fields.forEach(field => {
    (item.evidence?.[field] || []).forEach(entry => {
      if (seen.has(entry.key)) return;
      seen.add(entry.key);
      entries.push(entry);
    });
  });

  return entries.slice(0, limit);
}

function flattenEvidence(evidence) {
  if (!evidence) return [];
  return Object.values(evidence).flat();
}

function calculateEvidenceCoverage(items) {
  if (!items?.length) return 0;
  const covered = items.filter(item => flattenEvidence(item.evidence).length > 0).length;
  return covered / items.length;
}

function formatPercent(value) {
  if (!Number.isFinite(value)) return "0%";
  return `${Math.round(value * 100)}%`;
}

function buildThemeOptions(currentTheme, overrideTheme) {
  const options = unique([
    "",
    ...(state.raw || []).map(item => item.theme),
    ...CLASSIFICATION_RULES.map(rule => rule.theme),
    currentTheme,
    overrideTheme
  ]);
  return options
    .map(value => {
      const selected = value === (overrideTheme || "") ? " selected" : "";
      const label = value || "No override";
      return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(label)}</option>`;
    })
    .join("");
}

function buildOwnerOptions(currentOwner, overrideOwner) {
  const owners = unique([
    "",
    ...(state.raw || []).map(item => item.primaryOwner),
    currentOwner,
    overrideOwner,
    "Compliance",
    "Risk",
    "Operations",
    "Technology",
    "Finance",
    "Legal",
    "Product",
    "HR"
  ]);

  return owners
    .map(value => {
      const selected = value === (overrideOwner || "") ? " selected" : "";
      const label = value || "No override";
      return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(label)}</option>`;
    })
    .join("");
}

function buildReviewStatusOptions(selectedStatus) {
  const options = [
    "Not reviewed",
    "Reviewed",
    "Action defined",
    "Escalate",
    "Monitor"
  ];

  return options
    .map(value => {
      const selected = value === (selectedStatus || "Not reviewed") ? " selected" : "";
      return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(value)}</option>`;
    })
    .join("");
}

function fieldLabel(field) {
  return {
    title: "title",
    lead: "lead",
    description: "description",
    milestones: "milestone",
    impact: "impact",
    consumer: "consumer",
    timing: "timing",
    isNew: "new",
    general: "source"
  }[field] || field;
}

function topCounts(values, limit) {
  const counts = {};
  values.filter(Boolean).forEach(value => {
    counts[value] = (counts[value] || 0) + 1;
  });

  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([name, count]) => ({ name, count }));
}

function scoreSignalGroups(groups, blob) {
  let score = 0;
  const reasons = [];

  groups.forEach(group => {
    const matched = group.terms.filter(term => blob.includes(term.toLowerCase()));
    if (!matched.length) return;
    score += group.weight;
    reasons.push(group.reason);
  });

  return {
    score,
    reasons: dedupeStrings(reasons)
  };
}

function parseTimelineDate(text) {
  const source = String(text || "").toLowerCase();
  if (!source) return null;

  const dayMonthYearMatch = source.match(
    /\b(\d{1,2})\s+(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{4})\b/
  );
  if (dayMonthYearMatch) {
    const [, day, month, year] = dayMonthYearMatch;
    const date = new Date(Number(year), MONTH_LOOKUP[month], Number(day));
    return {
      label: `${day} ${capitalise(month)} ${year}`,
      dateValue: date.getTime()
    };
  }

  const monthYearMatch = source.match(
    /\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{4})\b/
  );
  if (monthYearMatch) {
    const [, month, year] = monthYearMatch;
    const date = new Date(Number(year), MONTH_LOOKUP[month], 1);
    return {
      label: `${capitalise(month)} ${year}`,
      dateValue: date.getTime()
    };
  }

  const quarterMatch = source.match(/\b(q[1-4]|h[12])\s+(\d{4})\b/);
  if (quarterMatch) {
    const [, period, yearText] = quarterMatch;
    const year = Number(yearText);
    const date = period === "q1"
      ? new Date(year, 0, 1)
      : period === "q2"
      ? new Date(year, 3, 1)
      : period === "q3"
      ? new Date(year, 6, 1)
      : period === "q4"
      ? new Date(year, 9, 1)
      : period === "h1"
      ? new Date(year, 0, 1)
      : new Date(year, 6, 1);
    return {
      label: `${period.toUpperCase()} ${year}`,
      dateValue: date.getTime()
    };
  }

  return null;
}

function extractTimelineMarkers(text) {
  const source = normaliseWs(text);
  if (!source) return [];

  const fragments = source
    .split(/\s+\|\s+|(?<=[.!?])\s+/)
    .map(normaliseWs)
    .filter(Boolean);
  const markers = [];
  const seen = new Set();

  const pushMarker = (entry, detail) => {
    if (!entry?.label || !Number.isFinite(entry.dateValue)) return;
    const cleanDetail = truncateText(normaliseWs(detail || source), 170);
    const key = `${entry.label}|${entry.dateValue}|${slugify(cleanDetail).slice(0, 36)}`;
    if (seen.has(key)) return;
    seen.add(key);
    markers.push({
      ...entry,
      detail: cleanDetail
    });
  };

  const candidates = fragments.length ? fragments : [source];
  candidates.forEach(fragment => {
    collectTimelineDateMatches(fragment).forEach(entry => {
      pushMarker(entry, fragment);
    });
  });

  if (!markers.length) {
    collectTimelineDateMatches(source).forEach(entry => {
      pushMarker(entry, source);
    });
  }

  return markers.sort((a, b) => a.dateValue - b.dateValue);
}

function collectTimelineDateMatches(text) {
  const source = String(text || "").toLowerCase();
  if (!source) return [];

  const matches = [];
  const push = entry => {
    if (!entry?.label || !Number.isFinite(entry.dateValue)) return;
    if (!matches.find(match => match.label === entry.label && match.dateValue === entry.dateValue)) {
      matches.push(entry);
    }
  };

  const quarterPattern = /\b(q[1-4]|h[12])\s+(\d{4})\b/gi;
  let match = quarterPattern.exec(source);
  while (match) {
    push(parseTimelineDate(match[0]));
    match = quarterPattern.exec(source);
  }

  const monthPattern = /\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}\b/gi;
  match = monthPattern.exec(source);
  while (match) {
    push(parseTimelineDate(match[0]));
    match = monthPattern.exec(source);
  }

  const dayMonthPattern = /\b\d{1,2}\s+(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}\b/gi;
  match = dayMonthPattern.exec(source);
  while (match) {
    push(parseTimelineDate(match[0]));
    match = dayMonthPattern.exec(source);
  }

  const gridRanges = [
    { pattern: /\boct-dec\s+(\d{4})\b/gi, label: year => `Oct-Dec ${year}`, month: 9 },
    { pattern: /\bjan-mar\s+(\d{4})\b/gi, label: year => `Jan-Mar ${year}`, month: 0 },
    { pattern: /\bapr-jun\s+(\d{4})\b/gi, label: year => `Apr-Jun ${year}`, month: 3 },
    { pattern: /\bjul-sept\s+(\d{4})\b/gi, label: year => `Jul-Sep ${year}`, month: 6 },
    { pattern: /\bjul-dec\s+(\d{4})\b/gi, label: year => `Jul-Dec ${year}`, month: 6 },
    { pattern: /\bjan-jun\s+(\d{4})\b/gi, label: year => `Jan-Jun ${year}`, month: 0 },
    { pattern: /\bjuly to december\s+(\d{4})\b/gi, label: year => `Jul-Dec ${year}`, month: 6 },
    { pattern: /\bpost july\s+(\d{4})\b/gi, label: year => `Post July ${year}`, month: 6 }
  ];

  gridRanges.forEach(range => {
    let rangeMatch = range.pattern.exec(source);
    while (rangeMatch) {
      const year = Number(rangeMatch[1]);
      push({
        label: range.label(year),
        dateValue: new Date(year, range.month, 1).getTime()
      });
      rangeMatch = range.pattern.exec(source);
    }
  });

  const halfYearPatterns = [
    { pattern: /\bfirst half of (\d{4})\b/gi, label: year => `H1 ${year}`, month: 0 },
    { pattern: /\bsecond half of (\d{4})\b/gi, label: year => `H2 ${year}`, month: 6 },
    { pattern: /\bmid[- ](\d{4})\b/gi, label: year => `Mid-${year}`, month: 5 },
    { pattern: /\bend of (\d{4})\b/gi, label: year => `End ${year}`, month: 11 },
    { pattern: /\bwinter (\d{4})-(\d{2})\b/gi, label: (year, nextYear) => `Winter ${year}-${nextYear}`, month: 11 }
  ];

  halfYearPatterns.forEach(patternEntry => {
    let timeMatch = patternEntry.pattern.exec(source);
    while (timeMatch) {
      const year = Number(timeMatch[1]);
      const nextYear = timeMatch[2] ? Number(`20${timeMatch[2]}`) : null;
      push({
        label: patternEntry.label(year, nextYear),
        dateValue: new Date(year, patternEntry.month, 1).getTime()
      });
      timeMatch = patternEntry.pattern.exec(source);
    }
  });

  return matches.sort((a, b) => a.dateValue - b.dateValue);
}

function estimateTimingBucketDate(bucket) {
  const now = new Date();
  if (bucket === "Near Term") return now.getTime() + 30 * 24 * 60 * 60 * 1000;
  if (bucket === "Medium Term") return now.getTime() + 180 * 24 * 60 * 60 * 1000;
  if (bucket === "Longer Term") return now.getTime() + 420 * 24 * 60 * 60 * 1000;
  return Number.POSITIVE_INFINITY;
}

function buildBlob(item) {
  return [
    item.sectionName,
    item.subcategory,
    item.leadRegulator,
    item.initiativeTitle,
    item.initiativeDescription,
    item.expectedKeyMilestones,
    item.publicationStatus,
    item.sourceLink,
    (item.structuredTimeline?.quarterSignals || []).map(signal => `${signal.label} ${signal.status}`).join(" "),
    item.rawText
  ]
    .join(" ")
    .toLowerCase();
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
    /\b(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}\b[^.]*\.?/gi,
    /\b\d{1,2}\s+(?:january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}\b[^.]*\.?/gi
  ];

  patterns.forEach(pattern => {
    const found = text.match(pattern);
    if (found) matches.push(...found);
  });

  return dedupeStrings(matches.map(normaliseWs)).slice(0, 5).join(" | ");
}

function segmentIsBold(segment) {
  return /bold|black|heavy|semibold|demi/i.test(String(segment?.fontName || ""));
}

function segmentsForCell(row, key) {
  const cell = row.cells?.[key];
  if (!cell || cell.x1 == null || cell.x2 == null) return [];

  return (row.segments || []).filter(segment => {
    const centre = segment.x + segment.width / 2;
    return centre >= cell.x1 - 2 && centre <= cell.x2 + 2;
  });
}

function extractInitiativeCellParts(row) {
  const segments = segmentsForCell(row, "initiative");
  if (!segments.length) {
    const fallback = cleanStructuredText(cellText(row, "initiative"));
    return {
      boldText: "",
      bodyText: "",
      fullText: fallback
    };
  }

  let boldText = "";
  let bodyText = "";
  let fullText = "";
  let titleClosed = false;

  segments.forEach(segment => {
    fullText = joinInline(fullText, segment.text);
    if (segmentIsBold(segment) && !titleClosed) {
      boldText = joinInline(boldText, segment.text);
    } else {
      titleClosed = true;
      bodyText = joinInline(bodyText, segment.text);
    }
  });

  return {
    boldText: cleanTitleCandidate(boldText),
    bodyText: cleanStructuredText(bodyText),
    fullText: cleanStructuredText(fullText)
  };
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

function cleanTitleCandidate(value) {
  let title = normaliseWs(value);
  const pattern = new RegExp(
    `^(?:${REGULATOR_CODES.join("|")})(?:\\/(?:${REGULATOR_CODES.join("|")}))*\\s+`,
    "i"
  );
  title = title.replace(pattern, "");
  return normaliseWs(title);
}

function isPlausibleTitle(value) {
  const title = normaliseWs(value);
  if (!title || title.length < 6) return false;
  if (title.length > 160) return false;
  if (isLeadToken(title)) return false;
  if (looksLikeBodyCopy(title)) return false;
  if (ROW_NOISE_PATTERNS.some(pattern => pattern.test(title))) return false;
  return !HEADER_TEXT_PATTERNS.some(pattern => pattern.test(title));
}

function shouldTreatAsTitleContinuation(accumulator, value, row) {
  if (!value) return false;
  if (accumulator.initiativeTitle.length > 140) return false;
  if (accumulator.titleLineCount >= 3) return false;
  if (cellText(row, "milestones") || cellText(row, "impact") || cellText(row, "consumer")) return false;
  if (looksLikeBodyCopy(value) && !row.hasBold) return false;
  if (/[.:;]/.test(value) && !row.hasBold) return false;
  if (looksLikeMilestoneText(value)) return false;
  if (isQuarterHeaderText(value)) return false;
  if (row.hasBold) return true;
  if (!/[a-z]/.test(value) || value.length > 95) return false;
  return /^[A-Z0-9]/.test(value) && !/[.!?]$/.test(value);
}

function looksLikeMilestoneText(text) {
  const value = String(text || "").toLowerCase();
  return /\b(?:q[1-4]|h[12]|january|february|march|april|may|june|july|august|september|october|november|december|202[5-9])\b/.test(value);
}

function isQuarterHeaderText(text) {
  const value = normaliseWs(text).toLowerCase();
  if (!value) return false;

  if (
    /^(oct-dec|jan-mar|apr-jun|jul-sept|jul-dec|jan-jun|post july)\s+\d{4}$/.test(value) ||
    /^july to december \d{4}$/.test(value) ||
    /^beyond (that )?horizon$/.test(value)
  ) {
    return true;
  }

  const hits = [
    /\boct-dec \d{4}\b/.test(value),
    /\bjan-mar \d{4}\b/.test(value),
    /\bapr-jun \d{4}\b/.test(value),
    /\bjul-sept \d{4}\b/.test(value),
    /\bjul-dec \d{4}\b/.test(value),
    /\bjan-jun \d{4}\b/.test(value),
    /\bpost july \d{4}\b/.test(value)
  ].filter(Boolean).length;

  return hits >= 2;
}

function cleanStructuredText(value) {
  const text = normaliseWs(value);
  if (!text) return "";
  if (ROW_NOISE_PATTERNS.some(pattern => pattern.test(text))) return "";
  if (HEADER_TEXT_PATTERNS.some(pattern => pattern.test(text))) return "";
  if (isQuarterHeaderText(text)) return "";
  if (/^(h|l|u|m|yes|no)$/i.test(text)) return "";
  return text;
}

function looksLikeBodyCopy(value) {
  const text = normaliseWs(value);
  if (!text) return false;
  const wordCount = text.split(/\s+/).filter(Boolean).length;
  if (/[.!?]/.test(text) && wordCount >= 9) return true;
  if (
    wordCount >= 14 &&
    /\b(the fca|will|would|should|review|consultation|published|update|further engagement|statement)\b/i.test(text)
  ) {
    return true;
  }
  return false;
}

function normaliseInitiativeDescription(value, title) {
  const text = cleanStructuredText(value);
  if (!text) return "";
  const withoutTitle = stripRepeatedTitle(text, title);
  return cleanStructuredText(withoutTitle);
}

function isLikelyTitleContinuationText(value, row) {
  const text = cleanStructuredText(value);
  if (!text) return false;
  if (cellText(row, "milestones") || cellText(row, "impact") || cellText(row, "consumer")) return false;
  if (looksLikeMilestoneText(text)) return false;
  if (looksLikeBodyCopy(text) && !row.hasBold) return false;
  if (row.hasBold) return true;
  return text.length <= 96 && !/[.!?]$/.test(text);
}

function cellText(row, key) {
  return normaliseWs(row.cells?.[key]?.text || "");
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

function buildCanonicalKey(item) {
  return [
    normaliseLookupKey(canonicaliseSectionName(item.sectionName)),
    normaliseLookupKey(item.leadRegulator),
    looseTitleKey(item.initiativeTitle)
  ].join("|");
}

function normaliseLookupKey(value) {
  return normaliseWs(value)
    .toLowerCase()
    .replace(/\u00a0/g, " ")
    .replace(/[’']/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function groupBy(items, getKey) {
  return items.reduce((acc, item) => {
    const key = getKey(item);
    if (!acc[key]) acc[key] = [];
    acc[key].push(item);
    return acc;
  }, {});
}

function clonePdfBinary(value) {
  if (!value) {
    throw new Error("No PDF binary was provided.");
  }

  if (value instanceof Uint8Array) {
    return new Uint8Array(value);
  }

  if (ArrayBuffer.isView(value)) {
    return new Uint8Array(
      value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength)
    );
  }

  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value.slice(0));
  }

  throw new Error("Unsupported PDF binary format.");
}

function csvEscape(value) {
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
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

function splitLines(value) {
  return String(value || "")
    .split(",")
    .map(normaliseWs)
    .filter(Boolean);
}

function unique(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b));
}

function dedupeStrings(values) {
  return [...new Set(values.map(value => normaliseWs(value)).filter(Boolean))];
}

function appendUnique(target, value) {
  if (value == null || value === "") return;
  if (Array.isArray(target) && !target.includes(value)) {
    target.push(value);
  }
}

function truncateText(value, maxLength) {
  const text = String(value || "");
  if (text.length <= maxLength) return text;
  return `${text.slice(0, Math.max(maxLength - 3, 0))}...`;
}

function buildCardDescription(item) {
  const sourceText = normaliseWs(item.initiativeDescription || "");
  if (!sourceText) {
    return truncateText(item.fmrukReadThrough || item.rationale || "No summary available.", 220);
  }

  const firstSentence = sourceText.split(/(?<=[.!?])\s+/).map(normaliseWs).filter(Boolean)[0] || sourceText;
  if (firstSentence.length >= 70) {
    return truncateText(firstSentence, 220);
  }

  return truncateText(sourceText, 220);
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

function capitalise(value) {
  const text = String(value || "");
  return text.charAt(0).toUpperCase() + text.slice(1);
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
