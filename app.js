const STORAGE_KEY = "fmruk_reg_intelligence_v8";
const PROFILE_KEY = "fmruk_profile_v8";
const FEEDBACK_KEY = "fmruk_feedback_v8";
const PDF_DB_NAME = "fmruk_pdf_workspace";
const PDF_STORE_NAME = "pdfs";
const PDF_RECORD_KEY = "current_pdf";
const APP_VERSION = "1.2.0";
const APP_UPDATED_AT = "29 March 2026";
const APP_ENVIRONMENT_LABEL = "Browser workspace";

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

const state = {
  baseItems: [],
  raw: [],
  filtered: [],
  selectedItemId: null,
  datasetMeta: null,
  profile: null,
  feedback: {},
  currentPage: "home",
  roleView: "Compliance",
  dashboardMode: "workspace",
  activePreset: "all",
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

const DASHBOARD_MODES = {
  workspace: {
    label: "Analyst Workspace",
    summary:
      "Analyst workspace keeps the full intake, filtering, review and evidence toolset visible."
  },
  committee: {
    label: "Committee View",
    summary:
      "Committee view reduces lower-value controls and keeps the narrative, priority items, timing and evidence in clearer focus."
  },
  executive: {
    label: "Executive Summary",
    summary:
      "Executive summary strips the workspace back to portfolio position, material items, movement, ownership and near-term decisions."
  },
  evidence: {
    label: "Evidence Review",
    summary:
      "Evidence review gives more room to the source PDF and evidence trail for challenge, verification and line-by-line checking."
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

const ROLE_VIEWS = {
  Compliance: {
    summary:
      "Compliance view prioritises policy interpretation, governance, surveillance, supervisory exposure, and implementation readiness.",
    ownerBoosts: ["Compliance", "Risk", "Legal"],
    obligationBoosts: ["Policy & Procedures", "Governance & Accountability", "Market Conduct & Surveillance", "Training & Communications"],
    stageBoosts: ["Final Rules / Implementation", "Supervisory / Thematic Review", "Consultation / Policy Development"]
  },
  COO: {
    summary:
      "COO view surfaces operational delivery risk, outsourcing dependencies, reporting readiness, data quality, and programme mobilisation.",
    ownerBoosts: ["Operations", "Risk", "Technology", "Finance"],
    obligationBoosts: ["Reporting & MI", "Outsourcing & Third Parties", "Operational Resilience", "Operating Model"],
    stageBoosts: ["Reporting / Data Request", "Final Rules / Implementation", "Supervisory / Thematic Review"]
  },
  Legal: {
    summary:
      "Legal view emphasises perimeter changes, disclosures, interpretation questions, contractual implications, and governance obligations.",
    ownerBoosts: ["Legal", "Compliance"],
    obligationBoosts: ["Legal Entity Perimeter", "Disclosure & Communications", "Policy & Procedures", "Governance & Accountability"],
    stageBoosts: ["Legislative / Perimeter Change", "Consultation / Policy Development", "Final Rules / Implementation"]
  },
  Executive: {
    summary:
      "Executive view compresses the portfolio into material items, change themes, near-term decisions, and the few issues that need leadership attention.",
    ownerBoosts: ["Compliance", "Risk", "Operations", "Legal"],
    obligationBoosts: ["Governance & Accountability", "Operational Resilience", "Reporting & MI", "Legal Entity Perimeter"],
    stageBoosts: ["Final Rules / Implementation", "Reporting / Data Request", "Supervisory / Thematic Review"]
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

const SECTION_LOOKUP = new Map(
  KNOWN_SECTIONS.map(value => [value.toLowerCase(), value])
);

const SUBCATEGORY_LOOKUP = new Map(
  KNOWN_SUBCATEGORIES.map(value => [value.toLowerCase(), value])
);

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
  /^jan-jun \d{4}$/i,
  /^post july \d{4}$/i,
  /^e\s*formal engagement planned/i
];

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
      reason: "Touches cross-cutting governance, resilience or outsourcing controls.",
      terms: ["operational resilience", "outsourcing", "third party", "governance", "smcr", "dear ceo"]
    },
    {
      weight: 14,
      reason: "Touches regulatory reporting, data or submissions that often land on the UK entity.",
      terms: ["regulatory reporting", "reporting", "data collection", "return", "attestation", "notification"]
    },
    {
      weight: 12,
      reason: "Touches financial crime, sanctions or client onboarding controls.",
      terms: ["aml", "anti-money laundering", "sanctions", "kyc", "financial crime"]
    },
    {
      weight: 10,
      reason: "Touches sustainability, disclosures or technology change with cross-cutting UK impact.",
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
  await hydratePdfFromStorage();
  renderAll();
}

function mapEls() {
  const ids = [
    "landingVersion",
    "landingDatasetStatus",
    "landingDatasetCopy",
    "homeLastScan",
    "homeComparisonStatus",
    "utilityEnvironment",
    "utilityMode",
    "utilityViewer",
    "utilityPdfStatus",
    "utilityLastScan",
    "utilityComparison",
    "utilityEvidenceCoverage",
    "dashboardModeButtons",
    "dashboardModeLabel",
    "dashboardModeSummary",
    "headerMeta",
    "datasetInfo",
    "comparisonInfo",
    "portfolioNarrative",
    "currentRoleLabel",
    "reloadBtn",
    "exportJsonBtn",
    "exportCsvBtn",
    "exportBoardBriefBtn",
    "exportOwnerPackBtn",
    "clearStorageBtn",
    "fileInput",
    "uploadBtn",
    "uploadStatus",
    "roleViewButtons",
    "roleSummary",
    "presetButtons",
    "activePresetLabel",
    "searchInput",
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
    "pdfPreviewStatus",
    "evidenceTrail",
    "pdfPreview",
    "askInput",
    "askBtn",
    "askAnswer",
    "askResults",
    "kpiTotal",
    "kpiImmediateAction",
    "kpiHighRelevance",
    "kpiNeedsReview",
    "kpiChanges",
    "footerVersion",
    "footerUpdated",
    "footerParser",
    "footerDataset",
    "footerStorageNotice"
  ];

  ids.forEach(id => {
    els[id] = document.getElementById(id);
  });

  PROFILE_FIELDS.forEach(([, domId]) => {
    els[domId] = document.getElementById(domId);
  });
}

function bindEvents() {
  els.uploadBtn.addEventListener("click", handleUpload);
  els.reloadBtn.addEventListener("click", () => renderAll());
  els.exportJsonBtn.addEventListener("click", exportJson);
  els.exportCsvBtn.addEventListener("click", exportCsv);
  els.exportBoardBriefBtn.addEventListener("click", exportBoardBrief);
  els.exportOwnerPackBtn.addEventListener("click", exportOwnerPack);
  els.clearStorageBtn.addEventListener("click", clearSavedData);
  els.dashboardModeButtons.addEventListener("click", handleDashboardModeClick);
  els.roleViewButtons.addEventListener("click", handleRoleViewClick);
  els.presetButtons.addEventListener("click", handlePresetClick);
  document.querySelectorAll("[data-page-target]").forEach(button => {
    button.addEventListener("click", () => navigateToPage(button.dataset.pageTarget));
  });
  window.addEventListener("hashchange", syncPageFromHash);
  els.askBtn.addEventListener("click", runAskQuery);
  els.askInput.addEventListener("keydown", event => {
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
  return window.location.hash === "#review" ? "review" : "home";
}

function syncPageFromHash() {
  state.currentPage = getPageFromHash();
  renderPageState();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function navigateToPage(page) {
  if (!["home", "review"].includes(page)) return;

  const nextHash = `#${page}`;
  if (window.location.hash !== nextHash) {
    window.location.hash = nextHash;
    return;
  }

  state.currentPage = page;
  renderPageState();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function handleDashboardModeClick(event) {
  const button = event.target.closest("[data-dashboard-mode]");
  if (!button) return;

  state.dashboardMode = button.dataset.dashboardMode;
  renderAll();
}

function handleRoleViewClick(event) {
  const button = event.target.closest("[data-role-view]");
  if (!button) return;

  state.roleView = button.dataset.roleView;
  renderRoleButtons();
  renderAll();
}

function handlePresetClick(event) {
  const button = event.target.closest("[data-preset]");
  if (!button) return;

  state.activePreset = button.dataset.preset;
  renderPresetButtons();
  applyFilters();
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
      "Dataset is too large for browser storage. The current session will still work, but persistence may be limited.";
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
      els.uploadStatus.textContent = "Saved browser data cleared.";
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

  state.profile = mergeProfile(nextProfile);
  localStorage.setItem(PROFILE_KEY, JSON.stringify(state.profile));
  els.uploadStatus.textContent = "FMRUK profile saved. Re-scoring the current dataset.";
  reanalyseCurrentDataset();
}

function resetProfile() {
  state.profile = structuredClone(DEFAULT_PROFILE);
  localStorage.setItem(PROFILE_KEY, JSON.stringify(state.profile));
  renderProfileForm();
  els.uploadStatus.textContent = "FMRUK profile reset to the default model.";
  reanalyseCurrentDataset();
}

function renderProfileForm() {
  PROFILE_FIELDS.forEach(([field, domId]) => {
    if (!els[domId]) return;
    els[domId].value = (state.profile?.[field] || []).join(", ");
  });
}

function mergeProfile(profile) {
  const merged = structuredClone(DEFAULT_PROFILE);
  PROFILE_FIELDS.forEach(([field]) => {
    merged[field] = Array.isArray(profile?.[field])
      ? profile[field].map(normaliseWs).filter(Boolean)
      : structuredClone(DEFAULT_PROFILE[field]);
  });
  return merged;
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
  const file = els.fileInput.files[0];
  if (!file) {
    els.uploadStatus.textContent = "Please select the FCA initiatives PDF first.";
    return;
  }

  const ext = file.name.split(".").pop().toLowerCase();
  if (ext !== "pdf") {
    els.uploadStatus.textContent = "PDF-only mode is enabled.";
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
    const parsedItems = await parsePdfFile(buffer, file.name);
    const deduped = dedupeItems(parsedItems);

    if (!deduped.length) {
      els.uploadStatus.textContent =
        "No initiatives were detected in the PDF. Review the source formatting or parser assumptions.";
      return;
    }

    await saveStoredPdf(buffer);
    state.pdfBuffer = buffer;
    state.pdfDocument = await loadPdfDocumentFromBuffer(buffer);

    els.uploadStatus.textContent = `Analysing ${deduped.length} extracted initiatives...`;
    const analysed = analyseRows(deduped);
    const comparison = compareWithPreviousDataset(previousItems, analysed);

    state.baseItems = deduped;
    state.raw = sortItems(comparison.items);
    state.datasetMeta = {
      fileName: file.name,
      uploadedAt: new Date().toISOString(),
      rowCount: state.raw.length,
      fileType: "PDF",
      parserVersion: "v8",
      previousFileName: previousMeta?.fileName || "",
      comparisonSummary: comparison.summary
    };
    state.selectedItemId = state.raw[0]?.id || null;
    state.currentPdfHighlightKey = "";

    saveToStorage();
    els.uploadStatus.textContent = `Loaded ${state.raw.length} initiatives from ${file.name}.`;
    renderAll();
  } catch (err) {
    console.error(err);
    els.uploadStatus.textContent = `Upload failed. ${err.message || "The PDF could not be parsed."}`;
  }
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
    const nextTitle = cleanTitleCandidate(cellText(nextRow, "initiative") || nextRow.text);
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
  const directLead = normaliseLeadToken(cellText(row, "lead"));
  const directTitle = cleanTitleCandidate(cellText(row, "initiative"));

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

  const embedded = splitLeadAndTitle(cellText(row, "initiative"));
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

  const initiativeCell = stripRepeatedTitle(cellText(row, "initiative"), accumulator.initiativeTitle);
  const overflowText = normaliseWs(cellText(row, "overflow"));
  const leadResidue = stripLeadPrefix(cellText(row, "lead"), accumulator.leadRegulator);

  if (opts.isStartRow) {
    if (leadResidue) {
      appendUnique(accumulator.initiativeDescriptionParts, leadResidue);
      recordEvidence(accumulator, "description", row, "lead", leadResidue);
    }
    if (initiativeCell && !shouldTreatAsTitleContinuation(accumulator, initiativeCell, row)) {
      appendUnique(accumulator.initiativeDescriptionParts, initiativeCell);
      recordEvidence(accumulator, "description", row, "initiative", initiativeCell);
    }
  } else if (opts.consumeAsTitle && initiativeCell) {
    accumulator.initiativeTitle = normaliseWs(`${accumulator.initiativeTitle} ${initiativeCell}`);
    recordEvidence(accumulator, "title", row, "initiative", initiativeCell);
  } else if (initiativeCell) {
    if (shouldTreatAsTitleContinuation(accumulator, initiativeCell, row)) {
      accumulator.initiativeTitle = normaliseWs(`${accumulator.initiativeTitle} ${initiativeCell}`);
      recordEvidence(accumulator, "title", row, "initiative", initiativeCell);
    } else {
      appendUnique(accumulator.initiativeDescriptionParts, initiativeCell);
      recordEvidence(accumulator, "description", row, "initiative", initiativeCell);
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
  const description = dedupeStrings(accumulator.initiativeDescriptionParts).join(" ");
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

  items.forEach(item => {
    const key = item.canonicalKey || buildCanonicalKey(item);
    if (seen.has(key)) return;
    seen.add(key);
    output.push(item);
  });

  return output;
}

function analyseRows(items) {
  return items.map(item => {
    const classification = classifyItem(item);
    const stage = determineStage(item);
    const obligations = mapObligations(item, classification, stage);
    const impact = determineImpactLevel(item, classification, stage, obligations);
    const relevance = evaluateFmrukRelevance(item, classification, stage, obligations);
    const uncertainty = determineUncertainty(item, classification, stage, relevance);
    const suggestedActions = buildSuggestedActions(
      item,
      classification,
      stage,
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
    const timeline = buildTimelinePoint(item);
    const clusterLabel = determineCluster(item, classification, obligations);

    let analysed = {
      ...item,
      theme: classification.theme,
      internalSubTheme: classification.subTheme,
      primaryOwner: classification.primaryOwner,
      secondaryOwner: classification.secondaryOwner,
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
      relevanceSignals: relevance.positiveReasons,
      relevanceNegativeSignals: relevance.negativeReasons,
      profileReasons: relevance.profileReasons,
      whyNotRelevant: relevance.whyNotRelevant,
      uncertaintyFlags: uncertainty.flags,
      needsReview: uncertainty.needsReview,
      potentialBusinessImpact,
      suggestedActions,
      suggestedAction: suggestedActions.join(" "),
      immediateActionRequired:
        relevance.isRelevant && impact.level === "High" && stage.id !== "monitoring",
      rationale: buildRationale(item, classification, stage, relevance, obligations),
      briefBlocks: buildBriefBlocks(
        item,
        classification,
        stage,
        obligations,
        relevance,
        suggestedActions,
        uncertainty
      ),
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
    signals: ["No strong stage markers were detected, so the item is being treated as watchlist monitoring."]
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
        reason: "Default obligation map when the initiative has not yet been tied to a more specific control domain."
      }
    ];
  }

  return scored.map(entry => ({
    name: entry.name,
    reason: entry.reasons[0] || `${entry.name} is in scope based on the source text and stage.`
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
  let score = 30;
  const positive = scoreSignalGroups(RELEVANCE_SIGNAL_GROUPS.positive, blob);
  const negative = scoreSignalGroups(RELEVANCE_SIGNAL_GROUPS.negative, blob);
  const profileMatch = scoreProfileMatch(item, state.profile);

  score += positive.score;
  score -= negative.score;
  score += profileMatch.score;
  score += RELEVANCE_SIGNAL_GROUPS.sectionAdjustments[item.sectionName] || 0;

  if (
    ["Transaction Reporting", "Best Execution", "Market Abuse / Surveillance", "Capital / Liquidity", "Operational Resilience", "Outsourcing / Third Party Risk", "SMCR / Governance", "Regulatory Reporting / Data Collection"].includes(
      classification.subTheme
    )
  ) {
    score += 8;
  }

  if (obligations.some(obligation => obligation.name === "Reporting & MI")) score += 4;
  if (obligations.some(obligation => obligation.name === "Governance & Accountability")) score += 4;
  if (stage.id === "reporting" || stage.id === "final_rules") score += 4;
  if (item.leadRegulator.includes("FCA")) score += 3;
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

  const isRelevant = score >= 58 || (score >= 50 && !negativeReasons.length);
  const band = score >= 80 ? "High" : score >= 58 ? "Medium" : "Low";
  const whyNotRelevant =
    !isRelevant && negativeReasons.length
      ? `Current view: low FMRUK relevance because ${negativeReasons.join(" ")}`
      : "";

  return {
    score,
    band,
    isRelevant,
    positiveReasons,
    negativeReasons,
    profileReasons: profileMatch.reasons,
    whyNotRelevant
  };
}

function scoreProfileMatch(item, profile) {
  const blob = buildBlob(item);
  const categories = [
    ["businessLines", 16, "Matches FMRUK business line"],
    ["permissions", 14, "Matches FMRUK permission set"],
    ["products", 12, "Matches FMRUK product scope"],
    ["distributionModel", 10, "Matches FMRUK distribution model"],
    ["outsourcedProviders", 10, "Matches outsourced provider exposure"],
    ["reportingObligations", 14, "Matches reporting obligation"],
    ["committees", 8, "Matches governance committee touchpoint"],
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

function determineUncertainty(item, classification, stage, relevance) {
  const flags = [];
  if ((item.parseConfidence || 0) < 70) flags.push("PDF extraction confidence is low.");
  if ((classification.confidence || 0) < 60) flags.push("Classification confidence is modest.");
  if (classification.ambiguity) flags.push(classification.ambiguity);
  if ((stage.confidence || 0) < 58) flags.push("Initiative stage is uncertain.");
  if (!item.expectedKeyMilestones) flags.push("Milestones were not clearly extracted.");
  if (relevance.score >= 45 && relevance.score <= 65) {
    flags.push("Relevance sits in the judgement range and needs manual review.");
  }
  if (relevance.positiveReasons.length && relevance.negativeReasons.length) {
    flags.push("The initiative shows both FMRUK and non-FMRUK relevance signals.");
  }
  return {
    flags: dedupeStrings([...flags, ...(item.parseWarnings || [])]),
    needsReview: flags.length > 0 || (item.parseWarnings || []).length > 0
  };
}

function buildPotentialBusinessImpact(item, classification, stage, obligations) {
  const obligationNames = obligations.map(obligation => obligation.name).slice(0, 3);
  const obligationText = obligationNames.length
    ? ` The main control domains in scope are ${obligationNames.join(", ")}.`
    : "";
  const stageNote = {
    consultation:
      " The main near-term need is understanding scope, shaping response and avoiding late delivery surprises.",
    final_rules:
      " The item is close to implementation, so the focus should move quickly from interpretation into delivery and control readiness.",
    supervisory:
      " The main exposure is evidencing current-state compliance and fixing weak controls before supervisory scrutiny increases.",
    reporting:
      " The practical risk sits in data quality, ownership, reconciliation and readiness to make accurate submissions.",
    legislation:
      " The key question is whether the perimeter or legal interpretation changes the entity obligations that sit underneath implementation.",
    monitoring:
      " This currently looks more like a watchlist item than an immediate delivery programme."
  }[stage.id] || "";

  return `${classification.impactStatement}${obligationText}${stageNote}`;
}

function buildSuggestedActions(item, classification, stage, impact, relevance, obligations, uncertainty) {
  const actions = [];
  const stageActions = GENERIC_STAGE_ACTIONS[stage.id] || GENERIC_STAGE_ACTIONS.monitoring;

  if (relevance.isRelevant) {
    actions.push(
      `Lead review: ${classification.primaryOwner}, with ${classification.secondaryOwner} support.`
    );
  } else {
    actions.push(
      `Keep this with ${classification.primaryOwner} on the watchlist until FMRUK scope is clearer.`
    );
  }

  actions.push(stageActions[0]);
  actions.push(classification.coreActions[0]);

  if (impact.level === "High" || stage.id === "final_rules" || stage.id === "reporting") {
    actions.push(classification.coreActions[1]);
  }

  if (obligations.length) {
    actions.push(`Focus the first review on ${obligations.slice(0, 2).map(itemObligation => itemObligation.name).join(" and ")}.`);
  }

  if (item.expectedKeyMilestones) {
    actions.push(`Track milestone text captured from the PDF: ${truncateText(item.expectedKeyMilestones, 160)}.`);
  }

  if (uncertainty.needsReview) {
    actions.push("Validate the extracted row against the source PDF before relying on the detail for delivery planning.");
  }

  return dedupeStrings(actions).slice(0, 5);
}

function buildRationale(item, classification, stage, relevance, obligations) {
  const positives = relevance.positiveReasons.length
    ? `Positive signals: ${relevance.positiveReasons.join(" ")}`
    : "No strong positive relevance signals were detected.";
  const negatives = relevance.negativeReasons.length
    ? ` Counter-signals: ${relevance.negativeReasons.join(" ")}`
    : "";
  const obligationText = obligations.length
    ? ` Obligation map: ${obligations.map(obligation => obligation.name).join(", ")}.`
    : "";

  return `Classification: ${classification.theme} / ${classification.subTheme}. Stage: ${stage.label}. ${positives}${negatives}${obligationText}`;
}

function buildBriefBlocks(item, classification, stage, obligations, relevance, actions, uncertainty) {
  const whatIsThis =
    `${stage.label} item in ${classification.theme}, led by ${item.leadRegulator || "the relevant regulator"} and listed under ${item.sectionName || "the FCA grid"}.`;
  const whatIsChanging =
    item.initiativeDescription
      ? `The source text covers ${truncateText(item.initiativeDescription, 220)}`
      : classification.impactStatement;
  const whyNow =
    item.expectedKeyMilestones
      ? `Timing is ${item.timingBucket.toLowerCase()}. Milestone reference: ${truncateText(item.expectedKeyMilestones, 190)}`
      : `Timing is ${item.timingBucket.toLowerCase()}. The item is currently tracked as ${stage.label.toLowerCase()}.`;
  const whyFmrukMayCare =
    relevance.isRelevant
      ? `Main relevance drivers: ${truncateText(relevance.positiveReasons.join(" "), 220)}`
      : item.whyNotRelevant || "No strong FMRUK relevance drivers are currently recorded.";
  const whoOwnsIt =
    `Lead owner: ${classification.primaryOwner}. Support: ${classification.secondaryOwner}. Control areas: ${obligations.map(obligation => obligation.name).join(", ")}.`;
  const whatNeedsToHappenNext = actions.slice(0, 3).join(" ");
  const whatIsStillUnclear = uncertainty.flags.length
    ? uncertainty.flags.join(" ")
    : "No material open points are currently flagged.";

  return [
    {
      title: "What this is",
      copy: whatIsThis,
      evidenceFields: ["title", "lead", "general"]
    },
    {
      title: "What is changing",
      copy: whatIsChanging,
      evidenceFields: ["description", "impact", "general"]
    },
    {
      title: "Why now",
      copy: whyNow,
      evidenceFields: ["milestones", "timing", "isNew"]
    },
    {
      title: "Why FMRUK may care",
      copy: whyFmrukMayCare,
      evidenceFields: ["description", "milestones", "general"]
    },
    {
      title: "Who owns it",
      copy: whoOwnsIt,
      evidenceFields: ["title", "description"]
    },
    {
      title: "What needs to happen next",
      copy: whatNeedsToHappenNext,
      evidenceFields: ["milestones", "description", "general"]
    },
    {
      title: "What is still unclear",
      copy: whatIsStillUnclear,
      evidenceFields: ["general", "milestones", "description"]
    }
  ];
}

function buildTimelinePoint(item) {
  const match = parseTimelineDate(item.expectedKeyMilestones || item.rawText || "");
  return {
    label: match?.label || item.timingBucket,
    sortValue: match?.dateValue || estimateTimingBucketDate(item.timingBucket),
    raw: item.expectedKeyMilestones || item.timingBucket || "To Be Confirmed"
  };
}

function determineCluster(item, classification, obligations) {
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
    next.whyNotRelevant = "Analyst review has marked this initiative as out of scope for FMRUK at present.";
  }

  if (feedback.markUrgent) {
    next.impactLevel = "High";
    next.impactScore = Math.max(next.impactScore || 0, 82);
    next.immediateActionRequired = true;
  }

  if (feedback.markParseIssue) {
    next.parseConfidence = Math.max((next.parseConfidence || 70) - 18, 20);
    next.parseConfidenceBand = next.parseConfidence >= 85 ? "High" : next.parseConfidence >= 70 ? "Medium" : "Low";
    next.parseWarnings = dedupeStrings([
      ...(next.parseWarnings || []),
      "Analyst feedback indicates a parsing issue."
    ]);
  }

  if (feedback.note) {
    next.uncertaintyFlags = dedupeStrings([
      ...(next.uncertaintyFlags || []),
      `Analyst note: ${feedback.note}`
    ]);
  }

  next.needsReview = Boolean(
    next.needsReview ||
      feedback.markParseIssue ||
      feedback.note ||
      feedback.themeOverride ||
      feedback.ownerOverride
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
      return {
        ...item,
        changeStatus: "New",
        changeNarrative: "New initiative compared with the previous upload."
      };
    }

    const deltaType = determineDeltaType(previous, item);
    if (deltaType === "Accelerated") acceleratedCount += 1;
    else if (deltaType === "Delayed") delayedCount += 1;
    else if (deltaType === "Changed") changedCount += 1;

    return {
      ...item,
      changeStatus: deltaType,
      changeNarrative: buildChangeNarrative(previous, item, deltaType)
    };
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
  if (deltaType === "New") return "New initiative compared with the previous upload.";
  if (deltaType === "Accelerated") return "Milestone timing is earlier than in the previous upload.";
  if (deltaType === "Delayed") return "Milestone timing is later than in the previous upload.";
  if (deltaType === "Changed") return "Description, staging or triage outcome differs from the previous upload.";
  return "No material delta detected against the previous upload.";
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
        changeNarrative: item.changeNarrative || "No comparison note is currently available."
      }
    ])
  );

  state.raw = sortItems(
    analyseRows(state.baseItems).map(item => ({
      ...item,
      ...(deltaMap.get(item.canonicalKey) || {
        changeStatus: "Existing",
        changeNarrative: "No comparison note is currently available."
      })
    }))
  );

  saveToStorage();
  renderAll();
}

function renderAll() {
  renderPageState();
  renderDashboardModeButtons();
  updateMeta();
  renderRoleButtons();
  renderPresetButtons();
  populateFilters();
  applyFilters();
}

function renderPageState() {
  document.body.dataset.page = state.currentPage;
  document.querySelectorAll(".page-switch[data-page-target]").forEach(button => {
    button.classList.toggle("is-current", button.dataset.pageTarget === state.currentPage);
  });
}

function updateMeta() {
  const roleView = ROLE_VIEWS[state.roleView];
  const dashboardMode = DASHBOARD_MODES[state.dashboardMode] || DASHBOARD_MODES.workspace;
  els.currentRoleLabel.textContent = `${state.roleView} lens active`;
  els.roleSummary.textContent = roleView.summary;
  els.utilityEnvironment.textContent = APP_ENVIRONMENT_LABEL;
  els.utilityMode.textContent = dashboardMode.label;
  els.utilityViewer.textContent = state.roleView;
  els.footerVersion.textContent = APP_VERSION;
  els.footerUpdated.textContent = APP_UPDATED_AT;
  els.landingVersion.textContent = `Version ${APP_VERSION}`;
  els.footerStorageNotice.textContent =
    "The workspace stores the current dataset, evidence trail and review settings locally in this browser.";

  if (!state.datasetMeta) {
    els.headerMeta.textContent = "No dataset loaded";
    els.datasetInfo.textContent = "No saved dataset found.";
    els.comparisonInfo.textContent = "Upload a PDF to establish the first comparison baseline.";
    els.landingDatasetStatus.textContent = "Awaiting PDF upload";
    els.landingDatasetCopy.textContent =
      "Open the review workspace to load a PDF and generate the current portfolio view.";
    els.homeLastScan.textContent = "Not run";
    els.homeComparisonStatus.textContent = "No baseline";
    els.utilityPdfStatus.textContent = "Awaiting upload";
    els.utilityLastScan.textContent = "Not run";
    els.utilityComparison.textContent = "No baseline";
    els.utilityEvidenceCoverage.textContent = "0%";
    els.footerParser.textContent = "Awaiting dataset";
    els.footerDataset.textContent = "No PDF loaded";
    return;
  }

  els.headerMeta.textContent = `${state.datasetMeta.fileName} | ${state.datasetMeta.rowCount} initiatives | uploaded ${formatDate(state.datasetMeta.uploadedAt)}`;
  els.datasetInfo.textContent = `Stored in this browser only. File type: ${state.datasetMeta.fileType}. Parser: ${state.datasetMeta.parserVersion}.`;
  els.landingDatasetStatus.textContent = `${state.datasetMeta.rowCount} initiatives in scope`;
  els.utilityPdfStatus.textContent = truncateText(state.datasetMeta.fileName, 34);
  els.homeLastScan.textContent = formatDate(state.datasetMeta.uploadedAt);
  els.utilityLastScan.textContent = formatDate(state.datasetMeta.uploadedAt);
  els.utilityEvidenceCoverage.textContent = formatPercent(calculateEvidenceCoverage(state.raw));
  els.footerParser.textContent = state.datasetMeta.parserVersion || "Current parser";
  els.footerDataset.textContent = `${state.datasetMeta.fileName} | ${state.datasetMeta.rowCount} items`;

  const summary = state.datasetMeta.comparisonSummary;
  if (summary) {
    const comparisonBits = [
      `${summary.newCount} new`,
      `${summary.changedCount} changed`,
      `${summary.acceleratedCount} accelerated`,
      `${summary.delayedCount} delayed`,
      `${summary.removedCount} removed`
    ];
    const movedCount = state.raw.filter(item => item.changeStatus !== "Existing").length;
    els.comparisonInfo.textContent = `Comparison summary: ${comparisonBits.join(", ")}.`;
    els.homeComparisonStatus.textContent = `${movedCount} items moved`;
    els.utilityComparison.textContent = `${movedCount} items moved`;
    els.landingDatasetCopy.textContent =
      `Latest scan uploaded ${formatDate(state.datasetMeta.uploadedAt)}. Comparison baseline available with ${comparisonBits.join(", ")}.`;
  } else {
    els.comparisonInfo.textContent = "No comparison baseline is available yet.";
    els.homeComparisonStatus.textContent = "Baseline not set";
    els.utilityComparison.textContent = "Baseline not set";
    els.landingDatasetCopy.textContent =
      `Latest scan uploaded ${formatDate(state.datasetMeta.uploadedAt)}. No comparison baseline is available yet.`;
  }
}

function renderDashboardModeButtons() {
  const mode = DASHBOARD_MODES[state.dashboardMode] || DASHBOARD_MODES.workspace;
  document.body.dataset.dashboardMode = state.dashboardMode;
  els.dashboardModeLabel.textContent = mode.label;
  els.dashboardModeSummary.textContent = mode.summary;

  const buttons = els.dashboardModeButtons.querySelectorAll("[data-dashboard-mode]");
  buttons.forEach(button => {
    button.classList.toggle("active", button.dataset.dashboardMode === state.dashboardMode);
  });
}

function renderRoleButtons() {
  const buttons = els.roleViewButtons.querySelectorAll("[data-role-view]");
  buttons.forEach(button => {
    button.classList.toggle("active", button.dataset.roleView === state.roleView);
  });
}

function renderPresetButtons() {
  const buttons = els.presetButtons.querySelectorAll("[data-preset]");
  buttons.forEach(button => {
    button.classList.toggle("active", button.dataset.preset === state.activePreset);
  });
  els.activePresetLabel.textContent = `Preset: ${presetLabel(state.activePreset)}`;
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
      item.potentialBusinessImpact,
      item.rationale,
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
      (!needsReviewOnly || item.needsReview === true) &&
      matchesActivePreset(item)
    );
  });

  state.filtered = sortItems(items, state.roleView);

  if (!state.filtered.find(item => item.id === state.selectedItemId)) {
    state.selectedItemId = state.filtered[0]?.id || null;
  }

  renderSummary(state.filtered);
  renderInitiativeList(state.filtered);
  const selectedItem = state.filtered.find(item => item.id === state.selectedItemId);
  renderDetail(selectedItem);
  renderTimelineList(state.filtered);
  renderPdfEvidence(selectedItem);
  renderAskResults();
}

function matchesActivePreset(item) {
  switch (state.activePreset) {
    case "high_priority":
      return item.immediateActionRequired || (item.relevanceScore || 0) >= 80 || item.impactLevel === "High";
    case "needs_review":
      return item.needsReview === true;
    case "delta":
      return item.changeStatus !== "Existing";
    case "reporting":
      return item.obligations?.some(obligation => obligation.name === "Reporting & MI");
    default:
      return true;
  }
}

function sortItems(items, roleView = state.roleView) {
  const role = ROLE_VIEWS[roleView] || ROLE_VIEWS.Compliance;
  return [...items].sort((a, b) => {
    const roleScoreA = rolePriorityScore(a, role);
    const roleScoreB = rolePriorityScore(b, role);
    if (roleScoreB !== roleScoreA) return roleScoreB - roleScoreA;
    if ((b.immediateActionRequired ? 1 : 0) !== (a.immediateActionRequired ? 1 : 0)) {
      return (b.immediateActionRequired ? 1 : 0) - (a.immediateActionRequired ? 1 : 0);
    }
    if ((b.relevanceScore || 0) !== (a.relevanceScore || 0)) {
      return (b.relevanceScore || 0) - (a.relevanceScore || 0);
    }
    return (b.impactScore || 0) - (a.impactScore || 0);
  });
}

function rolePriorityScore(item, role) {
  let score = 0;
  if (role.ownerBoosts.includes(item.primaryOwner)) score += 12;
  if ((item.obligations || []).some(obligation => role.obligationBoosts.includes(obligation.name))) score += 10;
  if (role.stageBoosts.includes(item.stageLabel)) score += 8;
  if (state.roleView === "Executive" && item.changeStatus !== "Existing") score += 8;
  if (item.immediateActionRequired) score += 12;
  if (item.needsReview) score += 4;
  return score;
}

function renderSummary(items) {
  els.kpiTotal.textContent = items.length;
  els.kpiImmediateAction.textContent = items.filter(item => item.immediateActionRequired).length;
  els.kpiHighRelevance.textContent = items.filter(item => (item.relevanceScore || 0) >= 80).length;
  els.kpiNeedsReview.textContent = items.filter(item => item.needsReview).length;
  els.kpiChanges.textContent = items.filter(item => item.changeStatus !== "Existing").length;
  els.listMeta.textContent = `${items.length} items in scope`;
  els.portfolioNarrative.textContent = buildPortfolioNarrative(items);
  renderClusterList(items);
  renderDeltaList(items);
}

function buildPortfolioNarrative(items) {
  if (!items.length) {
    return "No initiatives match the current filters.";
  }

  const topThemes = topCounts(items.map(item => item.theme), 3)
    .map(entry => entry.name)
    .join(", ");
  const topOwners = topCounts(items.map(item => item.primaryOwner), 2)
    .map(entry => entry.name)
    .join(" and ");
  const highPriority = items.filter(item => item.immediateActionRequired).length;
  const delta = items.filter(item => item.changeStatus !== "Existing").length;
  const roleView = ROLE_VIEWS[state.roleView];

  return `${roleView.summary} ${items.length} initiatives are in scope. ${highPriority} require immediate action and ${delta} show movement against the previous upload. Leading themes: ${topThemes || "none recorded"}. Main owners: ${topOwners || "no clear ownership pattern"}.`;
}

function renderClusterList(items) {
  const clusters = topCounts(items.map(item => item.clusterLabel), 6);
  els.clusterList.innerHTML = "";

  if (!clusters.length) {
    els.clusterList.innerHTML = "<li>No cluster summary is available yet.</li>";
    return;
  }

  clusters.forEach(cluster => {
    const li = document.createElement("li");
    li.innerHTML = `<strong>${escapeHtml(cluster.name)}</strong><br />${cluster.count} initiatives are grouped in this programme.`;
    els.clusterList.appendChild(li);
  });
}

function renderDeltaList(items) {
  const summary = state.datasetMeta?.comparisonSummary;
  els.deltaList.innerHTML = "";

  if (!summary) {
    els.deltaList.innerHTML = "<li>No comparison baseline is available yet.</li>";
    return;
  }

  const lines = [
    `${summary.newCount} initiatives are new in the latest upload.`,
    `${summary.changedCount} initiatives changed in wording, stage or triage outcome.`,
    `${summary.acceleratedCount} initiatives have earlier timing than the previous upload.`,
    `${summary.delayedCount} initiatives have later timing than the previous upload.`,
    summary.removedItems?.length
      ? `Removed from the latest upload: ${summary.removedItems.join(", ")}.`
      : `${summary.removedCount} initiatives no longer appear in the latest upload.`
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

  const nextItems = items
    .filter(item => Number.isFinite(item.timeline?.sortValue))
    .sort((a, b) => a.timeline.sortValue - b.timeline.sortValue)
    .slice(0, 8);

  if (!nextItems.length) {
    els.timelineList.innerHTML = "<li>No milestone timeline is available for the current filter set.</li>";
    return;
  }

  nextItems.forEach(item => {
    const li = document.createElement("li");
    li.innerHTML = `
      <span class="timeline-date">${escapeHtml(item.timeline.label || item.timingBucket)}</span>
      <strong>${escapeHtml(item.initiativeTitle)}</strong><br />
      ${escapeHtml(truncateText(item.expectedKeyMilestones || item.stageLabel, 160))}
    `;
    els.timelineList.appendChild(li);
  });
}

function renderInitiativeList(items) {
  els.initiativeList.innerHTML = "";

  if (!items.length) {
    els.initiativeList.innerHTML = `<div class="empty-list">No initiatives match the current filters.</div>`;
    return;
  }

  items.forEach(item => {
    const card = document.createElement("article");
    card.className = "initiative-card";
    if (item.id === state.selectedItemId) card.classList.add("selected");

    card.innerHTML = `
      <div class="initiative-topline">
        <div class="chip-row">
          ${chipHtml(item.relevanceBand, item.relevanceBand)}
          ${chipHtml(item.stageLabel, stageChipClass(item.stageLabel))}
          ${chipHtml(item.changeStatus, changeChipClass(item.changeStatus))}
        </div>
        <span class="section-note">${escapeHtml(item.primaryOwner)}</span>
      </div>
      <h3 class="initiative-title">${escapeHtml(item.initiativeTitle)}</h3>
      <p class="initiative-copy">${escapeHtml(truncateText(item.briefBlocks?.[1]?.copy || item.initiativeDescription, 190))}</p>
      <div class="chip-row">
        ${chipHtml(`Parse ${item.parseConfidence || 0}`, parseChipClass(item.parseConfidence || 0))}
        ${chipHtml(item.clusterLabel, "chip-neutral")}
        ${item.immediateActionRequired ? chipHtml("Immediate action", "chip-high") : ""}
      </div>
      <div class="chip-row">
        ${(item.obligations || []).slice(0, 3).map(obligation => chipHtml(obligation.name, "chip-neutral")).join("")}
      </div>
    `;

    card.addEventListener("click", () => {
      state.selectedItemId = item.id;
      renderInitiativeList(state.filtered);
      renderDetail(item);
      renderPdfEvidence(item);
    });

    els.initiativeList.appendChild(card);
  });
}

function renderDetail(item) {
  if (!item) {
    els.detailMeta.textContent = "No initiative selected";
    els.detailPanel.innerHTML =
      "Upload the FCA Grid PDF, then select an initiative to review the briefing note.";
    return;
  }

  const sourcePages = item.sourcePages?.length ? item.sourcePages.join(", ") : "N/A";
  els.detailMeta.textContent = `${item.primaryOwner} lead | ${item.stageLabel} | pages ${sourcePages}`;
  const feedback = state.feedback[item.canonicalKey] || {};
  const evidenceCount = flattenEvidence(item.evidence).length;

  const heroChips = [
    chipHtml(`${item.relevanceScore} relevance`, chipClassForBand(item.relevanceBand)),
    chipHtml(item.impactLevel, chipClassForBand(item.impactLevel)),
    chipHtml(`${item.parseConfidence} parse`, parseChipClass(item.parseConfidence)),
    chipHtml(item.changeStatus, changeChipClass(item.changeStatus)),
    item.isFmrukRelevant ? chipHtml("FMRUK relevant", "chip-high") : chipHtml("Low FMRUK relevance", "chip-low")
  ].join("");

  const briefCards = item.briefBlocks
    .map(block => renderBriefCard(item, block))
    .join("");

  const decisionStrip = [
    buildDecisionStripCard(
      "Lead and ownership",
      `${item.primaryOwner || "TBC"} lead with ${item.secondaryOwner || "no secondary owner recorded"} support.`
    ),
    buildDecisionStripCard(
      "Timing and movement",
      `${item.timeline?.label || item.timingBucket || "Timing to be confirmed"} | ${item.changeStatus || "Existing"}`
    ),
    buildDecisionStripCard(
      "Evidence basis",
      `Pages ${sourcePages} | Parse confidence ${item.parseConfidence || 0} | ${evidenceCount} evidence anchor${evidenceCount === 1 ? "" : "s"}`
    )
  ].join("");

  const decisionCards = [
    buildDecisionCard(
      "FMRUK read-through",
      item.briefBlocks?.[3]?.copy || item.rationale || "No FMRUK read-through is currently recorded."
    ),
    buildDecisionCard("Decision required", buildDecisionRequiredCopy(item)),
    buildDecisionCard("Governance route", buildGovernanceRoute(item))
  ].join("");

  const metaCards = [
    ["Section", item.sectionName || "N/A"],
    ["Subcategory", item.subcategory || "N/A"],
    ["Lead regulator", item.leadRegulator || "N/A"],
    ["Primary owner", item.primaryOwner || "N/A"],
    ["Secondary owner", item.secondaryOwner || "N/A"],
    ["Timing", item.timeline?.label || item.timingBucket || "N/A"],
    ["Source pages", sourcePages],
    ["Change note", item.changeNarrative || "N/A"]
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

  const uncertaintyList = item.uncertaintyFlags?.length
    ? `<ul class="detail-list">${item.uncertaintyFlags.map(flag => `<li>${escapeHtml(flag)}</li>`).join("")}</ul>`
    : `<p>No material review points are currently flagged.</p>`;

  const whyNotRelevant = item.whyNotRelevant
    ? `<div class="support-card"><h4>Current relevance view</h4><p>${escapeHtml(item.whyNotRelevant)}</p></div>`
    : "";

  els.detailPanel.innerHTML = `
    <section class="detail-hero">
      <div class="chip-row">${heroChips}</div>
      <h3>${escapeHtml(item.initiativeTitle)}</h3>
      <p class="hero-copy">${escapeHtml(item.briefBlocks?.[0]?.copy || item.rationale)}</p>
    </section>

    <section class="decision-strip">${decisionStrip}</section>

    <section class="decision-note-grid">${decisionCards}</section>

    <section class="support-card">
      <div class="detail-section-heading">
        <h4>Structured brief</h4>
        <span class="section-note">Evidence-linked summary</span>
      </div>
      <div class="brief-grid">${briefCards}</div>
    </section>

    <section class="detail-meta-grid">${metaCards}</section>

    <section class="support-card">
      <h4>Impacted domains</h4>
      <div class="obligation-grid">
        ${(item.obligations || []).map(obligation => chipHtml(obligation.name, "chip-neutral")).join("")}
      </div>
      <p class="support-copy">${escapeHtml(item.potentialBusinessImpact || "N/A")}</p>
    </section>

    <section class="support-card">
      <h4>Outstanding checks</h4>
      ${uncertaintyList}
    </section>

    <section class="support-card">
      <h4>Basis for prioritisation</h4>
      <div class="chip-column">
        <div class="list-caption">Positive signals</div>
        <div class="signal-grid">${renderSignalSet(item.relevanceSignals, "chip-high")}</div>
        <div class="list-caption">Counter-signals</div>
        <div class="signal-grid">${renderSignalSet(item.relevanceNegativeSignals, "chip-low")}</div>
        <div class="list-caption">Profile matches</div>
        <div class="signal-grid">${renderSignalSet(item.profileReasons, "chip-neutral")}</div>
      </div>
    </section>

    ${whyNotRelevant}

    <section class="feedback-panel">
      <h4>Analyst review controls</h4>
      <p class="feedback-note">
        Record urgency, parsing issues, ownership changes, classification changes and review notes. Saved review decisions are reused when the same initiative appears again.
      </p>
      <div class="feedback-actions">
        <button class="feedback-chip ${feedback.markUrgent ? "active" : ""}" id="feedbackUrgentBtn" type="button">Mark Urgent</button>
        <button class="feedback-chip ${feedback.markNotRelevant ? "active" : ""}" id="feedbackNotRelevantBtn" type="button">Mark Not Relevant</button>
        <button class="feedback-chip ${feedback.markParseIssue ? "active" : ""}" id="feedbackParseBtn" type="button">Parsing Issue</button>
        <button class="feedback-chip" id="clearFeedbackBtn" type="button">Clear Feedback</button>
      </div>
      <div class="filter-grid">
        <label>
          <span>Theme Override</span>
          <select id="themeOverrideSelect">${buildThemeOptions(item.theme, feedback.themeOverride)}</select>
        </label>
        <label>
          <span>Owner Override</span>
          <select id="ownerOverrideSelect">${buildOwnerOptions(item.primaryOwner, feedback.ownerOverride)}</select>
        </label>
      </div>
      <label>
        <span>Analyst Note</span>
        <textarea id="feedbackNoteInput" rows="3">${escapeHtml(feedback.note || "")}</textarea>
      </label>
      <div class="inline-actions">
        <button class="secondary-btn" id="saveFeedbackBtn" type="button">Save Feedback</button>
      </div>
    </section>
  `;

  attachDetailEvents(item);
}

function buildDecisionStripCard(label, copy) {
  return `
    <article class="decision-card">
      <div class="meta-label">${escapeHtml(label)}</div>
      <p class="decision-note-copy">${escapeHtml(copy)}</p>
    </article>
  `;
}

function buildDecisionCard(title, copy) {
  return `
    <article class="decision-card">
      <h4>${escapeHtml(title)}</h4>
      <p class="decision-note-copy">${escapeHtml(copy)}</p>
    </article>
  `;
}

function buildDecisionRequiredCopy(item) {
  if (item.immediateActionRequired) {
    return `Immediate triage is recommended. Confirm applicability, assign delivery ownership under ${item.primaryOwner || "the lead owner"}, and set an implementation response against the next milestone.`;
  }

  if (item.needsReview) {
    return `Manual review is recommended before formal mobilisation. Confirm scope, ownership, evidence quality and timing against the source text.`;
  }

  if (!item.isFmrukRelevant) {
    return `No immediate mobilisation is indicated on the current read-through. Keep the item visible for monitoring in case timing, scope or business impact changes.`;
  }

  return `Keep the item in active review with ${item.primaryOwner || "the lead owner"}. Confirm whether a policy, governance or delivery response is needed before the next milestone.`;
}

function buildGovernanceRoute(item) {
  const committee = selectGovernanceCommittee(item);

  if (item.immediateActionRequired) {
    return `Take the item through ${committee} with ${item.primaryOwner || "the lead owner"} as sponsor and provide an executive update if implementation risk or timing escalates.`;
  }

  if (item.stageLabel === "Reporting / Data Request") {
    return `Route through ${committee} and the relevant delivery forum so reporting readiness, control ownership and resourcing can be confirmed early.`;
  }

  if (item.needsReview) {
    return `Keep the item within ${committee} until scope, relevance and ownership are confirmed. Escalate more widely if the review identifies a material change for FMRUK.`;
  }

  return `Maintain oversight through ${committee}, led by ${item.primaryOwner || "the primary owner"}, and revisit escalation if the initiative moves closer or broadens in scope.`;
}

function selectGovernanceCommittee(item) {
  const committees = state.profile?.committees?.length
    ? state.profile.committees
    : DEFAULT_PROFILE.committees;

  const lowerCommittees = committees.map(value => String(value || "").trim()).filter(Boolean);
  const hasObligation = name =>
    item.obligations?.some(obligation => obligation.name === name);

  const choose = hint =>
    lowerCommittees.find(entry => entry.toLowerCase().includes(hint));

  if (item.immediateActionRequired) {
    return choose("board") || choose("risk") || choose("compliance") || lowerCommittees[0] || "the relevant governance forum";
  }

  if (hasObligation("Operational Resilience") || hasObligation("Outsourcing & Third Parties")) {
    return choose("risk") || choose("oversight") || lowerCommittees[0] || "the relevant governance forum";
  }

  if (hasObligation("Disclosure & Communications") || item.primaryOwner === "Legal") {
    return choose("compliance") || choose("board") || lowerCommittees[0] || "the relevant governance forum";
  }

  return choose("compliance") || choose("oversight") || lowerCommittees[0] || "the relevant governance forum";
}

function renderBriefCard(item, block) {
  const evidence = getEvidenceForFields(item, block.evidenceFields, 3);
  return `
    <article class="brief-card">
      <h4>${escapeHtml(block.title)}</h4>
      <p class="brief-copy">${escapeHtml(block.copy)}</p>
      <div class="brief-evidence">
        <div class="chip-row">
          ${evidence.length ? evidence.map(entry => evidenceButtonHtml(entry)).join("") : chipHtml("No direct evidence link", "chip-neutral")}
        </div>
      </div>
    </article>
  `;
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
  const note = normaliseWs(document.getElementById("feedbackNoteInput")?.value || "");

  const next = {
    ...current,
    themeOverride: themeOverride || "",
    ownerOverride: ownerOverride || "",
    note
  };

  state.feedback[item.canonicalKey] = next;
  saveFeedback();
  els.uploadStatus.textContent = "Analyst feedback saved and applied to the dataset.";
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
    els.pdfPreviewStatus.textContent = "Select an initiative to review its linked evidence.";
    return;
  }

  const evidence = getEvidenceForFields(item, ["title", "lead", "description", "milestones", "impact", "timing", "isNew"], 20);
  if (!evidence.length) {
    els.pdfPreviewStatus.textContent = "No evidence references were captured for the selected initiative.";
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
    button.textContent = `P${entry.pageNumber} ${fieldLabel(entry.field)}: ${truncateText(entry.excerpt, 52)}`;
    button.addEventListener("click", () => {
      state.currentPdfHighlightKey = entry.key;
      renderPdfEvidence(item);
    });
    els.evidenceTrail.appendChild(button);
  });

  renderPdfPreview(item, evidence).catch(err => {
    console.error(err);
    els.pdfPreviewStatus.textContent = "The PDF preview could not be rendered for this selection.";
  });
}

async function renderPdfPreview(item, evidence) {
  els.pdfPreview.innerHTML = "";

  if (!state.pdfDocument) {
    els.pdfPreviewStatus.textContent =
      "The PDF preview is not loaded in this browser session. Upload the PDF again to restore the live preview.";
    return;
  }

  const pages = item.sourcePages?.length ? item.sourcePages.slice(0, 4) : [];
  if (!pages.length) {
    els.pdfPreviewStatus.textContent = "No source pages were captured for this initiative.";
    return;
  }

  els.pdfPreviewStatus.textContent = `Showing source pages ${pages.join(", ")} with linked highlight bands.`;
  const activeEvidence = evidence.find(entry => entry.key === state.currentPdfHighlightKey);
  const token = ++state.pdfRenderToken;

  for (const pageNumber of pages) {
    const page = await state.pdfDocument.getPage(pageNumber);
    if (token !== state.pdfRenderToken) return;

    const viewport = page.getViewport({ scale: 1.1 });
    const wrapper = document.createElement("div");
    wrapper.className = "pdf-page";

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

    wrapper.appendChild(canvas);
    wrapper.appendChild(overlay);
    els.pdfPreview.appendChild(wrapper);
  }
}

function runAskQuery() {
  const query = normaliseWs(els.askInput.value);
  state.ask.query = query;

  if (!query) {
    state.ask.answer =
      "Enter a search query to scan initiatives, obligations and source evidence.";
    state.ask.results = [];
    renderAskResults();
    return;
  }

  const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
  const scored = state.raw
    .map(item => {
      const haystack = [
        item.initiativeTitle,
        item.theme,
        item.internalSubTheme,
        item.rationale,
        item.potentialBusinessImpact,
        item.whyNotRelevant,
        item.primaryOwner,
        item.stageLabel,
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
    state.ask.answer = `No initiatives matched "${query}" in the current filtered set.`;
    renderAskResults();
    return;
  }

  const topThemes = topCounts(state.ask.results.map(item => item.theme), 2)
    .map(entry => entry.name)
    .join(", ");
  const topOwners = topCounts(state.ask.results.map(item => item.primaryOwner), 2)
    .map(entry => entry.name)
    .join(" and ");

  state.ask.answer = `${state.ask.results.length} initiatives match "${query}". Main themes: ${topThemes || "none"}. Main owners: ${topOwners || "none"}.`;
  renderAskResults();
}

function renderAskResults() {
  els.askAnswer.textContent = state.ask.answer || "Enter a search query to scan the filtered portfolio.";
  els.askResults.innerHTML = "";

  if (!state.ask.results.length) {
    els.askResults.innerHTML = "<li>No search results are available.</li>";
    return;
  }

  state.ask.results.forEach(item => {
    const li = document.createElement("li");
    li.innerHTML = `<strong>${escapeHtml(item.initiativeTitle)}</strong><br />${escapeHtml(
      truncateText(item.briefBlocks?.[3]?.copy || item.rationale, 180)
    )}`;
    li.addEventListener("click", () => {
      state.selectedItemId = item.id;
      renderInitiativeList(state.filtered);
      renderDetail(item);
      renderPdfEvidence(item);
    });
    els.askResults.appendChild(li);
  });
}

function exportJson() {
  if (!state.filtered.length) {
    els.uploadStatus.textContent = "There is no filtered dataset to export.";
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
    els.uploadStatus.textContent = "There is no filtered dataset to export.";
    return;
  }

  const columns = [
    "changeStatus",
    "clusterLabel",
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
    "potentialBusinessImpact",
    "rationale",
    "suggestedAction"
  ];

  const lines = [
    columns.join(","),
    ...state.filtered.map(item =>
      columns.map(column => csvEscape(item[column])).join(",")
    )
  ];

  downloadTextFile("fmruk-regulatory-intelligence.csv", lines.join("\n"), "text/csv;charset=utf-8;");
}

function exportBoardBrief() {
  const items = state.filtered.slice(0, 12);
  if (!items.length) {
    els.uploadStatus.textContent = "There are no filtered initiatives available for a board brief.";
    return;
  }

  const JsPDF = getPdfConstructor();
  if (!JsPDF) {
    els.uploadStatus.textContent = "The PDF export library did not load.";
    return;
  }

  const doc = buildReportPdf({
    kind: "board",
    title: "FMRUK Regulatory Board Brief",
    subtitle: "Regulatory developments, management actions and assurance points for FMRUK.",
    audience: `${state.roleView} lens`,
    items,
    summaryText: buildPortfolioNarrative(items)
  });

  const stamp = new Date().toISOString().slice(0, 10);
  doc.save(`fmruk-board-brief-${stamp}.pdf`);
}

function exportOwnerPack() {
  const role = state.roleView;
  const relevantItems = state.filtered.filter(item =>
    role === "Executive" ? item.immediateActionRequired || item.changeStatus !== "Existing" : true
  );

  if (!relevantItems.length) {
    els.uploadStatus.textContent = "There are no filtered initiatives available for an owner pack.";
    return;
  }

  const JsPDF = getPdfConstructor();
  if (!JsPDF) {
    els.uploadStatus.textContent = "The PDF export library did not load.";
    return;
  }

  const doc = buildReportPdf({
    kind: "owner",
    title: "FMRUK Owner Pack",
    subtitle: "Owner actions, control implications and review points by function.",
    audience: `${role} lens`,
    items: relevantItems,
    summaryText: buildPortfolioNarrative(relevantItems)
  });

  const stamp = new Date().toISOString().slice(0, 10);
  doc.save(`fmruk-owner-pack-${slugify(role)}-${stamp}.pdf`);
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
  startPdfPage(ctx, "Portfolio overview");
  addOverviewSection(ctx, config);

  if (config.kind === "board") {
    addMovementSection(ctx, config);
    addMaterialInitiativeSection(ctx, config.items);
    addPortfolioAuditSection(ctx, config.items, "Board assurance and audit points");
  } else {
    addOwnerSummarySection(ctx, config.items);
    addOwnerDetailSections(ctx, config.items);
    addPortfolioAuditSection(ctx, config.items, "Owner assurance and audit points");
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
      navy: [17, 59, 103],
      navyDeep: [13, 45, 77],
      teal: [13, 141, 134],
      gold: [178, 124, 24],
      red: [176, 67, 67],
      green: [45, 122, 76],
      text: [22, 35, 52],
      soft: [95, 111, 132],
      line: [214, 223, 233],
      panel: [245, 248, 252],
      white: [255, 255, 255]
    },
    datasetLabel: state.datasetMeta?.fileName || "Current filtered review set",
    audience: config.audience,
    activeHeader: config.title
  };
}

function drawReportCover(ctx, config) {
  const { doc, pageWidth, pageHeight, colors, margin } = ctx;

  doc.setFillColor(...colors.navyDeep);
  doc.rect(0, 0, pageWidth, 182, "F");

  doc.setFillColor(...colors.teal);
  doc.rect(0, 182, pageWidth, 6, "F");

  doc.setTextColor(...colors.white);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text("FIDELITY INVESTMENTS", margin, 56);

  doc.setFont("times", "bold");
  doc.setFontSize(28);
  doc.text(config.title, margin, 98);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(12);
  const subtitleLines = doc.splitTextToSize(config.subtitle, pageWidth - margin * 2);
  doc.text(subtitleLines, margin, 126);

  const summaryTop = 220;
  const summaryHeight = 126;
  doc.setFillColor(...colors.white);
  doc.roundedRect(margin, summaryTop, pageWidth - margin * 2, summaryHeight, 18, 18, "F");
  doc.setDrawColor(...colors.line);
  doc.roundedRect(margin, summaryTop, pageWidth - margin * 2, summaryHeight, 18, 18, "S");

  const metaLines = [
    `Prepared: ${formatDate(new Date().toISOString())}`,
    `Source PDF: ${state.datasetMeta?.fileName || "Current browser session"}`,
    `Audience: ${config.audience}`,
    `Parsing engine: ${state.datasetMeta?.parserVersion || "Current version"}`
  ];

  doc.setTextColor(...colors.text);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("Document context", margin + 18, summaryTop + 28);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  doc.text(metaLines, margin + 18, summaryTop + 48, { lineHeightFactor: 1.6 });

  const metrics = buildPdfMetrics(config.items);
  drawPdfMetrics(
    ctx,
    metrics,
    summaryTop + summaryHeight + 22,
    110,
    2
  );

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(...colors.soft);
  doc.text(
    "Prepared from the current filtered dataset in the FMRUK Regulatory Review.",
    margin,
    pageHeight - 38
  );
}

function buildPdfMetrics(items) {
  return [
    { label: "Items in scope", value: String(items.length) },
    { label: "Immediate action", value: String(items.filter(item => item.immediateActionRequired).length) },
    { label: "Needs review", value: String(items.filter(item => item.needsReview).length) },
    { label: "Changed / moved", value: String(items.filter(item => item.changeStatus !== "Existing").length) }
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
    doc.roundedRect(x, top, columnWidth, boxHeight, 16, 16, "FD");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.setTextColor(...colors.soft);
    doc.text(metric.label.toUpperCase(), x + 16, top + 24);
    doc.setFont("times", "bold");
    doc.setFontSize(26);
    doc.setTextColor(...colors.navyDeep);
    doc.text(metric.value, x + 16, top + 66);
  });
}

function startPdfPage(ctx, label) {
  const { doc, pageWidth, margin, colors } = ctx;
  doc.addPage();
  ctx.activeHeader = label;
  ctx.y = 70;

  doc.setFillColor(...colors.panel);
  doc.rect(0, 0, pageWidth, 58, "F");
  doc.setDrawColor(...colors.line);
  doc.line(margin, 58, pageWidth - margin, 58);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(...colors.teal);
  doc.text(label.toUpperCase(), margin, 28);

  doc.setFont("times", "bold");
  doc.setFontSize(20);
  doc.setTextColor(...colors.navyDeep);
  doc.text(ctx.datasetLabel, margin, 48);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(...colors.soft);
  doc.text(ctx.audience, pageWidth - margin, 28, { align: "right" });
}

function ensurePdfSpace(ctx, neededHeight) {
  if (ctx.y + neededHeight <= ctx.pageHeight - 56) return;
  startPdfPage(ctx, ctx.activeHeader);
}

function addPdfSectionTitle(ctx, kicker, title) {
  ensurePdfSpace(ctx, 52);
  const { doc, margin, colors } = ctx;
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(...colors.teal);
  doc.text(kicker.toUpperCase(), margin, ctx.y);
  ctx.y += 18;
  doc.setFont("times", "bold");
  doc.setFontSize(22);
  doc.setTextColor(...colors.navyDeep);
  doc.text(title, margin, ctx.y);
  ctx.y += 24;
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
  addPdfSectionTitle(ctx, "Portfolio overview", "Executive summary");
  addPdfParagraph(ctx, config.summaryText, { color: ctx.colors.text, size: 11.5 });

  const topThemes = topCounts(config.items.map(item => item.theme), 4);
  const clusterRows = topCounts(config.items.map(item => item.clusterLabel), 5).map(entry => [
    entry.name,
    String(entry.count)
  ]);

  addTableSection(
    ctx,
    "Theme distribution",
    ["Theme or cluster", "Items"],
    [
      ...topThemes.map(entry => [entry.name, String(entry.count)]),
      ...clusterRows
    ].slice(0, 8)
  );
}

function addMovementSection(ctx, config) {
  addPdfSectionTitle(ctx, "Change intelligence", "Movement snapshot");

  const summary = state.datasetMeta?.comparisonSummary;
  const bullets = summary
    ? [
        `${summary.newCount} initiatives are new in the current upload.`,
        `${summary.changedCount} initiatives changed in wording, stage or triage outcome.`,
        `${summary.acceleratedCount} initiatives have earlier timing and ${summary.delayedCount} have later timing than in the previous upload.`,
        `${summary.removedCount} initiatives no longer appear in the latest upload.`
      ]
    : ["No prior upload comparison is available."];

  addPdfBulletList(ctx, bullets, { color: ctx.colors.text });

  const movementRows = config.items
    .filter(item => item.changeStatus !== "Existing")
    .slice(0, 8)
    .map(item => [
      item.changeStatus,
      truncateText(item.initiativeTitle, 52),
      item.primaryOwner,
      item.timeline?.label || item.timingBucket
    ]);

  if (movementRows.length) {
    addTableSection(ctx, "Material movements", ["Status", "Initiative", "Owner", "Timing"], movementRows);
  }
}

function addMaterialInitiativeSection(ctx, items) {
  addPdfSectionTitle(ctx, "Material initiatives", "Board briefing notes");
  items.forEach((item, index) => {
    addInitiativeCard(ctx, item, index + 1, "board");
  });
}

function addOwnerSummarySection(ctx, items) {
  addPdfSectionTitle(ctx, "Owner summary", "Action coverage");
  addPdfParagraph(
    ctx,
    "This pack is grouped by primary owner so each function can review its queue, required actions, control areas and key review points.",
    { color: ctx.colors.text, size: 11.5 }
  );

  const grouped = groupBy(items, item => item.primaryOwner);
  const rows = Object.entries(grouped)
    .sort((a, b) => b[1].length - a[1].length)
    .map(([owner, ownerItems]) => [
      owner,
      String(ownerItems.length),
      String(ownerItems.filter(item => item.immediateActionRequired).length),
      String(ownerItems.filter(item => item.needsReview).length)
    ]);

  addTableSection(ctx, "Owner queue", ["Owner", "Items", "Immediate", "Needs review"], rows);
}

function addOwnerDetailSections(ctx, items) {
  const grouped = groupBy(items, item => item.primaryOwner);
  Object.entries(grouped)
    .sort((a, b) => b[1].length - a[1].length)
    .forEach(([owner, ownerItems]) => {
      addPdfSectionTitle(ctx, owner, "Owner action pack");
      addPdfParagraph(
        ctx,
        `${owner} currently owns ${ownerItems.length} initiatives in the filtered scope. The table and notes below support action tracking, review and assurance.`,
        { color: ctx.colors.text, size: 11.5 }
      );

      const trackerRows = ownerItems.slice(0, 12).map(item => [
        truncateText(item.initiativeTitle, 44),
        item.stageLabel,
        item.changeStatus,
        truncateText(item.suggestedActions?.[0] || item.suggestedAction, 66)
      ]);

      addTableSection(
        ctx,
        `${owner} action tracker`,
        ["Initiative", "Stage", "Change", "Primary next action"],
        trackerRows
      );

      ownerItems.forEach((item, index) => {
        addInitiativeCard(ctx, item, index + 1, "owner");
      });
    });
}

function addPortfolioAuditSection(ctx, items, title) {
  addPdfSectionTitle(ctx, "Audit and assurance", title);

  addPdfBulletList(ctx, buildPortfolioAuditPoints(items), {
    color: ctx.colors.text
  });

  const flaggedRows = items
    .filter(item => item.needsReview || (item.parseConfidence || 0) < 70)
    .slice(0, 10)
    .map(item => [
      truncateText(item.initiativeTitle, 46),
      String(item.parseConfidence || 0),
      item.primaryOwner,
      truncateText((item.uncertaintyFlags || [])[0] || "Needs manual review.", 70)
    ]);

  if (flaggedRows.length) {
    addTableSection(
      ctx,
      "Items needing the strongest challenge",
      ["Initiative", "Parse", "Owner", "Primary review point"],
      flaggedRows
    );
  }
}

function addMethodologySection(ctx, config) {
  addPdfSectionTitle(ctx, "Method and assumptions", "How to read this report");

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
      "This report is based on the current filtered dataset and reflects the active role lens and profile settings.",
      "Each initiative summary uses parsed FCA PDF content, linked source rows, rule-based classification and any saved analyst overrides.",
      "Audit and review points highlight parsing confidence, open items, missing milestones and scope questions that require human review.",
      ...profileSummary.slice(0, 4)
    ],
    { color: ctx.colors.text }
  );

  addPdfParagraph(
    ctx,
    config.kind === "board"
      ? "Board note: this report supports management discussion and challenge. It does not replace legal or compliance judgement."
      : "Owner note: this pack supports mobilisation, documented judgement and review by compliance, legal, risk and audit.",
    { color: ctx.colors.soft, size: 10.5 }
  );
}

function addInitiativeCard(ctx, item, index, mode) {
  const { doc, margin, contentWidth, colors } = ctx;
  const title = `${index}. ${item.initiativeTitle}`;
  const meta = `${item.theme} | ${item.stageLabel} | ${item.primaryOwner} / ${item.secondaryOwner} | ${item.timeline?.label || item.timingBucket} | pages ${item.sourcePages.join(", ")}`;
  const summary = item.briefBlocks?.[3]?.copy || item.rationale;
  const action = item.suggestedActions?.[0] || item.suggestedAction;
  const obligations = (item.obligations || []).map(obligation => obligation.name).join(", ");
  const auditPoints = buildItemAuditPoints(item, mode);

  doc.setFont("times", "bold");
  doc.setFontSize(16);
  const titleLines = doc.splitTextToSize(title, contentWidth - 30);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  const metaLines = doc.splitTextToSize(meta, contentWidth - 30);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10.5);
  const summaryLines = doc.splitTextToSize(`Summary: ${summary}`, contentWidth - 30);
  const actionLines = doc.splitTextToSize(`Action: ${action}`, contentWidth - 30);
  const obligationLines = doc.splitTextToSize(`Obligations: ${obligations || "Not yet mapped."}`, contentWidth - 30);
  const auditHeight = auditPoints.length * 15 + 8;
  const boxHeight =
    28 +
    titleLines.length * 18 +
    metaLines.length * 12 +
    summaryLines.length * 12 +
    actionLines.length * 12 +
    obligationLines.length * 12 +
    auditHeight +
    24;

  ensurePdfSpace(ctx, boxHeight + 12);

  doc.setFillColor(...colors.panel);
  doc.setDrawColor(...colors.line);
  doc.roundedRect(margin, ctx.y, contentWidth, boxHeight, 16, 16, "FD");

  let cursorY = ctx.y + 22;
  doc.setTextColor(...colors.navyDeep);
  doc.setFont("times", "bold");
  doc.setFontSize(16);
  doc.text(titleLines, margin + 16, cursorY);
  cursorY += titleLines.length * 18;

  doc.setTextColor(...colors.soft);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text(metaLines, margin + 16, cursorY);
  cursorY += metaLines.length * 12 + 8;

  doc.setTextColor(...colors.text);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10.5);
  doc.text(summaryLines, margin + 16, cursorY);
  cursorY += summaryLines.length * 12 + 8;
  doc.text(actionLines, margin + 16, cursorY);
  cursorY += actionLines.length * 12 + 8;
  doc.text(obligationLines, margin + 16, cursorY);
  cursorY += obligationLines.length * 12 + 10;

  doc.setFont("helvetica", "bold");
  doc.setTextColor(...colors.teal);
  doc.text("Audit and assurance points", margin + 16, cursorY);
  cursorY += 14;

  doc.setFont("helvetica", "normal");
  doc.setTextColor(...colors.text);
  auditPoints.forEach(point => {
    const lines = doc.splitTextToSize(`- ${point}`, contentWidth - 30);
    doc.text(lines, margin + 16, cursorY);
    cursorY += lines.length * 12 + 3;
  });

  ctx.y += boxHeight + 12;
}

function buildItemAuditPoints(item, mode) {
  const points = [
    `Source evidence captured on PDF pages ${item.sourcePages.join(", ")}.`,
    `Parse confidence is ${item.parseConfidence} (${item.parseConfidenceBand}).`,
    `Change signal is ${item.changeStatus}: ${item.changeNarrative}`,
    `Primary review points: ${(item.uncertaintyFlags || []).slice(0, 2).join(" ") || "No material review points flagged."}`
  ];

  if (mode === "owner") {
    points.push(
      `Owner check: confirm ${item.primaryOwner} has accepted the action and agrees the obligation map.`
    );
  } else {
    points.push(
      "Board check: confirm that timing, ownership and actions remain proportionate to the assessed relevance and impact."
    );
  }

  return points.slice(0, 5);
}

function buildPortfolioAuditPoints(items) {
  const summary = state.datasetMeta?.comparisonSummary;
  const reviewCount = items.filter(item => item.needsReview).length;
  const lowParseCount = items.filter(item => (item.parseConfidence || 0) < 70).length;
  const missingMilestoneCount = items.filter(item => !item.expectedKeyMilestones).length;
  const changedCount = items.filter(item => item.changeStatus !== "Existing").length;
  const urgentCount = items.filter(item => item.immediateActionRequired).length;

  return [
    `${reviewCount} initiatives currently carry manual review flags and require source validation before action is closed.`,
    `${lowParseCount} initiatives have lower parsing confidence and should be checked against the original PDF.`,
    `${missingMilestoneCount} initiatives do not yet have clear milestone extraction, creating timing uncertainty.`,
    `${changedCount} initiatives show movement compared with the previous upload and should be reviewed for business impact and priority.`,
    `${urgentCount} initiatives currently require immediate action based on impact, timing, and relevance.`,
    summary?.removedCount
      ? `${summary.removedCount} initiatives no longer appear in the latest upload and should be checked for completion, movement or removal from the grid.`
      : "No removed initiatives were detected in the current comparison set."
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

function renderSignalSet(values, className) {
  if (!values?.length) {
    return chipHtml("None", "chip-neutral");
  }
  return values.map(value => chipHtml(value, className)).join("");
}

function evidenceButtonHtml(entry) {
  return `<button class="evidence-chip" data-evidence-key="${escapeHtml(entry.key)}" type="button">P${entry.pageNumber} ${escapeHtml(fieldLabel(entry.field))}</button>`;
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
    isNew: "new flag",
    general: "source"
  }[field] || field;
}

function chipHtml(value, className) {
  return `<span class="chip ${className}">${escapeHtml(value)}</span>`;
}

function chipClassForBand(value) {
  if (value === "High") return "chip-high";
  if (value === "Medium") return "chip-medium";
  if (value === "Low") return "chip-low";
  return "chip-neutral";
}

function stageChipClass(value) {
  if (/final|reporting/i.test(value)) return "chip-high";
  if (/consultation|supervisory/i.test(value)) return "chip-medium";
  return "chip-neutral";
}

function changeChipClass(value) {
  if (value === "New" || value === "Accelerated") return "chip-high";
  if (value === "Changed" || value === "Delayed") return "chip-medium";
  return "chip-neutral";
}

function parseChipClass(score) {
  if (score >= 85) return "chip-high";
  if (score >= 70) return "chip-medium";
  return "chip-low";
}

function presetLabel(value) {
  return {
    all: "All Items",
    high_priority: "High Priority",
    needs_review: "Needs Review",
    delta: "New This Upload",
    reporting: "Reporting Items"
  }[value] || "All Items";
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
  if (isLeadToken(title)) return false;
  if (ROW_NOISE_PATTERNS.some(pattern => pattern.test(title))) return false;
  return !HEADER_TEXT_PATTERNS.some(pattern => pattern.test(title));
}

function shouldTreatAsTitleContinuation(accumulator, value, row) {
  if (!value) return false;
  if (accumulator.initiativeTitle.length > 120) return false;
  if (cellText(row, "milestones") || cellText(row, "impact") || cellText(row, "consumer")) return false;
  if (/[.:;]/.test(value)) return false;
  if (looksLikeMilestoneText(value)) return false;
  return value.length <= 90;
}

function looksLikeMilestoneText(text) {
  const value = String(text || "").toLowerCase();
  return /\b(?:q[1-4]|h[12]|january|february|march|april|may|june|july|august|september|october|november|december|202[5-9])\b/.test(value);
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
    normaliseWs(item.sectionName).toLowerCase(),
    normaliseWs(item.leadRegulator).toLowerCase(),
    normaliseWs(item.initiativeTitle).toLowerCase()
  ].join("|");
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
