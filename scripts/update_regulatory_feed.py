import json
import re
from datetime import datetime, timezone
from pathlib import Path
from hashlib import md5

import feedparser
import pandas as pd
from bs4 import BeautifulSoup


BASE_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = BASE_DIR / "data"
OUTPUT_FILE = DATA_DIR / "regulatory_updates.json"

FCA_RSS = "https://www.fca.org.uk/news/rss.xml"
PRA_RSS = "https://www.bankofengland.co.uk/rss/prudential-regulation-publications"
GRID_XLSX = "https://www.fca.org.uk/publication/corporate/regulatory-initiatives-grid-dec-2025.xlsx"


THEME_RULES = [
    {
        "theme": "Market Conduct / Trading",
        "sub_theme": "Transaction Reporting",
        "keywords": [
            "transaction reporting", "mifir reporting", "reporting obligation", "reporting fields",
            "arm", "approved reporting mechanism", "regulatory reporting"
        ],
        "impact": "May require review of transaction reporting controls, reconciliations, exception monitoring and governance.",
        "primary_owner": "Compliance",
        "secondary_owner": "Operations"
    },
    {
        "theme": "Market Conduct / Trading",
        "sub_theme": "Best Execution",
        "keywords": [
            "best execution", "execution quality", "execution policy", "rts 28", "venue analysis"
        ],
        "impact": "Could affect best execution governance, monitoring, policy wording and oversight of execution arrangements.",
        "primary_owner": "Compliance",
        "secondary_owner": "Legal"
    },
    {
        "theme": "Market Conduct / Trading",
        "sub_theme": "Market Abuse / Surveillance",
        "keywords": [
            "market abuse", "mar", "surveillance", "inside information", "personal account dealing"
        ],
        "impact": "May require updates to surveillance controls, market abuse risk assessments, training and internal escalation processes.",
        "primary_owner": "Compliance",
        "secondary_owner": "Legal"
    },
    {
        "theme": "Operational Resilience",
        "sub_theme": "Operational Resilience",
        "keywords": [
            "operational resilience", "important business service", "impact tolerance", "resilience"
        ],
        "impact": "May require review of resilience mapping, testing, scenario analysis and governance over important business services.",
        "primary_owner": "Risk",
        "secondary_owner": "Compliance"
    },
    {
        "theme": "Operational Resilience",
        "sub_theme": "Outsourcing / Third Party Risk",
        "keywords": [
            "outsourcing", "third party", "critical third party", "service provider", "vendor risk"
        ],
        "impact": "Could require updates to outsourcing registers, due diligence, contractual controls and oversight of service providers.",
        "primary_owner": "Risk",
        "secondary_owner": "Technology"
    },
    {
        "theme": "Prudential / MIFIDPRU",
        "sub_theme": "Capital / Liquidity",
        "keywords": [
            "mifidpru", "capital", "liquidity", "own funds", "icara", "prudential", "concentration risk"
        ],
        "impact": "May affect prudential assessments, capital planning, ICARA assumptions, monitoring and governance.",
        "primary_owner": "Finance",
        "secondary_owner": "Risk"
    },
    {
        "theme": "Prudential / MIFIDPRU",
        "sub_theme": "Remuneration",
        "keywords": [
            "remuneration", "bonus", "pay", "incentive", "compensation", "malus", "clawback"
        ],
        "impact": "Could require updates to remuneration frameworks, governance, documentation and control testing.",
        "primary_owner": "HR",
        "secondary_owner": "Compliance"
    },
    {
        "theme": "Consumer / Conduct",
        "sub_theme": "Consumer Duty",
        "keywords": [
            "consumer duty", "fair value", "good outcomes", "vulnerable customers", "consumer"
        ],
        "impact": "May require review of product governance, distribution oversight, client communications and conduct monitoring.",
        "primary_owner": "Compliance",
        "secondary_owner": "Product"
    },
    {
        "theme": "Consumer / Conduct",
        "sub_theme": "Complaints / Client Treatment",
        "keywords": [
            "complaints", "redress", "disclosure", "client communication", "retail investors"
        ],
        "impact": "Could affect complaints handling, customer disclosures, oversight controls and record keeping.",
        "primary_owner": "Compliance",
        "secondary_owner": "Operations"
    },
    {
        "theme": "AML / Sanctions / Financial Crime",
        "sub_theme": "AML / KYC / Sanctions",
        "keywords": [
            "aml", "anti-money laundering", "sanctions", "kyc", "financial crime", "fraud", "bribery"
        ],
        "impact": "May require review of AML controls, KYC processes, sanctions screening, escalation and training.",
        "primary_owner": "Compliance",
        "secondary_owner": "Legal"
    },
    {
        "theme": "Governance / SMCR",
        "sub_theme": "SMCR / Governance",
        "keywords": [
            "smcr", "senior managers", "certification regime", "governance", "board", "committee", "dear ceo"
        ],
        "impact": "Could affect governance frameworks, accountabilities, committee reporting and senior management oversight.",
        "primary_owner": "Compliance",
        "secondary_owner": "Company Secretariat"
    },
    {
        "theme": "ESG / Sustainability",
        "sub_theme": "Sustainability / SDR",
        "keywords": [
            "sustainability", "sdr", "greenwashing", "climate", "esg", "transition plan", "tnfd"
        ],
        "impact": "May require updates to sustainability disclosures, product communications, governance and control evidence.",
        "primary_owner": "Legal",
        "secondary_owner": "Compliance"
    },
    {
        "theme": "Data / AI / Technology",
        "sub_theme": "AI / Data / Cyber",
        "keywords": [
            "artificial intelligence", "ai", "machine learning", "cyber", "data", "model risk", "technology"
        ],
        "impact": "Could affect governance over AI use cases, data controls, cyber oversight and internal policy requirements.",
        "primary_owner": "Technology",
        "secondary_owner": "Compliance"
    }
]


FMRUK_KEYWORDS = [
    "uk", "fca", "pra", "investment firm", "asset management", "asset manager",
    "mifid", "mifidpru", "consumer duty", "transaction reporting", "market abuse",
    "best execution", "outsourcing", "operational resilience", "financial crime",
    "sanctions", "governance", "remuneration", "sustainability", "disclosure",
    "regulatory reporting", "firm", "firms"
]


def clean_html(raw: str) -> str:
    if not raw:
        return ""
    soup = BeautifulSoup(raw, "html.parser")
    text = soup.get_text(" ", strip=True)
    return re.sub(r"\s+", " ", text).strip()


def parse_date(date_str: str):
    if not date_str:
        return None
    try:
        dt = pd.to_datetime(date_str, errors="coerce", utc=True)
        if pd.isna(dt):
            return None
        return dt.date().isoformat()
    except Exception:
        return None


def stable_id(*parts: str) -> str:
    joined = "||".join([p or "" for p in parts])
    return md5(joined.encode("utf-8")).hexdigest()[:12]


def infer_type(title: str, summary: str) -> str:
    blob = f"{title} {summary}".lower()

    checks = [
        ("Policy Statement", "policy statement"),
        ("Consultation", "consultation"),
        ("Supervisory Statement", "supervisory statement"),
        ("Dear CEO", "dear ceo"),
        ("Handbook Notice", "handbook notice"),
        ("Speech", "speech"),
        ("Discussion Paper", "discussion paper"),
        ("Final Rules / Policy", "final rules")
    ]

    for label, term in checks:
        if term in blob:
            return label

    return "Publication"


def classify_item(title: str, summary: str):
    blob = f"{title} {summary}".lower()

    for rule in THEME_RULES:
      if any(keyword in blob for keyword in rule["keywords"]):
        return {
            "theme": rule["theme"],
            "sub_theme": rule["sub_theme"],
            "potential_business_impact": rule["impact"],
            "primary_owner": rule["primary_owner"],
            "secondary_owner": rule["secondary_owner"]
        }

    return {
        "theme": "General UK Regulatory Change",
        "sub_theme": "General",
        "potential_business_impact": "May require initial triage to determine whether any policy, control, governance or implementation response is needed.",
        "primary_owner": "Compliance",
        "secondary_owner": "Legal"
    }


def determine_impact_level(title: str, summary: str, item_type: str, theme: str) -> str:
    blob = f"{title} {summary} {item_type} {theme}".lower()

    high_terms = [
        "final rules", "policy statement", "consultation", "mifidpru",
        "transaction reporting", "consumer duty", "operational resilience",
        "outsourcing", "sanctions", "market abuse", "capital", "liquidity"
    ]
    medium_terms = [
        "speech", "discussion paper", "dear ceo", "governance",
        "sustainability", "cyber", "data", "ai"
    ]

    if any(term in blob for term in high_terms):
        return "High"
    if any(term in blob for term in medium_terms):
        return "Medium"
    return "Low"


def determine_status(item_type: str, published_date: str = None, due_date: str = None) -> str:
    today = datetime.now(timezone.utc).date()

    if item_type == "Consultation":
        if due_date:
            due = pd.to_datetime(due_date, errors="coerce")
            if pd.notna(due):
                return "Open Consultation" if due.date() >= today else "Closed / Historic"
        return "Open Consultation"

    if item_type in {"Policy Statement", "Final Rules / Policy", "Supervisory Statement", "Handbook Notice"}:
        return "Final Rules / Policy"

    if item_type == "Dear CEO":
        return "Supervisory Attention"

    if published_date:
        published = pd.to_datetime(published_date, errors="coerce")
        if pd.notna(published):
            delta_days = (today - published.date()).days
            if delta_days <= 14:
                return "New"
            if delta_days <= 90:
                return "Upcoming Implementation"

    return "Closed / Historic"


def score_relevance(title: str, summary: str, authority: str, item_type: str, theme: str, sub_theme: str, source: str) -> int:
    blob = f"{title} {summary} {theme} {sub_theme}".lower()
    score = 20

    if authority in {"FCA", "PRA", "Bank of England"}:
        score += 20

    if source == "Regulatory Initiatives Grid":
        score += 10

    if item_type in {"Consultation", "Policy Statement", "Final Rules / Policy", "Supervisory Statement", "Dear CEO"}:
        score += 15

    for keyword in FMRUK_KEYWORDS:
        if keyword in blob:
            score += 4

    if sub_theme in {
        "Transaction Reporting",
        "Best Execution",
        "Market Abuse / Surveillance",
        "Operational Resilience",
        "Outsourcing / Third Party Risk",
        "Capital / Liquidity",
        "Consumer Duty",
        "AML / KYC / Sanctions"
    }:
        score += 12

    return min(score, 100)


def is_fmruk_relevant(title: str, summary: str, authority: str, theme: str, sub_theme: str) -> bool:
    if authority not in {"FCA", "PRA", "Bank of England"}:
        return False

    blob = f"{title} {summary} {theme} {sub_theme}".lower()
    return any(keyword in blob for keyword in FMRUK_KEYWORDS) or theme != "General UK Regulatory Change"


def build_rationale(authority: str, theme: str, sub_theme: str, impact: str) -> str:
    return (
        f"{authority} item classified under {theme} / {sub_theme}. "
        f"This is likely relevant where it could affect UK entity policy, controls, governance, oversight, reporting or implementation planning. "
        f"Potential operational implication: {impact}"
    )


def build_suggested_action(item_type: str, impact_level: str, primary_owner: str, secondary_owner: str) -> str:
    if impact_level == "High":
        return f"Immediate triage by {primary_owner}; engage {secondary_owner} and assess whether policy, procedure, control or governance updates are required."
    if item_type == "Consultation":
        return f"Triage by {primary_owner}; consider whether {secondary_owner} should support an applicability assessment or consultation response."
    return f"Review by {primary_owner}; determine whether monitoring, interpretation or follow-up with {secondary_owner} is needed."


def parse_due_date_from_summary(summary: str):
    patterns = [
        r"responses?\s+by\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        r"deadline\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        r"closes?\s+on\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})"
    ]

    for pattern in patterns:
        match = re.search(pattern, summary, flags=re.IGNORECASE)
        if match:
            parsed = parse_date(match.group(1))
            if parsed:
                return parsed
    return None


def get_feed_items(feed_url: str, authority_label: str, source_label: str):
    feed = feedparser.parse(feed_url)
    items = []

    for entry in feed.entries[:100]:
        title = clean_html(entry.get("title", ""))
        summary = clean_html(entry.get("summary", "") or entry.get("description", ""))
        url = entry.get("link", "")
        published_date = parse_date(entry.get("published", "") or entry.get("updated", ""))
        due_date = parse_due_date_from_summary(summary)
        item_type = infer_type(title, summary)
        classification = classify_item(title, summary)
        impact_level = determine_impact_level(title, summary, item_type, classification["theme"])
        status = determine_status(item_type, published_date, due_date)
        relevance = score_relevance(
            title=title,
            summary=summary,
            authority=authority_label,
            item_type=item_type,
            theme=classification["theme"],
            sub_theme=classification["sub_theme"],
            source=source_label
        )

        impact = classification["potential_business_impact"]

        record = {
            "id": stable_id(source_label, title, url),
            "source": source_label,
            "authority": authority_label,
            "type": item_type,
            "title": title,
            "summary": summary[:700],
            "url": url,
            "published_date": published_date,
            "due_date": due_date,
            "status": status,
            "theme": classification["theme"],
            "sub_theme": classification["sub_theme"],
            "impact_level": impact_level,
            "potential_business_impact": impact,
            "primary_owner": classification["primary_owner"],
            "secondary_owner": classification["secondary_owner"],
            "relevance_score": relevance,
            "is_fmruk_relevant": is_fmruk_relevant(
                title, summary, authority_label, classification["theme"], classification["sub_theme"]
            ),
            "rationale": build_rationale(authority_label, classification["theme"], classification["sub_theme"], impact),
            "suggested_action": build_suggested_action(
                item_type, impact_level, classification["primary_owner"], classification["secondary_owner"]
            )
        }

        items.append(record)

    return items


def normalise_col(value) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def pick_col(columns, options):
    for option in options:
        for col in columns:
            if option in col:
                return col
    return None


def normalise_authority(value: str) -> str:
    v = str(value).lower()
    if "fca" in v:
        return "FCA"
    if "pra" in v:
        return "PRA"
    if "bank" in v:
        return "Bank of England"
    return str(value).strip() or "Regulator"


def fetch_grid_items():
    items = []

    try:
        excel = pd.ExcelFile(GRID_XLSX)
        frames = []

        for sheet in excel.sheet_names:
            try:
                df = pd.read_excel(excel, sheet_name=sheet)
                if len(df.columns) >= 3 and len(df) > 0:
                    frames.append(df)
            except Exception:
                continue

        if not frames:
            return items

        grid_df = pd.concat(frames, ignore_index=True)
        grid_df.columns = [normalise_col(c) for c in grid_df.columns]

        title_col = pick_col(grid_df.columns, ["initiative", "title", "name"])
        authority_col = pick_col(grid_df.columns, ["authority", "regulator"])
        summary_col = pick_col(grid_df.columns, ["description", "summary", "detail"])
        timing_col = pick_col(grid_df.columns, ["timing", "planned_timing", "milestone", "date"])
        sector_col = pick_col(grid_df.columns, ["sector"])
        impact_col = pick_col(grid_df.columns, ["impact"])

        if not title_col:
            return items

        grid_df = grid_df[grid_df[title_col].notna()].copy()

        for _, row in grid_df.iterrows():
            title = str(row.get(title_col, "")).strip()
            if not title or title.lower() == "nan":
                continue

            authority = normalise_authority(row.get(authority_col, "Regulatory Initiatives Forum"))

            summary_parts = []
            for col in [summary_col, timing_col, sector_col, impact_col]:
                if col and pd.notna(row.get(col)):
                    summary_parts.append(f"{col.replace('_', ' ').title()}: {str(row.get(col)).strip()}")

            summary = " | ".join(summary_parts)[:700]
            item_type = "Pipeline Initiative"
            classification = classify_item(title, summary)
            impact_level = "High" if classification["sub_theme"] in {
                "Transaction Reporting", "Operational Resilience", "Outsourcing / Third Party Risk",
                "Capital / Liquidity", "Consumer Duty"
            } else "Medium"

            relevance = score_relevance(
                title=title,
                summary=summary,
                authority=authority,
                item_type=item_type,
                theme=classification["theme"],
                sub_theme=classification["sub_theme"],
                source="Regulatory Initiatives Grid"
            )

            items.append({
                "id": stable_id("grid", title, authority, summary),
                "source": "Regulatory Initiatives Grid",
                "authority": authority,
                "type": item_type,
                "title": title,
                "summary": summary or "Forward-looking pipeline item from the Regulatory Initiatives Grid.",
                "url": "https://www.fca.org.uk/publications/corporate-documents/regulatory-initiatives-grid/dashboard",
                "published_date": "2025-12-11",
                "due_date": None,
                "status": "Upcoming Implementation",
                "theme": classification["theme"],
                "sub_theme": classification["sub_theme"],
                "impact_level": impact_level,
                "potential_business_impact": classification["potential_business_impact"],
                "primary_owner": classification["primary_owner"],
                "secondary_owner": classification["secondary_owner"],
                "relevance_score": min(relevance + 5, 100),
                "is_fmruk_relevant": is_fmruk_relevant(
                    title, summary, authority, classification["theme"], classification["sub_theme"]
                ),
                "rationale": (
                    f"Forward-looking pipeline item from the Regulatory Initiatives Grid under "
                    f"{classification['theme']} / {classification['sub_theme']}. "
                    f"Useful for implementation planning, sequencing and resource allocation for FMRUK."
                ),
                "suggested_action": build_suggested_action(
                    item_type, impact_level, classification["primary_owner"], classification["secondary_owner"]
                )
            })

    except Exception as exc:
        print(f"Grid fetch failed: {exc}")

    return items


def dedupe_items(items):
    seen = set()
    deduped = []

    for item in items:
        key = (item.get("title"), item.get("url"))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(item)

    return deduped


def sort_items(items):
    def sort_key(item):
        relevance = item.get("relevance_score", 0)
        date_value = item.get("due_date") or item.get("published_date") or "1900-01-01"
        return (-relevance, str(date_value))

    return sorted(items, key=sort_key)


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    fca_items = get_feed_items(FCA_RSS, "FCA", "FCA RSS")
    pra_items = get_feed_items(PRA_RSS, "PRA", "PRA Prudential Publications RSS")
    grid_items = fetch_grid_items()

    all_items = dedupe_items(fca_items + pra_items + grid_items)
    all_items = sort_items(all_items)

    payload = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "item_count": len(all_items),
        "items": all_items
    }

    OUTPUT_FILE.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {len(all_items)} items to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
