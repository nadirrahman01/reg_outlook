import json
import re
from datetime import datetime, timezone
from pathlib import Path
from hashlib import md5

import feedparser
import pandas as pd
import requests
from bs4 import BeautifulSoup


BASE_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = BASE_DIR / "data"
OUTPUT_FILE = DATA_DIR / "regulatory_updates.json"

FCA_RSS = "https://www.fca.org.uk/news/rss.xml"
PRA_RSS = "https://www.bankofengland.co.uk/rss/prudential-regulation-publications"
GRID_XLSX = "https://www.fca.org.uk/publication/corporate/regulatory-initiatives-grid-dec-2025.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; FMRUK-Regulatory-Tracker/1.0; +https://github.com/)"
}


THEME_RULES = {
    "Operational Resilience": [
        "operational resilience", "third party", "critical third party",
        "outsourcing", "resilience", "business continuity"
    ],
    "Consumer / Conduct": [
        "consumer duty", "consumer", "retail", "complaints", "redress",
        "fair value", "vulnerable customers"
    ],
    "Prudential / MIFIDPRU": [
        "mifidpru", "prudential", "capital", "liquidity", "icara",
        "remuneration", "concentration risk"
    ],
    "Market Conduct / Trading": [
        "market abuse", "best execution", "transaction reporting",
        "mifid", "short selling", "uk emir", "sftr", "securitisation"
    ],
    "AML / Sanctions / Financial Crime": [
        "financial crime", "aml", "anti-money laundering", "sanctions",
        "fraud", "crime", "kyc"
    ],
    "ESG / Sustainability": [
        "sustainability", "sdr", "esg", "climate", "transition plan",
        "tnfd", "greenwashing"
    ],
    "Data / AI / Technology": [
        "artificial intelligence", "ai", "machine learning", "cyber",
        "data", "model risk", "technology"
    ],
    "Governance / SMCR": [
        "smcr", "senior managers", "certification regime", "governance",
        "board", "dear ceo"
    ]
}


OWNER_RULES = {
    "Operational Resilience": "Operational Risk / UK Compliance / Technology Risk",
    "Consumer / Conduct": "UK Compliance / Legal / Product Governance",
    "Prudential / MIFIDPRU": "Finance / Prudential Compliance / Risk",
    "Market Conduct / Trading": "Market Compliance / Trading Oversight",
    "AML / Sanctions / Financial Crime": "Financial Crime Compliance",
    "ESG / Sustainability": "Sustainability / Product / Compliance",
    "Data / AI / Technology": "Technology Risk / Data Governance / Compliance",
    "Governance / SMCR": "Company Secretariat / Compliance / HR"
}


def clean_html(raw: str) -> str:
    if not raw:
        return ""
    soup = BeautifulSoup(raw, "html.parser")
    text = soup.get_text(" ", strip=True)
    return re.sub(r"\s+", " ", text).strip()


def parse_date(date_str: str):
    if not date_str:
        return None
    for fmt in [
        "%a, %d %b %Y %H:%M:%S %Z",
        "%Y-%m-%d",
        "%d %B %Y",
        "%d %b %Y",
        "%Y/%m/%d"
    ]:
        try:
            return datetime.strptime(date_str, fmt).date().isoformat()
        except Exception:
            continue
    try:
        return pd.to_datetime(date_str, errors="coerce").date().isoformat()
    except Exception:
        return None


def score_theme(text: str) -> str:
    blob = text.lower()
    for theme, keywords in THEME_RULES.items():
        for keyword in keywords:
            if keyword in blob:
                return theme
    return "General UK Regulatory Change"


def relevance_score(text: str, authority: str, source: str) -> int:
    blob = text.lower()
    score = 30

    if authority in {"FCA", "PRA", "Bank of England"}:
        score += 20

    high_value_terms = [
        "consultation", "policy statement", "final rules", "supervisory statement",
        "dear ceo", "handbook", "mifidpru", "consumer duty", "operational resilience",
        "financial crime", "sanctions", "market abuse", "outsourcing", "capital",
        "liquidity", "remuneration", "transaction reporting", "best execution"
    ]
    for term in high_value_terms:
        if term in blob:
            score += 6

    if "grid" in source.lower():
        score += 10

    if "asset management" in blob or "investment firm" in blob or "firms" in blob:
        score += 8

    return min(score, 100)


def build_rationale(title: str, summary: str, theme: str, authority: str) -> str:
    base = f"{authority} item tagged to {theme}. "
    if theme in OWNER_RULES:
        return (
            base
            + "This is likely relevant where it could change FMRUK policy, controls, disclosures, monitoring, governance or implementation planning."
        )
    return (
        base
        + "This may still matter if it changes UK regulatory expectations, timing, operational workload or assurance requirements."
    )


def is_fmruk_relevant(title: str, summary: str, authority: str, theme: str) -> bool:
    blob = f"{title} {summary}".lower()

    if authority not in {"FCA", "PRA", "Bank of England"}:
        return False

    likely_relevant = [
        "firm", "investment", "asset", "consumer", "conduct", "operational resilience",
        "outsourcing", "financial crime", "sanctions", "market abuse", "transaction reporting",
        "mifid", "capital", "liquidity", "governance", "remuneration", "dear ceo",
        "sustainability", "ai", "data"
    ]
    if any(term in blob for term in likely_relevant):
        return True

    return theme != "General UK Regulatory Change"


def status_from_dates(published_date: str, due_date: str = None) -> str:
    today = datetime.now(timezone.utc).date()

    if due_date:
        try:
            due = pd.to_datetime(due_date).date()
            return "Open" if due >= today else "Closed"
        except Exception:
            pass

    if published_date:
        try:
            published = pd.to_datetime(published_date).date()
            delta_days = (today - published).days
            return "Upcoming" if delta_days <= 30 else "Pipeline"
        except Exception:
            pass

    return "Pipeline"


def stable_id(*parts: str) -> str:
    joined = "||".join([p or "" for p in parts])
    return md5(joined.encode("utf-8")).hexdigest()[:12]


def get_feed_items(feed_url: str, authority_label: str, source_label: str):
    feed = feedparser.parse(feed_url)
    items = []

    for entry in feed.entries[:80]:
        title = clean_html(entry.get("title", ""))
        summary = clean_html(entry.get("summary", "") or entry.get("description", ""))
        url = entry.get("link", "")
        published = parse_date(entry.get("published", "") or entry.get("updated", ""))
        text_blob = f"{title} {summary}"
        theme = score_theme(text_blob)
        score = relevance_score(text_blob, authority_label, source_label)

        record = {
            "id": stable_id(source_label, title, url),
            "source": source_label,
            "authority": authority_label,
            "type": infer_type(title, summary),
            "title": title,
            "summary": summary[:600],
            "url": url,
            "published_date": published,
            "due_date": None,
            "status": status_from_dates(published),
            "theme": theme,
            "relevance_score": score,
            "is_fmruk_relevant": is_fmruk_relevant(title, summary, authority_label, theme),
            "rationale": build_rationale(title, summary, theme, authority_label),
            "suggested_owner": OWNER_RULES.get(theme, "UK Compliance / Legal")
        }
        items.append(record)

    return items


def infer_type(title: str, summary: str) -> str:
    blob = f"{title} {summary}".lower()

    mapping = [
        ("Policy Statement", "policy statement"),
        ("Consultation", "consultation"),
        ("Supervisory Statement", "supervisory statement"),
        ("Dear CEO", "dear ceo"),
        ("Handbook Notice", "handbook notice"),
        ("Statement", " statement"),
        ("Speech", "speech"),
        ("News", "news")
    ]
    for label, term in mapping:
        if term in blob:
            return label
    return "Publication"


def fetch_grid_items():
    items = []

    try:
        # Read all sheets and merge what we can.
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

        title_col = pick_col(grid_df.columns, ["initiative", "name", "title"])
        authority_col = pick_col(grid_df.columns, ["authority", "regulator"])
        summary_col = pick_col(grid_df.columns, ["description", "detail", "summary"])
        sector_col = pick_col(grid_df.columns, ["sector"])
        impact_col = pick_col(grid_df.columns, ["impact", "expected_impact"])
        timing_col = pick_col(grid_df.columns, ["timing", "planned_timing", "date", "milestone"])
        consumer_col = pick_col(grid_df.columns, ["consumer"])

        if not title_col:
            return items

        grid_df = grid_df[grid_df[title_col].notna()].copy()

        for _, row in grid_df.iterrows():
            title = str(row.get(title_col, "")).strip()
            if not title or title.lower() == "nan":
                continue

            authority = str(row.get(authority_col, "Regulatory Initiatives Forum")).strip()
            summary_parts = []

            for col in [summary_col, sector_col, impact_col, timing_col, consumer_col]:
                if col and pd.notna(row.get(col)):
                    summary_parts.append(f"{col.replace('_', ' ').title()}: {str(row.get(col)).strip()}")

            summary = " | ".join(summary_parts)[:700]
            theme = score_theme(f"{title} {summary}")
            score = relevance_score(f"{title} {summary}", authority, "Regulatory Initiatives Grid")

            items.append({
                "id": stable_id("grid", title, authority, summary),
                "source": "Regulatory Initiatives Grid",
                "authority": normalise_authority(authority),
                "type": "Pipeline Initiative",
                "title": title,
                "summary": summary or "Pipeline item from the Regulatory Initiatives Grid.",
                "url": "https://www.fca.org.uk/publications/corporate-documents/regulatory-initiatives-grid/dashboard",
                "published_date": "2025-12-11",
                "due_date": None,
                "status": "Pipeline",
                "theme": theme,
                "relevance_score": min(score + 5, 100),
                "is_fmruk_relevant": is_fmruk_relevant(title, summary, normalise_authority(authority), theme),
                "rationale": (
                    "Forward-looking pipeline item from the Regulatory Initiatives Grid. "
                    "Useful for implementation planning, sequencing and assessing future compliance demand on FMRUK."
                ),
                "suggested_owner": OWNER_RULES.get(theme, "UK Compliance / Legal")
            })

    except Exception as exc:
        print(f"Grid fetch failed: {exc}")

    return items


def normalise_col(value) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def pick_col(columns, keyword_options):
    for option in keyword_options:
        for col in columns:
            if option in col:
                return col
    return None


def normalise_authority(value: str) -> str:
    v = value.lower()
    if "fca" in v:
        return "FCA"
    if "pra" in v:
        return "PRA"
    if "bank" in v:
        return "Bank of England"
    return value or "Regulator"


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
        due = item.get("due_date") or item.get("published_date") or ""
        return (-relevance, due)

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
