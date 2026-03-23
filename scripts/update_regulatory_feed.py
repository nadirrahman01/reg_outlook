import os
import re
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Any, Optional

import pandas as pd
import pdfplumber
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
DB_PATH = BASE_DIR / "regulatory_dashboard.db"

UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
app.config["UPLOAD_FOLDER"] = str(UPLOAD_DIR)
ALLOWED_EXTENSIONS = {"pdf", "xlsx", "xls"}


KNOWN_SECTIONS = [
    "Multi-sector",
    "Banking, credit and lending",
    "Payments and cryptoassets",
    "Insurance and reinsurance",
    "Investment management",
    "Pensions and retirement income",
    "Retail investments",
    "Wholesale financial markets",
    "Annex: initiatives completed/stopped",
]

KNOWN_SUBCATEGORIES = [
    "Competition, innovation and other",
    "Conduct",
    "Cross-cutting/omnibus",
    "Sustainable finance",
    "Financial resilience",
    "Operational resilience",
    "Other single-sector initiatives",
    "Repeal and replacement of assimilated law under FSMA 2023",
]

LEAD_PATTERN = re.compile(
    r"^(FCA|PRA|BoE|HMT|ICO|TPR|FRC|PSR|CMA|FOS|FSCS)(/[A-Z][A-Za-z]+|/[A-Z]{2,5}|/[A-Za-z]{2,10})*$"
)

THEME_RULES = [
    {
        "theme": "Market Conduct / Trading",
        "sub_theme": "Transaction Reporting",
        "keywords": [
            "transaction reporting", "mifir", "regulatory reporting",
            "reporting fields", "approved reporting mechanism", "arm"
        ],
        "impact": "May require review of transaction reporting controls, reconciliations, exception monitoring and governance.",
        "primary_owner": "Compliance",
        "secondary_owner": "Operations",
    },
    {
        "theme": "Market Conduct / Trading",
        "sub_theme": "Best Execution",
        "keywords": ["best execution", "rts 28", "execution quality", "venue analysis", "execution policy"],
        "impact": "Could affect best execution governance, monitoring, policy wording and oversight of execution arrangements.",
        "primary_owner": "Compliance",
        "secondary_owner": "Legal",
    },
    {
        "theme": "Market Conduct / Trading",
        "sub_theme": "Market Abuse / Surveillance",
        "keywords": ["market abuse", "mar", "surveillance", "inside information", "personal account dealing"],
        "impact": "May require updates to surveillance controls, market abuse risk assessments, training and escalation processes.",
        "primary_owner": "Compliance",
        "secondary_owner": "Legal",
    },
    {
        "theme": "Operational Resilience",
        "sub_theme": "Operational Resilience",
        "keywords": ["operational resilience", "important business service", "impact tolerance", "resilience"],
        "impact": "May require review of resilience mapping, testing, scenario analysis and governance over important business services.",
        "primary_owner": "Risk",
        "secondary_owner": "Compliance",
    },
    {
        "theme": "Operational Resilience",
        "sub_theme": "Outsourcing / Third Party Risk",
        "keywords": ["outsourcing", "third party", "critical third party", "service provider", "vendor risk"],
        "impact": "Could require updates to outsourcing registers, due diligence, contractual controls and oversight of service providers.",
        "primary_owner": "Risk",
        "secondary_owner": "Technology",
    },
    {
        "theme": "Prudential / MIFIDPRU",
        "sub_theme": "Capital / Liquidity",
        "keywords": ["mifidpru", "capital", "liquidity", "own funds", "icara", "prudential", "concentration risk"],
        "impact": "May affect prudential assessments, capital planning, ICARA assumptions, monitoring and governance.",
        "primary_owner": "Finance",
        "secondary_owner": "Risk",
    },
    {
        "theme": "Prudential / MIFIDPRU",
        "sub_theme": "Remuneration",
        "keywords": ["remuneration", "bonus", "pay", "compensation", "malus", "clawback", "incentive"],
        "impact": "Could require updates to remuneration frameworks, governance, documentation and control testing.",
        "primary_owner": "HR",
        "secondary_owner": "Compliance",
    },
    {
        "theme": "Consumer / Conduct",
        "sub_theme": "Consumer Duty",
        "keywords": ["consumer duty", "fair value", "good outcomes", "vulnerable customers", "consumer"],
        "impact": "May require review of product governance, distribution oversight, client communications and conduct monitoring.",
        "primary_owner": "Compliance",
        "secondary_owner": "Product",
    },
    {
        "theme": "AML / Sanctions / Financial Crime",
        "sub_theme": "AML / KYC / Sanctions",
        "keywords": ["aml", "anti-money laundering", "sanctions", "kyc", "financial crime", "fraud", "bribery"],
        "impact": "May require review of AML controls, KYC processes, sanctions screening, escalation and training.",
        "primary_owner": "Compliance",
        "secondary_owner": "Legal",
    },
    {
        "theme": "Governance / SMCR",
        "sub_theme": "SMCR / Governance",
        "keywords": ["smcr", "senior managers", "certification regime", "governance", "board", "committee", "dear ceo"],
        "impact": "Could affect governance frameworks, accountabilities, committee reporting and senior management oversight.",
        "primary_owner": "Compliance",
        "secondary_owner": "HR",
    },
    {
        "theme": "ESG / Sustainability",
        "sub_theme": "Sustainability / SDR",
        "keywords": ["sustainability", "sdr", "greenwashing", "climate", "esg", "transition plan", "tnfd", "tcfd", "srs"],
        "impact": "May require updates to sustainability disclosures, product communications, governance and control evidence.",
        "primary_owner": "Legal",
        "secondary_owner": "Compliance",
    },
    {
        "theme": "Data / AI / Technology",
        "sub_theme": "AI / Data / Cyber",
        "keywords": ["artificial intelligence", "ai", "machine learning", "cyber", "data", "ict", "technology"],
        "impact": "Could affect governance over AI use cases, data controls, cyber oversight and internal policy requirements.",
        "primary_owner": "Technology",
        "secondary_owner": "Compliance",
    },
]

FMRUK_KEYWORDS = [
    "uk", "fca", "pra", "investment firm", "asset management", "investment management",
    "mifid", "mifidpru", "consumer duty", "transaction reporting", "market abuse",
    "best execution", "outsourcing", "operational resilience", "financial crime",
    "sanctions", "governance", "remuneration", "sustainability", "disclosure",
    "regulatory reporting", "firm", "firms", "data collections"
]


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            saved_path TEXT NOT NULL,
            file_type TEXT NOT NULL,
            uploaded_at TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS initiatives (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_id INTEGER NOT NULL,
            section_name TEXT,
            subcategory TEXT,
            lead_regulator TEXT,
            initiative_title TEXT,
            initiative_description TEXT,
            expected_key_milestones TEXT,
            indicative_impact_on_firms TEXT,
            consumer_interest TEXT,
            timing_updated TEXT,
            is_new TEXT,
            timing_bucket TEXT,
            theme TEXT,
            internal_sub_theme TEXT,
            impact_level TEXT,
            potential_business_impact TEXT,
            primary_owner TEXT,
            secondary_owner TEXT,
            relevance_score INTEGER,
            is_fmruk_relevant INTEGER,
            rationale TEXT,
            suggested_action TEXT,
            raw_text TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY (upload_id) REFERENCES uploads (id)
        )
    """)

    conn.commit()
    conn.close()


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def normalise_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def title_caseish(text: str) -> str:
    return normalise_ws(text)


def extract_pdf_text(path: Path) -> List[Dict[str, Any]]:
    pages = []
    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            pages.append({"page": i, "text": text})
    return pages


def parse_section_and_subcategory(line: str, current_section: Optional[str], current_subcategory: Optional[str]):
    clean = normalise_ws(line)

    if clean in KNOWN_SECTIONS:
        return clean, None

    if clean in KNOWN_SUBCATEGORIES:
        return current_section, clean

    return current_section, current_subcategory


def looks_like_lead_line(line: str) -> bool:
    cleaned = normalise_ws(line)
    if not cleaned:
        return False
    if len(cleaned) > 40:
        return False
    return bool(LEAD_PATTERN.match(cleaned))


def parse_pdf_initiatives(path: Path) -> List[Dict[str, Any]]:
    pages = extract_pdf_text(path)
    initiatives: List[Dict[str, Any]] = []

    current_section = None
    current_subcategory = None
    current_item = None

    skip_lines = {
        "Lead Initiative Expected key milestones",
        "Indicative impact on firms",
        "Consumer interest",
        "Timing updated",
        "New",
        "Key",
    }

    for page in pages:
        raw_lines = [normalise_ws(x) for x in page["text"].splitlines()]
        lines = [x for x in raw_lines if x]

        for idx, line in enumerate(lines):
            if line in skip_lines:
                continue
            if line.startswith("Regulatory Initiatives Grid |"):
                continue

            current_section, current_subcategory = parse_section_and_subcategory(
                line, current_section, current_subcategory
            )

            if looks_like_lead_line(line):
                next_line = lines[idx + 1] if idx + 1 < len(lines) else ""
                if next_line and len(next_line) > 3 and next_line not in KNOWN_SECTIONS and next_line not in KNOWN_SUBCATEGORIES:
                    if current_item:
                        initiatives.append(current_item)

                    current_item = {
                        "section_name": current_section,
                        "subcategory": current_subcategory,
                        "lead_regulator": line,
                        "initiative_title": title_caseish(next_line),
                        "initiative_description": "",
                        "expected_key_milestones": "",
                        "indicative_impact_on_firms": "",
                        "consumer_interest": "",
                        "timing_updated": "",
                        "is_new": "",
                        "timing_bucket": "",
                        "raw_text": f"{line}\n{next_line}\n",
                    }
                    continue

            if current_item:
                if line == current_item["initiative_title"] or line == current_item["lead_regulator"]:
                    continue

                current_item["raw_text"] += line + "\n"

    if current_item:
        initiatives.append(current_item)

    cleaned = []
    for item in initiatives:
        populated = enrich_from_raw_text(item)
        if populated["initiative_title"] and len(populated["initiative_title"]) > 3:
            cleaned.append(populated)

    return dedupe_initiatives(cleaned)


def enrich_from_raw_text(item: Dict[str, Any]) -> Dict[str, Any]:
    raw = normalise_ws(item.get("raw_text", ""))
    text = raw

    impact_match = re.search(r"\b(H|L|U)\b", text)
    if impact_match and not item.get("indicative_impact_on_firms"):
        item["indicative_impact_on_firms"] = impact_match.group(1)

    new_match = re.search(r"\bnew\b", text, flags=re.IGNORECASE)
    item["is_new"] = "Yes" if new_match else item.get("is_new") or "No"

    timing_update_match = re.search(r"timing", text, flags=re.IGNORECASE)
    item["timing_updated"] = "Possible" if timing_update_match else item.get("timing_updated") or "No"

    milestones = []
    milestone_patterns = [
        r"(Q[1-4]\s+\d{4}:[^.]+(?:\.)?)",
        r"(\b(?:January|February|March|April|May|June|July|August|September|October|November|December)[^.]+(?:\.)?)",
        r"(\b\d{1,2}\s+[A-Z][a-z]+\s+\d{4}[^.]*\.)",
    ]
    for pattern in milestone_patterns:
        milestones.extend(re.findall(pattern, raw))

    if milestones:
        deduped_m = []
        seen = set()
        for m in milestones:
            m2 = normalise_ws(m)
            if m2 not in seen and len(m2) > 8:
                seen.add(m2)
                deduped_m.append(m2)
        item["expected_key_milestones"] = " | ".join(deduped_m[:6])

    desc = raw
    desc = desc.replace(item.get("lead_regulator", ""), "")
    desc = desc.replace(item.get("initiative_title", ""), "")
    desc = normalise_ws(desc)
    item["initiative_description"] = desc[:1200]

    bucket = infer_timing_bucket(raw)
    item["timing_bucket"] = bucket

    return item


def infer_timing_bucket(text: str) -> str:
    blob = text.lower()
    if "q1 2026" in blob or "q2 2026" in blob or "january" in blob or "april" in blob:
        return "Near Term"
    if "q3 2026" in blob or "q4 2026" in blob or "2026" in blob:
        return "Medium Term"
    if "2027" in blob or "post july 2027" in blob:
        return "Longer Term"
    return "To Be Confirmed"


def parse_excel_initiatives(path: Path) -> List[Dict[str, Any]]:
    initiatives = []
    xls = pd.ExcelFile(path)

    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(path, sheet_name=sheet_name)
        except Exception:
            continue

        if df.empty:
            continue

        cols = [str(c).strip().lower() for c in df.columns]
        df.columns = cols

        def col_like(options):
            for c in cols:
                for o in options:
                    if o in c:
                        return c
            return None

        lead_col = col_like(["lead"])
        init_col = col_like(["initiative", "title", "name"])
        desc_col = col_like(["description"])
        milestone_col = col_like(["expected key milestones", "milestone"])
        impact_col = col_like(["impact on firms", "indicative impact"])
        consumer_col = col_like(["consumer interest"])
        timing_updated_col = col_like(["timing updated", "change in timing"])
        new_col = col_like(["new"])
        sector_col = col_like(["sector", "section"])
        subcat_col = col_like(["subcategory", "sub-category", "category"])
        timing_bucket_col = col_like(["post july", "q1", "q2", "q3", "q4"])

        if not init_col:
            continue

        for _, row in df.iterrows():
            title = normalise_ws(str(row.get(init_col, "")))
            if not title or title.lower() == "nan":
                continue

            initiatives.append({
                "section_name": normalise_ws(str(row.get(sector_col, sheet_name))) if sector_col else sheet_name,
                "subcategory": normalise_ws(str(row.get(subcat_col, ""))) if subcat_col else "",
                "lead_regulator": normalise_ws(str(row.get(lead_col, ""))) if lead_col else "",
                "initiative_title": title,
                "initiative_description": normalise_ws(str(row.get(desc_col, ""))) if desc_col else "",
                "expected_key_milestones": normalise_ws(str(row.get(milestone_col, ""))) if milestone_col else "",
                "indicative_impact_on_firms": normalise_ws(str(row.get(impact_col, ""))) if impact_col else "",
                "consumer_interest": normalise_ws(str(row.get(consumer_col, ""))) if consumer_col else "",
                "timing_updated": normalise_ws(str(row.get(timing_updated_col, ""))) if timing_updated_col else "",
                "is_new": normalise_ws(str(row.get(new_col, ""))) if new_col else "",
                "timing_bucket": normalise_ws(str(row.get(timing_bucket_col, ""))) if timing_bucket_col else "",
                "raw_text": json.dumps({k: str(v) for k, v in row.to_dict().items()}, ensure_ascii=False),
            })

    return dedupe_initiatives(initiatives)


def dedupe_initiatives(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen = set()
    out = []
    for item in items:
        key = (
            normalise_ws(item.get("section_name", "")).lower(),
            normalise_ws(item.get("lead_regulator", "")).lower(),
            normalise_ws(item.get("initiative_title", "")).lower(),
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out


def classify_item(title: str, description: str, section_name: str, subcategory: str) -> Dict[str, str]:
    blob = f"{title} {description} {section_name} {subcategory}".lower()

    for rule in THEME_RULES:
        if any(keyword in blob for keyword in rule["keywords"]):
            return {
                "theme": rule["theme"],
                "internal_sub_theme": rule["sub_theme"],
                "potential_business_impact": rule["impact"],
                "primary_owner": rule["primary_owner"],
                "secondary_owner": rule["secondary_owner"],
            }

    section_blob = f"{section_name} {subcategory}".lower()

    if "investment management" in section_blob:
        return {
            "theme": "Investment Management / Product",
            "internal_sub_theme": "Investment Management",
            "potential_business_impact": "May require review of product governance, disclosures, oversight or firm implementation planning.",
            "primary_owner": "Compliance",
            "secondary_owner": "Legal",
        }

    if "wholesale financial markets" in section_blob:
        return {
            "theme": "Market Structure / Wholesale",
            "internal_sub_theme": "Wholesale Markets",
            "potential_business_impact": "Could affect market-facing controls, policy interpretation and implementation planning for wholesale business lines.",
            "primary_owner": "Compliance",
            "secondary_owner": "Legal",
        }

    return {
        "theme": "General UK Regulatory Change",
        "internal_sub_theme": "General",
        "potential_business_impact": "May require initial triage to determine whether any policy, control, governance or implementation response is needed.",
        "primary_owner": "Compliance",
        "secondary_owner": "Legal",
    }


def determine_impact_level(indicative_impact: str, title: str, description: str) -> str:
    impact_flag = (indicative_impact or "").strip().upper()
    blob = f"{title} {description}".lower()

    if impact_flag == "H":
        return "High"
    if impact_flag == "L":
        return "Low"

    high_terms = [
        "consumer duty", "transaction reporting", "operational resilience", "outsourcing",
        "mifidpru", "capital", "liquidity", "market abuse", "sanctions", "data collections"
    ]
    medium_terms = ["consultation", "disclosure", "governance", "cyber", "ict", "sustainability"]

    if any(t in blob for t in high_terms):
        return "High"
    if any(t in blob for t in medium_terms):
        return "Medium"
    return "Medium" if impact_flag == "U" else "Low"


def score_relevance(title: str, description: str, section_name: str, theme: str, sub_theme: str) -> int:
    blob = f"{title} {description} {section_name} {theme} {sub_theme}".lower()
    score = 25

    for kw in FMRUK_KEYWORDS:
        if kw in blob:
            score += 4

    if section_name in {"Investment management", "Multi-sector", "Wholesale financial markets"}:
        score += 10

    if sub_theme in {
        "Transaction Reporting",
        "Operational Resilience",
        "Outsourcing / Third Party Risk",
        "Capital / Liquidity",
        "Consumer Duty",
        "AML / KYC / Sanctions",
        "AI / Data / Cyber"
    }:
        score += 12

    return min(score, 100)


def is_fmruk_relevant(title: str, description: str, section_name: str, theme: str) -> bool:
    blob = f"{title} {description} {section_name} {theme}".lower()

    if section_name in {"Investment management", "Multi-sector", "Wholesale financial markets"}:
        return True

    return any(kw in blob for kw in FMRUK_KEYWORDS)


def build_rationale(section_name: str, theme: str, sub_theme: str, impact: str) -> str:
    return (
        f"Classified from the uploaded Grid under {section_name}. "
        f"Mapped to {theme} / {sub_theme}. "
        f"This is likely relevant to FMRUK where it could affect UK entity policy, controls, governance, reporting or implementation planning. "
        f"Potential operational implication: {impact}"
    )


def build_suggested_action(impact_level: str, primary_owner: str, secondary_owner: str, timing_bucket: str) -> str:
    if impact_level == "High":
        return f"Immediate triage by {primary_owner}; engage {secondary_owner}; assess whether policy, procedure, control or governance updates are required. Horizon: {timing_bucket}."
    if impact_level == "Medium":
        return f"Review by {primary_owner}; confirm applicability with {secondary_owner}; track milestones and assign implementation owner if needed. Horizon: {timing_bucket}."
    return f"Monitor through {primary_owner}; retain on watchlist and reassess if timing or scope changes. Horizon: {timing_bucket}."


def analyse_initiatives(items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    analysed = []
    for item in items:
        classification = classify_item(
            item.get("initiative_title", ""),
            item.get("initiative_description", ""),
            item.get("section_name", ""),
            item.get("subcategory", ""),
        )
        impact_level = determine_impact_level(
            item.get("indicative_impact_on_firms", ""),
            item.get("initiative_title", ""),
            item.get("initiative_description", "")
        )
        relevance = score_relevance(
            item.get("initiative_title", ""),
            item.get("initiative_description", ""),
            item.get("section_name", ""),
            classification["theme"],
            classification["internal_sub_theme"],
        )
        relevant = is_fmruk_relevant(
            item.get("initiative_title", ""),
            item.get("initiative_description", ""),
            item.get("section_name", ""),
            classification["theme"],
        )
        rationale = build_rationale(
            item.get("section_name", ""),
            classification["theme"],
            classification["internal_sub_theme"],
            classification["potential_business_impact"],
        )
        suggested_action = build_suggested_action(
            impact_level,
            classification["primary_owner"],
            classification["secondary_owner"],
            item.get("timing_bucket", "To Be Confirmed"),
        )

        analysed.append({
            **item,
            "theme": classification["theme"],
            "internal_sub_theme": classification["internal_sub_theme"],
            "impact_level": impact_level,
            "potential_business_impact": classification["potential_business_impact"],
            "primary_owner": classification["primary_owner"],
            "secondary_owner": classification["secondary_owner"],
            "relevance_score": relevance,
            "is_fmruk_relevant": 1 if relevant else 0,
            "rationale": rationale,
            "suggested_action": suggested_action,
        })

    analysed.sort(key=lambda x: (-x["relevance_score"], x.get("initiative_title", "")))
    return analysed


def clear_existing_upload_data(conn, upload_id: int):
    conn.execute("DELETE FROM initiatives WHERE upload_id = ?", (upload_id,))
    conn.commit()


def save_upload(filename: str, saved_path: str, file_type: str) -> int:
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO uploads (filename, saved_path, file_type, uploaded_at) VALUES (?, ?, ?, ?)",
        (filename, saved_path, file_type, datetime.now(timezone.utc).isoformat())
    )
    upload_id = cur.lastrowid
    conn.commit()
    conn.close()
    return upload_id


def save_initiatives(upload_id: int, items: List[Dict[str, Any]]):
    conn = get_db()
    cur = conn.cursor()

    for item in items:
        cur.execute("""
            INSERT INTO initiatives (
                upload_id, section_name, subcategory, lead_regulator, initiative_title,
                initiative_description, expected_key_milestones, indicative_impact_on_firms,
                consumer_interest, timing_updated, is_new, timing_bucket, theme,
                internal_sub_theme, impact_level, potential_business_impact, primary_owner,
                secondary_owner, relevance_score, is_fmruk_relevant, rationale,
                suggested_action, raw_text, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            upload_id,
            item.get("section_name"),
            item.get("subcategory"),
            item.get("lead_regulator"),
            item.get("initiative_title"),
            item.get("initiative_description"),
            item.get("expected_key_milestones"),
            item.get("indicative_impact_on_firms"),
            item.get("consumer_interest"),
            item.get("timing_updated"),
            item.get("is_new"),
            item.get("timing_bucket"),
            item.get("theme"),
            item.get("internal_sub_theme"),
            item.get("impact_level"),
            item.get("potential_business_impact"),
            item.get("primary_owner"),
            item.get("secondary_owner"),
            item.get("relevance_score"),
            item.get("is_fmruk_relevant"),
            item.get("rationale"),
            item.get("suggested_action"),
            item.get("raw_text"),
            datetime.now(timezone.utc).isoformat(),
        ))

    conn.commit()
    conn.close()


def latest_upload_id() -> Optional[int]:
    conn = get_db()
    row = conn.execute("SELECT id FROM uploads ORDER BY id DESC LIMIT 1").fetchone()
    conn.close()
    return row["id"] if row else None


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file part found."}), 400

    f = request.files["file"]
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "No file selected."}), 400

    if not allowed_file(f.filename):
        return jsonify({"ok": False, "error": "Unsupported file type. Upload PDF or XLSX."}), 400

    filename = secure_filename(f.filename)
    ext = filename.rsplit(".", 1)[1].lower()
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
    saved_name = f"{timestamp}_{filename}"
    saved_path = UPLOAD_DIR / saved_name
    f.save(saved_path)

    upload_id = save_upload(filename=filename, saved_path=str(saved_path), file_type=ext)

    try:
        if ext in {"xlsx", "xls"}:
            parsed = parse_excel_initiatives(saved_path)
        else:
            parsed = parse_pdf_initiatives(saved_path)

        analysed = analyse_initiatives(parsed)
        save_initiatives(upload_id, analysed)

        return jsonify({
            "ok": True,
            "upload_id": upload_id,
            "filename": filename,
            "file_type": ext,
            "initiative_count": len(analysed)
        })
    except Exception as e:
        return jsonify({"ok": False, "error": f"Parsing failed: {str(e)}"}), 500


@app.route("/api/uploads")
def list_uploads():
    conn = get_db()
    rows = conn.execute("""
        SELECT u.id, u.filename, u.file_type, u.uploaded_at,
               COUNT(i.id) AS initiative_count
        FROM uploads u
        LEFT JOIN initiatives i ON i.upload_id = u.id
        GROUP BY u.id, u.filename, u.file_type, u.uploaded_at
        ORDER BY u.id DESC
    """).fetchall()
    conn.close()

    return jsonify([dict(r) for r in rows])


@app.route("/api/initiatives")
def get_initiatives():
    upload_id = request.args.get("upload_id", type=int) or latest_upload_id()
    if not upload_id:
        return jsonify([])

    conn = get_db()
    rows = conn.execute("""
        SELECT *
        FROM initiatives
        WHERE upload_id = ?
        ORDER BY relevance_score DESC, initiative_title ASC
    """, (upload_id,)).fetchall()
    conn.close()

    return jsonify([dict(r) for r in rows])


@app.route("/api/initiatives/<int:item_id>")
def get_initiative_detail(item_id: int):
    conn = get_db()
    row = conn.execute("SELECT * FROM initiatives WHERE id = ?", (item_id,)).fetchone()
    conn.close()

    if not row:
        return jsonify({"ok": False, "error": "Item not found."}), 404

    return jsonify(dict(row))


@app.route("/api/summary")
def get_summary():
    upload_id = request.args.get("upload_id", type=int) or latest_upload_id()
    if not upload_id:
        return jsonify({
            "total": 0,
            "high_relevance": 0,
            "high_impact": 0,
            "fmruk_relevant": 0,
            "top_themes": [],
            "top_owners": []
        })

    conn = get_db()
    rows = conn.execute("""
        SELECT theme, primary_owner, impact_level, relevance_score, is_fmruk_relevant
        FROM initiatives
        WHERE upload_id = ?
    """, (upload_id,)).fetchall()
    conn.close()

    items = [dict(r) for r in rows]

    theme_counts = {}
    owner_counts = {}
    for item in items:
        theme = item.get("theme") or "Unknown"
        owner = item.get("primary_owner") or "Unknown"
        theme_counts[theme] = theme_counts.get(theme, 0) + 1
        owner_counts[owner] = owner_counts.get(owner, 0) + 1

    top_themes = sorted(theme_counts.items(), key=lambda x: (-x[1], x[0]))[:5]
    top_owners = sorted(owner_counts.items(), key=lambda x: (-x[1], x[0]))[:5]

    return jsonify({
        "total": len(items),
        "high_relevance": sum(1 for x in items if (x.get("relevance_score") or 0) >= 80),
        "high_impact": sum(1 for x in items if x.get("impact_level") == "High"),
        "fmruk_relevant": sum(1 for x in items if x.get("is_fmruk_relevant") == 1),
        "top_themes": [{"name": k, "count": v} for k, v in top_themes],
        "top_owners": [{"name": k, "count": v} for k, v in top_owners],
    })


@app.route("/api/download/<int:upload_id>")
def download_uploaded_file(upload_id: int):
    conn = get_db()
    row = conn.execute("SELECT filename, saved_path FROM uploads WHERE id = ?", (upload_id,)).fetchone()
    conn.close()

    if not row:
        return jsonify({"ok": False, "error": "Upload not found."}), 404

    saved_path = Path(row["saved_path"])
    return send_from_directory(saved_path.parent, saved_path.name, as_attachment=True, download_name=row["filename"])


if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5000)
