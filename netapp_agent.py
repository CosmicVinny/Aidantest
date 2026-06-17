#!/usr/bin/env python3
"""
NetApp AI/Data Prospecting Agent v2.2
Scores accounts, assigns tiers/angles, detects AIDE targets,
and builds a multi-tab Excel prospecting workbook.

Usage:
    python3 netapp_agent.py --csv accounts.csv --rep "Vinny" --outdir ./output
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

try:
    import pandas as pd
    import openpyxl
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, GradientFill
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    print("Installing required packages...")
    os.system("pip install openpyxl pandas --break-system-packages -q")
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

# ── Column detection map ─────────────────────────────────────────────────────
COLUMN_MAP = {
    "company":       ["company", "account name", "account", "name", "organization", "org name", "client"],
    "city":          ["city", "town", "municipality"],
    "state":         ["state", "province", "region", "st"],
    "address":       ["address", "street", "street address", "addr"],
    "industry":      ["industry", "vertical", "sector", "industry vertical", "primary industry"],
    "employees":     ["employees", "employee count", "headcount", "# employees", "emp count", "size"],
    "intent_topics": ["intent topics", "intent", "zoominfo intent", "intent signals", "topics"],
    "ai_keywords":   ["ai keywords", "keywords", "ai tags", "tags", "key terms"],
    "description":   ["description", "company description", "about", "overview", "summary", "notes"],
    "acct_num":      ["acct #", "account #", "account number", "zoominfo id", "crm id", "id", "acct id"],
}

# ── Scoring sets ─────────────────────────────────────────────────────────────
HIGH_INTENT_TOPICS = {
    "generative ai", "machine learning", "autonomous intelligence",
    "agentic ai", "autonomous intelligence and agentic ai",
    "high-performance computing", "all flash array",
}
MED_INTENT_TOPICS = {
    "artificial intelligence", "backup & recovery", "cloud",
    "cloud computing", "data management", "data analytics",
}

HIGH_FIT_INDUSTRIES = {
    "software", "biotech", "genomics", "life sciences", "pharmaceuticals",
    "media & internet", "defense", "aerospace", "manufacturing",
    "telecommunications", "semiconductor", "computer software", "information technology",
    "technology", "it services", "media",
}
MED_FIT_INDUSTRIES = {
    "healthcare", "finance", "business services", "education",
    "hospitals", "professional services", "government",
    "hospitals & physicians clinics", "healthcare services",
}

AI_KEYWORDS = {
    "ai", "ml", "machine learning", "deep learning", "neural", "gpu", "llm",
    "generative", "genomic", "sequencing", "bioinformatics", "hpc", "render",
    "vfx", "post-production", "defense", "aerospace", "autonomous", "robotics",
    "data pipeline", "object storage", "nfs", "cloud-native", "kubernetes",
    "inference", "training", "biotech", "pharma", "semiconductor", "eda",
    "simulation", "broadcast", "post production", "media production",
}

AIDE_TRIGGER_TOPICS = {
    "generative ai", "machine learning", "autonomous intelligence",
    "agentic ai", "autonomous intelligence and agentic ai",
    "high-performance computing", "artificial intelligence",
}
AIDE_TRIGGER_KEYWORDS = {
    "ai", "ml", "machine learning", "deep learning", "neural", "gpu", "llm",
    "generative", "hpc", "inference", "training", "autonomous", "agentic",
}

# ── Angle templates ───────────────────────────────────────────────────────────
ANGLE_MAP = {
    "generative ai": "AI Infrastructure",
    "agentic ai":    "AI Infrastructure",
    "autonomous intelligence and agentic ai": "AI Infrastructure",
    "machine learning": "AI Infrastructure",
    "high-performance computing": "HPC / AI Compute",
    "all flash array": "Flash Storage Modernization",
    "backup & recovery": "Data Protection",
    "cloud": "Hybrid Cloud",
    "cloud computing": "Hybrid Cloud",
    "data management": "Data Management",
    "data analytics": "Data & Analytics",
    "artificial intelligence": "AI Infrastructure",
    "kubernetes": "Cloud-Native / Kubernetes",
    "object storage": "Object Storage / S3",
    "cybersecurity": "Data Security",
}

EMAIL_TEMPLATES = {
    "AI Infrastructure": {
        "subject": "NetApp + {company}: Powering Your AI Workloads",
        "body": (
            "Hi {first_name},\n\n"
            "I noticed {company} has been actively exploring Generative AI and Machine Learning — "
            "impressive moves in a fast-moving space.\n\n"
            "NetApp's AI-ready infrastructure (ONTAP, StorageGRID, AFF C-Series) is purpose-built "
            "for the storage throughput and data pipeline demands of modern AI/ML workloads. "
            "We're helping teams like yours cut time-to-insight by removing storage as the bottleneck.\n\n"
            "Would you have 20 minutes to explore how NetApp can accelerate {company}'s AI roadmap?\n\n"
            "Best,\nVinny"
        ),
    },
    "HPC / AI Compute": {
        "subject": "Scaling HPC & AI at {company} — NetApp Can Help",
        "body": (
            "Hi {first_name},\n\n"
            "High-performance computing environments demand fast, scalable, parallel storage — "
            "and that's exactly where NetApp shines.\n\n"
            "Our ONTAP and EF-Series platforms deliver the IOPS and throughput needed for demanding "
            "simulation, rendering, and AI training workloads. "
            "I'd love to show you what we've done for similar {industry} teams.\n\n"
            "Open to a quick call this week?\n\nBest,\nVinny"
        ),
    },
    "Flash Storage Modernization": {
        "subject": "Modernize {company}'s Storage with NetApp AFF",
        "body": (
            "Hi {first_name},\n\n"
            "As storage demands grow, all-flash infrastructure is becoming the new baseline. "
            "NetApp's AFF portfolio delivers NVMe-powered performance with industry-leading efficiency.\n\n"
            "I'd be happy to walk through how {company} could reduce latency and TCO "
            "while future-proofing for AI and cloud workloads.\n\n"
            "Worth a 20-minute conversation?\n\nBest,\nVinny"
        ),
    },
    "Data Protection": {
        "subject": "Ransomware-Proof Your Data at {company}",
        "body": (
            "Hi {first_name},\n\n"
            "With ransomware threats on the rise, immutable backups and fast recovery "
            "are non-negotiable. NetApp's SnapLock and BlueXP backup solutions offer "
            "air-gapped protection and sub-hour RTOs.\n\n"
            "I'd love to share how we've helped organizations like {company} "
            "build bulletproof data protection strategies.\n\n"
            "Do you have time for a brief call?\n\nBest,\nVinny"
        ),
    },
    "Hybrid Cloud": {
        "subject": "NetApp + {company}: Seamless Hybrid Cloud Storage",
        "body": (
            "Hi {first_name},\n\n"
            "I see {company} is investing in cloud — NetApp makes hybrid cloud seamless "
            "with consistent data management across AWS, Azure, and GCP.\n\n"
            "Our Cloud Volumes ONTAP and BlueXP platform give your team a unified view "
            "of data wherever it lives, with enterprise-grade protection.\n\n"
            "Would a quick intro call make sense?\n\nBest,\nVinny"
        ),
    },
    "Default": {
        "subject": "NetApp + {company}: Smarter Data Infrastructure",
        "body": (
            "Hi {first_name},\n\n"
            "I help companies in the {state} area modernize their data infrastructure "
            "for AI, cloud, and hybrid workloads.\n\n"
            "NetApp's unified storage platform is trusted by thousands of enterprises "
            "to deliver performance, resilience, and flexibility.\n\n"
            "Would you be open to a 20-minute call to explore fit?\n\nBest,\nVinny"
        ),
    },
}

# ── Territory detection ───────────────────────────────────────────────────────
TERRITORY_MAP = {
    frozenset(["Illinois", "Indiana", "Wisconsin"]): "Midwest",
    frozenset(["California", "Nevada", "Hawaii"]): "West Coast",
    frozenset(["New York", "New Jersey", "Connecticut", "Massachusetts"]): "Northeast",
    frozenset(["Texas", "Oklahoma", "Louisiana", "Arkansas"]): "South Central",
    frozenset(["Florida", "Georgia", "North Carolina", "South Carolina"]): "Southeast",
    frozenset(["Washington", "Oregon", "Idaho", "Montana"]): "Pacific Northwest",
    frozenset(["Colorado", "Utah", "New Mexico", "Arizona"]): "Mountain West",
    frozenset(["Ohio", "Michigan", "Pennsylvania", "Minnesota"]): "Great Lakes",
}

STATE_ABBR = {
    "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR",
    "California": "CA", "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE",
    "Florida": "FL", "Georgia": "GA", "Hawaii": "HI", "Idaho": "ID",
    "Illinois": "IL", "Indiana": "IN", "Iowa": "IA", "Kansas": "KS",
    "Kentucky": "KY", "Louisiana": "LA", "Maine": "ME", "Maryland": "MD",
    "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN", "Mississippi": "MS",
    "Missouri": "MO", "Montana": "MT", "Nebraska": "NE", "Nevada": "NV",
    "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM", "New York": "NY",
    "North Carolina": "NC", "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK",
    "Oregon": "OR", "Pennsylvania": "PA", "Rhode Island": "RI", "South Carolina": "SC",
    "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX", "Utah": "UT",
    "Vermont": "VT", "Virginia": "VA", "Washington": "WA", "West Virginia": "WV",
    "Wisconsin": "WI", "Wyoming": "WY", "District of Columbia": "DC",
}

# ── Excel style helpers ───────────────────────────────────────────────────────
def _font(bold=False, size=11, color="000000", name="Calibri"):
    return Font(bold=bold, size=size, color=color, name=name)

def _fill(hex_color):
    return PatternFill(fill_type="solid", fgColor=hex_color)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

TIER_COLORS = {
    "S-Tier": ("1F3864", "FFFFFF"),  # dark navy bg, white text
    "A-Tier": ("2E75B6", "FFFFFF"),  # blue bg, white text
    "B-Tier": ("BDD7EE", "000000"),  # light blue bg, black text
    "C-Tier": ("EDEDED", "666666"),  # gray bg, dark gray text
}

AIDE_COLOR = ("7030A0", "FFFFFF")   # purple bg, white text

def style_header_row(ws, row_num, bg_hex="1F3864", fg_hex="FFFFFF", size=11):
    for cell in ws[row_num]:
        if cell.value is not None:
            cell.font = _font(bold=True, size=size, color=fg_hex)
            cell.fill = _fill(bg_hex)
            cell.alignment = _align(h="center", v="center")
            cell.border = _border()

def auto_width(ws, min_width=10, max_width=60):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = str(cell.value or "")
                max_len = max(max_len, len(val))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)

# ── Column detection ──────────────────────────────────────────────────────────
def detect_columns(headers):
    header_lower = {h.lower().strip(): h for h in headers}
    mapping = {}
    for field, aliases in COLUMN_MAP.items():
        for alias in aliases:
            if alias in header_lower:
                mapping[field] = header_lower[alias]
                break
    return mapping

# ── Data loading ──────────────────────────────────────────────────────────────
def load_csv(csv_path):
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
    if not rows:
        sys.exit(f"ERROR: CSV is empty: {csv_path}")
    col_map = detect_columns(list(rows[0].keys()))
    print(f"  Loaded {len(rows)} rows from {csv_path}")
    print(f"  Detected columns: {col_map}")
    return rows, col_map

def get_field(row, col_map, field, default=""):
    col = col_map.get(field)
    if col and col in row:
        return (row[col] or "").strip()
    return default

# ── Scoring ───────────────────────────────────────────────────────────────────
def parse_topics(topics_str):
    if not topics_str:
        return []
    return [t.strip().lower() for t in re.split(r"[;,|]", topics_str) if t.strip()]

def score_intent(topics):
    score = 0
    for t in topics:
        if t in HIGH_INTENT_TOPICS:
            score += 12
        elif t in MED_INTENT_TOPICS:
            score += 6
        else:
            score += 2
    return min(score, 35)

def score_industry(industry):
    ind_lower = industry.lower()
    if any(h in ind_lower for h in HIGH_FIT_INDUSTRIES):
        return 15
    if any(m in ind_lower for m in MED_FIT_INDUSTRIES):
        return 8
    return 0

def score_ai_keywords(description, topics_str, industry):
    text = f"{description} {topics_str} {industry}".lower()
    hits = sum(1 for kw in AI_KEYWORDS if kw in text)
    return min(hits * 2, 10)

def score_employees(emp_str):
    try:
        emp = int(str(emp_str).replace(",", "").strip())
        if emp >= 5000:
            return 5
        if emp >= 1000:
            return 4
        if emp >= 500:
            return 3
        if emp >= 100:
            return 2
        return 1
    except (ValueError, TypeError):
        return 0

def score_acct(acct_num):
    return 5 if str(acct_num).strip() else 0

def score_account(row, col_map):
    company     = get_field(row, col_map, "company")
    industry    = get_field(row, col_map, "industry")
    employees   = get_field(row, col_map, "employees")
    topics_str  = get_field(row, col_map, "intent_topics")
    description = get_field(row, col_map, "description")
    acct_num    = get_field(row, col_map, "acct_num")

    topics = parse_topics(topics_str)

    s_intent   = score_intent(topics)
    s_industry = score_industry(industry)
    s_ai       = score_ai_keywords(description, topics_str, industry)
    s_emp      = score_employees(employees)
    s_acct     = score_acct(acct_num)

    total = s_intent + s_industry + s_ai + s_emp + s_acct

    return {
        "score": total,
        "score_intent": s_intent,
        "score_industry": s_industry,
        "score_ai": s_ai,
        "score_emp": s_emp,
        "score_acct": s_acct,
        "topics": topics,
        "topics_str": topics_str,
    }

def assign_tier(score):
    # Thresholds scaled to actual max score of 70
    if score >= 49:
        return "S-Tier"
    if score >= 35:
        return "A-Tier"
    if score >= 21:
        return "B-Tier"
    return "C-Tier"

def is_aide_target(topics, description, industry):
    text = f"{' '.join(topics)} {description} {industry}".lower()
    for kw in AIDE_TRIGGER_KEYWORDS:
        if kw in text:
            return True
    for t in topics:
        if t in AIDE_TRIGGER_TOPICS:
            return True
    return False

def best_angle(topics, industry):
    # Prioritize high-value angles first
    priority_order = [
        "generative ai", "agentic ai", "autonomous intelligence and agentic ai",
        "machine learning", "high-performance computing", "artificial intelligence",
        "all flash array", "backup & recovery", "kubernetes", "object storage",
        "cloud", "cloud computing", "data management", "data analytics", "cybersecurity",
    ]
    topic_set = set(topics)
    for key in priority_order:
        if key in topic_set and key in ANGLE_MAP:
            return ANGLE_MAP[key]
    for t in topics:
        if t in ANGLE_MAP:
            return ANGLE_MAP[t]
    ind_lower = industry.lower()
    if any(h in ind_lower for h in ["software", "technology", "it", "semiconductor"]):
        return "AI Infrastructure"
    if any(h in ind_lower for h in ["healthcare", "biotech", "pharma", "genomics"]):
        return "AI Infrastructure"
    if any(h in ind_lower for h in ["manufacturing"]):
        return "Data Management"
    if any(h in ind_lower for h in ["finance", "financial"]):
        return "Data Protection"
    return "Default"

# ── Territory detection ───────────────────────────────────────────────────────
def detect_territory(rows, col_map):
    states = Counter(get_field(r, col_map, "state") for r in rows if get_field(r, col_map, "state"))
    if not states:
        return "Unknown"
    top_states = {s for s, _ in states.most_common(5)}
    best_match = None
    best_overlap = 0
    for territory_states, name in TERRITORY_MAP.items():
        overlap = len(top_states & territory_states)
        if overlap > best_overlap:
            best_overlap = overlap
            best_match = name
    if best_match and best_overlap >= 1:
        return best_match
    # Fallback: top state abbreviation
    top_state = states.most_common(1)[0][0]
    return STATE_ABBR.get(top_state, top_state[:2].upper())

# ── Build scored records ───────────────────────────────────────────────────────
def build_records(rows, col_map):
    records = []
    for row in rows:
        s = score_account(row, col_map)
        tier = assign_tier(s["score"])
        company     = get_field(row, col_map, "company")
        city        = get_field(row, col_map, "city")
        state       = get_field(row, col_map, "state")
        industry    = get_field(row, col_map, "industry")
        employees   = get_field(row, col_map, "employees")
        description = get_field(row, col_map, "description")
        acct_num    = get_field(row, col_map, "acct_num")

        aide = is_aide_target(s["topics"], description, industry)
        angle = best_angle(s["topics"], industry)

        records.append({
            "company":      company,
            "city":         city,
            "state":        state,
            "industry":     industry,
            "employees":    employees,
            "acct_num":     acct_num,
            "intent_topics": s["topics_str"],
            "description":  description,
            "score":        s["score"],
            "tier":         tier,
            "aide_target":  "YES" if aide else "",
            "angle":        angle,
            "score_intent":   s["score_intent"],
            "score_industry": s["score_industry"],
            "score_ai":       s["score_ai"],
            "score_emp":      s["score_emp"],
            "score_acct":     s["score_acct"],
        })

    records.sort(key=lambda x: x["score"], reverse=True)
    return records

# ── Tab 1: Account Tracker ─────────────────────────────────────────────────────
def write_account_tracker(wb, records):
    ws = wb.create_sheet("Account Tracker")
    ws.freeze_panes = "A2"

    headers = [
        "Rank", "Company", "City", "State", "Industry", "Employees",
        "Score", "Tier", "AIDE", "Angle",
        "Intent Topics", "Description", "Acct #",
        "Intent Pts", "Industry Pts", "AI Pts", "Emp Pts", "Acct Pts",
    ]
    ws.append(headers)
    style_header_row(ws, 1, bg_hex="1F3864", fg_hex="FFFFFF", size=11)

    for rank, rec in enumerate(records, start=1):
        row_data = [
            rank,
            rec["company"],
            rec["city"],
            rec["state"],
            rec["industry"],
            rec["employees"],
            rec["score"],
            rec["tier"],
            rec["aide_target"],
            rec["angle"],
            rec["intent_topics"],
            rec["description"],
            rec["acct_num"],
            rec["score_intent"],
            rec["score_industry"],
            rec["score_ai"],
            rec["score_emp"],
            rec["score_acct"],
        ]
        ws.append(row_data)

        row_num = rank + 1
        tier = rec["tier"]
        bg, fg = TIER_COLORS.get(tier, ("FFFFFF", "000000"))

        # Color tier column (H)
        tier_cell = ws.cell(row=row_num, column=8)
        tier_cell.fill = _fill(bg)
        tier_cell.font = _font(bold=True, color=fg, size=10)
        tier_cell.alignment = _align(h="center")

        # Color AIDE column (I) if flagged
        aide_cell = ws.cell(row=row_num, column=9)
        if rec["aide_target"]:
            aide_cell.fill = _fill(AIDE_COLOR[0])
            aide_cell.font = _font(bold=True, color=AIDE_COLOR[1], size=10)
            aide_cell.alignment = _align(h="center")

        # Score column (G)
        score_cell = ws.cell(row=row_num, column=7)
        score_cell.alignment = _align(h="center")
        score_cell.font = _font(bold=True, size=10)

        # Alternate row shading
        if rank % 2 == 0:
            for col in range(1, len(headers) + 1):
                c = ws.cell(row=row_num, column=col)
                if c.fill.fill_type == "none" or (c.fill.fgColor.rgb in ("00000000", "FFFFFFFF")):
                    c.fill = _fill("F5F8FF")

    auto_width(ws)
    ws.column_dimensions["K"].width = 45  # Intent Topics
    ws.column_dimensions["L"].width = 55  # Description

    print(f"  Account Tracker: {len(records)} rows")

# ── Tab 2: AIDE Targets ───────────────────────────────────────────────────────
def write_aide_targets(wb, records):
    ws = wb.create_sheet("AIDE Targets")
    ws.freeze_panes = "A2"

    aide_records = [r for r in records if r["aide_target"]]

    headers = [
        "Rank", "Company", "City", "State", "Industry", "Employees",
        "Score", "Tier", "Angle", "Intent Topics", "Description", "Acct #",
    ]
    ws.append(headers)
    style_header_row(ws, 1, bg_hex="7030A0", fg_hex="FFFFFF", size=11)

    for rank, rec in enumerate(aide_records, start=1):
        row_data = [
            rank,
            rec["company"],
            rec["city"],
            rec["state"],
            rec["industry"],
            rec["employees"],
            rec["score"],
            rec["tier"],
            rec["angle"],
            rec["intent_topics"],
            rec["description"],
            rec["acct_num"],
        ]
        ws.append(row_data)

        row_num = rank + 1
        tier = rec["tier"]
        bg, fg = TIER_COLORS.get(tier, ("FFFFFF", "000000"))

        tier_cell = ws.cell(row=row_num, column=8)
        tier_cell.fill = _fill(bg)
        tier_cell.font = _font(bold=True, color=fg, size=10)
        tier_cell.alignment = _align(h="center")

        if rank % 2 == 0:
            for col in range(1, len(headers) + 1):
                c = ws.cell(row=row_num, column=col)
                if c.fill.fill_type == "none" or (c.fill.fgColor.rgb in ("00000000", "FFFFFFFF")):
                    c.fill = _fill("F5EEFF")

    auto_width(ws)
    ws.column_dimensions["J"].width = 45
    ws.column_dimensions["K"].width = 55

    print(f"  AIDE Targets: {len(aide_records)} rows")

# ── Tab 3: Email Templates ─────────────────────────────────────────────────────
def write_email_templates(wb, records, rep_name):
    ws = wb.create_sheet("Email Templates")
    ws.freeze_panes = "A2"

    headers = [
        "Company", "State", "Industry", "Tier", "Angle",
        "Email Subject", "Email Body",
    ]
    ws.append(headers)
    style_header_row(ws, 1, bg_hex="2E75B6", fg_hex="FFFFFF", size=11)

    # Only top 100 by score for email templates
    top = records[:100]

    for rank, rec in enumerate(top, start=1):
        angle = rec["angle"]
        tmpl = EMAIL_TEMPLATES.get(angle, EMAIL_TEMPLATES["Default"])

        first_name = "{First Name}"
        subj = tmpl["subject"].format(
            company=rec["company"],
            state=rec["state"] or "your area",
            industry=rec["industry"] or "your industry",
            first_name=first_name,
        )
        body = tmpl["body"].format(
            company=rec["company"],
            state=rec["state"] or "your area",
            industry=rec["industry"] or "your industry",
            first_name=first_name,
        )

        row_data = [
            rec["company"],
            rec["state"],
            rec["industry"],
            rec["tier"],
            angle,
            subj,
            body,
        ]
        ws.append(row_data)

        row_num = rank + 1
        tier = rec["tier"]
        bg, fg = TIER_COLORS.get(tier, ("FFFFFF", "000000"))

        tier_cell = ws.cell(row=row_num, column=4)
        tier_cell.fill = _fill(bg)
        tier_cell.font = _font(bold=True, color=fg, size=10)
        tier_cell.alignment = _align(h="center")

        # Wrap body text
        body_cell = ws.cell(row=row_num, column=7)
        body_cell.alignment = _align(h="left", v="top", wrap=True)

    auto_width(ws)
    ws.column_dimensions["F"].width = 50
    ws.column_dimensions["G"].width = 80
    ws.row_dimensions[1].height = 20

    print(f"  Email Templates: {len(top)} rows")

# ── Tab 4: Buy Signal Cheat Sheet ─────────────────────────────────────────────
def write_cheat_sheet(wb, records):
    ws = wb.create_sheet("Buy Signal Cheat Sheet")

    # Header
    ws.merge_cells("A1:D1")
    title = ws["A1"]
    title.value = "NetApp Buy Signal Cheat Sheet"
    title.font = _font(bold=True, size=14, color="FFFFFF")
    title.fill = _fill("1F3864")
    title.alignment = _align(h="center", v="center")
    ws.row_dimensions[1].height = 28

    # Intent topic frequency from our data
    all_topics = []
    for rec in records:
        all_topics.extend(parse_topics(rec["intent_topics"]))
    topic_counts = Counter(all_topics)

    ws.append([""])  # blank row

    # Topic frequency table
    ws.append(["Intent Topic", "# Accounts", "NetApp Angle", "Key Products"])
    style_header_row(ws, 3, bg_hex="2E75B6", fg_hex="FFFFFF")

    product_map = {
        "generative ai":      "AFF C-Series, StorageGRID, ONTAP AI",
        "machine learning":   "AFF C-Series, ONTAP AI, EF-Series",
        "autonomous intelligence and agentic ai": "AFF C-Series, ONTAP AI",
        "agentic ai":         "AFF C-Series, ONTAP AI",
        "high-performance computing": "EF-Series, A-Series, AFF",
        "all flash array":    "AFF A-Series, AFF C-Series",
        "backup & recovery":  "SnapLock, BlueXP Backup, ONTAP SnapVault",
        "cloud":              "Cloud Volumes ONTAP, BlueXP, ANF/CVO",
        "cloud computing":    "Cloud Volumes ONTAP, BlueXP",
        "artificial intelligence": "AFF C-Series, ONTAP AI, StorageGRID",
        "data management":    "ONTAP, BlueXP, Cloud Volumes ONTAP",
        "data analytics":     "ONTAP, StorageGRID, Cloud Volumes ONTAP",
        "kubernetes":         "Trident, Cloud Volumes ONTAP",
        "object storage":     "StorageGRID, ONTAP S3",
        "cybersecurity":      "SnapLock, ONTAP ransomware protection",
    }

    row_idx = 4
    for topic, count in topic_counts.most_common(25):
        angle = ANGLE_MAP.get(topic, "General Data Infrastructure")
        products = product_map.get(topic, "ONTAP, AFF, StorageGRID")
        ws.append([topic.title(), count, angle, products])
        if row_idx % 2 == 0:
            for col in range(1, 5):
                ws.cell(row=row_idx, column=col).fill = _fill("F0F4FF")
        row_idx += 1

    ws.append([""])  # blank

    # Scoring legend
    leg_row = ws.max_row + 1
    ws.merge_cells(f"A{leg_row}:D{leg_row}")
    leg_title = ws.cell(row=leg_row, column=1)
    leg_title.value = "Tier Scoring Legend"
    leg_title.font = _font(bold=True, size=12, color="FFFFFF")
    leg_title.fill = _fill("1F3864")
    leg_title.alignment = _align(h="center")
    ws.row_dimensions[leg_row].height = 22

    tiers = [
        ("S-Tier", "70–100", "Priority accounts — immediate outreach", "1F3864", "FFFFFF"),
        ("A-Tier", "50–69",  "High value — schedule this week",        "2E75B6", "FFFFFF"),
        ("B-Tier", "30–49",  "Nurture — include in cadence",           "BDD7EE", "000000"),
        ("C-Tier", "0–29",   "Low priority — monitor intent",          "EDEDED", "666666"),
    ]
    for tier, rng, desc, bg, fg in tiers:
        r = ws.max_row + 1
        ws.cell(row=r, column=1, value=tier).fill = _fill(bg)
        ws.cell(row=r, column=1).font = _font(bold=True, color=fg)
        ws.cell(row=r, column=1).alignment = _align(h="center")
        ws.cell(row=r, column=2, value=rng).alignment = _align(h="center")
        ws.cell(row=r, column=3, value=desc)
        ws.cell(row=r, column=4, value="")

    auto_width(ws)
    print(f"  Buy Signal Cheat Sheet: {len(topic_counts)} topics")

# ── Tab 5: Pipeline Summary ────────────────────────────────────────────────────
def write_pipeline_summary(wb, records, rep_name, territory):
    ws = wb.create_sheet("Pipeline Summary")

    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value = f"Pipeline Summary — {rep_name} | {territory} | {datetime.now().strftime('%B %Y')}"
    t.font = _font(bold=True, size=14, color="FFFFFF")
    t.fill = _fill("1F3864")
    t.alignment = _align(h="center", v="center")
    ws.row_dimensions[1].height = 28

    ws.append([""])

    # Tier counts
    tier_counts = Counter(r["tier"] for r in records)
    aide_count  = sum(1 for r in records if r["aide_target"])
    with_intent = sum(1 for r in records if r["intent_topics"])
    with_zi     = sum(1 for r in records if r["acct_num"])

    ws.append(["Metric", "Count", "% of Total", "", ""])
    style_header_row(ws, 3, bg_hex="2E75B6", fg_hex="FFFFFF")

    total = len(records)
    summary_rows = [
        ("Total Accounts",        total,                       "100%"),
        ("S-Tier Accounts",       tier_counts.get("S-Tier",0), f"{tier_counts.get('S-Tier',0)/total*100:.1f}%"),
        ("A-Tier Accounts",       tier_counts.get("A-Tier",0), f"{tier_counts.get('A-Tier',0)/total*100:.1f}%"),
        ("B-Tier Accounts",       tier_counts.get("B-Tier",0), f"{tier_counts.get('B-Tier',0)/total*100:.1f}%"),
        ("C-Tier Accounts",       tier_counts.get("C-Tier",0), f"{tier_counts.get('C-Tier',0)/total*100:.1f}%"),
        ("AIDE Targets",          aide_count,                  f"{aide_count/total*100:.1f}%"),
        ("Accounts w/ Intent",    with_intent,                 f"{with_intent/total*100:.1f}%"),
        ("ZoomInfo Enriched",     with_zi,                     f"{with_zi/total*100:.1f}%"),
    ]

    for i, (metric, count, pct) in enumerate(summary_rows, start=4):
        ws.append([metric, count, pct, "", ""])
        if i % 2 == 0:
            for col in range(1, 4):
                ws.cell(row=i, column=col).fill = _fill("F0F4FF")

    ws.append([""])

    # Top industries
    ind_row = ws.max_row + 1
    ws.merge_cells(f"A{ind_row}:E{ind_row}")
    sec = ws.cell(row=ind_row, column=1)
    sec.value = "Top Industries by Account Count"
    sec.font = _font(bold=True, size=12, color="FFFFFF")
    sec.fill = _fill("2E75B6")
    sec.alignment = _align(h="center")
    ws.row_dimensions[ind_row].height = 20

    ws.append(["Industry", "# Accounts", "S/A-Tier", "AIDE Targets", ""])
    style_header_row(ws, ws.max_row, bg_hex="BDD7EE", fg_hex="000000")

    industry_stats = defaultdict(lambda: {"total": 0, "sa": 0, "aide": 0})
    for r in records:
        ind = r["industry"] or "Unknown"
        industry_stats[ind]["total"] += 1
        if r["tier"] in ("S-Tier", "A-Tier"):
            industry_stats[ind]["sa"] += 1
        if r["aide_target"]:
            industry_stats[ind]["aide"] += 1

    sorted_inds = sorted(industry_stats.items(), key=lambda x: x[1]["total"], reverse=True)
    for i, (ind, stats) in enumerate(sorted_inds[:15]):
        r_idx = ws.max_row + 1
        ws.append([ind, stats["total"], stats["sa"], stats["aide"], ""])
        if i % 2 == 0:
            for col in range(1, 5):
                ws.cell(row=r_idx, column=col).fill = _fill("F0F4FF")

    ws.append([""])

    # Top states
    state_row = ws.max_row + 1
    ws.merge_cells(f"A{state_row}:E{state_row}")
    sec2 = ws.cell(row=state_row, column=1)
    sec2.value = "Account Distribution by State"
    sec2.font = _font(bold=True, size=12, color="FFFFFF")
    sec2.fill = _fill("2E75B6")
    sec2.alignment = _align(h="center")
    ws.row_dimensions[state_row].height = 20

    ws.append(["State", "# Accounts", "S/A-Tier", "AIDE Targets", ""])
    style_header_row(ws, ws.max_row, bg_hex="BDD7EE", fg_hex="000000")

    state_stats = defaultdict(lambda: {"total": 0, "sa": 0, "aide": 0})
    for r in records:
        st = r["state"] or "Unknown"
        state_stats[st]["total"] += 1
        if r["tier"] in ("S-Tier", "A-Tier"):
            state_stats[st]["sa"] += 1
        if r["aide_target"]:
            state_stats[st]["aide"] += 1

    sorted_states = sorted(state_stats.items(), key=lambda x: x[1]["total"], reverse=True)
    for i, (st, stats) in enumerate(sorted_states[:15]):
        r_idx = ws.max_row + 1
        ws.append([st, stats["total"], stats["sa"], stats["aide"], ""])
        if i % 2 == 0:
            for col in range(1, 5):
                ws.cell(row=r_idx, column=col).fill = _fill("F0F4FF")

    auto_width(ws)
    print(f"  Pipeline Summary: {total} accounts, {aide_count} AIDE targets")

# ── Tab 6: Territory Profile ───────────────────────────────────────────────────
def write_territory_profile(wb, records, rep_name, territory):
    ws = wb.create_sheet("Territory Profile")

    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value = f"Territory Profile: {territory} | Rep: {rep_name}"
    t.font = _font(bold=True, size=14, color="FFFFFF")
    t.fill = _fill("7030A0")
    t.alignment = _align(h="center", v="center")
    ws.row_dimensions[1].height = 28

    ws.append([""])

    # Top intent topics
    sec_row = ws.max_row + 1
    ws.merge_cells(f"A{sec_row}:D{sec_row}")
    sec = ws.cell(row=sec_row, column=1)
    sec.value = "Active Intent Signals in Territory"
    sec.font = _font(bold=True, size=12, color="FFFFFF")
    sec.fill = _fill("2E75B6")
    sec.alignment = _align(h="center")
    ws.row_dimensions[sec_row].height = 20

    ws.append(["Intent Topic", "# Accounts", "High Priority?", "Recommended Angle"])
    style_header_row(ws, ws.max_row, bg_hex="BDD7EE", fg_hex="000000")

    all_topics = []
    for rec in records:
        all_topics.extend(parse_topics(rec["intent_topics"]))
    topic_counts = Counter(all_topics)

    for i, (topic, count) in enumerate(topic_counts.most_common(20)):
        priority = "YES" if topic in HIGH_INTENT_TOPICS else ("MED" if topic in MED_INTENT_TOPICS else "")
        angle = ANGLE_MAP.get(topic, "General Data Infrastructure")
        r_idx = ws.max_row + 1
        ws.append([topic.title(), count, priority, angle])
        if priority == "YES":
            ws.cell(row=r_idx, column=3).fill = _fill("FFD700")
            ws.cell(row=r_idx, column=3).font = _font(bold=True)
        elif priority == "MED":
            ws.cell(row=r_idx, column=3).fill = _fill("FFF2CC")
        if i % 2 == 0:
            for col in [1, 2, 4]:
                ws.cell(row=r_idx, column=col).fill = _fill("F5EEFF")

    ws.append([""])

    # Angle breakdown
    sec2_row = ws.max_row + 1
    ws.merge_cells(f"A{sec2_row}:D{sec2_row}")
    sec2 = ws.cell(row=sec2_row, column=1)
    sec2.value = "Recommended Angles by Account Count"
    sec2.font = _font(bold=True, size=12, color="FFFFFF")
    sec2.fill = _fill("2E75B6")
    sec2.alignment = _align(h="center")
    ws.row_dimensions[sec2_row].height = 20

    ws.append(["Angle", "# Accounts", "S/A-Tier", "AIDE Targets"])
    style_header_row(ws, ws.max_row, bg_hex="BDD7EE", fg_hex="000000")

    angle_stats = defaultdict(lambda: {"total": 0, "sa": 0, "aide": 0})
    for r in records:
        a = r["angle"]
        angle_stats[a]["total"] += 1
        if r["tier"] in ("S-Tier", "A-Tier"):
            angle_stats[a]["sa"] += 1
        if r["aide_target"]:
            angle_stats[a]["aide"] += 1

    for i, (angle, stats) in enumerate(sorted(angle_stats.items(), key=lambda x: x[1]["total"], reverse=True)):
        r_idx = ws.max_row + 1
        ws.append([angle, stats["total"], stats["sa"], stats["aide"]])
        if i % 2 == 0:
            for col in range(1, 5):
                ws.cell(row=r_idx, column=col).fill = _fill("F5EEFF")

    auto_width(ws)
    print(f"  Territory Profile: {territory}")

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="NetApp AI/Data Prospecting Agent v2.2")
    parser.add_argument("--csv",    required=True, help="Path to enriched accounts CSV")
    parser.add_argument("--rep",    required=True, help="Rep name (e.g. 'Vinny')")
    parser.add_argument("--outdir", default="./output", help="Output directory")
    args = parser.parse_args()

    print("\n" + "="*60)
    print("  NetApp AI/Data Prospecting Agent v2.2")
    print("="*60)

    # Load CSV
    print(f"\n[1/7] Loading CSV: {args.csv}")
    rows, col_map = load_csv(args.csv)

    # Score accounts
    print(f"\n[2/7] Scoring {len(rows)} accounts...")
    records = build_records(rows, col_map)
    tier_counts = Counter(r["tier"] for r in records)
    aide_count  = sum(1 for r in records if r["aide_target"])
    print(f"  S-Tier: {tier_counts.get('S-Tier',0)}")
    print(f"  A-Tier: {tier_counts.get('A-Tier',0)}")
    print(f"  B-Tier: {tier_counts.get('B-Tier',0)}")
    print(f"  C-Tier: {tier_counts.get('C-Tier',0)}")
    print(f"  AIDE Targets: {aide_count}")

    # Detect territory
    print(f"\n[3/7] Detecting territory...")
    territory = detect_territory(rows, col_map)
    print(f"  Territory: {territory}")

    # Build workbook
    print(f"\n[4/7] Building Excel workbook...")
    wb = openpyxl.Workbook()
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    write_account_tracker(wb, records)
    write_aide_targets(wb, records)
    write_email_templates(wb, records, args.rep)
    write_cheat_sheet(wb, records)
    write_pipeline_summary(wb, records, args.rep, territory)
    write_territory_profile(wb, records, args.rep, territory)

    # Save
    os.makedirs(args.outdir, exist_ok=True)
    rep_last = args.rep.split()[-1]
    outfile = os.path.join(
        args.outdir,
        f"{rep_last}_NetApp_{territory.replace(' ','_')}_Prospecting_Tracker_CHECKPOINT.xlsx"
    )

    print(f"\n[5/7] Saving workbook to: {outfile}")
    wb.save(outfile)

    print(f"\n[6/7] Writing scored CSV...")
    scored_csv = os.path.join(args.outdir, f"{rep_last}_NetApp_Scored_Accounts.csv")
    fieldnames = [
        "rank", "company", "city", "state", "industry", "employees",
        "score", "tier", "aide_target", "angle",
        "intent_topics", "description", "acct_num",
        "score_intent", "score_industry", "score_ai", "score_emp", "score_acct",
    ]
    with open(scored_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for rank, rec in enumerate(records, start=1):
            writer.writerow({"rank": rank, **{k: rec[k] for k in fieldnames if k != "rank"}})
    print(f"  Scored CSV: {scored_csv}")

    print(f"\n[7/7] Done!")
    print(f"\n{'='*60}")
    print(f"  Output: {outfile}")
    print(f"  Accounts: {len(records)}")
    print(f"  S-Tier: {tier_counts.get('S-Tier',0)} | A-Tier: {tier_counts.get('A-Tier',0)}")
    print(f"  AIDE Targets: {aide_count}")
    print(f"  Territory: {territory}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
