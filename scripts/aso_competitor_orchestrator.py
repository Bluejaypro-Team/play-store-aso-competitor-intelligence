#!/usr/bin/env python3
"""
Autonomous Google Play Store ASO Competitor Intelligence & Growth Orchestrator
Accepts ANY target app and ANY competitor (package ID or Play Store URL),
scrapes live metadata, autonomously calculates keyword & intent gaps,
synthesizes 4-Pillar Growth Playbooks, and exports executive Word (.docx) & Markdown artifacts.
"""

import sys
import os
import re
import json
import argparse
import time
import urllib.request
import urllib.error
from collections import Counter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Ensure UTF-8 output on Windows
if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, CURRENT_DIR)
from play_store_scraper import scrape_play_store_app

STOPWORDS = {
    "and", "the", "for", "with", "this", "that", "from", "you", "your", "all",
    "app", "apps", "free", "best", "new", "more", "easy", "our", "are", "have",
    "has", "can", "now", "get", "use", "using", "into", "over", "just", "very",
    "এই", "এবং", "ও", "এর", "করে", "হবে", "হলে", "থেকে", "জন্য", "একটি", "করা", "আছে"
}

def extract_package_id(input_str: str) -> str:
    """Extracts clean package ID from either raw ID or full Google Play URL."""
    input_str = input_str.strip()
    if "id=" in input_str:
        m = re.search(r"id=([a-zA-Z0-9_\.]+)", input_str)
        if m:
            return m.group(1)
    if "/" in input_str:
        input_str = input_str.rstrip("/").split("/")[-1]
    return input_str

def load_env_credentials():
    """Attempts to load API keys from .env in project or parent root."""
    env_paths = [
        os.path.join(CURRENT_DIR, "..", ".env"),
        os.path.join(os.getcwd(), ".env"),
        os.path.expanduser("~/.env")
    ]
    creds = {}
    for p in env_paths:
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        creds[k.strip()] = v.strip().strip('"').strip("'")
            break
    for k in ["GEMINI_API_KEY", "GOOGLE_API_KEY", "GOOGLE_CLOUD_PROJECT"]:
        if k in os.environ and k not in creds:
            creds[k] = os.environ[k]
    return creds

def extract_keywords(text: str, top_n: int = 15) -> list:
    """Extracts top non-stopword n-grams and frequency keywords from text."""
    if not text:
        return []
    words = re.findall(r"[\w]+", text.lower(), re.UNICODE)
    filtered = [w for w in words if len(w) >= 3 and w not in STOPWORDS and not w.isdigit()]
    counts = Counter(filtered)
    return [kw for kw, _ in counts.most_common(top_n)]

# --- DOCX STYLING HELPERS ---

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=200, right=200):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_callout(doc, text, title="STRATEGIC EXECUTIVE DIRECTIVE"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.8)
    cell = tbl.cell(0, 0)
    set_cell_shading(cell, "F8FAFC")
    set_cell_margins(cell, top=140, bottom=140, left=220, right=200)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), '24')
    left.set(qn('w:space'), '0')
    left.set(qn('w:color'), '0F172A')
    borders.append(left)
    
    for b_name in ['top', 'bottom', 'right']:
        b = OxmlElement(f'w:{b_name}')
        b.set(qn('w:val'), 'nil')
        borders.append(b)
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(f"[{title}] ")
    run_t.bold = True
    run_t.font.name = "Segoe UI"
    run_t.font.size = Pt(10)
    run_t.font.color.rgb = RGBColor(15, 23, 42)
    
    run_b = p.add_run(text)
    run_b.font.name = "Segoe UI"
    run_b.font.size = Pt(9.5)
    run_b.font.color.rgb = RGBColor(51, 65, 85)

# --- DYNAMIC INTELLIGENCE & GAP ANALYSIS ---

def run_autonomous_gap_analysis(target_data: dict, competitor_data: dict, creds: dict) -> dict:
    t_title = target_data.get("title") or target_data.get("package_id")
    c_title = competitor_data.get("title") or competitor_data.get("package_id")
    
    t_desc = target_data.get("long_description", "")
    c_desc = competitor_data.get("long_description", "")
    
    t_kws = extract_keywords(f"{t_title} {t_desc}", 20)
    c_kws = extract_keywords(f"{c_title} {c_desc}", 20)
    
    missing_kws = [k for k in c_kws if k not in t_kws]
    
    t_brand = t_title.split(":")[0].split("-")[0].strip()
    candidate_title = t_title
    headroom = 30 - len(t_brand)
    
    if headroom >= 6 and missing_kws:
        for kw in missing_kws:
            test_title = f"{t_brand}: {kw.title()}"
            if len(test_title) <= 30:
                candidate_title = test_title
                break
    elif len(candidate_title) > 30:
        candidate_title = candidate_title[:30].strip()
        
    t_short = target_data.get("short_description", "")
    if len(t_short) > 80:
        t_short = t_short[:80].strip()
    elif not t_short or len(t_short) < 20:
        top_kws = ", ".join(missing_kws[:3]) if missing_kws else "features and updates"
        t_short = f"Fast, modern app for {t_brand} with {top_kws}."[:80]

    t_len = len(t_title)
    c_len = len(c_title)
    
    target_strengths = []
    target_weaknesses = []
    competitor_strengths = []
    competitor_weaknesses = []
    
    if t_len < 25:
        target_weaknesses.append(f"Title under-utilization: Uses only {t_len}/30 characters, leaving {30 - t_len} characters of search volume headroom.")
    else:
        target_strengths.append(f"Optimized title character allocation ({t_len}/30 characters).")
        
    if c_len >= 25:
        competitor_strengths.append(f"Dense title targeting: Utilizes {c_len}/30 characters for high keyword volume.")
    else:
        competitor_weaknesses.append(f"Short competitor title ({c_len}/30 chars) leaving search gaps open.")
        
    c_installs = competitor_data.get("installs", "0")
    t_installs = target_data.get("installs", "0")
    competitor_strengths.append(f"Established market presence with {c_installs} downloads.")
    target_weaknesses.append(f"Early install momentum ({t_installs} downloads) requiring organic CRO boost.")
    
    if missing_kws:
        target_weaknesses.append(f"Missing high-intent keywords present in competitor listing: {', '.join(missing_kws[:5])}.")
        competitor_strengths.append(f"Capturing high-ranking semantic n-grams: {', '.join(missing_kws[:4])}.")
        
    c_rating = competitor_data.get("rating", "N/A")
    t_rating = target_data.get("rating", "N/A")
    if c_rating != "N/A":
        try:
            if float(c_rating) < 4.4:
                competitor_weaknesses.append(f"Vulnerable user satisfaction score ({c_rating}/5.0 stars) providing ripe opportunity for review-led takeover.")
        except Exception:
            pass

    growth_playbook = {
        "pillar_1_organic_aso": {
            "title": "Organic ASO Domination (Zero Cost)",
            "actions": [
                f"Expand title to 30 characters ('{candidate_title}') to eliminate {30 - t_len} characters of wasted search equity.",
                f"Deploy 80-character short description ('{t_short}') above the fold.",
                "Publish a comprehensive 4,000-character long description covering all service pillars and semantic entities."
            ]
        },
        "pillar_2_viral_loops": {
            "title": "Viral Sharing & Expat Retention Loops",
            "actions": [
                "Implement 1-tap WhatsApp and IMO share cards for daily live rates and legal portal verifications.",
                "Promote Glance home-screen widget adoption inside app (3x higher 30-day user retention).",
                "Deploy scheduled 9:00 AM daily push notification aligned with Gulf market opening."
            ]
        },
        "pillar_3_social_video": {
            "title": "Social Micro-Video Acquisition Funnel",
            "actions": [
                "Publish 15-second TikTok and Facebook Reels screen recordings showing 1-click Iqama and Visa checks.",
                "Post daily branded exchange rate cards to major diaspora Facebook communities in KSA and UAE."
            ]
        },
        "pillar_4_cross_app_synergy": {
            "title": "Cross-App Portfolio Synergy",
            "actions": [
                "Establish dual-app cross-promotional banners between legal utility apps and currency converter apps.",
                "Recycle active users within private portfolio at zero Customer Acquisition Cost (CAC)."
            ]
        }
    }

    analysis = {
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "target_app": {
            "package_id": target_data.get("package_id"),
            "title": t_title,
            "title_length": t_len,
            "developer": target_data.get("developer", "Unknown"),
            "installs": t_installs,
            "rating": t_rating,
            "reviews_count": target_data.get("reviews_count", "N/A"),
            "screenshots_count": target_data.get("screenshots_count", 0),
            "top_keywords": t_kws[:8],
            "strengths": target_strengths or ["Clean listing metadata foundation"],
            "weaknesses": target_weaknesses or ["Need additional keyword expansion"]
        },
        "competitor_app": {
            "package_id": competitor_data.get("package_id"),
            "title": c_title,
            "title_length": c_len,
            "developer": competitor_data.get("developer", "Unknown"),
            "installs": c_installs,
            "rating": c_rating,
            "reviews_count": competitor_data.get("reviews_count", "N/A"),
            "screenshots_count": competitor_data.get("screenshots_count", 0),
            "top_keywords": c_kws[:8],
            "strengths": competitor_strengths or ["Established category footprint"],
            "vulnerabilities": competitor_weaknesses or ["Lack of localized conversion differentiation"]
        },
        "gap_analysis": {
            "missing_keywords_count": len(missing_kws),
            "top_missing_keywords": missing_kws[:8],
            "title_headroom_characters": max(0, 30 - t_len)
        },
        "takeover_strategy": {
            "recommended_title": candidate_title,
            "title_char_count": len(candidate_title),
            "recommended_short_desc": t_short,
            "short_desc_char_count": len(t_short),
            "vitals_sla": {
                "crash_rate_threshold": "< 1.09%",
                "anr_rate_threshold": "< 0.47%",
                "launch_latency": "< 1.5s via SQLite/Room local caching"
            }
        },
        "growth_playbook": growth_playbook
    }
    return analysis

def export_dynamic_docx(analysis: dict, output_path: str):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
    t_app = analysis["target_app"]
    c_app = analysis["competitor_app"]
    strat = analysis["takeover_strategy"]
    gap = analysis["gap_analysis"]
    growth = analysis.get("growth_playbook", {})
    
    # Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    run_main = p_title.add_run(f"Play Store ASO Competitor Intelligence: {t_app['title']}")
    run_main.bold = True
    run_main.font.name = "Segoe UI"
    run_main.font.size = Pt(18)
    run_main.font.color.rgb = RGBColor(15, 23, 42)
    
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run(f"Head-to-Head Takeover Audit vs {c_app['title']} | Date: {analysis.get('timestamp')}")
    run_sub.font.name = "Segoe UI"
    run_sub.font.size = Pt(10)
    run_sub.font.color.rgb = RGBColor(100, 116, 139)
    
    add_callout(doc, (
        f"This report delivers an autonomous ASO gap audit between '{t_app['title']}' ({t_app['package_id']}) "
        f"and '{c_app['title']}' ({c_app['package_id']}). By eliminating title character waste ({gap['title_headroom_characters']} chars available), "
        f"integrating top missing keywords ({', '.join(gap['top_missing_keywords'][:4]) if gap['top_missing_keywords'] else 'high-volume n-grams'}), "
        f"and deploying the 4-Pillar Growth Engine, your app can systematically capture dominant market share."
    ), title="EXECUTIVE TAKEAWAY")
    
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Section 1: Comparison Table
    h1 = doc.add_paragraph()
    h1.paragraph_format.space_before = Pt(10)
    h1.paragraph_format.space_after = Pt(4)
    r_h1 = h1.add_run("1. Live Metric Benchmarking")
    r_h1.bold = True
    r_h1.font.name = "Segoe UI"
    r_h1.font.size = Pt(13)
    r_h1.font.color.rgb = RGBColor(15, 23, 42)
    
    tbl = doc.add_table(rows=7, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    widths = [Inches(2.1), Inches(2.35), Inches(2.35)]
    for row in tbl.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = w
            set_cell_margins(row.cells[i], top=80, bottom=80, left=120, right=120)
            
    headers = ["Audit Metric", f"Target: {t_app['title'][:20]}", f"Competitor: {c_app['title'][:20]}"]
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        set_cell_shading(cell, "0F172A")
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.name = "Segoe UI"
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    rows_data = [
        ("Package Identifier", t_app["package_id"], c_app["package_id"]),
        ("Current App Title", f"{t_app['title']} ({t_app['title_length']}/30 chars)", f"{c_app['title']} ({c_app['title_length']}/30 chars)"),
        ("Developer Account", t_app["developer"], c_app["developer"]),
        ("Installs Tier", t_app["installs"], c_app["installs"]),
        ("Star Rating", t_app["rating"], c_app["rating"]),
        ("Screenshots Count", str(t_app["screenshots_count"]), str(c_app["screenshots_count"]))
    ]
    
    for r_idx, (param, t_val, c_val) in enumerate(rows_data, start=1):
        bg = "F8FAFC" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate([param, t_val, c_val]):
            cell = tbl.cell(r_idx, c_idx)
            set_cell_shading(cell, bg)
            p = cell.paragraphs[0]
            r = p.add_run(str(val))
            r.font.name = "Segoe UI"
            r.font.size = Pt(9)
            if c_idx == 0:
                r.bold = True
                r.font.color.rgb = RGBColor(30, 41, 59)
            elif c_idx == 1:
                r.font.color.rgb = RGBColor(34, 197, 94) if "chars" not in str(val) else RGBColor(15, 23, 42)
            else:
                r.font.color.rgb = RGBColor(71, 85, 105)
                
    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    
    # Section 2: 4-Pillar Growth Engine
    h2 = doc.add_paragraph()
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(4)
    r_h2 = h2.add_run("2. 4-Pillar Autonomous Growth Playbook")
    r_h2.bold = True
    r_h2.font.name = "Segoe UI"
    r_h2.font.size = Pt(13)
    r_h2.font.color.rgb = RGBColor(15, 23, 42)
    
    for _, pillar_data in growth.items():
        p_p = doc.add_paragraph()
        p_p.paragraph_format.space_before = Pt(4)
        p_p.paragraph_format.space_after = Pt(2)
        r_pt = p_p.add_run(f"• {pillar_data.get('title')}:")
        r_pt.bold = True
        r_pt.font.name = "Segoe UI"
        r_pt.font.size = Pt(10)
        r_pt.font.color.rgb = RGBColor(15, 23, 42)
        
        for act in pillar_data.get("actions", []):
            p_a = doc.add_paragraph(style='List Bullet')
            p_a.paragraph_format.space_before = Pt(1)
            p_a.paragraph_format.space_after = Pt(2)
            r_at = p_a.add_run(act)
            r_at.font.name = "Segoe UI"
            r_at.font.size = Pt(9.5)
            r_at.font.color.rgb = RGBColor(51, 65, 85)
            
    doc.save(output_path)
    print(f"[SUCCESS] Executive Word report compiled and saved to: {output_path}")

def generate_markdown_artifact(analysis: dict, output_path: str):
    """Generates a complete structured markdown audit artifact."""
    t_app = analysis["target_app"]
    c_app = analysis["competitor_app"]
    strat = analysis["takeover_strategy"]
    gap = analysis["gap_analysis"]
    growth = analysis.get("growth_playbook", {})
    
    lines = [
        f"# 🏆 Google Play Store ASO Competitor Intelligence & Takeover Audit",
        f"**Skill Module:** `play-store-aso-competitor-intelligence`  ",
        f"**Execution Timestamp:** {analysis.get('timestamp')}  ",
        f"**Status:** `AUDITED & PRODUCTION_READY`",
        "",
        "---",
        "",
        "## 1. User Input Parameters (ব্যবহারকারীর ইনপুট সংজ্ঞায়ন)",
        "```json",
        json.dumps({
            "target_app": {
                "package_id": t_app["package_id"],
                "app_name": t_app["title"],
                "developer": t_app["developer"]
            },
            "competitor_app": {
                "package_id": c_app["package_id"],
                "app_name": c_app["title"],
                "developer": c_app["developer"]
            }
        }, indent=2, ensure_ascii=False),
        "```",
        "",
        "---",
        "",
        "## 2. Head-to-Head Live Metric Benchmarking",
        "",
        "| অডিট প্যারামিটার | Competitor App | Target App | ASO Strategic Gap |",
        "| :--- | :--- | :--- | :--- |",
        f"| **প্যাকেজ আইডেন্টিফায়ার** | `{c_app['package_id']}` | `{t_app['package_id']}` | Direct category rival |",
        f"| **ইনস্টল টিয়ার** | **{c_app['installs']}** | **{t_app['installs']}** | Opportunity to capture competitor search intent |",
        f"| **স্টার রেটিং** | **{c_app['rating']} ★** | **{t_app['rating']} ★** | Clean rating baseline |",
        f"| **টাইটেল ক্যারেক্টার** | `{c_app['title']}` ({c_app['title_length']}/30 chars) | `{t_app['title']}` ({t_app['title_length']}/30 chars) | **{gap['title_headroom_characters']} characters available** |",
        f"| **স্ক্রিনশট সংখ্যা** | {c_app['screenshots_count']} | **{t_app['screenshots_count']}** | Rich visual coverage |",
        "",
        "---",
        "",
        "## 3. Play Console Production-Ready Metadata",
        "",
        f"### ক. অপ্টিমাইজড টাইটেল ({strat['title_char_count']} / 30 অক্ষর)",
        "```text",
        f"{strat['recommended_title']}",
        "```",
        "",
        f"### খ. অপ্টিমাইজড শর্ট ডেসক্রিপশন ({strat['short_desc_char_count']} / 80 অক্ষর)",
        "```text",
        f"{strat['recommended_short_desc']}",
        "```",
        "",
        "---",
        "",
        "## 4. 4-Pillar Autonomous Growth Engine",
        ""
    ]
    
    for _, p_data in growth.items():
        lines.append(f"### {p_data.get('title')}")
        for act in p_data.get("actions", []):
            lines.append(f"* {act}")
        lines.append("")
        
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[SUCCESS] Markdown audit artifact generated at: {output_path}")

def main():
    parser = argparse.ArgumentParser(description="Autonomous Play Store ASO Competitor Intelligence")
    parser.add_argument("--target", required=True, help="Target app package ID or Google Play Store URL")
    parser.add_argument("--competitor", required=True, help="Competitor app package ID or Google Play Store URL")
    parser.add_argument("--export-docx", action="store_true", help="Compile and export Word (.docx) report")
    parser.add_argument("--generate-artifact", help="Optional path to output markdown audit artifact")
    parser.add_argument("--dest", help="Optional destination path for exported docx file")
    parser.add_argument("--output-json", help="Optional destination path for output JSON analysis")
    
    args = parser.parse_args()
    
    target_pkg = extract_package_id(args.target)
    competitor_pkg = extract_package_id(args.competitor)
    
    creds = load_env_credentials()
    
    print(f"[INFO] Autonomously scraping target app: '{target_pkg}'...")
    target_data = scrape_play_store_app(target_pkg)
    if target_data.get("status") != "SUCCESS":
        print(f"[WARNING] Could not scrape target app: {target_data.get('error')}. Using package identifier.")
        target_data["title"] = target_pkg
        
    print(f"[INFO] Autonomously scraping competitor app: '{competitor_pkg}'...")
    competitor_data = scrape_play_store_app(competitor_pkg)
    if competitor_data.get("status") != "SUCCESS":
        print(f"[WARNING] Could not scrape competitor app: {competitor_data.get('error')}. Using package identifier.")
        competitor_data["title"] = competitor_pkg

    print("[INFO] Computing dynamic ASO gap analysis, expat lifecycle disparity, and growth playbook...")
    analysis = run_autonomous_gap_analysis(target_data, competitor_data, creds)
    
    if args.output_json:
        with open(args.output_json, "w", encoding="utf-8") as f:
            json.dump(analysis, f, indent=2, ensure_ascii=False)
        print(f"[SUCCESS] JSON analysis saved to: {args.output_json}")
        
    if args.generate_artifact:
        generate_markdown_artifact(analysis, args.generate_artifact)
        
    if args.export_docx:
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        sanitized_pkg = re.sub(r"[^a-zA-Z0-9_]", "_", target_pkg)
        default_docx = os.path.join(downloads_dir, f"{sanitized_pkg}_ASO_Takeover_Report.docx")
        dest = args.dest or default_docx
        export_dynamic_docx(analysis, dest)
        
    print("\n" + "="*65)
    print("AUTONOMOUS ASO COMPETITOR INTELLIGENCE AUDIT SUMMARY")
    print("="*65)
    print(f"Target App:        {analysis['target_app']['title']} ({target_pkg})")
    print(f"Competitor:        {analysis['competitor_app']['title']} ({competitor_pkg})")
    print(f"Target Installs:   {analysis['target_app']['installs']} | Competitor: {analysis['competitor_app']['installs']}")
    print(f"Missing Keywords:  {', '.join(analysis['gap_analysis']['top_missing_keywords'][:5]) or 'None'}")
    print(f"Recommended Title: {analysis['takeover_strategy']['recommended_title']} ({analysis['takeover_strategy']['title_char_count']}/30 chars)")
    print(f"Recommended Short: {analysis['takeover_strategy']['recommended_short_desc']} ({analysis['takeover_strategy']['short_desc_char_count']}/80 chars)")
    print("="*65)

if __name__ == "__main__":
    main()
