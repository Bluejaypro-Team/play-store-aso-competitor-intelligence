---
name: play-store-aso-competitor-intelligence
version: "1.0.0"
description: "Autonomous Google Play Store ASO competitor scraping, gap analysis, and conversion engineering studio. Ingests target and competitor package IDs, extracts live store metadata, performs head-to-head keyword and feature gap audits, generates policy-compliant metadata, drafts corridor Custom Store Listings (CSL), and exports executive .docx reports directly to the user's Downloads folder."
---

# Google Play Store ASO & Competitor Intelligence Studio

This skill provides an autonomous pipeline for scraping live Play Store listings, performing head-to-head competitive gap audits, verifying strict policy boundaries, and compiling executive Word (`.docx`) reports directly to the user's Downloads directory.

---

## 1. Capabilities & Workflow Architecture

```
[Target App Package ID] + [Competitor Package ID]
                     │
                     ▼
       ┌───────────────────────────┐
       │   play_store_scraper.py   │ ──► Rate-limited HTML extraction (1 req/sec)
       └─────────────┬─────────────┘
                     │
                     ▼
       ┌───────────────────────────┐
       │ aso_competitor_orchestrator│──► Head-to-Head Benchmarking & Gap Matrix
       └─────────────┬─────────────┘
                     │
         ┌───────────┴───────────┐
         ▼                       ▼
┌──────────────────┐    ┌──────────────────────────────────────────────┐
│ Play Console .md │    │ Executive .docx Report                       │
│ (Scratch folder) │    │ (C:/Users/User/Downloads/<App>_ASO_Report.docx│
└──────────────────┘    └──────────────────────────────────────────────┘
```

1. **Deterministic Metadata Ingestion**: Scrapes live Play Store titles, short/long descriptions, download tiers, star ratings, and review counts.
2. **Head-to-Head Competitive Benchmarking**: Directly contrasts keyword usage, title lengths (against the 30-character boundary), and vulnerability levers.
3. **Automated Policy & Limit Gates**: Verifies Title $\le 30$ chars, Short Description $\le 80$ chars, Long Description $\le 4,000$ chars, and screens for banned superlatives (`"Best"`, `"Free"`, `"Top"`, `"#1"`).
4. **GCC Diaspora Corridor Generation**: Drafts targeted Custom Store Listings (CSL) for Saudi Arabia (KSA) and United Arab Emirates (UAE).
5. **Executive Word Document Delivery**: Compiles beautifully formatted `.docx` reports with callouts and comparison tables directly to `C:/Users/User/Downloads/`.

---

## 2. Command Reference

### A. Run Full End-to-End Audit & Export Word Document
```bash
python C:/Users/User/.gemini/config/skills/play-store-aso-competitor-intelligence/scripts/aso_competitor_orchestrator.py --target "com.probashi.taka" --competitor "com.jameel.bdttakarrate" --export-docx
```

### B. Scrape a Single Play Store App
```bash
python C:/Users/User/.gemini/config/skills/play-store-aso-competitor-intelligence/scripts/play_store_scraper.py --package "com.probashi.taka" --output "C:/Users/User/.gemini/antigravity-ide/scratch/app_data.json"
```

### C. Output Comparative Analysis in JSON
```bash
python C:/Users/User/.gemini/config/skills/play-store-aso-competitor-intelligence/scripts/aso_competitor_orchestrator.py --target "com.probashi.taka" --competitor "com.jameel.bdttakarrate" --output-json "C:/Users/User/.gemini/antigravity-ide/scratch/gap_analysis.json"
```

---

## 3. Rules & Persistent Invariants

All metadata generation and competitive analysis authored by this skill must obey:
* **Rules Specification**: `C:/Users/User/.gemini/config/rules/play_store_aso_optimization_rules.md`
* **Skill Manifest**: `C:/Users/User/.gemini/config/skills/play-store-aso-competitor-intelligence/manifest.json`
* **Global Protocol**: `C:/Users/User/.gemini/config/AGENTS.md` (Protocol 25)

---

## 4. Dependencies
* `python-docx` ($\ge 1.1.0$)
* `play-store-aso-optimizer`
* `grounded-search-intent-architect`
* `prompt-vault-clipboard-manager`
