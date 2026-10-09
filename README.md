# Google Play Store ASO & Competitor Intelligence Studio

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Google Play ASO](https://img.shields.io/badge/Google%20Play-ASO%202026%20Compliant-brightgreen.svg)](https://play.google.com/console)
[![Status: Production](https://img.shields.io/badge/Status-Production%20Ready-success.svg)](#)

> Autonomous Google Play Store App Store Optimization (ASO) competitor scraping, gap analysis, and conversion engineering studio. Ingests target and competitor package IDs, extracts live store metadata, performs head-to-head keyword and feature gap audits, generates policy-compliant metadata, drafts corridor Custom Store Listings (CSL), and exports executive `.docx` reports.

---

## 🌟 Key Capabilities

* 🔍 **Live Play Store Scraping**: Rate-limited (1 req/sec) HTML metadata extractor with exponential backoff for public Google Play Store web listings.
* ⚖️ **Head-to-Head Competitor Benchmarking**: Directly contrasts keyword usage, title character lengths (against the strict 30-character boundary), install velocity tiers, and star ratings.
* 🛡️ **Automated Policy Boundary Gates**:
  * **Title Limit**: Strictly $\le 30$ characters.
  * **Short Description Limit**: Strictly $\le 80$ characters.
  * **Long Description Limit**: Strictly $\le 4,000$ characters.
  * **Superlative Ban**: Automatically detects and flags prohibited claims (`"Best"`, `"Free"`, `"Top"`, `"#1"`, `"No.1"`, `"Download Now"`).
  * **N-Gram Density**: Enforces natural 2.0% - 2.5% semantic keyword density to prevent algorithmic keyword stuffing penalties.
* 🌍 **Regional Custom Store Listings (CSL)**: Formats localized corridor store listings tailored to expatriate diaspora hubs (e.g., Saudi Arabia SAR / STC Pay, UAE AED / Al Ansari, Qatar QAR).
* 📄 **Executive Word (.docx) Delivery**: Automatically compiles professional, styled executive reports with tables, strategy callouts, and recommendations directly to the user's `Downloads` folder.

---

## 🏗️ Architecture

```
[Target Package ID] + [Competitor Package ID]
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

---

## 🚀 Quick Start

### 1. Prerequisites & Environment Setup
* Python 3.9+
* Required packages:
  ```bash
  pip install python-docx
  ```
* **Environment Configuration (`.env`)**:
  Copy the provided `.env.example` to `.env`:
  ```bash
  cp .env.example .env
  ```
  Fill in your specific API keys and Google Cloud Project identifiers:
  * `GEMINI_API_KEY`: Used for search grounding and semantic intent research.
  * `GOOGLE_CLOUD_PROJECT`: Your GCP project ID (for Vertex AI and Application Default Credentials).
  * `CLOUD_NATURAL_LANGUAGE_API_KEY`: For competitor review sentiment & entity salience audits.
  * `GOOGLE_PLAY_SERVICE_ACCOUNT_JSON`: (Optional) For automated Play Console deployments.

### 2. End-to-End Competitor Audit & .docx Export (100% Autonomous for ANY App)
Run the full comparative audit for **any Android app** and **any competitor** in any category (Finance, Health, Utilities, Education, E-Commerce, Gaming):

You can pass either **raw Package IDs** or **full Play Store URLs**:

```bash
# Example A: Using Package IDs (Any Category)
python scripts/aso_competitor_orchestrator.py \
  --target "<YOUR_TARGET_PACKAGE_ID>" \
  --competitor "<COMPETITOR_PACKAGE_ID>" \
  --export-docx

# Example B: Using Direct Google Play Web URLs
python scripts/aso_competitor_orchestrator.py \
  --target "https://play.google.com/store/apps/details?id=<YOUR_APP_ID>" \
  --competitor "https://play.google.com/store/apps/details?id=<COMPETITOR_APP_ID>" \
  --export-docx
```

The orchestrator will autonomously:
1. **Scrape live metadata** for both applications (Title, Developer, Installs, Ratings, Reviews, Descriptions, Screenshots) with built-in rate-limiting.
2. **Execute Head-to-Head Gap Analysis**: Measures title character headroom ($\le 30$ chars), detects missing semantic keywords, and analyzes rating disparities.
3. **Formulate Takeover Strategy**: Synthesizes a policy-compliant 30-char Title, 80-char Short Description, and Android Vitals targets.
4. **Compile Executive Word Report**: Saves the complete audit directly to `C:\Users\User\Downloads\<target_app>_ASO_Takeover_Report.docx`.

### 3. Standalone App Scraping
To inspect any individual app on Google Play:
```bash
python scripts/play_store_scraper.py --package "<PACKAGE_ID_OR_PLAY_STORE_URL>" --output "app_metadata.json"
```

---

## 📋 Google Play Store Compliance Matrix

| Rule ID | Rule Invariant Name | Boundary & Policy Standard |
| :--- | :--- | :--- |
| **RULE-ASO-1** | Strict Character Limits | Title $\le 30$, Short Description $\le 80$, Long Description $\le 4,000$ |
| **RULE-ASO-2** | Superlative Prohibition | Zero tolerance for `"Best"`, `"Free"`, `"Top"`, `"#1"`, `"No.1"`, `"Download Now"` |
| **RULE-ASO-3** | Semantic N-Gram Density | Primary target keywords strictly bounded to **2.0% - 2.5%** natural density |
| **RULE-ASO-4** | Visual CRO 3-Second Rule | Top 30% area of all 5 screenshots reserved for bold, high-contrast localized copy |
| **RULE-ASO-5** | Corridor CSL Invariant | Dedicated CSLs for KSA (SAR/STC Pay) and UAE (AED/Al Ansari) with localized hotlines |
| **RULE-ASO-6** | Android Vitals SLA | Crash Rate $< 1.09\%$, ANR Rate $< 0.47\%$, Cold Launch $< 1.5$ seconds via SQLite caching |
| **RULE-ASO-7** | Listing Experiment Gate | Mandatory 50/50 Play Console A/B testing for 7-14 days before 100% full rollout |

---

## 📁 Repository Structure

```
├── README.md                                # Project documentation & quick start guide
├── LICENSE                                  # MIT License (2026 Bluejaypro Team)
├── manifest.json                            # Antigravity skill manifest & address points
├── SKILL.md                                 # Complete agent skill specification
├── .gitignore                               # Git ignore configuration
├── references/
│   └── policy_gates.json                    # Play Store boundary limits & banned words
└── scripts/
    ├── aso_competitor_orchestrator.py       # Master CLI orchestrator & .docx compiler
    └── play_store_scraper.py                # Rate-limited Play Store web scraper
```

---

## 📜 License

Distributed under the **MIT License**. See [`LICENSE`](LICENSE) for more information.

---

<p align="center">
  Developed by <b>Bluejaypro Team</b> | Engineered for High-Converting App Store Dominance
</p>
