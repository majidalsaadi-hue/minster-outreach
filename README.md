# Ministry of Investment — Investor Relations CRM

A professional, single-user web application for the Relationship Manager at the
Minister of Investment of Saudi Arabia. Consolidates investor data, tracks
action items and meetings, monitors the opportunity pipeline, and auto-generates
executive reports for the Minister's office.

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Launch the tool

```bash
streamlit run app.py
```

The app opens automatically at `http://localhost:8501` in your browser.

---

## How to use the tool

### Upload updated data

1. Click **"Upload Updated Excel"** in the left sidebar
2. Select your `.xlsx` file
3. The tool validates the schema and loads all data immediately
4. A confirmation banner shows: investor count, meeting count, opportunity count, action item count

The tool accepts **two file formats**:
- **New format** (recommended): 5-sheet workbook — see *Excel Template* below
- **Legacy format**: the original per-investor Action Item Tracker sheets (e.g. `Action Items GA`, `Action Items Blackrock`) — automatically converted

### Navigate the tool

Use the left sidebar to switch between sections:

| Section | Purpose |
|---------|---------|
| **Dashboard** | Executive overview with KPI cards, alerts, pipeline charts |
| **Investors** | Full investor master list with drill-down profiles |
| **Meetings** | Meeting log — add and view all meetings |
| **Opportunities** | Investment opportunity pipeline |
| **Action Items** | Consolidated action item tracker across all investors |
| **RM Tasks** | Personal task list for the Relationship Manager |
| **Export & Reports** | Generate PowerPoint, PDF, and Excel exports |

### Switch language

Click **"العربية"** (Arabic) or **"English"** in the sidebar to toggle the interface.
Arabic layout switches to RTL automatically.

---

## How to generate and export reports

Go to **Export & Reports** in the sidebar:

### PowerPoint Report
- Mirrors the Executive Account Management Dashboard format
- Includes: title slide, executive KPI summary, pipeline overview, meeting outcomes,
  opportunities dashboard, sector/geography breakdown, per-investor slides, alerts slide
- Named: `MoI_Investor_Status_YYYY-MM-DD.pptx`
- Available in English or Arabic

### PDF Executive Summary
- Clean single-document for Minister briefing
- Includes: KPI banner, investor portfolio, active opportunities, overdue actions
- Named: `MoI_Executive_Summary_YYYY-MM-DD.pdf`
- Available in English or Arabic

### Excel Status Export
- Full data export: all 5 sheets formatted for archiving and review
- Named: `MoI_Status_Export_YYYY-MM-DD.xlsx`

### Excel Template Download
- Enhanced 5-sheet blank template with:
  - All dropdown validations (sectors, countries, tiers, statuses, etc.)
  - Pre-filled sample rows based on real investor data (GA, BlackRock, Barclays, Brookfield)
  - Frozen headers and colour-coded sheets
- Named: `MoI_CRM_Template_YYYY-MM-DD.xlsx`

---

## Excel Template Structure

The enhanced template contains 5 sheets:

| Sheet | Purpose |
|-------|---------|
| **Investor Master** | One row per investor — the spine of the system |
| **Meeting Log** | One row per meeting — replaces buried meeting tracking |
| **Opportunity Pipeline** | One row per investment opportunity |
| **Action Items** | Consolidated action items across all investors |
| **RM Tasks** | Personal task list for the RM |
| *(Reference Lists)* | Hidden sheet powering all dropdowns |

### Key fields added vs. legacy format

- Investor tier (Tier 1 Strategic / Tier 2 High Potential / Tier 3 General)
- Investment journey stage (6-stage gate: Qualification → Aftercare)
- Estimated investment value + actual commitment (SAR)
- Escalation flag (per investor, meeting, and action)
- Confidence level on opportunities
- Global IDs (INV-001, MTG-001, OPP-001, ACT-001) linking all sheets
- Controlled dropdowns replacing free-text fields

---

## How to add/manage tasks

1. Go to **RM Tasks** in the sidebar
2. Click **"Add Task"**
3. Fill in: title, linked investor, priority, due date, notes
4. Tasks appear in the **Pending Tasks** tab
5. Click ✅ next to any task to mark it complete
6. The **Dashboard** Strategic Alerts panel shows overdue and due-soon tasks

---

## Project structure

```
minster-outreach/
├── app.py                    # Main Streamlit application
├── requirements.txt          # Python dependencies
├── README.md                 # This file
├── config/
│   ├── settings.py           # MISA branding, constants, dropdown lists
│   └── translations.py       # All UI text in Arabic and English (168 keys)
├── data/
│   └── schema.py             # Expected Excel schema for validation
├── modules/
│   ├── data_loader.py        # Excel upload, legacy conversion, validation
│   ├── dashboard.py          # All dashboard panels and charts
│   ├── investors.py          # Investor master list and profile drill-down
│   ├── meetings.py           # Meeting log
│   ├── opportunities.py      # Opportunity pipeline
│   ├── actions.py            # Action items (consolidated)
│   └── tasks.py              # RM task management
├── exports/
│   ├── pptx_generator.py     # PowerPoint auto-generation
│   ├── pdf_generator.py      # PDF executive summary (reportlab)
│   └── excel_exporter.py     # Excel export + template generator
└── assets/
    └── styles.css            # MISA branding, RTL support, component styling
```

---

## Architectural decisions

**Tech stack: Python + Streamlit**
Chosen for single-user local deployment with zero infrastructure. No server, database, or
network required. Launches with one command. All state lives in the browser session.

**Excel as the data source**
The Excel file remains the master — the tool reads from it on upload and never writes back.
This preserves the existing AM workflow while adding a professional interface on top.

**Legacy format support**
The data loader auto-detects and converts the original `Action Items <Company>` per-sheet
format into the 5-sheet structure on upload, ensuring backward compatibility.

**Plotly for charts**
Selected for interactive, publication-quality charts that work natively in Streamlit.
All charts use the MISA brand colour palette.

**python-pptx + reportlab for exports**
Fully offline generation — no cloud APIs. The PPTX mirrors the Executive Account
Management Dashboard structure with per-investor slides. The PDF uses reportlab's
platypus layout engine for clean, paginated output.

**168-key translation dictionary**
All user-facing strings are in `config/translations.py`. Adding a new language requires
only adding a new key to each entry — no code changes.

---

## Branding colours

| Token | Hex | Use |
|-------|-----|-----|
| MISA Green | `#1B5C3F` | Primary — headers, buttons, charts |
| MISA Gold | `#C9974A` | Accent — Tier 1, highlights, export buttons |
| MISA Dark Green | `#0F3D2A` | Sidebar, gradient backgrounds |
| Off White | `#F7F7F2` | Page background, alternating table rows |

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| streamlit | ≥1.35 | Web UI framework |
| pandas | ≥2.0 | Data processing |
| plotly | ≥5.18 | Interactive charts |
| openpyxl | ≥3.1 | Excel read/write |
| python-pptx | ≥0.6.23 | PowerPoint generation |
| reportlab | ≥4.0 | PDF generation |
| xlrd | ≥2.0 | Legacy .xls support |
