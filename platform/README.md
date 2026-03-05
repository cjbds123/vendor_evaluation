# CPMS Evaluation Platform

A **local web-based platform** for configuring, running, scoring, and reporting CPMS vendor evaluations.

## Quick Start

### Option 1: Double-click
1. Double-click **`run.bat`** in the `platform` folder
2. Open **http://127.0.0.1:5000** in your browser
3. Click **"Quick Import CPMS Suite"** on the dashboard

### Option 2: Command line
```bash
cd platform
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## Features

### ✅ Test Suite Management
- Auto-import from the CPMS Excel file (Functional + Non-Functional tests)
- Browse, filter, and search all 200+ test cases
- Edit test cases (capability, scenario, criteria, priority, weight)
- Full category management (create, edit, delete, hierarchy)

### 📊 Evaluation Workflow
- Create evaluation projects with multiple vendors
- Score each vendor against every test case using configurable scoring scale
- Inline scoring with auto-calculation of weighted scores
- Bulk update support (set scoring/status for multiple tests at once)
- Test execution status tracking (Not Started → In Progress → Blocked → Submitted → Approved)

### 📎 Evidence Management
- Upload files (PDF, images, documents, ZIP)
- Add external links (URLs)
- Add text notes/transcripts
- Evidence status tracking (pending/accepted/rejected)

### 📈 Reporting & Comparison
- **Scorecard**: Overall weighted score per vendor, category breakdown, gating failures
- **Vendor Comparison**: Side-by-side category comparison with radar chart
- **Export**: Download vendor results as Excel
- **Charts**: Bar charts and radar charts for visual comparison

### ⚙️ Configuration
- **Scoring Scale**: Fully configurable (OOB=5, Configurable=4, Custom=2, Roadmap=1, Not Supported=0, N/A)
- **Categories**: Full CRUD – add, edit, rename, delete, set weight multipliers
- **Audit Log**: Full trail of all changes (scores, edits, approvals)

## Tech Stack
- **Backend**: Python Flask + SQLAlchemy
- **Database**: SQLite (zero configuration)
- **Frontend**: Bootstrap 5 + Chart.js
- **Import/Export**: openpyxl (Excel)

## Data
All data is stored locally in `platform/instance/cpms.db` (SQLite).
Evidence files are stored in `platform/uploads/`.

## Resetting
To start fresh, delete:
- `platform/instance/cpms.db` (database)
- `platform/uploads/` (evidence files)

Then restart the app and re-import the Excel file.
