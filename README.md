# Failure Tracking System
### Manufacturing QC — Pareto Analysis Dashboard

A production-ready web app built with **Python + Streamlit + OpenPyXL**.
All failure data is stored in a real Microsoft Excel file (`failure_log.xlsx`).

---

## Quick Start

### 1. Install Python (if not already installed)
Download Python 3.10+ from https://www.python.org/downloads/

### 2. Install dependencies
Open a terminal in this folder and run:
```bash
pip install -r requirements.txt
```

### 3. Run the app
```bash
streamlit run app.py
```

The app will open automatically in your browser at `http://localhost:8501`

---

## Features

### 📋 Data Entry
- Form with all required fields: Date, Model, Serial Number, Process, Failure Type, Description, Remark, Photo
- Photo upload — saved to `photos/` folder, file path stored in Excel
- Duplicate serial number detection with visual highlighting in Excel
- Validation for all required fields

### 📊 Pareto Analysis
- Date range filter (Start Date → End Date)
- Auto-generated Pareto chart (bar + cumulative line)
- Summary metrics: total failures, types, top failure, 80/20 rule count
- **Export Pareto Report to Excel** — styled report with embedded charts

### 📁 Failure Log
- Search by Serial Number
- Filter by Failure Type or Process
- Duplicate unit indicator (⚠️ ×N for repeated SNs, ✅ New for first-time)
- Delete individual records
- **Download Full Excel Log** from sidebar at any time

---

## File Structure
```
failure_tracker/
├── app.py                  # Main Streamlit application
├── requirements.txt        # Python dependencies
├── README.md               # This file
├── failure_log.xlsx        # Auto-created on first run (Excel data store)
└── photos/                 # Uploaded photos saved here
```

## Excel File Format
The `failure_log.xlsx` file is fully editable in Microsoft Excel.

**Sheet: Failure_Log**

| Date | Model | Serial Number | Process | Failure Type | Description | Remark | Photo Path |
|------|-------|---------------|---------|--------------|-------------|--------|------------|

- Rows alternate shading for readability
- Duplicate serial numbers highlighted in yellow/orange
- Header row frozen for easy scrolling

---

## Deployment (optional)

To share with your team, you can deploy to **Streamlit Community Cloud** (free):

1. Push this folder to a GitHub repository
2. Go to https://share.streamlit.io
3. Connect your GitHub repo and select `app.py`
4. Done — accessible via a public URL

> **Note:** For persistent Excel storage on the cloud, replace `failure_log.xlsx`
> with a cloud storage backend (Google Sheets, Azure Blob, AWS S3, etc.)

---

## Requirements
- Python 3.10+
- streamlit >= 1.35
- pandas >= 2.0
- openpyxl >= 3.1
- matplotlib >= 3.8
- pillow >= 10.0
