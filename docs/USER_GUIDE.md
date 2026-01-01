# 📘 DocuMint Expert Manual (v3.0)
*The Enterprise-Grade Document Automation Platform*

**Version**: 3.0.0 (International Standard)
**Last Updated**: January 1, 2026

---

## 🏗️ 1. Architecture & Process (Deep Dive)
DocuMint is built on a **Pipeline Architecture**. Unlike simple scripts, it processes data in isolated stages to ensure data integrity and speed.

### The Pipeline Flow
```mermaid
graph TD
    A[Excel Data Source] -->|1. Validation| B{Integrity Check}
    B -->|Fail| C[Stop & Alert User]
    B -->|Pass| D[2. Generation Engine]
    D -->|Template Injection| E[Word Doc (.docx)]
    E -->|Conversion| F[PDF Document (.pdf)]
    F -->|3. Dispatcher| G{Gateway Selection}
    G -->|Custom SMTP| H[Parallel Sending (5x Speed)]
    G -->|Outlook App| I[Serial Sending (Safe Mode)]
    H --> J[Analytics DB (SQLite)]
    I --> J
```

### Key Components
1.  **Validator**: Scans every row in your Excel file against the Word Template placeholders *before* starting. This prevents crashing at row 499 of 500.
2.  **ThreadPoolExecutor (New in v3.0)**: When using SMTP, DocuMint spins up 5 concurrent threads. This means it sends 5 emails simultaneously, drastically reducing wait times for large batches.
3.  **SQLite Recorder (New in v3.0)**: Every job is logged to a local database (`history.db`). This allows the "Analytics" tab to show you historical success rates.

---

## 🚀 2. Quick Start Guide

### Step 1: Launch
Double-click `run_studio.bat`. The Web Studio will open at `http://localhost:5000`.

### Step 2: Prepare Assets (The "Examples")
Look in the `examples/` folder.
*   **Data**: `student_data.xlsx` (Edit this with your list).
*   **Template**: `admit_card_template.docx` (Edit this with your design).

### Step 3: Configure Job
1.  **Profiles (New!)**: If you have saved settings before, pick them from the "Load Profile" dropdown.
2.  **File Paths**: Paste the full paths to your Excel and Word files.
3.  **PDF Output**: Choose where to save the generated files.

### Step 4: Run
Click **Start Engine**. Watch the "System Logs" or switch to the "Analytics" tab to watch the counter go up.

---

## 💾 3. Configuration Profiles
Stop typing the same paths every day.
1.  Set up your job (Paths, Subject, Body, SMTP settings).
2.  Click **"Save Profile"**.
3.  Name it (e.g., `Monthly_Invoices`).
4.  Next time, just select `Monthly_Invoices` from the dropdown.

---

## 📊 4. Analytics Dashboard
Click the **"Analytics"** button in the bottom panel.
*   **Total Jobs**: How many batches you have run.
*   **Success Rate**: Tracks if emails were delivered or bounced.
*   **History Table**: Shows the timestamp and size of previous jobs.

---

## 🔧 5. Troubleshooting
| Issue | Solution |
| :--- | :--- |
| **"SMTP Auth Failed"** | You typically need an **App Password** (not your login password) for Gmail/Yahoo. |
| **"OutlookRPC Error"** | Ensure the "New Outlook" (Web wrapper) is OFF. DocuMint needs Classic Outlook. |
| **"Placeholder Missing"** | Run "Check Integrity". It will tell you exactly which Excel header is missing. |

---
**Privacy Note**: DocuMint runs 100% offline. No data leaves your machine except via your own Email Gateway.
