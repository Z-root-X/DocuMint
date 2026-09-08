# DocuMint Engineering & User Guide

A comprehensive, production-grade guide to high-volume document generation, vector PDF compilation, and multi-channel dispatch with DocuMint.

---

## Table of Contents

1. [Real-World Problem Analysis](#1-real-world-problem-analysis)
   - [1.1 Academic Institutions & Universities](#11-academic-institutions--universities)
   - [1.2 Corporate HR & Payroll Operations](#12-corporate-hr--payroll-operations)
   - [1.3 Healthcare & Diagnostic Centers](#13-healthcare--diagnostic-centers)
   - [1.4 Conferences, Summits & Training Certifications](#14-conferences-summits--training-certifications)
2. [Why Legacy and Cloud Approaches Fail](#2-why-legacy-and-cloud-approaches-fail)
3. [Template Design Specifications](#3-template-design-specifications)
   - [3.1 Supported Placeholder Syntaxes](#31-supported-placeholder-syntaxes)
   - [3.2 Formatting Preservation Rules](#32-formatting-preservation-rules)
   - [3.3 Working with Tables and Structured Grids](#33-working-with-tables-and-structured-grids)
4. [Dataset Standards & Validation](#4-dataset-standards--validation)
   - [4.1 Excel and CSV Formatting Requirements](#41-excel-and-csv-formatting-requirements)
   - [4.2 Pre-Flight Validation Engine](#42-pre-flight-validation-engine)
5. [Universal PDF Conversion Pipeline](#5-universal-pdf-conversion-pipeline)
   - [5.1 Multi-Tier Conversion Hierarchy](#51-multi-tier-conversion-hierarchy)
   - [5.2 Headless Linux & Docker Deployment](#52-headless-linux--docker-deployment)
6. [Multi-Channel Dispatch Engine](#6-multi-channel-dispatch-engine)
   - [6.1 Direct SMTP Relay (Gmail, Microsoft 365, Amazon SES)](#61-direct-smtp-relay-gmail-microsoft-365-amazon-ses)
   - [6.2 Native Windows Outlook Automation](#62-native-windows-outlook-automation)
   - [6.3 Personalized Attachment Binding](#63-personalized-attachment-binding)
7. [Persistent Audit & Telemetry](#7-persistent-audit--telemetry)
8. [CLI and Python API Usage](#8-cli-and-python-api-usage)

---

## 1. Real-World Problem Analysis

High-volume document generation is a fundamental operational requirement across numerous industries. However, organizations frequently encounter significant technical friction, compliance hurdles, and manual overhead.

### 1.1 Academic Institutions & Universities
- **The Challenge:** Generating thousands of course certificates, examination admit cards, grade transcripts, and provisional credentials following semester completions or graduation ceremonies.
- **The Friction:** Manual mail merges in Word require days of human supervision. Mistakes such as swapped student registration numbers or mismatched grades cannot be caught automatically before generation.
- **DocuMint Solution:** Performs automated pre-flight column schema validation against the student registrar database, produces individualized vector PDFs with tamper-resistant hashes, and dispatches them directly to student inboxes in parallel threads.

### 1.2 Corporate HR & Payroll Operations
- **The Challenge:** Generating monthly salary slips, employment verification letters, bonus confirmations, and annual tax certificates for distributed workforces.
- **The Friction:** Uploading employee salary details and national identification numbers to third-party cloud SaaS document generators introduces serious privacy vulnerabilities and violates GDPR/FERPA regulations.
- **DocuMint Solution:** 100% local execution on the internal corporate network. Zero cloud leakage ensures total data sovereignty.

### 1.3 Healthcare & Diagnostic Centers
- **The Challenge:** Compiling and delivering personalized lab reports, health checkup summaries, and appointment reminders.
- **The Friction:** Data security regulations (such as HIPAA) strictly prohibit exposing patient diagnostic data to external unvetted APIs.
- **DocuMint Solution:** Local generation of patient-specific PDFs with immediate local email relay or desktop Outlook integration.

### 1.4 Conferences, Summits & Training Certifications
- **The Challenge:** Event organizers needing to generate branded attendance certificates and delegate badges with custom attendee names, track titles, and dates.
- **The Friction:** SaaS tools charge expensive per-document subscription fees ($0.20 to $1.50 per generated certificate), scaling poorly for thousands of attendees.
- **DocuMint Solution:** Open-source (MIT licensed) platform with zero recurring costs and unrestricted throughput.

---

## 2. Why Legacy and Cloud Approaches Fail

| Vector | Legacy MS Word Mail Merge | Cloud SaaS APIs (PandaDoc / Zapier) | DocuMint Local Engine |
| :--- | :--- | :--- | :--- |
| **Privacy & Security** | Local, but manual GUI only | High risk: Transmits data to third-party cloud | 100% Local-first, zero cloud exposure |
| **Batch PDF Generation** | Manual "Print to PDF" per document | Cloud-rendered (cost per generation) | Multi-tier automated vector PDF compiler |
| **Formatting Integrity** | Frequently strips run formatting | Requires proprietary web template builders | Run-preserving OpenXML parser |
| **Dispatch Automation** | Single-threaded Outlook only | API rate limits and paid tiers | Concurrent asynchronous worker pool |
| **Pre-Flight Validation** | None (crashes mid-job on errors) | Basic schema checks | Pre-flight missing column validator |
| **Audit Ledger** | None | Hosted cloud logs | Local ACID SQLite database (`history.db`) |

---

## 3. Template Design Specifications

### 3.1 Supported Placeholder Syntaxes
DocuMint accommodates diverse workflow conventions by supporting three standard placeholder patterns in Microsoft Word (`.docx`) files:

1. **Angle Bracket Syntax (Default):**
   ```text
   <Recipient_Name>
   <Course_Title>
   <Registration_ID>
   <Issue_Date>
   ```
2. **Jinja / Mustache Syntax:**
   ```text
   {{Recipient_Name}}
   {{Course_Title}}
   {{Registration_ID}}
   ```
3. **Square Bracket Syntax:**
   ```text
   [Recipient_Name]
   [Course_Title]
   [Registration_ID]
   ```

### 3.2 Formatting Preservation Rules
In the OpenXML standard (`.docx`), paragraphs are divided into `runs` (individual contiguous text chunks sharing identical formatting). Traditional Python string replacement scripts replace text at the paragraph level (`paragraph.text = new_text`), which strips character-level styles such as bold, italics, font size, and RGB text color.

DocuMint resolves this by performing a two-pass replacement:
1. **Pass 1 (Run-Level):** Inspects individual `run.text` objects. If the placeholder is contained within a single run, the text is substituted directly within that run, completely preserving all styling properties.
2. **Pass 2 (Boundary Fallback):** If Word split the placeholder across multiple runs due to spell-check or cursor edits, DocuMint re-stitches the paragraph cleanly to guarantee complete text replacement.

### 3.3 Working with Tables and Structured Grids
Placeholders positioned inside Word table cells, headers, footers, and callout boxes are recursively discovered and merged with identical fidelity to body paragraphs.

---

## 4. Dataset Standards & Validation

### 4.1 Excel and CSV Formatting Requirements
- **First Row Headers:** The first row of your spreadsheet (`.xlsx`, `.xls`, or `.csv`) must contain the column header names matching your template placeholders.
- **Column Normalization:** Column names are automatically trimmed of leading and trailing whitespace.
- **Data Types:** Dates and numeric values are stringified accurately without exponential notation bugs.

### 4.2 Pre-Flight Validation Engine
Before beginning batch processing, invoke `validate_placeholders(data_path, template_path)` to ensure every template placeholder corresponds to an existing column header in the dataset.

```python
from documint.core import validate_placeholders

is_valid, missing_columns = validate_placeholders("students.xlsx", "certificate_template.docx")

if not is_valid:
    print(f"Validation Error: Missing columns in spreadsheet: {missing_columns}")
else:
    print("Pre-flight check passed. All template placeholders exist in dataset.")
```

---

## 5. Universal PDF Conversion Pipeline

### 5.1 Multi-Tier Conversion Hierarchy
The `UniversalDocumentConverter` class implements a robust fallback hierarchy to ensure document conversion works seamlessly across any operating environment:

```text
[Input .docx]
      │
      ├──> Tier 1: Windows Word COM (win32com.client) ──> High-Fidelity Vector PDF
      │         (Automatic on Windows with Office installed)
      │
      ├──> Tier 2: Pure-Python docx2pdf
      │         (Fallback if COM automation is busy)
      │
      └──> Tier 3: Headless LibreOffice CLI (libreoffice / soffice)
                (Default on Linux, macOS, and Docker containers)
```

### 5.2 Headless Linux & Docker Deployment
To run DocuMint on Linux servers or Docker containers, install LibreOffice:

```bash
# Ubuntu / Debian
sudo apt-get update && sudo apt-get install -y libreoffice libreoffice-writer

# Red Hat / Rocky Linux / Fedora
sudo dnf install -y libreoffice-writer

# Test headless conversion
libreoffice --headless --convert-to pdf template.docx
```

---

## 6. Multi-Channel Dispatch Engine

### 6.1 Direct SMTP Relay
Configure connection parameters in your profile JSON or environment variables:

- **Gmail / Google Workspace:**
  - Host: `smtp.gmail.com`
  - Port: `587` (STARTTLS) or `465` (SSL)
  - Authentication: Use an App Password if 2-Step Verification is enabled.
- **Microsoft 365 / Outlook:**
  - Host: `smtp.office365.com`
  - Port: `587` (STARTTLS)

### 6.2 Native Windows Outlook Automation
When operating on Windows corporate workstations, DocuMint can dispatch emails directly through the active Microsoft Outlook desktop profile without storing raw SMTP credentials in code.

### 6.3 Personalized Attachment Binding
The dispatch worker automatically attaches each recipient's uniquely generated PDF to their email message while injecting their personalized variables into the email's HTML body.

---

## 7. Persistent Audit & Telemetry

Every batch job logs its activity to `history.db` (SQLite) with the following schema:

```sql
CREATE TABLE IF NOT EXISTS job_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
    recipient_email TEXT NOT NULL,
    recipient_name TEXT,
    document_path TEXT,
    status TEXT NOT NULL, -- 'SUCCESS', 'FAILED', 'PENDING'
    error_message TEXT,
    dispatch_method TEXT
);
```

---

## 8. CLI and Python API Usage

### Python API Example
```python
from documint.core import (
    replace_placeholders_in_doc,
    UniversalDocumentConverter,
    SmtpEmailSender
)
from docx import Document

# 1. Load template and replace variables
doc = Document("template.docx")
replace_placeholders_in_doc(doc, {
    "Student_Name": "Zihad Hasan",
    "Course": "Deep Learning Systems",
    "Grade": "Distinction"
})
doc.save("output_temp.docx")

# 2. Compile to PDF
converter = UniversalDocumentConverter()
converter.convert_to_pdf("output_temp.docx", "Zihad_Hasan_Certificate.pdf")

# 3. Dispatch via SMTP
sender = SmtpEmailSender(
    host="smtp.gmail.com",
    port=587,
    username="notifications@institution.edu",
    password="app-specific-password",
    use_tls=True
)
sender.send_email(
    to_email="student@institution.edu",
    subject="Your Academic Certificate",
    body_html="<p>Dear Zihad,<br>Please find your certificate attached.</p>",
    attachments=["Zihad_Hasan_Certificate.pdf"]
)
```

### Launching the Web Studio
```bash
# Launch local dashboard
python src/main.py

# Open your browser at http://localhost:5000
```

---

## License & Attribution

Authored by **Zihad Hasan** ([https://github.com/zihaaaad](https://github.com/zihaaaad)).  
Distributed under the **MIT License**.
