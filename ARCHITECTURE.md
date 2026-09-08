# DocuMint System Architecture Specification

## Overview
DocuMint is structured into decoupled, single-responsibility Python modules designed for high throughput, memory safety, and cross-platform document compilation.

```text
┌────────────────────────────────────────────────────────┐
│                   Input Ingestion Layer                │
│    Spreadsheets (.xlsx / .csv)  +  Templates (.docx)   │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│             Pre-Flight Validation Engine               │
│         Schema verification & Missing Column Alert     │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│            Run-Preserving OpenXML Parser               │
│       Granular replacement preserving fonts & styles   │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│          Universal PDF Conversion Pipeline             │
│    WinWord COM  ─►  docx2pdf  ─►  LibreOffice Headless │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│           Concurrent Asynchronous Dispatcher           │
│     SMTP SSL/TLS Pool  or  Local Outlook COM Driver    │
└───────────────────────────┬────────────────────────────┘
                            │
                            ▼
┌────────────────────────────────────────────────────────┐
│           ACID SQLite Audit & Telemetry Ledger         │
│          Full record of delivery hashes & retries      │
└────────────────────────────────────────────────────────┘
```

---

## Architectural Subsystems

### 1. Document Parsing Subsystem (`src/documint/core.py`)
- **OpenXML Tree Inspection:** Traverses `document.paragraphs`, `document.tables`, and nested `cell.paragraphs`.
- **Run-Level Token Replacement:** Substitutes variables within individual `run` text nodes to maintain font family, point size, bold, italic, underline, and color attributes.
- **Descending Length Sorting:** Replacement maps are sorted by key length in descending order, guaranteeing that bracketed tokens (e.g., `{{Name}}` or `<Name>`) are parsed before bare variable names.

### 2. Universal Document Converter (`UniversalDocumentConverter`)
- **Design Pattern:** Strategy Pattern utilizing an abstract base class (`DocumentConverter`).
- **Windows Subsystem:** Dispatches headless `Word.Application` COM instance with `Visible = False`, opens the temporary file, executes `SaveAs(..., FileFormat=17)`, and guarantees process termination in a `finally` block.
- **POSIX Subsystem:** Spawns a subprocess to invoke `libreoffice --headless --convert-to pdf` with directory redirection.

### 3. Asynchronous Dispatch Engine (`EmailSender`)
- **Connection Management:** Abstracted interface with concrete implementations:
  - `SmtpEmailSender`: Uses standard library `smtplib` with MIME multipart message packing.
  - `WinOutlookSender`: Invokes `win32com.client.Dispatch("Outlook.Application")`.
- **MIME Formatting:** Attachments are dynamically detected with `mimetypes.guess_type` and base64-encoded.

### 4. Local Web Studio (`src/web/`)
- Built with lightweight **Flask** routing serving a clean dashboard for non-technical administrators.
- Supports profile loading/saving from `profiles/*.json`.
- Queries `history.db` for real-time delivery graphs and audit logs.
