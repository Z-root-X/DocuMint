# DocuMint

Universal document processing, automated template generation, and batch dispatch platform built in Python. Engineered for high-volume automated document generation from Excel/Word templates, local analytics tracking, and concurrent multi-threaded dispatch pipelines.

**Live Interactive Showcase & Documentation:** [https://zihaaaad.github.io/DocuMint/](https://zihaaaad.github.io/DocuMint/)

---

## Real-World Problem Statement

Bulk document generation and personalized email delivery in enterprise and educational organizations typically face severe limitations:
1. **Formatting Corruption:** Standard Microsoft Word mail merge frequently alters typeface hierarchies, destroys table alignments, and requires manual clicks for every batch.
2. **Privacy Violations & Cloud Data Exposure:** Cloud SaaS tools (PandaDoc, DocuSign, Zapier plugins) require transmitting confidential student grades, employee salaries, and identity numbers to third-party multi-tenant servers.
3. **Dispatch Disconnect:** Traditional tools require generating documents in one software, converting them to PDF in a second tool, and manually attaching each PDF to separate emails in a third tool.

**DocuMint** resolves these bottlenecks by providing an automated, 100% local-first Python pipeline with zero cloud leakage, multi-tier cross-platform PDF compiling, run-preserving OpenXML placeholder replacement, and asynchronous batch dispatch.

---

## Architectural Highlights

- **Dynamic Template Engine:** Merges tabular dataset inputs (Excel/CSV) directly with formatted Microsoft Word (`.docx`) templates while strictly preserving font styling and table structures.
- **Multi-Syntax Placeholder Support:** Accepts `<Variable>`, `{{Variable}}`, and `[Variable]` notation.
- **Universal PDF Converter:** Multi-tier conversion hierarchy dynamically leveraging Microsoft Word COM automation, `docx2pdf`, or headless `LibreOffice` on Linux/Docker.
- **Parallel Dispatch Worker:** Asynchronous worker architecture utilizing parallel threads for high-throughput dispatch with custom SMTP endpoints or local Outlook integration.
- **Persistent Analytics & Audit Logging:** Built-in SQLite database (`history.db`) tracking job execution metrics, failure retries, and delivery statuses.
- **Pre-Flight Validation Engine:** Automatically checks template variables against spreadsheet headers before execution to prevent runtime failures.

---

## Documentation & Guides

- [Engineering & User Guide (GUIDE.md)](GUIDE.md) - Deep dive into real-world use cases, template authoring, SMTP configuration, and headless Linux deployment.
- [System Architecture (ARCHITECTURE.md)](ARCHITECTURE.md) - Detailed breakdown of internal subsystems, OpenXML parsing mechanics, and failure resilience.
- [Wikipedia-Style Web Knowledgebase](https://zihaaaad.github.io/DocuMint/) - Live interactive document simulator and complete technical specification.

---

## Project Structure

```text
├── docs/            # Interactive GitHub Pages Wikipedia-style documentation
├── examples/        # Sample spreadsheet datasets and Word template files
├── profiles/        # Saved job configurations and preset profiles (JSON)
├── scripts/         # Dataset generation and schema migration utilities
├── src/             # Core Python application logic and dispatch engine
│   ├── documint/    # Core document parsing, conversion, and dispatch modules
│   └── web/         # Local Flask Web Studio interface and API routes
├── tests/           # Automated pytest unit test suite
├── history.db       # Local SQLite execution statistics database
├── GUIDE.md         # Comprehensive engineering and usage guide
├── ARCHITECTURE.md  # System architecture specification
└── run_studio.bat   # Local execution launcher
```

---

## Getting Started

### Prerequisites

- Python 3.10+
- Microsoft Word (Windows) or LibreOffice (Linux/macOS) for PDF conversion

### Setup & Execution

1. **Clone the repository:**
   ```bash
   git clone https://github.com/zihaaaad/DocuMint.git
   cd DocuMint
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   # Or using uv:
   uv sync
   ```

3. **Run Unit Tests:**
   ```bash
   pytest tests/
   ```

4. **Launch the Studio:**
   ```bash
   # Windows batch launcher:
   run_studio.bat

   # Or run directly via Python:
   python src/main.py
   ```

5. **Access the local dashboard:**
   Navigate to `http://localhost:5000` in your web browser.

---

## Authors & Maintenance

Authored and maintained by **Zihad Hasan** ([https://github.com/zihaaaad](https://github.com/zihaaaad)).  
Developer Platform: [https://zihadhasan.web.app/](https://zihadhasan.web.app/)

---

## License

This project is licensed under the [MIT License](LICENSE).
