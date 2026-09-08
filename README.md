# DocuMint

Universal document processing, automated template generation, and batch dispatch platform built in Python. Engineered for high-volume automated document generation from Excel/Word templates, local analytics tracking, and concurrent multi-threaded dispatch pipelines.

---

## Architectural Highlights

- **Dynamic Template Engine:** Merges tabular dataset inputs (Excel/XLSX) directly with formatted Microsoft Word (.docx) templates to generate customized certificates, admit cards, and invoices.
- **Parallel Dispatch Worker:** Asynchronous worker architecture utilizing parallel threads for high-throughput dispatch with custom SMTP endpoints and local Outlook integration.
- **Persistent Analytics & Audit Logging:** Built-in SQLite database (`history.db`) tracking job execution metrics, failure retries, and delivery statuses.
- **Configurable Profiles:** JSON-based profile configuration system for instant switching between client workflows.

---

## Project Structure

```text
├── examples/        # Sample spreadsheet datasets and Word template files
├── profiles/        # Saved job configurations and preset profiles (JSON)
├── scripts/         # Dataset generation and schema migration utilities
├── src/             # Core Python application logic and dispatch engine
├── history.db       # Local SQLite execution statistics database
└── run_studio.bat   # Local execution launcher
```

---

## Getting Started

### Prerequisites

- Python 3.10+
- Microsoft Word (for previewing generated `.docx` templates)

### Setup & Execution

1. **Clone the repository:**
   ```bash
   git clone https://github.com/zihaaaad/DocuMint.git
   cd DocuMint
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Launch the Studio:**
   ```bash
   # Windows batch launcher:
   run_studio.bat

   # Or run directly via Python:
   python src/main.py
   ```

4. **Access the local dashboard:**
   Navigate to `http://localhost:5000` in your web browser.

---

## License

This project is licensed under the [MIT License](LICENSE).