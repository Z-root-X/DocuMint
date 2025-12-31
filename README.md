# 🍃 DocuMint

> **The Universal Document Automation Engine.**  
> *Batch Generate. Convert. Distribute.*

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Platform](https://img.shields.io/badge/platform-windows%20%7C%20linux-lightgrey.svg)]()
[![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)]()

---

## 🌟 Overview

**DocuMint** is an industrial-grade automation tool designed to eliminate manual document workflows. Whether you are a university issuing admit cards, a company sending generated contracts, or an event organizer distributing tickets, DocuMint orchestrates the entire process.

It seamlessly integrates **Excel Datasources**, **Word Templates**, and **Universal Email Protocols** (SMTP/Outlook) into a single, reliable pipeline.

---

## ✨ Key Features

### 🚄 **Core Engine**
*   **Dynamic Placeholders**: Inject data into your documents using simple tags like `<Name>` or `<Department>`.
*   **Universal Email**: Native support for **Gmail**, **Yahoo**, **Office365** (SMTP), and legacy **Outlook Desktop** automation.
*   **Smart Conversion**: High-fidelity `.docx` to `.pdf` conversion pipeline.

### 🛡️ **Reliability & Safety**
*   **Pre-Flight Validation**: The `Validate Files` engine scans your templates against your data before execution, preventing batch failures.
*   **Resilience**: Built-in retry mechanisms and delay throttling to respect API rate limits.
*   **Detailed Logging**: Comprehensive audit trails generated in Excel format.

### 🎨 **User Experience**
*   **Modern UI**: A responsive, dark-themed interface designed for professionals.
*   **Profiles**: Intelligent configuration persistence—pick up exactly where you left off.
*   **Standalone**: Deploy as a single portable `.exe` file. No Python installation required.

---

## � Quick Start

### 1. Installation

**Option A: Standalone Executable (Windows)**
Download the latest release, unzip, and run `DocuMint.exe`.

**Option B: Python Source**
```bash
git clone https://github.com/Z-root-X/DocuMint.git
cd DocuMint
pip install .
```

### 2. Your First Job
1.  **Prepare Data**: Create an Excel file (`data.xlsx`) with columns like `Name`, `Email`, `ID`.
2.  **Prepare Template**: Create a Word doc (`template.docx`) and use tags like `<Name>` where you want data to appear.
3.  **Run DocuMint**:
    ```bash
    python src/main.py
    ```
4.  **Configure**: Select your files.
5.  **Validate**: Click **"Validate Files"** to ensure all tags match your Excel headers.
6.  **Launch**: Go to the Run tab and click **Start Process**.

---

## 📚 Documentation

Detailed guides are available in the [docs](docs/) directory:

*   [� User Guide](docs/USER_GUIDE.md): Deep dive into advanced configuration, SMTP setup, and template design.
*   [💻 Developer Guide](CONTRIBUTING.md): How to build, test, and contribute.

---

## ⚙️ Configuration Hints

| Setting | Description |
| :--- | :--- |
| **PDF Format** | Define output filenames like `Card_{<ID>}_{<Name>}`. |
| **SMTP Host** | e.g. `smtp.gmail.com` for Gmail (Port 465 or 587). |
| **Delay** | Seconds to wait between emails. Recommended `2` seconds to avoid spam filters. |

---

## 🤝 Contributing

We welcome contributions! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for details.

## �‍💻 Created By

**Zihad Hasan**  
🚀 *Full Stack Developer | Python Expert*  
🌐 Portfolio: [zihadhasan.web.app](https://zihadhasan.web.app)

## �📄 License

Copyright © 2025. Released under the [MIT License](LICENSE).