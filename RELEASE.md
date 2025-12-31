# 🚀 Release v2.0.0

**DocuMint v2.0.0** is a major milestone that transforms the application from a Windows-specific tool into a Universal Document Engine.

## 🌟 Highlights
*   **Universal Email**: Now supports **Gmail**, **Yahoo**, **Office365**, and any custom SMTP server. You are no longer required to have Outlook installed!
*   **Pre-Flight Validation**: A new safety engine checks your files before processing. It scans your Word template for placeholders and ensures they exist in your Excel data.
*   **Modern UI**: A completely redesigned dark-theme interface with "Midnight Blue" and "Emerald Green" accents for a professional look.
*   **Portable Mode**: Includes a standalone `DocuMint.exe` that runs without Python.

## 📋 Full Changelog

### Added
*   `src/documint/core.py`: Added `SMTPSender` class for cross-platform email support.
*   `src/documint/gui.py`: Added "Validate Files" button and logic.
*   `src/documint/gui.py`: Added `ToolTip` class for in-app help.
*   `RELEASE.md`: Added release documentation.
*   `docs/USER_GUIDE.md`: Comprehensive user manual.

### Changed
*   **Colors**: Updated branding to Professional Navy/Green palette.
*   **Structure**: Refactored entire codebase into `src/documint` package.
*   **Build**: Updated `DocuMint.spec` for reliable .exe generation.

### Fixed
*   Fixed crashing issue when `config.json` was corrupted (added try-except block).
*   Fixed `pywin32` dependency issues on non-Windows systems (by making imports conditional).

## 📥 Download
*   **Source Code**: [v2.0.0.zip](https://github.com/Z-root-X/DocuMint/archive/refs/tags/v2.0.0.zip)
*   **Portable EXE**: [DocuMint_Portable.zip](https://github.com/Z-root-X/DocuMint/raw/master/DocuMint_Portable.zip)
