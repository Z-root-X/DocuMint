# User Guide for DocuMint
*Automated Document Generation System*

## 1. Quick Start
DocuMint has two interfaces:
1.  **Web Studio (Recommended)**: Modern, browser-based interface.
2.  **Desktop App**: Classic Windows GUI.

---

## 2. Web Studio Guide
### Step 1: Launch
1.  Open your terminal in the project folder.
2.  Run: `python src/web/app.py`
3.  Go to: `http://localhost:5000`

### Step 2: Configure Job
*   **File Resources**: Paste the full absolute path to your `.xlsx` data file and `.docx` template.
    *   *Tip: Use the "Check Integrity" button to verify files exist.*
*   **Email Composer**: Write your email subject and body.
    *   *Feature: Click "Preview" to see how the HTML body renders.*
*   **Target Folders**: Specify where to save the generated PDFs.
*   **Gateway**: Choose **SMTP** (Universal) or **Outlook**.
    *   *For Gmail: Use App Password, not your login password.*

### Step 3: Run
*   Click **Start Engine**.
*   Monitor the **System Logs** panel at the bottom for real-time status.

---

## 3. Data Preparation
### Excel File (.xlsx)
*   **Row 1 MUST be headers.**
*   **Required Header**: `Email` (case-insensitive) is required if sending emails.
*   **Other Headers**: Columns like `Name`, `ID`, `Dept` become variables.

### Word Template (.docx)
*   Use placeholders formatted as `<HeaderName>`.
*   Example: `Dear <Name>, your ID is <ID>.`
*   The system matches `<Name>` in Word to the `Name` column in Excel.

---

## 4. Email Configuration
### Gmail Setup
1.  Go to Google Account > Security.
2.  Enable 2-Factor Authentication.
3.  Search for "App Passwords".
4.  Generate a new App Password (select "Mail" and "Windows Computer").
5.  Use this 16-character code as the **App Password** in DocuMint.

### Outlook Setup
*   Ensure the "New Outlook" toggle is OFF (Classic Outlook application must be running).
*   DocuMint will use your default Outlook profile to send emails.

---
*Created by [Zihad Hasan](https://zihadhasan.web.app)*
