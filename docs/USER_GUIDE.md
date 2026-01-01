# 📘 DocuMint Expert Manual
*Automated Document Generation & Emailing System*

**Version:** 2.1.0 (`Web Studio Edition`)
**Last Updated:** January 1, 2026
**Author:** Zihad Hasan

> [!IMPORTANT]
> **Privacy First**: DocuMint runs entirely on your local machine. No data (Excel/Word) is ever uploaded to the cloud.

---

## 📋 Table of Contents
1.  [Concept Overview](#concept)
2.  [Quick Start (Web Studio)](#quick-start)
3.  [Designing Your Templates](#templates)
4.  [Email Gateway Setup](#gateway)
5.  [Advanced Configuration](#advanced)
6.  [Troubleshooting Guide](#troubleshooting)

---

## <a id="concept"></a>1. Concept Overview
DocuMint automates the "Mail Merge" process. It takes:
1.  **Data Source (Excel)**: A list of people (Name, ID, Email).
2.  **Template (Word)**: A document with placeholders (e.g., `<Name>`).

**The Output**:
*   A personalized PDF for every row in Excel.
*   An email sent to that person with the PDF attached (optional).

---

## <a id="quick-start"></a>2. Quick Start (Web Studio)
The Web Studio is the modern interface for DocuMint.

### Step 1: Run the Server
Open your terminal in the project folder and run:
```bash
python src/web/app.py
```
Open **[http://localhost:5000](http://localhost:5000)** in Chrome/Edge.

### Step 2: Get Templates (Optional)
If you are new, go to the **Documentation** page (Sidebar > Documentation) and click **"Download Template"** to get starter files.

### Step 3: Configure the Job
1.  **Excel Path**: Full path to your `.xlsx` file (e.g., `D:\MyFiles\data.xlsx`).
2.  **Word Path**: Full path to your `.docx` file.
    *   *Tip: Click "Check Integrity" to verify your files match.*
3.  **PDF Output**: Folder where PDFs will be saved (automatically created if missing).

### Step 4: Compose & Send
1.  **Email Composer**: Write your Subject and Body.
    *   *Tip: Use the "Preview" button to see your HTML email live.*
2.  **Gateway**: Choose "Custom SMTP" (Gmail/Yahoo) or "Outlook App".
3.  **Launch**: Click "Start Engine".

---

## <a id="templates"></a>3. Designing Your Templates

### Excel Data Rules
*   **Row 1 is Reserved**: The first row MUST contain headers (e.g., `Name`, `ID`, `Designation`).
*   **Case Insensitive**: `Name` and `name` are treated the same.
*   **Required Column**: You MUST have a column named `Email` if you want to send emails.
*   **No Merged Cells**: Ensure your data is a simple flat table.

### Word Template Rules
*   **Placeholders**: Use angle brackets: `<HeaderName>`.
*   **Styling**: Apply Bold/Color/Fonts directly to the `< >` text in Word. The replaced text will inherit that style.
*   **Images/Tables**: Fully supported. The script preserves all layout.

---

## <a id="gateway"></a>4. Email Gateway Setup

### Option A: Gmail (Universal SMTP) 🔥 *Recommended*
1.  **Host**: `smtp.gmail.com`
2.  **Port**: 465 (Auto-handled).
3.  **Password**: You strictly need an **App Password**.
    *   Go to Google Account > Security > 2-Step Verification > App Passwords.
    *   Generate one for "Mail". Copy the 16-digit code.
    *   Paste it into DocuMint. **Do not use your login password.**

### Option B: Outlook Desktop App
1.  **Requirement**: "Classic" Outlook must be installed and running.
2.  **Auth**: No password needed. It uses your active desktop session.
3.  **Limitation**: Slow for bulk sending (Outlook has security delays).

---

## <a id="advanced"></a>5. Advanced Configuration
When sending bulk emails (e.g., 500+), server reputation matters.

*   **Delay**: The default delay is **2 seconds** between emails. Increase this to 5-10s for large lists to avoid spam filters.
*   **Retries**: If an email fails (e.g., internet blip), the system retries **2 times** automatically.

---

## <a id="troubleshooting"></a>6. Troubleshooting Guide

| Issue | Solution |
| :--- | :--- |
| **"Column Not Found"** | Check your Excel header. ensure no trailing spaces (e.g. "Name "). |
| **"SMTP Auth Failed"** | You are likely using your login password. Use a Google App Password. |
| **"Outlook RPC Error"** | Ensure the "New Outlook" toggle is OFF. DocuMint needs Classic COM automation. |
| **Bangla Font Issues** | Web preview uses standard web fonts. The final PDF uses your system's `Hind Siliguri`. Ensure it is installed in Windows. |

---
*Created by [Zihad Hasan](https://zihadhasan.web.app)*
