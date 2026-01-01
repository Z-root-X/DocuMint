# 📘 DocuMint v2.0 Professional User Guide

**Version:** 2.0.0 (Clean Enterprise Edition)
**Last Updated:** January 1, 2026
**Author:** Zihad Hasan

---

## 📋 Table of Contents
1.  [Introduction](#1-introduction)
2.  [Installation & Setup](#2-installation--setup)
3.  [The Web Studio (Recommended)](#3-the-web-studio-recommended)
    *   [Interface Overview](#31-interface-overview)
    *   [Step-by-Step Execution](#32-step-by-step-execution)
4.  [The Desktop Application](#4-the-desktop-application)
5.  [Preparing Your Resources](#5-preparing-your-resources)
    *   [Excel Data Structure](#51-excel-data-structure)
    *   [Word Template Design](#52-word-template-design)
6.  [Email Gateway Configuration](#6-email-gateway-configuration)
7.  [Troubleshooting & FAQ](#7-troubleshooting--faq)

---

## 1. Introduction
**DocuMint** is an enterprise-grade automation tool designed to generate personalized documents (PDFs) from Word templates and distribute them via email. It replaces manual mail merge processes with a robust, automated engine.

### Key Features
*   **Universal Email**: Works with Gmail, Yahoo, Office365, and minimal SMTP servers.
*   **Live Preview**: visual editor for HTML emails.
*   **Bangla Support**: First-class support for Unicode/Bangla fonts.
*   **Privacy**: All processing happens locally on your machine.

---

## 2. Installation & Setup
### Prerequisites
*   **OS**: Windows 10/11 (Preferred) or macOS/Linux (Web Studio only).
*   **Python**: Version 3.8 or higher.
*   **Microsoft Office**: Word (required for PDF conversion).

### Quick Install
1.  Download the project source code or `DocuMint_Portable.zip`.
2.  Open your terminal/command prompt.
3.  Install dependencies:
    ```bash
    pip install .
    ```

---

## 3. The Web Studio (Recommended)
The Web Studio is the new, modern way to use DocuMint. It features a clean, professional interface with real-time logs.

### 3.1 Interface Overview
*   **Sidebar**: Quick access to Dashboard, Documentation, and Developer Profile.
*   **File Resources Card**: Where you link your local Excel and Word files.
*   **Email Composer**: A split-screen editor where you write your email. Toggle between **"Write"** (Editor) and **"Preview"** modes.
*   **Terminal**: A black console at the bottom showing real-time system operations.

### 3.2 Step-by-Step Execution

#### Step 1: Launch the Server
```bash
python src/web/app.py
```
Open your browser to: `http://localhost:5000`

#### Step 2: Connect Your Data
1.  Locate your **Excel Data File**. Copy its full path (e.g., `D:\MyFiles\students.xlsx`).
2.  Paste it into the **Excel Data Path** field.
3.  Locate your **Word Template**. Copy its path (e.g., `D:\MyFiles\template.docx`).
4.  Paste it into the **Word Template Path** field.
5.  **Pro Tip**: Click the **"Check Integrity"** button. The system will read both files and confirm they match (e.g., if your Word file asks for `<Name>`, it checks if Excel has a `Name` column).

#### Step 3: Compose Your Email
1.  Enter a **Subject** line.
2.  Write your **Body**. You can use HTML formatting (`<b>bold</b>`, `<br>`).
3.  Click the **Preview** button to see exactly how it will look to the recipient.

#### Step 4: Configure Gateway & Send
1.  Select **Method**:
    *   **Custom SMTP** (Recommended for Gmail/Yahoo).
    *   **Outlook App** (Uses your local Outlook program).
2.  For SMTP, enter your **Host**, **User Email**, and **App Password**.
3.  Click **Start Engine**.
4.  Watch the **System Logs**. You will see:
    *   `Started processing...`
    *   `Generating PDF for: [Name]...`
    *   `Email sent to: [Email]...`

---

## 4. The Desktop Application
For users who prefer a native Windows window, the legacy GUI is still supported.

1.  Run `python src/main.py`.
2.  Follow the Tabbed interface (Welcome -> Files -> Email -> Run).
3.  The functionality is identical to the Web Studio but without Live Preview.

---

## 5. Preparing Your Resources

### 5.1 Excel Data Structure
Your Excel file acts as the database.
*   **Row 1 (Headers)**: This row defines your variable names.
    *   Example: `Name`, `EmployeeID`, `Department`, `Email`.
*   **Rows 2+ (Data)**: The actual data for each person.
*   **Mandatory Column**: You MUST have a column named `Email` (case-insensitive) if you plan to send emails.

| Name | ID | Department | Email |
| :--- | :--- | :--- | :--- |
| Zihad | 001 | Engineering | zihad@example.com |

### 5.2 Word Template Design
Your Word document serves as the design blueprint.
*   **Placeholders**: Use angle brackets `< >`.
*   **Matching**: The text inside `< >` must EXACTLY match an Excel header.
    *   Excel: `Name` -> Word: `<Name>`
    *   Excel: `EmployeeID` -> Word: `<EmployeeID>`
*   **Formatting**: You can style the placeholders (Bold, Red, Font Size) in Word. The replaced text will keep that style.

---

## 6. Email Gateway Configuration

### Using Gmail (Universal SMTP)
Gmail requires an **App Password** (not your login password).
1.  Go to **Google Account** -> **Security**.
2.  Enable **2-Step Verification**.
3.  Search for **"App Passwords"**.
4.  Create new: Select App="Mail", Device="Windows Computer".
5.  Copy the 16-character code (e.g., `abcd efgh ijkl mnop`).
6.  Use this as your **Password** in DocuMint.
7.  Host: `smtp.gmail.com`

### Using Outlook Desktop App
*   **No Password Required**: This method uses your installed Outlook program. You do NOT need an App Password.
*   **Prerequisite**: Microsoft Outlook must be installed and logged in.
*   **Classic Mode**: Ensure the "New Outlook" toggle is OFF (Classic Outlook application works best).

---

## 7. Troubleshooting & FAQ

**Q: The system says "Column <X> not found in Excel".**
*   **A**: Check your Excel file. Ensure Row 1 has that exact name. Check for hidden spaces (e.g., "Name " instead of "Name").

**Q: PDFs are generated, but Emails fail.**
*   **A**: Verify your App Password. Also, check internet connection. Firewall might block port 465.

**Q: Bangla font looks broken in the Web Preview.**
*   **A**: The Web Preview uses web fonts. Ensure your system has Bangla fonts installed for the final PDF generation (Word uses system fonts).

---
*For further support, contact the developer via the links in the Web Studio Sidebar.*
