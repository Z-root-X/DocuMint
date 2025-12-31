# 📘 DocuMint User Guide

Welcome to the comprehensive guide for DocuMint. This document covers advanced usage, troubleshooting, and tips for getting the most out of the application.

## Table of Contents
1.  [Template Design](#1-template-design)
2.  [Excel Data Preparation](#2-excel-data-preparation)
3.  [Email Configuration](#3-email-configuration)
    *   [Outlook Mode](#outlook-mode)
    *   [SMTP Mode (Gmail, Yahoo, etc.)](#smtp-mode)
4.  [Validation & Safety](#4-validation--safety)

---

## 1. Template Design
DocuMint uses **Placeholders** to inject data into your Word Documents.

### Syntax
Format: `<ColumnHeader>`

*   **Example**: If your Excel header is `StudentName`, use `<StudentName>` in the Word doc.
*   **Styling**: You can bold, color, or change the font of the placeholder in Word. The replaced text will inherit the same style.

> **💡 Pro Tip**: Ensure your placeholders do not contain spaces inside the brackets if your Excel headers don't have them.

---

## 2. Excel Data Preparation
Your data source must be an `.xlsx` or `.xls` file.

*   **Row 1**: Must contain **Headers**. These are the keys for your placeholders.
*   **Email Column**: One column must contain email addresses. DocuMint defaults to looking for `Email`, `E-mail`, or you can assume the column named `Email` is used by the logic.
*   **Clean Data**: Ensure there are no empty rows in the middle of your dataset.

---

## 3. Email Configuration

### Outlook Mode
*   **Requirement**: Microsoft Outlook Desktop App installed and logged in.
*   **Pros**: Uses your existing Outlook signature and sent folder.
*   **Cons**: Slower, Windows only.

### SMTP Mode
Allows sending without Outlook. Great for bulk sending from a specific server.

#### Gmail Setup
1.  **Host**: `smtp.gmail.com`
2.  **Port**: `465` (SSL)
3.  **User**: Your full gmail address.
4.  **Password**: You **cannot** use your normal password. You must generate an **App Password**:
    *   Go to Google Account > Security.
    *   Enable 2-Step Verification.
    *   Search for "App Passwords" and create one for DocuMint.
    *   Use that 16-character code here.

---

## 4. Validation & Safety
Before running a large batch:
1.  Load your files.
2.  Click **Validate Files** in the setup screen.
3.  DocuMint scans every `<Tag>` in your `.docx`.
4.  It checks if that `Tag` exists as a column in your `.xlsx`.
5.  If any are missing, it stops you. **This prevents sending 500 emails with "Dear <Name>" instead of the actual name.**
