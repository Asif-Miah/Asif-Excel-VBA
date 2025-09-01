# 📧 Excel VBA Email Automation

This repository contains **Excel VBA macros** to automate Outlook email management and Excel data organization.  

---

## ⚡ Features

- 📨 Categorize sent emails by recipient into separate sheets.  
- 📊 Split column data into multiple sheets based on unique values.  
- 🔔 Generate a report of all unread emails from all Outlook folders.  
- 🛠️ Clean sheet names to avoid Excel naming issues.  

---

## 🏁 Key Macros

- **MergeMacros** – Master macro to run all tasks sequentially.  
- **ListSentEmailsToSeparateSheetsByRecipient** – Organizes sent emails by recipient.  
- **SplitColumnBIntoSheets** – Splits active sheet data by unique values in Column B.  
- **CleanSheetName** – Helper to sanitize Excel sheet names.  
- **ListAllUnreadEmails_AllFolders** – Generates a report of all unread emails.  
- **ScanFoldersForUnread** – Recursively scans Outlook folders for unread emails.  

---

## 🛠️ Usage

1. Save Excel file as **macro-enabled** (`.xlsm`).  
2. Press `Alt + F11` → Insert this code into a module.  
3. Run `MergeMacros` to execute all steps.  
4. Ensure column B exists in the active sheet before splitting.  
5. Outlook must be installed and configured.  

---

**Made with ❤️ by Asif Miah**
