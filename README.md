# Gemini-Outreach-automation
An automated Google Apps Script that analyzes prospect websites using the Gemini API, dynamically drafts highly personalized B2B outreach emails, and sends them automatically via Gmail.

# AI-Driven Sales & Marketing Operations Automation

An automated workflow engine built with **Google Apps Script** and **Generative AI (Gemini API)**. This system transforms a static Google Sheet into an intelligent CRM and autonomous outreach agent by scanning prospect websites, understanding their business context, and generating highly personalized email campaigns sent directly through **Gmail**.

---

## 🧭 Project Architecture & Workflow

The script executes an automated pipeline divided into four distinct phases:

1. **Data Ingestion & Concurrency Control:** Reads active lead data from Google Sheets. It utilizes Google’s `LockService` to prevent race conditions or overlapping executions if multiple triggers fire simultaneously.
2. **Contextual AI Analysis:** Extracts the prospect's company website and feeds it into the `Gemini API`. The AI analyzes the target's business model and identifies key pain points.
3. **Hyper-Personalized Copywriting:** Based on the website analysis, the AI dynamically writes a custom B2B outreach email tailored specifically to the prospect's name, role, and industry.
4. **Automated Delivery & Logging:** Sends the message cleanly via the user's connected `Gmail` account and writes an instantaneous audit log back to the spreadsheet to prevent duplicate outreach.

---

## 🛠️ Tech Stack & Key APIs

* **Runtime Environment:** Google Apps Script (V8 Engine / JavaScript)
* **Generative AI Layer:** Gemini API (`UrlFetchApp` HTTPS orchestration)
* **Database Layer:** Google Sheets API (`SpreadsheetApp`)
* **Communication Layer:** Gmail API (`GmailApp`)
* **System Safety:** Core execution wrapped in `LockService` routines

---

## 📦 Step-by-Step Installation & Configuration

### Step 1: Secure Your Repository (Crucial Security Action)
Before uploading your code to a public GitHub repository, protect your private access credentials:
* Remove your raw **Google Sheet ID** string.
* Remove your raw **Gemini API Key** string.
* Replace them with safe environment variable handlers or variable placeholders.

### Step 2: Set Up Script Properties in Google Workspace
To run the script without hardcoding your private API key:
1. Inside your Google Apps Script project editor, click on the **Project Settings** (the gear icon ⚙️ on the left menu).
2. Scroll down to the **Script Properties** section.
3. Click **Add script property**.
4. Set the **Property** name to `GEMINI_API_KEY` and paste your actual Gemini token into the **Value** field.
5. Save the properties. The script will now safely fetch the key using:
   ```javascript
   const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
