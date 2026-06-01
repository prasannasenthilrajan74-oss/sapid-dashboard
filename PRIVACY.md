# Privacy Policy for SkillTrack Analyzer

**Last Updated: June 1, 2026**

SkillTrack Analyzer is built as a cross-platform desktop application using the Tauri framework. We take your privacy very seriously. This document describes how your data is handled.

---

## 1. Local-First Processing
All spreadsheet data, employee ledger information, training records, and analytics configurations are parsed and processed **entirely client-side on your local machine**. 
- No spreadsheet rows or employee metrics are uploaded to third-party databases, tracking platforms, or advertising networks.
- The raw spreadsheets you select or drop into the file uploader are loaded directly into the webview's temporary browser memory and are never saved or sent to any remote server.

## 2. Direct Integration with Google
When you synchronize with a public Google Sheet or Google Apps Script Web App:
- The network requests are routed **directly** from your desktop client to Google's official API servers (`docs.google.com` or `script.google.com`).
- There are no intermediary proxy servers, middleware, or SaaS layers collecting your sync requests.
- For domain-restricted or private corporate Google Workspace URLs, the fetch is completed directly through the WebView using your local browser Google authentication session. No authorization tokens or user login credentials are ever seen, read, or stored by the application.

## 3. Persistent Local Storage
The application uses local web browser database files (`localStorage`) to remember:
- The URL of the Google Sheet you synced last.
- A local cache of the parsed sheet data, enabling the app to load instantly and operate in offline modes without a network connection.

You can fully clear this data at any time by pressing the **Reset to Default Data** button in the header sync bar, or by importing a different configuration.

## 4. Local Logging
Application runtime logs (e.g. sync failures or parsing issues) are appended to a plain text log file (`error.log`) saved in your local operating system's designated app data directory (e.g. `AppData/Local/com.skilltrack.analyzer/logs/error.log`).
- These logs are designed to aid in manual diagnostic and troubleshooting procedures.
- They are kept strictly local and are never auto-transmitted.
- You can review and clear these logs directly from the **Error Logs** panel in the footer.
