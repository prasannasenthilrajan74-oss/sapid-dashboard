# SAPID license analyser User Manual

Welcome to **SAPID license analyser**, an interactive, dark-mode desktop dashboard designed to visualize, filter, and analyze SAP user licensing distributions and purchased progress.

---

## Table of Contents
1. [Connecting a Standard Google Sheet](#1-connecting-a-standard-google-sheet)
2. [Connecting an Enterprise Google Sheet (Apps Script)](#2-connecting-an-enterprise-google-sheet-apps-script)
3. [Uploading Local Excel Files](#3-uploading-local-excel-files)
4. [Using Dashboard Metrics and Filters](#4-using-dashboard-metrics-and-filters)
5. [License Distribution Progress & Purchased Quotas](#5-license-distribution-progress--purchased-quotas)
6. [SAP User Group Explorer](#6-sap-user-group-explorer)
7. [Navigating the SAP User Ledger](#7-navigating-the-sap-user-ledger)
8. [Exporting and Importing Configurations](#8-exporting-and-importing-configurations)
9. [Reviewing Error Logs](#9-reviewing-error-logs)

---

## 1. Connecting a Standard Google Sheet
If your spreadsheet is publicly shared or viewable by anyone with the link:
1. Open your Google Sheet in your web browser.
2. Click the **Share** button in the top right, and set general access to **"Anyone with the link can view"**.
3. Copy the URL from the browser address bar.
4. Paste the URL into the **Google Sheets Link...** input box in the header sync bar.
5. Click **Sync** (or press Enter). The dashboard will reload and visualize the active sheet data.

*Note: The sheet must follow the SAP License schema with columns: SAPID, Name, License, Department, Function, Last Logon, and Group.*

---

## 2. Connecting an Enterprise Google Sheet (Apps Script)
For enterprise domains that restrict public link sharing, configure a Google Apps Script fetcher:
1. Open your sheet, click **Extensions** -> **Apps Script**.
2. Replace any code with this script:
   ```javascript
   function doGet() {
     var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
     var data = sheet.getDataRange().getValues();
     return ContentService.createTextOutput(JSON.stringify(data))
       .setMimeType(ContentService.MimeType.JSON);
   }
   ```
3. Click **Deploy** -> **New Deployment**. Set the type to **Web App**.
4. Configure deployment settings:
   - **Execute as**: Me (your email)
   - **Who has access**: Anyone (essential for WebView browser fetches)
5. Deploy and copy the generated **Web App URL** (ends in `/exec`).
6. Paste the URL into the sync input box and click **Sync**.
7. Ensure you are signed in to the correct Google Workspace account in your browser so that Tauri's internal WebView can fetch the restricted URL.

---

## 3. Uploading Local Excel Files
To work entirely offline without cloud links:
1. Drag and drop your `.xlsx` or `.xls` spreadsheet file onto the **Upload Excel Sheet** area in the top-right header, or click the area to choose a file via the system dialog.
2. The application will parse the binary file in browser memory.
3. Once loaded, a **Reset to Default Data** button will appear in the sync bar, allowing you to return to the preloaded 500-user SAP dataset.

---

## 4. Using Dashboard Metrics and Filters
- **KPI Summary Cards**:
  - **Total Users**: Counts the total allocated users matching current active filters.
  - **Total Purchased**: The sum of all purchased license quotas originally bought from the head office.
  - **Available Licenses**: Remaining available licenses (Total Purchased - Total Allocated Users).
  - **Top License Type**: License ID (AX, AY, FX, HC, HD) representing the largest share.
- **Interactive Filters**: Refine the metrics by License ID or Department. All users are treated as active.

---

## 5. License Distribution Progress & Purchased Quotas
The right column displays the allocation progress for each license type:
- **Editable Purchased Quotas**: Every progress card features an inline **Limit** input field. You can directly edit the total purchased license quota originally bought from the head office.
- **Recalculations**: Modifying a limit instantly updates the **Total Purchased** and **Available Licenses** KPI cards and recalculates the progress bar percentage.
- **Persistence**: Your edited limits are automatically persisted locally in `localStorage` under `sap_license_purchased_[licName]`, meaning your custom numbers remain intact across application sessions.

---

## 6. SAP User Group Explorer
Select any SAP functional group from the button grid on the left. The assigned users ledger on the right will update to list unique enrolled users sorted by department. Click any row in the ledger to view the raw record details.

---

## 7. Navigating the SAP User Ledger
Click **Employee Ledger** in the navigation bar to view the database:
- **Search bar**: Search for users by Name or SAPID.
- **Filters**: Filter list by License ID or Department.
- **Sorting**: Click any column header (SAPID, User Name, License, Department, User Group, Last Logon) to toggle sorting.
- **Detail Drilldown**: Click any row in the ledger to display a detail modal of that user's profile.

---

## 8. Exporting and Importing Configurations
You can back up or transfer your active dashboard URL and configurations:
- **Export Config**: Click **Export Config** in the footer. This generates a `sap_license_config_[date].json` file containing your saved sheet URL and cached data.
- **Import Config**: Click **Import Config** and select a previously exported `.json` file to restore the configuration instantly.

---

## 9. Reviewing Error Logs
If a sync fails or data fails to parse, the system writes a diagnostic log entry:
- Click **Error Logs** in the footer to review recent log messages.
- Click **Open Logs Folder** to open your native OS folder explorer containing the `error.log` file.
- Click **Clear Logs** to truncate the file and keep disk usage minimal.
