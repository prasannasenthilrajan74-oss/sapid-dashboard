# SkillTrack Analyzer User Manual

Welcome to **SkillTrack Analyzer**, an interactive, dark-mode desktop dashboard designed to visualize, filter, and drill down into employee training mandays and course enrolments.

---

## Table of Contents
1. [Connecting a Standard Google Sheet](#1-connecting-a-standard-google-sheet)
2. [Connecting an Enterprise Google Sheet (Apps Script)](#2-connecting-an-enterprise-google-sheet-apps-script)
3. [Uploading Local Excel/CSV Files](#3-uploading-local-excelcsv-files)
4. [Using Dashboard Metrics and Filters](#4-using-dashboard-metrics-and-filters)
5. [Understanding Dynamic Range Auto-Scaling](#5-understanding-dynamic-range-auto-scaling)
6. [Navigating the Employee Ledger](#6-navigating-the-employee-ledger)
7. [Exporting and Importing Configurations](#7-exporting-and-importing-configurations)
8. [Reviewing Error Logs](#8-reviewing-error-logs)

---

## 1. Connecting a Standard Google Sheet
If your spreadsheet is publicly shared or viewable by anyone with the link:
1. Open your Google Sheet in your web browser.
2. Click the **Share** button in the top right, and set general access to **"Anyone with the link can view"**.
3. Copy the URL from the browser address bar.
4. Paste the URL into the **Google Sheets Link...** input box in the header sync bar.
5. Click **Sync** (or press Enter). The dashboard will reload and visualize the active sheet data.

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

## 3. Uploading Local Excel/CSV Files
To work entirely offline without setting up any cloud links:
1. Drag and drop your `.xlsx` or `.csv` spreadsheet file onto the **Upload Excel Sheet** area in the top-right header, or click the area to choose a file via the system dialog.
2. The application will parse the binary file in browser memory.
3. Once loaded, a **Reset to Default Data** button will appear in the sync bar, allowing you to return to the default cloud data at any time.

---

## 4. Using Dashboard Metrics and Filters
- **KPI Summary Cards**: Total Mandays, Total Unique Employees, Avg Mandays per Employee, and Top Plant are calculated dynamically.
- **Interactive Filters**: Refine the metrics by Plant, Department, or Course Attendance using the dropdown panels. All KPI cards and visualizations will transition instantly.
- **Split-Row Leadership Track**: Select any course from the button grid on the left. The participant ledger on the right will update to list enrolled participants sorted by department. Click any row in the ledger to view the raw 13-column record.

---

## 5. Understanding Dynamic Range Auto-Scaling
The range donut charts aggregate employee training mandays and segment them into intervals:
- **Auto-Scaling**: If employee mandays totals are small (<= 50) or large (> 50), the intervals adjust automatically (e.g. scaling to `0-50`, `50-100`... up to `200-250` for large totals). This prevents chart compression and ensures all employees are accounted for.
- **Slice Click Drilldown**: Click any color slice on any range chart to open a modal view showing the complete spreadsheet entries for matching employees.

---

## 6. Navigating the Employee Ledger
Click **Employee Ledger** in the navigation bar to view the database:
- **Search bar**: Search for employees by Name or Employee ID.
- **Sorting**: Click any column header (Employee ID, Name, Total Mandays, etc.) to toggle ascending or descending sorting.
- **Drilldown**: Click any row in the ledger to display a chronological modal history of all training courses that employee attended.

---

## 7. Exporting and Importing Configurations
You can back up or transfer your active dashboard URL and configurations:
- **Export Config**: Click **Export Config** in the footer. This generates a `skilltrack_config_[date].json` file containing your saved sheet URL and cached data.
- **Import Config**: Click **Import Config** and select a previously exported `.json` file to restore the configuration instantly.

---

## 8. Reviewing Error Logs
If a sync fails or data fails to parse, the system writes a diagnostic log entry:
- Click **Error Logs** in the footer to review recent log messages.
- Click **Open Logs Folder** to open your native OS folder explorer containing the `error.log` file.
- Click **Clear Logs** to truncate the file and keep disk usage minimal.
