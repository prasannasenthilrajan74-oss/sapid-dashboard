# SAPID license analyser

An interactive, premium, dark-mode cross-platform desktop application designed to visualize, filter, and analyze SAP user licensing distributions and logon recency activity. The dashboard is built with **modern glassmorphic aesthetics** and updates dynamically by fetching real-time data directly from a **Google Sheet** (or Google Apps Script Web App) at runtime, with offline support for local Excel uploads.

---

## Key Features

### 1. Live Google Sheets Integration & Sync
- **Dynamic Fetching**: Automatically fetches live data on page load.
- **Header URL Sync Bar**: Paste any public Google Sheet link (or Apps Script Web App link) directly into the header sync input box and click **Sync** to instantly refresh the dashboard.
- **Persistent Settings**: Successfully synced URLs are cached in `localStorage` to persist across launches.
- **Offline Fallbacks**: If the fetch fails, local file uploads remain available.

### 2. Google Apps Script Web App Support (For Private Domains)
- Built-in support for private/restricted organization domains (such as enterprise Google Workspace accounts).
- Detects Google Apps Script Web App links and automatically converts 2D array matrix responses (`[[headers], [row1], [row2]]`) into the correct object ledger format dynamically.

### 3. Logon Recency Analysis
- Aggregates user last logon dates relative to the benchmark date (`June 2, 2026`) and distributes them into **5 range-analysis donut charts**: Active (0-30 days), Recent (31-60 days), Inactive (61-90 days), Long Inactive (91-120 days), and Dormant (120+ days).
- **Slice Click Drilldown**: Clicking any pie/donut slice opens an interactive detailed modal showing the **entire profile details** of matching users (SAPID, Name, License, Department, Function, Last Logon, Group), sorted by department.
- **Toggle Grouping**: Instantly toggle slices to group users by either SAP License or Department.

### 4. SAP User Group Explorer
- Dedicated functional group explorer taking up the left-to-middle layout space (60-65% width) side-by-side with secondary widgets.
- **Left Side**: Dynamic group buttons showing unique counts of assigned users.
- **Right Side**: Renders the assigned user list (User Name, License, Department) for the selected group, **sorted by department**. Clicking any row pops up the **entire raw profile** of details.

### 5. Secondary Column Analytics (Right Side)
- **Logon Activity Target**: Plan target (100 users) vs. dynamically computed completions Actuals (sum of active users with logon within 30 days) with counting animations.
- **License Distribution Progress**: Dynamically lists unique user counts per license type (AX, AY, FX, HC, HD) with count-up micro-animations.

### 6. Searchable & Sortable SAP User Ledger
- Lists all users with their SAPID, User Name, License, Department, User Group, and Last Logon.
- Support for text searches (by Name or SAPID) and pagination control.

### 7. Drag-and-Drop Local File Uploader
- Allows users to drag and drop or select local `.xlsx` or `.xls` files to render metrics fully client-side inside the browser memory.

---

## Technical Stack & Libraries
- **Desktop Wrapper**: [Tauri v2](https://tauri.app/) (Rust-backed cross-platform native wrapper)
- **Frontend Core**: HTML5, Vanilla JavaScript (ES6)
- **Styles**: Vanilla CSS3 (Custom animations, glassmorphism, responsive grids down to `992px` screen width)
- **Chart Visualizations**: [Chart.js](https://www.chartjs.org/)
- **Spreadsheet Parsing**: [SheetJS (XLSX)](https://sheetjs.com/)

---

## Getting Started

### Local Development
To run this project locally, start a lightweight web server inside this directory or run Tauri dev:
```bash
# Start Tauri development environment
npm run tauri dev
```

---

## Connecting Your Google Sheet

### Option A: Standard Google Sheets (Publicly Shared)
1. Open your Google Sheet.
2. Click **Share** at the top right, and set General access to **"Anyone with the link can view"**.
3. Copy the URL from the browser address bar.
4. Paste it into the dashboard URL box and click **Sync**.

### Option B: Restricted Domain Sheets (e.g. Enterprise Organization)
If your organization restricts external sharing, you can fetch data securely by deploying a micro Google Apps Script Web App:

1. In your Google Sheet, open **Extensions** -> **Apps Script**.
2. Replace any existing code with the following snippet:
   ```javascript
   function doGet() {
     var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
     var data = sheet.getDataRange().getValues();
     return ContentService.createTextOutput(JSON.stringify(data))
       .setMimeType(ContentService.MimeType.JSON);
   }
   ```
3. Click **Deploy** -> **New Deployment**.
4. Set "Select type" to **Web app**.
5. Configure the deployment settings:
   - **Description**: SAPID license analyser Fetcher
   - **Execute as**: Me (your email)
   - **Who has access**: Anyone (this lets the client-side JavaScript request data without a login prompt)
6. Click **Deploy**.
7. Copy the generated **Web App URL** (the link ending in `/exec`).
8. Paste that URL directly into the dashboard's URL Sync input box and click **Sync**.
