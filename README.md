# Mandays Training Analytics Dashboard

An interactive, premium, dark-mode web application designed to visualize, filter, and drill down into employee training mandays. The dashboard is built with **modern glassmorphic aesthetics** and updates dynamically by fetching real-time data directly from a **Google Sheet** (or Google Apps Script Web App) at runtime, with fallback support for offline local file uploads.

---

## Key Features

### 1. Live Google Sheets Integration & Sync
- **Dynamic Fetching**: Automatically fetches live data on page load using the browser's native `fetch()` API.
- **Header URL Sync Bar**: Paste any public Google Sheet link (or published web link) directly into the header sync input box and press **Sync** to instantly refresh the dashboard.
- **Sticky Settings**: Successfully loaded URLs are stored in `localStorage`, meaning the dashboard remembers your preferred data sheet across browser refreshes.
- **Safe Fallbacks**: If the fetch fails (due to offline states or network boundaries), the uploader remains ready for local files.

### 2. Google Apps Script Web App Support (For Restricted Domains)
- Built-in support for private/restricted organization domains (such as `@vitstudent.ac.in` or enterprise Google Workspace).
- Detects Google Apps Script Web App links and automatically converts 2D array matrix responses (`[[headers], [row1], [row2]]`) into the correct object ledger format dynamically.

### 3. Dynamic Range Auto-Scaling (Pie Charts)
- Aggregates employee training mandays across all rows and distributes them into **5 range-analysis donut charts**.
- **Auto-Scaling**: Automatically detects if employee mandays totals are small (<= 50) or large (> 50) and scales the chart intervals accordingly (e.g., automatically shifting to `0-50, 50-100, ..., 200-250` intervals for larger sheets), ensuring **100% of employees are captured and visualised**.
- **Slice Click Drilldown**: Clicking any pie/donut slice opens an interactive detailed modal showing the **entire row (all 13 columns)** of all matching Excel rows, **sorted by department**.

### 4. Split-Row Leadership Track
- Dedicated split section taking up the left-to-middle layout space (60-65% width) side-by-side with secondary widgets.
- **Left Side**: Dynamic course buttons showing unique counts of enrolled participants.
- **Right Side**: Renders the participant ledger (Employee Name, Corp/Plant, Department) for the selected course, **sorted by department**. Clicking any row pops up the **entire raw row (all 13 columns)** of details, sorted by department.

### 5. Secondary Column Analytics (Right Side)
- **Business Excellence & TQM**: Target Plan (200 mandays) vs. dynamically computed completions Actuals (sum of mandays for rows with `attended === 'Yes'`) with counting animations.
- **Category Progress**: Dynamically lists unique person counts per training category with count-up micro-animations.

### 6. Searchable & Sortable Employee Ledger
- Lists all employees with their plant, department, and aggregated metrics.
- Default-sorted by **Department**.
- Support for text searches (by Name or Employee ID) and pagination control.

### 7. Drag-and-Drop Local File Uploader
- Allows users to drag and drop or select local `.xlsx` or `.csv` files to render metrics fully client-side inside the browser memory.

---

## Technical Stack & Libraries
- **Structure & Logic**: HTML5, Vanilla JavaScript (ES6)
- **Styles**: Vanilla CSS3 (Custom animations, glassmorphism, responsive grids down to `992px` screen width)
- **Chart Visualizations**: [Chart.js](https://www.chartjs.org/)
- **Spreadsheet Parsing**: [SheetJS (XLSX)](https://sheetjs.com/)

---

## Getting Started

### Local Development
To run this project locally, start a lightweight web server inside this directory:
```bash
# Using http-server
npx http-server -p 8080
```
Open **[http://localhost:8080](http://localhost:8080)** in your default browser.

---

## Connecting Your Google Sheet

### Option A: Standard Google Sheets (Publicly Shared)
1. Open your Google Sheet.
2. Click **Share** at the top right, and set General access to **"Anyone with the link can view"**.
3. Copy the URL from the browser address bar.
4. Paste it into the dashboard URL box and click **Sync**.

### Option B: Restricted Domain Sheets (e.g. VIT Student / Enterprise)
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
   - **Description**: Training Dashboard Fetcher
   - **Execute as**: Me (your email)
   - **Who has access**: Anyone (this lets the client-side JavaScript request data without a login prompt)
6. Click **Deploy**.
7. Copy the generated **Web App URL** (the link ending in `/exec`).
8. Paste that URL directly into the dashboard's URL Sync input box and click **Sync**.

---

## Local Repository Commands
To stage and commit these changes locally:
```bash
git add app.js index.html styles.css README.md .gitignore
git commit -m "Configure dynamic live Google Sheets syncing with Apps Script support"
```
