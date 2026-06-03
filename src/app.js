// ─── Tauri IPC Bridge ─────────────────────────────────────────────────────────
// Uses window.__TAURI__.core.invoke — injected by Tauri into the WebView at runtime.
// Falls back gracefully to null if running outside Tauri (e.g. plain browser).
async function tauriInvoke(command, args) {
  if (window.__TAURI__ && window.__TAURI__.core) {
    return window.__TAURI__.core.invoke(command, args);
  }
  throw new Error('Tauri IPC not available. Run this app via Tauri.');
}

// HTML Escaper to prevent DOM XSS
function escapeHTML(str) {
  if (str === null || str === undefined) return '';
  const s = String(str);
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

const MAX_CONFIG_FILE_BYTES = 2 * 1024 * 1024;
const MAX_EXCEL_FILE_BYTES = 10 * 1024 * 1024;
const MAX_CACHED_ROWS = 25000;

// Creates an option with DOM APIs so spreadsheet values cannot become executable HTML.
function appendSafeOption(selectElement, value, label) {
  const option = document.createElement('option');
  option.value = String(value);
  option.textContent = String(label);
  selectElement.appendChild(option);
}

// Validates imported row arrays before writing them to localStorage or app state.
function validateImportedRows(rows) {
  if (!Array.isArray(rows)) {
    throw new Error('Cached sheet data must be an array.');
  }
  if (rows.length > MAX_CACHED_ROWS) {
    throw new Error(`Cached sheet data is too large. Maximum allowed rows: ${MAX_CACHED_ROWS}.`);
  }
  return rows;
}

// Stores sheet data only after size checks so localStorage cannot be filled accidentally.
function cacheSheetData(rows) {
  const validatedRows = validateImportedRows(rows);
  localStorage.setItem('cached_sheet_data', JSON.stringify(validatedRows));
}

// Keeps imported sync URLs limited to the Google sources accepted by the Tauri backend.
function isAllowedGoogleDataUrl(urlValue) {
  try {
    const parsed = new URL(urlValue);
    const path = parsed.pathname;
    const isSheet = parsed.hostname === 'docs.google.com' && path.startsWith('/spreadsheets/d/');
    const isScript = parsed.hostname === 'script.google.com'
      && (path.startsWith('/macros/s/') || path.startsWith('/a/macros/'))
      && path.endsWith('/exec');
    return parsed.protocol === 'https:' && (isSheet || isScript);
  } catch {
    return false;
  }
}

// Returns the sheet name matching priority keywords (case-insensitive, trimmed), falling back to the first sheet.
function getBestSheetName(workbook) {
  if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
    return null;
  }
  const targets = [
    'live users',
    'live user',
    'acctive user',
    'active user',
    'sapidinfo',
    'sapid info',
    'sapid',
    'total id'
  ];
  for (const target of targets) {
    const matched = workbook.SheetNames.find(name => {
      if (typeof name !== 'string') return false;
      return name.trim().toLowerCase() === target.toLowerCase();
    });
    if (matched) return matched;
  }
  return workbook.SheetNames[0];
}

// Show a premium glassmorphic toast notification
function showToast(message, type = 'success') {
  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }
  
  const toast = document.createElement('div');
  toast.className = `toast-notification ${type}`;
  toast.innerHTML = `
    <span class="toast-icon">${type === 'success' ? '✅' : '❌'}</span>
    <span class="toast-message">${escapeHTML(message)}</span>
  `;
  
  container.appendChild(toast);
  
  // Trigger transition
  setTimeout(() => {
    toast.classList.add('show');
  }, 10);
  
  // Auto-remove after 3 seconds
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => {
      toast.remove();
    }, 300);
  }, 3000);
}

// Dashboard Application State
let rawData = [];
let employeesList = []; // Aggregated employees list
let currentFilters = {
  license: 'ALL',
  dept: 'ALL'
};

let selectedExplorerGroup = '';
let groupCurrentPage = 1;
const groupPageSize = 8;

// Chart Instances
let licenseChartInstance = null;
let deptChartInstance = null;
let plantIdChartInstance = null;
let leadershipChartInstance = null;

// Default purchased license totals (originally bought from the head office)
const DEFAULT_LICENSE_PURCHASED = {
  'AX': 78,
  'AY': 6,
  'FX': 37,
  'HC': 83,
  'HD': 41,
  'Other': 150
};

// Color Palettes
const LICENSE_COLORS = {
  'AX': '#0ea5e9',       // Sky Blue
  'AY': '#6366f1',       // Indigo
  'FX': '#a855f7',       // Purple
  'HC': '#ec4899',       // Pink
  'HD': '#14b8a6',       // Teal
  'Other': '#64748b'     // Muted Slate
};

const DEPT_COLORS = {
  'SD': '#f43f5e',
  'QM': '#10b981',
  'PM': '#f59e0b',
  'R&D': '#3b82f6',
  'MM': '#8b5cf6',
  'HCM': '#06b6d4',
  'PP': '#84cc16',
  'MED': '#ec4899',
  'FICO': '#14b8a6',
  'Other': '#64748b'
};

// --- ROBUST EXCEL COLUMN KEY ACCESSORS ---
function getRobustRowValue(row, targetKeys, defaultValue = '') {
  if (!row) return defaultValue;
  for (const key of targetKeys) {
    if (row[key] !== undefined && row[key] !== null) {
      return String(row[key]).trim();
    }
  }
  const cleanKey = (k) => String(k).toLowerCase().replace(/[\s\-_]/g, '');
  const cleanTargetKeys = targetKeys.map(cleanKey);
  for (const actualKey of Object.keys(row)) {
    if (cleanTargetKeys.includes(cleanKey(actualKey))) {
      const val = row[actualKey];
      if (val !== undefined && val !== null) {
        return String(val).trim();
      }
    }
  }
  return defaultValue;
}

function getRowLicense(row) {
  return getRobustRowValue(row, ['License', 'license', 'license type', 'licenseid', 'license id'], 'Unknown');
}

function getRowDepartment(row) {
  return getRobustRowValue(row, ['Department', 'department', 'dept', 'dep', 'division', 'div'], 'Unknown');
}

function getRowGroup(row) {
  return getRobustRowValue(row, ['Group', 'group', 'user group', 'usergroup', 'grp'], '');
}

function getRowSAPID(row) {
  return getRobustRowValue(row, ['SAPID', 'sapid', 'sap id', 'sap_id', 'id', 'user id', 'userid', 'emp id', 'empid', 'employee id'], '');
}

function getRowUserName(row) {
  return getRobustRowValue(row, ['Name', 'name', 'user name', 'username', 'employee name', 'emp name', 'full name', 'fullname'], 'Unknown');
}

function getRowFunction(row) {
  return getRobustRowValue(row, ['Function', 'function', 'role', 'job title', 'jobtitle', 'group', 'user group'], 'Unknown');
}

function getRowLastLogon(row) {
  return getRobustRowValue(row, ['Last Logon', 'last logon', 'lastlogon', 'last_logon', 'logon', 'last login', 'lastlogin'], '-');
}

// Function to convert Google Sheets URL to CSV export URL
function getGoogleSheetsCsvUrl(urlStr) {
  if (!urlStr) return null;
  
  // If it's a published Google Sheet ("Publish to web")
  if (urlStr.includes('/pub')) {
    if (urlStr.includes('output=csv')) {
      return urlStr;
    }
    // Convert HTML published URL to CSV output
    return urlStr.replace(/\/pub([?#].*)?$/, '/pub?output=csv');
  }

  const idMatch = urlStr.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch) return null;
  const sheetId = idMatch[1];
  
  let gid = '0';
  const gidMatch = urlStr.match(/[?&#]gid=([0-9]+)/);
  if (gidMatch) {
    gid = gidMatch[1];
  }
  
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
}

// ─── Main Data Fetch (via Tauri commands) ─────────────────────────────────────
async function fetchGoogleSheetsData(sheetUrl) {
  // Show a visual loading state in the sync button
  const syncBtn = document.getElementById('sheet-sync-btn');
  const originalText = syncBtn ? syncBtn.innerText : 'Sync';
  if (syncBtn) {
    syncBtn.innerText = 'Syncing...';
    syncBtn.disabled = true;
  }

  // Detect URL type
  const isDomainScript = sheetUrl.includes('script.google.com/a/macros/');
  const isAppsScript   = isDomainScript || sheetUrl.includes('script.google.com/macros/s/');

  try {
    let parsedData = [];

    if (isDomainScript) {
      // ── Domain-restricted Apps Script (/a/macros/) ────────────────────────
      // Step 1: Validate against whitelist via Tauri command (server never fetches the URL itself)
      const validateResult = await tauriInvoke('validate_url', { url: sheetUrl });
      if (!validateResult.success) {
        throw new Error(validateResult.message || '🔒 Access denied: This link is not authorized.');
      }

      // Step 2: Fetch directly from the browser — the user's Google session cookie
      //         is automatically sent via the WebView, so Google accepts the domain-restricted URL.
      const directRes = await fetch(sheetUrl, {
        credentials: 'include',   // sends the user's Google cookies
        mode: 'cors'
      });

      if (!directRes.ok) {
        throw new Error(`Google returned status ${directRes.status}. Make sure you are logged in to the correct Google Workspace account in this browser.`);
      }

      // Try to detect if we got an HTML login-redirect instead of JSON
      const contentType = directRes.headers.get('content-type') || '';
      if (contentType.includes('text/html')) {
        throw new Error('Google redirected to a login page. Please ensure you are signed in to your organization\'s Google account in this browser, then try again.');
      }

      const jsonData = await directRes.json();

      // Check if it's a 2D array (e.g. [[header1, header2], [val1, val2]])
      if (Array.isArray(jsonData) && jsonData.length > 0 && Array.isArray(jsonData[0])) {
        const headers = jsonData[0];
        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          const obj = {};
          headers.forEach((header, colIdx) => {
            obj[header] = row[colIdx] !== undefined ? row[colIdx] : '';
          });
          parsedData.push(obj);
        }
      } else if (Array.isArray(jsonData)) {
        parsedData = jsonData;
      } else {
        throw new Error('Unexpected JSON format from Google Apps Script.');
      }

    } else {
      // ── Regular Apps Script or Google Sheets — route through Tauri command ─
      const result = await tauriInvoke('fetch_sheet_data', { url: sheetUrl });

      if (!result.success) {
        throw new Error(result.message || 'Failed to fetch sheet data.');
      }

      if (result.is_json && result.data) {
        // JSON response from regular public Apps Script
        const jsonData = result.data;
        if (Array.isArray(jsonData) && jsonData.length > 0 && Array.isArray(jsonData[0])) {
          const headers = jsonData[0];
          for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            const obj = {};
            headers.forEach((header, colIdx) => {
              obj[header] = row[colIdx] !== undefined ? row[colIdx] : '';
            });
            parsedData.push(obj);
          }
        } else if (Array.isArray(jsonData)) {
          parsedData = jsonData;
        } else {
          throw new Error('Unexpected JSON format from Google Apps Script.');
        }
      } else if (result.csv) {
        // CSV response from Google Sheets
        const workbook = XLSX.read(result.csv, { type: 'string' });
        const sheetName = getBestSheetName(workbook);
        const worksheet = workbook.Sheets[sheetName];
        parsedData = XLSX.utils.sheet_to_json(worksheet);
      }
    }

    if (parsedData.length === 0) {
      throw new Error('The data is empty or could not be parsed. Check that your Apps Script returns a non-empty JSON array.');
    }

    rawData = parsedData;

    // Save successfully loaded URL and data to localStorage
    localStorage.setItem('dashboard_sheet_url', sheetUrl);
    cacheSheetData(parsedData);

    // Update reset button visibility
    const resetBtn = document.getElementById('reset-btn');
    if (resetBtn) {
      resetBtn.classList.remove('hidden');
    }

    initDashboard();
  } catch (err) {
    console.error('Error fetching Google Sheets data:', err);
    tauriInvoke('append_log', { level: 'error', message: `Google Sheet fetch error: ${err.message}` }).catch(console.error);
    alert(`Failed to sync: ${err.message}\nFalling back to local cached data.`);

    // Fallback to local storage cache
    const cachedDataStr = localStorage.getItem('cached_sheet_data');
    if (cachedDataStr) {
      try {
        rawData = JSON.parse(cachedDataStr);
        initDashboard();
        return;
      } catch (parseErr) {
        console.error('Failed to parse cached data:', parseErr);
      }
    }

    if (rawData.length === 0) {
      rawData = [];
      initDashboard();
    }
  } finally {
    if (syncBtn) {
      syncBtn.innerText = originalText;
      syncBtn.disabled = false;
    }
  }
}

// Initialize Application
window.addEventListener('DOMContentLoaded', () => {
  // Setup Chart.js global defaults for dark mode if library is loaded
  if (typeof Chart !== 'undefined') {
    Chart.defaults.color = '#94a3b8';
    Chart.defaults.font.family = "'Inter', sans-serif";
    Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(15, 23, 42, 0.95)';
    Chart.defaults.plugins.tooltip.titleColor = '#fff';
    Chart.defaults.plugins.tooltip.bodyColor = '#e2e8f0';
    Chart.defaults.plugins.tooltip.borderColor = 'rgba(255, 255, 255, 0.1)';
    Chart.defaults.plugins.tooltip.borderWidth = 1;
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.cornerRadius = 8;
  } else {
    console.warn("Chart.js library is not loaded. Visual charts will be disabled.");
  }

  // Setup Event Listeners
  setupEventListeners();

  // Clear legacy default URL on startup
  localStorage.removeItem('dashboard_sheet_url');
  localStorage.removeItem('cached_sheet_data');
  let savedUrl = null;
  const urlInput = document.getElementById('sheet-url-input');
  if (urlInput) {
    urlInput.value = savedUrl || '';
  }

  // Initialize with saved URL or fallback to local preloaded SAP data
  if (savedUrl) {
    fetchGoogleSheetsData(savedUrl);
  } else {
    const cachedDataStr = localStorage.getItem('cached_sheet_data');
    if (cachedDataStr) {
      try {
        rawData = JSON.parse(cachedDataStr);
        // Quick verification that cache contains some data
        if (rawData.length > 0) {
          // If cached URL was set, show reset button
          const resetBtn = document.getElementById('reset-btn');
          if (resetBtn && savedUrl) {
            resetBtn.classList.remove('hidden');
          }
          initDashboard();
          return;
        }
      } catch (parseErr) {
        console.error('Failed to parse cached data:', parseErr);
      }
    }
    
    rawData = [];
    initDashboard();
  }
});

// Setup event handlers
function setupEventListeners() {
  // File Uploader
  const fileInput = document.getElementById('excel-file');
  fileInput.addEventListener('change', handleFileUpload);

  // Reset Button
  const resetBtn = document.getElementById('reset-btn');
  resetBtn.addEventListener('click', () => {
    localStorage.removeItem('dashboard_sheet_url');
    localStorage.removeItem('cached_sheet_data');
    const urlInput = document.getElementById('sheet-url-input');
    if (urlInput) {
      urlInput.value = '';
    }
    fileInput.value = '';
    resetBtn.classList.add('hidden');
    
    rawData = [];
    initDashboard();
  });

  // Google Sheets Sync
  const syncBtn = document.getElementById('sheet-sync-btn');
  const urlInput = document.getElementById('sheet-url-input');
  
  if (syncBtn && urlInput) {
    const triggerSync = () => {
      const enteredUrl = urlInput.value.trim();
      if (enteredUrl) {
        fetchGoogleSheetsData(enteredUrl);
      }
    };
    
    syncBtn.addEventListener('click', triggerSync);
    urlInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        triggerSync();
      }
    });
  }

  // Filter Changes
  document.getElementById('filter-license').addEventListener('change', (e) => {
    currentFilters.license = e.target.value;
    updateDashboard();
  });
  document.getElementById('filter-dept').addEventListener('change', (e) => {
    currentFilters.dept = e.target.value;
    updateDashboard();
  });



  // Modal Close
  document.getElementById('modal-close-btn').addEventListener('click', closeModal);
  document.getElementById('drilldown-modal').addEventListener('click', (e) => {
    if (e.target.id === 'drilldown-modal') closeModal();
  });

  // SAP User Group Explorer filter category change event
  document.getElementById('filter-explorer-license').addEventListener('change', () => {
    groupCurrentPage = 1;
    renderLeadershipChart();
  });

  // SAP User Group Explorer pagination events
  const btnPrevGroups = document.getElementById('btn-prev-groups');
  const btnNextGroups = document.getElementById('btn-next-groups');
  if (btnPrevGroups && btnNextGroups) {
    btnPrevGroups.addEventListener('click', () => {
      if (groupCurrentPage > 1) {
        groupCurrentPage--;
        renderLeadershipChart();
      }
    });
    btnNextGroups.addEventListener('click', () => {
      groupCurrentPage++;
      renderLeadershipChart();
    });
  }

  // Input handler for editable purchased limits in category progress list
  const progressContainer = document.getElementById('category-progress-container');
  if (progressContainer) {
    progressContainer.addEventListener('input', (e) => {
      if (e.target.classList.contains('limit-input')) {
        const lic = e.target.getAttribute('data-license');
        let val = parseInt(e.target.value);
        if (isNaN(val) || val < 0) {
          val = 0;
        }
        localStorage.setItem(`sap_license_purchased_${lic}`, val);
        updateDashboard();
      }
    });
  }

  // ─── Settings Footer & Modals Listeners ─────────────────────────────────────
  
  // Privacy Policy modal
  const linkPrivacy = document.getElementById('link-privacy-policy');
  const modalPrivacy = document.getElementById('privacy-modal');
  const btnClosePrivacy = document.getElementById('btn-close-privacy');
  
  if (linkPrivacy && modalPrivacy && btnClosePrivacy) {
    linkPrivacy.addEventListener('click', (e) => {
      e.preventDefault();
      modalPrivacy.classList.add('active');
    });
    btnClosePrivacy.addEventListener('click', () => {
      modalPrivacy.classList.remove('active');
    });
    modalPrivacy.addEventListener('click', (e) => {
      if (e.target === modalPrivacy) modalPrivacy.classList.remove('active');
    });
  }

  // Licenses modal
  const linkLicenses = document.getElementById('link-licenses');
  const modalLicenses = document.getElementById('licenses-modal');
  const btnCloseLicenses = document.getElementById('btn-close-licenses');
  
  if (linkLicenses && modalLicenses && btnCloseLicenses) {
    linkLicenses.addEventListener('click', (e) => {
      e.preventDefault();
      modalLicenses.classList.add('active');
    });
    btnCloseLicenses.addEventListener('click', () => {
      modalLicenses.classList.remove('active');
    });
    modalLicenses.addEventListener('click', (e) => {
      if (e.target === modalLicenses) modalLicenses.classList.remove('active');
    });
  }

  // Error Logs modal
  const btnViewLogs = document.getElementById('btn-view-logs');
  const modalLogs = document.getElementById('logs-modal');
  const btnCloseLogs = document.getElementById('btn-close-logs');
  const logContentArea = document.getElementById('log-content-area');
  const btnOpenLogsFolder = document.getElementById('btn-open-logs-folder');
  const btnClearLogs = document.getElementById('btn-clear-logs');

  const loadLogs = async () => {
    if (logContentArea) {
      logContentArea.innerText = 'Loading logs...';
      try {
        const logs = await tauriInvoke('read_logs');
        logContentArea.innerText = logs;
      } catch (err) {
        logContentArea.innerText = `Error loading logs: ${err.message}`;
      }
    }
  };

  if (btnViewLogs && modalLogs && btnCloseLogs) {
    btnViewLogs.addEventListener('click', (e) => {
      e.preventDefault();
      modalLogs.classList.add('active');
      loadLogs();
    });
    btnCloseLogs.addEventListener('click', () => {
      modalLogs.classList.remove('active');
    });
    modalLogs.addEventListener('click', (e) => {
      if (e.target === modalLogs) modalLogs.classList.remove('active');
    });
  }

  if (btnOpenLogsFolder) {
    btnOpenLogsFolder.addEventListener('click', async () => {
      try {
        await tauriInvoke('open_log_directory');
      } catch (err) {
        alert(`Failed to open log folder: ${err.message}`);
      }
    });
  }

  if (btnClearLogs) {
    btnClearLogs.addEventListener('click', async () => {
      if (confirm('Are you sure you want to clear all error logs?')) {
        try {
          await tauriInvoke('clear_logs');
          loadLogs();
        } catch (err) {
          alert(`Failed to clear logs: ${err.message}`);
        }
      }
    });
  }

  // User manual click handler
  const linkUserManual = document.getElementById('link-user-manual');
  if (linkUserManual) {
    linkUserManual.addEventListener('click', (e) => {
      e.preventDefault();
      window.open('USER_MANUAL.pdf', '_blank');
    });
  }

  // PDF Capture/Download — true full-page screenshot approach
  const btnDownloadPdf = document.getElementById('btn-download-pdf');
  if (btnDownloadPdf) {
    btnDownloadPdf.addEventListener('click', async () => {
      const originalLabel = btnDownloadPdf.textContent;
      btnDownloadPdf.textContent = '⏳ Generating...';
      btnDownloadPdf.disabled = true;

      try {
        const element = document.querySelector('.dashboard-wrapper');

        // Step 1: Render the entire dashboard to a canvas (full scroll height)
        const canvas = await html2canvas(element, {
          scale: 2,
          useCORS: true,
          backgroundColor: '#0b0f19',
          scrollX: 0,
          scrollY: -window.scrollY,
          windowWidth: document.documentElement.offsetWidth,
          windowHeight: element.scrollHeight,
          height: element.scrollHeight,
          onclone: (clonedDoc) => {
            // Ensure the cloned wrapper expands to full content height
            const wrapper = clonedDoc.querySelector('.dashboard-wrapper');
            if (wrapper) {
              wrapper.style.maxWidth = 'none';
              wrapper.style.overflow = 'visible';
            }
          }
        });

        const imgData = canvas.toDataURL('image/jpeg', 0.95);

        // Step 2: Build a PDF whose page dimensions exactly match the screenshot
        // pdf coordinate system: px -> mm at 96 dpi
        const pxToMm = (px) => px * 25.4 / 96;
        const pageWidthMm  = pxToMm(canvas.width  / 2);  // /2 because scale:2
        const pageHeightMm = pxToMm(canvas.height / 2);

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({
          orientation: pageWidthMm > pageHeightMm ? 'landscape' : 'portrait',
          unit: 'mm',
          format: [pageWidthMm, pageHeightMm]
        });

        pdf.addImage(imgData, 'JPEG', 0, 0, pageWidthMm, pageHeightMm);

        const today = new Date().toISOString().split('T')[0];
        pdf.save(`SAPID_Dashboard_${today}.pdf`);

        showToast('PDF downloaded successfully!');
      } catch (err) {
        console.error('PDF Generation failed:', err);
        showToast('PDF generation failed: ' + err.message, 'error');
        tauriInvoke('append_log', { level: 'error', message: `PDF export failed: ${err.message}` }).catch(console.error);
      } finally {
        btnDownloadPdf.textContent = originalLabel;
        btnDownloadPdf.disabled = false;
      }
    });
  }

  // Configuration Backup/Export
  const btnExport = document.getElementById('btn-export-config');
  if (btnExport) {
    btnExport.addEventListener('click', () => {
      try {
        const config = {
          dashboard_sheet_url: localStorage.getItem('dashboard_sheet_url') || '',
          cached_sheet_data: JSON.parse(localStorage.getItem('cached_sheet_data') || '[]'),
          exported_at: new Date().toISOString(),
          app: "SAPID license analyser",
          version: "1.0.0"
        };
        const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `sap_license_config_${new Date().toISOString().split('T')[0]}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast('Configuration exported successfully!');
      } catch (err) {
        alert(`Export failed: ${err.message}`);
        tauriInvoke('append_log', { level: 'error', message: `Config export failed: ${err.message}` }).catch(console.error);
      }
    });
  }

  // Configuration Import
  const btnImport = document.getElementById('btn-import-config');
  const importFile = document.getElementById('import-config-file');
  if (btnImport && importFile) {
    btnImport.addEventListener('click', () => {
      importFile.click();
    });

    importFile.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (!file) return;
      if (file.size > MAX_CONFIG_FILE_BYTES) {
        alert(`Import failed: configuration file is larger than ${MAX_CONFIG_FILE_BYTES / 1024 / 1024} MB.`);
        importFile.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onload = async (evt) => {
        try {
          const config = JSON.parse(evt.target.result);
          if (!config.dashboard_sheet_url && (!config.cached_sheet_data || !Array.isArray(config.cached_sheet_data))) {
            throw new Error('Invalid file format. Must contain a dashboard URL or cached sheet data.');
          }

          if (config.dashboard_sheet_url) {
            if (!isAllowedGoogleDataUrl(config.dashboard_sheet_url)) {
              throw new Error('Imported dashboard URL must be a supported HTTPS Google Sheets or Apps Script /exec link.');
            }
            localStorage.setItem('dashboard_sheet_url', config.dashboard_sheet_url);
            const urlInput = document.getElementById('sheet-url-input');
            if (urlInput) urlInput.value = config.dashboard_sheet_url;
          }
          if (config.cached_sheet_data) {
            rawData = validateImportedRows(config.cached_sheet_data);
            cacheSheetData(rawData);
          }

          alert('Configuration imported successfully! Refreshing dashboard...');
          initDashboard();
        } catch (err) {
          alert(`Import failed: ${err.message}`);
          tauriInvoke('append_log', { level: 'error', message: `Config import failed: ${err.message}` }).catch(console.error);
        } finally {
          importFile.value = '';
        }
      };
      reader.readAsText(file);
    });
  }
}

// Handle Excel Upload
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  if (file.size > MAX_EXCEL_FILE_BYTES) {
    alert(`The uploaded Excel file is larger than ${MAX_EXCEL_FILE_BYTES / 1024 / 1024} MB.`);
    e.target.value = '';
    return;
  }

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = getBestSheetName(workbook);
      const worksheet = workbook.Sheets[sheetName];
      const parsedData = XLSX.utils.sheet_to_json(worksheet);

      if (parsedData.length === 0) {
        alert("The uploaded excel sheet is empty!");
        return;
      }
      validateImportedRows(parsedData);

      rawData = parsedData;
      
      // Show reset button
      document.getElementById('reset-btn').classList.remove('hidden');

      initDashboard();
    } catch (err) {
      console.error(err);
      tauriInvoke('append_log', { level: 'error', message: `Local Excel upload parse error: ${err.message}` }).catch(console.error);
      alert("Error parsing excel file: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Aesthetics and modal header updates are statically handled in HTML now

// Initialize components (options, dropdowns)
function initDashboard() {
  // Populate filter options dynamically
  const licenses = new Set();
  const depts = new Set();

  rawData.forEach(row => {
    const license = getRowLicense(row);
    const dept = getRowDepartment(row);
    if (license) licenses.add(license);
    if (dept) depts.add(dept);
  });

  // Populate License select
  const licenseSelect = document.getElementById('filter-license');
  if (licenseSelect) {
    licenseSelect.textContent = '';
    appendSafeOption(licenseSelect, 'ALL', 'All Licenses');
    Array.from(licenses).sort().forEach(license => {
      appendSafeOption(licenseSelect, license, license);
    });
  }

  // Populate Dept select
  const deptSelect = document.getElementById('filter-dept');
  if (deptSelect) {
    deptSelect.textContent = '';
    appendSafeOption(deptSelect, 'ALL', 'All Departments');
    Array.from(depts).sort().forEach(dept => {
      appendSafeOption(deptSelect, dept, dept);
    });
  }

  // Populate SAP User Group Explorer filter dynamically
  const explorerLicenseSelect = document.getElementById('filter-explorer-license');
  if (explorerLicenseSelect) {
    explorerLicenseSelect.textContent = '';
    appendSafeOption(explorerLicenseSelect, 'ALL', 'All Licenses');
    Array.from(licenses).sort().forEach(license => {
      appendSafeOption(explorerLicenseSelect, license, license);
    });
    explorerLicenseSelect.value = 'ALL';
  }

  // Reset Filters to ALL
  currentFilters = { license: 'ALL', dept: 'ALL' };
  if (licenseSelect) licenseSelect.value = 'ALL';
  if (deptSelect) deptSelect.value = 'ALL';

  updateDashboard();
}

// Update dashboard metrics and charts
function updateDashboard() {
  // 1. Filter Raw Data
  const filteredRows = rawData.filter(row => {
    const matchLicense = currentFilters.license === 'ALL' || getRowLicense(row) === currentFilters.license;
    const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
    return matchLicense && matchDept;
  });

  // 2. Map to employees list (logon recency)
  employeesList = filteredRows.map(row => {
    return {
      id: getRowSAPID(row),
      name: getRowUserName(row),
      license: getRowLicense(row),
      dept: getRowDepartment(row)
    };
  });

  // 3. Render KPIs
  renderKPIs(filteredRows, employeesList);

  // 4. Render Main Charts
  renderLicenseChart(filteredRows);
  renderDeptChart(filteredRows);
  renderPlantIdChart(filteredRows);

  // Render User Group Explorer
  renderLeadershipChart();

  // Render License Distribution Progress on the right column
  renderRightColumnAnalytics(filteredRows);
}

// Render KPI Summary Cards
function renderKPIs(rows, emps) {
  // Total Users (Allocated)
  document.getElementById('kpi-total-users').innerText = rows.length;

  // Compute total purchased licenses limit
  let totalPurchased = 0;
  const uniqueLicenses = ['AX', 'AY', 'FX', 'HC', 'HD'];
  uniqueLicenses.forEach(lic => {
    const stored = localStorage.getItem(`sap_license_purchased_${lic}`);
    const limit = stored ? parseInt(stored) : (DEFAULT_LICENSE_PURCHASED[lic] || 150);
    totalPurchased += limit;
  });
  document.getElementById('kpi-total-purchased').innerText = totalPurchased;

  // Available Licenses: Total Purchased - Total Users (Allocated)
  const totalAllocated = rawData.length;
  const available = Math.max(0, totalPurchased - totalAllocated);
  document.getElementById('kpi-available-licenses').innerText = available;

  // Find Top License
  const licenseCounts = {};
  rows.forEach(row => {
    const license = getRowLicense(row);
    if (license) {
      licenseCounts[license] = (licenseCounts[license] || 0) + 1;
    }
  });

  let topLicense = '-';
  let maxLicenseCount = 0;
  Object.keys(licenseCounts).forEach(l => {
    if (licenseCounts[l] > maxLicenseCount) {
      maxLicenseCount = licenseCounts[l];
      topLicense = l;
    }
  });
  
  if (topLicense !== '-') {
    document.getElementById('kpi-top-license').innerText = `${topLicense} (${maxLicenseCount} users)`;
  } else {
    document.getElementById('kpi-top-license').innerText = '-';
  }
}

// Render First Graph: License Distribution
function renderLicenseChart(rows) {
  if (typeof Chart === 'undefined') return;

  const licenseCounts = {};
  rows.forEach(row => {
    const license = getRowLicense(row);
    licenseCounts[license] = (licenseCounts[license] || 0) + 1;
  });

  const labels = Object.keys(licenseCounts).sort();
  const dataValues = labels.map(label => licenseCounts[label]);
  const colors = labels.map(label => LICENSE_COLORS[label] || LICENSE_COLORS['Other']);

  if (licenseChartInstance) {
    licenseChartInstance.destroy();
  }

  const ctx = document.getElementById('chart-license-distribution').getContext('2d');
  licenseChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Users count',
        data: dataValues,
        backgroundColor: colors,
        borderRadius: 6,
        borderWidth: 0,
        maxBarThickness: 45
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => ` Assigned Users: ${context.raw}`
          }
        }
      },
      scales: {
        x: {
          grid: { display: false, drawBorder: false },
          ticks: { font: { weight: '600' } }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)', drawBorder: false },
          beginAtZero: true,
          ticks: {
            precision: 0
          }
        }
      }
    }
  });
}

// Render Second Graph of Segment 1: Department distribution
function renderDeptChart(rows) {
  if (typeof Chart === 'undefined') return;

  const deptCounts = {};
  rows.forEach(row => {
    const dept = getRowDepartment(row);
    deptCounts[dept] = (deptCounts[dept] || 0) + 1;
  });

  const labels = Object.keys(deptCounts).sort((a,b) => deptCounts[b] - deptCounts[a]);
  const dataValues = labels.map(label => deptCounts[label]);
  const colors = labels.map(label => DEPT_COLORS[label] || DEPT_COLORS['Other']);

  if (deptChartInstance) {
    deptChartInstance.destroy();
  }

  const ctx = document.getElementById('chart-dept-distribution').getContext('2d');
  deptChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Users count',
        data: dataValues,
        backgroundColor: colors,
        borderRadius: 6,
        borderWidth: 0,
        maxBarThickness: 45
      }]
    },
    options: {
      indexAxis: 'y', // Horizontal bars
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => ` Users: ${context.raw}`
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)', drawBorder: false },
          beginAtZero: true,
          ticks: {
            precision: 0
          }
        },
        y: {
          grid: { display: false, drawBorder: false },
          ticks: { font: { weight: '600' } }
        }
      }
    }
  });
}

// Render Third Graph: SAP ID Prefix Distribution (First 4 Characters/Numbers)
// Render Third Graph: Plant ID Distribution (First 4 Characters/Numbers of SAP ID)
function renderPlantIdChart(rows) {
  if (typeof Chart === 'undefined') return;

  const plantIdCounts = {};
  rows.forEach(row => {
    const sapId = getRowSAPID(row);
    const prefix = sapId && sapId.length >= 4 ? sapId.substring(0, 4) : (sapId || 'Unknown');
    plantIdCounts[prefix] = (plantIdCounts[prefix] || 0) + 1;
  });

  const labels = Object.keys(plantIdCounts).sort();
  const dataValues = labels.map(label => plantIdCounts[label]);
  
  // Vibrant gradients/colors for prefixes
  const palette = ['#0ea5e9', '#6366f1', '#a855f7', '#ec4899', '#14b8a6', '#f59e0b', '#10b981'];
  const colors = labels.map((_, i) => palette[i % palette.length]);

  if (plantIdChartInstance) {
    plantIdChartInstance.destroy();
  }

  const canvas = document.getElementById('chart-plantid-distribution');
  if (!canvas) return;

  const ctx = canvas.getContext('2d');
  plantIdChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Users count',
        data: dataValues,
        backgroundColor: colors,
        borderRadius: 6,
        borderWidth: 0,
        maxBarThickness: 45
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => ` Users: ${context.raw}`
          }
        }
      },
      onClick: (e, elements) => {
        if (elements && elements.length > 0) {
          const index = elements[0].index;
          const clickedPrefix = labels[index];
          showPlantIdDetailsModal(clickedPrefix);
        }
      },
      scales: {
        x: {
          grid: { display: false, drawBorder: false },
          ticks: { font: { weight: '600' } }
        },
        y: {
          grid: { color: 'rgba(255, 255, 255, 0.05)', drawBorder: false },
          beginAtZero: true,
          ticks: {
            precision: 0
          }
        }
      }
    }
  });
}

function showPlantIdDetailsModal(prefix) {
  // Find matching users from the filtered/current state of rawData
  const matchingRows = rawData.filter(row => {
    const matchLicense = currentFilters.license === 'ALL' || getRowLicense(row) === currentFilters.license;
    const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
    if (!matchLicense || !matchDept) return false;

    const sapId = getRowSAPID(row);
    const rowPrefix = sapId && sapId.length >= 4 ? sapId.substring(0, 4) : (sapId || 'Unknown');
    return rowPrefix === prefix;
  });

  const modalTitle = document.getElementById('modal-title');
  const modalSubtitle = document.getElementById('modal-subtitle');
  const modalTableBody = document.getElementById('modal-table-body');

  modalTitle.innerText = `Plant ID Group: ${prefix}`;
  modalSubtitle.innerText = `Users matching Plant ID (sorted by Department)`;

  modalTableBody.innerHTML = '';
  
  // Sort by department, then name
  matchingRows.sort((a, b) => {
    const deptA = getRowDepartment(a);
    const deptB = getRowDepartment(b);
    if (deptA !== deptB) return deptA.localeCompare(deptB);
    const nameA = getRowUserName(a);
    const nameB = getRowUserName(b);
    return nameA.localeCompare(nameB);
  });

  matchingRows.forEach(row => {
    const trModal = document.createElement('tr');
    trModal.innerHTML = `
      <td style="padding: 0.8rem 1rem; font-weight: 600;">${escapeHTML(getRowSAPID(row))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowUserName(row))}</td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-license">${escapeHTML(getRowLicense(row))}</span></td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-dept">${escapeHTML(getRowDepartment(row))}</span></td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowFunction(row))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowLastLogon(row))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowGroup(row) || '-')}</td>
    `;
    modalTableBody.appendChild(trModal);
  });

  document.getElementById('drilldown-modal').classList.add('active');
}

function closeModal() {
  document.getElementById('drilldown-modal').classList.remove('active');
}

// Render User Group Explorer: displays groups as buttons
function renderLeadershipChart() {
  const selectedCat = document.getElementById('filter-explorer-license').value;

  const filteredRows = rawData.filter(row => {
    const matchLicense = currentFilters.license === 'ALL' || getRowLicense(row) === currentFilters.license;
    const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
    const matchCat = selectedCat === 'ALL' || getRowLicense(row) === selectedCat;
    return matchLicense && matchDept && matchCat;
  });

  const groupData = {};
  filteredRows.forEach(row => {
    const groupName = getRowFunction(row);
    if (!groupName) return;
    const sapId = getRowSAPID(row);
    const name = getRowUserName(row);
    const license = getRowLicense(row);
    const dept = getRowDepartment(row);

    if (!groupData[groupName]) {
      groupData[groupName] = {
        name: groupName,
        employeeMap: {}
      };
    }

    if (!groupData[groupName].employeeMap[sapId]) {
      groupData[groupName].employeeMap[sapId] = {
        id: sapId,
        name: name,
        corp: license,
        dept: dept
      };
    }
  });

  const groups = Object.keys(groupData).sort();

  const buttonsContainer = document.getElementById('group-buttons-container');
  buttonsContainer.innerHTML = '';

  const prevBtn = document.getElementById('btn-prev-groups');
  const nextBtn = document.getElementById('btn-next-groups');
  const pageInfo = document.getElementById('group-page-info');

  if (groups.length === 0) {
    buttonsContainer.innerHTML = `
      <div style="grid-column: 1/-1; text-align: center; color: var(--text-muted); padding: 2rem 0;">
        No groups found for this license
      </div>
    `;
    document.getElementById('selected-group-title').innerText = 'Assigned Users';
    document.getElementById('selected-group-subtitle').innerText = 'Select a group to view details';
    document.getElementById('explorer-users-body').innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--text-muted); padding: 2rem 0; font-size: 0.85rem;">
          No groups available
        </td>
      </tr>
    `;
    if (prevBtn) prevBtn.disabled = true;
    if (nextBtn) nextBtn.disabled = true;
    if (pageInfo) pageInfo.innerText = 'Page 1 of 1';
    return;
  }

  // Calculate pages for groups
  const totalGroupPages = Math.ceil(groups.length / groupPageSize) || 1;
  if (groupCurrentPage > totalGroupPages) {
    groupCurrentPage = totalGroupPages;
  }
  if (groupCurrentPage < 1) {
    groupCurrentPage = 1;
  }

  // Update button states and text
  if (prevBtn) prevBtn.disabled = groupCurrentPage === 1;
  if (nextBtn) nextBtn.disabled = groupCurrentPage === totalGroupPages;
  if (pageInfo) pageInfo.innerText = `Page ${groupCurrentPage} of ${totalGroupPages}`;

  const startIndex = (groupCurrentPage - 1) * groupPageSize;
  const endIndex = Math.min(startIndex + groupPageSize, groups.length);
  const pageGroups = groups.slice(startIndex, endIndex);

  // Set selected group to first group of current page if not on the page or not valid
  if (!selectedExplorerGroup || !groups.includes(selectedExplorerGroup) || !pageGroups.includes(selectedExplorerGroup)) {
    selectedExplorerGroup = pageGroups[0];
  }

  pageGroups.forEach((groupName, index) => {
    const count = Object.keys(groupData[groupName].employeeMap).length;
    const btn = document.createElement('div');
    btn.className = `group-card-btn ${selectedExplorerGroup === groupName ? 'active' : ''}`;
    btn.innerHTML = `
      <span class="group-btn-title">${escapeHTML(groupName)}</span>
      <span class="group-btn-value" id="explorer-group-val-${index}">0</span>
    `;

    btn.addEventListener('click', () => {
      const allBtns = buttonsContainer.querySelectorAll('.group-card-btn');
      allBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      selectedExplorerGroup = groupName;
      renderExplorerEmployeesList(groupData[groupName], selectedCat);
    });

    buttonsContainer.appendChild(btn);

    const valEl = btn.querySelector('.group-btn-value');
    animateCountValue(valEl, 0, count, 600 + index * 100);
  });

  renderExplorerEmployeesList(groupData[selectedExplorerGroup], selectedCat);
}

// Render right side users list for selected group
function renderExplorerEmployeesList(groupInfo, selectedCat) {
  const titleEl = document.getElementById('selected-group-title');
  const subtitleEl = document.getElementById('selected-group-subtitle');
  const tbody = document.getElementById('explorer-users-body');

  if (!groupInfo) {
    titleEl.innerText = 'Assigned Users';
    subtitleEl.innerText = 'Select a group to view details';
    tbody.innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--text-muted); padding: 2rem 0; font-size: 0.85rem;">
          No group selected
        </td>
      </tr>
    `;
    return;
  }

  const groupName = groupInfo.name;
  const employeeMap = groupInfo.employeeMap;
  const employees = Object.values(employeeMap);

  titleEl.innerText = groupName;
  subtitleEl.innerText = `Unique Assigned Persons: ${employees.length} (Click user row to view details)`;

  tbody.innerHTML = '';
  if (employees.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--text-muted); padding: 2rem 0; font-size: 0.85rem;">
          No users assigned
        </td>
      </tr>
    `;
    return;
  }

  employees.sort((a, b) => a.dept.localeCompare(b.dept));

  employees.forEach(emp => {
    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.innerHTML = `
      <td><span style="font-weight: 600;">${escapeHTML(emp.name)}</span> <span style="color: var(--text-secondary); font-size: 0.75rem; margin-left: 0.4rem;">(${escapeHTML(emp.id)})</span></td>
      <td><span class="badge-license">${escapeHTML(emp.corp)}</span></td>
      <td><span class="badge-dept">${escapeHTML(emp.dept)}</span></td>
    `;

    tr.addEventListener('click', () => {
      const matchingRows = rawData.filter(row => {
        const sapId = getRowSAPID(row);
        if (sapId !== emp.id) return false;

        const matchLicense = currentFilters.license === 'ALL' || getRowLicense(row) === currentFilters.license;
        const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
        const matchCat = selectedCat === 'ALL' || getRowLicense(row) === selectedCat;
        const matchGroup = getRowFunction(row) === groupName;
        return matchLicense && matchDept && matchCat && matchGroup;
      });

      matchingRows.sort((a, b) => {
        const deptA = getRowDepartment(a);
        const deptB = getRowDepartment(b);
        if (deptA !== deptB) return deptA.localeCompare(deptB);
        const nameA = getRowUserName(a);
        const nameB = getRowUserName(b);
        return nameA.localeCompare(nameB);
      });

      const modalTitle = document.getElementById('modal-title');
      const modalSubtitle = document.getElementById('modal-subtitle');
      const modalTableBody = document.getElementById('modal-table-body');

      modalTitle.innerText = 'SAP User Group Assignment';
      modalSubtitle.innerText = `User: ${emp.name} | Group: ${groupName} (sorted by Department)`;

      modalTableBody.innerHTML = '';
      matchingRows.forEach(row => {
        const trModal = document.createElement('tr');
        trModal.innerHTML = `
          <td style="padding: 0.8rem 1rem; font-weight: 600;">${escapeHTML(getRowSAPID(row))}</td>
          <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowUserName(row))}</td>
          <td style="padding: 0.8rem 1rem;"><span class="badge-license">${escapeHTML(getRowLicense(row))}</span></td>
          <td style="padding: 0.8rem 1rem;"><span class="badge-dept">${escapeHTML(getRowDepartment(row))}</span></td>
          <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowFunction(row))}</td>
          <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowLastLogon(row))}</td>
          <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowGroup(row) || '-')}</td>
        `;
        modalTableBody.appendChild(trModal);
      });

      document.getElementById('drilldown-modal').classList.add('active');
    });

    tbody.appendChild(tr);
  });
}

// Helper to animate numbers count-up effect
function animateCountValue(elementId, start, end, duration = 800) {
  const obj = typeof elementId === 'string' ? document.getElementById(elementId) : elementId;
  if (!obj) return;
  
  let startTimestamp = null;
  const step = (timestamp) => {
    if (!startTimestamp) startTimestamp = timestamp;
    const progress = Math.min((timestamp - startTimestamp) / duration, 1);
    const current = Math.floor(progress * (end - start) + start);
    
    obj.innerHTML = current;
    
    if (progress < 1) {
      window.requestAnimationFrame(step);
    }
  };
  window.requestAnimationFrame(step);
}

// Render target and license progress in the right column
function renderRightColumnAnalytics(filteredRows) {
  // Get all unique licenses from the entire rawData so all categories are always visible
  const allLicenses = Array.from(new Set(rawData.map(getRowLicense))).filter(Boolean).sort();

  // Calculate allocation counts per license type in filteredRows
  const licenseCounts = {};
  allLicenses.forEach(lic => {
    licenseCounts[lic] = 0;
  });
  filteredRows.forEach(row => {
    const lic = getRowLicense(row);
    if (lic && licenseCounts[lic] !== undefined) {
      licenseCounts[lic]++;
    }
  });

  const container = document.getElementById('category-progress-container');
  if (container) {
    if (allLicenses.length === 0) {
      container.innerHTML = `
        <div style="text-align: center; color: var(--text-muted); padding: 1rem 0; font-size: 0.85rem;">
          No license data
        </div>
      `;
      return;
    }

    // Remove placeholder message if it was displayed
    if (container.querySelector('div') && container.querySelector('div').textContent.includes('No license data')) {
      container.innerHTML = '';
    }

    allLicenses.forEach((lic) => {
      const count = licenseCounts[lic];
      const stored = localStorage.getItem(`sap_license_purchased_${lic}`);
      const limit = stored ? parseInt(stored) : (DEFAULT_LICENSE_PURCHASED[lic] || 150);
      const percentage = limit > 0 ? Math.min(100, Math.round((count / limit) * 100)) : 0;

      let card = document.getElementById(`progress-card-${lic}`);
      if (!card) {
        card = document.createElement('div');
        card.className = 'category-progress-card';
        card.id = `progress-card-${lic}`;
        card.innerHTML = `
          <div class="cat-progress-header">
            <div class="cat-progress-meta">
              <span class="cat-progress-name" style="color: ${LICENSE_COLORS[lic] || LICENSE_COLORS['Other']};">${escapeHTML(lic)}</span>
              <span class="cat-progress-details">
                <span class="allocated-count" id="cat-allocated-${lic}" style="font-weight: 700; color: var(--text-primary);">${count}</span> / 
                <span class="limit-count" id="cat-limit-val-${lic}">${limit}</span> Assigned (<span class="percentage-val" id="cat-pct-${lic}">${percentage}</span>%)
              </span>
            </div>
            <div class="cat-progress-input-wrapper" data-html2canvas-ignore="true">
              <label for="limit-input-${lic}">Limit:</label>
              <input type="number" id="limit-input-${lic}" class="limit-input" data-license="${lic}" value="${limit}" min="0">
            </div>
          </div>
          <div class="cat-progress-remaining-line" style="font-size: 0.8rem; color: var(--text-secondary); margin-bottom: 0.35rem; display: flex; justify-content: space-between;">
            <span>Remaining Count:</span>
            <span id="cat-remaining-${lic}" style="font-weight: 600; color: ${limit - count > 0 ? '#10b981' : '#f43f5e'};">${limit - count}</span>
          </div>
          <div class="cat-progress-bar-track">
            <div class="cat-progress-bar-fill" id="progress-fill-${lic}" style="width: ${percentage}%; background-color: ${LICENSE_COLORS[lic] || LICENSE_COLORS['Other']};"></div>
          </div>
        `;
        container.appendChild(card);
      } else {
        // Update in-place to avoid visual resets and preserve keyboard focus
        const allocatedEl = document.getElementById(`cat-allocated-${lic}`);
        if (allocatedEl) allocatedEl.textContent = count;

        const limitEl = document.getElementById(`cat-limit-val-${lic}`);
        if (limitEl) limitEl.textContent = limit;

        const pctEl = document.getElementById(`cat-pct-${lic}`);
        if (pctEl) pctEl.textContent = percentage;

        const inputEl = document.getElementById(`limit-input-${lic}`);
        if (inputEl && document.activeElement !== inputEl) {
          inputEl.value = limit;
        }

        const remainingEl = document.getElementById(`cat-remaining-${lic}`);
        if (remainingEl) {
          remainingEl.textContent = limit - count;
          remainingEl.style.color = limit - count > 0 ? '#10b981' : '#f43f5e';
        }

        const fillEl = document.getElementById(`progress-fill-${lic}`);
        if (fillEl) {
          fillEl.style.width = `${percentage}%`;
        }
      }
    });

    // Remove any legacy cards that are no longer valid
    const existingCards = container.querySelectorAll('.category-progress-card');
    existingCards.forEach(c => {
      const id = c.id || '';
      const lic = id.replace('progress-card-', '');
      if (lic && !allLicenses.includes(lic)) {
        c.remove();
      }
    });
  }
}
