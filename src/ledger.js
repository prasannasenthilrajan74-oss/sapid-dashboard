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

// Ledger Application State
let rawData = [];
let employeesList = []; // Aggregated employees list
let currentFilters = {
  license: 'ALL',
  dept: 'ALL',
  search: ''
};

// Table State
let tableSortColumn = 'sapid'; // Default sort by SAPID
let tableSortDirection = 'asc';
let tableCurrentPage = 1;
let tablePageSize = 10;

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

// Group/User group accessor
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

  if (urlStr.includes('/pub')) {
    if (urlStr.includes('output=csv')) {
      return urlStr;
    }
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

// Fetch spreadsheet data via secure backend proxy (adapted for Tauri)
async function fetchSpreadsheetData(sheetUrl) {
  try {
    // Detect URL type
    const isDomainScript = sheetUrl.includes('script.google.com/a/macros/');
    const isAppsScript = isDomainScript || sheetUrl.includes('script.google.com/macros/s/');

    let parsedData = [];

    if (isDomainScript) {
      // ── Domain-restricted Apps Script (/a/macros/) ────────────────────────
      // Step 1: Validate against whitelist via Tauri command
      const validateResult = await tauriInvoke('validate_url', { url: sheetUrl });
      if (!validateResult.success) {
        throw new Error(validateResult.message || '🔒 Access denied: This link is not authorized.');
      }

      // Step 2: Fetch directly from the browser — the user's Google session cookie
      //         is automatically sent via the WebView.
      const directRes = await fetch(sheetUrl, {
        credentials: 'include',
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

      // Check if it's a 2D array
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
      throw new Error('The parsed data is empty.');
    }

    rawData = parsedData;
    cacheSheetData(parsedData);
    initLedger();
  } catch (err) {
    console.error('Error fetching sheet data:', err);
    tauriInvoke('append_log', { level: 'error', message: `Google Sheet fetch error: ${err.message}` }).catch(console.error);

    // Fallback to local storage cache
    const cachedDataStr = localStorage.getItem('cached_sheet_data');
    if (cachedDataStr) {
      try {
        rawData = JSON.parse(cachedDataStr);
        initLedger();
        return;
      } catch (parseErr) {
        console.error('Failed to parse cached data:', parseErr);
      }
    }

    const tbody = document.getElementById('employee-table-body');
    if (tbody) {
      tbody.innerHTML = `
        <tr>
          <td colspan="6" style="text-align: center; color: var(--accent-rose); padding: 3rem 0; font-weight: 500;">
            ❌ Failed to fetch spreadsheet data: ${err.message}
          </td>
        </tr>
      `;
    }
  }
}

// Aesthetics and modal header updates are statically handled in markup now

// Initialize options
function initLedger() {
  const licenses = new Set();
  const depts = new Set();

  rawData.forEach(row => {
    const license = getRowLicense(row);
    const dept = getRowDepartment(row);
    if (license) licenses.add(license);
    if (dept) depts.add(dept);
  });

  // License Select
  const licenseSelect = document.getElementById('filter-license');
  if (licenseSelect) {
    licenseSelect.textContent = '';
    appendSafeOption(licenseSelect, 'ALL', 'All Licenses');
    Array.from(licenses).sort().forEach(license => {
      appendSafeOption(licenseSelect, license, license);
    });
  }

  // Dept Select
  const deptSelect = document.getElementById('filter-dept');
  if (deptSelect) {
    deptSelect.textContent = '';
    appendSafeOption(deptSelect, 'ALL', 'All Departments');
    Array.from(depts).sort().forEach(dept => {
      appendSafeOption(deptSelect, dept, dept);
    });
  }

  // Reset Filters UI
  if (licenseSelect) licenseSelect.value = 'ALL';
  if (deptSelect) deptSelect.value = 'ALL';
  const searchInput = document.getElementById('search-employee');
  if (searchInput) searchInput.value = '';
  const pageSizeSelect = document.getElementById('table-page-size');
  if (pageSizeSelect) pageSizeSelect.value = '10';

  currentFilters = { license: 'ALL', dept: 'ALL', search: '' };
  tableCurrentPage = 1;
  tablePageSize = 10;
  tableSortColumn = 'sapid';
  tableSortDirection = 'asc';

  // Update table headers UI classes
  const headers = ['th-sapid', 'th-name', 'th-license', 'th-department', 'th-group', 'th-lastlogon'];
  headers.forEach(hId => {
    const el = document.getElementById(hId);
    if (el) {
      el.classList.remove('sorted-asc', 'sorted-desc');
    }
  });
  const defaultHeader = document.getElementById('th-sapid');
  if (defaultHeader) {
    defaultHeader.classList.add('sorted-asc');
  }

  updateLedger();
}

// Filter and map raw data to employee list
function updateLedger() {
  const filteredRows = rawData.filter(row => {
    const matchLicense = currentFilters.license === 'ALL' || getRowLicense(row) === currentFilters.license;
    const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
    return matchLicense && matchDept;
  });

  employeesList = filteredRows.map(row => {
    return {
      sapid: getRowSAPID(row),
      name: getRowUserName(row),
      license: getRowLicense(row),
      department: getRowDepartment(row),
      group: getRowGroup(row),
      lastlogon: getRowLastLogon(row)
    };
  });

  renderEmployeeTable();
}

// Get filtered, sorted list and render to DOM
function renderEmployeeTable() {
  let list = [...employeesList];

  // Apply Search query
  if (currentFilters.search) {
    const term = currentFilters.search;
    list = list.filter(emp => {
      return emp.name.toLowerCase().includes(term) || emp.sapid.toLowerCase().includes(term);
    });
  }

  // Sort list
  list.sort((a, b) => {
    const valA = a[tableSortColumn] || '';
    const valB = b[tableSortColumn] || '';

    if (typeof valA === 'string') {
      return tableSortDirection === 'asc'
        ? valA.localeCompare(valB)
        : valB.localeCompare(valA);
    } else {
      return tableSortDirection === 'asc'
        ? valA - valB
        : valB - valA;
    }
  });

  const totalItems = list.length;
  const totalPages = Math.ceil(totalItems / tablePageSize) || 1;

  if (tableCurrentPage > totalPages) {
    tableCurrentPage = totalPages;
  }
  if (tableCurrentPage < 1) {
    tableCurrentPage = 1;
  }

  const startIdx = (tableCurrentPage - 1) * tablePageSize;
  const endIdx = Math.min(startIdx + tablePageSize, totalItems);
  const pageItems = list.slice(startIdx, endIdx);

  const tbody = document.getElementById('employee-table-body');
  tbody.innerHTML = '';

  if (pageItems.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" style="text-align: center; color: var(--text-muted); padding: 3rem 0;">
          No matching employees found
        </td>
      </tr>
    `;
    document.getElementById('table-info').innerText = 'Showing 0 to 0 of 0 employees';
    document.getElementById('btn-prev-page').disabled = true;
    document.getElementById('btn-next-page').disabled = true;
    return;
  }

  pageItems.forEach(emp => {
    const tr = document.createElement('tr');
    tr.className = 'ledger-row';
    tr.innerHTML = `
      <td style="font-weight: 600;">${escapeHTML(emp.sapid)}</td>
      <td>${escapeHTML(emp.name)}</td>
      <td><span class="badge-license">${escapeHTML(emp.license)}</span></td>
      <td><span class="badge-dept">${escapeHTML(emp.department)}</span></td>
      <td style="text-align: left;">${escapeHTML(emp.group)}</td>
      <td style="text-align: left; font-weight: 600; color: #38bdf8;">${escapeHTML(emp.lastlogon)}</td>
    `;

    tr.addEventListener('click', () => {
      showEmployeeDetailModal(emp.sapid, emp.name);
    });

    tbody.appendChild(tr);
  });

  document.getElementById('table-info').innerText = `Showing ${startIdx + 1} to ${endIdx} of ${totalItems} employees`;
  document.getElementById('btn-prev-page').disabled = tableCurrentPage === 1;
  document.getElementById('btn-next-page').disabled = tableCurrentPage === totalPages;
}

// Show detailed user profile modal for selected user
function showEmployeeDetailModal(empId, empName) {
  const userRow = rawData.find(row => getRowSAPID(row) === empId);

  const modalTitle = document.getElementById('modal-title');
  const modalSubtitle = document.getElementById('modal-subtitle');
  const modalTableBody = document.getElementById('modal-table-body');

  modalTitle.innerText = `${empName} (${empId})`;
  modalSubtitle.innerText = 'SAP account profile details';

  modalTableBody.innerHTML = '';
  if (userRow) {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="padding: 0.8rem 1rem; font-weight: 600;">${escapeHTML(getRowSAPID(userRow))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowUserName(userRow))}</td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-license">${escapeHTML(getRowLicense(userRow))}</span></td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-dept">${escapeHTML(getRowDepartment(userRow))}</span></td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowFunction(userRow))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowLastLogon(userRow))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowGroup(userRow) || '-')}</td>
    `;
    modalTableBody.appendChild(tr);
  }

  document.getElementById('drilldown-modal').classList.add('active');
}

function closeModal() {
  document.getElementById('drilldown-modal').classList.remove('active');
}

// Setup Event Listeners
function setupLedgerListeners() {
  const searchBtn = document.getElementById('search-btn');
  const searchInput = document.getElementById('search-employee');

  const performSearch = () => {
    currentFilters.search = searchInput.value.toLowerCase().trim();
    tableCurrentPage = 1;
    renderEmployeeTable();
  };

  if (searchBtn && searchInput) {
    searchBtn.addEventListener('click', performSearch);
    searchInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        performSearch();
      }
    });
  }

  // Clear button
  document.getElementById('reset-filter-btn').addEventListener('click', () => {
    searchInput.value = '';
    document.getElementById('filter-license').value = 'ALL';
    document.getElementById('filter-dept').value = 'ALL';

    currentFilters = { license: 'ALL', dept: 'ALL', search: '' };
    tableCurrentPage = 1;
    updateLedger();
  });

  // Dropdown filter changes
  document.getElementById('filter-license').addEventListener('change', (e) => {
    currentFilters.license = e.target.value;
    tableCurrentPage = 1;
    updateLedger();
  });

  // Dropdown filter changes
  document.getElementById('filter-dept').addEventListener('change', (e) => {
    currentFilters.dept = e.target.value;
    tableCurrentPage = 1;
    updateLedger();
  });

  // Page Size selector
  document.getElementById('table-page-size').addEventListener('change', (e) => {
    tablePageSize = parseInt(e.target.value);
    tableCurrentPage = 1;
    renderEmployeeTable();
  });

  // Table Headers Sorting
  const headers = ['th-sapid', 'th-name', 'th-license', 'th-department', 'th-group', 'th-lastlogon'];
  headers.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('click', () => {
        const col = el.getAttribute('data-sort');
        if (tableSortColumn === col) {
          tableSortDirection = tableSortDirection === 'asc' ? 'desc' : 'asc';
        } else {
          tableSortColumn = col;
          tableSortDirection = 'asc';
        }

        headers.forEach(hId => {
          const hEl = document.getElementById(hId);
          if (hEl) {
            hEl.classList.remove('sorted-asc', 'sorted-desc');
          }
        });
        el.classList.add(tableSortDirection === 'asc' ? 'sorted-asc' : 'sorted-desc');

        renderEmployeeTable();
      });
    }
  });

  // Pagination buttons
  document.getElementById('btn-prev-page').addEventListener('click', () => {
    if (tableCurrentPage > 1) {
      tableCurrentPage--;
      renderEmployeeTable();
    }
  });

  document.getElementById('btn-next-page').addEventListener('click', () => {
    const totalPages = Math.ceil(employeesList.filter(emp => {
      if (!currentFilters.search) return true;
      return emp.name.toLowerCase().includes(currentFilters.search) || emp.sapid.toLowerCase().includes(currentFilters.search);
    }).length / tablePageSize);

    if (tableCurrentPage < totalPages) {
      tableCurrentPage++;
      renderEmployeeTable();
    }
  });

  // Modal Closing events
  document.getElementById('modal-close-btn').addEventListener('click', closeModal);
  document.getElementById('drilldown-modal').addEventListener('click', (e) => {
    if (e.target.id === 'drilldown-modal') closeModal();
  });

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

        // Step 1: Render the entire ledger page to a canvas (full scroll height)
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
            const wrapper = clonedDoc.querySelector('.dashboard-wrapper');
            if (wrapper) {
              wrapper.style.maxWidth = 'none';
              wrapper.style.overflow = 'visible';
            }
          }
        });

        const imgData = canvas.toDataURL('image/jpeg', 0.95);

        // Step 2: Build a PDF page sized exactly to the screenshot
        const pxToMm = (px) => px * 25.4 / 96;
        const pageWidthMm = pxToMm(canvas.width / 2);
        const pageHeightMm = pxToMm(canvas.height / 2);

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({
          orientation: pageWidthMm > pageHeightMm ? 'landscape' : 'portrait',
          unit: 'mm',
          format: [pageWidthMm, pageHeightMm]
        });

        pdf.addImage(imgData, 'JPEG', 0, 0, pageWidthMm, pageHeightMm);

        const today = new Date().toISOString().split('T')[0];
        pdf.save(`SAPID_Ledger_${today}.pdf`);

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
          }
          if (config.cached_sheet_data) {
            rawData = validateImportedRows(config.cached_sheet_data);
            cacheSheetData(rawData);
          }

          alert('Configuration imported successfully! Refreshing ledger...');
          initLedger();
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

// DOM content loaded
window.addEventListener('DOMContentLoaded', () => {
  setupLedgerListeners();

  // Clear legacy default URL on startup
  localStorage.removeItem('dashboard_sheet_url');
  localStorage.removeItem('cached_sheet_data');
  let activeUrl = null;

  if (activeUrl) {
    fetchSpreadsheetData(activeUrl);
  } else {
    const cachedDataStr = localStorage.getItem('cached_sheet_data');
    if (cachedDataStr) {
      try {
        rawData = JSON.parse(cachedDataStr);
        if (rawData.length > 0) {
          initLedger();
          return;
        }
      } catch (parseErr) {
        console.error('Failed to parse cached data:', parseErr);
      }
    }

    rawData = [];
    initLedger();
  }
});
