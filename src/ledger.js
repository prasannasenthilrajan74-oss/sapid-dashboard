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
  corp: 'ALL',
  dept: 'ALL',
  search: ''
};

// Table State
let tableSortColumn = 'id'; // Default sort by Employee ID
let tableSortDirection = 'asc';
let tableCurrentPage = 1;
let tablePageSize = 10;

// Default Google Sheets URL (matching app.js)
const DEFAULT_SHEET_URL = 'https://script.google.com/macros/s/AKfycbytaI36PDf09D7O2RicMWEkGn-JXiew3zPL6bc3OLGKTc0klmd0gUj9ZCfdg2JvY9Sb/exec';

// --- ROBUST EXCEL COLUMN KEY ACCESSORS ---
function getRowCorpPlant(row) {
  return row['Corp/Plant'] || row['Corp'] || row['Plant'] || row['corp'] || row['plant'] || row['division'] || row['Division'] || 'Unknown';
}

function getRowDepartment(row) {
  return row['Dept'] || row['Department'] || row['dept'] || row['department'] || 'Unknown';
}

function getRowCategory(row) {
  return row['Cat'] || row['Category'] || row['cat'] || row['category'] || '';
}

function getRowEmployeeId(row) {
  return row['GenID'] || row['GenId'] || row['Employee ID'] || row['EmployeeID'] || row['id'] || row['ID'] || '';
}

function getRowEmployeeName(row) {
  return row['Name'] || row['name'] || row['Employee Name'] || row['EmployeeName'] || 'Unknown';
}

function getRowCourseName(row) {
  return row['cname'] || row['Course Name'] || row['CourseName'] || row['course_name'] || row['Course'] || row['course'] || 'Unknown';
}

function getRowMandays(row) {
  const val = row['mandays'] !== undefined ? row['mandays'] : (row['Mandays'] !== undefined ? row['Mandays'] : (row['manday'] !== undefined ? row['manday'] : row['Manday']));
  return parseFloat(val) || 0;
}

function getRowAttended(row) {
  return row['attended'] || row['Attended'] || '';
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
    const isAppsScript   = isDomainScript || sheetUrl.includes('script.google.com/macros/s/');

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
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
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

    if (typeof window.PRELOADED_DATA !== 'undefined') {
      rawData = window.PRELOADED_DATA;
      initLedger();
    } else {
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
}

// Initialize options
function initLedger() {
  const corps = new Set();
  const depts = new Set();

  rawData.forEach(row => {
    const corp = getRowCorpPlant(row);
    const dept = getRowDepartment(row);
    if (corp) corps.add(corp);
    if (dept) depts.add(dept);
  });

  // Corp Select
  const corpSelect = document.getElementById('filter-corp');
  corpSelect.textContent = '';
  appendSafeOption(corpSelect, 'ALL', 'All Plants');
  Array.from(corps).sort().forEach(corp => {
    appendSafeOption(corpSelect, corp, corp);
  });

  // Dept Select
  const deptSelect = document.getElementById('filter-dept');
  deptSelect.textContent = '';
  appendSafeOption(deptSelect, 'ALL', 'All Departments');
  Array.from(depts).sort().forEach(dept => {
    appendSafeOption(deptSelect, dept, dept);
  });

  // Reset Filters UI
  corpSelect.value = 'ALL';
  deptSelect.value = 'ALL';
  document.getElementById('search-employee').value = '';
  document.getElementById('table-page-size').value = '10';
  
  currentFilters = { corp: 'ALL', dept: 'ALL', search: '' };
  tableCurrentPage = 1;
  tablePageSize = 10;
  tableSortColumn = 'id';
  tableSortDirection = 'asc';

  // Update table headers UI classes
  const headers = ['th-genid', 'th-name', 'th-corp', 'th-dept', 'th-courses', 'th-mandays'];
  headers.forEach(hId => {
    const el = document.getElementById(hId);
    if (el) {
      el.classList.remove('sorted-asc', 'sorted-desc');
    }
  });
  const defaultHeader = document.getElementById('th-genid');
  if (defaultHeader) {
    defaultHeader.classList.add('sorted-asc');
  }

  updateLedger();
}

// Aggregate and trigger rendering
function updateLedger() {
  const filteredRows = rawData.filter(row => {
    const matchCorp = currentFilters.corp === 'ALL' || getRowCorpPlant(row) === currentFilters.corp;
    const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
    return matchCorp && matchDept;
  });

  const employeeMap = {};
  filteredRows.forEach(row => {
    const genId = getRowEmployeeId(row) || `EMP-${getRowEmployeeName(row)}`;
    const name = getRowEmployeeName(row);
    const corp = getRowCorpPlant(row);
    const dept = getRowDepartment(row);
    const mandays = getRowMandays(row);
    const attended = getRowAttended(row) === 'Yes';

    if (!employeeMap[genId]) {
      employeeMap[genId] = {
        id: genId,
        name: name,
        corp: corp,
        dept: dept,
        mandaysSum: 0,
        coursesTotal: 0,
        coursesAttended: 0
      };
    }

    employeeMap[genId].mandaysSum += mandays;
    employeeMap[genId].coursesTotal += 1;
    if (attended) {
      employeeMap[genId].coursesAttended += 1;
    }
  });

  employeesList = Object.values(employeeMap);
  renderEmployeeTable();
}

// Get filtered, sorted list and render to DOM
function renderEmployeeTable() {
  let list = [...employeesList];

  // Apply Search query
  if (currentFilters.search) {
    const term = currentFilters.search;
    list = list.filter(emp => {
      return emp.name.toLowerCase().includes(term) || emp.id.toLowerCase().includes(term);
    });
  }

  // Sort list
  list.sort((a, b) => {
    let valA = a[tableSortColumn];
    let valB = b[tableSortColumn];

    if (tableSortColumn === 'courses') {
      valA = a.coursesTotal;
      valB = b.coursesTotal;
    }

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
      <td style="font-weight: 600;">${escapeHTML(emp.id)}</td>
      <td>${escapeHTML(emp.name)}</td>
      <td><span class="badge-plant">${escapeHTML(emp.corp)}</span></td>
      <td><span class="badge-dept">${escapeHTML(emp.dept)}</span></td>
      <td style="text-align: center;">
        <span style="font-weight: 500;">${escapeHTML(emp.coursesAttended)}</span>
        <span style="color: var(--text-muted);">/ ${escapeHTML(emp.coursesTotal)}</span>
      </td>
      <td style="text-align: right; font-weight: 700; color: #38bdf8;">${escapeHTML(emp.mandaysSum.toFixed(1))}</td>
    `;
    
    tr.addEventListener('click', () => {
      showEmployeeDetailModal(emp.id, emp.name);
    });

    tbody.appendChild(tr);
  });

  document.getElementById('table-info').innerText = `Showing ${startIdx + 1} to ${endIdx} of ${totalItems} employees`;
  document.getElementById('btn-prev-page').disabled = tableCurrentPage === 1;
  document.getElementById('btn-next-page').disabled = tableCurrentPage === totalPages;
}

// Show detailed courses modal for selected employee
function showEmployeeDetailModal(empId, empName) {
  const rawRowsList = rawData.filter(row => {
    const genId = getRowEmployeeId(row) || `EMP-${getRowEmployeeName(row)}`;
    return genId === empId;
  });

  rawRowsList.sort((a, b) => {
    const dateA = new Date(a.fd || a['From Date'] || 0);
    const dateB = new Date(b.fd || b['From Date'] || 0);
    return dateA - dateB;
  });

  const modalTitle = document.getElementById('modal-title');
  const modalSubtitle = document.getElementById('modal-subtitle');
  const modalTableBody = document.getElementById('modal-table-body');

  modalTitle.innerText = `${empName} (${empId})`;
  modalSubtitle.innerText = `Aggregated course enrolment records (${rawRowsList.length} items, sorted by date)`;

  modalTableBody.innerHTML = '';
  rawRowsList.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="padding: 0.8rem 1rem;">${index + 1}</td>
      <td style="padding: 0.8rem 1rem; font-weight: 600;">${escapeHTML(getRowEmployeeId(row))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowEmployeeName(row))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(row.Des || row['Designation'] || row['des'] || row['designation'] || '-')}</td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-dept">${escapeHTML(getRowDepartment(row))}</span></td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-plant">${escapeHTML(getRowCorpPlant(row))}</span></td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(getRowCategory(row) || '-')}</td>
      <td style="padding: 0.8rem 1rem; font-size: 0.85rem; max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHTML(getRowCourseName(row))}">${escapeHTML(getRowCourseName(row))}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(row.fd || row['From Date'] || row['fd'] || row['from_date'] || '-')}</td>
      <td style="padding: 0.8rem 1rem;">${escapeHTML(row.ed || row['To Date'] || row['ed'] || row['to_date'] || '-')}</td>
      <td style="padding: 0.8rem 1rem; font-size: 0.85rem;">${escapeHTML(row.organisor || row['Organisor'] || row['Organizer'] || row['organizer'] || '-')}</td>
      <td style="padding: 0.8rem 1rem; text-align: right; font-weight: 700; color: #38bdf8;">${escapeHTML(getRowMandays(row).toFixed(1))}</td>
      <td style="padding: 0.8rem 1rem; text-align: center;">
        <span style="color: ${getRowAttended(row) === 'Yes' ? '#10b981' : '#f43f5e'}; font-weight: 600;">
          ${escapeHTML(getRowAttended(row) || 'No')}
        </span>
      </td>
    `;
    modalTableBody.appendChild(tr);
  });

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
    document.getElementById('filter-corp').value = 'ALL';
    document.getElementById('filter-dept').value = 'ALL';
    
    currentFilters = { corp: 'ALL', dept: 'ALL', search: '' };
    tableCurrentPage = 1;
    updateLedger();
  });

  // Dropdown filter changes
  document.getElementById('filter-corp').addEventListener('change', (e) => {
    currentFilters.corp = e.target.value;
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
  const headers = ['th-genid', 'th-name', 'th-corp', 'th-dept', 'th-courses', 'th-mandays'];
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
      return emp.name.toLowerCase().includes(currentFilters.search) || emp.id.toLowerCase().includes(currentFilters.search);
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

  // Configuration Backup/Export
  const btnExport = document.getElementById('btn-export-config');
  if (btnExport) {
    btnExport.addEventListener('click', () => {
      try {
        const config = {
          dashboard_sheet_url: localStorage.getItem('dashboard_sheet_url') || '',
          cached_sheet_data: JSON.parse(localStorage.getItem('cached_sheet_data') || '[]'),
          exported_at: new Date().toISOString(),
          app: "SkillTrack Analyzer",
          version: "1.0.0"
        };
        const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `skilltrack_config_${new Date().toISOString().split('T')[0]}.json`;
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

  let activeUrl = localStorage.getItem('dashboard_sheet_url');
  if (!activeUrl || activeUrl === 'https://script.google.com/a/macros/vitstudent.ac.in/s/AKfycbx4BgQYzPr7dDzF4zDCYkc3eANVHwxXn0ztHedVPu5Fa9ldh4MaARYjtixV21nHJPlS/exec') {
    localStorage.removeItem('dashboard_sheet_url');
    activeUrl = DEFAULT_SHEET_URL;
  }
  
  fetchSpreadsheetData(activeUrl);
});
