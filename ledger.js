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

// Default Google Sheets URL
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

// Fetch spreadsheet data
async function fetchSpreadsheetData(sheetUrl) {
  const isAppsScript = sheetUrl.includes('script.google.com/macros/s/');
  let csvUrl = sheetUrl;
  if (!isAppsScript) {
    csvUrl = getGoogleSheetsCsvUrl(sheetUrl);
    if (!csvUrl) {
      alert('Invalid Google Sheets URL format.');
      return;
    }
  }

  try {
    const response = await fetch(csvUrl);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    let parsedData = [];
    if (isAppsScript) {
      const jsonData = await response.json();
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
        throw new Error('Unexpected JSON format.');
      }
    } else {
      const csvText = await response.text();
      const workbook = XLSX.read(csvText, { type: 'string' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      parsedData = XLSX.utils.sheet_to_json(worksheet);
    }

    if (parsedData.length === 0) {
      throw new Error('The parsed data is empty.');
    }

    rawData = parsedData;
    initLedger();
  } catch (err) {
    console.error('Error fetching sheet data:', err);
    if (typeof window.PRELOADED_DATA !== 'undefined') {
      rawData = window.PRELOADED_DATA;
      initLedger();
    } else {
      const tbody = document.getElementById('employee-table-body');
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
  corpSelect.innerHTML = '<option value="ALL">All Plants</option>';
  Array.from(corps).sort().forEach(corp => {
    corpSelect.innerHTML += `<option value="${corp}">${corp}</option>`;
  });

  // Dept Select
  const deptSelect = document.getElementById('filter-dept');
  deptSelect.innerHTML = '<option value="ALL">All Departments</option>';
  Array.from(depts).sort().forEach(dept => {
    deptSelect.innerHTML += `<option value="${dept}">${dept}</option>`;
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

  // Apply Search query (only updated on Search Button click or Enter)
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
      <td style="font-weight: 600;">${emp.id}</td>
      <td>${emp.name}</td>
      <td><span class="badge-plant">${emp.corp}</span></td>
      <td><span class="badge-dept">${emp.dept}</span></td>
      <td style="text-align: center;">
        <span style="font-weight: 500;">${emp.coursesAttended}</span>
        <span style="color: var(--text-muted);">/ ${emp.coursesTotal}</span>
      </td>
      <td style="text-align: right; font-weight: 700; color: #38bdf8;">${emp.mandaysSum.toFixed(1)}</td>
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
      <td style="padding: 0.8rem 1rem; font-weight: 600;">${getRowEmployeeId(row)}</td>
      <td style="padding: 0.8rem 1rem;">${getRowEmployeeName(row)}</td>
      <td style="padding: 0.8rem 1rem;">${row.Des || row['Designation'] || row['des'] || row['designation'] || '-'}</td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-dept">${getRowDepartment(row)}</span></td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-plant">${getRowCorpPlant(row)}</span></td>
      <td style="padding: 0.8rem 1rem;">${getRowCategory(row) || '-'}</td>
      <td style="padding: 0.8rem 1rem; font-size: 0.85rem; max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${getRowCourseName(row)}">${getRowCourseName(row)}</td>
      <td style="padding: 0.8rem 1rem;">${row.fd || row['From Date'] || row['fd'] || row['from_date'] || '-'}</td>
      <td style="padding: 0.8rem 1rem;">${row.ed || row['To Date'] || row['ed'] || row['to_date'] || '-'}</td>
      <td style="padding: 0.8rem 1rem; font-size: 0.85rem;">${row.organisor || row['Organisor'] || row['Organizer'] || row['organizer'] || '-'}</td>
      <td style="padding: 0.8rem 1rem; text-align: right; font-weight: 700; color: #38bdf8;">${getRowMandays(row).toFixed(1)}</td>
      <td style="padding: 0.8rem 1rem; text-align: center;">
        <span style="color: ${getRowAttended(row) === 'Yes' ? '#10b981' : '#f43f5e'}; font-weight: 600;">
          ${getRowAttended(row) || 'No'}
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
          hEl.classList.remove('sorted-asc', 'sorted-desc');
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
}

// DOM content loaded
window.addEventListener('DOMContentLoaded', () => {
  setupLedgerListeners();
  const activeUrl = localStorage.getItem('dashboard_sheet_url') || DEFAULT_SHEET_URL;
  fetchSpreadsheetData(activeUrl);
});
