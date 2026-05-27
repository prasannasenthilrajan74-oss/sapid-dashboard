// Dashboard Application State
let rawData = [];
let employeesList = []; // Aggregated employees list
let currentFilters = {
  corp: 'ALL',
  dept: 'ALL',
  attended: 'ALL'
};

// Toggle for range pie charts slice grouping: 'corp' or 'dept'
let pieSliceGrouping = 'corp';


let selectedLeadershipCourse = '';

// Chart Instances
let plantChartInstance = null;
let deptChartInstance = null;
let rangeChartInstances = {};
let leadershipChartInstance = null;

// Color Palettes
const CORP_COLORS = {
  'H0': '#0ea5e9',       // Sky Blue
  'P1': '#6366f1',       // Indigo
  'P2': '#a855f7',       // Purple
  'P3': '#ec4899',       // Pink
  'P4': '#14b8a6',       // Teal
  'Other': '#64748b'     // Muted Slate
};

const DEPT_COLORS = {
  'Admin': '#f43f5e',      // Rose
  'QA': '#10b981',         // Emerald
  'Operations': '#f59e0b', // Amber
  'HR': '#3b82f6',         // Blue
  'IT': '#8b5cf6',         // Purple
  'Finance': '#06b6d4',    // Cyan
  'Production': '#84cc16', // Lime
  'Other': '#64748b'
};

const RANGE_CONFIGS = [
  { id: '0-10', label: '0-10 Mandays', min: 0, max: 10, color: 'emerald' },
  { id: '10-20', label: '10-20 Mandays', min: 10.0001, max: 20, color: 'blue' },
  { id: '20-30', label: '20-30 Mandays', min: 20.0001, max: 30, color: 'amber' },
  { id: '30-40', label: '30-40 Mandays', min: 30.0001, max: 40, color: 'orange' },
  { id: '40-50', label: '40-50 Mandays', min: 40.0001, max: 50, color: 'rose' }
];

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

// Default Google Sheets URL
const DEFAULT_SHEET_URL = 'https://script.google.com/macros/s/AKfycbytaI36PDf09D7O2RicMWEkGn-JXiew3zPL6bc3OLGKTc0klmd0gUj9ZCfdg2JvY9Sb/exec';

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
  
  // Use /gviz/tq endpoint to prevent CORS redirects in browser environments
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}`;
}

// Function to fetch data from Google Sheets or Apps Script Web App
async function fetchGoogleSheetsData(sheetUrl) {
  const isAppsScript = sheetUrl.includes('script.google.com/macros/s/');
  
  let csvUrl = sheetUrl;
  if (!isAppsScript) {
    csvUrl = getGoogleSheetsCsvUrl(sheetUrl);
    if (!csvUrl) {
      alert('Invalid Google Sheets URL format. Please paste a valid link.');
      return;
    }
  }

  // Show a visual loading state in the sync button
  const syncBtn = document.getElementById('sheet-sync-btn');
  const originalText = syncBtn ? syncBtn.innerText : 'Sync';
  if (syncBtn) {
    syncBtn.innerText = 'Syncing...';
    syncBtn.disabled = true;
  }

  try {
    const response = await fetch(csvUrl);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    let parsedData = [];

    if (isAppsScript) {
      // It's a JSON response from Apps Script
      const jsonData = await response.json();
      
      // Check if it's a 2D array (e.g. [[header1, header2], [val1, val2]])
      if (Array.isArray(jsonData) && jsonData.length > 0 && Array.isArray(jsonData[0])) {
        const headers = jsonData[0];
        parsedData = [];
        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          const obj = {};
          headers.forEach((header, colIdx) => {
            obj[header] = row[colIdx] !== undefined ? row[colIdx] : '';
          });
          parsedData.push(obj);
        }
      } else if (Array.isArray(jsonData)) {
        // Already an array of objects
        parsedData = jsonData;
      } else {
        throw new Error('Unexpected JSON format from Google Apps Script.');
      }
    } else {
      // It's a CSV response from Google Sheets Visualization API
      const csvText = await response.text();
      // Parse CSV using SheetJS
      const workbook = XLSX.read(csvText, { type: 'string' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      parsedData = XLSX.utils.sheet_to_json(worksheet);
    }

    if (parsedData.length === 0) {
      throw new Error('The data is empty or could not be parsed.');
    }

    rawData = parsedData;
    
    // Save successfully loaded URL to localStorage
    localStorage.setItem('dashboard_sheet_url', sheetUrl);
    
    // Update reset button visibility
    const resetBtn = document.getElementById('reset-btn');
    if (resetBtn) {
      if (sheetUrl !== DEFAULT_SHEET_URL) {
        resetBtn.classList.remove('hidden');
      } else {
        resetBtn.classList.add('hidden');
      }
    }
    
    initDashboard();
  } catch (err) {
    console.error('Error fetching Google Sheets data:', err);
    alert(`Failed to sync: ${err.message}\nFalling back to default preloaded data.`);
    
    // Fallback to preloaded data if rawData is empty
    if (rawData.length === 0 && typeof window.PRELOADED_DATA !== 'undefined') {
      rawData = window.PRELOADED_DATA;
      initDashboard();
    }
  } finally {
    if (syncBtn) {
      syncBtn.innerText = originalText;
      syncBtn.disabled = false;
    }
  }
}

// --- AUTHENTICATION & LOGIN FLOW ---

// Check if user is logged in (session storage or remembered in local storage)
function checkAuthStatus() {
  return sessionStorage.getItem('skilltrack_logged_in') === 'true' || 
         localStorage.getItem('skilltrack_remembered_user') === 'rane';
}

// Setup login interface event listeners
function setupLoginListeners() {
  const loginForm = document.getElementById('login-form');
  const togglePasswordBtn = document.getElementById('toggle-password-btn');
  const passwordInput = document.getElementById('password');
  const usernameInput = document.getElementById('username');
  const errorMsg = document.getElementById('login-error-msg');
  const submitBtn = document.getElementById('login-submit-btn');

  if (togglePasswordBtn && passwordInput) {
    togglePasswordBtn.addEventListener('click', () => {
      const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
      passwordInput.setAttribute('type', type);
      
      const svg = togglePasswordBtn.querySelector('.eye-icon');
      if (svg) {
        if (type === 'text') {
          svg.style.opacity = '0.5';
        } else {
          svg.style.opacity = '1';
        }
      }
    });
  }

  if (loginForm) {
    loginForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      
      const usernameVal = usernameInput.value.trim();
      const passwordVal = passwordInput.value;
      
      // Clear error states
      errorMsg.classList.add('hidden');
      errorMsg.innerText = '';
      
      const originalBtnText = submitBtn.innerHTML;
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span>Verifying...</span>';
      
      try {
        const response = await fetch('/api/login', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            username: usernameVal,
            password: passwordVal
          })
        });

        const data = await response.json();
        
        if (response.ok && data.success) {
          // Store session/remember state
          const rememberMe = document.getElementById('remember-me').checked;
          if (rememberMe) {
            localStorage.setItem('skilltrack_remembered_user', usernameVal);
          } else {
            localStorage.removeItem('skilltrack_remembered_user');
          }
          sessionStorage.setItem('skilltrack_logged_in', 'true');
          
          // Animate and hide login screen
          const loginContainer = document.getElementById('login-container');
          loginContainer.classList.add('fade-out');
          
          // Show dashboard
          const dashboardWrapper = document.querySelector('.dashboard-wrapper');
          dashboardWrapper.classList.remove('hidden');
          
          // Fetch data from Google Sheet
          const savedUrl = localStorage.getItem('dashboard_sheet_url') || DEFAULT_SHEET_URL;
          fetchGoogleSheetsData(savedUrl);
          
          // Remove from layout after transition completes
          setTimeout(() => {
            loginContainer.classList.add('hidden');
          }, 400);
        } else {
          // Validation error
          errorMsg.classList.remove('hidden');
          errorMsg.innerText = data.message || 'Invalid username or password.';
          
          // Clear password input and focus
          passwordInput.value = '';
          passwordInput.focus();
          
          // Shake the card to indicate validation failure
          const card = document.querySelector('.login-card');
          if (card) {
            card.style.animation = 'none';
            void card.offsetWidth; // Trigger reflow to restart animation
            card.style.animation = 'shake 0.4s ease-in-out';
          }
        }
      } catch (err) {
        console.error('Authentication error:', err);
        errorMsg.classList.remove('hidden');
        errorMsg.innerText = 'Unable to connect to the authentication server. Please try again.';
      } finally {
        submitBtn.disabled = false;
        submitBtn.innerHTML = originalBtnText;
      }
    });
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
  setupLoginListeners();

  // Load URL from localStorage or default
  const savedUrl = localStorage.getItem('dashboard_sheet_url') || DEFAULT_SHEET_URL;
  const urlInput = document.getElementById('sheet-url-input');
  if (urlInput) {
    urlInput.value = savedUrl;
  }

  // Handle conditional initial data fetch depending on authorization status
  if (checkAuthStatus()) {
    // Hide login container overlay instantly
    const loginContainer = document.getElementById('login-container');
    if (loginContainer) {
      loginContainer.classList.add('hidden');
    }
    // Show main dashboard container
    const dashboardWrapper = document.querySelector('.dashboard-wrapper');
    if (dashboardWrapper) {
      dashboardWrapper.classList.remove('hidden');
    }
    // Fetch data immediately
    fetchGoogleSheetsData(savedUrl);
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
    const urlInput = document.getElementById('sheet-url-input');
    if (urlInput) {
      urlInput.value = DEFAULT_SHEET_URL;
    }
    fileInput.value = '';
    resetBtn.classList.add('hidden');
    fetchGoogleSheetsData(DEFAULT_SHEET_URL);
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
  document.getElementById('filter-corp').addEventListener('change', (e) => {
    currentFilters.corp = e.target.value;
    updateDashboard();
  });
  document.getElementById('filter-dept').addEventListener('change', (e) => {
    currentFilters.dept = e.target.value;
    updateDashboard();
  });
  document.getElementById('filter-attended').addEventListener('change', (e) => {
    currentFilters.attended = e.target.value;
    updateDashboard();
  });


  // Pie chart slice toggle (Group by Corp vs Group by Dept)
  const btnGroupCorp = document.getElementById('toggle-slice-corp');
  const btnGroupDept = document.getElementById('toggle-slice-dept');

  btnGroupCorp.addEventListener('click', () => {
    if (pieSliceGrouping !== 'corp') {
      pieSliceGrouping = 'corp';
      btnGroupCorp.classList.add('active');
      btnGroupDept.classList.remove('active');
      renderRangePieCharts();
    }
  });

  btnGroupDept.addEventListener('click', () => {
    if (pieSliceGrouping !== 'dept') {
      pieSliceGrouping = 'dept';
      btnGroupDept.classList.add('active');
      btnGroupCorp.classList.remove('active');
      renderRangePieCharts();
    }
  });



  // Modal Close
  document.getElementById('modal-close-btn').addEventListener('click', closeModal);
  document.getElementById('drilldown-modal').addEventListener('click', (e) => {
    if (e.target.id === 'drilldown-modal') closeModal();
  });

  // Leadership Track category change event
  document.getElementById('filter-leadership-cat').addEventListener('change', () => {
    renderLeadershipChart();
  });
}

// Handle Excel Upload
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const parsedData = XLSX.utils.sheet_to_json(worksheet);

      if (parsedData.length === 0) {
        alert("The uploaded excel sheet is empty!");
        return;
      }

      rawData = parsedData;
      
      // Show reset button
      document.getElementById('reset-btn').classList.remove('hidden');

      initDashboard();
    } catch (err) {
      console.error(err);
      alert("Error parsing excel file: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Initialize components (options, dropdowns)
function initDashboard() {
  // Populate filter options dynamically
  const corps = new Set();
  const depts = new Set();
  const cats = new Set();

  rawData.forEach(row => {
    const corp = getRowCorpPlant(row);
    const dept = getRowDepartment(row);
    const cat = getRowCategory(row);
    if (corp) corps.add(corp);
    if (dept) depts.add(dept);
    if (cat) cats.add(cat);
  });

  // Populate Corp select
  const corpSelect = document.getElementById('filter-corp');
  corpSelect.innerHTML = '<option value="ALL">All Plants</option>';
  Array.from(corps).sort().forEach(corp => {
    corpSelect.innerHTML += `<option value="${corp}">${corp}</option>`;
  });

  // Populate Dept select
  const deptSelect = document.getElementById('filter-dept');
  deptSelect.innerHTML = '<option value="ALL">All Departments</option>';
  Array.from(depts).sort().forEach(dept => {
    deptSelect.innerHTML += `<option value="${dept}">${dept}</option>`;
  });

  // Populate Leadership Track Category select dynamically
  const catSelect = document.getElementById('filter-leadership-cat');
  catSelect.innerHTML = '<option value="ALL">All Categories</option>';
  Array.from(cats).sort().forEach(cat => {
    catSelect.innerHTML += `<option value="${cat}">${cat}</option>`;
  });
  catSelect.value = 'ALL';

  // Reset Filters to ALL
  currentFilters = { corp: 'ALL', dept: 'ALL', attended: 'ALL' };
  document.getElementById('filter-corp').value = 'ALL';
  document.getElementById('filter-dept').value = 'ALL';
  document.getElementById('filter-attended').value = 'ALL';

  updateDashboard();
}

// Update dashboard metrics and charts
function updateDashboard() {
  // 1. Filter Raw Data
  const filteredRows = rawData.filter(row => {
    const matchCorp = currentFilters.corp === 'ALL' || getRowCorpPlant(row) === currentFilters.corp;
    const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
    const matchAttended = currentFilters.attended === 'ALL' || getRowAttended(row) === currentFilters.attended;
    return matchCorp && matchDept && matchAttended;
  });

  // 2. Aggregate Employees
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

  // 3. Render KPIs
  renderKPIs(filteredRows, employeesList);

  // 4. Render Main Charts
  renderPlantChart(filteredRows);
  renderDeptChart(filteredRows);

  // 5. Render Range Pie Charts (based on employees list)
  renderRangePieCharts();

  // Render Leadership Track Segment
  renderLeadershipChart();

  // Render TQM and Category Progress on the right column
  renderRightColumnAnalytics(filteredRows);


}

// Render KPI Summary Cards
function renderKPIs(rows, emps) {
  // Total Mandays
  const totalMandays = rows.reduce((acc, row) => acc + getRowMandays(row), 0);
  document.getElementById('kpi-total-mandays').innerText = totalMandays.toFixed(1);

  // Total Unique Employees
  document.getElementById('kpi-total-employees').innerText = emps.length;

  // Average Mandays per Employee
  const avgMandays = emps.length > 0 ? (totalMandays / emps.length) : 0;
  document.getElementById('kpi-avg-mandays').innerText = avgMandays.toFixed(1);

  // Find Top Plant (highest sum of mandays)
  const plantSums = {};
  rows.forEach(row => {
    const plant = getRowCorpPlant(row);
    if (plant) {
      plantSums[plant] = (plantSums[plant] || 0) + getRowMandays(row);
    }
  });

  let topPlant = '-';
  let maxPlantSum = 0;
  Object.keys(plantSums).forEach(p => {
    if (plantSums[p] > maxPlantSum) {
      maxPlantSum = plantSums[p];
      topPlant = p;
    }
  });
  
  if (topPlant !== '-') {
    document.getElementById('kpi-top-plant').innerText = `${topPlant} (${maxPlantSum.toFixed(1)} d)`;
  } else {
    document.getElementById('kpi-top-plant').innerText = '-';
  }
}

// Render First Graph: Sum of Mandays for each Corp/Plant
function renderPlantChart(rows) {
  if (typeof Chart === 'undefined') return;

  // Aggregate sum of mandays by Plant
  const plantSums = {};
  rows.forEach(row => {
    const plant = getRowCorpPlant(row);
    plantSums[plant] = (plantSums[plant] || 0) + getRowMandays(row);
  });

  const labels = Object.keys(plantSums).sort();
  const dataValues = labels.map(label => plantSums[label]);
  const colors = labels.map(label => CORP_COLORS[label] || CORP_COLORS['Other']);

  if (plantChartInstance) {
    plantChartInstance.destroy();
  }

  const ctx = document.getElementById('chart-plant-mandays').getContext('2d');
  plantChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Sum of Mandays',
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
            label: (context) => ` Total Mandays: ${context.raw.toFixed(1)}`
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
            callback: (val) => `${val} d`
          }
        }
      }
    }
  });
}

// Render Second Graph of Segment 1: Department breakdown
function renderDeptChart(rows) {
  if (typeof Chart === 'undefined') return;

  const deptSums = {};
  rows.forEach(row => {
    const dept = getRowDepartment(row);
    deptSums[dept] = (deptSums[dept] || 0) + getRowMandays(row);
  });

  // Sort departments by descending mandays
  const labels = Object.keys(deptSums).sort((a,b) => deptSums[b] - deptSums[a]);
  const dataValues = labels.map(label => deptSums[label]);
  const colors = labels.map(label => DEPT_COLORS[label] || DEPT_COLORS['Other']);

  if (deptChartInstance) {
    deptChartInstance.destroy();
  }

  const ctx = document.getElementById('chart-dept-mandays').getContext('2d');
  deptChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Sum of Mandays',
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
            label: (context) => ` Total Mandays: ${context.raw.toFixed(1)}`
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)', drawBorder: false },
          beginAtZero: true
        },
        y: {
          grid: { display: false, drawBorder: false },
          ticks: { font: { weight: '600' } }
        }
      }
    }
  });
}

// Helper to compute dynamic ranges based on employee mandays
function getDynamicRangeConfigs(maxMandays) {
  if (maxMandays <= 50) {
    return [
      { id: '0-10', label: '0-10 Mandays', min: 0, max: 10, color: 'emerald' },
      { id: '10-20', label: '10-20 Mandays', min: 10.0001, max: 20, color: 'blue' },
      { id: '20-30', label: '20-30 Mandays', min: 20.0001, max: 30, color: 'amber' },
      { id: '30-40', label: '30-40 Mandays', min: 30.0001, max: 40, color: 'orange' },
      { id: '40-50', label: '40-50 Mandays', min: 40.0001, max: 50, color: 'rose' }
    ];
  }
  
  // Scale intervals dynamically. Round interval up to nearest 10 for clean visual ranges.
  const interval = Math.ceil(maxMandays / 5 / 10) * 10;
  const configs = [];
  const colors = ['emerald', 'blue', 'amber', 'orange', 'rose'];
  
  for (let i = 0; i < 5; i++) {
    const min = i * interval;
    const max = (i + 1) * interval;
    configs.push({
      id: `${min}-${max}`,
      label: `${min}-${max} Mandays`,
      min: i === 0 ? min : min + 0.0001,
      max: max,
      color: colors[i]
    });
  }
  return configs;
}

// Render dynamic range pie charts representing intervals
function renderRangePieCharts() {
  const container = document.getElementById('ranges-pies-container');
  if (!container) return;

  // 1. Find max mandays to determine scaling boundaries
  const maxMandays = employeesList.reduce((max, emp) => Math.max(max, emp.mandaysSum), 0);
  
  // 2. Generate range configs
  const rangeConfigs = getDynamicRangeConfigs(maxMandays);

  // 3. Clear container and recreate cards dynamically
  container.innerHTML = '';
  
  rangeConfigs.forEach(range => {
    const cardEl = document.createElement('div');
    cardEl.className = 'glass-card pie-card';
    cardEl.id = `card-range-${range.id}`;
    
    // Choose bullet icon color
    const bullet = range.color === 'emerald' ? '🟢' : range.color === 'blue' ? '🔵' : range.color === 'amber' ? '🟡' : range.color === 'orange' ? '🟠' : '🔴';
    
    cardEl.innerHTML = `
      <div class="pie-card-header">
        <span class="pie-card-title">${bullet} ${range.label}</span>
        <span class="pie-card-badge" id="badge-range-${range.id}">0 Employees</span>
      </div>
      <div class="chart-container-donut">
        <canvas id="chart-range-${range.id}"></canvas>
      </div>
    `;
    container.appendChild(cardEl);
  });

  // 4. Render chart canvases
  rangeConfigs.forEach(range => {
    // Find employees whose mandaysSum fits in this range
    const employeesInRange = employeesList.filter(emp => {
      const s = emp.mandaysSum;
      return s >= range.min && s <= range.max;
    });

    // Update Badge text
    const badge = document.getElementById(`badge-range-${range.id}`);
    if (badge) {
      badge.innerText = `${employeesInRange.length} Employee${employeesInRange.length !== 1 ? 's' : ''}`;
    }

    const canvas = document.getElementById(`chart-range-${range.id}`);
    const cardEl = document.getElementById(`card-range-${range.id}`);

    // If there are no employees in this range, display empty state and skip drawing
    if (employeesInRange.length === 0) {
      if (rangeChartInstances[range.id]) {
        rangeChartInstances[range.id].destroy();
        delete rangeChartInstances[range.id];
      }
      
      canvas.style.display = 'none';
      const emptyState = document.createElement('div');
      emptyState.className = 'pie-empty-state';
      emptyState.innerHTML = `<i>🫙</i><p>No employees in this range</p>`;
      cardEl.appendChild(emptyState);
      return;
    }

    // Group employees in this range by Plant or Department depending on current toggle
    const sliceCounts = {};
    employeesInRange.forEach(emp => {
      const key = pieSliceGrouping === 'corp' ? emp.corp : emp.dept;
      sliceCounts[key] = (sliceCounts[key] || 0) + 1;
    });

    const labels = Object.keys(sliceCounts).sort();
    const dataValues = labels.map(label => sliceCounts[label]);
    
    // Choose appropriate color palette
    const colors = labels.map(label => {
      if (pieSliceGrouping === 'corp') {
        return CORP_COLORS[label] || CORP_COLORS['Other'];
      } else {
        return DEPT_COLORS[label] || DEPT_COLORS['Other'];
      }
    });

    if (rangeChartInstances[range.id]) {
      rangeChartInstances[range.id].destroy();
    }

    if (typeof Chart === 'undefined') return;

    const ctx = canvas.getContext('2d');
    rangeChartInstances[range.id] = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: dataValues,
          backgroundColor: colors,
          borderWidth: 2,
          borderColor: '#111827', // Card background color
          hoverOffset: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '60%',
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              boxWidth: 10,
              padding: 10,
              font: { size: 9 }
            }
          },
          tooltip: {
            callbacks: {
              label: (context) => ` ${context.label}: ${context.raw} employee${context.raw !== 1 ? 's' : ''}`
            }
          }
        },
        onClick: (evt, activeElements) => {
          if (activeElements.length > 0) {
            const index = activeElements[0].index;
            const clickedLabel = labels[index];
            showDrilldownDetail(range, pieSliceGrouping, clickedLabel, employeesInRange);
          }
        }
      }
    });
  });
}

// Open modal detailing employees in a specific slice (range + plant/dept)
function showDrilldownDetail(rangeInfo, groupType, groupValue, employeesInRange) {
  // Filter employees belonging to the clicked group (e.g. Corp 'H0' or Dept 'IT')
  const detailsList = employeesInRange.filter(emp => {
    const val = groupType === 'corp' ? emp.corp : emp.dept;
    return val === groupValue;
  });

  // Get employee IDs in detailsList
  const employeeIds = new Set(detailsList.map(emp => emp.id));

  // Filter rawData to matching rows under current active filters and matching employee IDs
  const rawRowsList = rawData.filter(row => {
    const genId = getRowEmployeeId(row) || `EMP-${getRowEmployeeName(row)}`;
    if (!employeeIds.has(genId)) return false;

    const matchCorp = currentFilters.corp === 'ALL' || getRowCorpPlant(row) === currentFilters.corp;
    const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
    const matchAttended = currentFilters.attended === 'ALL' || getRowAttended(row) === currentFilters.attended;
    return matchCorp && matchDept && matchAttended;
  });

  // Sort them by Department! (Requirement: "like that it shld be sort baes on the dept")
  rawRowsList.sort((a, b) => {
    const deptA = getRowDepartment(a);
    const deptB = getRowDepartment(b);
    if (deptA !== deptB) return deptA.localeCompare(deptB);
    const nameA = getRowEmployeeName(a);
    const nameB = getRowEmployeeName(b);
    return nameA.localeCompare(nameB);
  });

  // Populate Modal
  const modalTitle = document.getElementById('modal-title');
  const modalSubtitle = document.getElementById('modal-subtitle');
  const modalTableBody = document.getElementById('modal-table-body');

  modalTitle.innerText = `${rangeInfo.label} Breakdown`;
  modalSubtitle.innerText = `${groupType === 'corp' ? 'Plant' : 'Department'} : ${groupValue} (${rawRowsList.length} record${rawRowsList.length !== 1 ? 's' : ''}, sorted by Department)`;

  modalTableBody.innerHTML = '';
  rawRowsList.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="padding: 0.8rem 1rem;">${row.Sno || '-'}</td>
      <td style="padding: 0.8rem 1rem; font-weight: 600;">${getRowEmployeeId(row)}</td>
      <td style="padding: 0.8rem 1rem;">${getRowEmployeeName(row)}</td>
      <td style="padding: 0.8rem 1rem;">${row.Des || row['Designation'] || row['des'] || row['designation'] || '-'}</td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-dept">${getRowDepartment(row)}</span></td>
      <td style="padding: 0.8rem 1rem;"><span class="badge-plant">${getRowCorpPlant(row)}</span></td>
      <td style="padding: 0.8rem 1rem;">${getRowCategory(row) || '-'}</td>
      <td style="padding: 0.8rem 1rem; font-size: 0.85rem; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${getRowCourseName(row)}">${getRowCourseName(row)}</td>
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

  // Show Modal
  document.getElementById('drilldown-modal').classList.add('active');
}

function closeModal() {
  document.getElementById('drilldown-modal').classList.remove('active');
}



// Render Leadership Track Segment: displays courses as cards/buttons
function renderLeadershipChart() {
  const selectedCat = document.getElementById('filter-leadership-cat').value;

  // Filter raw data by current global filters (Plant, Attended, Search) and local Category filter
  const filteredRows = rawData.filter(row => {
    const matchCorp = currentFilters.corp === 'ALL' || getRowCorpPlant(row) === currentFilters.corp;
    const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
    const matchAttended = currentFilters.attended === 'ALL' || getRowAttended(row) === currentFilters.attended;
    const matchCat = selectedCat === 'ALL' || getRowCategory(row) === selectedCat;
    return matchCorp && matchDept && matchAttended && matchCat;
  });

  // Count unique persons (GenID) per Course Name (cname) and aggregate employee lists
  const courseData = {};
  filteredRows.forEach(row => {
    const course = getRowCourseName(row);
    if (!course) return;
    const genId = getRowEmployeeId(row) || `EMP-${getRowEmployeeName(row)}`;
    const name = getRowEmployeeName(row);
    const corp = getRowCorpPlant(row);
    const dept = getRowDepartment(row);

    if (!courseData[course]) {
      courseData[course] = {
        name: course,
        employeeMap: {}
      };
    }

    if (!courseData[course].employeeMap[genId]) {
      courseData[course].employeeMap[genId] = {
        id: genId,
        name: name,
        corp: corp,
        dept: dept
      };
    }
  });

  const courses = Object.keys(courseData).sort();

  // Populate dynamic course buttons
  const buttonsContainer = document.getElementById('course-buttons-container');
  buttonsContainer.innerHTML = '';

  if (courses.length === 0) {
    buttonsContainer.innerHTML = `
      <div style="grid-column: 1/-1; text-align: center; color: var(--text-muted); padding: 2rem 0;">
        No courses found in this category
      </div>
    `;
    // Clear right side
    document.getElementById('selected-course-title').innerText = 'Enrolled Employees';
    document.getElementById('selected-course-subtitle').innerText = 'Select a course to view details';
    document.getElementById('leadership-employees-body').innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--text-muted); padding: 2rem 0; font-size: 0.85rem;">
          No courses available
        </td>
      </tr>
    `;
    return;
  }

  // Reset or bound-check selected course
  if (!selectedLeadershipCourse || !courses.includes(selectedLeadershipCourse)) {
    selectedLeadershipCourse = courses[0];
  }

  courses.forEach((courseName, index) => {
    const count = Object.keys(courseData[courseName].employeeMap).length;
    const btn = document.createElement('div');
    btn.className = `course-card-btn ${selectedLeadershipCourse === courseName ? 'active' : ''}`;
    btn.innerHTML = `
      <span class="course-btn-title">${courseName}</span>
      <span class="course-btn-value" id="leadership-course-val-${index}">0</span>
    `;

    // Click handler to select course and render employees on right side
    btn.addEventListener('click', () => {
      // Update active styling
      const allBtns = buttonsContainer.querySelectorAll('.course-card-btn');
      allBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      selectedLeadershipCourse = courseName;
      renderLeadershipEmployeesList(courseData[courseName], selectedCat);
    });

    buttonsContainer.appendChild(btn);

    // Animate count-up for course numbers using internal element query
    const valEl = btn.querySelector('.course-btn-value');
    animateCountValue(valEl, 0, count, 600 + index * 100);
  });

  // Render employee list for selected course on right side
  renderLeadershipEmployeesList(courseData[selectedLeadershipCourse], selectedCat);
}

// Render right side enrolled employees list for selected course
function renderLeadershipEmployeesList(course, selectedCat) {
  const titleEl = document.getElementById('selected-course-title');
  const subtitleEl = document.getElementById('selected-course-subtitle');
  const tbody = document.getElementById('leadership-employees-body');

  if (!course) {
    titleEl.innerText = 'Enrolled Employees';
    subtitleEl.innerText = 'Select a course to view details';
    tbody.innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--text-muted); padding: 2rem 0; font-size: 0.85rem;">
          No course selected
        </td>
      </tr>
    `;
    return;
  }

  const courseName = course.name;
  const employeeMap = course.employeeMap;
  const employees = Object.values(employeeMap);

  titleEl.innerText = courseName;
  subtitleEl.innerText = `Unique Enrolled Persons: ${employees.length} (Click employee row to view full details)`;

  tbody.innerHTML = '';
  if (employees.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="3" style="text-align: center; color: var(--text-muted); padding: 2rem 0; font-size: 0.85rem;">
          No employees enrolled
        </td>
      </tr>
    `;
    return;
  }

  // Sort employees by Department! (Requirement: "like that it shld be sort baes on the dept")
  employees.sort((a, b) => {
    return a.dept.localeCompare(b.dept);
  });

  employees.forEach(emp => {
    const tr = document.createElement('tr');
    tr.className = 'clickable-row';
    tr.innerHTML = `
      <td><span style="font-weight: 600;">${emp.name}</span> <span style="color: var(--text-secondary); font-size: 0.75rem; margin-left: 0.4rem;">(${emp.id})</span></td>
      <td><span class="badge-plant">${emp.corp}</span></td>
      <td><span class="badge-dept">${emp.dept}</span></td>
    `;

    // Click handler to open entire row drilldown modal for this employee, course and category
    tr.addEventListener('click', () => {
      // Find all raw rows in the dataset matching: employee ID, course name, category, and other active filters
      const matchingRows = rawData.filter(row => {
        const genId = getRowEmployeeId(row) || `EMP-${getRowEmployeeName(row)}`;
        if (genId !== emp.id) return false;

        const matchCorp = currentFilters.corp === 'ALL' || getRowCorpPlant(row) === currentFilters.corp;
        const matchDept = currentFilters.dept === 'ALL' || getRowDepartment(row) === currentFilters.dept;
        const matchAttended = currentFilters.attended === 'ALL' || getRowAttended(row) === currentFilters.attended;
        const matchCat = selectedCat === 'ALL' || getRowCategory(row) === selectedCat;
        const matchCourse = getRowCourseName(row) === courseName;
        return matchCorp && matchDept && matchAttended && matchCat && matchCourse;
      });

      // Sort by Department
      matchingRows.sort((a, b) => {
        const deptA = getRowDepartment(a);
        const deptB = getRowDepartment(b);
        if (deptA !== deptB) return deptA.localeCompare(deptB);
        const nameA = getRowEmployeeName(a);
        const nameB = getRowEmployeeName(b);
        return nameA.localeCompare(nameB);
      });

      // Show in modal
      const modalTitle = document.getElementById('modal-title');
      const modalSubtitle = document.getElementById('modal-subtitle');
      const modalTableBody = document.getElementById('modal-table-body');

      modalTitle.innerText = `Leadership Track: Enrolment Record`;
      modalSubtitle.innerText = `Employee: ${emp.name} | Course: ${courseName} (${matchingRows.length} record${matchingRows.length !== 1 ? 's' : ''}, sorted by Department)`;

      modalTableBody.innerHTML = '';
      matchingRows.forEach(row => {
        const trModal = document.createElement('tr');
        trModal.innerHTML = `
          <td style="padding: 0.8rem 1rem;">${row.Sno || '-'}</td>
          <td style="padding: 0.8rem 1rem; font-weight: 600;">${getRowEmployeeId(row)}</td>
          <td style="padding: 0.8rem 1rem;">${getRowEmployeeName(row)}</td>
          <td style="padding: 0.8rem 1rem;">${row.Des || row['Designation'] || row['des'] || row['designation'] || '-'}</td>
          <td style="padding: 0.8rem 1rem;"><span class="badge-dept">${getRowDepartment(row)}</span></td>
          <td style="padding: 0.8rem 1rem;"><span class="badge-plant">${getRowCorpPlant(row)}</span></td>
          <td style="padding: 0.8rem 1rem;">${getRowCategory(row) || '-'}</td>
          <td style="padding: 0.8rem 1rem; font-size: 0.85rem; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${getRowCourseName(row)}">${getRowCourseName(row)}</td>
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
        modalTableBody.appendChild(trModal);
      });

      // Show Modal
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
    
    // For TQM completions, check if we need a decimal place
    if (obj.id === 'tqm-actual-val') {
      const currentDec = (progress * (end - start) + start);
      obj.innerHTML = currentDec.toFixed(1);
    } else {
      obj.innerHTML = current;
    }
    
    if (progress < 1) {
      window.requestAnimationFrame(step);
    }
  };
  window.requestAnimationFrame(step);
}

// Render TQM and Category Progress widgets in the right column
function renderRightColumnAnalytics(filteredRows) {
  // 1. Business Excellence & TQM completions mandays (rows with attended === 'Yes')
  const actualCompletions = filteredRows
    .filter(row => getRowAttended(row) === 'Yes')
    .reduce((acc, row) => acc + getRowMandays(row), 0);

  // We animate actual completions value from 0 to the target
  const actualValEl = document.getElementById('tqm-actual-val');
  if (actualValEl) {
    const currentVal = parseFloat(actualValEl.innerText) || 0;
    animateCountValue(actualValEl, currentVal, actualCompletions, 600);
  }

  // 2. Category Progress - Unique employee count per category
  // Group by category, count unique GenID
  const catEmployees = {};
  filteredRows.forEach(row => {
    const cat = getRowCategory(row);
    if (!cat) return;
    const genId = getRowEmployeeId(row) || `EMP-${getRowEmployeeName(row)}`;
    if (!catEmployees[cat]) {
      catEmployees[cat] = new Set();
    }
    catEmployees[cat].add(genId);
  });

  const categories = Object.keys(catEmployees).sort();
  const container = document.getElementById('category-progress-container');
  if (container) {
    container.innerHTML = '';
    if (categories.length === 0) {
      container.innerHTML = `
        <div style="text-align: center; color: var(--text-muted); padding: 1rem 0; font-size: 0.85rem;">
          No category data
        </div>
      `;
      return;
    }

    categories.forEach((cat, index) => {
      const count = catEmployees[cat].size;
      const card = document.createElement('div');
      card.className = 'category-progress-card';
      card.innerHTML = `
        <span class="cat-progress-name">${cat}</span>
        <span class="cat-progress-val" id="cat-progress-val-${index}">0</span>
      `;
      container.appendChild(card);

      // Animate category progress value using internal element query
      const valEl = card.querySelector('.cat-progress-val');
      animateCountValue(valEl, 0, count, 600 + index * 100);
    });
  }
}
