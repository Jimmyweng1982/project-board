// ===== Global State =====
let workbook = null;
let allProjects = {};      // { sheetName: [project, ...] }
let currentSheet = '';
let currentCustomer = 'all';

const STAGES = [
  'Definition & Planning',
  'Program Generation',
  'Design & Layout',
  'Fab & Build',
  'Program Debug',
  'Correlation & Optimize',
];

// Column indices (0-based: A=0, B=1, C=2, ...)
const COL_CUSTOMER = 2;   // Col C: Customer (tester)
const COL_PROJECT  = 3;   // Col D: Project name / POR / Actual
const STAGE_COLS   = [4, 5, 6, 7, 8, 9]; // Col E~J: 6 stages
const COL_RELEASE  = 10;  // Col K: Customer Release

// ===== DOM =====
const fileInput      = document.getElementById('excel-file');
const sheetTabsEl    = document.getElementById('sheet-tabs');
const customerFilterEl = document.getElementById('customer-filter');
const boardEl        = document.getElementById('board');
const exportBtn      = document.getElementById('export-btn');

// ===== File Import =====
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  try {
    const data = await file.arrayBuffer();

    // xlsm needs cellDates + bookVBA:false to avoid macro parse errors
    workbook = XLSX.read(data, {
      cellDates: true,
      cellNF: false,
      cellText: false,
      bookVBA: false,
    });

    allProjects = {};
    workbook.SheetNames.forEach(name => {
      allProjects[name] = parseSheet(workbook.Sheets[name]);
    });

    currentSheet = workbook.SheetNames[0];
    currentCustomer = 'all';

    renderSheetTabs();
    renderCustomerFilter();
    renderBoard();

    exportBtn.style.display = 'inline-block';
  } catch (err) {
    alert('讀取 Excel 失敗：' + err.message);
  }
});

// ===== Excel Parsing =====
function parseSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const projects = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const cellD = row[COL_PROJECT];
    if (!cellD) continue;

    const cellDStr = String(cellD).toUpperCase();

    // POR row detected → project row is above, actual row is below
    if (cellDStr.includes('POR')) {
      const pctRowIdx = findPrevDataRowIndex(rows, i);
      if (pctRowIdx < 0) continue;

      const pctRow = rows[pctRowIdx];
      const actRow = rows[i + 1] || [];

      projects.push(buildProject(pctRow, row, actRow));
    }
  }

  return projects;
}

function findPrevDataRowIndex(rows, fromIndex) {
  for (let i = fromIndex - 1; i >= 0; i--) {
    const row = rows[i];
    if (!row) continue;
    const d = row[COL_PROJECT];
    if (!d) continue;
    const s = String(d).toLowerCase();
    if (!s.includes('por') && !s.includes('actual') && !s.includes('forecast')) {
      return i;
    }
  }
  return -1;
}

function buildProject(pctRow, porRow, actRow) {
  const customer     = String(pctRow[COL_CUSTOMER] || '').trim() || 'Uncategorized';
  const projectName  = String(pctRow[COL_PROJECT]  || '').trim();
  const customerRelease = String(pctRow[COL_RELEASE] || '').trim();

  // POR exit date is embedded in the text of col D of the POR row
  const porText = String(porRow[COL_PROJECT] || '');
  const porExitMatch = porText.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  const porExitDate  = porExitMatch ? `${porExitMatch[1]}/${porExitMatch[2]}/${porExitMatch[3]}` : '';

  const stages = STAGES.map((name, idx) => {
    const col = STAGE_COLS[idx];
    return {
      name,
      percent:    parsePercent(pctRow[col]),
      porDate:    parseDate(porRow[col]),
      actualDate: parseDate(actRow[col]),
    };
  });

  const avg = Math.round(stages.reduce((s, st) => s + st.percent, 0) / stages.length);

  return { customer, projectName, porExitDate, customerRelease, stages, avg };
}

// ===== Value Parsers =====
function parsePercent(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') {
    // Stored as 0~1 decimal (percentage format) or 0~100
    return val <= 1 ? Math.round(val * 100) : Math.round(val);
  }
  const str = String(val).trim().replace('%', '');
  const match = str.match(/([\d.]+)/);
  if (!match) return 0;
  const num = parseFloat(match[1]);
  return num <= 1 ? Math.round(num * 100) : Math.round(num);
}

function parseDate(val) {
  if (val === null || val === undefined || val === '') return '';

  // JS Date object (cellDates: true)
  if (val instanceof Date) return formatDate(val);

  // Excel serial number (fallback when cellDates fails)
  if (typeof val === 'number' && val > 40000) {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return formatDate(d);
  }

  const str = String(val).trim();

  // YYYY/M/D or YYYY-M-D
  const m = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return `${m[1]}/${m[2]}/${m[3]}`;

  // Return as-is if it's a meaningful text (e.g. "Canceled", "Spirox")
  return str.length < 30 ? str : '';
}

function formatDate(d) {
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}

// ===== Date Comparison → Color =====
function getStatusColor(porDateStr, actualDateStr, percent) {
  if (percent >= 100) return 'green';
  if (!porDateStr || !actualDateStr) return 'green';

  const por = parseDateObj(porDateStr);
  const act = parseDateObj(actualDateStr);
  if (!por || !act) return 'green';

  const diffDays = (act - por) / (1000 * 60 * 60 * 24);
  if (diffDays >= 14) return 'red';
  if (diffDays >= 7)  return 'yellow';
  return 'green';
}

function parseDateObj(str) {
  if (!str) return null;
  const parts = str.split('/').map(Number);
  if (parts.length !== 3 || isNaN(parts[0])) return null;
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

// ===== Render Sheet Tabs =====
function renderSheetTabs() {
  sheetTabsEl.innerHTML = '';
  Object.keys(allProjects).forEach(name => {
    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (name === currentSheet ? ' active' : '');
    tab.textContent = name;
    tab.addEventListener('click', () => {
      currentSheet    = name;
      currentCustomer = 'all';
      renderSheetTabs();
      renderCustomerFilter();
      renderBoard();
    });
    sheetTabsEl.appendChild(tab);
  });
}

// ===== Render Customer Filter =====
function renderCustomerFilter() {
  customerFilterEl.innerHTML = '';
  const projects   = allProjects[currentSheet] || [];
  const customers  = [...new Set(projects.map(p => p.customer))];

  customerFilterEl.appendChild(createFilterBtn('全部', 'all'));
  customers.forEach(c => customerFilterEl.appendChild(createFilterBtn(c, c)));
}

function createFilterBtn(label, value) {
  const btn = document.createElement('button');
  btn.className = 'filter-btn' + (currentCustomer === value ? ' active' : '');
  btn.textContent = label;
  btn.addEventListener('click', () => {
    currentCustomer = value;
    renderCustomerFilter();
    renderBoard();
  });
  return btn;
}

// ===== Render Board =====
function renderBoard() {
  const projects = allProjects[currentSheet] || [];

  if (projects.length === 0) {
    boardEl.innerHTML = '<div class="empty-state"><p style="color:#aaa;">此工作表沒有可辨識的專案資料</p></div>';
    return;
  }

  const filtered = currentCustomer === 'all'
    ? projects
    : projects.filter(p => p.customer === currentCustomer);

  // Group by customer
  const grouped = {};
  filtered.forEach(p => {
    if (!grouped[p.customer]) grouped[p.customer] = [];
    grouped[p.customer].push(p);
  });

  boardEl.innerHTML = '';

  Object.entries(grouped).forEach(([customer, projs]) => {
    const group = document.createElement('div');
    group.className = 'customer-group';
    group.innerHTML = `
      <div class="customer-header">
        <h2>${escapeHtml(customer)}</h2>
        <span class="project-count">${projs.length} 個專案</span>
      </div>
    `;
    projs.forEach(proj => group.appendChild(createProjectCard(proj)));
    boardEl.appendChild(group);
  });
}

function createProjectCard(proj) {
  const card = document.createElement('div');
  card.className = 'project-card';
  const avgColor = getAvgColor(proj);

  card.innerHTML = `
    <div class="project-title-row">
      <div>
        <span class="project-name">${escapeHtml(proj.projectName)}</span>
        ${proj.porExitDate ? `<span style="color:#666;font-size:0.8rem;margin-left:12px;">POR Exit: ${escapeHtml(proj.porExitDate)}</span>` : ''}
      </div>
      <div class="project-avg">
        <span>平均進度</span>
        <div class="avg-bar-container">
          <div class="avg-bar-fill progress-${avgColor}" style="width:${proj.avg}%;"></div>
        </div>
        <strong>${proj.avg}%</strong>
      </div>
    </div>
    <div class="stages-grid"></div>
  `;

  const grid = card.querySelector('.stages-grid');
  proj.stages.forEach((stage, idx) => grid.appendChild(createStageItem(proj, stage, idx)));
  return card;
}

function createStageItem(proj, stage, stageIdx) {
  const item = document.createElement('div');
  item.className = 'stage-item';

  const color = getStatusColor(stage.porDate, stage.actualDate, stage.percent);
  const warnClass = color === 'red' ? 'late-danger' : (color === 'yellow' ? 'late-warning' : '');

  item.innerHTML = `
    <div class="stage-name">${escapeHtml(stage.name)}</div>
    <div class="progress-bar" data-stage-idx="${stageIdx}">
      <div class="progress-fill progress-${color}" style="width:${stage.percent}%;"></div>
      <span class="progress-text">${stage.percent}%</span>
    </div>
    <div class="stage-dates">
      <div class="date-row">
        <span class="date-label">POR</span>
        <span class="date-value">${escapeHtml(stage.porDate || '-')}</span>
      </div>
      <div class="date-row">
        <span class="date-label">Actual</span>
        <span class="date-value ${warnClass}">${escapeHtml(stage.actualDate || '-')}</span>
      </div>
    </div>
  `;

  setupDrag(item.querySelector('.progress-bar'), stage, proj);
  return item;
}

// ===== Progress Bar Drag =====
function setupDrag(barEl, stage, proj) {
  let dragging = false;

  function updateFromEvent(e) {
    const rect   = barEl.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    let pct = Math.round(((clientX - rect.left) / rect.width) * 100);
    pct = Math.max(0, Math.min(100, Math.round(pct / 5) * 5));

    stage.percent = pct;
    proj.avg = Math.round(proj.stages.reduce((s, st) => s + st.percent, 0) / proj.stages.length);

    const fill  = barEl.querySelector('.progress-fill');
    const text  = barEl.querySelector('.progress-text');
    const color = getStatusColor(stage.porDate, stage.actualDate, pct);
    fill.style.width  = pct + '%';
    fill.className    = 'progress-fill progress-' + color;
    text.textContent  = pct + '%';

    const card = barEl.closest('.project-card');
    if (card) {
      const avgFill  = card.querySelector('.avg-bar-fill');
      const avgText  = card.querySelector('.project-avg strong');
      const avgColor = getAvgColor(proj);
      avgFill.style.width = proj.avg + '%';
      avgFill.className   = 'avg-bar-fill progress-' + avgColor;
      avgText.textContent = proj.avg + '%';
    }
  }

  barEl.addEventListener('mousedown', e => { dragging = true; updateFromEvent(e); e.preventDefault(); });
  barEl.addEventListener('touchstart', e => { dragging = true; updateFromEvent(e); }, { passive: true });
  document.addEventListener('mousemove', e => { if (dragging) updateFromEvent(e); });
  document.addEventListener('touchmove', e => { if (dragging) updateFromEvent(e); }, { passive: true });
  document.addEventListener('mouseup',   () => { dragging = false; });
  document.addEventListener('touchend',  () => { dragging = false; });
}

// ===== Project Average Color =====
function getAvgColor(proj) {
  if (proj.stages.some(s => getStatusColor(s.porDate, s.actualDate, s.percent) === 'red'))    return 'red';
  if (proj.stages.some(s => getStatusColor(s.porDate, s.actualDate, s.percent) === 'yellow')) return 'yellow';
  return 'green';
}

// ===== Export =====
exportBtn.addEventListener('click', () => {
  if (!workbook) return;

  const projects = allProjects[currentSheet] || [];
  const sheet    = workbook.Sheets[currentSheet];
  const rows     = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  let projIdx = 0;
  for (let i = 0; i < rows.length && projIdx < projects.length; i++) {
    const row = rows[i];
    if (!row || !row[COL_PROJECT]) continue;
    if (String(row[COL_PROJECT]).toUpperCase().includes('POR')) {
      const pctRowIdx = findPrevDataRowIndex(rows, i);
      if (pctRowIdx >= 0) {
        const proj = projects[projIdx];
        proj.stages.forEach((stage, sIdx) => {
          const col     = STAGE_COLS[sIdx];
          const cellRef = XLSX.utils.encode_cell({ r: pctRowIdx, c: col });
          if (sheet[cellRef]) {
            sheet[cellRef].v = stage.percent / 100;
            sheet[cellRef].t = 'n';
          }
        });
      }
      projIdx++;
    }
  }

  XLSX.writeFile(workbook, 'project-board-export.xlsx');
});

// ===== Utility =====
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}
