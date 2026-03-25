// ===== Global State =====
let workbook = null;
let allProjects = {};
let currentSheet = '';
let currentCustomer = 'all';

// Active date popup
let activePopup = null;

const STAGES = [
  'Definition & Planning',
  'Program Generation',
  'Design & Layout',
  'Fab & Build',
  'Program Debug',
  'Correlation & Optimize',
];

const COL_CUSTOMER = 2;
const COL_PROJECT  = 3;
const STAGE_COLS   = [4, 5, 6, 7, 8, 9];
const COL_RELEASE  = 10;

// ===== DOM =====
const fileInput        = document.getElementById('excel-file');
const sheetTabsEl      = document.getElementById('sheet-tabs');
const customerFilterEl = document.getElementById('customer-filter');
const boardEl          = document.getElementById('board');
const exportBtn        = document.getElementById('export-btn');

// Close popup when clicking outside
document.addEventListener('click', (e) => {
  if (activePopup && !activePopup.contains(e.target)) {
    closePopup();
  }
});

// ===== File Import =====
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data, { cellDates: true, cellNF: false, cellText: false, bookVBA: false });

    allProjects = {};
    workbook.SheetNames.forEach(name => {
      allProjects[name] = parseSheet(workbook.Sheets[name]);
    });

    currentSheet    = workbook.SheetNames[0];
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
    if (!row || !row[COL_PROJECT]) continue;
    if (String(row[COL_PROJECT]).toUpperCase().includes('POR')) {
      const pctRowIdx = findPrevDataRowIndex(rows, i);
      if (pctRowIdx < 0) continue;
      projects.push(buildProject(rows[pctRowIdx], row, rows[i + 1] || []));
    }
  }
  return projects;
}

function findPrevDataRowIndex(rows, fromIndex) {
  for (let i = fromIndex - 1; i >= 0; i--) {
    const row = rows[i];
    if (!row || !row[COL_PROJECT]) continue;
    const s = String(row[COL_PROJECT]).toLowerCase();
    if (!s.includes('por') && !s.includes('actual') && !s.includes('forecast')) return i;
  }
  return -1;
}

function buildProject(pctRow, porRow, actRow) {
  const customer    = String(pctRow[COL_CUSTOMER] || '').trim() || 'Uncategorized';
  const projectName = String(pctRow[COL_PROJECT]  || '').trim();
  const porText     = String(porRow[COL_PROJECT]  || '');
  const m           = porText.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  const porExitDate = m ? `${m[1]}/${m[2]}/${m[3]}` : '';

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
  return { customer, projectName, porExitDate, stages, avg };
}

function parsePercent(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val <= 1 ? Math.round(val * 100) : Math.round(val);
  const str = String(val).replace('%', '').trim();
  const m   = str.match(/([\d.]+)/);
  if (!m) return 0;
  const num = parseFloat(m[1]);
  return num <= 1 ? Math.round(num * 100) : Math.round(num);
}

function parseDate(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) return formatDate(val);
  if (typeof val === 'number' && val > 40000) {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return formatDate(d);
  }
  const str = String(val).trim();
  const m = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return `${m[1]}/${m[2]}/${m[3]}`;
  return str.length < 30 ? str : '';
}

function formatDate(d) {
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}

// ===== Status Color =====
function getStatusColor(porDateStr, actualDateStr, percent) {
  if (percent >= 100) return 'green';
  if (!porDateStr || !actualDateStr) return 'green';
  const por = parseDateObj(porDateStr);
  const act = parseDateObj(actualDateStr);
  if (!por || !act) return 'green';
  const diffDays = (act - por) / 86400000;
  if (diffDays >= 14) return 'red';
  if (diffDays >= 7)  return 'yellow';
  return 'green';
}

function parseDateObj(str) {
  if (!str) return null;
  const p = str.split('/').map(Number);
  if (p.length !== 3 || isNaN(p[0])) return null;
  return new Date(p[0], p[1] - 1, p[2]);
}

// Date string → value for <input type="date"> (YYYY-MM-DD)
function dateToInput(str) {
  if (!str) return '';
  const p = str.split('/');
  if (p.length !== 3) return '';
  return `${p[0]}-${String(p[1]).padStart(2,'0')}-${String(p[2]).padStart(2,'0')}`;
}

// <input type="date"> value → display string
function inputToDate(val) {
  if (!val) return '';
  const p = val.split('-');
  if (p.length !== 3) return '';
  return `${p[0]}/${parseInt(p[1])}/${parseInt(p[2])}`;
}

// ===== Render Sheet Tabs =====
function renderSheetTabs() {
  sheetTabsEl.innerHTML = '';
  Object.keys(allProjects).forEach(name => {
    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (name === currentSheet ? ' active' : '');
    tab.textContent = name;
    tab.addEventListener('click', () => {
      currentSheet = name; currentCustomer = 'all';
      renderSheetTabs(); renderCustomerFilter(); renderBoard();
    });
    sheetTabsEl.appendChild(tab);
  });
}

// ===== Render Customer Filter =====
function renderCustomerFilter() {
  customerFilterEl.innerHTML = '';
  const customers = [...new Set((allProjects[currentSheet] || []).map(p => p.customer))];
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

  const grouped = {};
  filtered.forEach(p => {
    (grouped[p.customer] = grouped[p.customer] || []).push(p);
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
    const grid = document.createElement('div');
    grid.className = 'cards-grid';
    projs.forEach(proj => grid.appendChild(createProjectCard(proj)));
    group.appendChild(grid);
    boardEl.appendChild(group);
  });
}

// ===== Project Card =====
function createProjectCard(proj) {
  const card = document.createElement('div');
  card.className = 'project-card';

  const avgColor = getAvgColor(proj);
  card.innerHTML = `
    <div class="card-header">
      <div class="project-name"><span title="${escapeHtml(proj.projectName)}">${escapeHtml(proj.projectName)}</span></div>
      ${proj.porExitDate ? `<span class="por-exit-badge">POR Exit: ${escapeHtml(proj.porExitDate)}</span>` : ''}
      <div class="avg-block">
        <div class="avg-bar-container">
          <div class="avg-bar-fill progress-${avgColor}" style="width:${proj.avg}%;"></div>
        </div>
        <span class="avg-pct">${proj.avg}%</span>
      </div>
    </div>
    <div class="stage-rows"></div>
  `;

  const stageRows = card.querySelector('.stage-rows');
  proj.stages.forEach((stage, idx) => stageRows.appendChild(createStageRow(proj, stage, idx, card)));
  return card;
}

// ===== Stage Row =====
function createStageRow(proj, stage, stageIdx, card) {
  const row = document.createElement('div');
  row.className = 'stage-row';

  const color     = getStatusColor(stage.porDate, stage.actualDate, stage.percent);
  const warnClass = color === 'red' ? 'late-danger' : color === 'yellow' ? 'late-warning' : '';

  // Progress bar cell
  const barEl = document.createElement('div');
  barEl.className = 'progress-bar';
  barEl.innerHTML = `
    <div class="progress-fill progress-${color}" style="width:${stage.percent}%;"></div>
    <span class="progress-text">${stage.percent}%</span>
  `;

  // POR date chip
  const porChip = createDateChip(stage.porDate, '', 'POR', (newDate) => {
    stage.porDate = newDate;
    refreshRow(row, proj, stage, card);
  });

  // Actual date chip
  const actChip = createDateChip(stage.actualDate, warnClass, 'Actual', (newDate) => {
    stage.actualDate = newDate;
    refreshRow(row, proj, stage, card);
  });

  row.innerHTML = `<div class="stage-label" title="${escapeHtml(stage.name)}">${escapeHtml(stage.name)}</div>`;
  row.appendChild(barEl);
  row.appendChild(porChip);
  row.appendChild(actChip);

  setupDrag(barEl, stage, proj, card, row);
  return row;
}

// Refresh a row's colors after date or percent change
function refreshRow(row, proj, stage, card) {
  const color     = getStatusColor(stage.porDate, stage.actualDate, stage.percent);
  const warnClass = color === 'red' ? 'late-danger' : color === 'yellow' ? 'late-warning' : '';

  const fill = row.querySelector('.progress-fill');
  fill.className = 'progress-fill progress-' + color;

  const chips = row.querySelectorAll('.date-chip');
  // update actual chip color (2nd chip)
  chips[1].className = 'date-chip' + (warnClass ? ' ' + warnClass : '') + (stage.actualDate ? '' : ' no-date');

  refreshCardAvg(proj, card);
}

function refreshCardAvg(proj, card) {
  proj.avg = Math.round(proj.stages.reduce((s, st) => s + st.percent, 0) / proj.stages.length);
  const avgColor = getAvgColor(proj);
  const fill = card.querySelector('.avg-bar-fill');
  const pct  = card.querySelector('.avg-pct');
  fill.style.width = proj.avg + '%';
  fill.className   = 'avg-bar-fill progress-' + avgColor;
  pct.textContent  = proj.avg + '%';
}

// ===== Date Chip =====
function createDateChip(dateStr, extraClass, label, onChange) {
  const chip = document.createElement('div');
  const cls  = ['date-chip', extraClass, dateStr ? '' : 'no-date'].filter(Boolean).join(' ');
  chip.className = cls;
  chip.innerHTML = `<span class="date-chip-label">${escapeHtml(label)}</span><span>${dateStr || '—'}</span>`;
  chip.title = label;

  chip.addEventListener('click', (e) => {
    e.stopPropagation();
    openDatePopup(chip, dateStr, label, (newDate) => {
      dateStr = newDate;
      chip.innerHTML = `<span class="date-chip-label">${escapeHtml(label)}</span><span>${newDate || '—'}</span>`;
      onChange(newDate);
    });
  });

  return chip;
}

// ===== Date Popup =====
function openDatePopup(anchorEl, currentDate, label, onApply) {
  closePopup();

  const popup = document.createElement('div');
  popup.className = 'date-popup';
  popup.innerHTML = `
    <label>${label} 日期</label>
    <input type="date" value="${dateToInput(currentDate)}">
    <div class="date-popup-actions">
      <button class="btn-clear">清除</button>
      <button class="btn-apply">確定</button>
    </div>
  `;

  const input = popup.querySelector('input');

  popup.querySelector('.btn-clear').addEventListener('click', (e) => {
    e.stopPropagation();
    onApply('');
    closePopup();
  });

  popup.querySelector('.btn-apply').addEventListener('click', (e) => {
    e.stopPropagation();
    onApply(inputToDate(input.value));
    closePopup();
  });

  // Position near the chip
  document.body.appendChild(popup);
  const rect = anchorEl.getBoundingClientRect();
  let top  = rect.bottom + 6;
  let left = rect.left;
  if (left + 200 > window.innerWidth) left = window.innerWidth - 210;
  if (top  + 120 > window.innerHeight) top = rect.top - 126;
  popup.style.top  = top + 'px';
  popup.style.left = left + 'px';

  activePopup = popup;
  input.focus();
}

function closePopup() {
  if (activePopup) {
    activePopup.remove();
    activePopup = null;
  }
}

// ===== Drag Progress Bar =====
function setupDrag(barEl, stage, proj, card, row) {
  let dragging = false;

  function update(e) {
    const rect    = barEl.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    let pct = Math.round(((clientX - rect.left) / rect.width) * 100);
    pct = Math.max(0, Math.min(100, Math.round(pct / 5) * 5));
    stage.percent = pct;

    barEl.querySelector('.progress-fill').style.width = pct + '%';
    barEl.querySelector('.progress-text').textContent = pct + '%';
    refreshRow(row, proj, stage, card);
  }

  barEl.addEventListener('mousedown', e => { dragging = true; update(e); e.preventDefault(); });
  barEl.addEventListener('touchstart', e => { dragging = true; update(e); }, { passive: true });
  document.addEventListener('mousemove', e => { if (dragging) update(e); });
  document.addEventListener('touchmove', e => { if (dragging) update(e); }, { passive: true });
  document.addEventListener('mouseup',  () => { dragging = false; });
  document.addEventListener('touchend', () => { dragging = false; });
}

// ===== Average Color =====
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
        projects[projIdx].stages.forEach((stage, sIdx) => {
          const cellRef = XLSX.utils.encode_cell({ r: pctRowIdx, c: STAGE_COLS[sIdx] });
          if (sheet[cellRef]) { sheet[cellRef].v = stage.percent / 100; sheet[cellRef].t = 'n'; }
        });
      }
      projIdx++;
    }
  }
  XLSX.writeFile(workbook, 'project-board-export.xlsx');
});

// ===== Utility =====
function escapeHtml(text) {
  const d = document.createElement('div');
  d.textContent = text;
  return d.innerHTML;
}
