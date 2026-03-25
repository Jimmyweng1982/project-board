// ===== 全域狀態 =====
let workbook = null;       // 原始 Excel workbook
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

// 階段對應的 Excel 欄位索引 (0-based: C=2, D=3, E=4, F=5, G=6, H=7)
const STAGE_COLS = [2, 3, 4, 5, 6, 7];

// ===== DOM =====
const fileInput = document.getElementById('excel-file');
const sheetTabsEl = document.getElementById('sheet-tabs');
const customerFilterEl = document.getElementById('customer-filter');
const boardEl = document.getElementById('board');
const exportBtn = document.getElementById('export-btn');

// ===== 檔案匯入 =====
fileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const data = await file.arrayBuffer();
  workbook = XLSX.read(data, { cellDates: true });

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
});

// ===== Excel 解析 =====
function parseSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const projects = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[1]) continue;

    const cellB = String(row[1]);
    // 找到包含 POR 的列 → 上一列是百分比，下一列是 Actual
    if (cellB.toUpperCase().includes('POR')) {
      const pctRow = findPrevDataRow(rows, i);
      const actRow = rows[i + 1];

      if (pctRow) {
        projects.push(buildProject(pctRow, row, actRow));
      }
    }
  }

  return projects;
}

function findPrevDataRow(rows, fromIndex) {
  for (let i = fromIndex - 1; i >= 0; i--) {
    const row = rows[i];
    if (row && row[1] && !String(row[1]).toUpperCase().includes('POR')
        && !String(row[1]).toLowerCase().includes('actual')
        && !String(row[1]).toLowerCase().includes('forecast')) {
      return row;
    }
  }
  return null;
}

function buildProject(pctRow, porRow, actRow) {
  const customer = String(pctRow[0] || '').trim() || '未分類';
  const projectName = String(pctRow[1] || '').trim();
  const customerRelease = String(pctRow[8] || pctRow[STAGE_COLS.length + 2] || '').trim();

  // 解析 POR exit 日期
  const porText = String(porRow[1] || '');
  const porExitMatch = porText.match(/\d{4}\/\d{1,2}\/\d{1,2}/);
  const porExitDate = porExitMatch ? porExitMatch[0] : '';

  const stages = STAGES.map((name, idx) => {
    const col = STAGE_COLS[idx];
    return {
      name,
      percent: parsePercent(pctRow[col]),
      porDate: parseDate(porRow[col]),
      actualDate: parseDate(actRow ? actRow[col] : ''),
    };
  });

  const avg = Math.round(stages.reduce((s, st) => s + st.percent, 0) / stages.length);

  return { customer, projectName, porExitDate, customerRelease, stages, avg };
}

function parsePercent(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') {
    // Excel 可能存 0~1 的小數或 0~100 的整數
    return val <= 1 && val > 0 ? Math.round(val * 100) : Math.round(val);
  }
  const str = String(val).trim();
  const match = str.match(/([\d.]+)/);
  if (!match) return 0;
  const num = parseFloat(match[1]);
  return num <= 1 && num > 0 ? Math.round(num * 100) : Math.round(num);
}

function parseDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return formatDate(val);
  }
  const str = String(val).trim();
  const match = str.match(/\d{4}\/\d{1,2}\/\d{1,2}/);
  return match ? match[0] : str;
}

function formatDate(d) {
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}

// ===== 日期比較 → 顏色 =====
function getStatusColor(porDateStr, actualDateStr, percent) {
  if (percent >= 100) return 'green';
  if (!porDateStr || !actualDateStr) return 'green';

  const por = parseDateObj(porDateStr);
  const act = parseDateObj(actualDateStr);
  if (!por || !act) return 'green';

  const diffMs = act.getTime() - por.getTime();
  const diffDays = diffMs / (1000 * 60 * 60 * 24);

  if (diffDays >= 14) return 'red';
  if (diffDays >= 7) return 'yellow';
  return 'green';
}

function parseDateObj(str) {
  if (!str) return null;
  const parts = str.split('/').map(Number);
  if (parts.length !== 3) return null;
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

// ===== 渲染 Sheet 頁籤 =====
function renderSheetTabs() {
  sheetTabsEl.innerHTML = '';
  Object.keys(allProjects).forEach(name => {
    const tab = document.createElement('div');
    tab.className = 'sheet-tab' + (name === currentSheet ? ' active' : '');
    tab.textContent = name;
    tab.addEventListener('click', () => {
      currentSheet = name;
      currentCustomer = 'all';
      renderSheetTabs();
      renderCustomerFilter();
      renderBoard();
    });
    sheetTabsEl.appendChild(tab);
  });
}

// ===== 渲染客戶篩選 =====
function renderCustomerFilter() {
  customerFilterEl.innerHTML = '';
  const projects = allProjects[currentSheet] || [];
  const customers = [...new Set(projects.map(p => p.customer))];

  const allBtn = createFilterBtn('全部', 'all');
  customerFilterEl.appendChild(allBtn);

  customers.forEach(c => {
    customerFilterEl.appendChild(createFilterBtn(c, c));
  });
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

// ===== 渲染看板 =====
function renderBoard() {
  const projects = allProjects[currentSheet] || [];
  if (projects.length === 0) {
    boardEl.innerHTML = '<div class="empty-state"><p>此工作表沒有可辨識的專案資料</p></div>';
    return;
  }

  const filtered = currentCustomer === 'all'
    ? projects
    : projects.filter(p => p.customer === currentCustomer);

  // 依客戶分組
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

    projs.forEach((proj, projIdx) => {
      group.appendChild(createProjectCard(proj, projIdx));
    });

    boardEl.appendChild(group);
  });
}

function createProjectCard(proj, projIdx) {
  const card = document.createElement('div');
  card.className = 'project-card';

  const avgColor = getAvgColor(proj);

  card.innerHTML = `
    <div class="project-title-row">
      <div>
        <span class="project-name">${escapeHtml(proj.projectName)}</span>
        ${proj.porExitDate ? `<span style="color:#666; font-size:0.8rem; margin-left:12px;">POR Exit: ${escapeHtml(proj.porExitDate)}</span>` : ''}
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

  proj.stages.forEach((stage, stageIdx) => {
    grid.appendChild(createStageItem(proj, stage, stageIdx));
  });

  return card;
}

function createStageItem(proj, stage, stageIdx) {
  const item = document.createElement('div');
  item.className = 'stage-item';

  const color = getStatusColor(stage.porDate, stage.actualDate, stage.percent);
  const dateWarningClass = color === 'red' ? 'late-danger' : (color === 'yellow' ? 'late-warning' : '');

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
        <span class="date-value ${dateWarningClass}">${escapeHtml(stage.actualDate || '-')}</span>
      </div>
    </div>
  `;

  // 拖動進度條
  const bar = item.querySelector('.progress-bar');
  setupDrag(bar, stage, proj);

  return item;
}

// ===== 進度條拖動 =====
function setupDrag(barEl, stage, proj) {
  let dragging = false;

  function updateFromEvent(e) {
    const rect = barEl.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    let pct = Math.round(((clientX - rect.left) / rect.width) * 100);
    pct = Math.max(0, Math.min(100, pct));

    // 吸附到 5 的倍數
    pct = Math.round(pct / 5) * 5;

    stage.percent = pct;

    // 更新平均
    proj.avg = Math.round(proj.stages.reduce((s, st) => s + st.percent, 0) / proj.stages.length);

    // 更新 UI
    const fill = barEl.querySelector('.progress-fill');
    const text = barEl.querySelector('.progress-text');
    const color = getStatusColor(stage.porDate, stage.actualDate, pct);

    fill.style.width = pct + '%';
    fill.className = 'progress-fill progress-' + color;
    text.textContent = pct + '%';

    // 更新專案平均進度條
    const card = barEl.closest('.project-card');
    if (card) {
      const avgFill = card.querySelector('.avg-bar-fill');
      const avgText = card.querySelector('.project-avg strong');
      const avgColor = getAvgColor(proj);
      avgFill.style.width = proj.avg + '%';
      avgFill.className = 'avg-bar-fill progress-' + avgColor;
      avgText.textContent = proj.avg + '%';
    }
  }

  barEl.addEventListener('mousedown', (e) => {
    dragging = true;
    updateFromEvent(e);
    e.preventDefault();
  });

  barEl.addEventListener('touchstart', (e) => {
    dragging = true;
    updateFromEvent(e);
  }, { passive: true });

  document.addEventListener('mousemove', (e) => {
    if (dragging) updateFromEvent(e);
  });

  document.addEventListener('touchmove', (e) => {
    if (dragging) updateFromEvent(e);
  }, { passive: true });

  document.addEventListener('mouseup', () => { dragging = false; });
  document.addEventListener('touchend', () => { dragging = false; });
}

// ===== 專案平均顏色 =====
function getAvgColor(proj) {
  const hasRed = proj.stages.some(s => getStatusColor(s.porDate, s.actualDate, s.percent) === 'red');
  const hasYellow = proj.stages.some(s => getStatusColor(s.porDate, s.actualDate, s.percent) === 'yellow');
  if (hasRed) return 'red';
  if (hasYellow) return 'yellow';
  return 'green';
}

// ===== 匯出 =====
exportBtn.addEventListener('click', () => {
  if (!workbook) return;

  // 把修改後的百分比寫回
  const projects = allProjects[currentSheet] || [];
  const sheet = workbook.Sheets[currentSheet];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  let projIdx = 0;
  for (let i = 0; i < rows.length && projIdx < projects.length; i++) {
    const row = rows[i];
    if (!row || !row[1]) continue;
    const cellB = String(row[1]).toUpperCase();
    if (cellB.includes('POR')) {
      // 百分比在 POR 的前一列
      const pctRowIdx = findPrevDataRowIndex(rows, i);
      if (pctRowIdx >= 0) {
        const proj = projects[projIdx];
        proj.stages.forEach((stage, sIdx) => {
          const col = STAGE_COLS[sIdx];
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

function findPrevDataRowIndex(rows, fromIndex) {
  for (let i = fromIndex - 1; i >= 0; i--) {
    const row = rows[i];
    if (row && row[1] && !String(row[1]).toUpperCase().includes('POR')
        && !String(row[1]).toLowerCase().includes('actual')
        && !String(row[1]).toLowerCase().includes('forecast')) {
      return i;
    }
  }
  return -1;
}

// ===== 工具 =====
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}
