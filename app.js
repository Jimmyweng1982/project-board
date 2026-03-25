// 從 localStorage 讀取任務，沒有就用預設
let tasks = JSON.parse(localStorage.getItem('tasks')) || [
  { id: 1, title: '設計首頁 UI', desc: '畫出首頁的 wireframe', priority: 'high', status: 'todo' },
  { id: 2, title: '建立資料庫', desc: '設計 schema', priority: 'medium', status: 'todo' },
  { id: 3, title: '寫登入功能', desc: '', priority: 'medium', status: 'doing' },
  { id: 4, title: '部署到 Zeabur', desc: '設定好環境', priority: 'low', status: 'done' },
];

const modal = document.getElementById('modal');
const addBtn = document.getElementById('add-task-btn');
const cancelBtn = document.getElementById('cancel-btn');
const saveBtn = document.getElementById('save-btn');

// 儲存到 localStorage
function saveTasks() {
  localStorage.setItem('tasks', JSON.stringify(tasks));
}

// 渲染所有任務
function render() {
  ['todo', 'doing', 'done'].forEach(status => {
    const list = document.getElementById(status);
    list.innerHTML = '';
    tasks.filter(t => t.status === status).forEach(task => {
      list.appendChild(createCard(task));
    });
  });
}

// 建立任務卡片
function createCard(task) {
  const card = document.createElement('div');
  card.className = 'task-card';
  card.draggable = true;
  card.dataset.id = task.id;

  const priorityLabel = { high: '高', medium: '中', low: '低' };

  card.innerHTML = `
    <button class="delete-btn" title="刪除">&times;</button>
    <h4>${escapeHtml(task.title)}</h4>
    ${task.desc ? `<p>${escapeHtml(task.desc)}</p>` : ''}
    <span class="priority ${task.priority}">${priorityLabel[task.priority]}優先</span>
  `;

  // 刪除
  card.querySelector('.delete-btn').addEventListener('click', () => {
    tasks = tasks.filter(t => t.id !== task.id);
    saveTasks();
    render();
  });

  // 拖放事件
  card.addEventListener('dragstart', e => {
    e.dataTransfer.setData('text/plain', task.id);
    card.classList.add('dragging');
  });

  card.addEventListener('dragend', () => {
    card.classList.remove('dragging');
  });

  return card;
}

// 設定每個欄位的拖放
document.querySelectorAll('.column').forEach(col => {
  col.addEventListener('dragover', e => {
    e.preventDefault();
    col.classList.add('drag-over');
  });

  col.addEventListener('dragleave', () => {
    col.classList.remove('drag-over');
  });

  col.addEventListener('drop', e => {
    e.preventDefault();
    col.classList.remove('drag-over');
    const id = Number(e.dataTransfer.getData('text/plain'));
    const task = tasks.find(t => t.id === id);
    if (task) {
      task.status = col.dataset.status;
      saveTasks();
      render();
    }
  });
});

// 彈窗控制
addBtn.addEventListener('click', () => modal.classList.add('active'));
cancelBtn.addEventListener('click', closeModal);
modal.addEventListener('click', e => {
  if (e.target === modal) closeModal();
});

function closeModal() {
  modal.classList.remove('active');
  document.getElementById('task-title').value = '';
  document.getElementById('task-desc').value = '';
  document.getElementById('task-priority').value = 'medium';
}

// 儲存新任務
saveBtn.addEventListener('click', () => {
  const title = document.getElementById('task-title').value.trim();
  if (!title) return alert('請輸入任務名稱');

  tasks.push({
    id: Date.now(),
    title,
    desc: document.getElementById('task-desc').value.trim(),
    priority: document.getElementById('task-priority').value,
    status: 'todo',
  });

  saveTasks();
  render();
  closeModal();
});

// 防 XSS
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// 初始渲染
render();
