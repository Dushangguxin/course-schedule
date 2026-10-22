// 课程表查询软件 — 阶段 4：Excel 导入
// 数据保存在浏览器 localStorage（键名 course-schedule）
// Excel 解析依赖 SheetJS（xlsx.full.min.js，需先于本文件加载）

const STORAGE_KEY = 'course-schedule';
const SCHEMA_VERSION = 2; // 数据版本，升级时用于迁移
const TOTAL_WEEKS = 20;

const DAYS = ['周一', '周二', '周三', '周四', '周五', '周六', '周日'];

// 学校实际作息时间（5 大节）
const DEFAULT_SLOTS = [
  { no: 1, start: '08:00', end: '09:50' },
  { no: 2, start: '10:10', end: '12:00' },
  { no: 3, start: '14:00', end: '15:50' },
  { no: 4, start: '16:10', end: '18:00' },
  { no: 5, start: '19:00', end: '20:50' },
];

// —— 工具函数 ——
function range(start, end) {
  const arr = [];
  for (let i = start; i <= end; i++) arr.push(i);
  return arr;
}
const fullWeeks = () => range(1, TOTAL_WEEKS);
const oddWeeks = () => fullWeeks().filter((w) => w % 2 === 1);
const evenWeeks = () => fullWeeks().filter((w) => w % 2 === 0);
const sameWeeks = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function todayColumn() {
  const d = new Date().getDay(); // 0=周日 ... 6=周六
  return d === 0 ? 7 : d;         // 映射到 1=周一 ... 7=周日
}

function todayStr() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

// —— 数据读写 ——
function sampleCourses() {
  return [
    { id: uid(), name: '高等数学', teacher: '张老师', room: '教三201', day: 3, slot: 1, weeks: range(1, 16) },
    { id: uid(), name: '大学英语', teacher: '李老师', room: '外语楼305', day: 1, slot: 2, weeks: oddWeeks() },
    { id: uid(), name: '大学物理', teacher: '赵老师', room: '实验楼B2', day: 3, slot: 3, weeks: range(1, 8) },
    { id: uid(), name: '体育', teacher: '王老师', room: '操场', day: 5, slot: 3, weeks: evenWeeks() },
    { id: uid(), name: '线性代数', teacher: '刘老师', room: '教二108', day: 2, slot: 4, weeks: range(1, 8) },
  ];
}

function defaultState() {
  return {
    version: SCHEMA_VERSION,
    termStart: '',            // 开学日期（YYYY-MM-DD）
    timeSlots: DEFAULT_SLOTS,
    courses: sampleCourses(), // 首次打开时展示示例课程，可删除
  };
}

function migrate(s) {
  if (!s.version || s.version < SCHEMA_VERSION) {
    s.timeSlots = DEFAULT_SLOTS; // 更新为学校实际作息时间
    s.version = SCHEMA_VERSION;
  }
}

function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const s = JSON.parse(raw);
      migrate(s);
      save(s);
      return s;
    }
  } catch (e) {
    /* 数据损坏时重建 */
  }
  const s = defaultState();
  save(s);
  return s;
}

function save(s) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
}

let state = load();

// —— 周次计算 ——
function currentWeek() {
  if (!state.termStart) return null;
  const start = new Date(state.termStart + 'T00:00:00');
  const now = new Date();
  const diffDays = Math.floor((now - start) / 86400000);
  const week = Math.floor(diffDays / 7) + 1;
  return Math.max(1, Math.min(TOTAL_WEEKS, week));
}

// —— 渲染 ——
const scheduleEl = document.getElementById('schedule');

function formatWeeks(weeks) {
  if (!weeks || weeks.length === 0) return '';
  if (sameWeeks(weeks, fullWeeks())) return '全学期';
  if (sameWeeks(weeks, oddWeeks())) return '单周';
  if (sameWeeks(weeks, evenWeeks())) return '双周';
  const s = [...weeks].sort((a, b) => a - b);
  const parts = [];
  let start = s[0];
  let prev = s[0];
  for (let i = 1; i < s.length; i++) {
    if (s[i] === prev + 1) {
      prev = s[i];
    } else {
      parts.push(start === prev ? `${start}` : `${start}-${prev}`);
      start = prev = s[i];
    }
  }
  parts.push(start === prev ? `${start}` : `${start}-${prev}`);
  return parts.join(',') + '周';
}

function escapeHTML(str) {
  return String(str ?? '').replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  }[ch]));
}

// 精简教室显示：去掉「多媒体/智慧」等说明、把括号换成空格
function cleanRoom(room) {
  if (!room) return '';
  let s = String(room);
  s = s.replace(/（[^（）]*?(多媒体|智慧)[^（）]*?）/g, '');
  s = s.replace(/\([^()]*?(多媒体|智慧)[^()]*?\)/g, '');
  s = s.replace(/[（()）]/g, ' ');
  s = s.replace(/(楼)\s*(\d)/g, '$1 $2');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

function courseHTML(c) {
  const room = cleanRoom(c.room);
  const tip = escapeHTML([c.name, room].filter(Boolean).join(' · '));
  return `
    <div class="course" data-id="${c.id}" title="${tip}">
      <div class="course-name">${escapeHTML(c.name)}</div>
      ${room ? `<div class="course-room">${escapeHTML(room)}</div>` : ''}
    </div>`;
}

// 一个格子里的课程：最多显示 2 门，其余折叠为「+N」
const MAX_VISIBLE = 2;

function cellHTML(courses) {
  if (courses.length <= MAX_VISIBLE) {
    return courses.map(courseHTML).join('');
  }
  return courseHTML(courses[0]) + `<div class="more">+${courses.length - 1}</div>`;
}

function currentSlot() {
  const now = new Date();
  const minutes = now.getHours() * 60 + now.getMinutes();
  for (const s of state.timeSlots) {
    const [sh, sm] = s.start.split(':').map(Number);
    const [eh, em] = s.end.split(':').map(Number);
    if (minutes >= sh * 60 + sm && minutes < eh * 60 + em) return s.no;
  }
  return null;
}

function updateSubbar(week) {
  document.getElementById('week-el').textContent = week == null ? '第 ? 周' : `第 ${week} 周`;
  document.getElementById('weekday-el').textContent = DAYS[todayColumn() - 1];
  document.getElementById('date-el').textContent = todayStr();
}

function scrollToToday() {
  const wrap = document.querySelector('.schedule-wrap');
  const th = scheduleEl.querySelector('th.today');
  if (wrap && th) {
    wrap.scrollLeft += th.getBoundingClientRect().left - wrap.getBoundingClientRect().left - 66;
  }
}

let firstRender = true;

function render() {
  const today = todayColumn();
  const week = currentWeek();
  const nowSlot = currentSlot();
  const visible = (c) => week == null || c.weeks.includes(week);

  updateSubbar(week);

  let html = '<thead><tr><th class="corner">时间</th>';
  DAYS.forEach((d, i) => {
    const col = i + 1;
    const cls = col === today ? ' class="today"' : '';
    const badge = col === today ? '<span class="badge">今</span>' : '';
    html += `<th${cls}>${d}${badge}</th>`;
  });
  html += '</tr></thead><tbody>';

  state.timeSlots.forEach((slot) => {
    const isNow = slot.no === nowSlot;
    const nowMark = isNow ? '<span class="now-badge">进行中</span>' : '';
    html += `<tr><td class="time${isNow ? ' now' : ''}"><span class="slot-no">第${slot.no}节</span><span class="slot-range">${slot.start}–${slot.end}</span>${nowMark}</td>`;
    for (let col = 1; col <= 7; col++) {
      const classes = [];
      if (col === today) classes.push('today-col');
      if (col === today && isNow) classes.push('now');
      const cellCls = classes.length ? ` class="${classes.join(' ')}"` : '';
      const courses = state.courses.filter((c) => c.day === col && c.slot === slot.no && visible(c));
      html += `<td${cellCls} data-day="${col}" data-slot="${slot.no}">${cellHTML(courses)}</td>`;
    }
    html += '</tr>';
  });

  html += '</tbody>';
  scheduleEl.innerHTML = html;

  if (firstRender) {
    scrollToToday();
    firstRender = false;
  }
}

// —— 课程弹窗与表单 ——
const modal = document.getElementById('modal');
const modalTitle = document.getElementById('modal-title');
const form = document.getElementById('course-form');
const daySelect = form.elements.day;
const slotSelect = form.elements.slot;
const weekTypeSelect = form.elements.weekType;
const customWeeks = document.getElementById('custom-weeks');
const deleteBtn = document.getElementById('delete-btn');

DAYS.forEach((d, i) => daySelect.add(new Option(d, String(i + 1))));
state.timeSlots.forEach((s) => slotSelect.add(new Option(`第${s.no}节 ${s.start}–${s.end}`, String(s.no))));

weekTypeSelect.addEventListener('change', () => {
  customWeeks.hidden = weekTypeSelect.value !== 'custom';
});

let editingId = null;

function openAdd(day, slot) {
  editingId = null;
  modalTitle.textContent = '添加课程';
  deleteBtn.hidden = true;
  form.reset();
  daySelect.value = String(day);
  slotSelect.value = String(slot);
  weekTypeSelect.value = 'full';
  customWeeks.hidden = true;
  modal.hidden = false;
  form.elements.name.focus();
}

function weeksTypeOf(weeks) {
  if (sameWeeks(weeks, fullWeeks())) return 'full';
  if (sameWeeks(weeks, oddWeeks())) return 'odd';
  if (sameWeeks(weeks, evenWeeks())) return 'even';
  return 'custom';
}

function openEdit(id) {
  const c = state.courses.find((x) => x.id === id);
  if (!c) return;
  editingId = id;
  modalTitle.textContent = '编辑课程';
  deleteBtn.hidden = false;
  form.elements.name.value = c.name;
  form.elements.teacher.value = c.teacher || '';
  form.elements.room.value = c.room || '';
  daySelect.value = String(c.day);
  slotSelect.value = String(c.slot);
  const type = weeksTypeOf(c.weeks);
  weekTypeSelect.value = type;
  if (type === 'custom') {
    form.elements.weekStart.value = Math.min(...c.weeks);
    form.elements.weekEnd.value = Math.max(...c.weeks);
  }
  customWeeks.hidden = type !== 'custom';
  modal.hidden = false;
}

function weeksFromForm() {
  const type = weekTypeSelect.value;
  if (type === 'full') return fullWeeks();
  if (type === 'odd') return oddWeeks();
  if (type === 'even') return evenWeeks();
  let s = parseInt(form.elements.weekStart.value, 10) || 1;
  let e = parseInt(form.elements.weekEnd.value, 10) || s;
  s = Math.max(1, Math.min(TOTAL_WEEKS, s));
  e = Math.max(s, Math.min(TOTAL_WEEKS, e));
  return range(s, e);
}

function closeModal() {
  modal.hidden = true;
  editingId = null;
}

form.addEventListener('submit', (e) => {
  e.preventDefault();
  const data = {
    name: form.elements.name.value.trim(),
    teacher: form.elements.teacher.value.trim(),
    room: form.elements.room.value.trim(),
    day: parseInt(daySelect.value, 10),
    slot: parseInt(slotSelect.value, 10),
    weeks: weeksFromForm(),
  };
  if (!data.name) return;

  if (editingId) {
    const c = state.courses.find((x) => x.id === editingId);
    if (c) Object.assign(c, data);
  } else {
    state.courses.push({ id: uid(), ...data });
  }
  save(state);
  render();
  closeModal();
});

deleteBtn.addEventListener('click', () => {
  if (!editingId) return;
  state.courses = state.courses.filter((x) => x.id !== editingId);
  save(state);
  render();
  closeModal();
});

// —— 设置弹窗（开学日期 + 作息时间） ——
const settingsModal = document.getElementById('settings-modal');
const settingsForm = document.getElementById('settings-form');
const slotEditor = document.getElementById('slot-editor');

function renderSlotEditor() {
  slotEditor.innerHTML = state.timeSlots.map((s) => `
    <div class="slot-row">
      <span class="slot-label">第${s.no}节</span>
      <input type="time" value="${s.start}" data-slot="${s.no}" data-field="start" />
      <span class="slot-sep">–</span>
      <input type="time" value="${s.end}" data-slot="${s.no}" data-field="end" />
    </div>`).join('');
}

document.getElementById('settings-btn').addEventListener('click', () => {
  settingsForm.elements.termStart.value = state.termStart || '';
  renderSlotEditor();
  settingsModal.hidden = false;
});

document.getElementById('settings-cancel').addEventListener('click', () => {
  settingsModal.hidden = true;
});

settingsModal.addEventListener('click', (e) => {
  if (e.target === settingsModal) settingsModal.hidden = true;
});

settingsForm.addEventListener('submit', (e) => {
  e.preventDefault();
  state.termStart = settingsForm.elements.termStart.value;

  const newSlots = state.timeSlots.map((s) => {
    const start = slotEditor.querySelector(`input[data-slot="${s.no}"][data-field="start"]`).value;
    const end = slotEditor.querySelector(`input[data-slot="${s.no}"][data-field="end"]`).value;
    return { no: s.no, start: start || s.start, end: end || s.end };
  });
  for (const s of newSlots) {
    if (s.start && s.end && s.start >= s.end) {
      alert(`第${s.no}节：结束时间需晚于开始时间`);
      return;
    }
  }
  state.timeSlots = newSlots;

  save(state);
  render();
  settingsModal.hidden = true;
});

// —— 课程详情弹窗（折叠后点击展开） ——
const detailModal = document.getElementById('detail-modal');
const detailTitle = document.getElementById('detail-title');
const detailList = document.getElementById('detail-list');
let detailDay = 1;
let detailSlot = 1;

function detailItemHTML(c) {
  const meta = [c.teacher, cleanRoom(c.room)].filter(Boolean).join(' · ');
  return `
    <div class="detail-item">
      <div class="detail-head">
        <span class="detail-name">${escapeHTML(c.name)}</span>
        <button class="mini-btn" data-edit="${c.id}">编辑</button>
      </div>
      ${meta ? `<div class="detail-meta">${escapeHTML(meta)}</div>` : ''}
      <span class="tag">${escapeHTML(formatWeeks(c.weeks))}</span>
    </div>`;
}

function openCellDetail(day, slot) {
  detailDay = day;
  detailSlot = slot;
  const week = currentWeek();
  const visible = (c) => week == null || c.weeks.includes(week);
  const courses = state.courses.filter((c) => c.day === day && c.slot === slot && visible(c));
  detailTitle.textContent = `${DAYS[day - 1]} 第${slot}节`;
  detailList.innerHTML = courses.map(detailItemHTML).join('') || '<p class="detail-empty">暂无课程</p>';
  detailModal.hidden = false;
}

detailList.addEventListener('click', (e) => {
  const btn = e.target.closest('[data-edit]');
  if (!btn) return;
  detailModal.hidden = true;
  openEdit(btn.dataset.edit);
});

document.getElementById('detail-close').addEventListener('click', () => {
  detailModal.hidden = true;
});

document.getElementById('detail-add').addEventListener('click', () => {
  detailModal.hidden = true;
  openAdd(detailDay, detailSlot);
});

detailModal.addEventListener('click', (e) => {
  if (e.target === detailModal) detailModal.hidden = true;
});

// —— 交互绑定 ——
scheduleEl.addEventListener('click', (e) => {
  const moreEl = e.target.closest('.more');
  if (moreEl) {
    const cell = moreEl.closest('td[data-day]');
    openCellDetail(+cell.dataset.day, +cell.dataset.slot);
    return;
  }
  const courseEl = e.target.closest('.course');
  if (courseEl) {
    const cell = courseEl.closest('td[data-day]');
    openCellDetail(+cell.dataset.day, +cell.dataset.slot);
    return;
  }
  const cell = e.target.closest('td[data-day]');
  if (cell) openAdd(+cell.dataset.day, +cell.dataset.slot);
});

document.getElementById('add-btn').addEventListener('click', () => {
  openAdd(todayColumn(), 1);
});

document.getElementById('cancel-btn').addEventListener('click', closeModal);

modal.addEventListener('click', (e) => {
  if (e.target === modal) closeModal();
});

// —— Excel 导入 ——
function parseWeeks(str) {
  if (!str) return fullWeeks();
  const s = String(str).replace(/周$/, '').replace(/，/g, ',');
  const weeks = [];
  for (const part of s.split(',')) {
    const p = part.trim();
    if (!p) continue;
    const m = p.match(/^(\d+)\s*-\s*(\d+)$/);
    if (m) {
      for (let w = +m[1]; w <= +m[2]; w++) weeks.push(w);
    } else {
      const n = parseInt(p, 10);
      if (!isNaN(n)) weeks.push(n);
    }
  }
  return [...new Set(weeks)].sort((a, b) => a - b);
}

// 解析一个格子：每门课占两行 —— 第一行课程名，第二行「老师;周次;教室」
function parseCell(text, day, slot) {
  const courses = [];
  if (!text) return courses;
  const lines = String(text).split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
  for (let i = 0; i + 1 < lines.length; i += 2) {
    const name = lines[i];
    const parts = lines[i + 1].split(';');
    courses.push({
      name,
      teacher: (parts[0] || '').trim(),
      weeks: parseWeeks((parts[1] || '').trim()),
      room: (parts[2] || '').trim(),
      day,
      slot,
    });
  }
  return courses;
}

function parseSchedule(rows) {
  const courses = [];
  let headerIdx = -1;
  rows.forEach((r, i) => {
    if (r.some((c) => String(c).includes('星期一'))) headerIdx = i;
  });
  if (headerIdx < 0) return courses;

  let slot = 0;
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const label = String(row[0] || '');
    if (!/大节/.test(label)) continue;
    slot += 1;
    if (slot > state.timeSlots.length) break;
    for (let d = 1; d <= 7; d++) {
      courses.push(...parseCell(row[d], d, slot));
    }
  }
  return courses;
}

const importInput = document.getElementById('import-input');

document.getElementById('import-btn').addEventListener('click', () => importInput.click());

importInput.addEventListener('change', () => {
  const file = importInput.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const courses = parseSchedule(rows);

      if (courses.length === 0) {
        alert('没有解析到课程，请确认这是教务系统导出的课表文件。');
        return;
      }
      if (state.courses.length > 0 && !confirm(`将导入 ${courses.length} 门课程并替换当前课表，是否继续？`)) {
        return;
      }
      state.courses = courses.map((c) => ({ id: uid(), ...c }));
      save(state);
      render();
      alert(`成功导入 ${courses.length} 门课程！`);
    } catch (err) {
      alert('导入失败：' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
  importInput.value = '';
});

render();
