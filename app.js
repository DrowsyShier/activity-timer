/* ================================================================
 * 版本活动时间规划生成器 app.js
 * 纯原生 JavaScript，无构建工具、无框架依赖
 * 导出 Excel 依赖本地 libs/xlsx.full.min.js（需先于本文件加载）
 * ================================================================ */

/* ==== 1. 常量与色板 ==== */
// 6 种预置活动类型。色板来自 dataviz 规范（固定顺序，已通过色盲/对比度校验）；
// 块上文字颜色由 inkFor() 按明度自动选择白/黑，保证可读性。
const TYPE_DEFS = [
  { key: 'sign',     label: '签到活动',     color: '#2a78d6' },
  { key: 'festival', label: '节日/庆典活动', color: '#eb6834' },
  { key: 'recharge', label: '充值/累充活动', color: '#1baf7a' },
  { key: 'preheat',  label: '预热活动',     color: '#eda100' },
  { key: 'rotation', label: '常规轮转活动', color: '#e87ba4' },
  { key: 'limited',  label: '限时活动',     color: '#008300' },
];
const PX_PER_DAY = 36;   // 甘特图：每天像素宽度
const LANE_W = 190;      // 甘特图：左侧泳道标签宽度
const MAX_HISTORY = 20;  // 撤销栈上限

const typeLabel = key => (TYPE_DEFS.find(t => t.key === key) || {}).label || key;
const typeColor = key => (TYPE_DEFS.find(t => t.key === key) || {}).color || '#898781';

// 特殊日期标注（节日表）：md 为每年固定的阳历节日；date 为农历节日（已内置 2025-2028 年，可自行增补）
const HOLIDAYS = [
  { name: '元旦', md: '01-01' },
  { name: '情人节', md: '02-14' },
  { name: '妇女节', md: '03-08' },
  { name: '劳动节', md: '05-01' },
  { name: '儿童节', md: '06-01' },
  { name: '国庆节', md: '10-01' },
  { name: '万圣夜', md: '10-31' },
  { name: '双十一', md: '11-11' },
  { name: '双十二', md: '12-12' },
  { name: '平安夜', md: '12-24' },
  { name: '圣诞节', md: '12-25' },
  { name: '除夕', date: '2025-01-28' }, { name: '春节', date: '2025-01-29' }, { name: '元宵节', date: '2025-02-12' },
  { name: '端午节', date: '2025-05-31' }, { name: '中秋节', date: '2025-10-06' },
  { name: '除夕', date: '2026-02-16' }, { name: '春节', date: '2026-02-17' }, { name: '元宵节', date: '2026-03-03' },
  { name: '端午节', date: '2026-06-19' }, { name: '中秋节', date: '2026-09-25' },
  { name: '除夕', date: '2027-02-05' }, { name: '春节', date: '2027-02-06' }, { name: '元宵节', date: '2027-02-20' },
  { name: '端午节', date: '2027-06-09' }, { name: '中秋节', date: '2027-09-15' },
  { name: '除夕', date: '2028-01-25' }, { name: '春节', date: '2028-01-26' }, { name: '元宵节', date: '2028-02-04' },
  { name: '端午节', date: '2028-05-28' }, { name: '中秋节', date: '2028-10-03' },
];

// 单日几个活动算"密集"：每日上限 >= 3 时取 3；上限为 2 时取 2（满员即密集）；上限为 1 时永不触发
function denseThreshold() {
  const limit = AppState.rules.allowOverlap ? AppState.rules.dailyCap : 1;
  return limit >= 3 ? 3 : 2;
}

/* ==== 2. 工具函数 ==== */
let seq = 0;
const uid = prefix => prefix + '-' + (++seq) + '-' + Date.now().toString(36);

const clampInt = (v, lo, hi) => {
  const n = parseInt(v, 10);
  if (isNaN(n)) return lo;
  return Math.max(lo, Math.min(hi, n));
};

// 解析 'YYYY-MM-DD'（手动拆分，避免 new Date('YYYY-MM-DD') 的时区坑）
function parseDateStr(s) {
  if (!s) return null;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return { y: +m[1], m: +m[2], d: +m[3] };
}

// 版本开始日期 + 天索引 → Date 对象（天索引从 0 开始）
function dateAt(startStr, dayIndex) {
  const p = parseDateStr(startStr) || { y: 2026, m: 9, d: 1 };
  const dt = new Date(p.y, p.m - 1, p.d);
  dt.setDate(dt.getDate() + dayIndex);
  return dt;
}

const pad2 = n => String(n).padStart(2, '0');
const fmtDate = dt => dt.getFullYear() + '-' + pad2(dt.getMonth() + 1) + '-' + pad2(dt.getDate());
const fmtMD = dt => (dt.getMonth() + 1) + '/' + dt.getDate();
const WEEK_CN = ['日', '一', '二', '三', '四', '五', '六'];
const isWeekend = dt => dt.getDay() === 0 || dt.getDay() === 6;

// 查询某天是否为特殊日期，返回节日名（无则 null）
function holidayFor(dt) {
  const md = pad2(dt.getMonth() + 1) + '-' + pad2(dt.getDate());
  const full = fmtDate(dt);
  for (const h of HOLIDAYS) {
    if (h.date === full || h.md === md) return h.name;
  }
  return null;
}

// 两个日期之间的天数差（按 UTC 口径计算，避免夏令时误差）
function dayDiff(a, b) {
  return Math.round((Date.UTC(a.getFullYear(), a.getMonth(), a.getDate())
    - Date.UTC(b.getFullYear(), b.getMonth(), b.getDate())) / 86400000);
}

// HTML 转义（活动名称是用户输入，必须转义）
function escHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

const deepCopy = obj => JSON.parse(JSON.stringify(obj));

// 把 0 基天索引数组（升序）格式化为"第1-3、5、8-9天"样式的区间文本
function fmtDayRanges(days) {
  const parts = [];
  let s = 0;
  while (s < days.length) {
    let e = s;
    while (e + 1 < days.length && days[e + 1] === days[e] + 1) e++;
    parts.push(e === s ? '第' + (days[s] + 1) + '天' : '第' + (days[s] + 1) + '-' + (days[e] + 1) + '天');
    s = e + 1;
  }
  return parts.join('、');
}

// 根据颜色明度决定块上文字用白还是黑（保证文字与底色对比度 >= 3:1）
function inkFor(hex) {
  const h = String(hex).replace('#', '');
  if (!/^[0-9a-fA-F]{6}$/.test(h)) return '#ffffff';
  const r = parseInt(h.slice(0, 2), 16) / 255;
  const g = parseInt(h.slice(2, 4), 16) / 255;
  const b = parseInt(h.slice(4, 6), 16) / 255;
  const lin = c => c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  const L = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
  return L <= 0.28 ? '#ffffff' : '#0b0b0b';
}

/* ==== 3. 全局状态与默认数据 ==== */
const AppState = {
  version: { startDate: '2026-09-01', totalDays: 30 },
  rules: { allowOverlap: true, dailyCap: 3 },
  activities: [
    { id: 'a2', name: '周年庆典', type: 'festival', duration: 10, count: 1, intervalDays: 0, color: '#eb6834' },
    { id: 'a1', name: '每日签到', type: 'sign',     duration: 7,  count: 2, intervalDays: 0, color: '#2a78d6' },
    { id: 'a5', name: '轮转挑战', type: 'rotation', duration: 3,  count: 3, intervalDays: 4, color: '#e87ba4' },
    { id: 'a3', name: '累充返利', type: 'recharge', duration: 7,  count: 2, intervalDays: 3, color: '#1baf7a' },
    { id: 'a6', name: '限时商店', type: 'limited',  duration: 7,  count: 2, intervalDays: 3, color: '#008300' },
    { id: 'a4', name: '版本预热', type: 'preheat',  duration: 5,  count: 1, intervalDays: 0, color: '#eda100' },
  ],
  schedule: { instances: [], overflow: [] },
  history: [],
  ui: { editingId: null, drag: null },
};

/* ==== 4. DOM 引用 ==== */
const $ = id => document.getElementById(id);
const elStartDate = $('input-start-date'), elTotalDays = $('input-total-days'), elVersionSummary = $('version-summary');
const elAllowOverlap = $('input-allow-overlap'), elDailyCap = $('input-daily-cap');
const elActivityList = $('activity-list'), elAddActivity = $('btn-add-activity');
const elGenerate = $('btn-generate'), elIncremental = $('btn-incremental'), elUndo = $('btn-undo'), elExport = $('btn-export');
const elWarningBanner = $('warning-banner');
const elStatsBar = $('stats-bar');
const elGanttLegend = $('gantt-legend'), elGanttScroll = $('gantt-scroll'), elGanttInner = $('gantt-inner');
const elSuggestionList = $('suggestion-list');
const elPanelConflicts = $('panel-conflicts'), elConflictList = $('conflict-list');
const elModal = $('edit-modal'), elEditInfo = $('edit-info'), elEditStart = $('edit-start'), elEditDuration = $('edit-duration');
const elEditDelete = $('edit-delete'), elEditCancel = $('edit-cancel'), elEditOk = $('edit-ok');

/* ==== 5. 输入区渲染 ==== */
function renderActivities() {
  elActivityList.innerHTML = AppState.activities.map(a => {
    const typeOpts = TYPE_DEFS.map(t =>
      '<option value="' + t.key + '"' + (t.key === a.type ? ' selected' : '') + '>' + t.label + '</option>').join('');
    return '<div class="activity-item" data-id="' + escHtml(a.id) + '">'
      + '<input type="text" class="act-name" value="' + escHtml(a.name) + '" title="活动名称">'
      + '<select class="act-type" title="活动类型">' + typeOpts + '</select>'
      + '<label>持续 <input type="number" class="act-duration" min="1" max="90" value="' + a.duration + '"> 天</label>'
      + '<label>数量 <input type="number" class="act-count" min="1" max="30" value="' + a.count + '"> 期</label>'
      + '<label>间隔 <input type="number" class="act-interval" min="0" max="30" value="' + a.intervalDays + '" title="同一活动相邻两期之间的最小间隔天数"> 天</label>'
      + '<input type="color" class="act-color" value="' + a.color + '" title="颜色">'
      + '<button class="act-delete btn-danger">删除</button>'
      + '</div>';
  }).join('');
}

function renderToolbar() {
  elDailyCap.disabled = !AppState.rules.allowOverlap;
  elUndo.disabled = AppState.history.length === 0;
  // 补排按钮：所有活动都已排满时禁用
  const placedCount = new Map();
  for (const i of AppState.schedule.instances) placedCount.set(i.activityId, (placedCount.get(i.activityId) || 0) + 1);
  elIncremental.disabled = !AppState.activities.some(a => a.count > (placedCount.get(a.id) || 0));
}

function renderVersionSummary() {
  const { startDate, totalDays } = AppState.version;
  if (!parseDateStr(startDate)) { elVersionSummary.textContent = ''; return; }
  elVersionSummary.textContent = '→ 结束日期 ' + fmtDate(dateAt(startDate, totalDays - 1)) + '（共 ' + totalDays + ' 天）';
}

/* ==== 6. 排期算法 ==== */
// 每日负载：每天正在进行中的活动数
function computeDayLoad(instances, totalDays) {
  const load = new Array(totalDays).fill(0);
  for (const i of instances)
    for (let d = Math.max(0, i.startDay); d < i.startDay + i.duration && d < totalDays; d++) load[d]++;
  return load;
}

// 判断能否把某个实例放在第 day 天开始（load 为"当前已排入实例"的每日负载；
// intervalDays 为该活动自己设置的最小间隔，控制与相邻同活动实例的距离）
function canPlaceWith(day, duration, activityId, intervalDays, instances, load, totalDays, rules) {
  if (day < 0 || day + duration > totalDays) return false;
  const limit = rules.allowOverlap ? rules.dailyCap : 1;
  for (let d = day; d < day + duration; d++) if (load[d] >= limit) return false;
  // 同类活动最小间隔：与该活动其他已排实例之间至少隔 intervalDays 天
  if (intervalDays > 0) {
    for (const s of instances) {
      if (s.activityId !== activityId) continue;
      const gap = s.startDay > day
        ? s.startDay - (day + duration - 1) - 1
        : day - (s.startDay + s.duration - 1) - 1;
      if (gap < intervalDays) return false;
    }
  }
  return true;
}

function makeInstance(a, day, period) {
  return { id: uid('i'), activityId: a.id, name: a.name, type: a.type, color: a.color,
           startDay: day, duration: a.duration, intervalDays: a.intervalDays, period, manual: false };
}

// 从 anchorStart 起向前贪心寻找最早合法位置（找到后继续向后推进）
function placeActivityCompact(a, instances, overflow, load, totalDays, r, anchorStart) {
  let cursor = Math.min(anchorStart, totalDays - a.duration);
  let placed = 0;
  while (placed < a.count) {
    let found = -1;
    for (let d = cursor; d + a.duration <= totalDays; d++) {
      if (canPlaceWith(d, a.duration, a.id, a.intervalDays, instances, load, totalDays, r)) { found = d; break; }
    }
    if (found < 0) {
      overflow.push({ id: uid('o'), activityId: a.id, name: a.name, type: a.type, color: a.color,
                      duration: a.duration, period: placed + 1, reason: '周期容量不足或间隔约束无法满足' });
      break;
    }
    const inst = makeInstance(a, found, placed + 1);
    instances.push(inst);
    for (let d = found; d < found + a.duration; d++) load[d]++;
    placed++;
    // 更早的天已被占用或被同活动间隔约束挡住，直接跳过
    cursor = found + a.duration + a.intervalDays;
  }
}

// 从版本结束时间往前排：第一期待在版本末尾，后续期继续向前递推
function placeActivityFromEnd(a, instances, overflow, load, totalDays, r) {
  let cursorEnd = totalDays; // 下一期的结束日（不含当天）
  let placed = 0;
  while (placed < a.count) {
    const latestStart = Math.min(cursorEnd, totalDays) - a.duration;
    let found = -1;
    for (let d = latestStart; d >= 0; d--) {
      if (canPlaceWith(d, a.duration, a.id, a.intervalDays, instances, load, totalDays, r)) { found = d; break; }
    }
    if (found < 0) {
      overflow.push({ id: uid('o'), activityId: a.id, name: a.name, type: a.type, color: a.color,
                      duration: a.duration, period: placed + 1, reason: '周期容量不足或间隔约束无法满足' });
      break;
    }
    const inst = makeInstance(a, found, placed + 1);
    instances.push(inst);
    for (let d = found; d < found + a.duration; d++) load[d]++;
    placed++;
    // 下一期必须在本期开始前至少 intervalDays 天结束
    cursorEnd = found - a.intervalDays;
  }
}

// 紧凑排列，内置三条默认节奏：
// ① 周年庆典优先排入；② 累充活动锚定周年活动开启日（同时开启）；③ 预热活动从版本结束时间往前排
function generateCompact() {
  const { totalDays } = AppState.version, r = AppState.rules;
  const instances = [], overflow = [], load = new Array(totalDays).fill(0);
  const list = AppState.activities;
  for (const a of list) if (a.type === 'festival') placeActivityCompact(a, instances, overflow, load, totalDays, r, 0);
  for (const a of list) if (a.type === 'recharge') {
    const fest = instances.find(i => i.type === 'festival');
    placeActivityCompact(a, instances, overflow, load, totalDays, r, fest ? fest.startDay : 0);
  }
  for (const a of list) {
    if (a.type !== 'festival' && a.type !== 'recharge' && a.type !== 'preheat')
      placeActivityCompact(a, instances, overflow, load, totalDays, r, 0);
  }
  for (const a of list) if (a.type === 'preheat') placeActivityFromEnd(a, instances, overflow, load, totalDays, r);
  return { instances, overflow };
}

function generateSchedule(withHistory) {
  if (!parseDateStr(AppState.version.startDate)) { alert('请先选择版本开始日期'); return; }
  // 有手动调整时提醒：全量重排会覆盖它们（想保留请用"补排新活动"）
  if (withHistory && AppState.schedule.instances.some(i => i.manual)
      && !confirm('重新生成将覆盖你已手动调整过的排期，确定继续吗？\n（只想补充新活动，请取消后点「＋ 补排新活动」）')) return;
  if (withHistory) pushUndo();
  AppState.schedule = generateCompact();
  AppState.ui.editingId = null;
  renderAll();
}

// 增量补排：保留现有排期（含手动调整），只补排缺失的活动实例
function scheduleIncremental() {
  const { totalDays } = AppState.version, r = AppState.rules;
  const instances = AppState.schedule.instances.slice();
  const overflow = [];
  const load = computeDayLoad(instances, totalDays);
  const placedCount = new Map();
  let maxPeriodOf = aId => {
    let mp = 0;
    for (const i of instances) if (i.activityId === aId) mp = Math.max(mp, i.period);
    return mp;
  };
  for (const i of instances) placedCount.set(i.activityId, (placedCount.get(i.activityId) || 0) + 1);

  // 与全量生成相同的节奏：festival → recharge（锚定周年开启日）→ 其他 → preheat（末尾往前）
  const placeMissing = (a, anchorStart) => {
    const missing = a.count - (placedCount.get(a.id) || 0);
    if (missing <= 0) return;
    const periodBase = Math.max(placedCount.get(a.id) || 0, maxPeriodOf(a.id));
    let cursor = Math.min(anchorStart, totalDays - a.duration);
    let placed = 0;
    while (placed < missing) {
      let found = -1;
      for (let d = cursor; d + a.duration <= totalDays; d++) {
        if (canPlaceWith(d, a.duration, a.id, a.intervalDays, instances, load, totalDays, r)) { found = d; break; }
      }
      if (found < 0) {
        overflow.push({ id: uid('o'), activityId: a.id, name: a.name, type: a.type, color: a.color,
                        duration: a.duration, period: periodBase + placed + 1,
                        reason: '周期容量不足或间隔约束无法满足' });
        break;
      }
      const inst = makeInstance(a, found, periodBase + placed + 1);
      instances.push(inst);
      for (let d = found; d < found + a.duration; d++) load[d]++;
      placed++;
      cursor = found + a.duration + a.intervalDays;
    }
  };
  const placeMissingFromEnd = a => {
    const missing = a.count - (placedCount.get(a.id) || 0);
    if (missing <= 0) return;
    const periodBase = Math.max(placedCount.get(a.id) || 0, maxPeriodOf(a.id));
    let cursorEnd = totalDays;
    let placed = 0;
    while (placed < missing) {
      const latestStart = Math.min(cursorEnd, totalDays) - a.duration;
      let found = -1;
      for (let d = latestStart; d >= 0; d--) {
        if (canPlaceWith(d, a.duration, a.id, a.intervalDays, instances, load, totalDays, r)) { found = d; break; }
      }
      if (found < 0) {
        overflow.push({ id: uid('o'), activityId: a.id, name: a.name, type: a.type, color: a.color,
                        duration: a.duration, period: periodBase + placed + 1,
                        reason: '周期容量不足或间隔约束无法满足' });
        break;
      }
      const inst = makeInstance(a, found, periodBase + placed + 1);
      instances.push(inst);
      for (let d = found; d < found + a.duration; d++) load[d]++;
      placed++;
      cursorEnd = found - a.intervalDays;
    }
  };

  const beforeLen = instances.length, beforeOver = AppState.schedule.overflow.length;
  for (const a of AppState.activities) if (a.type === 'festival') placeMissing(a, 0);
  for (const a of AppState.activities) if (a.type === 'recharge') {
    const fest = instances.find(i => i.type === 'festival');
    placeMissing(a, fest ? fest.startDay : 0);
  }
  for (const a of AppState.activities) {
    if (a.type !== 'festival' && a.type !== 'recharge' && a.type !== 'preheat') placeMissing(a, 0);
  }
  for (const a of AppState.activities) if (a.type === 'preheat') placeMissingFromEnd(a);

  if (instances.length !== beforeLen || overflow.length !== beforeOver) pushUndo();
  AppState.schedule = { instances, overflow };
  AppState.ui.editingId = null;
  renderAll();
}

// 容量预检（输入变化时实时提示）
function capacityWarnings() {
  const ws = [];
  const { totalDays } = AppState.version, r = AppState.rules;
  let totalSlots = 0;
  for (const a of AppState.activities) {
    totalSlots += a.duration * a.count;
    if (a.duration > totalDays) ws.push('「' + a.name + '」单次持续 ' + a.duration + ' 天，超过版本周期 ' + totalDays + ' 天');
    const need = a.duration * a.count + a.intervalDays * (a.count - 1);
    if (need > totalDays) ws.push('「' + a.name + '」共 ' + a.count + ' 期，按最小间隔至少需要 ' + need + ' 天，超出周期长度（' + totalDays + ' 天）');
  }
  const capacity = totalDays * (r.allowOverlap ? r.dailyCap : 1);
  if (totalSlots > capacity) ws.push('活动总时长 ' + totalSlots + ' 天，超过周期容量 ' + capacity + ' 天（' + totalDays + ' 天 × 每日上限 ' + (r.allowOverlap ? r.dailyCap : 1) + '），必然溢出');
  return ws;
}

function renderWarningBanner() {
  const parts = [];
  for (const w of capacityWarnings()) parts.push('<div class="banner-item">⚠ ' + escHtml(w) + '</div>');
  for (const o of AppState.schedule.overflow) parts.push('<div class="banner-item">⛔ 「' + escHtml(o.name) + ' 第' + o.period + '期」未排入：' + escHtml(o.reason) + '</div>');
  elWarningBanner.classList.toggle('hidden', parts.length === 0);
  elWarningBanner.classList.toggle('has-error', AppState.schedule.overflow.length > 0);
  elWarningBanner.innerHTML = parts.length
    ? '<div class="banner-title">' + (AppState.schedule.overflow.length ? '排期警告' : '容量提醒') + '</div>' + parts.join('')
    : '';
}

/* ==== 7. 甘特图渲染 ==== */
// 行序 = 活动列表顺序 → 开始天 → 期数（排序稳定，编辑时行不跳动）
function sortedInstances() {
  const order = new Map(AppState.activities.map((a, idx) => [a.id, idx]));
  return AppState.schedule.instances.slice().sort((x, y) =>
    (order.get(x.activityId) ?? 999) - (order.get(y.activityId) ?? 999)
    || x.startDay - y.startDay || x.period - y.period
    || (x.id < y.id ? -1 : 1));
}

function renderGantt(conflicts) {
  const { startDate, totalDays } = AppState.version;
  const insts = sortedInstances();

  if (!insts.length) {
    elGanttInner.innerHTML = '<div class="gantt-empty">暂无排期数据 — 请点击「⚙ 自动生成排期」</div>';
    return;
  }

  const conflictIds = new Set();
  for (const c of conflicts) for (const id of c.instIds) conflictIds.add(id);

  // 日期轴（跨月/首日显示 M/D，其余只显示日号）+ 周末底色
  let axisCells = '', bgCells = '';
  for (let d = 0; d < totalDays; d++) {
    const dt = dateAt(startDate, d);
    const wCls = isWeekend(dt) ? ' is-weekend' : '';
    const dateText = (d === 0 || dt.getDate() === 1) ? fmtMD(dt) : String(dt.getDate());
    const hol = holidayFor(dt);
    axisCells += '<div class="axis-cell' + wCls + '"' + (hol ? ' title="' + hol + '"' : '') + '>'
      + dateText + '<span class="weekday">' + WEEK_CN[dt.getDay()] + '</span>'
      + (hol ? '<span class="holiday">' + hol + '</span>' : '')
      + '</div>';
    bgCells += '<div class="bg-cell' + wCls + '"></div>';
  }

  // 密度色带（0 空闲 / 1..DENSE-1 适中 / >=DENSE 密集）
  const DENSE = denseThreshold();
  const load = computeDayLoad(insts, totalDays);
  let densityCells = '';
  for (let d = 0; d < totalDays; d++) {
    const n = load[d];
    const dt = dateAt(startDate, d);
    const lvl = n === 0 ? '空闲' : n < DENSE ? '适中' : '密集';
    densityCells += '<div class="density-cell d' + (n === 0 ? 0 : n < DENSE ? 1 : 2)
      + '" title="' + fmtMD(dt) + ' · ' + n + ' 个活动 · ' + lvl + '">' + (n > 0 ? n : '') + '</div>';
  }

  // 泳道：每个实例一行
  let rows = '';
  for (const inst of insts) {
    const left = inst.startDay * PX_PER_DAY;
    const width = inst.duration * PX_PER_DAY;
    const cls = 'gantt-block' + (conflictIds.has(inst.id) ? ' conflict' : '') + (inst.duration < 3 ? ' narrow' : '');
    rows += '<div class="gantt-row">'
      + '<div class="lane-label" title="' + escHtml(inst.name) + ' 第' + inst.period + '期">'
      + '<span class="lane-name">' + escHtml(inst.name) + '</span><span class="period">第' + inst.period + '期</span></div>'
      + '<div class="gantt-track" style="width:' + (totalDays * PX_PER_DAY) + 'px">'
      + '<div class="' + cls + '" data-id="' + inst.id + '"'
      + ' title="' + escHtml(inst.name) + ' 第' + inst.period + '期 · ' + inst.duration + ' 天 · 拖拽移动 / 点击编辑"'
      + ' style="left:' + left + 'px;width:' + width + 'px;background:' + inst.color + ';color:' + inkFor(inst.color) + '">'
      + '<span class="block-name">' + escHtml(inst.name) + '</span><span class="block-days">' + inst.duration + '天</span>'
      + '</div></div></div>';
  }

  // 溢出（未排入）幽灵块
  let overflowHtml = '';
  if (AppState.schedule.overflow.length) {
    const ghosts = AppState.schedule.overflow.map(o =>
      '<div class="ghost-block" style="width:' + (o.duration * PX_PER_DAY) + 'px" title="未排入原因：' + escHtml(o.reason) + '">'
      + escHtml(o.name) + ' 第' + o.period + '期</div>').join('');
    overflowHtml = '<div class="gantt-overflow"><span class="overflow-title">未排入（溢出）：</span>' + ghosts + '</div>';
  }

  // 今天标记线
  let todayHtml = '';
  const todayIdx = dayDiff(new Date(), dateAt(startDate, 0));
  if (todayIdx >= 0 && todayIdx < totalDays) {
    todayHtml = '<div class="today-line" style="left:' + (LANE_W + todayIdx * PX_PER_DAY + PX_PER_DAY / 2) + 'px">'
      + '<span class="today-tag">今天</span></div>';
  }

  elGanttInner.innerHTML =
    '<div class="gantt-axis-row"><div class="axis-spacer"></div>' + axisCells + '</div>'
    + '<div class="gantt-density-row"><div class="density-spacer"></div>' + densityCells + '</div>'
    + '<div class="gantt-body">'
    + '<div class="bg-layer"><div class="bg-spacer"></div>' + bgCells + '</div>'
    + rows
    + '</div>'
    + overflowHtml
    + todayHtml;
  elGanttInner.style.width = (LANE_W + totalDays * PX_PER_DAY) + 'px';
}

/* ==== 8. 拖拽与编辑 ==== */
function findInst(id) { return AppState.schedule.instances.find(i => i.id === id); }

function pushUndo() {
  AppState.history.push(deepCopy(AppState.schedule.instances));
  if (AppState.history.length > MAX_HISTORY) AppState.history.shift();
}

function undo() {
  const prev = AppState.history.pop();
  if (!prev) return;
  AppState.schedule.instances = prev;
  renderAll();
}

elGanttScroll.addEventListener('pointerdown', e => {
  if (e.pointerType === 'mouse' && e.button !== 0) return;
  const block = e.target.closest ? e.target.closest('.gantt-block') : null;
  if (!block) return;
  const inst = findInst(block.dataset.id);
  if (!inst) return;
  e.preventDefault(); // 防止选中文字
  AppState.ui.drag = { instId: inst.id, origStartDay: inst.startDay, startX: e.clientX, moved: false, el: block };
  try { block.setPointerCapture(e.pointerId); } catch (err) { /* 忽略 */ }
  block.classList.add('dragging');
});

elGanttScroll.addEventListener('pointermove', e => {
  const drag = AppState.ui.drag;
  if (!drag) return;
  const inst = findInst(drag.instId);
  if (!inst) return;
  const dx = e.clientX - drag.startX;
  if (!drag.moved && Math.abs(dx) < 3) return; // 3px 阈值：区分"点击"与"拖拽"
  drag.moved = true;
  // 拖拽过程中只改样式，不写状态、不重绘
  const newDay = Math.max(0, Math.min(AppState.version.totalDays - inst.duration,
    Math.round(drag.origStartDay + dx / PX_PER_DAY)));
  drag.el.style.left = (newDay * PX_PER_DAY) + 'px';
});

function endDrag(e) {
  const drag = AppState.ui.drag;
  if (!drag) return;
  AppState.ui.drag = null;
  drag.el.classList.remove('dragging');
  const inst = findInst(drag.instId);
  if (!drag.moved) { if (inst) openEditModal(inst.id); return; } // 未移动 = 点击 → 打开编辑
  if (inst) {
    const newDay = Math.max(0, Math.min(AppState.version.totalDays - inst.duration,
      Math.round(drag.origStartDay + (e.clientX - drag.startX) / PX_PER_DAY)));
    if (newDay !== inst.startDay) { pushUndo(); inst.startDay = newDay; inst.manual = true; }
  }
  renderAll(); // 提交 → 校验 → 重绘
}

elGanttScroll.addEventListener('pointerup', endDrag);
elGanttScroll.addEventListener('pointercancel', () => {
  const drag = AppState.ui.drag;
  if (!drag) return;
  AppState.ui.drag = null;
  drag.el.classList.remove('dragging');
  renderAll();
});

/* —— 编辑弹窗 —— */
function openEditModal(id) {
  const inst = findInst(id);
  if (!inst) return;
  AppState.ui.editingId = id;
  elEditInfo.textContent = '「' + inst.name + ' 第' + inst.period + '期」';
  elEditStart.value = inst.startDay + 1;
  elEditDuration.value = inst.duration;
  renderModal();
}
function closeEditModal() { AppState.ui.editingId = null; renderModal(); }
function renderModal() { elModal.classList.toggle('hidden', !AppState.ui.editingId); }

elEditOk.addEventListener('click', () => {
  const inst = findInst(AppState.ui.editingId);
  if (!inst) { closeEditModal(); return; }
  const duration = clampInt(elEditDuration.value, 1, AppState.version.totalDays);
  const maxStart = AppState.version.totalDays - duration;
  const startDay = Math.max(0, Math.min(clampInt(elEditStart.value, 1, AppState.version.totalDays) - 1, maxStart));
  if (startDay !== inst.startDay || duration !== inst.duration) { pushUndo(); }
  inst.startDay = startDay; inst.duration = duration; inst.manual = true;
  closeEditModal();
  renderAll();
});
elEditDelete.addEventListener('click', () => {
  const idx = AppState.schedule.instances.findIndex(i => i.id === AppState.ui.editingId);
  if (idx >= 0) { pushUndo(); AppState.schedule.instances.splice(idx, 1); }
  closeEditModal();
  renderAll();
});
elEditCancel.addEventListener('click', closeEditModal);
elModal.addEventListener('click', e => { if (e.target === elModal) closeEditModal(); });

/* ==== 9. 冲突校验 ==== */
function validateSchedule() {
  const insts = AppState.schedule.instances;
  const { totalDays } = AppState.version, r = AppState.rules;
  const conflicts = [];
  if (!insts.length) return conflicts;
  const name = i => i.name + ' 第' + i.period + '期';
  const limit = r.allowOverlap ? r.dailyCap : 1;

  // ① 边界冲突：超出版本周期
  for (const i of insts) {
    if (i.startDay < 0 || i.startDay + i.duration > totalDays)
      conflicts.push({ type: 'bounds', instIds: [i.id],
        text: '「' + name(i) + '」超出版本周期（第 ' + (i.startDay + 1) + ' 天开始，持续 ' + i.duration + ' 天）' });
  }

  // ② 重叠冲突：逐日检查
  const seen = new Set();
  for (let d = 0; d < totalDays; d++) {
    const actives = insts.filter(i => i.startDay <= d && d < i.startDay + i.duration);
    if (actives.length <= limit) continue;
    if (!r.allowOverlap) {
      for (let a = 0; a < actives.length; a++) for (let b = a + 1; b < actives.length; b++) {
        const key = 'overlap|' + [actives[a].id, actives[b].id].sort().join(',');
        if (seen.has(key)) continue;
        seen.add(key);
        conflicts.push({ type: 'overlap', instIds: [actives[a].id, actives[b].id],
          text: '「' + name(actives[a]) + '」与「' + name(actives[b]) + '」在第 ' + (d + 1) + ' 天重叠（当前不允许重叠）' });
      }
    } else {
      conflicts.push({ type: 'overcap', instIds: actives.map(x => x.id),
        text: '第 ' + (d + 1) + ' 天有 ' + actives.length + ' 个活动同时进行，超过每日上限 ' + limit });
    }
  }

  // ③ 间隔冲突：同活动相邻两期之间
  const byActivity = new Map();
  for (const i of insts) {
    if (!byActivity.has(i.activityId)) byActivity.set(i.activityId, []);
    byActivity.get(i.activityId).push(i);
  }
  for (const list of byActivity.values()) {
    list.sort((a, b) => a.startDay - b.startDay);
    for (let k = 0; k + 1 < list.length; k++) {
      const a = list[k], b = list[k + 1];
      const need = Math.max(a.intervalDays || 0, b.intervalDays || 0);
      const gap = b.startDay - (a.startDay + a.duration - 1) - 1;
      if (gap < need)
        conflicts.push({ type: 'gap', instIds: [a.id, b.id],
          text: '「' + name(a) + '」与「' + name(b) + '」间隔不足（要求 ≥ ' + need + ' 天，实际 ' + gap + ' 天）' });
    }
  }
  return conflicts;
}

function renderConflicts(conflicts) {
  elPanelConflicts.classList.toggle('hidden', conflicts.length === 0);
  elConflictList.innerHTML = conflicts.length
    ? '<p class="conflict-total">共 ' + conflicts.length + ' 处冲突（对应块已在图中标红 ⚠）</p>'
      + conflicts.map(c => '<div class="conflict-item">' + escHtml(c.text) + '</div>').join('')
    : '';
}

/* ==== 10. 统计与建议 ==== */
function renderStats() {
  const insts = AppState.schedule.instances;
  if (!insts.length) {
    elStatsBar.innerHTML = '<div class="stat-tile stat-empty">生成排期后，这里会显示活动密度统计</div>';
    return;
  }
  const { totalDays, startDate } = AppState.version;
  const DENSE = denseThreshold();
  const load = computeDayLoad(insts, totalDays);
  const sum = load.reduce((a, b) => a + b, 0);
  const maxLoad = Math.max(...load);
  const densest = [];
  load.forEach((v, d) => { if (v === maxLoad) densest.push(d); });
  const denseDays = load.filter(v => v >= DENSE).length;
  let longestRun = 0, run = 0;
  for (const v of load) { run = v >= DENSE ? run + 1 : 0; longestRun = Math.max(longestRun, run); }
  const fmtDays = idxs => idxs.slice(0, 3).map(d => fmtMD(dateAt(startDate, d))).join('、') + (idxs.length > 3 ? ' 等' : '');

  const tiles = [
    ['已排入', insts.length + ' 期' + (AppState.schedule.overflow.length ? '（未排入 ' + AppState.schedule.overflow.length + '）' : '')],
    ['平均每日活动', (sum / totalDays).toFixed(1) + ' 个'],
    ['最密集日期', maxLoad ? fmtDays(densest) + '（' + maxLoad + ' 个）' : '—'],
    ['密集天数', denseDays + ' 天（≥' + DENSE + ' 个/天）'],
    ['最长密集连续', longestRun + ' 天'],
  ];
  elStatsBar.innerHTML = tiles.map(t =>
    '<div class="stat-tile"><div class="stat-value">' + escHtml(t[1]) + '</div><div class="stat-label">' + t[0] + '</div></div>').join('');
}

function betterScore(k, best) {
  for (let i = 0; i < k.length; i++) { if (k[i] !== best[i]) return k[i] < best[i]; }
  return false;
}

// 为每个高密度日上的实例搜索"降负载最优"的移动方案；
// 并生成"间隔不足时后续各期整体顺延"的链式调整建议
function computeSuggestions() {
  const insts = AppState.schedule.instances;
  const { totalDays } = AppState.version, r = AppState.rules;
  if (!insts.length) return { list: [], chain: [], guidanceDays: [], healthy: false };
  const DENSE = denseThreshold();
  const limit = r.allowOverlap ? r.dailyCap : 1;
  const load = computeDayLoad(insts, totalDays);
  const denseDays = [];
  load.forEach((v, d) => { if (v >= DENSE) denseDays.push(d); });

  // —— 链式调整建议：与密集无关，只要有间隔违规就生成 ——
  const chain = [];
  const byAct = new Map();
  for (const i of insts) {
    if (!byAct.has(i.activityId)) byAct.set(i.activityId, []);
    byAct.get(i.activityId).push(i);
  }
  for (const group of byAct.values()) {
    if (group.length < 2) continue;
    group.sort((x, y) => x.startDay - y.startDay || x.period - y.period);
    for (let k = 0; k + 1 < group.length; k++) {
      const a = group[k], b = group[k + 1];
      const need = Math.max(a.intervalDays || 0, b.intervalDays || 0);
      const gap = b.startDay - (a.startDay + a.duration - 1) - 1;
      if (gap === need) continue;
      // 方向：间隔不足 → 右移顺延；间隔偏大且经过手动调整 → 左移紧凑
      const shiftRight = gap < need;
      if (!shiftRight && !(a.manual || b.manual)) continue;
      // 从 b 开始，把该活动后续所有期按各自间隔重新排列
      const plan = [];
      let prev = a, feasible = true;
      for (let j = k + 1; j < group.length; j++) {
        const cur = group[j];
        const needJ = Math.max(prev.intervalDays || 0, cur.intervalDays || 0);
        const newStart = prev.startDay + prev.duration + needJ;
        if (shiftRight) {
          if (newStart + cur.duration > totalDays) { feasible = false; break; }
        } else if (newStart >= cur.startDay) break; // 后续各期已无需移动
        plan.push({ instId: cur.id, toDay: newStart });
        prev = { startDay: newStart, duration: cur.duration, intervalDays: cur.intervalDays };
      }
      if (!feasible || !plan.length) {
        if (shiftRight) break; // 右移失败：下游违规对也无法修复，放弃该组
        continue;              // 左移失败：继续检查后面的相邻对
      }
      // 可行性校验：其余实例的负载 + 链上实例
      const movedIds = new Set(plan.map(s => s.instId));
      const restInsts = insts.filter(i => !movedIds.has(i.id));
      const loadC = computeDayLoad(restInsts, totalDays);
      let ok = true;
      for (const step of plan) {
        const cur = insts.find(i => i.id === step.instId);
        for (let d = step.toDay; d < step.toDay + cur.duration; d++) {
          loadC[d]++;
          if (loadC[d] > limit) { ok = false; break; }
        }
        if (!ok) break;
      }
      if (!ok) {
        if (shiftRight) break;
        continue;
      }
      chain.push({
        activityId: a.activityId,
        text: shiftRight
          ? '「' + a.name + ' 第' + b.period + '期」与第' + a.period + '期间隔不足（要求 ≥ ' + need + ' 天，实际 ' + gap + ' 天）。'
            + '可一键将第' + b.period + '期及之后的 ' + plan.length + ' 期整体顺延'
            + '（第' + b.period + '期移至第 ' + (plan[0].toDay + 1) + ' 天' + (plan.length > 1 ? '，其余各期依次顺延' : '') + '）'
          : '「' + a.name + ' 第' + b.period + '期」与第' + a.period + '期间隔偏大（设定 ' + need + ' 天，实际 ' + gap + ' 天）。'
            + '可一键以第' + a.period + '期为基准，将第' + b.period + '期及之后的 ' + plan.length + ' 期按间隔重新紧凑排列'
            + '（第' + b.period + '期移至第 ' + (plan[0].toDay + 1) + ' 天' + (plan.length > 1 ? '，其余各期依次紧凑' : '') + '）',
        plan,
      });
      break; // 每组活动最多一条链式建议
    }
  }

  if (!denseDays.length) return { list: [], chain, guidanceDays: [], healthy: chain.length === 0 };

  const countDays = (l, fn) => l.reduce((n, v) => n + (fn(v) ? 1 : 0), 0);
  const viol = countDays(load, v => v > limit);   // 超限日数（超过每日上限）
  const dense = denseDays.length;                 // 密集日数
  const maxLoad = Math.max(...load);

  const cands = [];
  const seen = new Set();
  const movableDays = new Set();

  for (const d of denseDays) {
    const actives = insts.filter(i => i.startDay <= d && d < i.startDay + i.duration);
    for (const inst of actives) {
      // 去掉该实例后的基线负载
      const base = load.slice();
      for (let x = inst.startDay; x < inst.startDay + inst.duration && x < totalDays; x++) base[x]--;
      const others = insts.filter(i => i.id !== inst.id && i.activityId === inst.activityId);
      const span = totalDays - inst.duration;
      let best = null;
      // 按与当前位置的距离螺旋搜索（保证建议的移动距离尽量小）
      for (let radius = 1; radius <= totalDays; radius++) {
        for (const c of [inst.startDay - radius, inst.startDay + radius]) {
          if (c < 0 || c > span) continue;
          const load2 = base.slice();
          let ok = true;
          for (let x = c; x < c + inst.duration; x++) {
            load2[x]++;
            if (load2[x] > limit) { ok = false; break; }
          }
          if (!ok) continue;
          for (const s of others) {
            const gap = s.startDay > c ? s.startDay - (c + inst.duration - 1) - 1 : c - (s.startDay + s.duration - 1) - 1;
            if (gap < inst.intervalDays) { ok = false; break; }
          }
          if (!ok) continue;
          // 评分依次：超限日数 ↓ → 密集日数 ↓ → 最大负载 ↓ → 移动距离 ↓ → 位置靠前
          const viol2 = countDays(load2, v => v > limit);
          const dense2 = countDays(load2, v => v >= DENSE);
          const max2 = Math.max(...load2);
          const dist = Math.abs(c - inst.startDay);
          const score = [viol2, dense2, max2, dist, c];
          if (!best || betterScore(score, best.score)) best = { c, score };
        }
      }
      // 仅当移动后"超限日/密集日/最大负载"任一指标变好才值得建议
      if (best && (best.score[0] < viol || best.score[1] < dense || best.score[2] < maxLoad)) {
        movableDays.add(d);
        const key = inst.id + '@' + best.c;
        if (!seen.has(key)) {
          seen.add(key);
          const deltaDesc = viol > 0
            ? '可减少 ' + (viol - best.score[0]) + ' 处超限日'
            : dense - best.score[1] > 0
              ? '可减少 ' + (dense - best.score[1]) + ' 处密集日'
              : '最高单日负载从 ' + maxLoad + ' 降至 ' + best.score[2];
          cands.push({ instId: inst.id, name: inst.name, period: inst.period,
                       fromDay: inst.startDay, toDay: best.c, loadAt: load[d],
                       delta: viol > 0 ? viol - best.score[0] : dense - best.score[1], deltaDesc });
        }
      }
    }
  }

  // 排序：改善幅度大者优先，其次移动距离小者优先
  cands.sort((a, b) => b.delta - a.delta || Math.abs(a.fromDay - a.toDay) - Math.abs(b.fromDay - b.toDay));
  const list = cands.slice(0, 6).map(c => ({
    instId: c.instId, toDay: c.toDay,
    text: '第 ' + (c.fromDay + 1) + ' 天有 ' + c.loadAt + ' 个活动同时进行，建议将「' + c.name + ' 第' + c.period + '期」'
      + '从第 ' + (c.fromDay + 1) + ' 天移至第 ' + (c.toDay + 1) + ' 天（' + c.deltaDesc + '）',
  }));

  // 该密集日上没有任何实例能通过移动改善排期 → 记录后合并成一条指导性建议
  const guidanceDays = [];
  for (const d of denseDays) {
    if (movableDays.has(d)) continue;
    guidanceDays.push(d);
  }
  return { list, chain, guidanceDays, healthy: false };
}

let lastChainPlans = []; // 最近一次渲染的链式调整方案（供"一键应用"按钮查找）

function renderSuggestions() {
  if (!AppState.schedule.instances.length) {
    elSuggestionList.innerHTML = '<div class="muted">生成排期后，这里会给出密度提示与优化建议。</div>';
    return;
  }
  const DENSE = denseThreshold();
  const { list, chain, guidanceDays, healthy } = computeSuggestions();
  if (healthy) {
    elSuggestionList.innerHTML = '<div class="sug-healthy">🎉 当前排期健康：没有间隔违规，也没有单日达到 ' + DENSE + ' 个活动的日期。</div>';
    return;
  }
  lastChainPlans = chain;
  const chainItems = chain.map((c, idx) =>
    '<div class="suggestion-item chain"><span class="sug-text">🔗 ' + escHtml(c.text) + '</span>'
    + '<button class="sug-chain" data-chain-idx="' + idx + '">一键应用</button></div>').join('');
  const items = list.map(s =>
    '<div class="suggestion-item"><span class="sug-text">💡 ' + escHtml(s.text) + '</span>'
    + '<button class="sug-apply" data-inst="' + s.instId + '" data-to="' + s.toDay + '">应用</button></div>').join('');
  const guidanceHtml = guidanceDays.length
    ? '<div class="suggestion-item guidance"><span class="sug-text">📌 「' + fmtDayRanges(guidanceDays) + '」存在单日 ' + DENSE + ' 个及以上活动同时进行，但周期内找不到能改善排期的空位。'
      + '可尝试：增加版本天数 / 提高每日上限 / 减少活动数量或持续天数。</span></div>'
    : '';
  elSuggestionList.innerHTML = chainItems + items + guidanceHtml;
}

// 一键应用链式调整：把同活动后续各期整体顺延（应用前复查全部约束）
function applyChainPlan(plan) {
  if (!plan || !plan.length) return;
  const insts = AppState.schedule.instances;
  const { totalDays } = AppState.version, r = AppState.rules;
  const movedIds = new Set(plan.map(s => s.instId));
  const copy = deepCopy(insts);
  const rest = copy.filter(i => !movedIds.has(i.id));
  const load = computeDayLoad(rest, totalDays);
  const limit = r.allowOverlap ? r.dailyCap : 1;
  const chain = plan.map(s => {
    const inst = copy.find(i => i.id === s.instId);
    return inst ? { ...inst, startDay: s.toDay } : null;
  }).filter(Boolean).sort((x, y) => x.startDay - y.startDay);
  // ① 边界与每日上限
  for (const c of chain) {
    if (c.startDay < 0 || c.startDay + c.duration > totalDays) { alert('该建议已失效（超出版本周期），请重新查看建议'); return; }
    for (let d = c.startDay; d < c.startDay + c.duration; d++) {
      load[d]++;
      if (load[d] > limit) { alert('该建议已失效（超过每日上限），请重新查看建议'); return; }
    }
  }
  // ② 与未移动的同活动实例之间的间隔
  for (const c of chain) {
    for (const s of rest) {
      if (s.activityId !== c.activityId) continue;
      const gap = s.startDay > c.startDay ? s.startDay - (c.startDay + c.duration - 1) - 1 : c.startDay - (s.startDay + s.duration - 1) - 1;
      if (gap < Math.max(c.intervalDays || 0, s.intervalDays || 0)) { alert('该建议已失效（间隔不足），请重新查看建议'); return; }
    }
  }
  // ③ 链内相邻间隔（生成时已保证，此处兜底复核）
  for (let k = 0; k + 1 < chain.length; k++) {
    const a = chain[k], b = chain[k + 1];
    const gap = b.startDay - (a.startDay + a.duration - 1) - 1;
    if (gap < Math.max(a.intervalDays || 0, b.intervalDays || 0)) { alert('该建议已失效（间隔不足），请重新查看建议'); return; }
  }
  // 提交
  pushUndo();
  for (const step of plan) {
    const inst = copy.find(i => i.id === step.instId);
    if (inst) { inst.startDay = step.toDay; inst.manual = true; }
  }
  AppState.schedule.instances = copy;
  renderAll();
}

elSuggestionList.addEventListener('click', e => {
  const chainBtn = e.target.closest('.sug-chain');
  if (chainBtn) {
    const idx = parseInt(chainBtn.dataset.chainIdx, 10);
    applyChainPlan(lastChainPlans[idx] && lastChainPlans[idx].plan);
    return;
  }
  const btn = e.target.closest('.sug-apply');
  if (!btn) return;
  const inst = findInst(btn.dataset.inst);
  if (!inst) return;
  const toDay = parseInt(btn.dataset.to, 10);
  // 应用前复查约束（排期可能已被手动改动）
  const others = AppState.schedule.instances.filter(i => i.id !== inst.id);
  const load = computeDayLoad(others, AppState.version.totalDays);
  if (!canPlaceWith(toDay, inst.duration, inst.activityId, inst.intervalDays, others, load, AppState.version.totalDays, AppState.rules)) {
    alert('该建议已失效（排期可能已被修改），请重新查看建议');
    return;
  }
  pushUndo();
  inst.startDay = toDay;
  inst.manual = true;
  renderAll();
});

/* ==== 11. 导出 Excel ==== */
function exportExcel() {
  if (typeof XLSX === 'undefined') {
    alert('未加载 Excel 导出库（libs/xlsx.full.min.js）。\n请按 README.md 中的说明下载该文件到 libs/ 目录后刷新页面。');
    return;
  }
  const insts = AppState.schedule.instances;
  if (!insts.length) { alert('当前没有排期数据，请先生成排期'); return; }
  const { startDate, totalDays } = AppState.version;
  const DENSE = denseThreshold();

  // Sheet1 排期表：每个实例一行
  const rows1 = [['活动名称', '活动类型', '期数', '开始时间', '结束时间', '持续天数']];
  for (const i of insts) rows1.push([
    i.name, typeLabel(i.type), '第' + i.period + '期',
    fmtDate(dateAt(startDate, i.startDay)),
    fmtDate(dateAt(startDate, i.startDay + i.duration - 1)),
    i.duration,
  ]);
  const ws1 = XLSX.utils.aoa_to_sheet(rows1);
  ws1['!cols'] = [{ wch: 20 }, { wch: 16 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 10 }];

  // Sheet2 每日明细：每天一行
  const load = computeDayLoad(insts, totalDays);
  const rows2 = [['日期', '星期', '当日进行中活动', '活动数', '密度等级']];
  for (let d = 0; d < totalDays; d++) {
    const dt = dateAt(startDate, d);
    const names = insts.filter(i => i.startDay <= d && d < i.startDay + i.duration)
      .map(i => i.name + '(第' + i.period + '期)').join('、');
    rows2.push([fmtDate(dt), '周' + WEEK_CN[dt.getDay()], names, load[d],
      load[d] === 0 ? '空闲' : load[d] < DENSE ? '适中' : '密集']);
  }
  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  ws2['!cols'] = [{ wch: 12 }, { wch: 6 }, { wch: 44 }, { wch: 8 }, { wch: 10 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, '排期表');
  XLSX.utils.book_append_sheet(wb, ws2, '每日明细');
  const fname = '活动排期_' + startDate + '_至_' + fmtDate(dateAt(startDate, totalDays - 1)) + '.xlsx';
  try {
    XLSX.writeFile(wb, fname);
  } catch (err) {
    // 兜底：手动 Blob + 下载链接
    const blob = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })],
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = fname;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
}

/* ==== 12. 渲染管线与初始化 ==== */
function renderAll() {
  const conflicts = validateSchedule(); // 冲突先算：甘特图标红与冲突面板共用同一份
  renderActivities();
  renderToolbar();
  renderVersionSummary();
  renderWarningBanner();
  renderStats();
  renderGantt(conflicts);
  renderSuggestions();
  renderConflicts(conflicts);
  renderLegend();
  renderModal();
}

function renderLegend() {
  const DENSE = denseThreshold();
  const midLabel = DENSE === 2 ? '1' : '1-' + (DENSE - 1);
  const typeItems = TYPE_DEFS.map(t =>
    '<span class="legend-item"><span class="swatch" style="background:' + t.color + '"></span>' + t.label + '</span>').join('');
  elGanttLegend.innerHTML = typeItems
    + '<span class="legend-sep"></span>'
    + '<span class="legend-item"><span class="swatch swatch-empty"></span>0 个/天（空闲）</span>'
    + '<span class="legend-item"><span class="swatch d1"></span>' + midLabel + ' 个/天</span>'
    + '<span class="legend-item"><span class="swatch d2"></span>≥' + DENSE + ' 个/天（密集）</span>';
}

function syncInputsFromState() {
  elStartDate.value = AppState.version.startDate;
  elTotalDays.value = AppState.version.totalDays;
  elAllowOverlap.checked = AppState.rules.allowOverlap;
  elDailyCap.value = AppState.rules.dailyCap;
}

/* —— 版本与规则输入 —— */
elStartDate.addEventListener('change', () => {
  if (parseDateStr(elStartDate.value)) AppState.version.startDate = elStartDate.value;
  renderAll();
});
elTotalDays.addEventListener('change', () => { AppState.version.totalDays = clampInt(elTotalDays.value, 1, 90); renderAll(); });
elAllowOverlap.addEventListener('change', () => { AppState.rules.allowOverlap = elAllowOverlap.checked; renderAll(); });
elDailyCap.addEventListener('change', () => { AppState.rules.dailyCap = clampInt(elDailyCap.value, 1, 10); renderAll(); });

/* —— 活动列表（事件委托） —— */
elActivityList.addEventListener('change', e => {
  const item = e.target.closest('.activity-item');
  if (!item) return;
  const a = AppState.activities.find(x => x.id === item.dataset.id);
  if (!a) return;
  if (e.target.classList.contains('act-name')) a.name = e.target.value.trim() || '未命名活动';
  else if (e.target.classList.contains('act-type')) { a.type = e.target.value; a.color = typeColor(a.type); }
  else if (e.target.classList.contains('act-duration')) a.duration = clampInt(e.target.value, 1, 90);
  else if (e.target.classList.contains('act-count')) a.count = clampInt(e.target.value, 1, 30);
  else if (e.target.classList.contains('act-interval')) a.intervalDays = clampInt(e.target.value, 0, 30);
  else if (e.target.classList.contains('act-color')) a.color = e.target.value;
  renderAll();
});
elActivityList.addEventListener('click', e => {
  const del = e.target.closest('.act-delete');
  if (!del) return;
  const item = del.closest('.activity-item');
  const a = AppState.activities.find(x => x.id === item.dataset.id);
  if (!a) return;
  if (!confirm('确定删除活动「' + a.name + '」？\n已排入的实例也会一并移除。')) return;
  AppState.activities = AppState.activities.filter(x => x.id !== a.id);
  AppState.schedule.instances = AppState.schedule.instances.filter(i => i.activityId !== a.id);
  pushUndo();
  renderAll();
});
elAddActivity.addEventListener('click', () => {
  const t = TYPE_DEFS[TYPE_DEFS.length - 1];
  AppState.activities.push({ id: uid('a'), name: '新活动', type: t.key, duration: 7, count: 1, intervalDays: 0, color: t.color });
  renderAll();
});

/* —— 工具栏 —— */
elGenerate.addEventListener('click', () => generateSchedule(true));
elIncremental.addEventListener('click', scheduleIncremental);
elUndo.addEventListener('click', undo);
elExport.addEventListener('click', exportExcel);

/* —— 快捷键：Ctrl+Z 撤销、Esc 关闭弹窗 —— */
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
    const t = e.target;
    if (t && (t.tagName === 'INPUT' || t.tagName === 'SELECT' || t.tagName === 'TEXTAREA')) return;
    e.preventDefault();
    undo();
  }
  if (e.key === 'Escape' && AppState.ui.editingId) closeEditModal();
});

/* —— 启动 —— */
syncInputsFromState();
renderAll();
generateSchedule(false); // 首次载入自动生成一份示例排期，方便直接演示
