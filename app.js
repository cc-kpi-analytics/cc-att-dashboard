/* ============================================================
   Attendance Console — data load, aggregation, and rendering
   ============================================================ */

const MONTH_NAMES = ['January','February','March','April','May','June',
                      'July','August','September','October','November','December'];

// Event types whose "Sched hours" value should NOT count toward an agent's
// scheduled hours total (they are breaks / full-day absence line items).
const EXCLUDE_FROM_SCHED = new Set(['Lunch','PTO','UPL','FMLA','PFML','Bereavement','Alt Holiday']);

// Event types treated as attendance exceptions in the Daily Log.
const EXCEPTION_TYPES = new Set(['PTO','UPL','FMLA','PFML','Late','Left Early','UTO','Bereavement','Alt Holiday','NCNS']);

// Pill styling per exception type.
const PILL_CLASS = {
  'Late': 'warn',
  'Left Early': 'warn',
  'UTO': 'bad',
  'NCNS': 'bad',
  'PTO': 'info',
  'UPL': 'info',
  'FMLA': 'info',
  'PFML': 'info',
  'Bereavement': 'violet',
  'Alt Holiday': 'violet',
};

let STATE = {
  records: [],          // one entry per raw row, enriched
  dailyGroups: [],       // one entry per agent+date
  years: [],
  monthsByYear: new Map(),
  supervisors: [],
  programs: [],
  agentNames: [],
  overview: { year: '', month: '', supervisor: '' },
  dailylog: { year: '', month: '', search: '' },
  watchlist: { year: '', month: '', program: '' },
  showHours: false,
};

/* ---------------- helpers ---------------- */

function fmtHours(n) {
  if (n === null || n === undefined || isNaN(n)) return '0';
  const r = Math.round(n * 100) / 100;
  return r.toString();
}

function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return 'N/A';
  return (n * 100).toFixed(1) + '%';
}

// Attendance % = 100% − (Non-discretionary shrinkage ÷ Sched hours). Higher is better.
function pctBadgeClass(n) {
  if (n === null || n === undefined || isNaN(n)) return 'att-good';
  if (n >= 0.95) return 'att-good';
  if (n >= 0.85) return 'att-warn';
  return 'att-bad';
}

function fmtDate(d) {
  if (!d) return '';
  return d.toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric' });
}

function dateKey(d) {
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

function showError(msg) {
  const el = document.getElementById('errBanner');
  el.textContent = msg;
  el.style.display = 'block';
  setTimeout(() => { el.style.display = 'none'; }, 6000);
}

function setLoading(msg) {
  const overlay = document.getElementById('loadingOverlay');
  if (msg === null) { overlay.classList.add('hidden'); return; }
  overlay.classList.remove('hidden');
  document.getElementById('loadingMsg').textContent = msg;
}

/* ---------------- data loading ---------------- */

async function loadWorkbook() {
  setLoading('READING attendance_raw.xlsx …');
  const resp = await fetch('./attendance_raw.xlsx');
  if (!resp.ok) throw new Error('Could not fetch attendance_raw.xlsx (HTTP ' + resp.status + '). Make sure the file sits in the same folder as index.html.');
  const buf = await resp.arrayBuffer();

  setLoading('PARSING workbook …');
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });

  const need = ['raw', 'supervisor', 'program'];
  for (const s of need) {
    if (!wb.SheetNames.includes(s)) throw new Error('Workbook is missing the required "' + s + '" sheet.');
  }

  const rawRows = XLSX.utils.sheet_to_json(wb.Sheets['raw'], { defval: null });
  const supRows = XLSX.utils.sheet_to_json(wb.Sheets['supervisor'], { defval: null });
  const progRows = XLSX.utils.sheet_to_json(wb.Sheets['program'], { defval: null });

  setLoading('BUILDING agent / program index …');
  const supMap = new Map();
  supRows.forEach(r => {
    const name = r['Name'];
    const sup = r['Supervisor'];
    if (name) supMap.set(String(name).trim(), sup ? String(sup).trim() : 'Unassigned');
  });

  const progMap = new Map();
  progRows.forEach(r => {
    const sup = r['Supervisor'];
    const prog = r['Program'];
    if (sup) progMap.set(String(sup).trim(), prog !== null && prog !== undefined ? String(prog).trim() : 'Unassigned');
  });

  setLoading('PROCESSING ' + rawRows.length.toLocaleString() + ' rows …');
  const records = [];
  const yearSet = new Set();
  const monthsByYear = new Map();
  const supSet = new Set();
  const progSet = new Set();
  const agentSet = new Set();

  for (const row of rawRows) {
    const agent = row['Agent Name'] ? String(row['Agent Name']).trim() : null;
    if (!agent) continue;
    const dateVal = row['Date'];
    if (!dateVal || !(dateVal instanceof Date) || isNaN(dateVal.getTime())) continue;

    const eventType = row['Event type'] ? String(row['Event type']).trim() : '';
    const schedRaw = Number(row['Sched hours']) || 0;
    const workHours = Number(row['Work hours']) || 0;
    const nonDisc = Number(row['Non-discretionary shrinkage']) || 0;

    const supervisor = supMap.get(agent) || 'Unassigned';
    const program = progMap.get(supervisor) || 'Unassigned';
    const schedAdj = EXCLUDE_FROM_SCHED.has(eventType) ? 0 : schedRaw;

    const year = dateVal.getFullYear();
    const month = dateVal.getMonth() + 1;

    records.push({
      agent, supervisor, program, eventType,
      schedRaw, schedAdj, workHours, nonDisc,
      date: dateVal, year, month, dkey: dateKey(dateVal),
    });

    yearSet.add(year);
    if (!monthsByYear.has(year)) monthsByYear.set(year, new Set());
    monthsByYear.get(year).add(month);
    supSet.add(supervisor);
    progSet.add(program);
    agentSet.add(agent);
  }

  STATE.records = records;
  STATE.years = Array.from(yearSet).sort((a,b)=>a-b);
  STATE.monthsByYear = monthsByYear;
  STATE.supervisors = Array.from(supSet).sort();
  STATE.programs = Array.from(progSet).sort();
  STATE.agentNames = Array.from(agentSet).sort();

  setLoading('GROUPING daily records …');
  buildDailyGroups();

  setLoading(null);
}

function buildDailyGroups() {
  const map = new Map();
  for (const r of STATE.records) {
    const key = r.agent + '||' + r.dkey;
    let g = map.get(key);
    if (!g) {
      g = { agent: r.agent, supervisor: r.supervisor, program: r.program, date: r.date, dkey: r.dkey, year: r.year, month: r.month,
            events: [], workHoursSum: 0, schedRawSum: 0, schedAdjSum: 0, nonDiscSum: 0 };
      map.set(key, g);
    }
    g.workHoursSum += r.workHours;
    g.schedRawSum += r.schedRaw;
    g.schedAdjSum += r.schedAdj;
    g.nonDiscSum += r.nonDisc;
    if (EXCEPTION_TYPES.has(r.eventType)) {
      g.events.push({ type: r.eventType, hours: r.schedRaw });
    }
  }

  const groups = [];
  for (const g of map.values()) {
    let status;
    const realEvents = g.events.filter(e => e.hours > 0); // drop zero-hour placeholder rows (e.g. a 0-hr PTO stub alongside a real Late entry)
    if (realEvents.length > 0) {
      status = realEvents;
    } else if (g.workHoursSum > 0 && g.schedRawSum > 0) {
      status = 'present';
    } else {
      continue; // no work, no exception, no real shift info — skip
    }
    groups.push({
      agent: g.agent, supervisor: g.supervisor, program: g.program,
      date: g.date, dkey: g.dkey, year: g.year, month: g.month, status,
      pct: g.schedAdjSum > 0 ? 1 - (g.nonDiscSum / g.schedAdjSum) : null,
    });
  }
  groups.sort((a,b) => (a.date - b.date) || a.agent.localeCompare(b.agent));
  STATE.dailyGroups = groups;
}

/* ---------------- aggregation ---------------- */

function aggregateOverview(filter) {
  const byAgent = new Map();
  for (const r of STATE.records) {
    if (filter.year !== '' && r.year !== filter.year) continue;
    if (filter.month !== '' && r.month !== filter.month) continue;
    if (filter.supervisor !== '' && r.supervisor !== filter.supervisor) continue;
    let a = byAgent.get(r.agent);
    if (!a) {
      a = { agent: r.agent, supervisor: r.supervisor, program: r.program, sched: 0, absence: 0 };
      byAgent.set(r.agent, a);
    }
    a.sched += r.schedAdj;
    a.absence += r.nonDisc;
  }
  const out = Array.from(byAgent.values()).map(a => ({
    ...a,
    pct: a.sched > 0 ? 1 - (a.absence / a.sched) : null,
  }));
  out.sort((a,b) => a.agent.localeCompare(b.agent));
  return out;
}

function aggregateWatchlist(filter) {
  const byAgent = new Map();
  for (const r of STATE.records) {
    if (filter.year !== '' && r.year !== filter.year) continue;
    if (filter.month !== '' && r.month !== filter.month) continue;
    if (filter.program !== '' && r.program !== filter.program) continue;
    let a = byAgent.get(r.agent);
    if (!a) {
      a = { agent: r.agent, supervisor: r.supervisor, program: r.program, sched: 0, absence: 0 };
      byAgent.set(r.agent, a);
    }
    a.sched += r.schedAdj;
    a.absence += r.nonDisc;
  }
  const out = Array.from(byAgent.values())
    .map(a => ({ ...a, pct: a.sched > 0 ? 1 - (a.absence / a.sched) : null }))
    .filter(a => a.pct !== null);
  out.sort((a,b) => a.pct - b.pct); // lowest attendance % (worst) first
  return out.slice(0, 10);
}

// (per-date Attendance % is now computed directly in buildDailyGroups)

function filterDailyLog(filter) {
  const q = filter.search.trim().toLowerCase();
  return STATE.dailyGroups.filter(g => {
    if (filter.year !== '' && g.year !== filter.year) return false;
    if (filter.month !== '' && g.month !== filter.month) return false;
    if (q && !g.agent.toLowerCase().includes(q)) return false;
    return true;
  });
}

/* ---------------- rendering ---------------- */

function renderOverview() {
  const rows = aggregateOverview(STATE.overview);
  const tbody = document.getElementById('ov-tbody');

  if (rows.length === 0) {
    tbody.innerHTML = '';
    document.querySelector('#view-overview .table-scroll').innerHTML =
      '<div class="empty-state"><div class="big">No records match these filters</div><div class="small">Try a different year, month, or supervisor.</div></div>';
    return;
  }
  ensureTable('view-overview', 'ov-tbody');

  tbody.innerHTML = rows.map(a => `
    <tr>
      <td class="name-cell">${esc(a.agent)}</td>
      <td>${esc(a.supervisor)}</td>
      <td>${esc(a.program)}</td>
      ${STATE.showHours ? `<td class="num">${fmtHours(a.sched)}</td><td class="num">${fmtHours(a.absence)}</td>` : ''}
      <td class="num"><span class="att-badge ${pctBadgeClass(a.pct)}">${fmtPct(a.pct)}</span></td>
    </tr>
  `).join('');
}

function renderWatchlist() {
  const rows = aggregateWatchlist(STATE.watchlist);
  const tbody = document.getElementById('wl-tbody');

  if (rows.length === 0) {
    tbody.innerHTML = '';
    document.querySelector('#view-watchlist .table-scroll').innerHTML =
      '<div class="empty-state"><div class="big">No ranked agents for these filters</div><div class="small">Try a different year, month, or program.</div></div>';
    return;
  }
  ensureTable('view-watchlist', 'wl-tbody');

  tbody.innerHTML = rows.map((a, i) => `
    <tr>
      <td class="rank-cell"><span class="rank-num">${i+1}</span></td>
      <td class="name-cell">${esc(a.agent)}</td>
      <td>${esc(a.supervisor)}</td>
      <td>${esc(a.program)}</td>
      ${STATE.showHours ? `<td class="num">${fmtHours(a.sched)}</td><td class="num">${fmtHours(a.absence)}</td>` : ''}
      <td class="num"><span class="att-badge ${pctBadgeClass(a.pct)}">${fmtPct(a.pct)}</span></td>
    </tr>
  `).join('');
}

function statusCellHtml(status) {
  if (status === 'present') {
    return '<span class="pill present">Present</span>';
  }
  return status.map(e => {
    const cls = PILL_CLASS[e.type] || 'info';
    return `<span class="pill ${cls}">${esc(e.type)} (${fmtHours(e.hours)} hrs)</span>`;
  }).join(' ');
}

function renderDailyLog() {
  const rows = filterDailyLog(STATE.dailylog);
  const tbody = document.getElementById('dl-tbody');

  if (rows.length === 0) {
    tbody.innerHTML = '';
    document.querySelector('#view-dailylog .table-scroll').innerHTML =
      '<div class="empty-state"><div class="big">No daily records match these filters</div><div class="small">Try a different year, month, or search term.</div></div>';
    return;
  }
  ensureTable('view-dailylog', 'dl-tbody');

  const MAX_ROWS = 2000;
  const shown = rows.slice(0, MAX_ROWS);

  tbody.innerHTML = shown.map(g => {
    return `
    <tr>
      <td class="mono">${fmtDate(g.date)}</td>
      <td class="name-cell">${esc(g.agent)}</td>
      <td>${esc(g.supervisor)}</td>
      <td>${esc(g.program)}</td>
      <td class="status-cell">${statusCellHtml(g.status)}</td>
      <td class="num"><span class="att-badge ${pctBadgeClass(g.pct)}">${fmtPct(g.pct)}</span></td>
    </tr>
  `;
  }).join('');

  if (rows.length > MAX_ROWS) {
    document.getElementById('dl-note').textContent =
      `Showing first ${MAX_ROWS.toLocaleString()} of ${rows.length.toLocaleString()} rows — narrow filters to see more`;
  } else {
    document.getElementById('dl-note').textContent =
      "Attendance % is calculated per date (that day's absence hours ÷ that day's adjusted sched hours)";
  }
}

function ensureTable(viewId, tbodyId) {
  // Re-inject the table markup if an empty-state replaced it.
  const scroll = document.querySelector('#' + viewId + ' .table-scroll');
  if (!document.getElementById(tbodyId)) {
    location.reload(); // safety net, shouldn't normally trigger
  }
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s === null || s === undefined ? '' : String(s);
  return d.innerHTML;
}

/* ---------------- filter UI wiring ---------------- */

function opt(value, label) {
  const o = document.createElement('option');
  o.value = value;
  o.textContent = label;
  return o;
}

function populateYearMonth(prefix, onChange) {
  const yearSel = document.getElementById(prefix + '-year');
  const monthSel = document.getElementById(prefix + '-month');

  yearSel.appendChild(opt('', 'All years'));
  STATE.years.forEach(y => yearSel.appendChild(opt(y, y)));

  function refreshMonths() {
    const yVal = yearSel.value;
    monthSel.innerHTML = '';
    monthSel.appendChild(opt('', 'All months'));
    let monthNums;
    if (yVal === '') {
      const all = new Set();
      STATE.monthsByYear.forEach(set => set.forEach(m => all.add(m)));
      monthNums = Array.from(all);
    } else {
      monthNums = Array.from(STATE.monthsByYear.get(Number(yVal)) || []);
    }
    monthNums.sort((a,b)=>a-b).forEach(m => monthSel.appendChild(opt(m, MONTH_NAMES[m-1])));
  }
  refreshMonths();

  yearSel.addEventListener('change', () => { refreshMonths(); onChange(); });
  monthSel.addEventListener('change', onChange);
}

function setupOverviewFilters() {
  populateYearMonth('ov', () => {
    STATE.overview.year = document.getElementById('ov-year').value === '' ? '' : Number(document.getElementById('ov-year').value);
    STATE.overview.month = document.getElementById('ov-month').value === '' ? '' : Number(document.getElementById('ov-month').value);
    renderOverview();
    updateMeta('ov');
  });

  const supSel = document.getElementById('ov-supervisor');
  supSel.appendChild(opt('', 'All supervisors'));
  STATE.supervisors.forEach(s => supSel.appendChild(opt(s, s)));
  supSel.addEventListener('change', () => {
    STATE.overview.supervisor = supSel.value;
    renderOverview();
    updateMeta('ov');
  });

  document.getElementById('ov-clear').addEventListener('click', () => {
    document.getElementById('ov-year').value = '';
    document.getElementById('ov-year').dispatchEvent(new Event('change'));
    supSel.value = '';
    STATE.overview = { year: '', month: '', supervisor: '' };
    renderOverview();
    updateMeta('ov');
  });

  updateMeta('ov');
}

function setupWatchlistFilters() {
  populateYearMonth('wl', () => {
    STATE.watchlist.year = document.getElementById('wl-year').value === '' ? '' : Number(document.getElementById('wl-year').value);
    STATE.watchlist.month = document.getElementById('wl-month').value === '' ? '' : Number(document.getElementById('wl-month').value);
    renderWatchlist();
    updateMeta('wl');
  });

  const progSel = document.getElementById('wl-program');
  progSel.appendChild(opt('', 'All programs'));
  STATE.programs.forEach(p => progSel.appendChild(opt(p, p)));
  progSel.addEventListener('change', () => {
    STATE.watchlist.program = progSel.value;
    renderWatchlist();
    updateMeta('wl');
  });

  document.getElementById('wl-clear').addEventListener('click', () => {
    document.getElementById('wl-year').value = '';
    document.getElementById('wl-year').dispatchEvent(new Event('change'));
    progSel.value = '';
    STATE.watchlist = { year: '', month: '', program: '' };
    renderWatchlist();
    updateMeta('wl');
  });

  updateMeta('wl');
}

function setupDailyLogFilters() {
  populateYearMonth('dl', () => {
    STATE.dailylog.year = document.getElementById('dl-year').value === '' ? '' : Number(document.getElementById('dl-year').value);
    STATE.dailylog.month = document.getElementById('dl-month').value === '' ? '' : Number(document.getElementById('dl-month').value);
    renderDailyLog();
    updateMeta('dl');
  });

  const searchInput = document.getElementById('dl-search');
  const suggestBox = document.getElementById('dl-suggest');
  let hiIndex = -1;

  function closeSuggest() { suggestBox.classList.remove('open'); suggestBox.innerHTML = ''; hiIndex = -1; }

  function openSuggest(matches) {
    if (matches.length === 0) {
      suggestBox.innerHTML = '<div class="empty">No matching agents</div>';
    } else {
      suggestBox.innerHTML = matches.slice(0, 8).map(n => `<button type="button">${esc(n)}</button>`).join('');
    }
    suggestBox.classList.add('open');
  }

  searchInput.addEventListener('input', () => {
    const q = searchInput.value.trim().toLowerCase();
    STATE.dailylog.search = searchInput.value;
    renderDailyLog();
    updateMeta('dl');
    if (q === '') { closeSuggest(); return; }
    const matches = STATE.agentNames.filter(n => n.toLowerCase().includes(q));
    openSuggest(matches);
  });

  suggestBox.addEventListener('click', (e) => {
    const btn = e.target.closest('button');
    if (!btn) return;
    searchInput.value = btn.textContent;
    STATE.dailylog.search = btn.textContent;
    closeSuggest();
    renderDailyLog();
    updateMeta('dl');
  });

  document.addEventListener('click', (e) => {
    if (!e.target.closest('.search-wrap')) closeSuggest();
  });

  document.getElementById('dl-clear').addEventListener('click', () => {
    document.getElementById('dl-year').value = '';
    document.getElementById('dl-year').dispatchEvent(new Event('change'));
    searchInput.value = '';
    STATE.dailylog = { year: '', month: '', search: '' };
    closeSuggest();
    renderDailyLog();
    updateMeta('dl');
  });

  updateMeta('dl');
}

function updateMeta(prefix) {
  const el = document.getElementById(prefix + '-meta');
  if (!el) return;
  const parts = [];
  const s = STATE[prefix === 'ov' ? 'overview' : prefix === 'wl' ? 'watchlist' : 'dailylog'];
  if (s.year !== '') parts.push(s.year);
  if (s.month !== '') parts.push(MONTH_NAMES[s.month-1]);
  if (prefix === 'ov' && s.supervisor) parts.push(s.supervisor);
  if (prefix === 'wl' && s.program) parts.push(s.program);
  if (prefix === 'dl' && s.search) parts.push('"' + s.search + '"');
  el.textContent = parts.length ? parts.join(' · ') : 'All records';
}

/* ---------------- tabs ---------------- */

/* ---------------- secret hours toggle ----------------
   Type "hours" anywhere on the page (outside a text field) to show/hide
   the Sched Hours and Absence Hours columns in Overview and Watchlist.
   No visible button/menu on purpose — type it again to hide. */

function updateHoursHeaders() {
  const rows = [document.getElementById('ov-head-row'), document.getElementById('wl-head-row')];
  rows.forEach(row => {
    if (!row) return;
    row.querySelectorAll('.hours-th').forEach(el => el.remove());
    if (STATE.showHours) {
      const anchor = row.querySelector('.attendance-th');
      const schedTh = document.createElement('th');
      schedTh.className = 'num hours-th';
      schedTh.textContent = 'Sched Hours';
      const absTh = document.createElement('th');
      absTh.className = 'num hours-th';
      absTh.textContent = 'Absence Hours';
      anchor.parentNode.insertBefore(schedTh, anchor);
      anchor.parentNode.insertBefore(absTh, anchor);
    }
  });
}

function showToast(msg) {
  const el = document.getElementById('secretToast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => el.classList.remove('show'), 1400);
}

function toggleHoursColumns() {
  STATE.showHours = !STATE.showHours;
  updateHoursHeaders();
  renderOverview();
  renderWatchlist();
  showToast(STATE.showHours ? 'Hours columns: shown' : 'Hours columns: hidden');
}

function setupSecretToggle() {
  let buffer = '';
  document.addEventListener('keydown', (e) => {
    const t = e.target;
    const isTyping = t && (t.tagName === 'INPUT' || t.tagName === 'SELECT' || t.tagName === 'TEXTAREA');
    if (isTyping || e.key.length !== 1) return;
    buffer = (buffer + e.key).slice(-10).toLowerCase();
    if (buffer.endsWith('hours')) {
      toggleHoursColumns();
      buffer = '';
    }
  });
}

function setupTabs() {
  const btns = document.querySelectorAll('.tab-btn');
  btns.forEach(btn => {
    btn.addEventListener('click', () => {
      btns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
      document.getElementById('view-' + btn.dataset.view).classList.add('active');
    });
  });
}

/* ---------------- boot ---------------- */

async function boot() {
  try {
    await loadWorkbook();
    document.getElementById('recordCount').textContent =
      STATE.records.length.toLocaleString() + ' rows · ' + STATE.agentNames.length + ' agents';

    setupTabs();
    setupSecretToggle();
    setupOverviewFilters();
    setupDailyLogFilters();
    setupWatchlistFilters();

    renderOverview();
    renderDailyLog();
    renderWatchlist();
  } catch (err) {
    console.error(err);
    setLoading(null);
    showError(err.message || 'Something went wrong loading the dashboard.');
  }
}

boot();
