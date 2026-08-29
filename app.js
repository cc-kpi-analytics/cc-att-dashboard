/* ============================================================
   Attendance Console — data load, aggregation, and rendering
   ============================================================ */

const MONTH_NAMES = ['January','February','March','April','May','June',
                      'July','August','September','October','November','December'];

// Event types whose duration should NOT count toward an agent's scheduled
// hours total — these are unscheduled/non-working time (approved leave
// types, unpaid personal time, admin time, and lunch).
// NOTE: 'Admin' here is the plain "Admin" type, distinct from "Work - Admin"
// (which IS real work and counts toward Sched hours as normal).
const EXCLUDE_FROM_SCHED = new Set([
  'Lunch', 'PTO', 'UPL', 'FMLA', 'PFML', 'Bereavement', 'Alt Holiday',
  'Jury Duty', 'Admin', 'Traumatic Leave',
]);

// Event types treated as attendance/absence statuses in the Daily Log.
// Note: 'Admin' here means the plain "Admin" type only — "Work - Admin" is
// real work and is never treated as a status.
const EXCEPTION_TYPES = new Set([
  'PTO', 'UPL', 'FMLA', 'PFML', 'Late', 'Left Early', 'UTO', 'Bereavement',
  'Alt Holiday', 'NCNS', 'Traumatic Leave', 'Jury Duty', 'Admin',
]);

// "Unapproved shrink" — the numerator of Attendance %.
const NON_DISCRETIONARY_TYPES = new Set(['Late', 'Left Early', 'UTO', 'NCNS']);

// Pill styling per absence status.
const PILL_CLASS = {
  'Late': 'warn',
  'Left Early': 'warn',
  'UTO': 'bad',
  'NCNS': 'bad',
  'PTO': 'info',
  'UPL': 'info',
  'FMLA': 'info',
  'PFML': 'info',
  'Jury Duty': 'info',
  'Admin': 'info',
  'Bereavement': 'violet',
  'Alt Holiday': 'violet',
  'Traumatic Leave': 'violet',
};

/* ---------------- weekly data source config ---------------- */

const DATA_DIR = './data/';

// How far back/forward to look for weekly files when auto-discovering what's
// available. Widen these if you have older history or need to see further
// into the future — see README.
const DISCOVERY_YEARS_BACK = 2;
const DISCOVERY_DAYS_FORWARD = 120;
const DISCOVERY_CONCURRENCY = 16;

let STATE = {
  supMap: new Map(),
  progMap: new Map(),
  agentNames: [],
  supervisors: [],
  programs: [],

  availableWeeks: [],     // [{ sunday, saturday, filename, key }], sorted ascending
  weekPromises: new Map(),// key -> Promise<enriched records[]>, doubles as a cache

  records: [],            // flattened enriched interval records from every fetched week
  dailyGroups: [],        // one entry per agent+date, rebuilt whenever records grow

  years: [],
  monthsByYear: new Map(),

  overview: { year: '', month: '', supervisor: '' },
  dailylog: { year: '', month: '', search: '' },
  watchlist: { year: '', month: '', program: '' },
  showHours: false,

  renderGen: { overview: 0, watchlist: 0, dailylog: 0 },
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

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s === null || s === undefined ? '' : String(s);
  return d.innerHTML;
}

/* ---------------- roster (supervisor / program mapping) ---------------- */

async function loadRoster() {
  setLoading('READING attendance_raw.xlsx …');
  const resp = await fetch('./attendance_raw.xlsx');
  if (!resp.ok) throw new Error('Could not fetch attendance_raw.xlsx (HTTP ' + resp.status + '). Make sure the file sits in the same folder as index.html.');
  const buf = await resp.arrayBuffer();

  setLoading('PARSING roster …');
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  for (const s of ['supervisor', 'program']) {
    if (!wb.SheetNames.includes(s)) throw new Error('Workbook is missing the required "' + s + '" sheet.');
  }
  const supRows = XLSX.utils.sheet_to_json(wb.Sheets['supervisor'], { defval: null });
  const progRows = XLSX.utils.sheet_to_json(wb.Sheets['program'], { defval: null });

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

  STATE.supMap = supMap;
  STATE.progMap = progMap;
  STATE.agentNames = Array.from(new Set(supRows.map(r => r['Name']).filter(Boolean).map(n => String(n).trim()))).sort();
  STATE.supervisors = Array.from(new Set(supRows.map(r => { const s = r['Supervisor']; return s ? String(s).trim() : 'Unassigned'; }))).sort();
  STATE.programs = Array.from(new Set(progRows.map(r => { const p = r['Program']; return p !== null && p !== undefined ? String(p).trim() : 'Unassigned'; }))).sort();
}

/* ---------------- weekly file discovery ---------------- */

function weekFilenameFor(sunday) {
  const yyyy = sunday.getFullYear();
  const mm = String(sunday.getMonth() + 1).padStart(2, '0');
  const dd = String(sunday.getDate()).padStart(2, '0');
  return `WB_${yyyy}_${mm}_${dd}.mhtml`;
}

function computeCandidateSundays() {
  const start = new Date(new Date().getFullYear() - DISCOVERY_YEARS_BACK, 0, 1);
  while (start.getDay() !== 0) start.setDate(start.getDate() + 1); // roll to first Sunday
  const end = new Date();
  end.setDate(end.getDate() + DISCOVERY_DAYS_FORWARD);

  const list = [];
  const cur = new Date(start);
  while (cur <= end) {
    list.push(new Date(cur));
    cur.setDate(cur.getDate() + 7);
  }
  return list;
}

async function headExists(url) {
  try {
    const resp = await fetch(url, { method: 'HEAD', cache: 'no-store' });
    return resp.ok;
  } catch (e) {
    return false;
  }
}

async function discoverAvailableWeeks(progressCb) {
  const candidates = computeCandidateSundays();
  const found = [];
  let idx = 0;
  let checked = 0;

  async function worker() {
    while (idx < candidates.length) {
      const my = idx++;
      const sunday = candidates[my];
      const filename = weekFilenameFor(sunday);
      const ok = await headExists(DATA_DIR + filename);
      checked++;
      if (progressCb) progressCb(checked, candidates.length);
      if (ok) {
        const saturday = new Date(sunday);
        saturday.setDate(saturday.getDate() + 6);
        found.push({ sunday, saturday, filename, key: dateKey(sunday) });
      }
    }
  }

  const workers = [];
  for (let i = 0; i < DISCOVERY_CONCURRENCY; i++) workers.push(worker());
  await Promise.all(workers);

  found.sort((a, b) => a.sunday - b.sunday);
  return found;
}

function computeYearMonthIndex(weeks) {
  const yearSet = new Set();
  const monthsByYear = new Map();
  weeks.forEach(w => {
    for (let i = 0; i < 7; i++) {
      const d = new Date(w.sunday);
      d.setDate(d.getDate() + i);
      const y = d.getFullYear(), m = d.getMonth() + 1;
      yearSet.add(y);
      if (!monthsByYear.has(y)) monthsByYear.set(y, new Set());
      monthsByYear.get(y).add(m);
    }
  });
  return { years: Array.from(yearSet).sort((a, b) => a - b), monthsByYear };
}

function weeksOverlapping(year, month) {
  if (year === '') return STATE.availableWeeks.slice();
  let rangeStart, rangeEnd;
  if (month === '') {
    rangeStart = new Date(year, 0, 1);
    rangeEnd = new Date(year, 11, 31);
  } else {
    rangeStart = new Date(year, month - 1, 1);
    rangeEnd = new Date(year, month, 0); // last day of that month
  }
  return STATE.availableWeeks.filter(w => w.saturday >= rangeStart && w.sunday <= rangeEnd);
}

/* ---------------- weekly file fetch + parse + cache ---------------- */

function fetchAndParseWeek(weekMeta) {
  if (STATE.weekPromises.has(weekMeta.key)) return STATE.weekPromises.get(weekMeta.key);

  const p = (async () => {
    const resp = await fetch(DATA_DIR + weekMeta.filename);
    if (!resp.ok) throw new Error('Failed to fetch ' + weekMeta.filename + ' (HTTP ' + resp.status + ').');
    const text = await resp.text();
    const intervals = parseWeeklyMhtml(text); // from mhtml-parser.js

    const enriched = [];
    for (const iv of intervals) {
      const agent = iv.agent;
      const supervisor = STATE.supMap.get(agent) || 'Unassigned';
      const program = STATE.progMap.get(supervisor) || 'Unassigned';
      const schedAdj = EXCLUDE_FROM_SCHED.has(iv.type) ? 0 : iv.duration;
      const nonDisc = NON_DISCRETIONARY_TYPES.has(iv.type) ? iv.duration : 0;
      const year = iv.date.getFullYear();
      const month = iv.date.getMonth() + 1;

      enriched.push({
        agent, supervisor, program, eventType: iv.type,
        schedRaw: iv.duration, schedAdj, nonDisc,
        date: iv.date, year, month, dkey: dateKey(iv.date),
      });
    }

    STATE.records = STATE.records.concat(enriched);
    buildDailyGroups();
    return enriched;
  })();

  STATE.weekPromises.set(weekMeta.key, p);
  return p;
}

/** Fetches any not-yet-loaded weeks in weekMetas, showing progress inside the given view's table. */
async function ensureWeeksLoadedForView(viewId, weekMetas) {
  const toFetch = weekMetas.filter(w => !STATE.weekPromises.has(w.key));
  if (toFetch.length > 0) {
    const label = toFetch.length === 1 ? '1 new week' : toFetch.length + ' new weeks';
    setTableMessage(viewId, `Loading ${label} of data…`, `0 / ${toFetch.length} fetched`, viewId + '-loadprog');
  }
  let done = 0;
  await Promise.all(weekMetas.map(async w => {
    const wasCached = STATE.weekPromises.has(w.key);
    await fetchAndParseWeek(w);
    if (!wasCached) {
      done++;
      const el = document.getElementById(viewId + '-loadprog');
      if (el) el.textContent = `${done} / ${toFetch.length} fetched`;
    }
  }));
}

/* ---------------- daily grouping ---------------- */

function buildDailyGroups() {
  const map = new Map();
  for (const r of STATE.records) {
    const key = r.agent + '||' + r.dkey;
    let g = map.get(key);
    if (!g) {
      g = { agent: r.agent, supervisor: r.supervisor, program: r.program, date: r.date, dkey: r.dkey, year: r.year, month: r.month,
            eventHours: new Map(), schedAdjSum: 0, nonDiscSum: 0 };
      map.set(key, g);
    }
    g.schedAdjSum += r.schedAdj;
    g.nonDiscSum += r.nonDisc;
    if (EXCEPTION_TYPES.has(r.eventType)) {
      // Merge same-type exceptions split across multiple intervals the same day
      // (e.g. FMLA taken in a morning block and an afternoon block) into one total.
      g.eventHours.set(r.eventType, (g.eventHours.get(r.eventType) || 0) + r.schedRaw);
    }
  }

  const groups = [];
  for (const g of map.values()) {
    let status;
    const realEvents = Array.from(g.eventHours.entries())
      .map(([type, hours]) => ({ type, hours }))
      .filter(e => e.hours > 0); // drop zero-hour placeholder rows
    if (realEvents.length > 0) {
      status = realEvents;
    } else if (g.schedAdjSum > 0) {
      status = 'present';
    } else {
      continue; // no absence status and no scheduled hours — nothing to report that day
    }
    groups.push({
      agent: g.agent, supervisor: g.supervisor, program: g.program,
      date: g.date, dkey: g.dkey, year: g.year, month: g.month, status,
      pct: g.schedAdjSum > 0 ? 1 - (g.nonDiscSum / g.schedAdjSum) : null,
    });
  }
  groups.sort((a, b) => (a.date - b.date) || a.agent.localeCompare(b.agent));
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
  out.sort((a, b) => a.agent.localeCompare(b.agent));
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
  out.sort((a, b) => a.pct - b.pct); // lowest attendance % (worst) first
  return out.slice(0, 10);
}

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

function colCountFor(section) {
  if (section === 'overview') return STATE.showHours ? 6 : 4;
  if (section === 'watchlist') return STATE.showHours ? 7 : 5;
  return 6; // daily log
}

/** Replaces a table's <tbody> with a single spanning message row (loading/empty states). */
function setTableMessage(viewId, bigText, smallText, smallId) {
  const section = viewId === 'view-overview' ? 'overview' : viewId === 'view-watchlist' ? 'watchlist' : 'dailylog';
  const tbodyId = viewId === 'view-overview' ? 'ov-tbody' : viewId === 'view-watchlist' ? 'wl-tbody' : 'dl-tbody';
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  const colspan = colCountFor(section);
  tbody.innerHTML = `<tr><td colspan="${colspan}"><div class="empty-state"><div class="big">${esc(bigText)}</div>` +
    (smallText ? `<div class="small"${smallId ? ` id="${esc(smallId)}"` : ''}>${esc(smallText)}</div>` : '') +
    `</div></td></tr>`;
}

async function renderOverview() {
  const myGen = ++STATE.renderGen.overview;
  const filter = STATE.overview;
  const needed = weeksOverlapping(filter.year, filter.month);

  if (needed.length === 0) {
    setTableMessage('view-overview', 'No data available for this period', 'No weekly files were found for this year/month.');
    return;
  }

  try {
    await ensureWeeksLoadedForView('view-overview', needed);
  } catch (err) {
    if (myGen !== STATE.renderGen.overview) return;
    setTableMessage('view-overview', 'Could not load data for this period', err.message || 'Try again in a moment.');
    return;
  }
  if (myGen !== STATE.renderGen.overview) return; // a newer filter change superseded this render

  const rows = aggregateOverview(filter);
  if (rows.length === 0) {
    setTableMessage('view-overview', 'No records match these filters', 'Try a different year, month, or supervisor.');
    return;
  }

  const tbody = document.getElementById('ov-tbody');
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

async function renderWatchlist() {
  const myGen = ++STATE.renderGen.watchlist;
  const filter = STATE.watchlist;
  const needed = weeksOverlapping(filter.year, filter.month);

  if (needed.length === 0) {
    setTableMessage('view-watchlist', 'No data available for this period', 'No weekly files were found for this year/month.');
    return;
  }

  try {
    await ensureWeeksLoadedForView('view-watchlist', needed);
  } catch (err) {
    if (myGen !== STATE.renderGen.watchlist) return;
    setTableMessage('view-watchlist', 'Could not load data for this period', err.message || 'Try again in a moment.');
    return;
  }
  if (myGen !== STATE.renderGen.watchlist) return;

  const rows = aggregateWatchlist(filter);
  if (rows.length === 0) {
    setTableMessage('view-watchlist', 'No ranked agents for these filters', 'Try a different year, month, or program.');
    return;
  }

  const tbody = document.getElementById('wl-tbody');
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
  }).join(', ');
}

async function renderDailyLog() {
  const myGen = ++STATE.renderGen.dailylog;
  const filter = STATE.dailylog;
  const needed = weeksOverlapping(filter.year, filter.month);

  if (needed.length === 0) {
    setTableMessage('view-dailylog', 'No data available for this period', 'No weekly files were found for this year/month.');
    return;
  }

  try {
    await ensureWeeksLoadedForView('view-dailylog', needed);
  } catch (err) {
    if (myGen !== STATE.renderGen.dailylog) return;
    setTableMessage('view-dailylog', 'Could not load data for this period', err.message || 'Try again in a moment.');
    return;
  }
  if (myGen !== STATE.renderGen.dailylog) return;

  const rows = filterDailyLog(filter);
  if (rows.length === 0) {
    setTableMessage('view-dailylog', 'No daily records match these filters', 'Try a different year, month, or search term.');
    return;
  }

  const MAX_ROWS = 2000;
  const shown = rows.slice(0, MAX_ROWS);
  const tbody = document.getElementById('dl-tbody');
  tbody.innerHTML = shown.map(g => `
    <tr>
      <td class="mono">${fmtDate(g.date)}</td>
      <td class="name-cell">${esc(g.agent)}</td>
      <td>${esc(g.supervisor)}</td>
      <td>${esc(g.program)}</td>
      <td class="status-cell">${statusCellHtml(g.status)}</td>
      <td class="num"><span class="att-badge ${pctBadgeClass(g.pct)}">${fmtPct(g.pct)}</span></td>
    </tr>
  `).join('');

  if (rows.length > MAX_ROWS) {
    document.getElementById('dl-note').textContent =
      `Showing first ${MAX_ROWS.toLocaleString()} of ${rows.length.toLocaleString()} rows — narrow filters to see more`;
  } else {
    document.getElementById('dl-note').textContent =
      "Attendance % is calculated per date (that day's absence hours ÷ that day's adjusted sched hours)";
  }
}

/* ---------------- filter UI wiring ---------------- */

function opt(value, label) {
  const o = document.createElement('option');
  o.value = value;
  o.textContent = label;
  return o;
}

function populateYearMonth(prefix, currentFilter, onChange) {
  const yearSel = document.getElementById(prefix + '-year');
  const monthSel = document.getElementById(prefix + '-month');

  yearSel.appendChild(opt('', 'All years'));
  STATE.years.forEach(y => yearSel.appendChild(opt(y, y)));
  yearSel.value = currentFilter.year === '' ? '' : String(currentFilter.year);

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
  monthSel.value = currentFilter.month === '' ? '' : String(currentFilter.month);

  yearSel.addEventListener('change', () => { refreshMonths(); onChange(); });
  monthSel.addEventListener('change', onChange);
}

function setupOverviewFilters() {
  populateYearMonth('ov', STATE.overview, () => {
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
    STATE.overview.supervisor = '';
    renderOverview();
    updateMeta('ov');
  });

  updateMeta('ov');
}

function setupWatchlistFilters() {
  populateYearMonth('wl', STATE.watchlist, () => {
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
    STATE.watchlist.program = '';
    renderWatchlist();
    updateMeta('wl');
  });

  updateMeta('wl');
}

function setupDailyLogFilters() {
  populateYearMonth('dl', STATE.dailylog, () => {
    STATE.dailylog.year = document.getElementById('dl-year').value === '' ? '' : Number(document.getElementById('dl-year').value);
    STATE.dailylog.month = document.getElementById('dl-month').value === '' ? '' : Number(document.getElementById('dl-month').value);
    renderDailyLog();
    updateMeta('dl');
  });

  const searchInput = document.getElementById('dl-search');
  const suggestBox = document.getElementById('dl-suggest');

  function closeSuggest() { suggestBox.classList.remove('open'); suggestBox.innerHTML = ''; }

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
    STATE.dailylog.search = '';
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

/* ---------------- tabs ---------------- */

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
    await loadRoster();

    setLoading('DISCOVERING weekly files …');
    STATE.availableWeeks = await discoverAvailableWeeks((checked, total) => {
      setLoading(`DISCOVERING weekly files … (${checked}/${total})`);
    });

    if (STATE.availableWeeks.length === 0) {
      throw new Error('No weekly data files were found in ./data/. Upload at least one WB_YYYY_MM_DD.mhtml file to get started.');
    }

    const idx = computeYearMonthIndex(STATE.availableWeeks);
    STATE.years = idx.years;
    STATE.monthsByYear = idx.monthsByYear;

    // Default every section to the most recently available week's period,
    // rather than "All", so first paint only has to fetch 1-2 weeks.
    const latest = STATE.availableWeeks[STATE.availableWeeks.length - 1];
    const defaultYear = latest.sunday.getFullYear();
    const defaultMonth = latest.sunday.getMonth() + 1;
    STATE.overview.year = defaultYear; STATE.overview.month = defaultMonth;
    STATE.watchlist.year = defaultYear; STATE.watchlist.month = defaultMonth;
    STATE.dailylog.year = defaultYear; STATE.dailylog.month = defaultMonth;

    document.getElementById('recordCount').textContent =
      STATE.availableWeeks.length.toLocaleString() + (STATE.availableWeeks.length === 1 ? ' week' : ' weeks') +
      ' available · ' + STATE.agentNames.length.toLocaleString() + ' agents';

    setupTabs();
    setupSecretToggle();
    setupOverviewFilters();
    setupDailyLogFilters();
    setupWatchlistFilters();

    setLoading(null);

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
