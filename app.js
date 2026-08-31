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

  currentUser: null,        // signed-in username
  allowedPrograms: null,    // Set of allowed programs for this session, or null = unrestricted (admin)

  availableWeeks: [],     // [{ sunday, saturday, filename, key }], sorted ascending
  weekPromises: new Map(),// key -> Promise<enriched records[]>, doubles as a cache

  records: [],            // flattened enriched interval records from every fetched week
  dailyGroups: [],        // one entry per agent+date, rebuilt whenever records grow

  years: [],
  monthsByYear: new Map(),

  overview: { year: '', month: '', supervisor: '' },
  dailylog: { year: '', month: '', date: '', search: '' },
  watchlist: { year: '', month: '', program: '' },
  programsView: { year: '', month: '' },
  showHours: false,

  renderGen: { overview: 0, watchlist: 0, dailylog: 0, programsView: 0 },
};

/* ---------------- helpers ---------------- */

function fmtHours(n) {
  if (n === null || n === undefined || isNaN(n)) return '0';
  const r = Math.round(n * 100) / 100;
  return r.toString();
}

function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return (n * 100).toFixed(1) + '%';
}

// Attendance % = 100% − (Non-discretionary shrinkage ÷ Sched hours). Higher is better.
function pctBadgeClass(n) {
  if (n === null || n === undefined || isNaN(n)) return 'att-neutral';
  if (n >= 0.95) return 'att-good';
  if (n >= 0.85) return 'att-warn';
  return 'att-bad';
}

/**
 * Colors rows by rank against each other rather than fixed thresholds —
 * used for small aggregate panels (like Absenteeism by Day of Week) where
 * every value is a pooled average across many agents and naturally
 * clusters into a tight band, so the usual 85%/95% cutoffs barely
 * differentiate anything. Mutates each row with a `badgeClass`.
 */
function assignRelativeBadges(rows) {
  const withPct = rows.filter(r => r.pct !== null && r.pct !== undefined);
  const sorted = [...withPct].sort((a, b) => a.pct - b.pct); // worst first
  const n = sorted.length;
  sorted.forEach((r, i) => {
    const frac = n > 1 ? i / (n - 1) : 1; // 0 = worst, 1 = best
    if (frac < 1 / 3) r.badgeClass = 'att-bad';
    else if (frac < 2 / 3) r.badgeClass = 'att-warn';
    else r.badgeClass = 'att-good';
  });
  rows.forEach(r => { if (r.pct === null || r.pct === undefined) r.badgeClass = 'att-neutral'; });
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

/** Like esc(), but also escapes quotes — required for safely embedding text inside an HTML attribute value. */
function escAttr(s) {
  return esc(s).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Renders a Supervisor cell, with a hover tooltip explaining "Unassigned" when it applies. */
function fmtSupervisorCell(agent, supervisor) {
  if (supervisor !== 'Unassigned') return esc(supervisor);
  const tip = `${agent} isn't listed in the roster's "supervisor" sheet yet — add them there once their supervisor is known.`;
  return `<span class="unassigned-tag" data-tooltip="${escAttr(tip)}">Unassigned</span>`;
}

/** Renders a Program cell, with a hover tooltip explaining "Unassigned" when it applies. */
function fmtProgramCell(supervisor, program) {
  if (program !== 'Unassigned') return esc(program);
  const tip = supervisor === 'Unassigned'
    ? `No supervisor is assigned, so a program can't be determined either.`
    : `${supervisor} isn't listed in the roster's "program" sheet yet — add them there once their program is known.`;
  return `<span class="unassigned-tag" data-tooltip="${escAttr(tip)}">Unassigned</span>`;
}

/** Shows "signed in as X · Sign out" in the header. */
function updateSignedInStatus() {
  if (!STATE.currentUser) return;
  const wrap = document.getElementById('authStatus');
  const divider = document.getElementById('authDivider');
  if (!wrap || !divider) return;
  wrap.innerHTML = `${esc(STATE.currentUser)} <button type="button" id="signOutBtn" class="sign-out-btn">Sign out</button>`;
  wrap.style.display = '';
  divider.style.display = '';
  document.getElementById('signOutBtn').addEventListener('click', signOutUser);
}

/* ---------------- roster (supervisor / program mapping) ---------------- */

async function fetchRosterWorkbook() {
  setLoading('READING attendance_raw.xlsx …');
  const resp = await fetch('./attendance_raw.xlsx');
  if (!resp.ok) throw new Error('Could not fetch attendance_raw.xlsx (HTTP ' + resp.status + '). Make sure the file sits in the same folder as index.html.');
  const buf = await resp.arrayBuffer();

  setLoading('PARSING roster …');
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  for (const s of ['supervisor', 'program']) {
    if (!wb.SheetNames.includes(s)) throw new Error('Workbook is missing the required "' + s + '" sheet.');
  }
  return wb;
}

function buildRosterFromWorkbook(wb) {
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

/**
 * Scopes the roster (and, from then on, every fetched week's records) down
 * to the signed-in user's allowed program(s). Pass null for unrestricted
 * (admin) access.
 */
function applyProgramRestriction(allowedPrograms) {
  STATE.allowedPrograms = allowedPrograms || null;
  if (!STATE.allowedPrograms) return;

  const allowed = STATE.allowedPrograms;
  STATE.programs = STATE.programs.filter(p => allowed.has(p));
  STATE.supervisors = STATE.supervisors.filter(s => allowed.has(STATE.progMap.get(s) || 'Unassigned'));
  STATE.agentNames = STATE.agentNames.filter(a => {
    const sup = STATE.supMap.get(a) || 'Unassigned';
    return allowed.has(STATE.progMap.get(sup) || 'Unassigned');
  });
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

/** Parses an <input type="date"> value ("YYYY-MM-DD") as a local date, avoiding the UTC-shift pitfall of `new Date(str)`. */
function parseISODateLocal(str) {
  const [y, m, d] = str.split('-').map(Number);
  return new Date(y, m - 1, d);
}

/** Which week(s) does the Daily Log need for its current filter — a single exact date, or the usual year/month range. */
function weeksForDailyLog(filter) {
  if (filter.date) {
    const d = parseISODateLocal(filter.date);
    const match = STATE.availableWeeks.find(w => d >= w.sunday && d <= w.saturday);
    return match ? [match] : [];
  }
  return weeksOverlapping(filter.year, filter.month);
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
      if (STATE.allowedPrograms && !STATE.allowedPrograms.has(program)) continue; // out of scope for this signed-in user
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

/** Sums Sched/Absence hours across records matching filterFn and returns the combined Attendance %. */
function computeGroupAttendance(filterFn) {
  let sched = 0, absence = 0;
  for (const r of STATE.records) {
    if (!filterFn(r)) continue;
    sched += r.schedAdj;
    absence += r.nonDisc;
  }
  return sched > 0 ? 1 - (absence / sched) : null;
}

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
  return out.slice(0, 15);
}

/** Admin-only: one row per program, aggregated across every agent in it. */
function aggregateProgramSummary(filter) {
  const byProgram = new Map();
  for (const r of STATE.records) {
    if (filter.year !== '' && r.year !== filter.year) continue;
    if (filter.month !== '' && r.month !== filter.month) continue;
    let p = byProgram.get(r.program);
    if (!p) {
      p = { program: r.program, sched: 0, absence: 0, agents: new Set() };
      byProgram.set(r.program, p);
    }
    p.sched += r.schedAdj;
    p.absence += r.nonDisc;
    p.agents.add(r.agent);
  }
  const out = Array.from(byProgram.values()).map(p => ({
    program: p.program,
    agentCount: p.agents.size,
    sched: p.sched,
    absence: p.absence,
    pct: p.sched > 0 ? 1 - (p.absence / p.sched) : null,
  }));
  out.sort((a, b) => a.program.localeCompare(b.program));
  return out;
}

/** Admin-only: how the "unapproved shrink" hours break down by exception type, across all programs. */
function aggregateAbsenceBreakdown(filter) {
  const byType = new Map();
  let total = 0;
  for (const r of STATE.records) {
    if (filter.year !== '' && r.year !== filter.year) continue;
    if (filter.month !== '' && r.month !== filter.month) continue;
    if (!NON_DISCRETIONARY_TYPES.has(r.eventType) || r.schedRaw <= 0) continue;
    byType.set(r.eventType, (byType.get(r.eventType) || 0) + r.schedRaw);
    total += r.schedRaw;
  }
  const rows = Array.from(byType.entries()).map(([type, hours]) => ({
    type, hours, share: total > 0 ? hours / total : 0,
  }));
  rows.sort((a, b) => b.hours - a.hours);
  return { rows, total };
}

/** Admin-only: the dates with the most absence hours, across all programs. */
const DOW_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/**
 * Admin-only: pools every Sunday together, every Monday together, etc.
 * (across every week in the selected period) and shows each weekday's
 * average absenteeism — sorted with the worst (lowest Attendance %) first.
 */
function aggregateAbsenteeismByWeekday(filter) {
  const byDow = new Map(); // 0=Sunday .. 6=Saturday
  for (const r of STATE.records) {
    if (filter.year !== '' && r.year !== filter.year) continue;
    if (filter.month !== '' && r.month !== filter.month) continue;
    const dow = r.date.getDay();
    let d = byDow.get(dow);
    if (!d) { d = { dow, sched: 0, absence: 0, dates: new Set() }; byDow.set(dow, d); }
    d.sched += r.schedAdj;
    d.absence += r.nonDisc;
    d.dates.add(r.dkey);
  }

  const out = [];
  for (let dow = 0; dow < 7; dow++) {
    const d = byDow.get(dow);
    if (!d || d.dates.size === 0) continue;
    const occurrences = d.dates.size;
    out.push({
      name: DOW_NAMES[dow],
      occurrences,
      avgAbsence: d.absence / occurrences,
      pct: d.sched > 0 ? 1 - (d.absence / d.sched) : null,
    });
  }
  assignRelativeBadges(out);
  // out is already in Sunday→Saturday order from the loop above.
  return out;
}

function filterDailyLog(filter) {
  const q = filter.search.trim().toLowerCase();
  return STATE.dailyGroups.filter(g => {
    if (filter.date) {
      if (g.dkey !== filter.date) return false;
    } else {
      if (filter.year !== '' && g.year !== filter.year) return false;
      if (filter.month !== '' && g.month !== filter.month) return false;
    }
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
    updateOverviewSummary(filter);
    return;
  }

  try {
    await ensureWeeksLoadedForView('view-overview', needed);
  } catch (err) {
    if (myGen !== STATE.renderGen.overview) return;
    setTableMessage('view-overview', 'Could not load data for this period', err.message || 'Try again in a moment.');
    updateOverviewSummary(filter);
    return;
  }
  if (myGen !== STATE.renderGen.overview) return; // a newer filter change superseded this render

  const rows = aggregateOverview(filter);
  updateOverviewSummary(filter);
  if (rows.length === 0) {
    setTableMessage('view-overview', 'No records match these filters', 'Try a different year, month, or supervisor.');
    return;
  }

  const tbody = document.getElementById('ov-tbody');
  tbody.innerHTML = rows.map(a => `
    <tr>
      <td class="name-cell">${esc(a.agent)}</td>
      <td>${fmtSupervisorCell(a.agent, a.supervisor)}</td>
      <td>${fmtProgramCell(a.supervisor, a.program)}</td>
      ${STATE.showHours ? `<td class="num">${fmtHours(a.sched)}</td><td class="num">${fmtHours(a.absence)}</td>` : ''}
      <td class="num"><span class="att-badge ${pctBadgeClass(a.pct)}">${fmtPct(a.pct)}</span></td>
    </tr>
  `).join('');
}

function updateOverviewSummary(filter) {
  const el = document.getElementById('ov-summary');
  if (!el) return;
  if (filter.supervisor === '') { el.innerHTML = ''; return; }
  const pct = computeGroupAttendance(r =>
    (filter.year === '' || r.year === filter.year) &&
    (filter.month === '' || r.month === filter.month) &&
    r.supervisor === filter.supervisor
  );
  if (pct === null) { el.innerHTML = ''; return; }
  el.innerHTML = `${esc(filter.supervisor)} team attendance: <span class="att-badge ${pctBadgeClass(pct)}">${fmtPct(pct)}</span>`;
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
      <td>${fmtSupervisorCell(a.agent, a.supervisor)}</td>
      <td>${fmtProgramCell(a.supervisor, a.program)}</td>
      ${STATE.showHours ? `<td class="num">${fmtHours(a.sched)}</td><td class="num">${fmtHours(a.absence)}</td>` : ''}
      <td class="num"><span class="att-badge ${pctBadgeClass(a.pct)}">${fmtPct(a.pct)}</span></td>
    </tr>
  `).join('');
}

/** Puts the same message into all three Programs-view tables at once (loading/empty/error states). */
function setProgramsMessage(bigText, smallText) {
  const specs = [['pg-tbody', 5], ['pg-breakdown-tbody', 3], ['pg-days-tbody', 4]];
  const html = `<tr><td colspan="__COLSPAN__"><div class="empty-state"><div class="big">${esc(bigText)}</div>` +
    (smallText ? `<div class="small">${esc(smallText)}</div>` : '') + `</div></td></tr>`;
  specs.forEach(([id, colspan]) => {
    const tbody = document.getElementById(id);
    if (tbody) tbody.innerHTML = html.replace('__COLSPAN__', colspan);
  });
}

async function renderProgramsView() {
  const myGen = ++STATE.renderGen.programsView;
  const filter = STATE.programsView;
  const needed = weeksOverlapping(filter.year, filter.month);

  if (needed.length === 0) {
    setProgramsMessage('No data available for this period', 'No weekly files were found for this year/month.');
    return;
  }

  try {
    await ensureWeeksLoadedForView('view-programs', needed);
  } catch (err) {
    if (myGen !== STATE.renderGen.programsView) return;
    setProgramsMessage('Could not load data for this period', err.message || 'Try again in a moment.');
    return;
  }
  if (myGen !== STATE.renderGen.programsView) return;

  const programRows = aggregateProgramSummary(filter);
  const { rows: breakdownRows } = aggregateAbsenceBreakdown(filter);
  const dayRows = aggregateAbsenteeismByWeekday(filter);

  if (programRows.length === 0) {
    setProgramsMessage('No records match these filters', 'Try a different year or month.');
    return;
  }

  document.getElementById('pg-tbody').innerHTML = programRows.map(p => `
    <tr>
      <td class="name-cell">${p.program === "Unassigned" ? `<span class="unassigned-tag" data-tooltip="${escAttr('These agents\u2019 supervisors aren\u2019t listed in the program sheet yet (or the agents themselves have no supervisor assigned).')}">Unassigned</span>` : esc(p.program)}</td>
      <td class="num">${p.agentCount.toLocaleString()}</td>
      <td class="num">${fmtHours(p.sched)}</td>
      <td class="num">${fmtHours(p.absence)}</td>
      <td class="num"><span class="att-badge ${pctBadgeClass(p.pct)}">${fmtPct(p.pct)}</span></td>
    </tr>
  `).join('');

  const breakdownBody = document.getElementById('pg-breakdown-tbody');
  if (breakdownRows.length === 0) {
    breakdownBody.innerHTML = `<tr><td colspan="3"><div class="empty-state"><div class="small">No Late/UTO/NCNS/Left Early hours in this period.</div></div></td></tr>`;
  } else {
    breakdownBody.innerHTML = breakdownRows.map(b => `
      <tr>
        <td><span class="pill ${PILL_CLASS[b.type] || 'info'}">${esc(b.type)}</span></td>
        <td class="num">${fmtHours(b.hours)}</td>
        <td class="num">${(b.share * 100).toFixed(1)}%</td>
      </tr>
    `).join('');
  }

  const daysBody = document.getElementById('pg-days-tbody');
  if (dayRows.length === 0) {
    daysBody.innerHTML = `<tr><td colspan="4"><div class="empty-state"><div class="small">No absences recorded in this period.</div></div></td></tr>`;
  } else {
    daysBody.innerHTML = dayRows.map(d => `
      <tr>
        <td class="name-cell">${esc(d.name)}</td>
        <td class="num">${d.occurrences.toLocaleString()}</td>
        <td class="num">${fmtHours(d.avgAbsence)}</td>
        <td class="num"><span class="att-badge ${d.badgeClass}">${fmtPct(d.pct)}</span></td>
      </tr>
    `).join('');
  }
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
  const needed = weeksForDailyLog(filter);

  if (needed.length === 0) {
    setTableMessage('view-dailylog', 'No data available for this period', filter.date ? 'No weekly file covers this date.' : 'No weekly files were found for this year/month.');
    updateDailyLogSummary(filter);
    return;
  }

  try {
    await ensureWeeksLoadedForView('view-dailylog', needed);
  } catch (err) {
    if (myGen !== STATE.renderGen.dailylog) return;
    setTableMessage('view-dailylog', 'Could not load data for this period', err.message || 'Try again in a moment.');
    updateDailyLogSummary(filter);
    return;
  }
  if (myGen !== STATE.renderGen.dailylog) return;

  const rows = filterDailyLog(filter);
  updateDailyLogSummary(filter);
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
      <td>${fmtSupervisorCell(g.agent, g.supervisor)}</td>
      <td>${fmtProgramCell(g.supervisor, g.program)}</td>
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

function updateDailyLogSummary(filter) {
  const el = document.getElementById('dl-summary');
  if (!el) return;
  const q = filter.search.trim();
  const exactAgent = q ? STATE.agentNames.find(n => n.toLowerCase() === q.toLowerCase()) : null;
  if (!exactAgent) { el.innerHTML = ''; return; }
  const pct = computeGroupAttendance(r =>
    (filter.year === '' || r.year === filter.year) &&
    (filter.month === '' || r.month === filter.month) &&
    r.agent === exactAgent
  );
  if (pct === null) { el.innerHTML = ''; return; }
  el.innerHTML = `${esc(exactAgent)} attendance: <span class="att-badge ${pctBadgeClass(pct)}">${fmtPct(pct)}</span>`;
}

/* ---------------- filter UI wiring ---------------- */

function opt(value, label) {
  const o = document.createElement('option');
  o.value = value;
  o.textContent = label;
  return o;
}

/** Rebuilds a section's Month <option> list to match its Year select's current value. */
function refreshMonthOptions(prefix) {
  const yearSel = document.getElementById(prefix + '-year');
  const monthSel = document.getElementById(prefix + '-month');
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

/** Programmatically sets a section's Year/Month selects (and rebuilds Month options) without dispatching change events. */
function setYearMonthSelects(prefix, year, month) {
  const yearSel = document.getElementById(prefix + '-year');
  const monthSel = document.getElementById(prefix + '-month');
  yearSel.value = year === '' ? '' : String(year);
  refreshMonthOptions(prefix);
  monthSel.value = month === '' ? '' : String(month);
}

function populateYearMonth(prefix, currentFilter, onChange) {
  const yearSel = document.getElementById(prefix + '-year');
  const monthSel = document.getElementById(prefix + '-month');

  yearSel.appendChild(opt('', 'All years'));
  STATE.years.forEach(y => yearSel.appendChild(opt(y, y)));
  yearSel.value = currentFilter.year === '' ? '' : String(currentFilter.year);

  refreshMonthOptions(prefix);
  monthSel.value = currentFilter.month === '' ? '' : String(currentFilter.month);

  yearSel.addEventListener('change', () => { refreshMonthOptions(prefix); onChange(); });
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

function setupProgramsFilters() {
  populateYearMonth('pg', STATE.programsView, () => {
    STATE.programsView.year = document.getElementById('pg-year').value === '' ? '' : Number(document.getElementById('pg-year').value);
    STATE.programsView.month = document.getElementById('pg-month').value === '' ? '' : Number(document.getElementById('pg-month').value);
    renderProgramsView();
    updateMeta('pg');
  });

  document.getElementById('pg-clear').addEventListener('click', () => {
    document.getElementById('pg-year').value = '';
    document.getElementById('pg-year').dispatchEvent(new Event('change'));
    renderProgramsView();
    updateMeta('pg');
  });

  updateMeta('pg');
}

function setupDailyLogFilters() {
  populateYearMonth('dl', STATE.dailylog, () => {
    STATE.dailylog.year = document.getElementById('dl-year').value === '' ? '' : Number(document.getElementById('dl-year').value);
    STATE.dailylog.month = document.getElementById('dl-month').value === '' ? '' : Number(document.getElementById('dl-month').value);
    if (STATE.dailylog.date) {
      STATE.dailylog.date = '';
      document.getElementById('dl-date').value = '';
    }
    renderDailyLog();
    updateMeta('dl');
  });

  const dateInput = document.getElementById('dl-date');
  if (STATE.availableWeeks.length > 0) {
    const earliest = STATE.availableWeeks[0].sunday;
    const latest = STATE.availableWeeks[STATE.availableWeeks.length - 1].saturday;
    dateInput.min = dateKey(earliest);
    dateInput.max = dateKey(latest);
  }
  dateInput.addEventListener('change', () => {
    STATE.dailylog.date = dateInput.value; // already "YYYY-MM-DD", matches dkey format directly
    if (dateInput.value) {
      const d = parseISODateLocal(dateInput.value);
      STATE.dailylog.year = d.getFullYear();
      STATE.dailylog.month = d.getMonth() + 1;
      setYearMonthSelects('dl', STATE.dailylog.year, STATE.dailylog.month);
    }
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
    dateInput.value = '';
    STATE.dailylog.date = '';
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
  const s = STATE[prefix === 'ov' ? 'overview' : prefix === 'wl' ? 'watchlist' : prefix === 'pg' ? 'programsView' : 'dailylog'];
  if (prefix === 'dl' && s.date) {
    parts.push(parseISODateLocal(s.date).toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric' }));
  } else {
    if (s.year !== '') parts.push(s.year);
    if (s.month !== '') parts.push(MONTH_NAMES[s.month-1]);
  }
  if (prefix === 'ov' && s.supervisor) parts.push(s.supervisor);
  if (prefix === 'wl' && s.program) parts.push(s.program);
  if (prefix === 'dl' && s.search) parts.push('"' + s.search + '"');
  el.textContent = parts.length ? parts.join(' · ') : 'All records';
}

/* ---------------- secret typed-word triggers ----------------
   Type "hours" anywhere on the page (outside a text field) to show/hide
   the Sched Hours and Absence Hours columns in Overview and Watchlist.
   Type "gwapo" to capture a screenshot of whichever table is currently on
   screen (same action as the Konami code below — just an easier-to-type
   alternative). No visible button/menu for either, on purpose — type
   "hours" again to hide the columns. */

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
    } else if (buffer.endsWith('gwapo')) {
      captureVisibleTableScreenshot();
      buffer = '';
    }
  });
}

/* ---------------- secret screenshot (Konami code) ----------------
   ↑ ↑ ↓ ↓ ← → ← → B A anywhere outside a text field captures whichever
   table is currently on screen — respecting its current filters — as a
   clean, compact PNG. It's copied straight to the clipboard so it can be
   pasted directly into an email; if the browser won't allow that, it
   downloads instead. No visible button or hint anywhere in the UI. */

const KONAMI_SEQUENCE = ['ArrowUp','ArrowUp','ArrowDown','ArrowDown','ArrowLeft','ArrowRight','ArrowLeft','ArrowRight','b','a'];
const SCREENSHOT_MAX_ROWS = 300;

function buildFilterTitle(label, filter) {
  const parts = [];
  if (filter.year !== undefined && filter.year !== '') parts.push(String(filter.year));
  if (filter.month !== undefined && filter.month !== '') parts.push(MONTH_NAMES[filter.month - 1]);
  if (filter.supervisor) parts.push(filter.supervisor);
  if (filter.program) parts.push(filter.program);
  if (filter.search) parts.push('"' + filter.search + '"');
  return label + (parts.length ? ' — ' + parts.join(' · ') : ' — All records');
}

/** Pulls together whichever table is currently visible, using the data already loaded (no new fetch). */
function getVisibleTableData() {
  const activeBtn = document.querySelector('.tab-btn.active');
  const view = activeBtn ? activeBtn.dataset.view : 'overview';

  if (view === 'overview') {
    const filter = STATE.overview;
    const rows = aggregateOverview(filter);
    const columns = ['Agent Name', 'Supervisor', 'Program'];
    if (STATE.showHours) columns.push('Sched Hours', 'Absence Hours');
    columns.push('Attendance %');
    const dataRows = rows.slice(0, SCREENSHOT_MAX_ROWS).map(a => {
      const r = [a.agent, a.supervisor, a.program];
      if (STATE.showHours) r.push(fmtHours(a.sched), fmtHours(a.absence));
      r.push(fmtPct(a.pct));
      return r;
    });
    return { title: buildFilterTitle('Overview', filter), columns, rows: dataRows, totalRows: rows.length, filename: 'overview' };
  }

  if (view === 'watchlist') {
    const filter = STATE.watchlist;
    const rows = aggregateWatchlist(filter);
    const columns = ['#', 'Agent Name', 'Supervisor', 'Program'];
    if (STATE.showHours) columns.push('Sched Hours', 'Absence Hours');
    columns.push('Attendance %');
    const dataRows = rows.slice(0, SCREENSHOT_MAX_ROWS).map((a, i) => {
      const r = [String(i + 1), a.agent, a.supervisor, a.program];
      if (STATE.showHours) r.push(fmtHours(a.sched), fmtHours(a.absence));
      r.push(fmtPct(a.pct));
      return r;
    });
    return { title: buildFilterTitle('Watchlist', filter), columns, rows: dataRows, totalRows: rows.length, filename: 'watchlist' };
  }

  if (view === 'programs') {
    const filter = STATE.programsView;
    const rows = aggregateProgramSummary(filter);
    const columns = ['Program', 'Agents', 'Sched Hours', 'Absence Hours', 'Attendance %'];
    const dataRows = rows.slice(0, SCREENSHOT_MAX_ROWS).map(p => [
      p.program, String(p.agentCount), fmtHours(p.sched), fmtHours(p.absence), fmtPct(p.pct),
    ]);
    return { title: buildFilterTitle('Program Summary', filter), columns, rows: dataRows, totalRows: rows.length, filename: 'program-summary' };
  }

  // Daily Log
  const filter = STATE.dailylog;
  const rows = filterDailyLog(filter);
  const columns = ['Date', 'Agent Name', 'Supervisor', 'Program', 'Status', 'Attendance %'];
  const dataRows = rows.slice(0, SCREENSHOT_MAX_ROWS).map(g => [
    fmtDate(g.date), g.agent, g.supervisor, g.program,
    g.status === 'present' ? 'Present' : g.status.map(e => `${e.type} (${fmtHours(e.hours)} hrs)`).join(', '),
    fmtPct(g.pct),
  ]);
  return { title: buildFilterTitle('Daily Log', filter), columns, rows: dataRows, totalRows: rows.length, filename: 'daily-log' };
}

/** Draws a plain, print/email-friendly table (no dark theme, no colored pills) onto a canvas. */
function renderCompactTableToCanvas(data) {
  const DESCALE = 0.576; // 0.8 × 0.8 × 0.9 — three successive reductions (20%, 20%, 10%) from the original layout
  const PADDING = 16 * DESCALE;
  const ROW_H = 26 * DESCALE;
  const HEADER_H = 30 * DESCALE;
  const NOTE_H = data.totalRows > data.rows.length ? 18 * DESCALE : 0;
  const CELL_PAD = 10 * DESCALE;
  const FONT = `${(12 * DESCALE).toFixed(1)}px Arial, sans-serif`;
  const HEADER_FONT = `bold ${(12 * DESCALE).toFixed(1)}px Arial, sans-serif`;
  const NOTE_FONT = `${(11 * DESCALE).toFixed(1)}px Arial, sans-serif`;

  const scratch = document.createElement('canvas').getContext('2d');
  const colWidths = data.columns.map((col, i) => {
    scratch.font = HEADER_FONT;
    let max = scratch.measureText(col).width;
    scratch.font = FONT;
    for (const row of data.rows) {
      const w = scratch.measureText(String(row[i] ?? '')).width;
      if (w > max) max = w;
    }
    return Math.ceil(max) + CELL_PAD * 2;
  });

  const tableWidth = colWidths.reduce((a, b) => a + b, 0);
  const width = tableWidth + PADDING * 2;
  const height = NOTE_H + HEADER_H + data.rows.length * ROW_H + PADDING * 2;

  const scale = 2; // render at 2x for a crisp paste
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.ceil(width * scale));
  canvas.height = Math.max(1, Math.ceil(height * scale));
  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);
  ctx.textBaseline = 'middle';

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  let y = PADDING;

  if (NOTE_H) {
    ctx.fillStyle = '#9aa0ae';
    ctx.font = NOTE_FONT;
    ctx.fillText(`Showing first ${data.rows.length} of ${data.totalRows} rows`, PADDING, y + NOTE_H / 2);
    y += NOTE_H;
  }

  const tableTop = y;
  ctx.fillStyle = '#eef0f4';
  ctx.fillRect(PADDING, y, tableWidth, HEADER_H);
  ctx.fillStyle = '#1e2a47';
  ctx.font = HEADER_FONT;
  let x = PADDING;
  data.columns.forEach((col, i) => {
    ctx.fillText(col, x + CELL_PAD, y + HEADER_H / 2);
    x += colWidths[i];
  });
  y += HEADER_H;

  ctx.font = FONT;
  data.rows.forEach((row, rIdx) => {
    if (rIdx % 2 === 1) {
      ctx.fillStyle = '#f7f8fa';
      ctx.fillRect(PADDING, y, tableWidth, ROW_H);
    }
    ctx.fillStyle = '#1d2233';
    let cx = PADDING;
    row.forEach((cell, i) => {
      ctx.fillText(String(cell ?? ''), cx + CELL_PAD, y + ROW_H / 2);
      cx += colWidths[i];
    });
    y += ROW_H;
  });

  ctx.strokeStyle = '#d8dae0';
  ctx.lineWidth = 1;
  const tableHeight = HEADER_H + data.rows.length * ROW_H;
  ctx.strokeRect(PADDING + 0.5, tableTop + 0.5, tableWidth - 1, tableHeight - 1);
  let ly = tableTop + HEADER_H;
  for (let i = 0; i < data.rows.length; i++) {
    ctx.beginPath();
    ctx.moveTo(PADDING, ly + 0.5);
    ctx.lineTo(PADDING + tableWidth, ly + 0.5);
    ctx.stroke();
    ly += ROW_H;
  }

  return canvas;
}

async function captureVisibleTableScreenshot() {
  const data = getVisibleTableData();
  if (!data || data.rows.length === 0) {
    showToast('Nothing to capture');
    return;
  }
  const canvas = renderCompactTableToCanvas(data);
  canvas.toBlob(async (blob) => {
    if (!blob) { showToast('Screenshot failed'); return; }
    try {
      if (!navigator.clipboard || !window.ClipboardItem) throw new Error('Clipboard image API unavailable');
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
      showToast('Screenshot copied — paste it into your email');
    } catch (err) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${data.filename}-${dateKey(new Date())}.png`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      showToast('Screenshot downloaded');
    }
  }, 'image/png');
}

function setupKonamiScreenshot() {
  let buffer = [];
  document.addEventListener('keydown', (e) => {
    const t = e.target;
    const isTyping = t && (t.tagName === 'INPUT' || t.tagName === 'SELECT' || t.tagName === 'TEXTAREA');
    if (isTyping) return;
    const key = e.key.length === 1 ? e.key.toLowerCase() : e.key;
    buffer.push(key);
    if (buffer.length > KONAMI_SEQUENCE.length) buffer.shift();
    if (buffer.length === KONAMI_SEQUENCE.length && buffer.every((k, i) => k === KONAMI_SEQUENCE[i])) {
      buffer = [];
      captureVisibleTableScreenshot();
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
    const wb = await fetchRosterWorkbook();
    const accessMap = parseAccessSheet(wb);

    setLoading(null); // hide the loading overlay so the login screen (if needed) is visible
    const { username, allowedPrograms } = await runLoginGate(accessMap);
    STATE.currentUser = username;

    setLoading('PARSING roster …');
    buildRosterFromWorkbook(wb);
    applyProgramRestriction(allowedPrograms);

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
    STATE.programsView.year = defaultYear; STATE.programsView.month = defaultMonth;

    document.getElementById('weeksCount').textContent =
      STATE.availableWeeks.length.toLocaleString() + (STATE.availableWeeks.length === 1 ? ' week available' : ' weeks available');
    document.getElementById('agentsCount').textContent =
      STATE.agentNames.length.toLocaleString() + ' agents';
    updateSignedInStatus();

    setupTabs();
    setupSecretToggle();
    setupKonamiScreenshot();
    setupOverviewFilters();
    setupDailyLogFilters();
    setupWatchlistFilters();

    const isAdmin = STATE.allowedPrograms === null;
    if (isAdmin) {
      document.getElementById('tab-programs').style.display = '';
      setupProgramsFilters();
    }

    setLoading(null);

    renderOverview();
    renderDailyLog();
    renderWatchlist();
    if (isAdmin) renderProgramsView();
  } catch (err) {
    console.error(err);
    setLoading(null);
    showError(err.message || 'Something went wrong loading the dashboard.');
  }
}

boot();
