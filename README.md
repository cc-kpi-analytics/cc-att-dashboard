# Attendance Console

A single-page, fixed-frame attendance dashboard that reads weekly schedule
exports directly in the browser — no backend, no build step.

## Files
- `index.html` — page structure, 3 tabbed sections
- `style.css` — design system / styling
- `mhtml-parser.js` — turns a saved Five9 schedule report (.mhtml) into plain interval records
- `app.js` — discovers weekly files, fetches/parses them on demand, aggregates, renders
- `attendance_raw.xlsx` — **only used for the `supervisor` and `program` sheets** now (see below)
- `data/WB_YYYY_MM_DD.mhtml` — one file per week of raw schedule data (you add a new one every week)

## The weekly data workflow

Every week, from the Five9 "Published Schedule Detail" report (Printer View
By Agent), save the page as **Webpage, Single File** — this produces a
`.mhtml` file. Rename it to match the week's **Sunday** date and drop it in
`data/`:

```
data/WB_2026_08_23.mhtml   ← the week of Sun Aug 23 – Sat Aug 29, 2026
data/WB_2026_08_30.mhtml   ← the following week
```

**Naming is strict:** `WB_` + 4-digit year + `_` + 2-digit month + `_` +
2-digit day (of the Sunday the week starts on), extension `.mhtml`. Nothing
else needs to change — no manifest to edit, no code to touch. Push the file
to your repo and the app finds it automatically the next time someone loads
the page.

### How the app finds your weekly files
Netlify (like any static host) can't list a folder's contents on request, so
the app **auto-discovers** which weeks exist: on load, it computes every
Sunday in a configurable date range and sends a lightweight `HEAD` request
for each one's expected filename — no file, no data downloaded, it just
silently moves on. Only weeks that exist are downloaded in full, and only
when a section's filters actually need them.

That range is set in `app.js`:
```js
const DISCOVERY_YEARS_BACK = 2;      // how far back to look, from Jan 1 of (this year − 2)
const DISCOVERY_DAYS_FORWARD = 120;  // how far into the future to look, from today
```
If you ever need older history than that, widen `DISCOVERY_YEARS_BACK`.

### Why fetching is lazy, and what that means day-to-day
Each `.mhtml` file is a few MB (mostly page markup, not data — see below), so
the app doesn't fetch everything up front. Instead:
- On load, it only fetches the **most recently available week's** month for
  each section, so the first thing you see loads fast.
- When you change a Year/Month filter, it fetches whichever weeks overlap
  that period (a week can span two months — both get counted correctly) and
  shows a short "Loading N weeks of data…" message in that section's table
  while it works.
- Once a week is fetched, it's cached in memory for the rest of the session —
  switching back and forth between periods you've already viewed is instant.
- Picking "All years" fetches every available week — that's a lot of data on
  a site with a long history, so it'll take longer the first time.

## How the data is derived (this changed from the old Excel-based version)

The old `raw` sheet had pre-computed columns (Sched hours, Work hours,
Non-discretionary shrinkage, etc.) for each agent/day/event-type combo. The
weekly `.mhtml` report doesn't — it's raw **schedule intervals**: one row per
block of time, with a Start Time, End Time, an event `Type` (e.g. `Work -
CCS/Chat`, `Break`, `Lunch`, `PTO`, `Late`, `UTO`...), and a `Duration`. The
app derives the same figures it used to trust from Excel, now computed
directly from those intervals:

- **Sched hours** — sum of `Duration` for every interval that day, except
  intervals typed `Lunch`, `PTO`, `UPL`, `FMLA`, `PFML`, `Bereavement`, or
  `Alt Holiday` (same exclusion rule as before — these aren't real scheduled
  working time).
- **Absence hours** (non-discretionary shrinkage) — sum of `Duration` for
  intervals typed `Late`, `Left Early`, `UTO`, or `NCNS`.
- **Attendance %** — 100% − (Absence hours ÷ Sched hours), same formula as
  before, just fed by the numbers above.
- **"Work hours"** (used only to decide whether a day counts as *Present* in
  the Daily Log) — sum of `Duration` for intervals typed `Work - …`,
  `Meeting`, `Training`, or `Admin`. If an agent has no exceptions that day
  and this is greater than zero, the day shows **Present**.
- **Daily Log Status** — every exception-type interval that day (`PTO`,
  `UPL`, `FMLA`, `PFML`, `Late`, `Left Early`, `UTO`, `Bereavement`, `Alt
  Holiday`, `NCNS`) is **summed by type** (a day can have the same exception
  type split across two separate intervals — e.g. FMLA taken in a morning
  block and an afternoon block — these get merged into one total rather than
  shown as two separate entries) and displayed as `Type (X hrs)`, joined by
  `; ` if there's more than one type that day. Zero-hour placeholder rows are
  dropped since they carry no information.

If Five9 ever adds a new event type the app doesn't recognize, it falls
through safely: it counts toward Sched hours (like Break/Meeting do) but
isn't treated as an absence or as "work hours" — nothing breaks, it just
won't be specially categorized until you decide it should be.

## Section 1 — Overview
Filters: Year, Month, Supervisor. Shows Agent Name, Supervisor, Program, and
Attendance % (Sched/Absence Hours are hidden by default — see below).

## Section 2 — Daily Log
Filters: Year, Month, Employee Search (type-ahead). Shows Date, Agent Name,
Supervisor, Program, Status, and that date's Attendance %.

## Section 3 — Watchlist
Filters: Year, Month, Program. Shows the 10 agents with the **lowest**
Attendance % (worst attendance) for the selected period, ranked.

## Sched/Absence Hours (hidden by request)
Sched Hours and Absence Hours are hidden from Overview and Watchlist by
default. There's no visible button for this — on any page, just type the
word `hours` (no need to click or hold anything) and both columns toggle
on/off; type it again to hide them. It only fires outside of a text field,
so it won't trigger while typing in Employee Search. Nothing in the UI
hints that this exists.

## The Excel file's remaining role
`attendance_raw.xlsx` is still required, but now **only** for its
`supervisor` sheet (Agent → Supervisor) and `program` sheet (Supervisor →
Program). Its `raw` sheet is no longer read at all. Update this file
whenever your roster or program assignments change — it doesn't need to be
touched weekly the way the schedule data does.

### Data quirks handled
- Agents that appear in a weekly `.mhtml` file but aren't (yet) in the
  `supervisor` sheet show as Supervisor/Program = "Unassigned" rather than
  being dropped — useful when new hires show up in schedules before the
  roster spreadsheet catches up.
- Supervisors with no row in `program` show their agents' Program as
  "Unassigned".
- The `program` sheet's value `988` (which Excel sometimes stores as text,
  sometimes as a number) is normalized to text so it merges into one filter
  option instead of two.

## Deploying to GitHub + Netlify
1. Create a repo (or use your existing one) with this file layout:
   ```
   index.html
   style.css
   mhtml-parser.js
   app.js
   attendance_raw.xlsx
   data/
     WB_2026_08_23.mhtml
     WB_2026_08_30.mhtml
     ...
   ```
2. Commit and push.
3. In Netlify: **Add new site → Import an existing project**, connect the
   repo, leave the build command blank, and set the publish directory to
   wherever these files live (repo root, or that folder). No build step —
   it's static HTML/CSS/JS.
4. Every week, add the new `data/WB_YYYY_MM_DD.mhtml` file and push. That's
   the only ongoing maintenance the data side needs.

## A note on the .mhtml format
`.mhtml` ("Webpage, Single File") is a MIME-wrapped archive of a saved
webpage — it bundles the page's HTML together with its styling/images in one
file. The app only needs the HTML part, which `mhtml-parser.js` extracts and
decodes (Five9's export uses quoted-printable encoding) entirely client-side
before parsing it for schedule data. If your browser ever offers a lighter
plain-`.html` export instead, that would work too with a small tweak to the
fetch logic — let me know if that becomes available and I can wire it up.
