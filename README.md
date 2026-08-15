# Attendance Console

A single-page, fixed-frame attendance dashboard that reads `attendance_raw.xlsx`
directly in the browser (via SheetJS) — no backend, no build step.

## Files
- `index.html` — page structure, 3 tabbed sections
- `style.css` — design system / styling
- `app.js` — loads the workbook, aggregates data, wires up filters and tables
- `attendance_raw.xlsx` — your source data (**must stay in the same folder** as `index.html`)

## How it works
On load, the app fetches `./attendance_raw.xlsx` and parses three sheets:
- `raw` — the row-level event data
- `supervisor` — Agent → Supervisor mapping
- `program` — Supervisor → Program mapping

It then builds two in-memory tables:
1. **Row-level records** (one per raw row), used for Sections 1 and 3
2. **Daily groups** (one per agent + date), used for Section 2

### Section 1 — Overview
Groups records by agent (after Year/Month/Supervisor filters) and computes:
- **Sched hours** — sum of `Sched hours`, but rows with Event type `Lunch`, `PTO`,
  `UPL`, `FMLA`, `PFML`, `Bereavement`, or `Alt Holiday` are treated as **0** per
  your instructions (these aren't real working schedule time).
- **Absence hours** — sum of `Non-discretionary shrinkage`.
- **Attendance %** — 100% minus (Absence hours ÷ Sched hours). Higher is
  better. Shown as `N/A` when Sched hours is 0 for that agent/period (can't
  divide by zero).

### Section 2 — Daily Log
For each agent + date, collects every exception-type row (`PTO`, `UPL`, `FMLA`,
`PFML`, `Late`, `Left Early`, `UTO`, `Bereavement`, `Alt Holiday`, `NCNS`) and
shows them as `Type (X hrs)`, joined with `; ` if there's more than one that day.
Zero-hour placeholder rows (a data quirk in the source — e.g. a `0 hr` PTO stub
sitting alongside a real `Late` entry on the same day) are dropped from display
since they carry no information.

If an agent has no exceptions that day but does have both Sched hours and Work
hours, the status shows **Present**. Days with no data at all for an agent are
simply not shown (they weren't scheduled).

Each row also shows the agent's **Attendance %** for the selected Year/Month —
the same figure as Section 1, so you can see a given day's status alongside
that agent's overall attendance for the period. It's repeated on every row for
that agent (not a per-day metric, since attendance % only makes sense over a
period).

The Employee Search box filters by a "contains" match as you type, and the
dropdown suggestion list lets you jump straight to one agent.

### Section 3 — Watchlist
Same aggregation as Section 1, filtered by Program instead of Supervisor, then
sorted so the **lowest** Attendance % (worst attendance) is first, and capped
to the top 10. Agents with no Sched hours in the period (percentage undefined)
are excluded from ranking, since there's nothing to rank.

### Data quirks handled
- 2 agents (`Baguyo, Shane`, `Nee, Emily`) appear in `raw` but aren't in the
  `supervisor` sheet — they display as Supervisor/Program = "Unassigned" rather
  than being dropped.
- 7 supervisors in the `supervisor` sheet have no row in `program` — their
  agents show Program = "Unassigned".
- The `program` sheet has the value `988` stored as both text and a number in
  different rows — the app normalizes everything to text so they merge into a
  single filter option.

## Deploying to GitHub + Netlify
1. Create a new GitHub repo and add these four files to its root (or any single
   folder, as long as they're all together): `index.html`, `style.css`,
   `app.js`, `attendance_raw.xlsx`.
2. Commit and push.
3. In Netlify: **Add new site → Import an existing project**, connect the repo,
   leave the build command blank, and set the publish directory to wherever
   these files live (repo root, or that folder).
4. Deploy. No build step is required — it's static HTML/CSS/JS.

## Updating the data
To refresh the dashboard with a new period, just replace `attendance_raw.xlsx`
in the repo with your updated export (same sheet names: `raw`, `supervisor`,
`program`) and push. The Year/Month/Supervisor/Program filter options are all
derived from the file automatically — no code changes needed unless your
column headers change.

If a column header in your export ever changes (e.g. `Sched hours` becomes
`Scheduled Hours`), update the matching string in `app.js` — search for the
column name in quotes.
