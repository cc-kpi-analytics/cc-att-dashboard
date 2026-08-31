# Attendance Console

A single-page, fixed-frame attendance dashboard that reads weekly schedule
exports directly in the browser — no backend, no build step.

## Files
- `index.html` — page structure, login screen, 3 tabbed sections
- `style.css` — design system / styling
- `mhtml-parser.js` — turns a saved Five9 schedule report (.mhtml) into plain interval records
- `auth.js` — login gate, sourced from the `access` sheet in `attendance_raw.xlsx`
- `app.js` — discovers weekly files, fetches/parses them on demand, aggregates, renders
- `attendance_raw.xlsx` — the `supervisor` and `program` sheets, plus the `access` sheet (see "Signing in" below)
- `data/WB_YYYY_MM_DD.mhtml` — one file per week of raw schedule data (you add a new one every week)

## Signing in

The dashboard now sits behind a simple login screen, sourced entirely from
the `access` sheet in `attendance_raw.xlsx`:

| Username | Password | Program |
|---|---|---|
| rcl-leadership | Crisis*2026 | RCL |
| admin-attendance | WFM@crisis2026 | Admin |

- One row per account. A `Program` value scopes that account to just that
  program — Overview/Daily Log/Bottom 15 only show agents in it, including
  the filter dropdowns and Employee Search suggestions.
- A `Program` value of exactly `Admin` (any capitalization) grants full,
  unrestricted access to everything — that's what `admin-attendance` does
  by default.
- Usernames are matched without regard to case; passwords are matched
  exactly as typed.
- Once someone signs in, their browser remembers it for that session (it
  clears when they close the tab/browser) — they won't have to log in
  again on every page refresh. A "Sign out" link sits in the header next
  to the weeks/agents counter.
- To add, remove, or change an account, just edit the rows in the `access`
  sheet and push the updated `attendance_raw.xlsx` — no code changes
  needed.

**Please read this before treating it as real security.** This is a UI
convenience, not a lock. `attendance_raw.xlsx` — including every password
in that sheet, in plain text — sits in the same public, unauthenticated
`data/` folder as everything else in this app. Anyone who fetches that
file directly (or opens the browser's network tab while using the
dashboard) can read every username and password, restricted or not. This
gate stops someone from casually clicking into another program's data
inside the dashboard; it does not make the file itself inaccessible to a
determined or technical person. Don't reuse these passwords anywhere that
actually matters, and don't treat this as sufficient if the data ever
needs to be genuinely locked down — that would require a real backend in
front of the files, which is a meaningfully bigger project than this.

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

- **Sched hours** — sum of `Duration` for every interval that day, **except**
  intervals typed `Lunch`, `PTO`, `UPL`, `FMLA`, `PFML`, `Bereavement`, `Alt
  Holiday`, `Jury Duty`, `Traumatic Leave`, or plain `Admin` (as opposed to
  `Work - Admin`, which is real work and counts normally). These are treated
  as unscheduled/non-working time.
- **Absence hours** ("unapproved shrink") — sum of `Duration` for intervals
  typed `Late`, `Left Early`, `UTO`, or `NCNS`.
- **Attendance %** — 100% − (Absence hours ÷ Sched hours).
- **Present** — if an agent has no absence-status interval that day and
  Sched hours (as defined above) is greater than zero, the day shows
  **Present**.
- **Daily Log Status** — every absence-status interval that day (`Late`,
  `UTO`, `NCNS`, `Left Early`, `FMLA`, `PFML`, `UPL`, `Traumatic Leave`, `Alt
  Holiday`, `PTO`, `Bereavement`, `Jury Duty`, plain `Admin`) is **summed by
  type** (a day can have the same status split across two separate intervals
  — e.g. FMLA taken in a morning block and an afternoon block — these get
  merged into one total rather than shown as two separate entries) and
  displayed as `Type (X hrs)`, separated by a comma if there's more than one
  status that day. Zero-hour placeholder rows are dropped since they carry
  no information. `Admin` is included here (not just excluded from Sched
  hours) so an Admin-only day still shows something in the log instead of
  silently disappearing.

If Five9 ever adds a new event type the app doesn't recognize, it falls
through safely: it counts toward Sched hours (like Break/Meeting/Work-*
types do) but isn't treated as an absence status — nothing breaks, it just
won't be specially categorized until you tell me it needs to be one or the
other.

## Section 1 — Overview
Filters: Year, Month, Supervisor. Shows Agent Name, Supervisor, Program, and
Attendance % (Sched/Absence Hours are hidden by default — see below).

## Section 2 — Daily Log
Filters: Year, Month, a specific Date (optional — picking one narrows to
just that day and syncs Year/Month to match it; changing Year or Month
afterward clears the picked date), and Employee Search (type-ahead). Shows
Date, Agent Name, Supervisor, Program, Status, and that date's Attendance %.

## Section 3 — Bottom 15
Filters: Year, Month, Program. Shows the 15 agents with the **lowest**
Attendance % (worst attendance) for the selected period, ranked.

## Sched/Absence Hours (hidden by request)
Sched Hours and Absence Hours are hidden from Overview and Bottom 15 by
default. There's no visible button for this — on any page, just type the
word `hours` (no need to click or hold anything) and both columns toggle
on/off; type it again to hide them. It only fires outside of a text field,
so it won't trigger while typing in Employee Search. Nothing in the UI
hints that this exists.

## Screenshot-to-clipboard (type "gwapo", or the Konami code)
Two ways to trigger the same thing — whichever's easier to remember:
- Type the word **`gwapo`** anywhere on the page (outside a text field), or
- Press **↑ ↑ ↓ ↓ ← → ← → B A** (the Konami code)

Either one renders whichever table is currently on screen — respecting
whatever filters are applied — as a clean, compact PNG (plain white
background, black text, no dark theme or colored pills, no title/filter
text, sized to fit just the table itself and scaled down to about 58% of
its original size — three successive reductions of 20%, 20%, then 10%) and
copies it straight to your clipboard, so you can paste it directly into an
email or chat. If
your browser won't allow writing images to the clipboard (some browsers
restrict this), it downloads the PNG instead and tells you so via the same
small toast used for the `hours` trick.

It's capped at 300 rows to keep the image a reasonable size for something
you'd paste into an email — if the table has more than that, a small note
on the image says how many rows were left out.

## The Excel file's remaining role
`attendance_raw.xlsx` is still required, but now **only** for its
`supervisor` sheet (Agent → Supervisor), `program` sheet (Supervisor →
Program), and `access` sheet (login accounts — see "Signing in" above).
Its `raw` sheet is no longer read at all. Update this file whenever your
roster, program assignments, or accounts change — it doesn't need to be
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
   auth.js
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
