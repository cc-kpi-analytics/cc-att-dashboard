/* ============================================================
   Login gate — username/password sourced from the "access" sheet
   in attendance_raw.xlsx.
   ============================================================
   This is a UI convenience, not real security: attendance_raw.xlsx
   (including every plaintext password in it) is fetched over a public,
   unauthenticated URL exactly like every other file in this app, and
   could be read directly by anyone who has or guesses the link — signed
   in or not. This only gates what the DASHBOARD shows once someone has
   typed in valid-looking credentials; it does not make the underlying
   files themselves inaccessible. See the README for more on this
   tradeoff. Don't reuse these passwords anywhere that matters. */

const SESSION_KEY = 'attendanceConsoleUser';

/** Parses the "access" sheet into Map(usernameLower -> {username, password, programs:Set, isAdmin}). */
function parseAccessSheet(wb) {
  const map = new Map();
  if (!wb.SheetNames.includes('access')) return map;

  const rows = XLSX.utils.sheet_to_json(wb.Sheets['access'], { defval: null });
  rows.forEach(r => {
    const username = r['Username'];
    const password = r['Password'];
    const program = r['Program'];
    if (!username || password === null || password === undefined || password === '') return;

    const key = String(username).trim().toLowerCase();
    const progStr = program !== null && program !== undefined ? String(program).trim() : '';
    const isAdminRow = progStr.toLowerCase() === 'admin';

    let entry = map.get(key);
    if (!entry) {
      entry = { username: String(username).trim(), password: String(password), programs: new Set(), isAdmin: false };
      map.set(key, entry);
    }
    if (isAdminRow) entry.isAdmin = true;
    else if (progStr) entry.programs.add(progStr);
  });
  return map;
}

function getStoredUsername() {
  try { return sessionStorage.getItem(SESSION_KEY); } catch (e) { return null; }
}

function storeUsername(username) {
  try { sessionStorage.setItem(SESSION_KEY, username); } catch (e) { /* ignore (e.g. private browsing) */ }
}

function clearStoredUsername() {
  try { sessionStorage.removeItem(SESSION_KEY); } catch (e) { /* ignore */ }
}

/**
 * Resolves with { username, allowedPrograms } once the user is signed in —
 * either instantly from a still-valid stored session, or after they submit
 * correct credentials in the login screen. allowedPrograms is null for an
 * admin row (full access) or a Set of program names otherwise.
 */
function runLoginGate(accessMap) {
  return new Promise((resolve) => {
    const stored = getStoredUsername();
    if (stored) {
      const entry = accessMap.get(stored.toLowerCase());
      if (entry) {
        resolve({ username: entry.username, allowedPrograms: entry.isAdmin ? null : entry.programs });
        return;
      }
      clearStoredUsername(); // the stored account no longer exists in the sheet
    }

    const screen = document.getElementById('loginScreen');
    const form = document.getElementById('loginForm');
    const userInput = document.getElementById('loginUsername');
    const passInput = document.getElementById('loginPassword');
    const errorEl = document.getElementById('loginError');

    screen.classList.remove('hidden');
    userInput.focus();

    form.addEventListener('submit', function onSubmit(e) {
      e.preventDefault();
      const u = userInput.value.trim();
      const p = passInput.value;
      const entry = u ? accessMap.get(u.toLowerCase()) : null;

      if (!entry || entry.password !== p) {
        errorEl.textContent = 'Incorrect username or password.';
        errorEl.style.display = '';
        passInput.value = '';
        passInput.focus();
        return;
      }

      form.removeEventListener('submit', onSubmit);
      storeUsername(entry.username);
      screen.classList.add('hidden');
      resolve({ username: entry.username, allowedPrograms: entry.isAdmin ? null : entry.programs });
    });
  });
}

function signOutUser() {
  clearStoredUsername();
  window.location.reload();
}
