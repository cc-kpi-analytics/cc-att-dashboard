/* ============================================================
   Login gate — Netlify Identity for real authentication (hosted
   sign-up, sign-in, and self-serve "forgot password" — Netlify
   stores and hashes the passwords, nothing password-related ever
   touches attendance_raw.xlsx), with program access still driven
   by the "credentials" sheet in attendance_raw.xlsx.
   ============================================================
   What Netlify Identity gives us: an account system with no backend
   code to write. Sign up, sign in, email confirmation, and password
   reset are all handled by Netlify's hosted Identity (GoTrue) service
   through the netlify-identity-widget script tag in index.html — see
   README.md for the one-time Netlify dashboard setup.

   What it does NOT know about: which program an email belongs to.
   That mapping still lives in the "credentials" sheet of
   attendance_raw.xlsx (one column per program, listing the emails
   that belong to it; an "Admin" column grants full access). Anyone
   can create an Identity account, but the dashboard only unlocks for
   an email that also appears in that sheet — everyone else sees an
   "ask your admin" screen instead of any data. Because that sheet is
   fetched over a public, unauthenticated URL like every other file in
   this app, treat it as a directory, not a secret: it decides *scope*
   (which program someone sees), not identity or passwords — those are
   Netlify's job now. */

/**
 * Parses the "credentials" sheet into Map(emailLower -> {programs:Set, isAdmin}).
 * Layout: row 1 has one program name per column (one of them literally
 * "Admin"); every cell below a column is an email address that belongs
 * to that program. Columns are independent and can be different lengths.
 */
function parseCredentialsSheet(wb) {
  const map = new Map();
  if (!wb.SheetNames.includes('credentials')) return map;

  const rows = XLSX.utils.sheet_to_json(wb.Sheets['credentials'], { header: 1, defval: null });
  if (rows.length === 0) return map;

  const headers = rows[0].map(h => (h === null || h === undefined) ? '' : String(h).trim());

  for (let c = 0; c < headers.length; c++) {
    const program = headers[c];
    if (!program) continue;
    const isAdminCol = program.toLowerCase() === 'admin';

    for (let r = 1; r < rows.length; r++) {
      const cell = rows[r][c];
      if (cell === null || cell === undefined) continue;
      const email = String(cell).trim().toLowerCase();
      if (!email) continue;

      let entry = map.get(email);
      if (!entry) {
        entry = { programs: new Set(), isAdmin: false };
        map.set(email, entry);
      }
      if (isAdminCol) entry.isAdmin = true;
      else entry.programs.add(program);
    }
  }
  return map;
}

function signOutUser() {
  try {
    if (window.netlifyIdentity) {
      window.netlifyIdentity.logout();
      return; // the 'logout' handler below reloads the page
    }
  } catch (e) { /* fall through */ }
  window.location.reload();
}

/**
 * Resolves with { username, allowedPrograms } once a signed-in Identity
 * user is also found in the credentials sheet — either instantly (an
 * existing Netlify Identity session) or after they sign in / sign up via
 * the Identity widget modal. username is the user's email. allowedPrograms
 * is null for an admin entry (full access) or a Set of program names.
 */
function runLoginGate(credMap) {
  return new Promise((resolve) => {
    const identity = window.netlifyIdentity;
    if (!identity) {
      showLoginFatal('Netlify Identity didn’t load. Check your internet connection, or that Identity is enabled for this site (see README.md), then refresh.');
      return;
    }

    const screen = document.getElementById('loginScreen');
    const signedOutCard = document.getElementById('loginCard');
    const noAccessCard = document.getElementById('noAccessCard');
    const errorEl = document.getElementById('loginError');
    const noAccessEmailEl = document.getElementById('noAccessEmail');

    function showSignedOut() {
      screen.classList.remove('hidden');
      signedOutCard.classList.remove('hidden');
      noAccessCard.classList.add('hidden');
    }
    function showNoAccess(email) {
      screen.classList.remove('hidden');
      signedOutCard.classList.add('hidden');
      noAccessCard.classList.remove('hidden');
      noAccessEmailEl.textContent = email;
    }
    function hideLoginScreen() {
      screen.classList.add('hidden');
    }

    function tryEnter(user) {
      const email = (user.email || '').trim().toLowerCase();
      const entry = credMap.get(email);
      if (!entry) {
        showNoAccess(user.email || '(no email)');
        return false;
      }
      hideLoginScreen();
      resolve({ username: user.email, allowedPrograms: entry.isAdmin ? null : entry.programs });
      return true;
    }

    identity.on('init', (user) => {
      if (user) { if (!tryEnter(user)) { /* no-access screen already shown */ } }
      else showSignedOut();
    });

    identity.on('login', (user) => {
      errorEl.style.display = 'none';
      tryEnter(user);
      identity.close();
    });

    identity.on('signup', () => {
      // If email confirmation is required, Netlify's own widget shows the
      // "check your email" notice; nothing else to do here until they
      // click the confirmation link, which fires 'login' on return.
    });

    identity.on('logout', () => { window.location.reload(); });

    identity.on('error', (err) => {
      errorEl.textContent = (err && err.message) ? err.message : 'Something went wrong with sign-in. Please try again.';
      errorEl.style.display = '';
    });

    document.getElementById('loginSignInBtn').addEventListener('click', () => identity.open('login'));
    document.getElementById('loginSignUpBtn').addEventListener('click', () => identity.open('signup'));
    document.getElementById('noAccessSignOutBtn').addEventListener('click', signOutUser);
    document.getElementById('noAccessRefreshBtn').addEventListener('click', () => window.location.reload());

    identity.init(); // fires 'init' above with the current user (or null)
  });
}

function showLoginFatal(message) {
  const screen = document.getElementById('loginScreen');
  const signedOutCard = document.getElementById('loginCard');
  const noAccessCard = document.getElementById('noAccessCard');
  const errorEl = document.getElementById('loginError');
  screen.classList.remove('hidden');
  signedOutCard.classList.remove('hidden');
  noAccessCard.classList.add('hidden');
  errorEl.textContent = message;
  errorEl.style.display = '';
}
