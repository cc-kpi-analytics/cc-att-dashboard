/* ============================================================
   Login gate — Netlify Identity for real authentication (hosted
   sign-up, sign-in, and self-serve "forgot password" — Netlify
   stores and hashes the passwords, nothing password-related ever
   touches attendance_raw.xlsx), with program access still driven
   by the "credentials" sheet in attendance_raw.xlsx.
   ============================================================
   What Netlify Identity gives us: an account system with no backend
   code to write. Sign up, sign in, and password reset are handled by
   Netlify's hosted Identity (GoTrue) service through the
   netlify-identity-widget script tag in index.html — see README.md
   for the one-time Netlify dashboard setup (including turning off the
   signup confirmation email, since this app sends email only for
   password resets).

   Sign In and Create Account are our OWN forms here (not the widget's
   built-in modal) for two reasons: a "Confirm Password" field on
   signup, and — more importantly — so a brand-new signup can be
   checked against the credentials sheet BEFORE an account is created
   at all, which the widget's own signup tab has no way to do. We call
   straight into netlifyIdentity's underlying GoTrue client
   (`identity.gotrue`) to do this, which keeps the resulting session
   stored exactly where the widget itself expects it (so currentUser(),
   future page loads, and the widget's own automatic handling of a
   password-recovery email link all keep working normally).

   What none of this knows about is which program an email belongs to.
   That mapping still lives in the "credentials" sheet of
   attendance_raw.xlsx (one column per program, listing the emails
   that belong to it; an "Admin" column grants full access). Signing up
   is blocked client-side for an email that isn't in that sheet — this
   is the same UI-convenience tradeoff as the rest of this app accepts
   elsewhere: the sheet is fetched over a public, unauthenticated URL,
   so a technical person could still reach the signup API directly and
   create an unlisted account, but that account still couldn't see any
   data, since dashboard access checks the same sheet separately on
   every sign-in. See README.md for the hardened alternative (a small
   Netlify Function that enforces this server-side) if that residual
   gap matters for your use case. */

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

function friendlyAuthError(err) {
  const msg = (err && err.message) ? err.message : '';
  if (/already.*registered|user.*exists|already.*exist/i.test(msg)) return 'An account with that email already exists. Try Sign In instead.';
  if (/invalid.*login|invalid.*grant|invalid.*password|wrong.*password/i.test(msg)) return 'Incorrect email or password.';
  if (/password.*short|password.*least/i.test(msg)) return 'Password is too short — use at least 6 characters.';
  return msg || 'Something went wrong. Please try again.';
}

/**
 * Resolves with { username, allowedPrograms } once a signed-in Identity
 * user is also found in the credentials sheet — either instantly (an
 * existing Netlify Identity session), after they sign in through our
 * own form, after they create an account through our own form, or after
 * the widget's own password-recovery flow logs them in. username is the
 * user's email. allowedPrograms is null for an admin entry (full access)
 * or a Set of program names.
 */
function runLoginGate(credMap) {
  return new Promise((resolve) => {
    const identity = window.netlifyIdentity;
    if (!identity) {
      showLoginFatal('Netlify Identity didn’t load. Check your internet connection, or that Identity is enabled for this site (see README.md), then refresh.');
      return;
    }

    const screen = document.getElementById('loginScreen');
    const choiceCard = document.getElementById('loginCard');
    const signinCard = document.getElementById('signinCard');
    const signupCard = document.getElementById('signupCard');
    const noAccessCard = document.getElementById('noAccessCard');
    const loginErrorEl = document.getElementById('loginError');
    const noAccessEmailEl = document.getElementById('noAccessEmail');

    const allCards = [choiceCard, signinCard, signupCard, noAccessCard];
    function showOnly(card) {
      screen.classList.remove('hidden');
      allCards.forEach(c => c.classList.toggle('hidden', c !== card));
    }
    function hideLoginScreen() { screen.classList.add('hidden'); }

    function tryEnter(user) {
      const email = (user.email || '').trim().toLowerCase();
      const entry = credMap.get(email);
      if (!entry) {
        noAccessEmailEl.textContent = user.email || '(no email)';
        showOnly(noAccessCard);
        return false;
      }
      hideLoginScreen();
      resolve({ username: user.email, allowedPrograms: entry.isAdmin ? null : entry.programs });
      return true;
    }

    identity.on('init', (user) => {
      if (user) tryEnter(user);
      else showOnly(choiceCard);
    });

    // Fires when the widget's OWN internal flows complete a login behind
    // our backs — namely, after someone sets a new password from a
    // password-recovery email link (that UI is still the widget's own
    // built-in modal; see README). Signing in or signing up through our
    // own forms below call tryEnter() directly instead of relying on this.
    identity.on('login', (user) => { tryEnter(user); identity.close(); });
    identity.on('logout', () => { window.location.reload(); });
    identity.on('error', (err) => {
      loginErrorEl.textContent = friendlyAuthError(err);
      loginErrorEl.style.display = '';
    });

    /* ---------- choice card ---------- */
    document.getElementById('loginSignInBtn').addEventListener('click', () => {
      document.getElementById('signinError').style.display = 'none';
      document.getElementById('signinInfo').style.display = 'none';
      showOnly(signinCard);
      document.getElementById('signinEmail').focus();
    });
    document.getElementById('loginSignUpBtn').addEventListener('click', () => {
      document.getElementById('signupError').style.display = 'none';
      showOnly(signupCard);
      document.getElementById('signupEmail').focus();
    });

    /* ---------- sign in (custom form, straight to GoTrue) ---------- */
    const signinForm = document.getElementById('signinForm');
    const signinEmailInput = document.getElementById('signinEmail');
    const signinPasswordInput = document.getElementById('signinPassword');
    const signinErrorEl = document.getElementById('signinError');
    const signinInfoEl = document.getElementById('signinInfo');
    const signinSubmitBtn = document.getElementById('signinSubmitBtn');

    signinForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      signinErrorEl.style.display = 'none';
      signinInfoEl.style.display = 'none';

      const email = signinEmailInput.value.trim();
      const password = signinPasswordInput.value;
      if (!email || !password) {
        signinErrorEl.textContent = 'Enter your email and password.';
        signinErrorEl.style.display = '';
        return;
      }

      signinSubmitBtn.disabled = true;
      signinSubmitBtn.textContent = 'Signing in…';
      try {
        const user = await identity.gotrue.login(email, password, true);
        tryEnter(user);
      } catch (err) {
        signinErrorEl.textContent = friendlyAuthError(err);
        signinErrorEl.style.display = '';
      } finally {
        signinSubmitBtn.disabled = false;
        signinSubmitBtn.textContent = 'Sign In';
      }
    });

    document.getElementById('signinForgotBtn').addEventListener('click', async () => {
      signinErrorEl.style.display = 'none';
      const email = signinEmailInput.value.trim();
      if (!email) {
        signinErrorEl.textContent = 'Enter your email above first, then click "Forgot password?" again.';
        signinErrorEl.style.display = '';
        signinEmailInput.focus();
        return;
      }
      try {
        await identity.gotrue.requestPasswordRecovery(email);
        signinInfoEl.textContent = 'If that email has an account, a password reset link is on its way to it.';
        signinInfoEl.style.display = '';
      } catch (err) {
        signinErrorEl.textContent = friendlyAuthError(err);
        signinErrorEl.style.display = '';
      }
    });

    document.getElementById('signinBackBtn').addEventListener('click', () => showOnly(choiceCard));

    /* ---------- create account (custom form, gated by credentials sheet) ---------- */
    const signupForm = document.getElementById('signupForm');
    const signupEmailInput = document.getElementById('signupEmail');
    const signupPasswordInput = document.getElementById('signupPassword');
    const signupConfirmInput = document.getElementById('signupConfirmPassword');
    const signupErrorEl = document.getElementById('signupError');
    const signupSubmitBtn = document.getElementById('signupSubmitBtn');

    signupForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      signupErrorEl.style.display = 'none';

      const email = signupEmailInput.value.trim();
      const password = signupPasswordInput.value;
      const confirm = signupConfirmInput.value;

      if (!email || !password) {
        signupErrorEl.textContent = 'Enter an email and password.';
        signupErrorEl.style.display = '';
        return;
      }
      if (password !== confirm) {
        signupErrorEl.textContent = 'Passwords don’t match.';
        signupErrorEl.style.display = '';
        signupConfirmInput.value = '';
        signupConfirmInput.focus();
        return;
      }
      if (password.length < 6) {
        signupErrorEl.textContent = 'Password must be at least 6 characters.';
        signupErrorEl.style.display = '';
        return;
      }
      if (!credMap.has(email.toLowerCase())) {
        signupErrorEl.textContent = 'This email isn’t on the dashboard’s access list. Ask your admin to add it to the "credentials" sheet in attendance_raw.xlsx, then try again.';
        signupErrorEl.style.display = '';
        return;
      }

      signupSubmitBtn.disabled = true;
      signupSubmitBtn.textContent = 'Creating account…';
      try {
        await identity.gotrue.signup(email, password);
        const user = await identity.gotrue.login(email, password, true);
        tryEnter(user);
      } catch (err) {
        signupErrorEl.textContent = friendlyAuthError(err);
        signupErrorEl.style.display = '';
      } finally {
        signupSubmitBtn.disabled = false;
        signupSubmitBtn.textContent = 'Create Account';
      }
    });

    document.getElementById('signupBackBtn').addEventListener('click', () => showOnly(choiceCard));

    document.getElementById('noAccessSignOutBtn').addEventListener('click', signOutUser);
    document.getElementById('noAccessRefreshBtn').addEventListener('click', () => window.location.reload());

    identity.init(); // fires 'init' above with the current user (or null); also auto-handles a #recovery_token= link
  });
}

function showLoginFatal(message) {
  const screen = document.getElementById('loginScreen');
  const choiceCard = document.getElementById('loginCard');
  screen.classList.remove('hidden');
  ['loginCard', 'signinCard', 'signupCard', 'noAccessCard'].forEach(id => {
    document.getElementById(id).classList.toggle('hidden', id !== 'loginCard');
  });
  const errorEl = document.getElementById('loginError');
  errorEl.textContent = message;
  errorEl.style.display = '';
}
