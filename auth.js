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

   Sign In, Create Account, Forgot Password, and "set a new password
   after clicking the reset email link" are all OUR OWN forms here, not
   the widget's built-in modal — the modal turned out to render as an
   invisible, unclickable overlay on top of this page for at least one
   user, which is a known class of issue with embedding that widget
   in a custom page. We call straight into netlifyIdentity's underlying
   GoTrue client (`identity.gotrue`) for every step, which keeps the
   resulting session stored exactly where the widget itself expects it
   (so currentUser() and future page loads keep working normally) while
   never depending on the widget actually drawing anything itself. A
   "Confirm Password" field on signup, and checking a brand-new signup
   against the credentials sheet BEFORE an account is created at all,
   come along for free this way too.

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
  if (/not.*confirm|confirm.*email/i.test(msg)) return 'This account needs to be confirmed before signing in. If you just created it, wait a moment and try again — if it keeps happening, ask your admin to check the "autoconfirm" setting in the Netlify dashboard (see README).';
  if (/invalid.*login|invalid.*grant|invalid.*password|wrong.*password/i.test(msg)) return 'Incorrect email or password.';
  if (/password.*short|password.*least/i.test(msg)) return 'Password is too short — use at least 6 characters.';
  return msg || 'Something went wrong. Please try again.';
}

/** Small delay helper for the signup->login retry below. */
function wait(ms) { return new Promise(r => setTimeout(r, ms)); }

/**
 * Logs in right after a signup. With Netlify's "autoconfirm" setting on,
 * the account is usable immediately — but occasionally the very first
 * login attempt right after signup fails transiently (the account hasn't
 * finished propagating on Netlify's side yet) even though the account is
 * genuinely fine, as shown by a plain Sign In working a moment later. A
 * couple of short retries absorb that without the person ever seeing it.
 */
async function loginWithRetry(identity, email, password, attempts) {
  let lastErr;
  for (let i = 0; i < attempts; i++) {
    try {
      return await identity.gotrue.login(email, password, true);
    } catch (err) {
      lastErr = err;
      if (i < attempts - 1) await wait(600 * (i + 1));
    }
  }
  throw lastErr;
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
    const forgotCard = document.getElementById('forgotCard');
    const resetCard = document.getElementById('resetCard');
    const noAccessCard = document.getElementById('noAccessCard');
    const loginErrorEl = document.getElementById('loginError');
    const noAccessEmailEl = document.getElementById('noAccessEmail');

    const allCards = [choiceCard, signinCard, signupCard, forgotCard, resetCard, noAccessCard];
    function showOnly(card) {
      screen.classList.remove('hidden');
      allCards.forEach(c => c.classList.toggle('hidden', c !== card));
    }
    function hideLoginScreen() { screen.classList.add('hidden'); }

    // A password-reset email link lands back here as
    // "#recovery_token=...&type=recovery". An inline script at the very
    // top of index.html — which runs before this file, and before the
    // Identity widget's own script — already pulled the token out of the
    // URL and scrubbed the hash, so the widget's own internal handling
    // (which has proven unreliable on this site) never gets a chance to
    // see or react to it in the first place.
    const recoveryToken = window.__pendingRecoveryToken || null;
    delete window.__pendingRecoveryToken;
    let pendingRecoveryToken = null;

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
      if (recoveryToken) return; // the reset-password flow below owns the screen until that form is submitted
      if (user) tryEnter(user);
      else showOnly(choiceCard);
    });

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

      const email = signinEmailInput.value.trim().toLowerCase();
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

    document.getElementById('signinForgotBtn').addEventListener('click', () => {
      forgotErrorEl.style.display = 'none';
      forgotInfoEl.style.display = 'none';
      forgotEmailInput.value = signinEmailInput.value.includes('@') ? signinEmailInput.value.trim().toLowerCase() : '';
      showOnly(forgotCard);
      forgotEmailInput.focus();
    });

    document.getElementById('signinBackBtn').addEventListener('click', () => showOnly(choiceCard));

    /* ---------- forgot password (its own form — just an email + submit) ---------- */
    const forgotForm = document.getElementById('forgotForm');
    const forgotEmailInput = document.getElementById('forgotEmail');
    const forgotErrorEl = document.getElementById('forgotError');
    const forgotInfoEl = document.getElementById('forgotInfo');
    const forgotSubmitBtn = document.getElementById('forgotSubmitBtn');

    forgotForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      forgotErrorEl.style.display = 'none';
      forgotInfoEl.style.display = 'none';

      const email = forgotEmailInput.value.trim().toLowerCase();
      if (!email) {
        forgotErrorEl.textContent = 'Enter your account email.';
        forgotErrorEl.style.display = '';
        return;
      }

      forgotSubmitBtn.disabled = true;
      forgotSubmitBtn.textContent = 'Sending…';
      try {
        await identity.gotrue.requestPasswordRecovery(email);
        forgotInfoEl.textContent = 'If that email has an account, a password reset link is on its way to it. Check your inbox and click the link to set a new password.';
        forgotInfoEl.style.display = '';
      } catch (err) {
        forgotErrorEl.textContent = friendlyAuthError(err);
        forgotErrorEl.style.display = '';
      } finally {
        forgotSubmitBtn.disabled = false;
        forgotSubmitBtn.textContent = 'Send reset link';
      }
    });

    document.getElementById('forgotBackBtn').addEventListener('click', () => showOnly(signinCard));

    /* ---------- set a new password (after clicking a reset-email link) ---------- */
    const resetForm = document.getElementById('resetForm');
    const resetPasswordInput = document.getElementById('resetPassword');
    const resetConfirmInput = document.getElementById('resetConfirmPassword');
    const resetErrorEl = document.getElementById('resetError');
    const resetSubmitBtn = document.getElementById('resetSubmitBtn');
    const resetRequestNewBtn = document.getElementById('resetRequestNewBtn');

    // IMPORTANT: the recovery token is a one-time-use credential. Some
    // organizations' email security (notably Microsoft 365 Safe Links)
    // automatically opens email links in a scanning browser *before* the
    // person ever clicks them, to check they're not malicious — and if
    // that scan runs this page's JS, anything that auto-exchanges the
    // token the instant the page loads gets its token burned by the
    // scanner instead of the real person. So this deliberately does NOT
    // call identity.gotrue.recover() on page load. It only happens inside
    // the form's submit handler below, which requires someone to actually
    // type a password and click submit — something an automated link
    // scanner essentially never does. This is the standard mitigation for
    // this class of problem (the same one Supabase/Auth0/etc. recommend).
    function initiateRecoveryFlow(token) {
      showOnly(resetCard);
      resetErrorEl.style.display = 'none';
      resetRequestNewBtn.classList.add('hidden');
      pendingRecoveryToken = token;
    }

    resetForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      resetErrorEl.style.display = 'none';
      if (!pendingRecoveryToken) return;

      const password = resetPasswordInput.value;
      const confirm = resetConfirmInput.value;
      if (password !== confirm) {
        resetErrorEl.textContent = 'Passwords don’t match.';
        resetErrorEl.style.display = '';
        resetConfirmInput.value = '';
        resetConfirmInput.focus();
        return;
      }
      if (password.length < 6) {
        resetErrorEl.textContent = 'Password must be at least 6 characters.';
        resetErrorEl.style.display = '';
        return;
      }

      resetSubmitBtn.disabled = true;
      resetSubmitBtn.textContent = 'Saving…';
      try {
        // identity.gotrue.recover() calls a legacy "/verify" endpoint that
        // this Netlify site's hosted Identity backend 404'd on earlier. The
        // real GoTrue source (netlify/gotrue api/verify.go) confirms /verify
        // IS the correct, required step — it just has to be called the way
        // the backend actually expects: POST, JSON body {type, token}. That
        // call exchanges the one-time recovery token for a real session
        // (access_token); the recovery token itself is never valid as a
        // Bearer credential on its own, which is why using it directly
        // against /user (the previous attempt) kept reporting it as
        // invalid/expired even on a brand-new link. The actual password
        // update is a second, separate call using that session's token.
        const verifyRes = await fetch('/.netlify/identity/verify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ type: 'recovery', token: pendingRecoveryToken }),
        });
        if (!verifyRes.ok) {
          const err = new Error('verify failed');
          err.status = verifyRes.status;
          throw err;
        }
        const session = await verifyRes.json(); // { access_token, refresh_token, expires_in, user, ... }

        const updateRes = await fetch('/.netlify/identity/user', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + session.access_token },
          body: JSON.stringify({ password }),
        });
        if (!updateRes.ok) {
          let detail = '';
          try { const body = await updateRes.json(); detail = body.msg || body.error_description || body.error || ''; } catch (e2) { /* ignore */ }
          throw new Error(detail || ('Request failed with status ' + updateRes.status));
        }
        const updatedUser = await updateRes.json();
        pendingRecoveryToken = null;

        // Also log in normally (same as the signup flow does right after
        // creating an account) so this is a real, persisted Identity
        // session — not just a one-off password change.
        let sessionUser = updatedUser;
        try {
          sessionUser = await loginWithRetry(identity, updatedUser.email, password, 2);
        } catch (loginErr) { /* password is genuinely set either way — proceed with what we have */ }
        tryEnter(sessionUser);
      } catch (err) {
        if (err.status === 404 || err.status === 422) {
          resetErrorEl.textContent = 'This password reset link has already been used or has expired — this can happen if your email security automatically opens links before you click them. Request a fresh one below.';
          resetRequestNewBtn.classList.remove('hidden');
        } else {
          resetErrorEl.textContent = friendlyAuthError(err);
        }
        resetErrorEl.style.display = '';
      } finally {
        resetSubmitBtn.disabled = false;
        resetSubmitBtn.textContent = 'Set New Password';
      }
    });

    document.getElementById('resetRequestNewBtn').addEventListener('click', () => showOnly(forgotCard));

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

      const email = signupEmailInput.value.trim().toLowerCase();
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
        try {
          const user = await loginWithRetry(identity, email, password, 3);
          tryEnter(user);
        } catch (loginErr) {
          // The account was created successfully (signup() didn't throw) —
          // only the immediate auto-login didn't go through. Don't show a
          // scary error for an account that's actually fine; just hand
          // them to Sign In with the email already filled in.
          showOnly(signinCard);
          signinEmailInput.value = email;
          signinInfoEl.textContent = 'Account created. Sign in below to continue.';
          signinInfoEl.style.display = '';
          signinPasswordInput.focus();
        }
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

    if (recoveryToken) initiateRecoveryFlow(recoveryToken);
    identity.init(); // fires 'init' above with the current user (or null) — a no-op while a recovery flow is in progress
  });
}

function showLoginFatal(message) {
  const screen = document.getElementById('loginScreen');
  const choiceCard = document.getElementById('loginCard');
  screen.classList.remove('hidden');
  ['loginCard', 'signinCard', 'signupCard', 'forgotCard', 'noAccessCard'].forEach(id => {
    document.getElementById(id).classList.toggle('hidden', id !== 'loginCard');
  });
  const errorEl = document.getElementById('loginError');
  errorEl.textContent = message;
  errorEl.style.display = '';
}
