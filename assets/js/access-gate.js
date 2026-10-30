// ── Access gate (login) · client ──────────────────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: applyModulePermissions, initializeAfterAccess, setLoggedUserDisplayName
// (dichiarati nell'inline script) e refreshSideMenuAuthState (dichiarato in
// assets/js/side-menu.js). test/authService.test.js valuta la parte da
// "let _accessLoginInProgress" in poi direttamente da questo file.
function showAccessGate() {
    const overlay = document.getElementById('accessGateOverlay');
    const userInput = document.getElementById('accessGateUserInput');
    const input = document.getElementById('accessGateInput');
    const rememberUser = document.getElementById('accessGateRememberUser');
    const err = document.getElementById('accessGateError');
    if (!overlay) return;
    if (err) err.textContent = '';
    if (userInput) {
        let rememberedUsername = '';
        try { rememberedUsername = localStorage.getItem('afterkalk_remembered_username') || ''; } catch {}
        if (!String(userInput.value || '').trim() && rememberedUsername) {
            userInput.value = rememberedUsername;
        }
        if (rememberUser) rememberUser.checked = Boolean(rememberedUsername);
    }
    overlay.style.display = 'flex';
    refreshSideMenuAuthState();
    setTimeout(() => {
        if (userInput && !String(userInput.value || '').trim()) {
            userInput.focus();
            return;
        }
        if (input) input.focus();
    }, 30);
}

function hideAccessGate() {
    const overlay = document.getElementById('accessGateOverlay');
    if (!overlay) return;
    overlay.style.display = 'none';
    refreshSideMenuAuthState();
}

let _accessLoginInProgress = false;
let accessSessionRevision = 0;

async function forgetRememberedLogin() {
    const storageStatus = document.getElementById('accessGateLoginStorageStatus');
    try {
        if (window.GohDesktopLogin) await window.GohDesktopLogin.forget();
        document.getElementById('accessGateRememberPassword').checked = false;
        document.getElementById('accessGateForgetLogin').style.display = 'none';
        if (storageStatus) storageStatus.textContent = '';
    } catch (_) {
        if (storageStatus) storageStatus.textContent = 'Gemt login kunne ikke slettes. Prøv igen.';
    }
}

async function configureRememberedLogin() {
    if (!window.GohDesktopLogin) return null;
    try {
        const info = await window.GohDesktopLogin.info();
        document.getElementById('accessGateRememberPasswordLabel').style.display = info.available ? 'flex' : 'none';
        document.getElementById('accessGateRememberPassword').checked = info.saved;
        document.getElementById('accessGateForgetLogin').style.display = info.saved || info.damaged ? '' : 'none';
        if (info.saved) {
            document.getElementById('accessGateUserInput').value = info.username;
            document.getElementById('accessGateRememberUser').checked = true;
        }
        return info;
    } catch (_) { return null; }
}

function applyAuthenticatedSession(data) {
    if (!data.token || !data.user || !data.user.username) throw new Error('Ugyldig session');
    authToken = data.token;
    loggedUserRole = data.user.role || 'user';
    loggedUserPermissions = data.user.permissions || {};
    loggedUsername = String(data.user.username);
    setLoggedUserDisplayName(data.user.displayName || loggedUsername);
    for (const id of ['accessGateUserInput', 'sideMenuUserInput']) {
        const field = document.getElementById(id);
        if (field) field.value = loggedUsername;
    }
    accessGranted = true;
    hideAccessGate();
    refreshSideMenuAuthState();
    const adminCard = document.getElementById('administrationDashCard');
    if (adminCard) adminCard.style.display = loggedUserRole === 'superadmin' ? '' : 'none';
    applyModulePermissions();
    initializeAfterAccess();
}

let accessRestoreTask = null;
function restoreAccessSession() {
    if (accessGranted || accessRestoreTask) return accessRestoreTask;
    const revision = accessSessionRevision;
    accessRestoreTask = (async () => {
        _accessLoginInProgress = true;
        const button = document.getElementById('accessGateBtn');
        if (button) button.disabled = true;
        try {
            const saved = await configureRememberedLogin();
            const response = await fetch('/auth/session', { credentials: 'same-origin', cache: 'no-store', signal: AbortSignal.timeout(8000) });
            if (response.ok) {
                const data = await response.json();
                if (revision === accessSessionRevision) applyAuthenticatedSession(data);
            } else if (revision === accessSessionRevision && response.status === 401 && saved?.saved) {
                const restored = await window.GohDesktopLogin.restore();
                if (restored && revision === accessSessionRevision) applyAuthenticatedSession(restored);
            }
        } catch (_) {
            document.getElementById('accessGateLoginStorageStatus').textContent = 'Log ind for at fortsætte.';
        } finally {
            if (revision === accessSessionRevision) {
                _accessLoginInProgress = false;
                if (button) button.disabled = false;
                if (!accessGranted) showAccessGate();
            }
        }
    })();
    return accessRestoreTask;
}

async function submitAccessCode() {
    if (_accessLoginInProgress) return;
    const revision = accessSessionRevision;
    const userInput = document.getElementById('accessGateUserInput');
    const input = document.getElementById('accessGateInput');
    const err = document.getElementById('accessGateError');
    const btn = document.getElementById('accessGateBtn');
    const rememberUser = document.getElementById('accessGateRememberUser');
    const userName = String(userInput ? userInput.value : '').trim() || 'admin';
    const loginUserName = userName.toLowerCase() === 'superbruger' ? 'admin' : userName;
    const value = input ? String(input.value || '') : '';

    if (err) err.textContent = 'Åbner...';
    if (btn) {
        btn.disabled = true;
        btn.textContent = 'Åbner...';
    }
    _accessLoginInProgress = true;

    try {
        const response = await fetch('/auth/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ username: loginUserName, password: value })
        });
        const data = await response.json();
        if (!response.ok || data.error) throw new Error(data.error || 'Forkert brugernavn eller kode');
        const canonicalUsername = String(data.user && data.user.username || loginUserName);
        const rememberPassword = document.getElementById('accessGateRememberPassword');
        if (rememberPassword?.checked && rememberUser) rememberUser.checked = true;
        try {
            if (rememberUser && rememberUser.checked) localStorage.setItem('afterkalk_remembered_username', canonicalUsername);
            else localStorage.removeItem('afterkalk_remembered_username');
        } catch {}
        if (window.GohDesktopLogin) {
            try {
                if (rememberPassword?.checked) await window.GohDesktopLogin.remember({ username: canonicalUsername, password: value, token: data.token });
                else await window.GohDesktopLogin.forget();
            } catch (_) {
                if (rememberPassword) rememberPassword.checked = false;
                alert('Login er godkendt, men det gemte login kunne ikke opdateres. Brug Glem gemt login eller Log ud for at prøve igen.');
            }
        }
        if (input) input.value = '';
        if (revision === accessSessionRevision) applyAuthenticatedSession(data);
    } catch (e) {
        accessGranted = false;
        if (err) err.textContent = e && e.message ? e.message : 'Forkert brugernavn eller kode';
        if (input) input.focus();
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.textContent = 'Åbn dashboard';
        }
        const adminCard = document.getElementById('administrationDashCard');
        if (adminCard) adminCard.style.display = loggedUserRole === 'superadmin' ? '' : 'none';
        applyModulePermissions();
        _accessLoginInProgress = false;
    }
}
