// ── Auth / brugere ──────────────────────────────────────────────────────────
// Estratto verbatim da routes/apiRoutes.js: gestione utenti (users.json),
// hash password, sessioni bearer-token e guard middleware. Factory con fs
// iniettato per mantenere identici i call site del router.
const path = require('path');
const crypto = require('crypto');
const gohData = require('./gohDataService');
const USERS_STATE_KEY = 'app_users';
const SESSION_COOKIE_NAME = 'gantech_session';
const SESSION_TTL_MS = 8 * 60 * 60 * 1000;

function createAuthService({ fs, usersFile }) {
    const legacyUsersFile = usersFile || path.join(__dirname, '..', 'users.json');
    const dataDir = String(process.env.GANTECH_DATA_DIR || '').trim();
    const resolvedUsersFile = dataDir ? path.join(dataDir, 'users.json') : legacyUsersFile;
    const authSessions = new Map();

    function ensureUsersFile() {
        if (fs.existsSync(resolvedUsersFile)) return;
        if (resolvedUsersFile !== legacyUsersFile && fs.existsSync(legacyUsersFile)) {
            fs.mkdirSync(path.dirname(resolvedUsersFile), { recursive: true });
            fs.copyFileSync(legacyUsersFile, resolvedUsersFile);
            return;
        }
        fs.mkdirSync(path.dirname(resolvedUsersFile), { recursive: true });
        fs.writeFileSync(resolvedUsersFile, '[]\n', 'utf8');
    }

    function readUsers() {
        ensureUsersFile();
        const parsed = JSON.parse(fs.readFileSync(resolvedUsersFile, 'utf8'));
        return Array.isArray(parsed) ? parsed : [];
    }

    function writeUsers(users) {
        ensureUsersFile();
        fs.writeFileSync(resolvedUsersFile, JSON.stringify(users, null, 2) + '\n', 'utf8');
        gohData.setAppState(USERS_STATE_KEY, users).catch(() => {});
    }

    // All'avvio gli utenti condivisi su GOH vincono sul file locale (fail-soft).
    async function hydrateUsersFromDb() {
        const state = await gohData.getAppState(USERS_STATE_KEY);
        if (!state || !Array.isArray(state.payload)) {
            // Primo avvio con DB vuoto: pubblica gli utenti locali come base condivisa
            const localUsers = readUsers();
            if (localUsers.length > 0) gohData.setAppState(USERS_STATE_KEY, localUsers).catch(() => {});
            return false;
        }
        ensureUsersFile();
        fs.writeFileSync(resolvedUsersFile, JSON.stringify(state.payload, null, 2) + '\n', 'utf8');
        return true;
    }

    function safeUser(user) {
        return {
            username: user.username,
            displayName: user.displayName || user.username,
            role: user.role || (user.isSuperUser ? 'superadmin' : 'user'),
            active: user.active !== false,
            permissions: user.permissions || {}
        };
    }

    function makePasswordHash(password, salt) {
        return crypto.pbkdf2Sync(password, salt, 120000, 64, 'sha512').toString('hex');
    }

    function getBearerToken(req) {
        const authorization = String((req.headers && req.headers.authorization) || '');
        const match = authorization.match(/^Bearer\s+(.+)$/i);
        return match ? match[1].trim() : '';
    }

    function getCookieToken(req) {
        const cookieHeader = String((req.headers && req.headers.cookie) || '');
        for (const item of cookieHeader.split(';')) {
            const separator = item.indexOf('=');
            if (separator === -1) continue;
            const name = item.slice(0, separator).trim();
            if (name !== SESSION_COOKIE_NAME) continue;
            const value = item.slice(separator + 1).trim();
            try {
                return decodeURIComponent(value);
            } catch (_) {
                return '';
            }
        }
        return '';
    }

    function getSessionToken(req) {
        return getBearerToken(req) || getCookieToken(req);
    }

    function getSessionUser(req) {
        const token = getSessionToken(req);
        if (!token) return null;
        const session = authSessions.get(token);
        if (!session) return null;
        if (session.expiresAt <= Date.now()) {
            authSessions.delete(token);
            return null;
        }
        return session.user;
    }

    function requireAuthenticated(req, res, next) {
        if (!getSessionUser(req)) {
            return res.status(401).json({ error: 'Login kræves' });
        }
        return next();
    }

    function revokeSession(req) {
        const tokens = new Set([getBearerToken(req), getCookieToken(req)].filter(Boolean));
        for (const token of tokens) authSessions.delete(token);
    }

    function buildSessionCookie(token) {
        return SESSION_COOKIE_NAME + '=' + encodeURIComponent(token)
            + '; HttpOnly; SameSite=Strict; Path=/; Max-Age=' + Math.floor(SESSION_TTL_MS / 1000);
    }

    function buildExpiredSessionCookie() {
        return SESSION_COOKIE_NAME + '=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT';
    }

    function requireSuperadmin(req, res) {
        const user = getSessionUser(req);
        if (!user || user.role !== 'superadmin') {
            res.status(403).json({ error: 'Superadmin adgang kræves' });
            return null;
        }
        return user;
    }

    function requireModulePermission(permission) {
        return (req, res, next) => {
            const user = getSessionUser(req);
            if (!user || (user.role !== 'superadmin' && !(user.permissions && user.permissions[permission]))) {
                return res.status(403).json({ error: 'Adgang til modulet er ikke tilladt' });
            }
            return next();
        };
    }

    return {
        authSessions,
        readUsers,
        writeUsers,
        hydrateUsersFromDb,
        safeUser,
        makePasswordHash,
        getSessionUser,
        requireAuthenticated,
        requireSuperadmin,
        requireModulePermission,
        revokeSession,
        buildSessionCookie,
        buildExpiredSessionCookie,
        sessionTtlMs: SESSION_TTL_MS
    };
}

module.exports = { createAuthService };
