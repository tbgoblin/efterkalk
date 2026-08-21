// ── Auth / brugere ──────────────────────────────────────────────────────────
// Estratto verbatim da routes/apiRoutes.js: gestione utenti (users.json),
// hash password, sessioni bearer-token e guard middleware. Factory con fs
// iniettato per mantenere identici i call site del router.
const path = require('path');
const crypto = require('crypto');

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

    function getSessionUser(req) {
        const token = String(req.headers.authorization || '').replace(/^Bearer\s+/i, '');
        const session = authSessions.get(token);
        return session && session.expiresAt > Date.now() ? session.user : null;
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
        safeUser,
        makePasswordHash,
        getSessionUser,
        requireSuperadmin,
        requireModulePermission
    };
}

module.exports = { createAuthService };
