const test = require('node:test');
const assert = require('node:assert/strict');
const { createAuthService } = require('../services/authService');

function createService(cookieScope = '') {
    return createAuthService({ fs: {}, usersFile: 'unused-in-these-tests.json', cookieScope });
}

function responseRecorder() {
    return {
        statusCode: 200,
        body: null,
        status(code) {
            this.statusCode = code;
            return this;
        },
        json(body) {
            this.body = body;
            return this;
        }
    };
}

test('bearer authentication remains supported', () => {
    const service = createService();
    const user = { username: 'operator', role: 'user' };
    service.authSessions.set('bearer-token', { user, expiresAt: Date.now() + 60_000 });

    const actual = service.getSessionUser({ headers: { authorization: 'Bearer bearer-token' } });
    assert.equal(actual, user);
});

test('HttpOnly same-origin cookie authenticates separate application pages', () => {
    const service = createService();
    const user = { username: 'operator', role: 'user' };
    service.authSessions.set('cookie-token', { user, expiresAt: Date.now() + 60_000 });

    const actual = service.getSessionUser({
        headers: { cookie: 'theme=dark; gantech_session=cookie-token; other=value' }
    });
    assert.equal(actual, user);

    const cookie = service.buildSessionCookie('cookie-token');
    assert.match(cookie, /^gantech_session=cookie-token;/);
    assert.match(cookie, /HttpOnly/);
    assert.match(cookie, /SameSite=Strict/);
    assert.match(cookie, /Path=\//);
    assert.match(cookie, /Max-Age=28800/);
});

test('expired sessions are rejected and removed', () => {
    const service = createService();
    service.authSessions.set('expired', {
        user: { username: 'old' },
        expiresAt: Date.now() - 1
    });

    assert.equal(service.getSessionUser({ headers: { authorization: 'Bearer expired' } }), null);
    assert.equal(service.authSessions.has('expired'), false);
});

test('authentication middleware rejects anonymous writes with 401', () => {
    const service = createService();
    const response = responseRecorder();
    let nextCalled = false;

    service.requireAuthenticated({ headers: {} }, response, () => { nextCalled = true; });

    assert.equal(nextCalled, false);
    assert.equal(response.statusCode, 401);
    assert.deepEqual(response.body, { error: 'Login kræves' });
});

test('any-module middleware allows a user with one requested BOM permission', () => {
    const service = createService();
    service.authSessions.set('allowed', {
        user: { username: 'operator', role: 'user', permissions: { bomCalculator: true } },
        expiresAt: Date.now() + 60_000
    });
    let nextCalled = false;

    service.requireAnyModulePermission(['bomMaterials', 'bomCalculator'])(
        { headers: { authorization: 'Bearer allowed' } }, responseRecorder(), () => { nextCalled = true; }
    );

    assert.equal(nextCalled, true);
});

test('any-module middleware rejects a user without the requested BOM permission', () => {
    const service = createService();
    service.authSessions.set('denied', {
        user: { username: 'operator', role: 'user', permissions: { bomMaterials: true } },
        expiresAt: Date.now() + 60_000
    });
    const response = responseRecorder();

    service.requireAnyModulePermission('bomVismaPreview')(
        { headers: { authorization: 'Bearer denied' } }, response, () => assert.fail('next must not be called')
    );

    assert.equal(response.statusCode, 403);
    assert.deepEqual(response.body, { error: 'Adgang til BOM-området er ikke tilladt' });
});

test('any-module middleware always allows superadmin', () => {
    const service = createService();
    service.authSessions.set('superadmin', {
        user: { username: 'admin', role: 'superadmin', permissions: {} },
        expiresAt: Date.now() + 60_000
    });
    let nextCalled = false;

    service.requireAnyModulePermission('bomVismaPreview')(
        { headers: { authorization: 'Bearer superadmin' } }, responseRecorder(), () => { nextCalled = true; }
    );

    assert.equal(nextCalled, true);
});

test('logout revokes bearer and cookie sessions and emits an expired cookie', () => {
    const service = createService();
    const session = { user: { username: 'operator' }, expiresAt: Date.now() + 60_000 };
    service.authSessions.set('bearer-token', session);
    service.authSessions.set('cookie-token', session);

    service.revokeSession({
        headers: {
            authorization: 'Bearer bearer-token',
            cookie: 'gantech_session=cookie-token'
        }
    });

    assert.equal(service.authSessions.size, 0);
    assert.match(service.buildExpiredSessionCookie(), /Max-Age=0/);
});

test('session cookies are isolated between localhost ports including logout', () => {
    const first = createService('3000');
    const second = createService('3001');
    const cookie = first.buildSessionCookie('shared-token').split(';')[0];
    for (const service of [first, second]) service.authSessions.set('shared-token', {
        user: { username: 'operator' }, expiresAt: Date.now() + 60000
    });
    assert.equal(first.getSessionToken({ headers: { cookie } }), 'shared-token');
    assert.equal(first.getSessionUser({ headers: { cookie } }).username, 'operator');
    assert.equal(second.getSessionUser({ headers: { cookie } }), null);
    assert.equal(first.buildExpiredSessionCookie().split('=')[0], cookie.split('=')[0]);
    first.revokeSession({ headers: { cookie } });
    assert.equal(first.getSessionUser({ headers: { cookie } }), null);
    assert.equal(second.authSessions.size, 1);
});

test('session endpoint restores the safe user and token, rejects expired and revoked sessions without caching', () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const source = fs.readFileSync(path.join(__dirname, '../routes/apiRoutes.js'), 'utf8');
    const start = source.indexOf("    router.get('/auth/session',");
    const end = source.indexOf("    router.post('/auth/logout',", start);
    assert.ok(start > 0 && end > start);
    const service = createService();
    let handlers;
    new Function('router', 'requireAuthenticated', 'getSessionToken', 'getSessionUser', source.slice(start, end))(
        { get(route, ...callbacks) { handlers = callbacks; } }, service.requireAuthenticated, service.getSessionToken, service.getSessionUser);
    service.authSessions.set('restored', { user: { username: 'operator', role: 'user', permissions: {} }, expiresAt: Date.now() + 60000 });
    const req = { headers: { cookie: 'gantech_session=restored' } };
    const request = () => {
        const res = responseRecorder();
        res.headers = {};
        res.setHeader = (key, value) => { res.headers[key] = value; };
        handlers[0](req, res, () => handlers[1](req, res));
        return res;
    };
    const restored = request();
    assert.equal(restored.body.token, 'restored');
    assert.equal(restored.body.user.username, 'operator');
    assert.equal(restored.headers['Cache-Control'], 'no-store');
    service.authSessions.get('restored').expiresAt = 0;
    assert.equal(request().statusCode, 401);
    assert.equal(request().headers['Cache-Control'], 'no-store');
    assert.equal(service.authSessions.size, 0);
});

test('desktop credentials use encrypted bytes, stay in their own store and can be forgotten', () => {
    const { createDesktopCredentialStore } = require('../services/desktopCredentialService');
    const files = new Map();
    const plaintext = new Map();
    let available = true;
    const options = {
        fs: { existsSync: key => files.has(key), readFileSync: key => files.get(key),
            writeFileSync: (key, value) => files.set(key, value), unlinkSync: key => files.delete(key) },
        safeStorage: { isEncryptionAvailable: () => available,
            encryptString: value => { const bytes = Buffer.from('encrypted-' + plaintext.size); plaintext.set(bytes, value); return bytes; },
            decryptString: bytes => { if (!plaintext.has(bytes)) throw new Error('Corrupt'); return plaintext.get(bytes); } },
        enabled: true
    };
    const first = createDesktopCredentialStore({ ...options, filePath: 'userA-clientA.bin' });
    const second = createDesktopCredentialStore({ ...options, filePath: 'userA-clientB.bin' });
    const third = createDesktopCredentialStore({ ...options, filePath: 'userB-clientA.bin' });
    assert.deepEqual(first.info(), { available: true, saved: false, username: '' });
    first.save({ username: ' operator ', password: 'synthetic test value' });
    assert.ok(Buffer.isBuffer(files.get('userA-clientA.bin')));
    assert.equal(files.get('userA-clientA.bin').includes('synthetic'), false);
    assert.deepEqual(first.info(), { available: true, saved: true, username: 'operator' });
    assert.deepEqual(first.read(), { username: 'operator', password: 'synthetic test value' });
    assert.equal(second.read(), null);
    assert.equal(third.read(), null);
    available = false;
    assert.deepEqual(first.info(), { available: false, saved: false });
    assert.equal(first.read(), null);
    assert.throws(() => first.save({ username: 'operator', password: 'synthetic' }));
    available = true;
    files.set('userA-clientA.bin', Buffer.from('damaged'));
    assert.equal(first.info().damaged, true);
    assert.throws(() => first.read());
    first.forget();
    first.forget();
    assert.equal(files.size, 0);
    const disabled = createDesktopCredentialStore({ ...options, filePath: 'rds-without-client.bin', enabled: false });
    assert.equal(disabled.info().available, false);
    assert.throws(() => disabled.save({ username: 'operator', password: 'synthetic' }));
});

function accessClientFixture(options = {}) {
    const fs = require('node:fs');
    const path = require('node:path');
    const source = fs.readFileSync(path.join(__dirname, '../server.js'), 'utf8');
    const start = source.indexOf('            let _accessLoginInProgress = false;');
    const end = source.indexOf('            const ADMIN_MODULES =', start);
    const fields = new Map();
    const storage = new Map();
    let shown = 0;
    let initialized = 0;
    const document = { getElementById(id) {
        if (!fields.has(id)) fields.set(id, { value: '', checked: false, textContent: '', style: {}, focus() {} });
        return fields.get(id);
    } };
    const client = new Function('document', 'window', 'fetch', 'localStorage', 'showAccessGate', 'initializeAfterAccess', `
        let authToken = null, loggedUserRole = 'user', loggedUserPermissions = {}, loggedUsername = '', accessGranted = false;
        const setLoggedUserDisplayName = () => {}, hideAccessGate = () => {}, refreshSideMenuAuthState = () => {}, applyModulePermissions = () => {}, alert = () => {};
        ${source.slice(start, end)}
        return { restoreAccessSession, submitAccessCode, forgetRememberedLogin,
            cancel() { accessSessionRevision += 1; accessGranted = false; },
            state: () => ({ authToken, loggedUsername, accessGranted }) };
    `)(document, options.desktop ? { GohDesktopLogin: options.desktop } : {}, options.fetch,
        { setItem: (key, value) => storage.set(key, value), removeItem: key => storage.delete(key) },
        () => { shown += 1; }, () => { initialized += 1; });
    return { ...client, document, storage, counters: () => ({ shown, initialized }) };
}

test('main page restores an existing cookie session exactly once without password storage', async () => {
    let requests = 0;
    const client = accessClientFixture({ fetch: async url => {
        assert.equal(url, '/auth/session');
        requests += 1;
        return { ok: true, json: async () => ({ token: 'restored', user: { username: 'operator' } }) };
    } });
    assert.equal(client.restoreAccessSession(), client.restoreAccessSession());
    await client.restoreAccessSession();
    assert.equal(requests, 1);
    assert.deepEqual(client.counters(), { shown: 0, initialized: 1 });
    assert.equal(client.state().loggedUsername, 'operator');
    assert.equal(client.storage.size, 0);
});

test('expired cookie restores only an opted-in desktop login and cancellation rejects late results', async () => {
    let restores = 0;
    const desktop = { info: async () => ({ available: true, saved: true, username: 'operator' }),
        restore: async () => { restores += 1; return { token: 'native-token', user: { username: 'operator' } }; } };
    const expired = accessClientFixture({ desktop, fetch: async () => ({ ok: false, status: 401 }) });
    await expired.restoreAccessSession();
    assert.equal(restores, 1);
    assert.equal(expired.state().accessGranted, true);
    const anonymous = accessClientFixture({ fetch: async () => ({ ok: false, status: 401 }) });
    await anonymous.restoreAccessSession();
    assert.deepEqual(anonymous.counters(), { shown: 1, initialized: 0 });
    let resolve;
    const delayed = accessClientFixture({ fetch: () => new Promise(done => { resolve = done; }) });
    const task = delayed.restoreAccessSession();
    await Promise.resolve();
    delayed.cancel();
    resolve({ ok: true, json: async () => ({ token: 'late-token', user: { username: 'operator' } }) });
    await task;
    assert.equal(delayed.state().accessGranted, false);
    assert.equal(delayed.counters().initialized, 0);
});

test('password remembering is opt-in after successful login and never writes passwords to browser storage', async () => {
    for (const remember of [false, true]) {
        let remembered = null;
        let forgotten = 0;
        const client = accessClientFixture({ desktop: {
            remember: async data => { remembered = data; }, forget: async () => { forgotten += 1; }
        }, fetch: async (url, options) => {
            assert.equal(url, '/auth/login');
            assert.equal(JSON.parse(options.body).password, ' synthetic ');
            return { ok: true, json: async () => ({ token: 'new-token', user: { username: 'Operator' } }) };
        } });
        client.document.getElementById('accessGateUserInput').value = 'operator';
        client.document.getElementById('accessGateInput').value = ' synthetic ';
        client.document.getElementById('accessGateRememberPassword').checked = remember;
        await client.submitAccessCode();
        assert.equal(client.state().accessGranted, true);
        assert.equal(client.document.getElementById('accessGateInput').value, '');
        assert.equal(forgotten, remember ? 0 : 1);
        assert.equal(Boolean(remembered), remember);
        if (remembered) assert.deepEqual(remembered, { username: 'Operator', password: ' synthetic ', token: 'new-token' });
        assert.deepEqual([...client.storage.values()], remember ? ['Operator'] : []);
    }
});
