const test = require('node:test');
const assert = require('node:assert/strict');
const { createAuthService } = require('../services/authService');

function createService() {
    return createAuthService({ fs: {}, usersFile: 'unused-in-these-tests.json' });
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
