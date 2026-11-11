const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const crypto = require('node:crypto');
const path = require('node:path');

function fixture() {
    const source = fs.readFileSync(path.join(__dirname, '../routes/apiRoutes.js'), 'utf8');
    const start = source.indexOf('    const bilancioArchive =');
    const end = source.indexOf("    router.get('/aftercalc/:ordno'", start);
    const routes = new Map(), calls = [], pdf = Buffer.from('%PDF-original');
    const saved = { id: 'snapshot', report: { reportId: 'default', year: 2026, period: 4 },
        revision: 1, pdfBase64: pdf.toString('base64'), pdfSha256: crypto.createHash('sha256').update(pdf).digest('hex') };
    const archive = {
        async load(...args) { calls.push(['load', ...args]); return saved; },
        async save(...args) { calls.push(['save', ...args]); return saved; },
        async draft(...args) { calls.push(['draft', ...args]); return {}; },
        async list(...args) { calls.push(['list', ...args]); return []; }
    };
    const router = Object.fromEntries(['get', 'post'].map(method => [method, (url, ...handlers) => routes.set(method + url, handlers)]));
    const guard = module => { assert.equal(module, 'bilancio'); return (req, res, next) => req.allowed ? next() : res.status(403).json({ error: 'denied' }); };
    new Function('router', 'createBilancioArchiveService', 'gohData', 'bilancioDefinitionStore', 'bilancioService', 'generateBilancioPdf', 'pkgVersion',
        'requireModulePermission', 'getSessionUser', 'validateBilancioPeriod', 'logEvent', 'crypto', source.slice(start, end))(
        router, () => archive, {}, {}, {}, () => {}, '1.1.95', guard, () => ({ username: 'alice' }), () => {}, () => {}, crypto
    );
    return {
        calls, saved, pdf,
        async request(method, url, allowed, query = {}, body = {}) {
            const result = { status: 200, headers: {} };
            const res = {
                setHeader(name, value) { result.headers[name] = value; },
                status(code) { result.status = code; return this; },
                json(data) { result.data = data; }, send(data) { result.data = data; }
            };
            const [guardFn, handler] = routes.get(method + url);
            let proceed = false;
            const req = { allowed, query, body, params: { id: 'snapshot' } };
            guardFn(req, res, () => { proceed = true; });
            if (proceed) await handler(req, res);
            return result;
        }
    };
}
test('every archive and draft endpoint requires bilancio permission before accessing stored data', async () => {
    const f = fixture();
    for (const [method, url] of [['get', '/bilancio-trial/draft'], ['get', '/bilancio-trial/archive'],
        ['post', '/bilancio-trial/archive'], ['get', '/bilancio-trial/archive/:id'], ['get', '/bilancio-trial/archive/:id/download']]) {
        assert.equal((await f.request(method, url, false)).status, 403);
    }
    assert.equal(f.calls.length, 0);
});
test('downloads return stored PDF bytes, JSON excludes binary data, and corrupt PDF fails explicitly', async () => {
    const f = fixture();
    const response = await f.request('get', '/bilancio-trial/archive/:id/download', true, { reportId: 'default' });
    assert.deepEqual(response.data, f.pdf);
    assert.equal(response.headers['Content-Type'], 'application/pdf');
    const json = await f.request('get', '/bilancio-trial/archive/:id/download', true, { format: 'json' });
    assert.equal(json.data.pdfBase64, undefined);
    assert.equal(json.data.revision, 1);
    f.saved.pdfBase64 = Buffer.from('tampered').toString('base64');
    assert.equal((await f.request('get', '/bilancio-trial/archive/:id/download', true)).status, 503);
});
