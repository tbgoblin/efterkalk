const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const express = require('express');
const { createMileageService, calculateAllowance } = require('../services/mileageService');
const { parseMileageImport, parseCsv } = require('../services/mileageImportService');
const { createMileageRoutingService } = require('../services/mileageRoutingService');
const { generateMileagePdf } = require('../services/mileagePdf');
const { createMileageRouter } = require('../routes/mileageRoutes');

const worker = { username: 'worker', displayName: 'Test Worker', role: 'user', permissions: { mileage: true } };
const other = { username: 'other', displayName: 'Other Worker', role: 'user', permissions: { mileage: true } };
const approver = { username: 'approver', role: 'user', permissions: { mileage: true, mileageManage: true } };
const users = [worker, other, approver];
function store() {
    let payload;
    return {
        failWrites: false,
        async getAppState(_key, options) { assert.equal(options.strict, true); return payload ? { payload: structuredClone(payload) } : null; },
        async setAppState(_key, next, options) {
            if (this.failWrites || (payload?.version || 0) !== options.expectedVersion) return false;
            payload = structuredClone(next);
            return true;
        }
    };
}
async function fixture() {
    const gohData = store();
    const service = createMileageService({ gohData, readUsers: () => users, now: () => '2026-11-23T12:00:00Z' });
    await service.saveEmployee({ version: 0, name: worker.displayName, username: worker.username, plate: 'TEST', active: true }, approver);
    await service.saveEmployee({ version: 1, name: other.displayName, username: other.username, plate: 'OTHER', active: true }, approver);
    await service.saveRate({ version: 2, year: 2026, high: 4, low: 2 }, approver);
    const save = async (overrides = {}, user = worker) => {
        const data = await service.data(user);
        await service.saveTrip({ version: data.version, date: '2026-11-20', plate: 'TEST', purpose: 'Business visit',
            stops: [{ address: 'Test start', note: '' }, { address: 'Test stop', note: 'Pause' }, { address: 'Test end', note: '' }],
            km: 100, roundTrip: true, manualReason: 'Odometer, includes return', comment: '', ...overrides }, user);
        return (await service.data(user)).trips.at(-1);
    };
    const act = async (trip, action, user = approver, note = '') => service.transition(trip.id, {
        version: (await service.data(user)).version, action, note
    }, user);
    return { gohData, service, save, act };
}

test('allowance correctly splits the 20,000-km crossing and rounds only money', () => {
    assert.deepEqual(calculateAllowance(100, 19950, { year: 2026, high: 4, low: 2 }), {
        previousKm: 19950, highKm: 50, lowKm: 50, highRate: 4, lowRate: 2, amount: 300, year: 2026
    });
    assert.equal(calculateAllowance(10, 20000, { year: 2026, high: 4, low: 2 }).amount, 20);
    assert.equal(calculateAllowance(1.23, 0, { year: 2026, high: 3.795, low: 2 }).amount, 4.67);
});

test('trip workflow enforces ownership, independent approval, fixed rates, and payment once', async () => {
    const f = await fixture();
    const trip = await f.save();
    assert.equal((await f.service.data(worker)).trips[0].estimate.amount, 400);
    assert.equal((await f.service.data(other)).trips.length, 0);
    await assert.rejects(f.service.document(trip.id, other), { statusCode: 404 });
    await f.act(trip, 'submit', worker);
    await assert.rejects(f.act(trip, 'approve', worker), { statusCode: 403 });
    const selfManager = { ...worker, permissions: approver.permissions };
    await assert.rejects(f.act(trip, 'approve', selfManager), /egen kørsel/);
    const employee = (await f.service.data(approver)).employees.find(row => row.username === worker.username);
    await f.service.saveEmployee({ ...employee, version: (await f.service.data(approver)).version, username: '' }, approver);
    await assert.rejects(f.act(trip, 'approve', selfManager), /egen kørsel/);
    await f.service.saveEmployee({ ...employee, version: (await f.service.data(approver)).version }, approver);
    await f.act(trip, 'approve');
    let approved = (await f.service.data(worker)).trips[0];
    assert.equal(approved.allowance.amount, 400);
    assert.equal(approved.km, 100); // Total km already includes return; never double it again.
    await f.service.saveRate({ version: (await f.service.data(approver)).version, year: 2026, high: 5, low: 3 }, approver);
    approved = (await f.service.data(worker)).trips[0];
    assert.equal(approved.allowance.highRate, 4);
    await f.act(trip, 'pay');
    await assert.rejects(f.act(trip, 'pay'), { statusCode: 409 });
    await assert.rejects(f.save({ id: trip.id }), { statusCode: 409 });
    assert.equal((await f.service.data(worker)).trips[0].history.at(-1).action, 'pay');
});

test('reject requires a reason and allows the owner to correct and resubmit', async () => {
    const f = await fixture();
    const trip = await f.save();
    await f.act(trip, 'submit', worker);
    await assert.rejects(f.act(trip, 'reject'), /Begrundelse/);
    await f.act(trip, 'reject', approver, 'Explain route');
    await f.save({ id: trip.id, km: 80, manualReason: 'Corrected odometer reading' });
    await f.act(trip, 'submit', worker);
    await f.act(trip, 'approve');
    assert.equal((await f.service.data(worker)).trips[0].allowance.amount, 320);
});

test('backdated trips use their own calendar-year rate and exclude other years from the threshold', async () => {
    const f = await fixture();
    await f.service.saveRate({ version: (await f.service.data(approver)).version, year: 2025, high: 3, low: 1 }, approver);
    const previousYear = await f.save({ date: '2025-12-31', km: 10000 });
    await f.act(previousYear, 'submit', worker); await f.act(previousYear, 'approve');
    const thisYear = await f.save();
    await f.act(thisYear, 'submit', worker); await f.act(thisYear, 'approve');
    assert.equal((await f.service.document(previousYear.id, worker)).allowance.amount, 30000);
    assert.equal((await f.service.document(thisYear.id, worker)).allowance.previousKm, 0);
    await assert.rejects(f.save({ stops: [null, { address:'Test' }] }), { statusCode:400 });
});

test('stale or failed GOH writes never report success, and concurrent approvals share the annual limit', async () => {
    const f = await fixture();
    await assert.rejects(f.service.saveRate({ version: 0, year: 2026, high: 4, low: 2 }, approver), { statusCode: 409 });
    f.gohData.failWrites = true;
    await assert.rejects(f.save(), { statusCode: 409 });
    assert.equal((await f.service.data(worker)).trips.length, 0);
    f.gohData.failWrites = false;
    const a = await f.save({ km: 10000 });
    await f.act(a, 'submit', worker); await f.act(a, 'approve');
    const b = await f.save({ km: 9990 });
    await f.act(b, 'submit', worker); await f.act(b, 'approve');
    const c = await f.save({ km: 20 }); await f.act(c, 'submit', worker);
    const d = await f.save({ km: 20 }); await f.act(d, 'submit', worker);
    const version = (await f.service.data(approver)).version;
    const results = await Promise.allSettled([c, d].map(trip => f.service.transition(trip.id, { version, action: 'approve' }, approver)));
    assert.equal(results.filter(row => row.status === 'fulfilled').length, 1);
    const pending = (await f.service.data(worker)).trips.find(row => row.status === 'submitted');
    await f.act(pending, 'approve');
    const amounts = (await f.service.data(worker)).trips.slice(-2).map(row => row.allowance.amount).sort((a,b) => a-b);
    assert.deepEqual(amounts, [40, 60]);
});

const importFiles = () => ({
    employees: 'Bruger,Nummerplade,Aktiv,Admin\nTest Worker,TEST,Sandt,Sandt\n',
    trips: 'Id,Navn,Status,Nr_Plade,Dato_Dato,Beregnet_km,Retur,Udbetalt,Formaal\n1,Test Worker,Godkendt,TEST,20-11-2026,100,Sandt,Falsk,"Visit, with return"\n2,Test Worker,Godkendt,TEST,,50,Falsk,Sandt,Visit\n',
    addresses: 'Titel,Addresse,Firmanavn\nOffice,"Example road 1, Denmark",Gantech\n',
    purposes: 'Titel,Formaal\nVisit,Business visit\n'
});
test('legacy CSV preserves purpose and total km, does not invent amounts or grant permissions', () => {
    const result = parseMileageImport(importFiles());
    assert.equal(result.trips[0].purpose, 'Visit, with return');
    assert.equal(result.trips[0].km, 100);
    assert.equal(result.trips[0].allowance, null);
    assert.equal(result.trips[1].date, null);
    assert.equal(result.trips[1].legacyPaid, true);
    assert.equal(result.employees[0].username, '');
    assert.equal(Object.hasOwn(result.employees[0], 'admin'), false);
    assert.equal(result.warnings.some(row => row.includes('mangler dato')), true);
    const titleOnly = importFiles();
    titleOnly.purposes = 'Titel,Formaal\nCustomer visit,\n';
    assert.deepEqual(parseMileageImport(titleOnly).purposes, ['Customer visit']);
    assert.deepEqual(parseCsv('A;B\r\n"Line1\nLine2";"Quote ""yes"""\r\n'), [{ A: 'Line1\nLine2', B: 'Quote "yes"' }]);
    assert.throws(() => parseCsv('A,B\n"broken,b\n'), /Uafsluttet/);
    const files = importFiles(); files.trips += files.trips.split('\n')[1] + '\n';
    assert.throws(() => parseMileageImport(files), /Dublerede/);
});

test('one-time import is atomic and undated legacy trips block approval until corrected', async () => {
    const gohData = store();
    const service = createMileageService({ gohData, readUsers: () => users });
    await assert.rejects(service.importData({ version: 0, files: importFiles() }, worker), { statusCode: 403 });
    await service.importData({ version: 0, files: importFiles() }, approver);
    await assert.rejects(service.importData({ version: 1, files: importFiles() }, approver), { statusCode: 409 });
    let data = await service.data(approver);
    await service.saveEmployee({ ...data.employees[0], username: worker.username, version: data.version }, approver);
    data = await service.data(approver);
    await service.saveRate({ year: 2026, high: 4, low: 2, version: data.version }, approver);
    data = await service.data(worker);
    await service.saveTrip({ version: data.version, date: '2026-11-20', plate: 'TEST', purpose: 'Visit',
        km: 100, stops: [{ address: 'Start' }, { address: 'End' }], manualReason: 'Odometer' }, worker);
    let trip = (await service.data(worker)).trips.at(-1);
    await service.transition(trip.id, { version: (await service.data(worker)).version, action: 'submit' }, worker);
    await assert.rejects(service.transition(trip.id, { version: (await service.data(worker)).version, action: 'approve' }, approver), /uden dato/);
    const legacy = data.trips.find(row => !row.date);
    await service.correctLegacyDate(legacy.id, { version: (await service.data(approver)).version, date: '2026-01-01' }, approver);
    await service.transition(trip.id, { version: (await service.data(worker)).version, action: 'approve' }, approver);
    trip = await service.document(trip.id, worker);
    assert.equal(trip.allowance.previousKm, 150);
    await assert.rejects(service.transition(legacy.id, { version: (await service.data(worker)).version, action: 'pay' }, approver), /Historik|historik/);
});

test('routing is opt-in, validates provider results and routes return through reversed stops', async () => {
    const disabled = createMileageRoutingService({ apiKey: () => '' });
    await assert.rejects(disabled.search('Example'), /ikke konfigureret/);
    const calls = [];
    const routing = createMileageRoutingService({ apiKey: () => 'test-only-key', fetchImpl: async (url, options) => {
        calls.push({ url, options });
        if (url.includes('/geocode/')) return { ok:true, json:async() => ({ features:[
            { properties:{label:'Example address'}, geometry:{coordinates:[10,55]} },
            { properties:{label:'Invalid'}, geometry:{coordinates:[200,55]} }, null
        ] }) };
        return { ok: true, json: async () => ({ routes: [{ segments: [{ distance: 1000 }, { distance: 2000 }, { distance: 2500 }, { distance: 1100 }] }] }) };
    } });
    const points = [[10,55],[11,56],[12,57]];
    const result = await routing.route(points, true);
    assert.deepEqual(JSON.parse(calls[0].options.body).coordinates, [[10,55],[11,56],[12,57],[11,56],[10,55]]);
    assert.equal(JSON.parse(calls[0].options.body).instructions, true);
    assert.equal(result.km, 6.6);
    assert.deepEqual(await routing.search('Example address'), [{ address:'Example address',coordinates:[10,55] }]);
    assert.equal(new URL(calls[1].url).searchParams.get('api_key'), 'test-only-key');
    await assert.rejects(routing.route([[200,55],[12,57]],false), /gyldig adresse/);
    const broken = createMileageRoutingService({ apiKey: () => 'test-only-key', fetchImpl: async () => ({ ok: true, json: async () => null }) });
    await assert.rejects(broken.route(points,false), /alle delstrækninger/);
    await assert.rejects(broken.search('Example'), /ingen adresseliste/);
});

test('internal PDF is generated from authorized trip data including fixed allowance', async () => {
    const f = await fixture();
    const trip = await f.save();
    await f.act(trip, 'submit', worker); await f.act(trip, 'approve');
    const pdf = await generateMileagePdf([await f.service.document(trip.id, worker)]);
    assert.equal(pdf.subarray(0,5).toString(), '%PDF-');
    assert.ok(pdf.length > 2000);
});

test('API protects configuration and import separately from employee access', async t => {
    const app = express(), gohData = store();
    app.use((req,res,next) => req.path === '/mileage/import' ? next() : express.json()(req,res,next));
    const getSessionUser = req => users.find(row => row.username === req.headers['x-test-user']);
    const requireModulePermission = permission => (req,res,next) => getSessionUser(req)?.permissions[permission]
        ? next() : res.status(403).json({ error: 'Denied' });
    app.use('/mileage', createMileageRouter({ gohData, readUsers: () => users, getSessionUser, requireModulePermission, logEvent() {} }));
    const server = http.createServer(app);
    await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    t.after(() => new Promise(resolve => { server.close(resolve); server.closeAllConnections(); }));
    const base = 'http://127.0.0.1:' + server.address().port;
    assert.equal((await fetch(base+'/mileage/data')).status, 403);
    const response = await fetch(base+'/mileage/data', { headers: { 'x-test-user':'worker' } });
    assert.equal(response.status, 200);
    assert.deepEqual((await response.json()).trips, []);
    assert.equal((await fetch(base+'/mileage/rates', { method:'POST', headers:{ 'x-test-user':'worker','Content-Type':'application/json' }, body:JSON.stringify({year:2026}) })).status, 403);
    const preview = await fetch(base+'/mileage/import', { method:'POST', headers:{ 'x-test-user':'approver','Content-Type':'application/json' },
        body: JSON.stringify({preview:true,files:importFiles()}) });
    assert.equal(preview.status, 200);
    assert.equal((await preview.json()).counts.trips, 2);
    assert.equal((await gohData.getAppState('',{strict:true})), null);
});
