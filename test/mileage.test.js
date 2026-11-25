const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const express = require('express');
const { createMileageService, calculateAllowance } = require('../services/mileageService');
const { parseMileageImport, parseCsv } = require('../services/mileageImportService');
const { createMileageRoutingService, createMileageRoutingSettings } = require('../services/mileageRoutingService');
const { generateMileagePdf } = require('../services/mileagePdf');
const { createMileageRouter } = require('../routes/mileageRoutes');

const worker = { username: 'worker', displayName: 'Test Worker', role: 'user', permissions: { mileage: true } };
const other = { username: 'other', displayName: 'Other Worker', role: 'user', permissions: { mileage: true } };
const approver = { username: 'approver', role: 'user', permissions: { mileage: true, mileageManage: true } };
const users = [worker, other, approver];
function store() {
    const values = new Map();
    return {
        failWrites: false,
        async getAppState(key, options) { assert.equal(options.strict, true); return values.has(key) ? { payload: structuredClone(values.get(key)) } : null; },
        async setAppState(key, next, options) {
            if (this.failWrites || (values.get(key)?.version || 0) !== options.expectedVersion) return false;
            values.set(key, structuredClone(next));
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

test('shared addresses are manager-only and preserve imports while invalidating changed coordinates', async () => {
    const gohData = store();
    await gohData.setAppState('mileage_v1', { schema:1, version:0, employees:[], addresses:[
        { name:'Example Company', address:'Test Street 1' }
    ], purposes:[], rates:[], trips:[], imports:[] }, { expectedVersion:0 });
    const service = createMileageService({ gohData, readUsers: () => users });
    await assert.rejects(service.saveAddress({ version:0, name:'example company', address:'Test Street 1', coordinates:[10, 55] }, worker), { statusCode:403 });
    await service.saveAddress({ version:0, name:'example company', address:'Test Street 1', coordinates:[10, 55] }, approver);
    let data = await service.data(other);
    assert.equal(data.addresses.length, 1);
    assert.deepEqual(data.addresses[0].coordinates, [10, 55]);
    await service.saveAddress({ version:1, name:'Example Company', address:'Test Street 2',
        originalName:'Example Company', originalAddress:'Test Street 1' }, approver);
    data = await service.data(worker);
    assert.equal(data.addresses.length, 1);
    assert.equal(data.addresses[0].address, 'Test Street 2');
    assert.equal(data.addresses[0].coordinates, undefined);
    await assert.rejects(service.saveAddress({ version:2, name:'', address:'Test' }, approver), { statusCode:400 });
    await assert.rejects(service.saveAddress({ version:2, name:'Bad', address:'Test', coordinates:[181, 55] }, approver), { statusCode:400 });
    await assert.rejects(service.saveAddress({ version:1, name:'Stale', address:'Test' }, approver), { statusCode:409 });
    const results = await Promise.allSettled([
        service.saveAddress({ version:2, name:'One', address:'Test 1' }, approver),
        service.saveAddress({ version:2, name:'Two', address:'Test 2' }, approver)
    ]);
    assert.equal(results.filter(result => result.status === 'fulfilled').length, 1);
    assert.equal(results.filter(result => result.status === 'rejected' && result.reason.statusCode === 409).length, 1);
    assert.equal((await service.data(worker)).addresses.length, 2);
});

test('managers can edit and remove shared purposes and addresses without changing old trips', async () => {
    const f = await fixture();
    await assert.rejects(f.service.savePurpose({ version:3, purpose:'Customer visit' }, worker), { statusCode:403 });
    await f.service.savePurpose({ version:3, purpose:'Customer visit' }, approver);
    await f.service.savePurpose({ version:4, original:'Customer visit', purpose:'Customer meeting' }, approver);
    await assert.rejects(f.service.savePurpose({ version:5, purpose:'customer MEETING' }, approver), { statusCode:409 });
    assert.deepEqual((await f.service.data(worker)).purposes, ['Customer meeting']);
    await f.service.deletePurpose({ version:5, purpose:'Customer meeting' }, approver);
    await assert.rejects(f.service.deletePurpose({ version:6, purpose:'Customer meeting' }, approver), { statusCode:404 });
    assert.deepEqual((await f.service.data(worker)).purposes, []);

    await f.service.saveAddress({ version:6, name:'Office', address:'Main Road 1' }, approver);
    await f.save({ purpose:'Customer meeting' });
    await f.service.deleteAddress({ version:8, name:'Office', address:'Main Road 1' }, approver);
    assert.equal((await f.service.data(worker)).trips[0].purpose, 'Customer meeting');
    assert.deepEqual((await f.service.data(worker)).addresses, []);
});

test('trip workflow enforces permission, fixed rates, and payment once', async () => {
    const f = await fixture();
    const trip = await f.save();
    assert.equal((await f.service.data(worker)).trips[0].estimate.amount, 400);
    assert.equal((await f.service.data(other)).trips.length, 0);
    await assert.rejects(f.service.document(trip.id, other), { statusCode: 404 });
    await f.act(trip, 'submit', worker);
    await assert.rejects(f.act(trip, 'approve', worker), { statusCode: 403 });
    const selfManager = { ...worker, permissions: approver.permissions };
    await assert.rejects(f.act(trip, 'reject', selfManager, 'Test'), { statusCode: 409 });
    await f.act(trip, 'approve', selfManager);
    let approved = (await f.service.data(worker)).trips[0];
    assert.equal(approved.approvedBy, worker.username);
    assert.equal(approved.history.at(-1).by, worker.username);
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

test('a manager can approve a trip they created for another employee or one created for them', async () => {
    const f = await fixture();
    const otherId = (await f.service.data(other)).employeeId;
    const trip = await f.save({ employeeId: otherId }, worker);
    assert.equal(trip.employeeId, otherId);
    assert.equal((await f.service.data(other)).trips[0].id, trip.id);
    assert.equal((await f.service.document(trip.id, worker)).id, trip.id);
    await assert.rejects(f.save({ id: trip.id, employeeId: (await f.service.data(worker)).employeeId }, worker), /ikke ændres/);
    await f.save({ id: trip.id, employeeId: otherId, km: 80 }, other);
    await f.act(trip, 'submit', worker);
    const creatorManager = { ...worker, permissions: approver.permissions };
    const ownerManager = { ...other, permissions: approver.permissions };
    await f.act(trip, 'approve', creatorManager);
    const approved = await f.service.document(trip.id, other);
    assert.equal(approved.allowance.amount, 320);
    assert.equal(approved.approvedBy, worker.username);
    const second = await f.save({ employeeId: otherId }, worker);
    await f.act(second, 'submit', worker);
    await f.act(second, 'approve', ownerManager);
    assert.equal((await f.service.document(second.id, other)).approvedBy, other.username);
});

test('Visma synchronization preserves GOH links, tracks active employees and rejects ambiguous or different sources', async () => {
    const f = await fixture();
    const before = await f.service.data(approver);
    const rows = [{ EmpNo: 10, Nm: 'Test Worker', Gr2: 0 }, { EmpNo: 20, Nm: 'New Worker', Gr2: 0 }];
    assert.equal((await f.service.syncEmployees(rows, 'company-source', worker)).changed, true);
    let data = await f.service.data(approver);
    assert.equal(data.employees.find(row => row.vismaEmpNo === '10').id, before.employees[0].id);
    assert.equal(data.employees.find(row => row.vismaEmpNo === '10').username, worker.username);
    assert.equal(data.participants.some(row => row.name === 'New Worker'), true);
    assert.equal((await f.service.syncEmployees(rows, 'company-source', worker)).changed, false);
    assert.equal((await f.service.data(approver)).version, data.version);
    await f.service.syncEmployees([{ EmpNo: 10, Nm: 'Renamed Worker', Gr2: 99 }, rows[1]], 'company-source', worker);
    data = await f.service.data(approver);
    assert.equal(data.employees.find(row => row.vismaEmpNo === '10').username, worker.username);
    assert.equal(data.participants.some(row => row.name === 'Renamed Worker'), false);
    await assert.rejects(f.service.syncEmployees(rows, 'other-source', worker), /profilen afviger/);
    await assert.rejects(f.service.syncEmployees([{ EmpNo: 10, Nm: 'Duplicate', Gr2: 0 }, { EmpNo: 10, Nm: 'Different', Gr2: 0 }], 'company-source', worker), /dublerede/);
    await assert.rejects(f.service.syncEmployees([], 'company-source', worker), /ingen gyldig/);
    assert.equal((await f.service.data(approver)).version, data.version);
    const synced = data.employees.find(row => row.vismaEmpNo === '10');
    await assert.rejects(f.service.saveEmployee({ ...synced, name: 'Altered', version: data.version }, approver), /styres af Visma/);
});

test('duplicate Visma names do not attach ambiguous employees, and unlinked inactive entries are ignored', async () => {
    const f = await fixture();
    const result = await f.service.syncEmployees([
        { EmpNo:10, Nm:'Test Worker', Gr2:0 }, { EmpNo:11, Nm:'Test Worker', Gr2:99 },
        { EmpNo:20, Nm:'Past Worker', Gr2:99 }, { EmpNo:21, Nm:'Past Worker', Gr2:99 },
        { EmpNo:30, Nm:'Active New', Gr2:0 }
    ], 'company-source', worker);
    assert.match(result.warnings.join(' '), /navnesammenfald/);
    const data = await f.service.data(approver);
    assert.equal(data.employees.find(row => row.username === worker.username).vismaEmpNo, undefined);
    assert.equal(data.employees.some(row => row.name === 'Past Worker'), false);
    assert.equal(data.employees.find(row => row.vismaEmpNo === '30').name, 'Active New');
});

test('manager deletion is audited, excludes km and marks affected approved trips for review', async () => {
    const f = await fixture();
    const first = await f.save({ km: 10000 });
    await f.act(first, 'submit', worker); await f.act(first, 'approve'); await f.act(first, 'pay');
    const second = await f.save({ km: 100 });
    await f.act(second, 'submit', worker); await f.act(second, 'approve');
    const deleteInput = { version: (await f.service.data(approver)).version, note: 'Duplicate entry' };
    await assert.rejects(f.service.deleteTrip(first.id, deleteInput, worker), { statusCode: 403 });
    await assert.rejects(f.service.deleteTrip(first.id, { ...deleteInput, note: '' }, approver), /Begrundelse/);
    await f.service.deleteTrip(first.id, deleteInput, approver);
    assert.equal((await f.service.data(worker)).trips.some(row => row.id === first.id), false);
    const deleted = (await f.service.data(approver)).trips.find(row => row.id === first.id);
    assert.equal(deleted.status, 'deleted');
    assert.equal(deleted.deletedPreviousStatus, 'paid');
    assert.equal(deleted.history.at(-1).note, 'Duplicate entry');
    const affected = await f.service.document(second.id, worker);
    assert.equal(affected.reconciliationRequired, true);
    assert.equal(affected.allowance.previousKm, 10000);
    await assert.rejects(f.service.document(first.id, worker), { statusCode: 404 });
    await assert.rejects(f.act(first, 'pay'), { statusCode: 409 });
    const third = await f.save({ km: 20 });
    await f.act(third, 'submit', worker); await f.act(third, 'approve');
    assert.equal((await f.service.document(third.id, worker)).allowance.previousKm, 100);
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

test('import merges existing employees and trips while flagging affected approved amounts', async () => {
    const f = await fixture();
    const existing = await f.save();
    await f.act(existing, 'submit', worker); await f.act(existing, 'approve');
    const before = await f.service.data(approver);
    const preview = await f.service.previewImport(importFiles(), approver);
    assert.equal(preview.counts.trips, 2);
    assert.match(preview.warnings.join(' '), /skal gennemgås/);
    assert.equal((await f.service.data(approver)).version, before.version);
    await f.service.importData({ version: before.version, files: importFiles() }, approver);
    const after = await f.service.data(approver);
    assert.equal(after.employees.length, 2);
    assert.equal(after.employees.find(row => row.username === worker.username).id, before.employees[0].id);
    assert.equal(after.trips.length, 3);
    assert.ok(after.trips.filter(row => row.legacy).every(row => row.employeeId === before.employees[0].id));
    assert.equal(after.trips.find(row => row.id === existing.id).reconciliationRequired, true);
    assert.equal(after.trips.find(row => row.id === existing.id).allowance.amount, 400);
    await assert.rejects(f.service.previewImport(importFiles(), approver), /allerede importeret/);
});

test('routing is opt-in, validates provider results and routes return through reversed stops', async () => {
    const disabled = createMileageRoutingService({ apiKey: () => '' });
    await assert.rejects(disabled.search('Example'), /ikke konfigureret/);
    const calls = [];
    const routing = createMileageRoutingService({ apiKey: () => 'test-only-key', fetchImpl: async (url, options) => {
        calls.push({ url, options });
        if (url.includes('/geocode/')) return { ok:true, json:async() => ({ features:[
            { properties:{label:'Example address'}, geometry:{coordinates:[10,55]} },
            { properties:{name:'Hotel Example', label:'Hotel Example, Test City', street:'Testvej', housenumber:'25'}, geometry:{coordinates:[11,56]} },
            { properties:{label:'Testvej 25, Test City', street:'Testvej', housenumber:'25'}, geometry:{coordinates:[12,57]} },
            { properties:{label:'Other Hotel, Test City', street:'Hotelvej'}, geometry:{coordinates:[13,58]} },
            { properties:{label:'Invalid'}, geometry:{coordinates:[200,55]} }, null
        ] }) };
        return { ok: true, json: async () => ({ routes: [{ segments: [{ distance: 1000 }, { distance: 2000 }, { distance: 2500 }, { distance: 1100 }] }] }) };
    } });
    const points = [[10,55],[11,56],[12,57]];
    const result = await routing.route(points, true);
    assert.deepEqual(JSON.parse(calls[0].options.body).coordinates, [[10,55],[11,56],[12,57],[11,56],[10,55]]);
    assert.equal(JSON.parse(calls[0].options.body).instructions, true);
    assert.equal(result.km, 6.6);
    assert.deepEqual(await routing.search('Example address'), [
        { name:'',address:'Example address',streetAddressKnown:false,coordinates:[10,55] },
        { name:'Hotel Example',address:'Hotel Example, Test City · Testvej 25',streetAddressKnown:true,coordinates:[11,56] },
        { name:'',address:'Testvej 25, Test City',streetAddressKnown:true,coordinates:[12,57] },
        { name:'',address:'Other Hotel, Test City · Hotelvej',streetAddressKnown:false,coordinates:[13,58] }
    ]);
    assert.equal(new URL(calls[1].url).searchParams.get('api_key'), 'test-only-key');
    assert.equal(new URL(calls[1].url).searchParams.get('size'), '10');
    await assert.rejects(routing.route([[200,55],[12,57]],false), /gyldig adresse/);
    const broken = createMileageRoutingService({ apiKey: () => 'test-only-key', fetchImpl: async () => ({ ok: true, json: async () => null }) });
    await assert.rejects(broken.route(points,false), /alle delstrækninger/);
    await assert.rejects(broken.search('Example'), /ingen adresseliste/);
    for (const [status, message] of [
        [401, /API-nøglen eller adgangen/],
        [403, /API-nøglen eller adgangen/],
        [400, /afviste forespørgslen/],
        [500, /HTTP 500/]
    ]) {
        const rejected = createMileageRoutingService({
            apiKey: () => 'test-only-key',
            fetchImpl: async () => ({ ok: false, status })
        });
        await assert.rejects(rejected.search('Example'), message);
    }
});

test('routing settings share a write-only database key across instances and handle conflicts and removal', async () => {
    const gohData = store();
    const settings = createMileageRoutingSettings({ gohData, envKey: () => 'environment-test-key' });
    const second = createMileageRoutingSettings({ gohData, envKey: () => 'environment-test-key' });
    assert.deepEqual(await settings.status(), { version: 0, configured: true, source: 'environment' });
    assert.deepEqual(await settings.save({ version: 0, action: 'save', apiKey: 'database-test-key' }, approver),
        { version: 1, configured: true, source: 'database' });
    assert.equal(await second.key(), 'database-test-key');
    let usedKey;
    const routing = createMileageRoutingService({ apiKey: second.key, fetchImpl: async (url, options) => {
        usedKey = options.headers.Authorization;
        assert.equal(new URL(url).searchParams.get('api_key'), usedKey);
        return { ok: true, json: async () => ({ features: [] }) };
    } });
    await routing.search('Test address');
    assert.equal(usedKey, 'database-test-key');
    await assert.rejects(second.save({ version: 0, action: 'remove' }, approver), { statusCode: 409 });
    await assert.rejects(second.save({ version: 1, action: 'save', apiKey: 'bad\nkey' }, approver), { statusCode: 400 });
    gohData.failWrites = true;
    await assert.rejects(second.save({ version: 1, action: 'remove' }, approver), { statusCode: 409 });
    gohData.failWrites = false;
    assert.deepEqual(await second.save({ version: 1, action: 'remove' }, approver),
        { version: 2, configured: true, source: 'environment' });
    assert.equal(await settings.key(), 'environment-test-key');
    await routing.search('Test address');
    assert.equal(usedKey, 'environment-test-key');
    const noFallback = createMileageRoutingSettings({ gohData, envKey: () => '' });
    assert.equal((await noFallback.status()).configured, false);
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
    const superadmin = { username: 'admin-test', role: 'superadmin', permissions: {} };
    const getSessionUser = req => [...users, superadmin].find(row => row.username === req.headers['x-test-user']);
    const requireModulePermission = permission => (req,res,next) => getSessionUser(req)?.permissions[permission]
        ? next() : res.status(403).json({ error: 'Denied' });
    const requireSuperadmin = (req,res) => {
        const user = getSessionUser(req);
        if (user?.role === 'superadmin') return user;
        res.status(403).json({ error: 'Denied' });
        return null;
    };
    const readVismaEmployees = async () => ({ source:'test-company', rows:[{ EmpNo:10, Nm:'Test Worker', Gr2:0 }] });
    let customerReadFails = false;
    const readVismaMileageActors = async () => {
        if (customerReadFails) throw new Error('test failure');
        return [
            { ActorType:'customer', CustNo:'61209086', SupNo:'', Nm:'FarmDroid ApS', Ad1:'Industrivej 1', Ad2:'', PNo:'6600', PArea:'Vejen' },
            { ActorType:'supplier', CustNo:'', SupNo:'370488', Nm:'Ravema', Ad1:'Margretelundsväg', Ad2:'', PNo:'331 34', PArea:'Värnamo' }
        ];
    };
    app.use('/mileage', createMileageRouter({
        gohData, readUsers: () => users, readVismaEmployees, readVismaMileageActors,
        getSessionUser, requireModulePermission, requireSuperadmin, logEvent() {}
    }));
    const server = http.createServer(app);
    await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    t.after(() => new Promise(resolve => { server.close(resolve); server.closeAllConnections(); }));
    const base = 'http://127.0.0.1:' + server.address().port;
    assert.equal((await fetch(base+'/mileage/data')).status, 403);
    const response = await fetch(base+'/mileage/data', { headers: { 'x-test-user':'worker' } });
    assert.equal(response.status, 200);
    assert.deepEqual((await response.json()).trips, []);
    assert.equal((await (await fetch(base+'/mileage/data', { headers:{'x-test-user':'worker'} })).json()).isSuperadmin, false);
    assert.equal((await fetch(base+'/mileage/visma-actors')).status, 403);
    const customersResponse = await fetch(base+'/mileage/visma-actors', { headers:{'x-test-user':'worker'} });
    assert.equal(customersResponse.status, 200);
    assert.deepEqual((await customersResponse.json()).actors, [
        { ActorType:'customer', CustNo:'61209086', SupNo:'', Nm:'FarmDroid ApS', Ad1:'Industrivej 1', Ad2:'', PNo:'6600', PArea:'Vejen' },
        { ActorType:'supplier', CustNo:'', SupNo:'370488', Nm:'Ravema', Ad1:'Margretelundsväg', Ad2:'', PNo:'331 34', PArea:'Värnamo' }
    ]);
    customerReadFails = true;
    const failedCustomersResponse = await fetch(base+'/mileage/visma-actors', { headers:{'x-test-user':'worker'} });
    assert.equal(failedCustomersResponse.status, 503);
    assert.match((await failedCustomersResponse.json()).error, /Kunde- og leverandøradresser fra Visma kunne ikke hentes/);
    customerReadFails = false;
    assert.equal((await fetch(base+'/mileage/rates', { method:'POST', headers:{ 'x-test-user':'worker','Content-Type':'application/json' }, body:JSON.stringify({year:2026}) })).status, 403);
    const preview = await fetch(base+'/mileage/import', { method:'POST', headers:{ 'x-test-user':'approver','Content-Type':'application/json' },
        body: JSON.stringify({preview:true,files:importFiles()}) });
    assert.equal(preview.status, 200);
    assert.equal((await preview.json()).counts.trips, 2);
    assert.equal((await gohData.getAppState('',{strict:true})), null);
    const sync = await fetch(base+'/mileage/employees/sync', { method:'POST', headers:{ 'x-test-user':'worker','Content-Type':'application/json' }, body:'{}' });
    assert.equal(sync.status, 200);
    const syncedData = await (await fetch(base+'/mileage/data', { headers:{ 'x-test-user':'worker' } })).json();
    assert.equal(syncedData.participants.length, 1);
    const create = await fetch(base+'/mileage/trips', { method:'POST', headers:{ 'x-test-user':'worker','Content-Type':'application/json' },
        body:JSON.stringify({ version:syncedData.version, employeeId:syncedData.participants[0].id, date:'2025-01-01', plate:'TEST',
            purpose:'Visit', stops:[{address:'Start'},{address:'End'}], km:5, manualReason:'Odometer' }) });
    assert.equal(create.status, 200);
    let tripData = await (await fetch(base+'/mileage/data', { headers:{ 'x-test-user':'worker' } })).json();
    const tripId = tripData.trips[0].id;
    const deleteBody = JSON.stringify({ version:tripData.version, note:'Duplicate' });
    assert.equal((await fetch(base+'/mileage/trips/'+tripId+'/delete', { method:'POST', headers:{ 'x-test-user':'worker','Content-Type':'application/json' }, body:deleteBody })).status, 403);
    assert.equal((await fetch(base+'/mileage/trips/'+tripId+'/delete', { method:'POST', headers:{ 'x-test-user':'approver','Content-Type':'application/json' }, body:deleteBody })).status, 200);
    assert.equal((await fetch(base+'/mileage/report.pdf?year=2025&employeeId='+encodeURIComponent(tripData.participants[0].id), { headers:{ 'x-test-user':'approver' } })).status, 404);
    const current = await (await fetch(base+'/mileage/data', { headers:{ 'x-test-user':'worker' } })).json();
    const addressBody = JSON.stringify({ version:current.version, name:'Example', address:'Test Street 1' });
    assert.equal((await fetch(base+'/mileage/addresses', { method:'POST', headers:{ 'Content-Type':'application/json' }, body:addressBody })).status, 403);
    assert.equal((await fetch(base+'/mileage/addresses', { method:'POST', headers:{ 'x-test-user':'worker','Content-Type':'application/json' }, body:addressBody })).status, 403);
    assert.equal((await fetch(base+'/mileage/addresses', { method:'POST', headers:{ 'x-test-user':'approver','Content-Type':'application/json' }, body:addressBody })).status, 200);
    const otherData = await (await fetch(base+'/mileage/data', { headers:{ 'x-test-user':'other' } })).json();
    assert.equal(otherData.addresses[0].name, 'Example');
    const purposeBody = JSON.stringify({ version:otherData.version, purpose:'Customer visit' });
    assert.equal((await fetch(base+'/mileage/purposes', { method:'POST', headers:{ 'x-test-user':'worker','Content-Type':'application/json' }, body:purposeBody })).status, 403);
    assert.equal((await fetch(base+'/mileage/purposes', { method:'POST', headers:{ 'x-test-user':'approver','Content-Type':'application/json' }, body:purposeBody })).status, 200);
    const purposeData = await (await fetch(base+'/mileage/data', { headers:{ 'x-test-user':'worker' } })).json();
    assert.deepEqual(purposeData.purposes, ['Customer visit']);
    assert.equal((await fetch(base+'/mileage/addresses/delete', { method:'POST', headers:{ 'x-test-user':'worker','Content-Type':'application/json' },
        body:JSON.stringify({ version:purposeData.version, name:'Example', address:'Test Street 1' }) })).status, 403);
    const headers = { 'x-test-user':'admin-test', 'Content-Type':'application/json' };
    const denied = await fetch(base+'/mileage/routing-settings', { method:'POST', headers:{ ...headers, 'x-test-user':'approver' },
        body:JSON.stringify({ version:0, action:'save', apiKey:'database-test-key' }) });
    assert.equal(denied.status, 403);
    assert.equal((await fetch(base+'/mileage/routing-settings', { headers: { 'x-test-user':'worker' } })).status, 403);
    assert.equal((await fetch(base+'/mileage/routing-settings')).status, 403);
    const saved = await fetch(base+'/mileage/routing-settings', { method:'POST', headers,
        body:JSON.stringify({ version:0, action:'save', apiKey:'database-test-key' }) });
    assert.equal(saved.status, 200);
    assert.doesNotMatch(await saved.text(), /database-test-key/);
    const status = await fetch(base+'/mileage/routing-settings', { headers });
    assert.equal(status.headers.get('cache-control'), 'no-store');
    assert.doesNotMatch(await status.text(), /database-test-key/);
    const workerData = await (await fetch(base+'/mileage/data', { headers: { 'x-test-user':'worker' } })).text();
    assert.doesNotMatch(workerData, /database-test-key/);
    assert.equal(JSON.parse(workerData).routingConfigured, true);
});
