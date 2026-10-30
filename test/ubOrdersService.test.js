const test = require('node:test');
const assert = require('node:assert/strict');
const { createUbOrdersService } = require('../services/ubOrdersService');

function fixture(initial = {}) {
    const notes = structuredClone(initial);
    let reads = 0;
    const orders = [
        { OrdNo:1, OrdDt:20260801, LstInvDt:0, CustNo:42, CustomerName:'Kunde', InvoAm:0 },
        { OrdNo:2, OrdDt:20260901, LstInvDt:20261005, CustNo:43, InvoAm:100 }
    ];
    const service = createUbOrdersService({
        getConnection:async () => ({ request() {
            let ids;
            return { input(_name, _type, values) { ids=values.split(',').map(Number); return this; },
                async query() { reads++; return { recordset:orders.filter(order => ids.includes(order.OrdNo)) }; } };
        } }),
        sql:{ NVarChar:() => '', MAX:0 },
        orderNotes:{ getNote:id => notes[id] || null, getAllNotes:() => notes,
            setNote(id, note) { notes[id]=note; return note; } }
    });
    return { service, notes, reads:() => reads };
}

test('U/B can be registered without invoice or credit flag; month and general note are preserved', async () => {
    const f = fixture({ 1:{ status:'ok', text:'Eksisterende note', isCreditNote:false } });
    await f.service.add({ ordNo:1, month:'2026-10', text:'Prøve' });
    assert.equal(f.notes[1].isUB, true);
    assert.equal(f.notes[1].isCreditNote, false);
    assert.equal(f.notes[1].text, 'Eksisterende note');
    const october = await f.service.list('2026-10');
    assert.deepEqual(october.map(row => row.OrdNo), [1]);
    assert.equal(october[0].text, 'Prøve');
    assert.equal((await f.service.list('2026-09')).length, 0);
    await assert.rejects(() => f.service.remove({ ordNo:1, month:'2026-09' }), { statusCode:409 });
    assert.equal(f.notes[1].isUB, true);
    await f.service.remove({ ordNo:1, month:'2026-10' });
    assert.equal(f.notes[1].isUB, false);
    assert.equal(f.notes[1].text, 'Eksisterende note');
    assert.equal((await f.service.list('2026-10')).length, 0);
});

test('legacy U/B entries use invoice/order month; ordinary zero-valued orders are not automatically U/B', async () => {
    const f = fixture({ 2:{ isUB:true, isCreditNote:true, text:'Legacy' } });
    const rows = await f.service.list('2026-10');
    assert.deepEqual(rows.map(row => row.OrdNo), [2]);
    assert.equal(rows[0].text, 'Legacy');
    assert.equal((await f.service.list('2026-08')).length, 0);
});

test('invalid and missing orders do not create registrations or read invalid query parameters', async () => {
    const f = fixture();
    await assert.rejects(() => f.service.add({ ordNo:'1 OR 1=1', month:'2026-10' }), { statusCode:400 });
    await assert.rejects(() => f.service.add({ ordNo:1, month:'2026-13' }), { statusCode:400 });
    assert.equal(f.reads(), 0);
    await assert.rejects(() => f.service.add({ ordNo:999, month:'2026-10' }), { statusCode:404 });
    assert.deepEqual(f.notes, {});
});

test('repeated registration updates its note without duplication and cannot silently move another month', async () => {
    const f = fixture();
    await f.service.add({ ordNo:1, month:'2026-10', text:'Første note' });
    await f.service.add({ ordNo:1, month:'2026-10', text:'Rettet note' });
    assert.equal((await f.service.list('2026-10')).length, 1);
    assert.equal(f.notes[1].ubText, 'Rettet note');
    await assert.rejects(() => f.service.add({ ordNo:1, month:'2026-11' }), { statusCode:409 });
    assert.equal(f.notes[1].ubMonth, '2026-10');
});

test('order notes persist a standalone U/B flag and preserve its month when the general note is edited', () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const vm = require('node:vm');
    const source = fs.readFileSync(path.join(__dirname, '../services/orderNotesService.js'), 'utf8');
    let stored;
    const fakeFs = { mkdirSync() {}, existsSync:() => stored !== undefined,
        readFileSync:() => stored, writeFileSync(_file, value) { stored=value; }, copyFileSync() {} };
    function loadService() {
        const context = { module:{ exports:{} }, __dirname, process:{ env:{} }, console,
            require(name) {
                if (name === 'fs') return fakeFs;
                if (name === 'path') return path;
                if (name === 'process') return { env:{} };
                return { setAppState:async () => {} };
            } };
        vm.runInNewContext(source, context);
        return context.module.exports;
    }
    let service = loadService();
    service.setNote(1, { isUB:true, isCreditNote:false, ubMonth:'2026-10', ubText:'Prøve' });
    service = loadService();
    assert.equal(service.getNote(1).isUB, true);
    assert.equal(service.getNote(1).isCreditNote, false);
    service.setNote(1, { text:'Generel note', isUB:true });
    assert.equal(service.getNote(1).ubMonth, '2026-10');
    assert.equal(service.getNote(1).ubText, 'Prøve');
    service.setNote(1, { text:'Generel note', isUB:false });
    assert.equal(service.getNote(1).ubMonth, '');
    assert.equal(service.getNote(1).ubText, '');
});
