const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { createBilancioArchiveService } = require('../services/bilancioArchiveService');
const { generateBilancioPdf } = require('../services/bilancioPdf');
const { createDefinitionStore } = require('../services/bilancioDefinition');

function storage() {
    const records = new Map();
    return {
        records,
        async getAppState(key) { return records.has(key) ? { payload: structuredClone(records.get(key)) } : null; },
        async getAppStateKeysByPrefix(prefix) { return [...records.keys()].filter(key => key.startsWith(prefix)).map(key => ({ key })); },
        async setAppState(key, value, options = {}) {
            if (options.createOnly && records.has(key)) return false;
            if (options.expectedVersion != null && (records.get(key)?.version || 0) !== options.expectedVersion) return false;
            records.set(key, structuredClone(value)); return true;
        }
    };
}
function bundle(amount = 100) {
    const row = (n, type = 'subtotal') => ({
        id: 'result', type, name: 'Årets resultat', amounts: Array(n).fill(amount), percentages: Array(n).fill(100), accounts: []
    });
    const columns = Array.from({ length: 12 }, (_, i) => ({ year: 2026, period: i + 1 }));
    return {
        scope: 'scope', definition: { formula: '1' }, configurationRevision: { number: 1, id: 'a'.repeat(64) },
        report: { year: 2026, period: 4, reportId: 'default', reportName: 'Økonomirapport', generatedAt: '2026-11-11T10:00:00Z',
            rows: [row(4)], assets: { title: 'Balance', rows: [row(2)] }, revenueAmounts: Array(4).fill(amount), periodOnly: [] },
        monthly: { columns, rows: [row(12)] }, yearly: { columns: columns.map((_, i) => ({ year: 2015 + i, period: 12 })), rows: [row(12)] }
    };
}
function fixture() {
    const gohData = storage();
    let activeScope = 'scope', current = bundle(), time = Date.parse('2026-11-11T10:00:00Z');
    let calculations = 0;
    const args = {
        gohData, definitionStore: { scope: () => activeScope },
        bilancioService: { async bundle(year, period, reportId, template) {
            calculations++;
            const result = structuredClone(current);
            result.report.year = year; result.report.period = period; result.report.reportId = reportId;
            if (template) {
                result.definition = structuredClone(template.definition);
                result.configurationRevision = structuredClone(template.configurationRevision);
            }
            return result;
        } },
        generatePdf: generateBilancioPdf, appVersion: '1.1.95', now: () => time
    };
    const service = createBilancioArchiveService(args);
    return {
        service, gohData, args, calculations: () => calculations,
        change: () => { current = bundle(500); current.configurationRevision = { number: 2, id: 'b'.repeat(64) }; },
        switchScope: () => { activeScope = 'other'; },
        expire: () => { time += 3 * 60 * 60 * 1000; }
    };
}

test('saved reports keep all values, formulas and original PDF across changes and service restarts', async () => {
    const f = fixture();
    const draft = await f.service.draft(2026, 4, 'default', 'alice');
    draft.report.rows[0].amounts[0] = 9999;
    const first = await f.service.save({ draftId: draft.draftId, unit: 1000, note: '' }, 'alice');
    assert.equal(first.report.rows[0].amounts[0], 100);
    assert.equal(first.revision, 1);
    f.change();
    const next = await f.service.draft(2026, 4, 'default', 'alice');
    assert.equal(next.baseRevision, 1);
    await assert.rejects(f.service.save({ draftId: next.draftId, unit: 1000, note: '' }, 'alice'), e => e.statusCode === 400);
    const second = await f.service.save({ draftId: next.draftId, unit: 1000, note: 'Rettelse' }, 'alice');
    assert.equal(second.revision, 2);
    assert.equal(second.configurationRevision.number, 1);
    assert.equal(second.report.rows[0].amounts[0], 500);
    const reopened = await createBilancioArchiveService(f.args).load('default', first.id);
    assert.equal(reopened.pdfBase64, first.pdfBase64);
    assert.equal(reopened.report.rows[0].amounts[0], 100);
    assert.equal(reopened.definition.formula, '1');
    assert.equal(f.calculations(), 2);
    assert.equal(crypto.createHash('sha256').update(Buffer.from(first.pdfBase64, 'base64')).digest('hex'), first.pdfSha256);
    assert.deepEqual((await f.service.list('default')).map(row => row.revision), [2, 1]);
    assert.equal((await f.service.save({ draftId: draft.draftId, unit: 1, note: '' }, 'alice')).id, first.id);
});

test('archive refuses expired drafts, wrong users, different profiles and cross-report reads', async () => {
    const f = fixture();
    const draft = await f.service.draft(2026, 4, 'default', 'alice');
    await assert.rejects(f.service.save({ draftId: draft.draftId, unit: 1, note: '' }, 'bob'), e => e.statusCode === 403);
    const saved = await f.service.save({ draftId: draft.draftId, unit: 1, note: '' }, 'alice');
    await assert.rejects(f.service.load('other', saved.id), e => e.statusCode === 400);
    f.expire();
    await assert.rejects(f.service.save({ draftId: draft.draftId, unit: 1, note: '' }, 'alice'), e => e.statusCode === 410);
    f.switchScope();
    await assert.rejects(f.service.load('default', saved.id), e => e.statusCode === 400);
});

test('concurrent report revisions cannot overwrite each other; unavailable storage fails explicitly', async () => {
    const f = fixture();
    const [a, b] = await Promise.all(['alice', 'bob'].map(user => f.service.draft(2026, 4, 'default', user)));
    const outcomes = await Promise.allSettled([
        f.service.save({ draftId: a.draftId, unit: 1, note: '' }, 'alice'),
        f.service.save({ draftId: b.draftId, unit: 1, note: '' }, 'bob')
    ]);
    assert.equal(outcomes.filter(result => result.status === 'fulfilled').length, 1);
    assert.equal(outcomes.find(result => result.status === 'rejected').reason.statusCode, 409);
    assert.equal((await f.service.list('default')).length, 1);
    f.gohData.getAppStateKeysByPrefix = async () => null;
    await assert.rejects(f.service.list('default'), e => e.statusCode === 503);
});

test('failed PDF generation or failed persistence never reports a saved revision', async () => {
    const f = fixture();
    const pdfFailure = createBilancioArchiveService({ ...f.args, generatePdf: async () => { throw Error('PDF generation failed'); } });
    const draft = await pdfFailure.draft(2026, 4, 'default', 'alice');
    await assert.rejects(pdfFailure.save({ draftId: draft.draftId, unit: 1, note: '' }, 'alice'), /PDF generation failed/);
    assert.equal([...f.gohData.records.keys()].filter(key => key.startsWith('bilancio_snapshot_')).length, 0);
    const other = await f.service.draft(2026, 4, 'default', 'alice');
    f.gohData.setAppState = async () => false;
    await assert.rejects(f.service.save({ draftId: other.draftId, unit: 1, note: '' }, 'alice'), e => e.statusCode === 503);
    assert.equal([...f.gohData.records.keys()].filter(key => key.startsWith('bilancio_snapshot_')).length, 0);
});

test('first read pins only monthly template across restarts; amounts refresh and unopened months get new template', async () => {
    const f = fixture();
    const first = await f.service.draft(2026, 4, 'default', 'alice');
    f.change();
    const restarted = createBilancioArchiveService(f.args);
    const sameMonth = await restarted.draft(2026, 4, 'default', 'alice');
    assert.equal(sameMonth.configurationRevision.id, first.configurationRevision.id);
    assert.equal(sameMonth.report.rows[0].amounts[0], 500);
    assert.equal(sameMonth.templateLockedAt, first.templateLockedAt);
    const nextMonth = await restarted.draft(2026, 5, 'default', 'alice');
    assert.equal(nextMonth.configurationRevision.number, 2);
    const otherReport = await restarted.draft(2026, 4, 'other', 'alice');
    assert.equal(otherReport.configurationRevision.number, 2);
});

test('PDF contains actual mixed A4 page sizes, not just CSS orientation hints', async () => {
    const pdf = await generateBilancioPdf({ ...bundle(), unit: 1000, revision: 1, savedBy: 'alice', savedAt: '2026-11-11T10:00:00Z', appVersion: '1.1.95' });
    assert.equal(pdf.subarray(0, 5).toString(), '%PDF-');
    const boxes = [...pdf.toString('latin1').matchAll(/\/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/g)];
    assert.equal(boxes.length, 4);
    assert.deepEqual(boxes.map(match => Number(match[1]) > Number(match[2])), [true, false, true, true]);
});

test('long PDF tables paginate without dropping rows and retain their page orientation', async () => {
    const data = bundle();
    data.monthly.rows = Array.from({ length: 100 }, (_, i) => ({ ...data.monthly.rows[0], name: 'Kontogruppe ' + i }));
    const pdf = await generateBilancioPdf({ ...data, unit: 1, revision: 1, savedBy: 'alice', savedAt: '2026-11-11T10:00:00Z', appVersion: '1.1.95' });
    const boxes = [...pdf.toString('latin1').matchAll(/\/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/g)];
    assert.ok(boxes.length > 4);
    assert.equal(boxes.filter(match => Number(match[1]) < Number(match[2])).length, 1);
    // PDFKit writes text as hexadecimal strings, one for each cell.
    const zlib = require('node:zlib');
    const source = pdf.toString('latin1');
    let text = '';
    for (const match of source.matchAll(/stream\n([\s\S]*?)\nendstream/g)) {
        const stream = zlib.inflateSync(Buffer.from(match[1], 'latin1')).toString('latin1');
        text += [...stream.matchAll(/<([a-f0-9]+)>/gi)].map(m => Buffer.from(m[1], 'hex').toString('latin1')).join('');
    }
    assert.ok(text.includes('Kontogruppe 0'));
    assert.ok(text.includes('Kontogruppe 99'));
});

test('configuration revisions are immutable copies; stale configuration saves still conflict', async () => {
    const gohData = storage();
    const config = { schema: 1, version: 0, balanceTitle: 'Balance', revenueRow: 'a', beforeTaxRow: 'a',
        pnl: [{ id: 'a', name: 'Resultat', type: 'formula', formula: '1' }], balance: [{ id: 'h', name: 'Balance', type: 'heading' }] };
    const store = createDefinitionStore({ gohData, getProfile: () => ({ server: 's', database: 'd' }), defaultDefinition: config });
    const original = await store.load();
    const first = await store.save(original, { accounts: [], groups: [] }, 'alice');
    const firstRevision = await store.freeze(first);
    const archived = structuredClone(gohData.records.get('bilancio_revision_' + firstRevision.id));
    const second = await store.save({ ...first, pnl: [{ ...first.pnl[0], formula: '2' }] }, { accounts: [], groups: [] }, 'alice');
    assert.notEqual(first.revision.id, second.revision.id);
    assert.deepEqual(gohData.records.get('bilancio_revision_' + firstRevision.id), archived);
    await assert.rejects(store.save(first, { accounts: [], groups: [] }, 'bob'), e => e.statusCode === 409);
    assert.equal((await store.load()).pnl[0].formula, '2');
});
