const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function fixture() {
    const nodes = new Map(), calls = [];
    let saved = null, saveError = null;
    const row = { name: 'Resultat', type: 'subtotal', amounts: [100, 90, 300, 290], percentages: [100, 100, 100, 100], accounts: [] };
    const report = { reportId: 'default', reportName: 'Økonomirapport', year: 2026, period: 4, generatedAt: '2026-11-11T10:00:00Z',
        rows: [row], assets: { title: 'Balance', rows: [{ ...row, amounts: [10, 9] }] }, revenueAmounts: row.amounts };
    const bundle = { report, monthly: { columns: [{ year: 2026, period: 4 }], rows: [{ ...row, amounts: [100] }] },
        yearly: { columns: [{ year: 2026, period: 4 }], rows: [{ ...row, amounts: [300] }] },
        draftId: 'draft', baseRevision: 0, configurationRevision: { number: 7, id: 'a'.repeat(64) }, appVersion: '1.1.95' };
    function node(id) {
        if (!nodes.has(id)) nodes.set(id, {
            value: id === 'unit' ? '1000' : id === 'reportSelect' ? 'default' : '',
            disabled: false, textContent: '', innerHTML: '', handlers: {},
            addEventListener(event, handler) { this.handlers[event] = handler; },
            add() {}, requestSubmit() {}, querySelectorAll() { return []; }
        });
        return nodes.get(id);
    }
    const context = vm.createContext({
        document: { getElementById: node },
        window: { addEventListener() {} }, location: { search: '' },
        Intl, Date, URLSearchParams, Option: function() {}, confirm: () => true,
        fetch: async (url, options = {}) => {
            calls.push([url, options]);
            let data, status = 200;
            if (url === '/bilancio-trial/reports') data = { reports: [{ reportId: 'default', reportName: 'Økonomirapport' }] };
            else if (url.startsWith('/bilancio-trial/draft?')) data = structuredClone(bundle);
            else if (url === '/bilancio-trial/archive' && options.method === 'POST') {
                if (saveError) { status = 409; data = { error: saveError }; }
                else {
                    saved = { ...structuredClone(bundle), id: 'snapshot', revision: 1, savedBy: 'alice', savedAt: '2026-11-11T10:00:00Z',
                        note: 'Kontrolleret', unit: 1000 };
                    data = structuredClone(saved);
                }
            } else if (url.startsWith('/bilancio-trial/archive?')) data = { reports: saved ? [{ ...saved, ...saved.report }] : [] };
            else if (url.startsWith('/bilancio-trial/archive/snapshot?')) data = structuredClone(saved);
            else throw Error('Unexpected request: ' + url);
            return { ok: status === 200, status, json: async () => data };
        }
    });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../assets/js/bilancio-trial.js'), 'utf8'), context);
    return {
        node, calls, bundle,
        flush: () => new Promise(resolve => setImmediate(resolve)),
        submit: () => node('filters').handlers.submit({ preventDefault() {} }),
        click: id => node(id).handlers.click(),
        failSave: () => { saveError = 'En ny revision er allerede gemt.'; }
    };
}
test('client saves server draft token only, opens frozen history without new calculation, and marks revisions', async () => {
    const f = fixture();
    await f.flush();
    await f.submit();
    assert.match(f.node('revisionInfo').textContent, /KLADDE.*Konfiguration Rev\. 7/);
    assert.equal(f.node('saveReportBtn').disabled, false);
    assert.equal(f.node('downloadPdfBtn').disabled, true);
    await f.click('saveReportBtn');
    const body = JSON.parse(f.calls.find(([url, opts]) => url === '/bilancio-trial/archive' && opts.method === 'POST')[1].body);
    assert.deepEqual(body, { draftId: 'draft', unit: 1000, note: '' });
    assert.equal(f.node('saveReportBtn').disabled, true);
    assert.equal(f.node('downloadPdfBtn').disabled, false);
    assert.match(f.node('revisionInfo').textContent, /GEMT.*Rapport Rev\. 1/);
    assert.equal(f.node('unit').disabled, true);
    f.bundle.report.rows[0].amounts[0] = 99999;
    f.node('archiveSelect').value = 'snapshot';
    const calculations = f.calls.filter(([url]) => url.startsWith('/bilancio-trial/draft?')).length;
    await f.click('openArchiveBtn');
    assert.equal(f.calls.filter(([url]) => url.startsWith('/bilancio-trial/draft?')).length, calculations);
    assert.doesNotMatch(f.node('report').innerHTML, /99\.999/);
    assert.match(f.node('status').textContent, /ingen genberegning/);
});
test('client reports failed saves explicitly and leaves draft retryable', async () => {
    const f = fixture();
    await f.flush();
    await f.submit();
    f.failSave();
    await f.click('saveReportBtn');
    assert.match(f.node('error').textContent, /En ny revision/);
    assert.equal(f.node('saveReportBtn').disabled, false);
    assert.equal(f.node('downloadPdfBtn').disabled, true);
});
