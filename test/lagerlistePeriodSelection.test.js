const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createLagerlisteClient(options = {}) {
    const elements = {
        lagerlisteCompareA: { value: options.periodA || 'month:2026-08' },
        lagerlisteCompareB: { value: options.periodB || '' },
        lagerlisteCompareResults: { innerHTML: '', querySelectorAll: () => [] },
        lagerlisteResults: { innerHTML: '', querySelectorAll: () => [] },
        lagerlisteSnapshotStatus: { textContent: '' }
    };
    const fetchCalls = [];
    const historical = options.historical || {
        ok: true,
        generatedAt: '2026-08-31T23:59:00.000Z',
        totals: {},
        categories: {}
    };
    const context = vm.createContext({
        authToken: 'test-token',
        document: {
            getElementById: id => elements[id] || null,
            querySelector: () => null,
            querySelectorAll: () => []
        },
        fetch: async url => {
            fetchCalls.push(String(url));
            if (String(url).includes('/lagerliste/snapshot/2026-08')) {
                return { ok: true, status: 200, json: async () => ({ ok: true, current: historical }) };
            }
            throw new Error('Unexpected fetch: ' + url);
        },
        alert: () => {},
        confirm: () => true,
        console,
        setTimeout,
        clearTimeout
    });
    const source = fs.readFileSync(path.join(__dirname, '..', 'assets', 'js', 'lagerliste.js'), 'utf8');
    vm.runInContext(source + `
        ;globalThis.__lagerlisteTest = {
            compare: lagerlisteComparePeriods,
            resolve: lagerlisteResolvePeriod,
            setLive(payload) {
                lagerlisteLiveCurrent = payload;
                lagerlisteRender(payload, null, 'Aktuel');
            }
        };
    `, context, { filename: 'lagerliste.js' });
    return { api: context.__lagerlisteTest, elements, fetchCalls };
}

test('period A opens a single monthly closing when period B is empty', async () => {
    const client = createLagerlisteClient();

    await client.api.compare();

    assert.match(client.elements.lagerlisteResults.innerHTML, /Periode: Måned 2026-08/);
    assert.equal(client.elements.lagerlisteCompareResults.innerHTML, '');
    assert.match(client.elements.lagerlisteSnapshotStatus.textContent, /PDF udskriver denne periode/);
    assert.deepEqual(client.fetchCalls, ['/lagerliste/snapshot/2026-08']);
});

test('opening a historical month does not replace the live comparison dataset', async () => {
    const client = createLagerlisteClient();
    const live = { ok: true, generatedAt: '2026-09-08T08:00:00.000Z', totals: {}, categories: {}, marker: 'live' };
    client.api.setLive(live);

    await client.api.compare();
    const resolvedCurrent = await client.api.resolve('current');

    assert.equal(resolvedCurrent.payload.marker, 'live');
    assert.equal(client.fetchCalls.filter(url => url.includes('/lagerliste/current')).length, 0);
});
