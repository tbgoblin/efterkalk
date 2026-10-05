const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createLagerlisteClient(options = {}) {
    const select = value => ({ value, options: [{ value: '' }],
        set innerHTML(html) {
            this.html = html;
            this.options = Array.from(html.matchAll(/<option value="([^"]*)"/g), match => ({ value: match[1] }));
            this.value = '';
        }, get innerHTML() { return this.html || ''; } });
    const elements = {
        lagerlisteCompareA: select(options.periodA || 'month:2026-08'),
        lagerlisteCompareB: select(options.periodB || ''),
        lagerlistePeriodStatus: { textContent: '' },
        lagerlisteMonth: { value: '2026-09' },
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
            if (options.fetch) return options.fetch(url);
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
            refresh: refreshLagerlisteCompareOptions,
            refreshPeriods: refreshLagerlistePeriods,
            save: saveLagerlisteSnapshot,
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

test('monthly periods load even while daily snapshot history is stalled', async () => {
    const client = createLagerlisteClient({ fetch: async url => {
        if (url === '/lagerliste/snapshots') return new Promise(() => {});
        assert.equal(url, '/lagerliste/snapshot-months');
        return { ok: true, json: async () => ({ ok: true, months: ['2026-09', '2026-08'] }) };
    } });
    client.api.refreshPeriods();
    await new Promise(resolve => setImmediate(resolve));
    for (const id of ['lagerlisteCompareA', 'lagerlisteCompareB']) {
        assert.match(client.elements[id].innerHTML, /month:2026-09/);
        assert.match(client.elements[id].innerHTML, /Aktuel \(live\)/);
    }
});

test('failed period loading leaves live selectable and shows a retry message', async () => {
    const client = createLagerlisteClient({ fetch: async () => { throw Error('offline'); } });
    await client.api.refresh();
    assert.match(client.elements.lagerlisteCompareA.innerHTML, /current/);
    assert.match(client.elements.lagerlistePeriodStatus.textContent, /offline.*Opdater perioder/);
});

test('saving a month displays its Diverse instead of leaving the current month on screen', async () => {
    const current = { generatedAt: new Date().toISOString(),
        diverseStatus: { month: '2026-09', complete: true },
        categories: { diverse: [{ mode: 'amount', Value: 738746.2 }, { mode: 'visma', Value: 241298.76 }] },
        totals: { diverse: 980044.96 } };
    const client = createLagerlisteClient({ fetch: async url => ({ ok: true,
        json: async () => url.includes('snapshot-months') ? { ok: true, months: ['2026-09'] } : { ok: true, current }
    }) });
    await client.api.save();
    assert.match(client.elements.lagerlisteResults.innerHTML, /Periode: Måned 2026-09/);
    assert.match(client.elements.lagerlisteResults.innerHTML, /Diverse · måned 2026-09: Administration 738\.746,20/);
    assert.match(client.elements.lagerlisteResults.innerHTML, /Visma 44\/45\/46\/63 241\.298,76/);
    assert.equal(client.elements.lagerlisteCompareA.value, 'month:2026-09');
    assert.equal(client.elements.lagerlisteCompareB.value, '');
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
