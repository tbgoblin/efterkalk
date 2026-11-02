const test = require('node:test');
const assert = require('node:assert/strict');
const report = require('../assets/js/reconciliation-report');

function data(overrides = {}) {
    return {
        month: '2026-10', plateMode: 'fifo',
        metrics: { orders: 317, material: 2057612.51, material_only: 1500000, stang: 200000, parts: 300000, subcontract: 57612.51 },
        rows: [
            { id: 'booked', name: 'Bogført (Vareforbrug)', type: 'accounts', display: 'row', unit: 'dkk', sign: 1, value: 2400412.63, accounts: [
                { AcNo: 12070, Nm: 'Varekøb', Amount: 2319312.41 }, { AcNo: 12092, Nm: 'Salg jern/skrot', Amount: -50000 }, { AcNo: 12067, Nm: 'Paller og emballage', Amount: 131100.22 }] },
            { id: 'material', name: 'Beregnet (materiale)', type: 'source', source: 'material', display: 'row', unit: 'dkk', value: 2057612.51 },
            { id: 'difference', name: 'Forskel', type: 'formula', formula: '[booked] - [material]', display: 'row', unit: 'dkk', value: 342800.12 },
            { id: 'explanation', name: 'Forklaret af', type: 'heading', display: 'row' },
            { id: 'warehouse', name: 'Δ Varelager', type: 'source', source: 'warehouse_delta', display: 'row', unit: 'dkk', value: 335914.02 },
            { id: 'gap', name: 'Ikke fuldt faktureret', type: 'source', source: 'purchase_gap', display: 'row', unit: 'dkk', value: 20176.20 },
            { id: 'residual', name: 'Uforklaret rest', type: 'formula', formula: '[difference] - [warehouse] + [gap]', display: 'row', unit: 'dkk', value: 27062.30 }
        ],
        details: {
            purchaseGap: { missingValue: 15000, missingOrderCount: 3, partialValue: 5176.2, partialOrderCount: 1, totalValue: 20176.2 },
            varelager: {
                month: '2026-10', previousMonth: '2026-09', plateMode: 'fifo',
                current: { month: '2026-10', createdAt: '2026-11-01T19:51:31Z', valuationMode: 'live-at-save' },
                previous: { month: '2026-09', createdAt: '2026-10-01T03:55:35Z', valuationMode: 'live-at-save' },
                categories: [
                    { key: 'plates', previous: 3315305, current: 3033019, delta: -282286, movers: [{ key: '30101200355485', descr: '12 mm S355', previous: 284902, current: 98715, delta: -186187 }] },
                    { key: 'opfolgningvare', previous: 609721, current: 1197764, delta: 588043, movers: [{ key: '1005401474', descr: 'Rail <storage>', previous: 51315, current: 179959, delta: 128644 }] }
                ],
                workInProgress: [{ key: 'salgordreVia', previous: 890425, current: 1104007, delta: 213582 },
                    { key: 'finishedNotInvoiced', previous: 127188, current: 142398, delta: 15210 }]
            }
        },
        ubOrders: [{ OrdNo: 99, CustomerName: 'Kunde A/S', text: 'Prøve', material: 1234.5 }],
        ...overrides
    };
}

test('the report shows the result, every booked account and the formulas with row names', () => {
    const html = report.buildHtml(data());
    assert.match(html, /Afstemning af Vareforbrug · oktober 2026/);
    assert.match(html, /= \[Forskel\] - \[Δ Varelager\] \+ \[Ikke fuldt faktureret\]/);
    assert.match(html, /12092<\/td><td>Salg jern\/skrot/);
    assert.match(html, /<tr class="flag"><td>12067/);
    assert.doesNotMatch(html, /<tr class="flag"><td>12070/);
    assert.match(html, /Rail &lt;storage&gt;/);
    assert.match(html, /Kunde A\/S/);
    assert.doesNotMatch(html, /Nogle tal var ikke færdigberegnet/);
});

test('explanations quantify non-material accounts and finished, not invoiced orders but not VIA', () => {
    const html = report.buildHtml(data());
    // Flagged accounts: -50.000 + 131.100,22 = 81.100,22, so the residual would be 27.062,30 - 81.100,22.
    assert.match(html, /Uden dem ville resten være -54\.037,92 DKK i stedet for 27\.062,30 DKK/);
    assert.match(html, /Deres kost er ændret med \+15\.210/);
    assert.match(html, /Går i samme retning som resten/);
    assert.doesNotMatch(html, /\+213\.582/);
    assert.doesNotMatch(html, /dage efter månedens udgang/);
});

test('a late closing, missing values and a missing Varelager closing are called out', () => {
    const late = data();
    late.details.varelager.current.createdAt = '2026-11-06T09:00:00Z';
    late.rows[4].value = null;
    let html = report.buildHtml(late);
    assert.match(html, /Lagerlisten for oktober 2026 er gemt 5 dage efter månedens udgang/);
    assert.match(html, /Nogle tal var ikke færdigberegnet/);
    const empty = data();
    empty.details.varelager.categories = [];
    empty.ubOrders = [];
    html = report.buildHtml(empty);
    assert.match(html, /mangler en lagerliste-lukning/);
    assert.match(html, /Ingen U\/B-ordrer registreret/);
    assert.match(html, /bør de registreres som U\/B/);
});

test('measured finished-order material gets its own section, VIA only for information, and the adjusted residual', () => {
    const measured = data();
    measured.details.wip = { month: '2026-10', previousMonth: '2026-09', delta: 6418, viaDelta: 181452,
        previous: { via: 539026, viaOrders: 122, finished: 47320, finishedOrders: 13, missing: [] },
        current: { via: 720478, viaOrders: 85, finished: 53738, finishedOrders: 12, missing: [] } };
    let html = report.buildHtml(measured);
    assert.match(html, /7\. Færdige, afsendte ordrer, der endnu ikke er faktureret/);
    assert.match(html, /47\.320 \(13\).*53\.738 \(12\).*\+6\.418/);
    assert.match(html, /Til orientering: salgsordrer i arbejde \(VIA\), ikke medregnet/);
    assert.match(html, /Medregnet ville resten være 20\.644,30 DKK/);
    measured.rows.splice(6, 0, { id: 'finished', name: 'Δ Færdige, ikke fakturerede ordrer', type: 'source', source: 'finished_material', display: 'row', unit: 'dkk', value: 6418 });
    html = report.buildHtml(measured);
    assert.doesNotMatch(html, /Medregnet ville resten være/);
});
