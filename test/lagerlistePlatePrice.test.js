const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../assets/js/lagerliste.js'), 'utf8');
const context = vm.createContext({});
vm.runInContext(source, context);
const payload = {
    totals: { plates: 100, stang: 20 },
    categories: { plates: [{ Value: 100, FifoValue: 75 }], plateGroups: [{ FifoValue: 75 }] }
};

test('Pladelager selects standard or FIFO without double-counting groups or changing snapshots', () => {
    const original = JSON.stringify(payload);
    assert.equal(context.lagerlistePlateTotals(payload, 'standard').plates, 100);
    assert.equal(context.lagerlistePlateTotals(payload, 'fifo').plates, 75);
    assert.equal(context.lagerlistePlateTotals(payload, 'fifo').stang, 20);
    assert.equal(JSON.stringify(payload), original);
    assert.equal(context.lagerlistePlateTotals(payload, 'standard').plates, 100);
});

test('FIFO supports group-only snapshots and zero values but refuses missing costs', () => {
    assert.equal(context.lagerlistePlateTotals({ categories: { plateGroups: [{ FifoValue: 0 }] } }, 'fifo').plates, 0);
    assert.throws(() => context.lagerlistePlateTotals({ categories: { plates: [{}] } }, 'fifo'), /FIFO/);
    assert.throws(() => context.lagerlistePlateTotals({}, 'fifo'), /FIFO/);
});

test('FIFO updates warehouse/grand totals and uses same basis for previous period', () => {
    vm.runInContext("lagerlistePlatePriceMode = 'fifo'", context);
    const html = context.lagerlisteSummaryTable({
        generatedAt: 'today', totals: context.lagerlistePlateTotals(payload), categories: payload.categories,
        comparison: { totals: { plates: 200 }, categories: { plates: [{ FifoValue: 60 }] } }
    });
    assert.match(html, /75,00 DKK/);
    assert.match(html, /95,00 DKK/);
    assert.match(html, /60,00 DKK/);
    assert.match(html, /StcBal.PhCstPr/);
    assert.doesNotMatch(html, /200,00 DKK/);
});
