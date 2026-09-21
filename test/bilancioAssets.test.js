const test = require('node:test');
const assert = require('node:assert/strict');
const { resolveAssetGroups, buildAssetRows } = require('../services/bilancioService');
const chart = [
    { AcGr: '47_Driftsmidler', AgAcGr: '47_Driftsmidler_ialt' },
    { AcGr: '47_Afskrivninger_ialt', AgAcGr: '47_Driftsmidler_ialt' },
    { AcGr: '47_Afskrivninger', AgAcGr: '47_Afskrivninger_ialt' },
    { AcGr: 'unrelated', AgAcGr: 'another' }
];
function records() {
    return [2026, 2025].flatMap(Yr => [
        { AcNo: 66980, AcGr: '66_Tilgodehavender', Balance: 100 },
        { AcNo: 66981, AcGr: '66_Tilgodehavender', Balance: 200 },
        { AcNo: 66100, AcGr: '66_Tilgodehavender', Balance: 900 },
        { AcNo: 66110, AcGr: '66_Tilgodehavender', Balance: -50 },
        { AcNo: 66200, AcGr: '66_Tilgodehavender', Balance: 700 },
        { AcNo: 61100, AcGr: '61', Balance: 100 },
        { AcNo: 61120, AcGr: '61', Balance: 200 },
        { AcNo: 61850, AcGr: '61', Balance: 300 },
        { AcNo: 61900, AcGr: '61', Balance: -50 },
        { AcNo: 46000, AcGr: '46_Grunde_og_bygninger', Balance: 1000 },
        { AcNo: 47000, AcGr: '47_Driftsmidler', Balance: 800 },
        { AcNo: 47900, AcGr: '47_Afskrivninger', Balance: -300 }
    ].map(row => ({ ...row, Yr, Balance: row.Balance * (Yr === 2026 ? 1 : 2) })));
}
test('nested groups include depreciation once and subtotal sums the three asset rows', () => {
    const groups = resolveAssetGroups(chart);
    assert.equal(groups.has('47_Afskrivninger'), true);
    assert.equal(groups.has('unrelated'), false);
    const rows = buildAssetRows(records(), 2026, groups);
    assert.deepEqual(rows.map(row => row.amounts), [[300,600],[1000,2000],[500,1000],[1800,3600],[550,1100],[900,1800]]);
    assert.equal(rows[3].type, 'subtotal');
});
test('receivables include only 66100 in both years', () => {
    const rows = buildAssetRows(records(),2026,resolveAssetGroups(chart));
    const receivables = rows.find(row => row.name === 'Tilgodehavender fra salg');
    assert.deepEqual(receivables.accounts,[66100]);
    assert.deepEqual(receivables.amounts,[900,1800]);
    const accounts = rows.filter(row => row.type !== 'subtotal').flatMap(row => row.accounts);
    assert.equal(new Set(accounts).size,accounts.length);
    assert.deepEqual(rows[0].amounts,[300,600]);
    assert.deepEqual(rows[3].amounts,[1800,3600]);
});
test('duplicate balances and overlaps fail instead of inflating the total', () => {
    const source = records();
    assert.throws(() => buildAssetRows([...source,source[0]],2026,resolveAssetGroups(chart)), /dubleret/);
    source[0].AcGr = '47_Driftsmidler';
    assert.throws(() => buildAssetRows(source,2026,resolveAssetGroups(chart)), /flere aktivposter/);
});
test('group cycles terminate without repeating any group', () => {
    const groups = resolveAssetGroups([...chart,{AcGr:'47_Driftsmidler_ialt',AgAcGr:'47_Afskrivninger'}]);
    assert.equal(groups.size,4);
});
