const test = require('node:test');
const assert = require('node:assert/strict');
const { createBilancioService } = require('../services/bilancioService');

function fixture(amount) {
    const calls = [];
    const pool = { request: () => ({
        inputs: {},
        input(name, _type, value) { this.inputs[name] = value; return this; },
        async query() { calls.push(this.inputs); return { recordset: [{ Amount: amount }] }; }
    }) };
    const service = createBilancioService({ getConnection: async () => pool, sql: { Int: 'Int', NVarChar: () => 'NVarChar' } });
    return { service, calls };
}

test('Vareforbrug for a month is returned as a positive magnitude, not re-negated', async () => {
    // The raw AcTr sum for Vareforbrug accounts is already positive; only the bilancio P&L's
    // own sign:-1 flips it negative for its own display. This lookup must not repeat that flip.
    const { service } = fixture(1674478.21);
    assert.equal(await service.vareforbrugForMonth('2026-08'), 1674478.21);
});

test('calendar month maps to the correct fiscal year/period (fiscal year starts in July)', async () => {
    const { service, calls } = fixture(0);
    const purchaseAccount = 12070;
    await service.vareforbrugForMonth('2026-08'); // August 2026 -> fiscal year 2026, period 2
    assert.deepEqual(calls[0], { year: 2026, period: 2, purchaseAccount });
    await service.vareforbrugForMonth('2026-03'); // March 2026 -> fiscal year 2025, period 9
    assert.deepEqual(calls[1], { year: 2025, period: 9, purchaseAccount });
    await service.vareforbrugForMonth('2026-07'); // July 2026 -> fiscal year 2026, period 1
    assert.deepEqual(calls[2], { year: 2026, period: 1, purchaseAccount });
});

test('an invalid or missing month is rejected', async () => {
    const { service } = fixture(0);
    await assert.rejects(service.vareforbrugForMonth('2026-13'));
    await assert.rejects(service.vareforbrugForMonth(''));
    await assert.rejects(service.vareforbrugForMonth(undefined));
});

test('only account 12070 (Varekøb) is queried, not the whole 12_Vareforbrug group', async () => {
    // The group 12_Vareforbrug bundles accounts with unrelated meaning (12085 Intern køb, 12092
    // Salg jern/skrot, ...) alongside 12070 Varekøb, the actual purchase account. Filtering by
    // account keeps "Bogført" a clean, directly comparable purchase figure without a separate
    // "non-purchase" breakout — the user asked for this simplification over the group-wide total.
    const { service, calls } = fixture(2085024.31);
    const result = await service.vareforbrugForMonth('2026-09');
    assert.equal(result, 2085024.31);
    assert.equal(calls[0].purchaseAccount, 12070);
    assert.equal(calls[0].group, undefined);
});
