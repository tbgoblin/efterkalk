const test = require('node:test');
const assert = require('node:assert/strict');

const thresholds = require('../assets/js/omsaetning-daily-thresholds');
const thresholdStore = require('../services/omsaetningThresholdsService');

test('daily threshold defaults preserve the documented break-even and budget', () => {
    assert.deepEqual(thresholds.sanitizeConfig({ useDailyBudget: true }), {
        useDailyBudget: true,
        dailyBreakEvenDkk: 208335,
        dailyBudgetDkk: 280851
    });
});

test('working days exclude weekends and official Danish public holidays', () => {
    assert.equal(thresholds.countWorkingDaysInMonth('2026-08'), 21);
    assert.equal(thresholds.countWorkingDaysInMonth('2026-05'), 19);
    assert.equal(thresholds.countWorkingDaysInMonth('2026-12'), 22);
});

test('shared company holiday weeks reduce only weekdays in the selected month', () => {
    const withoutCompanyHoliday = thresholds.countWorkingDaysInMonth('2026-12');
    const withWeek52Holiday = thresholds.countWorkingDaysInMonth('2026-12', {
        holidayWeekKeys: new Set(['202652'])
    });
    assert.equal(withoutCompanyHoliday, 22);
    assert.equal(withWeek52Holiday, 18);
});

test('month targets multiply editable daily values by displayed working days', () => {
    const targets = thresholds.getMonthTargets('2026-08', {
        useDailyBudget: true,
        dailyBreakEvenDkk: 208335,
        dailyBudgetDkk: 280851
    });
    assert.equal(targets.workingDays, 21);
    assert.equal(targets.warnThresholdMio, 4.375035);
    assert.equal(targets.goodThresholdMio, 5.897871);
});

test('a manual monthly working-day value overrides the automatic calendar count', () => {
    const targets = thresholds.getMonthTargets('2026-08', {
        useDailyBudget: true,
        dailyBreakEvenDkk: 200000,
        dailyBudgetDkk: 300000
    }, { manualWorkingDays: 17 });
    assert.equal(targets.workingDays, 17);
    assert.equal(targets.warnThresholdMio, 3.4);
    assert.equal(targets.goodThresholdMio, 5.1);
    assert.equal(thresholds.countWorkingDaysInMonth('2026-08', { manualWorkingDays: 32 }), 21);
});

test('monthly working-day settings keep only valid month keys and integer values', () => {
    assert.deepEqual(thresholds.normalizeWorkingDaysByMonth({
        '2026-01': 20,
        '2026-02': '19',
        '2026-13': 10,
        invalid: 12,
        '2026-03': 32,
        '2026-04': 18.5
    }), { '2026-01': 20, '2026-02': 19 });
});

test('a budget below break-even is normalized to a truthful non-negative range', () => {
    const config = thresholds.sanitizeConfig({
        useDailyBudget: true,
        dailyBreakEvenDkk: 250000,
        dailyBudgetDkk: 200000
    });
    assert.equal(config.dailyBreakEvenDkk, 250000);
    assert.equal(config.dailyBudgetDkk, 250000);
});

test('stored customer thresholds remain backward compatible and persist daily mode fields', () => {
    const legacy = thresholdStore.normalizeThresholds(3, 5, {});
    assert.equal(legacy.useDailyBudget, false);
    assert.equal(legacy.dailyBreakEvenDkk, 208335);
    assert.equal(legacy.dailyBudgetDkk, 280851);

    const daily = thresholdStore.normalizeThresholds(3, 5, {
        useDailyBudget: true,
        dailyBreakEvenDkk: 210000,
        dailyBudgetDkk: 300000
    });
    assert.equal(daily.useDailyBudget, true);
    assert.equal(daily.dailyBreakEvenDkk, 210000);
    assert.equal(daily.dailyBudgetDkk, 300000);
});
