(function (root, factory) {
    const api = factory();
    if (typeof module === 'object' && module.exports) module.exports = api;
    if (root) root.OmsaetningDailyThresholds = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict';

    const DEFAULT_DAILY_BREAK_EVEN_DKK = 208335;
    const DEFAULT_DAILY_BUDGET_DKK = 280851;

    function sanitizeConfig(rawConfig) {
        const raw = rawConfig && typeof rawConfig === 'object' ? rawConfig : {};
        const breakEvenRaw = Number(raw.dailyBreakEvenDkk);
        const budgetRaw = Number(raw.dailyBudgetDkk);
        const dailyBreakEvenDkk = Number.isFinite(breakEvenRaw)
            ? Math.max(0, breakEvenRaw)
            : DEFAULT_DAILY_BREAK_EVEN_DKK;
        const dailyBudgetDkk = Number.isFinite(budgetRaw)
            ? Math.max(dailyBreakEvenDkk, budgetRaw)
            : Math.max(dailyBreakEvenDkk, DEFAULT_DAILY_BUDGET_DKK);
        return {
            useDailyBudget: raw.useDailyBudget === true,
            dailyBreakEvenDkk: Math.round(dailyBreakEvenDkk),
            dailyBudgetDkk: Math.round(dailyBudgetDkk)
        };
    }

    function addUtcDays(date, days) {
        const result = new Date(date.getTime());
        result.setUTCDate(result.getUTCDate() + Number(days || 0));
        return result;
    }

    function dateKey(date) {
        const year = date.getUTCFullYear();
        const month = String(date.getUTCMonth() + 1).padStart(2, '0');
        const day = String(date.getUTCDate()).padStart(2, '0');
        return year + '-' + month + '-' + day;
    }

    // Meeus/Jones/Butcher Gregorian Easter algorithm.
    function getEasterSundayUtc(yearValue) {
        const year = Number(yearValue);
        const a = year % 19;
        const b = Math.floor(year / 100);
        const c = year % 100;
        const d = Math.floor(b / 4);
        const e = b % 4;
        const f = Math.floor((b + 8) / 25);
        const g = Math.floor((b - f + 1) / 3);
        const h = (19 * a + b - d - g + 15) % 30;
        const i = Math.floor(c / 4);
        const k = c % 4;
        const l = (32 + 2 * e + 2 * i - h - k) % 7;
        const m = Math.floor((a + 11 * h + 22 * l) / 451);
        const month = Math.floor((h + l - 7 * m + 114) / 31);
        const day = ((h + l - 7 * m + 114) % 31) + 1;
        return new Date(Date.UTC(year, month - 1, day));
    }

    function getDanishPublicHolidayKeys(yearValue) {
        const year = Number(yearValue);
        const easter = getEasterSundayUtc(year);
        const keys = new Set([
            year + '-01-01',
            year + '-12-25',
            year + '-12-26',
            dateKey(addUtcDays(easter, -3)), // Skærtorsdag
            dateKey(addUtcDays(easter, -2)), // Langfredag
            dateKey(addUtcDays(easter, 1)),  // 2. påskedag
            dateKey(addUtcDays(easter, 39)), // Kristi himmelfartsdag
            dateKey(addUtcDays(easter, 50))  // 2. pinsedag
        ]);
        // Store bededag was a public holiday through 2023 and was abolished from 2024.
        if (year <= 2023) keys.add(dateKey(addUtcDays(easter, 26)));
        return keys;
    }

    function getIsoWeekKeyUtc(date) {
        const working = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
        const day = working.getUTCDay() || 7;
        working.setUTCDate(working.getUTCDate() + 4 - day);
        const isoYear = working.getUTCFullYear();
        const yearStart = new Date(Date.UTC(isoYear, 0, 1));
        const week = Math.ceil((((working - yearStart) / 86400000) + 1) / 7);
        return String(isoYear) + String(week).padStart(2, '0');
    }

    function normalizeHolidayWeekSet(values) {
        if (values instanceof Set) return values;
        if (!Array.isArray(values)) return new Set();
        return new Set(values.map(value => String(value || '').trim()).filter(value => /^\d{6}$/.test(value)));
    }

    function normalizeWorkingDaysByMonth(values) {
        const source = values && typeof values === 'object' ? values : {};
        const normalized = {};
        for (const [monthKey, rawValue] of Object.entries(source)) {
            if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(monthKey)) continue;
            const value = Number(rawValue);
            if (Number.isInteger(value) && value >= 0 && value <= 31) normalized[monthKey] = value;
        }
        return normalized;
    }

    function countWorkingDaysInMonth(monthKey, options) {
        const match = String(monthKey || '').trim().match(/^(\d{4})-(\d{2})$/);
        if (!match) return 0;
        const year = Number(match[1]);
        const month = Number(match[2]);
        if (month < 1 || month > 12) return 0;

        const safeOptions = options && typeof options === 'object' ? options : {};
        const manualWorkingDays = Number(safeOptions.manualWorkingDays);
        if (Number.isInteger(manualWorkingDays) && manualWorkingDays >= 0 && manualWorkingDays <= 31) {
            return manualWorkingDays;
        }
        const holidayWeeks = normalizeHolidayWeekSet(safeOptions.holidayWeekKeys);
        const excludeCompanyHolidayWeeks = safeOptions.excludeCompanyHolidayWeeks !== false;
        const publicHolidays = getDanishPublicHolidayKeys(year);
        const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
        let count = 0;

        for (let day = 1; day <= daysInMonth; day += 1) {
            const date = new Date(Date.UTC(year, month - 1, day));
            const weekday = date.getUTCDay();
            if (weekday === 0 || weekday === 6) continue;
            if (publicHolidays.has(dateKey(date))) continue;
            if (excludeCompanyHolidayWeeks && holidayWeeks.has(getIsoWeekKeyUtc(date))) continue;
            count += 1;
        }
        return count;
    }

    function getMonthTargets(monthKey, rawConfig, options) {
        const config = sanitizeConfig(rawConfig);
        const workingDays = countWorkingDaysInMonth(monthKey, options);
        return {
            workingDays,
            warnThresholdMio: (config.dailyBreakEvenDkk * workingDays) / 1000000,
            goodThresholdMio: (config.dailyBudgetDkk * workingDays) / 1000000,
            dailyBreakEvenDkk: config.dailyBreakEvenDkk,
            dailyBudgetDkk: config.dailyBudgetDkk
        };
    }

    return {
        DEFAULT_DAILY_BREAK_EVEN_DKK,
        DEFAULT_DAILY_BUDGET_DKK,
        sanitizeConfig,
        getEasterSundayUtc,
        getDanishPublicHolidayKeys,
        getIsoWeekKeyUtc,
        normalizeWorkingDaysByMonth,
        countWorkingDaysInMonth,
        getMonthTargets
    };
});
