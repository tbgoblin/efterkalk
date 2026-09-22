function monthKey(year, month) {
    const date = new Date(Date.UTC(year, month - 1, 1));
    return date.toISOString().slice(0, 7);
}
function sourceMonth(row, year, period, prior, today) {
    if (row.period === 'fixed') return row.month;
    if (row.period === 'current') {
        const [y, m] = today.split('-').map(Number);
        return monthKey(y - prior, m);
    }
    const calendarYear = year + (period > 6 ? 1 : 0) - prior;
    const calendarMonth = ((period + 5) % 12) + 1;
    return monthKey(calendarYear, calendarMonth - (row.period === 'previous' ? 1 : 0));
}
function lagerValue(payload, metric) {
    if (!payload?.totals) return null;
    const t = payload.totals, c = payload.categories || {};
    const valid = v => v != null && Number.isFinite(Number(v)) ? Number(v) : null;
    if (metric === 'sales') return valid(t.finishedNotInvoicedSales);
    if (metric === 'cost') return valid(t.finishedNotInvoiced);
    if (metric === 'margin') {
        const sales = valid(t.finishedNotInvoicedSales), cost = valid(t.finishedNotInvoiced);
        return sales == null || cost == null ? null : Math.round((sales - cost) * 100) / 100;
    }
    if (metric === 'warehouse') {
        return ['plates','restPlates','opfolgningvare','stang','diverse'].reduce((sum,k) => sum + Number(t[k] || 0), 0)
            + (c.gr5Items || []).reduce((sum,r) => sum + Number(r.FifoValue || 0), 0);
    }
    if (metric === 'via') {
        return Number(t.finishedNotInvoiced || 0) + (c.salgordreVia || []).reduce((sum,r) => sum + ['TimeCost','MaterialCost','StangCost','PurchasedPartCost'].reduce((s,k) => s + Number(r[k] || 0), 0), 0)
            + (c.nestingCutting || []).reduce((sum,r) => sum + (r.CountedValue != null ? Number(r.CountedValue) : Math.max(0, Number(r.Value || 0))), 0);
    }
    return null;
}
function createSourceResolver({ lagerlisteService, fs, year, period, today }) {
    const cache = new Map();
    function payload(month, currentSales) {
        const key = month + '/' + currentSales;
        if (!cache.has(key)) cache.set(key, (async () => {
            if (!lagerlisteService || !fs) return null;
            const snapshot = await lagerlisteService.loadMonthlySnapshot({ fs, month, includeCurrentSales: currentSales });
            if (snapshot) return snapshot.current;
            return month === today ? lagerlisteService.getCurrent() : null;
        })());
        return cache.get(key);
    }
    return async (rows, dimension) => {
        const values = new Map();
        for (const row of rows.filter(r => ['lager','manual'].includes(r.type))) {
            const pair = await Promise.all([0,1].map(async prior => {
                const month = sourceMonth(row, year, period, prior, today);
                const value = row.type === 'manual' ? row.monthValues[month] ?? null : lagerValue(await payload(month, row.currentSales), row.metric);
                return value != null && Number.isFinite(value) ? value : null;
            }));
            values.set(row.id, dimension === 4 ? [...pair, ...pair] : pair);
        }
        return values;
    };
}
module.exports = { sourceMonth, lagerValue, createSourceResolver };
