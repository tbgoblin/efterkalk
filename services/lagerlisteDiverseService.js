const { randomUUID } = require('crypto');
const names = ['Div. bolte', 'Paller', 'Forbrugsmatl. Pakkeri', 'Gasser', 'Forbrugsmatl. Svejseafd.', 'Kølevæske', 'Skrot Alu', 'Skrot RF', 'Skrot Sort'];
const monthNow = () => new Intl.DateTimeFormat('sv-SE', { timeZone: 'Europe/Copenhagen', year: 'numeric', month: '2-digit' }).format(new Date());
function validateMonth(month) {
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Ugyldig måned');
    return month;
}
function calculateRows(rows) {
    if (!Array.isArray(rows) || rows.length > 300) throw new Error('Ugyldige Diverse-linjer');
    for (const category of names) {
        const group = rows.filter(row => row.category === category);
        if (group.length > 1 && group.some(row => row.mode === 'amount') && !category.startsWith('Skrot ')) throw new Error(category + ': Brug enten ét direkte beløb eller detaljelinjer, ikke begge.');
    }
    return rows.map(row => {
        if (!names.includes(row.category)) throw new Error('Ukendt kategori');
        const mode = row.category.startsWith('Skrot ') ? 'pallets' : row.mode;
        if (!['amount', 'quantity', 'pallets'].includes(mode)) throw new Error('Ugyldig beregning');
        const number = key => {
            if (row[key] === '' || row[key] == null) return null;
            const value = Number(String(row[key]).replace(',', '.'));
            if (!Number.isFinite(value) || value < 0 || value > 1e10) throw new Error('Ugyldigt tal: ' + key);
            return value;
        };
        const amount = number('amount'), quantity = number('quantity'), price = number('price'), kg = number('kg');
        const inputs = mode === 'amount' ? [amount] : mode === 'quantity' ? [quantity, price] : [quantity, kg, price];
        const complete = inputs.every(value => value !== null);
        const Value = complete ? Math.round(inputs.reduce((a, b) => a * b, 1) * 100) / 100 : 0;
        if (!Number.isSafeInteger(Math.round(Value * 100))) throw new Error('Beløbet er for stort');
        return { category: row.category, Descr: String(row.Descr || row.category).slice(0, 150), mode, amount, quantity, price, kg, complete, Value };
    });
}
function defaultRows() {
    return calculateRows(names.map(category => ({ category, mode: 'amount' })));
}
function createLagerlisteDiverseService({ gohData, getConnection }) {
    async function load(month) {
        validateMonth(month);
        const stored = await gohData.getAppState('lagerliste_diverse_' + month);
        if (typeof gohData.isEnabled === 'function' && !gohData.isEnabled()) throw new Error('GOH er ikke tilgængelig. Diverse kan ikke hentes sikkert. Prøv igen senere.');
        return stored && stored.payload || { month, rows: defaultRows(), saved: false };
    }
    async function save(month, rows, username) {
        validateMonth(month);
        const validated = calculateRows(rows);
        if (names.some(name => !validated.some(row => row.category === name))) throw new Error('Alle kategorier skal være med');
        const payload = { month, rows: validated, saved: true, updatedAt: new Date().toISOString(), updatedBy: String(username || ''), revision: randomUUID() };
        if (!await gohData.setAppState('lagerliste_diverse_rev_' + payload.revision, payload, { createOnly: true })) throw new Error('GOH kunne ikke gemme revisionen');
        if (!await gohData.setAppState('lagerliste_diverse_' + month, payload)) throw new Error('GOH kunne ikke gemme månedens værdier');
        return payload;
    }
    async function current() {
        const manual = await load(monthNow());
        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT P.ProdNo, P.Inf2 AS TegnNr, P.Descr,
                B.PoPhStB AS Quantity,
                TRY_CONVERT(decimal(18,6), REPLACE(CONVERT(varchar(100), P.Inf), ',', '.')) AS Price
            FROM Prod P WITH(NOLOCK)
            JOIN StcBal B WITH(NOLOCK) ON B.ProdNo = P.ProdNo AND B.StcNo = 1
            WHERE (B.Bal + B.StcInc - (B.ShpRsv + B.ShpRsvIn) - B.PicNotR) <> 0
              AND P.Gr5 IN (2,3)
              AND (P.ProdNo LIKE '44%' OR P.ProdNo LIKE '45%' OR P.ProdNo LIKE '46%' OR P.ProdNo LIKE '63%')
        `);
        const labels = { '44': 'PEM (44)', '45': 'Sv. bolte (45)', '46': 'POP nitter (46)', '63': 'Muffer (63)' };
        const automatic = (result.recordset || []).map(row => {
            const complete = row.Price != null && row.Quantity != null && Number.isFinite(Number(row.Price)) && Number.isFinite(Number(row.Quantity));
            return { ...row, category: labels[String(row.ProdNo).slice(0, 2)], mode: 'visma', complete, Value: complete ? Math.round(Number(row.Quantity) * Number(row.Price) * 100) / 100 : 0 };
        });
        const rows = [...calculateRows(manual.rows), ...automatic];
        return { month: manual.month, revision: manual.revision || null, rows, complete: rows.every(row => row.complete), total: Math.round(rows.reduce((sum, row) => sum + row.Value, 0) * 100) / 100 };
    }
    return { load, save, current };
}
module.exports = { createLagerlisteDiverseService, calculateRows, defaultRows, monthNow };
