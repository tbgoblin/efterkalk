// ── Lagerliste ──────────────────────────────────────────────────────────────
// Read-only lager value by category. Monthly snapshots are intentionally kept
// separate from live values so closing a month remains reproducible.
const path = require('path');

function createLagerlisteService({ getConnection, sql, diskCache, fs, getSalgordreViaRows, getRestPrices, dataDir }) {
    const snapshotDir = dataDir || path.join(__dirname, '..', 'data', 'lagerliste');
    const cacheKey = 'lagerliste_v8';

    function toNumber(value) {
        const parsed = Number(value || 0);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function round(value) {
        return Math.round(toNumber(value) * 100) / 100;
    }

    function readSnapshotFile(fsRef, month) {
        const file = path.join(snapshotDir, String(month || '').replace(/[^0-9-]/g, '') + '.json');
        if (!fsRef.existsSync(file)) return null;
        return JSON.parse(fsRef.readFileSync(file, 'utf8'));
    }

    function writeSnapshotFile(fsRef, month, payload) {
        fsRef.mkdirSync(snapshotDir, { recursive: true });
        const file = path.join(snapshotDir, String(month || '').replace(/[^0-9-]/g, '') + '.json');
        fsRef.writeFileSync(file, JSON.stringify(payload, null, 2) + '\n', 'utf8');
        return file;
    }

    async function getCurrent({ requestedOrdNo = null } = {}) {
        const key = requestedOrdNo ? cacheKey + '_' + requestedOrdNo : cacheKey;
        const cached = diskCache.get(key);
        if (cached) return cached;

        const pool = await getConnection();
        const todayInt = Number(new Date().toISOString().slice(0, 10).replace(/-/g, ''));
        const plateResult = await pool.request().query(`
            SELECT
                P.ProdNo,
                P.Descr,
                LEFT(CONVERT(varchar(100), P.ProdNo), 3) AS MaterialType,
                CASE SUBSTRING(CONVERT(varchar(100), P.ProdNo), 14, 1)
                    WHEN '1' THEN 'Lille'
                    WHEN '2' THEN 'Mellem'
                    WHEN '3' THEN 'Stor'
                    WHEN '4' THEN 'Jumbo'
                    WHEN '5' THEN 'Mega'
                    ELSE 'Andet'
                END AS Format,
                P.HgtU AS Thickness,
                P.WdtU AS WidthM,
                P.LgtU AS LengthM,
                P.NWgtU,
                COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0) AS Quantity,
                COALESCE(TRY_CONVERT(decimal(18, 6), REPLACE(CONVERT(varchar(100), P.Inf), ',', '.')), 0) AS StandardPrice,
                COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS UnitCost,
                COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0)
                    * COALESCE(TRY_CONVERT(decimal(18, 6), REPLACE(CONVERT(varchar(100), P.Inf), ',', '.')), 0) AS Value,
                COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0)
                    * COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS FifoValue,
                'Plade' AS Category
            FROM Prod P WITH(NOLOCK)
            INNER JOIN StcBal B WITH(NOLOCK) ON B.ProdNo = P.ProdNo AND B.StcNo = 1
                        WHERE TRY_CONVERT(decimal(18, 6), P.Gr6) = 1
                            AND CONVERT(varchar(100), P.ProdNo) LIKE '3%'
                            AND SUBSTRING(CONVERT(varchar(100), P.ProdNo), 3, 1) = '1'
                            AND COALESCE(TRY_CONVERT(decimal(18, 6), P.ProdGr), 0) <> 99999
              AND P.ProdNo <> '301001'
              AND COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0) <> 0
        `);
        const stangResult = await pool.request()
            .input('closeDate', sql.Int, todayInt)
            .query(`
                SELECT
                    P.ProdNo,
                    LEFT(CONVERT(varchar(100), P.ProdNo), 3) AS MaterialType,
                    RIGHT(CONVERT(varchar(100), P.ProdNo), 1) AS DimensionCode,
                    P.Descr,
                    P.Gr6,
                    P.Inf,
                    P.NWgtU,
                    SUM(COALESCE(TRY_CONVERT(decimal(18, 6), T.StcMov), 0)) AS Quantity,
                    COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS UnitCost,
                    SUM(COALESCE(TRY_CONVERT(decimal(18, 6), T.StcMov), 0))
                        * COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS Value,
                    'Stang' AS Category
                FROM Prod P WITH(NOLOCK)
                INNER JOIN ProdTr T WITH(NOLOCK) ON T.ProdNo = P.ProdNo
                INNER JOIN StcBal B WITH(NOLOCK) ON B.ProdNo = P.ProdNo AND B.StcNo = 1
                WHERE TRY_CONVERT(int, T.FinDt) <= @closeDate
                  AND TRY_CONVERT(decimal(18, 6), P.Gr6) = 2
                  AND T.FrStc = 1
                  AND COALESCE(TRY_CONVERT(decimal(18, 6), P.ProdGr), 0) <> 99999
                GROUP BY P.ProdNo, P.Descr, P.Gr6, P.Inf, P.NWgtU, B.PhCstPr
                HAVING SUM(COALESCE(TRY_CONVERT(decimal(18, 6), T.StcMov), 0)) <> 0
            `);
        const restPrices = typeof getRestPrices === 'function' ? getRestPrices() : { '301': 3, '311': 10, '321': 15, '331': 4, '381': 10, '302': 1 };
        const restPlateResult = await pool.request().query(`
            SELECT
                F.ProdNo,
                F.OrdNo,
                F.Txt1,
                F.Txt2,
                F.Val5,
                F.Val8,
                LEFT(CONVERT(varchar(100), F.ProdNo), 3) AS PlateType,
                P.Descr,
                P.Gr6,
                LEFT(CONVERT(varchar(100), F.ProdNo), 3) AS PriceType
            FROM FreeInf1 F WITH(NOLOCK)
            LEFT JOIN Prod P WITH(NOLOCK) ON P.ProdNo = F.ProdNo
            WHERE F.FrInfTp = 120
              AND F.ProdNo <> '3021000843542'
        `);
        const finishedRequest = pool.request();
        finishedRequest.input('requestedOrdNo', sql.Numeric, requestedOrdNo);
        const finishedResult = await finishedRequest.query(`
            SELECT
                O.OrdNo,
                O.CustNo,
                A.Nm AS CustomerName,
                SUM(
                    COALESCE(TRY_CONVERT(decimal(18, 6), L.NoFin), 0)
                    * COALESCE(TRY_CONVERT(decimal(18, 6), L.CCstPr), 0)
                ) AS Value,
                COUNT(*) AS LineCount,
                'Færdig ikke faktureret' AS Category
            FROM Ord O WITH(NOLOCK)
            LEFT JOIN Actor A WITH(NOLOCK) ON A.CustNo = O.CustNo
            INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = O.OrdNo
            WHERE O.TrTp = 1
              AND (O.InvoNo IS NULL OR O.InvoNo = '')
              AND COALESCE(TRY_CONVERT(decimal(18, 6), L.NoFin), 0) > 0
              AND (@requestedOrdNo IS NULL OR O.OrdNo = @requestedOrdNo)
            GROUP BY O.OrdNo, O.CustNo, A.Nm
        `);

        const plates = (plateResult.recordset || []).map(row => ({
            ...row,
            PlateType: String(row.ProdNo || '').trim().slice(0, 3) || String(row.MaterialType || '-'),
            Quantity: toNumber(row.Quantity),
            StandardPrice: toNumber(row.StandardPrice),
            UnitCost: toNumber(row.UnitCost),
            Value: round(row.Value),
            FifoValue: round(row.FifoValue),
            PlateCount: toNumber(row.NWgtU) > 0 ? Math.round(toNumber(row.Quantity) / toNumber(row.NWgtU)) : 0
        }));
        const plateGroupMap = new Map();
        const plateTypeLabels = {
            '301': 'Stålplader',
            '311': 'Rustfri plader',
            '321': 'Aluminiums plader',
            '331': 'Galvaniseret plader'
            , '381': 'Kobber/Messing'
        };
        for (const row of plates) {
            const group = plateGroupMap.get(row.PlateType) || {
                PlateType: row.PlateType,
                PlateTypeLabel: plateTypeLabels[row.PlateType] || '?',
                Quantity: 0,
                PlateCount: 0,
                Value: 0,
                FifoValue: 0,
                FifoPriceWeighted: 0,
                details: []
            };
            group.Quantity += row.Quantity;
            group.PlateCount += row.PlateCount;
            group.Value += row.Value;
            group.FifoValue += row.FifoValue;
            group.FifoPriceWeighted += row.Quantity * row.UnitCost;
            group.details.push(row);
            plateGroupMap.set(row.PlateType, group);
        }
        const plateGroups = Array.from(plateGroupMap.values()).map(group => ({
            ...group,
            Quantity: round(group.Quantity),
            PlateCount: Math.round(group.PlateCount),
            Value: round(group.Value),
            FifoValue: round(group.FifoValue),
            FifoPrice: group.Quantity > 0 ? round(group.FifoPriceWeighted / group.Quantity) : 0
        })).sort((left, right) => String(left.PlateType).localeCompare(String(right.PlateType)));
        const stang = (stangResult.recordset || []).map(row => ({
            ...row,
            Quantity: toNumber(row.Quantity),
            UnitCost: toNumber(row.UnitCost),
            Value: round(row.Value)
        }));
        const finishedNotInvoiced = (finishedResult.recordset || []).map(row => ({ ...row, Value: round(row.Value) }));
        const restPlates = (restPlateResult.recordset || []).map(row => {
            const weight = toNumber(row.Val5) * toNumber(row.Val8);
            const pricePerKg = toNumber(restPrices[String(row.PriceType || '').trim()] || 0);
            return {
                ...row,
                Weight: round(weight),
                PricePerKg: round(pricePerKg),
                Value: round(weight * pricePerKg)
            };
        }).filter(row => row.Weight !== 0);
        const restGroupMap = new Map();
        for (const row of restPlates) {
            const key = row.PlateType + '|' + String(row.Descr || row.Txt1 || '').trim();
            const group = restGroupMap.get(key) || {
                PlateType: row.PlateType,
                Material: String(row.Descr || row.Txt1 || '').trim() || '?',
                Weight: 0,
                Value: 0,
                details: []
            };
            group.Weight += row.Weight;
            group.Value += row.Value;
            group.details.push(row);
            restGroupMap.set(key, group);
        }
        const restPlateGroups = Array.from(restGroupMap.values()).sort((left, right) =>
            String(left.PlateType + left.Material).localeCompare(String(right.PlateType + right.Material))
        );
        const salgordreViaRows = typeof getSalgordreViaRows === 'function'
            ? await getSalgordreViaRows({ getConnection, sql, requestedOrdNo: null })
            : [];
        const salgordreVia = salgordreViaRows.map(row => ({
            OrdNo: row.OrdNo,
            CustomerName: row.CustomerName,
            MaterialCost: round(row.MaterialCost),
            StangCost: round(row.StangCost),
            TimeCost: round(row.TimeCost),
            Value: round(toNumber(row.MaterialCost) + toNumber(row.StangCost) + toNumber(row.TimeCost))
        }));
        const payload = {
            generatedAt: new Date().toISOString(),
            categories: {
                plates,
                plateGroups,
                restPlates,
                restPlateGroups,
                stang,
                finishedNotInvoiced,
                salgordreVia,
                diverse: []
            },
            totals: {
                plates: round(plates.reduce((sum, row) => sum + row.Value, 0)),
                restPlates: round(restPlates.reduce((sum, row) => sum + row.Value, 0)),
                stang: round(stang.reduce((sum, row) => sum + row.Value, 0)),
                finishedNotInvoiced: round(finishedNotInvoiced.reduce((sum, row) => sum + row.Value, 0)),
                salgordreVia: round(salgordreVia.reduce((sum, row) => sum + row.Value, 0)),
                diverse: 0
            }
        };
        payload.totals.total = round(Object.values(payload.totals).reduce((sum, value) => sum + toNumber(value), 0));
        diskCache.set(key, payload, 5 * 60 * 1000);
        return payload;
    }

    async function saveMonthlySnapshot({ fs, month, diverse = [] }) {
        const current = await getCurrent();
        const payload = { month, createdAt: new Date().toISOString(), current, diverse };
        return { ...payload, file: writeSnapshotFile(fs, month, payload) };
    }

    function loadMonthlySnapshot({ fs, month }) {
        return readSnapshotFile(fs, month);
    }

    function scheduleMonthlySnapshot({ onError } = {}) {
        const runIfMonthEnd = async () => {
            const now = new Date();
            const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
            if (tomorrow.getMonth() === now.getMonth()) return;
            const month = now.toISOString().slice(0, 7);
            if (readSnapshotFile(fs, month)) return;
            try {
                await saveMonthlySnapshot({ fs, month, diverse: [] });
            } catch (err) {
                if (typeof onError === 'function') onError(err);
            }
        };
        runIfMonthEnd();
        return setInterval(runIfMonthEnd, 6 * 60 * 60 * 1000);
    }

    return { getCurrent, saveMonthlySnapshot, loadMonthlySnapshot, scheduleMonthlySnapshot };
}

module.exports = { createLagerlisteService };
