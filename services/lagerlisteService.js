// ── Lagerliste ──────────────────────────────────────────────────────────────
// Read-only lager value by category. Monthly snapshots are intentionally kept
// separate from live values so closing a month remains reproducible.
const path = require('path');

function createLagerlisteService({ getConnection, sql, diskCache, fs, getSalgordreViaRows, getOrComputeAftercalc, getProductionSummary, getRestPrices, dataDir }) {
    const snapshotDir = dataDir || path.join(__dirname, '..', 'data', 'lagerliste');
    const historyDir = path.join(snapshotDir, 'history');
    const cacheKey = 'lagerliste_v16';
    let currentMemoryCache = null;

    function toNumber(value) {
        if (typeof value === 'string') {
            const normalized = value.replace(/\s+/g, '').replace(',', '.');
            const parsedString = Number(normalized);
            return Number.isFinite(parsedString) ? parsedString : 0;
        }
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

    async function writeSnapshotFile(fsRef, month, payload) {
        fsRef.mkdirSync(snapshotDir, { recursive: true });
        const file = path.join(snapshotDir, String(month || '').replace(/[^0-9-]/g, '') + '.json');
        const content = JSON.stringify(payload) + '\n';
        if (fsRef.promises && typeof fsRef.promises.writeFile === 'function') {
            await fsRef.promises.writeFile(file, content, 'utf8');
        } else {
            fsRef.writeFileSync(file, content, 'utf8');
        }
        return file;
    }

    function buildSnapshotId(date) {
        const pad = value => String(value).padStart(2, '0');
        return String(date.getFullYear())
            + '-' + pad(date.getMonth() + 1)
            + '-' + pad(date.getDate())
            + '_' + pad(date.getHours())
            + '-' + pad(date.getMinutes())
            + '-' + pad(date.getSeconds());
    }

    async function writePointInTimeSnapshotFile(fsRef, snapshotId, payload) {
        fsRef.mkdirSync(historyDir, { recursive: true });
        const safeId = String(snapshotId || '').replace(/[^0-9A-Za-z_\-]/g, '');
        const file = path.join(historyDir, safeId + '.json');
        const content = JSON.stringify(payload) + '\n';
        if (fsRef.promises && typeof fsRef.promises.writeFile === 'function') {
            await fsRef.promises.writeFile(file, content, 'utf8');
        } else {
            fsRef.writeFileSync(file, content, 'utf8');
        }
        return file;
    }

    function listPointInTimeSnapshots(fsRef) {
        if (!fsRef.existsSync(historyDir)) return [];
        const files = fsRef.readdirSync(historyDir)
            .filter(name => String(name).toLowerCase().endsWith('.json'))
            .sort((left, right) => String(right).localeCompare(String(left)));

        return files.map(name => {
            const file = path.join(historyDir, name);
            let createdAt = null;
            let capturedAt = null;
            try {
                const stat = fsRef.statSync(file);
                createdAt = stat && stat.mtime ? new Date(stat.mtime).toISOString() : null;
            } catch (_err) {
                createdAt = null;
            }
            try {
                const parsed = JSON.parse(fsRef.readFileSync(file, 'utf8'));
                capturedAt = parsed && parsed.capturedAt ? parsed.capturedAt : null;
            } catch (_err) {
                capturedAt = null;
            }
            return {
                snapshotId: String(name).replace(/\.json$/i, ''),
                file,
                createdAt,
                capturedAt
            };
        });
    }

    function loadPointInTimeSnapshot(fsRef, snapshotId) {
        const safeId = String(snapshotId || '').replace(/[^0-9A-Za-z_\-]/g, '');
        if (!safeId) return null;
        const file = path.join(historyDir, safeId + '.json');
        if (!fsRef.existsSync(file)) return null;
        return JSON.parse(fsRef.readFileSync(file, 'utf8'));
    }

    async function mapWithConcurrency(items, worker, concurrency = 4) {
        const safeItems = Array.isArray(items) ? items : [];
        const out = new Array(safeItems.length);
        let next = 0;

        async function pump() {
            while (true) {
                const index = next++;
                if (index >= safeItems.length) return;
                out[index] = await worker(safeItems[index], index);
            }
        }

        const workers = Array.from({ length: Math.min(Math.max(1, concurrency), safeItems.length) }, pump);
        await Promise.all(workers);
        return out;
    }

    async function getCurrent({ requestedOrdNo = null, forceRefresh = false } = {}) {
        const key = requestedOrdNo ? cacheKey + '_stang_v6_' + requestedOrdNo : cacheKey + '_stang_v6';
        if (!forceRefresh) {
            const cached = diskCache.get(key);
            if (cached) {
                currentMemoryCache = cached;
                return cached;
            }
        } else {
            diskCache.del(key);
        }

        const pool = await getConnection();
        const todayInt = Number(new Date().toISOString().slice(0, 10).replace(/-/g, ''));
        const currentMonthDate = new Date();
        currentMonthDate.setDate(1);
        currentMonthDate.setMonth(currentMonthDate.getMonth() - 2);
        const currentMonthStart = Number(currentMonthDate.toISOString().slice(0, 7).replace('-', '') + '01');
        const nextMonthDate = new Date(Date.UTC(new Date().getUTCFullYear(), new Date().getUTCMonth() + 1, 1));
        const nextMonthStart = Number(nextMonthDate.toISOString().slice(0, 10).replace(/-/g, ''));
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
        const gr5Result = await pool.request().query(`
            SELECT
                P.ProdNo,
                P.Descr,
                P.Gr5,
                P.NWgtU,
                COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0) AS Quantity,
                COALESCE(TRY_CONVERT(decimal(18, 6), REPLACE(CONVERT(varchar(100), P.Inf), ',', '.')), 0) AS StandardPrice,
                COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS UnitCost,
                COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0)
                    * COALESCE(TRY_CONVERT(decimal(18, 6), REPLACE(CONVERT(varchar(100), P.Inf), ',', '.')), 0) AS Value,
                COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0)
                    * COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS FifoValue
            FROM Prod P WITH(NOLOCK)
            INNER JOIN StcBal B WITH(NOLOCK) ON B.ProdNo = P.ProdNo AND B.StcNo = 1
            WHERE TRY_CONVERT(decimal(18, 6), P.Gr5) = 11
              AND COALESCE(TRY_CONVERT(decimal(18, 6), P.ProdGr), 0) <> 99999
              AND COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0) <> 0
            ORDER BY P.ProdNo
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
                    COALESCE(TRY_CONVERT(decimal(18, 6), REPLACE(CONVERT(varchar(100), P.Inf), ',', '.')), 0) AS StandardPrice,
                    COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS UnitCost,
                    SUM(COALESCE(TRY_CONVERT(decimal(18, 6), T.StcMov), 0))
                        * COALESCE(TRY_CONVERT(decimal(18, 6), REPLACE(CONVERT(varchar(100), P.Inf), ',', '.')), 0) AS Value,
                    SUM(COALESCE(TRY_CONVERT(decimal(18, 6), T.StcMov), 0))
                        * COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS FifoValue,
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
                const opfolgningResult = await pool.request().query(`
                        SELECT
                                P.Gr9,
                                P.ProdNo,
                                P.Descr,
                                COALESCE(TRY_CONVERT(decimal(18, 6), B.PoPhStB), 0) AS PoPhStB,
                                P.ProdGr,
                                COALESCE(TRY_CONVERT(decimal(18, 6), B.Bal), 0)
                                    + COALESCE(TRY_CONVERT(decimal(18, 6), B.StcInc), 0)
                                    - COALESCE(TRY_CONVERT(decimal(18, 6), B.ShpRsv), 0) AS Beholdning,
                                COALESCE(TRY_CONVERT(decimal(18, 6), B.PhCstPr), 0) AS PhCstPr,
                                COALESCE(TRY_CONVERT(decimal(18, 6), B.Bal), 0) AS Bal,
                                COALESCE(TRY_CONVERT(decimal(18, 6), B.StcInc), 0) AS StcInc,
                                COALESCE(TRY_CONVERT(decimal(18, 6), B.ShpRsv), 0) AS ShpRsv,
                                COALESCE(TRY_CONVERT(decimal(18, 6), B.ShpRsvIn), 0) AS ShpRsvIn,
                                COALESCE(TRY_CONVERT(decimal(18, 6), B.PicNotR), 0) AS PicNotR,
                                SB.LatestRecDt AS ShpBal_RecDt_Ultimo
                        FROM Prod P WITH(NOLOCK)
                        INNER JOIN StcBal B WITH(NOLOCK) ON B.ProdNo = P.ProdNo AND B.StcNo = 1
                        LEFT JOIN (
                                SELECT ProdNo, MAX(RecDt) AS LatestRecDt
                                FROM ShpBal WITH(NOLOCK)
                                GROUP BY ProdNo
                        ) SB ON SB.ProdNo = P.ProdNo
                        WHERE TRY_CONVERT(decimal(18, 6), P.Gr9) = 1
                            AND CONVERT(varchar(100), P.ProdNo) LIKE '1%'
                            AND CONVERT(varchar(100), P.ProdNo) NOT LIKE '%L%'
                            AND COALESCE(TRY_CONVERT(decimal(18, 6), P.ProdGr), 0) IN (1, 2)
                            AND (
                                        COALESCE(TRY_CONVERT(decimal(18, 6), B.Bal), 0)
                                        + COALESCE(TRY_CONVERT(decimal(18, 6), B.StcInc), 0)
                                        - COALESCE(TRY_CONVERT(decimal(18, 6), B.ShpRsv), 0)
                                    ) <> 0
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
                const nestingCuttingRequest = pool.request();
                nestingCuttingRequest.input('currentMonthStart', sql.Int, currentMonthStart);
                nestingCuttingRequest.input('nextMonthStart', sql.Int, nextMonthStart);
                const nestingCuttingResult = await nestingCuttingRequest.query(`
                        SELECT
                                O.OrdNo,
                                O.OrdDt,
                                L.LnNo,
                                L.TrInf2,
                                L.TrInf4,
                                L.ProdNo,
                                L.TrTp,
                                L.NoOrg,
                                L.NoFin,
                                L.NoInvoAb,
                                L.CstPr,
                                L.IncCst
                        FROM Ord O WITH(NOLOCK)
                        INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = O.OrdNo
                        WHERE TRY_CONVERT(int, O.OrdDt) >= @currentMonthStart
                            AND TRY_CONVERT(int, O.OrdDt) < @nextMonthStart
                            AND TRY_CONVERT(decimal(18, 6), O.Gr3) = 2
                            AND L.TrTp IN (5, 7)
                            AND NULLIF(LTRIM(RTRIM(CONVERT(varchar(100), L.TrInf4))), '') IS NOT NULL
                        ORDER BY O.OrdNo, L.TrInf4, L.LnNo
                `);
        const finishedRequest = pool.request();
        finishedRequest.input('requestedOrdNo', sql.Numeric, requestedOrdNo);
        const finishedResult = await finishedRequest.query(`
            SELECT
                O.OrdNo,
                O.CustNo,
                A.Nm AS CustomerName,
                O.Gr4,
                SUM(
                    COALESCE(TRY_CONVERT(decimal(18, 6), L.NoFin), 0)
                    * COALESCE(TRY_CONVERT(decimal(18, 6), L.CCstPr), 0)
                ) AS LegacyValue,
                COUNT(*) AS LineCount,
                'Færdig ikke faktureret' AS Category
            FROM Ord O WITH(NOLOCK)
            LEFT JOIN Actor A WITH(NOLOCK) ON A.CustNo = O.CustNo
            INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = O.OrdNo
            WHERE O.TrTp = 1
              AND (O.InvoNo IS NULL OR O.InvoNo = '')
                            AND COALESCE(TRY_CONVERT(decimal(18, 6), L.NoFin), 0) > 0
              AND (@requestedOrdNo IS NULL OR O.OrdNo = @requestedOrdNo)
            GROUP BY O.OrdNo, O.CustNo, A.Nm, O.Gr4
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
        const gr5Items = (gr5Result.recordset || []).map(row => ({
            ...row,
            Quantity: toNumber(row.Quantity),
            StandardPrice: toNumber(row.StandardPrice),
            UnitCost: toNumber(row.UnitCost),
            Value: round(row.Value),
            FifoValue: round(row.FifoValue)
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
        const stang = (stangResult.recordset || []).map(row => {
            const quantity = toNumber(row.Quantity);
            const standardPrice = toNumber(row.StandardPrice);
            const unitCost = toNumber(row.UnitCost);
            const value = round(quantity * standardPrice);
            const fifoValue = round(quantity * unitCost);
            return {
                ...row,
                Quantity: quantity,
                StandardPrice: standardPrice,
                UnitCost: unitCost,
                Value: value,
                FifoValue: fifoValue
            };
        });
        const opfolgningvare = (opfolgningResult.recordset || []).map(row => {
            const beholdning = toNumber(row.Beholdning);
            const fifoPrice = toNumber(row.PhCstPr);
            const pophStB = toNumber(row.PoPhStB);
            const preciseValue = beholdning * fifoPrice;
            const precisePoPhStBValue = pophStB * fifoPrice;
            return {
                ...row,
                Beholdning: beholdning,
                PoPhStB: pophStB,
                PhCstPr: fifoPrice,
                Value: round(preciseValue),
                PoPhStBValue: round(precisePoPhStBValue),
                Diff: round(precisePoPhStBValue - preciseValue),
                _preciseValue: preciseValue
            };
        });
        const finishedNotInvoiced = await mapWithConcurrency(finishedResult.recordset || [], async row => {
            const legacyValue = round(row.LegacyValue);
            const ordNo = Number(row.OrdNo || 0);
            let effectiveCost = legacyValue;
            let hasValidAftercalcCost = false;

            if (typeof getOrComputeAftercalc === 'function' && Number.isFinite(ordNo) && ordNo > 0) {
                try {
                    const aftercalc = await getOrComputeAftercalc(ordNo, { priority: 'high' });
                    const aftercalcCost = Number(aftercalc && aftercalc.summary && aftercalc.summary.totalCost);
                    if (Number.isFinite(aftercalcCost) && !(aftercalcCost === 0 && legacyValue > 0)) {
                        effectiveCost = aftercalcCost;
                        hasValidAftercalcCost = true;
                    }
                } catch (_err) {
                    // Fall back to production summary, then legacy value.
                }
            }

            if (!hasValidAftercalcCost && (!Number.isFinite(effectiveCost) || effectiveCost === legacyValue) && typeof getProductionSummary === 'function' && Number.isFinite(ordNo) && ordNo > 0) {
                try {
                    const summary = await getProductionSummary(ordNo, new Set(), { orderGr4: Number(row.Gr4 || 0) });
                    const summaryCost = Number(summary && summary.totalCost);
                    if (Number.isFinite(summaryCost)) {
                        effectiveCost = summaryCost;
                    }
                } catch (_err) {
                    // Keep legacy value as last fallback when summary is unavailable.
                }
            }

            return {
                ...row,
                LegacyValue: legacyValue,
                Value: round(effectiveCost)
            };
        });
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
        const nestingCuttingGroups = new Map();
        for (const row of nestingCuttingResult.recordset || []) {
            const route = String(row.TrInf4 || '').trim();
            const key = String(row.OrdNo || '') + '|' + route;
            const group = nestingCuttingGroups.get(key) || {
                OrdNo: row.OrdNo,
                OrdDt: row.OrdDt,
                Route: route,
                plates: [],
                products: []
            };
            if (Number(row.TrTp) === 5) group.plates.push(row);
            if (Number(row.TrTp) === 7) group.products.push(row);
            nestingCuttingGroups.set(key, group);
        }
        const nestingCutting = [];
        for (const group of nestingCuttingGroups.values()) {
            if (!group.plates.length || !group.products.length) continue;
            const plateIsFinished = group.plates.every(row => toNumber(row.NoFin) > 0);
            const allProductsUnfinished = group.products.every(row => toNumber(row.NoFin) === 0);
            if (!plateIsFinished || !allProductsUnfinished) continue;
            for (const plate of group.plates) {
                const value = toNumber(plate.IncCst) || toNumber(plate.CstPr) * toNumber(plate.NoFin);
                nestingCutting.push({
                    OrdNo: group.OrdNo,
                    OrdDt: group.OrdDt,
                    Route: group.Route,
                    ProdNo: String(plate.ProdNo || '').trim(),
                    Quantity: round(plate.NoFin),
                    Value: round(value),
                    ProductCount: group.products.length
                });
            }
        }
        const salgordreViaRows = typeof getSalgordreViaRows === 'function'
            ? await getSalgordreViaRows({ getConnection, sql, requestedOrdNo: null })
            : [];
        const salgordreVia = salgordreViaRows.map(row => ({
            OrdNo: row.OrdNo,
            CustomerName: row.CustomerName,
            MaterialCost: round(row.MaterialCost),
            StangCost: round(row.StangCost),
            PurchasedPartCost: round(row.PurchasedPartCost),
            TimeCost: round(row.TimeCost),
            Value: round(toNumber(row.MaterialCost) + toNumber(row.StangCost) + toNumber(row.PurchasedPartCost) + toNumber(row.TimeCost))
        }));
        const payload = {
            generatedAt: new Date().toISOString(),
            categories: {
                plates,
                gr5Items,
                plateGroups,
                restPlates,
                restPlateGroups,
                stang,
                opfolgningvare,
                nestingCutting,
                finishedNotInvoiced,
                salgordreVia,
                diverse: []
            },
            totals: {
                plates: round(plates.reduce((sum, row) => sum + row.Value, 0)),
                restPlates: round(restPlates.reduce((sum, row) => sum + row.Value, 0)),
                stang: round(stang.reduce((sum, row) => sum + row.Value, 0)),
                opfolgningvare: round(opfolgningvare.reduce((sum, row) => sum + toNumber(row._preciseValue), 0)),
                finishedNotInvoiced: round(finishedNotInvoiced.reduce((sum, row) => sum + row.Value, 0)),
                salgordreVia: round(salgordreVia.reduce((sum, row) => sum + row.Value, 0)),
                diverse: 0
            }
        };
        payload.totals.total = round(Object.values(payload.totals).reduce((sum, value) => sum + toNumber(value), 0));
        currentMemoryCache = payload;
        diskCache.set(key, payload, 5 * 60 * 1000);
        return payload;
    }

    async function saveMonthlySnapshot({ fs, month, diverse = [], currentOverride = null }) {
        const current = currentOverride && typeof currentOverride === 'object'
            ? currentOverride
            : currentMemoryCache;
        if (!current) throw new Error('Lagerliste cache er ikke klar. Tryk Opdater lagerliste først.');
        const payload = { month, createdAt: new Date().toISOString(), current, diverse };
        const file = path.join(snapshotDir, String(month || '').replace(/[^0-9-]/g, '') + '.json');
        setImmediate(() => {
            writeSnapshotFile(fs, month, payload)
                .catch(err => console.warn('[lagerliste] monthly snapshot write failed:', err.message));
        });
        return { ...payload, file };
    }

    async function savePointInTimeSnapshot({ fs, capturedAt, note = '', diverse = [], currentOverride = null, forceRefresh = true }) {
        const now = capturedAt instanceof Date && !Number.isNaN(capturedAt.getTime())
            ? capturedAt
            : new Date();
        const snapshotId = buildSnapshotId(now);
        const current = currentOverride && typeof currentOverride === 'object'
            ? currentOverride
            : await (async () => {
                const cached = currentMemoryCache;
                if (!cached) throw new Error('Lagerliste cache er ikke klar. Tryk Opdater lagerliste først.');
                return cached;
            })();
        const payload = {
            snapshotId,
            kind: 'point-in-time',
            capturedAt: now.toISOString(),
            createdAt: new Date().toISOString(),
            note: String(note || '').trim(),
            current,
            diverse
        };
        const file = path.join(historyDir, String(snapshotId || '').replace(/[^0-9A-Za-z_\-]/g, '') + '.json');
        setImmediate(() => {
            writePointInTimeSnapshotFile(fs, snapshotId, payload)
                .catch(err => console.warn('[lagerliste] point snapshot write failed:', err.message));
        });
        return { ...payload, file };
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

    return {
        getCurrent,
        saveMonthlySnapshot,
        loadMonthlySnapshot,
        savePointInTimeSnapshot,
        listPointInTimeSnapshots,
        loadPointInTimeSnapshot,
        scheduleMonthlySnapshot
    };
}

module.exports = { createLagerlisteService };
