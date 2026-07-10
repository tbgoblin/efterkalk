function createOrdreindgangService({ getConnection, sql }) {
    function isValidWeekKey(value) {
        const raw = String(value || '').trim();
        const match = raw.match(/^(\d{4})(\d{2})$/);
        if (!match) return false;
        const week = Number(match[2]);
        return week >= 1 && week <= 53;
    }

    function sanitizeWeekRange(fraWeek, tilWeek) {
        const fraRaw = String(fraWeek || '').trim();
        const tilRaw = String(tilWeek || '').trim();

        if (!isValidWeekKey(fraRaw) || !isValidWeekKey(tilRaw)) {
            const error = new Error('Ugyldig ugeperiode. Brug format YYYYWW.');
            error.statusCode = 400;
            throw error;
        }

        const fromWeek = Number(fraRaw);
        const toWeek = Number(tilRaw);

        if (toWeek < fromWeek) {
            const error = new Error('Til-uge skal være samme eller efter fra-uge.');
            error.statusCode = 400;
            throw error;
        }

        return { fromWeek, toWeek, fraRaw, tilRaw };
    }

    async function getSummary({ fraWeek, tilWeek }) {
        const range = sanitizeWeekRange(fraWeek, tilWeek);
        const pool = await getConnection();

        // Source of truth: SQL copied from Ordreindgang.rdl CommandText.
        const weeklyResult = await pool.request()
            .input('fra', sql.Int, range.fromWeek)
            .input('til', sql.Int, range.toWeek)
            .query(`
                ;WITH OrdSum AS (
                    SELECT
                        FreeInf2.Val8,
                        SUM((Ord.InvoSF + Ord.InvoIF) * (Ord.ExRt / 100.0)) / 1000.0 AS SumOrd
                    FROM F0001.dbo.FreeInf2 FreeInf2
                    INNER JOIN F0001.dbo.Ord Ord ON Ord.OrdDt = FreeInf2.Dt1
                    WHERE
                        Ord.TrTp = 1
                        AND Ord.OrdTp = 1
                        AND FreeInf2.FrInfTp = 550
                        AND FreeInf2.Val8 >= @fra
                        AND FreeInf2.Val8 <= @til
                    GROUP BY FreeInf2.Val8
                ),
                TilbudSum AS (
                    SELECT
                        FreeInf2.Val8,
                        SUM((Ord.InvoSF + Ord.InvoIF) * (Ord.ExRt / 100.0) * Ord.Free1 / 100.0) / 1000.0 AS SumTilbud
                    FROM F0001.dbo.FreeInf2 FreeInf2
                    INNER JOIN F0001.dbo.Ord Ord ON Ord.OrdDt = FreeInf2.Dt1
                    WHERE
                        Ord.TrTp = 1
                        AND Ord.OrdTp = 5
                        AND FreeInf2.FrInfTp = 550
                        AND FreeInf2.Val8 >= @fra
                        AND FreeInf2.Val8 <= @til
                    GROUP BY FreeInf2.Val8
                ),
                Budget AS (
                    SELECT
                        FreeInf2.Val8,
                        SUM(FreeInf2.Val9 * FreeInf2.Val10) AS Budget,
                        SUM(FreeInf2.Val10) AS Val10
                    FROM F0001.dbo.FreeInf2 FreeInf2
                    WHERE
                        FreeInf2.FrInfTp = 550
                        AND FreeInf2.Val8 >= @fra
                        AND FreeInf2.Val8 <= @til
                    GROUP BY FreeInf2.Val8
                ),
                P9_Value AS (
                    SELECT
                        SUM((Ord.InvoSF + Ord.InvoIF) * (Ord.ExRt / 100.0)) / NULLIF(SUM(FreeInf2.Val10), 0) AS P9
                    FROM F0001.dbo.FreeInf2 FreeInf2
                    INNER JOIN F0001.dbo.Ord Ord ON Ord.OrdDt = FreeInf2.Dt1
                    WHERE
                        Ord.TrTp = 1
                        AND Ord.OrdTp = 1
                        AND FreeInf2.FrInfTp = 550
                        AND FreeInf2.Txt1 = '2023/24'
                        AND FreeInf2.Val8 <= @til
                ),
                OrdVariazione AS (
                    SELECT
                        FreeInf2.Val8,
                        SUM((Ord.InvoSF + Ord.InvoIF) * (Ord.ExRt / 100.0)) / 1000.0 AS OrdCurrent,
                        LAG(SUM((Ord.InvoSF + Ord.InvoIF) * (Ord.ExRt / 100.0)) / 1000.0, 1)
                            OVER (ORDER BY FreeInf2.Val8) AS OrdPrevious
                    FROM F0001.dbo.FreeInf2 FreeInf2
                    INNER JOIN F0001.dbo.Ord Ord ON Ord.OrdDt = FreeInf2.Dt1
                    WHERE
                        Ord.TrTp = 1
                        AND Ord.OrdTp = 1
                        AND FreeInf2.FrInfTp = 550
                        AND FreeInf2.Val8 >= @fra
                        AND FreeInf2.Val8 <= @til
                    GROUP BY FreeInf2.Val8
                ),
                AverageValue AS (
                    SELECT AVG(SumOrd) AS AvgSumOrd
                    FROM OrdSum
                )
                SELECT
                    B.Val8,
                    O.SumOrd AS TotalOrd,
                    T.SumTilbud AS TotalTilbud,
                    B.Budget,
                    (P.P9 * B.Val10 / 100.0) AS CalculatedValue,
                    CASE
                        WHEN OV.OrdPrevious IS NOT NULL AND OV.OrdPrevious > 0
                            THEN ((OV.OrdCurrent - OV.OrdPrevious) / OV.OrdPrevious) * 100.0
                        ELSE NULL
                    END AS VariazionePercentualeOrd,
                    A.AvgSumOrd
                FROM Budget B
                LEFT JOIN OrdSum O ON B.Val8 = O.Val8
                LEFT JOIN TilbudSum T ON B.Val8 = T.Val8
                LEFT JOIN OrdVariazione OV ON B.Val8 = OV.Val8
                CROSS JOIN P9_Value P
                CROSS JOIN AverageValue A
                ORDER BY B.Val8;
            `);

        // Keep customer table in app, but align week filtering with FreeInf2.Val8 range.
        const customerResult = await pool.request()
            .input('fromWeek', sql.Int, range.fromWeek)
            .input('toWeek', sql.Int, range.toWeek)
            .query(`
                WITH deduped AS (
                    SELECT DISTINCT
                        o.OrdNo,
                        o.CustNo,
                        LTRIM(RTRIM(ISNULL(a.Nm, ''))) AS CustNm,
                        CASE
                            WHEN o.TrTp = 1 AND o.OrdTp = 1 THEN 'ORDRE'
                            WHEN o.TrTp = 1 AND o.OrdTp = 5 THEN 'TILBUD'
                            ELSE 'OTHER'
                        END AS Bucket,
                        (o.InvoSF + o.InvoIF) * (o.ExRt / 100.0) AS OrdAmt
                    FROM F0001.dbo.Ord o
                    INNER JOIN F0001.dbo.FreeInf2 FreeInf2
                        ON o.OrdDt = FreeInf2.Dt1
                    LEFT JOIN F0001.dbo.Actor a
                        ON o.CustNo = a.CustNo
                    WHERE
                        FreeInf2.FrInfTp = 550
                        AND FreeInf2.Val8 >= @fromWeek
                        AND FreeInf2.Val8 <= @toWeek
                        AND o.TrTp = 1
                        AND o.OrdTp IN (1, 5)
                )
                SELECT TOP (50)
                    CustNo,
                    MAX(CustNm) AS CustNm,
                    SUM(CASE WHEN Bucket = 'ORDRE'  THEN OrdAmt ELSE 0 END) / 1000.0 AS OrdSum,
                    SUM(CASE WHEN Bucket = 'TILBUD' THEN OrdAmt ELSE 0 END) / 1000.0 AS TilbudSum,
                    SUM(CASE WHEN Bucket = 'ORDRE'  THEN 1 ELSE 0 END) AS CntOrd,
                    SUM(CASE WHEN Bucket = 'TILBUD' THEN 1 ELSE 0 END) AS CntTilbud
                FROM deduped
                GROUP BY CustNo
                HAVING SUM(CASE WHEN Bucket = 'ORDRE'  THEN OrdAmt ELSE 0 END) <> 0
                    OR SUM(CASE WHEN Bucket = 'TILBUD' THEN OrdAmt ELSE 0 END) <> 0
                ORDER BY SUM(CASE WHEN Bucket = 'ORDRE' THEN OrdAmt ELSE 0 END) DESC;
            `);

        const weeklyRows = (weeklyResult.recordset || []).map(row => {
            const rawWeek = String(row.Val8 || '').trim();
            const weekKey = /^\d{6}$/.test(rawWeek) ? rawWeek : rawWeek;
            const year = /^\d{6}$/.test(weekKey) ? Number(weekKey.slice(0, 4)) : null;
            const week = /^\d{6}$/.test(weekKey) ? Number(weekKey.slice(4, 6)) : null;

            return {
                year,
                week,
                weekKey,
                totalOrd: Number(row.TotalOrd || 0),
                totalTilbud: Number(row.TotalTilbud || 0),
                totalBudget: Number(row.Budget || 0),
                avgOrd: Number(row.AvgSumOrd || 0),
                countOrd: 0,
                countTilbud: 0,
                calculatedValue: Number(row.CalculatedValue || 0),
                variationPctOrd: row.VariazionePercentualeOrd === null || row.VariazionePercentualeOrd === undefined
                    ? null
                    : Number(row.VariazionePercentualeOrd)
            };
        });

        const customers = (customerResult.recordset || []).map(row => {
            const ordSum = Number(row.OrdSum || 0);
            const tilbudSum = Number(row.TilbudSum || 0);
            const countOrd = Number(row.CntOrd || 0);
            const countTilbud = Number(row.CntTilbud || 0);
            const conversionPct = tilbudSum > 0 ? (ordSum / tilbudSum) * 100 : null;

            return {
                custNo: Number(row.CustNo || 0),
                customerName: String(row.CustNm || '').trim(),
                ordSum,
                tilbudSum,
                countOrd,
                countTilbud,
                conversionPct
            };
        });

        const totalOrdSum = weeklyRows.reduce((sum, row) => sum + Number(row.totalOrd || 0), 0);
        const totalTilbudSum = weeklyRows.reduce((sum, row) => sum + Number(row.totalTilbud || 0), 0);
        const totalBudgetSum = weeklyRows.reduce((sum, row) => sum + Number(row.totalBudget || 0), 0);

        // SSRS textbox uses Avg(Fields!TotalOrd.Value) while chart uses AvgSumOrd from SQL.
        // For parity with the chart/toggle series, expose AvgSumOrd from query.
        const avgSumOrd = weeklyRows.length > 0 ? Number(weeklyRows[0].avgOrd || 0) : 0;
        const conversionPct = totalTilbudSum > 0 ? (totalOrdSum / totalTilbudSum) * 100 : null;

        return {
            filters: {
                fraWeek: range.fraRaw,
                tilWeek: range.tilRaw
            },
            kpis: {
                totalOrdSum,
                totalTilbudSum,
                totalBudgetSum,
                avgSumOrd,
                conversionPct,
                totalOrdCount: 0,
                totalTilbudCount: 0
            },
            weeklyRows,
            customerRows: customers
        };
    }

    return {
        getSummary
    };
}

module.exports = {
    createOrdreindgangService
};
