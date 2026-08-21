// ── Belastning (capacity load) ──────────────────────────────────────────────
// Estratto verbatim da routes/apiRoutes.js: normalizzatori parametri e query
// SQL per belastning/grafisk e belastning/detail. Le firme sono identiche ai
// call site esistenti ({ getConnection, sql, ... }).

function parseBelastningDate(raw) {
    const txt = String(raw || '').trim();
    if (!txt) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(txt)) return txt;
    return null;
}

function parseBelastningDays(raw, fallback = 30) {
    const parsed = Number(raw);
    if (!Number.isFinite(parsed)) return fallback;
    return Math.max(1, Math.min(180, Math.round(parsed)));
}

function normalizeResGrCsv(raw) {
    return String(raw || '')
        .split(',')
        .map(x => x.trim())
        .filter(Boolean)
        .join(',');
}

function normalizeBelastningOrderFilter(raw) {
    return String(raw || '').replace(/\D+/g, '').slice(0, 12);
}

function normalizeBelastningCustomerFilter(raw) {
    return String(raw || '').trim().replace(/\s+/g, ' ').slice(0, 80);
}

function intCsvFromValues(values) {
    const seen = new Set();
    const out = [];
    for (const value of values || []) {
        const n = Number(value);
        if (!Number.isFinite(n)) continue;
        const i = Math.trunc(n);
        if (i <= 0) continue;
        const key = String(i);
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(key);
    }
    return out.join(',');
}

async function fetchBelastningRows({ getConnection, sql, toDay, dage, resGrCsv, parity, orderNo, customerFilter }) {
    const pool = await getConnection();
    const request = pool.request()
        .input('ToDay', sql.DateTime, new Date(toDay + 'T00:00:00'))
        .input('Dage', sql.Int, dage)
        .input('ResGr', sql.VarChar, resGrCsv)
        .input('Parity', sql.Int, parity)
        .input('OrderNo', sql.VarChar, String(orderNo || '').trim())
        .input('CustomerFilter', sql.VarChar, String(customerFilter || '').trim());

    const result = await request.query(`
        SET NOCOUNT ON;
        SET DATEFORMAT DMY;

        DECLARE @DD INT = dbo.EGD_Date2Int(@ToDay);
        DECLARE @ToDt INT = dbo.EGD_Date2Int(DATEADD(day, @Dage + 1, @ToDay));
        DECLARE @FrDt INT = dbo.EGD_Date2Int(DATEADD(day, -360, @ToDay));
        DECLARE @OrdNoInt INT = TRY_CONVERT(int, NULLIF(@OrderNo, ''));

        DECLARE @GemOrd INT;
        DECLARE @SOrdre INT;
        DECLARE @Dato datetime;
        DECLARE @POrdre INT;
        DECLARE @Antal float;
        DECLARE @Resv float;
        DECLARE @FM float;
        DECLARE @SumFM float;
        DECLARE @Aften float;
        DECLARE @ResGrX varchar(40);
        DECLARE @GemResGrX varchar(40);

        DECLARE @TmpFree TABLE
        (
            ID int identity(1,1),
            BeforeDD int,
            ResGr varchar(20),
            Dato datetime,
            Resv float,
            Antal float,
            FM float,
            SOrdre int,
            SLnNo int,
            POrdre int,
            AntalRec int,
            Aften float default 0
        );

        DECLARE @TmpKap TABLE
        (
            ID int identity(1,1),
            BeforeDD int,
            ResGr varchar(20),
            Dato datetime,
            Resv float,
            Antal float,
            FM float,
            SOrdre int,
            POrdre int,
            Kunde varchar(100),
            LevMode varchar(100),
            LevDato datetime,
            ULDato datetime,
            RestResv float default 0,
            Aften float default 0,
            RestAften float default 0
        );

        INSERT INTO @TmpFree (BeforeDD, ResGr, Dato, Resv, Antal, FM, SOrdre, POrdre, SLnNo, Aften)
        SELECT
            (CASE WHEN f.Dt1 < @DD THEN 1 ELSE 0 END) AS BeforeDD,
            R7.MainR7 AS ResGr,
            dbo.EGD_Int2Date(f.Dt1) AS Dato,
            SUM(CONVERT(float, ABS(f.Val1))) AS Resv,
            AVG(CONVERT(float, l.NoInvoAb)) AS Antal,
            AVG(CONVERT(float, l.NoFin)) AS FM,
            so.OrdNo AS SOrdre,
            po.OrdNo AS POrdre,
            l.LnNo AS SLnNo,
            SUM(CONVERT(float, ABS(CASE WHEN f.Txt4 <> '' THEN f.Val1 ELSE 0 END))) AS Aften
        FROM FreeInf1 f WITH(NOLOCK)
        INNER JOIN OrdLn l WITH(NOLOCK)
            ON f.OrdNo = l.OrdNo
           AND f.OrdLnNo = l.LnNo
        INNER JOIN Ord po WITH(NOLOCK)
            ON f.OrdNo = po.OrdNo
        INNER JOIN Ord so WITH(NOLOCK)
            ON po.OrdBasNo = so.OrdNo
        INNER JOIN R7 WITH(NOLOCK)
            ON f.R7 = R7.RNo
        WHERE f.FrInfTp = 2
          AND f.Val1 < 0
          AND l.ProdTp4 IN (1,3)
          AND l.TransGr3 < 80
          AND l.NoInvoAb >= 0
          AND f.Dt1 BETWEEN @FrDt AND @ToDt
          AND R7.Gr10 > 0
          AND (R7.Gr10 % 2) = @Parity
          AND (@ResGr = '' OR R7.MainR7 IN (SELECT LTRIM(RTRIM(value)) FROM string_split(@ResGr, ',')))
          AND (
                @OrdNoInt IS NULL
                OR so.OrdNo = @OrdNoInt
                OR po.OrdNo = @OrdNoInt
                OR f.OrdNo = @OrdNoInt
              )
          AND (
                @CustomerFilter = ''
                OR so.Nm LIKE '%' + @CustomerFilter + '%'
              )
        GROUP BY f.Dt1, l.LnNo, po.OrdNo, so.OrdNo, R7.MainR7;

        UPDATE t
        SET
            FM = FM / ISNULL((SELECT COUNT(*) FROM @TmpFree k WHERE k.POrdre = t.POrdre AND k.ResGr = t.ResGr AND k.SLnNo = t.SLnNo), 1),
            Antal = Antal / ISNULL((SELECT COUNT(*) FROM @TmpFree k WHERE k.POrdre = t.POrdre AND k.ResGr = t.ResGr AND k.SLnNo = t.SLnNo), 1),
            AntalRec = ISNULL((SELECT COUNT(*) FROM @TmpFree k WHERE k.POrdre = t.POrdre AND k.ResGr = t.ResGr AND k.SLnNo = t.SLnNo), 1)
        FROM @TmpFree t;

        INSERT INTO @TmpKap (BeforeDD, ResGr, Dato, Resv, Antal, FM, Aften, SOrdre, POrdre, Kunde, LevMode, LevDato, ULDato)
        SELECT
            x.BeforeDD,
            x.ResGr,
            x.Dato,
            SUM(x.Resv) AS Resv,
            SUM(x.Antal) AS Antal,
            SUM(x.FM) AS FM,
            SUM(x.Aften) AS Aften,
            x.SOrdre,
            x.POrdre,
            so.Nm AS Kunde,
            ISNULL((SELECT TOP (1) Txt FROM Txt WITH(NOLOCK) WHERE Lang = 45 AND TxtTp = 5 AND TxtNo = so.DelMt), '') AS LevMode,
            (CASE WHEN so.DelDt > 19800101 THEN dbo.EGD_Int2Date(so.DelDt) ELSE NULL END) AS LevDato,
            (CASE WHEN po.ArDt > 19800101 THEN dbo.EGD_Int2Date(po.ArDt) ELSE NULL END) AS ULDato
        FROM @TmpFree x
        INNER JOIN Ord po WITH(NOLOCK)
            ON x.POrdre = po.OrdNo
        INNER JOIN Ord so WITH(NOLOCK)
            ON po.OrdBasNo = so.OrdNo
        GROUP BY x.BeforeDD, x.ResGr, x.Dato, x.SOrdre, x.POrdre, so.Nm, so.DelMt, so.DelDt, po.ArDt;

        DECLARE Tmp_cursor CURSOR STATIC FOR
        SELECT SOrdre, Dato, MIN(POrdre) AS POrdre, SUM(Antal) AS Antal, SUM(Resv) AS Resv, SUM(FM) AS FM, SUM(Aften) AS Aften, ResGr
        FROM @TmpKap
        GROUP BY SOrdre, ResGr, Dato
        ORDER BY SOrdre, ResGr, Dato;

        OPEN Tmp_cursor;
        FETCH NEXT FROM Tmp_cursor INTO @SOrdre, @Dato, @POrdre, @Antal, @Resv, @FM, @Aften, @ResGrX;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            IF @GemOrd IS NULL OR @GemOrd <> @SOrdre OR @GemResGrX <> @ResGrX
            BEGIN
                SET @GemOrd = @SOrdre;
                SET @GemResGrX = @ResGrX;
                SET @SumFM = ISNULL((SELECT SUM(FM) FROM @TmpKap WHERE SOrdre = @GemOrd AND ResGr = @ResGrX), 0);
            END

            IF @Antal > 0
            BEGIN
                IF @SumFM > @Resv
                BEGIN
                    SET @SumFM = @SumFM - @Resv;
                    SET @Resv = 0;
                    SET @Aften = 0;
                END
                ELSE IF @Resv > @SumFM
                BEGIN
                    SET @Resv = @Resv - @SumFM;

                    IF @Aften > 0 AND @Aften = @Resv
                        SET @Aften = @Aften;
                    ELSE IF @Aften > @SumFM AND @SumFM > 0
                        SET @Aften = @Aften - @SumFM;
                    ELSE
                        SET @Aften = 0;

                    SET @SumFM = 0;
                END

                UPDATE TOP (1) @TmpKap
                SET RestResv = @Resv,
                    RestAften = @Aften
                WHERE SOrdre = @SOrdre
                  AND Dato = @Dato
                  AND ResGr = @ResGrX;
            END

            FETCH NEXT FROM Tmp_cursor INTO @SOrdre, @Dato, @POrdre, @Antal, @Resv, @FM, @Aften, @ResGrX;
        END

        CLOSE Tmp_cursor;
        DEALLOCATE Tmp_cursor;

        SELECT
            x.ResGr,
            x.Dato,
            LEFT(CONVERT(varchar, x.Dato, 103), 5) AS DatoX,
            x.Nm,
            SUM(x.Resv) AS Resv,
            SUM(x.Kap) AS Kap,
            SUM(x.Aften) AS Aften,
            DENSE_RANK() OVER (ORDER BY x.ResGr) AS Ranking
        FROM (
            SELECT
                t.ResGr,
                (CASE WHEN dbo.EGD_Date2Int(t.Dato) < @DD THEN NULL ELSE t.Dato END) AS Dato,
                SUM(t.RestResv - t.RestAften) AS Resv,
                0 AS Kap,
                SUM(t.RestAften) AS Aften,
                (SELECT Nm FROM R7 WHERE RNo = t.ResGr) AS Nm
            FROM @TmpKap t
            WHERE (t.Antal > 0 OR (t.Antal = 0 AND t.RestResv > 0))
              AND (@OrdNoInt IS NULL OR t.SOrdre = @OrdNoInt OR t.POrdre = @OrdNoInt)
              AND (t.LevDato >= @ToDay OR ((t.LevDato IS NULL OR t.LevDato < @ToDay) AND t.RestResv > 0))
            GROUP BY t.ResGr, (CASE WHEN dbo.EGD_Date2Int(t.Dato) < @DD THEN NULL ELSE t.Dato END)

            UNION ALL

            SELECT
                R7.MainR7 AS ResGr,
                dbo.EGD_Int2Date(f.Dt1) AS Dato,
                0 AS Resv,
                CONVERT(float, ABS(SUM(f.Val1 * R7.Am1))) AS Kap,
                0 AS Aften,
                R7.Nm AS Nm
            FROM FreeInf1 f WITH(NOLOCK)
            INNER JOIN R7 WITH(NOLOCK)
                ON f.R7 = R7.RNo
            WHERE f.FrInfTp = 1
              AND f.Dt1 BETWEEN @DD AND @ToDt
              AND R7.Gr10 > 0
              AND (R7.Gr10 % 2) = @Parity
              AND (@ResGr = '' OR R7.MainR7 IN (SELECT LTRIM(RTRIM(value)) FROM string_split(@ResGr, ',')))
            GROUP BY f.Dt1, R7.Nm, R7.MainR7
            HAVING ABS(SUM(f.Val1 * R7.Am1)) > 0
        ) x
        GROUP BY x.ResGr, x.Dato, x.Nm
        ORDER BY x.ResGr, x.Dato;
    `);

    return Array.isArray(result.recordset) ? result.recordset : [];
}

async function fetchBelastningOrderRows({ getConnection, sql, toDay, dage, resGrCsv, parity, orderNo, customerFilter }) {
    const pool = await getConnection();
    const request = pool.request()
        .input('ToDay', sql.DateTime, new Date(toDay + 'T00:00:00'))
        .input('Dage', sql.Int, dage)
        .input('ResGr', sql.VarChar, resGrCsv)
        .input('Parity', sql.Int, parity)
        .input('OrderNo', sql.VarChar, String(orderNo || '').trim())
        .input('CustomerFilter', sql.VarChar, String(customerFilter || '').trim());

    const result = await request.query(`
        SET NOCOUNT ON;
        SET DATEFORMAT DMY;

        DECLARE @DD INT = dbo.EGD_Date2Int(@ToDay);
        DECLARE @ToDt INT = dbo.EGD_Date2Int(DATEADD(day, @Dage + 1, @ToDay));
        DECLARE @FrDt INT = dbo.EGD_Date2Int(DATEADD(day, -360, @ToDay));
        DECLARE @OrdNoInt INT = TRY_CONVERT(int, NULLIF(@OrderNo, ''));

        DECLARE @GemOrd INT;
        DECLARE @SOrdre INT;
        DECLARE @Dato datetime;
        DECLARE @POrdre INT;
        DECLARE @Antal float;
        DECLARE @Resv float;
        DECLARE @FM float;
        DECLARE @SumFM float;
        DECLARE @Aften float;
        DECLARE @ResGrX varchar(40);
        DECLARE @GemResGrX varchar(40);

        DECLARE @TmpFree TABLE
        (
            ID int identity(1,1),
            BeforeDD int,
            ResGr varchar(20),
            Dato datetime,
            Resv float,
            Antal float,
            FM float,
            SOrdre int,
            SLnNo int,
            POrdre int,
            AntalRec int,
            Aften float default 0
        );

        DECLARE @TmpKap TABLE
        (
            ID int identity(1,1),
            BeforeDD int,
            ResGr varchar(20),
            Dato datetime,
            Resv float,
            Antal float,
            FM float,
            SOrdre int,
            POrdre int,
            Kunde varchar(100),
            LevMode varchar(100),
            LevDato datetime,
            ULDato datetime,
            RestResv float default 0,
            Aften float default 0,
            RestAften float default 0
        );

        INSERT INTO @TmpFree (BeforeDD, ResGr, Dato, Resv, Antal, FM, SOrdre, POrdre, SLnNo, Aften)
        SELECT
            (CASE WHEN f.Dt1 < @DD THEN 1 ELSE 0 END) AS BeforeDD,
            R7.MainR7 AS ResGr,
            dbo.EGD_Int2Date(f.Dt1) AS Dato,
            SUM(CONVERT(float, ABS(f.Val1))) AS Resv,
            AVG(CONVERT(float, l.NoInvoAb)) AS Antal,
            AVG(CONVERT(float, l.NoFin)) AS FM,
            so.OrdNo AS SOrdre,
            po.OrdNo AS POrdre,
            l.LnNo AS SLnNo,
            SUM(CONVERT(float, ABS(CASE WHEN f.Txt4 <> '' THEN f.Val1 ELSE 0 END))) AS Aften
        FROM FreeInf1 f WITH(NOLOCK)
        INNER JOIN OrdLn l WITH(NOLOCK)
            ON f.OrdNo = l.OrdNo
           AND f.OrdLnNo = l.LnNo
        INNER JOIN Ord po WITH(NOLOCK)
            ON f.OrdNo = po.OrdNo
        INNER JOIN Ord so WITH(NOLOCK)
            ON po.OrdBasNo = so.OrdNo
        INNER JOIN R7 WITH(NOLOCK)
            ON f.R7 = R7.RNo
        WHERE f.FrInfTp = 2
          AND f.Val1 < 0
          AND l.ProdTp4 IN (1,3)
          AND l.TransGr3 < 80
          AND l.NoInvoAb >= 0
          AND f.Dt1 BETWEEN @FrDt AND @ToDt
          AND R7.Gr10 > 0
          AND (R7.Gr10 % 2) = @Parity
          AND (@ResGr = '' OR R7.MainR7 IN (SELECT LTRIM(RTRIM(value)) FROM string_split(@ResGr, ',')))
          AND (
                @OrdNoInt IS NULL
                OR so.OrdNo = @OrdNoInt
                OR po.OrdNo = @OrdNoInt
                OR f.OrdNo = @OrdNoInt
              )
          AND (
                @CustomerFilter = ''
                OR so.Nm LIKE '%' + @CustomerFilter + '%'
              )
        GROUP BY f.Dt1, l.LnNo, po.OrdNo, so.OrdNo, R7.MainR7;

        UPDATE t
        SET
            FM = FM / ISNULL((SELECT COUNT(*) FROM @TmpFree k WHERE k.POrdre = t.POrdre AND k.ResGr = t.ResGr AND k.SLnNo = t.SLnNo), 1),
            Antal = Antal / ISNULL((SELECT COUNT(*) FROM @TmpFree k WHERE k.POrdre = t.POrdre AND k.ResGr = t.ResGr AND k.SLnNo = t.SLnNo), 1),
            AntalRec = ISNULL((SELECT COUNT(*) FROM @TmpFree k WHERE k.POrdre = t.POrdre AND k.ResGr = t.ResGr AND k.SLnNo = t.SLnNo), 1)
        FROM @TmpFree t;

        INSERT INTO @TmpKap (BeforeDD, ResGr, Dato, Resv, Antal, FM, Aften, SOrdre, POrdre, Kunde, LevMode, LevDato, ULDato)
        SELECT
            x.BeforeDD,
            x.ResGr,
            x.Dato,
            SUM(x.Resv) AS Resv,
            SUM(x.Antal) AS Antal,
            SUM(x.FM) AS FM,
            SUM(x.Aften) AS Aften,
            x.SOrdre,
            x.POrdre,
            so.Nm AS Kunde,
            ISNULL((SELECT TOP (1) Txt FROM Txt WITH(NOLOCK) WHERE Lang = 45 AND TxtTp = 5 AND TxtNo = so.DelMt), '') AS LevMode,
            (CASE WHEN so.DelDt > 19800101 THEN dbo.EGD_Int2Date(so.DelDt) ELSE NULL END) AS LevDato,
            (CASE WHEN po.ArDt > 19800101 THEN dbo.EGD_Int2Date(po.ArDt) ELSE NULL END) AS ULDato
        FROM @TmpFree x
        INNER JOIN Ord po WITH(NOLOCK)
            ON x.POrdre = po.OrdNo
        INNER JOIN Ord so WITH(NOLOCK)
            ON po.OrdBasNo = so.OrdNo
        GROUP BY x.BeforeDD, x.ResGr, x.Dato, x.SOrdre, x.POrdre, so.Nm, so.DelMt, so.DelDt, po.ArDt;

        DECLARE Tmp_cursor CURSOR STATIC FOR
        SELECT SOrdre, Dato, MIN(POrdre) AS POrdre, SUM(Antal) AS Antal, SUM(Resv) AS Resv, SUM(FM) AS FM, SUM(Aften) AS Aften, ResGr
        FROM @TmpKap
        GROUP BY SOrdre, ResGr, Dato
        ORDER BY SOrdre, ResGr, Dato;

        OPEN Tmp_cursor;
        FETCH NEXT FROM Tmp_cursor INTO @SOrdre, @Dato, @POrdre, @Antal, @Resv, @FM, @Aften, @ResGrX;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            IF @GemOrd IS NULL OR @GemOrd <> @SOrdre OR @GemResGrX <> @ResGrX
            BEGIN
                SET @GemOrd = @SOrdre;
                SET @GemResGrX = @ResGrX;
                SET @SumFM = ISNULL((SELECT SUM(FM) FROM @TmpKap WHERE SOrdre = @GemOrd AND ResGr = @ResGrX), 0);
            END

            IF @Antal > 0
            BEGIN
                IF @SumFM > @Resv
                BEGIN
                    SET @SumFM = @SumFM - @Resv;
                    SET @Resv = 0;
                    SET @Aften = 0;
                END
                ELSE IF @Resv > @SumFM
                BEGIN
                    SET @Resv = @Resv - @SumFM;

                    IF @Aften > 0 AND @Aften = @Resv
                        SET @Aften = @Aften;
                    ELSE IF @Aften > @SumFM AND @SumFM > 0
                        SET @Aften = @Aften - @SumFM;
                    ELSE
                        SET @Aften = 0;

                    SET @SumFM = 0;
                END

                UPDATE TOP (1) @TmpKap
                SET RestResv = @Resv,
                    RestAften = @Aften
                WHERE SOrdre = @SOrdre
                  AND Dato = @Dato
                  AND ResGr = @ResGrX;
            END

            FETCH NEXT FROM Tmp_cursor INTO @SOrdre, @Dato, @POrdre, @Antal, @Resv, @FM, @Aften, @ResGrX;
        END

        CLOSE Tmp_cursor;
        DEALLOCATE Tmp_cursor;

        SELECT
            t.ResGr,
            (SELECT Nm FROM R7 WHERE RNo = t.ResGr) AS Nm,
            t.Dato,
            LEFT(CONVERT(varchar, t.Dato, 103), 10) AS DatoX,
            t.SOrdre,
            t.POrdre,
            t.POrdre AS OrdNo,
            t.POrdre AS PurcNo,
            t.Kunde,
            t.LevMode,
            t.LevDato,
            t.ULDato,
            CONVERT(float, t.Resv) AS Resv,
            CONVERT(float, t.RestResv) AS RestResv,
            CONVERT(float, t.Aften) AS Aften,
            CONVERT(float, t.RestAften) AS RestAften,
            CONVERT(float, t.Resv) AS ResvRaw,
            CONVERT(float, t.Aften) AS AftenRaw,
            CONVERT(float, t.RestResv) AS ResvNet,
            ISNULL((
                SELECT STUFF((
                    SELECT '-' + LEFT(R7.Nm, 3)
                    FROM OrdLn l WITH(NOLOCK)
                    INNER JOIN R7 WITH(NOLOCK)
                        ON l.R7 = R7.RNo
                    WHERE l.OrdNo = t.POrdre
                      AND l.ProdTp4 = 1
                    ORDER BY l.LnNo
                    FOR XML PATH('')
                ), 1, 1, '')
            ), '') AS Opr
        FROM @TmpKap t
        WHERE (t.Antal > 0 OR (t.Antal = 0 AND t.RestResv > 0))
          AND (@OrdNoInt IS NULL OR t.SOrdre = @OrdNoInt OR t.POrdre = @OrdNoInt)
          AND (t.LevDato >= @ToDay OR ((t.LevDato IS NULL OR t.LevDato < @ToDay) AND t.RestResv > 0))
        ORDER BY t.Dato, t.ResGr, t.POrdre, t.SOrdre;
    `);

    return Array.isArray(result.recordset) ? result.recordset : [];
}

async function fetchBelastningSubOrderRows({ getConnection, sql, subOrderCsv }) {
    if (!String(subOrderCsv || '').trim()) return [];

    const pool = await getConnection();
    const request = pool.request()
        .input('SubOrderCsv', sql.VarChar, subOrderCsv);

    const result = await request.query(`
        SET NOCOUNT ON;

        ;WITH SubOrders AS (
            SELECT DISTINCT TRY_CONVERT(int, LTRIM(RTRIM(value))) AS OrdNo
            FROM string_split(@SubOrderCsv, ',')
            WHERE TRY_CONVERT(int, LTRIM(RTRIM(value))) IS NOT NULL
        )
        SELECT
            l.OrdNo AS SubOrdNo,
            l.LnNo AS SubLnNo,
            l.ProdNo AS SubProdNo,
            l.Descr AS SubDescr,
            l.ProdTp4 AS SubProdTp4,
            l.PurcNo AS NextSubOrdNo,
            l.NoFin AS SubNoFin,
            l.NoOrg AS SubNoOrg,
            l.CCstPr AS SubCcstPr,
            l.DPrice AS SubDPrice
        FROM OrdLn l WITH(NOLOCK)
        INNER JOIN SubOrders s
            ON s.OrdNo = l.OrdNo
        ORDER BY l.OrdNo, l.LnNo;
    `);

    return Array.isArray(result.recordset) ? result.recordset : [];
}

async function fetchBelastningOrderLineRows({ getConnection, sql, orderCsv }) {
    if (!String(orderCsv || '').trim()) return [];

    const pool = await getConnection();
    const request = pool.request()
        .input('OrderCsv', sql.VarChar, orderCsv);

    const result = await request.query(`
        SET NOCOUNT ON;

        ;WITH SourceOrders AS (
            SELECT DISTINCT TRY_CONVERT(int, LTRIM(RTRIM(value))) AS OrdNo
            FROM string_split(@OrderCsv, ',')
            WHERE TRY_CONVERT(int, LTRIM(RTRIM(value))) IS NOT NULL
        )
        SELECT
            l.OrdNo,
            l.LnNo,
            l.ProdNo,
            l.Descr,
            l.ProdTp4,
            l.PurcNo,
            l.NoFin,
            l.NoOrg,
            l.CCstPr,
            l.DPrice
        FROM OrdLn l WITH(NOLOCK)
        INNER JOIN SourceOrders s
            ON s.OrdNo = l.OrdNo
        ORDER BY l.OrdNo, l.LnNo;
    `);

    return Array.isArray(result.recordset) ? result.recordset : [];
}

module.exports = {
    parseBelastningDate,
    parseBelastningDays,
    normalizeResGrCsv,
    normalizeBelastningOrderFilter,
    normalizeBelastningCustomerFilter,
    intCsvFromValues,
    fetchBelastningRows,
    fetchBelastningOrderRows,
    fetchBelastningSubOrderRows,
    fetchBelastningOrderLineRows
};
