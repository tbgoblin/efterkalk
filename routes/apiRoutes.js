const express = require('express');
const http = require('http');
const path = require('path');
const orderNotesService = require('../services/orderNotesService');

// ── Personalehåndbog crawler ────────────────────────────────────────────────
let phIndex   = [];          // [{url, title, text}]
let phStatus  = 'idle';      // 'idle' | 'indexing' | 'ready' | 'error'
let phIndexedAt = null;
let phError   = null;
const PH_BASE = 'http://apv/GHB/';

function phFetch(url) {
    return new Promise((resolve, reject) => {
        const req = http.get(url, { headers: { 'User-Agent': 'Gantech-Crawler/1.0' } }, (res) => {
            if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
                return resolve({ redirect: res.headers.location });
            }
            let body = '';
            res.setEncoding('utf8');
            res.on('data', c => body += c);
            res.on('end', () => resolve({ body, status: res.statusCode }));
        });
        req.setTimeout(10000, () => { req.destroy(); reject(new Error('timeout')); });
        req.on('error', reject);
    });
}

function phLinks(html, base) {
    const out = new Set();
    const re = /href=["']([^"']+)["']/gi;
    let m;
    while ((m = re.exec(html)) !== null) {
        try {
            const abs = new URL(m[1], base).href.split('?')[0].split('#')[0];
            if (abs.startsWith(PH_BASE) && !abs.match(/\.(jpg|jpeg|png|gif|svg|pdf|zip|docx?|xlsx?|css|js|ico|woff2?)$/i)) {
                out.add(abs);
            }
        } catch {}
    }
    return [...out];
}

function phTitle(html) {
    const m = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    return m ? m[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim() : '';
}

function phText(html) {
    return html
        .replace(/<script[\s\S]*?<\/script>/gi, ' ')
        .replace(/<style[\s\S]*?<\/style>/gi, ' ')
        .replace(/<nav[\s\S]*?<\/nav>/gi, ' ')
        .replace(/<header[\s\S]*?<\/header>/gi, ' ')
        .replace(/<footer[\s\S]*?<\/footer>/gi, ' ')
        .replace(/<[^>]+>/g, ' ')
        .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#\d+;/g, ' ')
        .replace(/\s+/g, ' ').trim();
}

async function crawlPH() {
    if (phStatus === 'indexing') return;
    phStatus = 'indexing';
    phError = null;
    phIndex = [];
    const visited = new Set();
    const queue = [PH_BASE];
    console.log('[PH-CRAWL] Starting crawl of', PH_BASE);
    while (queue.length > 0) {
        const url = queue.shift();
        if (visited.has(url)) continue;
        visited.add(url);
        try {
            const r = await phFetch(url);
            if (r.redirect) {
                try {
                    const abs = new URL(r.redirect, url).href.split('?')[0].split('#')[0];
                    if (abs.startsWith(PH_BASE) && !visited.has(abs)) queue.push(abs);
                } catch {}
                continue;
            }
            if (r.status !== 200) continue;
            const title = phTitle(r.body);
            const text  = phText(r.body);
            phIndex.push({ url, title, text });
            for (const link of phLinks(r.body, url)) {
                if (!visited.has(link)) queue.push(link);
            }
        } catch { /* skip unreachable page */ }
    }
    phStatus = 'ready';
    phIndexedAt = new Date().toISOString();
    console.log('[PH-CRAWL] Done:', phIndex.length, 'pages indexed');
}

// Start crawl in background after module load
setTimeout(() => crawlPH().catch(e => { phStatus = 'error'; phError = e.message; console.error('[PH-CRAWL] Error:', e.message); }), 3000);
// ────────────────────────────────────────────────────────────────────────────

const QMS_DATASET_PATH = path.join(__dirname, '..', 'data', 'qms-dataset.json');

function defaultQmsDataset() {
    return {
        version: 1,
        updatedAt: new Date().toISOString(),
        folders: [
            {
                id: 'startside',
                name: 'Startside',
                description: 'Overordnet introduktion til kvalitetsledelsessystemet',
                documents: [
                    {
                        id: 'qfp-00',
                        title: 'QFP-00 Kvalitetsledelsessystemets startsidestartside',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-00%20Kvalitetsledelsessystemets%20startsidestartside.aspx',
                        content: 'Kvalitetsledelsessystemet bygger på ISO 9001-principper med procesorienteret tilgang. Procesflowet er opdelt i teknisk forberedelse, fabrikation/service samt fakturering og sagsafslutning.',
                        tags: ['ISO9001', 'procesflow'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'q-001',
                        title: 'Q-001 Kvalitetsledelsessystemet - Gantech håndbogen',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/Q-001%20Kvalitetsledelsessystemet%20-%20Gantech%20h%C3%A5ndbogen.aspx',
                        content: 'Forord og ramme for samspillet mellem kvalitetsledelsessystem, personalehåndbog, arbejdsmiljøportal og serviceportal.',
                        tags: ['forord', 'styring'],
                        updatedAt: new Date().toISOString()
                    }
                ]
            },
            {
                id: 'ledelse',
                name: 'Ledelse',
                description: 'Politikker, retningslinjer og administration',
                documents: [
                    {
                        id: 'qfp-01',
                        title: 'QFP-01 Politikker og retningslinjer',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-01%20Politikker%20og%20retningslinjer.aspx',
                        content: 'Overblik over certificeringer, godkendelser og overordnede politikker som virksomhedens styrende dokumentgrundlag.',
                        tags: ['politik', 'certificering'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-02',
                        title: 'QFP-02 Administration',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-02%20Administration.aspx',
                        content: 'Årlig revision af administrative systemer og rutiner for at sikre effektivt workflow og tydelig styring.',
                        tags: ['administration'],
                        updatedAt: new Date().toISOString()
                    }
                ]
            },
            {
                id: 'procesflow',
                name: 'Procesflow',
                description: 'Forespørgsel til produktion, levering og fakturering',
                documents: [
                    {
                        id: 'qfp-15',
                        title: 'QFP-15 Forespørgsel',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-15%20Foresp%C3%B8rgsel.aspx',
                        content: 'Proces for vurdering og håndtering af kundeforepørgsler med systematisk sagsbehandling.',
                        tags: ['forespørgsel'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-16',
                        title: 'QFP-16 Tilbud',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-16%20Tilbud.aspx',
                        content: 'Tilbudsproces med kravspecifikation og tekniske afklaringer.',
                        tags: ['tilbud'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-17',
                        title: 'QFP-17 Ordre eller kontrakt gennemgang',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-17%20Ordre%20eller%20kontrakt%20gennemgang.aspx',
                        content: 'Ordre-/kontraktgennemgang med formel validering af krav før igangsættelse.',
                        tags: ['ordre', 'kontrakt'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-18',
                        title: 'QFP-18 Produktions forberedelse',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-18%20Produktions%20forberedelse.aspx',
                        content: 'Planlægning, produktionsværktøjer og klargøring før produktion.',
                        tags: ['produktion'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-19',
                        title: 'QFP-19 Godkendelse og levering',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-19%20Godkendelse%20og%20levering.aspx',
                        content: 'Kontrol, godkendelse og levering efter aftalte specifikationer.',
                        tags: ['levering', 'kontrol'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-21',
                        title: 'QFP-21 Fakturering og opfølgning',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-21%20Fakturering%20og%20opf%C3%B8lgning.aspx',
                        content: 'Fakturering, sagsafslutning og opfølgning efter levering.',
                        tags: ['fakturering'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-22',
                        title: 'QFP-22 Produktion',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-22%20Produktion.aspx',
                        content: 'Ordrestyring, produktionsforløb og overdragelse til levering.',
                        tags: ['produktion'],
                        updatedAt: new Date().toISOString()
                    }
                ]
            }
        ]
    };
}

function ensureQmsDatasetFile(fsRef) {
    const dataDir = path.dirname(QMS_DATASET_PATH);
    if (!fsRef.existsSync(dataDir)) {
        fsRef.mkdirSync(dataDir, { recursive: true });
    }
    if (!fsRef.existsSync(QMS_DATASET_PATH)) {
        const seed = defaultQmsDataset();
        fsRef.writeFileSync(QMS_DATASET_PATH, JSON.stringify(seed, null, 2), 'utf8');
    }
}

function readQmsDataset(fsRef) {
    ensureQmsDatasetFile(fsRef);
    const raw = fsRef.readFileSync(QMS_DATASET_PATH, 'utf8');
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.folders)) {
        throw new Error('QMS dataset format invalid');
    }
    return parsed;
}

function validateQmsDataset(payload) {
    if (!payload || typeof payload !== 'object') return 'Dataset mangler';
    if (!Array.isArray(payload.folders)) return 'folders skal være en liste';
    for (const folder of payload.folders) {
        if (!folder || typeof folder !== 'object') return 'Ugyldig mappe';
        if (!String(folder.id || '').trim()) return 'Mappe mangler id';
        if (!String(folder.name || '').trim()) return 'Mappe mangler navn';
        if (!Array.isArray(folder.documents)) return 'Mappe documents skal være en liste';
        for (const doc of folder.documents) {
            if (!doc || typeof doc !== 'object') return 'Ugyldigt dokument';
            if (!String(doc.id || '').trim()) return 'Dokument mangler id';
            if (!String(doc.title || '').trim()) return 'Dokument mangler titel';
            if (doc.tags && !Array.isArray(doc.tags)) return 'tags skal være en liste';
        }
    }
    return null;
}

function writeQmsDataset(fsRef, dataset) {
    ensureQmsDatasetFile(fsRef);
    const normalized = {
        version: Number(dataset.version || 1),
        updatedAt: new Date().toISOString(),
        folders: dataset.folders.map(folder => ({
            id: String(folder.id || '').trim(),
            name: String(folder.name || '').trim(),
            description: String(folder.description || '').trim(),
            documents: folder.documents.map(doc => ({
                id: String(doc.id || '').trim(),
                title: String(doc.title || '').trim(),
                url: String(doc.url || '').trim(),
                content: String(doc.content || '').trim(),
                tags: Array.isArray(doc.tags) ? doc.tags.map(t => String(t).trim()).filter(Boolean) : [],
                updatedAt: new Date().toISOString()
            }))
        }))
    };
    fsRef.writeFileSync(QMS_DATASET_PATH, JSON.stringify(normalized, null, 2), 'utf8');
    return normalized;
}

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

const omsaetningThresholdsService = require('../services/omsaetningThresholdsService');
const settingsService = require('../services/settingsService');
const getConnectionModule = require('../db');
const { createOmsaetningService } = require('../services/omsaetningService');
const { createOrdreindgangService } = require('../services/ordreindgangService');
const { createBomService } = require('../services/bomService');

function createApiRouter({
    getConnection,
    sql,
    fs,
    spawn,
    diskCache,
    logEvent,
    getOrComputeAftercalc,
    getOrComputeOrderMargin,
    getProductionSummary,
    AFTERCALC_CACHE_KEY_PREFIX,
    ORDER_MARGIN_CACHE_KEY_PREFIX,
    CACHE_TTL_ORDER_MARGIN_MS,
    CACHE_TTL_LASER_METRICS_MS,
    isHttpUrl,
    normalizeWindowsPath,
    isAbsoluteWindowsPath,
    isSupportedImagePath,
    buildImageItems,
    orderListCache,
    orderMarginCache,
    orderRefreshInFlight,
    orderRefreshStatus,
    orderMarginInFlight,
    afterCalcInFlight,
    warmupProgress,
    refreshOrderListCache,
    isOrderListCacheFresh,
    ORDER_LIST_DAYS_BACK,
    pkgVersion
}) {
    const router = express.Router();
    const legacyAftercalcPrefixes = ['aftercalc_v22_', 'aftercalc_v21_', 'aftercalc_v20_', 'aftercalc_v19_', 'aftercalc_v18_', 'aftercalc_v17_', 'aftercalc_'];
    const omsaetningService = createOmsaetningService({ getConnection, sql });
    const ordreindgangService = createOrdreindgangService({ getConnection, sql });
    const bomService = createBomService({ getConnection, sql, diskCache, logEvent });

    router.get('/aftercalc/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            logEvent('SEARCH: OrdNo=' + ordNo);
            const data = await getOrComputeAftercalc(ordNo, { priority: 'high' });
            if (!data || data.error) {
                return res.json(data);
            }

            if (!data.error) {
                logEvent('  -> Found: Revenue=' + data.summary.totalRevenue + ', Margin=' + data.summary.marginPercentage + '%');
            }
            res.json(data);
        } catch (err) {
            console.error('Errore API:', err);
            logEvent('ERROR: ' + err.message);
            res.status(500).json({ error: err.message });
        }
    });

    router.get('/salgordre-via', async (req, res) => {
        try {
            const requestedOrdNo = req.query.ordNo === undefined ? null : Number(req.query.ordNo);
            if (requestedOrdNo !== null && (!Number.isInteger(requestedOrdNo) || requestedOrdNo <= 0)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }

            const cacheKey = 'salgordre_via_v12';
            if (requestedOrdNo === null && req.query.force !== '1') {
                const cached = diskCache.get(cacheKey);
                if (cached) return res.json({ ...cached, cached: true });
            }
            if (requestedOrdNo !== null && req.query.force === '1') {
                diskCache.del(cacheKey);
            }

            const pool = await getConnection();
            const request = pool.request();
            request.timeout = 60000;
            request.input('requestedOrdNo', sql.Numeric, requestedOrdNo);
            const result = await request.query(`
                WITH OpenSalesOrders AS (
                    SELECT OrdNo, DelDt, CreUsr, CustNo
                    FROM Ord WITH(NOLOCK)
                                        WHERE OrdTp = 1
                                            AND TrTp = 1
                                            AND Gr12 <> 10
                                            AND (
                                                    OrdPrSt & 256 = 256
                                                    OR OrdPrSt = 0
                                                    OR OrdPrSt = 402653456
                                                    OR OrdPrSt = 134217728
                                                    OR OrdPrSt & 4194304 = 4194304
                                            )
                                                AND (@requestedOrdNo IS NULL OR OrdNo = @requestedOrdNo)
                ),
                ProductionOrders AS (
                    SELECT DISTINCT
                        P.OrdBasNo AS SalesOrderNo,
                        P.OrdNo
                    FROM Ord P WITH(NOLOCK)
                    INNER JOIN OpenSalesOrders S ON S.OrdNo = P.OrdBasNo
                    WHERE P.TrTp <> 6
                ),
                ResourceMinutes AS (
                    SELECT
                        ProductionOrders.SalesOrderNo,
                        L.OrdNo,
                        L.LnNo,
                        CASE
                            WHEN R.Nm LIKE '%laser%'
                             AND ISNULL(Nesting.TotalLaserLines, 0) > 0
                             AND Nesting.FinishedLaserLines = Nesting.TotalLaserLines
                            THEN 80
                            ELSE ISNULL(L.TransGr3, 0)
                        END AS ResourceStatus,
                        ISNULL(L.NoOrg, 0) AS OriginalMinutes,
                        COALESCE(NULLIF(ProdTr.FinishedMinutes, 0), ISNULL(L.NoFin, 0)) AS FinishedMinutes
                    FROM ProductionOrders
                    INNER JOIN OrdLn L WITH(NOLOCK) ON L.OrdNo = ProductionOrders.OrdNo
                    LEFT JOIN R7 R WITH(NOLOCK) ON R.RNo = L.R7
                    OUTER APPLY (
                        SELECT SUM(CAST(P.NoInvoAb AS decimal(18, 6))) AS FinishedMinutes
                        FROM ProdTr P WITH(NOLOCK)
                        WHERE P.OrdNo = L.OrdNo
                          AND P.OrdLnNo = L.LnNo
                    ) ProdTr
                    OUTER APPLY (
                        SELECT
                            COUNT(*) AS TotalLaserLines,
                            SUM(CASE WHEN ISNULL(N.NoFin, 0) > 0 THEN 1 ELSE 0 END) AS FinishedLaserLines
                        FROM OrdLn N WITH(NOLOCK)
                        WHERE N.TrInf2 = CONVERT(varchar(20), L.OrdNo)
                          AND N.TrTp = 7
                          AND N.ProdNo LIKE '%L%'
                    ) Nesting
                    WHERE L.ProdTp4 IN (1, 3)
                      AND L.R7 <> ''
                ),
                ActiveProduction AS (
                    SELECT
                        SalesOrderNo,
                        COUNT(DISTINCT CASE WHEN ResourceStatus < 80 THEN OrdNo END) AS OpenProductionOrders,
                        COUNT(*) AS TotalResources,
                        SUM(CASE WHEN ResourceStatus = 80 THEN 1 ELSE 0 END) AS CompletedResources,
                        SUM(CASE WHEN ResourceStatus < 80 THEN 1 ELSE 0 END) AS RemainingResources,
                        SUM(FinishedMinutes) AS CompletedResourceMinutes,
                        SUM(CASE WHEN ResourceStatus = 80 THEN FinishedMinutes ELSE OriginalMinutes END) AS EffectiveResourceMinutes
                    FROM ResourceMinutes
                    GROUP BY SalesOrderNo
                )
                SELECT
                    S.OrdNo,
                    S.DelDt AS DeliveryDate,
                    S.CreUsr AS SellerUsr,
                    C.Nm AS CustomerName,
                    Active.OpenProductionOrders,
                    Active.TotalResources,
                    Active.CompletedResources,
                    Active.RemainingResources,
                    Active.CompletedResourceMinutes,
                    Active.EffectiveResourceMinutes,
                    CAST(NULL AS datetime) AS PlannedDate,
                    CAST(NULL AS varchar(100)) AS ResourceName
                FROM OpenSalesOrders S
                LEFT JOIN Actor C WITH(NOLOCK) ON C.CustNo = S.CustNo
                LEFT JOIN ActiveProduction Active ON Active.SalesOrderNo = S.OrdNo
                ORDER BY
                    CASE WHEN S.DelDt > 19800101 THEN S.DelDt ELSE 99991231 END,
                    S.OrdNo
            `);
            const payload = { rows: result.recordset || [] };
            if (requestedOrdNo === null) diskCache.set(cacheKey, payload, 5 * 60 * 1000);
            res.json(payload);
        } catch (err) {
            logEvent('ERROR salgordre-via: ' + err.message);
            res.status(500).json({ error: err.message || 'SalgOrdre VIA fejl' });
        }
    });

    router.get('/ordreoversigt/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            if (!Number.isFinite(ordNo)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }

            const cacheKey = 'ordreoversigt_v5_' + ordNo;
            const cached = diskCache.get(cacheKey);
            if (cached) return res.json({ ...cached, cached: true });

            const pool = await getConnection();
            const [headerResult, linesResult, planningResult, customerNotesResult] = await Promise.all([
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                        SELECT
                            O.OrdNo,
                            O.CustNo,
                            O.Nm,
                            O.ReqNo,
                            O.LiaActNo,
                            O.SelBuy,
                            O.TrTp,
                            O.Rsp,
                            O.CreUsr AS SellerUsr,
                            O.Ad1 AS OrderAddress1,
                            O.Ad2 AS OrderAddress2,
                            O.PNo AS OrderPostCode,
                            O.PArea AS OrderCity,
                            O.DelMt,
                            O.DelDt,
                            O.DelNm,
                            O.DelAd1,
                            O.DelAd2,
                            O.DelPNo,
                            O.DelPArea,
                            O.DelTrm,
                            O.Inf2,
                            O.Gr4,
                            O.ShpActNo,
                            ISNULL((SELECT TOP (1) Txt FROM Txt WITH(NOLOCK) WHERE Lang = 45 AND TxtTp = 5 AND TxtNo = O.DelMt), '') AS DeliveryMode,
                            A.Nm AS CustomerName,
                            A.Ad1 AS CustomerAddress1,
                            A.Ad2 AS CustomerAddress2,
                            A.PNo AS CustomerPostCode,
                            A.PArea AS CustomerCity
                        FROM Ord O
                        LEFT JOIN Actor A ON A.CustNo = O.CustNo
                        WHERE O.OrdNo = @ordNo
                    `),
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                                                SELECT
                                                    Ord_1.OrdBasNo AS SalgsOrdre,
                                                    O.MainOrd AS HovOrd,
                                                    O.OrdNo AS ProdOrd,
                                                        L.OrdNo,
                                                        L.LnNo,
                                                        L.ProdNo,
                                                        L.Descr,
                                                    L.NoInvoAb + L.NoFin - L.NoInvo AS Ant,
                                                    L.NoInvoAb AS SaveAnt,
                                                        L.NoInvo,
                                                        L.NoOrg,
                                                        L.NoFin,
                                                    L.Un,
                                                    L.TrInf1,
                                                    L.TrInf3 AS Savelængde,
                                                    O.OrdBasNo AS OrdGrNo,
                                                    Ord_1.MainOrd AS main,
                                                        L.PurcNo,
                                                        L.ProdTp4,
                                                        L.TrTp,
                                                        L.TrInf2,
                                                        L.TrInf4,
                                                        L.WebPg,
                                                        L.PictFNm,
                                                        P.Inf2 AS DrawingNo,
                                                        P.Inf7 AS TegnNr,
                                                        P.Inf4 AS CustomerItemNo,
                                                        P.Gr6,
                                                        P.Gr5,
                                                        Material.ProdNo AS MaterialNo,
                                                        Material.Descr AS MaterialDescription,
                                                        FirstOperation.Descr AS FirstOperation,
                                                        FirstOperation.ProdNo AS FirstOperationNo
                                                FROM OrdLn L WITH(NOLOCK)
                                                INNER JOIN Ord O WITH(NOLOCK) ON O.OrdNo = L.OrdNo
                                                LEFT JOIN Ord Ord_1 WITH(NOLOCK) ON O.MainOrd = Ord_1.OrdNo
                                                LEFT JOIN Prod P WITH(NOLOCK) ON P.ProdNo = L.ProdNo
                                                OUTER APPLY (
                                                        SELECT TOP (1) M.ProdNo, M.Descr
                                                        FROM Struct S WITH(NOLOCK)
                                                        INNER JOIN Prod M WITH(NOLOCK) ON M.ProdNo = S.SubProd
                                                        WHERE S.ProdNo = L.ProdNo
                                                            AND S.SubProd LIKE '3%'
                                                        ORDER BY S.SubProd
                                                ) Material
                                                OUTER APPLY (
                                                        SELECT TOP (1) Child.Descr, Child.ProdNo
                                                        FROM OrdLn Child WITH(NOLOCK)
                                                        WHERE Child.OrdNo = L.PurcNo
                                                            AND Child.ProdTp4 IN (1, 3)
                                                        ORDER BY Child.LnNo
                                                ) FirstOperation
                                                WHERE L.OrdNo = @ordNo
                                                ORDER BY L.LnNo
                    `),
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                        SELECT
                            R7.MainR7 AS ResourceNo,
                            R7.Nm AS ResourceName,
                            dbo.EGD_Int2Date(F.Dt1) AS PlannedDate,
                            SUM(ABS(CONVERT(float, F.Val1))) AS PlannedHours
                        FROM FreeInf1 F WITH(NOLOCK)
                        INNER JOIN OrdLn L WITH(NOLOCK)
                            ON L.OrdNo = F.OrdNo AND L.LnNo = F.OrdLnNo
                        INNER JOIN Ord PO WITH(NOLOCK)
                            ON PO.OrdNo = F.OrdNo
                        INNER JOIN R7 WITH(NOLOCK)
                            ON R7.RNo = F.R7
                        WHERE F.FrInfTp = 2
                          AND F.Val1 < 0
                          AND L.ProdTp4 IN (1, 3)
                          AND (PO.OrdBasNo = @ordNo OR PO.OrdNo = @ordNo)
                        GROUP BY R7.MainR7, R7.Nm, F.Dt1
                        ORDER BY F.Dt1, R7.MainR7
                    `),
                pool.request()
                    .input('ordNo', sql.Numeric, ordNo)
                    .query(`
                        SELECT ActInf.LnNo, ActInf.Txt1
                        FROM Ord O WITH(NOLOCK)
                        INNER JOIN Actor A WITH(NOLOCK) ON A.CustNo = O.CustNo
                        INNER JOIN ActInf ActInf WITH(NOLOCK) ON ActInf.ActNo = A.ActNo
                        WHERE O.OrdNo = @ordNo
                          AND ActInf.InfTp = 10
                        ORDER BY ActInf.LnNo
                    `)
            ]);

            if (headerResult.recordset.length === 0) {
                return res.status(404).json({ error: 'Ordre ikke fundet' });
            }

            const lines = linesResult.recordset;
            const linkedOrders = [];
            const queuedOrders = lines
                .map(line => ({
                    ordNo: Number(line.PurcNo || 0),
                    parentOrderNo: ordNo
                }))
                .filter(order => Number.isFinite(order.ordNo) && order.ordNo > 0);
            const seenOrderNos = new Set();

            while (queuedOrders.length > 0 && linkedOrders.length < 250) {
                const nextOrder = queuedOrders.shift();
                const linkedOrderNo = nextOrder.ordNo;
                if (seenOrderNos.has(linkedOrderNo)) continue;
                seenOrderNos.add(linkedOrderNo);

                const [linkedHeaderResult, linkedLinesResult] = await Promise.all([
                    pool.request()
                        .input('linkedOrderNo', sql.Numeric, linkedOrderNo)
                        .query('SELECT OrdNo, TrTp, ArDt, OrdBasNo FROM Ord WHERE OrdNo = @linkedOrderNo'),
                    pool.request()
                        .input('linkedOrderNo', sql.Numeric, linkedOrderNo)
                        .query(`
                            SELECT
                                Ord_1.OrdBasNo AS SalgsOrdre,
                                O.MainOrd AS HovOrd,
                                O.OrdNo AS ProdOrd,
                                   L.OrdNo,
                                   L.LnNo,
                                   L.ProdNo,
                                   L.Descr,
                                   L.NoInvoAb + L.NoFin - L.NoInvo AS Ant,
                                   L.NoInvoAb AS SaveAnt,
                                   L.NoInvo,
                                   L.NoOrg,
                                   L.NoFin,
                                   L.Un,
                                   L.TrInf1,
                                   L.TrInf3 AS Savelængde,
                                   O.OrdBasNo AS OrdGrNo,
                                   Ord_1.MainOrd AS main,
                                   L.PurcNo,
                                   L.ProdTp4,
                                   L.TrTp,
                                   L.TrInf2,
                                   L.TrInf4,
                                   L.WebPg,
                                   L.PictFNm,
                                   P.Inf2 AS DrawingNo,
                                   P.Inf7 AS TegnNr,
                                   P.Inf4 AS CustomerItemNo,
                                   P.Gr6,
                                   P.Gr5,
                                   Material.ProdNo AS MaterialNo,
                                   Material.Descr AS MaterialDescription,
                                   FirstOperation.Descr AS FirstOperation,
                                   FirstOperation.ProdNo AS FirstOperationNo
                            FROM OrdLn L WITH(NOLOCK)
                            INNER JOIN Ord O WITH(NOLOCK) ON O.OrdNo = L.OrdNo
                            LEFT JOIN Ord Ord_1 WITH(NOLOCK) ON O.MainOrd = Ord_1.OrdNo
                            LEFT JOIN Prod P WITH(NOLOCK) ON P.ProdNo = L.ProdNo
                            OUTER APPLY (
                                SELECT TOP (1) M.ProdNo, M.Descr
                                FROM Struct S WITH(NOLOCK)
                                INNER JOIN Prod M WITH(NOLOCK) ON M.ProdNo = S.SubProd
                                WHERE S.ProdNo = L.ProdNo
                                  AND S.SubProd LIKE '3%'
                                ORDER BY S.SubProd
                            ) Material
                            OUTER APPLY (
                                SELECT TOP (1) Child.Descr, Child.ProdNo
                                FROM OrdLn Child WITH(NOLOCK)
                                WHERE Child.OrdNo = L.PurcNo
                                  AND Child.ProdTp4 IN (1, 3)
                                ORDER BY Child.LnNo
                            ) FirstOperation
                            WHERE L.OrdNo = @linkedOrderNo
                            ORDER BY L.LnNo
                        `)
                ]);
                const linkedHeader = linkedHeaderResult.recordset[0] || {};
                if (!linkedHeader.OrdNo) continue;

                const linkedOrder = {
                    ordNo: linkedOrderNo,
                    parentOrderNo: nextOrder.parentOrderNo,
                    type: Number(linkedHeader.TrTp || 0) === 6 ? 'purchase' : 'production',
                    ordBaseNo: Number(linkedHeader.OrdBasNo || 0) > 0 ? Number(linkedHeader.OrdBasNo) : null,
                    ulevDate: Number(linkedHeader.ArDt || 0) > 19800101 ? linkedHeader.ArDt : null,
                    lines: linkedLinesResult.recordset
                };
                linkedOrders.push(linkedOrder);

                for (const linkedLine of linkedOrder.lines) {
                    const childOrderNo = Number(linkedLine.PurcNo || 0);
                    if (Number.isFinite(childOrderNo) && childOrderNo > 0 && !seenOrderNos.has(childOrderNo)) {
                        queuedOrders.push({ ordNo: childOrderNo, parentOrderNo: linkedOrderNo });
                    }
                }
            }

            const productionOrderNos = linkedOrders
                .filter(order => order.type === 'production')
                .map(order => Number(order.ordNo || 0))
                .filter(value => Number.isFinite(value) && value > 0);
            const purchaseItems = linkedOrders
                .filter(order => order.type === 'production')
                .flatMap(order => (order.lines || [])
                    .filter(line => String(line.ProdTp4 || '') === '2' && !/L$/i.test(String(line.ProdNo || '').trim()))
                    .map(line => ({
                        ...line,
                        ProductionOrderNo: order.ordNo,
                        ParentOrderNo: order.parentOrderNo,
                        UlevDate: order.ulevDate
                    })));
            const purchaseProdNos = Array.from(new Set(
                purchaseItems.map(item => String(item.ProdNo || '').trim()).filter(Boolean)
            ));

            let laserNesting = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const laserResult = await request.query(`
                    SELECT DISTINCT
                        Ord.OrdNo AS ProdOrd,
                        OrdLn.ProdNo AS ProdNo,
                        Prod.Descr AS Descr,
                        Prod.Inf7 AS TegnNr,
                        OrdLn.TrInf3 AS SavLgd,
                        SUBSTRING(Struct.SubProd, 1, 13) AS MaterialNo,
                        SUBSTRING(Prod_1.Descr, 1, 50) AS MaterialDescription,
                        OrdLn.NoInvoAb AS Ant,
                        Prod_2.PictNo AS Pict,
                        OrdLn_1.TrInf1 AS OplNest,
                        Prod.Inf4 AS CustomerItemNo,
                        Struct.Descr AS StructDescr,
                        Struct_1.NoPerStr AS Deling,
                        ISNULL(ProdCat.Descr, '-') AS retn
                    FROM Ord Ord WITH(NOLOCK)
                    LEFT JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    LEFT JOIN Prod Prod WITH(NOLOCK) ON Prod.ProdNo = OrdLn.ProdNo
                    LEFT JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    LEFT JOIN Struct Struct WITH(NOLOCK) ON Struct.ProdNo = OrdLn.ProdNo
                    LEFT JOIN Prod Prod_1 WITH(NOLOCK) ON Prod_1.ProdNo = Struct.SubProd
                    LEFT JOIN Struct Struct_1 WITH(NOLOCK) ON Struct_1.SubProd = OrdLn.ProdNo
                    LEFT JOIN Prod Prod_2 WITH(NOLOCK) ON Struct_1.ProdNo = Prod_2.ProdNo
                    LEFT JOIN OrdLn OrdLn_1 WITH(NOLOCK) ON OrdLn_1.PurcNo = Ord_1.MainOrd
                    LEFT JOIN ProdCat ProdCat WITH(NOLOCK) ON Prod.PrCatNo = ProdCat.PrCatNo
                    WHERE OrdLn.TrTp = 5
                      AND OrdLn.ProdNo LIKE '%L%'
                      AND Ord_1.OrdBasNo = @ordNo
                      AND Struct.SubProd LIKE '3%'
                      AND OrdLn.NoInvoAb <> 0
                      AND Prod.Inf5 <> 'komb'
                    ORDER BY Ord.OrdNo, OrdLn.ProdNo
                `);
                laserNesting = laserResult.recordset || [];
            }

            let purchaseStockRows = [];
            if (purchaseProdNos.length > 0) {
                const request = pool.request();
                const productPlaceholders = purchaseProdNos.map((prodNo, index) => {
                    const parameter = 'purchaseProdNo' + index;
                    request.input(parameter, sql.VarChar, prodNo);
                    return '@' + parameter;
                }).join(', ');
                const stockResult = await request.query(`
                    SELECT
                        P.ProdNo,
                        P.Descr AS ProductDescription,
                        ISNULL(B.Bal, 0) AS StockBalance,
                        ISNULL(B.InProdO, 0) AS Reserved,
                        F.Sup AS SupplierNo,
                        A.Nm AS SupplierName
                    FROM Prod P WITH(NOLOCK)
                    LEFT JOIN StcBal B WITH(NOLOCK)
                        ON B.ProdNo = P.ProdNo
                       AND B.StcNo = 1
                    LEFT JOIN FreeInf1 F WITH(NOLOCK)
                        ON F.ProdNo = P.ProdNo
                       AND F.Sup <> 0
                    LEFT JOIN Actor A WITH(NOLOCK)
                        ON A.SupNo = F.Sup
                    WHERE CONVERT(varchar(100), P.ProdNo) IN (${productPlaceholders})
                    ORDER BY P.ProdNo, A.Nm
                `);
                purchaseStockRows = stockResult.recordset || [];
            }

            let plateInventory = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const plateResult = await request.query(`
                    SELECT DISTINCT
                        Prod_1.R3 AS Carrier,
                        Prod.ProdNo,
                        Prod.Descr,
                        ISNULL(StcBal.PoPhStB, 0) AS StockBalance,
                        Prod.NWgtU AS SheetWeight,
                        Prod.HgtU AS Thickness,
                        CASE WHEN ISNULL(Prod.NWgtU, 0) <> 0 THEN ISNULL(StcBal.PoPhStB, 0) / Prod.NWgtU ELSE 0 END AS SheetCount,
                        CASE WHEN ISNULL(Prod.NWgtU, 0) <> 0 THEN (ISNULL(StcBal.ShpRsv, 0) + ISNULL(StcBal.ShpRsvIn, 0)) / Prod.NWgtU ELSE 0 END AS Reserved,
                        ISNULL(StcBal.PhCstPr, 0) AS FifoPrice
                    FROM Ord Ord WITH(NOLOCK)
                    INNER JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    INNER JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    INNER JOIN Struct Struct WITH(NOLOCK) ON Struct.ProdNo = OrdLn.ProdNo
                    INNER JOIN Prod Prod_1 WITH(NOLOCK) ON Prod_1.ProdNo = Struct.SubProd
                    INNER JOIN Prod Prod WITH(NOLOCK) ON Prod_1.R3 = Prod.R3
                    INNER JOIN StcBal StcBal WITH(NOLOCK) ON StcBal.ProdNo = Prod.ProdNo
                    WHERE OrdLn.TrTp = 5
                      AND OrdLn.ProdNo LIKE '%L%'
                      AND Ord_1.OrdBasNo = @ordNo
                      AND Struct.SubProd LIKE '3%'
                      AND StcBal.StcNo = 1
                      AND Prod.ProdNo <> '301001'
                      AND Prod.ProdNo LIKE '3%'
                      AND OrdLn.NoInvoAb <> 0
                    ORDER BY Prod.ProdNo
                `);
                plateInventory = plateResult.recordset || [];
            }

            let lDescriptionRows = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const result = await request.query(`
                    SELECT
                        Ord.OrdNo AS ProdOrd,
                        OrdLn.ProdNo AS ProdNr,
                        Prod.Descr AS Beskrivelse,
                        ProdDesc.Descr AS EkstraInf,
                        ProdDesc.LnNo
                    FROM Ord Ord WITH(NOLOCK)
                    INNER JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    INNER JOIN Prod Prod WITH(NOLOCK) ON Prod.ProdNo = OrdLn.ProdNo
                    INNER JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    INNER JOIN ProdDesc ProdDesc WITH(NOLOCK) ON ProdDesc.ProdNo = Prod.ProdNo
                    WHERE OrdLn.TrTp = 5
                      AND OrdLn.ProdNo LIKE '%L%'
                      AND Ord_1.OrdBasNo = @ordNo
                      AND OrdLn.NoInvoAb <> 0
                    ORDER BY Ord.OrdNo, ProdDesc.LnNo
                `);
                lDescriptionRows = result.recordset || [];
            }

            let customerLaserRows = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const result = await request.query(`
                    SELECT
                        Ord.OrdNo AS ProdOrd,
                        OrdLn.ProdNo AS ProdNr,
                        Prod.Descr AS Beskrivelse,
                        Prod.Inf7 AS TegnNr,
                        OrdLn.TrInf3 AS SavLgd,
                        SUBSTRING(Struct.SubProd, 1, 13) AS Raavarenr,
                        SUBSTRING(Prod_1.Descr, 1, 50) AS Raavarebetegn,
                        OrdLn.NoInvoAb AS Ant,
                        Prod_2.PictNo AS Pict,
                        OrdLn_1.TrInf1 AS OplNest,
                        Prod.Inf4 AS KVarenr,
                        Struct.Descr AS StructDescr,
                        Struct_1.NoPerStr AS Deling
                    FROM Ord Ord WITH(NOLOCK)
                    INNER JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    INNER JOIN Prod Prod WITH(NOLOCK) ON Prod.ProdNo = OrdLn.ProdNo
                    INNER JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    INNER JOIN Struct Struct WITH(NOLOCK) ON Struct.ProdNo = OrdLn.ProdNo
                    INNER JOIN Prod Prod_1 WITH(NOLOCK) ON Prod_1.ProdNo = Struct.SubProd
                    INNER JOIN Struct Struct_1 WITH(NOLOCK) ON Struct_1.SubProd = OrdLn.ProdNo
                    INNER JOIN Prod Prod_2 WITH(NOLOCK) ON Struct_1.ProdNo = Prod_2.ProdNo
                    INNER JOIN OrdLn OrdLn_1 WITH(NOLOCK) ON OrdLn_1.PurcNo = Ord_1.MainOrd
                    WHERE OrdLn.TrTp = 5
                      AND OrdLn.ProdNo LIKE '%L%'
                      AND Ord_1.OrdBasNo = @ordNo
                      AND Struct.SubProd LIKE '3%'
                      AND OrdLn.NoInvoAb <> 0
                    ORDER BY Ord.OrdNo, OrdLn.ProdNo
                `);
                customerLaserRows = result.recordset || [];
            }

            let laserInfoFromRouteRows = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const result = await request.query(`
                    SELECT DISTINCT
                        Ord.OrdNo AS ProdOrd,
                        OrdLn.ProdNo AS ProdNr,
                        Struct_2.SubProd,
                        Prod.Inf2,
                        R7.Gr3
                    FROM BgtLn BgtLn WITH(NOLOCK)
                    INNER JOIN Ord Ord WITH(NOLOCK) ON Ord.OrdNo = BgtLn.OrdNo
                    INNER JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    INNER JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    INNER JOIN Struct Struct WITH(NOLOCK) ON OrdLn.ProdNo = Struct.SubProd
                    INNER JOIN Struct Struct_1 WITH(NOLOCK) ON Struct_1.ProdNo = Struct.ProdNo
                    INNER JOIN Struct Struct_2 WITH(NOLOCK) ON Struct_1.SubProd = Struct_2.ProdNo
                    INNER JOIN Prod Prod WITH(NOLOCK) ON Prod.ProdNo = Struct_2.SubProd
                    INNER JOIN R7 R7 WITH(NOLOCK) ON R7.RNo = BgtLn.R7
                    WHERE OrdLn.TrTp = 5
                      AND OrdLn.ProdNo LIKE '%L%'
                      AND Ord_1.OrdBasNo = @ordNo
                      AND OrdLn.NoInvoAb <> 0
                      AND Struct_1.SubProd LIKE 'V%'
                      AND Prod.Inf2 <> ''
                    ORDER BY Ord.OrdNo, OrdLn.ProdNo
                `);
                laserInfoFromRouteRows = result.recordset || [];
            }

            let excelVareLinierRows = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const result = await request.query(`
                    SELECT
                        Ord_1.OrdBasNo AS SalgsOrdre,
                        Ord.MainOrd AS HovOrd,
                        Ord.OrdNo AS ProdOrd,
                        OrdLn.NoInvoAb + OrdLn.NoFin - OrdLn.NoInvo AS Ant,
                        OrdLn.ProdNo AS ProdNr,
                        Ord.OrdBasNo AS OrdGrNo,
                        OrdLn.Descr,
                        Ord_1.MainOrd AS main,
                        Prod.Inf7 AS TegnNr,
                        OrdLn.Un,
                        OrdLn.TrInf1
                    FROM Ord Ord WITH(NOLOCK)
                    INNER JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    INNER JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    INNER JOIN Prod Prod WITH(NOLOCK) ON Prod.ProdNo = OrdLn.ProdNo
                    WHERE OrdLn.TrTp = 7
                      AND OrdLn.ProdNo NOT LIKE 'V%'
                      AND Ord_1.OrdBasNo = @ordNo
                      AND OrdLn.NoInvoAb + OrdLn.NoFin - OrdLn.NoInvo <> 0
                    ORDER BY Ord.OrdNo, OrdLn.LnNo
                `);
                excelVareLinierRows = result.recordset || [];
            }

            let excelSaveListeRows = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const result = await request.query(`
                    SELECT
                        Ord_1.OrdBasNo AS SalgsOrdre,
                        Ord.MainOrd AS HovOrd,
                        Ord.OrdNo AS ProdOrd,
                        OrdLn.NoInvoAb AS Ant,
                        OrdLn.ProdNo AS ProdNr,
                        Ord.OrdBasNo AS OrdGrNo,
                        OrdLn.Descr,
                        Ord_1.MainOrd AS main,
                        Prod.Inf7 AS TegnNr,
                        OrdLn.TrInf3 AS Savelaengde,
                        Ord.Gr5
                    FROM Ord Ord WITH(NOLOCK)
                    INNER JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    INNER JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    INNER JOIN Prod Prod WITH(NOLOCK) ON Prod.ProdNo = OrdLn.ProdNo
                    WHERE OrdLn.TrTp = 7
                      AND OrdLn.ProdNo NOT LIKE 'V%'
                      AND Ord_1.OrdBasNo = @ordNo
                    ORDER BY Ord.OrdNo, OrdLn.LnNo
                `);
                excelSaveListeRows = result.recordset || [];
            }

            let excelIndkLstRows = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const result = await request.query(`
                    SELECT DISTINCT
                        OrdLn.ProdNo AS ProdNr,
                        OrdLn.Descr AS Beskrivelse,
                        Txt.Txt AS Enh,
                        Prod.Gr6,
                        StcBal.PoPhStB AS Beholdning,
                        StcBal.InProdO AS Reserveret,
                        Prod.Gr5
                    FROM Ord Ord WITH(NOLOCK)
                    INNER JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    INNER JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    INNER JOIN Prod Prod WITH(NOLOCK) ON Prod.ProdNo = OrdLn.ProdNo
                    INNER JOIN StcBal StcBal WITH(NOLOCK) ON StcBal.ProdNo = OrdLn.ProdNo
                    INNER JOIN Txt Txt WITH(NOLOCK) ON Prod.StSaleUn = Txt.TxtNo
                    WHERE OrdLn.TrTp = 5
                      AND Ord_1.OrdBasNo = @ordNo
                      AND Prod.Gr5 IN (2, 3, 11)
                      AND OrdLn.ProdNo NOT LIKE '%L%'
                      AND Txt.TxtTp = 16
                      AND Txt.Lang = 45
                      AND Prod.Gr6 <> 1

                    UNION

                    SELECT
                        OrdLn.ProdNo,
                        OrdLn.Descr,
                        Txt.Txt,
                        Prod.Gr6,
                        StcBal.PoPhStB,
                        StcBal.InProdO,
                        Prod.Gr5
                    FROM OrdLn OrdLn WITH(NOLOCK)
                    INNER JOIN Prod Prod WITH(NOLOCK) ON Prod.ProdNo = OrdLn.ProdNo
                    INNER JOIN StcBal StcBal WITH(NOLOCK) ON StcBal.ProdNo = Prod.ProdNo
                    INNER JOIN Txt Txt WITH(NOLOCK) ON Prod.StSaleUn = Txt.TxtNo
                    WHERE Txt.TxtTp = 16
                      AND Txt.Lang = 45
                      AND OrdLn.OrdNo = @ordNo
                      AND Prod.Gr5 = 3
                      AND Prod.Gr6 <> 1
                `);
                excelIndkLstRows = result.recordset || [];
            }

            let routeRows = [];
            {
                const request = pool.request();
                request.input('ordNo', sql.Numeric, ordNo);
                const result = await request.query(`
                    SELECT
                        Ord.OrdNo AS OrdNr,
                        OrdLn.ProdNo AS ProdNr,
                        OrdLn.NoInvoAb AS Ant,
                        Ord.MainOrd AS HovedOrdre,
                        Ord.OrdBasNo AS OrdGrNo,
                        OrdLn.Descr,
                        SUBSTRING(R7.Nm, 1, 3) AS ResourceShortName,
                        Struct_1.TrInf4,
                        Struct_1.Descr AS OperationDescr,
                        OrdLn.CfDelDt,
                        R7.RNo,
                        R7.Gr11,
                        Struct_1.TrInf3 AS PrgNr,
                        R7.Gr3
                    FROM Ord Ord WITH(NOLOCK)
                    INNER JOIN Ord Ord_1 WITH(NOLOCK) ON Ord.MainOrd = Ord_1.OrdNo
                    INNER JOIN OrdLn OrdLn WITH(NOLOCK) ON Ord.OrdNo = OrdLn.OrdNo
                    INNER JOIN Struct Struct WITH(NOLOCK) ON OrdLn.ProdNo = Struct.ProdNo
                    INNER JOIN Struct Struct_1 WITH(NOLOCK) ON Struct_1.ProdNo = Struct.SubProd
                    INNER JOIN R7 R7 WITH(NOLOCK) ON Struct_1.R7 = R7.RNo
                    WHERE OrdLn.TrTp = 7
                      AND OrdLn.ProdNo NOT LIKE 'V%'
                      AND Ord_1.OrdBasNo = @ordNo
                      AND Struct_1.ProdTp4 IN (1, 7)
                    ORDER BY OrdLn.ProdNo, Struct_1.TrInf4
                `);
                routeRows = result.recordset || [];
            }

            // Print flags: rowcounts from filtered result sets (Excel AB6, AB8, AB10)
            // saveListCount = SaveListe rows matching the Save*/Pladesaks* autofilter (Knap1_Klik)
            const saveListCount = excelSaveListeRows
                .filter(line => /^(Save|Pladesaks)/i.test(String(line.Descr || '').trim()))
                .length;
            
            // laserListCount = number of laser nesting lines from production orders
            const laserListCount = laserNesting.length;
            
            // purchaseListCount = rows in the real IndkLst query (Excel AB12)
            const purchaseListCount = excelIndkLstRows.length;
            
            // lDescriptionCount = rows in the real L-beskriv query (Excel AB10)
            const lDescriptionCount = lDescriptionRows.length;

            const deliveryDate = Number(headerResult.recordset[0].DelDt || 0) > 19800101
                ? headerResult.recordset[0].DelDt
                : null;
            const subDeliveryOrders = linkedOrders
                .filter(order => order.type === 'production')
                .map(order => order.ordNo);
            const warnings = [];
            if (subDeliveryOrders.length > 0) {
                warnings.push('U-lev ordrer: ' + subDeliveryOrders.join(', '));
            }
            if (lines.some(line => Number(line.NoOrg || 0) > 0 && Number(line.NoFin || 0) === 0)) {
                warnings.push('Ordren indeholder linjer, der endnu ikke er færdigmeldt.');
            }

            const payload = {
                order: {
                    ...headerResult.recordset[0],
                    DeliveryDate: deliveryDate
                },
                lines,
                linkedOrders,
                laserNesting,
                purchaseItems,
                purchaseStockRows,
                plateInventory,
                lDescriptionRows,
                customerLaserRows,
                laserInfoFromRouteRows,
                routeRows,
                excelVareLinierRows,
                excelSaveListeRows,
                excelIndkLstRows,
                customerNotes: customerNotesResult.recordset || [],
                printFlags: {
                    saveList: saveListCount,
                    purchaseList: purchaseListCount,
                    laserList: laserListCount,
                    lDescription: lDescriptionCount
                },
                planning: planningResult.recordset,
                warnings,
                cached: false
            };
            diskCache.set(cacheKey, payload, 5 * 60 * 1000);
            res.json(payload);
        } catch (err) {
            logEvent('ERROR ordreoversigt: ' + err.message);
            res.status(500).json({ error: err.message || 'Ordreoversigt fejl' });
        }
    });

    router.get('/order-margin/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            if (Number.isNaN(ordNo)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }
            const cacheKey = ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo;
            const cached = diskCache.get(cacheKey);
            if (cached) return res.json({ ...cached, cached: true });

            const marginInfo = await getOrComputeOrderMargin(ordNo);
            const result = {
                ordNo: marginInfo.ordNo,
                totalRevenue: marginInfo.totalRevenue,
                totalCost: marginInfo.totalCost,
                hasInvoiceWarning: Boolean(marginInfo.hasInvoiceWarning),
                cached: true
            };
            diskCache.set(cacheKey, result, CACHE_TTL_ORDER_MARGIN_MS);
            return res.json(result);
        } catch (err) {
            logEvent('ERROR order-margin: ' + err.message);
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/production-summary/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            if (Number.isNaN(ordNo)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }

            const orderGr4 = Number(req.query.gr4 || 0);
            const result = await getProductionSummary(ordNo, new Set(), { orderGr4 });
            return res.json(result);
        } catch (err) {
            console.error('Errore production-summary:', err);
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/nesting-detail/:ordno/:prodno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            const prodNo = req.params.prodno;
            if (Number.isNaN(ordNo) || !prodNo) {
                return res.status(400).json({ error: 'Ugyldige parametre' });
            }
            const pool = await getConnection();
            const result = await pool.request()
                .input('ordNo', sql.Numeric, ordNo)
                .input('prodNo', sql.VarChar, prodNo)
                .query(`
                    SELECT OrdNo, TrInf4, CstPr, NoFin, Descr
                    FROM OrdLn
                    WHERE TrInf2 = CAST(@ordNo AS VARCHAR(20))
                      AND ProdNo = @prodNo
                    ORDER BY OrdNo, LnNo
                `);
            res.json(result.recordset);
        } catch (err) {
            res.status(500).json({ error: err.message });
        }
    });

    router.get('/image-file', async (req, res) => {
        try {
            const rawPath = String(req.query.path || '').trim();
            if (!rawPath) {
                return res.status(400).json({ error: 'Billedsti mangler' });
            }

            if (isHttpUrl(rawPath)) {
                return res.redirect(rawPath);
            }

            const normalizedPath = normalizeWindowsPath(rawPath);
            if (!isAbsoluteWindowsPath(normalizedPath)) {
                return res.status(400).json({ error: 'Kun absolutte billedstier er tilladt' });
            }

            if (!isSupportedImagePath(normalizedPath)) {
                return res.status(400).json({ error: 'Filtypen understoettes ikke som billede' });
            }

            if (!fs.existsSync(normalizedPath)) {
                return res.status(404).json({ error: 'Billedfilen blev ikke fundet' });
            }

            return res.sendFile(normalizedPath, err => {
                if (err && !res.headersSent) {
                    res.status(err.statusCode || 500).json({ error: err.message });
                }
            });
        } catch (err) {
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/laser-route-metrics', async (req, res) => {
        try {
            const ordine = String(req.query.ordine || '').trim();
            const route = String(req.query.route || '').trim();
            const prodNoFilter = String(req.query.prodNo || '').trim();
            const normalizedProdNoFilter = prodNoFilter.toUpperCase();
            const showAllRoutes = req.query.showAllRoutes === '1';
            const orderGr4 = Number(req.query.gr4 || 0);
            const useSpecialLaserCost = orderGr4 === 3;

            if (!ordine) {
                return res.status(400).json({ error: 'Ugyldige parametre: ordine er paakraevet' });
            }

            const laserCacheKey = 'laser_v4_' + ordine + '_' + (route || 'all') + '_' + (prodNoFilter || 'all') + '_' + (showAllRoutes ? '1' : '0') + '_gr4_' + (useSpecialLaserCost ? '3' : '0');
            const cachedLaser = diskCache.get(laserCacheKey);
            if (cachedLaser) return res.json(cachedLaser);

            const pool = await getConnection();

            const candidateResult = await pool.request()
                .input('ordine', sql.VarChar, ordine)
                .query(`
                    SELECT OrdNo, TrInf4, ProdNo, NoFin
                    FROM OrdLn
                    WHERE TrInf2 = @ordine
                      AND TrTp = 7
                `);

            const candidates = candidateResult.recordset || [];
            const normalizedRoute = route ? route.toUpperCase() : '';
            const routeMatches = candidate => String(candidate.TrInf4 || '').trim().toUpperCase() === normalizedRoute;

            const withProd = prodNoFilter
                ? candidates.filter(c => String(c.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                : candidates;
            const withRoute = normalizedRoute
                ? withProd.filter(routeMatches)
                : withProd;
            const routeOnly = normalizedRoute
                ? candidates.filter(routeMatches)
                : [];

            const selectedCandidates = showAllRoutes
                ? (withProd.length > 0 ? withProd : (routeOnly.length > 0 ? routeOnly : candidates))
                : (withRoute.length > 0
                    ? [withRoute[0]]
                    : (withProd.length > 0
                        ? [withProd[0]]
                        : (routeOnly.length > 0 ? [routeOnly[0]] : (candidates[0] ? [candidates[0]] : []))));

            if (selectedCandidates.length === 0) {
                return res.json({
                    ordine,
                    route: route || null,
                    nestingOrdNo: null,
                    nestingOrdNos: [],
                    prodNo: prodNoFilter || null,
                    summary: {
                        KgConsumati: null,
                        CostoLastre: null,
                        KgFiniti: null,
                        SfridoKg: null,
                        SfridoPct: null
                    },
                    products: []
                });
            }

            const nestingOrdNos = Array.from(new Set(
                selectedCandidates
                    .map(candidate => String(candidate.OrdNo || '').trim())
                    .filter(Boolean)
            ));
            const effectiveRoute = showAllRoutes ? '' : String(selectedCandidates[0].TrInf4 || '').trim();

            // Mappa OrdNo_TrInf4_ProdNo → NoFin dalla produzione: contiene il vero Færdigmeldt per rotta.
            // I nesting order rows hanno spesso NoFin=totale (es. 40 per tutte le rotte),
            // mentre questi record (TrInf2=produzione) hanno il valore corretto per singola rotta.
            const candidateNoFinMap = new Map();
            for (const c of candidates) {
                const k = String(c.OrdNo || '').trim() + '_' + String(c.TrInf4 || '').trim() + '_' + String(c.ProdNo || '').trim().toUpperCase();
                if (!candidateNoFinMap.has(k)) candidateNoFinMap.set(k, Number(c.NoFin || 0));
            }

            if (nestingOrdNos.length === 0 || (!showAllRoutes && !effectiveRoute)) {
                return res.json({
                    ordine,
                    route: effectiveRoute || null,
                    nestingOrdNo: nestingOrdNos[0] || null,
                    nestingOrdNos,
                    prodNo: prodNoFilter || null,
                    summary: {
                        KgConsumati: null,
                        CostoLastre: null,
                        KgFiniti: null,
                        SfridoKg: null,
                        SfridoPct: null
                    },
                    products: []
                });
            }

            const nestingRowsRequest = pool.request();
            nestingOrdNos.forEach((ordValue, index) => {
                nestingRowsRequest.input(`nestingOrdNo${index}`, sql.VarChar, ordValue);
            });
            const nestingPlaceholders = nestingOrdNos.map((_, index) => `@nestingOrdNo${index}`).join(', ');
            const result = await nestingRowsRequest.query(`
                    SELECT LnNo, OrdNo, TrInf2, TrInf4, ProdNo, TrTp, NoFin, Free3, IncCst, CstPr, WebPg, PictFNm
                    FROM OrdLn
                    WHERE OrdNo IN (${nestingPlaceholders})
                      AND TrTp IN (5, 7)
                `);

            const rows = result.recordset || [];
            const toNumber = (v) => {
                if (v === null || v === undefined || v === '') return 0;
                if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
                const raw = String(v).trim();
                if (!raw) return 0;
                let normalized = raw;
                if (normalized.includes(',') && normalized.includes('.')) {
                    normalized = normalized.replace(/\./g, '').replace(',', '.');
                } else if (normalized.includes(',')) {
                    normalized = normalized.replace(',', '.');
                }
                const parsed = Number(normalized);
                return Number.isFinite(parsed) ? parsed : 0;
            };
            const normalizeExpectedWeight = (v) => {
                const parsed = toNumber(v);
                return Math.abs(parsed) >= 1000 ? (parsed / 1000) : parsed;
            };
            const round = (v) => Number.isFinite(v) ? parseFloat(Number(v).toFixed(6)) : null;

            const inScopeRows = showAllRoutes
                ? rows
                : rows.filter(r => String(r.TrInf4 || '').trim() === effectiveRoute);

            const sheetRows = inScopeRows.filter(r => Number(r.TrTp) === 5);
            const finishedRows = inScopeRows.filter(r => Number(r.TrTp) === 7);

            const kgConsumati = sheetRows.reduce((s, r) => s + toNumber(r.NoFin), 0);
            const costoLastre = sheetRows.reduce((s, r) => s + toNumber(r.IncCst), 0);

            const filteredFinishedRows = (prodNoFilter
                ? finishedRows.filter(r => String(r.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                : finishedRows)
                .sort((a, b) => toNumber(a.LnNo) - toNumber(b.LnNo));

            const filteredNestingRows = (prodNoFilter
                ? finishedRows.filter(r => String(r.ProdNo || '').trim().toUpperCase() === normalizedProdNoFilter)
                : finishedRows)
                .sort((a, b) => toNumber(a.LnNo) - toNumber(b.LnNo));

            const structMap = new Map();
            try {
                const uniqueProdNos = Array.from(new Set(
                    filteredNestingRows.map(r => String(r.ProdNo || '').trim().toUpperCase())
                ));
                if (uniqueProdNos.length > 0) {
                    const placeholders = uniqueProdNos.map((_, i) => `@p${i}`).join(', ');
                    const request = pool.request();
                    uniqueProdNos.forEach((prodNo, i) => {
                        request.input(`p${i}`, sql.VarChar, prodNo);
                    });
                    const structResult = await request.query(`
                        SELECT ProdNo, NoPerStr
                        FROM Struct
                        WHERE ProdNo IN (${placeholders})
                          AND SubProd LIKE '3%'
                    `);
                    const structRows = structResult.recordset || [];
                    for (const sr of structRows) {
                        const prodKey = String(sr.ProdNo || '').trim().toUpperCase();
                        const noPerStr = toNumber(sr.NoPerStr);
                        if (!structMap.has(prodKey)) {
                            structMap.set(prodKey, noPerStr);
                            if (noPerStr > 0) {
                                logEvent(`DEBUG Struct: ProdNo=${prodKey}, NoPerStr=${noPerStr}`);
                            }
                        }
                    }
                }
            } catch (err) {
                logEvent(`Errore lettura Struct: ${err.message}`);
            }

            const getExpectedUnitWeight = (row) => {
                const free3Weight = normalizeExpectedWeight(row.Free3);
                if (free3Weight > 0) return free3Weight;
                const prodKey = String(row.ProdNo || '').trim().toUpperCase();
                return normalizeExpectedWeight(structMap.get(prodKey));
            };

            const kgFiniti = finishedRows.reduce((s, r) => s + (getExpectedUnitWeight(r) * toNumber(r.NoFin)), 0);
            const sfridoKg = kgConsumati - kgFiniti;
            const sfridoPct = kgConsumati > 0 ? (sfridoKg / kgConsumati) : null;

            const routeStats = new Map();
            const routeStatsKey = (row) => String(row.OrdNo || '').trim() + '|' + String(row.TrInf4 || '').trim();
            for (const row of inScopeRows) {
                const statsKey = routeStatsKey(row);
                if (!routeStats.has(statsKey)) {
                    routeStats.set(statsKey, {
                        kgConsumati: 0,
                        costoLastre: 0,
                        kgFiniti: 0,
                        cstPrSum: 0,
                        cstPrCount: 0
                    });
                }

                const stats = routeStats.get(statsKey);
                if (Number(row.TrTp) === 5) {
                    stats.kgConsumati += toNumber(row.NoFin);
                    stats.costoLastre += toNumber(row.IncCst);
                    const rowCstPr = toNumber(row.CstPr);
                    if (rowCstPr > 0) {
                        stats.cstPrSum += rowCstPr;
                        stats.cstPrCount += 1;
                    }
                } else if (Number(row.TrTp) === 7) {
                    stats.kgFiniti += getExpectedUnitWeight(row) * toNumber(row.NoFin);
                }
            }

            const products = filteredNestingRows.map(r => {
                const routeKey = String(r.TrInf4 || '').trim();
                const refFinished = Number(r.TrTp) === 7
                    ? r
                    : filteredFinishedRows.find(fr => String(fr.TrInf4 || '').trim() === routeKey && String(fr.OrdNo || '').trim() === String(r.OrdNo || '').trim());

                const prodKey = String(r.ProdNo || '').trim().toUpperCase();
                const candidateLookupKey = String(r.OrdNo || '').trim() + '_' + routeKey + '_' + prodKey;
                const candidateNoFin = candidateNoFinMap.has(candidateLookupKey) ? candidateNoFinMap.get(candidateLookupKey) : null;
                const rowNoFin = refFinished ? toNumber(refFinished.NoFin) : null;
                // Multiordre can have multiple finished rows with same ProdNo/Route but different quantities.
                // Keep row-level NoFin when present; only fallback to candidate map if row quantity is missing.
                const qtaPezzi = rowNoFin !== null && rowNoFin > 0
                    ? rowNoFin
                    : (candidateNoFin !== null && candidateNoFin > 0 ? candidateNoFin : null);
                const structNoPerStr = structMap.get(prodKey) || null;
                const oldExpectedUnitWeight = refFinished ? getExpectedUnitWeight(refFinished) : null;
                const expectedUnitWeight = (structNoPerStr !== null && structNoPerStr > 0)
                    ? structNoPerStr
                    : oldExpectedUnitWeight;
                const kgProdotto = (qtaPezzi !== null && expectedUnitWeight !== null)
                    ? (expectedUnitWeight * qtaPezzi)
                    : null;
                const stats = routeStats.get(String(r.OrdNo || '').trim() + '|' + routeKey) || { kgConsumati: 0, costoLastre: 0, kgFiniti: 0, cstPrSum: 0, cstPrCount: 0 };
                const nWgtUMedio = (qtaPezzi !== null && qtaPezzi > 0 && kgProdotto !== null) ? (kgProdotto / qtaPezzi) : null;
                const oldKgProdotto = (qtaPezzi !== null && oldExpectedUnitWeight !== null)
                    ? (oldExpectedUnitWeight * qtaPezzi)
                    : null;
                const oldNWgtUMedio = (qtaPezzi !== null && qtaPezzi > 0 && oldKgProdotto !== null) ? (oldKgProdotto / qtaPezzi) : null;
                const kgUtilizzatiEffettivi = (oldKgProdotto !== null && stats.kgFiniti > 0)
                    ? ((oldKgProdotto / stats.kgFiniti) * stats.kgConsumati)
                    : null;
                const kgPerPezzoEffettivo = (kgUtilizzatiEffettivi !== null && qtaPezzi !== null && qtaPezzi > 0)
                    ? (kgUtilizzatiEffettivi / qtaPezzi)
                    : null;
                const avgSheetCstPr = stats.cstPrCount > 0 ? (stats.cstPrSum / stats.cstPrCount) : null;
                const quotaCosto = useSpecialLaserCost
                    ? ((kgUtilizzatiEffettivi !== null && avgSheetCstPr !== null) ? (kgUtilizzatiEffettivi * avgSheetCstPr) : null)
                    : ((oldKgProdotto !== null && stats.kgFiniti > 0)
                        ? ((oldKgProdotto / stats.kgFiniti) * stats.costoLastre)
                        : null);
                const costoPerPezzo = (quotaCosto !== null && qtaPezzi !== null && qtaPezzi > 0) ? (quotaCosto / qtaPezzi) : null;
                const euroPerKgFinito = (costoPerPezzo !== null && nWgtUMedio !== null && nWgtUMedio > 0)
                    ? (costoPerPezzo / nWgtUMedio)
                    : null;
                const imageRow = refFinished || r;
                const imageItems = buildImageItems(imageRow ? imageRow.WebPg : null, imageRow ? imageRow.PictFNm : null);

                return {
                    LnNo: toNumber(r.LnNo),
                    NestingOrdNo: toNumber(r.OrdNo),
                    ProdNo: String(r.ProdNo || '').trim(),
                    Route: routeKey,
                    TrTp: toNumber(r.TrTp),
                    QtaPezzi: qtaPezzi === null ? null : round(qtaPezzi),
                    KgProdotto: kgProdotto === null ? null : round(kgProdotto),
                    OldNWgtU_medio: oldNWgtUMedio === null ? null : round(oldNWgtUMedio),
                    NWgtU_medio: nWgtUMedio === null ? null : round(nWgtUMedio),
                    KgUtilizzatiEffettivi: kgUtilizzatiEffettivi === null ? null : round(kgUtilizzatiEffettivi),
                    KgPerPezzoEffettivo: kgPerPezzoEffettivo === null ? null : round(kgPerPezzoEffettivo),
                    QuotaCosto: quotaCosto === null ? null : round(quotaCosto),
                    CostoPerPezzo: costoPerPezzo === null ? null : round(costoPerPezzo),
                    EuroPerKgFinito: euroPerKgFinito === null ? null : round(euroPerKgFinito),
                    ImageItems: imageItems,
                    WebPg: imageRow ? String(imageRow.WebPg || '').trim() : '',
                    PictFNm: imageRow ? String(imageRow.PictFNm || '').trim() : ''
                };
            });

            const _debugLookups = filteredNestingRows.map(r => {
                const rk = String(r.TrInf4 || '').trim();
                const pk = String(r.ProdNo || '').trim().toUpperCase();
                const lk = String(r.OrdNo || '').trim() + '_' + rk + '_' + pk;
                return { OrdNo: r.OrdNo, TrInf4: r.TrInf4, ProdNo: r.ProdNo, TrTp: r.TrTp, NoFin_row: r.NoFin, lookupKey: lk, found: candidateNoFinMap.has(lk), candidateNoFin: candidateNoFinMap.get(lk) };
            });
            logEvent('LASER_DEBUG ordine=' + ordine + ' candidates=' + JSON.stringify(candidates.map(c => ({ OrdNo: c.OrdNo, TrInf4: c.TrInf4, ProdNo: c.ProdNo, NoFin: c.NoFin }))));
            logEvent('LASER_DEBUG mapEntries=' + JSON.stringify(Array.from(candidateNoFinMap.entries()).map(([k, v]) => ({ key: k, noFin: v }))));
            logEvent('LASER_DEBUG lookups=' + JSON.stringify(_debugLookups));
            const laserResult = {
                ordine,
                route: showAllRoutes ? null : effectiveRoute,
                nestingOrdNo: showAllRoutes
                    ? (nestingOrdNos.length === 1 ? nestingOrdNos[0] : null)
                    : (nestingOrdNos[0] || null),
                nestingOrdNos,
                prodNo: prodNoFilter || null,
                showAllRoutes,
                summary: {
                    KgConsumati: round(kgConsumati),
                    CostoLastre: round(costoLastre),
                    KgFiniti: round(kgFiniti),
                    SfridoKg: round(sfridoKg),
                    SfridoPct: sfridoPct === null ? null : round(sfridoPct)
                },
                products,
                _debug: {
                    candidates: candidates.map(c => ({ OrdNo: c.OrdNo, TrInf4: c.TrInf4, ProdNo: c.ProdNo, NoFin: c.NoFin })),
                    candidateNoFinMapEntries: Array.from(candidateNoFinMap.entries()).map(([k, v]) => ({ key: k, noFin: v })),
                    lookups: _debugLookups
                }
            };
            diskCache.set(laserCacheKey, laserResult, CACHE_TTL_LASER_METRICS_MS);
            return res.json(laserResult);
        } catch (err) {
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/omsaetning/accounts', async (req, res) => {
        try {
            const accounts = await omsaetningService.getAccounts();

            return res.json({
                ok: true,
                accounts
            });
        } catch (err) {
            logEvent('ERROR omsaetning/accounts: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/omsaetning/customers', async (req, res) => {
        try {
            const queryText = String(req.query.q || '').trim();
            const limit = Number(req.query.limit || 20);
            const customers = await omsaetningService.searchCustomers({ queryText, limit });

            return res.json({
                ok: true,
                customers
            });
        } catch (err) {
            logEvent('ERROR omsaetning/customers: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/omsaetning/summary', async (req, res) => {
        try {
            const fra = String(req.query.fra || '').trim();
            const til = String(req.query.til || '').trim();
            const accountCsv = String(req.query.accounts || '').trim();
            const customerCsv = String(req.query.customers || '').trim();
            const summary = await omsaetningService.getSummary({ fra, til, accountCsv, customerCsv });

            return res.json({
                ok: true,
                ...summary
            });
        } catch (err) {
            if (err && err.statusCode) {
                return res.status(err.statusCode).json({ ok: false, error: err.message || 'Ugyldig forespørgsel' });
            }
            logEvent('ERROR omsaetning/summary: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/ordreindgang/summary', async (req, res) => {
        try {
            const fraWeek = String(req.query.fraWeek || '').trim();
            const tilWeek = String(req.query.tilWeek || '').trim();
            const summary = await ordreindgangService.getSummary({ fraWeek, tilWeek });

            return res.json({
                ok: true,
                ...summary
            });
        } catch (err) {
            if (err && err.statusCode) {
                return res.status(err.statusCode).json({ ok: false, error: err.message || 'Ugyldig forespørgsel' });
            }
            logEvent('ERROR ordreindgang/summary: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/belastning/resources', async (req, res) => {
        try {
            const parityRaw = String(req.query.parity || '').trim();
            const parity = parityRaw === '' ? null : (parityRaw === '1' ? 1 : (parityRaw === '0' ? 0 : null));
            const pool = await getConnection();
            const request = pool.request();

            let query = `
                SELECT DISTINCT MainR7, MainR7 + ' - ' + (SELECT Nm FROM R7 r2 WHERE r2.RNo = R7.MainR7) AS R7Nm, Gr10
                FROM R7
                WHERE Gr10 > 0
            `;
            if (parity !== null) {
                request.input('Parity', sql.Int, parity);
                query += ` AND (Gr10 % 2) = @Parity`;
            }
            query += ` ORDER BY MainR7`;

            const result = await request.query(query);
            return res.json({ ok: true, resources: result.recordset || [] });
        } catch (err) {
            logEvent('ERROR belastning/resources: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/belastning/grafisk', async (req, res) => {
        try {
            const today = parseBelastningDate(req.query.toDay) || new Date().toISOString().slice(0, 10);
            const dage = parseBelastningDays(req.query.dage, 30);
            const resGrCsv = normalizeResGrCsv(req.query.resGr);
            const orderNo = normalizeBelastningOrderFilter(req.query.ord);
            const customerFilter = normalizeBelastningCustomerFilter(req.query.kunde);

            const [oddRowsRaw, evenRowsRaw] = await Promise.all([
                fetchBelastningRows({ getConnection, sql, toDay: today, dage, resGrCsv, parity: 1, orderNo, customerFilter }),
                fetchBelastningRows({ getConnection, sql, toDay: today, dage, resGrCsv, parity: 0, orderNo, customerFilter })
            ]);

            const trimRowsForFocusedMode = (rows) => {
                if (!orderNo && !customerFilter) return rows;
                return rows.filter(row => Number(row && row.Resv || 0) > 0 || Number(row && row.Aften || 0) > 0);
            };

            const oddRows = trimRowsForFocusedMode(oddRowsRaw);
            const evenRows = trimRowsForFocusedMode(evenRowsRaw);

            const summarize = (rows) => {
                const map = new Map();
                for (const row of rows) {
                    const key = String(row.ResGr || '').trim();
                    if (!key) continue;
                    if (!map.has(key)) {
                        map.set(key, {
                            resGr: key,
                            nm: String(row.Nm || '').trim(),
                            totalResv: 0,
                            totalKap: 0,
                            totalAften: 0
                        });
                    }
                    const item = map.get(key);
                    item.totalResv += Number(row.Resv || 0);
                    item.totalKap += Number(row.Kap || 0);
                    item.totalAften += Number(row.Aften || 0);
                }
                return Array.from(map.values())
                    .sort((a, b) => a.resGr.localeCompare(b.resGr, 'da'));
            };

            return res.json({
                ok: true,
                toDay: today,
                dage,
                resGr: resGrCsv,
                ord: orderNo,
                kunde: customerFilter,
                odd: {
                    resources: summarize(oddRows),
                    rows: oddRows
                },
                even: {
                    resources: summarize(evenRows),
                    rows: evenRows
                }
            });
        } catch (err) {
            logEvent('ERROR belastning/grafisk: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/belastning/detail', async (req, res) => {
        try {
            const today = parseBelastningDate(req.query.toDay) || new Date().toISOString().slice(0, 10);
            const dage = parseBelastningDays(req.query.dage, 30);
            const resGr = String(req.query.resGr || '').trim();
            const parity = String(req.query.parity || '1').trim() === '0' ? 0 : 1;
            const orderNo = normalizeBelastningOrderFilter(req.query.ord);
            const customerFilter = normalizeBelastningCustomerFilter(req.query.kunde);
            if (!resGr) {
                return res.status(400).json({ ok: false, error: 'resGr er påkrævet' });
            }

            const rows = await fetchBelastningRows({
                getConnection,
                sql,
                toDay: today,
                dage,
                resGrCsv: resGr,
                parity,
                orderNo,
                customerFilter
            });

            const orderRows = await fetchBelastningOrderRows({
                getConnection,
                sql,
                toDay: today,
                dage,
                resGrCsv: resGr,
                parity,
                orderNo,
                customerFilter
            });

            const directSubOrderCsv = intCsvFromValues(orderRows.map(x => x && x.PurcNo));
            const subOrderRowsLevel1 = await fetchBelastningSubOrderRows({
                getConnection,
                sql,
                subOrderCsv: directSubOrderCsv
            });

            const nestedSubOrderCsv = intCsvFromValues(subOrderRowsLevel1.map(x => x && x.NextSubOrdNo));
            const subOrderRowsLevel2 = await fetchBelastningSubOrderRows({
                getConnection,
                sql,
                subOrderCsv: nestedSubOrderCsv
            });

            const subOrderRows = [...subOrderRowsLevel1, ...subOrderRowsLevel2];

            const sourceOrderCsv = intCsvFromValues(orderRows.map(x => x && x.OrdNo));
            const orderLineRows = await fetchBelastningOrderLineRows({
                getConnection,
                sql,
                orderCsv: sourceOrderCsv
            });

            return res.json({ ok: true, toDay: today, dage, parity, resGr, ord: orderNo, kunde: customerFilter, rows, orderRows, subOrderRows, orderLineRows });
        } catch (err) {
            logEvent('ERROR belastning/detail: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/omsaetning/customer-threshold/:custno', (req, res) => {
        const custNo = String(req.params.custno || '').trim();
        if (!/^\d{1,20}$/.test(custNo)) {
            return res.status(400).json({ ok: false, error: 'Ugyldigt kundenummer' });
        }

        const threshold = omsaetningThresholdsService.getThreshold(custNo);
        const meta = omsaetningThresholdsService.getStorageMeta();
        if (!threshold) {
            return res.json({
                ok: true,
                custNo,
                warnThreshold: meta.defaultWarnThreshold,
                goodThreshold: meta.defaultGoodThreshold,
                updatedAt: null,
                exists: false,
                storageFile: meta.filePath
            });
        }

        return res.json({
            ok: true,
            custNo,
            warnThreshold: threshold.warnThreshold,
            goodThreshold: threshold.goodThreshold,
            updatedAt: threshold.updatedAt,
            exists: true,
            storageFile: meta.filePath
        });
    });

    router.post('/omsaetning/customer-threshold/:custno', express.json(), (req, res) => {
        const custNo = String(req.params.custno || '').trim();
        if (!/^\d{1,20}$/.test(custNo)) {
            return res.status(400).json({ ok: false, error: 'Ugyldigt kundenummer' });
        }

        const { warnThreshold, goodThreshold } = req.body || {};
        const saved = omsaetningThresholdsService.setThreshold(custNo, { warnThreshold, goodThreshold });
        const meta = omsaetningThresholdsService.getStorageMeta();
        if (!saved) {
            return res.status(400).json({ ok: false, error: 'Ugyldige tærskelværdier' });
        }

        return res.json({
            ok: true,
            custNo,
            warnThreshold: saved.warnThreshold,
            goodThreshold: saved.goodThreshold,
            updatedAt: saved.updatedAt,
            storageFile: meta.filePath
        });
    });

    router.get('/cache-status', (req, res) => {
        const entries = diskCache.list();
        res.json({ count: entries.length, entries });
    });

    router.get('/health', (req, res) => {
        const activeProfile = settingsService.getActiveProfile();
        res.json({ ok: true, version: pkgVersion, db: activeProfile.database, dbServer: activeProfile.server, dbProfile: activeProfile.id, dbLabel: activeProfile.label });
    });

    // ── Settings: database profiler ─────────────────────────────────────────
    router.get('/settings/db-profiles', (_req, res) => {
        try {
            res.json({ ok: true, ...settingsService.getSettingsSummary() });
        } catch (err) {
            res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.post('/settings/db-profiles/active', express.json(), (req, res) => {
        try {
            const profileId = String((req.body && req.body.profileId) || '').trim();
            if (!profileId) return res.status(400).json({ ok: false, error: 'profileId mangler' });
            const profile = settingsService.setActiveProfile(profileId);
            if (typeof getConnectionModule.resetConnection === 'function') {
                getConnectionModule.resetConnection();
            }
            logEvent('SETTINGS: active DB profile changed to ' + profileId + ' (' + profile.label + ')');
            res.json({ ok: true, activeProfile: profile, ...settingsService.getSettingsSummary() });
        } catch (err) {
            res.status(400).json({ ok: false, error: err.message });
        }
    });

    router.post('/settings/db-profiles/upsert', express.json(), (req, res) => {
        try {
            const profile = settingsService.upsertProfile(req.body || {});
            logEvent('SETTINGS: profile upserted: ' + profile.id);
            res.json({ ok: true, profile, ...settingsService.getSettingsSummary() });
        } catch (err) {
            res.status(400).json({ ok: false, error: err.message });
        }
    });

    router.delete('/settings/db-profiles/:id', (req, res) => {
        try {
            const profileId = String(req.params.id || '').trim();
            settingsService.deleteProfile(profileId);
            logEvent('SETTINGS: profile deleted: ' + profileId);
            res.json({ ok: true, ...settingsService.getSettingsSummary() });
        } catch (err) {
            res.status(400).json({ ok: false, error: err.message });
        }
    });

    router.get('/warmup-status', (req, res) => {
        const done = warmupProgress.cached + warmupProgress.loaded + warmupProgress.failed;
        const pct = warmupProgress.total > 0 ? Math.round((done / warmupProgress.total) * 100) : 100;

        const marginOrdNos = (Array.isArray(orderListCache.data) ? orderListCache.data : [])
            .map(row => Number(row && row.OrdNo))
            .filter(ordNo => Number.isFinite(ordNo));
        let marginDone = 0;
        for (const ordNo of marginOrdNos) {
            if (orderMarginCache.has(ordNo)) {
                marginDone += 1;
                continue;
            }
            const cachedMargin = diskCache.get(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo)
                || diskCache.getStale(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo)
                || diskCache.getStale('order_margin_v6_' + ordNo);
            if (cachedMargin && cachedMargin.totalCost !== null && cachedMargin.totalCost !== undefined) {
                marginDone += 1;
            }
        }
        const marginTotal = marginOrdNos.length;

        const combinedTotal = (warmupProgress.total || 0) + marginTotal;
        const combinedDone = done + marginDone;
        const combinedPct = combinedTotal > 0 ? Math.round((combinedDone / combinedTotal) * 100) : 100;
        const ready = !orderListCache.loading
            && !warmupProgress.running
            && (warmupProgress.total === 0 || done >= warmupProgress.total)
            && (marginTotal === 0 || marginDone >= marginTotal);

        res.json({
            running: warmupProgress.running,
            total: warmupProgress.total,
            cached: warmupProgress.cached,
            loaded: warmupProgress.loaded,
            failed: warmupProgress.failed,
            done,
            pct,
            current: warmupProgress.current,
            marginDone,
            marginTotal,
            combinedDone,
            combinedTotal,
            combinedPct,
            ready
        });
    });

    router.post('/cache-refresh-order/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            if (Number.isNaN(ordNo)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }

            if (!orderRefreshInFlight.has(ordNo)) {
                const refreshPromise = (async () => {
                    logEvent('CACHE REFRESH ORDER: ordNo=' + ordNo + ' start');
                    orderRefreshStatus.set(ordNo, { status: 'running', startedAt: Date.now() });

                    diskCache.del(AFTERCALC_CACHE_KEY_PREFIX + ordNo);
                    for (const prefix of legacyAftercalcPrefixes) {
                        diskCache.del(prefix + ordNo);
                    }
                    diskCache.del('prod_summary_' + ordNo);
                    diskCache.del('prod_summary_' + ordNo + '_gr4_3');
                    diskCache.del(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo);
                    diskCache.del('order_margin_v6_' + ordNo);
                    orderMarginCache.delete(ordNo);
                    orderMarginInFlight.delete(ordNo);
                    afterCalcInFlight.delete(ordNo);

                    const aftercalc = await getOrComputeAftercalc(ordNo, { priority: 'high' });
                    if (aftercalc && !aftercalc.error) {
                        const marginInfo = {
                            ordNo,
                            totalRevenue: Number(aftercalc.summary && aftercalc.summary.totalRevenue || 0),
                            totalCost: Number(aftercalc.summary && aftercalc.summary.totalCost || 0),
                            computedAt: Date.now()
                        };
                        orderMarginCache.set(ordNo, marginInfo);
                        const marginResult = {
                            ordNo,
                            totalRevenue: marginInfo.totalRevenue,
                            totalCost: marginInfo.totalCost,
                            cached: true
                        };
                        diskCache.set(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo, marginResult, CACHE_TTL_ORDER_MARGIN_MS);
                        logEvent('CACHE REFRESH ORDER: ordNo=' + ordNo + ' margin updated');
                        const currentState = orderRefreshStatus.get(ordNo) || {};
                        orderRefreshStatus.set(ordNo, {
                            status: 'done',
                            startedAt: currentState.startedAt || Date.now(),
                            finishedAt: Date.now()
                        });
                    } else {
                        const errMsg = (aftercalc && aftercalc.error) ? aftercalc.error : 'unknown error';
                        const currentState = orderRefreshStatus.get(ordNo) || {};
                        orderRefreshStatus.set(ordNo, {
                            status: 'error',
                            error: errMsg,
                            startedAt: currentState.startedAt || Date.now(),
                            finishedAt: Date.now()
                        });
                    }

                    logEvent('CACHE REFRESH ORDER: ordNo=' + ordNo + ' done');
                })()
                    .catch(err => {
                        const currentState = orderRefreshStatus.get(ordNo) || {};
                        orderRefreshStatus.set(ordNo, {
                            status: 'error',
                            error: err.message,
                            startedAt: currentState.startedAt || Date.now(),
                            finishedAt: Date.now()
                        });
                        logEvent('ERROR cache-refresh-order worker ordNo=' + ordNo + ': ' + err.message);
                    })
                    .finally(() => {
                        orderRefreshInFlight.delete(ordNo);
                    });

                orderRefreshInFlight.set(ordNo, refreshPromise);
            } else {
                logEvent('CACHE REFRESH ORDER: ordNo=' + ordNo + ' already running');
            }

            return res.json({ ok: true, ordNo, started: true });
        } catch (err) {
            logEvent('ERROR cache-refresh-order: ' + err.message);
            return res.status(500).json({ error: err.message });
        }
    });

    router.get('/cache-refresh-order-status/:ordno', (req, res) => {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) {
            return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
        }
        const state = orderRefreshStatus.get(ordNo);
        if (!state) {
            return res.json({ ordNo, status: 'idle' });
        }
        return res.json({ ordNo, ...state });
    });

    router.post('/cache-clear', (req, res) => {
        const deleted = diskCache.clearAll();
        orderMarginCache.clear();
        orderMarginInFlight.clear();
        afterCalcInFlight.clear();
        orderListCache.data = [];
        orderListCache.loadedAt = 0;
        orderListCache.lastError = null;
        warmupProgress.running = false;
        warmupProgress.total = 0;
        warmupProgress.cached = 0;
        warmupProgress.loaded = 0;
        warmupProgress.failed = 0;
        warmupProgress.current = null;
        warmupProgress.startedAt = null;
        warmupProgress.completedAt = null;
        logEvent('CACHE CLEARED: ' + deleted + ' files deleted, in-memory caches reset');

        // Rebuild caches immediately after manual clear so dashboard warmup can continue.
        setTimeout(() => {
            refreshOrderListCache(true)
                .then(() => {
                    logEvent('CACHE CLEAR: forced order-list refresh completed');
                })
                .catch(err => {
                    logEvent('CACHE CLEAR: forced order-list refresh failed: ' + err.message);
                });
        }, 10);

        res.json({ ok: true, deleted });
    });

    router.post('/desktop-update-check', async (req, res) => {
        try {
            const checkFn = global.__desktopManualUpdateCheck;
            if (typeof checkFn !== 'function') {
                return res.status(503).json({ ok: false, status: 'unavailable', message: 'Opdateringskontrol er ikke tilgaengelig i denne mode.' });
            }

            const result = await checkFn();
            logEvent('MANUAL-UPDATE-CHECK: status=' + String(result && result.status || 'unknown') + ', ok=' + String(!!(result && result.ok)));
            return res.json(result || { ok: false, status: 'error', message: 'Tomt svar fra updater.' });
        } catch (err) {
            logEvent('MANUAL-UPDATE-CHECK ERROR: ' + err.message);
            return res.status(500).json({ ok: false, status: 'error', message: err.message });
        }
    });

    router.get('/desktop-update-status', (req, res) => {
        try {
            const statusFn = global.__desktopManualUpdateStatus;
            if (typeof statusFn !== 'function') {
                return res.status(503).json({
                    ok: false,
                    status: 'unavailable',
                    message: 'Opdateringsstatus er ikke tilgaengelig i denne mode.'
                });
            }

            const result = statusFn();
            return res.json(result || {
                ok: false,
                status: 'error',
                message: 'Tomt svar fra updater-status.'
            });
        } catch (err) {
            logEvent('DESKTOP-UPDATE-STATUS ERROR: ' + err.message);
            return res.status(500).json({ ok: false, status: 'error', message: err.message });
        }
    });

    router.post('/desktop-update-install', (req, res) => {
        try {
            const installFn = global.__desktopManualUpdateInstall;
            if (typeof installFn !== 'function') {
                return res.status(503).json({
                    ok: false,
                    status: 'unavailable',
                    message: 'Installering er ikke tilgaengelig i denne mode.'
                });
            }

            const result = installFn();
            logEvent('DESKTOP-UPDATE-INSTALL: status=' + String(result && result.status || 'unknown') + ', ok=' + String(!!(result && result.ok)));
            return res.json(result || {
                ok: false,
                status: 'error',
                message: 'Tomt svar fra install-funktion.'
            });
        } catch (err) {
            logEvent('DESKTOP-UPDATE-INSTALL ERROR: ' + err.message);
            return res.status(500).json({ ok: false, status: 'error', message: err.message });
        }
    });

    router.post('/open-drawing', (req, res) => {
        (async () => {
            try {
                const rawPath = String((req.body && req.body.path) || '').trim();
                const prodNo = String((req.body && req.body.prodNo) || '').trim();
                let candidatePath = rawPath;

                if (!candidatePath && prodNo) {
                    const pool = await getConnection();
                    const drawingRow = await pool.request()
                        .input('prodNo', sql.VarChar(100), prodNo)
                        .query(`
                            SELECT TOP 1 LTRIM(RTRIM(CONVERT(VARCHAR(1000), WebPg))) AS WebPg
                            FROM FreeInf2
                            WHERE LTRIM(RTRIM(CONVERT(VARCHAR(100), ProdNo))) = @prodNo
                              AND WebPg IS NOT NULL
                              AND LTRIM(RTRIM(CONVERT(VARCHAR(1000), WebPg))) <> ''
                            ORDER BY LTRIM(RTRIM(CONVERT(VARCHAR(1000), WebPg))) DESC
                        `);

                    const webPg = String((drawingRow.recordset && drawingRow.recordset[0] && drawingRow.recordset[0].WebPg) || '').trim();
                    if (webPg) {
                        candidatePath = webPg;
                    }
                }

                if (!candidatePath) {
                    return res.status(400).json({ ok: false, message: 'Path mangler.' });
                }

                const lower = candidatePath.toLowerCase();
                if (lower.indexOf('.pdf') === -1) {
                    return res.status(400).json({ ok: false, message: 'Kun PDF er tilladt.' });
                }

                const child = spawn('cmd', ['/c', 'start', '', candidatePath], {
                    windowsHide: true,
                    detached: true,
                    stdio: 'ignore'
                });
                child.unref();

                logEvent('OPEN-DRAWING: ' + candidatePath + (prodNo ? (' [prodNo=' + prodNo + ']') : ''));
                return res.json({ ok: true });
            } catch (err) {
                logEvent('OPEN-DRAWING ERROR: ' + err.message);
                return res.status(500).json({ ok: false, message: err.message });
            }
        })();
    });

    router.get('/prodtr/:ordno/:lnno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            const lnNo = parseInt(req.params.lnno);
            if (Number.isNaN(ordNo) || Number.isNaN(lnNo)) {
                return res.status(400).json({ error: 'Ugyldige parametre' });
            }

            const lnNosFromQuery = String(req.query.lnNos || '')
                .split(',')
                .map(v => parseInt(String(v || '').trim(), 10))
                .filter(v => Number.isFinite(v) && v > 0);
            const lnNoSet = new Set([lnNo, ...lnNosFromQuery]);
            const lnNos = Array.from(lnNoSet);

            const pool = await getConnection();
            const request = pool.request()
                .input('ordNo', sql.Numeric, ordNo);

            let whereLineFilter = 'P.OrdLnNo = @lnNo';
            if (lnNos.length <= 1) {
                request.input('lnNo', sql.Numeric, lnNo);
            } else {
                const placeholders = [];
                lnNos.forEach((lineNo, idx) => {
                    const key = 'lnNo' + idx;
                    request.input(key, sql.Numeric, lineNo);
                    placeholders.push('@' + key);
                });
                whereLineFilter = 'P.OrdLnNo IN (' + placeholders.join(', ') + ')';
            }

            const result = await request.query(`
                    SELECT
                        P.FinDt,
                        P.FinTm,
                        P.NoInvoAb,
                        A.Nm AS HvemNm
                    FROM ProdTr P
                    LEFT JOIN Actor A ON A.EmpNo = P.EmpNo
                    WHERE P.OrdNo = @ordNo AND ${whereLineFilter}
                    ORDER BY P.FinDt DESC, P.FinTm DESC
                `);
            res.json(result.recordset);
        } catch (err) {
            res.status(500).json({ error: err.message });
        }
    });

    // ── Efterkalk: kundefaktura-oversigt ─────────────────────────────────────
    router.get('/efterkalk/customers', async (req, res) => {
        try {
            const q = String(req.query.q || '').trim();
            const pool = await getConnection();
            const request = pool.request();
            let where = 'A.CustNo <> 0';
            if (q) {
                request.input('q', sql.VarChar, '%' + q + '%');
                where += " AND (A.Nm LIKE @q OR CAST(A.CustNo AS VARCHAR(20)) LIKE @q OR A.Shrt LIKE @q)";
            }
            const result = await request.query(`
                SELECT DISTINCT TOP 60
                    A.CustNo, A.Nm, A.Shrt
                FROM Actor A
                WHERE ${where}
                  AND EXISTS (
                      SELECT 1 FROM Ord O
                      WHERE O.CustNo = A.CustNo
                        AND O.InvoNo IS NOT NULL AND O.InvoNo <> ''
                        AND O.InvoAm > 0
                  )
                ORDER BY A.Nm
            `);
            res.json({ ok: true, rows: result.recordset || [] });
        } catch (err) {
            logEvent('ERROR efterkalk/customers: ' + err.message);
            res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/efterkalk/customer-invoices', async (req, res) => {
        try {
            const custNo = parseInt(req.query.custno);
            if (Number.isNaN(custNo) || custNo <= 0) {
                return res.status(400).json({ ok: false, error: 'custno ugyldigt' });
            }
            // from/to as YYYY-MM-DD, stored in Visma as INT YYYYMMDD
            const fromStr = String(req.query.from || '').trim();
            const toStr   = String(req.query.to   || '').trim();
            if (!fromStr) return res.status(400).json({ ok: false, error: 'from dato mangler' });
            const fromInt = parseInt(fromStr.replace(/-/g, ''));
            const toInt   = toStr ? parseInt(toStr.replace(/-/g, '')) : parseInt(new Date().toISOString().slice(0,10).replace(/-/g,''));
            if (Number.isNaN(fromInt) || Number.isNaN(toInt)) {
                return res.status(400).json({ ok: false, error: 'Ugyldig dato' });
            }

            const pool = await getConnection();
            const result = await pool.request()
                .input('custNo',   sql.Int, custNo)
                .input('fromDate', sql.Int, fromInt)
                .input('toDate',   sql.Int, toInt)
                .query(`
                    SELECT
                        O.OrdNo,
                        O.LstInvDt,
                        O.InvoAm,
                        O.Gr4,
                        O.InvoNo,
                        A_cust.Nm   AS CustomerName,
                        A_cust.Shrt AS CustomerShrt,
                        SU.Usr      AS SellerUsr
                    FROM Ord O
                    LEFT JOIN Actor A_cust ON A_cust.CustNo = O.CustNo
                    OUTER APPLY (
                        SELECT TOP 1 A.Usr FROM Actor A
                        WHERE LTRIM(RTRIM(CONVERT(VARCHAR(50), A.EmpNo))) = LTRIM(RTRIM(CONVERT(VARCHAR(50), O.SelBuy)))
                    ) SU
                    WHERE O.CustNo  = @custNo
                      AND O.InvoNo IS NOT NULL AND O.InvoNo <> ''
                      AND O.InvoAm  > 0
                      AND O.LstInvDt >= @fromDate
                      AND O.LstInvDt <= @toDate
                    ORDER BY O.LstInvDt DESC, O.OrdNo DESC
                `);
            const rows = (result.recordset || []).map(r => ({
                OrdNo:        r.OrdNo,
                LstInvDt:     r.LstInvDt,
                InvoAm:       Number(r.InvoAm || 0),
                Gr4:          r.Gr4,
                InvoNo:       r.InvoNo,
                CustomerName: r.CustomerName,
                CustomerShrt: r.CustomerShrt,
                SellerUsr:    r.SellerUsr
            }));
            const totalInvoAm = rows.reduce((s, r) => s + r.InvoAm, 0);
            res.json({ ok: true, rows, count: rows.length, totalInvoAm, custNo, fromInt, toInt });
        } catch (err) {
            logEvent('ERROR efterkalk/customer-invoices: ' + err.message);
            res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/order-list-check-time', async (req, res) => {
        try {
            const pool = await getConnection();
            const result = await pool.request().query(`
                SELECT MAX(CAST(LstInvDt AS INT)) as maxInvDate
                FROM Ord
                WHERE CAST(CAST(LstInvDt AS CHAR(8)) AS INT)
                    >= CONVERT(INT, FORMAT(DATEADD(DAY, -${ORDER_LIST_DAYS_BACK}, GETDATE()), 'yyyyMMdd'))
            `);

            const maxDate = result.recordset[0]?.maxInvDate || 0;
            const serverTime = Date.now();

            res.json({
                lastModifiedDate: maxDate,
                serverTime: serverTime,
                cacheLastModified: orderListCache.lastModifiedTime
            });
        } catch (err) {
            logEvent('ERROR order-list-check-time: ' + err.message);
            res.status(500).json({ error: err.message });
        }
    });

    router.get('/order-list', async (req, res) => {
        try {
            const forceRefresh = req.query.force === '1';
            logEvent('ORDER-LIST: force=' + (forceRefresh ? '1' : '0'));

            if (forceRefresh) {
                await refreshOrderListCache(true);
            } else if (!isOrderListCacheFresh()) {
                await refreshOrderListCache();
            }

            let marginFromMemory = 0;
            let marginFromDisk = 0;
            let marginFromDiskStale = 0;
            let marginMissing = 0;

            const rowHasWarnings = (aftercalc) => {
                if (!aftercalc || typeof aftercalc !== 'object') return false;
                if (aftercalc.hasWarnings === true) return true;

                const lineHasWarning = (lines) => Array.isArray(lines) && lines.some(line => {
                    if (!line || typeof line !== 'object') return false;
                    if (line.HasWarning) return true;
                    const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                    const noFinValue = Number(line.NoFin || 0);
                    const noOrgValue = Number(line.NoOrg || 0);
                    return prodNoKey.startsWith('3') && noFinValue === 0 && noOrgValue > 0;
                });

                const prodOrderHasWarning = Array.isArray(aftercalc.productionOrders)
                    && aftercalc.productionOrders.some(order => order && (order.hasWarnings || lineHasWarning(order.lines)));

                return lineHasWarning(aftercalc.salesOrderLines)
                    || lineHasWarning(aftercalc.salesLines)
                    || prodOrderHasWarning;
            };

            const data = orderListCache.data.map(row => {
                const ordNoNum = Number(row.OrdNo);
                let marginInfo = orderMarginCache.get(ordNoNum);
                let warningSource = diskCache.get(AFTERCALC_CACHE_KEY_PREFIX + ordNoNum)
                    || diskCache.getStale(AFTERCALC_CACHE_KEY_PREFIX + ordNoNum)
                    || diskCache.get('aftercalc_' + ordNoNum)
                    || diskCache.getStale('aftercalc_' + ordNoNum);
                const hasWarning = rowHasWarnings(warningSource);
                if (marginInfo) {
                    marginFromMemory += 1;
                }

                if (!marginInfo) {
                    const cachedMargin = diskCache.get(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNoNum);
                    if (cachedMargin && cachedMargin.totalCost !== null && cachedMargin.totalCost !== undefined) {
                        marginInfo = {
                            ordNo: ordNoNum,
                            totalRevenue: Number(cachedMargin.totalRevenue || row.InvoAm || 0),
                            totalCost: Number(cachedMargin.totalCost || 0),
                            computedAt: Date.now()
                        };
                        orderMarginCache.set(ordNoNum, marginInfo);
                        marginFromDisk += 1;
                    }
                }

                if (!marginInfo) {
                    const staleMargin = diskCache.getStale(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNoNum)
                        || diskCache.getStale('order_margin_v6_' + ordNoNum);
                    if (staleMargin && staleMargin.totalCost !== null && staleMargin.totalCost !== undefined) {
                        marginInfo = {
                            ordNo: ordNoNum,
                            totalRevenue: Number(staleMargin.totalRevenue || row.InvoAm || 0),
                            totalCost: Number(staleMargin.totalCost || 0),
                            computedAt: Date.now()
                        };
                        orderMarginCache.set(ordNoNum, marginInfo);
                        marginFromDiskStale += 1;
                    }
                }

                if (!marginInfo) {
                    marginMissing += 1;
                }

                return {
                    ...row,
                    HasWarning: hasWarning,
                    WarningText: hasWarning ? 'Ordren indeholder mindst én advarsel.' : '',
                    TotalCost: marginInfo ? marginInfo.totalCost : null
                };
            });

            orderListCache.lastModifiedTime = Date.now();

            if (!isOrderListCacheFresh() && !orderListCache.loading) {
                refreshOrderListCache(true).catch(err => {
                    logEvent('ERROR order-list refresh: ' + err.message);
                });
            }

            logEvent('ORDER-LIST: returned ' + data.length + ' rows (margin memory=' + marginFromMemory + ', disk=' + marginFromDisk + ', stale=' + marginFromDiskStale + ', missing=' + marginMissing + ')');
            res.json(data);
        } catch (err) {
            logEvent('ERROR order-list: ' + err.message);
            res.status(500).json({ error: err.message });
        }
    });

    // ── ORDER NOTES ─────────────────────────────────────────────────────────
    router.get('/order-note/:ordno', (req, res) => {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) return res.status(400).json({ error: 'Ugyldigt ordrenummer' });
        const note = orderNotesService.getNote(ordNo);
        res.json(note || { status: '', text: '', isCreditNote: false, isUB: false, updatedAt: null });
    });

    router.get('/order-notes-all', (req, res) => {
        res.json(orderNotesService.getAllNotes());
    });

    router.post('/order-note/:ordno', express.json(), (req, res) => {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) return res.status(400).json({ error: 'Ugyldigt ordrenummer' });
        const { status = '', text = '', isCreditNote = false, isUB = false } = req.body || {};
        const note = orderNotesService.setNote(ordNo, { status, text, isCreditNote, isUB });
        res.json(note || { status: '', text: '', isCreditNote: false, isUB: false, updatedAt: null });
    });

    router.delete('/order-note/:ordno', (req, res) => {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) return res.status(400).json({ error: 'Ugyldigt ordrenummer' });
        orderNotesService.deleteNote(ordNo);
        res.json({ ok: true });
    });

    // ── Personalehåndbog search API ─────────────────────────────────────────
    router.get('/ph/status', (_req, res) => {
        res.json({ status: phStatus, count: phIndex.length, indexedAt: phIndexedAt, error: phError });
    });

    router.get('/ph/search', (req, res) => {
        const q = String(req.query.q || '').toLowerCase().trim();
        if (!q) return res.json({ results: [], status: phStatus });
        if (phStatus !== 'ready') return res.json({ results: [], status: phStatus });
        const terms = q.split(/\s+/).filter(Boolean);
        const results = [];
        for (const page of phIndex) {
            const haystack = (page.title + ' ' + page.text).toLowerCase();
            const score = terms.reduce((s, t) => s + (haystack.split(t).length - 1), 0);
            if (score === 0) continue;
            const firstTerm = terms[0];
            const idx = page.text.toLowerCase().indexOf(firstTerm);
            let snippet = '';
            if (idx >= 0) {
                const s = Math.max(0, idx - 80);
                const e = Math.min(page.text.length, idx + 200);
                snippet = (s > 0 ? '…' : '') + page.text.slice(s, e) + (e < page.text.length ? '…' : '');
            } else {
                snippet = page.text.slice(0, 240) + '…';
            }
            results.push({ url: page.url, title: page.title || page.url, snippet, score });
        }
        results.sort((a, b) => b.score - a.score);
        res.json({ results: results.slice(0, 40), status: phStatus });
    });

    router.post('/ph/reindex', (_req, res) => {
        if (phStatus === 'indexing') return res.json({ ok: false, msg: 'Allerede i gang' });
        crawlPH().catch(e => { phStatus = 'error'; phError = e.message; });
        res.json({ ok: true });
    });

    router.get('/qms/dataset', (_req, res) => {
        try {
            const dataset = readQmsDataset(fs);
            res.json({ ok: true, dataset });
        } catch (err) {
            res.status(500).json({ ok: false, error: err.message || 'QMS dataset fejl' });
        }
    });

    router.get('/bom/customers', async (req, res) => {
        try {
            const q = String(req.query.q || '').trim();
            const limit = req.query.limit === undefined ? undefined : Number(req.query.limit);
            const payload = await bomService.fetchCustomers({ q, limit });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/customers: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM customers fejl' });
        }
    });

    router.get('/bom/products', async (req, res) => {
        try {
            const customerNo = String(req.query.customerNo || req.query.cust || '').trim();
            const customerCode = String(req.query.customerCode || req.query.gr || '').trim();
            if (!customerNo) {
                return res.status(400).json({ error: 'customerNo er paakraevet' });
            }
            const limit = req.query.limit === undefined ? undefined : Number(req.query.limit);
            const payload = await bomService.fetchProductsByCustomer({ customerNo, customerCode, limit });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/products: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM products fejl' });
        }
    });

    router.get('/bom/revisions/by-drawing', async (req, res) => {
        try {
            const tgn = String(req.query.tgn || '').trim();
            const customerNo = String(req.query.customerNo || req.query.cust || '').trim();
            const customerCode = String(req.query.customerCode || req.query.gr || '').trim();
            if (!tgn || !customerNo) {
                return res.status(400).json({ error: 'tgn og customerNo er paakraevet' });
            }
            const payload = await bomService.fetchRevisionsByDrawing({ tgn, customerNo, customerCode });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/revisions/by-drawing: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM revisions fejl' });
        }
    });

    router.get('/bom/resources', async (_req, res) => {
        try {
            const payload = await bomService.fetchResources();
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/resources: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM resources fejl' });
        }
    });

    router.get('/bom/materials', async (req, res) => {
        try {
            const q = String(req.query.q || '').trim();
            const limit = Number(req.query.limit || 2500);
            const payload = await bomService.fetchMaterials({ q, limit });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/materials: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM materials fejl' });
        }
    });

    router.get('/bom/calculators/laser-params', async (req, res) => {
        try {
            const machine = String(req.query.machine || '').trim();
            const payload = await bomService.fetchLaserParameters({ machine });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/calculators/laser-params: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM laser params fejl' });
        }
    });

    router.get('/bom/calculators/process-params', async (_req, res) => {
        try {
            const payload = await bomService.fetchProcessParameters();
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/calculators/process-params: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM process params fejl' });
        }
    });

    router.get('/bom/components', async (req, res) => {
        try {
            const q = String(req.query.q || '').trim();
            const limit = req.query.limit === undefined ? undefined : Number(req.query.limit);
            const payload = await bomService.fetchComponents({ q, limit });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/components: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM components fejl' });
        }
    });

    router.get('/bom/customer-notes', async (req, res) => {
        try {
            const customerCode = String(req.query.customerCode || req.query.gr || '').trim();
            if (!customerCode) {
                return res.status(400).json({ error: 'customerCode er paakraevet' });
            }
            const payload = await bomService.fetchCustomerNotes({ customerCode });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/customer-notes: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM customer notes fejl' });
        }
    });

    router.get('/bom/suppliers', async (req, res) => {
        try {
            const q = String(req.query.q || '').trim();
            const payload = await bomService.fetchSuppliers({ q });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/suppliers: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM suppliers fejl' });
        }
    });

    router.get('/bom/product-tree', async (req, res) => {
        try {
            const prodNo = String(req.query.prodNo || '').trim();
            if (!prodNo) {
                return res.status(400).json({ error: 'prodNo er paakraevet' });
            }
            const payload = await bomService.fetchProductTree({ prodNo });
            res.json(payload);
        } catch (err) {
            logEvent('ERROR bom/product-tree: ' + err.message);
            res.status(500).json({ error: err.message || 'BOM product tree fejl' });
        }
    });

    router.post('/bom/calc/nesting', express.json(), (req, res) => {
        try {
            const result = bomService.computeNesting(req.body || {});
            res.json(result);
        } catch (err) {
            res.status(400).json({ error: err.message || 'Nesting beregning fejl' });
        }
    });

    router.post('/bom/calc/quote', express.json(), async (req, res) => {
        try {
            const result = await bomService.computeQuote(req.body || {});
            res.json(result);
        } catch (err) {
            logEvent('ERROR bom/calc/quote: ' + err.message);
            res.status(400).json({ error: err.message || 'Prisberegning fejl' });
        }
    });

    router.post('/bom/analyze-file', express.json({ limit: '40mb' }), (req, res) => {
        try {
            const filename = String((req.body && req.body.filename) || '').trim();
            const dataBase64 = (req.body && req.body.data) || '';
            if (!filename || !dataBase64) {
                return res.status(400).json({ error: 'filename og data (base64) er paakraevet' });
            }
            const buffer = Buffer.from(dataBase64, 'base64');
            if (buffer.length === 0) {
                return res.status(400).json({ error: 'Tom fil' });
            }
            const result = bomService.analyzeDrawingFile(filename, buffer);
            res.json({ filename, sizeBytes: buffer.length, ...result });
        } catch (err) {
            logEvent('ERROR bom/analyze-file: ' + err.message);
            res.status(400).json({ error: err.message || 'Filanalyse fejl' });
        }
    });

    router.post('/bom/cache/invalidate', (req, res) => {
        try {
            const scope = String((req.body && req.body.scope) || req.query.scope || 'all');
            const result = bomService.invalidate(scope);
            res.json({ ok: true, ...result });
        } catch (err) {
            logEvent('ERROR bom/cache/invalidate: ' + err.message);
            res.status(500).json({ ok: false, error: err.message || 'BOM cache invalidate fejl' });
        }
    });

    // ── BOM: Opret produkter i Visma ────────────────────────────────────────
    router.post('/bom/create-products/preview', express.json(), async (req, res) => {
        try {
            const result = await bomService.previewCreateProducts(req.body || {});
            res.json({ ok: true, ...result });
        } catch (err) {
            logEvent('ERROR bom/create-products/preview: ' + err.message);
            const status = err.statusCode || 400;
            res.status(status).json({ ok: false, error: err.message });
        }
    });

    router.post('/bom/create-products/execute', express.json(), async (req, res) => {
        try {
            const result = await bomService.createProductsInVisma(req.body || {});
            logEvent('BOM CREATE: ' + (result.created || []).map(r => r.ProdNo).join(', '));
            res.json({ ok: true, ...result });
        } catch (err) {
            logEvent('ERROR bom/create-products/execute: ' + err.message);
            const status = err.statusCode || 500;
            res.status(status).json({ ok: false, error: err.message });
        }
    });

    router.put('/qms/dataset', (req, res) => {
        try {
            const dataset = req.body && req.body.dataset;
            const validationError = validateQmsDataset(dataset);
            if (validationError) {
                return res.status(400).json({ ok: false, error: validationError });
            }
            const saved = writeQmsDataset(fs, dataset);
            res.json({ ok: true, dataset: saved });
        } catch (err) {
            res.status(500).json({ ok: false, error: err.message || 'Kunne ikke gemme QMS dataset' });
        }
    });
    // ────────────────────────────────────────────────────────────────────────

    return router;
}

module.exports = {
    createApiRouter
};
