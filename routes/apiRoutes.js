const express = require('express');
const path = require('path');
const crypto = require('crypto');
const orderNotesService = require('../services/orderNotesService');
const aftercalcCostExclusionsService = require('../services/aftercalcCostExclusionsService');
const lagerlisteReconciliationService = require('../services/lagerlisteReconciliationService');
const phCrawler = require('../services/phCrawlerService');
const { readQmsDataset, validateQmsDataset, writeQmsDataset } = require('../services/qmsService');
const {
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
} = require('../services/belastningService');

const omsaetningThresholdsService = require('../services/omsaetningThresholdsService');
const omsaetningDailyThresholds = require('../assets/js/omsaetning-daily-thresholds');
const { createAuthService } = require('../services/authService');
const { fetchSalgordreViaRows } = require('../services/viaService');
const { createLagerlisteService } = require('../services/lagerlisteService');
const { createLagerliste2Service } = require('../services/lagerliste2Service');
const settingsService = require('../services/settingsService');
const getConnectionModule = require('../db');
const { createOmsaetningService } = require('../services/omsaetningService');
const { createOrdreindgangService } = require('../services/ordreindgangService');
const { createBomService } = require('../services/bomService');
const { openPdfTarget } = require('../services/pdfOpenService');

function createApiRouter({
    getConnection,
    sql,
    fs,
    spawn,
    diskCache,
    gohData,
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
    const {
        authSessions,
        readUsers,
        writeUsers,
        hydrateUsersFromDb,
        safeUser,
        makePasswordHash,
        getSessionUser,
        requireAuthenticated,
        requireSuperadmin,
        requireModulePermission,
        revokeSession,
        buildSessionCookie,
        buildExpiredSessionCookie,
        sessionTtlMs
    } = createAuthService({ fs, usersFile: path.join(__dirname, '..', 'users.json') });

    // Stato condiviso da GOH: DB vince sul file locale, fail-soft se non raggiungibile
    Promise.allSettled([
        hydrateUsersFromDb(),
        orderNotesService.hydrateFromDb(),
        aftercalcCostExclusionsService.hydrateFromDb(),
        omsaetningThresholdsService.hydrateFromDb(),
        lagerlisteReconciliationService.hydrateFromDb()
    ]).then(results => {
        const hydrated = results.filter(r => r.status === 'fulfilled' && r.value === true).length;
        logEvent('GOH-STATE: ' + hydrated + '/5 delte tilstande hentet fra GOH (users/noter/efterkalk-kost/graenser/afstemninger)');
    });

    const VIA_CACHE_KEY = 'salgordre_via_v32';
    const VIA_CACHE_TTL_MS = 5 * 60 * 1000;
    const VIA_WARM_INTERVAL_MS = 10 * 60 * 1000;
    let viaWarmRunning = false;
    async function warmSalgordreVia(label) {
        if (viaWarmRunning) return;
        viaWarmRunning = true;
        try {
            const rows = await fetchSalgordreViaRows({ getConnection, sql, requestedOrdNo: null });
            diskCache.set(VIA_CACHE_KEY, { rows }, VIA_CACHE_TTL_MS);
            logEvent('WARM-VIA (' + label + '): ' + rows.length + ' rækker cachet');
        } catch (err) {
            logEvent('WARM-VIA ERROR (' + label + '): ' + err.message);
        } finally {
            viaWarmRunning = false;
        }
    }
    // Startup con lieve ritardo (dopo order-list) + refresh periodico in background
    setTimeout(() => { warmSalgordreVia('startup'); }, 20000);
    setInterval(() => { warmSalgordreVia('interval'); }, VIA_WARM_INTERVAL_MS);

    router.post('/auth/login', express.json(), (req, res) => {
        const username = String(req.body && req.body.username || '').trim().toLowerCase();
        const password = String(req.body && req.body.password || '');
        const users = readUsers();
        const user = users.find(item => String(item.username || '').toLowerCase() === username);
        const bootstrapAdmin = username === 'admin' && password === '12345';
        const validHash = user && user.passwordSalt && user.passwordHash
            && crypto.timingSafeEqual(Buffer.from(makePasswordHash(password, user.passwordSalt), 'hex'), Buffer.from(user.passwordHash, 'hex'));
        if ((!bootstrapAdmin && !validHash) || !user || user.active === false) {
            return res.status(401).json({ error: 'Forkert brugernavn eller kode' });
        }
        const normalized = safeUser(user);
        if (bootstrapAdmin) normalized.role = 'superadmin';
        const token = crypto.randomBytes(32).toString('hex');
        authSessions.set(token, { user: normalized, expiresAt: Date.now() + sessionTtlMs });
        res.setHeader('Set-Cookie', buildSessionCookie(token));
        return res.json({ token, user: normalized });
    });

    router.post('/auth/logout', (req, res) => {
        revokeSession(req);
        res.setHeader('Set-Cookie', buildExpiredSessionCookie());
        return res.json({ ok: true });
    });

    router.get('/admin/users', (req, res) => {
        if (!requireSuperadmin(req, res)) return;
        return res.json({ users: readUsers().map(safeUser) });
    });

    router.post('/admin/users', express.json(), (req, res) => {
        if (!requireSuperadmin(req, res)) return;
        const username = String(req.body && req.body.username || '').trim().toLowerCase();
        const displayName = String(req.body && req.body.displayName || username).trim();
        const password = String(req.body && req.body.password || '');
        if (!username || !password) {
            return res.status(400).json({ error: 'Brugernavn og kode skal udfyldes' });
        }
        const users = readUsers();
        if (users.some(item => String(item.username || '').toLowerCase() === username)) {
            return res.status(409).json({ error: 'Brugernavn findes allerede' });
        }
        const salt = crypto.randomBytes(16).toString('hex');
        const user = { username, displayName, passwordSalt: salt, passwordHash: makePasswordHash(password, salt), role: 'user', active: true, permissions: {}, createdAt: new Date().toISOString() };
        users.push(user);
        writeUsers(users);
        return res.status(201).json({ user: safeUser(user) });
    });

    router.put('/admin/users/:username', express.json(), (req, res) => {
        if (!requireSuperadmin(req, res)) return;
        const username = String(req.params.username || '').toLowerCase();
        const users = readUsers();
        const user = users.find(item => String(item.username || '').toLowerCase() === username);
        if (!user) return res.status(404).json({ error: 'Bruger findes ikke' });
        user.displayName = String(req.body && req.body.displayName || user.displayName || user.username).trim();
        user.active = req.body && req.body.active !== false;
        user.role = username === 'admin' ? 'superadmin' : (req.body && req.body.role === 'superadmin' ? 'superadmin' : 'user');
        user.permissions = user.role === 'superadmin' ? {} : (req.body && req.body.permissions && typeof req.body.permissions === 'object' ? req.body.permissions : {});
        if (req.body && req.body.password) {
            user.passwordSalt = crypto.randomBytes(16).toString('hex');
            user.passwordHash = makePasswordHash(String(req.body.password), user.passwordSalt);
        }
        writeUsers(users);
        return res.json({ user: safeUser(user) });
    });

    function deleteAdminUserHandler(req, res) {
        try {
            if (!requireSuperadmin(req, res)) return;
            const username = String(req.params.username || '').toLowerCase();
            if (username === 'admin') return res.status(400).json({ error: 'Bootstrap superadmin kan ikke slettes' });
            const users = readUsers();
            const remaining = users.filter(item => String(item.username || '').toLowerCase() !== username);
            if (remaining.length === users.length) return res.status(404).json({ error: 'Bruger findes ikke' });
            writeUsers(remaining);
            for (const [token, session] of authSessions.entries()) {
                if (String(session.user && session.user.username || '').toLowerCase() === username) authSessions.delete(token);
            }
            return res.json({ ok: true });
        } catch (err) {
            logEvent('ERROR admin/users DELETE: ' + err.message);
            return res.status(500).json({ error: err.message || 'Kunne ikke slette bruger' });
        }
    }

    router.delete('/admin/users/:username', deleteAdminUserHandler);
    router.post('/admin/users/:username/delete', deleteAdminUserHandler);
    const legacyAftercalcPrefixes = ['aftercalc_v22_', 'aftercalc_v21_', 'aftercalc_v20_', 'aftercalc_v19_', 'aftercalc_v18_', 'aftercalc_v17_', 'aftercalc_'];
    const omsaetningService = createOmsaetningService({ getConnection, sql });
    const ordreindgangService = createOrdreindgangService({ getConnection, sql });
    const bomService = createBomService({
        getConnection,
        sql,
        diskCache,
        logEvent,
        getActiveProfile: settingsService.getActiveProfile
    });
    const lagerlisteService = createLagerlisteService({
        getConnection,
        sql,
        diskCache,
        gohData,
        fs,
        getSalgordreViaRows: fetchSalgordreViaRows,
        getOrComputeAftercalc,
        getProductionSummary,
        getRestPrices: settingsService.getRestPrices,
        dataDir: process.env.GANTECH_DATA_DIR ? path.join(process.env.GANTECH_DATA_DIR, 'lagerliste') : undefined
    });
    const lagerliste2Service = createLagerliste2Service({
        getConnection,
        sql,
        getRestPrices: settingsService.getRestPrices
    });
    lagerlisteService.scheduleMonthlySnapshot({ onError: err => logEvent('ERROR lagerliste monthly snapshot: ' + err.message) });

    router.get('/aftercalc/:ordno', async (req, res) => {
        try {
            const ordNo = parseInt(req.params.ordno);
            const cachedOnly = String(req.query && req.query.cached || '') === '1';
            if (!cachedOnly) logEvent('SEARCH: OrdNo=' + ordNo);
            const forceRefresh = String(req.query && req.query.force || '') === '1';
            const data = await getOrComputeAftercalc(ordNo, { priority: 'high', forceRefresh, cachedOnly });
            if (cachedOnly && !data) {
                return res.json({ notCached: true });
            }
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

    router.get('/salgordre-via', requireModulePermission('salgordreVia'), async (req, res) => {
        try {
            const requestedOrdNo = req.query.ordNo === undefined ? null : Number(req.query.ordNo);
            if (requestedOrdNo !== null && (!Number.isInteger(requestedOrdNo) || requestedOrdNo <= 0)) {
                return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
            }

            const cacheKey = VIA_CACHE_KEY;
            // Solo cache (anche scaduta): risposta immediata per stale-while-revalidate
            if (requestedOrdNo === null && String(req.query.cached || '') === '1') {
                const freshCached = diskCache.get(cacheKey);
                if (freshCached) return res.json({ ...freshCached, cached: true, fresh: true });
                const staleCached = diskCache.getStale(cacheKey);
                if (staleCached) return res.json({ ...staleCached, cached: true, fresh: false });
                return res.json({ notCached: true });
            }
            if (requestedOrdNo === null && req.query.force !== '1') {
                const cached = diskCache.get(cacheKey);
                if (cached) return res.json({ ...cached, cached: true });
            }
            if (requestedOrdNo !== null && req.query.force === '1') {
                diskCache.del(cacheKey);
            }

            const rows = await fetchSalgordreViaRows({ getConnection, sql, requestedOrdNo });
            const payload = { rows };
            if (requestedOrdNo === null) diskCache.set(cacheKey, payload, VIA_CACHE_TTL_MS);
            res.json(payload);
        } catch (err) {
            logEvent('ERROR salgordre-via: ' + err.message);
            res.status(500).json({ error: err.message || 'SalgOrdre VIA fejl' });
        }
    });

    router.get('/ordreindgang/holiday-settings', async (req, res) => {
        try {
            const state = await gohData.getAppState('ordreindgang_holiday_settings');
            return res.json({ ok: true, settings: state ? state.payload : null });
        } catch (err) {
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.post('/ordreindgang/holiday-settings', express.json(), async (req, res) => {
        try {
            const settings = {
                holidayWeeksText: String(req.body && req.body.holidayWeeksText || '').slice(0, 2000),
                ignoreHolidayWeeks: !(req.body && req.body.ignoreHolidayWeeks === false)
            };
            const saved = await gohData.setAppState('ordreindgang_holiday_settings', settings);
            return res.json({ ok: saved, settings });
        } catch (err) {
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/omsaetning/working-days', async (req, res) => {
        try {
            const state = await gohData.getAppState('omsaetning_working_days');
            const months = omsaetningDailyThresholds.normalizeWorkingDaysByMonth(state && state.payload && state.payload.months);
            return res.json({ ok: true, months, updatedAt: state && state.payload ? state.payload.updatedAt || null : null });
        } catch (err) {
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/omsaetning/daily-budget-settings', async (req, res) => {
        try {
            const state = await gohData.getAppState('omsaetning_daily_budget_settings');
            const settings = omsaetningDailyThresholds.sanitizeConfig(state && state.payload);
            return res.json({
                ok: true,
                useDailyBudget: settings.useDailyBudget,
                dailyBreakEvenDkk: settings.dailyBreakEvenDkk,
                dailyBudgetDkk: settings.dailyBudgetDkk,
                updatedAt: state && state.payload ? state.payload.updatedAt || null : null
            });
        } catch (err) {
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.post('/omsaetning/daily-budget-settings', express.json(), async (req, res) => {
        const user = getSessionUser(req);
        if (!user) return res.status(401).json({ ok: false, error: 'Login kræves' });
        try {
            const settings = omsaetningDailyThresholds.sanitizeConfig(req.body);
            const payload = {
                useDailyBudget: settings.useDailyBudget,
                dailyBreakEvenDkk: settings.dailyBreakEvenDkk,
                dailyBudgetDkk: settings.dailyBudgetDkk,
                updatedAt: new Date().toISOString(),
                updatedBy: String(user.username || '')
            };
            const saved = await gohData.setAppState('omsaetning_daily_budget_settings', payload);
            if (!saved) return res.status(503).json({ ok: false, error: 'Dagsmål kunne ikke gemmes i GOH-databasen' });
            return res.json({ ok: true, ...payload });
        } catch (err) {
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/admin/working-days', async (req, res) => {
        if (!requireSuperadmin(req, res)) return;
        try {
            const year = Number(req.query.year);
            if (!Number.isInteger(year) || year < 2000 || year > 2200) {
                return res.status(400).json({ ok: false, error: 'Ugyldigt år' });
            }
            const state = await gohData.getAppState('omsaetning_working_days');
            const allMonths = omsaetningDailyThresholds.normalizeWorkingDaysByMonth(state && state.payload && state.payload.months);
            const months = Object.fromEntries(Object.entries(allMonths).filter(([key]) => key.startsWith(String(year) + '-')));
            return res.json({ ok: true, year, months, updatedAt: state && state.payload ? state.payload.updatedAt || null : null });
        } catch (err) {
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.post('/admin/working-days', express.json(), async (req, res) => {
        if (!requireSuperadmin(req, res)) return;
        try {
            const year = Number(req.body && req.body.year);
            if (!Number.isInteger(year) || year < 2000 || year > 2200) {
                return res.status(400).json({ ok: false, error: 'Ugyldigt år' });
            }
            const submittedMonths = req.body && req.body.months && typeof req.body.months === 'object' ? req.body.months : {};
            const yearPrefix = String(year) + '-';
            const expectedKeys = Array.from({ length: 12 }, (_, index) => yearPrefix + String(index + 1).padStart(2, '0'));
            const normalizedSubmitted = omsaetningDailyThresholds.normalizeWorkingDaysByMonth(submittedMonths);
            const normalizedYear = Object.fromEntries(Object.entries(normalizedSubmitted).filter(([key]) => key.startsWith(yearPrefix)));
            if (expectedKeys.some(key => !Object.prototype.hasOwnProperty.call(normalizedYear, key))) {
                return res.status(400).json({ ok: false, error: 'Alle 12 måneder skal have 0-31 arbejdsdage' });
            }

            const state = await gohData.getAppState('omsaetning_working_days');
            const months = omsaetningDailyThresholds.normalizeWorkingDaysByMonth(state && state.payload && state.payload.months);
            for (const key of Object.keys(months)) {
                if (key.startsWith(yearPrefix)) delete months[key];
            }
            Object.assign(months, normalizedYear);
            const payload = { months, updatedAt: new Date().toISOString() };
            const saved = await gohData.setAppState('omsaetning_working_days', payload);
            if (!saved) return res.status(503).json({ ok: false, error: 'Arbejdsdage kunne ikke gemmes' });
            return res.json({ ok: true, year, months: normalizedYear, updatedAt: payload.updatedAt });
        } catch (err) {
            return res.status(500).json({ ok: false, error: err.message });
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
                styklisteFallbackCost: Number(marginInfo.styklisteFallbackCost || 0),
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

    router.get('/omsaetning/month-detail', async (req, res) => {
        try {
            const month = String(req.query.month || '').trim();
            const accountCsv = String(req.query.accounts || '').trim();
            const customerCsv = String(req.query.customers || '').trim();
            const detail = await omsaetningService.getMonthDetail({ month, accountCsv, customerCsv });
            const weekKeys = Array.isArray(detail.weekKeys) ? detail.weekKeys : [];
            let weeklyRows = [];

            if (weekKeys.length > 0) {
                const summary = await ordreindgangService.getSummary({
                    fraWeek: weekKeys[0],
                    tilWeek: weekKeys[weekKeys.length - 1]
                });
                const allowedWeeks = new Set(weekKeys);
                weeklyRows = (Array.isArray(summary.weeklyRows) ? summary.weeklyRows : [])
                    .filter(row => allowedWeeks.has(String(row.weekKey || '')));
            }

            const totalOrdK = weeklyRows.reduce((sum, row) => sum + Number(row.totalOrd || 0), 0);
            const totalTilbudK = weeklyRows.reduce((sum, row) => sum + Number(row.totalTilbud || 0), 0);
            return res.json({
                ok: true,
                ...detail,
                ordreindgang: {
                    unit: 'thousand-dkk',
                    totalOrdK,
                    totalTilbudK,
                    weeklyRows
                }
            });
        } catch (err) {
            if (err && err.statusCode) {
                return res.status(err.statusCode).json({ ok: false, error: err.message || 'Ugyldig forespørgsel' });
            }
            logEvent('ERROR omsaetning/month-detail: ' + err.message);
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

    router.get('/omsaetning/customer-threshold/:custno', async (req, res) => {
        const custNo = String(req.params.custno || '').trim();
        if (!/^\d{1,20}$/.test(custNo)) {
            return res.status(400).json({ ok: false, error: 'Ugyldigt kundenummer' });
        }

        try {
            const threshold = await omsaetningThresholdsService.getThreshold(custNo);
            const meta = omsaetningThresholdsService.getStorageMeta();
            if (!threshold) {
                return res.json({
                    ok: true,
                    custNo,
                    warnThreshold: meta.defaultWarnThreshold,
                    goodThreshold: meta.defaultGoodThreshold,
                    useDailyBudget: meta.defaultUseDailyBudget,
                    dailyBreakEvenDkk: meta.defaultDailyBreakEvenDkk,
                    dailyBudgetDkk: meta.defaultDailyBudgetDkk,
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
                useDailyBudget: threshold.useDailyBudget,
                dailyBreakEvenDkk: threshold.dailyBreakEvenDkk,
                dailyBudgetDkk: threshold.dailyBudgetDkk,
                updatedAt: threshold.updatedAt,
                exists: true,
                storageFile: meta.filePath
            });
        } catch (err) {
            return res.status(500).json({ ok: false, error: err.message || 'Tærskler kunne ikke hentes' });
        }
    });

    router.post('/omsaetning/customer-threshold/:custno', express.json(), async (req, res) => {
        const custNo = String(req.params.custno || '').trim();
        if (!/^\d{1,20}$/.test(custNo)) {
            return res.status(400).json({ ok: false, error: 'Ugyldigt kundenummer' });
        }

        try {
            const { warnThreshold, goodThreshold, useDailyBudget, dailyBreakEvenDkk, dailyBudgetDkk } = req.body || {};
            const saved = await omsaetningThresholdsService.setThreshold(custNo, {
                warnThreshold,
                goodThreshold,
                useDailyBudget,
                dailyBreakEvenDkk,
                dailyBudgetDkk
            });
            const meta = omsaetningThresholdsService.getStorageMeta();
            if (!saved) {
                return res.status(400).json({ ok: false, error: 'Ugyldige tærskelværdier' });
            }

            return res.json({
                ok: true,
                custNo,
                warnThreshold: saved.warnThreshold,
                goodThreshold: saved.goodThreshold,
                useDailyBudget: saved.useDailyBudget,
                dailyBreakEvenDkk: saved.dailyBreakEvenDkk,
                dailyBudgetDkk: saved.dailyBudgetDkk,
                updatedAt: saved.updatedAt,
                storageFile: meta.filePath
            });
        } catch (err) {
            const status = err && err.code === 'GOH_PERSIST_FAILED' ? 503 : 500;
            return res.status(status).json({ ok: false, error: err.message || 'Tærskler kunne ikke gemmes' });
        }
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

    router.post('/settings/rest-prices', express.json(), (req, res) => {
        try {
            if (!requireSuperadmin(req, res)) return;
            const restPrices = settingsService.updateRestPrices(req.body && req.body.restPrices);
            diskCache.del('lagerliste_v8');
            return res.json({ ok: true, restPrices });
        } catch (err) {
            return res.status(400).json({ ok: false, error: err.message });
        }
    });

    router.post('/settings/db-profiles/active', express.json(), requireAuthenticated, (req, res) => {
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

    router.post('/settings/db-profiles/upsert', express.json(), requireAuthenticated, (req, res) => {
        try {
            const profile = settingsService.upsertProfile(req.body || {});
            logEvent('SETTINGS: profile upserted: ' + profile.id);
            res.json({ ok: true, profile, ...settingsService.getSettingsSummary() });
        } catch (err) {
            res.status(400).json({ ok: false, error: err.message });
        }
    });

    router.delete('/settings/db-profiles/:id', requireAuthenticated, (req, res) => {
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

                    // forceRefresh: salta la cache GOH e ricalcola da Visma; il risultato
                    // viene riscritto sia in locale sia su GOH (mirror) per tutte le macchine.
                    const aftercalc = await getOrComputeAftercalc(ordNo, { priority: 'high', forceRefresh: true });
                    if (aftercalc && !aftercalc.error) {
                        const marginInfo = {
                            ordNo,
                            totalRevenue: Number(aftercalc.summary && aftercalc.summary.totalRevenue || 0),
                            totalCost: Number(aftercalc.summary && aftercalc.summary.totalCost || 0),
                            styklisteFallbackCost: Number(aftercalc.summary && aftercalc.summary.styklisteFallbackCost || 0),
                            computedAt: Date.now()
                        };
                        orderMarginCache.set(ordNo, marginInfo);
                        const marginResult = {
                            ordNo,
                            totalRevenue: marginInfo.totalRevenue,
                            totalCost: marginInfo.totalCost,
                            styklisteFallbackCost: marginInfo.styklisteFallbackCost,
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

    router.post('/open-drawing', requireAuthenticated, async (req, res) => {
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

            const opened = await openPdfTarget(candidatePath, {
                spawn,
                openPath: global.__desktopOpenPath,
                openExternal: global.__desktopOpenExternal
            });
            logEvent('OPEN-DRAWING: ' + opened.value + (prodNo ? (' [prodNo=' + prodNo + ']') : ''));
            return res.json({ ok: true });
        } catch (err) {
            logEvent('OPEN-DRAWING ERROR: ' + err.message);
            return res.status(err.statusCode || 500).json({ ok: false, message: err.message });
        }
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
                            styklisteFallbackCost: Number(cachedMargin.styklisteFallbackCost || 0),
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
                            styklisteFallbackCost: Number(staleMargin.styklisteFallbackCost || 0),
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
                    TotalCost: marginInfo ? marginInfo.totalCost : null,
                    StyklisteFallbackCost: marginInfo ? Number(marginInfo.styklisteFallbackCost || 0) : null
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

    // ── AFTERCALC COST EXCLUSIONS ────────────────────────────────────────────
    router.get('/aftercalc-cost-exclusions/:ordno', requireAuthenticated, async (req, res) => {
        try {
            const ordNo = Number(req.params.ordno);
            const result = await aftercalcCostExclusionsService.listForOrder(ordNo);
            res.json({ ok: true, ordNo, ...result });
        } catch (error) {
            res.status(error.statusCode || 500).json({ ok: false, error: error.message });
        }
    });

    router.post('/aftercalc-cost-exclusions/:ordno/:lineno', express.json(), requireAuthenticated, async (req, res) => {
        try {
            const ordNo = Number(req.params.ordno);
            const lineNo = Number(req.params.lineno);
            if (!req.body || typeof req.body.excluded !== 'boolean') {
                return res.status(400).json({ ok: false, error: 'Feltet excluded skal være true eller false' });
            }
            const orderData = await getOrComputeAftercalc(ordNo, { priority: 'high' });
            if (!orderData || orderData.error) {
                return res.status(404).json({ ok: false, error: (orderData && orderData.error) || 'Ordren blev ikke fundet' });
            }
            const lineExists = Array.isArray(orderData.salesOrderLines)
                && orderData.salesOrderLines.some(line => Number(line && line.LnNo) === lineNo);
            if (!lineExists) {
                return res.status(404).json({ ok: false, error: 'Salgslinjen blev ikke fundet på ordren' });
            }

            const user = getSessionUser(req) || {};
            const result = await aftercalcCostExclusionsService.setLine(
                ordNo,
                lineNo,
                req.body.excluded,
                user.displayName || user.username || ''
            );
            orderMarginCache.delete(ordNo);
            orderMarginInFlight.delete(ordNo);
            diskCache.del(ORDER_MARGIN_CACHE_KEY_PREFIX + ordNo);
            res.json(result);
        } catch (error) {
            res.status(error.statusCode || 500).json({ ok: false, error: error.message });
        }
    });

    // ── Personalehåndbog search API ─────────────────────────────────────────
    router.get('/ph/status', (_req, res) => {
        res.json(phCrawler.getPhStatus());
    });

    router.get('/ph/search', (req, res) => {
        res.json(phCrawler.searchPh(req.query.q));
    });

    router.post('/ph/reindex', (_req, res) => {
        if (phCrawler.isPhIndexing()) return res.json({ ok: false, msg: 'Allerede i gang' });
        phCrawler.crawlPH().catch(e => phCrawler.markPhError(e.message));
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

    router.post('/bom/create-products/execute', express.json(), requireAuthenticated, async (req, res) => {
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

    router.put('/qms/dataset', requireAuthenticated, (req, res) => {
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

    router.get('/lagerliste/current', requireModulePermission('lagerliste'), async (req, res) => {
        try {
            const forceRefresh = String(req.query && req.query.force || '') === '1';
            const forceAftercalc = String(req.query && req.query.aftercalc || '') === '1';
            return res.json({ ok: true, ...(await lagerlisteService.getCurrent({ forceRefresh, forceAftercalc })) });
        } catch (err) {
            logEvent('ERROR lagerliste/current: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Lagerliste fejl' });
        }
    });

    router.get('/lagerliste2/routes/current', requireModulePermission('lagerliste'), async (req, res) => {
        try {
            const forceRefresh = String(req.query && req.query.force || '') === '1';
            return res.json({ ok: true, ...(await lagerliste2Service.getCurrentRoutes({ forceRefresh })) });
        } catch (err) {
            logEvent('ERROR lagerliste2/routes/current: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Lagerliste 2 rutefejl' });
        }
    });

    router.get('/lagerliste2/reservations/current', requireModulePermission('lagerliste'), async (req, res) => {
        try {
            const forceRefresh = String(req.query && req.query.force || '') === '1';
            return res.json({ ok: true, ...(await lagerliste2Service.getCurrentReservations({ forceRefresh })) });
        } catch (err) {
            logEvent('ERROR lagerliste2/reservations/current: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Lagerliste 2 reservationsfejl' });
        }
    });

    router.post('/lagerliste2/movement-evidence', express.json(), requireModulePermission('lagerliste'), async (req, res) => {
        try {
            return res.json({ ok: true, ...(await lagerliste2Service.getMovementEvidence(req.body || {})) });
        } catch (err) {
            logEvent('ERROR lagerliste2/movement-evidence: ' + err.message);
            return res.status(err.statusCode || 500).json({ ok: false, error: err.message || 'Lagerliste 2 bevægelsesforklaring fejlede' });
        }
    });

    router.get('/lagerliste/vareopslag/:prodno', requireModulePermission('lagerliste'), async (req, res) => {
        try {
            const result = await lagerlisteService.lookupProduct(req.params.prodno);
            if (!result) return res.status(404).json({ ok: false, error: 'Varenummer ikke fundet: ' + String(req.params.prodno || '') });
            return res.json({ ok: true, ...result });
        } catch (err) {
            logEvent('ERROR lagerliste/vareopslag: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Vareopslag fejl' });
        }
    });

    // TEMP debug (read-only): discover how reservations (ShpRsv) link to orders
    router.get('/lagerliste/reservations-debug', (req, res, next) => {
        if (!requireSuperadmin(req, res)) return;
        return next();
    }, async (req, res) => {
        try {
            const pool = await getConnection();
            const out = {};

            // Iterative mode: ?table=X [&prodno=Y] — table name validated against INFORMATION_SCHEMA
            const reqTable = String(req.query && req.query.table || '').trim();
            if (reqTable) {
                const tblCheck = pool.request();
                tblCheck.input('tbl', sql.NVarChar, reqTable);
                const found = await tblCheck.query(`
                    SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @tbl`);
                if (!found.recordset.length) {
                    return res.status(404).json({ ok: false, error: 'Tabel ikke fundet: ' + reqTable });
                }
                const safeName = found.recordset[0].TABLE_NAME;
                const colReq = pool.request();
                colReq.input('tbl', sql.NVarChar, safeName);
                const cols = await colReq.query(`
                    SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_NAME = @tbl ORDER BY ORDINAL_POSITION`);
                out.table = safeName;
                out.columns = cols.recordset;
                const colNames = cols.recordset.map(c => c.COLUMN_NAME);
                const prodNo = String(req.query.prodno || '').trim();
                const ordNo = String(req.query.ordno || '').trim();
                const rowReq = pool.request();
                let where = '';
                if (prodNo && colNames.includes('ProdNo')) {
                    rowReq.input('prodNo', sql.NVarChar, prodNo);
                    where = ' WHERE ProdNo = @prodNo';
                } else if (ordNo && colNames.includes('OrdNo') && /^\d+$/.test(ordNo)) {
                    rowReq.input('ordNo', sql.Numeric, Number(ordNo));
                    where = ' WHERE OrdNo = @ordNo';
                }
                const rows = await rowReq.query(`SELECT TOP 300 * FROM [${safeName}] WITH(NOLOCK)${where}`);
                out.rows = rows.recordset;
                return res.json({ ok: true, ...out });
            }

            const tables = await pool.request().query(`
                SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_NAME LIKE '%Shp%' OR TABLE_NAME LIKE '%Rsv%' OR TABLE_NAME LIKE '%Parti%'
                ORDER BY TABLE_NAME`);
            out.tables = tables.recordset;

            const columns = await pool.request().query(`
                SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'ShpBal' ORDER BY ORDINAL_POSITION`);
            out.shpBalColumns = columns.recordset;

            const ordLnRsv = await pool.request().query(`
                SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_NAME = 'OrdLn' AND (COLUMN_NAME LIKE '%Rsv%' OR COLUMN_NAME LIKE '%Alloc%' OR COLUMN_NAME LIKE '%Pic%')
                ORDER BY ORDINAL_POSITION`);
            out.ordLnReservationColumns = ordLnRsv.recordset;

            const reserved = await pool.request().query(`
                SELECT TOP 8 B.ProdNo, B.Bal, B.StcInc, B.ShpRsv, B.PoPhStB
                FROM StcBal B WITH(NOLOCK)
                WHERE B.StcNo = 1 AND TRY_CONVERT(decimal(18,6), B.ShpRsv) > 0
                ORDER BY TRY_CONVERT(decimal(18,6), B.ShpRsv) DESC`);
            out.reservedProducts = reserved.recordset;

            out.shpBalRows = {};
            const prodNos = reserved.recordset.map(r => String(r.ProdNo).trim()).filter(Boolean).slice(0, 4);
            for (const prodNo of prodNos) {
                const request = pool.request();
                request.input('prodNo', sql.NVarChar, prodNo);
                const rows = await request.query(`SELECT TOP 10 * FROM ShpBal WITH(NOLOCK) WHERE ProdNo = @prodNo`);
                out.shpBalRows[prodNo] = rows.recordset;
            }

            return res.json({ ok: true, ...out });
        } catch (err) {
            logEvent('ERROR lagerliste/reservations-debug: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message });
        }
    });

    router.get('/lagerliste/snapshot-months', requireModulePermission('lagerliste'), (req, res) => {
        try {
            return res.json({ ok: true, months: lagerlisteService.listMonthlySnapshots(fs) });
        } catch (err) {
            logEvent('ERROR lagerliste/snapshot-months: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Snapshot måneder fejl' });
        }
    });

    router.get('/lagerliste/reconciliations/:periodKey', requireModulePermission('lagerliste'), (req, res) => {
        return res.json({ ok: true, rows: lagerlisteReconciliationService.list(req.params.periodKey) });
    });

    router.post('/lagerliste/reconciliations', (req, res, next) => {
        const user = requireSuperadmin(req, res);
        if (!user) return;
        req.reconciliationUser = user;
        return next();
    }, express.json(), (req, res) => {
        try {
            const row = lagerlisteReconciliationService.add({
                ...(req.body || {}),
                createdBy: req.reconciliationUser && req.reconciliationUser.username
            });
            return res.json({ ok: true, row });
        } catch (err) {
            return res.status(400).json({ ok: false, error: err.message });
        }
    });

    router.delete('/lagerliste/reconciliations/:id', (req, res, next) => {
        if (!requireSuperadmin(req, res)) return;
        return next();
    }, (req, res) => {
        const deleted = lagerlisteReconciliationService.remove(req.params.id);
        if (!deleted) return res.status(404).json({ ok: false, error: 'Manuel afstemning ikke fundet' });
        return res.json({ ok: true });
    });

    router.get('/lagerliste/snapshot/:month', requireModulePermission('lagerliste'), (req, res) => {
        try {
            const snapshot = lagerlisteService.loadMonthlySnapshot({ fs, month: req.params.month });
            if (!snapshot) return res.status(404).json({ ok: false, error: 'Snapshot ikke fundet' });
            return res.json({ ok: true, ...snapshot });
        } catch (err) {
            logEvent('ERROR lagerliste/snapshot: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Snapshot fejl' });
        }
    });

    router.post('/lagerliste/snapshot/:month', (req, res, next) => {
        if (!requireSuperadmin(req, res)) return;
        return next();
    }, async (req, res) => {
        try {
            const month = String(req.params.month || '').trim();
            if (!/^\d{4}-\d{2}$/.test(month)) {
                return res.status(400).json({ ok: false, error: 'Måned skal være YYYY-MM' });
            }
            const diverse = Array.isArray(req.body && req.body.diverse) ? req.body.diverse : [];
            const saved = await lagerlisteService.saveMonthlySnapshot({ fs, month, diverse });
            return res.json({
                ok: true,
                month: saved.month,
                createdAt: saved.createdAt
            });
        } catch (err) {
            logEvent('ERROR lagerliste/snapshot create: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Snapshot fejl' });
        }
    });

    router.get('/lagerliste/snapshots', requireModulePermission('lagerliste'), (req, res) => {
        try {
            const rows = lagerlisteService.listPointInTimeSnapshots(fs);
            return res.json({ ok: true, rows });
        } catch (err) {
            logEvent('ERROR lagerliste/snapshots list: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Snapshot liste fejl' });
        }
    });

    router.delete('/lagerliste/snapshots/:id', (req, res, next) => {
        if (!requireSuperadmin(req, res)) return;
        return next();
    }, (req, res) => {
        try {
            const snapshotId = String(req.params.id || '').trim();
            const deleted = lagerlisteService.deletePointInTimeSnapshot(fs, snapshotId);
            if (!deleted) return res.status(404).json({ ok: false, error: 'Snapshot ikke fundet' });
            logEvent('Lagerliste dags-snapshot slettet: ' + snapshotId);
            return res.json({ ok: true, snapshotId });
        } catch (err) {
            logEvent('ERROR lagerliste/snapshot delete: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Snapshot kunne ikke slettes' });
        }
    });

    router.get('/lagerliste/snapshots/:id', requireModulePermission('lagerliste'), (req, res) => {
        try {
            const snapshot = lagerlisteService.loadPointInTimeSnapshot(fs, req.params.id);
            if (!snapshot) return res.status(404).json({ ok: false, error: 'Snapshot ikke fundet' });
            return res.json({ ok: true, snapshot });
        } catch (err) {
            logEvent('ERROR lagerliste/snapshots load: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Snapshot hentning fejl' });
        }
    });

    router.post('/lagerliste/snapshots', (req, res, next) => {
        if (!requireSuperadmin(req, res)) return;
        return next();
    }, async (req, res) => {
        try {
            const capturedAtRaw = req.body && req.body.capturedAt;
            const capturedAt = capturedAtRaw ? new Date(capturedAtRaw) : new Date();
            if (Number.isNaN(capturedAt.getTime())) {
                return res.status(400).json({ ok: false, error: 'capturedAt er ugyldig' });
            }
            const note = String(req.body && req.body.note || '').trim();
            const diverse = Array.isArray(req.body && req.body.diverse) ? req.body.diverse : [];
            const currentOverride = req.body && req.body.current && typeof req.body.current === 'object'
                ? req.body.current
                : null;
            const forceRefresh = String(req.body && req.body.forceRefresh || '') === '1';
            const pad = value => String(value).padStart(2, '0');
            const snapshotId = capturedAt.getFullYear() + '-' + pad(capturedAt.getMonth() + 1) + '-' + pad(capturedAt.getDate())
                + '_' + pad(capturedAt.getHours()) + '-' + pad(capturedAt.getMinutes()) + '-' + pad(capturedAt.getSeconds());
            setImmediate(() => {
                lagerlisteService.savePointInTimeSnapshot({ fs, capturedAt, note, diverse, currentOverride, forceRefresh })
                    .catch(err => logEvent('ERROR lagerliste/snapshots background save: ' + err.message));
            });
            return res.json({
                ok: true,
                snapshotId,
                kind: 'point-in-time',
                capturedAt: capturedAt.toISOString(),
                createdAt: new Date().toISOString(),
                note
            });
        } catch (err) {
            logEvent('ERROR lagerliste/snapshots create: ' + err.message);
            return res.status(500).json({ ok: false, error: err.message || 'Snapshot oprettelse fejl' });
        }
    });
    // ────────────────────────────────────────────────────────────────────────

    return router;
}

module.exports = {
    createApiRouter
};
