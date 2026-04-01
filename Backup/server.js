const express = require('express');
const sql = require('mssql/msnodesqlv8');
const fs = require('fs');
const path = require('path');
const getConnection = require('./db');

const app = express();
const APP_VERSION = 'Gantech AS Beta - v0.3.2 (2026-03-19)';
const LOG_FILE = path.join(__dirname, 'gantech.log');
const ORDER_LIST_CACHE_TTL_MS = 10 * 60 * 1000;
const ORDER_LIST_MAX_ROWS = 150;
const ORDER_LIST_DAYS_BACK = 30;
const STARTUP_MARGIN_WARM_COUNT = 60;
const MAX_DB_CALC_CONCURRENCY = 2;
const SESSION_COOKIE_NAME = 'gantech_sid';
const SESSION_TTL_MS = 1000 * 60 * 60 * 8;
const INITIAL_SUPERUSER = process.env.GANTECH_ADMIN_USER || 'admin';
const INITIAL_SUPERPASS = process.env.GANTECH_ADMIN_PASS || 'admin123';

const orderListCache = {
    data: [],
    loadedAt: 0,
    loading: false,
    refreshPromise: null,
    lastError: null
};

const orderMarginCache = new Map();
const orderMarginInFlight = new Map();
const dbCalcQueue = [];
let activeDbCalcs = 0;

// Log di sistema
function logEvent(message) {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] ${message}\n`;
    fs.appendFileSync(LOG_FILE, logMessage);
    console.log(logMessage.trim());
}

logEvent('=== SERVER STARTED - ' + APP_VERSION + ' ===');

// Evita cache lato browser durante lo sviluppo: forza sempre il fetch dell'ultima UI/API.
app.use((req, res, next) => {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
    next();
});

// Funzione principale per calcolare i costi
async function getAfterCalc(ordNo) {
    const pool = await getConnection();
    
    try {
        // 1. Carica l'intestazione dell'ordine di vendita con il nome cliente
        const orderResult = await pool.request()
            .input('ordNo', sql.Numeric, ordNo)
            .query(`
                SELECT O.OrdNo, O.TrTp, O.InvoSF, A.Nm as CustomerName
                FROM Ord O
                LEFT JOIN Actor A ON O.CustNo = A.CustNo
                WHERE O.OrdNo = @ordNo
            `);
        
        if (orderResult.recordset.length === 0) {
            return { error: 'Ordine non trovato' };
        }
        
        const orderHeader = orderResult.recordset[0];

        // 1b. Carica TUTTE le linee dell'ordine di vendita (per visualizzazione e navigazione a cascata)
        const salesOrderLinesResult = await pool.request()
            .input('ordNo', sql.Numeric, ordNo)
            .query(`
                SELECT
                    OrdNo,
                    LnNo,
                    ProdNo,
                    Descr,
                    DPrice,
                    NoFin,
                    CCstPr,
                    PurcNo,
                    CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost
                FROM OrdLn
                WHERE OrdNo = @ordNo
                ORDER BY LnNo
            `);
        
        // 2. Carica le linee di VENDITA (senza PurcNo)
        const salesLinesResult = await pool.request()
            .input('ordNo', sql.Numeric, ordNo)
            .query(`
                SELECT 
                    OrdNo, 
                    LnNo, 
                    ProdNo, 
                    Descr, 
                    DPrice,
                    NoFin,
                    CCstPr,
                    CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost
                FROM OrdLn
                WHERE OrdNo = @ordNo AND PurcNo IS NULL
                ORDER BY LnNo
            `);
        
        const salesLines = salesLinesResult.recordset.map(line => {
            const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
            const isDiscountLine = lineSalesPrice === 0;
            const effectiveLineCost = isDiscountLine ? 0 : (line.LineCost || 0);
            return {
                ...line,
                IsDiscountLine: isDiscountLine,
                EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2))
            };
        });
        const salesLinesTotalCost = salesLines.reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);

        // 2b. Costo linee vendita senza P.O. (PurcNo nullo o 0) per il riepilogo finale
        const salesOrderLines = salesOrderLinesResult.recordset.map(line => {
            const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
            const isDiscountLine = lineSalesPrice === 0;
            const effectiveLineCost = isDiscountLine ? 0 : (line.LineCost || 0);
            return {
                ...line,
                IsDiscountLine: isDiscountLine,
                EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2))
            };
        });

        function calculateProductionOrderTotal(lines) {
            let total = 0;
            for (const line of lines) {
                const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                if (line.LnNo === 1 || key === '0' || key === '3' || key === '5') continue;

                if (key === '2') {
                    total += (line.NestingCost || 0) * (line.NoFin || 0);
                } else {
                    total += (line.LineCost || 0);
                }
            }
            return parseFloat(Number(total).toFixed(2));
        }
        
        // 3. Carica i numeri di produzione associati (quelli con PurcNo)
        const productionLinesResult = await pool.request()
            .input('ordNo', sql.Numeric, ordNo)
            .query(`
                SELECT DISTINCT PurcNo FROM OrdLn
                WHERE OrdNo = @ordNo AND PurcNo IS NOT NULL
            `);
        
        const productionOrders = [];
        for (const prodLine of productionLinesResult.recordset) {
            const purcNo = prodLine.PurcNo;
            
            // Carica l'intestazione dell'ordine di produzione
            const prodOrderResult = await pool.request()
                .input('purcNo', sql.Numeric, purcNo)
                .query(`
                    SELECT O.OrdNo, A.Nm, O.TrTp, O.InvoSF FROM Ord O JOIN Actor A ON O.CustNo=A.CustNo WHERE OrdNo = @purcNo
                `);
            
            if (prodOrderResult.recordset.length > 0) {
                const prodOrder = prodOrderResult.recordset[0];

                // Carica le linee dell'ordine di produzione
                const prodLinesResult = await pool.request()
                    .input('purcNo', sql.Numeric, purcNo)
                    .query(`
                        SELECT 
                            OrdNo, 
                            LnNo, 
                            ProdNo, 
                            Descr, 
                            DPrice,
                            NoFin,
                            CCstPr,
                            PurcNo,
                            ProdTp4,
                            CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost,
                            (SELECT AVG(n.CstPr) FROM OrdLn n WHERE n.TrInf2 = CAST(@purcNo AS VARCHAR(20)) AND n.ProdNo = OrdLn.ProdNo) AS NestingCost
                        FROM OrdLn
                        WHERE OrdNo = @purcNo
                        ORDER BY LnNo
                    `);

                const prodLines = prodLinesResult.recordset;
                const prodTotalCost = calculateProductionOrderTotal(prodLines);

                productionOrders.push({
                    ordNo: purcNo,
                    trTp: prodOrder.TrTp,
                    revenue: prodOrder.InvoSF || 0,
                    lines: prodLines,
                    totalCost: prodTotalCost
                });
            }
        }

        const productionTotalByOrdNo = new Map(
            productionOrders.map(po => [Number(po.ordNo), Number(po.totalCost || 0)])
        );

        const salesOrderLinesWithProductionTotal = salesOrderLines.map(line => {
            const purcNo = line.PurcNo ? Number(line.PurcNo) : 0;
            const productionTotal = purcNo ? productionTotalByOrdNo.get(purcNo) : undefined;

            if (productionTotal !== undefined && !line.IsDiscountLine) {
                const roundedTotal = parseFloat(Number(productionTotal).toFixed(2));
                return {
                    ...line,
                    ProductionOrderTotalCost: roundedTotal,
                    EffectiveLineCost: roundedTotal
                };
            }

            return {
                ...line,
                ProductionOrderTotalCost: productionTotal !== undefined ? parseFloat(Number(productionTotal).toFixed(2)) : null
            };
        });

        const salesNoPOLines = salesOrderLinesWithProductionTotal.filter(line => !line.PurcNo || line.PurcNo === 0);
        const salesNoPOTotalCost = salesNoPOLines.reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);

        const includedProductionOrdNos = new Set(
            salesOrderLinesWithProductionTotal
                .filter(line => line.PurcNo && line.PurcNo !== 0 && !line.IsDiscountLine)
                .map(line => Number(line.PurcNo))
        );
        
        // 4. Calcola i totali generali
        const productionTotalCost = productionOrders.reduce((sum, ord) => {
            if (!includedProductionOrdNos.has(Number(ord.ordNo))) return sum;
            return sum + (ord.totalCost || 0);
        }, 0);
        const totalCost = salesNoPOTotalCost + productionTotalCost;
        const totalRevenue = orderHeader.InvoSF || 0;
        const margin = totalRevenue - totalCost;
        const marginPercentage = totalCost > 0 ? ((totalRevenue / totalCost) * 100).toFixed(2) : 0;
        
        return {
            orderHeader,
            salesOrderLines: salesOrderLinesWithProductionTotal,
            salesLines,
            salesLinesTotalCost: parseFloat(salesLinesTotalCost.toFixed(2)),
            productionOrders,
            summary: {
                totalRevenue: parseFloat(totalRevenue.toFixed(2)),
                totalCost: parseFloat(totalCost.toFixed(2)),
                margin: parseFloat(margin.toFixed(2)),
                marginPercentage
            }
        };
    } catch (err) {
        console.error('Errore in getAfterCalc:', err);
        throw err;
    }
}

function isOrderListCacheFresh() {
    return orderListCache.loadedAt > 0 && (Date.now() - orderListCache.loadedAt) < ORDER_LIST_CACHE_TTL_MS;
}

function runWithDbCalcLimit(task) {
    return new Promise((resolve, reject) => {
        dbCalcQueue.push({ task, resolve, reject });
        pumpDbCalcQueue();
    });
}

function pumpDbCalcQueue() {
    while (activeDbCalcs < MAX_DB_CALC_CONCURRENCY && dbCalcQueue.length > 0) {
        const job = dbCalcQueue.shift();
        activeDbCalcs += 1;
        Promise.resolve()
            .then(job.task)
            .then(job.resolve)
            .catch(job.reject)
            .finally(() => {
                activeDbCalcs -= 1;
                pumpDbCalcQueue();
            });
    }
}

async function fetchOrderListBase() {
    const pool = await getConnection();
    const result = await pool.request().query(`
            WITH BaseOrders AS (
                SELECT TOP ${ORDER_LIST_MAX_ROWS}
                    O.SelBuy,
                    O.OrdNo,
                    C.Nm AS CustomerName,
                    SU.Usr AS SellerUsr,
                    O.LstInvDt,
                    O.InvoSF
                FROM Ord O
                LEFT JOIN Actor C ON C.CustNo = O.CustNo
                OUTER APPLY (
                    SELECT TOP 1 A.Usr
                    FROM Actor A
                    WHERE LTRIM(RTRIM(CONVERT(VARCHAR(50), A.EmpNo))) = LTRIM(RTRIM(CONVERT(VARCHAR(50), O.SelBuy)))
                ) SU
                WHERE O.InvoNo IS NOT NULL AND O.InvoNo <> ''
                  AND O.InvoSF > 0
                                    AND O.LstInvDt >= CAST(CONVERT(VARCHAR(8), DATEADD(day, -${ORDER_LIST_DAYS_BACK}, GETDATE()), 112) AS INT)
                ORDER BY O.LstInvDt DESC
            )
            SELECT
                B.SelBuy,
                B.OrdNo,
                B.CustomerName,
                B.SellerUsr,
                B.LstInvDt,
                B.InvoSF
            FROM BaseOrders B
            ORDER BY B.LstInvDt DESC, B.OrdNo DESC
        `);

    return result.recordset;
}

async function getOrComputeOrderMargin(ordNo, options = {}) {
    const forceRefresh = Boolean(options.forceRefresh);
    const key = Number(ordNo);
    if (!Number.isFinite(key)) {
        throw new Error('Ordrenummer ugyldigt');
    }

    if (!forceRefresh && orderMarginCache.has(key)) {
        return orderMarginCache.get(key);
    }

    if (!forceRefresh && orderMarginInFlight.has(key)) {
        return orderMarginInFlight.get(key);
    }

    const computePromise = runWithDbCalcLimit(async () => {
        const data = await getAfterCalc(key);
        if (data.error) {
            throw new Error(data.error);
        }

        const marginInfo = {
            ordNo: key,
            totalRevenue: Number(data.summary.totalRevenue || 0),
            totalCost: Number(data.summary.totalCost || 0),
            computedAt: Date.now()
        };

        orderMarginCache.set(key, marginInfo);
        return marginInfo;
    }).finally(() => {
        orderMarginInFlight.delete(key);
    });

    orderMarginInFlight.set(key, computePromise);
    return computePromise;
}

function warmMarginsInBackground(ordNos) {
    for (const ordNo of ordNos) {
        const numericOrdNo = Number(ordNo);
        if (!Number.isFinite(numericOrdNo)) continue;
        getOrComputeOrderMargin(numericOrdNo).catch(() => {});
    }
}

async function refreshOrderListCache(force = false) {
    if (!force && isOrderListCacheFresh()) {
        return;
    }

    if (orderListCache.loading) {
        if (orderListCache.refreshPromise) {
            await orderListCache.refreshPromise;
        }
        return;
    }

    orderListCache.loading = true;
    orderListCache.refreshPromise = (async () => {
        try {
            const rows = await fetchOrderListBase();
            orderListCache.data = rows;
            orderListCache.loadedAt = Date.now();
            orderListCache.lastError = null;

            const warmOrdNos = rows.slice(0, STARTUP_MARGIN_WARM_COUNT).map(r => r.OrdNo);
            warmMarginsInBackground(warmOrdNos);
        } catch (err) {
            orderListCache.lastError = err.message;
            throw err;
        } finally {
            orderListCache.loading = false;
            orderListCache.refreshPromise = null;
        }
    })();

    await orderListCache.refreshPromise;
}

// Endpoint API
app.get('/aftercalc/:ordno', async (req, res) => {
    try {
        const ordNo = parseInt(req.params.ordno);
        logEvent('SEARCH: OrdNo=' + ordNo);
        const data = await getAfterCalc(ordNo);
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

// Endpoint API: margine singolo ordine (usato per caricamento progressivo lista)
app.get('/order-margin/:ordno', async (req, res) => {
    try {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) {
            return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
        }

        const marginInfo = await getOrComputeOrderMargin(ordNo);
        return res.json({
            ordNo: marginInfo.ordNo,
            totalRevenue: marginInfo.totalRevenue,
            totalCost: marginInfo.totalCost,
            cached: true
        });
    } catch (err) {
        logEvent('ERROR order-margin: ' + err.message);
        return res.status(500).json({ error: err.message });
    }
});

// Endpoint API: riepilogo ordine figlio (solo righe con NoFin > 0)
app.get('/production-summary/:ordno', async (req, res) => {
    const pool = await getConnection();
    try {
        const ordNo = parseInt(req.params.ordno);
        if (Number.isNaN(ordNo)) {
            return res.status(400).json({ error: 'Ordrenummer ugyldigt' });
        }

        const linesResult = await pool.request()
            .input('ordNo', sql.Numeric, ordNo)
            .query(`
                SELECT
                    LnNo,
                    ProdNo,
                    Descr,
                    NoFin,
                    DPrice,
                    CCstPr,
                    ProdTp4,
                    PurcNo,
                    CAST(NoFin * CCstPr AS DECIMAL(10,2)) AS LineCost,
                    (SELECT AVG(n.CstPr)
                     FROM OrdLn n
                     WHERE n.TrInf2 = CAST(@ordNo AS VARCHAR(20))
                       AND n.ProdNo = OrdLn.ProdNo) AS NestingCost
                FROM OrdLn
                WHERE OrdNo = @ordNo AND NoFin > 0
                ORDER BY LnNo
            `);

        const lines = linesResult.recordset.map(line => {
            const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
            const effectiveLineCost = key === '2'
                ? ((line.NestingCost || 0) * (line.NoFin || 0))
                : (line.LineCost || 0);

            return {
                ...line,
                EffectiveLineCost: parseFloat(Number(effectiveLineCost).toFixed(2))
            };
        });

        const totalCost = lines.reduce((sum, line) => sum + (line.EffectiveLineCost || 0), 0);

        return res.json({
            ordNo,
            lines,
            totalCost: parseFloat(Number(totalCost).toFixed(2))
        });
    } catch (err) {
        console.error('Errore production-summary:', err);
        return res.status(500).json({ error: err.message });
    }
});

// Endpoint nesting-detail: OrdLn-linjer brugt til nesting-kost for et produkt
app.get('/nesting-detail/:ordno/:prodno', async (req, res) => {
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

// Endpoint ProdTr: transaktioner for en produktionslinje
app.get('/prodtr/:ordno/:lnno', async (req, res) => {
    try {
        const ordNo = parseInt(req.params.ordno);
        const lnNo = parseInt(req.params.lnno);
        if (Number.isNaN(ordNo) || Number.isNaN(lnNo)) {
            return res.status(400).json({ error: 'Ugyldige parametre' });
        }
        const pool = await getConnection();
        const result = await pool.request()
            .input('ordNo', sql.Numeric, ordNo)
            .input('lnNo', sql.Numeric, lnNo)
            .query(`
                SELECT
                    P.FinDt,
                    P.FinTm,
                    P.NoInvoAb,
                    A.Nm AS HvemNm
                FROM ProdTr P
                LEFT JOIN Actor A ON A.EmpNo = P.EmpNo
                WHERE P.OrdNo = @ordNo AND P.OrdLnNo = @lnNo
                ORDER BY P.FinDt DESC, P.FinTm DESC
            `);
        res.json(result.recordset);
    } catch (err) {
        res.status(500).json({ error: err.message });
    }
});

// Endpoint lista ordini recenti
app.get('/order-list', async (req, res) => {
    try {
        const forceRefresh = req.query.force === '1';
        logEvent('ORDER-LIST: force=' + (forceRefresh ? '1' : '0'));

        if (forceRefresh) {
            await refreshOrderListCache(true);
        } else if (!isOrderListCacheFresh()) {
            await refreshOrderListCache();
        }

        const data = orderListCache.data.map(row => {
            const marginInfo = orderMarginCache.get(Number(row.OrdNo));
            return {
                ...row,
                TotalCost: marginInfo ? marginInfo.totalCost : null
            };
        });

        // Se la cache e' vecchia ma presente, rispondi subito e aggiorna in background.
        if (!isOrderListCacheFresh() && !orderListCache.loading) {
            refreshOrderListCache(true).catch(err => {
                logEvent('ERROR order-list refresh: ' + err.message);
            });
        }

        logEvent('ORDER-LIST: returned ' + data.length + ' rows');
        res.json(data);
    } catch (err) {
        logEvent('ERROR order-list: ' + err.message);
        res.status(500).json({ error: err.message });
    }
});

// Endpoint per HTML
app.get('/', (req, res) => {
    res.send(`
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Efterkalkulation</title>
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { font-family: Arial, sans-serif; background: #f5f5f5; padding: 20px; }
            .container { max-width: 1200px; margin: 0 auto; }
            .header-banner-wrapper { background: #c0392b; color: #fff; font-weight: 800; font-size: 25px; padding: 10px 12px; border-radius: 6px; margin-bottom: 20px; letter-spacing: 0.2px; width: 100%; position: sticky; top: 0; z-index: 1200; display: flex; align-items: center; justify-content: space-between; gap: 12px; }
            .header-status-badge { display: inline-block; font-size: 12px; font-weight: 700; color: #8a6d3b; background: #fff3cd; border: 1px solid #fff3cd; border-radius: 999px; padding: 4px 10px; white-space: nowrap; }
            .search-box { background: #fff; padding: 20px; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); position: sticky; top: 58px; z-index: 1100; display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
            .search-box.collapsed { padding: 8px 12px; height: 36px; }
            .search-box.collapsed > * { display: none; }
            .search-box.collapsed > #collapseToggleBtn { display: inline-block; }
            #collapseToggleBtn { background: #1976d2; color: #fff; border: none; padding: 8px 12px; border-radius: 3px; cursor: pointer; font-weight: 600; font-size: 12px; }
            .build-badge { display: inline-block; font-size: 12px; color: #444; background: #f1f1f1; border: 1px solid #ddd; border-radius: 4px; padding: 4px 8px; }
            .build-banner { display: none; }
            .search-box input { padding: 8px 12px; font-size: 14px; width: 200px; border: 1px solid #ddd; border-radius: 3px; }
            .search-box button { padding: 8px 16px; background: #4CAF50; color: white; border: none; border-radius: 3px; cursor: pointer; margin-left: 10px; }
            .mode-btn { background: #0d47a1 !important; }
            .list-toggle-btn { background: #455a64 !important; }
            .filter-input { width: 260px !important; margin-left: 10px; }
            .filter-select { width: 180px; padding: 8px 10px; border: 1px solid #ddd; border-radius: 3px; background: #fff; }
            .section { background: white; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); padding: 20px; }
            .order-header { background: linear-gradient(135deg, #1976D2 0%, #1565C0 100%); color: white; padding: 25px; border-radius: 6px; margin-bottom: 25px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
            .order-header h2 { margin: 0 0 20px 0; font-size: 28px; font-weight: 700; }
            .order-header-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; }
            .order-header-item { display: flex; flex-direction: column; }
            .order-header-label { font-size: 12px; font-weight: 600; opacity: 0.9; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
            .order-header-value { font-size: 22px; font-weight: 700; color: #fff; }
            h3 { color: #333; margin-bottom: 15px; border-bottom: 2px solid #2196F3; padding-bottom: 10px; }
            table { width: 100%; border-collapse: collapse; margin-top: 10px; }
            th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
            th { background: #f0f0f0; font-weight: bold; }
            tr:hover { background: #fafafa; }
            .summary-row { font-weight: bold; background: #f9f9f9; }
            .summary-box { background: #e8f5e9; padding: 15px; border-radius: 4px; margin-top: 15px; }
            .summary-box div { margin: 8px 0; font-size: 14px; }
            .summary-box .total { font-size: 18px; color: #2196F3; font-weight: bold; }
            .margin-positive { color: green; }
            .margin-negative { color: red; }
            .error { color: red; padding: 20px; background: #ffebee; border-radius: 4px; }
            .loading { color: #666; padding: 20px; }
            .prod-link { color: #1976D2; text-decoration: underline; cursor: pointer; }
            .prod-link:hover { color: #0D47A1; }
            .po-highlight { box-shadow: 0 0 0 3px #90CAF9; }
            .prodtp4-group { border: 1px solid #e5e5e5; border-radius: 4px; margin-bottom: 10px; overflow: hidden; }
            .prodtp4-header { background: #f7f9fc; padding: 10px 12px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; font-weight: 600; }
            .prodtp4-header:hover { background: #eef4fb; }
            .prodtp4-label { color: #2b2b2b; }
            .prodtp4-subtotal { color: #1976D2; font-weight: 700; }
            .prodtp4-body { padding: 8px 12px 12px; }
            .po-total-row { margin-top: 10px; padding: 10px 12px; border-top: 1px solid #ddd; font-weight: 700; text-align: right; background: #fafafa; }
            .prodtp4-hint { color: #555; margin: 6px 0 10px; font-size: 13px; }
            .main-product-box { background: #eef6ff; border: 2px solid #90caf9; border-radius: 6px; padding: 10px 12px; margin: 8px 0 12px; }
            .main-product-box .value { font-size: 20px; font-weight: 800; color: #0d47a1; margin-top: 3px; }
            .inline-link { color: #1565c0; text-decoration: underline; cursor: pointer; }
            .inline-link:hover { color: #0d47a1; }
            .prod-no-link { color: #1565c0; text-decoration: underline; cursor: pointer; }
            .prod-no-link:hover { color: #0d47a1; }
            .modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.35); display: none; align-items: center; justify-content: center; z-index: 9999; }
            .modal-box { width: min(900px, 92vw); max-height: 75vh; overflow: auto; background: #fff; border-radius: 8px; box-shadow: 0 10px 30px rgba(0,0,0,0.2); padding: 16px; }
            .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
            .modal-close { border: none; background: #efefef; border-radius: 4px; padding: 6px 10px; cursor: pointer; }
            .modal-loading { color: #666; padding: 8px 0; }
            .order-list-section { background: #fff; padding: 16px 20px; margin-bottom: 20px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            .order-list-section h3 { color: #333; margin-bottom: 12px; border-bottom: 2px solid #2196F3; padding-bottom: 8px; }
            .order-list-table { width: 100%; border-collapse: collapse; font-size: 13px; }
            .order-list-table th { background: #1565C0; color: #fff; padding: 8px 10px; text-align: left; }
            .order-list-table td { padding: 8px 10px; border-bottom: 1px solid #e0e0e0; cursor: pointer; }
            .order-list-table tr:hover td { background: #e3f2fd; }
        </style>
    </head>
    <body>
        <div class="header-banner-wrapper">
            <span>🔷 ${APP_VERSION}</span>
            <span class="header-status-badge" id="systemStatusBadge">System loading...</span>
        </div>
        <div class="container">
            <div class="search-box" id="searchBox">
                <button id="collapseToggleBtn" onclick="toggleSearchBox()" style="display:none;">▼ Søg</button>
                <input type="number" id="orderInput" placeholder="Indtast ordrenummer..." />
                <button onclick="searchOrder()">Søg</button>
                <button id="refreshListBtn" class="list-toggle-btn" onclick="refreshOrderList()">Opdater liste</button>
                <button class="mode-btn" onclick="toggleMarginMode()">Skift margin-type</button>
                <button id="listToggleBtn" class="list-toggle-btn" onclick="toggleOrderList()">Skjul kundeliste</button>
                <select id="brugerFilterSelect" class="filter-select" onchange="setBrugerFilter()">
                    <option value="">Alle brugere</option>
                </select>
                <input type="text" id="customerFilterInput" class="filter-input" placeholder="Søg kunde i listen..." oninput="setOrderListFilter()" />
                <button id="collapseExpandBtn" class="list-toggle-btn" onclick="toggleSearchBox()" style="margin-left:auto;">▲ Luk</button>
            </div>
            <div id="orderList"></div>
            <div id="result"></div>
        </div>

        <div id="summaryModal" class="modal-overlay" onclick="closeSummaryModal(event)">
            <div class="modal-box" onclick="event.stopPropagation()">
                <div class="modal-header">
                    <h3 id="summaryModalTitle">Produktoversigt</h3>
                    <button class="modal-close" onclick="closeSummaryModal()">Luk</button>
                </div>
                <div id="summaryModalBody"></div>
            </div>
        </div>
        
        <script>
            function formatNumber(num) {
                const fixed = parseFloat(num).toFixed(2);
                const parts = fixed.split('.');
                const integerPart = parts[0];
                const decimalPart = parts[1];
                
                // Aggiungi punto come separatore migliaia da destra a sinistra
                let formatted = '';
                for (let i = integerPart.length - 1, count = 0; i >= 0; i--, count++) {
                    if (count > 0 && count % 3 === 0) {
                        formatted = '.' + formatted;
                    }
                    formatted = integerPart[i] + formatted;
                }
                
                return formatted + ',' + decimalPart;
            }

            let currentMarginMode = 'classic';
            let orderListData = [];
            let orderListVisible = true;
            let orderListFilter = '';
            let orderListBrugerFilter = '';
            let marginStateByOrdNo = {};
            let marginJobQueue = [];
            let marginWorkerActiveCount = 0;
            let orderListRerenderTimer = null;
            let orderListLoading = false;
            let orderListAutoRefreshTimer = null;
            const MARGIN_MAX_CONCURRENT = 2;
            const MARGIN_QUEUE_DELAY_MS = 120;
            const MARGIN_PREFETCH_ROWS = ${ORDER_LIST_MAX_ROWS};
            const ORDER_LIST_AUTO_REFRESH_MS = 2 * 60 * 1000;

            function setSystemStatus(text, bgColor, textColor) {
                const badge = document.getElementById('systemStatusBadge');
                if (!badge) return;
                badge.textContent = text;
                badge.style.background = bgColor;
                badge.style.color = textColor;
                badge.style.borderColor = bgColor;
            }

            function updateSystemStatusFromOrders(orders) {
                if (!orders || orders.length === 0) {
                    setSystemStatus('System Ready', '#e8f5e9', '#1b5e20');
                    return;
                }

                const visibleOrders = orders.slice(0, MARGIN_PREFETCH_ROWS);
                const total = visibleOrders.length;
                let completed = 0;

                for (const o of visibleOrders) {
                    const state = getMarginState(o.OrdNo);
                    if (state && (state.status === 'success' || state.status === 'error')) {
                        completed += 1;
                    }
                }

                if (completed >= total) {
                    setSystemStatus('System Ready', '#e8f5e9', '#1b5e20');
                    return;
                }

                setSystemStatus('System loading... ' + completed + '/' + total, '#fff3cd', '#8a6d3b');
            }

            function getMarginModeLabel() {
                return currentMarginMode === 'new'
                    ? 'Ny (Salg/Kost x 100)'
                    : 'Klassisk ((Salg-Kost)/Salg x 100)';
            }

            function calculateOrderMarginPercent(revenue, cost) {
                if (currentMarginMode === 'new') {
                    return cost > 0 ? ((revenue / cost) * 100) : 0;
                }
                return revenue > 0 ? (((revenue - cost) / revenue) * 100) : 0;
            }

            function calculateLineMarginPercent(salesPrice, lineCost) {
                if (currentMarginMode === 'new') {
                    return lineCost > 0 ? ((salesPrice / lineCost) * 100) : 0;
                }
                return salesPrice > 0 ? (((salesPrice - lineCost) / salesPrice) * 100) : 0;
            }

            function toggleMarginMode() {
                currentMarginMode = currentMarginMode === 'new' ? 'classic' : 'new';
                const ordNo = document.getElementById('orderInput').value;
                if (ordNo) searchOrder();
                renderOrderList();
            }

            function scheduleOrderListRerender() {
                if (orderListRerenderTimer) return;
                orderListRerenderTimer = setTimeout(() => {
                    orderListRerenderTimer = null;
                    renderOrderList();
                }, 120);
            }

            function getMarginState(ordNo) {
                return marginStateByOrdNo[String(ordNo)] || null;
            }

            function hydrateMarginStateFromOrderList(orders) {
                marginStateByOrdNo = {};
                for (const o of orders) {
                    const ordNo = Number(o.OrdNo);
                    if (!Number.isFinite(ordNo)) continue;

                    if (o.TotalCost !== null && o.TotalCost !== undefined) {
                        marginStateByOrdNo[String(ordNo)] = {
                            status: 'success',
                            totalRevenue: Number(o.InvoSF || 0),
                            totalCost: Number(o.TotalCost || 0)
                        };
                    }
                }
            }

            function queueMarginLoad(ordNos) {
                for (const ordNo of ordNos) {
                    const key = String(ordNo);
                    const existing = marginStateByOrdNo[key];
                    if (existing && (existing.status === 'success' || existing.status === 'loading')) {
                        continue;
                    }

                    marginStateByOrdNo[key] = { status: 'loading' };
                    marginJobQueue.push(Number(ordNo));
                }
                pumpMarginQueue();
                scheduleOrderListRerender();
            }

            async function loadSingleOrderMargin(ordNo) {
                const key = String(ordNo);
                try {
                    const response = await fetch('/order-margin/' + ordNo);
                    const data = await response.json();
                    if (!response.ok || data.error) {
                        marginStateByOrdNo[key] = { status: 'error' };
                        return;
                    }

                    marginStateByOrdNo[key] = {
                        status: 'success',
                        totalRevenue: Number(data.totalRevenue || 0),
                        totalCost: Number(data.totalCost || 0)
                    };
                } catch (err) {
                    marginStateByOrdNo[key] = { status: 'error' };
                }
            }

            function pumpMarginQueue() {
                while (marginWorkerActiveCount < MARGIN_MAX_CONCURRENT && marginJobQueue.length > 0) {
                    const ordNo = marginJobQueue.shift();
                    marginWorkerActiveCount += 1;

                    loadSingleOrderMargin(ordNo)
                        .finally(() => {
                            marginWorkerActiveCount -= 1;
                            scheduleOrderListRerender();
                            setTimeout(pumpMarginQueue, MARGIN_QUEUE_DELAY_MS);
                        });
                }
            }

            function toggleOrderList() {
                orderListVisible = !orderListVisible;
                renderOrderList();
            }

            function setOrderListFilter() {
                const input = document.getElementById('customerFilterInput');
                orderListFilter = (input && input.value ? input.value : '').trim().toLowerCase();
                if (!orderListVisible && orderListFilter) {
                    orderListVisible = true;
                }
                renderOrderList();
            }

            function setBrugerFilter() {
                const input = document.getElementById('brugerFilterSelect');
                orderListBrugerFilter = (input && input.value ? input.value : '').trim();
                if (!orderListVisible && orderListBrugerFilter) {
                    orderListVisible = true;
                }
                renderOrderList();
            }

            function populateBrugerFilterOptions() {
                const select = document.getElementById('brugerFilterSelect');
                if (!select) return;

                const selectedValue = orderListBrugerFilter;
                const users = Array.from(new Set(
                    orderListData
                        .map(o => String(o.SellerUsr || '').trim())
                        .filter(v => v)
                )).sort((a, b) => a.localeCompare(b));

                let html = '<option value="">Alle brugere</option>';
                for (const user of users) {
                    html += '<option value="' + user + '">' + user + '</option>';
                }
                select.innerHTML = html;
                select.value = selectedValue;
            }

            function getFilteredOrders() {
                return orderListData.filter(o => {
                    const bruger = String(o.SellerUsr || '').trim();
                    const customer = String(o.CustomerName || '').toLowerCase();
                    const ord = String(o.OrdNo || '');
                    const matchesText = !orderListFilter || customer.includes(orderListFilter) || ord.includes(orderListFilter);
                    const matchesBruger = !orderListBrugerFilter || bruger === orderListBrugerFilter;
                    return matchesText && matchesBruger;
                }).sort((a, b) => {
                    const dateA = Number(a.LstInvDt || 0);
                    const dateB = Number(b.LstInvDt || 0);
                    if (dateB !== dateA) return dateB - dateA;
                    return Number(b.OrdNo || 0) - Number(a.OrdNo || 0);
                });
            }

            function getMarginBadge(marginPercent) {
                const margin = parseFloat(marginPercent);
                if (currentMarginMode === 'new') {
                    if (margin >= 125) {
                        return '<span style="background:#2e7d32; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">✅ ' + marginPercent + '%</span>';
                    } else if (margin >= 105) {
                        return '<span style="background:#ff9800; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">⚠️ ' + marginPercent + '%</span>';
                    }
                    return '<span style="background:#d32f2f; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">❌ ' + marginPercent + '%</span>';
                }

                if (margin > 20) {
                    return '<span style="background:#2e7d32; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">✅ ' + marginPercent + '%</span>';
                } else if (margin >= 5) {
                    return '<span style="background:#ff9800; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">⚠️ ' + marginPercent + '%</span>';
                }
                return '<span style="background:#d32f2f; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">❌ ' + marginPercent + '%</span>';
            }

            async function searchOrder() {
                const ordNo = document.getElementById('orderInput').value;
                if (!ordNo) {
                    alert('Indtast et ordrenummer');
                    return;
                }
                
                const result = document.getElementById('result');
                result.innerHTML = '<div class="loading">Indlaeser...</div>';
                
                try {
                    const response = await fetch('/aftercalc/' + ordNo);
                    const data = await response.json();
                    
                    if (data.error) {
                        result.innerHTML = '<div class="error">Fejl: ' + data.error + '</div>';
                        return;
                    }

                    const orderMarginPercent = calculateOrderMarginPercent(data.summary.totalRevenue, data.summary.totalCost).toFixed(2);
                    
                    let html = '<div class="order-header">';
                    html += '<h2>Salgsordre: ' + data.orderHeader.OrdNo + ' - ' + (data.orderHeader.CustomerName || '-') + '</h2>';
                    html += '<div class="order-header-row">';
                    html += '<div class="order-header-item"><div class="order-header-label">Faktureret beløb</div><div class="order-header-value">' + formatNumber(data.summary.totalRevenue) + ' DKK</div></div>';
                    html += '<div class="order-header-item"><div class="order-header-label">Kostpris</div><div class="order-header-value">' + formatNumber(data.summary.totalCost) + ' DKK</div></div>';
                    html += '<div class="order-header-item"><div class="order-header-label">Margin (' + getMarginModeLabel() + ')</div><div class="order-header-value">' + getMarginBadge(orderMarginPercent) + '</div></div>';
                    html += '</div></div>';

                    // Sezione linee ORDINE DI VENDITA complete
                    if (data.salesOrderLines && data.salesOrderLines.length > 0) {
                        html += '<div class="section"><h3>Salgsordrelinjer</h3>';
                        html += '<table><tr><th>Linje</th><th>Produkt</th><th>Beskrivelse</th><th>Antal</th><th>Kostpris</th><th>Samlet kost</th><th>Salgspris</th><th>Margin (%)</th><th>Prod.ordre</th></tr>';

                        for (const line of data.salesOrderLines) {
                            const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
                            const lineCost = line.EffectiveLineCost || 0;
                            const lineMarginPercent = calculateLineMarginPercent(lineSalesPrice, lineCost).toFixed(2);
                            const lineMarginBadge = lineSalesPrice === 0
                                ? '<span style="background:#757575; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">Rabatt</span>'
                                : getMarginBadge(lineMarginPercent);
                            html += '<tr>';
                            html += '<td>' + (line.LnNo || 0) + '</td>';

                            if (line.PurcNo && line.PurcNo !== 0) {
                                html += '<td><span class="prod-link" onclick="openProduction(' + line.PurcNo + ')">' + (line.ProdNo || '-') + '</span></td>';
                            } else {
                                html += '<td>' + (line.ProdNo || '-') + '</td>';
                            }

                            html += '<td>' + (line.Descr || '') + '</td>';
                            html += '<td>' + formatNumber(line.NoFin || 0) + '</td>';
                            const displayKostpris = (line.PurcNo && line.PurcNo !== 0)
                                ? (line.ProductionOrderTotalCost || 0)
                                : (line.CCstPr || 0);
                            html += '<td>' + formatNumber(displayKostpris) + '</td>';
                            html += '<td><strong>' + formatNumber(lineCost) + '</strong></td>';
                            html += '<td>' + formatNumber(lineSalesPrice) + '</td>';
                            html += '<td>' + lineMarginBadge + '</td>';
                            html += '<td>' + ((line.PurcNo && line.PurcNo !== 0) ? line.PurcNo : '-') + '</td>';
                            html += '</tr>';
                        }

                        html += '</table></div>';
                    }
                    
                    // Sezione linee di vendita
                    if (data.salesLines.length > 0) {
                        html += '<div class="section"><h3>Salgslinjer (Ekstra produkter)</h3>';
                        html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Antal</th><th>Salgspris</th><th>Kostpris/enhed</th><th>Samlet kost</th></tr>';
                        
                        for (const line of data.salesLines) {
                            html += '<tr>';
                            html += '<td>' + (line.ProdNo || '-') + '</td>';
                            html += '<td>' + (line.Descr || '') + '</td>';
                            html += '<td>' + formatNumber(line.NoFin || 0) + '</td>';
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                            html += '<td>' + formatNumber(line.CCstPr || 0) + '</td>';
                            html += '<td><strong>' + formatNumber(line.EffectiveLineCost || 0) + '</strong></td>';
                            html += '</tr>';
                        }
                        
                        html += '<tr class="summary-row"><td colspan="5">Total salgslinjer:</td><td>' + formatNumber(data.salesLinesTotalCost) + ' DKK</td></tr>';
                        html += '</table></div>';
                    }
                    
                    // Sezione ordini di produzione
                    if (data.productionOrders.length > 0) {
                        html += '<div class="section"><h3>Produktionsordrer</h3>';
                        const prodTp4Labels = {
                            '1': 'Operation',
                            '2': 'Materiale Laser',
                            '4': 'Produkt dele',
                            '5': 'Rute',
                            '6': 'Ydelse',
                            '7': 'Underleverandor',
                            '8': 'Materiale fast antal',
                            'NA': 'Ikke sat'
                        };
                        
                        for (const prodOrder of data.productionOrders) {
                            const mainProductLine = prodOrder.lines.find(line => line.ProdTp4 === 0) || prodOrder.lines.find(line => line.LnNo === 1);
                            const mainProductText = mainProductLine
                                ? ((mainProductLine.ProdNo || '-') + ' - ' + (mainProductLine.Descr || ''))
                                : '-';

                            html += '<div id="po-' + prodOrder.ordNo + '" data-order="' + prodOrder.ordNo + '" style="margin-bottom: 20px; border: 1px solid #ddd; padding: 15px; border-radius: 4px;">';
                            html += '<h4>Produktionsordre: ' + prodOrder.ordNo + '</h4>';
                            html += '<div class="main-product-box">';
                            html += '<div class="value">' + mainProductText + '</div>';
                            html += '</div>';
                            html += '<div class="prodtp4-hint">Klik paa en linje for at aabne/lukke detaljer.</div>';

                            const groupedLines = {};
                            for (const line of prodOrder.lines) {
                                const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                                if (key === '0' || key === '3' || key === '5') continue;
                                if (!groupedLines[key]) groupedLines[key] = [];
                                groupedLines[key].push(line);
                            }

                            const groupKeys = Object.keys(groupedLines).sort((a, b) => {
                                if (a === 'NA') return 1;
                                if (b === 'NA') return -1;
                                return Number(a) - Number(b);
                            });

                            let orderVisibleTotal = 0;

                            for (let i = 0; i < groupKeys.length; i++) {
                                const key = groupKeys[i];
                                const lines = groupedLines[key];
                                const subtotal = key === '2'
                                    ? lines.filter(line => line.LnNo !== 1).reduce((sum, line) => sum + ((line.NestingCost || 0) * (line.NoFin || 0)), 0)
                                    : lines.filter(line => line.LnNo !== 1).reduce((sum, line) => sum + (line.LineCost || 0), 0);
                                const isOpenByDefault = false;
                                orderVisibleTotal += subtotal;

                                html += '<div class="prodtp4-group">';
                                html += '<div class="prodtp4-header" onclick="toggleProdTp4Group(' + prodOrder.ordNo + ', &quot;' + key + '&quot;)">';
                                html += '<span class="prodtp4-label"><span id="po-' + prodOrder.ordNo + '-group-' + key + '-icon">' + (isOpenByDefault ? '▾' : '▸') + '</span> ' + key + ' - ' + (prodTp4Labels[key] || 'Altro') + '</span>';
                                html += '<span class="prodtp4-subtotal">Delsum: ' + formatNumber(subtotal) + ' DKK</span>';
                                html += '</div>';

                                html += '<div id="po-' + prodOrder.ordNo + '-group-' + key + '" class="prodtp4-body" style="display:' + (isOpenByDefault ? '' : 'none') + ';">';
                                if (key === '2') {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Antal</th><th>Kostpris/enhed</th><th>Kostpris nesting</th><th>Samlet kost</th></tr>';
                                } else {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Antal</th><th>Salgspris</th><th>Kostpris/enhed</th><th>Samlet kost</th></tr>';
                                }

                                for (const line of lines) {
                                    html += '<tr>';
                                    if (String(key) === '4' && line.PurcNo && line.PurcNo !== 0) {
                                        html += '<td><span class="inline-link" onclick="showChildProductionSummary(' + line.PurcNo + ')">' + (line.ProdNo || '-') + '</span></td>';
                                    } else if (line.ProdNo) {
                                        const safeProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        html += '<td><span class="prod-no-link" data-prodno="' + safeProdNo + '" data-ordno="' + prodOrder.ordNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="' + key + '">' + safeProdNo + '</span></td>';
                                    } else {
                                        html += '<td>-</td>';
                                    }
                                    html += '<td>' + (line.Descr || '') + '</td>';
                                    html += '<td>' + formatNumber(line.NoFin || 0) + '</td>';
                                    if (key === '2') {
                                        const nestingSamlet = (line.NestingCost || 0) * (line.NoFin || 0);
                                        html += '<td>' + formatNumber(line.CCstPr || 0) + '</td>';
                                        html += '<td>' + formatNumber(line.NestingCost || 0) + '</td>';
                                        html += '<td><strong>' + formatNumber(nestingSamlet) + '</strong></td>';
                                    } else {
                                        html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                                        html += '<td>' + formatNumber(line.CCstPr || 0) + '</td>';
                                        html += '<td><strong>' + formatNumber(line.LineCost || 0) + '</strong></td>';
                                    }
                                    html += '</tr>';
                                }

                                html += '</table>';
                                html += '</div>';
                                html += '</div>';
                            }
                            
                            html += '<div class="po-total-row">Total ordre: <span id="po-total-' + prodOrder.ordNo + '">' + formatNumber(orderVisibleTotal) + ' DKK</span></div>';
                            html += '</div>';
                        }
                        
                        html += '</div>';
                    }
                    
                    // Riepilogo
                    html += '<div class="section"><h3>Ordresammendrag</h3>';
                    html += '<div class="summary-box">';
                    html += '<div><strong>Samlet faktureret beløb:</strong> ' + formatNumber(data.summary.totalRevenue) + ' DKK</div>';
                    html += '<div><strong>Samlet kost:</strong> ' + formatNumber(data.summary.totalCost) + ' DKK</div>';
                    let marginClass = data.summary.margin >= 0 ? 'margin-positive' : 'margin-negative';
                    html += '<div class="total"><span class="' + marginClass + '">Margin: ' + formatNumber(data.summary.margin) + ' DKK (' + orderMarginPercent + '%)</span></div>';
                    html += '</div></div>';
                    
                    result.innerHTML = html;
                } catch (err) {
                    result.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                }
            }

            async function onProductClick(prodNo, ordNo, lnNo, prodTp4) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');

                title.textContent = 'Produkt: ' + prodNo;
                modal.style.display = 'flex';

                if (String(prodTp4) === '1') {
                    body.innerHTML = '<div class="modal-loading">Indlaeser transaktioner...</div>';
                    try {
                        const response = await fetch('/prodtr/' + ordNo + '/' + lnNo);
                        const rows = await response.json();
                        if (!response.ok || rows.error) {
                            body.innerHTML = '<div class="error">Fejl: ' + (rows.error || 'Uventet fejl') + '</div>';
                            return;
                        }
                        if (!rows.length) {
                            body.innerHTML = '<div>Ingen ProdTr-linjer fundet.</div>';
                            return;
                        }
                        let html = '<table>';
                        html += '<tr><th>Færdigmeldingsdato</th><th>Færdigmeldingstid</th><th>Antal Minutter</th><th>Hvem</th></tr>';
                        for (const r of rows) {
                            const finDt = String(r.FinDt || '').split('T')[0];
                            const finTm = r.FinTm != null ? String(r.FinTm).padStart(4, '0').replace(/^(\d{2})(\d{2})$/, '$1:$2') : '-';
                            html += '<tr>';
                            html += '<td>' + (finDt || '-') + '</td>';
                            html += '<td>' + finTm + '</td>';
                            html += '<td>' + formatNumber(r.NoInvoAb || 0) + '</td>';
                            html += '<td>' + (r.HvemNm || '-') + '</td>';
                            html += '</tr>';
                        }
                        html += '</table>';
                        body.innerHTML = html;
                    } catch (err) {
                        body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                    }
                } else if (String(prodTp4) === '2') {
                    body.innerHTML = '<div class="modal-loading">Indlaeser nesting-detaljer...</div>';
                    try {
                        const encProdNo = encodeURIComponent(prodNo);
                        const response = await fetch('/nesting-detail/' + ordNo + '/' + encProdNo);
                        const rows = await response.json();
                        if (!response.ok || rows.error) {
                            body.innerHTML = '<div class="error">Fejl: ' + (rows.error || 'Uventet fejl') + '</div>';
                            return;
                        }
                        if (!rows.length) {
                            body.innerHTML = '<div>Ingen nesting-linjer fundet.</div>';
                            return;
                        }
                        const avgCst = rows.reduce((s, r) => s + (r.CstPr || 0), 0) / rows.length;
                        let html = '<table>';
                        html += '<tr><th>OrdNo</th><th>Beskrivelse</th><th>TrInf4</th><th>Antal (NoFin)</th><th>CstPr</th></tr>';
                        for (const r of rows) {
                            html += '<tr>';
                            html += '<td>' + (r.OrdNo ?? '-') + '</td>';
                            html += '<td>' + (r.Descr || '-') + '</td>';
                            html += '<td>' + (r.TrInf4 || '-') + '</td>';
                            html += '<td>' + formatNumber(r.NoFin || 0) + '</td>';
                            html += '<td>' + formatNumber(r.CstPr || 0) + '</td>';
                            html += '</tr>';
                        }
                        html += '<tr class="summary-row"><td colspan="4"><strong>Gennemsnit CstPr (brugt i beregning):</strong></td><td><strong>' + formatNumber(avgCst) + '</strong></td></tr>';
                        html += '</table>';
                        body.innerHTML = html;
                    } catch (err) {
                        body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                    }
                }
            }

            document.addEventListener('click', function(e) {
                const span = e.target.closest('.prod-no-link');
                if (!span) return;
                const prodNo = span.dataset.prodno;
                const ordNo = span.dataset.ordno;
                const lnNo = span.dataset.lnno;
                const prodTp4 = span.dataset.prodtp4;
                if (prodNo) onProductClick(prodNo, ordNo, lnNo, prodTp4);
            });

            function toggleProdTp4Group(orderNo, prodTp4Key) {
                const el = document.getElementById('po-' + orderNo + '-group-' + prodTp4Key);
                const icon = document.getElementById('po-' + orderNo + '-group-' + prodTp4Key + '-icon');
                if (!el) return;
                const isClosed = el.style.display === 'none';
                el.style.display = isClosed ? '' : 'none';
                if (icon) icon.textContent = isClosed ? '▾' : '▸';
            }

            async function showChildProductionSummary(childOrdNo) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');

                title.textContent = 'Produktoversigt for ordre ' + childOrdNo;
                body.innerHTML = '<div class="modal-loading">Indlaeser...</div>';
                modal.style.display = 'flex';

                try {
                    const response = await fetch('/production-summary/' + childOrdNo);
                    const data = await response.json();

                    if (!response.ok || data.error) {
                        body.innerHTML = '<div class="error">Fejl: ' + (data.error || 'Uventet fejl') + '</div>';
                        return;
                    }

                    if (!data.lines || data.lines.length === 0) {
                        body.innerHTML = '<div>Ingen linjer med positivt antal (NoFin > 0).</div>';
                        return;
                    }

                    let html = '';
                    html += '<table><tr><th>Linje</th><th>ProdTp4</th><th>Prod</th><th>Beskrivelse</th><th>Antal</th><th>Salgspris</th><th>Kostpris/enhed</th><th>Nesting/enhed</th><th>Samlet kost (beregnet)</th></tr>';
                    for (const line of data.lines) {
                        html += '<tr>';
                        html += '<td>' + (line.LnNo || 0) + '</td>';
                        html += '<td>' + (line.ProdTp4 === null || line.ProdTp4 === undefined ? '-' : line.ProdTp4) + '</td>';
                        html += '<td>' + (line.ProdNo || '-') + '</td>';
                        html += '<td>' + (line.Descr || '') + '</td>';
                        html += '<td>' + formatNumber(line.NoFin || 0) + '</td>';
                        html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                        html += '<td>' + formatNumber(line.CCstPr || 0) + '</td>';
                        html += '<td>' + formatNumber(line.NestingCost || 0) + '</td>';
                        html += '<td><strong>' + formatNumber(line.EffectiveLineCost || 0) + '</strong></td>';
                        html += '</tr>';
                    }
                    html += '<tr class="summary-row"><td colspan="8">Total beregnet kost:</td><td><strong>' + formatNumber(data.totalCost || 0) + ' DKK</strong></td></tr>';
                    html += '</table>';
                    body.innerHTML = html;
                } catch (err) {
                    body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                }
            }

            function closeSummaryModal(event) {
                if (event && event.target && event.target.id !== 'summaryModal') return;
                const modal = document.getElementById('summaryModal');
                modal.style.display = 'none';
            }

            function openProduction(ordNo) {
                const el = document.getElementById('po-' + ordNo);
                if (!el) {
                    alert('Produktionsordre ' + ordNo + ' blev ikke fundet i de indlaeste resultater.');
                    return;
                }

                el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                el.classList.add('po-highlight');
                setTimeout(() => el.classList.remove('po-highlight'), 1800);
            }
            
            function renderOrderList() {
                const el = document.getElementById('orderList');
                const toggleBtn = document.getElementById('listToggleBtn');

                if (!orderListVisible) {
                    if (toggleBtn) toggleBtn.textContent = 'Vis kundeliste';
                    el.innerHTML = '';
                    return;
                }

                if (toggleBtn) toggleBtn.textContent = 'Skjul kundeliste';
                if (!orderListData || orderListData.length === 0) {
                    el.innerHTML = '<div class="loading">Indlaeser ordreliste...</div>';
                    return;
                }

                const orders = getFilteredOrders();
                if (orders.length === 0) {
                    el.innerHTML = '<div class="order-list-section"><h3>Ingen kunder fundet</h3><div>Proev en anden soegning.</div></div>';
                    return;
                }

                let html = '<div class="order-list-section">';
                html += '<h3>Seneste fakturerede ordrer (${ORDER_LIST_DAYS_BACK} dage) &mdash; ' + orders.length + ' af ' + orderListData.length + ' ordrer</h3>';
                html += '<table class="order-list-table"><tr><th>Bruger</th><th>Ordrenr.</th><th>Kunde</th><th>Fakturadato</th><th>Fakturabelob</th><th>Margin</th></tr>';
                for (const o of orders) {
                    const marginState = getMarginState(o.OrdNo);
                    let marginHtml = '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>';
                    if (marginState && marginState.status === 'loading') {
                        marginHtml = '<span style="background:#546e7a; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">...</span>';
                    } else if (marginState && marginState.status === 'error') {
                        marginHtml = '<span style="background:#8d6e63; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">ERR</span>';
                    } else if (marginState && marginState.status === 'success') {
                        const margin = calculateOrderMarginPercent(marginState.totalRevenue || 0, marginState.totalCost || 0).toFixed(2);
                        marginHtml = getMarginBadge(margin);
                    }

                    const d = String(o.LstInvDt || '');
                    const invDate = d.length === 8 ? d.slice(0,4) + '-' + d.slice(4,6) + '-' + d.slice(6,8) : (d || '-');
                    html += '<tr onclick="selectOrder(' + o.OrdNo + ')">'
                    html += '<td>' + (o.SellerUsr || '-') + '</td>';
                    html += '<td><strong>' + o.OrdNo + '</strong></td>';
                    html += '<td>' + (o.CustomerName || '-') + '</td>';
                    html += '<td>' + invDate + '</td>';
                    html += '<td>' + formatNumber(o.InvoSF || 0) + ' DKK</td>';
                    html += '<td>' + marginHtml + '</td>';
                    html += '</tr>';
                }
                html += '</table></div>';
                el.innerHTML = html;

                // Carica i margini in coda per tutti gli ordini visibili.
                const queuedOrders = orders.slice(0, MARGIN_PREFETCH_ROWS);
                queueMarginLoad(queuedOrders.map(o => o.OrdNo));
                updateSystemStatusFromOrders(queuedOrders);
            }

            async function loadOrderList(forceRefresh = false) {
                const el = document.getElementById('orderList');
                if (!el) return;

                if (orderListLoading) return;
                orderListLoading = true;
                const previousHtml = el.innerHTML;
                setSystemStatus('System loading...', '#fff3cd', '#8a6d3b');
                if (!orderListData || orderListData.length === 0) {
                    el.innerHTML = '<div class="loading">Indlaeser ordreliste...</div>';
                }
                try {
                    const endpoint = forceRefresh
                        ? '/order-list?force=1&t=' + Date.now()
                        : '/order-list';
                    const response = await fetch(endpoint);
                    if (!response.ok) {
                        setSystemStatus('System error', '#fdecea', '#b71c1c');
                        if (previousHtml) {
                            el.innerHTML = previousHtml;
                        }
                        return;
                    }
                    const orders = await response.json();
                    if (!orders || orders.error) {
                        setSystemStatus('System error', '#fdecea', '#b71c1c');
                        if (previousHtml) {
                            el.innerHTML = previousHtml;
                        }
                        return;
                    }
                    orderListData = orders;
                    hydrateMarginStateFromOrderList(orders);
                    populateBrugerFilterOptions();
                    renderOrderList();
                } catch (err) {
                    console.error('Fejl i loadOrderList:', err);
                    setSystemStatus('System error', '#fdecea', '#b71c1c');
                    if (previousHtml) {
                        el.innerHTML = previousHtml;
                    }
                } finally {
                    orderListLoading = false;
                }
            }

            function startOrderListAutoRefresh() {
                if (orderListAutoRefreshTimer) return;
                orderListAutoRefreshTimer = setInterval(() => {
                    if (document.hidden) return;
                    loadOrderList(false); // Bruger server-cache, ingen DB-foresporgsel medmindre cachen er forældet
                }, ORDER_LIST_AUTO_REFRESH_MS);
            }

            async function refreshOrderList() {
                const btn = document.getElementById('refreshListBtn');
                if (btn) {
                    btn.disabled = true;
                    btn.textContent = 'Opdaterer...';
                }

                try {
                    await loadOrderList(true);
                } finally {
                    if (btn) {
                        btn.disabled = false;
                        btn.textContent = 'Opdater liste';
                    }
                }
            }

            function selectOrder(ordNo) {
                document.getElementById('orderInput').value = ordNo;
                orderListVisible = false;
                renderOrderList();
                searchOrder();
                setTimeout(() => {
                    const resultEl = document.getElementById('result');
                    if (resultEl) {
                        resultEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                }, 150);
            }

            function toggleSearchBox() {
                const searchBox = document.getElementById('searchBox');
                const collapseToggleBtn = document.getElementById('collapseToggleBtn');
                const collapseExpandBtn = document.getElementById('collapseExpandBtn');
                
                searchBox.classList.toggle('collapsed');
                if (searchBox.classList.contains('collapsed')) {
                    collapseToggleBtn.style.display = 'inline-block';
                    collapseExpandBtn.style.display = 'none';
                    collapseToggleBtn.textContent = '▼ Søg';
                } else {
                    collapseToggleBtn.style.display = 'none';
                    collapseExpandBtn.style.display = 'inline-block';
                }
            }

            // Soeg ved indlaesning hvis ordrenummer er i query string
            window.onload = function() {
                loadOrderList();
                startOrderListAutoRefresh();
                const params = new URLSearchParams(window.location.search);
                if (params.has('ord')) {
                    document.getElementById('orderInput').value = params.get('ord');
                    searchOrder();
                }
            };
        </script>
    </body>
    </html>
    `);
});

const PORT = 3000;
app.listen(PORT, () => {
    console.log('Server in ascolto su http://localhost:' + PORT);
    logEvent('Server started - Ready to accept requests');

    refreshOrderListCache(true)
        .then(() => {
            logEvent('Cache primed: order list loaded and margin warmup started');
        })
        .catch(err => {
            logEvent('ERROR cache warmup: ' + err.message);
        });
});
