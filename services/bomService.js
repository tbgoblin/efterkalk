function createBomService({ getConnection, sql, diskCache, logEvent }) {
    const memoryCache = new Map();

    const TTL = {
        customers: 8 * 60 * 60 * 1000,
        products: 2 * 60 * 60 * 1000,
        revisions: 30 * 60 * 1000,
        resources: 8 * 60 * 60 * 1000,
        materials: 15 * 60 * 60 * 1000,
        calculators: 8 * 60 * 60 * 1000
    };

    function nowMs() {
        return Date.now();
    }

    function makeKey(prefix, value) {
        return 'bom_v1_' + prefix + '_' + String(value || 'all');
    }

    function getCached(cacheKey) {
        const mem = memoryCache.get(cacheKey);
        if (mem && mem.expiresAt > nowMs()) {
            return mem.data;
        }
        const disk = diskCache.get(cacheKey);
        if (disk !== null && disk !== undefined) {
            return disk;
        }
        return null;
    }

    function setCached(cacheKey, ttlMs, data) {
        memoryCache.set(cacheKey, {
            expiresAt: nowMs() + ttlMs,
            data
        });
        diskCache.set(cacheKey, data, ttlMs);
    }

    function uniqueNonEmpty(values) {
        return Array.from(new Set(values.map(value => String(value || '').trim()).filter(Boolean)));
    }

    async function fetchCustomers(options = {}) {
        const q = String(options.q || '').trim().toLowerCase();
        const limit = Math.max(0, Math.min(20000, Number(options.limit || 0)));
        const key = makeKey('customers', q + '_' + limit);
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT
                Actor.Gr AS [Varenr.],
                Actor.CustNo AS [Kundernr.],
                Actor.Nm AS [Kundenavn],
                Actor.Ad1 AS [Adresse1],
                Actor.Ad2 AS [Adresse2],
                Actor.PNo AS [Postnr.],
                Actor.PArea AS [By],
                Actor.Phone AS [Tlf.],
                Actor.CustPrGr AS [Prisliste],
                Actor.Inf6 AS [Kunde],
                Actor.Shrt AS [Kortnavn]
            FROM Actor WITH(NOLOCK)
            WHERE Actor.CustNo <> 0
        `);

        let rows = (Array.isArray(result.recordset) ? result.recordset : []).map(row => ({
            ...row,
            Gr: row['Varenr.'],
            CustNo: row['Kundernr.'],
            Nm: row['Kundenavn'],
            Ad1: row['Adresse1'],
            Ad2: row['Adresse2'],
            PNo: row['Postnr.'],
            PArea: row['By'],
            Phone: row['Tlf.'],
            CustPrGr: row['Prisliste'],
            Inf6: row['Kunde'],
            Shrt: row['Kortnavn']
        }));
        if (q) {
            rows = rows.filter(row => {
                const haystack = [
                    row.CustNo,
                    row.Nm,
                    row.Shrt,
                    row.Gr,
                    row.Ad1,
                    row.Ad2,
                    row.PNo,
                    row.PArea,
                    row.Phone,
                    row.CustPrGr,
                    row.Inf6
                ].map(v => String(v || '').toLowerCase()).join(' ');
                return haystack.includes(q);
            });
        }
        if (limit > 0) {
            rows = rows.slice(0, limit);
        }

        const payload = {
            rows,
            count: rows.length,
            cached: true,
            source: 'db'
        };
        setCached(key, TTL.customers, payload);
        return payload;
    }

    async function fetchProductsByCustomer(options = {}) {
        const customerKeys = uniqueNonEmpty([options.customerNo, options.customerCode]);
        if (customerKeys.length === 0) {
            throw new Error('customerNo mangler');
        }

        const limit = Math.max(0, Math.min(20000, Number(options.limit || 0)));
        const key = makeKey('products', customerKeys.join('_') + '_' + limit);
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const request = pool.request();
        const filters = customerKeys.map((value, index) => {
            const paramName = 'customerKey' + index;
            request.input(paramName, sql.VarChar, value);
            return 'Prod.Inf3 = @' + paramName;
        });

        const result = await request.query(`
            SELECT
                Prod.ProdNo,
                Prod.Descr,
                Prod.Inf3,
                1 AS chck,
                Prod.Inf2,
                Prod.Inf4,
                Prod.Inf7,
                Prod.Inf8
            FROM Prod WITH(NOLOCK)
                        WHERE (${filters.join(' OR ')})
              AND Prod.ProdGr <> 99999
            ORDER BY Prod.ProdNo
        `);

        let rows = (Array.isArray(result.recordset) ? result.recordset : []).map(row => ({
            ...row,
            TgNo: row.Inf2,
            CustomerNo: row.Inf3,
            CustomerNoAlt: row.Inf4,
            RevNo: row.Inf7,
            PosNo: row.Inf8
        }));
        if (limit > 0) {
            rows = rows.slice(0, limit);
        }
        const payload = {
            customerNo: String(options.customerNo || '').trim(),
            customerCode: String(options.customerCode || '').trim(),
            rows,
            count: rows.length,
            cached: true,
            source: 'db'
        };
        setCached(key, TTL.products, payload);
        return payload;
    }

    async function fetchRevisionsByDrawing(options = {}) {
        const tgn = String(options.tgn || '').trim();
        const customerKeys = uniqueNonEmpty([options.customerNo, options.customerCode]);
        if (!tgn) {
            throw new Error('tgn mangler');
        }
        if (customerKeys.length === 0) {
            throw new Error('customerNo mangler');
        }

        const key = makeKey('revisions', customerKeys.join('_') + '_' + tgn);
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const request = pool.request()
            .input('tgn', sql.VarChar, tgn);
        const filters = customerKeys.map((value, index) => {
            const paramName = 'customerKey' + index;
            request.input(paramName, sql.VarChar, value);
            return 'Prod.Inf3 = @' + paramName;
        });

        const result = await request.query(`
            SELECT
                Prod.ProdNo AS ProdNo,
                Prod.Descr AS Beskrivelse,
                Prod.Inf4 AS KundeNo,
                Prod.Inf2 AS TgNo,
                Prod.Inf7 AS RevNo,
                Prod.Inf8 AS PosNo
            FROM Prod WITH(NOLOCK)
            WHERE Prod.Inf2 = @tgn
              AND Prod.Inf2 <> ''
                            AND (${filters.join(' OR ')})
        `);

        const rows = (Array.isArray(result.recordset) ? result.recordset : []).map(row => ({
            ...row,
            Descr: row.Beskrivelse,
            CustomerNoAlt: row.KundeNo
        }));
        const payload = {
            tgn,
            customerNo: String(options.customerNo || '').trim(),
            customerCode: String(options.customerCode || '').trim(),
            rows,
            count: rows.length,
            cached: true,
            source: 'db'
        };
        setCached(key, TTL.revisions, payload);
        return payload;
    }

    async function fetchResources() {
        const key = makeKey('resources', 'all');
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT
                Prod.ProdNo,
                Prod.Descr,
                PrDcMat.CstPr,
                PrDcMat.SalePr,
                Prod.ProdGr,
                Prod.DensU,
                BgtLn.R7,
                Prod.Gr4,
                Prod.Inf3 AS CustomerNo
            FROM BgtLn WITH(NOLOCK)
            JOIN PrDcMat WITH(NOLOCK)
                ON PrDcMat.ProdNo = BgtLn.ProdNo
            JOIN Prod WITH(NOLOCK)
                ON PrDcMat.ProdNo = Prod.ProdNo
            WHERE Prod.ProdGr = 3
            ORDER BY Prod.ProdNo
        `);

        const rows = Array.isArray(result.recordset) ? result.recordset : [];
        const payload = {
            rows,
            count: rows.length,
            cached: true,
            source: 'db'
        };
        setCached(key, TTL.resources, payload);
        return payload;
    }

    async function fetchMaterials(options = {}) {
        const q = String(options.q || '').trim().toLowerCase();
        const limit = Math.max(0, Math.min(10000, Number(options.limit || 0)));
        const key = makeKey('materials', q + '_' + limit);
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT
                Prod.ProdNo,
                Prod.Inf6 AS Lager,
                Prod.HgtU AS tykklese,
                Prod.Descr AS beskrivelse,
                Prod.WdtU AS Bredde,
                Prod.LgtU AS [Længde],
                Txt.Txt AS Enhed,
                Prod.Inf AS Pris,
                Prod.DensU AS [VægtFylde],
                Prod.R3 AS RV,
                Prod.Free2 AS Avance,
                Prod.NWgtU,
                StcBal.PoPhStB,
                StcBal.Bal,
                StcBal.StcInc
            FROM Prod WITH(NOLOCK)
            JOIN Txt WITH(NOLOCK)
                ON Txt.TxtNo = Prod.StSaleUn
            LEFT JOIN StcBal WITH(NOLOCK)
                ON StcBal.ProdNo = Prod.ProdNo
               AND StcBal.StcNo = 1
            WHERE Txt.TxtTp = 16
              AND Txt.Lang = 45
              AND Prod.ProdGr = 2
              AND Prod.ProdNo NOT LIKE '1%'
              AND Prod.ProdNo <> '301001'
            ORDER BY Prod.ProdNo
        `);

        let rows = (Array.isArray(result.recordset) ? result.recordset : []).map(row => ({
            ...row,
            Thickness: row.tykklese,
            Descr: row.beskrivelse,
            Width: row.Bredde,
            Length: row['Længde'],
            UnitText: row.Enhed,
            Price: row.Pris,
            Density: row['VægtFylde'],
            Stock: row.Bal,
            StockUnitFactor: row.PoPhStB,
            StockIncrement: row.StcInc
        }));
        if (q) {
            rows = rows.filter(row => {
                const haystack = [
                    row.ProdNo,
                    row.Descr,
                    row.UnitText,
                    row.RV,
                    row.Lager,
                    row.Thickness,
                    row.Width,
                    row.Length
                ]
                    .map(v => String(v || '').toLowerCase())
                    .join(' ');
                return haystack.includes(q);
            });
        }
        if (limit > 0) {
            rows = rows.slice(0, limit);
        }

        const payload = {
            rows,
            count: rows.length,
            cached: true,
            source: 'db'
        };
        setCached(key, TTL.materials, payload);
        return payload;
    }

    async function fetchLaserParameters(options = {}) {
        const machine = String(options.machine || '').trim().toLowerCase();
        const key = makeKey('laser_params', machine || 'all');
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT
                FreeInf2.ProdNo,
                Prod.Descr,
                Prod.HgtU AS Tykkelse,
                FreeInf2.Txt1 AS Maskine,
                FreeInf2.Val1 AS [Skærehast.],
                FreeInf2.Val2 AS Pircing,
                FreeInf2.Val3 AS [Tillæg],
                FreeInf2.Txt2 AS Linse
            FROM FreeInf2 WITH(NOLOCK), Prod WITH(NOLOCK)
            WHERE FreeInf2.ProdNo = Prod.ProdNo
              AND FreeInf2.FrInfTp = 100
        `);

        let rows = Array.isArray(result.recordset) ? result.recordset : [];
        if (machine) {
            rows = rows.filter(row => String(row.Maskine || '').toLowerCase().includes(machine));
        }

        const payload = {
            rows,
            count: rows.length,
            cached: true,
            source: 'db'
        };
        setCached(key, TTL.calculators, payload);
        return payload;
    }

    async function fetchProcessParameters() {
        const key = makeKey('process_params', 'all');
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT
                FreeInf1.FrInfTp,
                FreeInf1.Gr4,
                FreeInf1.Gr5,
                FreeInf1.Val1,
                FreeInf1.Val2,
                FreeInf1.Val3,
                FreeInf1.Val4,
                FreeInf1.Val5,
                FreeInf1.Val6,
                FreeInf1.Val7,
                FreeInf1.Val8,
                FreeInf1.Val9,
                FreeInf1.Val10,
                FreeInf1.Val11,
                FreeInf1.Val12,
                FreeInf1.Val13,
                FreeInf1.Val14,
                FreeInf1.Val15,
                FreeInf1.Val16,
                FreeInf1.Val17,
                FreeInf1.Val18
            FROM FreeInf1 WITH(NOLOCK)
            WHERE FreeInf1.FrInfTp = 61 OR FreeInf1.FrInfTp = 62
            ORDER BY FreeInf1.FrInfTp, FreeInf1.Gr4, FreeInf1.Gr5
        `);

        const rows = Array.isArray(result.recordset) ? result.recordset : [];
        const payload = {
            rows,
            count: rows.length,
            cached: true,
            source: 'db'
        };
        setCached(key, TTL.calculators, payload);
        return payload;
    }

    // ── Komp: komponentkatalog (workbook 'Komp' sheet, connection Visma4) ──
    async function fetchComponents(options = {}) {
        const q = String(options.q || '').trim().toLowerCase();
        const limit = Math.max(1, Math.min(5000, Number(options.limit || 500)));
        const key = makeKey('components', q + '_' + limit);
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT DISTINCT
                Prod.ProdNo,
                Prod.Descr,
                Prod.ProdTp,
                Prod.Gr5,
                Prod.Inf AS Pris,
                Txt.Txt AS Enhed,
                Prod.Inf3 AS KundeKode,
                Prod.Free2 AS Avance
            FROM Prod WITH(NOLOCK)
            JOIN Txt WITH(NOLOCK)
                ON Prod.StSaleUn = Txt.TxtNo
            WHERE Prod.Gr5 IN (2, 3, 6, 10, 11)
              AND Prod.ProdNo NOT LIKE '%L%'
              AND Prod.ProdNo NOT LIKE '%!%'
              AND Txt.TxtTp = 16
              AND Txt.Lang = 45
              AND Prod.ProdGr <> 99999
        `);

        let rows = Array.isArray(result.recordset) ? result.recordset : [];
        if (q) {
            rows = rows.filter(row =>
                [row.ProdNo, row.Descr, row.KundeKode]
                    .map(v => String(v || '').toLowerCase())
                    .join(' ')
                    .includes(q)
            );
        }
        rows = rows.slice(0, limit);

        const payload = { rows, count: rows.length, cached: true, source: 'db' };
        setCached(key, TTL.materials, payload);
        return payload;
    }

    // ── Kundenoter / BOM-instruktioner (ActInf InfTp=5, workbook RevStKun) ──
    async function fetchCustomerNotes(options = {}) {
        const customerCode = String(options.customerCode || '').trim();
        if (!customerCode) throw new Error('customerCode mangler');
        const key = makeKey('custnotes', customerCode);
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request()
            .input('gr', sql.VarChar, customerCode)
            .query(`
                SELECT Actor.Gr, ActInf.Txt1, ActInf.InfTp, ActInf.LnNo, ActInf.Qty1, Actor.CustNo
                FROM ActInf WITH(NOLOCK)
                JOIN Actor WITH(NOLOCK) ON Actor.ActNo = ActInf.ActNo
                WHERE ActInf.InfTp = 5 AND Actor.Gr = @gr
                ORDER BY ActInf.LnNo
            `);

        const rows = (Array.isArray(result.recordset) ? result.recordset : [])
            .filter(row => String(row.Txt1 || '').trim());
        const payload = { customerCode, rows, count: rows.length, cached: true, source: 'db' };
        setCached(key, TTL.revisions, payload);
        return payload;
    }

    // ── Leverandører (workbook 'Lev' sheet) ──
    async function fetchSuppliers(options = {}) {
        const q = String(options.q || '').trim().toLowerCase();
        const key = makeKey('suppliers', q || 'all');
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request().query(`
            SELECT Actor.SupNo, Actor.Nm, Actor.Ad1, Actor.Ad2, Actor.Ad3, Actor.Ad4, Actor.PArea
            FROM Actor WITH(NOLOCK)
            WHERE Actor.SupNo <> 0
            ORDER BY Actor.Nm
        `);

        let rows = Array.isArray(result.recordset) ? result.recordset : [];
        if (q) {
            rows = rows.filter(row =>
                [row.SupNo, row.Nm, row.PArea].map(v => String(v || '').toLowerCase()).join(' ').includes(q)
            );
        }
        const payload = { rows, count: rows.length, cached: true, source: 'db' };
        setCached(key, TTL.customers, payload);
        return payload;
    }

    // ── Produkt-træ: sottolivelli via ProdNo-mønster (far → -1, -2, V..., ...L) ──
    async function fetchProductTree(options = {}) {
        const prodNo = String(options.prodNo || '').trim();
        if (!prodNo) throw new Error('prodNo mangler');
        const baseNo = prodNo.replace(/L$/i, '').replace(/^V/i, '');
        const key = makeKey('prodtree', baseNo);
        const cached = getCached(key);
        if (cached) return cached;

        const pool = await getConnection();
        const result = await pool.request()
            .input('exact', sql.VarChar, baseNo)
            .input('lasered', sql.VarChar, baseNo + 'L')
            .input('route', sql.VarChar, 'V' + baseNo)
            .input('subPrefix', sql.VarChar, baseNo + '-%')
            .query(`
                SELECT Prod.ProdNo, Prod.Descr, Prod.ProdGr, Prod.Gr5,
                       Prod.Inf2 AS TgNo, Prod.Inf7 AS RevNo, Prod.Inf8 AS PosNo,
                       Prod.Inf AS Pris, Prod.Inf3 AS KundeKode
                FROM Prod WITH(NOLOCK)
                WHERE (Prod.ProdNo = @exact
                    OR Prod.ProdNo = @lasered
                    OR Prod.ProdNo = @route
                    OR Prod.ProdNo LIKE @subPrefix)
                  AND Prod.ProdGr <> 99999
                ORDER BY Prod.ProdNo
            `);

        const all = Array.isArray(result.recordset) ? result.recordset : [];
        const classify = (no) => {
            if (/^V/i.test(no)) return 'route';
            if (/L$/i.test(no) && no.replace(/L$/i, '') === baseNo) return 'laser';
            if (no === baseNo) return 'parent';
            return 'sublevel';
        };
        const rows = all.map(row => {
            const kind = classify(String(row.ProdNo || ''));
            const subMatch = String(row.ProdNo || '').match(/-(\d+)L?$/);
            return {
                ...row,
                Kind: kind,
                SubPos: subMatch ? Number(subMatch[1]) : null,
                IsLaserPart: /L$/i.test(String(row.ProdNo || ''))
            };
        });

        // gruppér sublevels: '...-1' og '...-1L' hører sammen
        const parent = rows.find(r => r.Kind === 'parent') || null;
        const route = rows.find(r => r.Kind === 'route') || null;
        const laser = rows.find(r => r.Kind === 'laser') || null;
        const subMap = new Map();
        rows.filter(r => r.Kind === 'sublevel').forEach(r => {
            const base = String(r.ProdNo).replace(/L$/i, '');
            if (!subMap.has(base)) subMap.set(base, { base, pos: r.SubPos, main: null, laser: null });
            const slot = subMap.get(base);
            if (r.IsLaserPart) slot.laser = r; else slot.main = r;
        });
        const sublevels = Array.from(subMap.values()).sort((a, b) => (a.pos || 0) - (b.pos || 0));

        const payload = {
            prodNo: baseNo,
            parent, route, laser, sublevels,
            flat: rows,
            count: rows.length,
            cached: true,
            source: 'db'
        };
        setCached(key, TTL.products, payload);
        return payload;
    }

    // ── Nesting: hvor mange emner passer på et pladeformat ──
    function computeNesting(input = {}) {
        const sheetW = Number(input.sheetWidth || 0);   // mm
        const sheetL = Number(input.sheetLength || 0);  // mm
        const pieceW = Number(input.pieceWidth || 0);   // mm
        const pieceL = Number(input.pieceLength || 0);  // mm
        const margin = Number(input.margin || 10);      // kant
        const gap = Number(input.gap || 5);             // afstand mellem emner
        if (sheetW <= 0 || sheetL <= 0 || pieceW <= 0 || pieceL <= 0) {
            throw new Error('sheetWidth, sheetLength, pieceWidth og pieceLength er paakraevet');
        }
        const usableW = sheetW - 2 * margin;
        const usableL = sheetL - 2 * margin;
        const pieceArea = pieceW * pieceL;
        const sheetArea = sheetW * sheetL;

        const fit = (uw, ul, pw, pl) => {
            if (pw > uw || pl > ul) return { cols: 0, rows: 0, count: 0 };
            const cols = Math.floor((uw + gap) / (pw + gap));
            const rows = Math.floor((ul + gap) / (pl + gap));
            return { cols, rows, count: cols * rows };
        };

        const usedSpan = (n, p) => (n > 0 ? n * p + (n - 1) * gap : 0);
        const normalizeRemnant = (w, l) => {
            const rw = Math.max(0, w);
            const rl = Math.max(0, l);
            return { width: Math.round(rw), length: Math.round(rl), areaMm2: Math.round(rw * rl) };
        };

        const makePlan = ({ label, rotation, pw, pl, mixed }) => {
            const base = fit(usableW, usableL, pw, pl);
            if (base.count <= 0) {
                return {
                    label, rotation, pw, pl,
                    cols: 0, rows: 0,
                    mixedExtra: 0,
                    total: 0,
                    fragmented: false,
                    reusableRemnant: normalizeRemnant(0, 0)
                };
            }

            let mixedExtra = 0;
            let fragmented = false;
            let usedL = usedSpan(base.rows, pl);
            const usedW = usedSpan(base.cols, pw);

            if (mixed) {
                const restL = usableL - usedL - gap;
                if (restL >= pw) {
                    const stripFit = fit(usableW, restL, pl, pw);
                    mixedExtra = stripFit.count;
                    if (mixedExtra > 0) fragmented = true;
                    const stripRows = stripFit.rows;
                    const stripUseL = usedSpan(stripRows, pw);
                    if (stripUseL > 0) {
                        usedL = Math.min(usableL, usedL + gap + stripUseL);
                    }
                }
            }

            const total = base.count + mixedExtra;

            // Preserved rectangular remnants on right or at top (sheet is filled from one side).
            const rightRem = normalizeRemnant(usableW - usedW, usableL);
            const topRem = normalizeRemnant(usableW, usableL - usedL);
            const reusableRemnant = rightRem.areaMm2 >= topRem.areaMm2 ? rightRem : topRem;

            return {
                label,
                rotation,
                pw,
                pl,
                cols: base.cols,
                rows: base.rows,
                mixedExtra,
                total,
                fragmented,
                reusableRemnant
            };
        };

        const candidates = [
            makePlan({ label: 'normal_compact', rotation: 0, pw: pieceW, pl: pieceL, mixed: false }),
            makePlan({ label: 'normal_mixed', rotation: 0, pw: pieceW, pl: pieceL, mixed: true }),
            makePlan({ label: 'rotated_compact', rotation: 90, pw: pieceL, pl: pieceW, mixed: false }),
            makePlan({ label: 'rotated_mixed', rotation: 90, pw: pieceL, pl: pieceW, mixed: true })
        ];

        const normal = candidates.find(c => c.label === 'normal_compact');
        const rotated = candidates.find(c => c.label === 'rotated_compact');

        const bestCount = Math.max(...candidates.map(c => c.total));
        // Rest-friendly: we can sacrifice up to 1 piece to preserve a better reusable rectangular remnant.
        const viable = candidates.filter(c => c.total >= Math.max(0, bestCount - 1));
        viable.sort((a, b) => {
            const fragA = a.fragmented ? 1 : 0;
            const fragB = b.fragmented ? 1 : 0;
            if (fragA !== fragB) return fragA - fragB;
            if (b.reusableRemnant.areaMm2 !== a.reusableRemnant.areaMm2) {
                return b.reusableRemnant.areaMm2 - a.reusableRemnant.areaMm2;
            }
            if (b.total !== a.total) return b.total - a.total;
            return (a.rotation || 0) - (b.rotation || 0);
        });
        const best = viable[0] || candidates[0];

        const total = best.total;
        const utilizationPct = sheetArea > 0 ? Math.round((total * pieceArea / sheetArea) * 1000) / 10 : 0;

        let restPreservationPct = 0;
        if (usableW > 0 && usableL > 0) {
            restPreservationPct = Math.round((best.reusableRemnant.areaMm2 / (usableW * usableL)) * 1000) / 10;
        }

        return {
            input: { sheetW, sheetL, pieceW, pieceL, margin, gap },
            normal: { cols: normal.cols, rows: normal.rows, count: normal.total },
            rotated: { cols: rotated.cols, rows: rotated.rows, count: rotated.total },
            best: {
                cols: best.cols,
                rows: best.rows,
                count: best.cols * best.rows,
                rotation: best.rotation,
                mixedExtra: best.mixedExtra,
                total: best.total,
                strategy: best.label,
                fragmented: best.fragmented,
                reusableRemnant: best.reusableRemnant,
                restPreservationPct
            },
            utilizationPct
        };
    }

    // ── Pris-beregner: materiale + laser + ressourcer (model fra Laserberegner) ──
    async function computeQuote(input = {}) {
        const materialProdNo = String(input.materialProdNo || '').trim();
        const pieceW = Number(input.pieceWidth || 0);     // mm
        const pieceL = Number(input.pieceLength || 0);    // mm
        const qty = Math.max(1, Number(input.qty || 1));
        const cutLengthM = Number(input.cutLengthM || 0); // skærelængde i meter pr emne
        const piercings = Math.max(0, Number(input.piercings || 1));
        const machine = String(input.machine || 'R1100').trim();
        const laserEnabled = input.laserEnabled !== false;
        const laserOpstartMinutes = Math.max(0, Number(input.laserOpstartMinutes || 0)); // opstart pr ordre
        const laserMinutesOverride = (input.laserMinutesOverride == null || input.laserMinutesOverride === '')
            ? null : Math.max(0, Number(input.laserMinutesOverride)); // sælger kan rette laser-tiden
        const priceBasis = String(input.priceBasis || 'sale').toLowerCase() === 'cost' ? 'cost' : 'sale'; // kundens pristype
        const opMinutes = Number(input.opMinutes || 0);   // bagudkompatibel: øvrige operationer i minutter pr emne
        const resourceRate = Number(input.resourceRate || 0); // dkk/min for øvrige operationer
        // operations: [{key, label, prodNo, minutes, rate}] — aktive processer som i Excel (Buk, Svejs, Flad, ...)
        const operations = Array.isArray(input.operations) ? input.operations : [];
        // components: [{prodNo, descr, qty, unitPrice}] — styklistelinjer (R8200) pr emne
        const components = Array.isArray(input.components) ? input.components : [];
        // minimum pr ordre (som professionelle tilbudssystemer)
        const minimumOrderAmount = Math.max(0, Number(input.minimumOrderAmount || 0)); // dkk
        const minimumQty = Math.max(0, Number(input.minimumQty || 0));                 // stk
        const margin = Number(input.margin || 10);
        const gap = Number(input.gap || 5);

        if (!materialProdNo) throw new Error('materialProdNo mangler');
        if (pieceW <= 0 || pieceL <= 0) throw new Error('pieceWidth og pieceLength er paakraevet');

        const pool = await getConnection();

        // 1) materiale-info
        const matResult = await pool.request()
            .input('prodNo', sql.VarChar, materialProdNo)
            .query(`
                SELECT TOP 1 Prod.ProdNo, Prod.Descr, Prod.HgtU AS Tykkelse,
                       Prod.WdtU AS BreddeM, Prod.LgtU AS LaengdeM,
                       Prod.Inf AS PrisKg, Prod.DensU AS Densitet,
                       Prod.R3 AS RV, Prod.Free2 AS Avance, Prod.NWgtU
                FROM Prod WITH(NOLOCK)
                WHERE Prod.ProdNo = @prodNo
            `);
        const mat = (matResult.recordset || [])[0];
        if (!mat) throw new Error('Materiale ' + materialProdNo + ' ikke fundet');

        const dbSheetW = Number(mat.BreddeM || 0) * 1000;  // m -> mm
        const dbSheetL = Number(mat.LaengdeM || 0) * 1000;
        const inputSheetW = Number(input.sheetWidth || 0);
        const inputSheetL = Number(input.sheetLength || 0);
        const hasCustomSheet = inputSheetW > 0 && inputSheetL > 0;
        const sheetW = hasCustomSheet ? inputSheetW : dbSheetW;
        const sheetL = hasCustomSheet ? inputSheetL : dbSheetL;
        const thickness = Number(mat.Tykkelse || 0);     // mm
        const density = Number(mat.Densitet || 7.9);     // kg/dm3
        const priceKg = Number(String(mat.PrisKg || '0').replace(',', '.')) || 0;
        const avancePct = Number(mat.Avance || 0);

        // 2) skæreparametre for materiale+maskine (FreeInf2 FrInfTp=100)
        const cutResult = await pool.request()
            .input('prodNo', sql.VarChar, materialProdNo)
            .query(`
                SELECT FreeInf2.Txt1 AS Maskine, FreeInf2.Val1 AS Skaerehast,
                       FreeInf2.Val2 AS Piercing, FreeInf2.Val3 AS Tillaeg, FreeInf2.Txt2 AS Linse
                FROM FreeInf2 WITH(NOLOCK)
                WHERE FreeInf2.ProdNo = @prodNo AND FreeInf2.FrInfTp = 100
            `);
        const cutRows = Array.isArray(cutResult.recordset) ? cutResult.recordset : [];
        const cutParam = cutRows.find(r => String(r.Maskine || '').toLowerCase().includes(machine.toLowerCase())) || cutRows[0] || null;

        // 3) nesting: ægte form-nesting hvis DXF-kontur er givet, ellers rektangulær
        const shapePolygon = Array.isArray(input.shapePolygon) && input.shapePolygon.length >= 3 ? input.shapePolygon : null;
        let nesting = null;
        if (sheetW > 0 && sheetL > 0) {
            if (shapePolygon) {
                try {
                    nesting = computeShapeNesting({ sheetWidth: sheetW, sheetLength: sheetL, polygon: shapePolygon, margin, gap });
                } catch (_) { nesting = null; }
            }
            if (!nesting) {
                try {
                    nesting = computeNesting({ sheetWidth: sheetW, sheetLength: sheetL, pieceWidth: pieceW, pieceLength: pieceL, margin, gap });
                } catch (_) { nesting = null; }
            }
        }

        // 4) materialepris pr emne: emnevægt = areal(dm2) * tykkelse(dm) * densitet(kg/dm3)
        // ved form-nesting bruges polygonets faktiske areal i stedet for w x l
        const pieceAreaDm2 = (nesting && nesting.mode === 'shape' && nesting.pieceAreaMm2 > 0)
            ? nesting.pieceAreaMm2 / 10000
            : (pieceW / 100) * (pieceL / 100);
        const pieceWeightKg = pieceAreaDm2 * (thickness / 100) * density;
        // pladeandel: hvis nesting kendt, brug pladepris/antal som alternativ
        const sheetAreaDm2 = (sheetW / 100) * (sheetL / 100);
        const sheetWeightKg = sheetAreaDm2 * (thickness / 100) * density;
        const sheetPrice = sheetWeightKg * priceKg;
        const piecesPerSheet = nesting && nesting.best.total > 0 ? nesting.best.total : null;
        const matCostByWeight = pieceWeightKg * priceKg;
        const matCostBySheetShare = piecesPerSheet ? sheetPrice / piecesPerSheet : null;
        const materialCost = matCostBySheetShare !== null ? Math.max(matCostByWeight, matCostBySheetShare) : matCostByWeight;
        const materialPrice = priceBasis === 'cost' ? materialCost : materialCost * (1 + avancePct / 100);

        // 5) laser-tid: skærelængde / hastighed + piercing
        // Enheder fra Visma: Skaerehast = m/min, Piercing = minutter pr piercing
        const cutSpeed = cutParam ? Number(cutParam.Skaerehast || 0) : 0;           // m/min
        const piercingMin = cutParam ? Number(cutParam.Piercing || 0) : 0;         // min pr piercing
        const tillaegPct = cutParam ? Number(cutParam.Tillaeg || 0) : 0;
        const cutMinutes = cutSpeed > 0 ? cutLengthM / cutSpeed : 0;
        const pierceMinutes = piercings * piercingMin;
        const autoLaserMinutes = (cutMinutes + pierceMinutes) * (1 + tillaegPct / 100);
        const laserMinutes = laserEnabled ? (laserMinutesOverride != null ? laserMinutesOverride : autoLaserMinutes) : 0;

        // 6) laser minutsats (FreeInf1 61/62 matrix har maskinsatser; fallback 12 dkk/min)
        const laserRate = Number(input.laserRate || 12);
        const laserCost = laserMinutes * laserRate;

        // 7) aktive processer (Buk, Svejs, Flad, Montage ...)
        const round2 = v => Math.round(v * 100) / 100;
        const operationLines = operations
            .map(op => {
                const minutes = Number(op.minutes || 0);
                const rate = Number(op.rate || 0);
                const opstartMinutes = Math.max(0, Number(op.opstartMinutes || 0));
                return {
                    key: String(op.key || '').trim(),
                    label: String(op.label || op.key || 'Operation').trim(),
                    prodNo: String(op.prodNo || '').trim(),
                    minutes,
                    rate,
                    cost: round2(minutes * rate),
                    opstartMinutes,
                    opstartCost: round2(opstartMinutes * rate)
                };
            })
            .filter(op => (op.minutes > 0 || op.opstartMinutes > 0) && op.rate >= 0);
        const operationsCost = operationLines.reduce((sum, op) => sum + op.cost, 0);
        const operationsMinutes = operationLines.reduce((sum, op) => sum + op.minutes, 0);

        // bagudkompatibel enkelt-felt
        const legacyResourceCost = opMinutes * resourceRate;
        const resourceCost = operationsCost + legacyResourceCost;
        const resourceMinutes = operationsMinutes + opMinutes;

        // styklistelinjer (R8200): komponenter pr emne, antal x stykpris
        const componentLines = components
            .map(c => {
                const cQty = Math.max(0, Number(c.qty || 0));
                const cPrice = Math.max(0, Number(c.unitPrice || 0));
                return {
                    prodNo: String(c.prodNo || '').trim(),
                    descr: String(c.descr || '').trim(),
                    qty: cQty,
                    unitPrice: round2(cPrice),
                    lineCost: round2(cQty * cPrice)
                };
            })
            .filter(c => c.qty > 0 && c.prodNo);
        const componentsCost = componentLines.reduce((sum, c) => sum + c.lineCost, 0);

        // opstart: engangsomkostning pr ordre (laseropstart + opstart pr proces), fordeles paa antal
        const laserOpstartCost = laserEnabled ? laserOpstartMinutes * laserRate : 0;
        const operationsOpstartCost = operationLines.reduce((sum, op) => sum + op.opstartCost, 0);
        const operationsOpstartMinutes = operationLines.reduce((sum, op) => sum + op.opstartMinutes, 0);
        const opstartCost = laserOpstartCost + operationsOpstartCost;
        const opstartShare = opstartCost / qty;

        const unitPrice = materialPrice + laserCost + resourceCost + componentsCost + opstartShare;

        // minimum pr ordre: anvend største af beregnet total / minimumsbeløb; minimumsantal hæver kun prisen pr ordre
        const effQty = minimumQty > 0 ? Math.max(qty, minimumQty) : qty;
        const rawTotal = unitPrice * qty;
        const qtyAdjustedTotal = effQty > qty ? (materialPrice + laserCost + resourceCost + componentsCost) * effQty + opstartCost : rawTotal;
        const totalPrice = Math.max(qtyAdjustedTotal, minimumOrderAmount);
        const minimumApplied = totalPrice > rawTotal + 0.005;

        return {
            material: {
                prodNo: mat.ProdNo, descr: mat.Descr, thickness, density, priceKg,
                sheetW, sheetL, sheetWeightKg: round2(sheetWeightKg), sheetPrice: round2(sheetPrice),
                sheetSource: hasCustomSheet ? 'custom' : 'material',
                avancePct, priceBasis
            },
            cutParam: cutParam ? {
                maskine: cutParam.Maskine, skaerehast: cutSpeed, piercingMin, tillaegPct, linse: cutParam.Linse
            } : null,
            nesting,
            operations: operationLines,
            components: componentLines,
            perPiece: {
                weightKg: round2(pieceWeightKg),
                materialCost: round2(materialCost),
                materialPrice: round2(materialPrice),
                laserMinutes: round2(laserMinutes),
                autoLaserMinutes: round2(autoLaserMinutes),
                laserMinutesOverridden: laserMinutesOverride != null,
                laserCost: round2(laserCost),
                resourceMinutes: round2(resourceMinutes),
                resourceCost: round2(resourceCost),
                componentsCost: round2(componentsCost),
                opstartShare: round2(opstartShare),
                unitPrice: round2(unitPrice)
            },
            perOrder: {
                laserOpstartMinutes,
                laserOpstartCost: round2(laserOpstartCost),
                operationsOpstartMinutes,
                operationsOpstartCost: round2(operationsOpstartCost),
                opstartCost: round2(opstartCost)
            },
            total: {
                qty,
                effectiveQty: effQty,
                totalPrice: round2(totalPrice),
                rawTotal: round2(rawTotal),
                minimumApplied,
                minimumOrderAmount,
                minimumQty,
                sheetsNeeded: piecesPerSheet ? Math.ceil(qty / piecesPerSheet) : null
            }
        };
    }

    // ── Fil-analyse: DXF / STEP bounding box + skærelængde ──
    function analyzeDrawingFile(filename, buffer) {
        const ext = String(filename || '').toLowerCase().split('.').pop();
        if (ext === 'dxf') return analyzeDxf(buffer.toString('latin1'));
        if (ext === 'step' || ext === 'stp') return analyzeStep(buffer.toString('latin1'));
        if (ext === 'pdf') return analyzePdf(buffer);
        throw new Error('Format .' + ext + ' understoettes ikke (brug dxf, step, stp eller pdf)');
    }

    function analyzeDxf(text) {
        const lines = text.split(/\r?\n/);
        let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        let cutLength = 0;
        let entityCount = 0;
        let circles = 0;
        const segs = []; // alle linjestykker (buer/cirkler approximeret) til form-nesting

        const addPt = (x, y) => {
            if (!isFinite(x) || !isFinite(y)) return;
            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
        };
        const addSeg = (x1, y1, x2, y2) => {
            if (![x1, y1, x2, y2].every(isFinite)) return;
            segs.push([x1, y1, x2, y2]);
        };

        let i = 0;
        const readPair = () => {
            const code = (lines[i] || '').trim();
            const value = (lines[i + 1] || '').trim();
            i += 2;
            return { code, value };
        };

        let current = null;
        let data = {};
        let polyPts = [];
        let polyClosed = false;

        const flush = () => {
            if (!current) return;
            entityCount += 1;
            if (current === 'LINE') {
                const x1 = Number(data['10']), y1 = Number(data['20']);
                const x2 = Number(data['11']), y2 = Number(data['21']);
                addPt(x1, y1); addPt(x2, y2);
                if ([x1, y1, x2, y2].every(isFinite)) { cutLength += Math.hypot(x2 - x1, y2 - y1); addSeg(x1, y1, x2, y2); }
            } else if (current === 'CIRCLE') {
                const cx = Number(data['10']), cy = Number(data['20']), r = Number(data['40']);
                if ([cx, cy, r].every(isFinite)) {
                    addPt(cx - r, cy - r); addPt(cx + r, cy + r);
                    cutLength += 2 * Math.PI * r;
                    circles += 1;
                    const n = 24;
                    for (let k = 0; k < n; k += 1) {
                        const a1 = (k / n) * 2 * Math.PI, a2 = ((k + 1) / n) * 2 * Math.PI;
                        addSeg(cx + r * Math.cos(a1), cy + r * Math.sin(a1), cx + r * Math.cos(a2), cy + r * Math.sin(a2));
                    }
                }
            } else if (current === 'ARC') {
                const cx = Number(data['10']), cy = Number(data['20']), r = Number(data['40']);
                const a1 = Number(data['50']) * Math.PI / 180, a2 = Number(data['51']) * Math.PI / 180;
                if ([cx, cy, r, a1, a2].every(isFinite)) {
                    let sweep = a2 - a1;
                    if (sweep <= 0) sweep += 2 * Math.PI;
                    cutLength += r * sweep;
                    addPt(cx - r, cy - r); addPt(cx + r, cy + r);
                    const n = Math.max(2, Math.ceil(sweep / 0.25));
                    for (let k = 0; k < n; k += 1) {
                        const t1 = a1 + sweep * (k / n), t2 = a1 + sweep * ((k + 1) / n);
                        addSeg(cx + r * Math.cos(t1), cy + r * Math.sin(t1), cx + r * Math.cos(t2), cy + r * Math.sin(t2));
                    }
                }
            } else if (current === 'LWPOLYLINE' || current === 'POLYLINE') {
                for (let p = 0; p < polyPts.length; p += 1) {
                    addPt(polyPts[p][0], polyPts[p][1]);
                    if (p > 0) {
                        cutLength += Math.hypot(polyPts[p][0] - polyPts[p - 1][0], polyPts[p][1] - polyPts[p - 1][1]);
                        addSeg(polyPts[p - 1][0], polyPts[p - 1][1], polyPts[p][0], polyPts[p][1]);
                    }
                }
                if (polyClosed && polyPts.length > 2) {
                    const first = polyPts[0], last = polyPts[polyPts.length - 1];
                    cutLength += Math.hypot(first[0] - last[0], first[1] - last[1]);
                    addSeg(last[0], last[1], first[0], first[1]);
                }
            }
            current = null;
            data = {};
            polyPts = [];
            polyClosed = false;
        };

        let pendingX = null;
        while (i < lines.length - 1) {
            const { code, value } = readPair();
            if (code === '0') {
                flush();
                if (['LINE', 'CIRCLE', 'ARC', 'LWPOLYLINE', 'POLYLINE', 'VERTEX'].includes(value)) {
                    if (value === 'VERTEX' && current === 'POLYLINE') {
                        // vertex hører til aktiv polyline
                    } else if (value !== 'VERTEX') {
                        current = value;
                    }
                }
                continue;
            }
            if (!current) continue;
            if (current === 'LWPOLYLINE' || current === 'POLYLINE') {
                if (code === '70') polyClosed = (Number(value) & 1) === 1;
                if (code === '10') pendingX = Number(value);
                if (code === '20' && pendingX !== null) {
                    polyPts.push([pendingX, Number(value)]);
                    pendingX = null;
                }
            } else {
                data[code] = value;
            }
        }
        flush();

        if (!isFinite(minX)) throw new Error('Ingen geometri fundet i DXF');
        const width = Math.round((maxX - minX) * 100) / 100;
        const length = Math.round((maxY - minY) * 100) / 100;
        // udtræk yderkontur (største lukkede loop) til form-nesting
        let polygon = null;
        let polygonAreaMm2 = null;
        try {
            const outline = extractOutline(segs);
            if (outline && outline.points.length >= 3) {
                let pts = outline.points.map(p => [Math.round((p[0] - minX) * 100) / 100, Math.round((p[1] - minY) * 100) / 100]);
                if (pts.length > 400) {
                    const step = Math.ceil(pts.length / 400);
                    pts = pts.filter((_, idx) => idx % step === 0);
                }
                polygon = pts;
                polygonAreaMm2 = Math.round(outline.area);
            }
        } catch (_) { /* form-nesting valgfri */ }
        return {
            format: 'dxf',
            widthMm: Math.min(width, length),
            lengthMm: Math.max(width, length),
            cutLengthM: Math.round(cutLength / 10) / 100,
            piercingsEstimate: Math.max(1, circles + 1),
            entityCount,
            boundingBox: { minX, minY, maxX, maxY },
            polygon,
            polygonAreaMm2
        };
    }

    // ── Geometri-hjælpere til form-nesting ──
    function polygonArea(pts) {
        let a = 0;
        for (let i = 0; i < pts.length; i += 1) {
            const p = pts[i], q = pts[(i + 1) % pts.length];
            a += p[0] * q[1] - q[0] * p[1];
        }
        return Math.abs(a) / 2;
    }
    function pointInPolygon(x, y, pts) {
        let inside = false;
        for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
            const xi = pts[i][0], yi = pts[i][1], xj = pts[j][0], yj = pts[j][1];
            if (((yi > y) !== (yj > y)) && (x < (xj - xi) * (y - yi) / (yj - yi) + xi)) inside = !inside;
        }
        return inside;
    }
    function extractOutline(segs) {
        // kæd segmenter sammen til lukkede loops (0.5 mm tolerance); største loop = yderkontur
        const key = (x, y) => Math.round(x * 2) + ',' + Math.round(y * 2);
        const adj = new Map();
        segs.forEach((s, idx) => {
            const k1 = key(s[0], s[1]), k2 = key(s[2], s[3]);
            if (k1 === k2) return;
            if (!adj.has(k1)) adj.set(k1, []);
            if (!adj.has(k2)) adj.set(k2, []);
            adj.get(k1).push({ idx, pt: [s[2], s[3]] });
            adj.get(k2).push({ idx, pt: [s[0], s[1]] });
        });
        const used = new Set();
        let best = null, bestArea = 0;
        segs.forEach((s, idx) => {
            if (used.has(idx)) return;
            const k1 = key(s[0], s[1]), k2 = key(s[2], s[3]);
            if (k1 === k2) return;
            used.add(idx);
            const pts = [[s[0], s[1]], [s[2], s[3]]];
            let curKey = k2;
            let guard = 0;
            while (curKey !== k1 && guard++ < 20000) {
                const nexts = (adj.get(curKey) || []).filter(e => !used.has(e.idx));
                if (!nexts.length) break;
                const e = nexts[0];
                used.add(e.idx);
                pts.push(e.pt);
                curKey = key(e.pt[0], e.pt[1]);
            }
            if (curKey === k1 && pts.length >= 3) {
                const a = polygonArea(pts);
                if (a > bestArea) { bestArea = a; best = pts; }
            }
        });
        return best ? { points: best, area: bestArea } : null;
    }

    // ── Form-nesting: raster-baseret greedy bottom-left med rotationer (0/90/180/270) ──
    // Respekterer margin (afstand til pladekant) og gap (afstand mellem emner) som i Laser.
    function computeShapeNesting(input = {}) {
        const sheetW = Number(input.sheetWidth || 0);   // mm (bredde, y)
        const sheetL = Number(input.sheetLength || 0);  // mm (længde, x)
        const margin = Number(input.margin == null ? 10 : input.margin);
        const gap = Number(input.gap == null ? 5 : input.gap);
        const poly = Array.isArray(input.polygon) ? input.polygon : null;
        if (!poly || poly.length < 3) throw new Error('polygon paakraevet');
        if (sheetW <= 0 || sheetL <= 0) throw new Error('sheetWidth og sheetLength er paakraevet');
        // normaliser polygon til origo
        let pminX = Infinity, pminY = Infinity, pmaxX = -Infinity, pmaxY = -Infinity;
        poly.forEach(p => {
            if (p[0] < pminX) pminX = p[0];
            if (p[0] > pmaxX) pmaxX = p[0];
            if (p[1] < pminY) pminY = p[1];
            if (p[1] > pmaxY) pmaxY = p[1];
        });
        const base = poly.map(p => [p[0] - pminX, p[1] - pminY]);
        const pw = pmaxX - pminX, ph = pmaxY - pminY;
        if (pw <= 0 || ph <= 0) throw new Error('polygon har ingen udstrækning');
        const pieceAreaMm2 = polygonArea(base);
        const cell = Math.max(1, Math.min(5, Math.floor(Math.min(pw, ph) / 40) || 1));
        const rotPt = (p, rot) => {
            if (rot === 90) return [ph - p[1], p[0]];
            if (rot === 180) return [pw - p[0], ph - p[1]];
            if (rot === 270) return [p[1], pw - p[0]];
            return [p[0], p[1]];
        };
        const gapCells = Math.max(1, Math.ceil(gap / cell));
        const buildMask = rot => {
            const rp = base.map(p => rotPt(p, rot));
            const rw = (rot === 90 || rot === 270) ? ph : pw;
            const rh = (rot === 90 || rot === 270) ? pw : ph;
            const mw = Math.ceil(rw / cell), mh = Math.ceil(rh / cell);
            const grid = new Uint8Array(mw * mh);
            for (let gy = 0; gy < mh; gy += 1) {
                for (let gx = 0; gx < mw; gx += 1) {
                    if (pointInPolygon((gx + 0.5) * cell, (gy + 0.5) * cell, rp)) grid[gy * mw + gx] = 1;
                }
            }
            // dilater 1 celle så konturen er dækket
            const dil = new Uint8Array(mw * mh);
            for (let gy = 0; gy < mh; gy += 1) {
                for (let gx = 0; gx < mw; gx += 1) {
                    if (!grid[gy * mw + gx]) continue;
                    for (let dy = -1; dy <= 1; dy += 1) {
                        for (let dx = -1; dx <= 1; dx += 1) {
                            const ny = gy + dy, nx = gx + dx;
                            if (ny >= 0 && ny < mh && nx >= 0 && nx < mw) dil[ny * mw + nx] = 1;
                        }
                    }
                }
            }
            const cells = [];        // footprint: skal være fri før placering
            const inflated = [];     // footprint + gap-halo: stemples som optaget
            for (let gy = 0; gy < mh; gy += 1) {
                for (let gx = 0; gx < mw; gx += 1) {
                    if (dil[gy * mw + gx]) cells.push([gx, gy]);
                }
            }
            const inflSet = new Set();
            cells.forEach(([gx, gy]) => {
                for (let dy = -gapCells; dy <= gapCells; dy += 1) {
                    for (let dx = -gapCells; dx <= gapCells; dx += 1) {
                        inflSet.add((gx + dx) + ':' + (gy + dy));
                    }
                }
            });
            inflSet.forEach(k => {
                const [gx, gy] = k.split(':').map(Number);
                inflated.push([gx, gy]);
            });
            return { rot, mw, mh, cells, inflated };
        };
        const masks = [0, 90, 180, 270].map(buildMask);
        const gw = Math.ceil(sheetL / cell), gh = Math.ceil(sheetW / cell);
        const sheet = new Uint8Array(gw * gh);
        const mCells = Math.ceil(margin / cell);
        for (let gy = 0; gy < gh; gy += 1) {
            for (let gx = 0; gx < gw; gx += 1) {
                if (gx < mCells || gy < mCells || gx >= gw - mCells || gy >= gh - mCells) sheet[gy * gw + gx] = 1;
            }
        }
        const placements = [];
        const maxPieces = 1000;
        for (let y = 0; y < gh && placements.length < maxPieces; y += 1) {
            for (let x = 0; x < gw && placements.length < maxPieces; x += 1) {
                for (const mask of masks) {
                    if (x + mask.mw > gw || y + mask.mh > gh) continue;
                    let fits = true;
                    for (let ci = 0; ci < mask.cells.length; ci += 1) {
                        const cx2 = x + mask.cells[ci][0], cy2 = y + mask.cells[ci][1];
                        if (sheet[cy2 * gw + cx2]) { fits = false; break; }
                    }
                    if (!fits) continue;
                    for (let ci = 0; ci < mask.inflated.length; ci += 1) {
                        const cx2 = x + mask.inflated[ci][0], cy2 = y + mask.inflated[ci][1];
                        if (cx2 >= 0 && cx2 < gw && cy2 >= 0 && cy2 < gh) sheet[cy2 * gw + cx2] = 1;
                    }
                    placements.push({ x: Math.round(x * cell), y: Math.round(y * cell), rot: mask.rot });
                    break;
                }
            }
        }
        const total = placements.length;
        return {
            mode: 'shape',
            input: { sheetW, sheetL, pieceW: Math.round(pw * 100) / 100, pieceL: Math.round(ph * 100) / 100, margin, gap },
            cell,
            polygon: base.map(p => [Math.round(p[0] * 100) / 100, Math.round(p[1] * 100) / 100]),
            pieceAreaMm2: Math.round(pieceAreaMm2),
            placements,
            best: { total, rotationsUsed: Array.from(new Set(placements.map(p => p.rot))) },
            utilizationPct: (sheetW * sheetL) > 0 ? Math.round((total * pieceAreaMm2 / (sheetW * sheetL)) * 1000) / 10 : 0
        };
    }

    function analyzeStep(text) {
        const re = /CARTESIAN_POINT\s*\(\s*'[^']*'\s*,\s*\(\s*([-\d.eE+]+)\s*,\s*([-\d.eE+]+)\s*(?:,\s*([-\d.eE+]+)\s*)?\)/g;
        let m;
        let minX = Infinity, minY = Infinity, minZ = Infinity;
        let maxX = -Infinity, maxY = -Infinity, maxZ = -Infinity;
        let count = 0;
        while ((m = re.exec(text)) !== null) {
            const x = Number(m[1]), y = Number(m[2]), z = Number(m[3] || 0);
            if (!isFinite(x) || !isFinite(y)) continue;
            count += 1;
            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
            if (z < minZ) minZ = z;
            if (z > maxZ) maxZ = z;
        }
        if (count === 0) throw new Error('Ingen CARTESIAN_POINT fundet i STEP-fil');
        const dims = [maxX - minX, maxY - minY, maxZ - minZ].map(v => Math.round(v * 100) / 100).sort((a, b) => a - b);
        return {
            format: 'step',
            thicknessMm: dims[0],
            widthMm: dims[1],
            lengthMm: dims[2],
            cutLengthM: null,
            piercingsEstimate: null,
            pointCount: count
        };
    }

    function analyzePdf(buffer) {
        // Best effort: udtræk tekst og find dimensioner som '123,5 x 456' eller '123.5x456'
        const text = buffer.toString('latin1');
        const chunks = [];
        const streamRe = /\(([^)]{2,})\)\s*Tj/g;
        let m;
        while ((m = streamRe.exec(text)) !== null) chunks.push(m[1]);
        const joined = chunks.join(' ');
        const dimRe = /(\d{1,5}(?:[.,]\d{1,2})?)\s*[xX×]\s*(\d{1,5}(?:[.,]\d{1,2})?)/g;
        const found = [];
        while ((m = dimRe.exec(joined)) !== null) {
            const a = Number(m[1].replace(',', '.'));
            const b = Number(m[2].replace(',', '.'));
            if (a > 1 && b > 1 && a < 20000 && b < 20000) found.push([a, b]);
        }
        if (found.length === 0) {
            return {
                format: 'pdf',
                widthMm: null,
                lengthMm: null,
                cutLengthM: null,
                piercingsEstimate: null,
                note: 'Kunne ikke finde maal automatisk i PDF - indtast manuelt'
            };
        }
        const biggest = found.sort((p, q) => (q[0] * q[1]) - (p[0] * p[1]))[0];
        return {
            format: 'pdf',
            widthMm: Math.min(biggest[0], biggest[1]),
            lengthMm: Math.max(biggest[0], biggest[1]),
            cutLengthM: null,
            piercingsEstimate: null,
            candidates: found.slice(0, 8),
            note: 'Maal udtrukket fra PDF-tekst - kontroller altid mod tegning'
        };
    }

    function invalidate(scope = 'all') {
        const normalized = String(scope || 'all').trim().toLowerCase();
        const prefixes = {
            all: ['bom_v1_'],
            customers: ['bom_v1_customers_'],
            products: ['bom_v1_products_'],
            revisions: ['bom_v1_revisions_'],
            resources: ['bom_v1_resources_'],
            materials: ['bom_v1_materials_'],
            components: ['bom_v1_components_'],
            calculators: ['bom_v1_laser_params_', 'bom_v1_process_params_']
        };
        const selected = prefixes[normalized] || prefixes.all;

        let clearedMemory = 0;
        for (const key of Array.from(memoryCache.keys())) {
            if (selected.some(prefix => key.startsWith(prefix))) {
                memoryCache.delete(key);
                clearedMemory += 1;
            }
        }

        let clearedDisk = 0;
        const files = diskCache.list();
        for (const entry of files) {
            const key = String(entry.key || '');
            if (selected.some(prefix => key.startsWith(prefix))) {
                diskCache.del(key);
                clearedDisk += 1;
            }
        }

        if (logEvent) {
            logEvent('BOM CACHE INVALIDATE scope=' + normalized + ' memory=' + clearedMemory + ' disk=' + clearedDisk);
        }

        return {
            scope: normalized,
            memory: clearedMemory,
            disk: clearedDisk
        };
    }

    // ── TgForm kode → Gr8 int (som Enh-ark: A4=1, A3=2, A2=3, A1=4, A0=5) ──
    const TG_FORM_MAP = { 'A4': 1, 'A3': 2, 'A2': 3, 'A1': 4, 'A0': 5, '-': 6 };
    function tgFormToGr8(tgForm) {
        return TG_FORM_MAP[String(tgForm || 'A4').trim().toUpperCase()] || 1;
    }

    // ── Byg liste af Prod-records ud fra input (som Vareoplysninger) ──
    function buildProductRecords(input) {
        const custCode  = String(input.customerCode  || input.customerNo || '').trim();
        const suffix    = String(input.prodNoSuffix  || '').trim();
        const descr     = String(input.descr || '').slice(0, 60);
        const tgNo      = String(input.tgNo  || '').trim();
        const revNo     = String(input.revNo || '').trim();
        const tgForm    = String(input.tgForm || 'A4').trim();
        const custNoAlt = String(input.customerNoAlt || '').trim();
        const lager     = String(input.lager || '').trim();
        const creDate   = new Date().toISOString().slice(0, 10).replace(/-/g, '');

        // Dimensioner (mm → meter for Visma-formatet på færdigvarer)
        const hgtU   = Number(input.thickness || 0);           // tykkelse mm
        const lgtU   = Number(input.lengthMm  || 0) / 1000;   // mm → m
        const wdtU   = Number(input.widthMm   || 0) / 1000;   // mm → m
        const densU  = Number(input.density   || 7.85);
        const free2  = Number(input.margin    || 0);
        const stSaleUn = Number(input.stSaleUn || 1);
        const prodPrGr = Number(input.prodPrGr || 0);
        const prCatNo  = Number(input.prCatNo  || 1);
        const gr8      = tgFormToGr8(tgForm);

        const mainProdNo  = custCode + suffix;
        const routeProdNo = 'V' + mainProdNo;
        const laserProdNo = mainProdNo + 'L';

        const base = {
            Inf2: tgNo, Inf3: custCode, Inf4: custNoAlt, Inf6: lager,
            Inf7: revNo, Inf8: tgForm, HgtU: hgtU, LgtU: lgtU, WdtU: wdtU,
            DensU: densU, Inf: 0, Free2: free2, StSaleUn: stSaleUn,
            ProdPrGr: prodPrGr, PrCatNo: prCatNo, Gr8: gr8,
            NWgtU: 0, CreDt: creDate, Rsp: 0
        };

        const records = [];

        // R2: Hovedvare (ProdGr=1, assembly/salgsartikel)
        if (mainProdNo) {
            records.push({ ...base, ProdNo: mainProdNo, Descr: descr, ProdGr: 1, _role: 'main' });
        }
        // R3: Rute (ProdGr=1, ProdNo=V+main, beskrivelse = "Rute for ...")
        if (input.createRoute && mainProdNo) {
            records.push({ ...base, ProdNo: routeProdNo, Descr: ('Rute for ' + descr).slice(0, 60), ProdGr: 1, _role: 'route' });
        }
        // R4: Laserpart (ProdGr=2, ProdNo=main+L)
        if (input.createLaserPart && mainProdNo) {
            const laserDescr = 'Laser ' + descr;
            records.push({ ...base, ProdNo: laserProdNo, Descr: laserDescr.slice(0, 60), ProdGr: 2, _role: 'laser' });
        }

        return records;
    }

    // ── Anteprima: hvad vil blive oprettet, tjek duplikater ──
    async function previewCreateProducts(input) {
        const records = buildProductRecords(input);
        if (records.length === 0) throw new Error('Ingen produkter at oprette — udfyld kundenr. og produktnr.-suffiks');

        const pool = await getConnection();
        const prodNos = records.map(r => r.ProdNo).filter(Boolean);
        const placeholders = prodNos.map((_, i) => '@pn' + i).join(', ');
        const req = pool.request();
        prodNos.forEach((pn, i) => req.input('pn' + i, sql.VarChar, pn));
        const existing = await req.query(`SELECT ProdNo, Descr FROM Prod WHERE ProdNo IN (${placeholders})`);
        const conflicts = (existing.recordset || []).map(r => String(r.ProdNo));

        return {
            records: records.map(r => ({ ...r })),
            conflicts,
            canCreate: conflicts.length === 0
        };
    }

    // ── Opret produkter i Visma via transaktion ──
    async function createProductsInVisma(input) {
        const preview = await previewCreateProducts(input);
        if (preview.conflicts.length > 0) {
            const err = new Error('Produkterne eksisterer allerede i Visma: ' + preview.conflicts.join(', '));
            err.statusCode = 409;
            throw err;
        }

        const pool = await getConnection();
        const transaction = new sql.Transaction(pool);
        await transaction.begin();
        const created = [];
        try {
            for (const rec of preview.records) {
                const req = new sql.Request(transaction);
                req.input('ProdNo',    sql.VarChar(50),  rec.ProdNo);
                req.input('Descr',     sql.VarChar(60),  rec.Descr);
                req.input('ProdGr',    sql.Int,           rec.ProdGr);
                req.input('Inf2',      sql.VarChar(60),  rec.Inf2);
                req.input('Inf3',      sql.VarChar(60),  rec.Inf3);
                req.input('Inf4',      sql.VarChar(60),  rec.Inf4);
                req.input('Inf6',      sql.VarChar(20),  rec.Inf6);
                req.input('Inf7',      sql.VarChar(20),  rec.Inf7);
                req.input('Inf8',      sql.VarChar(20),  rec.Inf8);
                req.input('HgtU',      sql.Float,         rec.HgtU);
                req.input('LgtU',      sql.Float,         rec.LgtU);
                req.input('WdtU',      sql.Float,         rec.WdtU);
                req.input('DensU',     sql.Float,         rec.DensU);
                req.input('Inf',       sql.Float,         rec.Inf);
                req.input('Free2',     sql.Float,         rec.Free2);
                req.input('StSaleUn',  sql.Int,           rec.StSaleUn);
                req.input('ProdPrGr',  sql.Int,           rec.ProdPrGr);
                req.input('PrCatNo',   sql.Int,           rec.PrCatNo);
                req.input('Gr8',       sql.Int,           rec.Gr8);
                req.input('NWgtU',     sql.Float,         rec.NWgtU);
                req.input('CreDt',     sql.VarChar(8),   rec.CreDt);
                req.input('Rsp',       sql.Float,         rec.Rsp);
                await req.query(`
                    INSERT INTO Prod
                        (ProdNo, Descr, ProdGr, Inf2, Inf3, Inf4, Inf6, Inf7, Inf8,
                         HgtU, LgtU, WdtU, DensU, Inf, Free2, StSaleUn,
                         ProdPrGr, PrCatNo, Gr8, NWgtU, CreDt, Rsp)
                    VALUES
                        (@ProdNo, @Descr, @ProdGr, @Inf2, @Inf3, @Inf4, @Inf6, @Inf7, @Inf8,
                         @HgtU, @LgtU, @WdtU, @DensU, @Inf, @Free2, @StSaleUn,
                         @ProdPrGr, @PrCatNo, @Gr8, @NWgtU, @CreDt, @Rsp)
                `);
                created.push({ ProdNo: rec.ProdNo, Descr: rec.Descr, role: rec._role });
            }
            await transaction.commit();
            // Invalider relevante caches
            invalidate('products');
            return { ok: true, created };
        } catch (err) {
            await transaction.rollback();
            throw err;
        }
    }

    return {
        fetchCustomers,
        fetchProductsByCustomer,
        fetchRevisionsByDrawing,
        fetchResources,
        fetchMaterials,
        fetchLaserParameters,
        fetchProcessParameters,
        fetchComponents,
        fetchCustomerNotes,
        fetchSuppliers,
        fetchProductTree,
        computeNesting,
        computeShapeNesting,
        computeQuote,
        analyzeDrawingFile,
        previewCreateProducts,
        createProductsInVisma,
        invalidate
    };
}

module.exports = {
    createBomService
};
