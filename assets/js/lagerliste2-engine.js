(function (root, factory) {
    const api = factory();
    if (typeof module === 'object' && module.exports) module.exports = api;
    else root.Lagerliste2Engine = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    const EPSILON = 0.005;

    function number(value) {
        const parsed = Number(value || 0);
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function round(value) {
        const result = Math.round(number(value) * 100) / 100;
        return Object.is(result, -0) ? 0 : result;
    }

    function unwrapPayload(value) {
        if (!value || typeof value !== 'object') return { categories: {}, totals: {} };
        if (value.current && value.current.categories) return value.current;
        if (value.snapshot) return unwrapPayload(value.snapshot);
        return value.categories ? value : { categories: {}, totals: {} };
    }

    function parseProducts(value) {
        return String(value || '').split(',').map(item => item.trim()).filter(Boolean);
    }

    function flatPlateRows(categories) {
        if (Array.isArray(categories.plates)) return categories.plates;
        return (categories.plateGroups || []).flatMap(group => group.details || []);
    }

    function flatRestRows(categories) {
        if (Array.isArray(categories.restPlates)) return categories.restPlates;
        return (categories.restPlateGroups || []).flatMap(group => group.details || []);
    }

    function orderAllocationSummary(payloadA, payloadB, context = {}) {
        const states = Array.isArray(context && context.evidence && context.evidence.orderStates)
            ? context.evidence.orderStates : [];
        const summarizeSide = payload => {
            const categories = unwrapPayload(payload).categories || {};
            return {
                via: new Set((categories.salgordreVia || []).map(row => number(row.OrdNo)).filter(Boolean)),
                finished: new Set((categories.finishedNotInvoiced || []).map(row => number(row.OrdNo)).filter(Boolean))
            };
        };
        const sideA = summarizeSide(payloadA);
        const sideB = summarizeSide(payloadB);
        return states.filter(state => {
            const orderNo = number(state.orderNo || state.OrdNo);
            return (sideA.via.has(orderNo) && sideA.finished.has(orderNo))
                || (sideB.via.has(orderNo) && sideB.finished.has(orderNo))
                || (number(state.packedRatio) < 0.999 && (sideA.finished.has(orderNo) || sideB.finished.has(orderNo)));
        }).map(state => ({
            orderNo: number(state.orderNo || state.OrdNo),
            packedRatio: number(state.packedRatio),
            partiallyPacked: Boolean(state.partiallyPacked),
            headerClosed: Boolean(state.headerClosed),
            invoiced: Boolean(state.invoiced),
            duplicateBefore: sideA.via.has(number(state.orderNo || state.OrdNo)) && sideA.finished.has(number(state.orderNo || state.OrdNo)),
            duplicateAfter: sideB.via.has(number(state.orderNo || state.OrdNo)) && sideB.finished.has(number(state.orderNo || state.OrdNo))
        }));
    }

    function movementSpecs(payload, context = {}) {
        const categories = unwrapPayload(payload).categories || {};
        const orderStates = new Map((Array.isArray(context && context.evidence && context.evidence.orderStates)
            ? context.evidence.orderStates : []).map(row => [number(row.orderNo || row.OrdNo), row]));
        const viaOrders = new Set((categories.salgordreVia || []).map(row => number(row.OrdNo)).filter(Boolean));
        const finishedOrders = new Set((categories.finishedNotInvoiced || []).map(row => number(row.OrdNo)).filter(Boolean));
        const allocation = (row, category) => {
            const orderNo = number(row.OrdNo);
            const state = orderStates.get(orderNo);
            if (!state) return 1;
            const packedRatio = Math.max(0, Math.min(1, number(state.packedRatio)));
            const hasVia = viaOrders.has(orderNo);
            const hasFinished = finishedOrders.has(orderNo);
            if (category === 'via') {
                if (hasVia && hasFinished) return round(1 - packedRatio);
                return 1;
            }
            if (category === 'finished') {
                if (hasVia && hasFinished) return packedRatio;
                if (!hasVia && hasFinished) return packedRatio;
            }
            return 1;
        };
        const unpackedFinishedRows = (categories.finishedNotInvoiced || []).filter(row => {
            const orderNo = number(row.OrdNo);
            const state = orderStates.get(orderNo);
            return state && !viaOrders.has(orderNo) && number(state.packedRatio) < 0.999;
        });
        return [
            { category: 'Pladelager', rows: flatPlateRows(categories), key: r => String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.FifoValue ?? r.Value), order: () => 0 },
            { category: 'Rest plader', rows: flatRestRows(categories), key: r => String(r.ProdNo || '') + '/' + String(r.OrdNo || ''), product: r => r.ProdNo, value: r => number(r.Value), order: r => number(r.OrdNo), restCode: r => r.Txt1 || r.RestCode },
            { category: 'Plader VIA', rows: categories.nestingCutting || [], key: r => String(r.OrdNo || '') + '/' + String(r.Route || '') + '/' + String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.CountedValue ?? (r.IsEstimatedRest ? 0 : r.Value)), order: r => number(r.OrdNo), sales: r => number(r.SalesOrdNo), route: r => r.Route, products: r => parseProducts(r.Products), label: r => r.Products || r.Route },
            { category: 'Stang materiale', rows: categories.stang || [], key: r => String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.FifoValue ?? r.Value), order: () => 0 },
            { category: 'Lager Komponenter (FIFO)', rows: categories.gr5Items || [], key: r => String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.FifoValue ?? r.Value), order: () => 0 },
            { category: 'Opfølgningsvarer', rows: categories.opfolgningvare || [], key: r => String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.Value), order: () => 0 },
            { category: 'VIA Laser', rows: categories.salgordreVia || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.MaterialCost) * allocation(r, 'via'), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.MainProdNo || r.CustomerName },
            { category: 'VIA Stang', rows: categories.salgordreVia || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.StangCost) * allocation(r, 'via'), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.MainProdNo || r.CustomerName },
            { category: 'Indkøbt dele', rows: categories.salgordreVia || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.PurchasedPartCost) * allocation(r, 'via'), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.MainProdNo || r.CustomerName },
            { category: 'VIA Tid', rows: categories.salgordreVia || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.TimeCost) * allocation(r, 'via'), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.MainProdNo || r.CustomerName },
            { category: 'VIA ikke pakket', rows: unpackedFinishedRows, key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.Value) * (1 - allocation(r, 'finished')), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.CustomerName },
            { category: 'Færdige SO', rows: categories.finishedNotInvoiced || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.Value) * allocation(r, 'finished'), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.CustomerName }
        ];
    }

    function buildMovements(payloadA, payloadB, context = {}) {
        const specsA = movementSpecs(payloadA, context);
        const specsB = movementSpecs(payloadB, context);
        const movements = [];
        for (let index = 0; index < specsA.length; index += 1) {
            const before = specsA[index];
            const after = specsB[index];
            const map = new Map();
            const add = (row, side, spec) => {
                const key = String(spec.key(row) || '').trim();
                if (!key) return;
                const entry = map.get(key) || { category: spec.category, key, valueA: 0, valueB: 0, productKey: '', orderNo: 0, salesOrderNo: 0, route: '', restCode: '', products: [], label: '' };
                entry[side] += number(spec.value(row));
                entry.productKey = entry.productKey || String(spec.product(row) || '').trim();
                entry.orderNo = entry.orderNo || number(spec.order(row));
                entry.salesOrderNo = entry.salesOrderNo || number(spec.sales ? spec.sales(row) : 0);
                entry.route = entry.route || String(spec.route ? spec.route(row) : '').trim();
                entry.restCode = entry.restCode || String(spec.restCode ? spec.restCode(row) : '').trim();
                entry.label = entry.label || String(spec.label ? spec.label(row) : row.Descr || row.Txt2 || '').trim();
                const products = spec.products ? spec.products(row) : [];
                entry.products = Array.from(new Set(entry.products.concat(products)));
                map.set(key, entry);
            };
            for (const row of before.rows) add(row, 'valueA', before);
            for (const row of after.rows) add(row, 'valueB', after);
            for (const entry of map.values()) {
                const diff = round(entry.valueB - entry.valueA);
                if (Math.abs(diff) <= EPSILON) continue;
                movements.push({ ...entry, diff, remaining: diff });
            }
        }
        const routesByKey = new Map((Array.isArray(context && context.routes) ? context.routes : [])
            .map(route => [String(route.nestingOrdNo) + '|' + String(route.route), route]));
        for (const movement of movements) {
            if (movement.category !== 'Plader VIA' || movement.salesOrderNo > 0) continue;
            const route = routesByKey.get(String(movement.orderNo) + '|' + String(movement.route));
            const salesOrders = route && Array.isArray(route.salesOrderNos) ? route.salesOrderNos.filter(Boolean) : [];
            if (salesOrders.length === 1) movement.salesOrderNo = number(salesOrders[0]);
        }
        return movements.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
    }

    function canonicalValueSummary(payload, context = {}) {
        const current = unwrapPayload(payload);
        const totals = current.totals || {};
        const specs = movementSpecs(current, context);
        const sumSpec = category => round(specs
            .filter(spec => spec.category === category)
            .reduce((sum, spec) => sum + spec.rows.reduce((rowSum, row) => rowSum + number(spec.value(row)), 0), 0));
        const categories = {
            plates: round(totals.plates),
            restPlates: round(totals.restPlates),
            stang: round(totals.stang),
            opfolgningvare: round(totals.opfolgningvare),
            viaLaser: sumSpec('VIA Laser'),
            viaStang: sumSpec('VIA Stang'),
            purchasedParts: sumSpec('Indkøbt dele'),
            viaTime: sumSpec('VIA Tid'),
            viaNotPacked: sumSpec('VIA ikke pakket'),
            finishedNotInvoiced: sumSpec('Færdige SO'),
            diverse: round(totals.diverse)
        };
        const total = round(Object.values(categories).reduce((sum, value) => sum + number(value), 0));
        const rawV1Total = round(totals.total !== undefined
            ? totals.total
            : ['plates', 'restPlates', 'stang', 'opfolgningvare', 'finishedNotInvoiced', 'salgordreVia', 'diverse']
                .reduce((sum, key) => sum + number(totals[key]), 0));
        return {
            categories,
            total,
            rawV1Total,
            duplicateReduction: round(Math.max(0, rawV1Total - total))
        };
    }

    function reconcileMovements(payloadA, payloadB, context = {}) {
        const movements = buildMovements(payloadA, payloadB, context);
        const transfers = [];
        const rules = [
            { from: ['Pladelager'], to: ['Plader VIA'], flow: 'Pladelager → VIA Plader', confidence: 'high', match: (a, b) => a.productKey && a.productKey === b.productKey },
            { from: ['Plader VIA'], to: ['Rest plader'], flow: 'VIA Plader → Rest plader', confidence: 'high', match: (a, b) => a.productKey && a.productKey === b.productKey && (!b.orderNo || a.orderNo === b.orderNo) },
            { from: ['Rest plader'], to: ['Plader VIA'], flow: 'Rest plader → VIA Plader', confidence: 'high', match: (a, b) => a.productKey && a.productKey === b.productKey && (!a.orderNo || a.orderNo === b.orderNo) },
            { from: ['Plader VIA'], to: ['VIA Laser'], flow: 'VIA Plader → VIA Laser', confidence: 'medium', match: (a, b) => a.salesOrderNo > 0 && a.salesOrderNo === b.salesOrderNo },
            { from: ['Plader VIA'], to: ['Opfølgningsvarer', 'Lager Komponenter (FIFO)'], flow: 'VIA Plader → lagerført produkt', confidence: 'medium', match: (a, b) => b.productKey && a.products.includes(b.productKey) },
            { from: ['VIA Laser', 'VIA Stang', 'Indkøbt dele', 'VIA Tid', 'VIA ikke pakket'], to: ['Færdige SO'], flow: 'VIA → Færdige SO (ordre færdig/pakket)', confidence: 'high', match: (a, b) => a.salesOrderNo > 0 && a.salesOrderNo === b.salesOrderNo },
            {
                from: ['Pladelager'],
                to: ['Rest plader'],
                flow: 'Pladelager → Rest plader',
                confidence: 'high',
                match: (a, b) => a.productKey && a.productKey === b.productKey && b.orderNo > 0
                    && movements.some(row => row.category === 'Plader VIA'
                        && row.diff > EPSILON
                        && row.productKey === a.productKey
                        && row.orderNo === b.orderNo)
            }
        ];

        for (const rule of rules) {
            const sources = movements.filter(row => row.remaining < -EPSILON && rule.from.includes(row.category));
            const targets = movements.filter(row => row.remaining > EPSILON && rule.to.includes(row.category));
            for (const source of sources) {
                for (const target of targets) {
                    if (source.remaining >= -EPSILON || target.remaining <= EPSILON || !rule.match(source, target)) continue;
                    const amount = round(Math.min(Math.abs(source.remaining), target.remaining));
                    if (amount <= EPSILON) continue;
                    source.remaining = round(source.remaining + amount);
                    target.remaining = round(target.remaining - amount);
                    transfers.push({
                        kind: 'transfer',
                        flow: rule.flow,
                        confidence: rule.confidence,
                        amount,
                        netAmount: 0,
                        sourceCategory: source.category,
                        sourceKey: source.key,
                        targetCategory: target.category,
                        targetKey: target.key,
                        productKey: source.productKey || target.productKey,
                        orderNo: source.orderNo || target.orderNo,
                        salesOrderNo: source.salesOrderNo || target.salesOrderNo,
                        route: source.route || target.route
                    });
                }
            }
        }

        // Arbejdstid er ikke lagerført materiale. En positiv VIA Tid-bevægelse
        // er ny værdi registreret på ordren og har derfor ingen lager-modpost.
        for (const movement of movements) {
            if (movement.category !== 'VIA Tid' || movement.remaining <= EPSILON) continue;
            const amount = round(movement.remaining);
            movement.remaining = 0;
            transfers.push({
                kind: 'addition',
                flow: 'Registreret arbejdstid → VIA Tid',
                confidence: 'high',
                amount,
                netAmount: amount,
                sourceCategory: 'Registreret arbejdstid',
                sourceKey: 'tilført værdi',
                targetCategory: movement.category,
                targetKey: movement.key,
                productKey: movement.productKey,
                orderNo: movement.orderNo,
                salesOrderNo: movement.salesOrderNo,
                route: movement.route
            });
        }

        const evidence = context && context.evidence || {};
        const evidenceByProduct = rows => {
            const map = new Map();
            for (const row of Array.isArray(rows) ? rows : []) {
                const key = String(row.prodNo || row.ProdNo || '').trim();
                if (!key) continue;
                const list = map.get(key) || [];
                list.push(row);
                map.set(key, list);
            }
            return map;
        };
        const purchasesByProduct = evidenceByProduct(evidence.purchases);
        const consumptionsByProduct = evidenceByProduct(evidence.plateConsumptions);
        const stockConsumptionValueByProductAndOrder = new Map();
        const stockConsumptionOrdersByProduct = new Map();
        for (const row of Array.isArray(evidence.stockConsumptions) ? evidence.stockConsumptions : []) {
            const productKey = String(row.prodNo || row.ProdNo || '').trim();
            const salesOrderNo = number(row.salesOrderNo || row.SalesOrderNo || row.orderNo || row.OrdNo);
            if (!productKey || !salesOrderNo) continue;
            const key = productKey + '|' + String(salesOrderNo);
            stockConsumptionValueByProductAndOrder.set(key, round(
                number(stockConsumptionValueByProductAndOrder.get(key)) + Math.abs(number(row.value ?? row.StcCst))
            ));
            const orders = stockConsumptionOrdersByProduct.get(productKey) || new Set();
            orders.add(salesOrderNo);
            stockConsumptionOrdersByProduct.set(productKey, orders);
        }
        const reservationsByProduct = evidenceByProduct(evidence.reservations);
        const invoicesByOrder = new Map((Array.isArray(evidence.invoices) ? evidence.invoices : []).map(row => [number(row.orderNo || row.OrdNo), row]));
        const searchPlateEvidence = new Map((Array.isArray(evidence.unregisteredSearchPlates) ? evidence.unregisteredSearchPlates : []).map(row => [
            String(number(row.orderNo || row.OrdNo)) + '|' + String(row.route || row.Route || '').trim() + '|' + String(row.prodNo || row.ProdNo || '').trim(),
            row
        ]));
        const registeredRestConsumptionByCode = new Map();
        for (const row of Array.isArray(evidence.registeredRestConsumptions) ? evidence.registeredRestConsumptions : []) {
            const code = String(row.restCode || row.RestCode || '').trim();
            if (!code) continue;
            const rows = registeredRestConsumptionByCode.get(code) || [];
            rows.push(row);
            registeredRestConsumptionByCode.set(code, rows);
        }
        const routesByKey = new Map((Array.isArray(context && context.routes) ? context.routes : []).map(route => [String(route.nestingOrdNo) + '|' + String(route.route), route]));

        // Opfølgningsvarer har ingen ordre i selve lagersnapshot-rækken. ProdTr
        // dokumenterer derimod den præcise lagerafgang og dens R4/salgsordre.
        // Brug kun denne registrerede forbindelse; produktnavn/prefix alene er
        // ikke tilstrækkeligt til at flytte værdien til Færdige SO.
        for (const source of movements) {
            if (source.category !== 'Opfølgningsvarer' || source.remaining >= -EPSILON || !source.productKey) continue;
            const evidencedOrders = stockConsumptionOrdersByProduct.get(source.productKey) || new Set();
            const targets = movements.filter(row => row.category === 'Færdige SO' && row.remaining > EPSILON && row.salesOrderNo > 0);
            for (const target of targets) {
                if (source.remaining >= -EPSILON) break;
                const evidenceKey = source.productKey + '|' + String(target.salesOrderNo);
                const evidenceRemaining = number(stockConsumptionValueByProductAndOrder.get(evidenceKey));
                if (evidenceRemaining <= EPSILON) continue;
                // Snapshotværdien er historisk, mens ProdTr.StcCst kan være
                // genberegnet siden da. Ved én entydig R4-modtager beviser
                // transaktionen forbindelsen, men dens aktuelle kostpris må
                // ikke efterlade små falske restbeløb på den historiske række.
                const evidenceLimit = evidencedOrders.size === 1 ? Math.abs(source.remaining) : evidenceRemaining;
                const amount = round(Math.min(Math.abs(source.remaining), target.remaining, evidenceLimit));
                if (amount <= EPSILON) continue;
                source.remaining = round(source.remaining + amount);
                target.remaining = round(target.remaining - amount);
                stockConsumptionValueByProductAndOrder.set(evidenceKey, round(Math.max(0, evidenceRemaining - amount)));
                transfers.push({
                    kind: 'transfer',
                    flow: 'Opfølgningsvarer → Færdige SO (registreret forbrug)',
                    confidence: 'high',
                    amount,
                    netAmount: 0,
                    sourceCategory: source.category,
                    sourceKey: source.key,
                    targetCategory: target.category,
                    targetKey: target.key,
                    productKey: source.productKey,
                    orderNo: target.orderNo,
                    salesOrderNo: target.salesOrderNo,
                    route: ''
                });
            }

            // Færdige SO kan allerede være udlignet af VIA-værdierne ovenfor.
            // Det gør ikke den registrerede komponentafgang ugyldig: når ProdTr
            // kun peger på én R4/salgsordre, er destinationen stadig entydig.
            // Vis derfor bevægelsen som et dokumenteret transittrin uden at
            // oprette endnu en Færdige SO-værdi (det ville være dobbeltoptælling).
            if (source.remaining < -EPSILON && evidencedOrders.size === 1) {
                const salesOrderNo = Array.from(evidencedOrders)[0];
                const amount = round(Math.abs(source.remaining));
                source.remaining = 0;
                transfers.push({
                    kind: 'transfer',
                    flow: 'Opfølgningsvarer → ordre/VIA (registreret forbrug)',
                    confidence: 'high',
                    amount,
                    netAmount: 0,
                    sourceCategory: source.category,
                    sourceKey: source.key,
                    targetCategory: 'Ordre/VIA',
                    targetKey: String(salesOrderNo),
                    productKey: source.productKey,
                    orderNo: salesOrderNo,
                    salesOrderNo,
                    route: ''
                });
            }
        }

        // Rsv er den faktiske reservation mellem lagerprodukt og ordre. Den
        // bruges kun, hvis der endnu ikke findes en ProdTr-lagerafgang for
        // produktet. NoPic/NoFin afgør, om varen stadig står reserveret eller
        // allerede er plukket/færdigmeldt; Rsv-værdien lægges aldrig til
        // lagertotalen som en ekstra post.
        for (const source of movements) {
            if (source.category !== 'Opfølgningsvarer' || source.remaining >= -EPSILON || !source.productKey) continue;
            if ((stockConsumptionOrdersByProduct.get(source.productKey) || new Set()).size) continue;
            const reservationRows = reservationsByProduct.get(source.productKey) || [];
            const destinations = new Set(reservationRows.map(row => number(row.salesOrderNo || row.SalesOrderNo || row.orderNo || row.OrdNo)).filter(Boolean));
            if (destinations.size !== 1) continue;
            const destinationOrderNo = Array.from(destinations)[0];
            const active = reservationRows.some(row => number(row.activeQty ?? row.ActiveQty) > EPSILON
                || number(row.awaitingPickQty ?? row.AwaitingPickQty) > EPSILON
                || number(row.pickedNotFinishedQty ?? row.PickedNotFinishedQty) > EPSILON);
            const amount = round(Math.abs(source.remaining));
            source.remaining = 0;
            transfers.push({
                kind: 'transfer',
                flow: active
                    ? 'Opfølgningsvarer → reserveret til ordre (Rsv)'
                    : 'Opfølgningsvarer → ordre/VIA (Rsv færdigmeldt)',
                confidence: 'high',
                amount,
                netAmount: 0,
                sourceCategory: source.category,
                sourceKey: source.key,
                targetCategory: active ? 'Reserveret lager' : 'Ordre/VIA',
                targetKey: String(destinationOrderNo),
                productKey: source.productKey,
                orderNo: number(reservationRows[0] && (reservationRows[0].orderNo || reservationRows[0].OrdNo)) || destinationOrderNo,
                salesOrderNo: number(reservationRows[0] && (reservationRows[0].salesOrderNo || reservationRows[0].SalesOrderNo)) || destinationOrderNo,
                route: ''
            });
        }

        // Et registreret REST-stykke beholder oprindelsesordren i FreeInf1,
        // mens et senere nesting naturligt har et nyt OrdNo. Den tilfældige
        // _SCR0-kode er derfor den sikre forbindelse på tværs af ordrer.
        for (const source of movements) {
            if (source.category !== 'Rest plader' || source.remaining >= -EPSILON || !source.restCode) continue;
            const candidates = (registeredRestConsumptionByCode.get(source.restCode) || [])
                .filter(row => String(row.prodNo || row.ProdNo || '').trim() === source.productKey);
            if (candidates.length !== 1) continue;
            const consumption = candidates[0];
            const nestingOrderNo = number(consumption.nestingOrderNo || consumption.OrdNo);
            const route = String(consumption.route || consumption.Route || '').trim();
            const target = movements.find(row => row.category === 'Plader VIA'
                && row.remaining > EPSILON
                && row.orderNo === nestingOrderNo
                && String(row.route) === route
                && row.productKey === source.productKey);
            if (!target) {
                const amount = round(Math.abs(source.remaining));
                source.remaining = 0;
                transfers.push({
                    kind: 'transfer',
                    flow: 'REST Plader → nesting (genbrugt REST)',
                    confidence: 'high',
                    amount,
                    netAmount: 0,
                    sourceCategory: source.category,
                    sourceKey: source.key,
                    targetCategory: consumption.routeStarted ? 'Nesting/produktion' : 'Åben nesting',
                    targetKey: String(nestingOrderNo) + (route ? '/' + route : ''),
                    productKey: source.productKey,
                    orderNo: nestingOrderNo,
                    salesOrderNo: 0,
                    route,
                    restCode: source.restCode
                });
                continue;
            }

            const transferred = round(Math.min(Math.abs(source.remaining), target.remaining));
            if (transferred > EPSILON) {
                source.remaining = round(source.remaining + transferred);
                target.remaining = round(target.remaining - transferred);
                transfers.push({
                    kind: 'transfer',
                    flow: 'REST Plader → VIA Plader (genbrugt REST)',
                    confidence: 'high',
                    amount: transferred,
                    netAmount: 0,
                    sourceCategory: source.category,
                    sourceKey: source.key,
                    targetCategory: target.category,
                    targetKey: target.key,
                    productKey: source.productKey,
                    orderNo: nestingOrderNo,
                    salesOrderNo: target.salesOrderNo,
                    route,
                    restCode: source.restCode
                });
            }

            const inputValue = Math.abs(number(consumption.value ?? consumption.Value));
            const revaluationLimit = round(Math.max(0, inputValue - transferred));
            const revaluation = round(Math.min(target.remaining, revaluationLimit));
            if (revaluation > EPSILON) {
                target.remaining = round(target.remaining - revaluation);
                transfers.push({
                    kind: 'addition',
                    flow: 'Genindvundet værdi ved REST-forbrug → VIA Plader',
                    confidence: 'high',
                    amount: revaluation,
                    netAmount: revaluation,
                    sourceCategory: 'Værdiregulering af genbrugt REST',
                    sourceKey: source.restCode,
                    targetCategory: target.category,
                    targetKey: target.key,
                    productKey: source.productKey,
                    orderNo: nestingOrderNo,
                    salesOrderNo: target.salesOrderNo,
                    route,
                    restCode: source.restCode
                });
            }
        }

        for (const movement of movements) {
            if (movement.category !== 'Pladelager' || movement.remaining <= EPSILON) continue;
            const rows = purchasesByProduct.get(movement.productKey) || [];
            const evidencedValue = round(rows.reduce((sum, row) => sum + Math.abs(number(row.value ?? row.StcCst)), 0));
            const amount = round(Math.min(movement.remaining, evidencedValue));
            if (amount <= EPSILON) continue;
            movement.remaining = round(movement.remaining - amount);
            const orderNos = Array.from(new Set(rows.map(row => number(row.orderNo || row.OrdNo)).filter(Boolean)));
            transfers.push({
                kind: 'addition',
                flow: 'Indkøb → Pladelager',
                confidence: 'high',
                amount,
                netAmount: amount,
                sourceCategory: 'Indkøbsordre',
                sourceKey: orderNos.join(', ') || 'registreret varemodtagelse',
                targetCategory: movement.category,
                targetKey: movement.key,
                productKey: movement.productKey,
                orderNo: orderNos[0] || 0,
                salesOrderNo: 0,
                route: ''
            });
        }

        for (const movement of movements) {
            if (movement.category !== 'Pladelager' || movement.remaining >= -EPSILON) continue;
            const rows = consumptionsByProduct.get(movement.productKey) || [];
            const evidencedValue = round(rows.reduce((sum, row) => sum + Math.abs(number(row.value ?? row.StcCst)), 0));
            const amount = round(Math.min(Math.abs(movement.remaining), evidencedValue));
            if (amount <= EPSILON) continue;
            movement.remaining = round(movement.remaining + amount);
            const nestingOrders = Array.from(new Set(rows.map(row => number(row.orderNo || row.OrdNo)).filter(Boolean)));
            transfers.push({
                kind: 'transfer',
                flow: 'Pladelager → VIA Plader (registreret nesting)',
                confidence: 'high',
                amount,
                netAmount: 0,
                sourceCategory: movement.category,
                sourceKey: movement.key,
                targetCategory: 'VIA Plader',
                targetKey: nestingOrders.join(', ') || 'registreret nesting',
                productKey: movement.productKey,
                orderNo: nestingOrders[0] || 0,
                salesOrderNo: 0,
                route: ''
            });
        }

        for (const movement of movements) {
            if (movement.category !== 'Plader VIA' || movement.remaining <= EPSILON) continue;
            const hasVisiblePlateReceipt = movements.some(row => row.category === 'Pladelager'
                && row.productKey === movement.productKey
                && row.diff > EPSILON);
            if (hasVisiblePlateReceipt) continue;
            const purchaseRows = purchasesByProduct.get(movement.productKey) || [];
            const consumptionRows = (consumptionsByProduct.get(movement.productKey) || [])
                .filter(row => number(row.orderNo || row.OrdNo) === movement.orderNo);
            const purchaseValue = round(purchaseRows.reduce((sum, row) => sum + Math.abs(number(row.value ?? row.StcCst)), 0));
            const consumptionValue = round(consumptionRows.reduce((sum, row) => sum + Math.abs(number(row.value ?? row.StcCst)), 0));
            const amount = round(Math.min(movement.remaining, purchaseValue, consumptionValue));
            if (amount <= EPSILON) continue;
            movement.remaining = round(movement.remaining - amount);
            const purchaseOrders = Array.from(new Set(purchaseRows.map(row => number(row.orderNo || row.OrdNo)).filter(Boolean)));
            transfers.push({
                kind: 'addition',
                flow: 'Indkøb → Pladelager → VIA Plader',
                confidence: 'high',
                amount,
                netAmount: amount,
                sourceCategory: 'Indkøbsordre',
                sourceKey: purchaseOrders.join(', ') || 'registreret varemodtagelse',
                targetCategory: movement.category,
                targetKey: movement.key,
                productKey: movement.productKey,
                orderNo: movement.orderNo,
                salesOrderNo: movement.salesOrderNo,
                route: movement.route
            });
        }

        // "Søg" på nestingens pladelinje betyder, at materialet er en fysisk
        // rest, som ikke tidligere er registreret i pladelager/REST. Den kan
        // derfor forklare en positiv VIA-værdi, men er en reel tilført værdi
        // og ikke en lageroverførsel med en skjult modpost.
        for (const movement of movements) {
            if (movement.category !== 'Plader VIA' || movement.remaining <= EPSILON) continue;
            const route = routesByKey.get(String(movement.orderNo) + '|' + String(movement.route));
            const evidenceKey = String(movement.orderNo) + '|' + String(movement.route) + '|' + movement.productKey;
            const searchPlate = searchPlateEvidence.get(evidenceKey) || (route && (route.plates || []).find(plate => plate.unregisteredRestSource
                && String(plate.prodNo || '').trim() === movement.productKey));
            if (!searchPlate) continue;
            const amount = round(movement.remaining);
            movement.remaining = 0;
            transfers.push({
                kind: 'addition',
                flow: 'Uregistreret REST (Søg) → VIA Plader',
                confidence: 'high',
                amount,
                netAmount: amount,
                sourceCategory: 'Ikke lagerregistreret restplade',
                sourceKey: String(searchPlate.sourceInfo || 'Søg').trim(),
                targetCategory: movement.category,
                targetKey: movement.key,
                productKey: movement.productKey,
                orderNo: movement.orderNo,
                salesOrderNo: movement.salesOrderNo,
                route: movement.route
            });
        }

        for (const movement of movements) {
            if (movement.category !== 'Færdige SO' || movement.remaining >= -EPSILON) continue;
            const invoice = invoicesByOrder.get(movement.orderNo);
            if (!invoice) continue;
            const amount = round(Math.abs(movement.remaining));
            movement.remaining = 0;
            transfers.push({
                kind: 'exit',
                flow: 'Færdige SO → Faktureret',
                confidence: 'high',
                amount,
                netAmount: -amount,
                sourceCategory: movement.category,
                sourceKey: movement.key,
                targetCategory: 'Faktureret',
                targetKey: String(invoice.invoiceNo || invoice.InvoNo || '').trim() || String(movement.orderNo),
                productKey: '',
                orderNo: movement.orderNo,
                salesOrderNo: movement.salesOrderNo,
                route: ''
            });
        }

        // Hvis en ordre både er flyttet ud af VIA og faktureret mellem de to
        // snapshots, findes der ingen synlig mellemtilstand i slut-snapshot'et.
        // Fakturaen dokumenterer da hele kæden uden at oprette en kunstig pluspost.
        for (const movement of movements) {
            if (!['VIA Laser', 'VIA Stang', 'Indkøbt dele', 'VIA Tid', 'VIA ikke pakket'].includes(movement.category)
                || movement.remaining >= -EPSILON) continue;
            const invoice = invoicesByOrder.get(movement.orderNo);
            if (!invoice) continue;
            const amount = round(Math.abs(movement.remaining));
            movement.remaining = 0;
            transfers.push({
                kind: 'exit',
                flow: 'VIA → Færdige SO → Faktureret',
                confidence: 'high',
                amount,
                netAmount: -amount,
                sourceCategory: movement.category,
                sourceKey: movement.key,
                targetCategory: 'Faktureret',
                targetKey: String(invoice.invoiceNo || invoice.InvoNo || '').trim() || String(movement.orderNo),
                productKey: '',
                orderNo: movement.orderNo,
                salesOrderNo: movement.salesOrderNo,
                route: ''
            });
        }

        for (const movement of movements) {
            if (movement.category !== 'Plader VIA' || movement.remaining >= -EPSILON) continue;
            const route = routesByKey.get(String(movement.orderNo) + '|' + String(movement.route));
            if (!route || route.status !== 'completed' || !(route.productionOrderNos || []).length) continue;
            const amount = round(Math.abs(movement.remaining));
            movement.remaining = 0;
            const productionOrders = route.productionOrderNos || [];
            const salesOrders = route.salesOrderNos || [];
            transfers.push({
                kind: 'transfer',
                flow: 'VIA Plader → produktions-/salgsordre',
                confidence: 'high',
                amount,
                netAmount: 0,
                sourceCategory: movement.category,
                sourceKey: movement.key,
                targetCategory: 'Produktionsordre',
                targetKey: 'PO ' + productionOrders.join(', ') + (salesOrders.length ? ' · SO ' + salesOrders.join(', ') : ''),
                productKey: movement.productKey,
                orderNo: productionOrders[0] || movement.orderNo,
                salesOrderNo: salesOrders[0] || movement.salesOrderNo,
                route: movement.route
            });
        }

        const unresolved = movements.filter(row => Math.abs(row.remaining) > EPSILON).map(row => ({ ...row, remaining: round(row.remaining) }));
        const transferredValue = round(transfers.filter(row => row.kind === 'transfer').reduce((sum, row) => sum + row.amount, 0));
        const addedValue = round(transfers.filter(row => row.kind === 'addition').reduce((sum, row) => sum + row.amount, 0));
        const removedValue = round(transfers.filter(row => row.kind === 'exit').reduce((sum, row) => sum + row.amount, 0));
        return {
            movements,
            transfers,
            unresolved,
            orderAllocations: orderAllocationSummary(payloadA, payloadB, context),
            transferredValue,
            addedValue,
            removedValue,
            reconciledValue: round(transferredValue + addedValue + removedValue),
            residualValue: round(unresolved.reduce((sum, row) => sum + row.remaining, 0))
        };
    }

    return { unwrapPayload, parseProducts, buildMovements, canonicalValueSummary, reconcileMovements };
});
