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
        return Math.round(number(value) * 100) / 100;
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

    function movementSpecs(payload) {
        const categories = unwrapPayload(payload).categories || {};
        return [
            { category: 'Pladelager', rows: flatPlateRows(categories), key: r => String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.FifoValue ?? r.Value), order: () => 0 },
            { category: 'Rest plader', rows: flatRestRows(categories), key: r => String(r.ProdNo || '') + '/' + String(r.OrdNo || ''), product: r => r.ProdNo, value: r => number(r.Value), order: r => number(r.OrdNo) },
            { category: 'Plader VIA', rows: categories.nestingCutting || [], key: r => String(r.OrdNo || '') + '/' + String(r.Route || '') + '/' + String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.CountedValue ?? (r.IsEstimatedRest ? 0 : r.Value)), order: r => number(r.OrdNo), sales: r => number(r.SalesOrdNo), route: r => r.Route, products: r => parseProducts(r.Products), label: r => r.Products || r.Route },
            { category: 'Stang materiale', rows: categories.stang || [], key: r => String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.FifoValue ?? r.Value), order: () => 0 },
            { category: 'Lager Komponenter (FIFO)', rows: categories.gr5Items || [], key: r => String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.FifoValue ?? r.Value), order: () => 0 },
            { category: 'Opfølgningsvarer', rows: categories.opfolgningvare || [], key: r => String(r.ProdNo || ''), product: r => r.ProdNo, value: r => number(r.Value), order: () => 0 },
            { category: 'VIA Laser', rows: categories.salgordreVia || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.MaterialCost), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.MainProdNo || r.CustomerName },
            { category: 'VIA Stang', rows: categories.salgordreVia || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.StangCost), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.MainProdNo || r.CustomerName },
            { category: 'Indkøbt dele', rows: categories.salgordreVia || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.PurchasedPartCost), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.MainProdNo || r.CustomerName },
            { category: 'VIA Tid', rows: categories.salgordreVia || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.TimeCost), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.MainProdNo || r.CustomerName },
            { category: 'Færdige SO', rows: categories.finishedNotInvoiced || [], key: r => String(r.OrdNo || ''), product: () => '', value: r => number(r.Value), order: r => number(r.OrdNo), sales: r => number(r.OrdNo), label: r => r.CustomerName }
        ];
    }

    function buildMovements(payloadA, payloadB) {
        const specsA = movementSpecs(payloadA);
        const specsB = movementSpecs(payloadB);
        const movements = [];
        for (let index = 0; index < specsA.length; index += 1) {
            const before = specsA[index];
            const after = specsB[index];
            const map = new Map();
            const add = (row, side, spec) => {
                const key = String(spec.key(row) || '').trim();
                if (!key) return;
                const entry = map.get(key) || { category: spec.category, key, valueA: 0, valueB: 0, productKey: '', orderNo: 0, salesOrderNo: 0, route: '', products: [], label: '' };
                entry[side] += number(spec.value(row));
                entry.productKey = entry.productKey || String(spec.product(row) || '').trim();
                entry.orderNo = entry.orderNo || number(spec.order(row));
                entry.salesOrderNo = entry.salesOrderNo || number(spec.sales ? spec.sales(row) : 0);
                entry.route = entry.route || String(spec.route ? spec.route(row) : '').trim();
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
        return movements.sort((a, b) => Math.abs(b.diff) - Math.abs(a.diff));
    }

    function reconcileMovements(payloadA, payloadB, context = {}) {
        const movements = buildMovements(payloadA, payloadB);
        const transfers = [];
        const rules = [
            { from: ['Pladelager'], to: ['Plader VIA'], flow: 'Pladelager → VIA Plader', confidence: 'high', match: (a, b) => a.productKey && a.productKey === b.productKey },
            { from: ['Plader VIA'], to: ['Rest plader'], flow: 'VIA Plader → Rest plader', confidence: 'high', match: (a, b) => a.productKey && a.productKey === b.productKey && (!b.orderNo || a.orderNo === b.orderNo) },
            { from: ['Rest plader'], to: ['Plader VIA'], flow: 'Rest plader → VIA Plader', confidence: 'high', match: (a, b) => a.productKey && a.productKey === b.productKey && (!a.orderNo || a.orderNo === b.orderNo) },
            { from: ['Plader VIA'], to: ['VIA Laser'], flow: 'VIA Plader → VIA Laser', confidence: 'medium', match: (a, b) => a.salesOrderNo > 0 && a.salesOrderNo === b.salesOrderNo },
            { from: ['Plader VIA'], to: ['Opfølgningsvarer', 'Lager Komponenter (FIFO)'], flow: 'VIA Plader → lagerført produkt', confidence: 'medium', match: (a, b) => b.productKey && a.products.includes(b.productKey) },
            { from: ['VIA Laser', 'VIA Stang', 'Indkøbt dele', 'VIA Tid'], to: ['Færdige SO'], flow: 'VIA → Færdige SO', confidence: 'high', match: (a, b) => a.salesOrderNo > 0 && a.salesOrderNo === b.salesOrderNo },
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
        const invoicesByOrder = new Map((Array.isArray(evidence.invoices) ? evidence.invoices : []).map(row => [number(row.orderNo || row.OrdNo), row]));
        const searchPlateEvidence = new Map((Array.isArray(evidence.unregisteredSearchPlates) ? evidence.unregisteredSearchPlates : []).map(row => [
            String(number(row.orderNo || row.OrdNo)) + '|' + String(row.route || row.Route || '').trim() + '|' + String(row.prodNo || row.ProdNo || '').trim(),
            row
        ]));
        const routesByKey = new Map((Array.isArray(context && context.routes) ? context.routes : []).map(route => [String(route.nestingOrdNo) + '|' + String(route.route), route]));

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
            transferredValue,
            addedValue,
            removedValue,
            reconciledValue: round(transferredValue + addedValue + removedValue),
            residualValue: round(unresolved.reduce((sum, row) => sum + row.remaining, 0))
        };
    }

    return { unwrapPayload, parseProducts, buildMovements, reconcileMovements };
});
