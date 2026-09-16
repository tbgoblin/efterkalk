(function (root) {
    'use strict';

    const catalog = [
        { id: 'invoice-kpi', title: 'Omsætning i regnskabsåret', module: 'omsaetning', kind: 'kpi' },
        { id: 'order-flow', title: 'Ordreflow / ordrebeholdning', module: 'omsaetning', source: 'order-flow', kind: 'flow' },
        { id: 'latest', title: 'Seneste fakturaordrer', module: 'efterkalk', kind: 'list' },
        { id: 'best', title: 'Største dækningsbidrag', module: 'efterkalk', kind: 'list' },
        { id: 'risk', title: 'Laveste dækningsbidrag', module: 'efterkalk', kind: 'list' },
        { id: 'customers', title: 'Kunder efter omsætning', module: 'omsaetning', kind: 'bar' },
        { id: 'customer-share', title: 'Kunder og aktivitet', module: 'omsaetning', kind: 'donut' },
        { id: 'recent-orders', title: 'Seneste ordreindgang', module: 'ordreindgang', source: 'recent-orders', kind: 'list' },
        { id: 'sellers', title: 'Ansvarlige efter fakturabeløb', module: 'efterkalk', kind: 'bar' },
        { id: 'trend', title: 'Omsætning pr. måned', module: 'omsaetning', kind: 'bar' },
        { id: 'coverage', title: 'Kostgrundlag', module: 'efterkalk', kind: 'kpi' },
        { id: 'via-kpi', title: 'Arbejde i gang', module: 'salgordre-via', kind: 'kpi' },
        { id: 'via-top', title: 'Største kapitalbinding', module: 'salgordre-via', kind: 'list' },
        { id: 'via-due', title: 'Overskredet levering', module: 'salgordre-via', kind: 'list' },
        { id: 'via-next', title: 'Næste leveringer', module: 'salgordre-via', kind: 'list' },
        { id: 'resources', title: 'Næste ressourcer', module: 'salgordre-via', kind: 'bar' },
        { id: 'via-cost', title: 'VIA fordelt på kosttype', module: 'salgordre-via', kind: 'bar' },
        { id: 'load-kpi', title: 'Kapacitet og planlagte timer', module: 'belastning', kind: 'kpi' },
        { id: 'load-resources', title: 'Belastning pr. ressource', module: 'belastning', kind: 'capacity' },
        { id: 'load-days', title: 'Planlagt arbejde og kapacitet pr. dag', module: 'belastning', kind: 'capacity' },
        { id: 'load-backlog', title: 'Restarbejde før i dag', module: 'belastning', kind: 'bar' }
    ];
    const templates = [
        { id: 'economy', name: 'Økonomi', widgets: ['invoice-kpi', 'order-flow', 'best', 'risk', 'coverage', 'trend', 'via-cost'] },
        { id: 'production', name: 'Produktion', widgets: ['load-kpi', 'load-resources', 'load-days', 'load-backlog', 'via-due', 'via-next'] },
        { id: 'sales', name: 'Salg', widgets: ['invoice-kpi', 'order-flow', 'customer-share', 'recent-orders', 'customers', 'latest', 'sellers', 'trend', 'best'] },
        { id: 'management', name: 'Ledelse', widgets: ['invoice-kpi', 'order-flow', 'load-kpi', 'trend', 'load-resources', 'risk', 'via-due'] }
    ];
    const knownIds = new Set(catalog.map(widget => widget.id));
    const number = value => value === null || value === undefined || value === '' || typeof value === 'boolean'
        ? null : (Number.isFinite(Number(value)) ? Number(value) : null);
    const sum = (rows, field) => rows.reduce((total, row) => total + (number(row[field]) || 0), 0);
    const text = value => String(value == null ? '' : value).trim();

    function dateKey(value) {
        const raw = text(value);
        const compact = /^\d{8}$/.test(raw) ? raw : raw.slice(0, 10).replace(/-/g, '');
        if (!/^\d{8}$/.test(compact)) return '';
        const year = Number(compact.slice(0, 4));
        const month = Number(compact.slice(4, 6));
        const day = Number(compact.slice(6, 8));
        const parsed = new Date(year, month - 1, day);
        return year >= 1900 && parsed.getFullYear() === year && parsed.getMonth() === month - 1 && parsed.getDate() === day
            ? compact.slice(0, 4) + '-' + compact.slice(4, 6) + '-' + compact.slice(6, 8) : '';
    }

    function storageKey(username, database) {
        return 'goh:dashboard:v1:' + encodeURIComponent(text(database) || 'production') + ':' + encodeURIComponent(text(username).toLowerCase());
    }

    function normalizeConfig(raw) {
        const config = raw && typeof raw === 'object' ? raw : {};
        const used = new Set();
        let customCount = 0;
        const boards = (Array.isArray(config.boards) ? config.boards : []).filter(board => {
            if (!board || used.has(board.id)) return false;
            const standard = templates.some(template => template.id === board.id);
            if (!standard && (!/^custom-[a-z0-9-]+$/.test(board.id) || customCount >= 12)) return false;
            if (!standard) customCount += 1;
            used.add(board.id);
            return true;
        }).map(board => ({
            id: board.id, name: text(board.name).slice(0, 48) || 'Min dashboard',
            widgets: [...new Set(Array.isArray(board.widgets) ? board.widgets : [])].filter(id => knownIds.has(id)),
            wide: [...new Set(Array.isArray(board.wide) ? board.wide : [])].filter(id => knownIds.has(id)),
            layout: normalizeLayout(board.layout, board.widgets),
            options: Object.fromEntries((Array.isArray(board.widgets) ? board.widgets : []).filter(id => knownIds.has(id) && board.options && Object.prototype.hasOwnProperty.call(board.options, id)).map(id => {
                const defaults = widgetOptions(id);
                return [id, Object.fromEntries(Object.entries(widgetOptions(id, board.options[id])).filter(([key, value]) => value !== defaults[key]))];
            }))
        }));
        const active = templates.some(item => item.id === config.active) || boards.some(item => item.id === config.active)
            ? config.active : 'management';
        const flow = config.flow && typeof config.flow === 'object' ? config.flow : {};
        return { active, boards, period: ['all', 'month', 'quarter', 'year'].includes(config.period) ? config.period : 'all', limit: [5, 10].includes(config.limit) ? config.limit : 5,
            theme: ['light', 'dark', 'system'].includes(config.theme) ? config.theme : 'light',
            query: text(config.query).replace(/\s+/g, ' ').slice(0, 80),
            flow: { month: /^\d{4}-(0[1-9]|1[0-2])$/.test(flow.month) ? flow.month : '', cohort: ['all', 'prior', 'new'].includes(flow.cohort) ? flow.cohort : 'all', metric: ['opening', 'incoming', 'invoiced', 'closing'].includes(flow.metric) ? flow.metric : 'closing', query: text(flow.query).slice(0, 120), page: Number.isInteger(flow.page) ? Math.max(0, Math.min(10000, flow.page)) : 0, completedOnly: flow.completedOnly === true } };
    }

    function widgetOptions(id, raw = {}) {
        const options = raw && typeof raw === 'object' ? raw : {};
        const bounded = (value, fallback, min, max) => Number.isFinite(value) ? Math.max(min, Math.min(max, Math.round(value))) : fallback;
        return {
            title: text(options.title).slice(0, 60), search: text(options.search).slice(0, 80),
            limit: [5, 10, 20, 50].includes(options.limit) ? options.limit : 0,
            days: bounded(options.days, id === 'recent-orders' ? 30 : 20, 1, 90),
            includeBacklog: options.includeBacklog !== false, showEvening: options.showEvening !== false,
            showCustomers: options.showCustomers !== false, showMonths: options.showMonths !== false,
            showDb: options.showDb !== false, showPercentage: options.showPercentage !== false,
            deliveryDays: bounded(options.deliveryDays, 0, 0, 365),
            maxDbPercent: Number.isFinite(options.maxDbPercent) ? Math.max(-100, Math.min(100, options.maxDbPercent)) : null,
            metric: options.metric === 'orders' ? 'orders' : 'revenue',
            top: bounded(options.top, 5, 3, 6), selected: text(options.selected).slice(0, 160)
        };
    }

    function loadHorizons(ids, options = {}) {
        return [...new Set(ids.filter(id => catalog.find(widget => widget.id === id)?.module === 'belastning').map(id => widgetOptions(id, options[id]).days))].sort((left, right) => left - right);
    }

    function customerShares(revenue, orders, metric = 'revenue', top = 5) {
        const groups = new Map();
        for (const row of metric === 'orders' ? orders : revenue) {
            const key = row.customerId || row.customer;
            if (!groups.has(key)) groups.set(key, { key, label: row.customer, value: 0, orders: new Set() });
            const customer = groups.get(key);
            if (metric === 'orders') { customer.orders.add(row.order); customer.value = customer.orders.size; }
            else customer.value += row.revenue;
        }
        const rows = [...groups.values()].map(({ orders: _orders, ...row }) => ({ ...row, unit: metric === 'orders' ? 'antal' : 'DKK' })).sort((left, right) => right.value - left.value || left.label.localeCompare(right.label, 'da'));
        const positive = rows.filter(row => row.value > 0);
        const segments = positive.slice(0, top).map(row => ({ ...row, members: [row] }));
        if (positive.length > top) segments.push({ key: '__others', label: 'Andre', value: sum(positive.slice(top), 'value'), members: positive.slice(top) });
        return { rows, segments, positiveTotal: sum(positive, 'value'), negativeTotal: sum(rows.filter(row => row.value < 0), 'value'), total: sum(rows, 'value') };
    }

    function createPreferenceStore(options) {
        const cached = options.cached;
        let current = normalizeConfig(cached?.config || options.legacy || options.initial);
        let version = Number.isInteger(cached?.version) ? cached.version : null;
        let pending = cached?.pending === true;
        let ready = false;
        let active = true;
        let loading = null;
        let writing = null;
        let generation = 0;
        const status = (value, error) => { if (active) options.onStatus(value, error); };
        const cache = () => { if (active) options.cache({ config: current, version, pending }); };

        async function flush() {
            if (writing || !ready || !pending || !active) return writing;
            writing = (async () => {
                while (pending && ready && active) {
                    const sequence = generation;
                    const snapshot = normalizeConfig(current);
                    const expected = version;
                    status('saving');
                    try {
                        const result = await options.save(snapshot, expected);
                        if (!active) return;
                        if (result.version !== expected + 1) throw new Error('Ugyldigt svar fra GOH');
                        version = result.version;
                        pending = sequence !== generation;
                        cache();
                        if (!pending) status('saved');
                    } catch (error) {
                        if (!active) return;
                        ready = false;
                        status(error.status === 409 ? 'conflict' : 'error', error);
                    }
                }
            })();
            try { await writing; } finally { writing = null; }
        }

        async function load(discard = false) {
            if (!active || loading || writing) return loading || writing;
            ready = false;
            status('loading');
            loading = (async () => {
                try {
                    const result = await options.load();
                    if (!active) return;
                    if (options.schemaVersion && result.schemaVersion !== options.schemaVersion) throw new Error('Genstart GOH-serveren og genindlæs siden for at gemme personlige indstillinger.');
                    if (!Number.isInteger(result.version) || result.version < 0 || (result.version > 0 && !result.config)) throw new Error('Ugyldigt svar fra GOH');
                    if (!discard && pending && version !== result.version && !(version === null && result.version === 0)) {
                        options.onConfig(current);
                        status('conflict');
                        return;
                    }
                    if (discard || !pending) {
                        current = normalizeConfig(result.config || (discard ? options.initial : current));
                        pending = result.version === 0;
                    }
                    version = result.version;
                    ready = true;
                    cache();
                    options.onConfig(current);
                    status(pending ? 'saving' : 'saved');
                } catch (error) {
                    if (!active) return;
                    cache();
                    options.onConfig(current);
                    status(error.status === 409 ? 'conflict' : 'error', error);
                }
            })();
            try { await loading; } finally { loading = null; }
            return flush();
        }

        return {
            load,
            set(value) {
                if (!active) return;
                current = normalizeConfig(value);
                pending = true;
                generation++;
                cache();
                return flush();
            },
            dispose() { active = false; }
        };
    }

    function normalizeLayout(raw, ids) {
        const result = {};
        for (const id of Array.isArray(ids) ? ids : []) {
            const rect = raw && Object.prototype.hasOwnProperty.call(raw, id) ? raw[id] : null;
            if (!knownIds.has(id) || !rect || !['x', 'y', 'w', 'h'].every(key => Number.isFinite(rect[key]))) continue;
            const width = Math.max(id === 'order-flow' ? 6 : 3, Math.min(12, Math.round(rect.w)));
            result[id] = { x: Math.max(0, Math.min(12 - width, Math.round(rect.x))), y: Math.max(0, Math.min(1000, Math.round(rect.y))), w: width, h: Math.max(8, Math.min(32, Math.round(rect.h))) };
        }
        return result;
    }

    function arrangeWidgets(ids, saved = {}, wide = [], pinned = '', compact = false) {
        const layout = normalizeLayout(saved, ids);
        const placed = {};
        const overlap = (left, right) => left.x < right.x + right.w && left.x + left.w > right.x && left.y < right.y + right.h && left.y + left.h > right.y;
        const order = [...ids].sort((left, right) => Number(right === pinned) - Number(left === pinned)
            || Number(Boolean(layout[right])) - Number(Boolean(layout[left]))
            || (layout[left]?.y || 0) - (layout[right]?.y || 0));
        for (const id of order) {
            const previous = layout[id];
            const rect = previous ? { ...previous } : { x: 0, y: 0, w: id === 'order-flow' || wide.includes(id) ? 8 : 4, h: id === 'order-flow' ? 22 : 12 };
            const reposition = !previous || compact && id !== pinned;
            const columns = reposition ? [...new Set([rect.x, ...Array.from({ length: 13 - rect.w }, (_, index) => index)])] : [rect.x];
            if (compact && id !== pinned) rect.y = 0;
            let found = false;
            while (!found) {
                for (const column of columns) {
                    rect.x = column;
                    if (!Object.values(placed).some(other => overlap(rect, other))) { found = true; break; }
                }
                if (!found) rect.y++;
            }
            placed[id] = rect;
        }
        return placed;
    }

    function allowedWidgets(ids, canAccess) {
        return ids.map(id => catalog.find(widget => widget.id === id)).filter(widget => widget && canAccess(widget.module)
            && (widget.module === 'belastning' || ['via-due', 'via-next'].includes(widget.id) || canAccess('omsaetning')));
    }

    function allowedTemplates(canAccess) {
        return templates.filter(template => (template.id === 'production' || canAccess('omsaetning')) && allowedWidgets(template.widgets, canAccess).length);
    }

    function availableBoards(config, canAccess) {
        return allowedTemplates(canAccess).map(template => config.boards.find(board => board.id === template.id) || template)
            .concat(config.boards.filter(board => board.id.startsWith('custom-')));
    }

    async function refreshVisibleSources(widgetIds, canAccess, loadSource, options = {}) {
        const modules = [...new Set(allowedWidgets(widgetIds, canAccess).flatMap(widget => widget.id === 'customer-share' && options[widget.id]?.metric === 'orders' && canAccess('efterkalk') ? ['efterkalk'] : [widget.source || widget.module]))];
        const results = await Promise.allSettled(modules.map(moduleKey => Promise.resolve().then(() => loadSource(moduleKey))));
        return modules.filter((moduleKey, index) => results[index].status === 'rejected');
    }

    function fiscalRange(today) {
        const date = dateKey(today);
        const year = Number(date.slice(0, 4)) - (Number(date.slice(5, 7)) < 7 ? 1 : 0);
        return { from: year + '-07-01', fra: String(year * 100 + 1), til: String((year + 1) * 100 + 1) };
    }

    function economicRange(today, period) {
        if (period !== 'year') return fiscalRange(today);
        const year = Number(dateKey(today).slice(0, 4));
        return { from: year + '-01-01', fra: String((year - 1) * 100 + 7), til: String(year * 100 + 7) };
    }

    function createSourceCache(now = Date.now) {
        const entries = new Map();
        return {
            peek: key => entries.get(key),
            clear: () => entries.clear(),
            put: (key, payload) => entries.set(key, { state: 'ready', payload, expires: now() + 15 * 60 * 1000 }),
            load(key, loader, force = false) {
                const previous = entries.get(key);
                if (previous && previous.promise) return previous.promise;
                if (!force && previous && now() < previous.expires) return previous.error ? Promise.reject(previous.error) : Promise.resolve(previous.payload);
                const entry = { state: 'loading', payload: previous && previous.payload, expires: Infinity };
                entries.set(key, entry);
                entry.promise = Promise.resolve().then(loader).then(payload => {
                    if (!payload || payload.ok === false) throw new Error('Data kunne ikke indlæses');
                    entry.payload = payload; entry.state = 'ready'; entry.expires = now() + 15 * 60 * 1000;
                    return payload;
                }).catch(error => {
                    entry.state = 'error'; entry.error = error; entry.expires = now() + 60000;
                    throw error;
                }).finally(() => { entry.promise = null; });
                return entry.promise;
            }
        };
    }

    function aggregate(rows, getKey, getLabel, valueField) {
        const groups = new Map();
        for (const row of rows) {
            const key = getKey(row);
            if (!groups.has(key)) groups.set(key, { label: getLabel(row), value: 0, count: 0 });
            const group = groups.get(key);
            group.value += number(row[valueField]) || 0;
            group.count++;
        }
        return [...groups.values()].sort((left, right) => right.value - left.value || left.label.localeCompare(right.label, 'da'));
    }

    function isCreditOrder(row, notes) {
        const note = notes && notes[String(row.OrdNo)];
        return number(row.InvoAm) < 0 || Boolean(note && note.isCreditNote === true);
    }

    async function hydrateOrderCosts(rows, fetchMargin, isActive, onUpdate) {
        const pending = rows.filter(row => number(row.TotalCost) === null);
        let next = 0;
        await Promise.all(Array.from({ length: Math.min(3, pending.length) }, async () => {
            while (next < pending.length && isActive()) {
                const row = pending[next++];
                try {
                    const margin = await fetchMargin(row.OrdNo);
                    if (!isActive()) return;
                    const cost = number(margin && margin.totalCost);
                    if (cost !== null && !margin.error && number(row.TotalCost) === null) row.TotalCost = cost + (number(margin.styklisteFallbackCost) || 0);
                } catch (_) {}
                if (isActive()) onUpdate();
            }
        }));
    }

    function loadSearchFilters(query, unfilteredLoad) {
        const value = text(query).replace(/\s+/g, ' ').slice(0, 80);
        const rows = [unfilteredLoad?.odd, unfilteredLoad?.even].flatMap(group => group?.rows || []);
        const resourceMatch = rows.some(row => (text(row.ResGr) + ' ' + text(row.Nm)).toLocaleLowerCase('da').includes(value.toLocaleLowerCase('da')));
        if (!value || resourceMatch) return { ord: '', kunde: '', resourceQuery: value };
        return /^\d+$/.test(value) ? { ord: value.slice(0, 12), kunde: '', resourceQuery: '' } : { ord: '', kunde: value, resourceQuery: '' };
    }

    function buildData(input) {
        const options = widgetOptions(input.widgetId, input.options);
        const today = dateKey(input.today) || dateKey(new Date().toISOString());
        const fiscalFrom = fiscalRange(today).from;
        const quarterMonth = String(Math.floor((Number(today.slice(5, 7)) - 1) / 3) * 3 + 1).padStart(2, '0');
        const from = input.period === 'month' ? today.slice(0, 7) + '-01' : input.period === 'quarter' ? today.slice(0, 4) + '-' + quarterMonth + '-01' : input.period === 'year' ? today.slice(0, 4) + '-01-01' : '';
        const periodFrom = from || fiscalFrom;
        const query = text(input.query).replace(/\s+/g, ' ').toLocaleLowerCase('da');
        const widgetQuery = text(options.search).toLocaleLowerCase('da');
        const matches = row => [query, widgetQuery].every(term => !term || [row.customer, row.seller, row.order].join(' ').toLocaleLowerCase('da').includes(term));
        const revenue = ((input.revenue && input.revenue.rows) || []).map(row => ({
            date: dateKey(row.date), revenue: (number(row.revenueMio) || 0) * 1000000,
            customer: text(row.customerName) || 'Uden kunde', customerId: text(row.custNo)
        })).filter(row => row.date >= periodFrom && row.date <= today
            && [query, widgetQuery].every(term => !term || (row.customer + ' ' + row.customerId).toLocaleLowerCase('da').includes(term)));
        const orders = (input.orders || []).filter(row => !isCreditOrder(row, input.notes)).map(row => {
            const state = input.margins && input.margins[String(row.OrdNo)];
            const cost = state && state.status === 'success' ? number(state.totalCost) : number(row.TotalCost);
            const revenue = number(row.InvoAm);
            return { order: number(row.OrdNo), customer: text(row.CustomerName) || 'Ukendt kunde', customerId: text(row.CustNo), seller: text(row.SellerUsr) || 'Ikke angivet', date: dateKey(row.LstInvDt), revenue, cost, db: cost !== null && revenue !== null ? revenue - cost : null };
        }).filter(row => row.order > 0 && matches(row) && row.date >= periodFrom && row.date <= today);
        const valued = orders.filter(row => row.db !== null);
        const sellerCosts = new Map();
        for (const row of valued) {
            if (!sellerCosts.has(row.seller)) sellerCosts.set(row.seller, { revenue: 0, db: 0, count: 0 });
            const totals = sellerCosts.get(row.seller);
            totals.revenue += row.revenue;
            totals.db += row.db;
            totals.count++;
        }
        const sellers = aggregate(orders, row => row.seller, row => row.seller, 'revenue').map(row => {
            const totals = sellerCosts.get(row.label);
            return { ...row, dbPercent: totals && totals.revenue > 0 ? totals.db / totals.revenue * 100 : null, costCount: totals ? totals.count : 0 };
        });
        const via = (input.via || []).map(row => ({
            order: number(row.OrdNo), customer: text(row.CustomerName) || 'Ukendt kunde', seller: text(row.SellerUsr) || 'Ikke angivet',
            date: dateKey(row.DeliveryDate), resource: text(row.ResourceName) || 'Ikke angivet',
            costDataAvailable: row.CostDataAvailable !== false,
            material: number(row.MaterialCost) || 0, bar: number(row.StangCost) || 0, purchased: number(row.PurchasedPartCost) || 0, time: number(row.TimeCost) || 0
        })).filter(row => row.order > 0 && matches(row)).map(row => ({ ...row, value: row.costDataAvailable ? row.material + row.bar + row.purchased + row.time : null }));
        const viaCostsComplete = via.every(row => row.costDataAvailable);
        const viaTotal = field => viaCostsComplete ? sum(via, field) : null;
        const limit = options.limit || (input.limit === 10 ? 10 : 5);
        const recentFrom = new Date(today + 'T12:00:00');
        recentFrom.setDate(recentFrom.getDate() - options.days + 1);
        const recentOrders = (input.recentOrders || []).map(row => ({ order: number(row.OrdNo), customer: text(row.CustomerName) || 'Ukendt kunde', customerId: text(row.CustNo), seller: '', date: dateKey(row.OrderDate), value: number(row.OrderValueDkk) }))
            .filter(row => row.order > 0 && (row.value === null || row.value >= 0) && !input.notes?.[row.order]?.isCreditNote && row.date >= dateKey(recentFrom.toISOString()) && row.date <= today && matches(row));
        const ranked = valued.slice().sort((left, right) => right.db - left.db || right.order - left.order);
        const overdue = via.filter(row => row.date && row.date < today).sort((left, right) => left.date.localeCompare(right.date) || right.value - left.value);
        const load = input.load || {};
        const serverLoadQuery = text(load.kunde || load.ord).toLocaleLowerCase('da');
        const resourceQuery = query === serverLoadQuery ? '' : query;
        const loadRows = [load.odd, load.even].flatMap(group => group && Array.isArray(group.rows) ? group.rows : []).map(row => ({
            resource: text(row.ResGr), name: text(row.Nm), date: dateKey(row.Dato),
            backlog: !row.Dato || dateKey(row.Dato) && dateKey(row.Dato) < (load.toDay || today),
            planned: (number(row.Resv) || 0) / 60, capacity: (number(row.Kap) || 0) / 60, evening: (number(row.Aften) || 0) / 60
        })).filter(row => row.resource && (row.date || row.backlog) && (options.includeBacklog || !row.backlog) && [resourceQuery, widgetQuery].every(term => !term || (row.resource + ' ' + row.name).toLocaleLowerCase('da').includes(term)));
        const resourceGroups = new Map();
        const dayGroups = new Map();
        for (const row of loadRows) {
            if (!resourceGroups.has(row.resource)) resourceGroups.set(row.resource, { label: row.resource + ' ' + row.name, value: 0, capacity: 0, evening: 0, backlog: 0, unit: 'timer' });
            const resource = resourceGroups.get(row.resource);
            resource.value += row.planned; resource.capacity += row.capacity; resource.evening += row.evening;
            if (row.backlog) resource.backlog += row.planned + row.evening;
            else {
                if (!dayGroups.has(row.date)) dayGroups.set(row.date, { label: row.date, value: 0, capacity: 0, evening: 0, unit: 'timer' });
                const day = dayGroups.get(row.date);
                day.value += row.planned; day.capacity += row.capacity; day.evening += row.evening;
            }
        }
        const resources = [...resourceGroups.values()];
        const loadRatio = row => row.capacity > 0 ? row.value / row.capacity : row.value > 0 ? Infinity : 0;
        const monthlyRevenue = new Map(aggregate(revenue, row => row.date.slice(0, 7), row => row.date.slice(0, 7), 'revenue').map(row => [row.label, row]));
        const monthCursor = new Date(periodFrom + 'T12:00:00');
        const trend = [];
        while (dateKey(monthCursor.toISOString()).slice(0, 7) <= today.slice(0, 7)) {
            const month = monthCursor.getFullYear() + '-' + String(monthCursor.getMonth() + 1).padStart(2, '0');
            trend.push(monthlyRevenue.get(month) || { label: month, value: 0 });
            monthCursor.setMonth(monthCursor.getMonth() + 1);
        }
        const revenueMonths = input.period === 'quarter' && options.showMonths ? Array.from({ length: 3 }, (_, index) => {
            const month = today.slice(0, 4) + '-' + String(Number(quarterMonth) + index).padStart(2, '0');
            const label = new Intl.DateTimeFormat('da-DK', { month: 'long' }).format(new Date(month + '-01T12:00:00'));
            return { label, value: month > today.slice(0, 7) ? null : monthlyRevenue.get(month)?.value || 0, unit: 'DKK' };
        }) : [{ label: 'Denne måned', value: sum(revenue.filter(row => row.date.slice(0, 7) === today.slice(0, 7)), 'revenue'), unit: 'DKK' }];
        return {
            orders, via, valued, revenue, recentOrders, loadRows, fiscalFrom, from, periodFrom, today,
            customerShares: customerShares(revenue, orders, options.metric, options.top),
            widgets: {
                'invoice-kpi': [{ label: 'Bogført omsætning', value: sum(revenue, 'revenue'), unit: 'DKK' }, ...revenueMonths, ...(options.showCustomers ? [{ label: 'Kunder', value: new Set(revenue.filter(row => row.customerId && row.customerId !== '0').map(row => row.customerId)).size, unit: 'antal' }] : [])],
                'recent-orders': recentOrders.slice().sort((left, right) => right.date.localeCompare(left.date) || right.order - left.order).slice(0, limit),
                'customer-share': customerShares(revenue, orders, options.metric, options.top).rows,
                latest: orders.slice().sort((left, right) => right.date.localeCompare(left.date) || right.order - left.order).slice(0, limit).map(row => ({ ...row, value: row.revenue })),
                best: ranked.slice(0, limit).map(row => ({ ...row, value: row.db })),
                risk: ranked.slice().reverse().filter(row => options.maxDbPercent === null || row.revenue > 0 && row.db / row.revenue * 100 <= options.maxDbPercent).slice(0, limit).map(row => ({ ...row, value: row.db })),
                customers: aggregate(revenue, row => row.customerId || row.customer, row => row.customer, 'revenue').slice(0, limit),
                sellers: sellers.slice(0, limit),
                trend,
                coverage: [{ label: 'Ordrer med kendt kost', value: valued.length, unit: 'antal' }, { label: 'Mangler kostgrundlag', value: orders.length - valued.length, unit: 'antal' }, ...(options.showPercentage ? [{ label: 'Dækning', value: orders.length ? valued.length / orders.length * 100 : null, unit: '%' }] : [])],
                'via-kpi': [{ label: 'Registreret VIA-kost', value: viaTotal('value'), unit: 'DKK' }, { label: 'Aktive ordrer', value: via.length, unit: 'antal' }, { label: 'Overskredet levering', value: overdue.length, unit: 'antal' }],
                'via-top': via.filter(row => row.costDataAvailable).sort((left, right) => right.value - left.value || right.order - left.order).slice(0, limit),
                'via-due': overdue.slice(0, limit),
                'via-next': via.filter(row => row.date && row.date >= today && (!options.deliveryDays || (Date.parse(row.date) - Date.parse(today)) / 86400000 < options.deliveryDays)).sort((left, right) => left.date.localeCompare(right.date) || left.order - right.order).slice(0, limit),
                resources: viaCostsComplete ? aggregate(via, row => row.resource, row => row.resource, 'value').slice(0, limit) : [],
                'via-cost': [{ label: 'Materiale', value: viaTotal('material') }, { label: 'Stang', value: viaTotal('bar') }, { label: 'Indkøbte dele', value: viaTotal('purchased') }, { label: 'Tid', value: viaTotal('time') }],
                'load-kpi': [{ label: options.includeBacklog ? 'Planlagt dagarbejde inkl. rest' : 'Planlagt dagarbejde', value: sum(loadRows, 'planned'), unit: 'timer' }, { label: 'Kapacitet', value: sum(loadRows, 'capacity'), unit: 'timer' }, ...(options.showEvening ? [{ label: 'Aftenarbejde', value: sum(loadRows, 'evening'), unit: 'timer' }] : []), { label: 'Ressourcer over kapacitet', value: resources.filter(row => loadRatio(row) > 1).length, unit: 'antal' }],
                'load-resources': resources.slice().sort((left, right) => loadRatio(right) - loadRatio(left) || right.value - left.value).slice(0, limit),
                'load-days': [...dayGroups.values()].sort((left, right) => left.label.localeCompare(right.label)),
                'load-backlog': resources.filter(row => row.backlog > 0).map(row => ({ label: row.label, value: row.backlog, unit: 'timer' })).sort((left, right) => right.value - left.value).slice(0, limit)
            }
        };
    }

    const model = { catalog, templates, normalizeConfig, normalizeLayout, arrangeWidgets, widgetOptions, loadHorizons, customerShares, createPreferenceStore, allowedWidgets, allowedTemplates, availableBoards, refreshVisibleSources, fiscalRange, economicRange, createSourceCache, isCreditOrder, hydrateOrderCosts, loadSearchFilters, storageKey, dateKey, buildData };
    if (typeof module === 'object' && module.exports) module.exports = model;
    else root.GohDashboard = { model };

    if (!root.document) return;
    const escape = value => text(value).replace(/[&<>"']/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]));
    const format = (value, unit = 'DKK') => value === null ? 'Afventer' : new Intl.NumberFormat('da-DK', { maximumFractionDigits: unit === 'antal' ? 0 : 2 }).format(value) + (unit === 'antal' ? '' : ' ' + unit);
    const displayDate = value => value ? value.split('-').reverse().join('.') : 'Dato mangler';
    const byId = id => root.document.getElementById(id);
    const moduleLabel = key => ({ efterkalk: 'Efterkalk', 'salgordre-via': 'SalgOrdre VIA', omsaetning: 'Omsætning', belastning: 'Belastning', 'order-flow': 'Ordreflow', ordreindgang: 'Ordreindgang', 'recent-orders': 'Seneste ordreindgang' }[key] || key);
    let context = null;
    let config = normalizeConfig();
    let scope = '';
    let query = '';
    let draft = null;
    let searchTimer = null;
    let opener = null;
    let refreshTask = null;
    let layoutDraft = null;
    let pointerDrag = null;
    let preferenceStore = null;
    let preferencesReady = false;
    let preferenceStatus = 'loading';
    let localRecoveryAvailable = true;
    let flowSaveTimer = null;
    let widgetDraft = null;

    function permitted(widget) { return context && context.authenticated && widget && allowedWidgets([widget.id], context.canAccess).length > 0; }
    function currentBoard() { return layoutDraft || config.boards.find(board => board.id === config.active) || templates.find(board => board.id === config.active) || templates[3]; }
    function notify(message, error = false) {
        const status = byId('gohDashStatus');
        if (status) { status.textContent = message; status.dataset.error = String(error); }
    }
    function persist() {
        if (!scope || !context.authenticated) return;
        config.query = query;
        preferenceStore?.set(config);
    }
    function saveBoard(board) {
        config.boards = config.boards.some(item => item.id === board.id)
            ? config.boards.map(item => item.id === board.id ? board : item) : config.boards.concat(board);
        config.active = board.id;
        config = normalizeConfig(config);
        persist();
    }
    function saveTheme(value) {
        if (!preferencesReady || !context?.authenticated) return;
        config.theme = normalizeConfig({ theme: value }).theme;
        persist();
    }
    function flushPendingPreferences() {
        if (!preferencesReady || !preferenceStore || (!searchTimer && !flowSaveTimer)) return;
        if (searchTimer) query = text(byId('gohDashSearch')?.value).replace(/\s+/g, ' ').slice(0, 80);
        clearTimeout(searchTimer); searchTimer = null;
        clearTimeout(flowSaveTimer); flowSaveTimer = null;
        config.query = query;
        preferenceStore.set(config);
    }
    function renderPreferenceStatus(value = preferenceStatus) {
        preferenceStatus = value;
        root.GohTheme?.bind(config.theme, preferencesReady ? saveTheme : null, value);
        const status = byId('gohDashPersistence');
        if (!status) return;
        status.textContent = { loading: 'Henter GOH-profil...', saving: 'Gemmer i GOH...', saved: 'Gemt i GOH', error: 'Ikke gemt i GOH', conflict: 'GOH-profil ændret på en anden postation' }[value];
        status.dataset.error = String(['error', 'conflict'].includes(value));
        if (['error', 'conflict'].includes(value)) status.textContent += localRecoveryAvailable ? ' · Lokal kopi bevaret' : ' · Kun i denne session';
        const retry = byId('gohDashRetryPreferences');
        retry.hidden = !['error', 'conflict'].includes(value);
        retry.title = value === 'conflict' ? 'Hent GOH-profil' : 'Prøv GOH-synkronisering igen';
        retry.setAttribute('aria-label', retry.title);
    }
    function optionsFor(widget) {
        const options = widgetOptions(widget.id, currentBoard().options?.[widget.id]);
        if (widget.id === 'customer-share' && !context.canAccess('efterkalk')) options.metric = 'revenue';
        return options;
    }
    function visibleOptions() {
        return Object.fromEntries(allowedWidgets(currentBoard().widgets, context.canAccess).map(widget => [widget.id, optionsFor(widget)]));
    }
    function data(widget) {
        const allowed = context.canAccess;
        const options = widget ? optionsFor(widget) : {};
        const periodReady = context.economicPeriod === undefined || context.economicPeriod === config.period;
        const loadReady = context.loadQuery === undefined || context.loadQuery === query;
        const selectedLoad = widget?.module === 'belastning' ? (context.loads ? context.loads[options.days]?.payload : (context.load?.dage || 20) === options.days ? context.load : null) : context.load;
        return buildData({ ...context, widgetId: widget?.id, options, orders: periodReady && allowed('efterkalk') && allowed('omsaetning') ? context.orders : [], recentOrders: allowed('ordreindgang') && allowed('omsaetning') ? context.recentOrders : [], margins: periodReady && allowed('efterkalk') && allowed('omsaetning') ? context.margins : {}, via: allowed('salgordre-via') ? context.via : [], revenue: periodReady && allowed('omsaetning') ? context.revenue : null, load: loadReady && allowed('belastning') ? selectedLoad : null, period: config.period, query, limit: config.limit });
    }
    function action(label, content, attributes = '') {
        return '<button type="button" class="goh-icon" title="' + escape(label) + '" aria-label="' + escape(label) + '" ' + attributes + '>' + content + '</button>';
    }
    function layoutIcon(kind) {
        return '<svg class="lucide" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + ({
            move: '<path d="m5 9-3 3 3 3m4-10 3-3 3 3m4 4 3 3-3 3m-4 4-3 3-3-3M2 12h20M12 2v20"/>',
            resize: '<path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/>',
            save: '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h12l4 4v12a2 2 0 0 1-2 2Z"/><path d="M17 21v-8H7v8M7 3v5h8"/>',
            cancel: '<path d="m18 6-12 12M6 6l12 12"/>',
            compact: '<path d="M3 3h7v7H3zM14 3h7v7h-7zM3 14h7v7H3zM14 14h7v7h-7z"/>',
            settings: '<path d="M21 4h-7M10 4H3M21 12H8M4 12H3M21 20h-3M14 20H3M14 2v4M8 10v4M18 18v4"/>'
        }[kind] || '') + '</svg>';
    }
    function renderCustomerShares(shares, options) {
        const colors = ['#167260', '#377fa3', '#c18a27', '#a45683', '#648541', '#c05a44', '#77827d'].map((color, index) => 'var(--goh-chart-' + (index + 1) + ', ' + color + ')');
        const unit = options.metric === 'orders' ? 'antal' : 'DKK';
        const metric = options.metric === 'orders' ? 'Fakturaordrer' : 'Omsætning';
        const selected = shares.segments.find(segment => segment.key === options.selected);
        const details = selected ? selected.members : options.selected === '__negative' ? shares.rows.filter(row => row.value < 0) : shares.rows;
        let offset = 0;
        const circles = shares.segments.map((segment, index) => {
            const percent = segment.value / shares.positiveTotal * 100;
            const circle = '<circle class="goh-donut-segment" cx="60" cy="60" r="44" pathLength="100" fill="none" stroke="' + colors[index] + '" stroke-width="' + (segment.key === options.selected ? 23 : 18) + '" stroke-dasharray="' + percent + ' ' + (100 - percent) + '" stroke-dashoffset="' + (-offset) + '" transform="rotate(-90 60 60)" role="button" tabindex="0" data-donut-segment="' + escape(segment.key) + '" aria-label="' + escape(segment.label + ': ' + format(segment.value, unit)) + '" aria-pressed="' + (segment.key === options.selected) + '"><title>' + escape(segment.label + ': ' + format(segment.value, unit)) + '</title></circle>';
            offset += percent;
            return circle;
        }).join('');
        const compact = new Intl.NumberFormat('da-DK', { notation: 'compact', maximumFractionDigits: 1 }).format(shares.positiveTotal);
        return (shares.negativeTotal < 0 ? '<div class="goh-share-balance"><strong>Netto: ' + escape(format(shares.total, unit)) + '</strong><button type="button" data-donut-segment="__negative">Negative kundesaldi: ' + escape(format(shares.negativeTotal, unit)) + '</button></div>' : '')
            + '<div class="goh-donut"><div class="goh-donut-plot"><svg viewBox="0 0 120 120" aria-label="Kundefordeling"><circle cx="60" cy="60" r="44" fill="none" stroke="#e4eae7" stroke-width="18"/>' + circles + '</svg><div class="goh-donut-center"><strong>' + escape(compact) + '</strong><span>' + (shares.negativeTotal < 0 ? 'Positive kundesaldi' : metric) + '</span></div></div><div class="goh-donut-legend">'
            + shares.segments.map((segment, index) => '<button type="button" data-donut-segment="' + escape(segment.key) + '" aria-pressed="' + (options.selected === segment.key) + '"><span class="goh-chart-swatch" style="background:' + colors[index] + '"></span><span>' + escape(segment.label) + '</span><strong>' + (segment.value / shares.positiveTotal * 100).toFixed(1).replace('.', ',') + ' %</strong></button>').join('')
            + '</div></div><div class="goh-chart-detail"><table class="goh-widget-table"><caption>' + escape(selected?.label || (options.selected === '__negative' ? 'Negative kundesaldi' : 'Alle kunder')) + ' · ' + details.length + ' kunder</caption><thead><tr><th>Kunde</th><th>' + metric + (unit === 'DKK' ? ' DKK' : '') + '</th></tr></thead><tbody>'
            + details.slice(0, options.limit || config.limit).map(row => '<tr><td>' + escape(row.label) + '</td><td class="' + (row.value < 0 ? 'goh-negative' : '') + '">' + escape(format(row.value, unit)) + '</td></tr>').join('') + '</tbody></table>'
            + (details.length > (options.limit || config.limit) ? '<small>Viser ' + (options.limit || config.limit) + ' af ' + details.length + '</small>' : '')
            + '<p class="goh-chart-note">Netto: ' + escape(format(shares.total, unit)) + (shares.negativeTotal < 0 ? ' · Positive kundesaldi: ' + escape(format(shares.positiveTotal, unit)) + ' · <button type="button" data-donut-segment="__negative">Negative saldi: ' + escape(format(shares.negativeTotal, unit)) + '</button>' : '') + '</p></div>';
    }
    function commitWidgetOptions(id, options) {
        const original = currentBoard();
        saveBoard({ ...original, options: { ...original.options, [id]: options } });
        renderControls(); renderWidgets();
        return true;
    }
    function openWidgetSettings(id, button) {
        const widget = catalog.find(item => item.id === id);
        if (!permitted(widget) || layoutDraft) return;
        opener = button;
        widgetDraft = { id, board: currentBoard().id, options: optionsFor(widget) };
        byId('gohWidgetSettingsTitle').textContent = widget.title;
        renderWidgetSettings();
        byId('gohWidgetSettings').showModal();
    }
    function renderWidgetSettings() {
        if (!widgetDraft) return;
        const widget = catalog.find(item => item.id === widgetDraft.id);
        const options = widgetDraft.options;
        const input = (label, key, type = 'text', attributes = '') => '<label>' + label + '<input data-widget-option="' + key + '" type="' + type + '" value="' + escape(options[key] ?? '') + '" ' + attributes + '></label>';
        const toggle = (label, key) => '<label class="goh-setting-check"><input type="checkbox" data-widget-option="' + key + '"' + (options[key] ? ' checked' : '') + '>' + label + '</label>';
        const select = (label, key, choices, numeric = false) => '<label>' + label + '<select data-widget-option="' + key + '"' + (numeric ? ' data-numeric="true"' : '') + '>' + choices.map(([value, name]) => '<option value="' + value + '"' + (String(options[key]) === String(value) ? ' selected' : '') + '>' + name + '</option>').join('') + '</select></label>';
        let fields = input('Titel', 'title', 'text', 'maxlength="60" placeholder="' + escape(widget.title) + '"') + input(widget.module === 'belastning' ? 'Ressourcefilter' : 'Lokalt kunde-/ordrefilter', 'search', 'search', 'maxlength="80"');
        if (['list', 'bar', 'donut', 'flow'].includes(widget.kind) && !['trend', 'via-cost'].includes(widget.id) || widget.id === 'load-resources') fields += select('Antal viste rækker', 'limit', [[0, widget.kind === 'flow' ? 'Standard (8)' : 'Følg dashboard'], [5, '5'], [10, '10'], [20, '20'], [50, '50']], true);
        if (widget.module === 'belastning' || widget.id === 'recent-orders') fields += input(widget.module === 'belastning' ? 'Dage frem' : 'Seneste antal dage', 'days', 'number', 'min="1" max="90" step="1" required');
        if (widget.module === 'belastning' && widget.id !== 'load-backlog') fields += toggle('Medtag restarbejde før i dag', 'includeBacklog') + toggle('Vis aftenarbejde', 'showEvening');
        if (widget.id === 'invoice-kpi') fields += toggle('Vis kvartalets måneder', 'showMonths') + toggle('Vis antal kunder', 'showCustomers');
        if (widget.id === 'risk') fields += input('Højeste DB % (tom = alle)', 'maxDbPercent', 'number', 'min="-100" max="100" step="0.1"');
        if (widget.id === 'via-next') fields += select('Leveringshorisont', 'deliveryDays', [[0, 'Alle fremtidige'], [7, '7 dage'], [14, '14 dage'], [30, '30 dage'], [60, '60 dage'], [90, '90 dage']], true);
        if (widget.id === 'sellers') fields += toggle('Vis vægtet DB %', 'showDb');
        if (widget.id === 'coverage') fields += toggle('Vis dækningsprocent', 'showPercentage');
        if (widget.kind === 'donut') fields += select('Fordeling efter', 'metric', [['revenue', 'Bogført omsætning'], ...(context.canAccess('efterkalk') ? [['orders', 'Antal fakturaordrer']] : [])]) + select('Største kunder + Andre', 'top', [[3, '3'], [4, '4'], [5, '5'], [6, '6']], true);
        byId('gohWidgetSettingsFields').innerHTML = fields;
    }
    function saveWidgetSettings(event) {
        event.preventDefault();
        if (!widgetDraft || widgetDraft.board !== currentBoard().id) return;
        const options = { ...widgetDraft.options, selected: '' };
        byId('gohWidgetSettingsFields').querySelectorAll('[data-widget-option]').forEach(field => {
            options[field.dataset.widgetOption] = field.type === 'checkbox' ? field.checked : field.type === 'number' || field.dataset.numeric ? field.value === '' ? null : Number(field.value) : field.value;
        });
        if (commitWidgetOptions(widgetDraft.id, options)) byId('gohWidgetSettings').close();
    }
    function selectCustomerSegment(element) {
        const id = element.closest('[data-widget-id]').dataset.widgetId;
        const options = optionsFor(catalog.find(widget => widget.id === id));
        options.selected = options.selected === element.dataset.donutSegment ? '' : element.dataset.donutSegment;
        commitWidgetOptions(id, options);
        byId('gohDashGrid').querySelector('[data-widget-id="' + id + '"] .goh-donut-legend button[aria-pressed="true"]')?.focus({ preventScroll: true });
    }
    function mount(host) {
        host.className = 'goh-dashboard';
        host.innerHTML = '<div class="goh-toolbar"><div id="gohDashTemplates" class="goh-templates" role="group" aria-label="Dashboards"></div>'
            + '<div class="goh-toolbar-actions">'
            + '<button type="button" class="goh-primary" data-dash-action="create">+ Ny dashboard</button>'
            + '<button type="button" data-dash-action="edit">Tilpas</button>'
            + action('Indret dashboard', layoutIcon('move'), 'id="gohDashArrange" data-dash-action="arrange"')
            + action('Pak widgets tæt', layoutIcon('compact'), 'id="gohDashCompact" data-dash-action="compact" hidden')
            + action('Gem layout', layoutIcon('save'), 'id="gohDashSaveLayout" data-dash-action="save-layout" hidden')
            + action('Annuller layout', layoutIcon('cancel'), 'id="gohDashCancelLayout" data-dash-action="cancel-layout" hidden')
            + action('Opdater dashboard-data', '<svg class="lucide lucide-refresh-cw" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/></svg>', 'id="gohDashRefresh" data-dash-action="refresh" aria-busy="false"')
            + action('Eksportér viste data som CSV', '&#8595;', 'data-dash-action="export"') + '</div></div>'
            + '<div class="goh-view-heading"><h3 id="gohDashTitle"></h3><span id="gohDashIdentity"></span></div>'
            + '<div class="goh-save-state"><span id="gohDashPersistence" role="status" aria-live="polite"></span>' + action('Prøv GOH-synkronisering igen', '&#8635;', 'id="gohDashRetryPreferences" data-dash-action="retry-preferences" hidden') + '</div>'
            + '<div class="goh-filters"><label id="gohDashPeriodLabel">Økonomiperiode<select id="gohDashPeriod"><option value="all">Fra regnskabsårets start (juli)</option><option value="month">Denne måned</option><option value="quarter">Dette kvartal</option><option value="year">Dette kalenderår</option></select></label>'
            + '<label class="goh-search">Kunde / ansvarlig / ordre / ressource<input id="gohDashSearch" type="search" placeholder="Søg i data" maxlength="120"></label>'
            + '<label>Top / seneste<select id="gohDashLimit"><option value="5">5</option><option value="10">10</option></select></label></div>'
            + '<div id="gohDashScope" class="goh-scope"></div><div id="gohDashGrid" class="goh-widget-grid"></div>'
            + '<div id="gohDashStatus" class="goh-status" role="status" aria-live="polite"></div>'
            + '<dialog id="gohDashEditor" class="goh-editor" aria-labelledby="gohDashEditorTitle"><form id="gohDashForm">'
            + '<header><h3 id="gohDashEditorTitle">Min dashboard</h3>' + action('Luk uden at gemme', '&times;', 'data-dash-action="cancel"') + '</header>'
            + '<div class="goh-editor-body"><label>Navn<input id="gohDashName" required maxlength="48" autocomplete="off"></label>'
            + '<label>Udgangspunkt<select id="gohDashBase"><option value="current">Aktuelt layout</option>' + templates.map(template => '<option value="' + template.id + '">' + escape(template.name) + '</option>').join('') + '<option value="empty">Tom dashboard</option></select></label>'
            + '<h4>Widgetbibliotek</h4><div id="gohDashCatalog" class="goh-catalog"></div>'
            + '<h4>Rækkefølge og størrelse</h4><div id="gohDashSelected"></div><p id="gohDashEditorError" role="alert"></p></div>'
            + '<footer><button type="button" class="goh-danger" data-dash-action="delete">Slet dashboard</button><span class="goh-spacer"></span><button type="button" data-dash-action="cancel">Annuller</button><button class="goh-primary" type="submit">Gem dashboard</button></footer></form></dialog>'
            + '<dialog id="gohWidgetSettings" class="goh-editor" aria-labelledby="gohWidgetSettingsTitle"><form id="gohWidgetSettingsForm"><header><h3 id="gohWidgetSettingsTitle">Widgetindstillinger</h3>' + action('Luk', '&times;', 'data-dash-action="cancel-widget"') + '</header><div id="gohWidgetSettingsFields" class="goh-editor-body"></div><footer><button type="button" data-dash-action="reset-widget">Standardindstillinger</button><span class="goh-spacer"></span><button type="button" data-dash-action="cancel-widget">Annuller</button><button type="submit" class="goh-primary">Gem indstillinger</button></footer></form></dialog>';
        host.addEventListener('click', onClick);
        host.addEventListener('change', onChange);
        host.addEventListener('pointerdown', startLayoutPointer);
        host.addEventListener('keydown', onLayoutKey);
        root.addEventListener('pagehide', flushPendingPreferences);
        byId('gohDashSearch').addEventListener('input', () => {
            clearTimeout(searchTimer);
            searchTimer = setTimeout(() => { searchTimer = null; query = byId('gohDashSearch').value.trim().replace(/\s+/g, ' ').slice(0, 80); persist(); renderWidgets(); }, 450);
        });
        byId('gohDashForm').addEventListener('submit', saveDraft);
        byId('gohWidgetSettingsForm').addEventListener('submit', saveWidgetSettings);
        byId('gohWidgetSettings').addEventListener('close', () => { widgetDraft = null; opener?.isConnected && opener.focus(); });
        byId('gohDashEditor').addEventListener('close', () => { draft = null; if (opener && opener.isConnected) opener.focus(); });
    }
    function renderControls() {
        const availableTemplates = allowedTemplates(context.canAccess);
        if (templates.some(item => item.id === config.active) && !availableTemplates.some(item => item.id === config.active)) config.active = availableTemplates[0]?.id || 'production';
        const board = currentBoard();
        byId('gohDashTemplates').innerHTML = availableBoards(config, context.canAccess).map(item => '<button type="button" data-template="' + item.id + '" aria-pressed="' + (config.active === item.id) + '">' + escape(item.name) + '</button>').join('');
        byId('gohDashBase').querySelectorAll('option').forEach(option => { option.hidden = templates.some(item => item.id === option.value) && !availableTemplates.some(item => item.id === option.value); });
        byId('gohDashPeriodLabel').hidden = !context.canAccess('omsaetning') || !allowedWidgets(board.widgets, context.canAccess).some(widget => ['efterkalk', 'omsaetning'].includes(widget.module) && widget.kind !== 'flow');
        byId('gohDashPeriod').value = config.period;
        byId('gohDashLimit').value = String(config.limit);
        byId('gohDashTitle').textContent = board.name;
        byId('gohDashIdentity').textContent = context.username + ' · ' + context.database + ' · Personlig GOH-profil';
        byId('dashboardOperationalOverview').dataset.template = templates.some(item => item.id === config.active) ? config.active : 'custom';
        byId('dashboardOperationalOverview').classList.toggle('goh-arranging', Boolean(layoutDraft));
        for (const id of ['gohDashCompact', 'gohDashSaveLayout', 'gohDashCancelLayout']) byId(id).hidden = !layoutDraft;
        byId('gohDashArrange').hidden = Boolean(layoutDraft);
        byId('gohDashArrange').disabled = preferenceStatus === 'loading';
        byId('dashboardOperationalOverview').querySelectorAll('.goh-templates button, .goh-filters input, .goh-filters select, [data-dash-action="create"], [data-dash-action="edit"]').forEach(control => { control.disabled = Boolean(layoutDraft) || preferenceStatus === 'loading'; });
        byId('gohDashRetryPreferences').disabled = Boolean(layoutDraft);
    }
    function renderRefreshControl() {
        const button = byId('gohDashRefresh');
        if (!button || !context) return;
        const widgets = allowedWidgets(currentBoard().widgets, context.canAccess);
        const unavailable = typeof context.refreshSources !== 'function';
        button.disabled = unavailable || Boolean(refreshTask) || Boolean(layoutDraft) || !widgets.length || widgets.some(widget => sourceState(widget) === 'loading');
        button.setAttribute('aria-busy', String(Boolean(refreshTask)));
        button.title = unavailable ? 'Genstart GOH for at opdatere dashboard-data' : refreshTask ? 'Opdaterer dashboard-data...' : 'Opdater dashboard-data';
    }
    async function refreshData() {
        if (refreshTask || !context.authenticated || typeof context.refreshSources !== 'function' || byId('gohDashRefresh').disabled) return;
        const task = {};
        const refreshScope = scope;
        refreshTask = task;
        renderRefreshControl();
        notify('Opdaterer data...');
        try {
            const failed = await context.refreshSources(allowedWidgets(currentBoard().widgets, context.canAccess).map(widget => widget.id), visibleOptions());
            if (refreshTask !== task || scope !== refreshScope) return;
            renderWidgets();
            notify(failed.length ? 'Kunne ikke opdatere: ' + failed.map(moduleLabel).join(', ') + '. Tidligere data bevares, hvis de findes.'
                : 'Data opdateret kl. ' + new Intl.DateTimeFormat('da-DK', { hour: '2-digit', minute: '2-digit', second: '2-digit' }).format(new Date()), failed.length > 0);
        } catch (_) {
            if (refreshTask === task && scope === refreshScope) notify('Opdatering fejlede. Prøv igen.', true);
        } finally {
            if (refreshTask === task) { refreshTask = null; renderRefreshControl(); }
        }
    }
    function sourceState(widget) {
        if (widget.module === 'belastning' && context.loadQuery !== undefined && context.loadQuery !== query) return 'loading';
        if (widget.kind === 'flow') return 'ready';
        if (['efterkalk', 'omsaetning'].includes(widget.module) && context.economicPeriod !== undefined && context.economicPeriod !== config.period) return 'loading';
        if (widget.id === 'customer-share' && optionsFor(widget).metric === 'orders') return context.orderState;
        if (widget.module === 'belastning' && context.loads) return context.loads[optionsFor(widget).days]?.state || 'loading';
        return { efterkalk: context.orderState, 'salgordre-via': context.viaState, omsaetning: context.revenueState, belastning: context.loadState, ordreindgang: context.recentState }[widget.module];
    }
    function sourceRows(widget, result) {
        if (widget.id === 'customer-share' && optionsFor(widget).metric === 'orders') return result.orders;
        return { efterkalk: result.orders, 'salgordre-via': result.via, omsaetning: result.revenue, belastning: result.loadRows, ordreindgang: result.recentOrders }[widget.module] || [];
    }
    function renderWidget(widget, rows, result) {
        const options = optionsFor(widget);
        const available = sourceRows(widget, result);
        const source = widget.id === 'customer-share' && options.metric === 'orders' ? 'Fakturaordrer · Fra ' + displayDate(result.periodFrom) + ' · Uden kreditnotaer' : { efterkalk: 'Efterkalk · Alle fakturaordrer fra ' + displayDate(result.periodFrom) + ' · ' + result.orders.length + ' ordrer · ' + result.valued.length + ' med kendt kost' + (context.costState === 'loading' ? ' · Beregner kost...' : context.costState === 'error' ? ' · Nogle kostberegninger mangler' : ''), 'salgordre-via': 'VIA · Aktuelle åbne ordrer', omsaetning: 'Omsætning · Fra ' + displayDate(result.periodFrom) + ' · Standardkonti', belastning: 'Belastning · ' + displayDate(result.today) + ' + ' + options.days + ' dage · Timer' + (options.search ? ' · Ressource: ' + options.search : ''), ordreindgang: 'Ordredato · Seneste ' + options.days + ' dage · Aktuel salgsværdi, ikke fakturering' }[widget.module];
        const title = options.title || (widget.id === 'invoice-kpi' ? ({ all: widget.title, year: 'Omsætning i kalenderåret', quarter: 'Omsætning i kvartalet', month: 'Omsætning i måneden' }[config.period]) : widget.title);
        const header = '<header title="Flyt widget"><div><span class="goh-widget-source">' + moduleLabel(widget.module) + '</span><h4>' + escape(title) + '</h4></div><div class="goh-widget-tools">' + action('Indstillinger: ' + title, layoutIcon('settings'), 'data-widget-settings="' + widget.id + '"' + (layoutDraft ? ' disabled' : '')) + action('Åbn ' + moduleLabel(widget.module), '&#8599;', 'data-open-module="' + widget.module + '"') + '</div></header>';
        if (widget.kind === 'flow') return header + '<div class="goh-widget-body"><div data-order-flow-host></div></div><footer>Egen månedsvælger · Salgsværdi, ikke VIA-kost</footer>';
        let body;
        const state = sourceState(widget);
        if (!available.length && state !== 'ready') {
            const label = state === 'loading' ? 'Henter data...' : state === 'error' ? 'Data kunne ikke indlæses' : 'Ingen indlæste data';
            body = '<div class="goh-empty" role="status">' + label + '<button type="button" data-open-module="' + widget.module + '">Åbn ' + moduleLabel(widget.module) + '</button></div>';
        } else if (!available.length || !rows.length) {
            body = '<div class="goh-empty">' + (['best', 'risk'].includes(widget.id) && context.costState === 'loading' ? 'Beregner kost for periodens ordrer...' : widget.id === 'via-due' && available.length ? 'Ingen overskredne leveringer med kendt dato' : widget.id === 'load-backlog' && available.length ? 'Intet restarbejde før i dag' : 'Ingen data i dette udsnit') + '</div>';
        } else if (widget.kind === 'donut') {
            body = renderCustomerShares(result.customerShares, options);
        } else if (widget.kind === 'kpi') {
            body = '<dl class="goh-kpis">' + rows.map((row, index) => '<div' + (index === 0 ? ' class="goh-kpi-main"' : '') + '><dt>' + escape(row.label) + '</dt><dd class="' + (row.value < 0 ? 'goh-negative' : '') + '">' + escape(format(row.value, row.unit)) + '</dd></div>').join('') + '</dl>';
        } else if (widget.kind === 'capacity') {
            const max = Math.max(1, ...rows.flatMap(row => [row.value, row.capacity]));
            body = '<div class="goh-capacity-legend"><span>Dagarbejde</span><span>Kapacitet</span></div><ol class="goh-bars goh-capacity">' + rows.map(row => {
                const ratio = row.capacity > 0 ? format(row.value / row.capacity * 100, '%') : 'Ingen kapacitet';
                return '<li><div><span>' + escape(row.label) + '</span><strong class="' + (row.value > row.capacity ? 'goh-negative' : '') + '">' + escape(ratio) + '</strong></div><div class="goh-track" title="' + escape('Dagarbejde: ' + format(row.value, 'timer')) + '"><span style="width:' + (row.value / max * 100).toFixed(2) + '%"></span></div><div class="goh-track goh-capacity-track" title="' + escape('Kapacitet: ' + format(row.capacity, 'timer')) + '"><span style="width:' + (row.capacity / max * 100).toFixed(2) + '%"></span></div><small>' + escape(format(row.value, 'timer') + ' / ' + format(row.capacity, 'timer') + (options.showEvening ? ' · Aften: ' + format(row.evening, 'timer') : '')) + '</small></li>';
            }).join('') + '</ol>';
        } else if (widget.id === 'sellers' && options.showDb) {
            const maxRevenue = Math.max(1, ...rows.map(row => Math.abs(row.value)));
            const maxPercent = Math.max(100, ...rows.map(row => Math.abs(row.dbPercent || 0)));
            body = '<ol class="goh-bars goh-sellers">' + rows.map(row => {
                const percent = row.dbPercent === null ? (row.value === 0 ? 'Intet fakturabeløb' : 'Afventer kost') : format(row.dbPercent, '%');
                const coverage = row.costCount < row.count ? ' · Kost: ' + row.costCount + '/' + row.count + ' ordrer' : '';
                return '<li><div><span>' + escape(row.label) + '</span><strong>' + escape(format(row.value)) + '</strong></div>'
                    + '<div class="goh-track" title="' + escape('Fakturabeløb: ' + format(row.value)) + '"><span style="width:' + (Math.abs(row.value) / maxRevenue * 100).toFixed(2) + '%"></span></div>'
                    + '<div class="goh-db-label"><span>DB %' + escape(coverage) + '</span><strong class="' + (row.dbPercent < 0 ? 'goh-negative' : '') + '">' + escape(percent) + '</strong></div>'
                    + '<div class="goh-track goh-seller-db" title="' + escape('DB: ' + percent + coverage) + '"><span class="' + (row.dbPercent < 0 ? 'goh-bar-negative' : '') + '" style="width:' + (Math.abs(row.dbPercent || 0) / maxPercent * 100).toFixed(2) + '%"></span></div></li>';
            }).join('') + '</ol>';
        } else if (widget.kind === 'bar') {
            const max = Math.max(1, ...rows.map(row => Math.abs(row.value)));
            body = '<ol class="goh-bars">' + rows.map(row => '<li><div><span>' + escape(row.label) + '</span><strong class="' + (row.value < 0 ? 'goh-negative' : '') + '">' + escape(format(row.value, row.unit)) + '</strong></div><div class="goh-track" aria-hidden="true"><span class="' + (row.value < 0 ? 'goh-bar-negative' : '') + '" style="width:' + (Math.abs(row.value) / max * 100).toFixed(2) + '%"></span></div></li>').join('') + '</ol>';
        } else {
            body = '<ol class="goh-rows">' + rows.map(row => '<li><button type="button" data-order="' + row.order + '" data-order-module="' + (widget.id === 'recent-orders' ? 'efterkalk' : widget.module) + '"' + (widget.id === 'recent-orders' && !context.canAccess('efterkalk') ? ' disabled' : '') + '><span><strong>' + row.order + ' · ' + escape(row.customer) + '</strong><small>' + escape(displayDate(row.date) + (row.resource || row.seller ? ' · ' + (row.resource || row.seller) : '')) + '</small></span>' + (context.canAccess('omsaetning') ? '<b class="' + (row.value < 0 ? 'goh-negative' : '') + '">' + escape(format(row.value)) + '</b>' : '') + '</button></li>').join('') + '</ol>';
        }
        const viaWarning = widget.module === 'salgordre-via'
            ? (result.via.some(row => !row.costDataAvailable) ? ' · Ufuldstændige kostdata' : '')
                + (context.viaMeta?.unknownCount > 0 ? ' · ' + context.viaMeta.unknownCount + ' ordrer med uafstemt historik udeladt' : '')
                + (context.viaMeta?.scope === 'production' ? ' · Historisk produktionsudvalg' : '')
            : '';
        return header + '<div class="goh-widget-body">' + body + '</div><footer>' + (state === 'error' ? '<strong class="goh-negative">Opdatering fejlede · </strong>' : state === 'loading' ? 'Opdaterer · ' : '') + escape(source + viaWarning) + '</footer>';
    }
    function renderWidgets() {
        if (!context || !context.authenticated) return;
        if (pointerDrag) return;
        const result = data();
        const board = currentBoard();
        const widgets = allowedWidgets(board.widgets, context.canAccess);
        const settings = visibleOptions();
        if (!layoutDraft && context.ensureSources) context.ensureSources([...new Set(widgets.filter(widget => widget.kind !== 'flow').map(widget => widget.id === 'customer-share' && settings[widget.id].metric === 'orders' ? 'efterkalk' : widget.source || widget.module))], widgets.map(widget => widget.id), config.period, query, settings);
        const scopeParts = [];
        if (widgets.some(widget => widget.module === 'omsaetning' && widget.kind !== 'flow')) scopeParts.push('Omsætning: ' + displayDate(result.periodFrom) + ' til indeværende måned · Kunde-filter');
        if (widgets.some(widget => widget.kind === 'flow')) scopeParts.push('Ordreflow: egen måned · Alle omsætningskonti · Kunde-/ordrefilter');
        if (widgets.some(widget => widget.module === 'belastning')) scopeParts.push('Belastning: ' + displayDate(result.today) + ' · Horisont pr. widget: ' + loadHorizons(widgets.map(widget => widget.id), settings).join(' / ') + ' dage · Samlet kapacitet · Uafhængig af økonomiperioden');
        if (widgets.some(widget => widget.id === 'recent-orders')) scopeParts.push('Seneste ordreindgang: ordredato, ikke fakturadato · Egen dagperiode');
        if (widgets.some(widget => widget.module === 'efterkalk')) scopeParts.push('Alle fakturaordrer: ' + displayDate(result.periodFrom) + ' – ' + displayDate(result.today) + ' · Ingen ordrebegrænsning');
        if (widgets.some(widget => widget.module === 'salgordre-via')) scopeParts.push('VIA: aktuel status, uafhængig af fakturaperiode');
        byId('gohDashScope').textContent = scopeParts.join(' | ');
        const grid = byId('gohDashGrid');
        const layout = arrangeWidgets(widgets.map(widget => widget.id), board.layout, board.wide);
        const previousFlow = grid.querySelector('[data-order-flow-host]');
        const flowSettings = JSON.stringify(settings['order-flow'] || {});
        const retainFlow = previousFlow && previousFlow.dataset.scope === context.flowScope && previousFlow.dataset.query === query && previousFlow.dataset.settings === flowSettings && !refreshTask;
        const flowFocus = retainFlow && previousFlow.contains(root.document.activeElement) ? root.document.activeElement : null;
        const scrollPositions = grid.dataset.board === board.id && grid.dataset.query === query
            ? new Map([...grid.querySelectorAll('[data-widget-id]')].map(element => [element.dataset.widgetId, element.querySelector('.goh-widget-body')?.scrollTop || 0])) : new Map();
        grid.innerHTML = widgets.length ? widgets.map(widget => { const widgetResult = data(widget); return '<article class="goh-widget" data-widget-id="' + widget.id + '">' + renderWidget(widget, widgetResult.widgets[widget.id], widgetResult)
            + action('Flyt ' + widget.title + ' (piletaster)', layoutIcon('move'), 'data-layout-handle="move"') + action('Ændr størrelse: ' + widget.title + ' (piletaster)', layoutIcon('resize'), 'data-layout-handle="resize"')
            + '</article>'; }).join('') : '<div class="goh-empty goh-no-widgets">Ingen tilgængelige widgets i denne dashboard.<button type="button" data-dash-action="edit">Tilpas dashboard</button></div>';
        applyGridLayout(layout);
        if (retainFlow && grid.querySelector('[data-order-flow-host]')) grid.querySelector('[data-order-flow-host]').replaceWith(previousFlow);
        grid.dataset.board = board.id;
        grid.dataset.query = query;
        grid.querySelectorAll('[data-widget-id]').forEach(element => { element.querySelector('.goh-widget-body').scrollTop = scrollPositions.get(element.dataset.widgetId) || 0; });
        if (flowFocus?.isConnected) flowFocus.focus({ preventScroll: true });
        const flowHost = byId('gohDashGrid').querySelector('[data-order-flow-host]');
        if (flowHost && !retainFlow) {
            flowHost.dataset.scope = context.flowScope || '';
            flowHost.dataset.query = query;
            flowHost.dataset.settings = flowSettings;
            if (root.GohOrderFlow && context.loadOrderFlow) {
                const current = context;
                root.GohOrderFlow.mount(flowHost, {
                    scope: current.flowScope, month: result.today.slice(0, 7), today: result.today, query: [query, settings['order-flow']?.search].filter(Boolean).join(' '), pageSize: settings['order-flow']?.limit || 8,
                    initialState: config.flow,
                    onStateChange: (state, debounce) => {
                        if (!context?.authenticated || context.flowScope !== current.flowScope) return;
                        config.flow = state;
                        clearTimeout(flowSaveTimer);
                        if (debounce) flowSaveTimer = setTimeout(() => { flowSaveTimer = null; persist(); }, 450);
                        else { flowSaveTimer = null; persist(); }
                    },
                    load: current.loadOrderFlow,
                    isActive: () => context.authenticated && context.flowScope === current.flowScope && context.canAccess('omsaetning') && preferenceStatus !== 'loading',
                    openOrder: current.canAccess('efterkalk') ? ordNo => current.openOrder(ordNo, 'efterkalk') : null
                });
            } else flowHost.textContent = 'Genstart GOH for at aktivere ordreflow.';
        }
        renderRefreshControl();
    }
    function applyGridLayout(layout) {
        const grid = byId('gohDashGrid');
        grid.classList.add('goh-positioned-grid');
        grid.querySelectorAll('[data-widget-id]').forEach(element => {
            const rect = layout[element.dataset.widgetId];
            if (!rect) return;
            element.style.gridColumn = (rect.x + 1) + ' / span ' + rect.w;
            element.style.gridRow = (rect.y + 1) + ' / span ' + rect.h;
        });
    }
    function startLayout(render = true) {
        if (preferenceStatus === 'loading') return false;
        const board = currentBoard();
        layoutDraft = { ...board, widgets: board.widgets.slice(), wide: (board.wide || []).slice(), layout: { ...board.layout, ...arrangeWidgets(allowedWidgets(board.widgets, context.canAccess).map(widget => widget.id), board.layout, board.wide) } };
        if (render) { renderControls(); renderWidgets(); notify('Layout redigeres · Gem eller annuller'); }
        return true;
    }
    function finishLayout(save) {
        if (!layoutDraft) return;
        if (save) saveBoard(layoutDraft);
        layoutDraft = null;
        renderControls(); renderWidgets();
        if (!save) notify('Layoutændringer annulleret');
    }
    function changeLayout(id, requested) {
        const ids = allowedWidgets(layoutDraft.widgets, context.canAccess).map(widget => widget.id);
        layoutDraft.layout = { ...layoutDraft.layout, ...arrangeWidgets(ids, { ...layoutDraft.layout, [id]: requested }, layoutDraft.wide, id, true) };
        applyGridLayout(layoutDraft.layout);
    }
    function startLayoutPointer(event) {
        const handle = event.target.closest('[data-layout-handle]') || (!event.target.closest('button, input, select, a') && event.target.closest('.goh-widget > header'));
        if (!handle || event.button !== 0 || pointerDrag || preferenceStatus === 'loading') return;
        const automatic = !layoutDraft;
        if (automatic && !startLayout(false)) return;
        const operation = handle.dataset.layoutHandle || 'move';
        event.preventDefault();
        const id = handle.closest('[data-widget-id]').dataset.widgetId;
        const rect = layoutDraft.layout[id];
        const bounds = byId('gohDashGrid').getBoundingClientRect();
        const start = { x: event.clientX, y: event.clientY, scroll: root.scrollY, rect: { ...rect }, layout: structuredClone(layoutDraft.layout) };
        pointerDrag = { handle, pointerId: event.pointerId };
        handle.setPointerCapture(event.pointerId);
        handle.closest('article').classList.add('goh-dragging');
        const move = current => {
            if (!layoutDraft || !pointerDrag) return;
            const columns = Math.round((current.clientX - start.x) / ((bounds.width + 16) / 12));
            const rows = Math.round((current.clientY - start.y + root.scrollY - start.scroll) / 40);
            const next = operation === 'move' ? { ...start.rect, x: start.rect.x + columns, y: start.rect.y + rows }
                : { ...start.rect, w: Math.min(12 - start.rect.x, start.rect.w + columns), h: start.rect.h + rows };
            layoutDraft.layout = structuredClone(start.layout);
            changeLayout(id, next);
            if (current.clientY > root.innerHeight - 48) root.scrollBy(0, 16);
            else if (current.clientY < 80) root.scrollBy(0, -16);
        };
        const end = current => {
            handle.removeEventListener('pointermove', move);
            handle.removeEventListener('pointerup', end);
            handle.removeEventListener('pointercancel', end);
            if (handle.hasPointerCapture(event.pointerId)) handle.releasePointerCapture(event.pointerId);
            pointerDrag = null;
            if (!layoutDraft) return;
            if (current.type === 'pointercancel') layoutDraft.layout = start.layout;
            if (automatic) finishLayout(current.type !== 'pointercancel' && JSON.stringify(layoutDraft.layout) !== JSON.stringify(start.layout));
            else renderWidgets();
            byId('gohDashGrid').querySelector('[data-widget-id="' + id + '"] [data-layout-handle="' + operation + '"]')?.focus();
        };
        handle.addEventListener('pointermove', move);
        handle.addEventListener('pointerup', end);
        handle.addEventListener('pointercancel', end);
        pointerDrag.stop = () => {
            handle.removeEventListener('pointermove', move);
            handle.removeEventListener('pointerup', end);
            handle.removeEventListener('pointercancel', end);
            if (handle.hasPointerCapture(event.pointerId)) handle.releasePointerCapture(event.pointerId);
        };
    }
    function onLayoutKey(event) {
        if (pointerDrag || preferenceStatus === 'loading') return;
        if (event.target.matches('[data-donut-segment]') && ['Enter', ' '].includes(event.key)) { event.preventDefault(); selectCustomerSegment(event.target); return; }
        if (event.key === 'Escape' && layoutDraft) { event.preventDefault(); finishLayout(false); return; }
        const handle = event.target.closest('[data-layout-handle]');
        if (!handle || !['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
        const automatic = !layoutDraft;
        if (automatic && !startLayout(false)) return;
        event.preventDefault();
        const id = handle.closest('[data-widget-id]').dataset.widgetId;
        const rect = { ...layoutDraft.layout[id] };
        const horizontal = ['ArrowLeft', 'ArrowRight'].includes(event.key);
        const delta = ['ArrowLeft', 'ArrowUp'].includes(event.key) ? -1 : 1;
        const field = handle.dataset.layoutHandle === 'move' ? (horizontal ? 'x' : 'y') : (horizontal ? 'w' : 'h');
        rect[field] += delta;
        if (field === 'w') rect.w = Math.min(12 - rect.x, rect.w);
        changeLayout(id, rect);
        if (automatic) {
            finishLayout(true);
            byId('gohDashGrid').querySelector('[data-widget-id="' + id + '"] [data-layout-handle="' + handle.dataset.layoutHandle + '"]')?.focus();
        }
    }
    function openEditor(isNew, button) {
        opener = button;
        const board = currentBoard();
        const existing = !isNew && config.boards.some(item => item.id === board.id);
        if (isNew && config.boards.filter(item => item.id.startsWith('custom-')).length >= 12) { notify('Maks. 12 personlige dashboards. Slet en dashboard først.', true); return; }
        const newId = () => typeof root.crypto.randomUUID === 'function' ? root.crypto.randomUUID() : Array.from(root.crypto.getRandomValues(new Uint32Array(4)), value => value.toString(16)).join('-');
        draft = { id: isNew ? 'custom-' + newId() : board.id, name: isNew ? board.name + ' · min visning' : board.name, widgets: board.widgets.slice(), wide: (board.wide || []).slice(), layout: structuredClone(board.layout || {}), options: structuredClone(board.options || {}), existing };
        byId('gohDashName').value = draft.name;
        byId('gohDashBase').value = 'current';
        byId('gohDashEditorError').textContent = '';
        byId('gohDashEditorTitle').textContent = isNew ? 'Ny personlig dashboard' : 'Rediger min dashboard';
        byId('gohDashEditor').querySelector('[data-dash-action="delete"]').hidden = !existing;
        byId('gohDashEditor').querySelector('[data-dash-action="delete"]').textContent = templates.some(item => item.id === draft.id) ? 'Nulstil min dashboard' : 'Slet dashboard';
        renderEditor();
        byId('gohDashEditor').showModal();
        byId('gohDashName').focus();
    }
    function renderEditor(focusSelector) {
        if (!draft) return;
        byId('gohDashCatalog').innerHTML = catalog.filter(permitted).map(widget => '<label><input type="checkbox" data-widget-toggle="' + widget.id + '"' + (draft.widgets.includes(widget.id) ? ' checked' : '') + '><span>' + escape(widget.title) + '<small>' + moduleLabel(widget.module) + '</small></span></label>').join('') || '<p>Ingen widgets for dine modulrettigheder.</p>';
        const selected = allowedWidgets(draft.widgets, context.canAccess);
        byId('gohDashSelected').innerHTML = selected.map((widget, index) => '<div class="goh-editor-row"><strong>' + (index + 1) + '. ' + escape(widget.title) + '</strong><label><input type="checkbox" data-widget-wide="' + widget.id + '"' + (draft.wide.includes(widget.id) ? ' checked' : '') + '> Bred</label>'
            + action('Flyt op: ' + widget.title, '&#8593;', 'data-move="up" data-widget="' + widget.id + '"' + (index === 0 ? ' disabled' : ''))
            + action('Flyt ned: ' + widget.title, '&#8595;', 'data-move="down" data-widget="' + widget.id + '"' + (index === selected.length - 1 ? ' disabled' : '')) + '</div>').join('') || '<div class="goh-empty">Ingen widgets valgt</div>';
        if (focusSelector) byId('gohDashEditor').querySelector(focusSelector)?.focus();
    }
    function saveDraft(event) {
        event.preventDefault();
        if (!draft || !context.authenticated) return;
        const name = byId('gohDashName').value.trim();
        if (!name || !allowedWidgets(draft.widgets, context.canAccess).length) { byId('gohDashEditorError').textContent = 'Angiv et navn og vælg mindst én tilgængelig widget.'; return; }
        const board = { id: draft.id, name, widgets: draft.widgets, wide: draft.wide, layout: draft.layout, options: draft.options };
        saveBoard(board);
        byId('gohDashEditor').close();
        renderControls();
        renderWidgets();
    }
    function onChange(event) {
        if (preferenceStatus === 'loading') return;
        const target = event.target;
        if (target.id === 'gohDashPeriod' || target.id === 'gohDashLimit') { config.period = byId('gohDashPeriod').value; config.limit = Number(byId('gohDashLimit').value); persist(); renderWidgets(); }
        if (!draft) return;
        if (target.id === 'gohDashBase' && target.value !== 'current') { draft.widgets = templates.find(item => item.id === target.value)?.widgets.slice() || []; draft.wide = []; draft.layout = {}; renderEditor(); }
        const id = target.dataset.widgetToggle;
        if (id && permitted(catalog.find(widget => widget.id === id))) { draft.widgets = draft.widgets.filter(value => value !== id); if (target.checked) draft.widgets.push(id); renderEditor('[data-widget-toggle="' + id + '"]'); }
        const wide = target.dataset.widgetWide;
        if (wide) { draft.wide = draft.wide.filter(value => value !== wide); if (target.checked) draft.wide.push(wide); delete draft.layout[wide]; }
    }
    function onClick(event) {
        const button = event.target.closest('button, [data-donut-segment]');
        if (!button || !context || !context.authenticated) return;
        const attrs = button.dataset;
        if (attrs.dashAction === 'retry-preferences') {
            const discard = preferenceStatus === 'conflict';
            if (!discard || root.confirm('Hent profilen fra GOH og erstat de lokale, ikke-gemte ændringer?')) preferenceStore?.load(discard);
            return;
        }
        if (preferenceStatus === 'loading') return;
        if (attrs.donutSegment !== undefined) { selectCustomerSegment(button); return; }
        if (attrs.widgetSettings) openWidgetSettings(attrs.widgetSettings, button);
        if (attrs.dashAction === 'cancel-widget') byId('gohWidgetSettings').close();
        if (attrs.dashAction === 'reset-widget' && widgetDraft) { widgetDraft.options = widgetOptions(widgetDraft.id); renderWidgetSettings(); }
        if (attrs.dashAction === 'arrange') startLayout();
        if (attrs.dashAction === 'save-layout') finishLayout(true);
        if (attrs.dashAction === 'cancel-layout') finishLayout(false);
        if (attrs.dashAction === 'compact' && layoutDraft) { layoutDraft.layout = { ...layoutDraft.layout, ...arrangeWidgets(allowedWidgets(layoutDraft.widgets, context.canAccess).map(widget => widget.id), layoutDraft.layout, layoutDraft.wide, '', true) }; renderWidgets(); }
        if (attrs.template && availableBoards(config, context.canAccess).some(item => item.id === attrs.template)) { config.active = attrs.template; persist(); renderControls(); renderWidgets(); byId('gohDashTemplates').querySelector('[data-template="' + attrs.template + '"]').focus(); }
        if (attrs.openModule && context.canAccess(attrs.openModule)) context.openModule(attrs.openModule);
        if (attrs.order && context.canAccess(attrs.orderModule)) context.openOrder(Number(attrs.order), attrs.orderModule);
        if (attrs.move && draft) {
            const visible = allowedWidgets(draft.widgets, context.canAccess).map(widget => widget.id);
            const adjacent = visible[visible.indexOf(attrs.widget) + (attrs.move === 'up' ? -1 : 1)];
            const index = draft.widgets.indexOf(attrs.widget);
            const destination = draft.widgets.indexOf(adjacent);
            if (destination >= 0) { [draft.widgets[index], draft.widgets[destination]] = [draft.widgets[destination], draft.widgets[index]]; draft.layout = {}; renderEditor('[data-widget="' + attrs.widget + '"][data-move="' + attrs.move + '"]:not(:disabled), [data-widget-wide="' + attrs.widget + '"]'); }
        }
        if (attrs.dashAction === 'create' || attrs.dashAction === 'edit') openEditor(attrs.dashAction === 'create', button);
        if (attrs.dashAction === 'cancel') byId('gohDashEditor').close();
        if (attrs.dashAction === 'delete' && draft && draft.existing) {
            const standard = templates.some(item => item.id === draft.id);
            if (root.confirm(standard ? 'Nulstil kun din visning af "' + draft.name + '"?' : 'Slet dashboard "' + draft.name + '"?')) {
                config.boards = config.boards.filter(board => board.id !== draft.id);
                config.active = standard ? draft.id : 'management';
                persist(); byId('gohDashEditor').close(); renderControls(); renderWidgets();
            }
        }
        if (attrs.dashAction === 'export') exportCsv();
        if (attrs.dashAction === 'refresh') refreshData();
    }
    function exportCsv() {
        const result = data();
        const rows = [['Dashboard', 'Widget', 'Ordre / kategori', 'Kunde', 'Værdi', 'Enhed', 'Kapacitet (timer)', 'Aften (timer)', 'DB (%)', 'Ordrer med kendt kost', 'Ordrer i alt']];
        for (const widget of allowedWidgets(currentBoard().widgets, context.canAccess)) {
            const result = data(widget);
            if (widget.kind === 'flow') continue;
            if (!sourceRows(widget, result).length) continue;
            const hideMoney = widget.module === 'salgordre-via' && !context.canAccess('omsaetning');
            const options = optionsFor(widget);
            const selected = result.customerShares.segments.find(segment => segment.key === options.selected);
            const visible = widget.kind === 'donut' ? (selected ? selected.members : options.selected === '__negative' ? result.customerShares.rows.filter(row => row.value < 0) : result.customerShares.rows) : result.widgets[widget.id];
            for (const row of visible) rows.push([currentBoard().name, options.title || widget.title, row.order || row.label, row.customer || '', hideMoney || row.value === null ? '' : row.value, hideMoney ? '' : row.unit || 'DKK', row.capacity ?? '', row.evening ?? '', widget.id === 'sellers' ? row.dbPercent ?? '' : '', widget.id === 'sellers' ? row.costCount : '', widget.id === 'sellers' ? row.count : '']);
        }
        const csv = '\ufeff' + rows.map(row => row.map(value => '"' + (typeof value === 'number' ? String(value).replace('.', ',') : text(value).replace(/^[=+\-@\t\r]/, match => "'" + match)).replace(/"/g, '""') + '"').join(';')).join('\r\n');
        const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
        const link = root.document.createElement('a');
        link.href = url; link.download = 'GOH-dashboard-' + result.today + '.csv'; link.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
    function update(next) {
        context = next;
        const host = byId('dashboardOperationalOverview');
        if (!host) return;
        if (!next.authenticated || !next.username) { reset(); return; }
        if (typeof next.ensureSources !== 'function' || typeof next.loadPreferences !== 'function' || typeof next.savePreferences !== 'function') {
            reset();
            host.textContent = 'Dashboard-opdatering afventer genstart af GOH-serveren. Genstart GOH, og genindlæs siden.';
            return;
        }
        const key = storageKey(next.username, next.database);
        if (scope !== key) {
            reset(); context = next; scope = key;
            let legacy, cached;
            try { legacy = JSON.parse(root.localStorage.getItem(key) || 'null'); cached = JSON.parse(root.localStorage.getItem(key + ':goh') || 'null'); } catch (_) { legacy = null; cached = null; }
            mount(host);
            preferenceStore = createPreferenceStore({
                cached, legacy, schemaVersion: 4, initial: { active: allowedTemplates(next.canAccess)[0]?.id || 'production' },
                load: next.loadPreferences, save: next.savePreferences,
                cache: value => {
                    try { root.localStorage.setItem(key + ':goh', JSON.stringify(value)); localRecoveryAvailable = true; }
                    catch (_) { localRecoveryAvailable = false; }
                },
                onStatus: (value, error) => { renderPreferenceStatus(value); if (error?.message?.startsWith('Genstart GOH')) notify(error.message, true); renderControls(); },
                onConfig: value => {
                    config = normalizeConfig(value); query = config.query; preferencesReady = true;
                    root.GohTheme?.bind(config.theme, saveTheme, preferenceStatus);
                    byId('gohDashSearch').value = query;
                    byId('gohDashGrid').replaceChildren();
                    renderControls(); renderWidgets();
                }
            });
            preferenceStore.load();
            return;
        }
        if (!preferencesReady) return;
        if (!byId('gohDashGrid')) mount(host);
        renderControls(); renderWidgets();
    }
    function reset() {
        flushPendingPreferences();
        clearTimeout(searchTimer);
        clearTimeout(flowSaveTimer);
        searchTimer = null;
        flowSaveTimer = null;
        root.removeEventListener('pagehide', flushPendingPreferences);
        preferenceStore?.dispose();
        preferenceStore = null;
        preferencesReady = false;
        preferenceStatus = 'loading';
        root.GohTheme?.bind('light', null);
        localRecoveryAvailable = true;
        layoutDraft = null;
        if (pointerDrag) pointerDrag.stop();
        pointerDrag = null;
        refreshTask = null;
        const host = byId('dashboardOperationalOverview');
        if (byId('gohDashEditor')?.open) byId('gohDashEditor').close();
        if (byId('gohWidgetSettings')?.open) byId('gohWidgetSettings').close();
        widgetDraft = null;
        if (host) { host.replaceChildren(); host.removeEventListener('click', onClick); host.removeEventListener('change', onChange); host.removeEventListener('pointerdown', startLayoutPointer); host.removeEventListener('keydown', onLayoutKey); }
        scope = ''; query = ''; draft = null; config = normalizeConfig();
    }
    root.GohDashboard.update = update;
    root.GohDashboard.reset = reset;
})(typeof window === 'undefined' ? globalThis : window);