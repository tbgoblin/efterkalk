(function (root) {
    'use strict';

    function build(rawOrders, history, firstDate) {
        const histories = new Map(history.map(row => [Number(row.OrdNo), row]));
        const rows = [];
        let unknownCount = 0;
        for (const order of rawOrders) {
            const value = Number(order.OrderValueDkk || 0);
            if (value <= 0) continue;
            const invoiceHistory = histories.get(Number(order.OrdNo));
            const currentInvoiced = Number(order.InvoicedDkk || 0);
            if ((!invoiceHistory && Math.abs(currentInvoiced) > 0.01)
                || (invoiceHistory && Math.abs(Number(invoiceHistory.InvoicedTotalDkk || 0) - currentInvoiced) > 1)) {
                unknownCount++;
                continue;
            }
            const received = Number(order.OrderDate) >= firstDate;
            const before = Number(invoiceHistory && invoiceHistory.InvoicedBeforeDkk || 0);
            const invoiced = Number(invoiceHistory && invoiceHistory.InvoicedInMonthDkk || 0);
            const opening = received ? 0 : Math.max(0, value - before);
            const closing = Math.max(0, value - before - invoiced);
            if (!received && opening <= 0.01 && closing <= 0.01 && Math.abs(invoiced) <= 0.01) continue;
            rows.push({
                ordNo: Number(order.OrdNo), orderDate: Number(order.OrderDate),
                customerName: String(order.CustomerName || '').trim(), custNo: Number(order.CustNo || 0),
                cohort: received ? 'new' : 'prior', value, opening,
                incoming: received ? value : 0, invoiced, closing,
                current: Math.max(0, Number(order.RemainingDkk || 0)),
                adjustment: closing - opening - (received ? value : 0) + invoiced
            });
        }
        return { rows, unknownCount, ...summarize(rows) };
    }

    function summarize(rows) {
        const total = selected => {
            const result = { count: selected.length, completed: 0, opening: 0, incoming: 0, invoiced: 0, closing: 0, current: 0, adjustment: 0 };
            for (const row of selected) {
                for (const key of ['opening', 'incoming', 'invoiced', 'closing', 'current', 'adjustment']) result[key] += row[key];
                if (row.closing <= 0.01 && row.invoiced > 0) result.completed++;
            }
            return result;
        };
        return { total: total(rows), prior: total(rows.filter(row => row.cohort === 'prior')), received: total(rows.filter(row => row.cohort === 'new')) };
    }

    const preferences = new Map();
    const escape = value => String(value == null ? '' : value).replace(/[&<>"']/g, character => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]));
    const money = value => new Intl.NumberFormat('da-DK', { maximumFractionDigits: 2 }).format(value);
    const shortMoney = value => new Intl.NumberFormat('da-DK', { maximumFractionDigits: 2 }).format(value / 1000000) + ' mio.';
    const labels = { opening: 'Primo', incoming: 'Tilgang', invoiced: 'Faktureret', closing: 'Ultimo' };
    const icon = name => '<svg class="lucide" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + ({
        previous: '<path d="m15 18-6-6 6-6"/>', next: '<path d="m9 18 6-6-6-6"/>',
        download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/>',
        refresh: '<path d="M3 12a9 9 0 0 1 15.7-6.3L21 8M21 3v5h-5M21 12a9 9 0 0 1-15.7 6.3L3 16M8 16H3v5"/>'
    }[name] || '') + '</svg>';

    function mount(host, options) {
        const stateKey = options.scope;
        let state = preferences.get(stateKey);
        if (!state || options.initialState || options.fixedMonth && state.month !== options.month) {
            state = { month: options.month, cohort: 'all', metric: 'closing', query: '', page: 0, completedOnly: false, ...options.initialState };
            if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(state.month) || state.month > options.today.slice(0, 7) || options.fixedMonth) state.month = options.month;
            preferences.set(stateKey, state);
        }
        let payload = null;
        let token = 0;
        let visibleRows = [];
        const active = () => host.isConnected && (!options.isActive || options.isActive());
        host.className = 'goh-order-flow';
        host.innerHTML = '<div class="flow-toolbar"><label>M\u00e5ned<input type="month" data-flow-month value="' + escape(state.month) + '" max="' + escape(options.today.slice(0, 7)) + '"' + (options.fixedMonth ? ' disabled' : '') + '></label><div class="flow-cohorts" role="group" aria-label="Ordregruppe">'
            + [['all', 'Alle'], ['prior', 'Tidligere'], ['new', 'Nye']].map(([key, label]) => '<button type="button" data-flow-cohort="' + key + '">' + label + '</button>').join('')
            + '</div><button type="button" class="flow-icon" data-flow-refresh title="Opdater ordreflow" aria-label="Opdater ordreflow">' + icon('refresh') + '</button></div>'
            + '<div class="flow-status" role="status" aria-live="polite"></div><div class="flow-content"></div>';
        const content = host.querySelector('.flow-content');
        const status = host.querySelector('.flow-status');

        function filtered() {
            const query = (state.query + ' ' + (options.query || '')).trim().toLocaleLowerCase('da-DK').split(/\s+/).filter(Boolean);
            return payload.rows.filter(row => (state.cohort === 'all' || row.cohort === state.cohort)
                && query.every(part => (row.ordNo + ' ' + row.customerName + ' ' + row.custNo).toLocaleLowerCase('da-DK').includes(part)));
        }

        function renderTable(rows) {
            visibleRows = rows.filter(row => Math.abs(row[state.metric]) > 0.01 && (!state.completedOnly || row.closing <= 0.01 && row.invoiced > 0)).sort((left, right) => right[state.metric] - left[state.metric] || left.ordNo - right.ordNo);
            const pageSize = [5, 10, 20, 50].includes(options.pageSize) ? options.pageSize : 8;
            state.page = Math.min(state.page, Math.max(0, Math.ceil(visibleRows.length / pageSize) - 1));
            const pageRows = visibleRows.slice(state.page * pageSize, (state.page + 1) * pageSize);
            host.querySelector('.flow-table').innerHTML = '<table><caption>' + labels[state.metric] + ' \u00b7 ' + visibleRows.length + ' ordrer' + (state.completedOnly ? ' \u00b7 f\u00e6rdigfaktureret' : '') + '</caption><thead><tr><th>Ordre / kunde</th><th>Gruppe</th><th>' + labels[state.metric] + ' DKK</th></tr></thead><tbody>'
                + pageRows.map(row => '<tr><td>' + (options.openOrder ? '<button type="button" data-flow-order="' + row.ordNo + '">' + row.ordNo + '</button>' : '<strong>' + row.ordNo + '</strong>') + '<span>' + escape(row.customerName || row.custNo) + '</span></td><td>' + (row.cohort === 'prior' ? 'Tidligere' : 'Ny') + '</td><td>' + money(row[state.metric]) + '</td></tr>').join('')
                + (pageRows.length ? '' : '<tr><td colspan="3">Ingen ordrer i dette udsnit.</td></tr>') + '</tbody></table>'
                + '<div class="flow-pagination"><span>' + (visibleRows.length ? state.page * pageSize + 1 : 0) + '\u2013' + Math.min((state.page + 1) * pageSize, visibleRows.length) + ' af ' + visibleRows.length + '</span><button type="button" class="flow-icon" data-flow-page="-1" aria-label="Forrige ordrer" title="Forrige ordrer"' + (state.page === 0 ? ' disabled' : '') + '>' + icon('previous') + '</button><button type="button" class="flow-icon" data-flow-page="1" aria-label="N\u00e6ste ordrer" title="N\u00e6ste ordrer"' + ((state.page + 1) * pageSize >= visibleRows.length ? ' disabled' : '') + '>' + icon('next') + '</button></div>';
        }

        function render() {
            if (!payload || !active()) return;
            const rows = filtered();
            const summary = summarize(rows);
            const total = summary.total;
            host.querySelectorAll('[data-flow-cohort]').forEach(button => button.setAttribute('aria-pressed', String(button.dataset.flowCohort === state.cohort)));
            const max = Math.max(1, ...Object.keys(labels).map(key => Math.abs(summary.prior[key]) + Math.abs(summary.received[key])));
            const cutoff = String(payload.asOf);
            const date = cutoff.slice(6, 8) + '.' + cutoff.slice(4, 6) + '.' + cutoff.slice(0, 4);
            if (!content.querySelector('.flow-chart')) {
                content.innerHTML = '<div class="flow-headline"></div><div class="flow-legend"><span class="flow-prior">Ordrer fra tidligere m\u00e5neder</span><span class="flow-new">Ordrer modtaget i m\u00e5neden</span></div><div class="flow-chart" role="group" aria-label="Ordreflow i DKK"></div><div class="flow-summary"></div><div class="flow-search"><label>Kunde / ordre<input type="search" data-flow-query maxlength="120" value="' + escape(state.query) + '"></label><button type="button" class="flow-icon" data-flow-export aria-label="Eksport\u00e9r viste ordrer" title="Eksport\u00e9r viste ordrer">' + icon('download') + '</button></div><label class="flow-completed"><input type="checkbox" data-flow-completed' + (state.completedOnly ? ' checked' : '') + '>Kun f\u00e6rdigfakturerede ordrer</label><div class="flow-table"></div><p class="flow-basis"></p>';
            }
            content.querySelector('.flow-headline').innerHTML = '<div><span>Ordrebeholdning pr. ' + date + '</span><strong>' + shortMoney(total.closing) + ' <small>DKK</small></strong></div><div><span>F\u00e6rdigfaktureret i m\u00e5neden</span><b>' + total.completed + ' <small>ordrer</small></b><span>Heraf ' + summary.received.completed + ' nye ordrer</span></div>';
            content.querySelector('.flow-chart').innerHTML = Object.entries(labels).map(([key, label]) => {
                const segments = [['prior', summary.prior[key]], ['new', summary.received[key]]];
                return '<button type="button" class="flow-column" data-flow-metric="' + key + '" aria-pressed="' + (state.metric === key) + '" title="' + label + ': ' + money(total[key]) + ' DKK" aria-label="' + label + ', ' + money(total[key]) + ' DKK"><strong>' + shortMoney(total[key]) + '</strong><span class="flow-plot">'
                    + segments.map(([cohort, amount]) => '<span class="flow-segment flow-' + cohort + (amount < 0 ? ' flow-negative' : '') + '" style="height:' + (Math.abs(amount) / max * 100).toFixed(3) + '%" title="' + (cohort === 'prior' ? 'Tidligere' : 'Nye') + ': ' + money(amount) + ' DKK"></span>').join('') + '</span><span class="flow-column-label">' + label + '</span></button>';
            }).join('');
            content.querySelector('.flow-summary').innerHTML = '<span>Primo + tilgang \u2212 faktureret' + (Math.abs(total.adjustment) > 0.01 ? ' + regulering ' + money(total.adjustment) + ' DKK' : '') + ' = ultimo</span><span>Rest i dag, disse ordrer: <strong>' + money(total.current) + ' DKK</strong></span>';
            content.querySelector('.flow-basis').textContent = 'Salgspriser, ikke VIA-kost. Alle oms\u00e6tningskonti; valgt kundeafgr\u00e6nsning. Historik rekonstrueret fra bogf\u00f8ring og aktuelle ordrev\u00e6rdier.';
            status.textContent = payload.unknownCount ? payload.unknownCount + ' ordrer uden afstemt fakturahistorik er udeladt. Viste bel\u00f8b er delsummer.' : '';
            status.dataset.warning = String(payload.unknownCount > 0);
            renderTable(rows);
        }

        async function load(force = false) {
            const requestToken = ++token;
            payload = null;
            content.replaceChildren();
            host.setAttribute('aria-busy', 'true');
            status.textContent = 'Henter ordreflow...';
            host.querySelector('[data-flow-refresh]').disabled = true;
            try {
                const result = await options.load(state.month, force);
                if (requestToken !== token || !active()) return;
                if (!result || !Array.isArray(result.rows) || !result.total) throw new Error('Genstart GOH for at hente ordreflow.');
                payload = result;
                render();
            } catch (error) {
                if (requestToken === token && active()) status.textContent = 'Ordreflow kunne ikke hentes: ' + error.message;
            } finally {
                if (requestToken === token && active()) {
                    host.setAttribute('aria-busy', 'false');
                    host.querySelector('[data-flow-refresh]').disabled = false;
                }
            }
        }

        host.addEventListener('click', event => {
            const button = event.target.closest('button');
            if (!button || !active()) return;
            if (button.dataset.flowCohort) { state.cohort = button.dataset.flowCohort; state.page = 0; render(); }
            if (button.dataset.flowMetric) { state.metric = button.dataset.flowMetric; state.completedOnly = false; state.page = 0; host.querySelector('[data-flow-completed]').checked = false; render(); host.querySelector('[data-flow-metric="' + state.metric + '"]')?.focus(); }
            if (button.dataset.flowPage) { state.page += Number(button.dataset.flowPage); renderTable(filtered()); }
            if (button.dataset.flowCohort || button.dataset.flowMetric || button.dataset.flowPage) options.onStateChange?.({ ...state }, false);
            if (button.hasAttribute('data-flow-refresh')) load(true);
            if (button.dataset.flowOrder && options.openOrder) options.openOrder(Number(button.dataset.flowOrder));
            if (button.hasAttribute('data-flow-export') && payload) {
                const csvRows = [['Ordre', 'Kunde', 'Gruppe', 'Primo DKK', 'Tilgang DKK', 'Faktureret DKK', 'Ultimo DKK', 'Rest i dag DKK'], ...visibleRows.map(row => [row.ordNo, row.customerName, row.cohort === 'new' ? 'Ny' : 'Tidligere', row.opening, row.incoming, row.invoiced, row.closing, row.current])];
                const csv = '\ufeff' + csvRows.map(row => row.map(value => '"' + (typeof value === 'number' ? String(value).replace('.', ',') : String(value).replace(/^[=+\-@\t\r]/, character => "'" + character)).replace(/"/g, '""') + '"').join(';')).join('\r\n');
                const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
                const link = root.document.createElement('a');
                link.href = url; link.download = 'Ordreflow-' + state.month + '.csv'; link.click();
                setTimeout(() => URL.revokeObjectURL(url), 1000);
            }
        });
        host.addEventListener('change', event => {
            if (!active()) return;
            if (event.target.hasAttribute('data-flow-completed')) {
                state.completedOnly = event.target.checked;
                if (state.completedOnly) state.metric = 'invoiced';
                state.page = 0; render();
            }
            if (event.target.hasAttribute('data-flow-month')) {
                const value = event.target.value;
                if (!/^\d{4}-\d{2}$/.test(value) || value > options.today.slice(0, 7)) { event.target.value = state.month; return; }
                state.month = value; state.page = 0; load();
            }
            if (event.target.hasAttribute('data-flow-completed') || event.target.hasAttribute('data-flow-month')) options.onStateChange?.({ ...state }, false);
        });
        host.addEventListener('input', event => {
            if (!active()) return;
            if (event.target.hasAttribute('data-flow-query')) { state.query = event.target.value; state.page = 0; render(); options.onStateChange?.({ ...state }, true); }
        });
        return load();
    }

    const api = { build, summarize, mount, selectedMonth: (scope, fallback) => preferences.get(scope)?.month || fallback, reset: () => preferences.clear() };
    if (typeof module === 'object' && module.exports) module.exports = api;
    else root.GohOrderFlow = api;
}(typeof window === 'undefined' ? globalThis : window));