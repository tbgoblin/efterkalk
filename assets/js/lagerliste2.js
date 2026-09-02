(function () {
    const state = { routes: [], unassignedRest: [], periods: new Map(), comparison: null, reservations: [], reservationSummary: {}, currentValuation: null };
    const byId = id => document.getElementById(id);
    const fmt = value => new Intl.NumberFormat('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(Number(value || 0));
    const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));

    async function fetchJson(url, options) {
        const response = await fetch(url, options);
        const data = await response.json().catch(() => ({}));
        if (!response.ok || data.ok === false) {
            const error = new Error(data.error || ('HTTP ' + response.status));
            error.status = response.status;
            throw error;
        }
        return data;
    }

    function authMessage(error) {
        return error && (error.status === 401 || error.status === 403)
            ? 'Login eller Lagerliste-adgang kræves. Gå tilbage til Operations Hub og log ind.'
            : String(error && error.message || error);
    }

    function statusMeta(status) {
        if (status === 'completed') return { icon: '✓', text: 'Færdig', cls: 'done' };
        if (status === 'partial') return { icon: '◐', text: 'Delvist færdig', cls: 'partial' };
        if (status === 'not_started') return { icon: '⏳', text: 'Ikke startet', cls: 'open' };
        return { icon: '?', text: 'Ukendt', cls: 'unknown' };
    }

    function itemList(rows, status, kind) {
        const meta = statusMeta(status);
        if (!rows || !rows.length) return '<span class="zero">–</span>';
        return '<div class="items">' + rows.map(row => {
            const qty = kind === 'product'
                ? (fmt(row.finishedQty) + ' / ' + fmt(row.plannedQty))
                : (kind === 'registeredRest' ? (fmt(row.weight) + ' kg') : fmt(Math.abs(Number(row.finishedQty || 0))));
            const searchRest = kind === 'plate' && row.unregisteredRestSource;
            const missingSalesRef = kind === 'product' && !row.hasSalesReference;
            const estimatedRest = kind === 'estimatedRest';
            const registeredRest = kind === 'registeredRest';
            const badge = estimatedRest
                ? ' <span class="pill info">↩ Forventet REST</span>'
                : (registeredRest ? ' <span class="pill done">✓ REST Plader</span>' : (searchRest
                    ? ' <span class="pill info">♻ Søg-rest</span>'
                    : (missingSalesRef ? ' <span class="pill info">🏭 Ingen R4</span>' : '')));
            const title = estimatedRest
                ? 'Forventet REST fra nestinglinjen' + (row.sourceCode ? ': ' + row.sourceCode : '')
                : (registeredRest ? 'REST faktisk registreret i FreeInf1' : (searchRest
                    ? 'Ikke lagerregistreret restplade: ' + (row.sourceInfo || 'Søg')
                    : (missingSalesRef ? 'Ingen salgsreference; mulig intern produktion eller lagerordre' : meta.text)));
            const icon = estimatedRest ? '↩' : (registeredRest ? '♻' : meta.icon);
            const detail = registeredRest ? (row.restCode || row.label || '') : (estimatedRest ? row.sourceCode : '');
            return '<div class="item" title="' + esc(title) + '">' + icon + ' <strong>' + esc(row.prodNo) + '</strong> <small>' + esc(qty) + '</small>' + badge
                + (detail ? '<br><small>' + esc(detail) + '</small>' : '') + '</div>';
        }).join('') + '</div>';
    }

    function restList(route) {
        const expected = itemList(route.estimatedRestLines || [], route.status, 'estimatedRest');
        const registered = itemList(route.restPlates || [], route.status, 'registeredRest');
        if ((!route.estimatedRestLines || !route.estimatedRestLines.length) && (!route.restPlates || !route.restPlates.length)) return expected;
        return '<div><small>Forventet fra nesting</small>' + expected + '<small>Registreret REST Plader</small>' + registered + '</div>';
    }

    function orderLinks(route) {
        const production = 'PO: ' + (route.productionOrderNos.join(', ') || '–');
        const references = Array.isArray(route.salesOrderReferences) ? route.salesOrderReferences : [];
        const sales = references.length
            ? references.map(ref => '<strong>' + esc(ref.orderNo) + '</strong> <small>(' + esc(ref.source || 'ukendt') + ')</small>').join('<br>')
            : (route.salesOrderNos.length ? route.salesOrderNos.map(esc).join(', ') : '–');
        const missing = Number(route.unlinkedProductCount || 0) > 0
            ? '<br><span class="pill info">🏭 ' + esc(route.unlinkedProductCount) + ' uden R4</span>'
            : '';
        return production + '<br>SO: ' + sales + missing;
    }

    function renderRoutes() {
        const query = String(byId('routeSearch').value || '').trim().toLowerCase();
        const status = byId('routeStatusFilter').value;
        const rows = state.routes.filter(route => {
            if (status === 'active' && route.status === 'completed') return false;
            if (status && status !== 'active' && route.status !== status) return false;
            if (!query) return true;
            const haystack = [route.nestingOrdNo, route.route]
                .concat(route.plates.map(row => row.prodNo + ' ' + row.descr))
                .concat(route.products.map(row => row.prodNo + ' ' + row.descr))
                .concat(route.productionOrderNos, route.salesOrderNos).join(' ').toLowerCase();
            return haystack.includes(query);
        });
        byId('routeRows').innerHTML = rows.length ? rows.map(route => {
            const meta = statusMeta(route.status);
            const links = orderLinks(route);
            const allocationResidual = route.materialAllocationResidual ?? route.residualValue;
            const residualClass = Math.abs(allocationResidual) <= 1 ? 'zero' : 'negative';
            const restStatus = route.restRegistrationStatus === 'missing'
                ? '<br><span class="pill error">⚠ REST ikke registreret</span>'
                : (route.restRegistrationStatus === 'partial'
                    ? '<br><span class="pill error">⚠ REST delvist registreret</span>'
                    : (route.restRegistrationStatus === 'finished_unregistered'
                        ? '<br><span class="pill info">✓ REST færdigmeldt</span><br><small>Ikke fundet i REST-lagerlisten endnu</small>'
                        : (route.restRegistrationStatus === 'finished_partially_registered'
                            ? '<br><span class="pill info">✓ REST færdigmeldt · delvist lagerregistreret</span>'
                            : (route.restRegistrationStatus === 'pending' || route.restRegistrationStatus === 'pending_partial'
                                ? '<br><span class="pill open">⏳ REST ikke færdigmeldt endnu</span>'
                                : ''))));
            return '<tr><td><span class="pill ' + meta.cls + '">' + meta.icon + ' ' + esc(meta.text) + '</span><br><small>' + esc(route.progress) + '%</small></td>'
                + '<td><strong>' + esc(route.nestingOrdNo) + '</strong><br>Route ' + esc(route.route) + '</td>'
                + '<td>' + itemList(route.plates, route.status, 'plate') + '</td>'
                + '<td>' + restList(route) + restStatus + '</td>'
                + '<td>' + itemList(route.products, route.status, 'product') + '</td>'
                + '<td>' + links + '</td>'
                + '<td class="num">' + fmt(route.plateValue) + '</td>'
                + '<td class="num">' + fmt(route.completedProductValue) + '</td>'
                + '<td class="num">' + fmt(route.estimatedRestFifoValue) + '</td>'
                + '<td class="num">' + fmt(route.restValue) + '</td>'
                + '<td class="num">' + fmt(route.restWriteDown) + '</td>'
                + '<td class="num ' + residualClass + '">' + fmt(allocationResidual) + '</td></tr>';
        }).join('') : '<tr><td colspan="12" class="empty">Ingen ruter matcher filteret.</td></tr>';
        byId('routeStatus').textContent = rows.length + ' af ' + state.routes.length + ' ruter vises.';
    }

    function renderUnassignedRest() {
        const panel = byId('unassignedRestPanel');
        const rows = state.unassignedRest;
        panel.classList.toggle('hidden', !rows.length);
        byId('unassignedRestSummary').textContent = rows.length + ' REST-rækker uden sikker routekobling';
        byId('unassignedRestRows').innerHTML = rows.map(row => '<tr><td>' + esc(row.nestingOrdNo || '–') + '</td><td><strong>' + esc(row.prodNo || '–') + '</strong></td>'
            + '<td>' + esc(row.label || '–') + '</td><td><span class="pill open">⚠ ' + esc(row.reason || 'ukendt') + '</span></td>'
            + '<td class="num">' + fmt(row.weight) + '</td><td class="num">' + fmt(row.value) + '</td></tr>').join('');
    }

    async function loadRoutes(force) {
        byId('routeStatus').textContent = 'Indlæser routeforbindelser…';
        byId('refreshRoutesBtn').disabled = true;
        try {
            const data = await fetchJson('/lagerliste2/routes/current' + (force ? '?force=1' : ''));
            state.routes = Array.isArray(data.routes) ? data.routes : [];
            state.unassignedRest = Array.isArray(data.unassignedRest) ? data.unassignedRest : [];
            byId('metricOpen').textContent = state.routes.filter(row => row.status === 'not_started').length;
            byId('metricPartial').textContent = state.routes.filter(row => row.status === 'partial').length;
            byId('metricDone').textContent = state.routes.filter(row => row.status === 'completed').length;
            renderRoutes();
            renderUnassignedRest();
            if (state.unassignedRest.length) {
                byId('routeStatus').textContent += ' ' + state.unassignedRest.length + ' REST-rækker kræver manuel routekobling.';
            }
        } catch (error) {
            state.unassignedRest = [];
            renderUnassignedRest();
            byId('routeStatus').textContent = 'Kunne ikke hente ruter: ' + authMessage(error);
            byId('routeRows').innerHTML = '<tr><td colspan="12" class="empty">Routevisningen er utilgængelig. Lagerliste 1 påvirkes ikke.</td></tr>';
        } finally {
            byId('refreshRoutesBtn').disabled = false;
        }
    }

    function addPeriod(value, label, loader) {
        state.periods.set(value, { label, loader, data: null });
    }

    async function loadPeriod(value) {
        const period = state.periods.get(value);
        if (!period) throw new Error('Ukendt periode');
        if (!period.data) period.data = await period.loader();
        return period.data;
    }

    async function loadPeriods() {
        const [monthsResult, snapshotsResult] = await Promise.allSettled([
            fetchJson('/lagerliste/snapshot-months'),
            fetchJson('/lagerliste/snapshots')
        ]);
        addPeriod('current', 'Aktuel lagerliste', () => fetchJson('/lagerliste/current'));
        if (snapshotsResult.status === 'fulfilled') {
            for (const row of snapshotsResult.value.rows || []) {
                const id = String(row.snapshotId || '');
                if (!id) continue;
                const label = 'Dags-snapshot ' + id.replace('_', ' ');
                addPeriod('snapshot:' + id, label, async () => (await fetchJson('/lagerliste/snapshots/' + encodeURIComponent(id))).snapshot);
            }
        }
        if (monthsResult.status === 'fulfilled') {
            for (const month of monthsResult.value.months || []) {
                addPeriod('month:' + month, 'Måned ' + month, () => fetchJson('/lagerliste/snapshot/' + encodeURIComponent(month)));
            }
        }
        const options = Array.from(state.periods.entries()).map(([value, period]) => '<option value="' + esc(value) + '">' + esc(period.label) + '</option>').join('');
        byId('periodA').innerHTML = options;
        byId('periodB').innerHTML = options;
        byId('periodB').value = 'current';
        const previous = Array.from(state.periods.keys()).find(value => value !== 'current');
        if (previous) byId('periodA').value = previous;
    }

    function reservationStatusMeta(row) {
        if (row.status === 'finished') return { cls: 'done', text: '✓ Færdigmeldt' };
        if (row.status === 'picked') return { cls: 'partial', text: '◐ Plukket' };
        if (row.status === 'mixed') return { cls: 'partial', text: '◐ Delvist plukket' };
        return { cls: 'open', text: '⏳ Reserveret' };
    }

    function renderReservations() {
        const filter = byId('reservationFilter').value;
        const rows = state.reservations.filter(row => filter === 'all'
            || (filter === 'active' && Number(row.activeQty || 0) > 0.005)
            || (filter === 'finished' && Number(row.activeQty || 0) <= 0.005));
        byId('reservationRows').innerHTML = rows.length ? rows.map(row => {
            const meta = reservationStatusMeta(row);
            const salesLink = row.salesOrderNo
                ? ('SO <strong>' + esc(row.salesOrderNo) + '</strong> <small>(' + esc(row.linkSource || 'Rsv') + ')</small>')
                : '<span class="pill info">🏭 Intern/lager?</span>';
            const operationalOrder = Number(row.orderNo || 0) && Number(row.orderNo) !== Number(row.salesOrderNo)
                ? '<br><small>Ordre ' + esc(row.orderNo) + '</small>' : '';
            return '<tr><td><span class="pill ' + meta.cls + '">' + meta.text + '</span></td>'
                + '<td><strong>' + esc(row.prodNo) + '</strong><br><small>' + esc(row.descr || '–') + '</small></td>'
                + '<td>' + salesLink + operationalOrder + '</td><td>' + esc(row.customerName || '–') + '</td>'
                + '<td class="num">' + fmt(row.physicalStock) + '</td><td class="num">' + fmt(row.v1AvailableQty) + '</td>'
                + '<td class="num">' + fmt(row.reservedQty) + '</td><td class="num">' + fmt(row.pickedQty) + '</td>'
                + '<td class="num">' + fmt(row.finishedQty) + '</td><td class="num">' + fmt(row.activeValue) + '</td></tr>';
        }).join('') : '<tr><td colspan="10" class="empty">Ingen reservationer matcher filteret.</td></tr>';
    }

    async function loadInventoryOverview(force) {
        byId('refreshReservationsBtn').disabled = true;
        byId('reservationStatus').textContent = 'Indlæser lager, NoPac og Rsv…';
        try {
            const currentPeriod = state.periods.get('current');
            if (!currentPeriod) throw new Error('Aktuel lagerliste mangler');
            if (force) currentPeriod.data = await fetchJson('/lagerliste/current?force=1');
            const current = await loadPeriod('current');
            const categories = Lagerliste2Engine.unwrapPayload(current).categories || {};
            const salesOrders = Array.from(new Set((categories.salgordreVia || []).concat(categories.finishedNotInvoiced || [])
                .map(row => Number(row.OrdNo || 0)).filter(Boolean)));
            const evidencePromise = salesOrders.length && current.generatedAt
                ? fetchJson('/lagerliste2/movement-evidence', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ from: current.generatedAt, to: current.generatedAt, salesOrders })
                })
                : Promise.resolve({ orderStates: [] });
            const [reservationResult, evidenceResult] = await Promise.allSettled([
                fetchJson('/lagerliste2/reservations/current' + (force ? '?force=1' : '')),
                evidencePromise
            ]);
            const reservationData = reservationResult.status === 'fulfilled' ? reservationResult.value : { rows: [], summary: {} };
            const evidence = evidenceResult.status === 'fulfilled' ? evidenceResult.value : { orderStates: [] };
            state.reservations = Array.isArray(reservationData.rows) ? reservationData.rows : [];
            state.reservationSummary = reservationData.summary || {};
            state.currentValuation = Lagerliste2Engine.canonicalValueSummary(current, { evidence });
            byId('valuationV1').textContent = fmt(state.currentValuation.rawV1Total);
            byId('valuationV2').textContent = fmt(state.currentValuation.total);
            byId('valuationDedup').textContent = fmt(state.currentValuation.duplicateReduction);
            byId('reservationActiveValue').textContent = fmt(state.reservationSummary.activeValue);
            renderReservations();
            const linked = Number(state.reservationSummary.linkedRowCount || 0);
            const total = Number(state.reservationSummary.rowCount || 0);
            const active = Number(state.reservationSummary.activeRowCount || 0);
            const finished = Number(state.reservationSummary.finishedRowCount || 0);
            const warnings = [];
            if (reservationResult.status === 'rejected') warnings.push('Rsv kunne ikke læses: ' + authMessage(reservationResult.reason));
            if (evidenceResult.status === 'rejected') warnings.push('NoPac kunne ikke læses: ' + authMessage(evidenceResult.reason));
            byId('reservationStatus').textContent = active + ' åbne Rsv-linjer. ' + linked + ' af ' + total
                + ' registreringer har sikker salgsordre; ' + finished + ' færdigmeldte Rsv-linjer vises kun som historik og tælles ikke igen.'
                + (warnings.length ? ' ' + warnings.join(' ') : '');
        } catch (error) {
            byId('reservationStatus').textContent = 'Lagerværdi/reservationer kunne ikke indlæses: ' + authMessage(error);
            byId('reservationRows').innerHTML = '<tr><td colspan="10" class="empty">Ingen data.</td></tr>';
        } finally {
            byId('refreshReservationsBtn').disabled = false;
        }
    }

    function currentRouteForMovement(row) {
        if (row.category !== 'Plader VIA') return null;
        return state.routes.find(route => Number(route.nestingOrdNo) === Number(row.orderNo) && String(route.route) === String(row.route));
    }

    function renderComparison(result) {
        state.comparison = result;
        byId('metricMatched').textContent = result.transfers.length;
        byId('metricResidual').textContent = result.unresolved.length;
        byId('matchedRows').innerHTML = result.transfers.length ? result.transfers.map(row => {
            const key = row.productKey || row.salesOrderNo || row.orderNo || '–';
            const confidence = row.confidence === 'high' ? 'Høj' : 'Mellem';
            const netAmount = Number(row.netAmount || 0);
            const net = netAmount > 0 ? ('+' + fmt(netAmount)) : fmt(netAmount);
            const netClass = netAmount > 0 ? 'positive' : (netAmount < 0 ? 'negative' : 'zero');
            return '<tr><td><strong>' + esc(row.flow) + '</strong></td><td>' + esc(key) + (row.route ? '<br><small>Route ' + esc(row.route) + '</small>' : '') + '</td>'
                + '<td>' + esc(row.sourceCategory + ' · ' + row.sourceKey) + '</td><td>' + esc(row.targetCategory + ' · ' + row.targetKey) + '</td>'
                + '<td><span class="pill ' + (row.confidence === 'high' ? 'done' : 'info') + '">' + confidence + '</span></td>'
                + '<td class="num">' + fmt(row.amount) + '</td><td class="num ' + netClass + '">' + net + '</td></tr>';
        }).join('') : '<tr><td colspan="7" class="empty">Ingen sikre automatiske overførsler fundet.</td></tr>';

        byId('residualRows').innerHTML = result.unresolved.length ? result.unresolved.map(row => {
            const openRoute = currentRouteForMovement(row);
            const expected = openRoute && openRoute.status !== 'completed';
            const possibleInternal = openRoute && Number(openRoute.unlinkedProductCount || 0) > 0 && !(openRoute.salesOrderNos || []).length;
            const status = expected
                ? '<span class="pill open">⏳ Åben route</span>'
                : (possibleInternal ? '<span class="pill info">🏭 Intern/lager?</span>' : '<span class="pill error">⚠ Kontroller</span>');
            const cls = row.remaining > 0 ? 'positive' : 'negative';
            return '<tr><td>' + status + '</td><td>' + esc(row.category) + '</td><td>' + esc(row.key) + '</td><td>' + esc(row.label || '–') + '</td>'
                + '<td class="num">' + fmt(row.valueA) + '</td><td class="num">' + fmt(row.valueB) + '</td><td class="num ' + cls + '">' + fmt(row.remaining) + '</td></tr>';
        }).join('') : '<tr><td colspan="7" class="empty">Alle bevægelser er afstemt.</td></tr>';

        byId('rawRows').innerHTML = result.movements.map(row => '<tr><td>' + esc(row.category) + '</td><td>' + esc(row.key) + '</td><td class="num">' + fmt(row.valueA) + '</td><td class="num">' + fmt(row.valueB) + '</td><td class="num ' + (row.diff > 0 ? 'positive' : 'negative') + '">' + fmt(row.diff) + '</td></tr>').join('');
        byId('matchedPanel').classList.toggle('hidden', !byId('showMatched').checked);
        byId('rawPanel').classList.toggle('hidden', !byId('showRaw').checked);
    }

    async function loadMovementEvidence(payloadA, payloadB, movements) {
        const from = Lagerliste2Engine.unwrapPayload(payloadA).generatedAt;
        const to = Lagerliste2Engine.unwrapPayload(payloadB).generatedAt;
        if (!from || !to) return { evidence: {}, warning: 'Snapshotdato mangler; transaktionsforklaringer kunne ikke hentes.' };
        const products = Array.from(new Set(movements
            .filter(row => ['Pladelager', 'Plader VIA', 'Opfølgningsvarer'].includes(row.category) && row.productKey)
            .map(row => row.productKey)));
        const salesOrders = Array.from(new Set(movements
            .filter(row => ['VIA Laser', 'VIA Stang', 'Indkøbt dele', 'VIA Tid', 'VIA ikke pakket', 'Færdige SO'].includes(row.category) && row.orderNo > 0)
            .map(row => row.orderNo)
            .concat(movements.filter(row => row.category === 'Plader VIA').flatMap(row => {
                const route = currentRouteForMovement(row);
                return route && Array.isArray(route.salesOrderNos) ? route.salesOrderNos : [];
            }))));
        const nestingOrders = Array.from(new Set(movements
            .filter(row => row.category === 'Plader VIA' && row.orderNo > 0)
            .map(row => row.orderNo)));
        const restCodes = Array.from(new Set(movements
            .filter(row => row.category === 'Rest plader' && row.remaining < 0 && row.restCode)
            .map(row => row.restCode)));
        if (!products.length && !salesOrders.length && !nestingOrders.length && !restCodes.length) return { evidence: {}, warning: '' };
        try {
            const evidence = await fetchJson('/lagerliste2/movement-evidence', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ from, to, products, salesOrders, nestingOrders, restCodes })
            });
            return { evidence, warning: '' };
        } catch (error) {
            return { evidence: {}, warning: 'Transaktionsbeviser kunne ikke hentes: ' + authMessage(error) };
        }
    }

    async function compare() {
        const from = byId('periodA').value;
        const to = byId('periodB').value;
        if (!from || !to || from === to) {
            byId('compareStatus').textContent = 'Vælg to forskellige perioder.';
            return;
        }
        byId('compareBtn').disabled = true;
        byId('compareStatus').textContent = 'Sammenligner perioder…';
        try {
            const [payloadA, payloadB] = await Promise.all([loadPeriod(from), loadPeriod(to)]);
            const movements = Lagerliste2Engine.buildMovements(payloadA, payloadB, { routes: state.routes });
            const evidenceResult = await loadMovementEvidence(payloadA, payloadB, movements);
            const result = Lagerliste2Engine.reconcileMovements(payloadA, payloadB, { routes: state.routes, evidence: evidenceResult.evidence });
            renderComparison(result);
            const allocationText = result.orderAllocations && result.orderAllocations.length
                ? ' ' + result.orderAllocations.length + ' ordre(r) er fordelt uden dobbelttælling efter NoPac.'
                : '';
            byId('compareStatus').textContent = result.transfers.length + ' bevægelser forklaret automatisk; ' + result.unresolved.length + ' restbevægelser beholdt til kontrol.' + allocationText + (evidenceResult.warning ? ' ' + evidenceResult.warning : '');
        } catch (error) {
            byId('compareStatus').textContent = 'Sammenligning fejlede: ' + authMessage(error);
        } finally {
            byId('compareBtn').disabled = false;
        }
    }

    byId('routeSearch').addEventListener('input', renderRoutes);
    byId('routeStatusFilter').addEventListener('change', renderRoutes);
    byId('refreshRoutesBtn').addEventListener('click', () => loadRoutes(true));
    byId('refreshReservationsBtn').addEventListener('click', () => loadInventoryOverview(true));
    byId('reservationFilter').addEventListener('change', renderReservations);
    byId('compareBtn').addEventListener('click', compare);
    byId('showMatched').addEventListener('change', () => byId('matchedPanel').classList.toggle('hidden', !byId('showMatched').checked));
    byId('showRaw').addEventListener('change', () => byId('rawPanel').classList.toggle('hidden', !byId('showRaw').checked));

    Promise.all([loadRoutes(false), loadPeriods()]).then(async () => {
        await loadInventoryOverview(false);
        if (state.periods.size > 1) await compare();
    }).catch(error => {
        byId('compareStatus').textContent = authMessage(error);
    });
})();
