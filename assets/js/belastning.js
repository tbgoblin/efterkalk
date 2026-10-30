// ── Belastning (capacity/load board) · client ────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: canAccessModule, escapeHtmlFE, sanitizeDisplayName,
// getDashboardSourceCache, requestBelastningSummary, scheduleDashboardWidgets
// (dichiarati nell'inline script).

function setBelastningStatus(message) {
    const el = document.getElementById('belastningStatus');
    if (el) el.textContent = String(message || '');
}

function scrollBelastningDetailIntoView() {
    const detailEl = document.getElementById('belastningDetailSvg') || document.getElementById('belastningDetailWrap');
    if (!detailEl) return;
    const rect = detailEl.getBoundingClientRect();
    const headerOffset = 86;
    const top = Math.max(0, (window.scrollY || 0) + rect.top - headerOffset);
    window.scrollTo({ top, behavior: 'smooth' });
}

function getBelastningFilters() {
    const todayInput = document.getElementById('belastningToDay');
    const daysInput = document.getElementById('belastningDage');
    const resGrInput = document.getElementById('belastningResGr');
    const orderInput = document.getElementById('belastningOrdre');
    const customerInput = document.getElementById('belastningKunde');
    const today = String(todayInput && todayInput.value || '').trim() || new Date().toISOString().slice(0, 10);
    const daysRaw = Number(daysInput && daysInput.value || 30);
    const dage = Number.isFinite(daysRaw) ? Math.max(1, Math.min(180, Math.round(daysRaw))) : 30;
    const resGr = String(resGrInput && resGrInput.value || '').trim();
    const ord = String(orderInput && orderInput.value || '').replace(/\D+/g, '').slice(0, 12);
    const kunde = String(customerInput && customerInput.value || '').trim().replace(/\s+/g, ' ').slice(0, 80);
    if (orderInput && orderInput.value !== ord) {
        orderInput.value = ord;
    }
    if (customerInput && customerInput.value !== kunde) {
        customerInput.value = kunde;
    }
    return { today, dage, resGr, ord, kunde };
}

function escapeJsSingle(value) {
    return String(value || '').split('\\').join('\\\\').split("'").join("\\'");
}

function formatBelastningMinutes(value) {
    const num = Number(value || 0);
    if (!Number.isFinite(num)) return '0';
    return new Intl.NumberFormat('da-DK', { minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(Math.round(num));
}

function getBelastningLayoutStorageKey() {
    return 'afterkalk_belastning_layout_v1_' + sanitizeDisplayName(loggedUserDisplayName).toLowerCase();
}

function getBelastningCardKey(item) {
    const rg = String(item && item.resGr || '').trim();
    const parity = item && Number(item.parity) === 0 ? 0 : 1;
    return rg ? (String(parity) + ':' + rg) : '';
}

function readBelastningLayoutOrder() {
    try {
        const raw = localStorage.getItem(getBelastningLayoutStorageKey());
        const parsed = JSON.parse(String(raw || '[]'));
        return Array.isArray(parsed) ? parsed.map(v => String(v || '').trim()).filter(Boolean) : [];
    } catch {
        return [];
    }
}

function writeBelastningLayoutOrder(keys) {
    try {
        localStorage.setItem(getBelastningLayoutStorageKey(), JSON.stringify(Array.from(new Set((Array.isArray(keys) ? keys : []).map(v => String(v || '').trim()).filter(Boolean)))));
    } catch {}
}

function sortBelastningItemsBySavedLayout(items) {
    const safeItems = Array.isArray(items) ? items.slice() : [];
    const order = readBelastningLayoutOrder();
    const orderIndex = new Map(order.map((key, index) => [key, index]));
    return safeItems.sort((a, b) => {
        const keyA = getBelastningCardKey(a);
        const keyB = getBelastningCardKey(b);
        const idxA = orderIndex.has(keyA) ? orderIndex.get(keyA) : Number.MAX_SAFE_INTEGER;
        const idxB = orderIndex.has(keyB) ? orderIndex.get(keyB) : Number.MAX_SAFE_INTEGER;
        if (idxA !== idxB) return idxA - idxB;
        return String(a && a.resGr || '').localeCompare(String(b && b.resGr || ''), 'da');
    });
}

function clearBelastningDragMarkers(root) {
    const host = root || document;
    host.querySelectorAll('.belastning-resource-chart.drag-target-before, .belastning-resource-chart.drag-target-after').forEach(el => {
        el.classList.remove('drag-target-before', 'drag-target-after');
    });
}

function persistBelastningLayoutFromDom(targetId) {
    const wrap = document.getElementById(targetId);
    if (!wrap) return;
    const visibleKeys = Array.from(wrap.querySelectorAll('.belastning-resource-chart[data-belastning-key]'))
        .map(el => String(el.getAttribute('data-belastning-key') || '').trim())
        .filter(Boolean);
    if (visibleKeys.length === 0) return;
    const previousKeys = readBelastningLayoutOrder().filter(key => !visibleKeys.includes(key));
    writeBelastningLayoutOrder([...visibleKeys, ...previousKeys]);
}

function attachBelastningDragAndDrop(targetId) {
    const wrap = document.getElementById(targetId);
    if (!wrap || wrap.dataset.dragReady === '1') return;
    wrap.dataset.dragReady = '1';

    wrap.addEventListener('dragstart', event => {
        const card = event.target && event.target.closest ? event.target.closest('.belastning-resource-chart[data-belastning-key]') : null;
        if (!card) return;
        belastningDraggedCardKey = String(card.getAttribute('data-belastning-key') || '').trim();
        if (!belastningDraggedCardKey) return;
        card.classList.add('is-dragging');
        if (event.dataTransfer) {
            event.dataTransfer.effectAllowed = 'move';
            event.dataTransfer.setData('text/plain', belastningDraggedCardKey);
        }
    });

    wrap.addEventListener('dragover', event => {
        if (!belastningDraggedCardKey) return;
        event.preventDefault();
        const target = event.target && event.target.closest ? event.target.closest('.belastning-resource-chart[data-belastning-key]') : null;
        clearBelastningDragMarkers(wrap);
        if (!target || String(target.getAttribute('data-belastning-key') || '').trim() === belastningDraggedCardKey) return;
        const rect = target.getBoundingClientRect();
        const insertBefore = event.clientY < rect.top + (rect.height / 2);
        target.classList.add(insertBefore ? 'drag-target-before' : 'drag-target-after');
    });

    wrap.addEventListener('drop', event => {
        if (!belastningDraggedCardKey) return;
        event.preventDefault();
        const dragged = wrap.querySelector('.belastning-resource-chart[data-belastning-key="' + cssEscape(belastningDraggedCardKey) + '"]');
        const target = event.target && event.target.closest ? event.target.closest('.belastning-resource-chart[data-belastning-key]') : null;
        clearBelastningDragMarkers(wrap);
        if (!dragged || !target || dragged === target) return;
        const rect = target.getBoundingClientRect();
        const insertBefore = event.clientY < rect.top + (rect.height / 2);
        if (insertBefore) {
            wrap.insertBefore(dragged, target);
        } else {
            wrap.insertBefore(dragged, target.nextSibling);
        }
        persistBelastningLayoutFromDom(targetId);
    });

    wrap.addEventListener('dragend', () => {
        clearBelastningDragMarkers(wrap);
        wrap.querySelectorAll('.belastning-resource-chart.is-dragging').forEach(el => el.classList.remove('is-dragging'));
        belastningDraggedCardKey = '';
    });
}

function cssEscape(value) {
    const bs = String.fromCharCode(92);
    return String(value || '')
        .split(bs).join(bs + bs)
        .split('"').join(bs + '"');
}

function normalizeBelastningDateKey(rawDate, dateLabel) {
    const label = String(dateLabel || '').trim();
    const fullDa = label.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (fullDa) {
        const dd = fullDa[1].padStart(2, '0');
        const mm = fullDa[2].padStart(2, '0');
        const yy = fullDa[3];
        return yy + '-' + mm + '-' + dd;
    }
    const parsed = rawDate ? new Date(rawDate) : null;
    if (parsed && !Number.isNaN(parsed.getTime())) {
        const y = parsed.getFullYear();
        const m = String(parsed.getMonth() + 1).padStart(2, '0');
        const d = String(parsed.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
    }
    return '';
}

function normalizeBelastningDisplayDate(rawDate, dateLabel) {
    const label = String(dateLabel || '').trim();
    if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(label)) return label;
    const parsed = rawDate ? new Date(rawDate) : null;
    if (parsed && !Number.isNaN(parsed.getTime())) {
        return parsed.toLocaleDateString('da-DK');
    }
    return label || '-';
}

function getBelastningDateSortValue(dateKey) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateKey || ''))) return Number.MAX_SAFE_INTEGER;
    const ts = new Date(dateKey + 'T00:00:00').getTime();
    return Number.isNaN(ts) ? Number.MAX_SAFE_INTEGER : ts;
}

async function onBelastningDayColumnClick(resGr, parity, dayKey, event) {
    if (event && typeof event.stopPropagation === 'function') {
        event.stopPropagation();
    }
    const safeDayKey = String(dayKey || '').trim();
    if (!safeDayKey) return;
    await loadBelastningDetail(resGr, parity, { focusDayKey: safeDayKey });
}

function scheduleBelastningAutoReload() {
    if (belastningAutoReloadTimer) {
        clearTimeout(belastningAutoReloadTimer);
    }
    belastningAutoReloadTimer = setTimeout(() => {
        belastningAutoReloadTimer = null;
        loadBelastningGrafisk({ forceRefresh: true, silentValidation: true });
    }, BELASTNING_FILTER_DEBOUNCE_MS);
}

function startBelastningPeriodicRefresh() {
    if (belastningPeriodicTimer) return;
    belastningPeriodicTimer = setInterval(() => {
        const root = document.getElementById('mainBelastning');
        const isVisible = root && root.style.display !== 'none';
        if (isVisible) {
            loadBelastningGrafisk({ forceRefresh: true, silentValidation: true });
        }
    }, BELASTNING_PERIODIC_REFRESH_MS);
}

function buildBelastningClusterSvg(dayRows, opts) {
    const activeDateInput = document.getElementById('belastningToDay');
    return window.GohReportCharts.belastningCluster(Array.isArray(dayRows) ? dayRows : [], {
        ...(opts || {}),
        today: activeDateInput && activeDateInput.value || new Date().toISOString().slice(0, 10),
        viewportWidth: window.innerWidth || 1200
    });
}

function renderBelastningDetailSvg(rows, opts) {
    const wrap = document.getElementById('belastningDetailSvg');
    if (!wrap) return;
    const safeRows = Array.isArray(rows) ? rows : [];
    if (safeRows.length === 0) {
        wrap.innerHTML = '';
        return;
    }
    wrap.innerHTML = buildBelastningClusterSvg(safeRows, opts);
}

function renderBelastningBars(targetId, items, allRows) {
    const wrap = document.getElementById(targetId);
    if (!wrap) return;
    const rows = sortBelastningItemsBySavedLayout(items);
    const fullRows = Array.isArray(allRows) ? allRows : [];
    if (rows.length === 0) {
        wrap.innerHTML = '<div class="qms-empty">Ingen data i valgt periode.</div>';
        const svgTarget = targetId === 'belastningBarsCombined' ? 'belastningSvgCombined' : null;
        const svgWrap = document.getElementById(svgTarget);
        if (svgWrap) svgWrap.innerHTML = '';
        return;
    }

    const grouped = new Map();
    for (const row of fullRows) {
        const key = String(row && row.ResGr || '').trim();
        if (!key) continue;
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key).push(row);
    }

    wrap.innerHTML = rows.map(item => {
        const kap = Number(item.totalKap || 0);
        const resv = Number(item.totalResv || 0);
        const aften = Number(item.totalAften || 0);
        const loadPct = kap > 0 ? Math.min(160, (resv / kap) * 100) : 0;
        const rg = String(item.resGr || '').trim();
        const chartRows = grouped.get(rg) || [];
        const itemParity = item && item.parity === 0 ? 0 : 1;
        const isActive = belastningDetailContext
            && String(belastningDetailContext.resGr || '') === rg
            && Number(belastningDetailContext.parity) === itemParity;
        const cardKey = getBelastningCardKey(item);
        return ''
            + '<div class="belastning-resource-chart' + (isActive ? ' is-active' : '') + '" draggable="true" data-belastning-key="' + escapeHtmlFE(cardKey) + '" onclick="loadBelastningDetail(\'' + escapeHtmlFE(rg) + '\',' + itemParity + ')">'
            + '<div class="belastning-card-top">'
            + '<h5>Kapacitetsbelastning: ' + escapeHtmlFE(rg + ' ' + String(item.nm || '')) + '</h5>'
            + '<span class="belastning-drag-chip" title="Træk kortet for at gemme din egen rækkefølge">Flyt</span>'
            + '</div>'
            + buildBelastningClusterSvg(chartRows, {
                clickable: true,
                resGr: rg,
                parity: itemParity,
                activeDayKey: isActive ? belastningSelectedDayKey : ''
            })
            + '<div class="belastning-mini-meta">'
            + '<span>Belastning: ' + escapeHtmlFE(formatPctDa(loadPct)) + '</span> · '
            + '<span>Resv: ' + escapeHtmlFE(formatBelastningMinutes(resv)) + '</span>'
            + '<span>Kap: ' + escapeHtmlFE(formatBelastningMinutes(kap)) + '</span>'
            + '<span>Aften: ' + escapeHtmlFE(formatBelastningMinutes(aften)) + '</span>'
            + '</div></div>';
    }).join('');

        attachBelastningDragAndDrop(targetId);

    const svgTarget = targetId === 'belastningBarsCombined' ? 'belastningSvgCombined' : null;
    const svgWrap = document.getElementById(svgTarget);
    if (svgWrap) svgWrap.innerHTML = '';
}

function renderBelastningDetailTable(detailData, resGr, parity) {
    const wrap = document.getElementById('belastningDetailTable');
    const title = document.getElementById('belastningDetailTitle');
    const card = document.getElementById('belastningDetailWrap');
    if (!wrap || !title || !card) return;
    const safeRows = Array.isArray(detailData)
        ? detailData
        : (detailData && Array.isArray(detailData.rows) ? detailData.rows : []);
    const orderRows = detailData && Array.isArray(detailData.orderRows) ? detailData.orderRows : [];
    const subOrderRows = detailData && Array.isArray(detailData.subOrderRows) ? detailData.subOrderRows : [];
    const orderLineRows = detailData && Array.isArray(detailData.orderLineRows) ? detailData.orderLineRows : [];
    const resourceLookup = belastningLastPayload
        ? [
            ...(belastningLastPayload.odd && Array.isArray(belastningLastPayload.odd.resources) ? belastningLastPayload.odd.resources : []),
            ...(belastningLastPayload.even && Array.isArray(belastningLastPayload.even.resources) ? belastningLastPayload.even.resources : [])
        ]
        : [];
    const resourceName = (() => {
        const match = resourceLookup.find(item => String(item && item.resGr || '').trim() === String(resGr || '').trim());
        return match ? String(match.nm || '').trim() : '';
    })();

    if (safeRows.length === 0 && orderRows.length === 0) {
        card.style.display = 'none';
        wrap.innerHTML = '';
        return;
    }
    const groupLabel = String(resGr || '').trim() === '51'
        ? 'Robotsvejs'
        : '';
    const titleParts = ['Ordreoverblik: ' + String(resGr || '') + (resourceName ? (' ' + resourceName) : '')];
    if (groupLabel) titleParts.push(groupLabel);
    if (detailData && detailData.ord) titleParts.push('Ordre ' + String(detailData.ord));
    if (detailData && detailData.kunde) titleParts.push('Kunde ' + String(detailData.kunde));
    title.textContent = titleParts.join(' · ');

    const subOrderMap = new Map();
    for (const subRow of subOrderRows) {
        const key = String(subRow && subRow.SubOrdNo || '').trim();
        if (!key) continue;
        if (!subOrderMap.has(key)) subOrderMap.set(key, []);
        subOrderMap.get(key).push(subRow);
    }

    const orderLineMap = new Map();
    for (const line of orderLineRows) {
        const key = String(line && line.OrdNo || '').trim();
        if (!key) continue;
        if (!orderLineMap.has(key)) orderLineMap.set(key, []);
        orderLineMap.get(key).push(line);
    }

    const groupedByDate = new Map();
    const orderRowsWork = orderRows.map(row => ({
        ...row,
        SOrdre: String((row && row.SOrdre) || '').trim(),
        POrdre: String((row && row.POrdre) || '').trim(),
        OrdNo: String((row && row.OrdNo) || '').trim(),
        Kunde: String((row && row.Kunde) || '').trim(),
        Opr: String((row && row.Opr) || '').trim()
    }));
    const activeDateInput = document.getElementById('belastningToDay');
    const todayCut = (() => {
        const raw = activeDateInput && activeDateInput.value
            ? activeDateInput.value
            : new Date().toISOString().slice(0, 10);
        const t = new Date(raw + 'T00:00:00').getTime();
        return Number.isNaN(t) ? Date.now() : t;
    })();
    const readResv = row => Number((row && (row.Resv !== undefined ? row.Resv : row.ResvRaw)) || 0);
    const readRest = row => Number((row && (row.RestResv !== undefined ? row.RestResv : row.ResvNet)) || 0);
    const readAften = row => Number((row && (row.RestAften !== undefined ? row.RestAften : (row.Aften !== undefined ? row.Aften : row.AftenRaw))) || 0);

    for (const row of orderRowsWork) {
        const dateKey = normalizeBelastningDateKey(row && row.Dato, row && row.DatoX);
        const dateSort = getBelastningDateSortValue(dateKey);
        const dateTxt = normalizeBelastningDisplayDate(row && row.Dato, row && row.DatoX);
        const mapKey = dateKey || dateTxt;
        if (!groupedByDate.has(mapKey)) {
            groupedByDate.set(mapKey, {
                dateKey,
                dateTxt,
                dateSort,
                rows: [],
                totalResv: 0,
                totalRest: 0,
                totalAften: 0
            });
        }
        const bucket = groupedByDate.get(mapKey);
        bucket.rows.push(row);
        bucket.totalResv += readResv(row);
        bucket.totalRest += readRest(row);
        bucket.totalAften += readAften(row);
    }

    const sortedDateGroups = Array.from(groupedByDate.values()).sort((a, b) => a.dateSort - b.dateSort);
    const fmtDate = value => {
        const parsed = value ? new Date(value) : null;
        return parsed && !Number.isNaN(parsed.getTime())
            ? parsed.toLocaleDateString('da-DK')
            : '-';
    };

    const tableRows = sortedDateGroups.map((dateGroup, groupIndex) => {
        const dateKey = 'beldate_' + groupIndex;
        const groupsBySOrdre = new Map();
        for (const row of dateGroup.rows) {
            const sOrdre = String((row && (row.SOrdre || row.OrdNo)) || '-').trim() || '-';
            if (!groupsBySOrdre.has(sOrdre)) {
                groupsBySOrdre.set(sOrdre, {
                    sOrdre,
                    rows: [],
                    totalResv: 0,
                    totalRest: 0,
                    totalAften: 0
                });
            }
            const bucket = groupsBySOrdre.get(sOrdre);
            bucket.rows.push(row);
            bucket.totalResv += readResv(row);
            bucket.totalRest += readRest(row);
            bucket.totalAften += readAften(row);
        }

        const orderRowsHtml = Array.from(groupsBySOrdre.values()).map((orderGroup, orderIndex) => {
            const firstRow = orderGroup.rows[0] || {};
            const detailKey = dateKey + '_ord_' + orderIndex;
            const beforeClass = dateGroup.dateSort < todayCut ? ' belastning-before-day' : '';

            const groupsByPOrdre = new Map();
            for (const row of orderGroup.rows) {
                const pOrdre = String((row && (row.POrdre || row.PurcNo || row.OrdNo)) || '-').trim() || '-';
                if (!groupsByPOrdre.has(pOrdre)) {
                    groupsByPOrdre.set(pOrdre, {
                        pOrdre,
                        rows: [],
                        totalResv: 0,
                        totalRest: 0,
                        totalAften: 0
                    });
                }
                const pBucket = groupsByPOrdre.get(pOrdre);
                pBucket.rows.push(row);
                pBucket.totalResv += readResv(row);
                pBucket.totalRest += readRest(row);
                pBucket.totalAften += readAften(row);
            }

            const pOrderGroups = Array.from(groupsByPOrdre.values());
            const hasChildren = pOrderGroups.length > 0;
            const routeList = Array.from(new Set(orderGroup.rows.map(x => String(x && x.Opr || '').trim()).filter(Boolean))).join(' ');
            const parentRow = '<tr class="belastning-order-row' + beforeClass + '" data-parent-date="' + dateKey + '" data-parent-order="' + detailKey + '">'
                + '<td style="text-align:center;">'
                + (hasChildren
                    ? ('<button type="button" class="belastning-order-toggle" data-order-key="' + detailKey + '" data-collapsed="1" onclick="toggleBelastningOrderNode(\'' + detailKey + '\', this)">+</button>')
                    : '<span class="belastning-order-sub">-</span>')
                + '</td>'
                + '<td>' + escapeHtmlFE(dateGroup.dateTxt) + '</td>'
                + '<td style="text-align:right;"><span class="belastning-order-id">' + escapeHtmlFE(String(firstRow.SOrdre || orderGroup.sOrdre || '-')) + '</span></td>'
                + '<td style="text-align:right;"><span class="belastning-order-sub">intern</span></td>'
                + '<td>' + escapeHtmlFE(String(firstRow.Kunde || '-')) + '</td>'
                + '<td>' + escapeHtmlFE(routeList || '-') + '</td>'
                + '<td>' + escapeHtmlFE(String(firstRow.LevMode || '-')) + '</td>'
                + '<td>' + escapeHtmlFE(fmtDate(firstRow.LevDato)) + '</td>'
                + '<td>' + escapeHtmlFE(fmtDate(firstRow.ULDato)) + '</td>'
                + '<td style="text-align:right;">' + escapeHtmlFE(formatBelastningMinutes(orderGroup.totalResv)) + '</td>'
                + '<td style="text-align:right;">' + escapeHtmlFE(formatBelastningMinutes(orderGroup.totalRest)) + '</td>'
                + '<td style="text-align:right;">' + escapeHtmlFE(formatBelastningMinutes(orderGroup.totalAften)) + '</td>'
                + '</tr>';

            const childRowsHtml = pOrderGroups.map((pOrderGroup, childIndex) => {
                const childKey = detailKey + '_child_' + childIndex;
                const childRow = pOrderGroup.rows[0] || {};
                const routeText = String(childRow.Opr || '').trim() || '-';
                return '<tr class="belastning-order-detail-row' + beforeClass + '" data-parent-date="' + dateKey + '" data-parent-order="' + detailKey + '" data-child-key="' + childKey + '" style="display:none;">'
                    + '<td></td>'
                    + '<td>' + escapeHtmlFE(dateGroup.dateTxt) + '</td>'
                    + '<td style="text-align:right;">' + escapeHtmlFE(String(childRow.SOrdre || orderGroup.sOrdre || '-')) + '</td>'
                    + '<td style="text-align:right;">' + escapeHtmlFE(String(pOrderGroup.pOrdre || '-')) + '</td>'
                    + '<td>' + escapeHtmlFE(String(childRow.Kunde || '-')) + '</td>'
                    + '<td>' + escapeHtmlFE(routeText) + '</td>'
                    + '<td>-</td>'
                    + '<td>-</td>'
                    + '<td>' + escapeHtmlFE(fmtDate(childRow.ULDato)) + '</td>'
                    + '<td style="text-align:right;">' + escapeHtmlFE(formatBelastningMinutes(pOrderGroup.totalResv)) + '</td>'
                    + '<td style="text-align:right;">-</td>'
                    + '<td style="text-align:right;">-</td>'
                    + '</tr>';
            }).join('');

            return parentRow + childRowsHtml;
        }).join('');

        const groupHeaderClass = dateGroup.dateSort < todayCut ? 'belastning-date-row belastning-before-day' : 'belastning-date-row';
        const groupHeader = '<tr class="' + groupHeaderClass + '" data-day-key="' + escapeHtmlFE(String(dateGroup.dateKey || '')) + '">'
            + '<td colspan="12">'
            + '<div class="belastning-date-header">'
            + '<button type="button" class="belastning-date-toggle" data-date-key="' + dateKey + '" data-collapsed="0" onclick="toggleBelastningDateGroup(\'' + dateKey + '\', this)">-</button>'
            + '<span class="belastning-date-badge">Dato: ' + escapeHtmlFE(dateGroup.dateTxt) + '</span>'
            + '<span class="belastning-date-meta">'
            + '<span class="belastning-date-meta-item"><span class="belastning-date-meta-lbl">S-Ordre:</span><span class="belastning-date-meta-val">' + escapeHtmlFE(String(groupsBySOrdre.size)) + '</span></span>'
            + '<span class="belastning-date-meta-item"><span class="belastning-date-meta-lbl">Minutter:</span><span class="belastning-date-meta-val">' + escapeHtmlFE(formatBelastningMinutes(dateGroup.totalResv)) + '</span></span>'
            + '<span class="belastning-date-meta-item"><span class="belastning-date-meta-lbl">Rest:</span><span class="belastning-date-meta-val">' + escapeHtmlFE(formatBelastningMinutes(dateGroup.totalRest)) + '</span></span>'
            + '<span class="belastning-date-meta-item"><span class="belastning-date-meta-lbl">Aften:</span><span class="belastning-date-meta-val">' + escapeHtmlFE(formatBelastningMinutes(dateGroup.totalAften)) + '</span></span>'
            + '</span>'
            + '</div></td></tr>';

        return groupHeader + orderRowsHtml;
    }).join('');

    const orderTable = sortedDateGroups.length > 0
        ? ('<div class="belastning-section-title">Ordreoverblik</div>'
            + '<div class="belastning-order-shell">'
            + '<table class="belastning-order-table">'
            + '<thead><tr>'
            + '<th></th><th>Dato</th><th>S-Ordre</th><th>P-Ordre</th><th>Kunde</th><th>Rute</th><th>Lev.måde</th><th>Lev.dato</th><th>U-dato</th><th>Minutter</th><th>Rest</th><th>Aften</th>'
            + '</tr></thead>'
            + '<tbody>' + tableRows + '</tbody></table>'
            + '</div>')
        : '<div class="qms-empty">Ingen ordrelinjer for valgt ressourcegruppe/periode.</div>';

    renderBelastningDetailSvg(safeRows, {
        clickable: true,
        resGr: String(resGr || '').trim(),
        parity: parity === 0 ? 0 : 1,
        activeDayKey: belastningSelectedDayKey
    });
    wrap.innerHTML = orderTable;
    card.style.display = 'block';
    focusBelastningDayInTable(belastningSelectedDayKey);
}

function focusBelastningDayInTable(dayKey) {
    const key = String(dayKey || '').trim();
    const isBeforeFocus = key === 'before';
    const headers = Array.from(document.querySelectorAll('tr.belastning-date-row'));
    headers.forEach(row => row.classList.remove('belastning-day-focus'));
    if (!key) return;

    if (isBeforeFocus) {
        headers.filter(row => row.classList.contains('belastning-before-day')).forEach(row => row.classList.add('belastning-day-focus'));
    } else {
        const targetHeader = document.querySelector('tr.belastning-date-row[data-day-key="' + key + '"]');
        if (!targetHeader) return;
        targetHeader.classList.add('belastning-day-focus');
    }

    const toggles = Array.from(document.querySelectorAll('.belastning-date-toggle'));
    toggles.forEach(btn => {
        const internalKey = String(btn.getAttribute('data-date-key') || '');
        const groupRows = Array.from(document.querySelectorAll('tr[data-parent-date="' + internalKey + '"]'));
        const headerRow = btn.closest('tr.belastning-date-row');
        const isTarget = headerRow && (isBeforeFocus
            ? headerRow.classList.contains('belastning-before-day')
            : headerRow.getAttribute('data-day-key') === key);
        if (isTarget) {
            btn.setAttribute('data-collapsed', '0');
            btn.textContent = '-';
            groupRows.forEach(row => {
                if (!row.classList.contains('belastning-order-detail-row')) {
                    row.style.display = 'table-row';
                    return;
                }
                const orderKey = row.getAttribute('data-parent-order') || '';
                const orderBtn = document.querySelector('.belastning-order-toggle[data-order-key="' + orderKey + '"]');
                const orderCollapsed = !orderBtn || orderBtn.getAttribute('data-collapsed') === '1';
                row.style.display = orderCollapsed ? 'none' : 'table-row';
            });
        } else {
            btn.setAttribute('data-collapsed', '1');
            btn.textContent = '+';
            groupRows.forEach(row => {
                row.style.display = 'none';
            });
        }
    });
}

function toggleBelastningDateGroup(dateKey, buttonEl) {
    const collapsed = buttonEl && buttonEl.getAttribute('data-collapsed') === '1';
    const nextCollapsed = !collapsed;
    if (buttonEl) {
        buttonEl.setAttribute('data-collapsed', nextCollapsed ? '1' : '0');
        buttonEl.textContent = nextCollapsed ? '+' : '-';
    }

    const rows = document.querySelectorAll('tr[data-parent-date="' + dateKey + '"]');
    rows.forEach(row => {
        if (nextCollapsed) {
            row.style.display = 'none';
            return;
        }
        if (row.classList.contains('belastning-order-detail-row')) {
            const orderKey = row.getAttribute('data-parent-order') || '';
            const orderBtn = document.querySelector('.belastning-order-toggle[data-order-key="' + orderKey + '"]');
            const orderCollapsed = !orderBtn || orderBtn.getAttribute('data-collapsed') === '1';
            row.style.display = orderCollapsed ? 'none' : 'table-row';
            return;
        }
        row.style.display = 'table-row';
    });
}

function toggleBelastningOrderNode(orderKey, buttonEl) {
    const detailRows = Array.from(document.querySelectorAll('tr.belastning-order-detail-row[data-parent-order="' + orderKey + '"]'));
    if (detailRows.length === 0) return;

    const collapsed = buttonEl && buttonEl.getAttribute('data-collapsed') === '1';
    const nextCollapsed = !collapsed;
    if (buttonEl) {
        buttonEl.setAttribute('data-collapsed', nextCollapsed ? '1' : '0');
        buttonEl.textContent = nextCollapsed ? '+' : '-';
    }

    const dateKey = detailRows[0].getAttribute('data-parent-date') || '';
    const dateBtn = document.querySelector('.belastning-date-toggle[data-date-key="' + dateKey + '"]');
    const dateCollapsed = dateBtn && dateBtn.getAttribute('data-collapsed') === '1';
    detailRows.forEach(row => {
        row.style.display = (nextCollapsed || dateCollapsed) ? 'none' : 'table-row';
    });
}

async function loadBelastningDetail(resGr, parity, options) {
    try {
        const safeOptions = options && typeof options === 'object' ? options : {};
        const filters = getBelastningFilters();
        const query = new URLSearchParams({
            toDay: filters.today,
            dage: String(filters.dage),
            resGr: String(resGr || ''),
            parity: String(parity === 0 ? 0 : 1),
            ord: filters.ord,
            kunde: filters.kunde
        });
        belastningDetailContext = { resGr: String(resGr || '').trim(), parity: parity === 0 ? 0 : 1 };
        belastningSelectedDayKey = String(safeOptions.focusDayKey || '').trim();
        const response = await fetch('/belastning/detail?' + query.toString());
        if (!response.ok) throw new Error('HTTP ' + response.status);
        const payload = await response.json();
        if (!payload.ok) throw new Error(payload.error || 'Belastning detail fejl');
        renderBelastningDetailTable(payload, resGr, parity === 0 ? 0 : 1);
        if (belastningLastPayload) {
            const oddItems = belastningLastPayload.odd && Array.isArray(belastningLastPayload.odd.resources)
                ? belastningLastPayload.odd.resources.map(item => ({ ...item, parity: 1 }))
                : [];
            const evenItems = belastningLastPayload.even && Array.isArray(belastningLastPayload.even.resources)
                ? belastningLastPayload.even.resources.map(item => ({ ...item, parity: 0 }))
                : [];
            const oddRows = belastningLastPayload.odd && Array.isArray(belastningLastPayload.odd.rows) ? belastningLastPayload.odd.rows : [];
            const evenRows = belastningLastPayload.even && Array.isArray(belastningLastPayload.even.rows) ? belastningLastPayload.even.rows : [];
            renderBelastningBars('belastningBarsCombined', [...oddItems, ...evenItems], [...oddRows, ...evenRows]);
        }
        scrollBelastningDetailIntoView();
        setTimeout(scrollBelastningDetailIntoView, 160);
    } catch (err) {
        setBelastningStatus('Fejl ved detailhentning: ' + (err && err.message ? err.message : 'ukendt'));
    }
}

async function initializeBelastningIfNeeded() {
    if (belastningInitialized) return;
    belastningInitialized = true;
    const todayInput = document.getElementById('belastningToDay');
    if (todayInput && !todayInput.value) {
        todayInput.value = new Date().toISOString().slice(0, 10);
    }
    startBelastningPeriodicRefresh();
    setBelastningStatus('Henter data...');
    await loadBelastningGrafisk();
}

async function loadBelastningGrafisk(options) {
    if (!canAccessModule('belastning')) return;
    const safeOptions = options && typeof options === 'object' ? options : {};
    const emptyEl = document.getElementById('belastningEmpty');
    const loadBtn = document.getElementById('belastningLoadBtn');
    const graphWrap = document.getElementById('belastningGrafiskWrap');
    const detailWrap = document.getElementById('belastningDetailWrap');
    if (loadBtn) loadBtn.disabled = true;
    if (detailWrap) detailWrap.style.display = 'none';
    if (emptyEl) {
        emptyEl.style.display = 'block';
        emptyEl.textContent = 'Henter belastning...';
    }

    try {
        const filters = getBelastningFilters();
        const payload = await getDashboardSourceCache().load(belastningSummaryKey(filters), () => requestBelastningSummary(filters), safeOptions.forceRefresh === true);
        belastningLastPayload = payload;
        scheduleDashboardWidgets();
        const oddItems = payload.odd && Array.isArray(payload.odd.resources) ? payload.odd.resources.map(item => ({ ...item, parity: 1 })) : [];
        const evenItems = payload.even && Array.isArray(payload.even.resources) ? payload.even.resources.map(item => ({ ...item, parity: 0 })) : [];
        const oddRows = payload.odd && Array.isArray(payload.odd.rows) ? payload.odd.rows : [];
        const evenRows = payload.even && Array.isArray(payload.even.rows) ? payload.even.rows : [];
        const combinedItems = [...oddItems, ...evenItems];
        const combinedRows = [...oddRows, ...evenRows];
        renderBelastningBars('belastningBarsCombined', combinedItems, combinedRows);
        if (graphWrap) graphWrap.style.display = combinedItems.length ? 'grid' : 'none';
        if (emptyEl) {
            if (combinedItems.length === 0) {
                emptyEl.style.display = 'block';
                emptyEl.textContent = 'Ingen belastningsdata i valgt periode.';
            } else {
                emptyEl.style.display = 'none';
            }
        }
        const orderStatus = filters.ord ? (' · Ordre-filter: ' + filters.ord) : '';
        const customerStatus = filters.kunde ? (' · Kunde-filter: ' + filters.kunde) : '';
        setBelastningStatus('Periode: ' + filters.today + ' + ' + filters.dage + ' dage · Ressourcer: ' + combinedItems.length + orderStatus + customerStatus + ' · Auto-opdater: 15 min');
    } catch (err) {
        setBelastningStatus('Fejl ved hentning af belastning.');
        if (emptyEl) {
            emptyEl.style.display = 'block';
            emptyEl.textContent = 'Fejl: ' + (err && err.message ? err.message : 'ukendt fejl');
        }
        if (graphWrap) graphWrap.style.display = 'none';
    } finally {
        if (loadBtn) loadBtn.disabled = false;
    }
}
