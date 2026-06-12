
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

            function isLaserLProdNo(prodNo) {
                return String(prodNo || '').trim().toUpperCase().endsWith('L');
            }

            function isInvoiceTrackedProdNo(prodNo) {
                return String(prodNo || '').trim().toUpperCase().startsWith('U');
            }

            function shouldFilterChildSummary(prodTp4, prodNo, purcNo) {
                const displayKey = getDisplayProdTp4Key(prodTp4, prodNo, purcNo);
                return displayKey === '6' || displayKey === '9' || isInvoiceTrackedProdNo(prodNo);
            }

            function isProductionSummaryExcludedLine(line) {
                if (!line) return false;
                const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                return Number(line.LnNo || 0) === 1 || key === '0' || key === '3' || key === '5';
            }

            function getDisplayProdTp4Key(prodTp4, prodNo, purcNo) {
                const rawKey = (prodTp4 === null || prodTp4 === undefined) ? 'NA' : String(prodTp4);
                if (rawKey === '3') return '1';
                if (rawKey === '2' && !isLaserLProdNo(prodNo)) {
                    return Number(purcNo || 0) > 0 ? '9' : '4';
                }
                return rawKey;
            }

            function isExcludedOperationProdNo(prodNo) {
                const normalized = String(prodNo || '').trim().toUpperCase();
                return normalized === 'R1090' || normalized === 'R8200';
            }

            function escapeHtml(value) {
                return String(value || '')
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
            }

            function formatCount(value) {
                const num = Math.max(0, Math.trunc(Number(value) || 0));
                return new Intl.NumberFormat('da-DK', { maximumFractionDigits: 0 }).format(num);
            }

            function getResourceDisplayLabel(prodNo, descr) {
                const code = String(prodNo || '').trim();
                const name = String(descr || '').trim();
                if (!code) return name || '-';
                if (!name) return code;
                if (code.toUpperCase().startsWith('R')) {
                    return code + ' - ' + name;
                }
                return code;
            }

            function collectWarningMessages(item, fallbackText) {
                const unique = [];
                const pushValue = (value) => {
                    const chunks = String(value || '').split('|');
                    for (const chunk of chunks) {
                        const text = String(chunk || '').trim();
                        if (text && !unique.includes(text)) unique.push(text);
                    }
                };

                if (Array.isArray(item)) {
                    for (const entry of item) {
                        if (!entry) continue;
                        if (entry.WarningText) pushValue(entry.WarningText);
                        if (entry.warningText) pushValue(entry.warningText);
                    }
                } else if (item) {
                    if (item.WarningText) pushValue(item.WarningText);
                    if (item.warningText) pushValue(item.warningText);
                }

                if (unique.length === 0 && fallbackText) pushValue(fallbackText);
                return unique;
            }

            function getWarningIconMeta(message) {
                const text = String(message || '').trim().toLowerCase();
                if (text.includes('faktura') || text.includes('noinvo')) {
                    return { key: 'invoice', icon: '🧾' };
                }
                if (text.includes('tilknyttet produktionsordre') || text.includes('underliggende produktionsordre')) {
                    return { key: 'linked-order', icon: '🏭' };
                }
                if (text.includes('inkonsekvens') || text.includes('afvig')) {
                    return { key: 'consistency', icon: '⚠️' };
                }
                return { key: 'general', icon: '⚠️' };
            }

            function getWarningFlagHtml(item, fallbackText) {
                const hasWarning = Array.isArray(item)
                    ? item.some(entry => entry && (entry.HasWarning || entry.hasWarnings || entry.WarningText || entry.warningText))
                    : Boolean(item && (item.HasWarning || item.hasWarnings || item.WarningText || item.warningText));
                if (!hasWarning) return '';

                const messages = collectWarningMessages(item, fallbackText);
                if (messages.length === 0) return '';

                const grouped = new Map();
                for (const message of messages) {
                    const meta = getWarningIconMeta(message);
                    if (!grouped.has(meta.key)) {
                        grouped.set(meta.key, { icon: meta.icon, messages: [] });
                    }
                    grouped.get(meta.key).messages.push(message);
                }

                return Array.from(grouped.values()).map(group => {
                    const title = escapeHtml(group.messages.join(' | '));
                    return ' <span class="warning-flag" title="' + title + '">' + group.icon + '</span>';
                }).join('');
            }

            function getTimeAdjustmentFlagHtml(item, fallbackText) {
                if (!item || (!item.UsesEstimatedOperationTime && !item.hasEstimatedOperationTime)) return '';
                const title = escapeHtml(item.EstimatedTimeText || item.estimatedTimeText || fallbackText || 'Færdigmeldt minutter var 0 og er beregnet ud fra Stykliste Minutter.');
                return ' <span class="warning-flag" title="' + title + '">🕒</span>';
            }

            function getInvoiceStatusFlagHtml(item, forceShow = false) {
                if (!item) return '';
                const isTracked = Boolean(forceShow || item.IsInvoiceTracked || item.isInvoiceTracked || isInvoiceTrackedProdNo(item.ProdNo));
                if (!isTracked) return '';
                const noInvoValue = Number(item.NoInvo || 0);
                const noFinValue = Number(item.NoFin || 0);
                const hasMissing = Boolean(item.UsesMissingInvoiceFallback || item.usesMissingInvoiceFallback || String(item.MissingInvoiceText || item.missingInvoiceText || '').trim() || (noInvoValue === 0 && noFinValue > 0));
                const hasInvoice = item.HasInvoice === true || item.hasInvoice === true || noInvoValue > 0;
                const warningText = String(item.WarningText || item.warningText || '').toLowerCase();
                if (hasMissing && (warningText.includes('faktura') || warningText.includes('noinvo'))) {
                    return '';
                }
                const title = escapeHtml(item.InvoiceStatusText || item.invoiceStatusText || item.MissingInvoiceText || item.missingInvoiceText || (hasMissing
                    ? 'Mangler faktura; NoInvo er 0 og NoFin bruges til kostberegning.'
                    : (hasInvoice ? ('Faktura registreret: NoInvo = ' + noInvoValue + '.') : 'Ingen fakturainfo fundet.')));
                const icon = hasMissing ? '🧾' : (hasInvoice ? '📄' : '❔');
                return ' <span class="warning-flag" title="' + title + '">' + icon + '</span>';
            }

            function getInvoiceStatusSummaryHtml(lines, forceShow = false) {
                if (!Array.isArray(lines) || lines.length === 0) return '';
                const trackedLines = lines.filter(line => line && (forceShow || line.IsInvoiceTracked || line.isInvoiceTracked || isInvoiceTrackedProdNo(line.ProdNo)));
                if (trackedLines.length === 0) return '';
                const missingLine = trackedLines.find(line => {
                    if (!line) return false;
                    if (line.UsesMissingInvoiceFallback || line.usesMissingInvoiceFallback) return true;
                    const noInvoValue = Number(line.NoInvo || 0);
                    const noFinValue = Number(line.NoFin || 0);
                    return noInvoValue === 0 && noFinValue > 0;
                });
                const referenceLine = missingLine || trackedLines[0];
                const noInvoValue = Number((referenceLine && referenceLine.NoInvo) || 0);
                const hasInvoice = Boolean((referenceLine && (referenceLine.HasInvoice === true || referenceLine.hasInvoice === true)) || noInvoValue > 0);
                const cssClass = missingLine ? 'warn' : 'ok';
                const icon = missingLine ? '🧾' : (hasInvoice ? '📄' : '❔');
                const text = escapeHtml((referenceLine && (referenceLine.InvoiceStatusText || referenceLine.invoiceStatusText || referenceLine.MissingInvoiceText || referenceLine.missingInvoiceText)) || (missingLine
                    ? 'Mangler faktura; NoInvo er 0 og NoFin bruges til kostberegning.'
                    : (hasInvoice ? ('Faktura registreret: NoInvo = ' + noInvoValue + '.') : 'Ingen fakturainfo fundet.')));
                return '<div class="invoice-status-banner ' + cssClass + '">' + icon + ' ' + text + '</div>';
            }

            function getLaserAllocationFlagHtml(item, fallbackText) {
                if (!item || (!item.UsesLaserAllocationSpread && !item.usesLaserAllocationSpread)) return '';
                const title = escapeHtml(item.LaserAllocationText || item.laserAllocationText || fallbackText || 'Laserkosten er fordelt på et andet antal stk end denne ordrelinje, så pris pr. stk kan afvige.');
                return ' <span class="allocation-flag" title="' + title + '">*</span>';
            }

            const laserNestCostHints = new Map();

            function setLaserNestCostHint(ordNo, prodNo, nestingCost) {
                const numericOrdNo = Number(ordNo || 0);
                const normalizedProdNo = String(prodNo || '').trim().toUpperCase();
                const numericCost = Number(nestingCost || 0);
                if (!numericOrdNo || !normalizedProdNo || !(numericCost > 0)) return;
                laserNestCostHints.set(numericOrdNo + '|' + normalizedProdNo, numericCost);
            }

            function getLaserNestCostHint(ordNo, prodNo) {
                const numericOrdNo = Number(ordNo || 0);
                const normalizedProdNo = String(prodNo || '').trim().toUpperCase();
                if (!numericOrdNo || !normalizedProdNo) return null;
                const value = laserNestCostHints.get(numericOrdNo + '|' + normalizedProdNo);
                return Number.isFinite(Number(value)) && Number(value) > 0 ? Number(value) : null;
            }

            function toDrawingUrl(rawPath) {
                const value = String(rawPath || '').trim();
                if (!value) return '';
                const lower = value.toLowerCase();
                if (lower.startsWith('http://') || lower.startsWith('https://') || lower.startsWith('file://')) return value;

                const bs = String.fromCharCode(92);
                if (value.startsWith(bs + bs)) {
                    const uncPath = value.slice(2).split(bs).join('/');
                    return 'file://' + encodeURI(uncPath);
                }

                const normalized = value.split(bs).join('/');
                const hasDrivePrefix = normalized.length >= 3
                    && ((normalized[0] >= 'A' && normalized[0] <= 'Z') || (normalized[0] >= 'a' && normalized[0] <= 'z'))
                    && normalized[1] === ':'
                    && normalized[2] === '/';
                if (hasDrivePrefix) {
                    return 'file:///' + encodeURI(normalized);
                }

                return encodeURI(normalized);
            }

            function openDrawingPdf(pathOrMeta) {
                const meta = (pathOrMeta && typeof pathOrMeta === 'object') ? pathOrMeta : { path: pathOrMeta };
                const value = String(meta.path || '').trim();
                const prodNo = String(meta.prodNo || '').trim();
                const ordNo = String(meta.ordNo || '').trim();
                if (!value && !prodNo) return;
                fetch('/open-drawing', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ path: value, prodNo, ordNo })
                })
                .then(async (r) => {
                    if (r.ok) return;
                    let msg = 'Kunne ikke aabne tegning.';
                    try {
                        const d = await r.json();
                        if (d && d.message) msg = d.message;
                    } catch (_) {}
                    throw new Error(msg);
                })
                .catch((err) => {
                    const url = toDrawingUrl(value);
                    if (url) {
                        window.open(url, '_blank');
                    } else {
                        alert('Fejl ved åbning af tegning: ' + err.message);
                    }
                });
            }

            function toggleLaserOrderSummary() {
                const panel = document.getElementById('laserOrderSummaryPanel');
                const btn = document.getElementById('laserOrderSummaryToggleBtn');
                if (!panel || !btn) return;
                const isClosed = panel.style.display === 'none';
                panel.style.display = isClosed ? '' : 'none';
                if (!isClosed) {
                    const laserImagePanel = document.getElementById('laserImagePanel');
                    if (laserImagePanel) {
                        laserImagePanel.innerHTML = '';
                        laserImagePanel.classList.add('hidden');
                    }
                }
                btn.textContent = isClosed ? 'Skjul laseroversigt' : 'Vis laseroversigt';
            }

            function toggleOperationOrderSummary() {
                const panel = document.getElementById('operationOrderSummaryPanel');
                const btn = document.getElementById('operationOrderSummaryToggleBtn');
                if (!panel || !btn) return;
                const isClosed = panel.style.display === 'none';
                panel.style.display = isClosed ? '' : 'none';
                btn.textContent = isClosed ? 'Skjul operationer' : 'Vis operationer';
            }

            function buildOversigtModalView(type) {
                const isLaser = type === 'laser';
                const totalsId = isLaser ? 'laserOrderSummaryTotals' : 'operationOrderSummaryTotals';
                const bodyId = isLaser ? 'laserOrderSummaryBody' : 'operationOrderSummaryBody';
                const titleEl = document.getElementById('oversigtModalTitle');
                const subtitleEl = document.getElementById('oversigtModalSubtitle');
                const modalBody = document.getElementById('oversigtModalBody');
                const totals = document.getElementById(totalsId);
                const body = document.getElementById(bodyId);
                if (!titleEl || !subtitleEl || !modalBody || !totals || !body) return;

                titleEl.textContent = isLaser ? 'Laseroversigt (L-linjer)' : 'Operation Oversigt';
                subtitleEl.textContent = isLaser
                    ? 'Nesting, vægt og kost i samlet driftsvisning'
                    : 'Operationstid, kapacitet og kost i samlet driftsvisning';

                modalBody.innerHTML = ''
                    + '<div class="oversigt-modal-layout">'
                    +   '<section class="oversigt-panel oversigt-kpi"><h5>Samlede KPI</h5>' + totals.innerHTML + '</section>'
                    +   '<section class="oversigt-panel oversigt-details"><h5>Detaljer</h5>' + body.innerHTML + '</section>'
                    + '</div>';
                applyMicroTablePolish(modalBody);
            }

            function openOversigtModal(type) {
                currentOversigtModalType = (type === 'operation') ? 'operation' : 'laser';
                const modal = document.getElementById('oversigtModal');
                if (!modal) return;
                buildOversigtModalView(currentOversigtModalType);
                modal.style.display = 'flex';
            }

            function closeOversigtModal(event) {
                if (event && event.target && event.target.id !== 'oversigtModal') return;
                const modal = document.getElementById('oversigtModal');
                const body = document.getElementById('oversigtModalBody');
                if (modal) modal.style.display = 'none';
                if (body) body.innerHTML = '';
                currentOversigtModalType = null;
            }

            function refreshActiveOversigtModal() {
                if (!currentSearchOrderData) return;
                if (currentOversigtModalType === 'laser') {
                    loadSalesOrderLaserSummary(currentSearchOrderData);
                } else if (currentOversigtModalType === 'operation') {
                    loadSalesOrderOperationSummary(currentSearchOrderData);
                }
            }

            let currentMarginMode = 'classic';
            let orderListData = [];
            let orderListVisible = true;
            const ORDER_LIST_DAYS_BACK_CLIENT = 30;
            const AFTERCALC_CLIENT_CACHE_TTL_MS = 2 * 60 * 1000;
            let activeSearchRequestId = 0;
            let prefetchOrderDebounceTimer = null;
            let currentSearchOrderData = null;
            let lastOrderReportHtml = '';
            let lastOrderReportTitle = 'Rapport';
            let reportOriginState = null;
            let orderListFilter = '';
            let orderListBrugerFilter = '';
            let orderListMinDkkEnabled = false;
            let orderListMinDkkValue = 0;
            let marginStateByOrdNo = {};
            let marginJobQueue = [];
            let marginWorkerActiveCount = 0;
            let orderListRerenderTimer = null;
            let orderListLoading = false;
            let orderListAutoRefreshTimer = null;
            let orderListSortField = 'date';
            let orderListSortDir = 'desc';
            let marginSortRefreshTimer = null;
            let currentOversigtModalType = null;
            const aftercalcClientCache = new Map();
            const routeMetricsClientCache = new Map();
            const ROUTE_METRICS_CLIENT_CACHE_TTL_MS = 2 * 60 * 1000;

            function normalizeOrdNoValue(ordNo) {
                return String(ordNo || '').trim();
            }

            function pruneAftercalcClientCache() {
                if (aftercalcClientCache.size <= 80) return;
                const keys = Array.from(aftercalcClientCache.keys());
                for (let i = 0; i < keys.length - 80; i++) {
                    aftercalcClientCache.delete(keys[i]);
                }
            }

            function pruneRouteMetricsClientCache() {
                if (routeMetricsClientCache.size <= 100) return;
                const keys = Array.from(routeMetricsClientCache.keys());
                for (let i = 0; i < keys.length - 100; i++) {
                    routeMetricsClientCache.delete(keys[i]);
                }
            }

            async function requestRouteMetricsData(endpoint, options = {}) {
                const cacheKey = String(endpoint || '').trim();
                if (!cacheKey) throw new Error('Route metrics endpoint mangler');

                const forceReload = Boolean(options.forceReload);
                const now = Date.now();
                const existing = routeMetricsClientCache.get(cacheKey);

                if (!forceReload && existing) {
                    if (existing.data && (now - Number(existing.ts || 0)) < ROUTE_METRICS_CLIENT_CACHE_TTL_MS) {
                        return existing.data;
                    }
                    if (existing.promise) {
                        return existing.promise;
                    }
                }

                const fetchPromise = (async () => {
                    const response = await fetch(cacheKey);
                    const data = await response.json();
                    if (!response.ok || (data && data.error)) {
                        throw new Error((data && data.error) ? data.error : ('HTTP ' + response.status));
                    }
                    routeMetricsClientCache.set(cacheKey, { data, ts: Date.now(), promise: null });
                    pruneRouteMetricsClientCache();
                    return data;
                })();

                routeMetricsClientCache.set(cacheKey, { data: null, ts: now, promise: fetchPromise });
                try {
                    return await fetchPromise;
                } catch (err) {
                    routeMetricsClientCache.delete(cacheKey);
                    throw err;
                }
            }

            function buildLaserRouteMetricsEndpoint(ordine, route, prodNo, showAllRoutes) {
                return '/laser-route-metrics?ordine=' + encodeURIComponent(String(ordine || '').trim())
                    + (showAllRoutes ? '' : ('&route=' + encodeURIComponent(String(route || '').trim())))
                    + '&prodNo=' + encodeURIComponent(String(prodNo || '').trim())
                    + '&showAllRoutes=' + (showAllRoutes ? '1' : '0')
                    + (currentSalesOrderGr4 === 3 ? '&gr4=3' : '');
            }

            async function prefetchRouteMetricsForProduct(prodNo, ordNo, trInf2, trInf4, showAllRoutes) {
                if (!prodNo) return;
                const effectiveOrdine = String(ordNo || trInf2 || '').trim();
                if (!effectiveOrdine) return;

                let effectiveRoute = String(trInf4 || '').trim();
                if (!showAllRoutes && !effectiveRoute) {
                    try {
                        const fallbackResponse = await fetch('/nesting-detail/' + encodeURIComponent(effectiveOrdine) + '/' + encodeURIComponent(prodNo));
                        const fallbackRows = await fallbackResponse.json();
                        if (fallbackResponse.ok && Array.isArray(fallbackRows) && fallbackRows.length > 0) {
                            effectiveRoute = String(fallbackRows[0].TrInf4 || '').trim();
                        }
                    } catch (_) {}
                }

                if (!showAllRoutes && !effectiveRoute) return;
                const endpoint = buildLaserRouteMetricsEndpoint(effectiveOrdine, effectiveRoute, prodNo, showAllRoutes);
                requestRouteMetricsData(endpoint).catch(() => {});
            }

            async function requestAftercalcData(ordNo, options = {}) {
                const normalizedOrdNo = normalizeOrdNoValue(ordNo);
                if (!normalizedOrdNo) throw new Error('Ordrenummer mangler');

                const forceReload = Boolean(options.forceReload);
                const now = Date.now();
                const cacheKey = normalizedOrdNo;
                const existing = aftercalcClientCache.get(cacheKey);

                if (!forceReload && existing) {
                    if (existing.data && (now - Number(existing.ts || 0)) < AFTERCALC_CLIENT_CACHE_TTL_MS) {
                        return existing.data;
                    }
                    if (existing.promise) {
                        return existing.promise;
                    }
                }

                const fetchPromise = (async () => {
                    const response = await fetch('/aftercalc/' + encodeURIComponent(normalizedOrdNo));
                    const data = await response.json();
                    if (!response.ok) {
                        throw new Error((data && data.error) ? data.error : ('HTTP ' + response.status));
                    }
                    aftercalcClientCache.set(cacheKey, { data, ts: Date.now(), promise: null });
                    pruneAftercalcClientCache();
                    return data;
                })();

                aftercalcClientCache.set(cacheKey, { data: null, ts: now, promise: fetchPromise });

                try {
                    return await fetchPromise;
                } catch (err) {
                    aftercalcClientCache.delete(cacheKey);
                    throw err;
                }
            }

            function prefetchAftercalcData(ordNo) {
                const normalizedOrdNo = normalizeOrdNoValue(ordNo);
                if (!normalizedOrdNo) return;
                requestAftercalcData(normalizedOrdNo).catch(() => {});
            }

            function applyMicroTablePolish(rootEl) {
                const root = rootEl || document;
                const tables = Array.from(root.querySelectorAll('table'));
                if (!tables.length) return;

                const rightPattern = /færdigmeldt|minutter|min.|kost|pris|margin|afvigelse|kg|%|beløb|antal|samlet|dkk|forbrugt|stykliste|icon vægt|nestkost|nestmulti/i;
                const centerPattern = /linje|rute|prodtp4|prod.ordre|prodordre|nestingordre|ordre$/i;
                const leftPattern = /produkt|beskrivelse|kunde|type|linjer\/ref|hvem|status|message|beskrivelse/i;

                for (const table of tables) {
                    table.classList.add('micro-grid-table');
                    const headerCells = Array.from(table.querySelectorAll('tr:first-child th'));
                    if (!headerCells.length) continue;

                    const alignByIndex = [];
                    for (let i = 0; i < headerCells.length; i++) {
                        const text = String(headerCells[i].textContent || '').trim().toLowerCase();
                        let align = 'left';
                        if (rightPattern.test(text)) {
                            align = 'right';
                        } else if (centerPattern.test(text)) {
                            align = 'center';
                        } else if (leftPattern.test(text)) {
                            align = 'left';
                        }
                        alignByIndex[i] = align;
                    }

                    const rows = Array.from(table.querySelectorAll('tr'));
                    for (const row of rows) {
                        const cells = Array.from(row.children);
                        for (let i = 0; i < cells.length; i++) {
                            const align = alignByIndex[i] || 'left';
                            cells[i].style.textAlign = align;
                            if (align === 'right') {
                                cells[i].style.fontVariantNumeric = 'tabular-nums';
                            }
                        }
                    }
                }
            }

            // ── ORDER NOTES ────────────────────────────────────────────────
            let orderNotesCache = {};  // ordNo(string) -> { status, text, updatedAt }

            async function loadAllNotes() {
                try {
                    const r = await fetch('/order-notes-all');
                    if (r.ok) orderNotesCache = await r.json();
                } catch {}
            }

            async function loadOrderNote(ordNo) {
                const numericOrdNo = Number(ordNo || 0);
                if (!numericOrdNo) return null;
                try {
                    const r = await fetch('/order-note/' + numericOrdNo);
                    if (!r.ok) return null;
                    const note = await r.json();
                    orderNotesCache[String(numericOrdNo)] = note || { status: '', text: '', updatedAt: null };
                    renderOrderNoteBanner(numericOrdNo);
                    updateOrderNoteCell(numericOrdNo);
                    return note;
                } catch {
                    return null;
                }
            }

            function getOrderNoteHtml(ordNo) {
                const note = orderNotesCache[String(ordNo)];
                if (!note || (!note.status && !note.text && !note.isCreditNote)) return '<span style="color:#bbb;font-size:12px;">-</span>';
                const icons = { ok: '✅', error: '❌', check: '⚠️', credit: '🧾' };
                const icon = note.isCreditNote ? icons.credit : (icons[note.status] || '📝');
                const cls = note.isCreditNote ? 'credit' : (note.status || 'text');
                const preview = note.text ? escapeHtmlFE(note.text.slice(0, 40)) + (note.text.length > 40 ? '…' : '') : '';
                return '<span class="note-badge ' + cls + '" onclick="event.stopPropagation();openNotePopup(' + Number(ordNo) + ')">'
                    + icon + (note.isCreditNote ? ' Kreditnota' : '') + (preview ? ' ' + preview : '') + '</span>';
            }

            function isOrderMarkedCreditNote(ordNo) {
                const note = orderNotesCache[String(ordNo)];
                return Boolean(note && note.isCreditNote === true);
            }

            function escapeHtmlFE(s) {
                return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
            }

            function openNotePopup(ordNo, fromOrderDetail = false) {
                const note = orderNotesCache[String(ordNo)] || { status: '', text: '' };
                const existing = document.getElementById('notePopupOverlay');
                if (existing) existing.remove();

                const overlay = document.createElement('div');
                overlay.id = 'notePopupOverlay';
                overlay.className = 'note-popup-overlay';
                overlay.innerHTML =
                    '<div class="note-popup">' +
                    '<h3>📝 Note for ordre <strong>' + ordNo + '</strong></h3>' +
                    '<label>Status</label>' +
                    '<select id="noteStatusSel">' +
                    '<option value="">— ingen status —</option>' +
                    '<option value="ok">✅ OK</option>' +
                    '<option value="error">❌ Fejl</option>' +
                    '<option value="check">⚠️ Tjek</option>' +
                    '</select>' +
                    '<label style="display:flex;align-items:center;gap:8px;margin:-2px 0 10px 0;font-weight:600;">' +
                    '<input id="noteCreditChk" type="checkbox" ' + (note.isCreditNote ? 'checked' : '') + ' style="width:16px;height:16px;" />' +
                    'Kreditnota (udeluk fra samlet resoconto)' +
                    '</label>' +
                    '<label>Note</label>' +
                    '<textarea id="noteTextArea" placeholder="Skriv en note til denne ordre...">' + escapeHtmlFE(note.text || '') + '</textarea>' +
                    (note.updatedAt ? '<div style="font-size:11px;color:#888;margin-bottom:10px;">Sidst opdateret: ' + note.updatedAt.slice(0,16).replace('T',' ') + '</div>' : '') +
                    '<div class="note-popup-actions">' +
                    '<button class="btn-note-delete" onclick="deleteOrderNote(' + ordNo + ',' + fromOrderDetail + ')">Slet</button>' +
                    '<button class="btn-note-cancel" onclick="document.getElementById(\'notePopupOverlay\').remove()">Annuller</button>' +
                    '<button class="btn-note-save" onclick="saveOrderNote(' + ordNo + ',' + fromOrderDetail + ')">Gem</button>' +
                    '</div></div>';

                document.body.appendChild(overlay);
                overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
                document.getElementById('noteStatusSel').value = note.status || '';
            }

            async function saveOrderNote(ordNo, fromOrderDetail) {
                const status = document.getElementById('noteStatusSel').value;
                const text = document.getElementById('noteTextArea').value.trim();
                const isCreditNote = Boolean(document.getElementById('noteCreditChk') && document.getElementById('noteCreditChk').checked);
                try {
                    const r = await fetch('/order-note/' + ordNo, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ status, text, isCreditNote })
                    });
                    if (r.ok) {
                        const note = await r.json();
                        orderNotesCache[String(ordNo)] = note;
                    }
                } catch {}
                document.getElementById('notePopupOverlay').remove();
                updateOrderNoteCell(ordNo);
                if (fromOrderDetail) renderOrderNoteBanner(ordNo);
                if (orderListVisible) renderOrderList();
            }

            async function deleteOrderNote(ordNo, fromOrderDetail) {
                try {
                    await fetch('/order-note/' + ordNo, { method: 'DELETE' });
                    delete orderNotesCache[String(ordNo)];
                } catch {}
                document.getElementById('notePopupOverlay').remove();
                updateOrderNoteCell(ordNo);
                if (fromOrderDetail) renderOrderNoteBanner(ordNo);
                if (orderListVisible) renderOrderList();
            }

            function updateOrderNoteCell(ordNo) {
                const listEl = document.getElementById('orderList');
                if (!listEl) return;
                const cells = listEl.querySelectorAll('.order-note-cell[data-ordno="' + ordNo + '"]');
                const html = getOrderNoteHtml(ordNo);
                for (const cell of cells) { cell.innerHTML = html; }
                updateOrderListSummaryPanel();
            }

            function renderOrderNoteBanner(ordNo) {
                const el = document.getElementById('order-note-banner-' + ordNo);
                if (!el) return;
                const note = orderNotesCache[String(ordNo)];
                if (!note || (!note.status && !note.text && !note.isCreditNote)) { el.style.display = 'none'; return; }
                const icons = { ok: '✅', error: '❌', check: '⚠️', credit: '🧾' };
                const icon = note.isCreditNote ? icons.credit : (icons[note.status] || '📝');
                const cls = note.isCreditNote ? 'credit' : (note.status || 'text');
                el.className = 'order-note-banner ' + cls;
                el.style.display = 'flex';
                const label = note.isCreditNote
                    ? 'Kreditnota'
                    : (note.status === 'ok' ? 'OK' : note.status === 'error' ? 'Fejl' : note.status === 'check' ? 'Tjek' : 'Note');
                el.innerHTML = '<span class="note-icon">' + icon + '</span><div class="note-body"><strong>' +
                    label +
                    '</strong>' + (note.text ? ': ' + escapeHtmlFE(note.text) : '') + '</div>' +
                    '<span style="font-size:11px;opacity:0.7;margin-left:auto;cursor:pointer;" onclick="openNotePopup(' + ordNo + ',true)">✏️ Rediger</span>';
            }
            let summaryModalHistory = [];
            let summaryImageRegistry = {};
            let summaryImageRegistryCounter = 0;
            const ACCESS_CODE = '12345';
            let accessGranted = false;
            let loggedUserDisplayName = 'Bruger';
            let sideMenuOpen = false;
            let dashboardUpdatePollTimer = null;
            const MARGIN_MAX_CONCURRENT = 2;
            const MARGIN_QUEUE_DELAY_MS = 120;
            const MARGIN_FETCH_TIMEOUT_MS = 20000;
            const MARGIN_PREFETCH_ROWS = 150;
            const ORDER_LIST_AUTO_REFRESH_MS = 2 * 60 * 1000;
            let lastOrderListCheckTime = 0;
            let lastOrderListRemoteTime = 0;
            let omsaetningInitialized = false;
            let omsaetningAccounts = [];
            let omsaetningSelectedAccounts = new Set();
            let omsaetningCustomerResults = [];
            let omsaetningSelectedCustomers = new Map();
            let omsaetningCustomerSearchToken = 0;
            let omsaetningCustomerSearchTimer = null;
            let omsaetningSelectedFiscalYears = new Set();
            let omsaetningAutoReloadTimer = null;
            let omsaetningThresholdsByCustomer = new Map();
            let omsaetningDetailsCollapsed = true;
            let omsaetningAccountsPanelOpen = false;
            const OMSAETNING_SSRS_DEFAULT_ACCOUNTS = new Set(['11012', '11015', '11040']);
            const OMSAETNING_AUTO_RELOAD_DELAY_MS = 280;
            const OMSAETNING_SUMMARY_CACHE_TTL_MS = 15 * 60 * 1000;
            const OMSAETNING_CUSTOMER_SEARCH_CACHE_TTL_MS = 120000;
            const OMSAETNING_CACHE_MAX_ITEMS = 30;
            const OMSAETNING_SHOW_THRESHOLD_SECTION = false;
            let omsaetningThresholdLoadToken = 0;
            const OMSAETNING_DEFAULT_WARN_THRESHOLD = 3;
            const OMSAETNING_DEFAULT_GOOD_THRESHOLD = 5;
            let omsaetningSummaryCache = new Map();
            let omsaetningSummaryInFlight = new Map();
            let omsaetningCustomerSearchCache = new Map();
            let omsaetningCustomerSearchInFlight = new Map();
            let ordreindgangInitialized = false;
            let ordreindgangAutoReloadTimer = null;
            let ordreindgangSummaryCache = new Map();
            let ordreindgangSummaryInFlight = new Map();
            let ordreindgangLastPayload = null;
            let ordreindgangWeeklyCollapsed = true;
            let ordreindgangCustomersCollapsed = true;
            let ordreindgangResizeTimer = null;
            const ORDREINDGANG_AUTO_RELOAD_DELAY_MS = 280;
            const ORDREINDGANG_SUMMARY_CACHE_TTL_MS = 15 * 60 * 1000;
            let belastningInitialized = false;
            let belastningAutoReloadTimer = null;
            let belastningPeriodicTimer = null;
            let belastningLastPayload = null;
            let belastningSelectedDayKey = '';
            let belastningDetailContext = { resGr: '', parity: 1 };
            let belastningDraggedCardKey = '';
            const BELASTNING_FILTER_DEBOUNCE_MS = 280;
            let _belastningKundeSuggestTimer = null;
            let _belastningKundeResults = [];

            function scheduleBelastningKundeSuggest() {
                if (_belastningKundeSuggestTimer) clearTimeout(_belastningKundeSuggestTimer);
                _belastningKundeSuggestTimer = setTimeout(doBelastningKundeSuggest, 250);
            }

            function hideBelastningKundeSuggestions() {
                setTimeout(function() {
                    var d = document.getElementById('belastningKundeDropdown');
                    if (d) d.style.display = 'none';
                }, 200);
            }

            async function doBelastningKundeSuggest() {
                var inp = document.getElementById('belastningKunde');
                var q = inp ? inp.value.trim() : '';
                var d = document.getElementById('belastningKundeDropdown');
                if (!d) return;
                if (q.length < 2) { d.style.display = 'none'; return; }
                try {
                    var resp = await fetch('/omsaetning/customers?q=' + encodeURIComponent(q) + '&limit=15');
                    if (!resp.ok) return;
                    var data = await resp.json();
                    var results = Array.isArray(data.customers) ? data.customers : [];
                    if (!results.length) { d.style.display = 'none'; return; }
                    _belastningKundeResults = results;
                    d.innerHTML = results.map(function(r, i) {
                        var nm = escapeHtmlFE(String(r.name || ''));
                        var no = escapeHtmlFE(String(r.custNo || ''));
                        return '<div class="belastning-kunde-option" onmousedown="selectBelastningKundeOption(' + i + ',event)">'
                            + nm + '<span class="bko-sub">' + no + '</span></div>';
                    }).join('');
                    d.style.display = 'block';
                } catch(e) { /* silent */ }
            }

            function selectBelastningKundeOption(idx, e) {
                if (e) e.preventDefault();
                var r = _belastningKundeResults[idx];
                var name = r ? String(r.name || '') : '';
                var d = document.getElementById('belastningKundeDropdown');
                if (d) d.style.display = 'none';
                var inp = document.getElementById('belastningKunde');
                if (inp) {
                    inp.value = name;
                    inp.blur();
                }
                scheduleBelastningAutoReload();
            }
            const BELASTNING_PERIODIC_REFRESH_MS = 15 * 60 * 1000;

            function sanitizeDisplayName(name) {
                const safe = String(name || '').trim();
                return safe ? safe.slice(0, 32) : 'Bruger';
            }

            function updateHeaderGreeting() {
                const greeting = document.getElementById('headerUserGreeting');
                if (!greeting) return;
                greeting.textContent = 'Hej, ' + sanitizeDisplayName(loggedUserDisplayName);
            }

            function setLoggedUserDisplayName(name, persist = true) {
                loggedUserDisplayName = sanitizeDisplayName(name);
                if (persist) {
                    try {
                        localStorage.setItem('afterkalk_logged_user_name', loggedUserDisplayName);
                    } catch {}
                }
                updateHeaderGreeting();
            }

            function toggleSideMenu() {
                if (sideMenuOpen) {
                    closeSideMenu();
                } else {
                    openSideMenu();
                }
            }

            function openSideMenu() {
                const overlay = document.getElementById('sideMenuOverlay');
                if (!overlay) return;
                overlay.classList.add('open');
                sideMenuOpen = true;
                refreshSideMenuAuthState();
                const input = document.getElementById('sideMenuLoginInput');
                if (!accessGranted && input) {
                    setTimeout(() => input.focus(), 30);
                }
            }

            function closeSideMenu(event) {
                if (event && event.target && event.target.id !== 'sideMenuOverlay') return;
                const overlay = document.getElementById('sideMenuOverlay');
                if (!overlay) return;
                overlay.classList.remove('open');
                sideMenuOpen = false;
            }

            function refreshSideMenuAuthState() {
                const userInput = document.getElementById('sideMenuUserInput');
                const input = document.getElementById('sideMenuLoginInput');
                const loginBtn = document.getElementById('sideMenuLoginBtn');
                const status = document.getElementById('sideMenuAuthStatus');
                const logoutBtn = document.getElementById('sideMenuLogoutBtn');
                if (!status) return;

                if (accessGranted) {
                    status.textContent = 'Logget ind som ' + sanitizeDisplayName(loggedUserDisplayName) + '.';
                    status.classList.add('ok');
                    if (userInput) userInput.disabled = true;
                    if (input) {
                        input.value = '';
                        input.disabled = true;
                    }
                    if (loginBtn) loginBtn.disabled = true;
                    if (logoutBtn) logoutBtn.disabled = false;
                } else {
                    status.textContent = 'Ikke logget ind.';
                    status.classList.remove('ok');
                    if (userInput) userInput.disabled = false;
                    if (input) input.disabled = false;
                    if (loginBtn) loginBtn.disabled = false;
                    if (logoutBtn) logoutBtn.disabled = true;
                }
            }

            function submitAccessCodeFromSideMenu() {
                const sideUserInput = document.getElementById('sideMenuUserInput');
                const sideInput = document.getElementById('sideMenuLoginInput');
                const gateInput = document.getElementById('accessGateInput');
                if (sideUserInput) {
                    const desiredName = sanitizeDisplayName(sideUserInput.value);
                    setLoggedUserDisplayName(desiredName);
                }
                if (sideInput && gateInput) {
                    gateInput.value = sideInput.value || '';
                }
                submitAccessCode();
            }

            function navigateFromSideMenu(target) {
                if (target === 'dashboard') {
                    goToDashboard();
                    closeSideMenu();
                    return;
                }
                if (target === 'brugermanual') {
                    openBrugermanual();
                    closeSideMenu();
                    return;
                }
                if (target === 'personalehåndbog') {
                    openPersonalehåndbog();
                    closeSideMenu();
                    return;
                }
                openModule(target);
                closeSideMenu();
            }

            function logoutFromSideMenu() {
                accessGranted = false;
                setLoggedUserDisplayName('Bruger');
                closeSideMenu();
                goToDashboard();
                showAccessGate();
                const gateInput = document.getElementById('accessGateInput');
                if (gateInput) gateInput.value = '';
                refreshSideMenuAuthState();
            }

            function openBrugermanual() {
                const modal = document.getElementById('brugermanualModal');
                const body = document.getElementById('brugermanualBody');
                if (!modal || !body) return;
                body.innerHTML = ''
                    + '<section class="manual-card">'
                    + '<h4>1. Dashboard</h4>'
                    + '<p>Overblik over makrokategorier og hurtig adgang til moduler.</p>'
                    + '<ul><li>Brug kortene til at åbne modul.</li><li>Brug "Ryd Efterkalk cache" kun ved dataproblemer.</li><li>Warmup-status viser baggrundsindlæsning.</li></ul>'
                    + '</section>'
                    + '<section class="manual-card">'
                    + '<h4>2. Efterkalkulation</h4>'
                    + '<p>Ordreliste, margin, produktion og rapportdetaljer.</p>'
                    + '<ul><li>Klik på ordrelinje for fuld rapport.</li><li>"Opdater" på en ordre rydder cache for netop den ordre og henter frisk beregning.</li><li>Hvis tal ikke ændrer sig, er kilde-data sandsynligvis uændret.</li></ul>'
                    + '</section>'
                    + '<section class="manual-card">'
                    + '<h4>3. Omsætning</h4>'
                    + '<p>Periode-, konto- og kundebaseret omsætningsanalyse.</p>'
                    + '<ul><li>Vælg periode og konti.</li><li>Tryk "Opdater" for nye tal.</li><li>Print fra modulet efter opdatering.</li></ul>'
                    + '</section>'
                    + '<section class="manual-card">'
                    + '<h4>4. Ordreindgang</h4>'
                    + '<p>Ugevis ordre- og tilbudsoverblik.</p>'
                    + '<ul><li>Vælg ugeinterval (YYYYWW).</li><li>Tryk "Opdater".</li><li>Brug tabeller/grafer til opfølgning.</li></ul>'
                    + '</section>'
                    + '<section class="manual-card">'
                    + '<h4>5. Datadifferencer (NestKost)</h4>'
                    + '<p>NestKost pr. stk kan afvige, hvis færdigmeldt antal på ordrelinjen ikke matcher nesting-fordeling.</p>'
                    + '<ul><li>Pris pr. stk beregnes fra samme kilde som linjens totale kost.</li><li>Routedetaljer kan vise et andet antal pga. fordeling/split på ruter.</li></ul>'
                    + '</section>'
                    + '<div class="manual-meta">Tip: Brug side-menuen (☰) til hurtig navigation mellem moduler og manual.</div>';
                modal.style.display = 'flex';
            }

            function closeBrugermanual(event) {
                if (event && event.target && event.target.id !== 'brugermanualModal') return;
                const modal = document.getElementById('brugermanualModal');
                if (modal) modal.style.display = 'none';
            }

            const PH_BASE_URL = 'http://apv/GHB/';

            function openPersonalehåndbog() {
                const modal = document.getElementById('personalehåndbogsModal');
                const iframe = document.getElementById('personalehåndbogsIframe');
                const input = document.getElementById('personalehåndbogsSearchInput');
                if (!modal || !iframe) return;
                if (!iframe.src || iframe.src === 'about:blank' || iframe.src === window.location.href) {
                    iframe.src = PH_BASE_URL;
                }
                modal.classList.add('open');
                document.body.style.overflow = 'hidden';
                phCheckStatus();
                if (input) setTimeout(() => input.focus(), 150);
            }

            function closePersonalehåndbog() {
                const modal = document.getElementById('personalehåndbogsModal');
                const iframe = document.getElementById('personalehåndbogsIframe');
                if (modal) modal.classList.remove('open');
                if (iframe) iframe.src = '';
                document.body.style.overflow = '';
            }

            let qmsDataset = null;
            let qmsFlatDocs = [];
            let qmsSelectedDocId = null;
            let qmsEditMode = false;

            function escapeHtml(str) {
                return String(str || '')
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
            }

            function makeQmsId(prefix) {
                return String(prefix || 'id') + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
            }

            async function loadQmsDataset(force = false) {
                if (qmsDataset && !force) return qmsDataset;
                const r = await fetch('/qms/dataset');
                const data = await r.json();
                if (!data.ok || !data.dataset) throw new Error(data.error || 'QMS dataset fejl');
                qmsDataset = data.dataset;
                return qmsDataset;
            }

            async function saveQmsDataset() {
                const r = await fetch('/qms/dataset', {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ dataset: qmsDataset })
                });
                const data = await r.json();
                if (!data.ok) throw new Error(data.error || 'Kunne ikke gemme dataset');
                qmsDataset = data.dataset;
                return qmsDataset;
            }

            function flattenQmsDataset() {
                const out = [];
                if (!qmsDataset || !Array.isArray(qmsDataset.folders)) return out;
                for (const folder of qmsDataset.folders) {
                    const docs = Array.isArray(folder.documents) ? folder.documents : [];
                    for (const doc of docs) {
                        out.push({
                            folderId: folder.id,
                            folderName: folder.name,
                            folderDescription: folder.description || '',
                            id: doc.id,
                            title: doc.title,
                            url: doc.url || '',
                            content: doc.content || '',
                            tags: Array.isArray(doc.tags) ? doc.tags : []
                        });
                    }
                }
                qmsFlatDocs = out;
                return out;
            }

            function getSelectedQmsDoc() {
                return qmsFlatDocs.find(d => d.id === qmsSelectedDocId) || null;
            }

            function renderQmsView(doc) {
                const view = document.getElementById('qmsView');
                if (!view) return;
                if (!doc) {
                    view.innerHTML = '<h3>Kvalitetsledelsessystem</h3><div class="qms-view-meta">Vælg et dokument i venstre side.</div>';
                    return;
                }
                if (qmsEditMode) {
                    view.innerHTML = ''
                        + '<h3>Rediger dokument</h3>'
                        + '<div class="qms-view-meta">' + escapeHtml(doc.folderName) + '</div>'
                        + '<div class="qms-editor">'
                        + '<label>Titel</label><input id="qmsEditTitle" value="' + escapeHtml(doc.title) + '" />'
                        + '<label>URL (valgfri)</label><input id="qmsEditUrl" value="' + escapeHtml(doc.url) + '" />'
                        + '<label>Indhold</label><textarea id="qmsEditContent">' + escapeHtml(doc.content) + '</textarea>'
                        + '<div class="qms-editor-actions">'
                        + '<button class="save" onclick="qmsSaveCurrentDoc()">Gem dokument</button>'
                        + '<button class="delete" onclick="qmsDeleteCurrentDoc()">Slet dokument</button>'
                        + '<button class="cancel" onclick="toggleQmsEditMode(false)">Afslut redigering</button>'
                        + '</div>'
                        + '</div>';
                    return;
                }
                view.innerHTML = ''
                    + '<h3>' + escapeHtml(doc.title) + '</h3>'
                    + '<div class="qms-view-meta">' + escapeHtml(doc.folderName) + '</div>'
                    + '<div class="qms-view-content">' + escapeHtml(doc.content) + '</div>'
                    + (doc.url ? '<div class="qms-view-link"><a href="' + escapeHtml(doc.url) + '" target="_blank" rel="noopener noreferrer">Åbn original reference</a></div>' : '');
            }

            function renderQmsList(query = '') {
                const list = document.getElementById('qmsList');
                const label = document.getElementById('qmsListLabel');
                if (!list || !label) return;
                const q = String(query || '').trim().toLowerCase();
                const docs = flattenQmsDataset().filter(doc => {
                    if (!q) return true;
                    return (doc.title + ' ' + doc.folderName + ' ' + doc.content + ' ' + doc.tags.join(' ')).toLowerCase().includes(q);
                });
                label.textContent = docs.length + ' dokumenter';
                if (docs.length === 0) {
                    list.innerHTML = '<div class="qms-empty">Ingen dokumenter matcher din søgning.</div>';
                    renderQmsView(null);
                    return;
                }
                list.innerHTML = docs.map(doc => (
                    '<div class="qms-item" data-doc-id="' + escapeHtml(doc.id) + '" onclick="openQmsPage(this)">' +
                    '<div class="qms-item-title">' + escapeHtml(doc.title) + '</div>' +
                    '<div class="qms-item-meta">' + escapeHtml(doc.folderName) + '</div>' +
                    '</div>'
                )).join('');
                if (!qmsSelectedDocId || !docs.some(d => d.id === qmsSelectedDocId)) {
                    qmsSelectedDocId = docs[0].id;
                }
                const active = list.querySelector('.qms-item[data-doc-id="' + CSS.escape(qmsSelectedDocId) + '"]') || list.querySelector('.qms-item');
                if (active) openQmsPage(active);
            }

            async function openKvalitetsledelsessystem() {
                const modal = document.getElementById('qmsModal');
                const input = document.getElementById('qmsSearchInput');
                if (!modal) return;
                modal.classList.add('open');
                document.body.style.overflow = 'hidden';
                try {
                    await loadQmsDataset(false);
                    renderQmsList('');
                } catch (err) {
                    const list = document.getElementById('qmsList');
                    if (list) list.innerHTML = '<div class="qms-empty">Kunne ikke læse QMS dataset: ' + escapeHtml(err.message || '') + '</div>';
                    renderQmsView(null);
                }
                if (input) {
                    input.value = '';
                    setTimeout(() => input.focus(), 120);
                }
            }

            function closeQmsModal() {
                const modal = document.getElementById('qmsModal');
                if (modal) modal.classList.remove('open');
                document.body.style.overflow = '';
            }

            function searchQmsPages() {
                const input = document.getElementById('qmsSearchInput');
                renderQmsList(input ? input.value : '');
            }

            function openQmsPage(el) {
                const docId = el && el.getAttribute ? el.getAttribute('data-doc-id') : '';
                if (!docId) return;
                qmsSelectedDocId = docId;
                document.querySelectorAll('#qmsList .qms-item').forEach(x => x.classList.remove('active'));
                el.classList.add('active');
                renderQmsView(getSelectedQmsDoc());
            }

            function toggleQmsEditMode(force) {
                if (typeof force === 'boolean') {
                    qmsEditMode = force;
                } else {
                    qmsEditMode = !qmsEditMode;
                }
                const btn = document.getElementById('qmsEditToggleBtn');
                if (btn) btn.textContent = qmsEditMode ? 'Visning' : 'Rediger';
                renderQmsView(getSelectedQmsDoc());
            }

            async function qmsSaveCurrentDoc() {
                const doc = getSelectedQmsDoc();
                if (!doc) return;
                const title = document.getElementById('qmsEditTitle');
                const url = document.getElementById('qmsEditUrl');
                const content = document.getElementById('qmsEditContent');
                const folder = (qmsDataset.folders || []).find(f => f.id === doc.folderId);
                if (!folder) return;
                const target = (folder.documents || []).find(d => d.id === doc.id);
                if (!target) return;
                target.title = String(title && title.value || '').trim() || target.title;
                target.url = String(url && url.value || '').trim();
                target.content = String(content && content.value || '').trim();
                try {
                    await saveQmsDataset();
                    renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
                } catch (err) {
                    alert('Kunne ikke gemme: ' + (err.message || err));
                }
            }

            async function qmsDeleteCurrentDoc() {
                const doc = getSelectedQmsDoc();
                if (!doc) return;
                if (!confirm('Slet dokumentet "' + doc.title + '"?')) return;
                const folder = (qmsDataset.folders || []).find(f => f.id === doc.folderId);
                if (!folder) return;
                folder.documents = (folder.documents || []).filter(d => d.id !== doc.id);
                qmsSelectedDocId = null;
                try {
                    await saveQmsDataset();
                    renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
                } catch (err) {
                    alert('Kunne ikke slette: ' + (err.message || err));
                }
            }

            async function qmsCreateFolder() {
                try {
                    await loadQmsDataset(false);
                    const name = prompt('Navn på ny mappe:');
                    if (!name || !name.trim()) return;
                    qmsDataset.folders.push({
                        id: makeQmsId('folder'),
                        name: name.trim(),
                        description: '',
                        documents: []
                    });
                    await saveQmsDataset();
                    renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
                } catch (err) {
                    alert('Kunne ikke oprette mappe: ' + (err.message || err));
                }
            }

            async function qmsCreateDocument() {
                try {
                    await loadQmsDataset(false);
                    if (!Array.isArray(qmsDataset.folders) || qmsDataset.folders.length === 0) {
                        alert('Opret først en mappe.');
                        return;
                    }
                    const title = prompt('Titel på nyt dokument:');
                    if (!title || !title.trim()) return;
                    let folder = (qmsDataset.folders || []).find(f => f.id === qmsSelectedDocId) || null;
                    const selected = getSelectedQmsDoc();
                    if (selected) {
                        folder = (qmsDataset.folders || []).find(f => f.id === selected.folderId) || null;
                    }
                    if (!folder) folder = qmsDataset.folders[0];
                    folder.documents = Array.isArray(folder.documents) ? folder.documents : [];
                    const docId = makeQmsId('doc');
                    folder.documents.push({
                        id: docId,
                        title: title.trim(),
                        url: '',
                        content: '',
                        tags: []
                    });
                    qmsSelectedDocId = docId;
                    await saveQmsDataset();
                    renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
                    qmsEditMode = true;
                    toggleQmsEditMode(true);
                } catch (err) {
                    alert('Kunne ikke oprette dokument: ' + (err.message || err));
                }
            }

            function phSetStatus(msg) {
                const lbl = document.getElementById('phResultsLabel');
                const msgEl = document.getElementById('phStatusMsg');
                const list = document.getElementById('phResultsList');
                if (list) list.innerHTML = '<div class="ph-status-msg" id="phStatusMsg">' + msg + '</div>';
                if (lbl) lbl.textContent = 'Resultater';
            }

            async function phCheckStatus() {
                try {
                    const r = await fetch('/ph/status');
                    const d = await r.json();
                    if (d.status === 'indexing') {
                        phSetStatus('⏳ Indekserer sitet, vent venligst…');
                        setTimeout(phCheckStatus, 2000);
                    } else if (d.status === 'ready') {
                        phSetStatus('Skriv en søgning og tryk Søg.<br><small style="color:#9aabcc">' + d.count + ' sider indekseret</small>');
                    } else if (d.status === 'idle') {
                        phSetStatus('Indeks ikke klar. Tryk ↺ Genindekser.');
                    } else if (d.status === 'error') {
                        phSetStatus('⚠️ Fejl ved indeksering: ' + (d.error || ''));
                    }
                } catch { phSetStatus('Kunne ikke kontakte serveren.'); }
            }

            async function phReindex() {
                phSetStatus('⏳ Indekserer sitet, vent venligst…');
                try {
                    await fetch('/ph/reindex', { method: 'POST' });
                    setTimeout(phCheckStatus, 1500);
                } catch { phSetStatus('⚠️ Fejl ved genindeksering.'); }
            }

            function phEscapeHtml(s) {
                return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
            }

            function phHighlight(text, terms) {
                let out = phEscapeHtml(text);
                for (const t of terms) {
                    if (!t) continue;
                    const lower = out.toLowerCase();
                    const tl = t.toLowerCase();
                    let i = 0, result = '', pos;
                    while ((pos = lower.indexOf(tl, i)) !== -1) {
                        result += out.slice(i, pos) + '<mark>' + out.slice(pos, pos + t.length) + '</mark>';
                        i = pos + t.length;
                    }
                    out = result + out.slice(i);
                }
                return out;
            }

            async function searchPersonalehåndbog() {
                const input = document.getElementById('personalehåndbogsSearchInput');
                const list  = document.getElementById('phResultsList');
                const lbl   = document.getElementById('phResultsLabel');
                if (!input || !list) return;
                const q = input.value.trim();
                if (!q) { phCheckStatus(); return; }
                phSetStatus('🔍 Søger…');
                try {
                    const r = await fetch('/ph/search?q=' + encodeURIComponent(q));
                    const d = await r.json();
                    if (d.status === 'indexing') { phSetStatus('⏳ Indekserer endnu, prøv igen om lidt…'); return; }
                    if (!d.results || d.results.length === 0) {
                        phSetStatus('Ingen resultater for <strong>' + phEscapeHtml(q) + '</strong>.');
                        if (lbl) lbl.textContent = '0 resultater';
                        return;
                    }
                    if (lbl) lbl.textContent = d.results.length + ' resultat' + (d.results.length !== 1 ? 'er' : '');
                    const terms = q.toLowerCase().split(/\s+/).filter(Boolean);
                    list.innerHTML = d.results.map((res, i) => {
                        const title = phHighlight(res.title || res.url, terms);
                        const snip  = phHighlight(res.snippet, terms);
                        const safeUrl = phEscapeHtml(res.url);
                        return '<div class="ph-result-item" data-url="' + safeUrl + '" onclick="phOpenResult(this)" title="' + safeUrl + '">'
                            + '<div class="ph-result-title">' + title + '</div>'
                            + '<div class="ph-result-url">' + safeUrl + '</div>'
                            + '<div class="ph-result-snippet">' + snip + '</div>'
                            + '</div>';
                    }).join('');
                    // Auto-load first result
                    const first = list.querySelector('.ph-result-item');
                    if (first) phOpenResult(first);
                } catch { phSetStatus('⚠️ Søgefejl. Prøv igen.'); }
            }

            function phOpenResult(el) {
                const url = el.getAttribute('data-url');
                if (!url) return;
                const iframe = document.getElementById('personalehåndbogsIframe');
                if (iframe) iframe.src = url;
                document.querySelectorAll('.ph-result-item').forEach(e => e.classList.remove('ph-active'));
                el.classList.add('ph-active');
            }

            function setOmsaetningCacheEntry(cacheMap, key, value) {
                cacheMap.set(key, {
                    ts: Date.now(),
                    value
                });
                if (cacheMap.size > OMSAETNING_CACHE_MAX_ITEMS) {
                    const oldestKey = cacheMap.keys().next().value;
                    if (oldestKey !== undefined) cacheMap.delete(oldestKey);
                }
            }

            function getOmsaetningCacheEntry(cacheMap, key, ttlMs) {
                const hit = cacheMap.get(key);
                if (!hit) return null;
                if ((Date.now() - Number(hit.ts || 0)) > ttlMs) {
                    cacheMap.delete(key);
                    return null;
                }
                return hit.value;
            }

            function buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, selectedCustomers) {
                const accounts = Array.from(new Set((Array.isArray(selectedAccounts) ? selectedAccounts : [])
                    .map(v => String(v || '').trim())
                    .filter(Boolean))).sort();
                const customers = Array.from(new Set((Array.isArray(selectedCustomers) ? selectedCustomers : [])
                    .map(v => String(v || '').trim())
                    .filter(Boolean))).sort();
                return JSON.stringify({ fra, til, accounts, customers });
            }

            async function fetchOmsaetningSummaryCached(fra, til, selectedAccounts, selectedCustomers, options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const forceRefresh = safeOptions.forceRefresh === true;
                const cacheKey = buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, selectedCustomers);
                if (!forceRefresh) {
                    const cached = getOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, OMSAETNING_SUMMARY_CACHE_TTL_MS);
                    if (cached) {
                        return cached;
                    }
                }

                const customerFilters = Array.isArray(selectedCustomers)
                    ? Array.from(new Set(selectedCustomers.map(v => String(v || '').trim()).filter(Boolean))).sort()
                    : [];
                if (!forceRefresh && customerFilters.length > 0) {
                    const allCustomersKey = buildOmsaetningSummaryCacheKey(fra, til, selectedAccounts, []);
                    const allCustomersCached = getOmsaetningCacheEntry(omsaetningSummaryCache, allCustomersKey, OMSAETNING_SUMMARY_CACHE_TTL_MS);
                    if (allCustomersCached && Array.isArray(allCustomersCached.rows)) {
                        const selectedSet = new Set(customerFilters);
                        const filteredRows = allCustomersCached.rows.filter(row => {
                            const custNo = row && row.custNo !== null && row.custNo !== undefined
                                ? String(row.custNo).trim()
                                : '';
                            return custNo && selectedSet.has(custNo);
                        });
                        const derivedPayload = {
                            ok: true,
                            filters: {
                                fra,
                                til,
                                accounts: Array.isArray(allCustomersCached.filters && allCustomersCached.filters.accounts)
                                    ? allCustomersCached.filters.accounts
                                    : [],
                                customers: customerFilters
                            },
                            totalRevenueMio: filteredRows.reduce((sum, row) => sum + Number((row && row.revenueMio) || 0), 0),
                            rows: filteredRows
                        };
                        setOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, derivedPayload);
                        return derivedPayload;
                    }
                }

                if (!forceRefresh) {
                    const inFlight = omsaetningSummaryInFlight.get(cacheKey);
                    if (inFlight) {
                        return await inFlight;
                    }
                }

                const query = new URLSearchParams({
                    fra,
                    til,
                    accounts: selectedAccounts.join(','),
                    customers: selectedCustomers.join(',')
                });

                const reqPromise = (async () => {
                    const response = await fetch('/omsaetning/summary?' + query.toString());
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    setOmsaetningCacheEntry(omsaetningSummaryCache, cacheKey, payload);
                    return payload;
                })();

                omsaetningSummaryInFlight.set(cacheKey, reqPromise);
                try {
                    return await reqPromise;
                } finally {
                    omsaetningSummaryInFlight.delete(cacheKey);
                }
            }

            async function searchOmsaetningCustomersCached(q) {
                const key = String(q || '').trim().toLowerCase();
                if (!key) return { customers: [] };

                const cached = getOmsaetningCacheEntry(omsaetningCustomerSearchCache, key, OMSAETNING_CUSTOMER_SEARCH_CACHE_TTL_MS);
                if (cached) {
                    return cached;
                }

                const inFlight = omsaetningCustomerSearchInFlight.get(key);
                if (inFlight) {
                    return await inFlight;
                }

                const reqPromise = (async () => {
                    const response = await fetch('/omsaetning/customers?q=' + encodeURIComponent(key) + '&limit=25');
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    setOmsaetningCacheEntry(omsaetningCustomerSearchCache, key, payload);
                    return payload;
                })();

                omsaetningCustomerSearchInFlight.set(key, reqPromise);
                try {
                    return await reqPromise;
                } finally {
                    omsaetningCustomerSearchInFlight.delete(key);
                }
            }

            function applyOmsaetningDetailsCollapsedState() {
                const tableWrap = document.getElementById('omsaetningTableWrap');
                const toggleBtn = document.getElementById('omsaetningDetailsToggleBtn');
                if (!tableWrap || !toggleBtn) return;
                tableWrap.style.display = omsaetningDetailsCollapsed ? 'none' : 'block';
                toggleBtn.textContent = omsaetningDetailsCollapsed ? 'Vis detaljer' : 'Skjul detaljer';
            }

            function renderOmsaetningAccountsSummary() {
                const summaryEl = document.getElementById('omsaetningAccountsSummary');
                const activeEl = document.getElementById('omsaetningAccountsActive');
                if (!summaryEl) return;
                const selectedCount = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean).length;
                const totalCount = Array.isArray(omsaetningAccounts) ? omsaetningAccounts.length : 0;
                summaryEl.textContent = String(selectedCount) + '/' + String(totalCount) + ' valgt';

                if (activeEl) {
                    const selectedValues = Array.from(omsaetningSelectedAccounts.values())
                        .map(v => String(v || '').trim())
                        .filter(Boolean)
                        .sort((a, b) => a.localeCompare(b));

                    if (selectedValues.length === 0) {
                        activeEl.innerHTML = '<span class="chip more">Ingen konti valgt</span>';
                    } else {
                        const visible = selectedValues.slice(0, 5).map(acNo => {
                            const account = Array.isArray(omsaetningAccounts)
                                ? omsaetningAccounts.find(a => String(a && a.acNo || '').trim() === acNo)
                                : null;
                            const label = account ? (acNo + ' ' + String(account.name || '').trim()) : acNo;
                            return '<span class="chip" title="' + escapeHtmlFE(label) + '">' + escapeHtmlFE(acNo) + '</span>';
                        });
                        if (selectedValues.length > 5) {
                            visible.push('<span class="chip more">+' + escapeHtmlFE(String(selectedValues.length - 5)) + '</span>');
                        }
                        activeEl.innerHTML = visible.join('');
                    }
                }
            }

            function formatSigned(value, digits) {
                const n = Number(value || 0);
                const fixed = n.toFixed(Number.isFinite(digits) ? digits : 1);
                return (n > 0 ? '+' : '') + fixed;
            }

            function buildOmsaetningGaugeData(amountMio, warnThreshold, goodThreshold) {
                const amount = Number(amountMio || 0);
                const warn = Number(warnThreshold || 0);
                const good = Math.max(warn + 0.0001, Number(goodThreshold || 0));
                const span = Math.max(0.0001, good - warn);

                const marginPct = ((amount - warn) / span) * 30;
                const scaleMin = -30;
                const scaleMax = 60;
                const toLeft = value => ((value - scaleMin) / (scaleMax - scaleMin)) * 100;

                const zeroLeft = toLeft(0);
                const targetLeft = toLeft(30);
                const clamped = Math.max(scaleMin, Math.min(scaleMax, marginPct));
                const pointLeft = toLeft(clamped);

                return {
                    marginPct,
                    pointLeft,
                    zeroLeft,
                    targetLeft,
                    fillLeft: Math.min(pointLeft, zeroLeft),
                    fillWidth: Math.abs(pointLeft - zeroLeft),
                    fillClass: pointLeft >= zeroLeft ? 'pos' : 'neg',
                    deltaWarn: amount - warn,
                    deltaGood: amount - good
                };
            }

            function applyOmsaetningAccountsPanelState() {
                const panel = document.getElementById('omsaetningAccountsPanel');
                const search = document.getElementById('omsaetningAccountSearch');
                const btn = document.getElementById('omsaetningAccountsToggleBtn');
                if (panel) panel.style.display = omsaetningAccountsPanelOpen ? 'block' : 'none';
                if (search) search.style.display = omsaetningAccountsPanelOpen ? 'block' : 'none';
                if (btn) btn.textContent = omsaetningAccountsPanelOpen ? 'Skjul konti' : 'Vis konti';
            }

            function toggleOmsaetningAccountsPanel() {
                omsaetningAccountsPanelOpen = !omsaetningAccountsPanelOpen;
                applyOmsaetningAccountsPanelState();
            }

            function toggleOmsaetningDetails() {
                omsaetningDetailsCollapsed = !omsaetningDetailsCollapsed;
                applyOmsaetningDetailsCollapsedState();
            }

            function formatMio(value) {
                const numeric = Number(value || 0);
                return numeric.toLocaleString('da-DK', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
            }

            function formatDkkFromMio(valueMio) {
                const numeric = Number(valueMio || 0) * 1000000;
                return numeric.toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            }

            function formatMonthDa(dateValue) {
                if (!dateValue) return '-';
                const dt = new Date(dateValue);
                if (Number.isNaN(dt.getTime())) return String(dateValue);
                return dt.toLocaleDateString('da-DK', { month: 'short', year: 'numeric' });
            }

            function normalizeOmsaetningMonthKey(dateValue) {
                const dt = new Date(dateValue);
                if (Number.isNaN(dt.getTime())) return String(dateValue || '').trim();
                const year = dt.getFullYear();
                const month = String(dt.getMonth() + 1).padStart(2, '0');
                return String(year) + '-' + month + '-01';
            }

            function parseMonthInputToPeriod(monthValue) {
                const raw = String(monthValue || '').trim();
                const match = raw.match(/^(\d{4})-(\d{2})$/);
                if (!match) return null;
                const year = Number(match[1]);
                const month = Number(match[2]);
                if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return null;
                return {
                    year,
                    month,
                    period: year * 100 + month
                };
            }

            function calendarMonthToFiscalYrPr(year, month) {
                const y = Number(year);
                const m = Number(month);
                if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) return null;
                if (m >= 7) {
                    return (y * 100) + (m - 6);
                }
                return ((y - 1) * 100) + (m + 6);
            }

            function getCurrentFiscalYearStart(referenceDate) {
                const now = referenceDate instanceof Date ? referenceDate : new Date();
                const year = now.getFullYear();
                const month = now.getMonth() + 1;
                return month >= 7 ? year : (year - 1);
            }

            function getFiscalYearRange(yearValue) {
                const startYear = Number(yearValue);
                if (!Number.isFinite(startYear)) return null;
                return {
                    fromMonth: String(startYear) + '-07',
                    toMonth: String(startYear + 1) + '-06',
                    fra: String(startYear * 100 + 7),
                    til: String((startYear + 1) * 100 + 7)
                };
            }

            function applySelectedFiscalYearsToInputs() {
                const fromEl = document.getElementById('omsaetningFraMonth');
                const toEl = document.getElementById('omsaetningTilMonth');
                const selectedYears = Array.from(omsaetningSelectedFiscalYears.values())
                    .map(y => Number(y))
                    .filter(y => Number.isFinite(y))
                    .sort((a, b) => a - b);
                if (selectedYears.length === 0) return;

                const firstRange = getFiscalYearRange(selectedYears[0]);
                const lastRange = getFiscalYearRange(selectedYears[selectedYears.length - 1]);
                if (!firstRange || !lastRange) return;

                if (fromEl) fromEl.value = firstRange.fromMonth;
                if (toEl) toEl.value = lastRange.toMonth;
            }

            function buildOmsaetningPeriodRange() {
                const fromEl = document.getElementById('omsaetningFraMonth');
                const toEl = document.getElementById('omsaetningTilMonth');
                const fromMeta = parseMonthInputToPeriod(fromEl ? fromEl.value : '');
                const toMeta = parseMonthInputToPeriod(toEl ? toEl.value : '');
                if (!fromMeta || !toMeta) return null;

                const fromDate = new Date(fromMeta.year, fromMeta.month - 1, 1);
                const toDate = new Date(toMeta.year, toMeta.month - 1, 1);
                if (fromDate.getTime() > toDate.getTime()) return null;

                const exclusiveToDate = new Date(toMeta.year, toMeta.month, 1);
                const fraFiscal = calendarMonthToFiscalYrPr(fromMeta.year, fromMeta.month);
                const tilFiscal = calendarMonthToFiscalYrPr(exclusiveToDate.getFullYear(), exclusiveToDate.getMonth() + 1);
                if (!Number.isFinite(fraFiscal) || !Number.isFinite(tilFiscal)) return null;
                return {
                    fra: String(fraFiscal),
                    til: String(tilFiscal)
                };
            }

            function buildOmsaetningMonthKeys(periodRange) {
                if (!periodRange) return [];
                const fraNum = Number(periodRange.fra);
                const tilNum = Number(periodRange.til);
                if (!Number.isFinite(fraNum) || !Number.isFinite(tilNum) || fraNum >= tilNum) return [];

                let year = Math.floor(fraNum / 100);
                let month = fraNum % 100;
                const keys = [];

                while ((year * 100 + month) < tilNum) {
                    const mm = String(month).padStart(2, '0');
                    keys.push(String(year) + '-' + mm + '-01');
                    month += 1;
                    if (month > 12) {
                        month = 1;
                        year += 1;
                    }
                }
                return keys;
            }

            function renderOmsaetningYearChips(centerYear) {
                const wrap = document.getElementById('omsaetningYears');
                if (!wrap) return;

                const currentFiscalYear = Number(centerYear || getCurrentFiscalYearStart());
                const years = [currentFiscalYear - 3, currentFiscalYear - 2, currentFiscalYear - 1, currentFiscalYear, currentFiscalYear + 1];
                wrap.innerHTML = years.map(fiscalYearStart => {
                    const activeCls = omsaetningSelectedFiscalYears.has(fiscalYearStart) ? ' active' : '';
                    const label = String(fiscalYearStart) + '/' + String(fiscalYearStart + 1).slice(-2);
                    return '<button type="button" class="omsaetning-year-btn' + activeCls + '" onclick="toggleOmsaetningFiscalYear(' + fiscalYearStart + ')">' +
                        (fiscalYearStart === currentFiscalYear ? 'Nu ' : '') + escapeHtmlFE(label) +
                        '</button>';
                }).join('');
            }

            function toggleOmsaetningFiscalYear(year) {
                const yearValue = Number(year);
                if (!Number.isFinite(yearValue)) return;
                if (omsaetningSelectedFiscalYears.has(yearValue) && omsaetningSelectedFiscalYears.size > 1) {
                    omsaetningSelectedFiscalYears.delete(yearValue);
                } else {
                    omsaetningSelectedFiscalYears.add(yearValue);
                }
                applySelectedFiscalYearsToInputs();
                renderOmsaetningYearChips(getCurrentFiscalYearStart());
                loadOmsaetningSummary();
            }

            function filterOmsaetningAccounts() {
                const qEl = document.getElementById('omsaetningAccountSearch');
                const q = String((qEl && qEl.value) || '').trim().toLowerCase();
                const rows = document.querySelectorAll('#omsaetningAccountsList .omsaetning-account-item');
                for (const row of rows) {
                    const text = String(row.getAttribute('data-search') || '').toLowerCase();
                    row.style.display = (!q || text.includes(q)) ? '' : 'none';
                }
            }

            function setAllOmsaetningAccounts(checked) {
                const list = document.getElementById('omsaetningAccountsList');
                if (!list) return;
                const boxes = list.querySelectorAll('input[type="checkbox"][data-accno]');
                for (const box of boxes) {
                    box.checked = !!checked;
                    const value = String(box.getAttribute('data-accno') || '').trim();
                    if (!value) continue;
                    if (checked) omsaetningSelectedAccounts.add(value);
                    else omsaetningSelectedAccounts.delete(value);
                }
                renderOmsaetningAccountsSummary();
                scheduleOmsaetningAutoReload();
            }

            function renderOmsaetningAccountsList() {
                const list = document.getElementById('omsaetningAccountsList');
                if (!list) return;
                if (!Array.isArray(omsaetningAccounts) || omsaetningAccounts.length === 0) {
                    list.innerHTML = '<div class="omsaetning-account-item"><span>Ingen konti fundet</span></div>';
                    return;
                }

                list.innerHTML = omsaetningAccounts.map(acc => {
                    const value = String(acc.acNo || '').trim();
                    const checked = omsaetningSelectedAccounts.has(value) ? ' checked' : '';
                    const search = (value + ' ' + String(acc.name || '')).replace(/"/g, '&quot;');
                    return '<label class="omsaetning-account-item" data-search="' + search + '">' +
                        '<input type="checkbox" data-accno="' + escapeHtmlFE(value) + '"' + checked + ' onchange="toggleOmsaetningAccount(this)" />' +
                        '<span>' + escapeHtmlFE(value + ' - ' + String(acc.name || '')) + '</span>' +
                        '</label>';
                }).join('');
                renderOmsaetningAccountsSummary();
                applyOmsaetningAccountsPanelState();
            }

            function toggleOmsaetningAccount(inputEl) {
                if (!inputEl) return;
                const value = String(inputEl.getAttribute('data-accno') || '').trim();
                if (!value) return;
                if (inputEl.checked) omsaetningSelectedAccounts.add(value);
                else omsaetningSelectedAccounts.delete(value);
                renderOmsaetningAccountsSummary();
                scheduleOmsaetningAutoReload();
            }

            function renderOmsaetningSelectedCustomers() {
                const wrap = document.getElementById('omsaetningSelectedCustomers');
                if (!wrap) return;

                const entries = Array.from(omsaetningSelectedCustomers.entries());
                if (entries.length === 0) {
                    wrap.innerHTML = '';
                    return;
                }

                wrap.innerHTML = entries.map(([custNo, name]) => (
                    '<span class="omsaetning-selected-chip">' +
                    escapeHtmlFE(String(custNo) + ' - ' + String(name || '')) +
                    '<button type="button" title="Fjern kunde" onclick="removeOmsaetningCustomer(' + Number(custNo) + ')">×</button>' +
                    '</span>'
                )).join('');
            }

            function removeOmsaetningCustomer(custNo) {
                const key = String(custNo || '').trim();
                if (!key) return;
                omsaetningSelectedCustomers.delete(key);
                renderOmsaetningSelectedCustomers();
                onOmsaetningSelectedCustomersChanged();
                renderOmsaetningCustomerResults();
            }

            function clearOmsaetningCustomerSearchUi() {
                const qEl = document.getElementById('omsaetningCustomerSearch');
                if (qEl) {
                    qEl.value = '';
                    qEl.blur();
                }
                omsaetningCustomerResults = [];
                renderOmsaetningCustomerResults();
            }

            function renderOmsaetningCustomerMode() {
                const modeEl = document.getElementById('omsaetningCustomerMode');
                if (!modeEl) return;

                const selectedCount = omsaetningSelectedCustomers.size;
                if (selectedCount === 0) {
                    modeEl.textContent = 'Ingen kunde valgt: viser normal visning for valgte år og konti.';
                    return;
                }

                if (selectedCount === 1) {
                    modeEl.textContent = '1 kunde valgt: rapporten filtreres på kunden.';
                    return;
                }

                modeEl.textContent = String(selectedCount) + ' kunder valgt: søjlediagram sammenligner kunder måned for måned.';
            }

            function toggleOmsaetningCustomer(custNo, name) {
                const key = String(custNo || '').trim();
                if (!key) return;
                const wasSelected = omsaetningSelectedCustomers.has(key);
                if (omsaetningSelectedCustomers.has(key)) {
                    omsaetningSelectedCustomers.delete(key);
                } else {
                    omsaetningSelectedCustomers.set(key, String(name || '').trim());
                }
                renderOmsaetningSelectedCustomers();
                onOmsaetningSelectedCustomersChanged();
                if (!wasSelected) clearOmsaetningCustomerSearchUi();
                renderOmsaetningCustomerResults();
            }

            function toggleOmsaetningCustomerByButton(buttonEl) {
                if (!buttonEl) return;
                const custNo = String(buttonEl.getAttribute('data-custno') || '').trim();
                const custName = String(buttonEl.getAttribute('data-custname') || '').trim();
                toggleOmsaetningCustomer(custNo, custName);
            }

            function renderOmsaetningCustomerResults() {
                const wrap = document.getElementById('omsaetningCustomerResults');
                if (!wrap) return;
                const qEl = document.getElementById('omsaetningCustomerSearch');
                const q = String((qEl && qEl.value) || '').trim();
                const selectedCount = omsaetningSelectedCustomers.size;

                if (q.length < 2) {
                    if (selectedCount === 0) {
                        wrap.innerHTML = '<div class="omsaetning-customer-empty">Ingen kunde valgt: viser normal visning for valgte år og konti.</div>';
                    } else {
                        wrap.innerHTML = '<div class="omsaetning-customer-empty">Skriv mindst 2 tegn for at tilføje flere kunder.</div>';
                    }
                    return;
                }

                if (!Array.isArray(omsaetningCustomerResults) || omsaetningCustomerResults.length === 0) {
                    wrap.innerHTML = '<div class="omsaetning-customer-empty">Ingen kunder matcher søgningen.</div>';
                    return;
                }

                wrap.innerHTML = omsaetningCustomerResults.map(row => {
                    const custNo = String(row.custNo || '').trim();
                    const custName = String(row.name || '').trim();
                    const selected = omsaetningSelectedCustomers.has(custNo);
                    return '<div class="omsaetning-customer-item">' +
                        '<div class="meta"><strong>' + escapeHtmlFE(custName || '(uden navn)') + '</strong><span>' + escapeHtmlFE(custNo) + '</span></div>' +
                        '<button type="button" class="pick' + (selected ? ' remove' : '') + '" data-custno="' + escapeHtmlFE(custNo) + '" data-custname="' + escapeHtmlFE(custName) + '" onclick="toggleOmsaetningCustomerByButton(this)">' + (selected ? 'Valgt' : 'Vælg') + '</button>' +
                        '</div>';
                }).join('');
            }

            function scheduleOmsaetningCustomerSearch() {
                if (omsaetningCustomerSearchTimer) {
                    clearTimeout(omsaetningCustomerSearchTimer);
                }
                omsaetningCustomerSearchTimer = setTimeout(() => {
                    searchOmsaetningCustomers();
                }, 220);
            }

            async function searchOmsaetningCustomers() {
                const qEl = document.getElementById('omsaetningCustomerSearch');
                const q = String((qEl && qEl.value) || '').trim();
                const token = ++omsaetningCustomerSearchToken;

                if (q.length < 2) {
                    omsaetningCustomerResults = [];
                    renderOmsaetningCustomerResults();
                    return;
                }

                try {
                    const payload = await searchOmsaetningCustomersCached(q);
                    if (token !== omsaetningCustomerSearchToken) return;
                    const rows = Array.isArray(payload.customers) ? payload.customers : [];
                    omsaetningCustomerResults = rows;
                    renderOmsaetningCustomerResults();
                } catch (err) {
                    if (token !== omsaetningCustomerSearchToken) return;
                    omsaetningCustomerResults = [];
                    const wrap = document.getElementById('omsaetningCustomerResults');
                    if (wrap) {
                        wrap.innerHTML = '<div class="omsaetning-customer-empty">Fejl ved kundesøgning.</div>';
                    }
                }
            }

            function getOmsaetningStatusClass(valueMio, warnThreshold, goodThreshold) {
                const n = Number(valueMio || 0);
                if (n >= goodThreshold) return 'good';
                if (n >= warnThreshold) return 'mid';
                return 'low';
            }

            function getOmsaetningThresholdInputs() {
                const warnThresholdInput = document.getElementById('omsaetningWarnThreshold');
                const goodThresholdInput = document.getElementById('omsaetningGoodThreshold');

                const warnRaw = Number((warnThresholdInput && warnThresholdInput.value) || OMSAETNING_DEFAULT_WARN_THRESHOLD);
                const warnThreshold = Number.isFinite(warnRaw) ? Math.max(0, warnRaw) : OMSAETNING_DEFAULT_WARN_THRESHOLD;

                const goodRaw = Number((goodThresholdInput && goodThresholdInput.value) || OMSAETNING_DEFAULT_GOOD_THRESHOLD);
                const goodThreshold = Number.isFinite(goodRaw)
                    ? Math.max(warnThreshold, goodRaw)
                    : Math.max(warnThreshold, OMSAETNING_DEFAULT_GOOD_THRESHOLD);

                return { warnThreshold, goodThreshold };
            }

            function applyOmsaetningThresholdInputs(warnThreshold, goodThreshold) {
                const warnThresholdInput = document.getElementById('omsaetningWarnThreshold');
                const goodThresholdInput = document.getElementById('omsaetningGoodThreshold');
                if (!warnThresholdInput || !goodThresholdInput) return;

                const warnRaw = Number(warnThreshold);
                const normalizedWarn = Number.isFinite(warnRaw) ? Math.max(0, warnRaw) : OMSAETNING_DEFAULT_WARN_THRESHOLD;

                const goodRaw = Number(goodThreshold);
                const normalizedGood = Number.isFinite(goodRaw)
                    ? Math.max(normalizedWarn, goodRaw)
                    : Math.max(normalizedWarn, OMSAETNING_DEFAULT_GOOD_THRESHOLD);

                warnThresholdInput.value = normalizedWarn.toFixed(1);
                goodThresholdInput.value = normalizedGood.toFixed(1);
            }

            async function loadOmsaetningThresholdForCustomer(custNo) {
                const key = String(custNo || '').trim();
                if (!/^\d{1,20}$/.test(key)) return;
                const token = ++omsaetningThresholdLoadToken;

                try {
                    const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(key));
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    if (token !== omsaetningThresholdLoadToken) return;
                    omsaetningThresholdsByCustomer.set(key, {
                        warnThreshold: Number(payload.warnThreshold || OMSAETNING_DEFAULT_WARN_THRESHOLD),
                        goodThreshold: Number(payload.goodThreshold || OMSAETNING_DEFAULT_GOOD_THRESHOLD)
                    });
                    applyOmsaetningThresholdInputs(payload.warnThreshold, payload.goodThreshold);
                    renderOmsaetningCustomerThresholds();
                } catch (err) {
                    if (token !== omsaetningThresholdLoadToken) return;
                    console.warn('loadOmsaetningThresholdForCustomer failed:', err && err.message ? err.message : err);
                }
            }

            function renderOmsaetningCustomerThresholds() {
                const wrap = document.getElementById('omsaetningCustomerThresholds');
                if (!wrap) return;

                const selectedEntries = Array.from(omsaetningSelectedCustomers.entries());
                if (selectedEntries.length === 0) {
                    wrap.innerHTML = '';
                    return;
                }

                const rows = selectedEntries.map(([custNo, custName]) => {
                    const key = String(custNo || '').trim();
                    const threshold = omsaetningThresholdsByCustomer.get(key);
                    if (!threshold) {
                        return '<div class="omsaetning-customer-threshold-row"><span class="cust">' +
                            escapeHtmlFE(String(custName || key)) + ' (' + escapeHtmlFE(key) + ')' +
                            '</span><span class="thr">tærskler: indlæser...</span></div>';
                    }
                    return '<div class="omsaetning-customer-threshold-row"><span class="cust">' +
                        escapeHtmlFE(String(custName || key)) + ' (' + escapeHtmlFE(key) + ')' +
                        '</span><span class="thr">tærskler: ' + escapeHtmlFE(Number(threshold.warnThreshold).toFixed(1)) +
                        ' / ' + escapeHtmlFE(Number(threshold.goodThreshold).toFixed(1)) + '</span></div>';
                });

                wrap.innerHTML = rows.join('');
            }

            async function refreshOmsaetningThresholdsForSelectedCustomers(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys())
                    .map(v => String(v || '').trim())
                    .filter(v => /^\d{1,20}$/.test(v));

                if (selectedCustomers.length === 0) {
                    omsaetningThresholdLoadToken += 1;
                    omsaetningThresholdsByCustomer = new Map();
                    renderOmsaetningCustomerThresholds();
                    return;
                }

                const token = ++omsaetningThresholdLoadToken;
                renderOmsaetningCustomerThresholds();

                const results = await Promise.all(selectedCustomers.map(async custNo => {
                    try {
                        const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(custNo));
                        if (!response.ok) throw new Error('HTTP ' + response.status);
                        const payload = await response.json();
                        return {
                            custNo,
                            warnThreshold: Number(payload.warnThreshold || OMSAETNING_DEFAULT_WARN_THRESHOLD),
                            goodThreshold: Number(payload.goodThreshold || OMSAETNING_DEFAULT_GOOD_THRESHOLD)
                        };
                    } catch (err) {
                        console.warn('refresh threshold failed for', custNo, err && err.message ? err.message : err);
                        return {
                            custNo,
                            warnThreshold: OMSAETNING_DEFAULT_WARN_THRESHOLD,
                            goodThreshold: OMSAETNING_DEFAULT_GOOD_THRESHOLD
                        };
                    }
                }));

                if (token !== omsaetningThresholdLoadToken) return;

                omsaetningThresholdsByCustomer = new Map(results.map(item => [item.custNo, {
                    warnThreshold: item.warnThreshold,
                    goodThreshold: item.goodThreshold
                }]));

                if (safeOptions.applySingleSelectionToInputs === true && selectedCustomers.length === 1) {
                    const single = omsaetningThresholdsByCustomer.get(selectedCustomers[0]);
                    if (single) {
                        applyOmsaetningThresholdInputs(single.warnThreshold, single.goodThreshold);
                    }
                }

                renderOmsaetningCustomerThresholds();
            }

            async function persistOmsaetningThresholdsForCustomers(customerNos, warnThreshold, goodThreshold) {
                const customerKeys = Array.from(new Set((Array.isArray(customerNos) ? customerNos : [])
                    .map(v => String(v || '').trim())
                    .filter(v => /^\d{1,20}$/.test(v))));
                if (customerKeys.length === 0) return;

                await Promise.all(customerKeys.map(async custNo => {
                    const response = await fetch('/omsaetning/customer-threshold/' + encodeURIComponent(custNo), {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ warnThreshold, goodThreshold })
                    });
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                }));
            }

            function showOmsaetningThresholdPersistDialog(customerKeys, warnThreshold, goodThreshold) {
                return new Promise(resolve => {
                    const listHtml = customerKeys.map(custNo => {
                        const name = String(omsaetningSelectedCustomers.get(custNo) || '').trim();
                        const label = custNo + (name ? (' - ' + name) : '');
                        return '<button type="button" class="omsaetning-persist-customer-option" data-cust="' + escapeHtmlFE(custNo) + '">' + escapeHtmlFE(label) + '</button>';
                    }).join('');

                    const defaultCustomer = String(customerKeys[0] || '');
                    let selectedCustomer = defaultCustomer;
                    const overlay = document.createElement('div');
                    overlay.className = 'omsaetning-persist-overlay';
                    overlay.innerHTML =
                        '<div class="omsaetning-persist-dialog" role="dialog" aria-modal="true" aria-label="Gem tærskler">' +
                            '<div class="omsaetning-persist-head"><h4>Gem tærskler for valgte kunder</h4></div>' +
                            '<div class="omsaetning-persist-body">' +
                                '<div>Flere kunder er valgt. Vælg hvor de nye tærskler skal gemmes.</div>' +
                                '<div class="omsaetning-persist-customer-list">' + listHtml + '</div>' +
                                '<div class="omsaetning-persist-thr">Nye tærskler: ' + escapeHtmlFE(Number(warnThreshold).toFixed(1)) + ' / ' + escapeHtmlFE(Number(goodThreshold).toFixed(1)) + '</div>' +
                                '<div class="omsaetning-persist-pick">Valgt kunde: <span id="omsaetningPersistPicked" class="omsaetning-persist-picked">' + escapeHtmlFE(defaultCustomer) + '</span></div>' +
                                '</div>' +
                                '<div class="omsaetning-persist-actions">' +
                                    '<button type="button" class="primary" data-action="all">Gem alle</button>' +
                                    '<button type="button" class="ghost" data-action="single">Gem valgt kunde</button>' +
                                    '<button type="button" class="danger" data-action="none">Gem ikke</button>' +
                                '</div>' +
                            '</div>' +
                        '</div>';

                    const closeWith = value => {
                        overlay.remove();
                        resolve(value);
                    };

                    overlay.addEventListener('click', ev => {
                        if (ev.target === overlay) closeWith('NONE');
                    });

                    overlay.querySelector('[data-action="all"]').addEventListener('click', () => closeWith('ALL'));
                    overlay.querySelector('[data-action="none"]').addEventListener('click', () => closeWith('NONE'));
                    overlay.querySelector('[data-action="single"]').addEventListener('click', () => {
                        closeWith(String(selectedCustomer || '').trim());
                    });

                    const pickedEl = overlay.querySelector('#omsaetningPersistPicked');
                    const customerButtons = Array.from(overlay.querySelectorAll('.omsaetning-persist-customer-option'));

                    const renderPicked = () => {
                        for (const btn of customerButtons) {
                            const btnCust = String(btn.getAttribute('data-cust') || '');
                            btn.classList.toggle('active', btnCust === selectedCustomer);
                        }
                        if (pickedEl) pickedEl.textContent = selectedCustomer || '-';
                    };

                    for (const btn of customerButtons) {
                        btn.addEventListener('click', () => {
                            selectedCustomer = String(btn.getAttribute('data-cust') || '').trim();
                            renderPicked();
                        });
                    }

                    overlay.addEventListener('keydown', ev => {
                        if (ev.key === 'Enter') {
                            ev.preventDefault();
                            closeWith(String(selectedCustomer || '').trim());
                        }
                        if (ev.key === 'Escape') {
                            ev.preventDefault();
                            closeWith('NONE');
                        }
                    });

                    document.body.appendChild(overlay);
                    renderPicked();
                    if (customerButtons.length > 0) {
                        customerButtons[0].focus();
                    }
                });
            }

            async function resolveOmsaetningThresholdPersistTargets(selectedCustomers, warnThreshold, goodThreshold, options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const customerKeys = Array.from(new Set((Array.isArray(selectedCustomers) ? selectedCustomers : [])
                    .map(v => String(v || '').trim())
                    .filter(v => /^\d{1,20}$/.test(v))));

                if (customerKeys.length === 0) return [];
                if (safeOptions.silentValidation === true) return [];
                if (safeOptions.persistThresholdsOnUpdate !== true) return [];

                if (customerKeys.length === 1) {
                    return customerKeys;
                }

                const answerRaw = await showOmsaetningThresholdPersistDialog(customerKeys, warnThreshold, goodThreshold);
                const answer = String(answerRaw || '').trim().toUpperCase();

                if (!answer || answer === 'NONE' || answer === 'NO' || answer === 'N') {
                    return [];
                }
                if (answer === 'ALL' || answer === 'A') {
                    return customerKeys;
                }

                const exact = customerKeys.find(c => c.toUpperCase() === answer);
                if (exact) return [exact];

                alert('Ugyldigt valg for tærskel-gemning. Ingen tærskler blev gemt.');
                return [];
            }

            function onOmsaetningSelectedCustomersChanged() {
                const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys()).filter(Boolean);
                renderOmsaetningCustomerMode();
                if (selectedCustomers.length === 0) {
                    omsaetningThresholdLoadToken += 1;
                    omsaetningThresholdsByCustomer = new Map();
                    applyOmsaetningThresholdInputs(OMSAETNING_DEFAULT_WARN_THRESHOLD, OMSAETNING_DEFAULT_GOOD_THRESHOLD);
                    renderOmsaetningCustomerThresholds();
                    scheduleOmsaetningAutoReload();
                    return;
                }

                refreshOmsaetningThresholdsForSelectedCustomers({ applySingleSelectionToInputs: true });
                scheduleOmsaetningAutoReload();
            }

            function scheduleOmsaetningAutoReload() {
                if (omsaetningAutoReloadTimer) {
                    clearTimeout(omsaetningAutoReloadTimer);
                }
                omsaetningAutoReloadTimer = setTimeout(() => {
                    omsaetningAutoReloadTimer = null;
                    loadOmsaetningSummary({ silentValidation: true });
                }, OMSAETNING_AUTO_RELOAD_DELAY_MS);
            }

            function getOmsaetningStatusLabel(statusClass) {
                if (statusClass === 'good') return 'Over mål';
                if (statusClass === 'mid') return 'Nær mål';
                return 'Under mål';
            }

            function getOmsaetningColor(index) {
                const palette = ['#1565c0', '#00acc1', '#00897b', '#7b1fa2', '#ef6c00', '#5e35b1', '#43a047', '#c62828'];
                return palette[index % palette.length];
            }

            function renderOmsaetningCharts(rows, forcedMonthKeys) {
                const chartsWrap = document.getElementById('omsaetningChartsWrap');
                const stackedSvg = document.getElementById('omsaetningStackedChart');
                const trendSvg = document.getElementById('omsaetningTrendChart');
                const legend = document.getElementById('omsaetningLegend');
                const stackedTitle = document.getElementById('omsaetningStackedTitle');
                if (!chartsWrap || !stackedSvg || !trendSvg || !legend) return;

                const safeRows = Array.isArray(rows) ? rows : [];
                const safeForcedMonths = Array.isArray(forcedMonthKeys) ? forcedMonthKeys.map(v => String(v || '').trim()).filter(Boolean) : [];

                if (safeRows.length === 0 && safeForcedMonths.length === 0) {
                    chartsWrap.style.display = 'none';
                    stackedSvg.innerHTML = '';
                    trendSvg.innerHTML = '';
                    legend.innerHTML = '';
                    if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';
                    return;
                }

                const monthMap = new Map();
                const accountOrder = [];
                const seenAccounts = new Set();
                for (const row of safeRows) {
                    const monthKey = normalizeOmsaetningMonthKey(row.date);
                    if (!monthMap.has(monthKey)) monthMap.set(monthKey, new Map());
                    const accountKey = String(row.acNo || '');
                    if (!seenAccounts.has(accountKey)) {
                        seenAccounts.add(accountKey);
                        accountOrder.push({ acNo: accountKey, name: String(row.name || '') });
                    }
                    const monthAcc = monthMap.get(monthKey);
                    monthAcc.set(accountKey, (monthAcc.get(accountKey) || 0) + Number(row.revenueMio || 0));
                }

                const monthKeys = (safeForcedMonths.length > 0 ? safeForcedMonths : Array.from(monthMap.keys())).sort((a, b) => String(a).localeCompare(String(b)));
                if (monthKeys.length === 0) {
                    chartsWrap.style.display = 'none';
                    stackedSvg.innerHTML = '';
                    trendSvg.innerHTML = '';
                    legend.innerHTML = '';
                    if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';
                    return;
                }
                const monthlyTotals = monthKeys.map(key => {
                    const m = monthMap.get(key) || new Map();
                    let t = 0;
                    for (const value of m.values()) t += Number(value || 0);
                    return t;
                });

                function buildScale(values, fallbackAbs) {
                    const safeValues = Array.isArray(values) ? values : [];
                    let min = 0;
                    let max = 0;
                    for (const raw of safeValues) {
                        const value = Number(raw);
                        if (!Number.isFinite(value)) continue;
                        if (value < min) min = value;
                        if (value > max) max = value;
                    }
                    if (min === 0 && max === 0) {
                        max = Number.isFinite(fallbackAbs) && fallbackAbs > 0 ? fallbackAbs : 0.1;
                    }
                    if (Math.abs(max - min) < 0.000001) {
                        if (max > 0) min = 0;
                        else if (min < 0) max = 0;
                        else max = 0.1;
                    }
                    return { min, max, span: max - min };
                }

                const compareCustomers = Array.from(omsaetningSelectedCustomers.entries())
                    .map(([custNo, name]) => ({ custNo: String(custNo || '').trim(), name: String(name || '').trim() }))
                    .filter(c => c.custNo);
                const showCustomerComparison = compareCustomers.length > 1;

                const leftPad = 48;
                const topPad = 16;
                const bottomPad = 42;
                const chartHeight = 190;
                const innerHeight = chartHeight - topPad - bottomPad;
                const barWidth = 34;
                const barGap = 18;
                const innerWidth = Math.max(560, monthKeys.length * (barWidth + barGap));
                const viewWidth = leftPad + innerWidth + 20;
                const viewHeight = chartHeight;

                function toY(value, scale) {
                    const safeScale = scale || { min: 0, max: 1, span: 1 };
                    const safeValue = Number.isFinite(Number(value)) ? Number(value) : 0;
                    const ratio = (safeScale.max - safeValue) / (safeScale.span || 1);
                    return topPad + (ratio * innerHeight);
                }

                function appendGrid(svgHtml, width, scale) {
                    let out = svgHtml;
                    const ticks = 4;
                    for (let i = 0; i <= ticks; i++) {
                        const ratio = i / ticks;
                        const tickValue = scale.max - (ratio * scale.span);
                        const y = topPad + (innerHeight * ratio);
                        out += '<line x1="' + leftPad + '" y1="' + y + '" x2="' + (leftPad + width) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
                        out += '<text x="' + (leftPad - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="#5f7892">' + escapeHtmlFE(formatMio(tickValue)) + '</text>';
                    }
                    const zeroY = toY(0, scale);
                    out += '<line x1="' + leftPad + '" y1="' + zeroY + '" x2="' + (leftPad + width) + '" y2="' + zeroY + '" stroke="#8fa8c2" stroke-width="1.2" />';
                    return out;
                }

                let stackedSvgHtml = '<g>';
                if (showCustomerComparison) {
                    if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (kunde-sammenligning)';

                    const byMonthCustomer = new Map();
                    for (const row of safeRows) {
                        const monthKey = normalizeOmsaetningMonthKey(row.date);
                        const custNo = String(row.custNo || '').trim();
                        if (!monthKey || !custNo) continue;
                        if (!byMonthCustomer.has(monthKey)) byMonthCustomer.set(monthKey, new Map());
                        const map = byMonthCustomer.get(monthKey);
                        map.set(custNo, (map.get(custNo) || 0) + Number(row.revenueMio || 0));
                    }

                    const groupedValues = [];
                    for (const monthKey of monthKeys) {
                        const values = byMonthCustomer.get(monthKey) || new Map();
                        for (const customer of compareCustomers) {
                            const v = Number(values.get(customer.custNo) || 0);
                            groupedValues.push(v);
                        }
                    }
                    const groupedScale = buildScale(groupedValues, 0.1);
                    const groupedZeroY = toY(0, groupedScale);

                    const groupedBarWidth = 14;
                    const groupedBarGap = 5;
                    const monthGroupGap = 18;
                    const perMonthGroupWidth = (compareCustomers.length * groupedBarWidth) + ((compareCustomers.length - 1) * groupedBarGap);
                    const groupedInnerWidth = Math.max(560, monthKeys.length * (perMonthGroupWidth + monthGroupGap));

                    stackedSvgHtml = appendGrid(stackedSvgHtml, groupedInnerWidth, groupedScale);

                    monthKeys.forEach((monthKey, monthIndex) => {
                        const monthX = leftPad + monthIndex * (perMonthGroupWidth + monthGroupGap);
                        const values = byMonthCustomer.get(monthKey) || new Map();
                        compareCustomers.forEach((customer, customerIndex) => {
                            const value = Number(values.get(customer.custNo) || 0);
                            const yValue = toY(value, groupedScale);
                            const y = value >= 0 ? yValue : groupedZeroY;
                            const h = Math.max(1, Math.abs(groupedZeroY - yValue));
                            const x = monthX + customerIndex * (groupedBarWidth + groupedBarGap);
                            const titleText = formatMonthDa(monthKey) + ' - ' + String(customer.name || customer.custNo) + ' (' + customer.custNo + '): ' + formatMio(value) + ' Mio DKK (' + formatDkkFromMio(value) + ' DKK)';
                            stackedSvgHtml += '<rect x="' + x + '" y="' + y + '" width="' + groupedBarWidth + '" height="' + h + '" fill="' + getOmsaetningColor(customerIndex) + '" rx="2"><title>' + escapeHtmlFE(titleText) + '</title></rect>';
                        });

                        const labelX = monthX + (perMonthGroupWidth / 2);
                        stackedSvgHtml += '<text x="' + labelX + '" y="' + (topPad + innerHeight + 14) + '" text-anchor="middle" font-size="10" fill="#47617c">' + escapeHtmlFE(formatMonthDa(monthKey)) + '</text>';
                    });

                    stackedSvgHtml += '</g>';
                    stackedSvg.setAttribute('viewBox', '0 0 ' + (leftPad + groupedInnerWidth + 20) + ' ' + viewHeight);
                    stackedSvg.innerHTML = stackedSvgHtml;

                    legend.innerHTML = compareCustomers.map((customer, idx) =>
                        '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:' + getOmsaetningColor(idx) + ';"></span>' +
                        escapeHtmlFE(String(customer.name || customer.custNo)) + ' (' + escapeHtmlFE(customer.custNo) + ')</span>'
                    ).join('');
                } else {
                    if (stackedTitle) stackedTitle.textContent = 'Omsætning pr. måned (stacked pr. konto)';

                    const stackedMonthTotals = monthKeys.map(monthKey => {
                        const values = monthMap.get(monthKey) || new Map();
                        let pos = 0;
                        let neg = 0;
                        for (const value of values.values()) {
                            const n = Number(value || 0);
                            if (n >= 0) pos += n;
                            else neg += n;
                        }
                        return { pos, neg };
                    });
                    const stackedScale = buildScale(
                        stackedMonthTotals.flatMap(t => [t.pos, t.neg]),
                        0.1
                    );
                    const stackedZeroY = toY(0, stackedScale);

                    stackedSvgHtml = appendGrid(stackedSvgHtml, innerWidth, stackedScale);

                    monthKeys.forEach((monthKey, monthIndex) => {
                        const x = leftPad + monthIndex * (barWidth + barGap);
                        const values = monthMap.get(monthKey) || new Map();
                        let positiveStack = 0;
                        let negativeStack = 0;
                        accountOrder.forEach((acc, accIndex) => {
                            const value = Number(values.get(acc.acNo) || 0);
                            if (value === 0) return;

                            let startValue;
                            let endValue;
                            if (value > 0) {
                                startValue = positiveStack;
                                endValue = positiveStack + value;
                                positiveStack = endValue;
                            } else {
                                startValue = negativeStack;
                                endValue = negativeStack + value;
                                negativeStack = endValue;
                            }

                            const yStart = toY(startValue, stackedScale);
                            const yEnd = toY(endValue, stackedScale);
                            const y = Math.min(yStart, yEnd);
                            const h = Math.max(1, Math.abs(yEnd - yStart));
                            const titleText = formatMonthDa(monthKey) + ' - ' + String(acc.acNo) + ' ' + String(acc.name || '') + ': ' + formatMio(value) + ' Mio DKK (' + formatDkkFromMio(value) + ' DKK)';
                            stackedSvgHtml += '<rect x="' + x + '" y="' + y + '" width="' + barWidth + '" height="' + h + '" fill="' + getOmsaetningColor(accIndex) + '" rx="2"><title>' + escapeHtmlFE(titleText) + '</title></rect>';
                        });

                        if (Math.abs(positiveStack) < 0.000001 && Math.abs(negativeStack) < 0.000001) {
                            stackedSvgHtml += '<line x1="' + x + '" y1="' + stackedZeroY + '" x2="' + (x + barWidth) + '" y2="' + stackedZeroY + '" stroke="#cddced" stroke-width="1" />';
                        }
                        stackedSvgHtml += '<text x="' + (x + barWidth / 2) + '" y="' + (topPad + innerHeight + 14) + '" text-anchor="middle" font-size="10" fill="#47617c">' + escapeHtmlFE(formatMonthDa(monthKey)) + '</text>';
                    });

                    stackedSvgHtml += '</g>';
                    stackedSvg.setAttribute('viewBox', '0 0 ' + viewWidth + ' ' + viewHeight);
                    stackedSvg.innerHTML = stackedSvgHtml;

                    legend.innerHTML = accountOrder.map((acc, idx) =>
                        '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:' + getOmsaetningColor(idx) + ';"></span>' +
                        escapeHtmlFE(String(acc.acNo)) + ' ' + escapeHtmlFE(acc.name || '') + '</span>'
                    ).join('');
                }

                const trendLeftPad = 42;
                const trendTopPad = 16;
                const trendBottomPad = 28;
                const trendHeight = 190;
                const trendInnerHeight = trendHeight - trendTopPad - trendBottomPad;
                const trendInnerWidth = Math.max(560, monthKeys.length * 54);
                const trendViewWidth = trendLeftPad + trendInnerWidth + 16;
                const trendViewHeight = trendHeight;
                const trendScale = buildScale(monthlyTotals, 0.1);

                function toTrendY(value) {
                    const safeValue = Number.isFinite(Number(value)) ? Number(value) : 0;
                    const ratio = (trendScale.max - safeValue) / (trendScale.span || 1);
                    return trendTopPad + (ratio * trendInnerHeight);
                }

                let trendSvgHtml = '<g>';
                for (let i = 0; i <= 4; i++) {
                    const ratio = i / 4;
                    const y = trendTopPad + (trendInnerHeight * ratio);
                    const tickValue = trendScale.max - (ratio * trendScale.span);
                    trendSvgHtml += '<line x1="' + trendLeftPad + '" y1="' + y + '" x2="' + (trendLeftPad + trendInnerWidth) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
                    trendSvgHtml += '<text x="' + (trendLeftPad - 6) + '" y="' + (y + 4) + '" text-anchor="end" font-size="10" fill="#5f7892">' + escapeHtmlFE(formatMio(tickValue)) + '</text>';
                    }

                const trendZeroY = toTrendY(0);
                trendSvgHtml += '<line x1="' + trendLeftPad + '" y1="' + trendZeroY + '" x2="' + (trendLeftPad + trendInnerWidth) + '" y2="' + trendZeroY + '" stroke="#8fa8c2" stroke-width="1.2" />';

                const points = monthKeys.map((monthKey, idx) => {
                    const x = trendLeftPad + (trendInnerWidth * (monthKeys.length === 1 ? 0.5 : (idx / (monthKeys.length - 1))));
                    const y = toTrendY(monthlyTotals[idx]);
                    return { x, y, monthKey, total: monthlyTotals[idx] };
                });
                if (points.length === 0) {
                    chartsWrap.style.display = 'none';
                    trendSvg.innerHTML = '';
                    return;
                }
                const linePath = points.map((p, idx) => (idx === 0 ? 'M' : 'L') + p.x + ' ' + p.y).join(' ');
                const areaPath = linePath + ' L ' + points[points.length - 1].x + ' ' + trendZeroY + ' L ' + points[0].x + ' ' + trendZeroY + ' Z';
                trendSvgHtml += '<path d="' + areaPath + '" fill="rgba(21,101,192,0.12)" />';
                trendSvgHtml += '<path d="' + linePath + '" fill="none" stroke="#1565c0" stroke-width="3" />';
                points.forEach(p => {
                    const pointTitle = formatMonthDa(p.monthKey) + ': ' + formatMio(p.total) + ' Mio DKK (' + formatDkkFromMio(p.total) + ' DKK)';
                    trendSvgHtml += '<circle cx="' + p.x + '" cy="' + p.y + '" r="3.5" fill="#0f3560"><title>' + escapeHtmlFE(pointTitle) + '</title></circle>';
                });
                trendSvgHtml += '</g>';

                trendSvg.setAttribute('viewBox', '0 0 ' + trendViewWidth + ' ' + trendViewHeight);
                trendSvg.innerHTML = trendSvgHtml;
                chartsWrap.style.display = 'grid';
            }

            function normalizeWeekKeyInput(value) {
                return String(value || '').replace(/[^0-9]/g, '').slice(0, 6);
            }

            function parseWeekKeyMeta(value) {
                const raw = normalizeWeekKeyInput(value);
                if (!/^[0-9]{6}$/.test(raw)) return null;
                const year = Number(raw.slice(0, 4));
                const week = Number(raw.slice(4, 6));
                if (!Number.isFinite(year) || !Number.isFinite(week) || week < 1 || week > 53) return null;
                return { raw, year, week };
            }

            function getIsoWeekMeta(dateValue) {
                const d = new Date(Date.UTC(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()));
                d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
                const isoYear = d.getUTCFullYear();
                const yearStart = new Date(Date.UTC(isoYear, 0, 1));
                const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
                return {
                    isoYear,
                    week,
                    weekKey: String(isoYear) + String(week).padStart(2, '0')
                };
            }

            function getIsoWeekStartDate(isoYear, week) {
                const jan4 = new Date(Date.UTC(isoYear, 0, 4));
                const jan4Day = jan4.getUTCDay() || 7;
                const mondayWeek1 = new Date(jan4);
                mondayWeek1.setUTCDate(jan4.getUTCDate() - jan4Day + 1);
                const weekStart = new Date(mondayWeek1);
                weekStart.setUTCDate(mondayWeek1.getUTCDate() + ((week - 1) * 7));
                return weekStart;
            }

            function formatWeekLabel(weekKey) {
                const meta = parseWeekKeyMeta(weekKey);
                if (!meta) return String(weekKey || '-');
                return String(meta.year) + '-W' + String(meta.week).padStart(2, '0');
            }

            function formatDkkDa(value) {
                return Number(value || 0).toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            }

            function formatPctDa(value) {
                if (value === null || value === undefined || !Number.isFinite(Number(value))) return '-';
                return Number(value).toLocaleString('da-DK', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + '%';
            }

            function shouldShowOrdreindgangTilbudLine() {
                const el = document.getElementById('ordreindgangShowTilbud');
                return !!el && el.checked === true;
            }

            function setOrdreindgangStatus(message) {
                const el = document.getElementById('ordreindgangStatus');
                if (el) el.textContent = String(message || '');
            }

            function buildModulePrintStyles(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const orientation = safeOptions.orientation === 'landscape' ? 'landscape' : 'portrait';
                const reportMaxWidth = orientation === 'landscape' ? '277mm' : '190mm';
                return '<style>' +
                    '@page { size: A4 ' + orientation + '; margin: 12mm; }' +
                    'body { font-family: Segoe UI, Arial, sans-serif; margin:0; color:#172b3c; background:#fff; }' +
                    '.report { max-width: ' + reportMaxWidth + '; margin:0 auto; }' +
                    '.report-head { border-bottom:2px solid #d9e6f5; padding:0 0 8px 0; margin:0 0 10px 0; }' +
                    '.report-title { font-size:22px; font-weight:800; color:#0f3560; margin:0; }' +
                    '.report-sub { margin:3px 0 0 0; font-size:12px; color:#4b6783; }' +
                    '.report-meta { display:flex; gap:8px 16px; flex-wrap:wrap; margin-top:8px; font-size:12px; color:#355675; }' +
                    '.report-meta strong { color:#0f3560; }' +
                    '.kpis { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:8px; margin:10px 0 12px 0; }' +
                    '.kpi { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#f8fbff; }' +
                    '.kpi .lbl { font-size:11px; color:#4f6d8c; text-transform:uppercase; font-weight:700; }' +
                    '.kpi .val { margin-top:3px; font-size:16px; font-weight:800; color:#0f3560; }' +
                    '.section { margin-top:10px; page-break-inside:avoid; }' +
                    '.section h3 { margin:0 0 6px 0; font-size:14px; color:#214867; border-bottom:1px solid #e2ebf7; padding-bottom:4px; }' +
                    '.chart-box { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#fff; }' +
                    '.chart-box svg { width:100%; height:auto; display:block; }' +
                    '.legend-line { margin:0 0 6px 0; font-size:12px; color:#355675; font-weight:700; }' +
                    '.omsaetning-legend-item,.ordreindgang-legend-item { display:inline-flex; align-items:center; gap:6px; margin-right:10px; font-size:12px; font-weight:700; color:#355675; }' +
                    '.omsaetning-legend-swatch,.ordreindgang-legend-swatch { width:12px; height:12px; border-radius:3px; display:inline-block; }' +
                    '.context-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px; }' +
                    '.context-card { border:1px solid #dbe8f9; border-radius:8px; padding:8px; background:#f8fbff; }' +
                    '.context-card h4 { margin:0 0 6px 0; font-size:12px; color:#214867; text-transform:uppercase; }' +
                    '.context-card p { margin:0; font-size:12px; color:#355675; }' +
                    '.context-line { margin-top:4px; font-size:12px; color:#355675; }' +
                    '.pill-list { display:flex; flex-wrap:wrap; gap:4px; margin-top:6px; }' +
                    '.pill { display:inline-block; padding:2px 7px; border-radius:999px; background:#e8f1fc; color:#1f476a; font-size:11px; font-weight:700; }' +
                    '.report-landscape .kpis { grid-template-columns:repeat(4,minmax(0,1fr)); }' +
                    '.report-landscape table { font-size:11px; }' +
                    '.table-wrap { border:1px solid #dbe8f9; border-radius:8px; overflow:hidden; }' +
                    'table { width:100%; border-collapse:collapse; font-size:12px; }' +
                    'th { background:#1565c0; color:#fff; text-align:left; padding:7px 8px; }' +
                    'td { border-bottom:1px solid #e7eef8; padding:6px 8px; }' +
                    'td[style*="text-align:right"], th[style*="text-align:right"] { text-align:right !important; }' +
                    '.muted { color:#6a829b; font-size:11px; }' +
                    '</style>';
            }

            function openModulePrintWindow(title, subtitle, metaHtml, kpiHtml, sectionsHtml, options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const orientation = safeOptions.orientation === 'landscape' ? 'landscape' : 'portrait';
                const html = '<!doctype html><html><head><meta charset="utf-8" />' +
                    '<title>' + escapeHtmlFE(title) + '</title>' +
                    buildModulePrintStyles({ orientation }) +
                    '</head><body>' +
                    '<div class="report report-' + orientation + '">' +
                    '<header class="report-head">' +
                    '<h1 class="report-title">' + escapeHtmlFE(title) + '</h1>' +
                    '<p class="report-sub">' + escapeHtmlFE(subtitle) + '</p>' +
                    '<div class="report-meta">' + metaHtml + '</div>' +
                    '</header>' +
                    '<section class="kpis">' + kpiHtml + '</section>' +
                    sectionsHtml +
                    '</div>' +
                    '</body></html>';

                const existing = document.getElementById('modulePrintFrame');
                if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

                const frame = document.createElement('iframe');
                frame.id = 'modulePrintFrame';
                frame.setAttribute('aria-hidden', 'true');
                frame.style.position = 'fixed';
                frame.style.right = '0';
                frame.style.bottom = '0';
                frame.style.width = '1px';
                frame.style.height = '1px';
                frame.style.border = '0';
                frame.style.opacity = '0';
                document.body.appendChild(frame);

                const cleanup = () => {
                    const current = document.getElementById('modulePrintFrame');
                    if (current && current.parentNode) current.parentNode.removeChild(current);
                };

                const runPrint = () => {
                    const w = frame.contentWindow;
                    if (!w) {
                        alert('Kunne ikke oprette print-visning. Prøv igen.');
                        cleanup();
                        return;
                    }

                    w.onafterprint = cleanup;
                    w.focus();
                    w.print();
                    setTimeout(cleanup, 2000);
                };

                frame.onload = () => {
                    setTimeout(runPrint, 120);
                };

                // srcdoc avoids popup blockers and is more reliable in Electron than window.open.
                frame.srcdoc = html;
            }

            function getOrientationLabelDa(orientation) {
                return orientation === 'landscape' ? 'Liggende' : 'Stående';
            }

            function getPrintOrientationPreference(moduleKey) {
                const id = moduleKey === 'ordreindgang' ? 'ordreindgangPrintOrientation' : 'omsaetningPrintOrientation';
                const raw = String((document.getElementById(id) || {}).value || 'auto').toLowerCase();
                if (raw === 'portrait' || raw === 'landscape') return raw;
                return 'auto';
            }

            function resolvePrintOrientation(preference, autoOrientation) {
                if (preference === 'portrait' || preference === 'landscape') return preference;
                return autoOrientation === 'landscape' ? 'landscape' : 'portrait';
            }

            function getOrientationSourceLabelDa(preference) {
                return preference === 'auto' ? 'Auto' : 'Manuel';
            }

            function chooseOmsaetningPrintOrientation(context) {
                const c = context && typeof context === 'object' ? context : {};
                if (c.detailsOpen || c.thresholdOpen) return 'landscape';
                if (Number(c.selectedCustomers || 0) > 1) return 'landscape';
                if (Number(c.selectedAccounts || 0) > 6) return 'landscape';
                if (Number(c.periods || 0) > 12) return 'landscape';
                return 'portrait';
            }

            function chooseOrdreindgangPrintOrientation(context) {
                const c = context && typeof context === 'object' ? context : {};
                if (c.weeklyOpen || c.customersOpen) return 'landscape';
                if (Number(c.weeks || 0) > 20) return 'landscape';
                return 'portrait';
            }

            function printOmsaetningReport() {
                const chartsWrap = document.getElementById('omsaetningChartsWrap');
                if (!chartsWrap || chartsWrap.style.display === 'none') {
                    alert('Ingen Omsætning-data at printe. Tryk Opdater først.');
                    return;
                }

                const detailsWrapEl = document.getElementById('omsaetningTableWrap');
                const thresholdWrapEl = document.getElementById('omsaetningThresholdWrap');
                const detailsOpen = !!detailsWrapEl && detailsWrapEl.style.display !== 'none';
                const thresholdOpen = !!thresholdWrapEl && thresholdWrapEl.style.display !== 'none';

                const fra = String((document.getElementById('omsaetningFraMonth') || {}).value || '-');
                const til = String((document.getElementById('omsaetningTilMonth') || {}).value || '-');
                const total = String((document.getElementById('omsaetningTotalMio') || {}).textContent || '-').trim();
                const rows = String((document.getElementById('omsaetningRowsCount') || {}).textContent || '-').trim();
                const periods = String((document.getElementById('omsaetningPeriodsCount') || {}).textContent || '-').trim();
                const stackedSvg = (document.getElementById('omsaetningStackedChart') || {}).outerHTML || '';
                const trendSvg = (document.getElementById('omsaetningTrendChart') || {}).outerHTML || '';
                const legend = (document.getElementById('omsaetningLegend') || {}).innerHTML || '';
                const thresholdTable = (document.getElementById('omsaetningThresholdTable') || {}).innerHTML || '';
                const detailsTable = (document.getElementById('omsaetningTableWrap') || {}).innerHTML || '<div class="muted">Ingen detaljetabel tilgængelig.</div>';
                const customerModeText = String((document.getElementById('omsaetningCustomerMode') || {}).textContent || '').trim();
                const customerThresholdsHtml = (document.getElementById('omsaetningCustomerThresholds') || {}).innerHTML || '';
                const selectedAccounts = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean).map(v => String(v));
                const selectedCustomerEntries = Array.from(omsaetningSelectedCustomers.entries());
                const numericPeriods = Number(periods.replace(/./g, '').replace(',', '.')) || 0;
                const orientationPreference = getPrintOrientationPreference('omsaetning');
                const autoOrientation = chooseOmsaetningPrintOrientation({
                    detailsOpen,
                    thresholdOpen,
                    selectedCustomers: selectedCustomerEntries.length,
                    selectedAccounts: selectedAccounts.length,
                    periods: numericPeriods
                });
                const orientation = resolvePrintOrientation(orientationPreference, autoOrientation);

                const accountPills = selectedAccounts.slice(0, 18).map(v => '<span class="pill">' + escapeHtmlFE(v) + '</span>').join('');
                const accountRest = selectedAccounts.length > 18
                    ? '<span class="pill">+' + escapeHtmlFE(String(selectedAccounts.length - 18)) + '</span>'
                    : '';
                const customerPills = selectedCustomerEntries.slice(0, 12).map(([custNo, custName]) => {
                    const label = String(custName || '').trim() || String(custNo || '').trim();
                    return '<span class="pill">' + escapeHtmlFE(label + ' (' + String(custNo || '') + ')') + '</span>';
                }).join('');
                const customerRest = selectedCustomerEntries.length > 12
                    ? '<span class="pill">+' + escapeHtmlFE(String(selectedCustomerEntries.length - 12)) + '</span>'
                    : '';

                const metaHtml =
                    '<div><strong>Periode:</strong> ' + escapeHtmlFE(fra + ' → ' + til) + '</div>' +
                    '<div><strong>Layoutvalg:</strong> ' + escapeHtmlFE(getOrientationSourceLabelDa(orientationPreference)) + '</div>' +
                    '<div><strong>Layout:</strong> ' + escapeHtmlFE(getOrientationLabelDa(orientation)) + '</div>' +
                    '<div><strong>Udskrevet:</strong> ' + escapeHtmlFE(new Date().toLocaleString('da-DK')) + '</div>';

                const kpiHtml =
                    '<div class="kpi"><div class="lbl">Omsætning (Mio)</div><div class="val">' + escapeHtmlFE(total) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Rækker</div><div class="val">' + escapeHtmlFE(rows) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Perioder</div><div class="val">' + escapeHtmlFE(periods) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Modul</div><div class="val">Omsætning</div></div>';

                const contextSection =
                    '<section class="section"><h3>Aktive filtre og visning</h3>' +
                        '<div class="context-grid">' +
                            '<div class="context-card">' +
                                '<h4>Konti</h4>' +
                                '<p>' + escapeHtmlFE(String(selectedAccounts.length)) + ' aktiv</p>' +
                                '<div class="pill-list">' + accountPills + accountRest + '</div>' +
                            '</div>' +
                            '<div class="context-card">' +
                                '<h4>Kunder</h4>' +
                                '<p>' + escapeHtmlFE(String(selectedCustomerEntries.length)) + ' valgt</p>' +
                                '<div class="pill-list">' + (selectedCustomerEntries.length > 0 ? (customerPills + customerRest) : '<span class="pill">Ingen kunde</span>') + '</div>' +
                                '<div class="context-line"><strong>Visning:</strong> ' + escapeHtmlFE(customerModeText || 'Standardvisning') + '</div>' +
                            '</div>' +
                        '</div>' +
                    '</section>';

                const customerCompareSection = selectedCustomerEntries.length > 1
                    ? '<section class="section"><h3>Kunde-sammenligning (aktiv)</h3>' +
                        '<div class="context-card"><p>Flere kunder er valgt. Graf og tabeller er baseret på sammenligning pr. måned.</p></div>' +
                        '<div class="context-line"></div>' +
                        '<div class="table-wrap">' + (customerThresholdsHtml || '<div class="muted" style="padding:8px;">Ingen kundetærskler tilgængelig.</div>') + '</div>' +
                    '</section>'
                    : '';

                const thresholdSection = thresholdTable
                    ? '<section class="section"><h3>Tærskel-tabel</h3><div class="table-wrap">' + thresholdTable + '</div></section>'
                    : '';

                const sectionsHtml =
                    contextSection +
                    customerCompareSection +
                    '<section class="section"><h3>Stacked graf</h3><div class="chart-box"><div class="legend-line">' + legend + '</div>' + stackedSvg + '</div></section>' +
                    '<section class="section"><h3>Trend graf</h3><div class="chart-box">' + trendSvg + '</div></section>' +
                    (thresholdOpen ? thresholdSection : '') +
                    (detailsOpen ? '<section class="section"><h3>Detaljer</h3><div class="table-wrap">' + detailsTable + '</div></section>' : '');

                openModulePrintWindow(
                    'Gantech Operations Hub - Omsætning',
                    'Rapportudskrift (' + getOrientationLabelDa(orientation).toLowerCase() + ')',
                    metaHtml,
                    kpiHtml,
                    sectionsHtml,
                    { orientation }
                );
            }

            function printOrdreindgangReport() {
                const chartsWrap = document.getElementById('ordreindgangChartsWrap');
                if (!chartsWrap || chartsWrap.style.display === 'none') {
                    alert('Ingen Ordreindgang-data at printe. Tryk Opdater først.');
                    return;
                }

                const weeklyWrapEl = document.getElementById('ordreindgangWeeklyTable');
                const customersWrapEl = document.getElementById('ordreindgangCustomersTable');
                const weeklyOpen = !!weeklyWrapEl && weeklyWrapEl.style.display !== 'none';
                const customersOpen = !!customersWrapEl && customersWrapEl.style.display !== 'none';

                const fraWeek = String((document.getElementById('ordreindgangFraWeek') || {}).value || '-');
                const tilWeek = String((document.getElementById('ordreindgangTilWeek') || {}).value || '-');
                const totalOrd = String((document.getElementById('ordreindgangTotalOrd') || {}).textContent || '-').trim();
                const totalTilbud = String((document.getElementById('ordreindgangTotalTilbud') || {}).textContent || '-').trim();
                const avgOrd = String((document.getElementById('ordreindgangAvgOrd') || {}).textContent || '-').trim();
                const conv = String((document.getElementById('ordreindgangConv') || {}).textContent || '-').trim();
                const trendSvg = (document.getElementById('ordreindgangTrendChart') || {}).outerHTML || '';
                const legend = (document.getElementById('ordreindgangLegend') || {}).innerHTML || '';
                const weeklyTable = (document.getElementById('ordreindgangWeeklyTable') || {}).innerHTML || '<div class="muted">Ingen ugetabel tilgængelig.</div>';
                const customerTable = (document.getElementById('ordreindgangCustomersTable') || {}).innerHTML || '<div class="muted">Ingen kundetabel tilgængelig.</div>';
                const statusText = String((document.getElementById('ordreindgangStatus') || {}).textContent || '').trim();
                const tilbudEnabled = !!((document.getElementById('ordreindgangShowTilbud') || {}).checked);
                const weekCount = Array.isArray(ordreindgangLastPayload && ordreindgangLastPayload.weeklyRows)
                    ? ordreindgangLastPayload.weeklyRows.length
                    : 0;
                const orientationPreference = getPrintOrientationPreference('ordreindgang');
                const autoOrientation = chooseOrdreindgangPrintOrientation({
                    weeklyOpen,
                    customersOpen,
                    weeks: weekCount
                });
                const orientation = resolvePrintOrientation(orientationPreference, autoOrientation);

                const metaHtml =
                    '<div><strong>Periode:</strong> ' + escapeHtmlFE(fraWeek + ' → ' + tilWeek) + '</div>' +
                    '<div><strong>Tilbud-linje:</strong> ' + (tilbudEnabled ? 'Aktiv' : 'Skjult') + '</div>' +
                    '<div><strong>Layoutvalg:</strong> ' + escapeHtmlFE(getOrientationSourceLabelDa(orientationPreference)) + '</div>' +
                    '<div><strong>Layout:</strong> ' + escapeHtmlFE(getOrientationLabelDa(orientation)) + '</div>' +
                    '<div><strong>Udskrevet:</strong> ' + escapeHtmlFE(new Date().toLocaleString('da-DK')) + '</div>';

                const kpiHtml =
                    '<div class="kpi"><div class="lbl">Total Ordre</div><div class="val">' + escapeHtmlFE(totalOrd) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Total Tilbud</div><div class="val">' + escapeHtmlFE(totalTilbud) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Gns. Ordre</div><div class="val">' + escapeHtmlFE(avgOrd) + '</div></div>' +
                    '<div class="kpi"><div class="lbl">Tilbud → Ordre</div><div class="val">' + escapeHtmlFE(conv) + '</div></div>';

                const contextSection =
                    '<section class="section"><h3>Aktive filtre og visning</h3>' +
                        '<div class="context-grid">' +
                            '<div class="context-card">' +
                                '<h4>Periode</h4>' +
                                '<p>' + escapeHtmlFE(fraWeek + ' → ' + tilWeek) + '</p>' +
                            '</div>' +
                            '<div class="context-card">' +
                                '<h4>Tilbud-linje</h4>' +
                                '<p>' + (tilbudEnabled ? 'Aktiv (vises i graf)' : 'Skjult (kun ordre)') + '</p>' +
                                '<div class="context-line"><strong>Status:</strong> ' + escapeHtmlFE(statusText || 'Ingen status') + '</div>' +
                            '</div>' +
                        '</div>' +
                    '</section>';

                const sectionsHtml =
                    contextSection +
                    '<section class="section"><h3>Ugeudvikling</h3><div class="chart-box"><div class="legend-line">' + legend + '</div>' + trendSvg + '</div></section>' +
                    (weeklyOpen ? '<section class="section"><h3>Ugetabel</h3><div class="table-wrap">' + weeklyTable + '</div></section>' : '') +
                    (customersOpen ? '<section class="section"><h3>Topkunder</h3><div class="table-wrap">' + customerTable + '</div></section>' : '');

                openModulePrintWindow(
                    'Gantech Operations Hub - Ordreindgang',
                    'Rapportudskrift (' + getOrientationLabelDa(orientation).toLowerCase() + ')',
                    metaHtml,
                    kpiHtml,
                    sectionsHtml,
                    { orientation }
                );
            }

            function applyOrdreindgangDefaultWeeks() {
                const fraEl = document.getElementById('ordreindgangFraWeek');
                const tilEl = document.getElementById('ordreindgangTilWeek');
                if (!fraEl || !tilEl) return;

                const today = new Date();
                const toWeek = getIsoWeekMeta(today);
                const fromYear = toWeek.isoYear - 1;
                fraEl.value = String(fromYear) + '04';
                tilEl.value = toWeek.weekKey;
            }

            function buildOrdreindgangRange() {
                const fraEl = document.getElementById('ordreindgangFraWeek');
                const tilEl = document.getElementById('ordreindgangTilWeek');
                const fraMeta = parseWeekKeyMeta(fraEl ? fraEl.value : '');
                const tilMeta = parseWeekKeyMeta(tilEl ? tilEl.value : '');
                if (fraEl) fraEl.value = normalizeWeekKeyInput(fraEl.value);
                if (tilEl) tilEl.value = normalizeWeekKeyInput(tilEl.value);
                if (!fraMeta || !tilMeta) return null;

                const fromStart = getIsoWeekStartDate(fraMeta.year, fraMeta.week);
                const toStart = getIsoWeekStartDate(tilMeta.year, tilMeta.week);
                if (toStart.getTime() < fromStart.getTime()) return null;

                return {
                    fraWeek: fraMeta.raw,
                    tilWeek: tilMeta.raw
                };
            }

            function buildOrdreindgangSummaryCacheKey(range) {
                return JSON.stringify({
                    fraWeek: String(range.fraWeek || ''),
                    tilWeek: String(range.tilWeek || '')
                });
            }

            async function fetchOrdreindgangSummaryCached(range, options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const forceRefresh = safeOptions.forceRefresh === true;
                const cacheKey = buildOrdreindgangSummaryCacheKey(range);

                if (!forceRefresh) {
                    const cached = getOmsaetningCacheEntry(ordreindgangSummaryCache, cacheKey, ORDREINDGANG_SUMMARY_CACHE_TTL_MS);
                    if (cached) return cached;
                }

                if (!forceRefresh) {
                    const inFlight = ordreindgangSummaryInFlight.get(cacheKey);
                    if (inFlight) return await inFlight;
                }

                const reqPromise = (async () => {
                    const query = new URLSearchParams({
                        fraWeek: String(range.fraWeek),
                        tilWeek: String(range.tilWeek)
                    });
                    const response = await fetch('/ordreindgang/summary?' + query.toString());
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    setOmsaetningCacheEntry(ordreindgangSummaryCache, cacheKey, payload);
                    return payload;
                })();

                ordreindgangSummaryInFlight.set(cacheKey, reqPromise);
                try {
                    return await reqPromise;
                } finally {
                    ordreindgangSummaryInFlight.delete(cacheKey);
                }
            }

            function renderOrdreindgangTrendChart(rows) {
                const wrap = document.getElementById('ordreindgangChartsWrap');
                const svg = document.getElementById('ordreindgangTrendChart');
                const legendEl = document.getElementById('ordreindgangLegend');
                const safeRows = Array.isArray(rows) ? rows : [];
                const showTilbudLine = shouldShowOrdreindgangTilbudLine();
                if (!wrap || !svg) return;

                if (safeRows.length === 0) {
                    wrap.style.display = 'none';
                    svg.innerHTML = '';
                    if (legendEl) legendEl.innerHTML = '';
                    return;
                }

                const labels = safeRows.map(r => {
                    const key = String(r.weekKey || '');
                    if (/^\d{6}$/.test(key)) return key.slice(0, 4) + '-' + key.slice(4, 6);
                    return formatWeekLabel(key);
                });
                const ordValues = safeRows.map(r => Number(r.totalOrd || 0));
                const tilbudValues = safeRows.map(r => Number(r.totalTilbud || 0));
                const avgValue = Number(safeRows[0] && safeRows[0].avgOrd || 0);
                const allValues = showTilbudLine
                    ? ordValues.concat(tilbudValues).concat([avgValue])
                    : ordValues.concat([avgValue]);
                const rawMax = Math.max(...allValues, 0);
                const chartMax = rawMax <= 0 ? 1000 : (Math.ceil(rawMax / 1000) * 1000);

                const leftPad = 72;
                const topPad = 36;
                const bottomPad = 108;
                const viewportWidth = Math.max(760, (wrap.clientWidth || 0) - 24);
                const visibleWeeksTarget = Math.max(12, Math.min(24, safeRows.length || 12));
                const slot = Math.max(18, Math.min(36, Math.floor(viewportWidth / visibleWeeksTarget)));
                const width = Math.max(viewportWidth, safeRows.length * slot);
                const height = Math.max(320, Math.min(440, Math.round(window.innerHeight * 0.46)));
                const innerHeight = height - topPad - bottomPad;
                const viewWidth = leftPad + width + 20;
                const toY = val => topPad + ((chartMax - Math.max(0, Number(val || 0))) / chartMax) * innerHeight;
                const toXCenter = idx => leftPad + (idx * slot) + (slot / 2);
                const yZero = toY(0);

                let html = '<g>';
                for (let i = 0; i <= 6; i++) {
                    const ratio = i / 6;
                    const y = topPad + (innerHeight * ratio);
                    const tickVal = chartMax * (1 - ratio);
                    html += '<line x1="' + leftPad + '" y1="' + y + '" x2="' + (leftPad + width) + '" y2="' + y + '" stroke="#d9e6f8" stroke-width="1" />';
                    html += '<text x="' + (leftPad - 8) + '" y="' + (y + 5) + '" text-anchor="end" font-size="12" fill="#5f7892">' + escapeHtmlFE(Math.round(tickVal).toLocaleString('da-DK')) + '</text>';
                }

                const ordBarWidth = showTilbudLine ? 8 : 12;
                const tilbudBarWidth = 7;

                const labelStep = safeRows.length > 70 ? 4 : (safeRows.length > 52 ? 3 : (safeRows.length > 36 ? 2 : 1));
                safeRows.forEach((row, idx) => {
                    const centerX = toXCenter(idx);
                    const ordY = toY(ordValues[idx]);
                    const ordH = Math.max(1, yZero - ordY);

                    if (showTilbudLine) {
                        const tilbudY = toY(tilbudValues[idx]);
                        const tilbudH = Math.max(1, yZero - tilbudY);
                        const tilbudX = centerX - tilbudBarWidth - 1;
                        html += '<rect x="' + tilbudX + '" y="' + tilbudY + '" width="' + tilbudBarWidth + '" height="' + tilbudH + '" fill="#8ec3f7" rx="1"><title>' +
                            escapeHtmlFE(labels[idx] + ' Tilbud: ' + formatDkkDa(tilbudValues[idx])) + '</title></rect>';
                    }

                    const ordX = showTilbudLine ? (centerX + 1) : (centerX - (ordBarWidth / 2));
                    html += '<rect x="' + ordX + '" y="' + ordY + '" width="' + ordBarWidth + '" height="' + ordH + '" fill="#2f5ea5" rx="1"><title>' +
                        escapeHtmlFE(labels[idx] + ' Ordre: ' + formatDkkDa(ordValues[idx])) + '</title></rect>';

                    if ((idx % labelStep) === 0 || idx === safeRows.length - 1) {
                        html += '<text x="' + centerX + '" y="' + (topPad + innerHeight + 50) + '" text-anchor="middle" transform="rotate(-90 ' + centerX + ' ' + (topPad + innerHeight + 50) + ')" font-size="12" fill="#5f7892">' + escapeHtmlFE(labels[idx]) + '</text>';
                    }
                });

                const avgY = toY(avgValue);
                html += '<line x1="' + leftPad + '" y1="' + avgY + '" x2="' + (leftPad + width) + '" y2="' + avgY + '" stroke="#3d6eb5" stroke-width="3" />';
                html += '</g>';

                svg.setAttribute('viewBox', '0 0 ' + viewWidth + ' ' + height);
                svg.innerHTML = html;
                if (legendEl) {
                    let legendHtml = '';
                    legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#2f5ea5"></span>Ordre</span>';
                    if (showTilbudLine) {
                        legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#8ec3f7"></span>Tilbud</span>';
                    }
                    legendHtml += '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:#3d6eb5"></span>Gennem.Ordre</span>';
                    legendEl.innerHTML = legendHtml;
                }
                wrap.style.display = 'grid';
            }

            function applyOrdreindgangWeeklyCollapsedState() {
                const tableWrap = document.getElementById('ordreindgangWeeklyTable');
                const toggleBtn = document.getElementById('ordreindgangWeeklyToggleBtn');
                if (!tableWrap || !toggleBtn) return;
                tableWrap.style.display = ordreindgangWeeklyCollapsed ? 'none' : 'block';
                toggleBtn.textContent = ordreindgangWeeklyCollapsed ? 'Vis tabel' : 'Skjul tabel';
            }

            function applyOrdreindgangCustomersCollapsedState() {
                const tableWrap = document.getElementById('ordreindgangCustomersTable');
                const toggleBtn = document.getElementById('ordreindgangCustomersToggleBtn');
                if (!tableWrap || !toggleBtn) return;
                tableWrap.style.display = ordreindgangCustomersCollapsed ? 'none' : 'block';
                toggleBtn.textContent = ordreindgangCustomersCollapsed ? 'Vis tabel' : 'Skjul tabel';
            }

            function toggleOrdreindgangWeeklyTable() {
                ordreindgangWeeklyCollapsed = !ordreindgangWeeklyCollapsed;
                applyOrdreindgangWeeklyCollapsedState();
            }

            function toggleOrdreindgangCustomersTable() {
                ordreindgangCustomersCollapsed = !ordreindgangCustomersCollapsed;
                applyOrdreindgangCustomersCollapsedState();
            }

            function renderOrdreindgangWeeklyTable(rows) {
                const wrapCard = document.getElementById('ordreindgangWeeklyWrap');
                const wrap = document.getElementById('ordreindgangWeeklyTable');
                const safeRows = Array.isArray(rows) ? rows : [];
                const showTilbudColumn = shouldShowOrdreindgangTilbudLine();
                if (!wrapCard || !wrap) return;
                if (safeRows.length === 0) {
                    wrapCard.style.display = 'none';
                    wrap.innerHTML = '';
                    return;
                }

                const body = safeRows.map(row => {
                    const cells = [
                        '<td>' + escapeHtmlFE(formatWeekLabel(row.weekKey)) + '</td>',
                        '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalOrd)) + '</td>'
                    ];
                    if (showTilbudColumn) {
                        cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalTilbud)) + '</td>');
                    }
                    cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.totalBudget)) + '</td>');
                    cells.push('<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.avgOrd)) + '</td>');
                    return '<tr>' +
                        cells.join('') +
                        '</tr>';
                }).join('');

                const headers = [
                    '<th>Uge</th>',
                    '<th class="omsaetning-cell-right">Ordre</th>'
                ];
                if (showTilbudColumn) {
                    headers.push('<th class="omsaetning-cell-right">Tilbud</th>');
                }
                headers.push('<th class="omsaetning-cell-right">Budget</th>');
                headers.push('<th class="omsaetning-cell-right">Gns. ordre</th>');

                const colgroup = showTilbudColumn
                    ? '<colgroup><col style="width:16%;" /><col style="width:21%;" /><col style="width:23%;" /><col style="width:20%;" /><col style="width:20%;" /></colgroup>'
                    : '<colgroup><col style="width:18%;" /><col style="width:28%;" /><col style="width:27%;" /><col style="width:27%;" /></colgroup>';

                wrap.innerHTML = '<table class="omsaetning-table ordreindgang-weekly-table">' +
                    colgroup +
                    '<thead><tr>' + headers.join('') + '</tr></thead>' +
                    '<tbody>' + body + '</tbody></table>';
                wrapCard.style.display = 'block';
                applyOrdreindgangWeeklyCollapsedState();
            }

            function renderOrdreindgangCustomersTable(rows) {
                const wrapCard = document.getElementById('ordreindgangCustomersWrap');
                const wrap = document.getElementById('ordreindgangCustomersTable');
                const safeRows = Array.isArray(rows) ? rows : [];
                if (!wrapCard || !wrap) return;
                if (safeRows.length === 0) {
                    wrapCard.style.display = 'none';
                    wrap.innerHTML = '';
                    return;
                }

                const body = safeRows.map(row => {
                    const label = String(row.customerName || '').trim() || String(row.custNo || '-');
                    return '<tr>' +
                        '<td>' + escapeHtmlFE(label) + '</td>' +
                        '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.ordSum)) + '</td>' +
                        '<td style="text-align:right;">' + escapeHtmlFE(formatDkkDa(row.tilbudSum)) + '</td>' +
                        '<td style="text-align:right;">' + escapeHtmlFE(formatPctDa(row.conversionPct)) + '</td>' +
                        '</tr>';
                }).join('');

                wrap.innerHTML = '<table class="omsaetning-table">' +
                    '<thead><tr><th>Kunde</th><th>Ordre</th><th>Tilbud</th><th>Tilbud → Ordre</th></tr></thead>' +
                    '<tbody>' + body + '</tbody></table>';
                wrapCard.style.display = 'block';
                applyOrdreindgangCustomersCollapsedState();
            }

            function scheduleOrdreindgangAutoReload() {
                if (ordreindgangAutoReloadTimer) {
                    clearTimeout(ordreindgangAutoReloadTimer);
                }
                ordreindgangAutoReloadTimer = setTimeout(() => {
                    ordreindgangAutoReloadTimer = null;
                    loadOrdreindgangSummary({ silentValidation: true });
                }, ORDREINDGANG_AUTO_RELOAD_DELAY_MS);
            }

            async function initializeOrdreindgangIfNeeded() {
                if (ordreindgangInitialized) return;
                ordreindgangInitialized = true;
                window.addEventListener('resize', () => {
                    if (ordreindgangResizeTimer) clearTimeout(ordreindgangResizeTimer);
                    ordreindgangResizeTimer = setTimeout(() => {
                        ordreindgangResizeTimer = null;
                        renderOrdreindgangFromLastPayload();
                    }, 120);
                });
                applyOrdreindgangDefaultWeeks();
                await loadOrdreindgangSummary({ forceRefresh: true });
            }

            function renderOrdreindgangFromLastPayload() {
                if (!ordreindgangLastPayload || !Array.isArray(ordreindgangLastPayload.weeklyRows)) return;
                renderOrdreindgangTrendChart(ordreindgangLastPayload.weeklyRows);
                renderOrdreindgangWeeklyTable(ordreindgangLastPayload.weeklyRows);
            }

            async function loadOrdreindgangSummary(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const emptyEl = document.getElementById('ordreindgangEmpty');
                const loadBtn = document.getElementById('ordreindgangLoadBtn');
                const range = buildOrdreindgangRange();

                if (!range) {
                    setOrdreindgangStatus('Ugyldig ugeperiode. Brug format YYYYWW.');
                    if (safeOptions.silentValidation !== true) {
                        alert('Ugyldig ugeperiode. Brug format YYYYWW.');
                    }
                    return;
                }

                if (loadBtn) loadBtn.disabled = true;
                if (emptyEl) {
                    emptyEl.style.display = 'block';
                    emptyEl.textContent = 'Henter ordreindgang...';
                }
                setOrdreindgangStatus('Henter data...');

                try {
                    const payload = await fetchOrdreindgangSummaryCached(range, { forceRefresh: safeOptions.forceRefresh === true });
                    ordreindgangLastPayload = payload;
                    const kpis = payload && payload.kpis ? payload.kpis : {};
                    const weeklyRows = Array.isArray(payload && payload.weeklyRows) ? payload.weeklyRows : [];
                    const customerRows = Array.isArray(payload && payload.customerRows) ? payload.customerRows : [];

                    const ordEl = document.getElementById('ordreindgangTotalOrd');
                    const tilbudEl = document.getElementById('ordreindgangTotalTilbud');
                    const avgEl = document.getElementById('ordreindgangAvgOrd');
                    const convEl = document.getElementById('ordreindgangConv');

                    if (ordEl) ordEl.textContent = formatDkkDa(kpis.totalOrdSum || 0);
                    if (tilbudEl) tilbudEl.textContent = formatDkkDa(kpis.totalTilbudSum || 0);
                    if (avgEl) avgEl.textContent = formatDkkDa(kpis.avgSumOrd || 0);
                    if (convEl) convEl.textContent = formatPctDa(kpis.conversionPct);

                    renderOrdreindgangTrendChart(weeklyRows);
                    renderOrdreindgangWeeklyTable(weeklyRows);
                    renderOrdreindgangCustomersTable(customerRows);

                    if (emptyEl) {
                        if (weeklyRows.length === 0) {
                            emptyEl.style.display = 'block';
                            emptyEl.textContent = 'Ingen data i valgt ugeperiode.';
                        } else {
                            emptyEl.style.display = 'none';
                        }
                    }

                    setOrdreindgangStatus('Periode: ' + formatWeekLabel(range.fraWeek) + ' til ' + formatWeekLabel(range.tilWeek) + ' · rækker: ' + String(weeklyRows.length));
                } catch (err) {
                    setOrdreindgangStatus('Fejl ved hentning af ordreindgang.');
                    if (emptyEl) {
                        emptyEl.style.display = 'block';
                        emptyEl.textContent = 'Fejl: ' + (err && err.message ? err.message : 'ukendt fejl');
                    }
                } finally {
                    if (loadBtn) loadBtn.disabled = false;
                }
            }

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
                const options = opts && typeof opts === 'object' ? opts : {};
                const clickable = options.clickable === true;
                const chartResGr = String(options.resGr || '').trim();
                const chartParity = options.parity === 0 ? 0 : 1;
                const activeDayKey = String(options.activeDayKey || '').trim();
                const sourceRows = Array.isArray(dayRows) ? dayRows : [];
                const collapseBeforeToday = options.collapseBeforeToday !== false;
                const activeDateInput = document.getElementById('belastningToDay');
                const todayRaw = (activeDateInput && activeDateInput.value)
                    ? activeDateInput.value
                    : new Date().toISOString().slice(0, 10);
                const todayCut = (() => {
                    const t = new Date(todayRaw + 'T00:00:00').getTime();
                    return Number.isNaN(t) ? Date.now() : t;
                })();

                let rows = [];
                if (collapseBeforeToday) {
                    const groupedByDay = new Map();
                    sourceRows.forEach((row, index) => {
                        const dateKey = normalizeBelastningDateKey(row && row.Dato, row && row.DatoX);
                        const dateSort = getBelastningDateSortValue(dateKey);
                        const isNullDate = !row.Dato && !String(row.DatoX || '').trim();
                        const isBefore = isNullDate || dateSort < todayCut;
                        const bucketKey = isBefore ? '__before__' : (dateKey || ('__unknown__' + index));
                        if (!groupedByDay.has(bucketKey)) {
                            groupedByDay.set(bucketKey, {
                                Dato: isBefore ? null : (row && row.Dato),
                                DatoX: isBefore ? '-' : normalizeBelastningDisplayDate(row && row.Dato, row && row.DatoX),
                                Kap: 0,
                                Resv: 0,
                                Aften: 0,
                                __dayKey: isBefore ? 'before' : dateKey,
                                __dateLabel: isBefore ? '-' : normalizeBelastningDisplayDate(row && row.Dato, row && row.DatoX),
                                __sort: isBefore ? Number.MIN_SAFE_INTEGER : dateSort
                            });
                        }
                        const bucket = groupedByDay.get(bucketKey);
                        bucket.Kap += isBefore ? 0 : Number(row && row.Kap || 0);
                        bucket.Resv += Number(row && row.Resv || 0);
                        bucket.Aften += Number(row && row.Aften || 0);
                    });
                    rows = Array.from(groupedByDay.values()).sort((a, b) => Number(a.__sort || 0) - Number(b.__sort || 0));
                } else {
                    rows = sourceRows.slice().map((row, index) => {
                        const dateKey = normalizeBelastningDateKey(row && row.Dato, row && row.DatoX);
                        return {
                            ...row,
                            __dayKey: dateKey || ('__unknown__' + index),
                            __dateLabel: normalizeBelastningDisplayDate(row && row.Dato, row && row.DatoX),
                            __sort: getBelastningDateSortValue(dateKey)
                        };
                    }).sort((a, b) => Number(a.__sort || 0) - Number(b.__sort || 0));
                }

                if (rows.length === 0) return '';

                const leftPad = 40;
                const rightPad = 12;
                const topPad = 22;
                const bottomPad = 86;
                const parentWidth = Math.max(420, (window.innerWidth || 1200) * 0.42);
                const targetDaysOnScreen = Math.max(12, Math.min(30, rows.length));
                const groupW = Math.max(16, Math.min(30, Math.floor(parentWidth / Math.max(1, targetDaysOnScreen))));
                const innerW = Math.max(parentWidth, rows.length * groupW);
                const innerH = 180;
                const svgW = leftPad + innerW + rightPad;
                const svgH = topPad + innerH + bottomPad;
                const maxVal = rows.reduce((max, row) => {
                    return Math.max(max, Number(row.Kap || 0), Number(row.Resv || 0), Number(row.Aften || 0));
                }, 1);

                const yFor = (v) => topPad + innerH - (Math.max(0, Number(v || 0)) / maxVal) * innerH;
                const xFor = (i) => leftPad + i * groupW + (groupW / 2);
                const barW = Math.max(3, Math.min(7, groupW / 3.4));

                const grid = [];
                for (let i = 0; i <= 4; i++) {
                    const y = topPad + (innerH * i / 4);
                    const val = formatCount(Math.round(maxVal * (1 - i / 4)));
                    grid.push('<line class="grid" x1="' + leftPad + '" y1="' + y + '" x2="' + (svgW - rightPad) + '" y2="' + y + '"></line>');
                    grid.push('<text class="label" x="' + (leftPad - 4) + '" y="' + (y + 3) + '" text-anchor="end">' + val + '</text>');
                }

                const bars = rows.map((row, i) => {
                    const x = xFor(i);
                    const kap = Number(row.Kap || 0);
                    const resv = Number(row.Resv || 0);
                    const aften = Number(row.Aften || 0);
                    const dateLabel = String(row.__dateLabel || normalizeBelastningDisplayDate(row.Dato, row.DatoX));
                    const dayKey = String(row.__dayKey || normalizeBelastningDateKey(row.Dato, row.DatoX));
                    const ky = yFor(kap);
                    const ry = yFor(resv);
                    const ay = yFor(aften);
                    const kH = topPad + innerH - ky;
                    const rH = topPad + innerH - ry;
                    const aH = topPad + innerH - ay;
                    const hitX = x - (barW * 1.9);
                    const hitW = Math.max(barW * 3.8, 10);
                    const canClick = clickable && chartResGr && dayKey;
                    const clickAttr = canClick
                        ? (' onclick="onBelastningDayColumnClick(\'' + escapeJsSingle(chartResGr) + '\',' + chartParity + ',\'' + escapeJsSingle(dayKey) + '\', event)"')
                        : '';
                    const bandClass = 'belastning-day-band' + (activeDayKey && dayKey === activeDayKey ? ' active' : '');
                    return ''
                        + '<rect class="' + bandClass + '" x="' + hitX + '" y="' + topPad + '" width="' + hitW + '" height="' + innerH + '"' + clickAttr + '></rect>'
                        + '<rect class="belastning-series-kap" x="' + (x - barW * 1.5) + '" y="' + ky + '" width="' + barW + '" height="' + kH + '"><title>' + escapeHtmlFE(dateLabel + ' Kapacitet: ' + formatBelastningMinutes(kap)) + '</title></rect>'
                        + '<rect class="belastning-series-resv" x="' + (x - barW * 0.5) + '" y="' + ry + '" width="' + barW + '" height="' + rH + '"><title>' + escapeHtmlFE(dateLabel + ' Reservationer: ' + formatBelastningMinutes(resv)) + '</title></rect>'
                        + '<rect class="belastning-series-aften" x="' + (x + barW * 0.5) + '" y="' + ay + '" width="' + barW + '" height="' + aH + '"><title>' + escapeHtmlFE(dateLabel + ' Rest Aften: ' + formatBelastningMinutes(aften)) + '</title></rect>';
                }).join('');

                // Show every day label; dates are already rotated to reduce overlap.
                const labelStep = 1;
                const labels = rows.map((row, i) => {
                    const isToday = row.__dayKey === todayRaw;
                    if (i % labelStep !== 0 && i !== rows.length - 1 && !isToday) return '';
                    const txt = escapeHtmlFE(String(row.__dateLabel || normalizeBelastningDisplayDate(row.Dato, row.DatoX) || ''));
                    const labelY = topPad + innerH + 44;
                    const cx = xFor(i);
                    return '<text class="label" x="' + cx + '" y="' + labelY + '" text-anchor="middle" transform="rotate(-90 ' + cx + ' ' + labelY + ')">' + txt + '</text>';
                }).join('');

                const legendX = leftPad + 4;
                const legendY = 10;
                const legend = ''
                    + '<rect class="belastning-series-kap" x="' + legendX + '" y="' + legendY + '" width="10" height="10"></rect><text class="label" x="' + (legendX + 14) + '" y="' + (legendY + 9) + '">Kapacitet</text>'
                    + '<rect class="belastning-series-resv" x="' + (legendX + 88) + '" y="' + legendY + '" width="10" height="10"></rect><text class="label" x="' + (legendX + 102) + '" y="' + (legendY + 9) + '">Reservationer</text>'
                    + '<rect class="belastning-series-aften" x="' + (legendX + 196) + '" y="' + legendY + '" width="10" height="10"></rect><text class="label" x="' + (legendX + 210) + '" y="' + (legendY + 9) + '">Rest Aften</text>';

                return '<svg class="belastning-svg" style="min-width:' + svgW + 'px" viewBox="0 0 ' + svgW + ' ' + svgH + '" preserveAspectRatio="xMinYMin meet">'
                    + '<line class="axis" x1="' + leftPad + '" y1="' + topPad + '" x2="' + leftPad + '" y2="' + (topPad + innerH) + '"></line>'
                    + '<line class="axis" x1="' + leftPad + '" y1="' + (topPad + innerH) + '" x2="' + (svgW - rightPad) + '" y2="' + (topPad + innerH) + '"></line>'
                    + grid.join('')
                    + bars
                    + labels
                    + legend
                    + '</svg>';
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

                for (const row of orderRows) {
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
                await loadBelastningGrafisk({ forceRefresh: true });
            }

            async function loadBelastningGrafisk(options) {
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
                    const query = new URLSearchParams({
                        toDay: filters.today,
                        dage: String(filters.dage),
                        resGr: filters.resGr,
                        ord: filters.ord,
                        kunde: filters.kunde
                    });
                    const response = await fetch('/belastning/grafisk?' + query.toString());
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    if (!payload.ok) throw new Error(payload.error || 'Belastning grafisk fejl');
                    belastningLastPayload = payload;
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

            function showAccessGate() {
                const overlay = document.getElementById('accessGateOverlay');
                const userInput = document.getElementById('accessGateUserInput');
                const input = document.getElementById('accessGateInput');
                const err = document.getElementById('accessGateError');
                if (!overlay) return;
                if (err) err.textContent = '';
                if (userInput) {
                    const currentName = sanitizeDisplayName(loggedUserDisplayName);
                    if (!String(userInput.value || '').trim() && currentName && currentName !== 'Bruger') {
                        userInput.value = currentName;
                    }
                }
                overlay.style.display = 'flex';
                refreshSideMenuAuthState();
                setTimeout(() => {
                    if (userInput && !String(userInput.value || '').trim()) {
                        userInput.focus();
                        return;
                    }
                    if (input) input.focus();
                }, 30);
            }

            function hideAccessGate() {
                const overlay = document.getElementById('accessGateOverlay');
                if (!overlay) return;
                overlay.style.display = 'none';
                refreshSideMenuAuthState();
            }

            function submitAccessCode() {
                const userInput = document.getElementById('accessGateUserInput');
                const input = document.getElementById('accessGateInput');
                const err = document.getElementById('accessGateError');
                const btn = document.getElementById('accessGateBtn');
                const userName = sanitizeDisplayName(userInput ? userInput.value : '');
                const value = input ? String(input.value || '').trim() : '';
                if (value !== ACCESS_CODE) {
                    if (err) err.textContent = 'Forkert kode.';
                    if (input) {
                        input.select();
                        input.focus();
                    }
                    return;
                }

                if (err) err.textContent = 'Åbner...';
                if (btn) {
                    btn.disabled = true;
                    btn.textContent = 'Åbner...';
                }

                setTimeout(() => {
                    try {
                        if (userName && userName !== 'Bruger') {
                            setLoggedUserDisplayName(userName);
                        }
                        accessGranted = true;
                        hideAccessGate();
                        refreshSideMenuAuthState();
                        initializeAfterAccess();
                    } catch (e) {
                        accessGranted = false;
                        showAccessGate();
                        if (err) err.textContent = 'Fejl ved åbning: ' + (e && e.message ? e.message : 'ukendt fejl');
                    } finally {
                        if (btn) {
                            btn.disabled = false;
                            btn.textContent = 'Åbn';
                        }
                    }
                }, 0);
            }

            function initializeAfterAccess() {
                startWarmupPolling();
                loadOrderList(false);
                setTimeout(() => {
                    if (!orderListData || orderListData.length === 0) {
                        loadOrderList(true);
                    }
                }, 2500);
                startOrderListAutoRefresh();
                startDashboardUpdatePolling();

                const params = new URLSearchParams(window.location.search);
                if (params.has('ord')) {
                    document.getElementById('orderInput').value = params.get('ord');
                    openModule('efterkalk');
                    searchOrder();
                    return;
                }
                goToDashboard();
            }

            function openModule(moduleKey) {
                const dashboard = document.getElementById('mainDashboard');
                const workspace = document.getElementById('mainWorkspace');
                const omsaetning = document.getElementById('mainOmsaetning');
                const ordreindgang = document.getElementById('mainOrdreindgang');
                const belastning = document.getElementById('mainBelastning');

                if (moduleKey === 'efterkalk') {
                    if (!warmupCombinedReady) {
                        const msg = warmupCombinedTotal > 0
                            ? ('Efterkalk er ikke klar endnu (' + warmupCombinedDone + '/' + warmupCombinedTotal + ', ' + warmupCombinedPct + '%). Vent til warmup er færdig.')
                            : 'Efterkalk er ikke klar endnu. Vent et øjeblik til warmup/calculations er færdige.';
                        alert(msg);
                        return;
                    }
                    if (dashboard) dashboard.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'none';
                    if (ordreindgang) ordreindgang.style.display = 'none';
                    if (belastning) belastning.style.display = 'none';
                    if (workspace) workspace.style.display = 'block';
                    closeSideMenu();
                    goBackToList();
                    setTimeout(syncStickyOffsets, 0);
                    return;
                }

                if (moduleKey === 'omsaetning') {
                    if (dashboard) dashboard.style.display = 'none';
                    if (workspace) workspace.style.display = 'none';
                    if (ordreindgang) ordreindgang.style.display = 'none';
                    if (belastning) belastning.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'block';
                    closeSideMenu();
                    initializeOmsaetningIfNeeded();
                    setTimeout(syncStickyOffsets, 0);
                    return;
                }

                if (moduleKey === 'ordreindgang') {
                    if (dashboard) dashboard.style.display = 'none';
                    if (workspace) workspace.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'none';
                    if (belastning) belastning.style.display = 'none';
                    if (ordreindgang) ordreindgang.style.display = 'block';
                    closeSideMenu();
                    initializeOrdreindgangIfNeeded();
                    setTimeout(syncStickyOffsets, 0);
                    return;
                }

                if (moduleKey === 'belastning') {
                    if (dashboard) dashboard.style.display = 'none';
                    if (workspace) workspace.style.display = 'none';
                    if (omsaetning) omsaetning.style.display = 'none';
                    if (ordreindgang) ordreindgang.style.display = 'none';
                    if (belastning) belastning.style.display = 'block';
                    closeSideMenu();
                    initializeBelastningIfNeeded();
                    setTimeout(syncStickyOffsets, 0);
                    return;
                }

                if (moduleKey !== 'efterkalk') {
                    alert('Dette modul er klar til næste fase. Når du sender logikken, bygger vi det visuelt og funktionelt.');
                    return;
                }
            }

            function goToDashboard() {
                const dashboard = document.getElementById('mainDashboard');
                const workspace = document.getElementById('mainWorkspace');
                const omsaetning = document.getElementById('mainOmsaetning');
                const ordreindgang = document.getElementById('mainOrdreindgang');
                const belastning = document.getElementById('mainBelastning');
                closeSideMenu();
                if (workspace) workspace.style.display = 'none';
                if (omsaetning) omsaetning.style.display = 'none';
                if (ordreindgang) ordreindgang.style.display = 'none';
                if (belastning) belastning.style.display = 'none';
                if (dashboard) dashboard.style.display = 'block';
                const detailModal = document.getElementById('orderDetailModal');
                const detailBody = document.getElementById('orderDetailModalBody');
                if (detailModal) detailModal.style.display = 'none';
                if (detailBody) detailBody.innerHTML = '';
                document.body.classList.remove('report-modal-open');
                setTimeout(syncStickyOffsets, 0);
            }

            async function initializeOmsaetningIfNeeded() {
                if (omsaetningInitialized) return;
                omsaetningInitialized = true;

                const currentFiscalYear = getCurrentFiscalYearStart();
                omsaetningSelectedFiscalYears = new Set([currentFiscalYear]);
                applySelectedFiscalYearsToInputs();
                renderOmsaetningYearChips(currentFiscalYear);
                applyOmsaetningThresholdInputs(OMSAETNING_DEFAULT_WARN_THRESHOLD, OMSAETNING_DEFAULT_GOOD_THRESHOLD);
                renderOmsaetningCustomerMode();
                renderOmsaetningCustomerResults();

                await loadOmsaetningAccounts();
                await loadOmsaetningSummary();
            }

            async function loadOmsaetningAccounts() {
                const list = document.getElementById('omsaetningAccountsList');
                if (!list) return;
                list.innerHTML = '<div class="omsaetning-account-item"><span>Indlæser konti...</span></div>';
                try {
                    const response = await fetch('/omsaetning/accounts');
                    if (!response.ok) throw new Error('HTTP ' + response.status);
                    const payload = await response.json();
                    const accounts = Array.isArray(payload.accounts) ? payload.accounts : [];
                    omsaetningAccounts = accounts;
                    const allAccounts = accounts.map(acc => String(acc.acNo || '').trim()).filter(Boolean);
                    const ssrsPreset = allAccounts.filter(acNo => OMSAETNING_SSRS_DEFAULT_ACCOUNTS.has(acNo));
                    if (ssrsPreset.length === 0) {
                        omsaetningSelectedAccounts = new Set(allAccounts);
                    } else {
                        omsaetningSelectedAccounts = new Set(ssrsPreset);
                    }
                    renderOmsaetningAccountsList();
                } catch (err) {
                    list.innerHTML = '<div class="omsaetning-account-item"><span>Fejl ved konti</span></div>';
                    console.error('loadOmsaetningAccounts failed:', err);
                }
            }

            async function loadOmsaetningSummary(options) {
                const safeOptions = options && typeof options === 'object' ? options : {};
                const silentValidation = safeOptions.silentValidation === true;
                const loadBtn = document.getElementById('omsaetningLoadBtn');
                const empty = document.getElementById('omsaetningEmpty');
                const tableWrap = document.getElementById('omsaetningTableWrap');
                const detailsWrap = document.getElementById('omsaetningDetailsWrap');
                const thresholdWrap = document.getElementById('omsaetningThresholdWrap');
                const thresholdTable = document.getElementById('omsaetningThresholdTable');
                const chartsWrap = document.getElementById('omsaetningChartsWrap');
                const totalEl = document.getElementById('omsaetningTotalMio');
                const rowsEl = document.getElementById('omsaetningRowsCount');
                const periodsEl = document.getElementById('omsaetningPeriodsCount');

                const periodRange = buildOmsaetningPeriodRange();
                if (!periodRange) {
                    if (!silentValidation) {
                        alert('Vælg gyldig periode (Fra måned skal være før eller lig Til måned).');
                    }
                    return;
                }

                const fra = periodRange.fra;
                const til = periodRange.til;
                const monthKeysForPeriod = [];

                const selected = Array.from(omsaetningSelectedAccounts.values()).filter(Boolean);
                if (selected.length === 0) {
                    if (!silentValidation) {
                        alert('Vælg mindst én konto.');
                    }
                    return;
                }
                const selectedCustomers = Array.from(omsaetningSelectedCustomers.keys()).filter(Boolean);

                const thresholdInputs = getOmsaetningThresholdInputs();
                const warnThreshold = thresholdInputs.warnThreshold;
                const goodThreshold = thresholdInputs.goodThreshold;
                applyOmsaetningThresholdInputs(warnThreshold, goodThreshold);

                if (loadBtn) {
                    loadBtn.disabled = true;
                    loadBtn.textContent = 'Indlæser...';
                }

                try {
                    const payload = await fetchOmsaetningSummaryCached(fra, til, selected, selectedCustomers, safeOptions);

                    const persistTargets = await resolveOmsaetningThresholdPersistTargets(
                        selectedCustomers,
                        warnThreshold,
                        goodThreshold,
                        safeOptions
                    );

                    if (persistTargets.length > 0) {
                        persistOmsaetningThresholdsForCustomers(persistTargets, warnThreshold, goodThreshold)
                            .catch(err => console.warn('persistOmsaetningThresholdsForCustomers failed:', err && err.message ? err.message : err));

                        for (const custNo of persistTargets) {
                            omsaetningThresholdsByCustomer.set(custNo, {
                                warnThreshold,
                                goodThreshold
                            });
                        }
                        renderOmsaetningCustomerThresholds();
                    }

                    const rows = Array.isArray(payload.rows) ? payload.rows : [];
                    const uniquePeriods = new Set(monthKeysForPeriod.length > 0 ? monthKeysForPeriod : rows.map(r => normalizeOmsaetningMonthKey(r.date)));

                    const monthTotals = new Map();
                    for (const monthKey of monthKeysForPeriod) {
                        monthTotals.set(String(monthKey), 0);
                    }
                    for (const row of rows) {
                        const monthKey = normalizeOmsaetningMonthKey(row.date);
                        const prev = monthTotals.get(monthKey) || 0;
                        monthTotals.set(monthKey, prev + Number(row.revenueMio || 0));
                    }

                    if (totalEl) totalEl.textContent = formatMio(payload.totalRevenueMio || 0);
                    if (rowsEl) rowsEl.textContent = formatCount(rows.length);
                    if (periodsEl) periodsEl.textContent = formatCount(uniquePeriods.size);

                    const sortedMonths = (monthKeysForPeriod.length > 0
                        ? monthKeysForPeriod.map(monthKey => [monthKey, Number(monthTotals.get(String(monthKey)) || 0)])
                        : Array.from(monthTotals.entries()).sort((a, b) => String(a[0]).localeCompare(String(b[0]))));
                    let thresholdHtml = '<table class="omsaetning-table"><thead><tr>' +
                        '<th>Måned</th><th class="omsaetning-cell-right">Omsætning (Mio)</th><th>Tærskel</th>' +
                        '</tr></thead><tbody>';
                    for (const [monthKey, amountMio] of sortedMonths) {
                        const statusClass = getOmsaetningStatusClass(amountMio, warnThreshold, goodThreshold);
                        const monthMioLabel = formatMio(amountMio);
                        const monthDkkLabel = formatDkkFromMio(amountMio);
                        const gauge = buildOmsaetningGaugeData(amountMio, warnThreshold, goodThreshold);
                        const marginPctLabel = formatSigned(gauge.marginPct, 1) + '%';
                        const deltaWarnLabel = formatSigned(gauge.deltaWarn, 3) + ' Mio';
                        const deltaGoodLabel = formatSigned(gauge.deltaGood, 3) + ' Mio';
                        thresholdHtml += '<tr>' +
                            '<td>' + escapeHtmlFE(formatMonthDa(monthKey)) + '</td>' +
                            '<td class="omsaetning-cell-right" title="' + escapeHtmlFE(monthDkkLabel + ' DKK') + '">' + escapeHtmlFE(monthMioLabel) + '</td>' +
                            '<td>' +
                                '<div class="omsaetning-gauge-wrap">' +
                                    '<div class="omsaetning-gauge-meta">' +
                                        '<span class="omsaetning-status ' + statusClass + '">' + escapeHtmlFE(getOmsaetningStatusLabel(statusClass)) + '</span>' +
                                        '<strong>' + escapeHtmlFE(marginPctLabel) + '</strong>' +
                                    '</div>' +
                                    '<div class="omsaetning-gauge-track">' +
                                        '<div class="omsaetning-gauge-fill ' + gauge.fillClass + '" style="left:' + gauge.fillLeft.toFixed(2) + '%;width:' + gauge.fillWidth.toFixed(2) + '%;"></div>' +
                                        '<span class="omsaetning-gauge-marker" style="left:' + gauge.zeroLeft.toFixed(2) + '%;"></span>' +
                                        '<span class="omsaetning-gauge-marker" style="left:' + gauge.targetLeft.toFixed(2) + '%;"></span>' +
                                        '<span class="omsaetning-gauge-point" style="left:' + gauge.pointLeft.toFixed(2) + '%;"></span>' +
                                    '</div>' +
                                    '<div class="omsaetning-gauge-legend"><span>-30%</span><span>0% (3)</span><span>30% (5)</span><span>60%</span></div>' +
                                    '<div class="omsaetning-gauge-delta">vs 3: <strong>' + escapeHtmlFE(deltaWarnLabel) + '</strong> · vs 5: <strong>' + escapeHtmlFE(deltaGoodLabel) + '</strong></div>' +
                                '</div>' +
                            '</td>' +
                            '</tr>';
                    }
                    thresholdHtml += '</tbody></table>';

                    let html = '<table class="omsaetning-table"><thead><tr>' +
                        '<th>Måned</th><th>Konto</th><th>Navn</th><th>Kunde</th><th>Kundenavn</th><th style="text-align:right;">Omsætning (Mio)</th>' +
                        '</tr></thead><tbody>';

                    if (rows.length === 0) {
                        html += '<tr><td colspan="6" style="color:#5f7892;">Ingen bevægelser i perioden (0 for alle måneder).</td></tr>';
                    } else {
                        for (const row of rows) {
                            const rowMioLabel = formatMio(row.revenueMio || 0);
                            const rowDkkLabel = formatDkkFromMio(row.revenueMio || 0);
                            html += '<tr>' +
                                '<td>' + escapeHtmlFE(formatMonthDa(row.date)) + '</td>' +
                                '<td>' + escapeHtmlFE(String(row.acNo || '')) + '</td>' +
                                '<td>' + escapeHtmlFE(String(row.name || '')) + '</td>' +
                                '<td>' + escapeHtmlFE(row.custNo === null || row.custNo === undefined ? '' : String(row.custNo)) + '</td>' +
                                '<td>' + escapeHtmlFE(String(row.customerName || '')) + '</td>' +
                                '<td style="text-align:right;" title="' + escapeHtmlFE(rowDkkLabel + ' DKK') + '">' + escapeHtmlFE(rowMioLabel) + '</td>' +
                                '</tr>';
                        }
                    }

                    html += '</tbody></table>';
                    if (OMSAETNING_SHOW_THRESHOLD_SECTION) {
                        if (thresholdTable) thresholdTable.innerHTML = thresholdHtml;
                        if (thresholdWrap) thresholdWrap.style.display = 'block';
                    } else {
                        if (thresholdTable) thresholdTable.innerHTML = '';
                        if (thresholdWrap) thresholdWrap.style.display = 'none';
                    }
                    if (tableWrap) {
                        tableWrap.innerHTML = html;
                    }
                    if (detailsWrap) {
                        detailsWrap.style.display = 'block';
                        applyOmsaetningDetailsCollapsedState();
                    }
                    renderOmsaetningCharts(rows, monthKeysForPeriod);
                    if (empty) empty.style.display = 'none';
                } catch (err) {
                    if (tableWrap) {
                        tableWrap.style.display = 'none';
                        tableWrap.innerHTML = '';
                    }
                    if (detailsWrap) detailsWrap.style.display = 'none';
                    if (thresholdWrap) thresholdWrap.style.display = 'none';
                    if (thresholdTable) thresholdTable.innerHTML = '';
                    if (chartsWrap) chartsWrap.style.display = 'none';
                    if (empty) {
                        empty.style.display = 'block';
                        empty.textContent = 'Fejl ved indlæsning: ' + (err && err.message ? err.message : 'ukendt fejl');
                    }
                    console.error('loadOmsaetningSummary failed:', err);
                } finally {
                    if (loadBtn) {
                        loadBtn.disabled = false;
                        loadBtn.textContent = 'Opdater';
                    }
                }
            }

            async function checkOrderListFreshness() {
                const now = Date.now();
                if (now - lastOrderListCheckTime < 30000) return;
                lastOrderListCheckTime = now;

                try {
                    const r = await fetch('/order-list-check-time');
                    if (!r.ok) return;
                    const d = await r.json();
                    const remoteMaxDate = Number(d.lastModifiedDate || 0);
                    
                    if (remoteMaxDate > 0 && remoteMaxDate !== lastOrderListRemoteTime) {
                        console.info('ORDER-LIST: Database has new/changed order (date=' + remoteMaxDate + ')');
                        lastOrderListRemoteTime = remoteMaxDate;
                        await loadOrderList(true);
                    }
                } catch (err) {
                    console.warn('checkOrderListFreshness failed:', err.message);
                }
            }

            function registerSummaryImageData(title, items) {
                if (!Array.isArray(items) || items.length === 0) return '';
                summaryImageRegistryCounter += 1;
                const key = 'img-' + summaryImageRegistryCounter;
                summaryImageRegistry[key] = {
                    title: title || 'Billeder',
                    items: items
                };
                return key;
            }

            function getSummaryImageSrc(item) {
                if (!item) return '';
                if (item.type === 'url') return item.value;
                return '/image-file?path=' + encodeURIComponent(item.value || '');
            }

            function selectCompactImageItems(items) {
                const source = Array.isArray(items) ? items : [];
                const clean = source.filter(item => item && String(item.value || '').trim());
                if (clean.length <= 2) return clean;

                const pickByLabel = (pattern) => clean.find(item => pattern.test(String(item.label || '')));
                const picked = [];
                const pushUnique = (item) => {
                    if (!item) return;
                    const value = String(item.value || '').trim().toLowerCase();
                    if (!value) return;
                    if (picked.some(x => String(x.value || '').trim().toLowerCase() === value)) return;
                    picked.push(item);
                };

                pushUnique(pickByLabel(/webpg|nesting/i));
                pushUnique(pickByLabel(/pictfnm|icon/i));
                for (const item of clean) {
                    pushUnique(item);
                    if (picked.length >= 2) break;
                }
                return picked.slice(0, 2);
            }

            function openCompactImageModal(imageKey) {
                const entry = summaryImageRegistry[imageKey];
                const modal = document.getElementById('compactImageModal');
                const titleEl = document.getElementById('compactImageTitle');
                const subtitleEl = document.getElementById('compactImageSubtitle');
                const bodyEl = document.getElementById('compactImageBody');
                if (!entry || !modal || !titleEl || !bodyEl) return;

                const items = selectCompactImageItems(entry.items);
                if (!items.length) return;

                titleEl.textContent = entry.title || 'Billeder';
                subtitleEl.textContent = 'Viser nesting + hovedbilleder';

                let html = '<div class="compact-image-grid">';
                for (const item of items) {
                    const src = getSummaryImageSrc(item);
                    html += '<div class="compact-image-card">';
                    html += '<div class="compact-image-label">' + escapeHtml(item.label || 'Billede') + '</div>';
                    html += '<img class="image-preview-zoomable" src="' + escapeHtml(src) + '" alt="' + escapeHtml(item.label || entry.title || 'Billede') + '" loading="lazy" data-fullsrc="' + escapeHtml(src) + '" data-title="' + escapeHtml(item.label || entry.title || 'Billede') + '" data-path="' + escapeHtml(item.value || '') + '" />';
                    html += '<div class="compact-image-path">' + escapeHtml(item.value || '') + '</div>';
                    html += '</div>';
                }
                html += '</div>';
                bodyEl.innerHTML = html;
                modal.classList.add('show');
            }

            function closeCompactImageModal(event) {
                if (event && event.target && event.target.id !== 'compactImageModal') return;
                const modal = document.getElementById('compactImageModal');
                const bodyEl = document.getElementById('compactImageBody');
                if (!modal) return;
                modal.classList.remove('show');
                if (bodyEl) bodyEl.innerHTML = '';
            }

            function closeSummaryImagePanel() {
                const panels = [
                    document.getElementById('summaryImagePanel'),
                    document.getElementById('laserImagePanel')
                ];
                for (const panel of panels) {
                    if (!panel) continue;
                    panel.innerHTML = '';
                    panel.classList.add('hidden');
                }
                updateSummaryImagePanelLayout();
            }

            function updateSummaryImagePanelLayout() {
                const wrap = document.querySelector('#summaryModal .modal-content-wrap');
                const summaryPanel = document.getElementById('summaryImagePanel');
                if (!wrap || !summaryPanel) return;
                const hasImages = !summaryPanel.classList.contains('hidden') && summaryPanel.innerHTML.trim() !== '';
                const shouldFocus = hasImages && window.matchMedia('(max-width: 1440px)').matches;
                wrap.classList.toggle('image-focus', shouldFocus);
            }

            function openSummaryImagePanel(imageKey, preferredPanelId) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const laserPanelWrap = document.getElementById('laserOrderSummaryPanel');
                const laserPanel = document.getElementById('laserImagePanel');
                const summaryPanel = document.getElementById('summaryImagePanel');
                const isLaserVisible = laserPanelWrap && laserPanelWrap.style.display !== 'none';
                const isVisible = (el) => {
                    if (!el) return false;
                    const s = getComputedStyle(el);
                    return s.display !== 'none' && s.visibility !== 'hidden' && el.getClientRects().length > 0;
                };
                let panel = null;
                if (preferredPanelId === 'laserImagePanel') {
                    panel = isVisible(laserPanel) ? laserPanel : summaryPanel;
                } else if (preferredPanelId === 'summaryImagePanel') {
                    panel = summaryPanel;
                } else {
                    panel = (isLaserVisible && laserPanel) ? laserPanel : summaryPanel;
                }
                const entry = summaryImageRegistry[imageKey];
                if (!panel || !entry || !Array.isArray(entry.items) || entry.items.length === 0) {
                    closeSummaryImagePanel();
                    return;
                }

                if (panel.id === 'summaryImagePanel' && title) {
                    title.textContent = entry.title || 'Billeder';
                }
                if (panel.id === 'summaryImagePanel' && modal && modal.style.display !== 'flex') {
                    modal.style.display = 'flex';
                }

                let html = '<div class="summary-image-panel-header">';
                html += '<div class="summary-image-panel-title">' + escapeHtml(entry.title) + '</div>';
                html += '<button class="summary-image-close" onclick="closeSummaryImagePanel()">Luk</button>';
                html += '</div>';
                html += '<div class="image-preview-gallery">';

                for (const item of entry.items) {
                    const src = getSummaryImageSrc(item);
                    html += '<div class="image-preview-card">';
                    html += '<div class="image-preview-label">' + escapeHtml(item.label || 'Billede') + '</div>';
                    html += '<img class="image-preview-zoomable" src="' + escapeHtml(src) + '" alt="' + escapeHtml(entry.title) + '" loading="lazy" data-fullsrc="' + escapeHtml(src) + '" data-title="' + escapeHtml(item.label || entry.title || 'Billede') + '" data-path="' + escapeHtml(item.value || '') + '" />';
                    html += '<div class="image-preview-path">' + escapeHtml(item.value || '') + '</div>';
                    html += '</div>';
                }

                html += '</div>';
                panel.innerHTML = html;
                panel.classList.remove('hidden');
                updateSummaryImagePanelLayout();
            }

            function openImageLightbox(src, title, pathText) {
                const lightbox = document.getElementById('imageLightbox');
                const img = document.getElementById('imageLightboxImg');
                const titleEl = document.getElementById('imageLightboxTitle');
                const pathEl = document.getElementById('imageLightboxPath');
                if (!lightbox || !img) return;

                img.src = src || '';
                img.alt = title || 'Billede';
                if (titleEl) titleEl.textContent = title || 'Billede';
                if (pathEl) pathEl.textContent = pathText || '';
                lightbox.classList.remove('hidden');
            }

            function closeImageLightbox(event) {
                if (event && event.target && event.target.id !== 'imageLightbox') return;
                const lightbox = document.getElementById('imageLightbox');
                const img = document.getElementById('imageLightboxImg');
                const pathEl = document.getElementById('imageLightboxPath');
                if (!lightbox || lightbox.classList.contains('hidden')) return;

                lightbox.classList.add('hidden');
                if (img) {
                    img.src = '';
                    img.alt = '';
                }
                if (pathEl) pathEl.textContent = '';
            }

            function updateSummaryModalBackBtn() {
                const backBtn = document.getElementById('summaryModalBackBtn');
                if (!backBtn) return;
                backBtn.classList.toggle('hidden', summaryModalHistory.length === 0);
            }

            function pushSummaryModalState() {
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');
                const imagePanel = document.getElementById('summaryImagePanel');
                if (!title || !body) return;
                summaryModalHistory.push({
                    title: title.textContent,
                    bodyHtml: body.innerHTML,
                    imageHtml: imagePanel ? imagePanel.innerHTML : '',
                    imageHidden: imagePanel ? imagePanel.classList.contains('hidden') : true
                });
                updateSummaryModalBackBtn();
            }

            function goSummaryModalBack() {
                if (summaryModalHistory.length === 0) return;
                const prev = summaryModalHistory.pop();
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');
                const imagePanel = document.getElementById('summaryImagePanel');
                if (title) title.textContent = prev.title;
                if (body) body.innerHTML = prev.bodyHtml;
                if (imagePanel) {
                    imagePanel.innerHTML = prev.imageHtml || '';
                    imagePanel.classList.toggle('hidden', prev.imageHidden !== false);
                }
                updateSummaryImagePanelLayout();
                updateSummaryModalBackBtn();
            }

            function setSystemStatus(text, bgColor, textColor) {
                // Header right area now shows user greeting instead of system status.
            }

            // Warmup progress bar polling
            let warmupPollTimer = null;
            let warmupTopBarHideScheduled = false;
            let warmupCombinedReady = false;
            let warmupCombinedPct = 0;
            let warmupCombinedDone = 0;
            let warmupCombinedTotal = 0;
            let showDashboardWarmupNotice = false;
            function startWarmupPolling() {
                if (warmupPollTimer) return;
                const wrap = document.getElementById('warmupBarWrap');
                const fill = document.getElementById('warmupBarFill');
                const txt  = document.getElementById('warmupBarText');
                const dashText = document.getElementById('dashboardWarmupText');
                const dashMeta = document.getElementById('dashboardWarmupMeta');
                const dashFill = document.getElementById('dashboardWarmupFill');
                const dashPct = document.getElementById('dashboardWarmupPct');
                const dashWrap = document.getElementById('dashboardWarmupNotice');
                if (!wrap && !dashText) return;

                warmupPollTimer = setInterval(async () => {
                    try {
                        const r = await fetch('/warmup-status');
                        if (!r.ok) return;
                        const d = await r.json();

                        const totalCombined = Number(d.combinedTotal || d.total || 0);
                        const doneCombined = Number(d.combinedDone || d.done || 0);
                        const pctCombined = Number(d.combinedPct || d.pct || 0);
                        const readyCombined = d.ready === true;

                        warmupCombinedPct = Math.max(0, Math.min(100, pctCombined));
                        warmupCombinedDone = Math.max(0, doneCombined);
                        warmupCombinedTotal = Math.max(0, totalCombined);
                        warmupCombinedReady = readyCombined || (!d.running && warmupCombinedTotal === 0);

                        if (dashFill) dashFill.style.width = String(Math.max(0, Math.min(100, pctCombined))) + '%';
                        if (dashPct) dashPct.textContent = String(Math.max(0, Math.min(100, pctCombined))) + '%';

                        // Keep warmup hidden on initial dashboard screen, unless user explicitly triggered cache reset.
                        if (dashWrap) {
                            const shouldShowDashWarmup = d.running || (!readyCombined && totalCombined > 0) || showDashboardWarmupNotice;
                            dashWrap.classList.toggle('hidden', !shouldShowDashWarmup);
                        }

                        if (dashText) {
                            if (d.running) {
                                dashText.textContent = 'Forbereder ' + doneCombined + '/' + totalCombined + ' ordredata...';
                                if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                            } else if (readyCombined && totalCombined > 0) {
                                dashText.textContent = 'Klar! Efterkalk-data er forberedt.';
                                if (dashMeta) dashMeta.textContent = 'Åbn Efterkalk når som helst.';
                                if (dashWrap) {
                                    setTimeout(() => {
                                        dashWrap.classList.add('hidden');
                                        showDashboardWarmupNotice = false;
                                    }, 1800);
                                }
                            } else if (totalCombined > 0) {
                                dashText.textContent = 'Afventer baggrundsjob...';
                                if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                            } else {
                                dashText.textContent = 'Venter på warmup-status...';
                                if (dashMeta) dashMeta.textContent = 'Du kan bruge andre moduler imens.';
                            }
                        }

                        if (d.total === 0) {
                            if (wrap) wrap.classList.remove('active');
                            return;
                        }

                        if (wrap) wrap.classList.add('active');
                        if (fill) fill.style.width = d.pct + '%';

                        if (d.running) {
                            if (txt) txt.textContent = 'Forberegner ' + d.done + '/' + d.total + ' ordrer...';
                            warmupTopBarHideScheduled = false;
                        } else {
                            if (txt) txt.textContent = 'Klar! ' + d.loaded + ' nye + ' + d.cached + ' fra cache';
                            if (fill) fill.style.width = '100%';
                            if (!warmupTopBarHideScheduled) {
                                warmupTopBarHideScheduled = true;
                                setTimeout(() => {
                                    if (wrap) wrap.classList.remove('active');
                                    warmupTopBarHideScheduled = false;
                                }, 3000);
                            }
                        }
                    } catch(e) {
                        // ignore polling errors silently
                    }
                }, 800);
            }
            function updateSystemStatusFromOrders(orders) {
                if (!orders || orders.length === 0) {
                    setSystemStatus('System klar', '#e8f5e9', '#1b5e20');
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
                    setSystemStatus('System klar', '#e8f5e9', '#1b5e20');
                    return;
                }

                setSystemStatus('System indlæser... ' + completed + '/' + total, '#fff3cd', '#8a6d3b');
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

            function getOrderInvoiceStatusHtml(ordNo) {
                const marginState = getMarginState(ordNo);
                if (!marginState || marginState.status !== 'success') return '<span style="color:#999;">-</span>';
                if (marginState.hasInvoiceWarning) {
                    return '<span title="En eller flere linjer mangler faktura (NoInvo=0); kostberegning bruger NoFin som fallback." style="display:inline-flex;align-items:center;gap:4px;background:#fff3e0;color:#e65100;font-size:12px;font-weight:600;padding:2px 7px;border-radius:10px;border:1px solid #ffcc80;">🧾 Mangler</span>';
                }
                return '<span style="color:#388e3c;font-size:13px;" title="Alle fakturaer registreret.">✓</span>';
            }

            function updateOrderInvoiceCell(ordNo) {
                const listEl = document.getElementById('orderList');
                if (!listEl) return;
                const cells = listEl.querySelectorAll('.order-invoice-cell[data-ordno="' + ordNo + '"]');
                const html = getOrderInvoiceStatusHtml(ordNo);
                for (const cell of cells) { cell.innerHTML = html; }
            }

            function getOrderMarginHtml(ordNo) {
                const marginState = getMarginState(ordNo);
                let marginHtml = '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>';
                let toneClass = 'na';
                if (marginState && marginState.status === 'loading') {
                    marginHtml = '<span style="background:#546e7a; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">...</span>';
                    toneClass = 'na';
                } else if (marginState && marginState.status === 'success') {
                    const marginValue = calculateOrderMarginPercent(marginState.totalRevenue || 0, marginState.totalCost || 0);
                    const margin = marginValue.toFixed(2);
                    marginHtml = getMarginBadge(margin);
                    if (currentMarginMode === 'new') {
                        toneClass = marginValue >= 125 ? 'ok' : (marginValue >= 105 ? 'warn' : 'bad');
                    } else {
                        toneClass = marginValue > 20 ? 'ok' : (marginValue >= 5 ? 'warn' : 'bad');
                    }
                }
                return '<span class="order-margin-wrap"><span class="order-kpi-tone ' + toneClass + '" title="KPI tone fra Rapport 2.0">↗</span>' + marginHtml + '</span>';
            }

            function updateOrderMarginCell(ordNo) {
                const listEl = document.getElementById('orderList');
                if (!listEl) return;
                const cells = listEl.querySelectorAll('.order-margin-cell[data-ordno="' + ordNo + '"]');
                if (!cells || cells.length === 0) return;
                const marginHtml = getOrderMarginHtml(ordNo);
                for (const cell of cells) {
                    cell.innerHTML = marginHtml;
                }
                updateOrderListSummaryPanel();
            }

            function refreshOrderListStatus() {
                if (!orderListVisible) return;
                const visibleOrders = getFilteredOrders().slice(0, MARGIN_PREFETCH_ROWS);
                updateSystemStatusFromOrders(visibleOrders);
            }

            function scheduleMarginSortRefresh() {
                if (orderListSortField !== 'margin') return;
                if (marginSortRefreshTimer) return;
                marginSortRefreshTimer = setTimeout(() => {
                    marginSortRefreshTimer = null;
                    renderOrderList();
                }, 350);
            }

            function hydrateMarginStateFromOrderList(orders) {
                marginStateByOrdNo = {};
                for (const o of orders) {
                    const ordNo = Number(o.OrdNo);
                    if (!Number.isFinite(ordNo)) continue;

                    if (o.TotalCost !== null && o.TotalCost !== undefined) {
                        marginStateByOrdNo[String(ordNo)] = {
                            status: 'success',
                            totalRevenue: Number(o.InvoAm || 0),
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
                    updateOrderMarginCell(ordNo);
                }
                pumpMarginQueue();
                refreshOrderListStatus();
            }

            async function loadSingleOrderMargin(ordNo) {
                const key = String(ordNo);
                const controller = new AbortController();
                const timeoutId = setTimeout(() => controller.abort(), MARGIN_FETCH_TIMEOUT_MS);
                try {
                    const response = await fetch('/order-margin/' + ordNo, { signal: controller.signal });
                    let data = null;
                    try {
                        data = await response.json();
                    } catch {
                        data = { error: 'Invalid JSON response' };
                    }
                    if (!response.ok || data.error) {
                        marginStateByOrdNo[key] = { status: 'error' };
                        updateOrderMarginCell(ordNo);
                        refreshOrderListStatus();
                        scheduleMarginSortRefresh();
                        return;
                    }

                    marginStateByOrdNo[key] = {
                        status: 'success',
                        totalRevenue: Number(data.totalRevenue || 0),
                        totalCost: Number(data.totalCost || 0),
                        hasInvoiceWarning: Boolean(data.hasInvoiceWarning)
                    };
                    updateOrderMarginCell(ordNo);
                    updateOrderInvoiceCell(ordNo);
                    refreshOrderListStatus();
                    scheduleMarginSortRefresh();
                } catch (err) {
                    marginStateByOrdNo[key] = { status: 'error' };
                    updateOrderMarginCell(ordNo);
                    refreshOrderListStatus();
                    scheduleMarginSortRefresh();
                } finally {
                    clearTimeout(timeoutId);
                }
            }

            function pumpMarginQueue() {
                while (marginWorkerActiveCount < MARGIN_MAX_CONCURRENT && marginJobQueue.length > 0) {
                    const ordNo = marginJobQueue.shift();
                    marginWorkerActiveCount += 1;

                    loadSingleOrderMargin(ordNo)
                        .finally(() => {
                            marginWorkerActiveCount -= 1;
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

            function setOrderValueFilter() {
                const enabledInput = document.getElementById('orderMinDkkEnabled');
                const thresholdInput = document.getElementById('orderMinDkkInput');
                orderListMinDkkEnabled = !!(enabledInput && enabledInput.checked);

                const raw = Number(thresholdInput && thresholdInput.value || 0);
                orderListMinDkkValue = Number.isFinite(raw) ? Math.max(0, Math.round(raw)) : 0;

                if (thresholdInput) {
                    thresholdInput.disabled = !orderListMinDkkEnabled;
                    if (!thresholdInput.disabled && String(thresholdInput.value || '') !== String(orderListMinDkkValue)) {
                        thresholdInput.value = String(orderListMinDkkValue);
                    }
                }

                if (!orderListVisible && orderListMinDkkEnabled && orderListMinDkkValue > 0) {
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

            function setOrderListSort(field) {
                if (orderListSortField === field) {
                    orderListSortDir = orderListSortDir === 'asc' ? 'desc' : 'asc';
                } else {
                    orderListSortField = field;
                    orderListSortDir = field === 'date' || field === 'ordno' || field === 'belob' || field === 'margin' ? 'desc' : 'asc';
                }
                renderOrderList();
            }

            function getMarginValue(ordNo) {
                const state = getMarginState(ordNo);
                if (!state || state.status !== 'success') return null;
                return calculateOrderMarginPercent(state.totalRevenue || 0, state.totalCost || 0);
            }

            function getFilteredOrders() {
                const filtered = orderListData.filter(o => {
                    const bruger = String(o.SellerUsr || '').trim();
                    const customer = String(o.CustomerName || '').toLowerCase();
                    const ord = String(o.OrdNo || '');
                    const matchesText = !orderListFilter || customer.includes(orderListFilter) || ord.includes(orderListFilter);
                    const matchesBruger = !orderListBrugerFilter || bruger === orderListBrugerFilter;
                    const invoDkk = Number(o.InvoAm || 0);
                    const matchesMinDkk = !orderListMinDkkEnabled || invoDkk >= orderListMinDkkValue;
                    return matchesText && matchesBruger && matchesMinDkk;
                });

                const dir = orderListSortDir === 'asc' ? 1 : -1;
                filtered.sort((a, b) => {
                    switch (orderListSortField) {
                        case 'bruger': {
                            const cmp = String(a.SellerUsr || '').localeCompare(String(b.SellerUsr || ''));
                            return cmp * dir || Number(b.LstInvDt || 0) - Number(a.LstInvDt || 0);
                        }
                        case 'ordno':
                            return (Number(a.OrdNo || 0) - Number(b.OrdNo || 0)) * dir;
                        case 'kunde': {
                            const cmp = String(a.CustomerName || '').localeCompare(String(b.CustomerName || ''));
                            return cmp * dir || Number(b.OrdNo || 0) - Number(a.OrdNo || 0);
                        }
                        case 'date': {
                            const d = (Number(a.LstInvDt || 0) - Number(b.LstInvDt || 0)) * dir;
                            return d || (Number(b.OrdNo || 0) - Number(a.OrdNo || 0)) * dir;
                        }
                        case 'belob':
                            return (Number(a.InvoAm || 0) - Number(b.InvoAm || 0)) * dir;
                        case 'margin': {
                            const ma = getMarginValue(a.OrdNo);
                            const mb = getMarginValue(b.OrdNo);
                            if (ma === null && mb === null) return 0;
                            if (ma === null) return 1;
                            if (mb === null) return -1;
                            return (ma - mb) * dir;
                        }
                        default:
                            return Number(b.LstInvDt || 0) - Number(a.LstInvDt || 0);
                    }
                });
                return filtered;
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

            function getFilteredOrderSummary(orders) {
                const safeOrders = Array.isArray(orders) ? orders : [];
                let considered = 0;
                let excludedCredit = 0;
                let pendingMargin = 0;
                let totalRevenue = 0;
                let totalCost = 0;

                for (const o of safeOrders) {
                    if (isOrderMarkedCreditNote(o.OrdNo)) {
                        excludedCredit += 1;
                        continue;
                    }
                    const state = getMarginState(o.OrdNo);
                    if (!state || state.status !== 'success') {
                        pendingMargin += 1;
                        continue;
                    }
                    considered += 1;
                    totalRevenue += Number(state.totalRevenue || 0);
                    totalCost += Number(state.totalCost || 0);
                }

                const marginAmount = totalRevenue - totalCost;
                const marginPct = totalCost > 0 ? calculateOrderMarginPercent(totalRevenue, totalCost).toFixed(2) : '0.00';
                return {
                    considered,
                    excludedCredit,
                    pendingMargin,
                    totalRevenue,
                    totalCost,
                    marginAmount,
                    marginPct
                };
            }

            function buildOrderListSummaryHtml(orders) {
                const listSummary = getFilteredOrderSummary(orders);
                const activeFilters = [];
                if (orderListFilter) activeFilters.push('kunde/søgning: "' + escapeHtml(orderListFilter) + '"');
                if (orderListBrugerFilter) activeFilters.push('bruger: "' + escapeHtml(orderListBrugerFilter) + '"');
                if (orderListMinDkkEnabled) activeFilters.push('minimum fakturabeløb: ' + formatNumber(orderListMinDkkValue) + ' DKK');
                const filterText = activeFilters.length > 0 ? activeFilters.join(', ') : 'ingen aktive filtre';
                let html = '<div><strong>Filtreret ordrelisteoversigt</strong> (vist: ' + orders.length + ', medtaget: ' + listSummary.considered + ', kreditnota udelukket: ' + listSummary.excludedCredit + ', mangler margin: ' + listSummary.pendingMargin + ')</div>';
                html += '<div style="margin-top:4px; font-size:12px; color:#57718f;">Genereret: ' + escapeHtml(new Date().toLocaleString('da-DK')) + ' • Filtre: ' + filterText + '</div>';
                html += '<div style="margin-top:6px;display:flex;gap:18px;flex-wrap:wrap;">';
                html += '<span>Samlet omsætning: <strong>' + formatNumber(listSummary.totalRevenue) + ' DKK</strong></span>';
                html += '<span>Samlet kost: <strong>' + formatNumber(listSummary.totalCost) + ' DKK</strong></span>';
                html += '<span>Margin: <strong>' + formatNumber(listSummary.marginAmount) + ' DKK (' + listSummary.marginPct + '%)</strong></span>';
                html += '</div>';
                html += '<div class="order-list-summary-actions">';
                html += '<button class="list-toggle-btn" onclick="openOrderListPrintPreview()" title="Vis forhåndsvisning af den filtrerede ordreliste">Forhåndsvisning / PDF</button>';
                html += '</div>';
                return html;
            }

            function buildOrderDetailReportHtml(orderData, orderMarginPercent, costToDateFromProduction) {
                const orderHeader = (orderData && orderData.orderHeader) || {};
                const productionOrders = Array.isArray(orderData && orderData.productionOrders) ? orderData.productionOrders : [];
                const salesOrderLines = Array.isArray(orderData && orderData.salesOrderLines) ? orderData.salesOrderLines : [];
                const salesLines = Array.isArray(orderData && orderData.salesLines) ? orderData.salesLines : [];

                let totalPlannedMinutes = 0;
                let totalUsedMinutes = 0;
                let totalOperationCost = 0;
                let totalLaserCost = 0;
                const rows = [];
                const exceptionMap = new Map();

                const pushException = (type, prodOrdNo, ref, message) => {
                    const text = String(message || '').trim();
                    if (!text) return;
                    const normalizedType = String(type || '-').trim() || '-';
                    const normalizedProdOrdNo = String(prodOrdNo || '-').trim() || '-';
                    const normalizedRef = String(ref || '-').trim() || '-';
                    const key = normalizedType + '|' + normalizedProdOrdNo + '|' + normalizedRef + '|' + text;
                    const existing = exceptionMap.get(key);
                    if (existing) {
                        existing.count += 1;
                        return;
                    }
                    exceptionMap.set(key, {
                        type: normalizedType,
                        prodOrdNo: normalizedProdOrdNo,
                        ref: normalizedRef,
                        message: text,
                        count: 1
                    });
                };

                for (const prodOrder of productionOrders) {
                    const lines = Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : [];
                    let plannedMinutes = 0;
                    let usedMinutes = 0;
                    let operationCost = 0;
                    let laserCost = 0;
                    let materialCost = 0;
                    let productLabel = '-';

                    for (const line of lines) {
                        const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                        const lnNo = Number((line && line.LnNo) || 0);
                        if (lnNo === 1 && line && line.ProdNo) {
                            productLabel = String(line.ProdNo || '-') + ' - ' + String(line.Descr || '');
                        }
                        if (lnNo === 1 || key === '0' || key === '3' || key === '5') continue;

                        const totalCost = Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                        if (key === '1') {
                            plannedMinutes += Number((line && line.NoOrg) || 0);
                            usedMinutes += Number((line && (line.EffectiveOperationMinutes ?? line.NoFin)) || 0);
                            operationCost += totalCost;
                        } else if (key === '2') {
                            laserCost += totalCost;
                            if (!isLaserLProdNo(line && line.ProdNo)) {
                                materialCost += totalCost;
                            }
                        }

                        if (line && line.HasWarning && line.WarningText) {
                            pushException('Advarsel', prodOrder.ordNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.WarningText);
                        }
                        if (line && line.IsInvoiceTracked && line.UsesMissingInvoiceFallback) {
                            pushException('Faktura', prodOrder.ordNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.MissingInvoiceText || 'Mangler faktura');
                        }
                        if (line && line.UsesEstimatedOperationTime) {
                            pushException('Tid', prodOrder.ordNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.EstimatedTimeText || 'Færdigmeldt minutter var 0 og blev estimeret');
                        }
                    }

                    totalPlannedMinutes += plannedMinutes;
                    totalUsedMinutes += usedMinutes;
                    totalOperationCost += operationCost;
                    totalLaserCost += laserCost;
                    rows.push({
                        ordNo: Number(prodOrder && prodOrder.ordNo) || 0,
                        productLabel,
                        plannedMinutes,
                        usedMinutes,
                        deltaMinutes: usedMinutes - plannedMinutes,
                        operationCost,
                        laserCost,
                        materialCost,
                        totalCost: Number(prodOrder && prodOrder.totalCost) || 0
                    });
                }

                for (const line of salesOrderLines) {
                    if (line && line.HasWarning && line.WarningText) {
                        pushException('Salgsordre', line.PurcNo || orderHeader.OrdNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.WarningText);
                    }
                    if (line && line.IsInvoiceTracked && line.UsesMissingInvoiceFallback) {
                        pushException('Faktura', line.PurcNo || orderHeader.OrdNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.MissingInvoiceText || 'Mangler faktura');
                    }
                    if (line && line.UsesEstimatedOperationTime) {
                        pushException('Tid', line.PurcNo || orderHeader.OrdNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.EstimatedTimeText || 'Færdigmeldt minutter var 0 og blev estimeret');
                    }
                }

                for (const line of salesLines) {
                    if (line && line.HasWarning && line.WarningText) {
                        pushException('Ekstra linje', line.PurcNo || orderHeader.OrdNo, String(line.ProdNo || ('L' + (line.LnNo || '-'))), line.WarningText);
                    }
                }

                const exceptionRows = Array.from(exceptionMap.values())
                    .sort((a, b) => {
                        const aProd = Number(a.prodOrdNo) || 0;
                        const bProd = Number(b.prodOrdNo) || 0;
                        if (aProd !== bProd) return aProd - bProd;
                        if (a.type !== b.type) return a.type.localeCompare(b.type, 'da');
                        if (a.ref !== b.ref) return a.ref.localeCompare(b.ref, 'da');
                        return a.message.localeCompare(b.message, 'da');
                    });

                const exceptionCompactMap = new Map();
                for (const ex of exceptionRows) {
                    const prodOrdNo = String(ex.prodOrdNo || '-');
                    const message = String(ex.message || '').trim() || '-';
                    const compactKey = prodOrdNo + '|' + message;
                    if (!exceptionCompactMap.has(compactKey)) {
                        exceptionCompactMap.set(compactKey, {
                            prodOrdNo,
                            message,
                            typeSet: new Set(),
                            refSet: new Set(),
                            count: 0
                        });
                    }
                    const row = exceptionCompactMap.get(compactKey);
                    row.typeSet.add(String(ex.type || '-'));
                    row.refSet.add(String(ex.ref || '-'));
                    row.count += Number(ex.count || 0);
                }

                const exceptionCompactRows = Array.from(exceptionCompactMap.values())
                    .map(row => ({
                        prodOrdNo: row.prodOrdNo,
                        types: Array.from(row.typeSet.values()).sort((a, b) => a.localeCompare(b, 'da')),
                        refs: Array.from(row.refSet.values()).sort((a, b) => a.localeCompare(b, 'da')),
                        message: row.message,
                        count: row.count
                    }))
                    .sort((a, b) => {
                        const aProd = Number(a.prodOrdNo) || 0;
                        const bProd = Number(b.prodOrdNo) || 0;
                        if (aProd !== bProd) return aProd - bProd;
                        if (b.count !== a.count) return b.count - a.count;
                        return a.message.localeCompare(b.message, 'da');
                    });

                const exceptionOccurrenceCount = exceptionRows.reduce((sum, item) => sum + Number(item.count || 0), 0);
                const exceptionGroupCount = exceptionCompactRows.length;

                const marginAmount = Number((orderData && orderData.summary && orderData.summary.margin) || 0);
                const marginPct = orderMarginPercent || '0.00';
                const revenue = Number((orderData && orderData.summary && orderData.summary.totalRevenue) || 0);
                const cost = Number((orderData && orderData.summary && orderData.summary.totalCost) || 0);
                const generatedAt = new Date().toLocaleString('da-DK');
                const orderTypeLabel = Number(orderHeader.Gr4 || 0) === 3 ? 'Multiordre' : 'Ordre';
                const statusLabel = Number(orderHeader.InvoAm || 0) === 0
                    ? 'I produktion'
                    : (Number(orderHeader.DInvoIF || 0) <= 0 ? 'Komplet faktureret' : 'Delvist faktureret');

                let html = '<div id="orderDetailReport" class="order-detail-report">';
                html += '<div class="order-report-toolbar">';
                html += '<div class="order-report-meta"><strong>' + escapeHtml(orderTypeLabel) + ' ' + escapeHtml(String(orderHeader.OrdNo || '-')) + '</strong><div class="report-subline">' + escapeHtml(String(orderHeader.CustomerName || '-')) + '</div><div class="report-badges"><span class="report-badge">Status: <strong>' + escapeHtml(statusLabel) + '</strong></span><span class="report-badge">Margin: <strong>' + escapeHtml(marginPct) + '%</strong></span></div></div>';
                html += '</div>';
                html += '<div class="report-hero">';
                html += '<div class="report-hero-top">';
                html += '<div class="report-hero-title">';
                html += '<div class="eyebrow">Ledelsesoverblik</div>';
                html += '<h1>Rapport for ordre ' + escapeHtml(String(orderHeader.OrdNo || '-')) + '</h1>';
                html += '<div class="context">' + escapeHtml(String(orderHeader.CustomerName || '-')) + ' • Genereret ' + escapeHtml(generatedAt) + '</div>';
                html += '<div class="report-arrow">Forbedringsretning i grøn KPI-tone</div>';
                html += '</div>';
                html += '<div class="report-hero-meta">';
                html += '<div class="stamp">' + escapeHtml(orderTypeLabel) + ' • ' + escapeHtml(statusLabel) + '</div>';
                html += '<div class="stamp">' + escapeHtml(currentMarginMode === 'new' ? 'Ny marginmodel' : 'Klassisk marginmodel') + '</div>';
                html += '</div>';
                html += '</div>';
                html += '<div class="report-pill-row">';
                html += '<span class="report-pill ok"><strong>' + formatNumber(revenue) + ' DKK</strong> omsætning</span>';
                html += '<span class="report-pill"><strong>' + formatNumber(cost) + ' DKK</strong> kost</span>';
                html += '<span class="report-pill"><strong>' + formatNumber(marginAmount) + ' DKK</strong> margin</span>';
                html += '<span class="report-pill warn"><strong>' + formatCount(exceptionGroupCount) + '</strong> spor/advarsler</span>';
                html += '</div>';
                html += '</div>';

                html += '<div class="order-report-grid">';
                html += '<div class="order-report-card"><div class="label">Samlet omsætning</div><div class="value">' + formatNumber(revenue) + ' DKK</div></div>';
                html += '<div class="order-report-card"><div class="label">Samlet kost</div><div class="value">' + formatNumber(cost) + ' DKK</div></div>';
                html += '<div class="order-report-card"><div class="label">Margin</div><div class="value">' + formatNumber(marginAmount) + ' DKK (' + marginPct + '%)</div></div>';
                html += '<div class="order-report-card"><div class="label">Produktionsordrer</div><div class="value">' + formatNumber(productionOrders.length) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Planlagte minutter</div><div class="value">' + formatNumber(totalPlannedMinutes) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Brugte minutter</div><div class="value">' + formatNumber(totalUsedMinutes) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Operation kost</div><div class="value">' + formatNumber(totalOperationCost) + ' DKK</div></div>';
                html += '<div class="order-report-card"><div class="label">Laser / materiale kost</div><div class="value">' + formatNumber(totalLaserCost) + ' DKK</div></div>';
                html += '</div>';

                html += '<table class="order-report-table">';
                html += '<tr><th>Produktionsordre</th><th>Produkt</th><th>Planlagt min.</th><th>Brugt min.</th><th>Afvigelse</th><th>Operation kost</th><th>Laser / materiale kost</th><th>Samlet kost</th></tr>';
                for (const row of rows) {
                    html += '<tr>';
                    html += '<td>' + row.ordNo + '</td>';
                    html += '<td>' + escapeHtml(row.productLabel || '-') + '</td>';
                    html += '<td>' + formatNumber(row.plannedMinutes || 0) + '</td>';
                    html += '<td>' + formatNumber(row.usedMinutes || 0) + '</td>';
                    html += '<td>' + formatNumber(row.deltaMinutes || 0) + '</td>';
                    html += '<td>' + formatNumber(row.operationCost || 0) + ' DKK</td>';
                    html += '<td>' + formatNumber(row.laserCost || 0) + ' DKK</td>';
                    html += '<td><strong>' + formatNumber(row.totalCost || 0) + ' DKK</strong></td>';
                    html += '</tr>';
                }
                if (rows.length === 0) {
                    html += '<tr><td colspan="8">Ingen produktionsordrer fundet.</td></tr>';
                }
                html += '<tr class="summary-row"><td colspan="2">Samlet</td><td>' + formatNumber(totalPlannedMinutes || 0) + '</td><td>' + formatNumber(totalUsedMinutes || 0) + '</td><td>' + formatNumber(totalUsedMinutes - totalPlannedMinutes) + '</td><td>' + formatNumber(totalOperationCost || 0) + ' DKK</td><td>' + formatNumber(totalLaserCost || 0) + ' DKK</td><td><strong>' + formatNumber(cost || 0) + ' DKK</strong></td></tr>';
                html += '</table>';

                html += '<div class="order-report-grid" style="margin-top:12px;">';
                html += '<div class="order-report-card"><div class="label">Salgsordrer</div><div class="value">' + formatNumber(salesOrderLines.length) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Ekstra salgslinjer</div><div class="value">' + formatNumber(salesLines.length) + '</div></div>';
                html += '<div class="order-report-card"><div class="label">Kost til dato</div><div class="value">' + formatNumber(costToDateFromProduction || 0) + ' DKK</div></div>';
                html += '<div class="order-report-card"><div class="label">Marginprocent</div><div class="value">' + marginPct + '%</div></div>';
                html += '</div>';

                html += '<div class="order-report-card" style="margin-top:12px;">';
                html += '<div class="label">Exceptioner og spor</div>';
                html += '<div class="value" style="font-size:14px; font-weight:600; margin-bottom:8px;">Advarsler: ' + formatCount(exceptionGroupCount) + ' grupper • ' + formatCount(exceptionOccurrenceCount) + ' forekomster</div>';
                if (exceptionCompactRows.length > 0) {
                    html += '<table class="order-report-table" style="margin-top:0;">';
                    html += '<tr><th>Prod.ordre</th><th>Type</th><th>Linjer/Ref</th><th>Beskrivelse</th><th>Antal</th></tr>';
                    for (const ex of exceptionCompactRows) {
                        const refsPreview = ex.refs.length <= 4
                            ? ex.refs.join(', ')
                            : (ex.refs.slice(0, 4).join(', ') + ' +' + (ex.refs.length - 4));
                        const refsTitle = ex.refs.join(', ');
                        html += '<tr>';
                        html += '<td>' + escapeHtml(ex.prodOrdNo || '-') + '</td>';
                        html += '<td>' + escapeHtml(ex.types.join(' + ') || '-') + '</td>';
                        html += '<td title="' + escapeHtml(refsTitle) + '">' + escapeHtml(refsPreview || '-') + '</td>';
                        html += '<td>' + escapeHtml(ex.message || '-') + '</td>';
                        html += '<td>' + formatCount(ex.count || 0) + '</td>';
                        html += '</tr>';
                    }
                    html += '</table>';
                } else {
                    html += '<div style="color:#4b5563; font-size:13px;">Ingen ekceptioner fundet i den aktuelle ordre.</div>';
                }
                html += '</div>';
                html += '</div>';
                return html;
            }

            function toggleOrderDetailReport() {
                const report = document.getElementById('orderDetailReport');
                if (!report) return;
                const isClosed = report.style.display === 'none';
                report.style.display = isClosed ? '' : 'none';
            }

            let currentPrintPreviewMode = null;
            let reportPrintRestoreState = null;

            function isOrderDetailReportViewActive() {
                const bodyEl = document.getElementById('orderDetailModalBody');
                return Boolean(bodyEl && bodyEl.querySelector('#orderDetailReport'));
            }

            function updateOrderDetailModalBackButton() {
                const btn = document.getElementById('orderDetailModalBackBtn');
                if (!btn) return;
                const show = Boolean(reportOriginState && isOrderDetailReportViewActive());
                btn.style.display = show ? '' : 'none';
            }

            function restoreOrderDetailFromReport() {
                if (!reportOriginState) return false;
                const snapshot = reportOriginState;
                reportOriginState = null;
                openOrderDetailModal(snapshot.html, snapshot.title, snapshot.subtitle);
                return true;
            }

            function goBackFromReportToOrder() {
                restoreOrderDetailFromReport();
            }

            function openOrderDetailModal(html, titleText, subtitleText) {
                const overlay = document.getElementById('orderDetailModal');
                const titleEl = document.getElementById('orderDetailModalTitle');
                const subtitleEl = document.getElementById('orderDetailModalSubtitle');
                const bodyEl = document.getElementById('orderDetailModalBody');
                if (!overlay || !titleEl || !subtitleEl || !bodyEl) return;
                titleEl.textContent = titleText || 'Ordre-rapport';
                subtitleEl.textContent = subtitleText || 'Manager-oversigt med produktion, cost og sporbarhed';
                bodyEl.innerHTML = html || '';
                applyMicroTablePolish(bodyEl);
                overlay.style.display = 'flex';
                document.body.classList.add('report-modal-open');
                updateOrderDetailModalBackButton();
            }

            function closeOrderDetailModal(event) {
                if (event && event.target && event.target.id !== 'orderDetailModal') return;
                if (isOrderDetailReportViewActive() && reportOriginState) {
                    restoreOrderDetailFromReport();
                    return;
                }
                const overlay = document.getElementById('orderDetailModal');
                const bodyEl = document.getElementById('orderDetailModalBody');
                if (overlay) overlay.style.display = 'none';
                if (bodyEl) bodyEl.innerHTML = '';
                document.body.classList.remove('report-modal-open');
                reportOriginState = null;
                updateOrderDetailModalBackButton();
                goBackToList();
            }

            function buildStandaloneReportPrintCss() {
                return [
                    '@page { size: A4 portrait; margin: 12mm; }',
                    'body { margin: 0; background: #fff; color: #10253f; font-family: Arial, sans-serif; }',
                    '.order-detail-report { display:block !important; }',
                    '.order-report-toolbar,',
                    '.order-report-actions,',
                    '.list-toggle-btn,',
                    'button { display:none !important; }',
                    '.section { break-inside: avoid; page-break-inside: avoid; margin-bottom: 14px; }',
                    '.order-report-grid { grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; }',
                    '.order-report-card { break-inside: avoid; page-break-inside: avoid; }',
                    '.order-report-table { width:100%; border-collapse:collapse; }',
                    '.order-report-table th, .order-report-table td { border-bottom:1px solid #dfe8f3; padding:8px 10px; }',
                    '.order-report-table th { background:#eef5ff; }',
                    '.summary-box { break-inside: avoid; page-break-inside: avoid; }'
                ].join('\n');
            }

            function buildStandaloneListPrintCss() {
                return [
                    '@page { size: A4 portrait; margin: 7mm; }',
                    'body { margin:0; font-family: Arial, sans-serif; color:#1f2937; font-size:10px; line-height:1.2; }',
                    '.order-list-summary-actions, .order-report-actions, .search-box, .header-banner-wrapper { display:none !important; }',
                    '.order-list-section { margin:0 !important; box-shadow:none !important; border:none !important; padding:0 !important; }',
                    '.order-list-summary { margin:0 0 8px 0 !important; padding:7px 9px !important; font-size:10px !important; }',
                    '.order-list-section h3 { margin:0 0 7px 0 !important; padding:0 0 5px 0 !important; font-size:12px !important; }',
                    '.order-list-table { width:100%; font-size:9.3px; border-collapse:collapse; table-layout:auto; }',
                    '.order-list-table th, .order-list-table td { padding:3px 4px; line-height:1.15; }',
                    '.order-list-table th:nth-child(8), .order-list-table td:nth-child(8), .order-list-table th:nth-child(9), .order-list-table td:nth-child(9) { display:none; }',
                    '.order-list-table tr { break-inside: avoid; page-break-inside: avoid; }'
                ].join('\n');
            }

            function printStandaloneHtml(title, html, cssText) {
                const iframe = document.createElement('iframe');
                iframe.setAttribute('aria-hidden', 'true');
                iframe.style.position = 'fixed';
                iframe.style.right = '0';
                iframe.style.bottom = '0';
                iframe.style.width = '0';
                iframe.style.height = '0';
                iframe.style.border = '0';
                iframe.style.opacity = '0';
                document.body.appendChild(iframe);

                const styles = Array.from(document.querySelectorAll('style, link[rel="stylesheet"]'))
                    .map(el => el.outerHTML)
                    .join('\n');
                const safeTitle = escapeHtml(title || 'Udskrift');
                const doc = iframe.contentWindow.document;
                doc.open();
                doc.write('<!doctype html><html><head><meta charset="utf-8"><title>' + safeTitle + '</title>' + styles + '<style>' + (cssText || '') + '</style></head><body>' + (html || '') + '</body></html>');
                doc.close();

                setTimeout(() => {
                    try {
                        iframe.contentWindow.focus();
                        iframe.contentWindow.print();
                    } finally {
                        setTimeout(() => iframe.remove(), 1500);
                    }
                }, 250);
            }

            function printOrderDetailReport() {
                const bodyEl = document.getElementById('orderDetailModalBody');
                if (!bodyEl) return;
                const titleEl = document.getElementById('orderDetailModalTitle');
                const reportTitle = titleEl ? titleEl.textContent : 'Ordre-rapport';
                printStandaloneHtml(reportTitle, bodyEl.innerHTML, buildStandaloneReportPrintCss());
            }

            function renderPrintPreview(title, html, mode) {
                const overlay = document.getElementById('printPreviewOverlay');
                const titleEl = document.getElementById('printPreviewTitle');
                const bodyEl = document.getElementById('printPreviewBody');
                if (!overlay || !titleEl || !bodyEl) return;
                currentPrintPreviewMode = mode;
                titleEl.textContent = title;
                bodyEl.innerHTML = html;
                overlay.style.display = 'flex';
                document.body.classList.add('print-preview-lock');
                document.body.classList.add('print-preview-mode');
                const dialog = overlay.querySelector('.print-preview-dialog');
                if (dialog && dialog.focus) dialog.focus();
            }

            function closePrintPreview(event) {
                if (event && event.target && event.target.id !== 'printPreviewOverlay') return;
                const overlay = document.getElementById('printPreviewOverlay');
                if (overlay) overlay.style.display = 'none';
                const bodyEl = document.getElementById('printPreviewBody');
                if (bodyEl) bodyEl.innerHTML = '';
                document.body.classList.remove('print-preview-lock');
                document.body.classList.remove('print-preview-mode');
                currentPrintPreviewMode = null;
            }

            function confirmPrintFromPreview() {
                const bodyEl = document.getElementById('printPreviewBody');
                if (!bodyEl) return;
                const titleEl = document.getElementById('printPreviewTitle');
                const title = titleEl ? titleEl.textContent : 'Forhåndsvisning';
                const cssText = currentPrintPreviewMode === 'list'
                    ? buildStandaloneListPrintCss()
                    : buildStandaloneReportPrintCss();
                printStandaloneHtml(title, bodyEl.innerHTML, cssText);
            }

            document.addEventListener('keydown', function(event) {
                if (event.key === 'Escape' && document.body.classList.contains('print-preview-mode')) {
                    closePrintPreview();
                }
            });

            function openOrderListPrintPreview() {
                if (!orderListVisible) {
                    toggleOrderList();
                }
                const listEl = document.getElementById('orderList');
                if (!listEl) return;
                renderPrintPreview('Forhåndsvisning - ordreliste', listEl.innerHTML, 'list');
            }

            function openOrderDetailPrintPreview() {
                const report = document.getElementById('orderDetailReport');
                if (report) {
                    renderPrintPreview('Forhåndsvisning - rapport', report.outerHTML, 'report');
                    return;
                }
                if (lastOrderReportHtml) {
                    renderPrintPreview(lastOrderReportTitle || 'Forhåndsvisning - rapport', lastOrderReportHtml, 'report');
                }
            }

            function openLatestOrderReportPreview() {
                if (!lastOrderReportHtml) {
                    alert('Rapporten er ikke klar endnu. Søg efter en ordre først.');
                    return;
                }

                const overlay = document.getElementById('orderDetailModal');
                const bodyEl = document.getElementById('orderDetailModalBody');
                const titleEl = document.getElementById('orderDetailModalTitle');
                const subtitleEl = document.getElementById('orderDetailModalSubtitle');
                const modalOpen = overlay && getComputedStyle(overlay).display === 'flex';
                if (modalOpen && bodyEl && !isOrderDetailReportViewActive() && !reportOriginState) {
                    reportOriginState = {
                        html: bodyEl.innerHTML,
                        title: titleEl ? titleEl.textContent : 'Ordre-rapport',
                        subtitle: subtitleEl ? subtitleEl.textContent : 'Manager-oversigt med produktion, cost og sporbarhed'
                    };
                }

                openOrderDetailModal(
                    lastOrderReportHtml,
                    lastOrderReportTitle || 'Rapport 2.0',
                    'Manager-oversigt med produktion, cost og sporbarhed'
                );
            }

            function updateReportOpenButtonState(isReady, ordNo) {
                const btn = document.getElementById('openReportBtn');
                if (!btn) return;
                const ready = Boolean(isReady && lastOrderReportHtml);
                btn.disabled = !ready;
                if (ready) {
                    btn.title = 'Åbn seneste rapport i separat visning';
                    btn.textContent = ordNo ? ('Rapport ' + ordNo) : 'Rapport 2.0';
                } else {
                    btn.title = 'Søg efter en ordre for at aktivere rapport';
                    btn.textContent = 'Rapport 2.0';
                }
            }

            window.addEventListener('afterprint', function() {
                document.body.classList.remove('print-report-mode');
                document.body.classList.remove('print-list-mode');
                document.body.classList.remove('print-preview-mode');
                const report = document.getElementById('orderDetailReport');
                if (report && reportPrintRestoreState !== null) {
                    report.style.display = reportPrintRestoreState;
                    reportPrintRestoreState = null;
                }
                const overlay = document.getElementById('printPreviewOverlay');
                if (overlay) overlay.style.display = 'none';
                const bodyEl = document.getElementById('printPreviewBody');
                if (bodyEl) bodyEl.innerHTML = '';
                currentPrintPreviewMode = null;
            });

            function updateOrderListSummaryPanel() {
                const summaryEl = document.getElementById('orderListSummary');
                if (!summaryEl || !orderListVisible) return;
                const orders = getFilteredOrders();
                summaryEl.innerHTML = buildOrderListSummaryHtml(orders);
            }

            async function searchOrder() {
                const ordNo = document.getElementById('orderInput').value;
                if (!ordNo) {
                    alert('Indtast et ordrenummer');
                    return;
                }

                const requestId = ++activeSearchRequestId;

                // Keep customer list visible during direct order search.
                const result = document.getElementById('result');
                result.innerHTML = '<div class="loading">Indlæser...</div>';
                openOrderDetailModal(
                    '<div class="section"><h3>Ordre ' + escapeHtml(String(ordNo)) + '</h3><div class="loading">Henter ordredata...</div></div>',
                    'Ordre ' + escapeHtml(String(ordNo)) + ' - indlæser...',
                    'Forbereder produktion, cost og sporbarhed...'
                );
                
                try {
                    const data = await requestAftercalcData(ordNo);
                    if (requestId !== activeSearchRequestId) return;
                    
                    if (data.error) {
                        openOrderDetailModal(
                            '<div class="section"><div class="error">Fejl: ' + escapeHtml(String(data.error)) + '</div></div>',
                            'Ordre ' + escapeHtml(String(ordNo)) + ' - fejl',
                            'Kunne ikke hente data for ordren'
                        );
                        result.innerHTML = '<div class="error">Fejl: ' + data.error + '</div>';
                        return;
                    }

                    // NOTE: Gr4 is order type (e.g., Multiordre). Do not change Gr4 logic here.
                    currentSalesOrderGr4 = Number((data.orderHeader && data.orderHeader.Gr4) || 0);
                    currentSearchOrderData = data;
                    const orderMarginPercent = calculateOrderMarginPercent(data.summary.totalRevenue, data.summary.totalCost).toFixed(2);
                    const _invoAm = Number(data.orderHeader.InvoAm || 0);
                    const _dInvoIF = Number(data.orderHeader.DInvoIF || 0);
                    let invoiceStatusBadge, invoiceStatusSub = '';
                    if (_invoAm === 0) {
                        invoiceStatusBadge = '<span class="invoice-status-badge status-in-production">🔧 I produktion</span>';
                    } else if (_dInvoIF <= 0) {
                        invoiceStatusBadge = '<span class="invoice-status-badge status-fully-invoiced">✅ Komplet faktureret</span>';
                    } else {
                        invoiceStatusBadge = '<span class="invoice-status-badge status-partial-invoiced">⏳ Delvist faktureret</span>';
                        invoiceStatusSub = '<div style="font-size:12px; opacity:0.85; margin-top:5px;">Faktureret: ' + formatNumber(_invoAm) + ' | Mangler: ' + formatNumber(_dInvoIF) + '</div>';
                    }
                    const productionOrderByOrdNo = new Map((Array.isArray(data.productionOrders) ? data.productionOrders : []).map(order => [Number(order.ordNo || 0), order]));
                    const costToDateFromProduction = (Array.isArray(data.productionOrders) ? data.productionOrders : [])
                        .reduce((sum, order) => sum + Number((order && order.totalCost) || 0), 0);
                    const getSalesLineCostBreakdown = (purcNo) => {
                        const prodOrder = productionOrderByOrdNo.get(Number(purcNo || 0));
                        const lines = Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : [];
                        let operationTotal = 0;
                        let laserTotal = 0;
                        let materialTotal = 0;

                        for (const line of lines) {
                            const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                            const lnNo = Number((line && line.LnNo) || 0);
                            if (lnNo === 1 || key === '0' || key === '3' || key === '5') continue;
                            const effectiveCost = Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                            if (key === '1') {
                                operationTotal += effectiveCost;
                            } else if (key === '2') {
                                laserTotal += effectiveCost;
                                if (!isLaserLProdNo(line && line.ProdNo)) {
                                    materialTotal += effectiveCost;
                                }
                            }
                        }

                        return {
                            operationTotal,
                            laserTotal
                        };
                    };
                    
                    let html = '<div class="order-header">';
                    html += '<h2>Salgsordre: ' + data.orderHeader.OrdNo + ' - ' + (data.orderHeader.CustomerName || '-') + '</h2>';
                    const _noteOrdNo = Number(data.orderHeader.OrdNo);
                    const _existingNote = orderNotesCache[String(_noteOrdNo)];
                    const _noteIcons = { ok: '✅', error: '❌', check: '⚠️', credit: '🧾' };
                    const _noteIcon = _existingNote && _existingNote.isCreditNote
                        ? _noteIcons.credit
                        : (_existingNote && _existingNote.status ? (_noteIcons[_existingNote.status] || '📝') : '📝');
                    const _noteCls = _existingNote && _existingNote.isCreditNote
                        ? 'credit'
                        : (_existingNote && _existingNote.status ? _existingNote.status : 'text');
                    const _noteDisplay = _existingNote && (_existingNote.status || _existingNote.text || _existingNote.isCreditNote) ? 'flex' : 'none';
                    html += '<div id="order-note-banner-' + _noteOrdNo + '" class="order-note-banner ' + _noteCls + '" style="display:' + _noteDisplay + ';" onclick="openNotePopup(' + _noteOrdNo + ',true)">';
                    if (_existingNote && (_existingNote.status || _existingNote.text || _existingNote.isCreditNote)) {
                        html += '<span class="note-icon">' + _noteIcon + '</span><div class="note-body"><strong>' + (_existingNote.isCreditNote ? 'Kreditnota' : (_existingNote.status === 'ok' ? 'OK' : _existingNote.status === 'error' ? 'Fejl' : _existingNote.status === 'check' ? 'Tjek' : 'Note')) + '</strong>' + (_existingNote.text ? ': ' + escapeHtml(_existingNote.text) : '') + '</div><span style="font-size:11px;opacity:0.7;margin-left:auto;">✏️ Rediger</span>';
                    }
                    html += '</div>';
                    html += '<button onclick="openNotePopup(' + _noteOrdNo + ',true)" style="border:none;background:transparent;cursor:pointer;font-size:12px;color:#888;padding:0 0 8px 0;">📝 ' + (_existingNote && (_existingNote.status || _existingNote.text || _existingNote.isCreditNote) ? 'Rediger note' : 'Tilføj note') + '</button>';
                    html += '<div class="order-report-actions" style="display:flex; gap:8px; flex-wrap:wrap; margin:6px 0 12px 0;">';
                    html += '<button class="list-toggle-btn" onclick="openLatestOrderReportPreview()" title="Åbn rapporten i separat visning">Rapport 2.0</button>';
                    html += '</div>';
                    loadOrderNote(_noteOrdNo).catch(() => {});
                    html += '<div class="order-header-row">';
                    if (_invoAm === 0) {
                        // I Produktion: show cost to date + projected margin if DInvoIF available
                        html += '<div class="order-header-item"><div class="order-header-label">Kost til dato (estimat)</div><div class="order-header-value">' + formatNumber(costToDateFromProduction) + ' DKK</div></div>';
                        if (_dInvoIF > 0) {
                            const projectedMargin = _dInvoIF - costToDateFromProduction;
                            const projectedMarginPct = costToDateFromProduction > 0 ? calculateOrderMarginPercent(_dInvoIF, costToDateFromProduction).toFixed(2) : '0.00';
                            html += '<div class="order-header-item"><div class="order-header-label">Forventet salgsbeløb</div><div class="order-header-value">' + formatNumber(_dInvoIF) + ' DKK</div></div>';
                            html += '<div class="order-header-item"><div class="order-header-label">Forventet margin (prognose)</div><div class="order-header-value">' + getMarginBadge(projectedMarginPct) + '<div style="font-size:13px; opacity:0.85; margin-top:4px;">' + formatNumber(projectedMargin) + ' DKK</div></div></div>';
                        } else {
                            html += '<div class="order-header-item"><div class="order-header-label">Forventet salgsbeløb</div><div class="order-header-value" style="opacity:0.6; font-size:16px;">— (ukendt)</div></div>';
                            html += '<div class="order-header-item"><div class="order-header-label">Margin</div><div class="order-header-value"><span style="background:rgba(255,255,255,0.15); color:#fff; font-weight:bold; padding:2px 8px; border-radius:4px; font-size:14px;">— Ingen data</span></div></div>';
                        }
                    } else {
                        html += '<div class="order-header-item"><div class="order-header-label">Faktureret beløb</div><div class="order-header-value">' + formatNumber(data.summary.totalRevenue) + ' DKK</div></div>';
                        html += '<div class="order-header-item"><div class="order-header-label">Kostpris</div><div class="order-header-value">' + formatNumber(data.summary.totalCost) + ' DKK</div></div>';
                        html += '<div class="order-header-item"><div class="order-header-label">Margin (' + getMarginModeLabel() + ')</div><div class="order-header-value">' + getMarginBadge(orderMarginPercent) + '</div></div>';
                    }
                    html += '<div class="order-header-item"><div class="order-header-label">Fakturastatus</div><div class="order-header-value">' + invoiceStatusBadge + invoiceStatusSub + '</div></div>';
                    html += '</div></div>';

                    html += '<div class="section oversigt-launcher-section">';
                    html += '<h3>Produktionsoversigter</h3>';
                    html += '<div class="oversigt-launcher-grid">';
                    html += '<article class="oversigt-launcher-card">';
                    html += '<h4>Laseroversigt (L-linjer)</h4>';
                    html += '<div class="desc">Nesting, vægtafvigelser og kost på tværs af ruter.</div>';
                    html += '<div id="laserOversigtSummaryTeaser" class="oversigt-launcher-kpi"><span class="loading">Indlæser laser-KPI...</span></div>';
                    html += '<button class="list-toggle-btn" onclick="openOversigtModal(\'laser\')" title="Åbn detaljeret laseroversigt">Åbn laseroversigt</button>';
                    html += '</article>';
                    html += '<article class="oversigt-launcher-card">';
                    html += '<h4>Operation Oversigt</h4>';
                    html += '<div class="desc">Operationstid, afvigelser og omkostninger i én driftsvisning.</div>';
                    html += '<div id="operationOversigtSummaryTeaser" class="oversigt-launcher-kpi"><span class="loading">Indlæser operations-KPI...</span></div>';
                    html += '<button class="list-toggle-btn" onclick="openOversigtModal(\'operation\')" title="Åbn detaljeret operationsoversigt">Åbn operationer</button>';
                    html += '</article>';
                    html += '</div>';
                    html += '</div>';

                    html += '<div id="laserOrderSummaryPanel" style="display:none;" aria-hidden="true">';
                    html += '<div id="laserOrderSummaryTotals" class="summary-box"><div class="loading">Indlæser totaler...</div></div>';
                    html += '<div class="laser-summary-layout">';
                    html += '<div id="laserOrderSummaryBody" class="loading">Indlæser laserdata...</div>';
                    html += '<aside id="laserImagePanel" class="laser-image-panel hidden"></aside>';
                    html += '</div>';
                    html += '</div>';

                    html += '<div id="operationOrderSummaryPanel" style="display:none;" aria-hidden="true">';
                    html += '<div id="operationOrderSummaryTotals" class="summary-box"><div class="loading">Indlæser totaler...</div></div>';
                    html += '<div id="operationOrderSummaryBody" class="loading">Indlæser operationsdata...</div>';
                    html += '</div>';

                    lastOrderReportHtml = buildOrderDetailReportHtml(data, orderMarginPercent, costToDateFromProduction);
                    lastOrderReportTitle = 'Rapport 2.0 - ordre ' + String(data.orderHeader.OrdNo || '-');
                    updateReportOpenButtonState(true, String(data.orderHeader.OrdNo || ''));

                    // Sezione linee ORDINE DI VENDITA complete
                    if (data.salesOrderLines && data.salesOrderLines.length > 0) {
                        const hasSalesOrderDrawing = data.salesOrderLines.some(line => !!line.DrawingWebPg);
                        const salesOrderColspan = hasSalesOrderDrawing ? 11 : 10;
                        html += '<div class="section"><h3>Salgsordrelinjer</h3>';
                        html += '<table><tr><th>Linje</th><th>Produkt</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Kostpris</th><th>Samlet kost</th><th>Salgspris/enhed</th><th>Salgspris</th><th>Margin (%)</th><th>Prod.ordre</th>' + (hasSalesOrderDrawing ? '<th>Vis tegning</th>' : '') + '</tr>';

                        for (const line of data.salesOrderLines) {
                            const lineSalesPrice = (line.DPrice || 0) * (line.NoFin || 0);
                            const lineCost = line.EffectiveLineCost || 0;
                            const lineProdNo = String(line.ProdNo || '').trim();
                            const includeForMargin = lineProdNo.startsWith('1') || lineProdNo.startsWith('3');
                            const lineMarginValue = calculateLineMarginPercent(lineSalesPrice, lineCost);
                            const isExactlyHundred = Math.abs(lineMarginValue - 100) < 0.0001;
                            const lineMarginPercent = lineMarginValue.toFixed(2);
                            const hasProductionOrder = Boolean(line.PurcNo && line.PurcNo !== 0);
                            const breakdownRowId = 'sales-line-breakdown-' + String(data.orderHeader.OrdNo || '0') + '-' + String(line.LnNo || 0);
                            const breakdownInfo = hasProductionOrder
                                ? getSalesLineCostBreakdown(line.PurcNo)
                                : { operationTotal: 0, laserTotal: 0 };
                            // Rabat-badge fjernet: linjer med salgspris=0 (underlinjer af hovedprodukt) vises som N/A.
                            const lineMarginBadge = (!includeForMargin || lineSalesPrice === 0 || isExactlyHundred)
                                ? '<span style="background:#607d8b; color:#fff; font-weight:bold; padding:2px 6px; border-radius:4px;">N/A</span>'
                                : getMarginBadge(lineMarginPercent);
                            html += '<tr>';
                            html += '<td>' + (hasProductionOrder
                                ? ('<button type="button" onclick="toggleSalesLineBreakdown(\'' + breakdownRowId + '\', this)" title="Vis kost-opdeling" style="margin-right:6px; width:22px; height:22px; border:1px solid #90caf9; background:#e3f2fd; color:#0d47a1; border-radius:4px; cursor:pointer; font-weight:700;">+</button>')
                                : '') + (line.LnNo || 0) + '</td>';

                            const salesWarningFlag = getWarningFlagHtml(line, 'Tilknyttet produktionsordre har en advarsel.');
                            if (line.PurcNo && line.PurcNo !== 0) {
                                html += '<td><span class="prod-link" onclick="openProduction(' + line.PurcNo + ')">' + (line.ProdNo || '-') + '</span>' + salesWarningFlag + '</td>';
                            } else {
                                html += '<td>' + (line.ProdNo || '-') + salesWarningFlag + '</td>';
                            }

                            const displaySalesQty = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                                ? line.DisplayQuantity
                                : (line.NoFin || 0);
                            html += '<td>' + (line.Descr || '') + '</td>';
                            html += '<td>' + formatNumber(displaySalesQty) + '</td>';
                            const productionTotalCost = Number(line.ProductionOrderTotalCost || 0);
                            const lineQty = Number(line.NoFin || 0);
                            const displayKostpris = (line.PurcNo && line.PurcNo !== 0)
                                ? (lineQty > 0 ? (productionTotalCost / lineQty) : productionTotalCost)
                                : (line.CCstPr || 0);
                            html += '<td>' + formatNumber(displayKostpris) + '</td>';
                            html += '<td><strong>' + formatNumber(lineCost) + '</strong></td>';
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                            html += '<td>' + formatNumber(lineSalesPrice) + '</td>';
                            html += '<td>' + lineMarginBadge + '</td>';
                            html += '<td>' + ((line.PurcNo && line.PurcNo !== 0) ? line.PurcNo : '-') + '</td>';
                            if (hasSalesOrderDrawing) {
                                if (line.DrawingWebPg) {
                                    html += '<td><button class="list-toggle-btn drawing-open-btn" data-drawing-path="' + escapeHtml(String(line.DrawingWebPg || '')) + '" data-prod-no="' + escapeHtml(String(line.ProdNo || '')) + '" data-ord-no="' + escapeHtml(String(line.PurcNo || data.orderHeader.OrdNo || '')) + '" style="padding:4px 8px; margin-left:0;">Vis tegning</button></td>';
                                } else {
                                    html += '<td></td>';
                                }
                            }
                            html += '</tr>';
                            if (hasProductionOrder) {
                                html += '<tr id="' + breakdownRowId + '" style="display:none; background:#f8fbff;">';
                                html += '<td colspan="' + salesOrderColspan + '" style="padding:10px 16px; border-top:none;">';
                                html += '<div style="display:grid; gap:6px; color:#1f2937;">';
                                html += '<div><strong>Operation:</strong> ' + formatNumber(breakdownInfo.operationTotal || 0) + ' DKK</div>';
                                html += '<div><strong>Laser / materiale:</strong> ' + formatNumber(breakdownInfo.laserAndMaterialTotal || 0) + ' DKK</div>';
                                if ((breakdownInfo.materialTotal || 0) > 0) {
                                    html += '<div><strong>Materiale (ikke L):</strong> ' + formatNumber(breakdownInfo.materialTotal || 0) + ' DKK</div>';
                                }
                                html += '</div>';
                                html += '</td>';
                                html += '</tr>';
                            }
                        }

                        html += '</table></div>';
                    }
                    
                    // Sezione linee di vendita
                    if (data.salesLines.length > 0) {
                        const hasSalesLinesDrawing = data.salesLines.some(line => !!line.DrawingWebPg);
                        html += '<div class="section"><h3>Salgslinjer (Ekstra produkter)</h3>';
                        html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Salgspris</th><th>Kostpris/enhed</th><th>Samlet kost</th>' + (hasSalesLinesDrawing ? '<th>Vis tegning</th>' : '') + '</tr>';
                        
                        for (const line of data.salesLines) {
                            const salesExtraWarningFlag = getWarningFlagHtml(line, 'Inkonsekvens på salgslinje.');
                            const displaySalesExtraQty = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                                ? line.DisplayQuantity
                                : (line.NoFin || 0);
                            html += '<tr>';
                            html += '<td>' + (line.ProdNo || '-') + salesExtraWarningFlag + '</td>';
                            html += '<td>' + (line.Descr || '') + '</td>';
                            html += '<td>' + formatNumber(displaySalesExtraQty) + '</td>';
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                            html += '<td>' + formatNumber(line.CCstPr || 0) + '</td>';
                            html += '<td><strong>' + formatNumber(line.EffectiveLineCost || 0) + '</strong></td>';
                            if (hasSalesLinesDrawing) {
                                if (line.DrawingWebPg) {
                                    html += '<td><button class="list-toggle-btn drawing-open-btn" data-drawing-path="' + escapeHtml(String(line.DrawingWebPg || '')) + '" data-prod-no="' + escapeHtml(String(line.ProdNo || '')) + '" data-ord-no="' + escapeHtml(String(line.PurcNo || data.orderHeader.OrdNo || '')) + '" style="padding:4px 8px; margin-left:0;">Vis tegning</button></td>';
                                } else {
                                    html += '<td></td>';
                                }
                            }
                            html += '</tr>';
                        }
                        
                        html += '<tr class="summary-row"><td colspan="5">Total salgslinjer:</td><td>' + formatNumber(data.salesLinesTotalCost) + ' DKK</td>' + (hasSalesLinesDrawing ? '<td></td>' : '') + '</tr>';
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
                            '9': 'Indkøbt dele',
                            'NA': 'Ikke sat'
                        };
                        
                        for (const prodOrder of data.productionOrders) {
                            const mainProductLine = prodOrder.lines.find(line => line.ProdTp4 === 0) || prodOrder.lines.find(line => line.LnNo === 1);
                            const mainProductText = mainProductLine
                                ? ((mainProductLine.ProdNo || '-') + ' - ' + (mainProductLine.Descr || ''))
                                : '-';

                            html += '<div id="po-' + prodOrder.ordNo + '" data-order="' + prodOrder.ordNo + '" style="margin-bottom: 20px; border: 1px solid #ddd; padding: 15px; border-radius: 4px;">';
                            const prodOrderTimeFlagHtml = getTimeAdjustmentFlagHtml({
                                hasEstimatedOperationTime: !!prodOrder.hasEstimatedOperationTime,
                                EstimatedTimeText: 'Mindst én operation er genberegnet ud fra Stykliste Minutter, fordi Færdigmeldt var 0.'
                            });
                            html += '<h4>Produktionsordre: ' + prodOrder.ordNo + prodOrderTimeFlagHtml + getWarningFlagHtml({ HasWarning: !!prodOrder.hasWarnings, WarningText: prodOrder.warningText || '' }, 'Denne produktionsordre indeholder mindst en advarselslinje.') + '</h4>';
                            html += '<div class="main-product-box">';
                            html += '<div class="value">' + mainProductText + '</div>';
                            html += '</div>';
                            html += '<div class="prodtp4-hint">Klik paa en linje for at aabne/lukke detaljer.</div>';

                            const groupedLines = {};
                            const operationMergeMap = new Map();
                            const pendingNoOrgFromTp3 = new Map();
                            for (const line of prodOrder.lines) {
                                const rawKey = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                                if (rawKey === '0' || rawKey === '5') continue;

                                // Merge operation rows where ProdTp4 is 1 or 3 and ProdNo is the same.
                                const prodNoKey = String(line.ProdNo || '').trim().toUpperCase();
                                const normalizedKey = getDisplayProdTp4Key(rawKey, prodNoKey, line.PurcNo);


                                // R1090/R8200 must be fully excluded from Operations: no row and no cost contribution.
                                if (normalizedKey === '1' && isExcludedOperationProdNo(prodNoKey)) {
                                    continue;
                                }

                                // R-products under Produkt dele must never be shown or counted.
                                if (normalizedKey === '4' && prodNoKey.startsWith('R')) {
                                    continue;
                                }

                                if (normalizedKey === '1') {
                                    if (prodNoKey) {
                                        const mergeKey = normalizedKey + '|' + prodNoKey;
                                        if (rawKey === '3') {
                                            const extraNoOrg = Number(line.NoOrg || 0);
                                            if (operationMergeMap.has(mergeKey)) {
                                                const mergedLine = operationMergeMap.get(mergeKey);
                                                mergedLine.NoOrg = Number(mergedLine.NoOrg || 0) + extraNoOrg;
                                            } else {
                                                pendingNoOrgFromTp3.set(mergeKey, Number(pendingNoOrgFromTp3.get(mergeKey) || 0) + extraNoOrg);
                                            }
                                            continue;
                                        }

                                        if (!operationMergeMap.has(mergeKey)) {
                                            const extraNoOrg = Number(pendingNoOrgFromTp3.get(mergeKey) || 0);
                                            const mergedLine = {
                                                ...line,
                                                ProdTp4: 1,
                                                NoOrg: Number(line.NoOrg || 0) + extraNoOrg,
                                                NoFin: Number(line.NoFin || 0),
                                                LineCost: Number(line.LineCost || 0),
                                                EffectiveLineCost: Number(line.EffectiveLineCost || 0)
                                            };
                                            operationMergeMap.set(mergeKey, mergedLine);
                                            if (!groupedLines[normalizedKey]) groupedLines[normalizedKey] = [];
                                            groupedLines[normalizedKey].push(mergedLine);
                                        } else {
                                            const mergedLine = operationMergeMap.get(mergeKey);
                                            mergedLine.NoOrg = Number(mergedLine.NoOrg || 0) + Number(line.NoOrg || 0);
                                            mergedLine.NoFin = Number(mergedLine.NoFin || 0) + Number(line.NoFin || 0);
                                            mergedLine.LineCost = Number(mergedLine.LineCost || 0) + Number(line.LineCost || 0);
                                            mergedLine.EffectiveLineCost = Number(mergedLine.EffectiveLineCost || 0) + Number(line.EffectiveLineCost || 0);
                                            if ((!mergedLine.Descr || mergedLine.Descr === '-') && line.Descr) {
                                                mergedLine.Descr = line.Descr;
                                            }
                                        }
                                        continue;
                                    }
                                }

                                if (!groupedLines[normalizedKey]) groupedLines[normalizedKey] = [];
                                groupedLines[normalizedKey].push({ ...line, ProdTp4: normalizedKey === '1' ? 1 : line.ProdTp4 });
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
                                    ? lines.filter(line => line.LnNo !== 1).reduce((sum, line) => {
                                        if (!isLaserLProdNo(line.ProdNo)) {
                                            return sum + (line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null ? (line.EffectiveLineCost || 0) : (line.LineCost || 0));
                                        }
                                        if (line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null) {
                                            return sum + (line.EffectiveLineCost || 0);
                                        }
                                        const hasNestingCost = Number(line.NestingCost || 0) > 0;
                                        return sum + (hasNestingCost
                                            ? ((line.NestingCost || 0) * (line.NoFin || 0))
                                            : (line.LineCost || 0));
                                    }, 0)
                                    : lines.filter(line => line.LnNo !== 1).reduce((sum, line) => {
                                        const pn = String(line.ProdNo || '').toUpperCase();
                                        if (pn === 'R6200' && String(key) === '1') {
                                            return sum + ((line.NoOrg || 0) * (line.CCstPr || 0));
                                        }
                                        return sum + (line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null ? (line.EffectiveLineCost || 0) : (line.LineCost || 0));
                                    }, 0);
                                const isOpenByDefault = false;
                                orderVisibleTotal += subtotal;
                                const groupWarningFlagHtml = getWarningFlagHtml(lines, 'Denne gruppe indeholder mindst en advarselslinje.');

                                html += '<div class="prodtp4-group">';
                                html += '<div class="prodtp4-header" onclick="toggleProdTp4Group(' + prodOrder.ordNo + ', &quot;' + key + '&quot;)">';
                                html += '<span class="prodtp4-label"><span id="po-' + prodOrder.ordNo + '-group-' + key + '-icon">' + (isOpenByDefault ? '▾' : '▸') + '</span> ' + key + ' - ' + (prodTp4Labels[key] || 'Altro') + groupWarningFlagHtml + '</span>';
                                html += '<span class="prodtp4-subtotal">Delsum: ' + formatNumber(subtotal) + ' DKK</span>';
                                html += '</div>';

                                html += '<div id="po-' + prodOrder.ordNo + '-group-' + key + '" class="prodtp4-body" style="display:' + (isOpenByDefault ? '' : 'none') + ';">';
                                if (key === '2') {
                                    const laserCostHeader = currentSalesOrderGr4 === 3 ? 'NestMultiPris' : 'Kostpris nesting';
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>' + laserCostHeader + '</th><th>Samlet kost</th></tr>';
                                } else if (key === '1') {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Stykliste Minutter</th><th>Færdigmeldt minutter</th><th>Kostpris/enhed</th><th>Samlet kost</th></tr>';
                                } else if (key === '6' || key === '9') {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Pris/enhed</th><th>Samlet kost</th></tr>';
                                } else {
                                    html += '<table><tr><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Kostpris/enhed</th><th>Samlet kost</th></tr>';
                                }

                                for (const line of lines) {
                                    html += '<tr>';
                                    const warningFlagHtml = getWarningFlagHtml(line);
                                    const invoiceStatusFlagHtml = getInvoiceStatusFlagHtml(line);
                                    const timeAdjustFlagHtml = getTimeAdjustmentFlagHtml(line);
                                    const laserAllocationFlagHtml = getLaserAllocationFlagHtml(line);
                                    const hasChildProductionOrder = Number(line.PurcNo || 0) > 0;
                                    if (String(key) === '1' && line.ProdNo) {
                                        const safeProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeProdLabel = escapeHtml(getResourceDisplayLabel(line.ProdNo, line.Descr));
                                        const trInf2Value = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : prodOrder.ordNo);
                                        const safeTrInf2 = trInf2Value.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeTrInf4 = String(line.TrInf4 || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        html += '<td><span class="prod-no-link" data-prodno="' + safeProdNo + '" data-ordno="' + prodOrder.ordNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="' + key + '" data-trinf2="' + safeTrInf2 + '" data-trinf4="' + safeTrInf4 + '">' + safeProdLabel + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (String(key) === '2' && line.ProdNo && isLaserLProdNo(line.ProdNo)) {
                                        const safeProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeProdLabel = escapeHtml(getResourceDisplayLabel(line.ProdNo, line.Descr));
                                        const trInf2Value = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : prodOrder.ordNo);
                                        const safeTrInf2 = trInf2Value.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const safeTrInf4 = String(line.TrInf4 || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const linkNoFin = Number(line.NoFin || 0);
                                        const linkHasNestingCost = Number(line.NestingCost || 0) > 0;
                                        const linkHasEffectiveLaserCost = linkNoFin > 0
                                            && line.EffectiveLineCost !== undefined
                                            && line.EffectiveLineCost !== null;
                                        const linkDisplayLaserUnitCost = linkHasEffectiveLaserCost
                                            ? ((line.EffectiveLineCost || 0) / linkNoFin)
                                            : (linkHasNestingCost ? (line.NestingCost || 0) : (line.CCstPr || 0));
                                        html += '<td><span class="prod-no-link" data-prodno="' + safeProdNo + '" data-ordno="' + prodOrder.ordNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="' + key + '" data-trinf2="' + safeTrInf2 + '" data-trinf4="' + safeTrInf4 + '" data-showallroutes="1" data-nofin="' + linkNoFin + '" data-nestingcost="' + Number(linkDisplayLaserUnitCost || 0) + '">' + safeProdLabel + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (hasChildProductionOrder) {
                                        const safeChildProdNoForSummary = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                                        const childSummaryArgs = shouldFilterChildSummary(key, line.ProdNo, line.PurcNo)
                                            ? (Number(line.PurcNo || 0) + ', &quot;' + safeChildProdNoForSummary + '&quot;, true')
                                            : Number(line.PurcNo || 0);
                                        html += '<td><span class="inline-link" onclick="showChildProductionSummary(' + childSummaryArgs + ')">' + (line.ProdNo || '-') + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else if (line.ProdNo) {
                                        html += '<td>' + (line.ProdNo || '-') + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    } else {
                                        html += '<td>-' + timeAdjustFlagHtml + warningFlagHtml + '</td>';
                                    }
                                    html += '<td>' + (line.Descr || '') + '</td>';
                                    if (key === '1') {
                                        const effectiveNoFin = (line.EffectiveOperationMinutes !== undefined && line.EffectiveOperationMinutes !== null)
                                            ? (line.EffectiveOperationMinutes || 0)
                                            : (line.UsesEstimatedOperationTime ? (line.NoOrg || 0) : (line.NoFin || 0));
                                        html += '<td>' + formatNumber(line.NoOrg || 0) + '</td>';
                                        html += '<td>' + formatNumber(effectiveNoFin) + '</td>';
                                        const displayUnitCost1 = (line.CCstPr || 0);
                                        const displayTotalCost1 = (line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null)
                                            ? (line.EffectiveLineCost || 0)
                                            : (effectiveNoFin * (line.CCstPr || 0));
                                        html += '<td>' + formatNumber(displayUnitCost1) + '</td>';
                                        html += '<td><strong>' + formatNumber(displayTotalCost1) + '</strong></td>';
                                    } else {
                                        const displayQty = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                                            ? line.DisplayQuantity
                                            : (line.NoFin || 0);
                                        html += '<td>' + formatNumber(displayQty) + '</td>';
                                    }
                                    if (key === '2') {
                                        const isLaserLine = isLaserLProdNo(line.ProdNo);
                                        const hasNestingCost = Number(line.NestingCost || 0) > 0;
                                        const hasEffectiveLaserCost = Number(line.NoFin || 0) > 0
                                            && line.EffectiveLineCost !== undefined
                                            && line.EffectiveLineCost !== null;
                                        const nestingUnitCost = isLaserLine
                                            ? (hasEffectiveLaserCost
                                                ? ((line.EffectiveLineCost || 0) / (line.NoFin || 0))
                                                : (hasNestingCost ? (line.NestingCost || 0) : (line.CCstPr || 0)))
                                            : (line.CCstPr || 0);
                                        const nestingSamlet = isLaserLine
                                            ? (hasEffectiveLaserCost
                                                ? (line.EffectiveLineCost || 0)
                                                : (hasNestingCost
                                                    ? ((line.NestingCost || 0) * (line.NoFin || 0))
                                                    : (line.LineCost || 0)))
                                            : (line.LineCost || 0);
                                        html += '<td>' + formatNumber(nestingUnitCost) + '</td>';
                                        html += '<td><strong>' + formatNumber(nestingSamlet) + '</strong></td>';
                                    } else if (key !== '1') {
                                        const displayQtyNonOperation = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                                            ? Number(line.DisplayQuantity || 0)
                                            : Number(line.NoFin || 0);
                                        const displayUnitCost = (displayQtyNonOperation > 0 && line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null)
                                            ? ((line.EffectiveLineCost || 0) / displayQtyNonOperation)
                                            : ((line.DisplayUnitCost !== undefined && line.DisplayUnitCost !== null)
                                                ? line.DisplayUnitCost
                                                : (line.CCstPr || line.DPrice || 0));
                                        const displayTotalCost = line.EffectiveLineCost !== undefined && line.EffectiveLineCost !== null
                                            ? (line.EffectiveLineCost || 0)
                                            : (line.LineCost || 0);
                                        html += '<td>' + formatNumber(displayUnitCost) + '</td>';
                                        html += '<td><strong>' + formatNumber(displayTotalCost) + '</strong></td>';
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
                    
                    reportOriginState = null;
                    openOrderDetailModal(
                        html,
                        'Ordre ' + data.orderHeader.OrdNo + ' - ' + (data.orderHeader.CustomerName || '-'),
                        'Produktion, cost og sporbarhed i en separat rapportvisning'
                    );
                    result.innerHTML = '';
                    loadSalesOrderLaserSummary(data);
                    loadSalesOrderOperationSummary(data);
                } catch (err) {
                    if (requestId !== activeSearchRequestId) return;
                    openOrderDetailModal(
                        '<div class="section"><div class="error">Fejl: ' + escapeHtml(String(err.message || err)) + '</div></div>',
                        'Ordre ' + escapeHtml(String(ordNo)) + ' - fejl',
                        'Der opstod en fejl under indlæsning'
                    );
                    result.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                }
            }

            function loadSalesOrderOperationSummary(orderData) {
                const body = document.getElementById('operationOrderSummaryBody');
                const totals = document.getElementById('operationOrderSummaryTotals');
                const teaser = document.getElementById('operationOversigtSummaryTeaser');
                if (!body || !totals) return;

                try {
                    const productionOrders = Array.isArray(orderData && orderData.productionOrders) ? orderData.productionOrders : [];
                    const groupedRows = new Map();
                    let totalOperationCost = 0;
                    let totalStyklisteMinutes = 0;
                    let totalFinishedMinutes = 0;

                    for (const prodOrder of productionOrders) {
                        const lines = Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : [];

                        for (const line of lines) {
                            const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                            const lnNo = Number((line && line.LnNo) || 0);
                            if (lnNo === 1 || key !== '1') continue;

                            const prodNo = String((line && line.ProdNo) || '').trim();
                            if (!prodNo) continue;
                            const totalCost = Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                            const qty = Number((line && (line.DisplayQuantity ?? line.NoFin)) || 0);
                            const styklisteMinutes = Number((line && line.NoOrg) || 0);
                            const effectiveMinutes = Number((line && (line.EffectiveOperationMinutes ?? line.NoFin)) || 0);

                            totalOperationCost += totalCost;
                            totalStyklisteMinutes += styklisteMinutes;
                            totalFinishedMinutes += effectiveMinutes;
                            if (!groupedRows.has(prodNo)) {
                                groupedRows.set(prodNo, {
                                    prodNo,
                                    descr: String((line && line.Descr) || '').trim(),
                                    styklisteQty: 0,
                                    qty: 0,
                                    totalCost: 0,
                                    occurrences: 0
                                });
                            }

                            const group = groupedRows.get(prodNo);
                            group.styklisteQty += Number((line && line.NoOrg) || 0);
                            group.qty += qty;
                            group.totalCost += totalCost;
                            group.occurrences += 1;
                            if ((!group.descr || group.descr === '-') && line && line.Descr) {
                                group.descr = String(line.Descr).trim();
                            }
                        }
                    }

                    const rows = Array.from(groupedRows.values()).sort((a, b) => a.prodNo.localeCompare(b.prodNo));

                    if (rows.length === 0) {
                        body.innerHTML = '<div>Ingen operationer fundet for denne salgsordre.</div>';
                        totals.innerHTML = '<div><strong>Samlet Operation kost:</strong> 0,00 DKK</div><div><strong>Ordre stykliste minutter:</strong> 0,00</div><div><strong>Ordre færdigmeldt minutter:</strong> 0,00</div><div><strong>Afvigelse minutter:</strong> 0,00</div><div><strong>Samlet afvigelse %:</strong> NULL</div>';
                        if (teaser) teaser.innerHTML = '<strong>Ingen operationer fundet</strong>';
                        if (currentOversigtModalType === 'operation') buildOversigtModalView('operation');
                        return;
                    }

                    let html = '<table class="oversigt-table-operation">';
                    html += '<tr><th>Operation</th><th>Beskrivelse</th><th>Linjer</th><th>Stykliste</th><th>Færdig</th><th>Kost/enh.</th><th>Samlet</th></tr>';
                    for (const row of rows) {
                        const unitCost = row.qty > 0 ? (row.totalCost / row.qty) : 0;
                        html += '<tr>';
                        html += '<td>' + (row.prodNo || '-') + '</td>';
                        html += '<td>' + (row.descr || '-') + '</td>';
                        html += '<td>' + formatNumber(row.occurrences || 0) + '</td>';
                        html += '<td>' + formatNumber(row.styklisteQty || 0) + '</td>';
                        html += '<td>' + formatNumber(row.qty || 0) + '</td>';
                        html += '<td>' + formatNumber(unitCost || 0) + '</td>';
                        html += '<td><strong>' + formatNumber(row.totalCost || 0) + '</strong></td>';
                        html += '</tr>';
                    }
                    html += '</table>';
                    body.innerHTML = html;
                    const deltaMinutes = totalFinishedMinutes - totalStyklisteMinutes;
                    const deltaPct = totalStyklisteMinutes > 0
                        ? ((deltaMinutes / totalStyklisteMinutes) * 100)
                        : null;
                    totals.innerHTML = ''
                        + '<div><strong>Samlet Operation kost:</strong> ' + formatNumber(totalOperationCost) + ' DKK</div>'
                        + '<div><strong>Ordre stykliste minutter:</strong> ' + formatNumber(totalStyklisteMinutes) + '</div>'
                        + '<div><strong>Ordre færdigmeldt minutter:</strong> ' + formatNumber(totalFinishedMinutes) + '</div>'
                        + '<div><strong>Afvigelse minutter:</strong> ' + formatNumber(deltaMinutes) + '</div>'
                        + '<div><strong>Samlet afvigelse %:</strong> ' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + '</div>';
                    if (teaser) {
                        teaser.innerHTML = ''
                            + '<div><strong>' + formatNumber(totalOperationCost) + ' DKK</strong> samlet operation kost</div>'
                            + '<div>Afvigelse: <strong>' + formatNumber(deltaMinutes) + ' min</strong> (' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + ')</div>';
                    }
                    if (currentOversigtModalType === 'operation') buildOversigtModalView('operation');
                } catch (err) {
                    body.innerHTML = '<div class="error">Fejl operationsoversigt: ' + err.message + '</div>';
                    totals.innerHTML = '<div class="error">Fejl i samlet operationsoversigt: ' + err.message + '</div>';
                    if (teaser) teaser.innerHTML = '<strong>Fejl i operation KPI:</strong> ' + err.message;
                }
            }

            async function loadSalesOrderLaserSummary(orderData) {
                const body = document.getElementById('laserOrderSummaryBody');
                const totals = document.getElementById('laserOrderSummaryTotals');
                const teaser = document.getElementById('laserOversigtSummaryTeaser');
                if (!body || !totals) return;
                // NOTE: Preserve existing Gr4 branching exactly; Gr4=3 indicates Multiordre type.
                const orderGr4 = Number((orderData && orderData.orderHeader && orderData.orderHeader.Gr4) || currentSalesOrderGr4 || 0);

                try {
                    const requests = [];
                    const targetDedupe = new Set();
                    const visitedProdOrders = new Set();
                    const productionOrderLinesByOrdNo = new Map();
                    const productionOrders = Array.isArray(orderData.productionOrders) ? orderData.productionOrders : [];
                    const laserTargets = [];

                    function getOperationCostFromLines(lines) {
                        let total = 0;
                        for (const line of (Array.isArray(lines) ? lines : [])) {
                            const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
                            const lnNo = Number((line && line.LnNo) || 0);
                            if (lnNo === 1 || key === '0' || key === '3' || key === '5') continue;
                            if (key !== '1') continue;
                            total += Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
                        }
                        return total;
                    }

                    function addLaserTarget(targetOrdNo, prodNo, nestingCost) {
                        const cleanedOrdNo = Number(targetOrdNo || 0);
                        const cleanedProdNo = String(prodNo || '').trim();
                        const cleanedNestingCost = Number(nestingCost || 0);
                        if (!cleanedOrdNo || !cleanedProdNo) return;

                        if (cleanedNestingCost > 0) {
                            setLaserNestCostHint(cleanedOrdNo, cleanedProdNo, cleanedNestingCost);
                        }

                        const key = cleanedOrdNo + '|' + cleanedProdNo;
                        if (targetDedupe.has(key)) return;
                        targetDedupe.add(key);
                        laserTargets.push({ ordNo: cleanedOrdNo, prodNo: cleanedProdNo, nestingCost: cleanedNestingCost > 0 ? cleanedNestingCost : null });
                    }

                    function collectLaserTargetsFromLines(sourceOrdNo, lines) {
                        const childOrdNos = [];
                        for (const line of (Array.isArray(lines) ? lines : [])) {
                            const key = (line.ProdTp4 === null || line.ProdTp4 === undefined) ? 'NA' : String(line.ProdTp4);
                            const prodNo = String(line.ProdNo || '').trim();
                            if (key === '2' && isLaserLProdNo(prodNo)) {
                                addLaserTarget(sourceOrdNo, prodNo, line.NestingCost);
                            }
                            if (key === '4' && Number(line.PurcNo || 0) > 0) {
                                childOrdNos.push(Number(line.PurcNo || 0));
                            }
                        }
                        return childOrdNos;
                    }

                    async function fetchProductionSummarySafe(childOrdNo) {
                        try {
                            const response = await fetch('/production-summary/' + childOrdNo + (orderGr4 === 3 ? '?gr4=3' : ''));
                            const data = await response.json();
                            if (!response.ok || !data || data.error) return null;
                            return data;
                        } catch (_) {
                            return null;
                        }
                    }

                    const pendingChildOrdNos = [];

                    for (const prodOrder of productionOrders) {
                        const currentOrdNo = Number(prodOrder && prodOrder.ordNo || 0);
                        if (!currentOrdNo || visitedProdOrders.has(currentOrdNo)) continue;
                        visitedProdOrders.add(currentOrdNo);
                        productionOrderLinesByOrdNo.set(currentOrdNo, Array.isArray(prodOrder && prodOrder.lines) ? prodOrder.lines : []);
                        const discoveredChildOrdNos = collectLaserTargetsFromLines(currentOrdNo, prodOrder.lines);
                        for (const childOrdNo of discoveredChildOrdNos) {
                            if (!visitedProdOrders.has(childOrdNo)) {
                                pendingChildOrdNos.push(childOrdNo);
                            }
                        }
                    }

                    while (pendingChildOrdNos.length > 0) {
                        const childOrdNo = Number(pendingChildOrdNos.shift() || 0);
                        if (!childOrdNo || visitedProdOrders.has(childOrdNo)) continue;
                        visitedProdOrders.add(childOrdNo);

                        const childSummary = await fetchProductionSummarySafe(childOrdNo);
                        if (!childSummary) continue;

                        productionOrderLinesByOrdNo.set(childOrdNo, Array.isArray(childSummary.lines) ? childSummary.lines : []);

                        const discoveredChildOrdNos = collectLaserTargetsFromLines(childOrdNo, childSummary.lines);
                        for (const nestedOrdNo of discoveredChildOrdNos) {
                            if (!visitedProdOrders.has(nestedOrdNo)) {
                                pendingChildOrdNos.push(nestedOrdNo);
                            }
                        }
                    }

                    for (const target of laserTargets) {
                        const endpoint = '/laser-route-metrics?ordine=' + encodeURIComponent(target.ordNo)
                            + '&prodNo=' + encodeURIComponent(target.prodNo)
                            + '&showAllRoutes=1'
                            + (orderGr4 === 3 ? '&gr4=3' : '');

                        requests.push(
                            fetch(endpoint)
                                .then(r => r.json().then(data => ({ ok: r.ok, data })))
                                .then(({ ok, data }) => ({ ok, data, prodOrderNo: target.ordNo, requestedProdNo: target.prodNo, requestedRoute: null, requestedNestingCost: target.nestingCost }))
                                .catch(() => null)
                        );
                    }

                    if (requests.length === 0) {
                        body.innerHTML = '<div>Ingen L-linjer fundet for denne salgsordre.</div>';
                        totals.innerHTML = '<div><strong>Samlet L-kost (NestKost):</strong> 0,00 DKK</div><div><strong>Ordre stykliste kg:</strong> 0,00 kg</div><div><strong>Ordre forbrugt kg:</strong> 0,00 kg</div><div><strong>Afvigelse kg:</strong> 0,00 kg</div><div><strong>Samlet afvigelse %:</strong> NULL</div>';
                        if (teaser) teaser.innerHTML = '<strong>Ingen laserlinjer fundet</strong>';
                        if (currentOversigtModalType === 'laser') buildOversigtModalView('laser');
                        return;
                    }

                    const results = await Promise.all(requests);
                    const rows = [];

                    for (const item of results) {
                        if (!item || !item.ok || !item.data || item.data.error) continue;
                        const products = Array.isArray(item.data.products) ? item.data.products : [];
                        for (const p of products) {
                            const expected = p.NWgtU_medio;
                            const effective = p.KgPerPezzoEffettivo;
                            const hintedNestCost = getLaserNestCostHint(item.prodOrderNo, p.ProdNo);
                            const hasHintedNestCost = hintedNestCost !== null && hintedNestCost !== undefined && Number(hintedNestCost) > 0;
                            const routeSpecificCostPerPiece = hasHintedNestCost
                                ? hintedNestCost
                                : ((p.CostoPerPezzo !== null && p.CostoPerPezzo !== undefined)
                                    ? p.CostoPerPezzo
                                    : null);
                            const extraPct = (expected !== null && expected !== undefined && expected > 0 && effective !== null && effective !== undefined)
                                ? (((effective - expected) / expected) * 100)
                                : null;

                            rows.push({
                                prodOrderNo: item.prodOrderNo,
                                nestingOrdNo: p.NestingOrdNo || item.data.nestingOrdNo,
                                prodNo: p.ProdNo,
                                route: p.Route || item.data.route || item.requestedRoute,
                                noFin: p.QtaPezzi,
                                oldNWgtU_medio: p.OldNWgtU_medio,
                                expected,
                                effective,
                                costPerPiece: routeSpecificCostPerPiece,
                                quotaCost: p.QuotaCosto,
                                extraPct,
                                imageItems: Array.isArray(p.ImageItems) ? p.ImageItems : []
                            });
                        }
                    }

                    if (rows.length === 0) {
                        body.innerHTML = '<div>Ingen laserberegninger tilgaengelige for denne salgsordre.</div>';
                        totals.innerHTML = '<div><strong>Samlet L-kost (NestKost):</strong> 0,00 DKK</div><div><strong>Ordre stykliste kg:</strong> 0,00 kg</div><div><strong>Ordre forbrugt kg:</strong> 0,00 kg</div><div><strong>Afvigelse kg:</strong> 0,00 kg</div><div><strong>Samlet afvigelse %:</strong> NULL</div>';
                        if (teaser) teaser.innerHTML = '<strong>Ingen laserdata klar</strong>';
                        if (currentOversigtModalType === 'laser') buildOversigtModalView('laser');
                        return;
                    }

                    const multiNestHeader = orderGr4 === 3 ? 'NestMultiPris' : 'NestKost pr. stk';
                    let html = '<table class="oversigt-table-laser">';
                    html += '<tr><th>Prod.ordre</th><th>Nesting</th><th>Produkt</th><th>Rute</th><th>Færdig</th><th>Icon kg</th><th>Stykl. kg</th><th>Forbrug kg</th><th>' + (orderGr4 === 3 ? 'Multi/stk' : 'Nest/stk') + '</th><th>Samlet</th><th>Afvig. %</th><th>Vis</th></tr>';
                    let totalKgUtilizzati = 0;
                    let totalKgPrevisti = 0;
                    let totalKgIcon = 0;
                    let totalLaserCost = 0;
                    for (const r of rows) {
                        const rowNoFin = Number(r.noFin || 0);
                        const rowIcon = Number(r.oldNWgtU_medio || 0);
                        const rowExpected = Number(r.expected || 0);
                        const rowEffective = Number(r.effective || 0);
                        const rowCostPerPiece = Number(r.costPerPiece || 0);
                        const rowTotalCost = (r.costPerPiece !== null && r.costPerPiece !== undefined && rowNoFin > 0)
                            ? (rowNoFin * rowCostPerPiece)
                            : ((r.costPerPiece === null || r.costPerPiece === undefined || r.noFin === null || r.noFin === undefined)
                                ? null
                                : (rowNoFin * rowCostPerPiece));
                        totalKgIcon += rowNoFin * rowIcon;
                        totalKgPrevisti += rowNoFin * rowExpected;
                        totalKgUtilizzati += rowNoFin * rowEffective;
                        totalLaserCost += rowTotalCost || 0;

                        html += '<tr>';
                        html += '<td>' + (r.prodOrderNo || '-') + '</td>';
                        html += '<td>' + (r.nestingOrdNo || '-') + '</td>';
                        html += '<td>' + (r.prodNo || '-') + '</td>';
                        html += '<td>' + (r.route || '-') + '</td>';
                        html += '<td>' + (r.noFin === null || r.noFin === undefined ? 'NULL' : formatNumber(r.noFin)) + '</td>';
                        html += '<td>' + (r.oldNWgtU_medio === null || r.oldNWgtU_medio === undefined ? 'NULL' : formatNumber(r.oldNWgtU_medio)) + '</td>';
                        html += '<td>' + (r.expected === null || r.expected === undefined ? 'NULL' : formatNumber(r.expected)) + '</td>';
                        html += '<td>' + (r.effective === null || r.effective === undefined ? 'NULL' : formatNumber(r.effective)) + '</td>';
                        html += '<td>' + (r.costPerPiece === null || r.costPerPiece === undefined ? 'NULL' : formatNumber(r.costPerPiece)) + '</td>';
                        html += '<td>' + (rowTotalCost === null ? 'NULL' : formatNumber(rowTotalCost)) + '</td>';
                        html += '<td>' + (r.extraPct === null || r.extraPct === undefined ? 'NULL' : (formatNumber(r.extraPct) + '%')) + '</td>';
                        if (Array.isArray(r.imageItems) && r.imageItems.length > 0) {
                            const imageKey = registerSummaryImageData('Billeder for ' + (r.prodNo || 'produkt') + ' / rute ' + (r.route || '-'), r.imageItems);
                            html += '<td><button class="image-preview-btn" data-image-mode="compact" data-image-key="' + imageKey + '">Vis</button></td>';
                        } else {
                            html += '<td>-</td>';
                        }
                        html += '</tr>';
                    }
                    const deltaKg = totalKgUtilizzati - totalKgPrevisti;
                    const deltaPct = totalKgPrevisti > 0
                        ? ((deltaKg / totalKgPrevisti) * 100)
                        : null;
                    html += '</table>';
                    body.innerHTML = html;
                    applyMicroTablePolish(body);
                    totals.innerHTML = ''
                        + '<div><strong>Samlet L-kost (' + (orderGr4 === 3 ? 'NestMultiPris' : 'NestKost') + '):</strong> ' + formatNumber(totalLaserCost) + ' DKK</div>'
                        + '<div><strong>Ordre icon kg:</strong> ' + formatNumber(totalKgIcon) + ' kg</div>'
                        + '<div><strong>Ordre stykliste kg:</strong> ' + formatNumber(totalKgPrevisti) + ' kg</div>'
                        + '<div><strong>Ordre forbrugt kg:</strong> ' + formatNumber(totalKgUtilizzati) + ' kg</div>'
                        + '<div><strong>Afvigelse kg:</strong> ' + formatNumber(deltaKg) + ' kg</div>'
                        + '<div><strong>Samlet afvigelse %:</strong> ' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + '</div>';
                    if (teaser) {
                        teaser.innerHTML = ''
                            + '<div><strong>' + formatNumber(totalLaserCost) + ' DKK</strong> samlet L-kost</div>'
                            + '<div>Afvigelse: <strong>' + formatNumber(deltaKg) + ' kg</strong> (' + (deltaPct === null ? 'NULL' : (formatNumber(deltaPct) + '%')) + ')</div>';
                    }
                    if (currentOversigtModalType === 'laser') buildOversigtModalView('laser');
                } catch (err) {
                    body.innerHTML = '<div class="error">Fejl laseroversigt: ' + err.message + '</div>';
                    totals.innerHTML = '<div class="error">Fejl i samlet laseroversigt: ' + err.message + '</div>';
                    if (teaser) teaser.innerHTML = '<strong>Fejl i laser KPI:</strong> ' + err.message;
                }
            }

            async function onProductClick(prodNo, ordNo, lnNo, prodTp4, trInf2, trInf4, showAllRoutes, clickedNoFin, clickedNestingCost) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');

                const modalWasOpen = modal.style.display === 'flex';
                if (modalWasOpen) {
                    pushSummaryModalState();
                } else {
                    summaryModalHistory = [];
                    updateSummaryModalBackBtn();
                }

                closeSummaryImagePanel();

                title.textContent = 'Produkt: ' + prodNo;
                modal.style.display = 'flex';

                if (String(prodTp4) === '1') {
                    body.innerHTML = '<div class="modal-loading">Indlæser transaktioner...</div>';
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
                        html += '<tr><th>Færdigmeldingsdato</th><th>Færdigmeldingstid</th><th>Minutter</th><th>Hvem</th></tr>';
                        for (const r of rows) {
                            const rawFinDt = String(r.FinDt || '').trim();
                            const compactFinDt = rawFinDt.split('T')[0].replace(/-/g, '');
                            let finDt = '-';
                            if (/^\d{8}$/.test(compactFinDt)) {
                                finDt = compactFinDt.slice(6, 8) + '-' + compactFinDt.slice(4, 6) + '-' + compactFinDt.slice(0, 4);
                            } else if (rawFinDt) {
                                finDt = rawFinDt;
                            }
                            const rawFinTm = r.FinTm != null ? String(r.FinTm).trim() : '';
                            const finTm = rawFinTm
                                ? rawFinTm.padStart(4, '0').replace(/^(\d{2})(\d{2})$/, '$1:$2')
                                : '-';
                            html += '<tr>';
                            html += '<td>' + finDt + '</td>';
                            html += '<td>' + finTm + '</td>';
                            html += '<td>' + formatNumber(r.NoInvoAb || 0) + '</td>';
                            html += '<td>' + (r.HvemNm || '-') + '</td>';
                            html += '</tr>';
                        }
                        html += '</table>';
                        body.innerHTML = html;
                        applyMicroTablePolish(body);
                    } catch (err) {
                        body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                    }
                } else if (String(prodTp4) === '2') {
                    body.innerHTML = '<div class="modal-loading">Indlaeser ruteberegning...</div>';
                    try {
                        const effectiveOrdine = String(ordNo || trInf2 || '').trim();
                        let effectiveRoute = String(trInf4 || '').trim();

                        if (!effectiveOrdine) {
                            body.innerHTML = '<div class="error">Fejl: OrdNo/TrInf2 mangler paa den valgte linje.</div>';
                            return;
                        }

                        if (!showAllRoutes && !effectiveRoute) {
                            const encProdNo = encodeURIComponent(prodNo || '');
                            const fallbackResponse = await fetch('/nesting-detail/' + encodeURIComponent(effectiveOrdine) + '/' + encProdNo);
                            const fallbackRows = await fallbackResponse.json();
                            if (fallbackResponse.ok && Array.isArray(fallbackRows) && fallbackRows.length > 0) {
                                effectiveRoute = String(fallbackRows[0].TrInf4 || '').trim();
                            }
                        }

                        if (!showAllRoutes && !effectiveRoute) {
                            body.innerHTML = '<div class="error">Fejl: TrInf4 (route) mangler paa den valgte linje.</div>';
                            return;
                        }

                        const endpoint = buildLaserRouteMetricsEndpoint(effectiveOrdine, effectiveRoute, prodNo, showAllRoutes);
                        const data = await requestRouteMetricsData(endpoint);

                        const finalData = data;
                        const usedProdFilter = Boolean(prodNo);

                        const s = finalData.summary || {};
                        const products = Array.isArray(finalData.products) ? finalData.products : [];
                        const formatNullable = (value, suffix = '') => {
                            return value === null || value === undefined
                                ? 'NULL'
                                : (formatNumber(value) + suffix);
                        };

                        if (!products.length) {
                            body.innerHTML = usedProdFilter
                                ? '<div>Ingen faerdigvarer (TrTp=7) fundet for valgt produkt/route.</div>'
                                : '<div>Ingen faerdigvarer (TrTp=7) fundet for valgt rute.</div>';
                            return;
                        }

                        const multiNestHeader = currentSalesOrderGr4 === 3 ? 'NestMultiPris' : 'NestKost pr. stk';
                        let html = '<table>';
                        html += '<tr><th>Nestingordre</th><th>Produkt</th><th>Rute</th><th>Færdigmeldt</th><th>Icon vægt (kg/stk)</th><th>Stykliste vaegt (kg/stk)</th><th>Forbrugt (kg/stk)</th><th>' + multiNestHeader + '</th><th>Samlet kost</th><th>Afvigelse (%)</th><th>Billeder</th></tr>';
                        let totalKgPrevisti = 0;
                        let totalKgUtilizzati = 0;
                        let totalKgIcon = 0;
                        let totalLaserCost = 0;
                        const clickedNoFinNum = Number(clickedNoFin || 0);
                        const clickedNestingCostNum = Number(clickedNestingCost || 0);

                        for (const rowProduct of products) {
                            const oldExpected = rowProduct ? rowProduct.OldNWgtU_medio : null;
                            const expected = rowProduct ? rowProduct.NWgtU_medio : null;
                            const effective = rowProduct ? rowProduct.KgPerPezzoEffettivo : null;
                            const routeNoFin = rowProduct ? rowProduct.QtaPezzi : null;
                            const prodNoForCost = rowProduct ? (rowProduct.ProdNo || prodNo) : prodNo;
                            const isClickedProd = String(prodNoForCost || '').trim().toUpperCase() === String(prodNo || '').trim().toUpperCase();
                            const hasClickedNestCost = !showAllRoutes && isClickedProd && clickedNestingCostNum > 0;
                            const noFin = (!showAllRoutes && hasClickedNestCost && clickedNoFinNum > 0) ? clickedNoFinNum : routeNoFin;
                            const hintedNestCost = getLaserNestCostHint(effectiveOrdine, prodNoForCost);
                            const costPerPiece = hasClickedNestCost
                                ? clickedNestingCostNum
                                : ((rowProduct && rowProduct.CostoPerPezzo !== null && rowProduct.CostoPerPezzo !== undefined)
                                    ? rowProduct.CostoPerPezzo
                                    : hintedNestCost);
                            const noFinNum = Number(noFin || 0);
                            const expectedNum = Number(expected || 0);
                            const effectiveNum = Number(effective || 0);
                            const baseTotalCost = hasClickedNestCost
                                ? (noFinNum > 0 ? (noFinNum * Number(costPerPiece || 0)) : null)
                                : ((rowProduct && rowProduct.QuotaCosto !== null && rowProduct.QuotaCosto !== undefined)
                                    ? rowProduct.QuotaCosto
                                : ((costPerPiece === null || costPerPiece === undefined || noFin === null || noFin === undefined)
                                    ? null
                                    : (noFinNum * Number(costPerPiece || 0))));
                            let totalCost = baseTotalCost;
                            let displayCostPerPiece = costPerPiece;
                            totalKgIcon += noFinNum * Number(oldExpected || 0);
                            totalKgPrevisti += noFinNum * expectedNum;
                            totalKgUtilizzati += noFinNum * effectiveNum;
                            totalLaserCost += totalCost || 0;
                            const extraPct = (expected !== null && expected !== undefined && expected > 0 && effective !== null && effective !== undefined)
                                ? (((effective - expected) / expected) * 100)
                                : null;

                            html += '<tr>';
                            html += '<td>' + ((rowProduct && rowProduct.NestingOrdNo) || finalData.nestingOrdNo || '-') + '</td>';
                            html += '<td>' + (rowProduct ? (rowProduct.ProdNo || '-') : '-') + '</td>';
                            html += '<td>' + (rowProduct ? (rowProduct.Route || '-') : (finalData.route || '-')) + '</td>';
                            html += '<td>' + formatNullable(noFin) + '</td>';
                            html += '<td>' + formatNullable(oldExpected) + '</td>';
                            html += '<td>' + formatNullable(expected) + '</td>';
                            html += '<td>' + formatNullable(effective) + '</td>';
                            html += '<td>' + formatNullable(displayCostPerPiece) + '</td>';
                            html += '<td>' + formatNullable(totalCost) + '</td>';
                            html += '<td>' + (extraPct === null ? 'NULL' : (formatNumber(extraPct) + '%')) + '</td>';
                            if (Array.isArray(rowProduct.ImageItems) && rowProduct.ImageItems.length > 0) {
                                const imageKey = registerSummaryImageData('Billeder for ' + (rowProduct.ProdNo || 'produkt') + ' / rute ' + (rowProduct.Route || '-'), rowProduct.ImageItems);
                                html += '<td><button class="image-preview-btn" data-image-mode="compact" data-image-key="' + imageKey + '">Vis</button></td>';
                            } else {
                                html += '<td>-</td>';
                            }
                            html += '</tr>';
                        }
                        html += '</table>';
                        const displayedLaserTotal = (showAllRoutes && clickedNoFinNum > 0 && clickedNestingCostNum > 0)
                            ? (clickedNoFinNum * clickedNestingCostNum)
                            : totalLaserCost;
                        html += '<div class="summary-box" style="margin-top:12px;">'
                            + '<div><strong>Samlet L-kost (NestKost):</strong> ' + formatNumber(displayedLaserTotal) + ' DKK</div>'
                            + '<div><strong>Ordre icon kg:</strong> ' + formatNumber(totalKgIcon) + ' kg</div>'
                            + '<div><strong>Ordre stykliste kg:</strong> ' + formatNumber(totalKgPrevisti) + ' kg</div>'
                            + '<div><strong>Ordre forbrugt kg:</strong> ' + formatNumber(totalKgUtilizzati) + ' kg</div>'
                            + '</div>';
                        body.innerHTML = html;
                        applyMicroTablePolish(body);
                    } catch (err) {
                        body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                    }
                }
            }

            function handleProdNoClick(e) {
                const span = e.target.closest('.prod-no-link');
                if (!span) return;
                const prodNo = span.dataset.prodno;
                const ordNo = span.dataset.ordno;
                const lnNo = span.dataset.lnno;
                const prodTp4 = span.dataset.prodtp4;
                const trInf2 = span.dataset.trinf2;
                const trInf4 = span.dataset.trinf4;
                const showAllRoutes = span.dataset.showallroutes === '1';
                const noFin = span.dataset.nofin;
                const nestingCost = span.dataset.nestingcost;
                if (prodNo) onProductClick(prodNo, ordNo, lnNo, prodTp4, trInf2, trInf4, showAllRoutes, noFin, nestingCost);
            }

            function handleProdNoHover(e) {
                const span = e.target.closest('.prod-no-link');
                if (!span) return;
                if (span.dataset.prodtp4 !== '2') return;
                if (span.dataset.routePrefetchStarted === '1') return;
                span.dataset.routePrefetchStarted = '1';
                prefetchRouteMetricsForProduct(
                    span.dataset.prodno,
                    span.dataset.ordno,
                    span.dataset.trinf2,
                    span.dataset.trinf4,
                    span.dataset.showallroutes === '1'
                );
            }

            function handleImagePreviewClick(e) {
                const btn = e.target.closest('.image-preview-btn');
                if (!btn) return;
                const imageKey = btn.dataset.imageKey;
                if (imageKey) {
                    if (btn.dataset.imageMode === 'compact') {
                        openCompactImageModal(imageKey);
                        return;
                    }
                    const inLaserPanel = !!e.target.closest('#laserOrderSummaryPanel');
                    const preferredPanelId = inLaserPanel ? 'laserImagePanel' : 'summaryImagePanel';
                    openSummaryImagePanel(imageKey, preferredPanelId);
                }
            }

            function handleDrawingOpenClick(e) {
                const btn = e.target.closest('.drawing-open-btn');
                if (!btn) return;
                openDrawingPdf({
                    path: btn.dataset.drawingPath || '',
                    prodNo: btn.dataset.prodNo || '',
                    ordNo: btn.dataset.ordNo || ''
                });
            }

            function handlePreviewImageZoom(e) {
                const image = e.target.closest('.image-preview-zoomable');
                if (!image) return;
                openImageLightbox(
                    image.dataset.fullsrc || image.getAttribute('src') || '',
                    image.dataset.title || image.getAttribute('alt') || 'Billede',
                    image.dataset.path || ''
                );
            }

            // Outside modal content.
            document.addEventListener('click', handleProdNoClick);
            document.addEventListener('mouseover', handleProdNoHover);
            document.addEventListener('focusin', handleProdNoHover);
            document.addEventListener('click', handleImagePreviewClick);
            document.addEventListener('click', handleDrawingOpenClick);
            document.addEventListener('click', handlePreviewImageZoom);
            document.addEventListener('keydown', function(event) {
                if (event.key === 'Escape') {
                    closeSideMenu();
                    closeImageLightbox();
                    closeCompactImageModal();
                    closeOversigtModal();
                }
            });
            // Inside modal content (document listener is blocked by modal stopPropagation).
            const summaryModalBodyEl = document.getElementById('summaryModalBody');
            if (summaryModalBodyEl) {
                summaryModalBodyEl.addEventListener('click', handleProdNoClick);
                summaryModalBodyEl.addEventListener('mouseover', handleProdNoHover);
                summaryModalBodyEl.addEventListener('focusin', handleProdNoHover);
                summaryModalBodyEl.addEventListener('click', handleImagePreviewClick);
                summaryModalBodyEl.addEventListener('click', handlePreviewImageZoom);
            }
            const summaryImagePanelEl = document.getElementById('summaryImagePanel');
            if (summaryImagePanelEl) {
                summaryImagePanelEl.addEventListener('click', handlePreviewImageZoom);
            }
            window.addEventListener('resize', updateSummaryImagePanelLayout);
            const oversigtModalBodyEl = document.getElementById('oversigtModalBody');
            if (oversigtModalBodyEl) {
                oversigtModalBodyEl.addEventListener('click', handleProdNoClick);
                oversigtModalBodyEl.addEventListener('click', handleImagePreviewClick);
                oversigtModalBodyEl.addEventListener('click', handleDrawingOpenClick);
                oversigtModalBodyEl.addEventListener('click', handlePreviewImageZoom);
            }
            const orderDetailModalBodyEl = document.getElementById('orderDetailModalBody');
            if (orderDetailModalBodyEl) {
                // Clicks inside the order detail modal do not bubble to document because the shell stops propagation.
                orderDetailModalBodyEl.addEventListener('click', handleProdNoClick);
                orderDetailModalBodyEl.addEventListener('click', handleImagePreviewClick);
                orderDetailModalBodyEl.addEventListener('click', handleDrawingOpenClick);
                orderDetailModalBodyEl.addEventListener('click', handlePreviewImageZoom);
            }
            const laserImagePanelEl = document.getElementById('laserImagePanel');
            if (laserImagePanelEl) {
                laserImagePanelEl.addEventListener('click', handlePreviewImageZoom);
            }

            function toggleSalesLineBreakdown(rowId, buttonEl) {
                const row = document.getElementById(rowId);
                if (!row) return;
                const isClosed = row.style.display === 'none';
                row.style.display = isClosed ? 'table-row' : 'none';
                if (buttonEl) buttonEl.textContent = isClosed ? '−' : '+';
            }

            function toggleProdTp4Group(orderNo, prodTp4Key) {
                const el = document.getElementById('po-' + orderNo + '-group-' + prodTp4Key);
                const icon = document.getElementById('po-' + orderNo + '-group-' + prodTp4Key + '-icon');
                if (!el) return;
                const isClosed = el.style.display === 'none';
                el.style.display = isClosed ? '' : 'none';
                if (icon) icon.textContent = isClosed ? '▾' : '▸';
            }

            async function showChildProductionSummary(childOrdNo, targetProdNo, forceInvoiceStatus) {
                const modal = document.getElementById('summaryModal');
                const title = document.getElementById('summaryModalTitle');
                const body = document.getElementById('summaryModalBody');

                const modalWasOpen = modal.style.display === 'flex';
                if (modalWasOpen) {
                    pushSummaryModalState();
                } else {
                    summaryModalHistory = [];
                    updateSummaryModalBackBtn();
                }
                closeSummaryImagePanel();
                const normalizedTargetProdNo = String(targetProdNo || '').trim();
                title.textContent = normalizedTargetProdNo
                    ? ('Produktoversigt for ordre ' + childOrdNo + ' - ' + normalizedTargetProdNo)
                    : ('Produktoversigt for ordre ' + childOrdNo);
                body.innerHTML = '<div class="modal-loading">Indlaeser...</div>';
                modal.style.display = 'flex';

                try {
                    const response = await fetch('/production-summary/' + childOrdNo + (currentSalesOrderGr4 === 3 ? '?gr4=3' : ''));
                    const data = await response.json();

                    if (!response.ok || data.error) {
                        body.innerHTML = '<div class="error">Fejl: ' + (data.error || 'Uventet fejl') + '</div>';
                        return;
                    }

                    if (!data.lines || data.lines.length === 0) {
                        body.innerHTML = '<div>Ingen linjer fundet for denne produktionsordre.</div>';
                        return;
                    }

                    const filteredLines = normalizedTargetProdNo
                        ? data.lines.filter(line => String(line && line.ProdNo || '').trim().toUpperCase() === normalizedTargetProdNo.toUpperCase())
                        : data.lines;

                    if (!filteredLines || filteredLines.length === 0) {
                        body.innerHTML = '<div>Det valgte produkt blev ikke fundet i denne produktionsordre.</div>';
                        return;
                    }

                    const baseTitleText = normalizedTargetProdNo
                        ? ('Produktoversigt for ordre ' + childOrdNo + ' - ' + normalizedTargetProdNo)
                        : ('Produktoversigt for ordre ' + childOrdNo);
                    const titleFlags = [
                        data.hasEstimatedOperationTime ? '🕒' : '',
                        data.hasWarnings ? '⚠️' : ''
                    ].filter(Boolean).join(' ');
                    title.textContent = titleFlags
                        ? (baseTitleText + ' ' + titleFlags)
                        : baseTitleText;

                    const isYdelseFilteredView = !!normalizedTargetProdNo;
                    const modalTotalCost = normalizedTargetProdNo
                        ? filteredLines.reduce((sum, line) => sum + Number(line && line.EffectiveLineCost || 0), 0)
                        : Number(data.totalCost || 0);

                    const shouldShowInvoiceStatus = Boolean(forceInvoiceStatus) || filteredLines.some(line => line && (line.IsInvoiceTracked || line.isInvoiceTracked || isInvoiceTrackedProdNo(line.ProdNo)));

                    let html = '';
                    html += getInvoiceStatusSummaryHtml(filteredLines, shouldShowInvoiceStatus);
                    html += isYdelseFilteredView
                        ? '<table><tr><th>Linje</th><th>ProdTp4</th><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Pris/enhed</th><th>Samlet kost (beregnet)</th></tr>'
                        : '<table><tr><th>Linje</th><th>ProdTp4</th><th>Prod</th><th>Beskrivelse</th><th>Færdigmeldt</th><th>Salgspris</th><th>Kostpris/enhed</th><th>Nesting/enhed</th><th>Samlet kost (beregnet)</th></tr>';
                    for (const line of filteredLines) {
                        const lineExcludedFromTotal = !isYdelseFilteredView && isProductionSummaryExcludedLine(line);
                        const displayLineCost = lineExcludedFromTotal
                            ? null
                            : Number(line.EffectiveLineCost || 0);
                        const warningFlagHtml = getWarningFlagHtml(line);
                        const invoiceStatusFlagHtml = getInvoiceStatusFlagHtml(line, shouldShowInvoiceStatus);
                        const timeAdjustmentFlagHtml = getTimeAdjustmentFlagHtml(line);
                        const laserAllocationFlagHtml = getLaserAllocationFlagHtml(line);
                        html += '<tr>';
                        html += '<td>' + (line.LnNo || 0) + '</td>';
                        const displayProdTp4 = getDisplayProdTp4Key(line.ProdTp4, line.ProdNo, line.PurcNo);
                        html += '<td>' + (displayProdTp4 === 'NA' ? '-' : displayProdTp4) + '</td>';
                        const childHasPurcNo = Number(line.PurcNo || 0) > 0;
                        if (String(line.ProdTp4 || '') === '1' && line.ProdNo) {
                            const safeChildProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const trInf2FromLine = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : childOrdNo);
                            const trInf4FromLine = String(line.TrInf4 || '');
                            const safeChildTrInf2 = trInf2FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const safeChildTrInf4 = trInf4FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            html += '<td><span class="prod-no-link" data-prodno="' + safeChildProdNo + '" data-ordno="' + childOrdNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="1" data-trinf2="' + safeChildTrInf2 + '" data-trinf4="' + safeChildTrInf4 + '">' + safeChildProdNo + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else if (line.ProdNo && String(line.ProdNo).trim().toUpperCase().endsWith('L')) {
                            const safeChildProdNo = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const trInf2FromLine = String((line.TrInf2 !== null && line.TrInf2 !== undefined && String(line.TrInf2).trim() !== '') ? line.TrInf2 : childOrdNo);
                            const trInf4FromLine = String(line.TrInf4 || '');
                            const safeChildTrInf2 = trInf2FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const safeChildTrInf4 = trInf4FromLine.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            html += '<td><span class="prod-no-link" data-prodno="' + safeChildProdNo + '" data-ordno="' + childOrdNo + '" data-lnno="' + (line.LnNo || 0) + '" data-prodtp4="2" data-trinf2="' + safeChildTrInf2 + '" data-trinf4="' + safeChildTrInf4 + '" data-showallroutes="1" data-nofin="' + Number(line.NoFin || 0) + '" data-nestingcost="' + Number(line.NestingCost || 0) + '">' + safeChildProdNo + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else if (childHasPurcNo) {
                            const safeChildProdNoForSummary = String(line.ProdNo || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;');
                            const childSummaryArgs = shouldFilterChildSummary(displayProdTp4, line.ProdNo, line.PurcNo)
                                ? (Number(line.PurcNo || 0) + ', &quot;' + safeChildProdNoForSummary + '&quot;, true')
                                : Number(line.PurcNo || 0);
                            html += '<td><span class="inline-link" onclick="showChildProductionSummary(' + childSummaryArgs + ')">' + (line.ProdNo || '-') + '</span>' + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        } else {
                            html += '<td>' + (line.ProdNo || '-') + invoiceStatusFlagHtml + laserAllocationFlagHtml + timeAdjustmentFlagHtml + warningFlagHtml + '</td>';
                        }
                        const displayQty = (line.DisplayQuantity !== undefined && line.DisplayQuantity !== null)
                            ? line.DisplayQuantity
                            : (line.NoFin || 0);
                        const displayUnitCost = (Number(displayQty || 0) > 0 && displayLineCost !== undefined && displayLineCost !== null)
                            ? ((displayLineCost || 0) / displayQty)
                            : ((line.DisplayUnitCost !== undefined && line.DisplayUnitCost !== null)
                                ? line.DisplayUnitCost
                                : (line.CCstPr || 0));
                        const isLaserProdLine = isLaserLProdNo(line.ProdNo);
                        html += '<td>' + (line.Descr || '') + '</td>';
                        html += '<td>' + formatNumber(displayQty) + '</td>';
                        if (isYdelseFilteredView) {
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                        } else {
                            html += '<td>' + formatNumber(line.DPrice || 0) + '</td>';
                            html += '<td>' + (isLaserProdLine ? '-' : formatNumber(displayUnitCost)) + '</td>';
                            const hasLaserAllocationSpread = Boolean(line.UsesLaserAllocationSpread || line.usesLaserAllocationSpread);
                            const allocationTitle = hasLaserAllocationSpread
                                ? ' title="Nesting-fordeling bruger et andet antal end ordrelinjen; pris pr. stk kan afvige."'
                                : '';
                            const allocationHint = hasLaserAllocationSpread ? ' <span style="color:#b26a00; font-weight:700;">*</span>' : '';
                            html += '<td' + allocationTitle + '>' + formatNumber(line.NestingCost || 0) + allocationHint + '</td>';
                        }
                        html += '<td><strong>' + (displayLineCost === null ? '-' : formatNumber(displayLineCost)) + '</strong></td>';
                        html += '</tr>';
                    }
                    html += isYdelseFilteredView
                        ? '<tr class="summary-row"><td colspan="6">Total beregnet kost:</td><td><strong>' + formatNumber(modalTotalCost || 0) + ' DKK</strong></td></tr>'
                        : '<tr class="summary-row"><td colspan="8">Total beregnet kost:</td><td><strong>' + formatNumber(modalTotalCost || 0) + ' DKK</strong></td></tr>';
                    html += '</table>';
                    body.innerHTML = html;
                    applyMicroTablePolish(body);
                } catch (err) {
                    body.innerHTML = '<div class="error">Fejl: ' + err.message + '</div>';
                }
            }

            function closeSummaryModal(event) {
                if (event && event.target && event.target.id !== 'summaryModal') return;
                const modal = document.getElementById('summaryModal');
                modal.style.display = 'none';
                summaryModalHistory = [];
                closeSummaryImagePanel();
                updateSummaryModalBackBtn();
            }

            function scrollToElementWithStickyOffset(el) {
                if (!el) return;
                const header = document.querySelector('.header-banner-wrapper');
                const searchBox = document.getElementById('searchBox');
                const headerH = header ? header.offsetHeight : 0;
                const searchH = searchBox ? searchBox.offsetHeight : 0;
                const extraGap = 14;
                const targetTop = window.pageYOffset + el.getBoundingClientRect().top - headerH - searchH - extraGap;
                window.scrollTo({ top: Math.max(targetTop, 0), behavior: 'auto' });
            }

            function syncStickyOffsets() {
                const root = document.documentElement;
                const header = document.querySelector('.header-banner-wrapper');
                const searchBox = document.getElementById('searchBox');

                const headerH = header ? Math.ceil(header.getBoundingClientRect().height || 0) : 0;
                const searchVisible = !!searchBox && getComputedStyle(searchBox).display !== 'none';
                const searchH = searchVisible ? Math.ceil(searchBox.getBoundingClientRect().height || 0) : 0;

                // Header height can increase on smaller viewports; keep sticky controls below it.
                const searchTop = Math.max(headerH + 8, 58);
                const tableTop = Math.max(headerH + (searchVisible ? searchH : 0) + 8, 0);

                root.style.setProperty('--search-sticky-top', String(searchTop) + 'px');
                root.style.setProperty('--table-sticky-top', String(tableTop) + 'px');
            }

            function openProduction(ordNo) {
                const el = document.getElementById('po-' + ordNo);
                if (!el) {
                    alert('Produktionsordre ' + ordNo + ' blev ikke fundet i de indlaeste resultater.');
                    return;
                }

                const modal = document.getElementById('orderDetailModal');
                const modalBody = document.getElementById('orderDetailModalBody');
                const modalIsOpen = modal && getComputedStyle(modal).display === 'flex';
                if (modalIsOpen && modalBody && modalBody.contains(el)) {
                    // When browsing inside the report modal, scroll the modal body instead of the page.
                    const top = Math.max(el.offsetTop - 16, 0);
                    modalBody.scrollTo({ top, behavior: 'auto' });
                } else {
                    scrollToElementWithStickyOffset(el);
                }
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
                    el.innerHTML = '<div class="order-list-section"><h3>Ingen kunder fundet</h3><div>Prøv en anden søgning.</div></div>';
                    return;
                }

                let html = '<div class="order-list-section">';
                html += '<div id="orderListSummary" class="order-list-summary">';
                html += buildOrderListSummaryHtml(orders);
                html += '</div>';
                html += '<h3>Seneste fakturerede ordrer (' + ORDER_LIST_DAYS_BACK_CLIENT + ' dage) &mdash; ' + orders.length + ' af ' + orderListData.length + ' ordrer</h3>';
                const sortMark = (field) => {
                    if (orderListSortField !== field) return ' <span style="opacity:0.4;">^v</span>';
                    return orderListSortDir === 'asc'
                        ? ' <span style="color:#1976d2;">^</span>'
                        : ' <span style="color:#1976d2;">v</span>';
                };
                html += '<table class="order-list-table"><tr>';
                html += '<th class="order-sortable-header" data-sort-field="bruger" style="cursor:pointer; user-select:none;">Bruger' + sortMark('bruger') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="ordno" style="cursor:pointer; user-select:none;">Ordrenr.' + sortMark('ordno') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="kunde" style="cursor:pointer; user-select:none;">Kunde' + sortMark('kunde') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="date" style="cursor:pointer; user-select:none;">Fakturadato' + sortMark('date') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="belob" style="cursor:pointer; user-select:none;">Fakturabelob' + sortMark('belob') + '</th>';
                html += '<th class="order-sortable-header" data-sort-field="margin" style="cursor:pointer; user-select:none;">Margin' + sortMark('margin') + '</th>';
                html += '<th>Faktura</th>';
                html += '<th>Note</th>';
                html += '<th>Opdater</th>';
                html += '</tr>';
                for (const o of orders) {
                    const marginHtml = getOrderMarginHtml(o.OrdNo);

                    const d = String(o.LstInvDt || '');
                    const invDate = d.length === 8 ? d.slice(0,4) + '-' + d.slice(4,6) + '-' + d.slice(6,8) : (d || '-');
                    const orderWarningFlag = getWarningFlagHtml(o, 'Ordren indeholder mindst én advarsel.');
                    // NOTE: Visual badge only. No logic change: Gr4=3 => Multiordre order type.
                    const gr4TypeBadge = Number(o.Gr4 || 0) === 3
                        ? '<span title="Ordretype: Multiordre (Gr4=3)" aria-label="Ordretype: Multiordre" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:50%;background:#1565c0;color:#fff;font-size:11px;font-weight:700;margin-left:6px;vertical-align:middle;">M</span>'
                        : '';
                    html += '<tr data-ordno="' + o.OrdNo + '" class="order-list-row">'
                    html += '<td>' + (o.SellerUsr || '-') + '</td>';
                    html += '<td><strong>' + o.OrdNo + '</strong>' + gr4TypeBadge + orderWarningFlag + '</td>';
                    html += '<td>' + (o.CustomerName || '-') + '</td>';
                    html += '<td>' + invDate + '</td>';
                    html += '<td>' + formatNumber(o.InvoAm || 0) + ' DKK</td>';
                    html += '<td class="order-margin-cell" data-ordno="' + o.OrdNo + '">' + marginHtml + '</td>';
                    html += '<td class="order-invoice-cell" data-ordno="' + o.OrdNo + '">' + getOrderInvoiceStatusHtml(o.OrdNo) + '</td>';
                    html += '<td class="order-note-cell" data-ordno="' + o.OrdNo + '" onclick="event.stopPropagation();openNotePopup(' + o.OrdNo + ')">' + getOrderNoteHtml(o.OrdNo) + '</td>';
                    html += '<td class="order-refresh-cell"><button class="list-toggle-btn order-refresh-one-btn" data-ordno="' + o.OrdNo + '" style="padding:4px 8px; margin-left:0; background:#00695c !important;" title="Opdater cache for denne ordre">Opdater</button></td>';
                    html += '</tr>';
                }
                html += '</table>';

                html += '</div>';
                el.innerHTML = html;

                // Carica i margini in coda per tutti gli ordini visibili.
                const queuedOrders = orders.slice(0, MARGIN_PREFETCH_ROWS);
                queueMarginLoad(queuedOrders.map(o => o.OrdNo));
                updateSystemStatusFromOrders(queuedOrders);
            }

            async function loadOrderList(forceRefresh = false) {
                const el = document.getElementById('orderList');
                if (!el) return;

                const showOrderListError = (message) => {
                    el.innerHTML = '<div class="order-list-section"><h3>Ordreliste kunne ikke indlæses</h3><div>' + escapeHtml(message) + '</div><div style="margin-top:8px;"><button class="list-toggle-btn" onclick="refreshOrderList()">Prøv igen</button></div></div>';
                };

                if (orderListLoading && !forceRefresh) return;
                if (!forceRefresh && orderListData && orderListData.length > 0) {
                    renderOrderList();
                    return;
                }

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
                        } else {
                            showOrderListError('Serveren svarede med fejl (HTTP ' + response.status + ').');
                        }
                        return;
                    }
                    const orders = await response.json();
                    if (!orders || orders.error) {
                        setSystemStatus('System error', '#fdecea', '#b71c1c');
                        if (previousHtml) {
                            el.innerHTML = previousHtml;
                        } else {
                            showOrderListError((orders && orders.error) ? String(orders.error) : 'Ugyldigt svar fra serveren.');
                        }
                        return;
                    }
                    orderListData = orders;
                    hydrateMarginStateFromOrderList(orders);
                    populateBrugerFilterOptions();
                    loadAllNotes().then(() => {
                        if (orderListVisible) renderOrderList();
                    }).catch(() => {});
                    renderOrderList();
                    checkOrderListFreshness();
                } catch (err) {
                    console.error('Fejl i loadOrderList:', err);
                    setSystemStatus('System error', '#fdecea', '#b71c1c');
                    if (previousHtml) {
                        el.innerHTML = previousHtml;
                    } else {
                        showOrderListError(err && err.message ? err.message : 'Ukendt fejl.');
                    }
                } finally {
                    orderListLoading = false;
                }
            }

            function startOrderListAutoRefresh() {
                if (orderListAutoRefreshTimer) return;
                orderListAutoRefreshTimer = setInterval(() => {
                    if (document.hidden) return;
                    checkOrderListFreshness();
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

            async function refreshSingleOrderCache() {
                const ordNo = String(document.getElementById('orderInput').value || '').trim();
                if (!ordNo) {
                    alert('Indtast et ordrenummer foerst.');
                    return;
                }

                return refreshSingleOrderCacheByOrdNo(ordNo, true);
            }

            async function refreshSingleOrderCacheByOrdNo(ordNo, openAfter = false, clickedBtn = null) {
                const normalizedOrdNo = String(ordNo || '').trim();
                if (!normalizedOrdNo) return;
                const ordNoNum = Number(normalizedOrdNo);
                aftercalcClientCache.delete(normalizedOrdNo);
                let refreshSucceeded = false;

                const btn = document.getElementById('refreshSingleOrderBtn');
                if (btn) {
                    btn.disabled = true;
                    btn.textContent = 'Opdaterer ordre...';
                }
                if (clickedBtn) {
                    clickedBtn.disabled = true;
                    clickedBtn.textContent = '...';
                }

                try {
                    const r = await fetch('/cache-refresh-order/' + encodeURIComponent(normalizedOrdNo), { method: 'POST' });
                    const d = await r.json();
                    if (!r.ok || d.error) throw new Error((d && d.error) ? d.error : ('HTTP ' + r.status));

                    const startedAt = Date.now();
                    while (true) {
                        await new Promise(resolve => setTimeout(resolve, 1000));
                        const sr = await fetch('/cache-refresh-order-status/' + encodeURIComponent(normalizedOrdNo));
                        const sd = await sr.json();
                        if (sd && sd.status === 'done') {
                            break;
                        }
                        if (sd && sd.status === 'error') {
                            throw new Error(sd.error || 'Order refresh failed');
                        }
                        if (Date.now() - startedAt > 120000) {
                            throw new Error('Timeout waiting for order refresh');
                        }
                    }

                    await loadOrderList(true);
                    refreshSucceeded = true;
                    if (openAfter && Number.isFinite(ordNoNum)) {
                        await searchOrder();
                    } else {
                        const currentInputOrdNo = String((document.getElementById('orderInput') || {}).value || '').trim();
                        const detailModal = document.getElementById('orderDetailModal');
                        const detailOpen = detailModal && detailModal.style.display === 'flex';
                        if (detailOpen && currentInputOrdNo === normalizedOrdNo) {
                            await searchOrder();
                        }
                    }
                    if (clickedBtn) {
                        alert('Ordre ' + normalizedOrdNo + ' er opdateret fra kilden.');
                    }
                } catch (e) {
                    alert('Fejl ved ordre-cache opdatering: ' + e.message);
                } finally {
                    if (btn) {
                        btn.disabled = false;
                        btn.textContent = 'Opdater ordre-cache';
                    }
                    if (clickedBtn) {
                        clickedBtn.disabled = false;
                        clickedBtn.textContent = refreshSucceeded ? 'Opdateret' : 'Opdater';
                        if (refreshSucceeded) {
                            setTimeout(() => {
                                if (clickedBtn && clickedBtn.isConnected) clickedBtn.textContent = 'Opdater';
                            }, 1400);
                        }
                    }
                }
            }

            async function clearAppCache() {
                const confirmed = confirm('Er du sikker? Dette vil slette alt cache og tage lang tid at genindlæse data.');
                if (!confirmed) return;
                
                const btn = document.getElementById('clearCacheBtn');
                const dashBtn = document.getElementById('dashboardClearCacheBtn');
                if (btn) { btn.disabled = true; btn.textContent = 'Rydder...'; }
                if (dashBtn) { dashBtn.disabled = true; dashBtn.textContent = 'Rydder cache...'; }
                try {
                    showDashboardWarmupNotice = true;
                    warmupCombinedReady = false;
                    warmupCombinedPct = 0;
                    warmupCombinedDone = 0;
                    warmupCombinedTotal = 0;
                    const r = await fetch('/cache-clear', { method: 'POST' });
                    if (!r.ok) throw new Error('HTTP ' + r.status);
                    const d = await r.json();
                    alert('Cache ryddet: ' + (d.deleted || 0) + ' filer slettet.');
                } catch (e) {
                    alert('Fejl ved cache-rydning: ' + e.message);
                } finally {
                    if (btn) { btn.disabled = false; btn.textContent = 'Ryd cache'; }
                    if (dashBtn) { dashBtn.disabled = false; dashBtn.textContent = 'Ryd Efterkalk cache'; }
                }
            }

            async function checkDesktopUpdateNow() {
                const btn = document.getElementById('checkUpdateBtn');
                const dashBtn = document.getElementById('dashboardUpdateCheckBtn');
                if (btn) { btn.disabled = true; btn.textContent = 'Tjekker...'; }
                if (dashBtn) { dashBtn.disabled = true; dashBtn.textContent = 'Tjekker...'; }

                try {
                    const r = await fetch('/desktop-update-check', { method: 'POST' });
                    const d = await r.json();
                    if (!r.ok) throw new Error((d && d.message) ? d.message : ('HTTP ' + r.status));

                    // Show status only in dashboard update section (no popup alerts).
                    applyDashboardUpdateNotice({
                        status: d && d.status ? d.status : 'checking',
                        latestVersion: d && (d.latestVersion || d.version) ? (d.latestVersion || d.version) : undefined,
                        currentVersion: d && d.currentVersion ? d.currentVersion : undefined,
                        downloaded: d && d.downloaded === true,
                        canInstallNow: d && d.canInstallNow === true,
                        message: d && d.message ? d.message : 'Opdateringskontrol sendt.'
                    });
                } catch (e) {
                    applyDashboardUpdateNotice({ status: 'error', message: 'Fejl ved opdateringskontrol: ' + e.message });
                } finally {
                    if (btn) { btn.disabled = false; btn.textContent = 'Tjek opdatering nu'; }
                    if (dashBtn) { dashBtn.disabled = false; dashBtn.textContent = 'Tjek nu'; }
                    refreshDashboardUpdateNotice().catch(() => {});
                }
            }

            function applyDashboardUpdateNotice(state) {
                const wrap = document.getElementById('dashboardUpdateNotice');
                const titleEl = document.getElementById('dashboardUpdateTitle');
                const textEl = document.getElementById('dashboardUpdateText');
                const installBtn = document.getElementById('dashboardUpdateInstallBtn');
                if (!wrap || !titleEl || !textEl || !installBtn) return;

                const safe = state && typeof state === 'object' ? state : {};
                const status = String(safe.status || 'unavailable');
                const latest = safe.latestVersion ? String(safe.latestVersion) : null;
                const current = safe.currentVersion ? String(safe.currentVersion) : null;
                const canInstall = safe.canInstallNow === true || safe.downloaded === true || status === 'downloaded';

                let title = 'Programopdatering';
                let line = safe.message ? String(safe.message) : 'Ingen status endnu.';

                if (status === 'downloaded') {
                    title = 'Ny version klar';
                    line = 'Version ' + (latest || '?') + ' er hentet og klar til installation.';
                } else if (status === 'available') {
                    title = 'Ny version fundet';
                    line = 'Version ' + (latest || '?') + ' er fundet og hentes i baggrunden.';
                } else if (status === 'up-to-date') {
                    title = 'Programmet er opdateret';
                    line = current ? ('Aktuel version: ' + current + '.') : 'Du har den nyeste version.';
                } else if (status === 'checking' || status === 'busy') {
                    title = 'Søger efter opdateringer';
                    line = safe.message ? String(safe.message) : 'Tjekker...';
                } else if (status === 'unsupported' || status === 'unavailable') {
                    title = 'Opdatering ikke tilgængelig';
                } else if (status === 'installing') {
                    title = 'Installerer opdatering';
                } else if (status === 'error') {
                    title = 'Opdateringsfejl';
                }

                titleEl.textContent = title;
                textEl.textContent = line;
                wrap.classList.add('active');
                installBtn.style.display = canInstall ? 'inline-block' : 'none';
                installBtn.disabled = !canInstall;
            }

            async function refreshDashboardUpdateNotice() {
                const r = await fetch('/desktop-update-status');
                const d = await r.json();
                if (!r.ok) {
                    applyDashboardUpdateNotice({
                        status: 'error',
                        message: (d && d.message) ? d.message : ('HTTP ' + r.status)
                    });
                    return;
                }
                applyDashboardUpdateNotice(d || { status: 'unavailable', message: 'Ingen status.' });
            }

            function startDashboardUpdatePolling() {
                if (dashboardUpdatePollTimer) return;
                refreshDashboardUpdateNotice().catch(() => {});
                dashboardUpdatePollTimer = setInterval(() => {
                    if (document.hidden) return;
                    refreshDashboardUpdateNotice().catch(() => {});
                }, 20000);
            }

            async function installDesktopUpdateNow() {
                const installBtn = document.getElementById('dashboardUpdateInstallBtn');
                if (installBtn) {
                    installBtn.disabled = true;
                    installBtn.textContent = 'Installerer...';
                }

                try {
                    const r = await fetch('/desktop-update-install', { method: 'POST' });
                    const d = await r.json();
                    if (!r.ok) throw new Error((d && d.message) ? d.message : ('HTTP ' + r.status));
                    alert((d && d.message) ? d.message : 'Installering starter.');
                } catch (e) {
                    alert('Fejl ved installering: ' + e.message);
                } finally {
                    if (installBtn) {
                        installBtn.disabled = false;
                        installBtn.textContent = 'Installer nu';
                    }
                    refreshDashboardUpdateNotice().catch(() => {});
                }
            }

            function selectOrder(ordNo) {
                document.getElementById('orderInput').value = ordNo;
                prefetchAftercalcData(ordNo);
                orderListVisible = false;
                renderOrderList();
                searchOrder();
                window.scrollTo({ top: 0, behavior: 'auto' });
            }

            function goBackToList() {
                document.getElementById('result').innerHTML = '';
                orderListVisible = true;
                renderOrderList();
                setTimeout(() => {
                    const listEl = document.getElementById('orderList');
                    if (listEl) scrollToElementWithStickyOffset(listEl);
                }, 50);
            }

            function toggleSearchBox() {
                const searchBox = document.getElementById('searchBox');
                const collapseToggleBtn = document.getElementById('collapseToggleBtn');
                const collapseExpandBtn = document.getElementById('collapseExpandBtn');
                
                searchBox.classList.toggle('collapsed');
                if (searchBox.classList.contains('collapsed')) {
                    collapseToggleBtn.style.display = 'inline-block';
                    collapseExpandBtn.style.display = 'none';
                    collapseToggleBtn.textContent = '↗ Søg';
                } else {
                    collapseToggleBtn.style.display = 'none';
                    collapseExpandBtn.style.display = 'inline-block';
                }
                setTimeout(syncStickyOffsets, 0);
            }

            let uiBootstrapped = false;
            function bootstrapUiAfterLoad() {
                if (uiBootstrapped) return;
                uiBootstrapped = true;
                try {
                    const storedName = localStorage.getItem('afterkalk_logged_user_name');
                    if (storedName) {
                        loggedUserDisplayName = sanitizeDisplayName(storedName);
                    }
                } catch {}
                updateHeaderGreeting();
                showAccessGate();
                syncStickyOffsets();
                window.addEventListener('resize', syncStickyOffsets);
                const orderInput = document.getElementById('orderInput');
                const accessGateBtn = document.getElementById('accessGateBtn');
                if (orderInput) {
                    orderInput.addEventListener('keydown', function(event) {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            if (!accessGranted) {
                                submitAccessCode();
                                return;
                            }
                            searchOrder();
                        }
                    });
                    orderInput.addEventListener('input', function() {
                        clearTimeout(prefetchOrderDebounceTimer);
                        const ordNo = String(orderInput.value || '').trim();
                        if (!ordNo || ordNo.length < 4) return;
                        prefetchOrderDebounceTimer = setTimeout(() => {
                            prefetchAftercalcData(ordNo);
                        }, 260);
                    });
                }
                if (accessGateBtn) {
                    accessGateBtn.addEventListener('click', function(event) {
                        event.preventDefault();
                        submitAccessCode();
                    });
                }
                const accessGateInput = document.getElementById('accessGateInput');
                if (accessGateInput) {
                    accessGateInput.addEventListener('keydown', function(event) {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            submitAccessCode();
                        }
                    });
                }
                const accessGateUserInput = document.getElementById('accessGateUserInput');
                if (accessGateUserInput) {
                    accessGateUserInput.addEventListener('keydown', function(event) {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            const codeInput = document.getElementById('accessGateInput');
                            if (codeInput) {
                                codeInput.focus();
                                codeInput.select();
                            }
                        }
                    });
                }
                const sideMenuLoginInput = document.getElementById('sideMenuLoginInput');
                if (sideMenuLoginInput) {
                    sideMenuLoginInput.addEventListener('keydown', function(event) {
                        if (event.key === 'Enter') {
                            event.preventDefault();
                            submitAccessCodeFromSideMenu();
                        }
                    });
                }
                const sideMenuUserInput = document.getElementById('sideMenuUserInput');
                if (sideMenuUserInput) {
                    sideMenuUserInput.value = sanitizeDisplayName(loggedUserDisplayName);
                }
                updateReportOpenButtonState(Boolean(lastOrderReportHtml));
                refreshSideMenuAuthState();
                const orderListEl = document.getElementById('orderList');
                if (orderListEl) {
                    orderListEl.addEventListener('pointerdown', function(e) {
                        const sortHeader = e.target.closest('.order-sortable-header');
                        if (!sortHeader) return;
                        e.preventDefault();
                        const field = sortHeader.getAttribute('data-sort-field');
                        if (field) setOrderListSort(field);
                    });

                    orderListEl.addEventListener('click', function(e) {
                        const sortHeader = e.target.closest('.order-sortable-header');
                        if (sortHeader) {
                            return;
                        }
                        const refreshBtn = e.target.closest('.order-refresh-one-btn');
                        if (refreshBtn) {
                            e.preventDefault();
                            e.stopPropagation();
                            const ordNo = refreshBtn.getAttribute('data-ordno');
                            refreshSingleOrderCacheByOrdNo(ordNo, false, refreshBtn);
                            return;
                        }
                        if (e.target.closest('.order-refresh-cell')) {
                            e.preventDefault();
                            e.stopPropagation();
                            return;
                        }
                        const tr = e.target.closest('tr[data-ordno]');
                        if (tr) selectOrder(Number(tr.dataset.ordno));
                    });

                    orderListEl.addEventListener('mouseover', function(e) {
                        const tr = e.target.closest('tr[data-ordno]');
                        if (!tr) return;
                        const ordNo = String(tr.dataset.ordno || '').trim();
                        if (!ordNo) return;
                        prefetchAftercalcData(ordNo);
                    });
                }
            }

            // Soeg ved indlaesning hvis ordrenummer er i query string
            window.addEventListener('load', bootstrapUiAfterLoad, { once: true });
            document.addEventListener('DOMContentLoaded', bootstrapUiAfterLoad, { once: true });
            if (document.readyState === 'complete' || document.readyState === 'interactive') {
                setTimeout(bootstrapUiAfterLoad, 0);
            }
        
