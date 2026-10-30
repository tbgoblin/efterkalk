// ── buildOrderDetailReportHtml (generatore HTML del report ordine stampabile) · client ──
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, formatNumber, formatCount, isLaserLProdNo
// (dichiarati nell'inline script).
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

            const totalCost = key === '1'
                ? (Number((line && (line.EffectiveOperationMinutes ?? line.NoFin)) || 0) * Number((line && line.CCstPr) || 0))
                : Number((line && (line.EffectiveLineCost ?? line.LineCost)) || 0);
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

        for (const line of lines) {
            const key = (line && line.ProdTp4 !== null && line.ProdTp4 !== undefined) ? String(line.ProdTp4) : 'NA';
            const lnNo = Number((line && line.LnNo) || 0);
            if (lnNo === 1 || key !== '3') continue;
            usedMinutes += Number((line && (line.EffectiveOperationMinutes ?? line.NoFin)) || 0);
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
    const costExclusion = (orderData && orderData.costExclusion) || {};
    const excludedCost = Number(costExclusion.excludedCost || 0);
    const excludedLineCount = Number(costExclusion.excludedLineCount || 0);
    const originalCost = Number.isFinite(Number(costExclusion.originalCost))
        ? Number(costExclusion.originalCost)
        : cost;
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
    if (excludedLineCount > 0) {
        html += '<div class="report-cost-adjustment"><strong>Permanent kostjustering:</strong> kost fra ' + formatCount(excludedLineCount) + ' salgslinje(r) er markeret som udeladt. Fratrukket kost: ' + formatNumber(excludedCost) + ' DKK; beregnet kost før justering: ' + formatNumber(originalCost) + ' DKK. Salgsprisen er uændret. Valget er gemt i GOH.</div>';
    }
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
