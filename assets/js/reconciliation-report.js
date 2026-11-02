// Printable explanation of "Vareforbrug vs. materiale" for one month (saved as PDF via the print dialog).
(function (root, factory) {
    if (typeof module === 'object' && module.exports) module.exports = factory();
    else root.GohReconciliationReport = factory();
}(typeof globalThis === 'undefined' ? this : globalThis, function () {
    'use strict';
    const MONTHS = ['januar','februar','marts','april','maj','juni','juli','august','september','oktober','november','december'];
    const CATEGORY_LABELS = { plates:'Pladelager', opfolgningvare:'Opfølgningsvarer', stang:'Stang', gr5Items:'Lagerkomponenter (gr. 5)', diverse:'Diverse', restPlates:'Restplader' };
    const WIP_LABELS = { salgordreVia:'Salgsordrer i arbejde (VIA)', finishedNotInvoiced:'Færdige salgsordrer, ikke faktureret (kost)' };
    // Accounts in this chart of accounts that are not material purchases comparable with the orders' calculated material.
    const ACCOUNT_NOTES = {
        12070: { text:'Hovedkontoen for materialeindkøb.' },
        12075: { text:'Underleverandørarbejde indgår også i det beregnede materiale.' },
        12078: { text:'Overfladebehandling er underleverandørarbejde og indgår i det beregnede materiale.' },
        12067: { text:'Paller og emballage indgår ikke i ordrernes beregnede materiale.', distorts:true },
        12080: { text:'Lagerregulering er bogføringens egen lagerændring. Δ Varelager måler det samme, så beløbet tælles dobbelt.', distorts:true },
        12085: { text:'Intern køb: kontrollér, om det svarer til materiale i ordrerne.' },
        12090: { text:'Vikar er løn, ikke materiale.', distorts:true },
        12092: { text:'Salg af jern/skrot er en indtægt, ikke et indkøb.', distorts:true }
    };
    const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
    const number = (value, decimals = 2) => value == null || !Number.isFinite(Number(value)) ? '—'
        : new Intl.NumberFormat('da-DK', { minimumFractionDigits:decimals, maximumFractionDigits:decimals }).format(Number(value));
    const dkk = (value, decimals = 2) => value == null ? '—' : number(value, decimals) + ' DKK';
    const signed = (value, decimals = 0) => value == null ? '—' : (value > 0 ? '+' : '') + number(value, decimals);
    const monthName = key => { const [year, month] = String(key || '').split('-').map(Number); return MONTHS[month - 1] ? MONTHS[month - 1] + ' ' + year : String(key || ''); };
    const dateTime = iso => {
        const date = new Date(iso);
        return Number.isFinite(date.getTime()) ? new Intl.DateTimeFormat('da-DK', { timeZone:'Europe/Copenhagen', dateStyle:'short', timeStyle:'short' }).format(date) : '—';
    };
    const valueText = row => row.value == null ? '—' : row.unit === 'count' ? number(row.value, 0) : row.unit === 'percent' ? number(row.value) + ' %' : dkk(row.value);

    function daysAfterMonthEnd(meta) {
        if (!meta || !meta.createdAt) return null;
        const [year, month] = meta.month.split('-').map(Number);
        const monthEnd = Date.UTC(year, month, 1) - 2 * 3600 * 1000; // midnight Copenhagen, summer-time safe enough for whole days
        return (new Date(meta.createdAt).getTime() - monthEnd) / 86400000;
    }

    function snapshotText(meta) {
        if (!meta) return 'mangler lagerliste-lukning';
        if (meta.live) return 'levende beregning (måneden er ikke lukket)';
        const mode = meta.valuationMode === 'live-at-save' ? 'manuel lukning, værdi på gemmetidspunktet'
            : meta.valuationMode === 'live-during-calculation' ? 'automatisk lukning kl. 23:59' : 'lukning';
        return 'gemt ' + dateTime(meta.createdAt) + ' (' + mode + ')';
    }

    function formulaText(formula, names) {
        return String(formula || '').replace(/\[([A-Za-z][A-Za-z0-9_-]*)\]/g, (match, id) => names.has(id) ? '[' + names.get(id) + ']' : match);
    }

    function analyse(data) {
        const rows = data.rows || [];
        const bySource = source => rows.find(row => row.type === 'source' && row.source === source);
        const residualRow = [...rows].reverse().find(row => row.type === 'formula');
        const residual = residualRow ? residualRow.value : null;
        const accountRows = rows.filter(row => row.type === 'accounts');
        const distorting = [];
        for (const row of accountRows) {
            for (const account of row.accounts || []) {
                const note = ACCOUNT_NOTES[Number(account.AcNo)];
                const amount = Number(account.Amount) * row.sign;
                if (note && note.distorts && Math.abs(amount) >= 0.5) distorting.push({ row:row.name, AcNo:account.AcNo, Nm:account.Nm, amount });
            }
        }
        return { bySource, residualRow, residual, accountRows, distorting };
    }

    function explanations(data, info) {
        const items = [];
        const { residual, distorting } = info;
        const details = data.details || {};
        const varelager = details.varelager || {};
        const direction = value => residual == null || value == null || Math.abs(value) < 0.5 ? ''
            : Math.sign(value) === Math.sign(residual) ? ' Går i samme retning som resten og kan forklare en del af den.'
                : ' Går i modsat retning af resten og forklarer den ikke.';
        if (distorting.length) {
            const total = distorting.reduce((sum, item) => sum + item.amount, 0);
            items.push('<b>Konti der ikke er materialeindkøb.</b> "' + esc(distorting[0].row) + '" indeholder '
                + distorting.map(item => esc(item.AcNo + ' ' + item.Nm) + ' (' + esc(signed(item.amount)) + ')').join(', ')
                + '. Uden dem ville resten være ' + esc(dkk(residual == null ? null : residual - total)) + ' i stedet for ' + esc(dkk(residual)) + '.');
        }
        for (const meta of [varelager.current, varelager.previous]) {
            const days = daysAfterMonthEnd(meta);
            if (days != null && days > 1.5) items.push('<b>Lagerlisten for ' + esc(monthName(meta.month)) + ' er gemt ' + esc(number(days, 0))
                + ' dage efter månedens udgang.</b> Bevægelser i de dage (modtagelser, forbrug) er talt med i månedens lagerværdi.');
        }
        const wip = (varelager.workInProgress || []).filter(item => Math.abs(item.delta) >= 0.5);
        if (wip.length) {
            const total = wip.reduce((sum, item) => sum + item.delta, 0);
            items.push('<b>Materiale i ordrer, der endnu ikke er faktureret.</b> Det er købt og bogført, men hverken med i det beregnede materiale eller i Varelager. Ændring i måneden: '
                + wip.map(item => esc(WIP_LABELS[item.key] || item.key) + ' ' + esc(signed(item.delta))).join(', ')
                + '. Værdierne indeholder også tid, så kun en del er materiale.' + direction(total));
        }
        items.push('<b>Leverandørfakturaer for varer modtaget i en tidligere måned.</b> De øger Bogført i denne måned uden at være med i "Ikke fuldt faktureret". Kan ikke beregnes bagud, fordi faktureringsstatus kun kendes, som den er i dag.');
        items.push('<b>Spild, skrot og prisforskelle.</b> Spild af plader og stang, der ikke er registreret på ordrerne, og forskellen mellem FIFO-kostprisen i ordrerne og den fakturerede indkøbspris.');
        if (!data.ubOrders || !data.ubOrders.length) items.push('<b>Ordrer leveret uden faktura.</b> Der er ingen U/B-ordrer registreret i måneden. Hvis der er leveret ordrer uden faktura (garanti, prøver, internt brug), bør de registreres som U/B, så deres materiale kommer med.');
        return items;
    }

    function buildHtml(data) {
        const info = analyse(data);
        const details = data.details || {};
        const varelager = details.varelager || {};
        const gap = details.purchaseGap;
        const metrics = data.metrics || {};
        const rows = data.rows || [];
        const names = new Map(rows.map(row => [row.id, row.name]));
        const plateLabel = data.plateMode === 'standard' ? 'standardpris' : 'FIFO';
        const missing = rows.some(row => row.display !== 'hidden' && row.type !== 'heading' && row.value == null);
        const parts = [];
        parts.push('<header><h1>Afstemning af Vareforbrug · ' + esc(monthName(data.month)) + '</h1>'
            + '<p>Udskrevet ' + esc(dateTime(new Date().toISOString())) + ' · pladelager vurderet til ' + esc(plateLabel) + '</p></header>');
        if (missing) parts.push('<p class="warning">Nogle tal var ikke færdigberegnet, da rapporten blev lavet (vist som —). Vent til alle ordrer er beregnet, og udskriv igen.</p>');

        parts.push('<h2>1. Resultat</h2><table><tbody>' + rows.filter(row => row.display !== 'hidden').map(row => row.type === 'heading'
            ? '<tr class="heading"><td colspan="2">' + esc(row.name) + '</td></tr>'
            : '<tr' + (row === info.residualRow ? ' class="total"' : '') + '><td>' + esc(row.name)
                + (row.type === 'formula' ? '<div class="note">= ' + esc(formulaText(row.formula, names)) + '</div>' : '')
                + '</td><td class="num">' + esc(valueText(row)) + '</td></tr>').join('') + '</tbody></table>');
        const booked = info.accountRows[0];
        if (info.residual != null && booked && booked.value) {
            parts.push('<p>Uforklaret rest svarer til ' + esc(number(Math.abs(info.residual) / Math.abs(booked.value) * 100, 1)) + ' % af "' + esc(booked.name) + '".</p>');
        }

        parts.push('<h2>2. Bogført Vareforbrug — konti</h2>');
        if (!info.accountRows.length) parts.push('<p>Opsætningen har ingen række med konti.</p>');
        for (const row of info.accountRows) {
            const posted = (row.accounts || []).filter(account => Math.abs(Number(account.Amount)) >= 0.005);
            const empty = (row.accounts || []).length - posted.length;
            parts.push('<h3>' + esc(row.name) + (row.sign < 0 ? ' (fortegn vendt)' : '') + '</h3><table><thead><tr><th>Konto</th><th>Navn</th><th class="num">Beløb</th><th>Bemærkning</th></tr></thead><tbody>'
                + posted.map(account => {
                    const note = ACCOUNT_NOTES[Number(account.AcNo)];
                    const amount = Number(account.Amount) * row.sign;
                    return '<tr' + (note && note.distorts && Math.abs(amount) >= 0.5 ? ' class="flag"' : '') + '><td>' + esc(account.AcNo) + '</td><td>' + esc(account.Nm) + '</td><td class="num">'
                        + esc(number(amount)) + '</td><td class="note">' + esc(note ? note.text : '') + '</td></tr>';
                }).join('')
                + '<tr class="total"><td colspan="2">I alt</td><td class="num">' + esc(number(row.value)) + '</td><td class="note">'
                + (empty ? esc(empty + (empty === 1 ? ' konto' : ' konti') + ' uden bevægelser er udeladt.') : '') + '</td></tr></tbody></table>');
        }

        parts.push('<h2>3. Beregnet materiale</h2><p>Summen af materialet på månedens fakturerede ordrer, inklusive kreditnotaer. U/B-ordrer er ikke med.</p><table><tbody>'
            + [['Materiale (plader)', metrics.material_only], ['Stang', metrics.stang], ['Indkøbte dele', metrics.parts], ['Underleverandør', metrics.subcontract]]
                .map(([label, value]) => '<tr><td>' + esc(label) + '</td><td class="num">' + esc(dkk(value)) + '</td></tr>').join('')
            + '<tr class="total"><td>I alt (' + esc(number(metrics.orders, 0)) + ' ordrer)</td><td class="num">' + esc(dkk(metrics.material)) + '</td></tr></tbody></table>');

        parts.push('<h2>4. Δ Varelager — ' + esc(monthName(varelager.previousMonth)) + ' → ' + esc(monthName(varelager.month)) + '</h2>'
            + '<p>Værdien af lageret ved udgangen af ' + esc(monthName(varelager.month)) + ' minus værdien ved udgangen af ' + esc(monthName(varelager.previousMonth))
            + '. Positiv betyder, at lageret er steget: en del af det købte ligger stadig på lager.</p><ul class="meta">'
            + '<li>' + esc(monthName(varelager.month)) + ': ' + esc(snapshotText(varelager.current)) + '</li>'
            + '<li>' + esc(monthName(varelager.previousMonth)) + ': ' + esc(snapshotText(varelager.previous)) + '</li></ul>');
        const categories = varelager.categories || [];
        if (!categories.length) parts.push('<p class="warning">Δ Varelager kan ikke beregnes, fordi en af månederne mangler en lagerliste-lukning.</p>');
        else {
            const sum = key => categories.reduce((total, item) => total + item[key], 0);
            parts.push('<table><thead><tr><th>Kategori</th><th class="num">' + esc(monthName(varelager.previousMonth)) + '</th><th class="num">' + esc(monthName(varelager.month)) + '</th><th class="num">Ændring</th></tr></thead><tbody>'
                + categories.map(item => '<tr><td>' + esc(CATEGORY_LABELS[item.key] || item.key) + (item.key === 'plates' ? ' (' + esc(plateLabel) + ')' : '') + '</td><td class="num">' + esc(number(item.previous, 0))
                    + '</td><td class="num">' + esc(number(item.current, 0)) + '</td><td class="num">' + esc(signed(item.delta)) + '</td></tr>').join('')
                + '<tr class="total"><td>Varelager i alt</td><td class="num">' + esc(number(sum('previous'), 0)) + '</td><td class="num">' + esc(number(sum('current'), 0)) + '</td><td class="num">' + esc(signed(sum('delta'))) + '</td></tr></tbody></table>');
            const moving = categories.filter(item => item.movers.length && Math.abs(item.delta) >= 1000);
            if (moving.length) parts.push('<h3>Største ændringer pr. kategori</h3>');
            for (const item of moving) {
                parts.push('<table class="movers"><colgroup><col style="width:17%"><col style="width:44%"><col style="width:13%"><col style="width:13%"><col style="width:13%"></colgroup><thead><tr><th colspan="2">' + esc(CATEGORY_LABELS[item.key] || item.key) + ' (' + esc(signed(item.delta)) + ')</th><th class="num">Før</th><th class="num">Efter</th><th class="num">Ændring</th></tr></thead><tbody>'
                    + item.movers.map(mover => '<tr><td>' + esc(mover.key) + '</td><td>' + esc(mover.descr) + '</td><td class="num">' + esc(number(mover.previous, 0)) + '</td><td class="num">'
                        + esc(number(mover.current, 0)) + '</td><td class="num">' + esc(signed(mover.delta)) + '</td></tr>').join('') + '</tbody></table>');
            }
        }

        parts.push('<h2>5. Modtaget indkøb, ikke fuldt faktureret</h2><p>Indkøbsordrelinjer modtaget i måneden, hvor leverandørfakturaen endnu ikke er bogført. Materialet kan være brugt, men er endnu ikke med i Bogført. Underleverandørvarer (varenummer U…) er udeladt.</p>');
        parts.push(gap ? '<table><tbody><tr><td>Ingen faktura (' + esc(number(gap.missingOrderCount, 0)) + ' ordrer)</td><td class="num">' + esc(dkk(gap.missingValue)) + '</td></tr>'
            + '<tr><td>Delvis faktureret (' + esc(number(gap.partialOrderCount, 0)) + ' ordrer)</td><td class="num">' + esc(dkk(gap.partialValue)) + '</td></tr>'
            + '<tr class="total"><td>I alt</td><td class="num">' + esc(dkk(gap.totalValue)) + '</td></tr></tbody></table>' : '<p>—</p>');

        parts.push('<h2>6. U/B-ordrer</h2><p>Ordrer leveret uden faktura. Materialet er købt og bogført, men ikke med i det beregnede materiale.</p>');
        const ub = data.ubOrders || [];
        parts.push(ub.length ? '<table><thead><tr><th>Ordre</th><th>Kunde</th><th>Bemærkning</th><th class="num">Materiale</th></tr></thead><tbody>'
            + ub.map(order => '<tr><td>' + esc(order.OrdNo) + '</td><td>' + esc(order.CustomerName || order.CustNo || '') + '</td><td>' + esc(order.text || '') + '</td><td class="num">' + esc(number(order.material)) + '</td></tr>').join('')
            + '<tr class="total"><td colspan="3">I alt</td><td class="num">' + esc(number(ub.every(order => order.material != null) ? ub.reduce((total, order) => total + order.material, 0) : null)) + '</td></tr></tbody></table>'
            : '<p>Ingen U/B-ordrer registreret i måneden.</p>');

        parts.push('<h2>7. Mulige forklaringer på resten</h2><p>Punkterne er ikke beregnet ind i resten. De er de mest sandsynlige årsager, med beløb hvor de kan måles.</p><ol>'
            + explanations(data, info).map(item => '<li>' + item + '</li>').join('') + '</ol>');
        return parts.join('');
    }

    function print(data) {
        let host = document.getElementById('kfReconPrint');
        if (!host) {
            host = document.createElement('div');
            host.id = 'kfReconPrint';
            document.body.appendChild(host);
        }
        host.innerHTML = buildHtml(data);
        document.body.classList.add('kf-recon-printing');
        const done = () => { document.body.classList.remove('kf-recon-printing'); window.removeEventListener('afterprint', done); };
        window.addEventListener('afterprint', done);
        requestAnimationFrame(() => window.print());
    }

    return { buildHtml, print };
}));
