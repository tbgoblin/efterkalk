// Printable explanation of "Vareforbrug vs. materiale" for one month (saved as PDF via the print dialog).
(function (root, factory) {
    if (typeof module === 'object' && module.exports) module.exports = factory();
    else root.GohReconciliationReport = factory();
}(typeof globalThis === 'undefined' ? this : globalThis, function () {
    'use strict';
    const MONTHS = ['januar','februar','marts','april','maj','juni','juli','august','september','oktober','november','december'];
    const CATEGORY_LABELS = { plates:'Pladelager', opfolgningvare:'Opfølgningsvarer', stang:'Stang', gr5Items:'Lagerkomponenter (gr. 5)', diverse:'Diverse', restPlates:'Restplader' };
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
        const measured = details.wip && details.wip.delta != null ? details.wip.delta : null;
        const finished = (varelager.workInProgress || []).find(item => item.key === 'finishedNotInvoiced' && Math.abs(item.delta) >= 0.5);
        if (info.bySource('finished_material')) {
            // Already part of the formula above; nothing left to explain here.
        } else if (measured != null) {
            items.push('<b>Færdige, afsendte ordrer, der endnu ikke er faktureret.</b> Materialet (uden tid) i dem er ændret med '
                + esc(signed(measured)) + ' (se afsnit 7). Det er købt og forbrugt, men hverken med i det beregnede materiale eller i Varelager. Medregnet ville resten være '
                + esc(dkk(residual == null ? null : residual - measured)) + '.');
        } else if (finished) {
            items.push('<b>Færdige, afsendte ordrer, der endnu ikke er faktureret.</b> Deres kost er ændret med ' + esc(signed(finished.delta))
                + '. Det er købt og forbrugt, men hverken med i det beregnede materiale eller i Varelager. Værdien indeholder også tid, så kun en del er materiale.' + direction(finished.delta));
        }
        const timing = details.timing && Number.isFinite(details.timing.value) ? details.timing.value : null;
        if (info.bySource('purchase_timing')) {
            // Already part of the formula above.
        } else if (timing != null) {
            items.push('<b>Fakturaer bogført i en anden måned end varemodtagelsen.</b> Målt til ' + esc(signed(timing))
                + ' (se afsnit 8). Medregnet ville resten være ' + esc(dkk(residual == null ? null : residual - timing)) + '.');
        } else {
            items.push('<b>Leverandørfakturaer for varer modtaget i en tidligere måned.</b> De øger Bogført i denne måned uden at være med i "Ikke fuldt faktureret".');
        }
        items.push('<b>Spild, skrot og prisforskelle.</b> Spild af plader og stang, der ikke er registreret på ordrerne, og forskellen mellem FIFO-kostprisen i ordrerne og den fakturerede indkøbspris.');
        if (!data.ubOrders || !data.ubOrders.length) items.push('<b>Ordrer leveret uden faktura.</b> Der er ingen U/B-ordrer registreret i måneden. Hvis der er leveret ordrer uden faktura (garanti, prøver, internt brug), bør de registreres som U/B, så deres materiale kommer med.');
        return items;
    }

    const millions = (value, decimals) => new Intl.NumberFormat('da-DK', { minimumFractionDigits:decimals, maximumFractionDigits:decimals }).format(Number(value) / 1e6) + ' mio. kr.';
    const kroner = value => number(Math.abs(value), 0) + ' kr.';

    // Largest invoice pushing the timing in the same direction as its total.
    function timingDetail(timing, value) {
        const top = (timing && timing.items || []).find(item => Math.sign(item.amount) === Math.sign(value)) || null;
        if (!top) return '';
        const received = top.receiptMonths.map(monthName).join(' og ');
        const text = top.text && !/^indkøbs faktura$/i.test(top.text) ? ' (' + top.text + ')' : '';
        return ', især ' + (top.referencedInvoNo ? 'en kreditnota' : 'faktura ' + top.invoNo) + ' på ' + kroner(top.amount)
            + ' for varer modtaget i ' + received + text;
    }

    // One sentence per explaining row; wording follows the sign of the row's value.
    function explainRow(row, details) {
        const v = row.value;
        switch (row.source) {
            case 'warehouse_delta': return v >= 0 ? 'Lageret er steget: en del af det købte ligger stadig på lager.'
                : 'Lageret er faldet: der er brugt materiale fra lageret, som er købt i tidligere måneder.';
            case 'purchase_gap': return 'Varer er modtaget, men leverandørfakturaen er endnu ikke bogført.';
            case 'finished_material': return v >= 0 ? 'Mere materiale sidder i færdige, afsendte ordrer, der endnu ikke er faktureret, end ved månedens start.'
                : 'Færdige ordrer fra tidligere måneder er faktureret i måneden, så deres materiale er nu med i det beregnede.';
            case 'via_material': return v >= 0 ? 'Mere materiale sidder i ordrer i produktion end ved månedens start.'
                : 'Ordrer i produktion er blevet færdige, så mindre materiale sidder i igangværende ordrer.';
            case 'ub_material': return 'Materiale er brugt på ordrer leveret uden faktura (U/B).';
            case 'purchase_timing': return (v >= 0 ? 'Der er bogført mere end månedens varemodtagelser' : 'Der er bogført mindre end månedens varemodtagelser')
                + timingDetail(details.timing, v) + '.';
            default: return row.name + '.';
        }
    }

    function buildSummary(data, info) {
        const rows = data.rows || [];
        const metrics = data.metrics || {};
        const signs = data.signs || {};
        const booked = info.accountRows[0];
        const material = info.bySource('material');
        const residual = info.residualRow;
        if (!booked || !material || !residual || [booked.value, material.value, residual.value, metrics.invoiced, metrics.cost].some(v => v == null)) return '';
        const difference = booked.value - material.value;
        const start = rows.findIndex(row => row.type === 'formula');
        const explaining = rows.slice(start + 1, rows.indexOf(residual))
            .filter(row => row.type !== 'heading' && row.type !== 'formula' && signs[row.id] && row.value != null && Math.abs(row.value) >= 0.5);
        const db = metrics.invoiced - metrics.cost;
        const parts = [];
        parts.push('<p>I ' + esc(monthName(data.month)) + ' er der faktureret <b>' + esc(millions(metrics.invoiced, 1)) + '</b>'
            + (metrics.orders != null ? ' (' + esc(number(metrics.orders, 0)) + ' ordrer)' : '') + ' med en kost på <b>' + esc(millions(metrics.cost, 1))
            + '</b>, svarende til et dækningsbidrag på <b>' + esc(millions(db, 1)) + ' (ca. ' + esc(number(metrics.invoiced ? db / metrics.invoiced * 100 : 0, 0)) + ' %)</b>.</p>');
        parts.push('<p>Materialet på de fakturerede ordrer udgør <b>' + esc(millions(material.value, 2)) + '</b>, mens der i bogføringen er <b>'
            + esc(millions(booked.value, 2)) + '</b> på vareforbrugskontiene. Der er altså bogført <b>' + esc(kroner(difference)) + ' '
            + (difference >= 0 ? 'mere' : 'mindre') + '</b>, end der er brugt på ordrerne.' + (explaining.length ? ' Forskellen forklares således:' : '') + '</p>');
        if (explaining.length) {
            parts.push('<ul class="summary-list">' + explaining.map(row => {
                const effect = (signs[row.id] === '−' ? -1 : 1) * row.value;
                return '<li><b>' + (effect < 0 ? '−' : '+') + ' ' + esc(kroner(effect)) + '</b> ' + esc(explainRow(row, data.details || {})) + '</li>';
            }).join('') + '</ul>');
        }
        const share = booked.value ? Math.abs(residual.value) / Math.abs(booked.value) * 100 : null;
        const positive = 'der er bogført mere, end der er brugt eller ligger på lager, fx spild eller køb, der ikke er registreret på ordrerne';
        const negative = 'ordrerne bærer mere materiale, end der er bogført, fx fordi varer er forbrugt til en højere kostpris end den endelige fakturapris';
        const meaning = residual.value < 0
            ? 'Resten er negativ. Det betyder, at ' + negative + '. En positiv rest ville betyde, at ' + positive + '.'
            : 'Resten er positiv. Det betyder, at ' + positive + '. En negativ rest ville betyde, at ' + negative + '.';
        parts.push('<p>Tilbage er en uforklaret rest på <b>' + esc(number(residual.value, 0)) + ' kr.</b>'
            + (share != null ? ', svarende til <b>' + esc(number(share, 1)) + ' %</b> af det bogførte vareforbrug' : '') + '. ' + esc(meaning) + '</p>');
        return '<section class="summary"><h2>Sammenfatning</h2>' + parts.join('') + '</section>';
    }

    // Monthly controls. input: { month, invoicedAll (incl. U/B), controls (server), timing, multiMonth, missingCost }
    function buildControls(input) {
        const controls = input.controls;
        if (!controls) return [];
        const list = [];
        const revenueDiff = controls.revenueBooked - Number(input.invoicedAll || 0);
        const other = controls.otherRevenue || [];
        const otherItems = other.map(item => item.account + ' ' + item.name + ' · ' + signed(item.amount, 2) + (item.text ? ' · ' + item.text : '') + ' (ikke en faktura, indgår ikke)');
        list.push(Math.abs(revenueDiff) < 1
            ? { key:'revenue', title:'Faktureret = bogført omsætning', status:'ok', text:'Fakturaerne i Månedens DB (inkl. U/B) svarer til den bogførte omsætning på fakturaer: ' + dkk(controls.revenueBooked) + '.', items: otherItems }
            : { key:'revenue', title:'Faktureret = bogført omsætning', status:'warn', text:'Bogført omsætning på fakturaer ' + dkk(controls.revenueBooked) + ', fakturaer i Månedens DB ' + dkk(input.invoicedAll)
                + ' (forskel ' + signed(revenueDiff, 2) + '). En ordre mangler eller er talt med forkert beløb.', items: otherItems });
        const timing = input.timing;
        if (timing && Array.isArray(timing.unlinked)) {
            const count = timing.unlinkedCount ?? timing.unlinked.length;
            list.push(count ? { key:'unlinked', title:'Bogføringer uden indkøbsordre', status:'warn',
                text:count + ' bogføringer på 12070, 12075 og 12078 uden tilknyttet indkøbsordre, i alt ' + dkk(timing.unlinkedTotal) + '. Kontrollér om de er materiale til ordrerne eller hører på en anden konto.',
                items: timing.unlinked.slice(0, 8).map(item => (item.supplier || 'Ukendt leverandør') + ' · faktura ' + item.invoNo + ' · ' + signed(item.amount) + (item.text && !/^indkøbs faktura$/i.test(item.text) ? ' · ' + item.text : '')) }
                : { key:'unlinked', title:'Bogføringer uden indkøbsordre', status:'ok', text:'Alle bogføringer på 12070, 12075 og 12078 er koblet til en indkøbsordre.' });
        }
        const missing = input.missingCost || [], multi = input.multiMonth || [];
        list.push(missing.length
            ? { key:'orders', title:'Ordrer med kostdata', status:'warn', text:missing.length + ' fakturerede ordrer mangler kostdata, så det beregnede materiale er ufuldstændigt: ' + missing.slice(0, 12).join(', ') + '.' }
            : { key:'orders', title:'Ordrer med kostdata', status: multi.length ? 'info' : 'ok',
                text:'Alle fakturerede ordrer har kostdata.' + (multi.length ? ' ' + multi.length + ' ordrer er faktureret over flere måneder; kun månedens andel er talt med.' : ''),
                items: multi.slice(0, 8).map(order => 'Ordre ' + order.OrdNo + ': ' + signed(order.InvoAm, 2) + ' af ' + number(order.InvoAmTotal, 2) + ' i måneden, kostandel ' + number(order.InvoiceShare * 100, 0) + ' %') });
        const snapshot = meta => {
            if (!meta || meta.missing) return 'mangler lagerliste-lukning for ' + monthName(meta && meta.month);
            if (meta.live) return null;
            const days = daysAfterMonthEnd(meta);
            return days != null && days > 1.5 ? 'lukningen for ' + monthName(meta.month) + ' er gemt ' + number(Math.floor(days), 0) + ' dage efter månedens udgang' : null;
        };
        const problems = [snapshot(controls.snapshots.current), snapshot(controls.snapshots.previous)].filter(Boolean);
        list.push(problems.length
            ? { key:'snapshot', title:'Lagerliste-lukninger', status:'warn', text:problems.join('; ') + '.' }
            : { key:'snapshot', title:'Lagerliste-lukninger', status: controls.snapshots.current.live ? 'info' : 'ok',
                text: controls.snapshots.current.live ? 'Måneden er ikke lukket; Varelager er den levende beregning.'
                    : 'Begge lukninger er gemt ved månedsskiftet (' + dateTime(controls.snapshots.previous.createdAt) + ' og ' + dateTime(controls.snapshots.current.createdAt) + ').' });
        const unchanged = controls.diverseUnchanged || [];
        list.push(unchanged.length
            ? { key:'diverse', title:'Diverse optalt', status:'warn', text:unchanged.length + ' rækker i Diverse (over 5.000 kr.) har nøjagtig samme værdi som sidste måned. Kontrollér om de er optalt.',
                items: unchanged.slice(0, 8).map(row => row.descr + ' · ' + number(row.value, 0) + ' kr.') }
            : { key:'diverse', title:'Diverse optalt', status:'ok', text:'Alle større rækker i Diverse er ændret siden sidste måned.' });
        return list;
    }

    function controlsHtml(list) {
        if (!list.length) return '';
        const label = { ok:'OK', warn:'Afvigelse', info:'Info' };
        return '<h2>Kontroller</h2><table class="controls" data-universal-sort="off"><tbody>' + list.map(item => '<tr class="ctl-' + item.status + '"><td class="ctl-status">' + label[item.status]
            + '</td><td><b>' + esc(item.title) + '.</b> ' + esc(item.text)
            + (item.items && item.items.length ? '<ul>' + item.items.map(text => '<li>' + esc(text) + '</li>').join('') + '</ul>' : '') + '</td></tr>').join('') + '</tbody></table>';
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
        parts.push(buildSummary(data, info));
        parts.push(controlsHtml(data.controlList || []));

        parts.push('<h2>1. Resultat</h2><table><tbody>' + rows.filter(row => row.display !== 'hidden').map(row => row.type === 'heading'
            ? '<tr class="heading"><td colspan="2">' + esc(row.name) + '</td></tr>'
            : '<tr' + (row === info.residualRow ? ' class="total"' : '') + '><td>' + (data.signs && data.signs[row.id] ? '<span class="sign">' + esc(data.signs[row.id]) + '</span>' : '') + esc(row.name)
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

        const wip = details.wip;
        parts.push('<h2>7. Færdige, afsendte ordrer, der endnu ikke er faktureret</h2><p>Materiale, stang, indkøbte dele og underleverandør (uden tid) i ordrer, der ved månedens udgang er færdige og afsendt, men ikke faktureret. '
            + 'Materialet er taget fra lageret, så Varelager er faldet, men det kommer først med i det beregnede materiale, når ordren faktureres. Ændringen i måneden er det, der påvirker resten.'
            + (data.ubOrders && data.ubOrders.length ? ' Månedens U/B-ordrer er udeladt, da de står i afsnit 6.' : '') + '</p>');
        if (wip && wip.current && wip.previous) {
            const line = (label, key, count, cls) => '<tr' + (cls ? ' class="' + cls + '"' : '') + '><td>' + esc(label) + '</td><td class="num">' + esc(number(wip.previous[key], 0)) + ' (' + esc(wip.previous[count]) + ')</td><td class="num">'
                + esc(number(wip.current[key], 0)) + ' (' + esc(wip.current[count]) + ')</td><td class="num">'
                + esc(wip.previous[key] == null || wip.current[key] == null ? '—' : signed(wip.current[key] - wip.previous[key])) + '</td></tr>';
            parts.push('<table><thead><tr><th>Materiale (antal ordrer)</th><th class="num">' + esc(monthName(wip.previousMonth)) + '</th><th class="num">' + esc(monthName(wip.month))
                + '</th><th class="num">Ændring</th></tr></thead><tbody>' + line('Færdige salgsordrer, ikke faktureret', 'finished', 'finishedOrders', 'total')
                + (info.bySource('via_material') ? line('Salgsordrer i arbejde (VIA)', 'via', 'viaOrders', 'total')
                    : line('Til orientering: salgsordrer i arbejde (VIA), ikke medregnet', 'via', 'viaOrders', 'note')) + '</tbody></table>');
            const missing = [...(wip.previous.missing || []), ...(wip.current.missing || [])];
            if (missing.length) parts.push('<p class="warning">Materialet kunne ikke beregnes for ordre ' + esc(missing.join(', ')) + '.</p>');
        } else parts.push('<p class="warning">Kan ikke beregnes, fordi en af månederne mangler en lagerliste-lukning.</p>');

        const timing = details.timing;
        parts.push('<h2>8. Fakturaer bogført i en anden måned end varemodtagelsen</h2><p>Bogføringer på 12070, 12075 og 12078 er koblet til indkøbsordrelinjerne via leverandørens fakturanummer. '
            + 'Det, der er bogført i måneden for varer modtaget i andre måneder, trækkes ud, og det, der er bogført i andre måneder (±3) for varer modtaget i måneden, lægges til. '
            + 'Kreditnotaer uden egen indkøbsordre følges til den faktura, de henviser til. Varer modtaget men endnu ikke faktureret står i afsnit 5.</p>');
        if (timing && Number.isFinite(timing.value)) {
            parts.push('<table><tbody><tr><td>Bogført i måneden for varer modtaget i andre måneder</td><td class="num">' + esc(dkk(timing.bookedForOtherMonths)) + '</td></tr>'
                + '<tr><td>Bogført i andre måneder for varer modtaget i måneden</td><td class="num">' + esc(dkk(timing.bookedElsewhereForThisMonth)) + '</td></tr>'
                + '<tr class="total"><td>Periodeforskydning (trækkes fra)</td><td class="num">' + esc(dkk(timing.value)) + '</td></tr></tbody></table>');
            if ((timing.items || []).length) {
                parts.push('<table class="movers"><colgroup><col style="width:14%"><col style="width:16%"><col style="width:18%"><col style="width:36%"><col style="width:16%"></colgroup><thead><tr><th>Bogført</th><th>Faktura</th><th>Modtaget</th><th>Tekst</th><th class="num">Effekt</th></tr></thead><tbody>'
                    + timing.items.map(item => '<tr><td>' + esc(monthName(item.bookedMonth)) + '</td><td>' + esc(item.invoNo) + (item.referencedInvoNo ? '<div class="note">vedr. ' + esc(item.referencedInvoNo) + '</div>' : '')
                        + '</td><td>' + esc(item.receiptMonths.map(monthName).join(', ')) + '</td><td>' + esc(item.text) + '</td><td class="num">' + esc(signed(item.amount)) + '</td></tr>').join('') + '</tbody></table>');
            }
        } else parts.push('<p>—</p>');

        parts.push('<h2>9. Mulige forklaringer på resten</h2><p>Punkterne er ikke beregnet ind i resten. De er de mest sandsynlige årsager, med beløb hvor de kan måles.</p><ol>'
            + explanations(data, info).map(item => '<li>' + item + '</li>').join('') + '</ol>');
        // Opt out of assets/js/table-sort.js, which otherwise adds sort arrows and resizers to the printed tables.
        return parts.join('').replace(/<table(?=[\s>])/g, '<table data-universal-sort="off"');
    }

    function fileTitle(month) {
        return 'Report ' + monthName(month);
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
        // The print dialog proposes the document title as the PDF file name.
        const previousTitle = document.title;
        document.title = fileTitle(data.month);
        const done = () => {
            document.body.classList.remove('kf-recon-printing');
            document.title = previousTitle;
            window.removeEventListener('afterprint', done);
        };
        window.addEventListener('afterprint', done);
        requestAnimationFrame(() => window.print());
    }

    return { buildHtml, buildControls, controlsHtml, print, fileTitle };
}));
