let diverseAdminRows = [];
let diverseAdminMonth = null;
let diverseAdminTemplate = [];
function diverseAdminRender() {
    const root = document.getElementById('diverseAdminRows');
    const input = (row, key, disabled = false) => '<input style="width:90px" type="number" min="0" step="any" data-field="' + key + '" value="' + (row[key] ?? '') + '"' + (disabled ? ' disabled' : '') + '>';
    root.innerHTML = '<table><thead><tr><th>Kategori / beskrivelse</th><th>Beregning</th><th>Beløb</th><th>Antal / paller</th><th>Kg/palle</th><th>Pris / kg-pris</th><th>Handling</th><th>Værdi</th></tr></thead><tbody>'
        + diverseAdminRows.map((row, index) => '<tr data-index="' + index + '"><td>' + lagerlisteEscape(row.category)
            + '<br><input data-field="Descr" value="' + lagerlisteEscape(row.Descr) + '"></td><td><select data-field="mode">'
            + (row.category.startsWith('Skrot ') ? ['amount', 'pallets'] : ['amount', 'quantity']).map(mode => '<option value="' + mode + '"' + (row.mode === mode ? ' selected' : '') + '>' + ({ amount: 'Direkte beløb', quantity: 'Antal × pris', pallets: 'Paller × kg × pris' })[mode] + '</option>').join('')
            + '</select></td><td>' + input(row, 'amount', row.mode !== 'amount') + '</td><td>' + input(row, 'quantity', row.mode === 'amount') + '</td><td>' + input(row, 'kg', row.mode !== 'pallets') + '</td><td>' + input(row, 'price', row.mode === 'amount') + '</td><td><button type="button" onclick="diverseAdminRemove(' + index + ')">Fjern</button></td></tr>').join('') + '</tbody></table>';
    root.querySelectorAll('[data-field="mode"]').forEach(select => select.onchange = () => { diverseAdminCollect(); diverseAdminRender(); });
    root.querySelectorAll('thead th').forEach(th => { th.style.position = 'sticky'; th.style.top = '0'; th.style.background = '#eaf2ff'; });
    root.querySelectorAll('tr[data-index]').forEach(tr => tr.insertAdjacentHTML('beforeend', '<td data-value></td>'));
    root.insertAdjacentHTML('beforeend', '<p data-manual-total></p>');
    root.querySelectorAll('input').forEach(input => input.oninput = diverseAdminPreview);
    diverseAdminPreview();
}
function diverseAdminPreview() {
    diverseAdminCollect();
    let total = 0, complete = true;
    document.querySelectorAll('#diverseAdminRows tr[data-index]').forEach(tr => {
        const row = diverseAdminRows[Number(tr.dataset.index)];
        const fields = row.mode === 'amount' ? ['amount'] : row.mode === 'pallets' ? ['quantity', 'kg', 'price'] : ['quantity', 'price'];
        const valid = fields.every(key => row[key] !== '' && row[key] != null && Number.isFinite(Number(row[key])) && Number(row[key]) >= 0);
        const value = valid ? Math.round(fields.reduce((sum, key) => sum * Number(row[key]), 1) * 100) / 100 : 0;
        total += value;
        complete = complete && valid;
        tr.querySelector('[data-value]').textContent = valid ? lagerlisteFormat(value) : 'Ikke udfyldt';
    });
    const manualStock = diverseAdminRows.some(row => row.category === 'PEM (44)');
    document.querySelector('#diverseAdminRows [data-manual-total]').textContent = 'Manuelle værdier: ' + lagerlisteFormat(total) + (complete ? '' : ' (foreløbigt – tomme felter mangler)') + (manualStock ? '. Denne måned: 44/45/46/63 indtastes manuelt; Visma tilføjes IKKE.' : '. Visma 44/45/46/63 tilføjes i Lagerliste.');
}
function diverseAdminCollect() {
    document.querySelectorAll('#diverseAdminRows tr[data-index]').forEach(tr => {
        const row = diverseAdminRows[Number(tr.dataset.index)];
        tr.querySelectorAll('[data-field]').forEach(input => row[input.dataset.field] = input.value);
    });
}
async function diverseAdminLoad() {
    const status = document.getElementById('diverseAdminStatus');
    const month = document.getElementById('diverseAdminMonth').value;
    try {
        if (diverseAdminMonth && !confirm('Hent måned og erstat eventuelle ikke-gemte ændringer?')) return;
        const response = await fetch('/admin/lagerliste-diverse/' + encodeURIComponent(month), { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'Kunne ikke hente');
        diverseAdminRows = data.data.rows;
        diverseAdminTemplate = data.template || [];
        diverseAdminMonth = month;
        diverseAdminRender();
        status.textContent = 'Viser ' + month + '. ' + (data.data.saved ? 'Gemte værdier.' : 'Ikke udfyldt endnu.');
    } catch (err) { status.textContent = err.message; }
}
async function diverseAdminCopyPrevious() {
    const status = document.getElementById('diverseAdminStatus');
    try {
        const target = document.getElementById('diverseAdminMonth').value;
        if (!/^\d{4}-\d{2}$/.test(target)) throw new Error('Vælg mål-måned først.');
        const date = new Date(target + '-01T12:00:00Z');
        date.setUTCMonth(date.getUTCMonth() - 1);
        const previous = date.toISOString().slice(0, 7);
        if (!confirm('Kopiér ' + previous + ' til ' + target + '? Ikke-gemte ændringer erstattes. Kontrollér alle værdier før Gem.')) return;
        const response = await fetch('/admin/lagerliste-diverse/' + previous, { headers: { Authorization: 'Bearer ' + String(authToken || '') } });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'Kunne ikke hente');
        if (!data.data.saved) throw new Error('Forrige måned har ingen gemte Diverse-værdier.');
        diverseAdminMonth = target;
        diverseAdminRows = data.data.rows;
        diverseAdminTemplate = data.template || [];
        diverseAdminRender();
        status.textContent = 'Kopieret fra ' + previous + ' – IKKE GEMT. Kontrollér månedens antal, priser og beløb før Gem.';
    } catch (err) { status.textContent = err.message; }
}
function diverseAdminAdd() {
    if (!diverseAdminMonth) return;
    diverseAdminCollect();
    const category = document.getElementById('diverseAdminCategory').value;
    if (diverseAdminRows.some(row => row.category === category && row.mode === 'amount')) {
        document.getElementById('diverseAdminStatus').textContent = 'Skift kategorien fra direkte beløb til Antal × pris før du tilføjer detaljer.';
        return;
    }
    diverseAdminRows.push({ category, Descr: '', mode: category.startsWith('Skrot ') ? 'pallets' : 'quantity' });
    diverseAdminRender();
}
function diverseAdminUseTemplate() {
    if (!diverseAdminMonth) return;
    const category = document.getElementById('diverseAdminCategory').value;
    const rows = diverseAdminTemplate.filter(row => row.category === category);
    if (!rows.length) return;
    if (!confirm('Erstat linjerne for ' + category + ' med skabelonens tomme detaljer? Eksisterende beløb i denne kategori erstattes først når du gemmer.')) return;
    diverseAdminCollect();
    diverseAdminRows = [...diverseAdminRows.filter(row => row.category !== category), ...rows.map(row => ({ ...row }))];
    diverseAdminRender();
}
function diverseAdminRemove(index) {
    diverseAdminCollect();
    const row = diverseAdminRows[index];
    if (diverseAdminRows.filter(item => item.category === row.category).length === 1) {
        document.getElementById('diverseAdminStatus').textContent = 'Behold mindst én linje pr. kategori. Indtast 0 hvis kategorien er tom.';
        return;
    }
    diverseAdminRows.splice(index, 1);
    diverseAdminRender();
}
async function diverseAdminSave() {
    const status = document.getElementById('diverseAdminStatus');
    try {
        if (!diverseAdminMonth || document.getElementById('diverseAdminMonth').value !== diverseAdminMonth) throw new Error('Hent den valgte måned først.');
        diverseAdminCollect();
        const response = await fetch('/admin/lagerliste-diverse/' + diverseAdminMonth, { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + String(authToken || '') }, body: JSON.stringify({ rows: diverseAdminRows }) });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'Kunne ikke gemme');
        diverseAdminRows = data.data.rows;
        diverseAdminRender();
        status.textContent = 'Gemt i GOH: ' + diverseAdminMonth + '. ' + (diverseAdminRows.every(row => row.complete) ? 'Alle manuelle linjer udfyldt.' : 'Udfyld de tomme felter før månedslukning.') + ' Åbn måneden i Lagerliste: Diverse vises som månedens tillæg. Original lukning bevares.';
        if (typeof lagerlisteLiveCurrent !== 'undefined') lagerlisteLiveCurrent = null;
        if (typeof lagerlisteDisplayedLabel !== 'undefined' && lagerlisteDisplayedLabel === 'Måned ' + diverseAdminMonth) {
            try {
                const period = await lagerlisteResolvePeriod('month:' + diverseAdminMonth);
                lagerlisteRender(period.payload, null, period.label);
            } catch (err) { status.textContent += ' Rapporten kunne ikke opdateres: ' + err.message; }
        }
    } catch (err) { status.textContent = err.message; }
}
