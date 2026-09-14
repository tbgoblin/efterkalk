let diverseAdminRows = [];
let diverseAdminMonth = null;
function diverseAdminRender() {
    const root = document.getElementById('diverseAdminRows');
    const input = (row, key, disabled = false) => '<input style="width:90px" type="number" min="0" step="any" data-field="' + key + '" value="' + (row[key] ?? '') + '"' + (disabled ? ' disabled' : '') + '>';
    root.innerHTML = '<table><thead><tr><th>Kategori / beskrivelse</th><th>Beregning</th><th>Beløb</th><th>Antal / paller</th><th>Kg/palle</th><th>Pris / kg-pris</th></tr></thead><tbody>'
        + diverseAdminRows.map((row, index) => '<tr data-index="' + index + '"><td>' + lagerlisteEscape(row.category)
            + '<br><input data-field="Descr" value="' + lagerlisteEscape(row.Descr) + '"></td><td><select data-field="mode"' + (row.category.startsWith('Skrot ') ? ' disabled' : '') + '>'
            + ['amount', 'quantity', 'pallets'].filter(mode => mode !== 'pallets' || row.category.startsWith('Skrot ')).map(mode => '<option value="' + mode + '"' + (row.mode === mode ? ' selected' : '') + '>' + ({ amount: 'Direkte beløb', quantity: 'Antal × pris', pallets: 'Paller × kg × pris' })[mode] + '</option>').join('')
            + '</select></td><td>' + input(row, 'amount', row.mode !== 'amount') + '</td><td>' + input(row, 'quantity', row.mode === 'amount') + '</td><td>' + input(row, 'kg', row.mode !== 'pallets') + '</td><td>' + input(row, 'price', row.mode === 'amount') + '</td><td><button type="button" onclick="diverseAdminRemove(' + index + ')">Fjern</button></td></tr>').join('') + '</tbody></table>';
    root.querySelectorAll('[data-field="mode"]').forEach(select => select.onchange = () => { diverseAdminCollect(); diverseAdminRender(); });
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
        diverseAdminMonth = month;
        diverseAdminRender();
        status.textContent = 'Viser ' + month + '. ' + (data.data.saved ? 'Gemte værdier.' : 'Ikke udfyldt endnu.');
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
        status.textContent = 'Gemt i GOH: ' + diverseAdminMonth + '. ' + (diverseAdminRows.every(row => row.complete) ? 'Alle manuelle linjer udfyldt.' : 'Udfyld de tomme felter før månedslukning.') + ' Opdater Lagerliste. Eksisterende lukninger ændres ikke.';
    } catch (err) { status.textContent = err.message; }
}
