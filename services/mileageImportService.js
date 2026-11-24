const crypto = require('crypto');
const { fail, validDate } = require('./mileageService');

function parseCsv(input) {
    if (typeof input !== 'string' || input.length > 2000000) throw fail('CSV-filen mangler eller er for stor (maks. 2 MB).');
    const text = input.replace(/^\uFEFF/, '');
    const separator = text.slice(0, text.indexOf('\n')).includes(';') ? ';' : ',';
    const rows = [];
    let row = [], value = '', quoted = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === '"') {
            if (quoted && text[i + 1] === '"') { value += '"'; i++; }
            else if (!quoted && value.length) throw fail('Ugyldige anførselstegn i CSV.');
            else quoted = !quoted;
        } else if (!quoted && (ch === separator || ch === '\n')) {
            row.push(value.replace(/\r$/, '')); value = '';
            if (ch === '\n') { if (row.some(Boolean)) rows.push(row); row = []; }
        } else value += ch;
    }
    if (quoted) throw fail('Uafsluttet tekstfelt i CSV.');
    row.push(value.replace(/\r$/, ''));
    if (row.some(Boolean)) rows.push(row);
    const headers = rows.shift();
    if (!headers?.length || new Set(headers).size !== headers.length) throw fail('CSV mangler entydige kolonnenavne.');
    return rows.map((cells, i) => {
        if (cells.length !== headers.length) throw fail('Forkert antal CSV-felter på række ' + (i + 2) + '.');
        return Object.fromEntries(headers.map((key, index) => [key, cells[index].trim()]));
    });
}

const key = value => String(value || '').normalize('NFC').trim().toLowerCase();
const stableId = value => crypto.createHash('sha256').update(key(value)).digest('hex').slice(0, 32);
function boolean(value) {
    if (['sandt', 'true', '1'].includes(key(value))) return true;
    if (['falsk', 'false', '0'].includes(key(value))) return false;
    throw fail('Ugyldig ja/nej-værdi i importen.');
}
function decimal(value) {
    const raw = String(value).trim();
    const parsed = Number(raw.includes(',') ? raw.replace(/\./g, '').replace(',', '.') : raw);
    if (!raw || !Number.isFinite(parsed) || parsed < 0 || parsed > 10000) throw fail('Ugyldigt antal km i importen.');
    return parsed;
}
function date(value) {
    if (!value) return null;
    const match = /^(\d{2})-(\d{2})-(\d{4})$/.exec(value);
    const iso = match ? `${match[3]}-${match[2]}-${match[1]}` : value;
    if (!validDate(iso)) throw fail('Ugyldig dato i historikken.');
    return iso;
}
function requireColumns(rows, columns, label) {
    if (!rows.length || columns.some(col => !Object.hasOwn(rows[0], col))) throw fail(label + ': CSV-kolonner mangler (' + columns.join(', ') + ').');
}
function parseMileageImport(files) {
    if (!files || typeof files !== 'object') throw fail('Vælg de fire CSV-filer.');
    const employeeRows = parseCsv(files.employees);
    const tripRows = parseCsv(files.trips);
    const addressRows = parseCsv(files.addresses);
    const purposeRows = parseCsv(files.purposes);
    requireColumns(employeeRows, ['Bruger', 'Nummerplade', 'Aktiv'], 'Medarbejdere');
    requireColumns(tripRows, ['Id', 'Navn', 'Status', 'Nr_Plade', 'Dato_Dato', 'Beregnet_km', 'Retur', 'Udbetalt', 'Formaal'], 'Kørsler');
    requireColumns(addressRows, ['Titel', 'Addresse', 'Firmanavn'], 'Adresser');
    requireColumns(purposeRows, ['Titel', 'Formaal'], 'Formål');
    const employees = employeeRows.map(row => {
        if (!row.Bruger) throw fail('Medarbejdernavn mangler i CSV.');
        return { id: stableId(row.Bruger), name: row.Bruger, plate: row.Nummerplade, active: boolean(row.Aktiv), username: '' };
    });
    if (new Set(employees.map(row => row.id)).size !== employees.length) throw fail('Dublerede medarbejdernavne i CSV. Ret før import.');
    const byName = new Map(employees.map(row => [key(row.name), row]));
    const trips = tripRows.map(row => {
        const employee = byName.get(key(row.Navn));
        if (!employee) throw fail('Historikken indeholder medarbejdere, som ikke findes i medarbejderfilen.');
        if (!row.Id || !['godkendt', 'udbetalt'].includes(key(row.Status))) throw fail('Historikken har en ukendt status eller manglende ID; den skal afstemmes før import.');
        return { id: 'legacy-' + stableId(row.Id), legacy: true, sourceId: row.Id, employeeId: employee.id, employeeName: employee.name,
            date: date(row.Dato_Dato), km: decimal(row.Beregnet_km), plate: row.Nr_Plade, purpose: row.Formaal,
            legacyCreatedBy: row['Oprettet af'] || row['Oprattet af'] || '', legacyCreatedAt: row.Oprettet || '',
            comment: row.BrugerKommentar || '', roundTrip: boolean(row.Retur), stops: row.Fra && row.Til ? [{ address: row.Fra, note: '' }, { address: row.Til, note: '' }] : [],
            status: 'legacy', legacyPaid: boolean(row.Udbetalt), allowance: null, history: [],
            manualReason: 'Importeret beregnet km fra Power Apps; retur er allerede indeholdt i det historiske km-tal.' };
    });
    if (new Set(trips.map(row => row.id)).size !== trips.length) throw fail('Dublerede kørsels-ID i CSV.');
    const warnings = [
        'Historiske km importeres uændret, også ved retur. Ingen historiske beløb er beregnet eller udbetalt igen.',
        'Formål bevares ordret. Kontrollér historikken og tilknyt GOH-brugere før nye kørsler godkendes.',
        'Gamle Admin-markeringer giver ingen GOH-rettigheder.'
    ];
    const undated = trips.filter(row => !row.date).length;
    if (undated) warnings.push(undated + ' historiske kørsler mangler dato; de berørte medarbejderes nye godkendelser blokeres indtil datoen er rettet.');
    return { employees, trips,
        addresses: addressRows.map(row => ({ name: row.Firmanavn || row.Titel, address: row.Addresse })).filter(row => row.address),
        purposes: [...new Set(purposeRows.map(row => row.Formaal || row.Titel).filter(Boolean))], warnings };
}
module.exports = { parseCsv, parseMileageImport };
