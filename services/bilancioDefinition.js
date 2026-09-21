const crypto = require('crypto');

function defaults(lines) {
    const contributions = [];
    const pnl = lines.map((line, index) => {
        const id = 'p' + index;
        if (line.type === 'subtotal') return { id, name: line.name, type: 'sum', sources: [...contributions] };
        if (line.type === 'computed') {
            const row = { id, name: line.name, type: 'percent', sources: ['p16'], rate: line.rate * 100 };
            contributions.push(id); return row;
        }
        contributions.push(id);
        return { id, name: line.name, type: 'accounts', accounts: [], groups: line.codes, exclude: [], children: false, sign: -1 };
    });
    const accountRow = (id, name, accounts, groups = [], children = false) => ({ id, name, type: 'accounts', accounts, groups, exclude: [], children, sign: 1 });
    return { schema: 1, version: 0, revenueRow: 'p0', beforeTaxRow: 'p16', balanceTitle: 'Aktiver', pnl, balance: [
        accountRow('a0', 'Depusitum husleje + driftsmidler', [66980, 66981]),
        accountRow('a1', 'Grunde og bygninger', [], ['46_Grunde_og_bygninger']),
        accountRow('a2', 'Driftsmidler i alt', [], ['47_Driftsmidler_ialt'], true),
        { id: 'a3', name: 'Anlægsaktiver i alt', type: 'sum', sources: ['a0', 'a1', 'a2'] },
        accountRow('a4', 'Varebeholdninger', [61100, 61120, 61850, 61900]),
        accountRow('a5', 'Tilgodehavender fra salg', [66100])
    ] };
}
function invalid(message) { const error = new Error(message); error.statusCode = 400; throw error; }
function validateDefinition(input) {
    if (!input || input.schema !== 1 || !Number.isInteger(input.version) || input.version < 0) invalid('Ugyldig rapportversion. Genindlæs opsætningen.');
    const text = (value, label) => {
        if (typeof value !== 'string' || !value.trim() || value.length > 160) invalid(label + ' mangler eller er for langt.');
        return value.trim();
    };
    const numbers = values => {
        if (!Array.isArray(values) || values.length > 1000 || values.some(n => !Number.isSafeInteger(n) || n <= 0)) invalid('Kontonumre skal være positive heltal.');
        return [...new Set(values)];
    };
    const section = (rows, label) => {
        if (!Array.isArray(rows) || rows.length > 150 || !rows.length) invalid(label + ': tilføj mindst én række (maks. 150).');
        const prior = new Map();
        return rows.map(row => {
            if (!row || !/^[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(row.id) || prior.has(row.id)) invalid(label + ': ugyldigt eller gentaget række-id.');
            const result = { id: row.id, name: text(row.name, 'Navn'), type: row.type };
            if (row.type === 'accounts') {
                if (!Array.isArray(row.groups) || row.groups.length > 200) invalid('Ugyldige kontogrupper.');
                result.accounts = numbers(row.accounts);
                result.groups = [...new Set(row.groups.map(v => text(v, 'Kontogruppe')))];
                result.exclude = numbers(row.exclude);
                result.children = row.children === true;
                if (![1, -1].includes(row.sign)) invalid('Vælg fortegn +1 eller -1.');
                result.sign = row.sign;
                if (!result.accounts.length && !result.groups.length) invalid(result.name + ': vælg konti eller grupper.');
            } else if (['sum', 'percent'].includes(row.type)) {
                if (!Array.isArray(row.sources) || !row.sources.length || new Set(row.sources).size !== row.sources.length) invalid(result.name + ': vælg forskellige rækker til beregningen.');
                for (const id of row.sources) if (!prior.has(id) || prior.get(id).type === 'heading') invalid(result.name + ': beregningen må kun bruge talrækker ovenfor.');
                result.sources = [...row.sources];
                if (row.type === 'percent') {
                    if (typeof row.rate !== 'number' || !Number.isFinite(row.rate) || Math.abs(row.rate) > 1000) invalid('Ugyldig procentsats.');
                    result.rate = row.rate;
                }
                // A subtotal must not add both a subtotal and its underlying rows.
                const leaves = id => {
                    const source = prior.get(id);
                    return source.type === 'sum' ? source.sources.flatMap(leaves) : [id];
                };
                const inputs = result.sources.flatMap(leaves);
                if (new Set(inputs).size !== inputs.length) invalid(result.name + ': den samme række medregnes flere gange via subtotaler.');
            } else if (row.type !== 'heading') invalid('Ukendt rækketype.');
            prior.set(row.id, result); return result;
        });
    };
    const pnl = section(input.pnl, 'Resultatopgørelse');
    const balance = section(input.balance, 'Balance');
    for (const key of ['revenueRow', 'beforeTaxRow']) {
        if (!pnl.some(row => row.id === input[key] && row.type !== 'heading')) invalid('Vælg rækken for ' + (key === 'revenueRow' ? 'omsætning' : 'resultat før skat') + '.');
    }
    return { schema: 1, version: input.version, revenueRow: input.revenueRow, beforeTaxRow: input.beforeTaxRow,
        balanceTitle: text(input.balanceTitle, 'Balancetitel'), pnl, balance };
}
function compileDefinition(input, catalog) {
    const config = validateDefinition(input);
    const accounts = new Map(catalog.accounts.map(row => [Number(row.AcNo), row]));
    const groupCodes = new Set([...catalog.groups.map(row => String(row.AcGr).trim()), ...catalog.accounts.map(row => String(row.AcGr || '').trim())]);
    const compile = rows => {
        const used = new Map();
        return rows.map(row => {
            if (row.type !== 'accounts') return row;
            for (const ac of [...row.accounts, ...row.exclude]) if (!accounts.has(ac)) invalid(row.name + ': konto ' + ac + ' findes ikke.');
            const groups = new Set(row.groups);
            for (const group of groups) if (!groupCodes.has(group)) invalid(row.name + ': kontogruppe ' + group + ' findes ikke.');
            if (row.children) {
                let changed = true;
                while (changed) {
                    changed = false;
                    for (const group of catalog.groups) {
                        const code = String(group.AcGr).trim();
                        if (groups.has(String(group.AgAcGr || '').trim()) && !groups.has(code)) { groups.add(code); changed = true; }
                    }
                }
            }
            const selected = new Set([...row.accounts, ...catalog.accounts.filter(ac => groups.has(String(ac.AcGr || '').trim())).map(ac => Number(ac.AcNo))]);
            row.exclude.forEach(ac => selected.delete(ac));
            for (const ac of selected) {
                if (used.has(ac)) invalid(`Konto ${ac} medregnes i både "${used.get(ac)}" og "${row.name}". Fjern den fra en af rækkerne eller brug Udelad konti.`);
                used.set(ac, row.name);
            }
            return { ...row, selectedAccounts: [...selected].sort((a, b) => a - b) };
        });
    };
    return { ...config, pnl: compile(config.pnl), balance: compile(config.balance) };
}
function evaluateRows(rows, records, dimension, fields) {
    const data = new Map();
    for (const record of records) {
        const ac = Number(record.AcNo);
        if (data.has(ac)) invalid('Dubleret konto i beregningen: ' + ac);
        data.set(ac, record);
    }
    const values = new Map();
    return rows.map(row => {
        let amounts = Array(dimension).fill(0), details = [];
        if (row.type === 'accounts') {
            details = row.selectedAccounts.map(ac => {
                const record = data.get(ac);
                if (!record) invalid('Kontodata mangler for ' + ac);
                const numbers = fields.map(field => {
                    if (record[field] == null || !Number.isFinite(Number(record[field]))) invalid('Ugyldig saldo for ' + ac);
                    return Number(record[field]) * row.sign;
                });
                return { account: ac, name: String(record.Nm || '').trim(), amounts: numbers };
            });
            amounts = amounts.map((_, i) => details.reduce((sum, d) => sum + d.amounts[i], 0));
        } else if (row.type === 'sum' || row.type === 'percent') {
            amounts = amounts.map((_, i) => row.sources.reduce((sum, id) => sum + values.get(id)[i], 0) * (row.type === 'percent' ? row.rate / 100 : 1));
        }
        values.set(row.id, amounts);
        return { id: row.id, name: row.name, type: ({ accounts: 'group', sum: 'subtotal', percent: 'computed', heading: 'heading' })[row.type], accounts: details, amounts };
    });
}
function createDefinitionStore({ gohData, getProfile, defaultDefinition }) {
    const key = () => {
        const p = getProfile();
        return 'bilancio_definition_' + crypto.createHash('sha256').update(JSON.stringify([p.server.toLowerCase(), p.database.toLowerCase()])).digest('hex').slice(0, 32);
    };
    async function load() {
        const scope = key();
        const state = await gohData.getAppState(scope, { strict: true });
        if (!state) return { ...structuredClone(defaultDefinition), scope };
        return { ...validateDefinition(state.payload), scope, updatedBy: state.payload.updatedBy, updatedAt: state.payload.updatedAt };
    }
    async function save(input, catalog, username) {
        const scope = key();
        if (input?.scope !== scope) invalid('Databaseprofilen er ændret. Genindlæs opsætningen.');
        const config = validateDefinition(input);
        compileDefinition(config, catalog);
        const payload = { ...config, scope, version: config.version + 1, updatedBy: username, updatedAt: new Date().toISOString() };
        if (!await gohData.setAppState(scope, payload, { expectedVersion: config.version })) {
            const error = new Error('Opsætningen blev ikke gemt. En anden bruger kan have ændret den, eller GOH er utilgængelig. Genindlæs før du prøver igen.');
            error.statusCode = 409; throw error;
        }
        return payload;
    }
    return { load, save, scope: key };
}
module.exports = { defaults, validateDefinition, compileDefinition, evaluateRows, createDefinitionStore };
