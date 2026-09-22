const crypto = require('crypto');
const { parseFormula, evaluateFormula } = require('./bilancioFormula');

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
    const section = (rows, label, externalIds = new Set()) => {
        if (!Array.isArray(rows) || rows.length > 150 || !rows.length) invalid(label + ': tilføj mindst én række (maks. 150).');
        const prior = new Map();
        return rows.map(row => {
            if (!row || !/^[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(row.id) || prior.has(row.id)) invalid(label + ': ugyldigt eller gentaget række-id.');
            const result = { id: row.id, name: text(row.name, 'Navn'), type: row.type };
            result.visibility = row.visibility || 'both';
            if (!['both', 'period', 'ytd', 'hidden'].includes(result.visibility)) invalid('Ugyldig visning.');
            result.bold = row.bold === true;
            result.checkZero = row.checkZero === true;
            if (row.type === 'formula') {
                let parsed;
                try { parsed = parseFormula(row.formula); } catch (err) { invalid(result.name + ': ' + err.message); }
                for (const id of parsed.refs) {
                    const local = prior.get(id);
                    if ((!local || local.type === 'heading') && !externalIds.has(id)) invalid(result.name + ': referér kun til talrækker ovenfor' + (externalIds.size ? ' eller i resultatopgørelsen' : '') + '.');
                }
                result.formula = row.formula;
            } else if (row.type === 'lager' || row.type === 'manual') {
                result.period = row.period || 'selected';
                if (!['selected', 'previous', 'current', 'fixed'].includes(result.period)) invalid('Ugyldig dataperiode.');
                if (result.period === 'fixed') {
                    if (!/^20\d{2}-(0[1-9]|1[0-2])$/.test(row.month || '')) invalid('Vælg fast måned.');
                    result.month = row.month;
                }
                if (row.type === 'lager') {
                    if (!['sales', 'cost', 'margin', 'warehouse', 'via'].includes(row.metric)) invalid('Vælg Lagerliste-værdi.');
                    result.metric = row.metric;
                    result.currentSales = row.currentSales === true;
                } else {
                    if (!row.monthValues || typeof row.monthValues !== 'object' || Array.isArray(row.monthValues) || Object.keys(row.monthValues).length > 600) invalid('Ugyldige månedsbeløb.');
                    result.monthValues = {};
                    for (const [month, value] of Object.entries(row.monthValues)) {
                        if (!/^20\d{2}-(0[1-9]|1[0-2])$/.test(month) || typeof value !== 'number' || !Number.isFinite(value)) invalid('Brug YYYY-MM og et gyldigt beløb.');
                        result.monthValues[month] = value;
                    }
                }
            } else
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
    const pnlIds = new Set(pnl.filter(row => row.type !== 'heading').map(row => row.id));
    const balance = section(input.balance, 'Balance', pnlIds);
    for (const key of ['revenueRow', 'beforeTaxRow']) {
        if (!pnl.some(row => row.id === input[key] && row.type !== 'heading')) invalid('Vælg rækken for ' + (key === 'revenueRow' ? 'omsætning' : 'resultat før skat') + '.');
    }
    const reportId = input.reportId || 'default';
    if (typeof reportId !== 'string' || !/^[a-zA-Z][a-zA-Z0-9_-]{0,39}$/.test(reportId)) invalid('Ugyldigt rapport-id.');
    return { schema: 1, version: input.version, reportId, reportName: text(input.reportName || 'Økonomirapport', 'Rapportnavn'), showPnl: input.showPnl !== false, legacyPeriodRows: input.legacyPeriodRows !== false, revenueRow: input.revenueRow, beforeTaxRow: input.beforeTaxRow,
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
function evaluateRows(rows, records, dimension, fields, external = new Map(), crossSection = new Map()) {
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
            amounts = amounts.map((_, i) => row.sources.some(id => values.get(id)[i] == null) ? null : row.sources.reduce((sum, id) => sum + values.get(id)[i], 0) * (row.type === 'percent' ? row.rate / 100 : 1));
        } else if (row.type === 'formula') {
            const { tree } = parseFormula(row.formula);
            amounts = amounts.map((_, i) => evaluateFormula(tree, id => values.has(id) ? values.get(id)[i] : (crossSection.get(id) || [])[i] ?? null));
        } else if (row.type === 'lager' || row.type === 'manual') {
            amounts = external.get(row.id) || Array(dimension).fill(null);
        }
        values.set(row.id, amounts);
        return { id: row.id, name: row.name, visibility: row.visibility || 'both', bold: row.bold, checkZero: row.checkZero, type: ({ accounts: 'group', sum: 'subtotal', percent: 'computed', heading: 'heading', formula: 'computed', lager: 'computed', manual: 'computed' })[row.type], accounts: details, amounts };
    });
}
function createDefinitionStore({ gohData, getProfile, defaultDefinition }) {
    const key = () => {
        const p = getProfile();
        return 'bilancio_definition_' + crypto.createHash('sha256').update(JSON.stringify([p.server.toLowerCase(), p.database.toLowerCase()])).digest('hex').slice(0, 32);
    };
    function storageKey(id = 'default') {
        if (typeof id !== 'string' || !/^[a-zA-Z][a-zA-Z0-9_-]{0,39}$/.test(id)) invalid('Ugyldigt rapport-id.');
        return id === 'default' ? key() : key() + '_report_' + id;
    }
    async function load(reportId = 'default') {
        const scope = key();
        const state = await gohData.getAppState(storageKey(reportId), { strict: true });
        if (!state) {
            if (reportId !== 'default') { const err = new Error('Rapporten findes ikke.'); err.statusCode = 404; throw err; }
            return { ...structuredClone(defaultDefinition), reportId, reportName: 'Økonomirapport', scope };
        }
        return { ...validateDefinition(state.payload), reportId, scope, updatedBy: state.payload.updatedBy, updatedAt: state.payload.updatedAt };
    }
    async function list() {
        const prefix = key() + '_report_';
        const keys = await gohData.getAppStateKeysByPrefix(prefix);
        if (!Array.isArray(keys)) throw new Error('Rapportlisten kunne ikke hentes fra GOH. Prøv igen.');
        const reports = await Promise.all(['default', ...keys.map(row => row.key.slice(prefix.length))].map(id => load(id)));
        return reports.map(r => ({ reportId: r.reportId, reportName: r.reportName }));
    }
    async function save(input, catalog, username) {
        const scope = key();
        if (input?.scope !== scope) invalid('Databaseprofilen er ændret. Genindlæs opsætningen.');
        const config = validateDefinition(input);
        compileDefinition(config, catalog);
        const payload = { ...config, scope, version: config.version + 1, updatedBy: username, updatedAt: new Date().toISOString() };
        if (!await gohData.setAppState(storageKey(config.reportId), payload, { expectedVersion: config.version })) {
            const error = new Error('Opsætningen blev ikke gemt. En anden bruger kan have ændret den, eller GOH er utilgængelig. Genindlæs før du prøver igen.');
            error.statusCode = 409; throw error;
        }
        return payload;
    }
    return { load, save, list, scope: key };
}
module.exports = { defaults, validateDefinition, compileDefinition, evaluateRows, createDefinitionStore };
