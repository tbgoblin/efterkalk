const crypto = require('crypto');

const STATE_KEY = 'mileage_v1';
const fail = (message, statusCode = 400) => Object.assign(new Error(message), { statusCode });
const manager = user => user?.role === 'superadmin' || user?.permissions?.mileageManage === true;
const username = user => String(user?.username || '').trim().toLowerCase();
const round = value => Math.round((value + Number.EPSILON) * 100) / 100;
const validDate = value => typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)
    && Number.isFinite(Date.parse(value)) && new Date(value).toISOString().slice(0, 10) === value;

function text(value, label, max = 500, required = true) {
    if (typeof value !== 'string' || value.trim().length > max || (required && !value.trim())) throw fail(label + ' mangler eller er ugyldig.');
    return value.trim();
}

function positive(value, label, max = 100000) {
    if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0 || value > max) throw fail(label + ' skal være et positivt tal.');
    return value;
}

function calculateAllowance(km, previousKm, rate) {
    const highKm = round(Math.min(km, Math.max(0, 20000 - previousKm)));
    const lowKm = round(km - highKm);
    return { previousKm: round(previousKm), highKm, lowKm, highRate: rate.high, lowRate: rate.low,
        amount: round(highKm * rate.high + lowKm * rate.low), year: rate.year };
}

function createMileageService({ gohData, readUsers, now = () => new Date().toISOString(), id = () => crypto.randomUUID() }) {
    async function load() {
        const saved = await gohData.getAppState(STATE_KEY, { strict: true });
        const state = saved?.payload || { schema: 1, version: 0, employees: [], addresses: [], purposes: [], rates: [], trips: [], imports: [] };
        if (state.schema !== 1 || !Number.isSafeInteger(state.version)
            || ['employees', 'addresses', 'purposes', 'rates', 'trips', 'imports'].some(key => !Array.isArray(state[key]))) {
            throw fail('Kørselsdata er ugyldige. Kontakt administratoren.', 503);
        }
        return structuredClone(state);
    }

    async function persist(state, version, user) {
        state.version = version + 1;
        state.updatedAt = now();
        state.updatedBy = username(user);
        if (!await gohData.setAppState(STATE_KEY, state, { expectedVersion: version })) {
            throw fail('Kunne ikke gemme i GOH. Data kan være ændret af en anden bruger. Genindlæs og prøv igen.', 409);
        }
        return state;
    }

    function checkVersion(state, version) {
        if (!Number.isSafeInteger(version) || version !== state.version) throw fail('Data er ændret. Genindlæs og prøv igen.', 409);
    }

    function ownEmployee(state, user) {
        return state.employees.find(row => row.username === username(user) && row.active);
    }

    function requireManager(user) {
        if (!manager(user)) throw fail('Adgang til kørselsadministration kræves.', 403);
    }

    function history(trip, action, user, note = '') {
        trip.history.push({ action, by: username(user), at: now(), note });
    }

    async function data(user) {
        const state = await load();
        const canManage = manager(user);
        const employee = ownEmployee(state, user);
        return { version: state.version, user: { username: username(user), displayName: user.displayName || user.username },
            canManage, employeeId: employee?.id || null, employees: canManage ? state.employees : employee ? [employee] : [],
            addresses: state.addresses, purposes: state.purposes, rates: state.rates,
            trips: state.trips.filter(row => canManage || row.employeeId === employee?.id).map(trip => {
                if (trip.legacy || trip.allowance) return trip;
                const year = Number(trip.date?.slice(0, 4));
                const rate = state.rates.find(row => row.year === year);
                const historyIncomplete = state.trips.some(row => row.employeeId === trip.employeeId && row.legacy && !row.date);
                const previousKm = state.trips.filter(row => row.employeeId === trip.employeeId
                    && ['approved', 'paid', 'legacy'].includes(row.status) && Number(row.date?.slice(0, 4)) === year)
                    .reduce((sum, row) => sum + row.km, 0);
                return { ...trip, estimate: rate && !historyIncomplete ? calculateAllowance(trip.km, previousKm, rate) : null };
            }),
            users: canManage ? readUsers().filter(row => row.active !== false).map(row => ({ username: row.username, name: row.displayName || row.username })) : [] };
    }

    async function saveEmployee(input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const name = text(input.name, 'Navn', 200);
        const linkedUser = text(input.username || '', 'GOH-bruger', 100, false).toLowerCase();
        if (linkedUser && !readUsers().some(row => username(row) === linkedUser && row.active !== false)) throw fail('GOH-brugeren findes ikke eller er inaktiv.');
        if (linkedUser && state.employees.some(row => row.id !== input.id && row.username === linkedUser)) throw fail('GOH-brugeren er allerede knyttet til en medarbejder.');
        const employee = input.id ? state.employees.find(row => row.id === input.id) : { id: id() };
        if (!employee) throw fail('Medarbejderen findes ikke.', 404);
        Object.assign(employee, { name, username: linkedUser, plate: text(input.plate || '', 'Nummerplade', 30, false), active: input.active === true });
        if (!input.id) state.employees.push(employee);
        return persist(state, input.version, user);
    }

    async function saveRate(input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        if (!Number.isInteger(input.year) || input.year < 2000 || input.year > 2200) throw fail('Ugyldigt satsår.');
        const rate = { year: input.year, high: positive(input.high, 'Sats til og med 20.000 km', 100), low: positive(input.low, 'Sats over 20.000 km', 100) };
        const index = state.rates.findIndex(row => row.year === rate.year);
        if (index < 0) state.rates.push(rate); else state.rates[index] = rate;
        return persist(state, input.version, user);
    }

    async function saveTrip(input, user) {
        const state = await load();
        checkVersion(state, input.version);
        const employee = ownEmployee(state, user);
        if (!employee) throw fail('En administrator skal knytte din GOH-bruger til en aktiv medarbejder.', 403);
        let trip = input.id ? state.trips.find(row => row.id === input.id) : null;
        if (input.id && (!trip || trip.employeeId !== employee.id)) throw fail('Kørslen findes ikke.', 404);
        if (trip && !['draft', 'rejected'].includes(trip.status)) throw fail('Kun kladder og afviste kørsler kan ændres.', 409);
        if (!validDate(input.date) || input.date > now().slice(0, 10)) throw fail('Vælg en gyldig kørselsdato, ikke i fremtiden.');
        if (!Array.isArray(input.stops) || input.stops.length < 2 || input.stops.length > 12) throw fail('Angiv fra, til og højst 10 mellemstop.');
        const stops = input.stops.map(stop => {
            if (!stop || typeof stop !== 'object' || Array.isArray(stop)) throw fail('Ugyldigt stop.');
            return { address: text(stop.address, 'Adresse', 500), note: text(stop.note || '', 'Stop/pausenote', 500, false) };
        });
        const km = round(positive(input.km, 'Samlet antal km', 10000));
        if (km <= 0) throw fail('Samlet antal km er for lille.');
        const purpose = text(input.purpose, 'Formål', 2000);
        const plate = text(input.plate, 'Nummerplade', 30);
        const manualReason = text(input.manualReason || '', 'Afstandsgrundlag', 1000, false);
        if (!manualReason) throw fail('Angiv hvordan km er fastlagt, eller hvorfor de er rettet.');
        const saved = { date: input.date, stops, km, purpose, plate, roundTrip: input.roundTrip === true,
            comment: text(input.comment || '', 'Kommentar', 2000, false), manualReason };
        if (!trip) {
            trip = { id: id(), employeeId: employee.id, employeeName: employee.name, createdBy: username(user), createdAt: now(), history: [], status: 'draft' };
            state.trips.push(trip);
        }
        Object.assign(trip, saved, { status: 'draft', updatedAt: now() });
        history(trip, 'save', user);
        return persist(state, input.version, user);
    }

    async function transition(tripId, input, user) {
        const state = await load();
        checkVersion(state, input.version);
        const trip = state.trips.find(row => row.id === tripId);
        if (!trip) throw fail('Kørslen findes ikke.', 404);
        const employee = ownEmployee(state, user);
        const own = employee?.id === trip.employeeId;
        const action = input.action;
        if (action === 'submit') {
            if (!own || !['draft', 'rejected'].includes(trip.status)) throw fail('Kørslen kan ikke indsendes.', 403);
            if (!state.rates.some(row => row.year === Number(trip.date.slice(0, 4)))) throw fail('Årets satser skal oprettes før indsendelse.');
            trip.status = 'submitted';
        } else {
            requireManager(user);
            if (trip.legacy) throw fail('Importeret historik kan ikke godkendes eller udbetales igen.', 409);
            if (action === 'approve') {
                if (own || trip.createdBy === username(user)) throw fail('Du kan ikke godkende din egen kørsel.', 403);
                if (trip.status !== 'submitted') throw fail('Kun indsendte kørsler kan godkendes.', 409);
                const year = Number(trip.date.slice(0, 4));
                const rate = state.rates.find(row => row.year === year);
                if (!rate) throw fail('Satser mangler for kørselsåret.');
                const undated = state.trips.some(row => row.employeeId === trip.employeeId && row.legacy && !row.date);
                if (undated) throw fail('Medarbejderens historik har en kørsel uden dato. Ret importgrundlaget før godkendelse.');
                const previousKm = state.trips.filter(row => row.employeeId === trip.employeeId
                    && ['approved', 'paid', 'legacy'].includes(row.status) && Number(row.date?.slice(0, 4)) === year)
                    .reduce((sum, row) => sum + row.km, 0);
                trip.allowance = calculateAllowance(trip.km, previousKm, rate);
                trip.approvedBy = username(user);
                trip.approvedAt = now();
                trip.status = 'approved';
            } else if (action === 'reject') {
                if (own || trip.createdBy === username(user) || trip.status !== 'submitted') throw fail('Kørslen kan ikke afvises.', 409);
                text(input.note, 'Begrundelse', 1000);
                trip.status = 'rejected';
            } else if (action === 'pay') {
                if (trip.status !== 'approved' || !trip.allowance) throw fail('Kun godkendte kørsler kan markeres udbetalt.', 409);
                trip.status = 'paid';
            } else throw fail('Ugyldig handling.');
        }
        history(trip, action, user, text(input.note || '', 'Kommentar', 1000, false));
        trip.updatedAt = now();
        return persist(state, input.version, user);
    }

    async function document(tripId, user) {
        const state = await load();
        const employee = ownEmployee(state, user);
        const trip = state.trips.find(row => row.id === tripId && (manager(user) || row.employeeId === employee?.id));
        if (!trip) throw fail('Kørslen findes ikke.', 404);
        return trip;
    }

    async function importData(input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const { parseMileageImport } = require('./mileageImportService');
        const imported = parseMileageImport(input.files);
        if (state.imports.length) throw fail('Historik er allerede importeret. Gentagen import er blokeret for at undgå dubletter.', 409);
        if (state.trips.length || state.employees.length) throw fail('Import kræver tomme kørselsdata. Importér historikken før nye medarbejdere og kørsler.', 409);
        Object.assign(state, { employees: imported.employees, trips: imported.trips, addresses: imported.addresses, purposes: imported.purposes });
        for (const trip of state.trips) history(trip, 'import', user, 'Power Apps-id: ' + trip.sourceId);
        state.imports.push({ at: now(), by: username(user), count: imported.trips.length, warnings: imported.warnings });
        await persist(state, input.version, user);
        return { warnings: imported.warnings };
    }

    async function correctLegacyDate(tripId, input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const trip = state.trips.find(row => row.id === tripId && row.legacy);
        if (!trip || trip.date) throw fail('Kun en manglende historisk dato kan suppleres.', 409);
        if (!validDate(input.date) || input.date > now().slice(0, 10)) throw fail('Ugyldig historisk dato.');
        trip.date = input.date;
        history(trip, 'legacy-date', user, 'Manglende dato suppleret efter import.');
        return persist(state, input.version, user);
    }

    return { data, saveEmployee, saveRate, saveTrip, transition, document, importData, correctLegacyDate };
}

module.exports = { createMileageService, calculateAllowance, manager, fail, validDate };
