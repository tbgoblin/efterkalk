const crypto = require('crypto');

const STATE_KEY = 'mileage_v1';
const fail = (message, statusCode = 400) => Object.assign(new Error(message), { statusCode });
const manager = user => user?.role === 'superadmin' || user?.permissions?.mileageManage === true;
const username = user => String(user?.username || '').trim().toLowerCase();
const nameKey = value => String(value || '').normalize('NFC').trim().toLowerCase();
const round = value => Math.round((value + Number.EPSILON) * 100) / 100;
const validDate = value => typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)
    && Number.isFinite(Date.parse(value)) && new Date(value).toISOString().slice(0, 10) === value;
const validCoordinates = value => Array.isArray(value) && value.length === 2
    && value.every(n => typeof n === 'number' && Number.isFinite(n)) && Math.abs(value[0]) <= 180 && Math.abs(value[1]) <= 90;

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

    function userNames(users) {
        return new Map(users.map(row => [username(row), row.displayName || row.name || row.username]));
    }

    function savedAddressNames(state) {
        const names = new Map();
        for (const address of state.addresses) {
            const key = nameKey(address.address);
            if (!names.has(key)) names.set(key, address.name);
            else if (names.get(key) !== address.name) names.set(key, null);
        }
        return names;
    }

    function addDisplayNames(trip, names, addressNames) {
        const stops = (trip.stops || []).map(stop => {
            const name = stop.name || addressNames.get(nameKey(stop.address));
            return name ? { ...stop, name } : stop;
        });
        const history = (trip.history || []).map(item => ({ ...item, byName: names.get(username({ username: item.by })) || item.by }));
        return {
            ...trip,
            createdByName: names.get(username({ username: trip.createdBy })) || trip.createdBy,
            approvedByName: names.get(username({ username: trip.approvedBy })) || trip.approvedBy,
            stops,
            history
        };
    }

    function requireManager(user) {
        if (!manager(user)) throw fail('Adgang til kørselsadministration kræves.', 403);
    }

    function history(trip, action, user, note = '') {
        trip.history.push({ action, by: username(user), at: now(), note });
    }

    function markForReview(state, employeeId, year, user, reason, excludedId, marked = new Set()) {
        for (const trip of state.trips) {
            if (trip.id !== excludedId && trip.employeeId === employeeId && ['approved', 'paid'].includes(trip.status)
                && (!year || trip.date?.slice(0, 4) === year) && !marked.has(trip.id)) {
                trip.reconciliationRequired = true;
                history(trip, 'review-required', user, reason);
                marked.add(trip.id);
            }
        }
        return marked;
    }

    async function data(user) {
        const state = await load();
        const canManage = manager(user);
        const employee = ownEmployee(state, user);
        const users = readUsers();
        const names = userNames(users);
        const addressNames = savedAddressNames(state);
        const trips = state.trips.filter(row => canManage || (row.status !== 'deleted'
            && (row.employeeId === employee?.id || row.createdBy === username(user)))).map(trip => {
            const visibleTrip = addDisplayNames(trip, names, addressNames);
            if (!canManage) {
                delete visibleTrip.allowance;
                delete visibleTrip.estimate;
                return visibleTrip;
            }
            if (trip.legacy || trip.allowance) return visibleTrip;
            const year = Number(trip.date?.slice(0, 4));
            const rate = state.rates.find(row => row.year === year);
            const previousKm = state.trips.filter(row => row.employeeId === trip.employeeId
                && ['approved', 'paid'].includes(row.status) && Number(row.date?.slice(0, 4)) === year)
                .reduce((sum, row) => sum + row.km, 0);
            return { ...visibleTrip, estimate: rate ? calculateAllowance(trip.km, previousKm, rate) : null };
        });
        return { version: state.version, user: { username: username(user), displayName: user.displayName || user.username },
            isSuperadmin: user?.role === 'superadmin',
            canManage, employeeId: employee?.id || null, employees: canManage ? state.employees : employee ? [employee] : [],
            participants: state.employees.filter(row => row.active).map(row => ({ id: row.id, name: row.name, vismaEmpNo: row.vismaEmpNo || null })),
            addresses: state.addresses, purposes: state.purposes, rates: canManage ? state.rates : [],
            trips,
            users: canManage ? users.filter(row => row.active !== false).map(row => ({ username: row.username, name: row.displayName || row.username })) : [] };
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
        if (employee.vismaEmpNo && (name !== employee.name || input.active !== employee.active)) {
            throw fail('Navn og aktiv status styres af Visma.');
        }
        Object.assign(employee, { name, username: linkedUser, plate: text(input.plate || '', 'Nummerplade', 30, false), active: input.active === true });
        if (!input.id) state.employees.push(employee);
        return persist(state, input.version, user);
    }

    async function syncEmployees(rows, source, user) {
        if (!Array.isArray(rows) || !rows.length || typeof source !== 'string' || !source) {
            throw fail('Visma returnerede ingen gyldig medarbejderliste. Ingen ændringer blev gemt.', 503);
        }
        const incoming = [], ids = new Set(), names = new Map();
        for (const row of rows) {
            const empNo = String(row.EmpNo ?? '').trim();
            const name = typeof row.Nm === 'string' ? row.Nm.trim() : '';
            const group = Number(row.Gr2);
            if (!/^[1-9]\d*$/.test(empNo) || !name || name.length > 200 || row.Gr2 == null || ![0, 99].includes(group)
                || ids.has(empNo)) {
                throw fail('Visma-listen indeholder ugyldige eller dublerede medarbejdernumre. Ingen ændringer blev gemt.', 409);
            }
            ids.add(empNo);
            names.set(nameKey(name), (names.get(nameKey(name)) || 0) + 1);
            incoming.push({ empNo, name, active: group === 0 });
        }
        for (let attempt = 0; attempt < 2; attempt++) {
            const state = await load();
            if (state.vismaSource && state.vismaSource !== source) throw fail('Visma-profilen afviger fra den registrerede medarbejderkilde. Kontakt administratoren.', 409);
            const knownIds = new Map(), existingNames = new Map();
            for (const employee of state.employees) {
                if (employee.vismaEmpNo) {
                    if (knownIds.has(employee.vismaEmpNo)) throw fail('Dublerede Visma-medarbejdernumre i GOH.', 409);
                    knownIds.set(employee.vismaEmpNo, employee);
                }
                const key = nameKey(employee.name);
                if (!existingNames.has(key)) existingNames.set(key, []);
                existingNames.get(key).push(employee);
            }
            let changed = state.vismaSource !== source;
            const ambiguous = new Set();
            for (const row of incoming) {
                let employee = knownIds.get(row.empNo);
                const matches = existingNames.get(nameKey(row.name)) || [];
                if (!employee) {
                    const unlinked = matches.filter(match => !match.vismaEmpNo);
                    if (unlinked.length > 1 || (unlinked.length && names.get(nameKey(row.name)) > 1)) {
                        if (row.active || unlinked.some(match => match.active)) ambiguous.add(row.name);
                        continue;
                    }
                    employee = unlinked[0];
                    if (!employee) {
                        if (!row.active) continue;
                        employee = { id: id(), name: row.name, username: '', plate: '', active: row.active };
                        state.employees.push(employee);
                        if (!existingNames.has(nameKey(row.name))) existingNames.set(nameKey(row.name), []);
                        existingNames.get(nameKey(row.name)).push(employee);
                    }
                    employee.vismaEmpNo = row.empNo;
                    knownIds.set(row.empNo, employee);
                    changed = true;
                }
                if (employee.name !== row.name || employee.active !== row.active) {
                    employee.name = row.name;
                    employee.active = row.active;
                    changed = true;
                }
            }
            const warnings = ambiguous.size ? [
                ambiguous.size + ' medarbejdernavne kan ikke knyttes sikkert til Visma pga. navnesammenfald.'
                + (manager(user) ? ' Kontrollér medarbejderne i opsætningen: ' + [...ambiguous].join(', ') + '.' : ' Kontakt administratoren.')
            ] : [];
            if (!changed) return { changed: false, warnings };
            state.vismaSource = source;
            const version = state.version;
            try {
                await persist(state, version, user);
                return { changed: true, warnings };
            } catch (error) {
                if (attempt || error.statusCode !== 409) throw error;
            }
        }
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

    async function saveAddress(input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const name = text(input.name, 'Navn', 200);
        const address = text(input.address, 'Adresse', 500);
        if (input.coordinates != null && !validCoordinates(input.coordinates)) throw fail('Ugyldige adressekoordinater.');
        const editing = input.originalName != null || input.originalAddress != null;
        if (editing && (input.originalName == null || input.originalAddress == null)) throw fail('Den oprindelige adresse mangler.');
        const originalName = editing ? text(input.originalName, 'Oprindeligt navn', 200, false) : null;
        const originalAddress = editing ? text(input.originalAddress, 'Oprindelig adresse', 500) : null;
        let saved = editing
            ? state.addresses.find(row => nameKey(row.name) === nameKey(originalName) && nameKey(row.address) === nameKey(originalAddress))
            : state.addresses.find(row => nameKey(row.name) === nameKey(name) && nameKey(row.address) === nameKey(address));
        if (editing && !saved) throw fail('Adressen findes ikke længere. Genindlæs listen.', 404);
        if (editing && state.addresses.some(row => row !== saved && nameKey(row.name) === nameKey(name) && nameKey(row.address) === nameKey(address))) {
            throw fail('En adresse med dette navn og denne adresse findes allerede.', 409);
        }
        if (!saved) {
            if (state.addresses.length >= 5000) throw fail('Adressebogen er fuld.');
            saved = {};
            state.addresses.push(saved);
        }
        if (saved.coordinates && nameKey(saved.address) !== nameKey(address) && input.coordinates == null) delete saved.coordinates;
        Object.assign(saved, { name, address });
        if (input.coordinates != null) saved.coordinates = input.coordinates.slice();
        saved.updatedAt = now();
        saved.updatedBy = username(user);
        return persist(state, input.version, user);
    }

    async function deleteAddress(input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const name = text(input.name, 'Navn', 200, false);
        const address = text(input.address, 'Adresse', 500);
        const index = state.addresses.findIndex(row => nameKey(row.name) === nameKey(name) && nameKey(row.address) === nameKey(address));
        if (index < 0) throw fail('Adressen findes ikke længere. Genindlæs listen.', 404);
        state.addresses.splice(index, 1);
        return persist(state, input.version, user);
    }

    async function savePurpose(input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const purpose = text(input.purpose, 'Formål', 2000);
        const editing = input.original != null;
        const index = editing ? state.purposes.findIndex(row => nameKey(row) === nameKey(input.original)) : -1;
        if (editing && index < 0) throw fail('Formålet findes ikke længere. Genindlæs listen.', 404);
        if (state.purposes.some((row, rowIndex) => rowIndex !== index && nameKey(row) === nameKey(purpose))) {
            throw fail('Formålet findes allerede.', 409);
        }
        if (!editing && state.purposes.length >= 5000) throw fail('Listen over formål er fuld.');
        if (editing) state.purposes[index] = purpose;
        else state.purposes.push(purpose);
        return persist(state, input.version, user);
    }

    async function deletePurpose(input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const purpose = text(input.purpose, 'Formål', 2000);
        const index = state.purposes.findIndex(row => nameKey(row) === nameKey(purpose));
        if (index < 0) throw fail('Formålet findes ikke længere. Genindlæs listen.', 404);
        state.purposes.splice(index, 1);
        return persist(state, input.version, user);
    }

    async function saveTrip(input, user) {
        const state = await load();
        checkVersion(state, input.version);
        let trip = input.id ? state.trips.find(row => row.id === input.id) : null;
        const own = ownEmployee(state, user);
        if (input.id && (!trip || (trip.employeeId !== own?.id && trip.createdBy !== username(user)))) throw fail('Kørslen findes ikke.', 404);
        if (trip && !['draft', 'rejected'].includes(trip.status)) throw fail('Kun kladder og afviste kørsler kan ændres.', 409);
        if (trip && input.employeeId && input.employeeId !== trip.employeeId) throw fail('Medarbejderen kan ikke ændres på en eksisterende kørsel.');
        const employee = trip ? state.employees.find(row => row.id === trip.employeeId)
            : state.employees.find(row => row.id === (input.employeeId || own?.id) && row.active);
        if (!employee) throw fail('Vælg en aktiv medarbejder.');
        if (!validDate(input.date) || input.date > now().slice(0, 10)) throw fail('Vælg en gyldig kørselsdato, ikke i fremtiden.');
        if (!Array.isArray(input.stops) || input.stops.length < 2 || input.stops.length > 12) throw fail('Angiv fra, til og højst 10 mellemstop.');
        const stops = input.stops.map(stop => {
            if (!stop || typeof stop !== 'object' || Array.isArray(stop)) throw fail('Ugyldigt stop.');
            return { name: text(stop.name || '', 'Virksomhedsnavn', 500, false),
                address: text(stop.address, 'Adresse', 500), note: text(stop.note || '', 'Stop/pausenote', 500, false) };
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
        const creator = trip.createdBy === username(user);
        const action = input.action;
        if (action === 'submit') {
            if ((!own && !creator) || !['draft', 'rejected'].includes(trip.status)) throw fail('Kørslen kan ikke indsendes.', 403);
            if (!state.rates.some(row => row.year === Number(trip.date.slice(0, 4)))) throw fail('Årets satser skal oprettes før indsendelse.');
            trip.status = 'submitted';
        } else {
            requireManager(user);
            if (trip.legacy || trip.status === 'deleted') throw fail('Importeret eller slettet historik kan ikke godkendes eller udbetales igen.', 409);
            if (action === 'approve') {
                if (trip.status !== 'submitted') throw fail('Kun indsendte kørsler kan godkendes.', 409);
                const year = Number(trip.date.slice(0, 4));
                const rate = state.rates.find(row => row.year === year);
                if (!rate) throw fail('Satser mangler for kørselsåret.');
                const previousKm = state.trips.filter(row => row.employeeId === trip.employeeId
                    && ['approved', 'paid'].includes(row.status) && Number(row.date?.slice(0, 4)) === year)
                    .reduce((sum, row) => sum + row.km, 0);
                trip.allowance = calculateAllowance(trip.km, previousKm, rate);
                trip.approvedBy = username(user);
                trip.approvedAt = now();
                trip.status = 'approved';
            } else if (action === 'reject') {
                if (own || creator || trip.status !== 'submitted') throw fail('Kørslen kan ikke afvises.', 409);
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
        const trip = state.trips.find(row => row.id === tripId && row.status !== 'deleted'
            && (manager(user) || row.employeeId === employee?.id || row.createdBy === username(user)));
        if (!trip) throw fail('Kørslen findes ikke.', 404);
        const visibleTrip = addDisplayNames(trip, userNames(readUsers()), savedAddressNames(state));
        if (manager(user)) return visibleTrip;
        delete visibleTrip.allowance;
        delete visibleTrip.estimate;
        return visibleTrip;
    }

    function mergeImport(state, files, user) {
        const { parseMileageImport } = require('./mileageImportService');
        const imported = parseMileageImport(files);
        if (state.imports.length) throw fail('Historik er allerede importeret. Gentagen import er blokeret for at undgå dubletter.', 409);
        const names = new Map();
        for (const employee of state.employees) {
            const name = nameKey(employee.name);
            if (!names.has(name)) names.set(name, []);
            names.get(name).push(employee);
        }
        const ids = new Set(state.employees.map(row => row.id));
        const idMap = new Map();
        let matched = 0;
        for (const employee of imported.employees) {
            const matches = names.get(nameKey(employee.name)) || [];
            if (matches.length > 1) throw fail('En importeret medarbejder matcher flere eksisterende medarbejdere. Afklar navnene før import.', 409);
            let existing = matches[0];
            if (existing) matched++;
            if (!existing) {
                if (ids.has(employee.id)) throw fail('Medarbejder-ID er allerede i brug. Afklar før import.', 409);
                existing = employee;
                state.employees.push(existing);
                names.set(nameKey(existing.name), [existing]);
                ids.add(existing.id);
            }
            idMap.set(employee.id, existing.id);
        }
        const tripIds = new Set(state.trips.map(row => row.id));
        for (const trip of imported.trips) {
            if (tripIds.has(trip.id)) throw fail('Et historisk kørsels-ID findes allerede. Importen er stoppet for at undgå dubletter.', 409);
            trip.employeeId = idMap.get(trip.employeeId);
            trip.employeeName = state.employees.find(row => row.id === trip.employeeId).name;
            history(trip, 'import', user, 'Power Apps-id: ' + trip.sourceId);
            state.trips.push(trip);
        }
        const addresses = new Set(state.addresses.map(row => nameKey(row.address)));
        for (const row of imported.addresses) if (!addresses.has(nameKey(row.address))) {
            state.addresses.push(row); addresses.add(nameKey(row.address));
        }
        const purposes = new Set(state.purposes.map(nameKey));
        for (const purpose of imported.purposes) if (!purposes.has(nameKey(purpose))) {
            state.purposes.push(purpose); purposes.add(nameKey(purpose));
        }
        imported.warnings.push(matched + ' medarbejdere er matchet på navn; ' + (imported.employees.length - matched) + ' nye medarbejdere oprettes. Kontrollér navne og GOH-tilknytninger.');
        state.imports.push({ at: now(), by: username(user), count: imported.trips.length, warnings: imported.warnings });
        return imported;
    }

    async function previewImport(files, user) {
        requireManager(user);
        const imported = mergeImport(await load(), files, user);
        return { counts: { employees: imported.employees.length, trips: imported.trips.length,
            addresses: imported.addresses.length, purposes: imported.purposes.length }, warnings: imported.warnings };
    }

    async function importData(input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const imported = mergeImport(state, input.files, user);
        await persist(state, input.version, user);
        return { warnings: imported.warnings };
    }

    async function deleteTrip(tripId, input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const trip = state.trips.find(row => row.id === tripId && row.status !== 'deleted');
        if (!trip) throw fail('Kørslen findes ikke eller er allerede slettet.', 404);
        const reason = text(input.note, 'Begrundelse for sletning', 1000);
        const previousStatus = trip.status;
        if (['approved', 'paid'].includes(previousStatus)) {
            markForReview(state, trip.employeeId, trip.date?.slice(0, 4), user, 'Km er fjernet efter godkendelsen. Kontrollér beregningen.', trip.id);
        }
        trip.status = 'deleted';
        trip.deletedPreviousStatus = previousStatus;
        trip.deletedAt = now();
        trip.deletedBy = username(user);
        history(trip, 'delete', user, reason);
        return persist(state, input.version, user);
    }

    async function correctLegacyDate(tripId, input, user) {
        requireManager(user);
        const state = await load();
        checkVersion(state, input.version);
        const trip = state.trips.find(row => row.id === tripId && row.status === 'legacy');
        if (!trip || trip.date) throw fail('Kun en manglende historisk dato kan suppleres.', 409);
        if (!validDate(input.date) || input.date > now().slice(0, 10)) throw fail('Ugyldig historisk dato.');
        trip.date = input.date;
        history(trip, 'legacy-date', user, 'Manglende dato suppleret efter import.');
        return persist(state, input.version, user);
    }

    return { data, syncEmployees, saveEmployee, saveRate, saveAddress, deleteAddress, savePurpose, deletePurpose,
        saveTrip, transition, document, previewImport, importData, deleteTrip, correctLegacyDate };
}

module.exports = { createMileageService, calculateAllowance, manager, fail, validDate, validCoordinates };
