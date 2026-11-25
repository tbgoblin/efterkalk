const { fail, validCoordinates } = require('./mileageService');
const proj4 = require('proj4');

function createMileageRoutingSettings({ gohData, envKey = () => process.env.GOH_ORS_API_KEY || '' }) {
    const stateKey = 'mileage_routing_settings_v1';
    async function read() {
        const saved = await gohData.getAppState(stateKey, { strict: true });
        const state = saved?.payload || { version: 0, apiKey: '' };
        if (!Number.isSafeInteger(state.version) || state.version < 0 || typeof state.apiKey !== 'string') {
            throw fail('Ruteindstillingerne er ugyldige. Kontakt administratoren.', 503);
        }
        return state;
    }
    async function status() {
        const state = await read();
        const source = state.apiKey ? 'database' : envKey().trim() ? 'environment' : 'none';
        return { version: state.version, configured: source !== 'none', source };
    }
    async function key() {
        const state = await read();
        return state.apiKey || envKey().trim();
    }
    async function save(input, user) {
        if (!input || !Number.isSafeInteger(input.version) || input.version < 0 || input.version >= 2147483647
            || !['save', 'remove'].includes(input.action)) throw fail('Ugyldig forespørgsel.');
        const apiKey = input.action === 'remove' ? '' : typeof input.apiKey === 'string' ? input.apiKey.trim() : '';
        if (input.action === 'save' && (!apiKey || apiKey.length > 512 || !/^[A-Za-z0-9._~+\/=-]+$/.test(apiKey))) {
            throw fail('Angiv en gyldig API-nøgle.');
        }
        const current = await read();
        if (current.version !== input.version) throw fail('Indstillingerne er ændret. Genindlæs og prøv igen.', 409);
        const state = { version: current.version + 1, apiKey, updatedAt: new Date().toISOString(), updatedBy: String(user.username || '') };
        if (!await gohData.setAppState(stateKey, state, { expectedVersion: current.version })) {
            throw fail('Indstillingerne kunne ikke gemmes i GOH. Genindlæs og prøv igen.', 409);
        }
        return status();
    }
    return { status, key, save };
}

function createMileageRoutingService({ fetchImpl = global.fetch, apiKey = () => process.env.GOH_ORS_API_KEY || '' } = {}) {
    const apiBase = 'https://api.heigit.org';
    async function getKey() {
        const key = (await apiKey()).trim();
        if (!key) throw fail('Automatisk afstand er ikke konfigureret. Kontakt en administrator, eller angiv km manuelt.', 503);
        return key;
    }
    const configured = async () => Boolean((await apiKey()).trim());
    async function request(url, key, options = {}) {
        let response;
        try {
            response = await fetchImpl(url, { ...options, headers: { Authorization: key, 'Content-Type': 'application/json' },
                signal: AbortSignal.timeout(15000), redirect: 'error' });
        } catch {
            throw fail('Ruteservicen kunne ikke kontaktes. Prøv igen eller angiv km manuelt.', 502);
        }
        if (!response.ok) {
            if (response.status === 429) throw fail('Ruteservicens forespørgselsgrænse er nået. Prøv senere eller angiv km manuelt.', 502);
            if (response.status === 401 || response.status === 403) {
                throw fail('Ruteservicen afviste API-nøglen eller adgangen. Bed superadministratoren kontrollere eller udskifte nøglen.', 502);
            }
            if (response.status === 400) {
                throw fail('Ruteservicen afviste forespørgslen. Kontrollér indtastningen og prøv igen, eller angiv km manuelt.', 502);
            }
            throw fail(`Ruteservicen returnerede HTTP ${response.status}. Prøv igen eller angiv km manuelt.`, 502);
        }
        try { return await response.json(); }
        catch { throw fail('Ruteservicen returnerede et ugyldigt svar.', 502); }
    }
    async function search(query) {
        if (typeof query !== 'string' || query.trim().length < 3 || query.length > 500) throw fail('Skriv mindst tre tegn i adressen.');
        const key = await getKey();
        const params = new URLSearchParams({ api_key: key, text: query.trim(), size: '10' });
        const result = await request(apiBase + '/pelias/v1/search?' + params, key);
        if (!Array.isArray(result?.features)) throw fail('Ruteservicen returnerede ingen adresseliste.', 502);
        return result.features.filter(row => validCoordinates(row?.geometry?.coordinates) && typeof row.properties?.label === 'string')
            .map(row => {
                const properties = row.properties;
                const label = properties.label.trim();
                const street = typeof properties.street === 'string' ? properties.street.trim() : '';
                const number = typeof properties.housenumber === 'string' ? properties.housenumber.trim() : '';
                const streetAddress = [street, number].filter(Boolean).join(' ');
                const labelKey = label.toLocaleLowerCase();
                const alreadyShown = streetAddress && (labelKey.includes(streetAddress.toLocaleLowerCase())
                    || (street && number && labelKey.includes((number + ' ' + street).toLocaleLowerCase())));
                const name = typeof properties.name === 'string' ? properties.name.trim() : '';
                return { name, address: streetAddress && !alreadyShown ? label + ' · ' + streetAddress : label,
                    streetAddressKnown: Boolean(street && number), coordinates: row.geometry.coordinates };
            }).filter(row => row.address);
    }
    async function searchDanish(query) {
        if (typeof query !== 'string' || query.trim().length < 3) throw fail('Skriv mindst tre tegn i adressen.');
        if (query.trim().length > 73) throw fail('Den danske adressesøgning kan højst modtage 73 tegn. Forkort adressen og prøv igen.');
        await getKey();
        const base = 'https://adressevaelger.dk';
        const token = 'adressevaelger123';
        async function danishRequest(url) {
            let response;
            try {
                response = await fetchImpl(url, { signal: AbortSignal.timeout(15000), redirect: 'error' });
            } catch {
                throw fail('Den danske adressetjeneste kunne ikke kontaktes. Prøv igen eller angiv km manuelt.', 502);
            }
            if (!response.ok) throw fail(`Den danske adressetjeneste returnerede HTTP ${response.status}. Prøv igen eller angiv km manuelt.`, 502);
            let result;
            try { result = await response.json(); }
            catch { throw fail('Den danske adressetjeneste returnerede et ugyldigt svar.', 502); }
            if (result?.status !== 'ok') throw fail('Den danske adressetjeneste kunne ikke fuldføre søgningen. Prøv igen eller angiv km manuelt.', 502);
            return result;
        }
        const params = new URLSearchParams({ tekst: query.trim(), token, maksimum: '10' });
        const result = await danishRequest(base + '/husnumre/soeg?' + params);
        if (!Array.isArray(result.fund)) throw fail('Den danske adressetjeneste returnerede ingen adresseliste.', 502);
        const matches = result.fund.filter(row => row?.type === 'husnummer').slice(0, 10);
        return Promise.all(matches.map(async row => {
            if (typeof row.id !== 'string' || !/^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(row.id)) {
                throw fail('Den danske adressetjeneste returnerede et ugyldigt adresse-id.', 502);
            }
            const details = await danishRequest(base + '/husnumre/' + row.id + '?' + new URLSearchParams({ token }));
            const house = details.husnummer;
            const geometry = house?.adgangspunkt?.geometri;
            const coordinates = geometry?.coordinates;
            if (geometry?.crs?.properties?.name !== 'EPSG:25832' || !Array.isArray(coordinates) || coordinates.length !== 2
                || !coordinates.every(Number.isFinite) || typeof house?.adgangsadressebetegnelse !== 'string'
                || !house.adgangsadressebetegnelse.trim() || typeof house.vejnavn !== 'string' || !house.vejnavn.trim()
                || typeof house.husnummertekst !== 'string' || !house.husnummertekst.trim()) {
                throw fail('Den danske adressetjeneste returnerede en adresse uden gyldige koordinater eller adresseoplysninger.', 502);
            }
            const converted = proj4('+proj=utm +zone=32 +ellps=GRS80 +units=m +no_defs', 'WGS84', coordinates);
            if (!validCoordinates(converted)) throw fail('Den danske adressetjeneste returnerede ugyldige koordinater.', 502);
            return { name: '', address: house.adgangsadressebetegnelse.trim(), streetAddressKnown: true, coordinates: converted };
        }));
    }
    async function searchCombined(query) {
        if (typeof query !== 'string' || query.trim().length < 3 || query.length > 500) throw fail('Skriv mindst tre tegn i adressen.');
        await getKey();
        const [danish, general] = await Promise.allSettled([searchDanish(query), search(query)]);
        const failures = [danish, general].filter(result => result.status === 'rejected');
        const warnings = failures.map(result => result.reason?.statusCode
            ? result.reason.message : 'En adressetjeneste kunne ikke fuldføre søgningen. Prøv igen.');
        if (failures.length === 2) throw fail(warnings.join(' '), 502);
        const official = danish.status === 'fulfilled' ? danish.value : [];
        const places = general.status === 'fulfilled' ? general.value : [];
        const addresses = [];
        for (let i = 0; i < Math.max(official.length, places.length); i++) {
            if (official[i]) addresses.push({ ...official[i], source: 'danish' });
            if (places[i]) addresses.push({ ...places[i], source: 'general' });
        }
        return { addresses, warnings };
    }
    async function route(points, roundTrip) {
        if (!Array.isArray(points) || points.length < 2 || points.length > 12 || !points.every(validCoordinates)) throw fail('Vælg en gyldig adresse for hvert stop.');
        // Calculate the actual reverse itinerary too; road restrictions can make it differ from 2x outward distance.
        const itinerary = roundTrip === true ? points.concat(points.slice(0, -1).reverse()) : points;
        const result = await request(apiBase + '/openrouteservice/v2/directions/driving-car', await getKey(), {
            // ORS omits segment distances when instructions are disabled.
            method: 'POST', body: JSON.stringify({ coordinates: itinerary, instructions: true, geometry: false })
        });
        const segments = result?.routes?.[0]?.segments;
        if (!Array.isArray(segments) || segments.length !== itinerary.length - 1
            || segments.some(row => !Number.isFinite(row?.distance) || row.distance < 0)) throw fail('Ruteservicen returnerede ikke afstand for alle delstrækninger.', 502);
        const km = Math.round(segments.reduce((sum, row) => sum + row.distance, 0) / 10) / 100;
        if (km <= 0 || km > 10000) throw fail('Den beregnede afstand ligger uden for det tilladte interval.', 502);
        return { km, legs: segments.map(row => Math.round(row.distance / 10) / 100), provider: 'openrouteservice',
            attribution: '© openrouteservice.org by HeiGIT | Map data © OpenStreetMap contributors' };
    }
    return { configured, search, searchDanish, searchCombined, route };
}
module.exports = { createMileageRoutingService, createMileageRoutingSettings };
