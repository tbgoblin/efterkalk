const { fail } = require('./mileageService');

function coordinates(value) {
    return Array.isArray(value) && value.length === 2 && value.every(n => typeof n === 'number' && Number.isFinite(n))
        && Math.abs(value[0]) <= 180 && Math.abs(value[1]) <= 90;
}

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
        if (!response.ok) throw fail(response.status === 429 ? 'Ruteservicens forespørgselsgrænse er nået. Prøv senere eller angiv km manuelt.'
            : 'Ruteservicen afviste forespørgslen. Kontrollér adresser og serverens API-konfiguration.', 502);
        try { return await response.json(); }
        catch { throw fail('Ruteservicen returnerede et ugyldigt svar.', 502); }
    }
    async function search(query) {
        if (typeof query !== 'string' || query.trim().length < 3 || query.length > 500) throw fail('Skriv mindst tre tegn i adressen.');
        const key = await getKey();
        const params = new URLSearchParams({ api_key: key, text: query.trim(), size: '5' });
        const result = await request('https://api.openrouteservice.org/geocode/search?' + params, key);
        if (!Array.isArray(result?.features)) throw fail('Ruteservicen returnerede ingen adresseliste.', 502);
        return result.features.filter(row => coordinates(row?.geometry?.coordinates) && typeof row.properties?.label === 'string')
            .map(row => ({ address: row.properties.label, coordinates: row.geometry.coordinates }));
    }
    async function route(points, roundTrip) {
        if (!Array.isArray(points) || points.length < 2 || points.length > 12 || !points.every(coordinates)) throw fail('Vælg en gyldig adresse for hvert stop.');
        // Calculate the actual reverse itinerary too; road restrictions can make it differ from 2x outward distance.
        const itinerary = roundTrip === true ? points.concat(points.slice(0, -1).reverse()) : points;
        const result = await request('https://api.openrouteservice.org/v2/directions/driving-car', await getKey(), {
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
    return { configured, search, route };
}
module.exports = { createMileageRoutingService, createMileageRoutingSettings };
