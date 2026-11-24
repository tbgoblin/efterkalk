const { fail } = require('./mileageService');

function coordinates(value) {
    return Array.isArray(value) && value.length === 2 && value.every(n => typeof n === 'number' && Number.isFinite(n))
        && Math.abs(value[0]) <= 180 && Math.abs(value[1]) <= 90;
}

function createMileageRoutingService({ fetchImpl = global.fetch, apiKey = () => process.env.GOH_ORS_API_KEY || '' } = {}) {
    const configured = () => Boolean(apiKey().trim());
    async function request(url, options = {}) {
        if (!configured()) throw fail('Automatisk afstand er ikke konfigureret. Administratoren skal sætte GOH_ORS_API_KEY på serveren. Brug km manuelt indtil da.', 503);
        let response;
        try {
            response = await fetchImpl(url, { ...options, headers: { Authorization: apiKey().trim(), 'Content-Type': 'application/json' },
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
        const params = new URLSearchParams({ api_key: apiKey().trim(), text: query.trim(), size: '5' });
        const result = await request('https://api.openrouteservice.org/geocode/search?' + params);
        if (!Array.isArray(result?.features)) throw fail('Ruteservicen returnerede ingen adresseliste.', 502);
        return result.features.filter(row => coordinates(row?.geometry?.coordinates) && typeof row.properties?.label === 'string')
            .map(row => ({ address: row.properties.label, coordinates: row.geometry.coordinates }));
    }
    async function route(points, roundTrip) {
        if (!Array.isArray(points) || points.length < 2 || points.length > 12 || !points.every(coordinates)) throw fail('Vælg en gyldig adresse for hvert stop.');
        // Calculate the actual reverse itinerary too; road restrictions can make it differ from 2x outward distance.
        const itinerary = roundTrip === true ? points.concat(points.slice(0, -1).reverse()) : points;
        const result = await request('https://api.openrouteservice.org/v2/directions/driving-car', {
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
module.exports = { createMileageRoutingService };
