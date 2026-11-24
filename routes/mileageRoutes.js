const express = require('express');
const { createMileageService, fail } = require('../services/mileageService');
const { createMileageRoutingService } = require('../services/mileageRoutingService');
const { parseMileageImport } = require('../services/mileageImportService');
const { generateMileagePdf } = require('../services/mileagePdf');

function createMileageRouter({ gohData, readUsers, getSessionUser, requireModulePermission, logEvent }) {
    const router = express.Router();
    const service = createMileageService({ gohData, readUsers });
    const routing = createMileageRoutingService();
    const limits = new Map();
    router.use(requireModulePermission('mileage'));
    router.use((_req, res, next) => { res.setHeader('Cache-Control', 'no-store'); next(); });
    const run = action => async (req, res) => {
        try {
            if (req.method === 'POST' && (!req.body || typeof req.body !== 'object' || Array.isArray(req.body))) throw fail('En gyldig JSON-forespørgsel kræves.');
            await action(req, res, getSessionUser(req));
        }
        catch (error) {
            // Do not log addresses, employee records, request bodies or provider credentials.
            logEvent('ERROR mileage: HTTP ' + (error.statusCode || 500));
            res.status(error.statusCode || 500).json({ error: error.statusCode ? error.message : 'Kørselsmodulet kunne ikke behandle forespørgslen.' });
        }
    };
    const management = requireModulePermission('mileageManage');
    const throttle = (req, _res, next) => {
        const key = String(getSessionUser(req).username).toLowerCase();
        const now = Date.now();
        for (const [user, value] of limits) if (now - value.start >= 60000) limits.delete(user);
        const entry = limits.get(key) || { start: now, count: 0 };
        limits.set(key, entry);
        if (++entry.count > 30) return next(fail('For mange adresse- eller rutekald. Vent et minut.', 429));
        next();
    };
    router.get('/data', run(async (_req, res, user) => res.json({ ok: true, ...await service.data(user), routingConfigured: routing.configured() })));
    router.post('/employees', management, run(async (req, res, user) => { await service.saveEmployee(req.body, user); res.json({ ok: true }); }));
    router.post('/rates', management, run(async (req, res, user) => { await service.saveRate(req.body, user); res.json({ ok: true }); }));
    router.post('/trips', run(async (req, res, user) => { await service.saveTrip(req.body, user); res.json({ ok: true }); }));
    router.post('/trips/:id/action', run(async (req, res, user) => { await service.transition(req.params.id, req.body, user); res.json({ ok: true }); }));
    router.post('/trips/:id/legacy-date', management, run(async (req, res, user) => { await service.correctLegacyDate(req.params.id, req.body, user); res.json({ ok: true }); }));
    router.post('/import', management, express.json({ limit: '5mb' }), run(async (req, res, user) => {
        if (req.body.preview === true) {
            const parsed = parseMileageImport(req.body.files);
            return res.json({ ok: true, counts: { employees: parsed.employees.length, trips: parsed.trips.length,
                addresses: parsed.addresses.length, purposes: parsed.purposes.length }, warnings: parsed.warnings });
        }
        res.json({ ok: true, ...await service.importData(req.body, user) });
    }));
    router.post('/addresses/search', throttle, run(async (req, res) => res.json({ ok: true, addresses: await routing.search(req.body.query) })));
    router.post('/route', throttle, run(async (req, res) => res.json({ ok: true, ...await routing.route(req.body.points, req.body.roundTrip) })));
    router.get('/trips/:id/document.pdf', run(async (req, res, user) => {
        const trip = await service.document(req.params.id, user);
        const pdf = await generateMileagePdf([trip]);
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename="Gantech-korsel.pdf"');
        res.send(pdf);
    }));
    router.get('/report.pdf', run(async (req, res, user) => {
        if (!/^\d{4}$/.test(String(req.query.year || '')) || typeof req.query.employeeId !== 'string') throw fail('Vælg år og medarbejder.');
        const data = await service.data(user);
        const trips = data.trips.filter(row => row.employeeId === req.query.employeeId && row.date?.slice(0, 4) === req.query.year);
        if (!trips.length) throw fail('Ingen kørsler i det valgte år.', 404);
        const pdf = await generateMileagePdf(trips.sort((a, b) => a.date.localeCompare(b.date)));
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename="Gantech-korsel-aarsbilag.pdf"');
        res.send(pdf);
    }));
    router.use((error, _req, res, _next) => {
        logEvent('ERROR mileage middleware: HTTP ' + (error.statusCode || error.status || 500));
        res.status(error.statusCode || error.status || 500).json({ error: error.statusCode ? error.message : 'Ugyldig eller for stor forespørgsel.' });
    });
    return router;
}
module.exports = { createMileageRouter };
