const crypto = require('node:crypto');
const { validatePeriod } = require('./bilancioService');

function failure(message, statusCode = 409) {
    const error = new Error(message);
    error.statusCode = statusCode;
    throw error;
}
function createBilancioArchiveService({ gohData, definitionStore, bilancioService, generatePdf, appVersion, now = Date.now }) {
    const drafts = new Map();
    const scope = () => definitionStore.scope();
    const prefix = reportId => {
        if (typeof reportId !== 'string' || !/^[a-zA-Z][a-zA-Z0-9_-]{0,39}$/.test(reportId)) failure('Ugyldigt rapport-id.', 400);
        return 'bilancio_snapshot_' + crypto.createHash('sha256').update(JSON.stringify([scope(), reportId])).digest('hex').slice(0, 32) + '_';
    };
    async function keys(reportId) {
        const result = await gohData.getAppStateKeysByPrefix(prefix(reportId));
        if (!Array.isArray(result)) failure('Rapportarkivet kunne ikke hentes fra GOH.', 503);
        return result.map(row => row.key).sort().reverse();
    }
    const monthPrefix = (reportId, year, period) => prefix(reportId) + year + '-' + String(period).padStart(2, '0') + '_';
    async function latestRevision(reportId, year, period) {
        const month = monthPrefix(reportId, year, period);
        const entries = (await keys(reportId)).filter(key => key.startsWith(month));
        return entries.length ? Number(entries[0].slice(month.length)) : 0;
    }
    async function load(reportId, id) {
        const currentScope = scope();
        if (typeof id !== 'string' || !id.startsWith(prefix(reportId)) || !/^bilancio_snapshot_[a-f0-9]{32}_\d{4}-\d{2}_\d{6}$/.test(id)) failure('Ugyldig arkivreference.', 400);
        const result = await gohData.getAppState(id, { strict: true });
        if (!result) failure('Rapporten findes ikke i arkivet.', 404);
        if (scope() !== currentScope || result.payload.scope !== currentScope) failure('Databaseprofilen er ændret. Hent arkivet igen.');
        return structuredClone(result.payload);
    }
    function summary(saved) {
        return {
            id: saved.id, reportId: saved.report.reportId, reportName: saved.report.reportName,
            year: saved.report.year, period: saved.report.period, revision: saved.revision,
            configurationRevision: saved.configurationRevision, savedAt: saved.savedAt,
            savedBy: saved.savedBy, note: saved.note, appVersion: saved.appVersion, unit: saved.unit
        };
    }
    return {
        load,
        async list(reportId) {
            const currentScope = scope();
            const result = [];
            for (const id of await keys(reportId)) result.push(summary(await load(reportId, id)));
            if (scope() !== currentScope) failure('Databaseprofilen er ændret. Hent arkivet igen.');
            return result;
        },
        async draft(year, period, reportId, username) {
            validatePeriod(year, period);
            prefix(reportId);
            for (const [id, draft] of drafts) if (now() - draft.createdAt > 2 * 60 * 60 * 1000) drafts.delete(id);
            if (drafts.size >= 64) failure('For mange åbne kladder. Prøv igen senere.', 503);
            const currentScope = scope();
            const baseRevision = await latestRevision(reportId, year, period);
            const templateKey = monthPrefix(reportId, year, period).replace('bilancio_snapshot_', 'bilancio_template_').slice(0, -1);
            let template = (await gohData.getAppState(templateKey, { strict: true }))?.payload;
            let bundle = await bilancioService.bundle(year, period, reportId, template || null);
            if (!template) {
                template = {
                    schema: 1, scope: currentScope, definition: bundle.definition,
                    configurationRevision: bundle.configurationRevision, lockedAt: new Date(now()).toISOString(), lockedBy: username
                };
                if (scope() !== currentScope) failure('Databaseprofilen er ændret. Hent rapporten igen.');
                if (!await gohData.setAppState(templateKey, template, { createOnly: true })) {
                    const concurrent = await gohData.getAppState(templateKey, { strict: true });
                    if (!concurrent) failure('Månedens rapportskabelon kunne ikke låses på GOH.', 503);
                    template = concurrent.payload;
                    bundle = await bilancioService.bundle(year, period, reportId, template);
                }
            }
            if (scope() !== currentScope || bundle.scope !== currentScope) failure('Databaseprofilen er ændret. Hent rapporten igen.');
            bundle.templateLockedAt = template.lockedAt;
            const draftId = crypto.randomUUID();
            drafts.set(draftId, { bundle: structuredClone(bundle), username, createdAt: now(), baseRevision, savedId: null });
            return { ...bundle, draftId, baseRevision, appVersion, expiresAt: new Date(now() + 2 * 60 * 60 * 1000).toISOString() };
        },
        async save({ draftId, unit, note }, username) {
            if (![1, 1000].includes(unit)) failure('Vælg DKK eller t.kr.', 400);
            if (typeof note !== 'string' || note.length > 1000) failure('Noten må højst være 1000 tegn.', 400);
            const draft = drafts.get(draftId);
            if (!draft || now() - draft.createdAt > 2 * 60 * 60 * 1000) failure('Kladden er udløbet. Hent rapporten igen.', 410);
            if (draft.username !== username) failure('Kladden tilhører en anden bruger.', 403);
            if (draft.bundle.scope !== scope()) failure('Databaseprofilen er ændret. Hent rapporten igen.');
            const { report } = draft.bundle;
            if (draft.savedId) return load(report.reportId, draft.savedId);
            const latest = await latestRevision(report.reportId, report.year, report.period);
            if (latest !== draft.baseRevision) failure('En ny revision er allerede gemt. Hent rapporten igen før du opretter en rettelse.');
            if (latest && !note.trim()) failure('Skriv en begrundelse for den nye rapportrevision.', 400);
            if (latest >= 999999) failure('Maksimalt antal revisioner er nået.');
            const id = monthPrefix(report.reportId, report.year, report.period) + String(latest + 1).padStart(6, '0');
            const saved = {
                schema: 1, ...structuredClone(draft.bundle), id, revision: latest + 1, unit,
                note: note.trim(), savedBy: username, savedAt: new Date(now()).toISOString(), appVersion
            };
            const pdf = await generatePdf(saved);
            saved.pdfBase64 = pdf.toString('base64');
            saved.pdfSha256 = crypto.createHash('sha256').update(pdf).digest('hex');
            if (saved.scope !== scope()) failure('Databaseprofilen er ændret. Rapporten blev ikke gemt.');
            if (!await gohData.setAppState(id, saved, { createOnly: true })) {
                const concurrent = await gohData.getAppState(id, { strict: true });
                if (!concurrent) failure('Rapporten kunne ikke gemmes på GOH. Prøv igen.', 503);
                failure('En anden bruger har gemt revisionen. Hent rapporten igen.');
            }
            draft.savedId = id;
            return structuredClone(saved);
        }
    };
}
module.exports = { createBilancioArchiveService };
