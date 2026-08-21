// ── QMS dataset ─────────────────────────────────────────────────────────────
// Estratto verbatim da routes/apiRoutes.js: lettura/scrittura/validazione del
// dataset QMS locale (data/qms-dataset.json). Le firme con fsRef sono
// mantenute identiche ai call site esistenti.
const path = require('path');

const QMS_DATASET_PATH = path.join(__dirname, '..', 'data', 'qms-dataset.json');

function defaultQmsDataset() {
    return {
        version: 1,
        updatedAt: new Date().toISOString(),
        folders: [
            {
                id: 'startside',
                name: 'Startside',
                description: 'Overordnet introduktion til kvalitetsledelsessystemet',
                documents: [
                    {
                        id: 'qfp-00',
                        title: 'QFP-00 Kvalitetsledelsessystemets startsidestartside',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-00%20Kvalitetsledelsessystemets%20startsidestartside.aspx',
                        content: 'Kvalitetsledelsessystemet bygger på ISO 9001-principper med procesorienteret tilgang. Procesflowet er opdelt i teknisk forberedelse, fabrikation/service samt fakturering og sagsafslutning.',
                        tags: ['ISO9001', 'procesflow'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'q-001',
                        title: 'Q-001 Kvalitetsledelsessystemet - Gantech håndbogen',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/Q-001%20Kvalitetsledelsessystemet%20-%20Gantech%20h%C3%A5ndbogen.aspx',
                        content: 'Forord og ramme for samspillet mellem kvalitetsledelsessystem, personalehåndbog, arbejdsmiljøportal og serviceportal.',
                        tags: ['forord', 'styring'],
                        updatedAt: new Date().toISOString()
                    }
                ]
            },
            {
                id: 'ledelse',
                name: 'Ledelse',
                description: 'Politikker, retningslinjer og administration',
                documents: [
                    {
                        id: 'qfp-01',
                        title: 'QFP-01 Politikker og retningslinjer',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-01%20Politikker%20og%20retningslinjer.aspx',
                        content: 'Overblik over certificeringer, godkendelser og overordnede politikker som virksomhedens styrende dokumentgrundlag.',
                        tags: ['politik', 'certificering'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-02',
                        title: 'QFP-02 Administration',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-02%20Administration.aspx',
                        content: 'Årlig revision af administrative systemer og rutiner for at sikre effektivt workflow og tydelig styring.',
                        tags: ['administration'],
                        updatedAt: new Date().toISOString()
                    }
                ]
            },
            {
                id: 'procesflow',
                name: 'Procesflow',
                description: 'Forespørgsel til produktion, levering og fakturering',
                documents: [
                    {
                        id: 'qfp-15',
                        title: 'QFP-15 Forespørgsel',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-15%20Foresp%C3%B8rgsel.aspx',
                        content: 'Proces for vurdering og håndtering af kundeforepørgsler med systematisk sagsbehandling.',
                        tags: ['forespørgsel'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-16',
                        title: 'QFP-16 Tilbud',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-16%20Tilbud.aspx',
                        content: 'Tilbudsproces med kravspecifikation og tekniske afklaringer.',
                        tags: ['tilbud'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-17',
                        title: 'QFP-17 Ordre eller kontrakt gennemgang',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-17%20Ordre%20eller%20kontrakt%20gennemgang.aspx',
                        content: 'Ordre-/kontraktgennemgang med formel validering af krav før igangsættelse.',
                        tags: ['ordre', 'kontrakt'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-18',
                        title: 'QFP-18 Produktions forberedelse',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-18%20Produktions%20forberedelse.aspx',
                        content: 'Planlægning, produktionsværktøjer og klargøring før produktion.',
                        tags: ['produktion'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-19',
                        title: 'QFP-19 Godkendelse og levering',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-19%20Godkendelse%20og%20levering.aspx',
                        content: 'Kontrol, godkendelse og levering efter aftalte specifikationer.',
                        tags: ['levering', 'kontrol'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-21',
                        title: 'QFP-21 Fakturering og opfølgning',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-21%20Fakturering%20og%20opf%C3%B8lgning.aspx',
                        content: 'Fakturering, sagsafslutning og opfølgning efter levering.',
                        tags: ['fakturering'],
                        updatedAt: new Date().toISOString()
                    },
                    {
                        id: 'qfp-22',
                        title: 'QFP-22 Produktion',
                        url: 'https://gantech.sharepoint.com/handbook/Sider/QFP-22%20Produktion.aspx',
                        content: 'Ordrestyring, produktionsforløb og overdragelse til levering.',
                        tags: ['produktion'],
                        updatedAt: new Date().toISOString()
                    }
                ]
            }
        ]
    };
}

function ensureQmsDatasetFile(fsRef) {
    const dataDir = path.dirname(QMS_DATASET_PATH);
    if (!fsRef.existsSync(dataDir)) {
        fsRef.mkdirSync(dataDir, { recursive: true });
    }
    if (!fsRef.existsSync(QMS_DATASET_PATH)) {
        const seed = defaultQmsDataset();
        fsRef.writeFileSync(QMS_DATASET_PATH, JSON.stringify(seed, null, 2), 'utf8');
    }
}

function readQmsDataset(fsRef) {
    ensureQmsDatasetFile(fsRef);
    const raw = fsRef.readFileSync(QMS_DATASET_PATH, 'utf8');
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.folders)) {
        throw new Error('QMS dataset format invalid');
    }
    return parsed;
}

function validateQmsDataset(payload) {
    if (!payload || typeof payload !== 'object') return 'Dataset mangler';
    if (!Array.isArray(payload.folders)) return 'folders skal være en liste';
    for (const folder of payload.folders) {
        if (!folder || typeof folder !== 'object') return 'Ugyldig mappe';
        if (!String(folder.id || '').trim()) return 'Mappe mangler id';
        if (!String(folder.name || '').trim()) return 'Mappe mangler navn';
        if (!Array.isArray(folder.documents)) return 'Mappe documents skal være en liste';
        for (const doc of folder.documents) {
            if (!doc || typeof doc !== 'object') return 'Ugyldigt dokument';
            if (!String(doc.id || '').trim()) return 'Dokument mangler id';
            if (!String(doc.title || '').trim()) return 'Dokument mangler titel';
            if (doc.tags && !Array.isArray(doc.tags)) return 'tags skal være en liste';
        }
    }
    return null;
}

function writeQmsDataset(fsRef, dataset) {
    ensureQmsDatasetFile(fsRef);
    const normalized = {
        version: Number(dataset.version || 1),
        updatedAt: new Date().toISOString(),
        folders: dataset.folders.map(folder => ({
            id: String(folder.id || '').trim(),
            name: String(folder.name || '').trim(),
            description: String(folder.description || '').trim(),
            documents: folder.documents.map(doc => ({
                id: String(doc.id || '').trim(),
                title: String(doc.title || '').trim(),
                url: String(doc.url || '').trim(),
                content: String(doc.content || '').trim(),
                tags: Array.isArray(doc.tags) ? doc.tags.map(t => String(t).trim()).filter(Boolean) : [],
                updatedAt: new Date().toISOString()
            }))
        }))
    };
    fsRef.writeFileSync(QMS_DATASET_PATH, JSON.stringify(normalized, null, 2), 'utf8');
    return normalized;
}

module.exports = {
    QMS_DATASET_PATH,
    defaultQmsDataset,
    ensureQmsDatasetFile,
    readQmsDataset,
    validateQmsDataset,
    writeQmsDataset
};
