(function (root, factory) {
    if (typeof module === 'object' && module.exports) module.exports = factory(require('./bilancio-formula'));
    else root.GohMonthlySummaryModel = factory(root.GohReportFormula);
}(typeof globalThis === 'undefined' ? this : globalThis, function (formula) {
    'use strict';
    const sources = {
        orders:'Antal fakturaordrer', invoiced:'Faktureret inkl. kreditnotaer', cost:'Kost (valgt kostvisning)',
        db:'Dækningsbidrag', dg:'Dækningsgrad', credits:'Kreditnotaer, beløb', credit_count:'Antal kreditnotaer',
        material:'Materiale, stang, indkøbte dele og underleverandør', material_only:'Materiale', stang:'Stang',
        time:'Tid', parts:'Indkøbte dele', subcontract:'Underleverandør',
        warehouse_delta:'Ændring i varelager', purchase_gap:'Modtaget indkøb, ikke fuldt faktureret', ub_orders:'Antal U/B-ordrer',
        ub_material:'Materiale i U/B-ordrer', finished_material:'Ændring i materiale i færdige, ikke fakturerede ordrer',
        via_material:'Ændring i materiale i igangværende ordrer (VIA)',
        purchase_timing:'Fakturaer bogført i en anden måned end varemodtagelsen'
    };
    const descriptions = {
        orders:'Antal fakturerede ordrer i måneden, inkl. kreditnotaer. U/B-ordrer tælles ikke med.',
        invoiced:'Fakturerede beløb i måneden, inkl. kreditnotaer og uden U/B-ordrer.',
        cost:'Ordrernes samlede kost efter den valgte kostvisning (kun registreret eller inkl. manglende tid fra stykliste).',
        db:'Faktureret minus kost.',
        dg:'Dækningsbidrag i procent af faktureret.',
        credits:'Summen af månedens kreditnotaer (negativt beløb).',
        credit_count:'Antal kreditnotaer i måneden.',
        material:'Materiale, stang, indkøbte dele og underleverandør fra efterkalkulationen på månedens fakturerede ordrer: forbrugt antal × kostpris (FIFO). Uden tid.',
        material_only:'Plademateriale fra efterkalkulationen på månedens fakturerede ordrer.',
        stang:'Stang fra efterkalkulationen på månedens fakturerede ordrer.',
        time:'Registreret tid på månedens fakturerede ordrer.',
        parts:'Indkøbte dele på månedens fakturerede ordrer.',
        subcontract:'Underleverandørarbejde på månedens fakturerede ordrer.',
        warehouse_delta:'Lagerværdi ved månedens udgang minus forrige måneds udgang, fra lagerliste-lukningerne: plader, stang, opfølgningsvarer, lagerkomponenter, restplader og diverse. Positiv = lageret er steget. Ordrer i arbejde og færdige ordrer er ikke med.',
        purchase_gap:'Indkøbsordrelinjer modtaget i måneden, hvor leverandørfakturaen endnu ikke er bogført (som det ser ud i dag). Underleverandørvarer (U…) er udeladt.',
        ub_orders:'Antal ordrer registreret som U/B (uden beregning) i måneden.',
        ub_material:'Materiale (uden tid) i månedens U/B-ordrer. Det er købt og bogført, men kommer aldrig med i det beregnede materiale.',
        finished_material:'Ændring i materiale (uden tid) i ordrer, der er færdige og afsendt, men ikke faktureret: månedens udgang minus forrige måneds. Månedens U/B-ordrer er udeladt.',
        via_material:'Ændring i materiale (uden tid) i ordrer, der stadig er i produktion (VIA): månedens udgang minus forrige måneds.',
        purchase_timing:'Bogført på 12070, 12075 og 12078 i måneden for varer modtaget i andre måneder, minus det, der er bogført i andre måneder (±3) for varer modtaget i måneden. Koblet via leverandørens fakturanummer; kreditnotaer følger den faktura, de henviser til. Positiv = der er bogført mere end månedens modtagelser.'
    };
    // Tooltip text for a row: what it is and how it is calculated.
    function describe(row, names = new Map()) {
        if (row.type === 'source') return descriptions[row.source] || '';
        if (row.type === 'accounts') return 'Bogførte beløb i måneden på de valgte konti'
            + (row.accounts && row.accounts.length ? ': ' + row.accounts.map(account => account.AcNo + ' ' + account.Nm).join(', ') : '') + '.';
        if (row.type === 'formula') return '= ' + String(row.formula || '').replace(/\[([A-Za-z][A-Za-z0-9_-]*)\]/g, (match, id) => names.has(id) ? '[' + names.get(id) + ']' : match);
        if (row.type === 'manual') return 'Beløb indtastet manuelt for måneden.';
        return '';
    }
    function defaults() {
        const source = (id, name, metric, display='card', unit='dkk') => ({ id,name,type:'source',source:metric,display,unit,section:'Vareforbrug vs. materiale' });
        const calculation = (id, name, expression) => ({ id,name,type:'formula',formula:expression,display:'row',unit:'dkk',section:'Vareforbrug vs. materiale' });
        return { schema:1, version:0, title:'Månedens overblik', rows:[
            source('orders','Ordrer','orders','card','count'), source('invoiced','Samlet faktureret','invoiced'),
            source('cost','Samlet kost','cost'), source('dg','Dækningsgrad','dg','card','percent'),
            {id:'booked',name:'Bogført (Vareforbrug)',type:'accounts',accounts:[12070],groups:[],exclude:[],children:false,sign:1,display:'row',unit:'dkk',section:'Vareforbrug vs. materiale'},
            source('material','Beregnet (materiale)','material','row'),
            calculation('difference','Forskel','[booked] - [material]'),
            {id:'explanation',name:'Forklaret af',type:'heading',display:'row',unit:'dkk',section:'Vareforbrug vs. materiale'},
            source('warehouse','Δ Varelager','warehouse_delta','row'),
            source('gap','Ikke fuldt faktureret','purchase_gap','row'),
            source('ub_material','Materiale i U/B-ordrer','ub_material','row'),
            source('finished_material','Δ Færdige, ikke fakturerede ordrer','finished_material','row'),
            source('timing','Faktureret i anden måned end modtaget','purchase_timing','row'),
            calculation('residual','Uforklaret rest','[difference] - [warehouse] + [gap] - [ub_material] - [finished_material] - [timing]')
        ] };
    }
    function invalid(message) { throw Object.assign(new Error(message), { statusCode:400 }); }
    function validate(input) {
        if (!input || input.schema !== 1 || !Number.isSafeInteger(input.version) || input.version < 0) invalid('Ugyldig version. Genindlæs opsætningen.');
        const text = (value, label, empty=false) => {
            if (typeof value !== 'string' || (!empty && !value.trim()) || value.length > 160) invalid(label + ': brug 1–160 tegn.');
            return value.trim();
        };
        if (!Array.isArray(input.rows) || !input.rows.length || input.rows.length > 60) invalid('Tilføj 1–60 rækker.');
        const prior = new Map();
        const rows = input.rows.map(row => {
            if (!row || !/^[A-Za-z][A-Za-z0-9_-]{0,63}$/.test(row.id) || prior.has(row.id)) invalid('Ugyldigt eller gentaget række-id.');
            const result = { id:row.id, name:text(row.name,'Navn'), type:row.type,
                display:row.display || 'row', unit:row.unit || 'dkk', section:text(row.section || '', 'Sektion',true) };
            if (!['card','row','hidden'].includes(result.display) || !['dkk','count','percent'].includes(result.unit)) invalid('Ugyldig visning eller enhed.');
            if (row.type === 'source') {
                if (!Object.hasOwn(sources,row.source)) invalid('Vælg en datakilde.');
                result.source=row.source;
            } else if (row.type === 'accounts') {
                const numbers = values => {
                    if (!Array.isArray(values) || values.length > 1000 || values.some(value => !Number.isSafeInteger(value) || value <= 0)) invalid('Kontonumre skal være positive heltal.');
                    return [...new Set(values)];
                };
                result.accounts=numbers(row.accounts || []); result.exclude=numbers(row.exclude || []);
                if (!Array.isArray(row.groups) || row.groups.length > 200) invalid('Ugyldige kontogrupper.');
                result.groups=[...new Set(row.groups.map(value => text(value,'Kontogruppe')))];
                if (!result.accounts.length && !result.groups.length) invalid(result.name + ': vælg mindst én konto eller gruppe.');
                if (![1,-1].includes(row.sign)) invalid('Vælg fortegn + eller −.');
                result.sign=row.sign; result.children=row.children === true;
            } else if (row.type === 'formula') {
                let parsed;
                try { parsed=formula.parseFormula(row.formula); } catch(error) { invalid(result.name + ': ' + error.message); }
                for (const ref of parsed.refs) if (!prior.has(ref) || prior.get(ref).type === 'heading') invalid(result.name + ': brug kun talrækker ovenfor i formlen.');
                result.formula=row.formula;
            } else if (row.type === 'manual') {
                if (!row.monthValues || typeof row.monthValues !== 'object' || Array.isArray(row.monthValues) || Object.keys(row.monthValues).length > 600) invalid('Ugyldige månedsbeløb.');
                result.monthValues={};
                for (const [month,value] of Object.entries(row.monthValues)) {
                    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month) || typeof value !== 'number' || !Number.isFinite(value)) invalid('Brug YYYY-MM og et gyldigt beløb.');
                    result.monthValues[month]=value;
                }
            } else if (row.type !== 'heading') invalid('Ukendt rækketype.');
            prior.set(row.id,result); return result;
        });
        return { schema:1,version:input.version,title:text(input.title,'Titel'),rows,scope:input.scope };
    }
    function evaluate(definition, { month, metrics={}, selections={}, accounts=[] } = {}) {
        const values=new Map();
        const data=new Map(accounts.map(row => [Number(row.AcNo),row]));
        return definition.rows.map(row => {
            let value=null, details=[];
            if (row.type === 'source') value=metrics[row.source] ?? null;
            else if (row.type === 'accounts') {
                const ids=selections[row.id];
                if (Array.isArray(ids)) {
                    details=ids.map(id => data.get(Number(id)));
                    if (details.every(account => account && account.Amount != null && Number.isFinite(Number(account.Amount)))) {
                        value=details.reduce((sum,account) => sum + Number(account.Amount)*row.sign,0);
                    }
                }
            } else if (row.type === 'formula') value=formula.evaluateFormula(formula.parseFormula(row.formula).tree,id => values.get(id) ?? null);
            else if (row.type === 'manual') value=row.monthValues[month] ?? null;
            if (value != null && !Number.isFinite(value)) value=null;
            values.set(row.id,value);
            return { ...row,value,accounts:details.filter(Boolean) };
        });
    }
    // Sign each row carries in the formula that uses it ("[a] - [b] + [c]"). Formulas with other
    // operators, numbers or parentheses give no signs, since a single +/− would be misleading there.
    function formulaSigns(rows) {
        const signs=new Map();
        for (const row of rows) {
            const text=String(row.type==='formula' && row.formula || '').trim();
            if (!text || !/^[+-]?\s*\[[A-Za-z][A-Za-z0-9_-]*\](\s*[+-]\s*\[[A-Za-z][A-Za-z0-9_-]*\])*$/.test(text)) continue;
            for (const match of text.matchAll(/([+-]?)\s*\[([A-Za-z][A-Za-z0-9_-]*)\]/g)) signs.set(match[2],match[1]==='-'?'−':'+');
        }
        return signs;
    }
    return { sources,descriptions,describe,defaults,validate,evaluate,formulaSigns };
}));
