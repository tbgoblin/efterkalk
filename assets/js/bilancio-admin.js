(() => {
    const el = id => document.getElementById(id);
    const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    let definition, catalog, selected = 0, dirty = false, busy = false;
    const labels = { accounts: 'Konti / grupper', sum: 'Sum af valgte rækker', percent: 'Procent af valgte rækker', formula: 'Formel', lager: 'Lagerliste', manual: 'Manuelt beløb pr. måned', heading: 'Overskrift' };
    const metrics = { sales: 'Færdige SO salgpris (hele ordren)', cost: 'Færdige SO kostpris', margin: 'SO salgpris − SO kostpris', warehouse: 'Varelager (inkl. rest, standard pladepris)', via: 'Vare i arbejde / VIA' };
    const periods = { selected: 'Måneden valgt i rapporten', previous: 'Måneden før rapportperioden', current: 'Aktuel kalendermåned', fixed: 'Fast måned' };
    function extraEditor(row) {
        if (row.type === 'formula') {
            const chip = r => `<button type="button" data-insert="${r.id}">${esc(r.name)} [${r.id}]</button>`;
            const localChips = rows().slice(0,selected).filter(r=>r.type!=='heading').map(chip).join('');
            const crossChips = el('section').value==='balance' ? definition.pnl.filter(r=>r.type!=='heading').map(chip).join('') : '';
            return `<label>Formel<textarea id="formula" rows="4" maxlength="2000">${esc(row.formula || '')}</textarea></label><div class="chips">${localChips}${crossChips?`<span class="muted" style="width:100%">Fra resultatopgørelsen (år til dato):</span>${crossChips}`:''}</div>`;
        }
        if (!['lager','manual'].includes(row.type)) return '';
        return (row.type === 'lager' ? `<label>Værdi<select id="metric">${Object.entries(metrics).map(([v,l])=>`<option value="${v}" ${row.metric===v?'selected':''}>${l}</option>`).join('')}</select></label><label class="check"><input id="currentSales" type="checkbox" ${row.currentSales?'checked':''}>Hent manglende salgspriser fra aktuelle Visma-ordrer</label>` : `<label>Beløb pr. måned (DKK, én linje pr. måned)<textarea id="monthValues" rows="6" placeholder="2026-08 = 1500.50">${esc(Object.entries(row.monthValues||{}).map(([m,v])=>m+' = '+v).join('\n'))}</textarea></label>`) + `<label>Dataperiode<select id="sourcePeriod">${Object.entries(periods).map(([v,l])=>`<option value="${v}" ${(row.period||'selected')===v?'selected':''}>${l}</option>`).join('')}</select></label><label>Fast måned<input id="sourceMonth" type="month" value="${esc(row.month||'')}" ${row.period==='fixed'?'':'disabled'}></label>`;
    }
    const rows = () => definition[el('section').value];
    const current = () => rows()[selected];
    const uid = prefix => prefix + Date.now().toString(36) + Math.random().toString(36).slice(2,8);
    function askText(title, defaultValue = '') {
        return new Promise(resolve => {
            const dialog = el('textPromptDialog'), input = el('textPromptInput');
            el('textPromptTitle').textContent = title; input.value = defaultValue; dialog.returnValue = '';
            const onClose = () => { dialog.removeEventListener('close', onClose); resolve(dialog.returnValue === 'ok' ? input.value : null); };
            dialog.addEventListener('close', onClose); dialog.showModal(); input.focus(); input.select();
        });
    }
    el('textPromptCancel').onclick = () => el('textPromptDialog').close('cancel');
    let reportList = [];
    function drawReports() {
        const id=definition.reportId||'default';
        const list=reportList.filter(r=>r.reportId!==id).concat([{reportId:id,reportName:definition.reportName||'Økonomirapport'}]);
        el('reportSelect').innerHTML=list.map(r=>`<option value="${esc(r.reportId)}">${esc(r.reportName)}</option>`).join('');el('reportSelect').value=id;
        el('reportName').value=definition.reportName||'Økonomirapport';el('showPnl').checked=definition.showPnl!==false;
        document.querySelector('a[href^="/assets/bilancio-trial.html"]').href='/assets/bilancio-trial.html?reportId='+encodeURIComponent(id);
    }
    const message = (value, error = false) => { el('message').textContent = value; el('message').className = error ? 'error' : ''; };
    function changed() { dirty = true; el('preview').innerHTML = ''; el('revision').textContent = `Version ${definition.version} · Ikke gemte ændringer`; }
    async function api(url, options) {
        const response = await fetch(url, { cache: 'no-store', ...options });
        const data = await response.json();
        if (!response.ok) throw Error(response.status === 401 || response.status === 403 ? 'Log ind som superadmin i Operations Hub.' : data.error || 'Handlingen mislykkedes.');
        return data;
    }
    function lock(value) {
        busy = value;
        el('workspace').inert = value;
        el('save').disabled = value || !definition;
        el('reload').disabled = value;
        el('showPreview').disabled = value || !definition;
    }
    function drawRoles() {
        const options = definition.pnl.filter(row => row.type !== 'heading').map(row => `<option value="${row.id}">${esc(row.name)}</option>`).join('');
        el('revenue').innerHTML = options; el('revenue').value = definition.revenueRow;
        el('beforeTax').innerHTML = options; el('beforeTax').value = definition.beforeTaxRow;
        const previous=el('assetTotal').value;
        el('assetTotal').innerHTML='<option value="">Opret total af eksisterende balancekonti</option>'+definition.balance.filter(r=>r.type!=='heading').map(r=>`<option value="${r.id}">${esc(r.name)}</option>`).join('');
        el('assetTotal').value=previous;
    }
    function drawRows() {
        el('rows').innerHTML = rows().map((row, index) => row.visibility==='hidden' && !el('showHelpers').checked && index!==selected ? '' : `<div class="row ${row.type==='heading'?'is-heading':''}"><button type="button" data-index="${index}" class="${index===selected?'selected':''}">${esc(row.name)}<small>${row.visibility==='hidden'?'Skjult hjælperække':labels[row.type]}</small></button></div>`).join('');
        el('rows').querySelectorAll('[data-index]').forEach(button => button.onclick = () => { selected = Number(button.dataset.index); draw(); });
        el('up').disabled = selected === 0; el('down').disabled = selected === rows().length - 1;
    }
    function drawEditor() {
        const row = current();
        if (!row) { el('editor').innerHTML = '<p>Tryk + Række for at begynde.</p>'; el('catalogPanel').hidden=true;return; }
        el('editor').innerHTML = `<label>Navn i rapporten<input id="rowName" maxlength="160" value="${esc(row.name)}"></label><label>Type<select id="rowType">${Object.entries(labels).map(([v,l])=>`<option value="${v}" ${row.type===v?'selected':''}>${l}</option>`).join('')}</select></label>` +
            (row.type === 'accounts' ? `<label>Fortegn<select id="sign"><option value="1" ${row.sign===1?'selected':''}>Behold kontofortegn (typisk balance)</option><option value="-1" ${row.sign===-1?'selected':''}>Vend fortegn (typisk resultatopgørelse)</option></select></label><label class="check"><input id="children" type="checkbox" ${row.children?'checked':''}>Medtag undergrupper</label><div>Konti<div class="chips">${row.accounts.map(ac=>`<button class="chip" data-remove="accounts" data-value="${ac}">${ac} ×</button>`).join('')||'<span class="muted">Ingen direkte konti</span>'}</div></div><div>Kontogrupper<div class="chips">${row.groups.map((g,i)=>`<button class="chip" data-remove="groups" data-value="${i}">${esc(g)} ×</button>`).join('')||'<span class="muted">Ingen grupper</span>'}</div></div><div>Udelad konti<div class="chips">${row.exclude.map(ac=>`<button class="chip" data-remove="exclude" data-value="${ac}">${ac} ×</button>`).join('')||'<span class="muted">Ingen undtagelser</span>'}</div></div>` :
            ['sum','percent'].includes(row.type) ? `<div>Medregn disse rækker<div class="source-list">${rows().slice(0,selected).filter(r=>r.type!=='heading').map(r=>`<label><input type="checkbox" data-source="${r.id}" ${row.sources.includes(r.id)?'checked':''}>${esc(r.name)}</label>`).join('')||'<p class="muted">Flyt rækken under de rækker, du vil summere.</p>'}</div></div>${row.type==='percent'?`<label>Procent (f.eks. −22 til skat)<input id="rate" type="number" step="any" min="-1000" max="1000" value="${row.rate}"></label>`:''}` : (row.type === 'heading' ? '<p class="muted">En overskrift har intet beløb og ændrer ikke totalerne.</p>' : '')) + '<button id="deleteRow" type="button">Fjern række</button>';
        el('deleteRow').insertAdjacentHTML('beforebegin', extraEditor(row) + `<p class="muted">Formelreference: <code>[${row.id}]</code></p><label>Vis i rapporten<select id="visibility">${Object.entries({both:'Begge tabeller',period:'Kun periode',ytd:'Kun år til dato',hidden:'Skjult hjælperække'}).map(([v,l])=>`<option value="${v}" ${(row.visibility||'both')===v?'selected':''}>${l}</option>`).join('')}</select></label><label class="check"><input id="bold" type="checkbox" ${row.bold?'checked':''}>Fed skrift</label>`);
        el('visibility').onchange = e => { row.visibility=e.target.value;changed(); };
        el('bold').onchange = e => { row.bold=e.target.checked;changed(); };
        el('deleteRow').insertAdjacentHTML('beforebegin', `<label class="check"><input id="checkZero" type="checkbox" ${row.checkZero?'checked':''}>Nul-kontrol (OK ved 0,00 kr)</label>`);
        el('checkZero').onchange=e=>{row.checkZero=e.target.checked;changed();};
        if (el('formula')) el('formula').oninput = e => { row.formula=e.target.value;changed(); };
        el('editor').querySelectorAll('[data-insert]').forEach(b=>b.onclick=()=>{
            const input=el('formula');input.setRangeText('['+b.dataset.insert+']',input.selectionStart,input.selectionEnd,'end');row.formula=input.value;changed();input.focus();
        });
        if (el('metric')) el('metric').onchange=e=>{row.metric=e.target.value;changed();};
        if (el('currentSales')) el('currentSales').onchange=e=>{row.currentSales=e.target.checked;changed();};
        if (el('sourcePeriod')) el('sourcePeriod').onchange=e=>{row.period=e.target.value;changed();el('sourceMonth').disabled=row.period!=='fixed';el('sourceMonth').closest('label').hidden=row.period!=='fixed';};
        if (el('sourceMonth')) el('sourceMonth').onchange=e=>{row.month=e.target.value;changed();};
        if (el('monthValues')) el('monthValues').oninput=e=>{
            const values={};let valid=true;
            for(const line of e.target.value.split('\n').filter(l=>l.trim())) {
                const match=line.trim().match(/^(20\d{2}-(?:0[1-9]|1[0-2]))\s*=\s*(-?\d+(?:[.,]\d+)?)$/);
                if(!match || Object.hasOwn(values,match[1])) {valid=false;break;}
                values[match[1]]=Number(match[2].replace(',','.'));
            }
            row.monthValues=valid?values:null;changed();message(valid?'Kladden er ændret.':'Brug én unik måned pr. linje: 2026-08 = 1500.50',!valid);
        };
        el('rowName').oninput = event => { row.name = event.target.value; changed(); drawRows(); drawRoles(); };
        el('rowType').onchange = event => {
            const type = event.target.value;
            Object.assign(row, { type });
            if (type === 'accounts') Object.assign(row,{accounts:row.accounts||[],groups:row.groups||[],exclude:row.exclude||[],sign:row.sign||(el('section').value==='pnl'?-1:1),children:row.children||false});
            if (['sum','percent'].includes(type)) Object.assign(row,{sources:row.sources||[],rate:row.rate??-22});
            if (type === 'formula') row.formula = row.formula || '0';
            if (['lager','manual'].includes(type)) Object.assign(row,{period:row.period||'selected',metric:row.metric||'margin',monthValues:row.monthValues||{}});
            changed(); draw();
        };
        if (el('sign')) el('sign').onchange = event => { row.sign = Number(event.target.value); changed(); };
        if (el('children')) el('children').onchange = event => { row.children = event.target.checked; changed(); };
        if (el('rate')) el('rate').oninput = event => { row.rate = event.target.value === '' ? null : Number(event.target.value); changed(); };
        el('editor').querySelectorAll('[data-remove]').forEach(button => button.onclick = () => {
            const key = button.dataset.remove, value = Number(button.dataset.value);
            row[key] = row[key].filter((v,i) => key === 'groups' ? i !== value : v !== value); changed(); drawEditor();
        });
        el('editor').querySelectorAll('[data-source]').forEach(box => box.onchange = () => { row.sources = [...el('editor').querySelectorAll('[data-source]:checked')].map(b=>b.dataset.source); changed(); });
        el('deleteRow').onclick = () => {
            const sourceUsers = rows().filter(r => (r.sources || []).includes(row.id));
            const formulaUsers = [...definition.pnl, ...definition.balance].filter(r => r.type==='formula' && (r.formula||'').includes('['+row.id+']'));
            const dependents = [...new Set([...sourceUsers, ...formulaUsers])];
            if (dependents.length) return message('Fjern først rækken fra beregningerne: ' + dependents.map(r=>r.name).join(', '),true);
            if (el('section').value === 'pnl' && [definition.revenueRow,definition.beforeTaxRow].includes(row.id)) return message('Vælg først en anden række som omsætning / resultat før skat.',true);
            rows().splice(selected,1); selected = Math.max(0,selected-1); changed(); draw();
        };
        simplifyEditor(row);
    }
    function simplifyEditor(row) {
        el('editorTitle').textContent=row.type==='heading'?'Redigér overskrift':'Redigér række';
        el('rowType').parentElement.firstChild.textContent='Indhold';
        const advanced=document.createElement('details');advanced.className='advanced';
        advanced.innerHTML='<summary>Flere muligheder</summary><div class="form"></div>';
        const target=advanced.lastElementChild;
        for(const id of ['sign','children','visibility','bold','checkZero','currentSales']) {
            const input=el(id);if(input)target.append(input.closest('label'));
        }
        const reference=el('editor').querySelector('p.muted code')?.parentElement;if(reference)target.append(reference);
        const exclusions=el('editor').querySelector('[data-remove="exclude"]');
        if(exclusions)target.append(exclusions.parentElement.parentElement);
        else for(const child of [...el('editor').children])if(child.tagName==='DIV' && child.firstChild?.textContent==='Udelad konti')target.append(child);
        target.append(el('deleteRow'));el('editor').append(advanced);
        if(el('sourceMonth'))el('sourceMonth').closest('label').hidden=row.period!=='fixed';
        if(row.type==='formula') {
            const formulaInput=el('formula'), chips=formulaInput.closest('label').nextElementSibling;
            const details=document.createElement('details');details.className='advanced';details.innerHTML='<summary>Redigér formel</summary>';
            const summary=document.createElement('div');summary.className='calculation-summary';
            const names=new Map([...definition.pnl,...definition.balance].map(r=>[r.id,r.name]));
            const update=()=>{summary.textContent=(row.formula||'').replace(/\[([^\]]+)\]/g,(_,id)=>names.get(id)||id);};update();
            formulaInput.addEventListener('input',update);
            formulaInput.closest('label').before(summary,details);details.append(formulaInput.closest('label'),chips);
        }
        el('catalogPanel').hidden=row.type!=='accounts';drawCatalog();
    }
    function drawCatalog() {
        if (!catalog) return;
        const query = el('catalogSearch').value.toLocaleLowerCase();
        if(!query.trim()){el('catalog').innerHTML='<p class="muted">Søg efter en konto eller gruppe. Tryk + for at tilføje den til rækken.</p>';return;}
        const items = [ ...catalog.accounts.map(a=>({label:`${a.AcNo} · ${a.Nm}`,sub:a.AcGr,value:Number(a.AcNo),kind:'account'})),
            ...catalog.groups.map(g=>({label:g.AcGr,sub:g.AgAcGr?`Samles i: ${g.AgAcGr}`:'Kontogruppe',value:g.AcGr,kind:'group'})) ];
        const matches = items.filter(i=>(i.label+' '+i.sub).toLocaleLowerCase().includes(query));
        el('catalog').innerHTML = matches.slice(0,80).map((item,index)=>`<div class="pick"><span>${esc(item.label)}</span><small>${esc(item.sub)}</small><div><button data-add="${index}">+ Medtag ${item.kind==='group'?'gruppe':'konto'}</button>${item.kind==='account'?`<button data-exclude="${index}">Udelad</button>`:''}</div></div>`).join('') + (matches.length>80?'<p class="muted">Viser de første 80. Søg for at indsnævre.</p>':'');
        const add = (index, exclude) => {
            const row = current(), item = matches[index];
            if (!row || row.type !== 'accounts') return message('Vælg først en række af typen Konti / grupper.',true);
            const key = exclude?'exclude':item.kind==='group'?'groups':'accounts';
            if (!row[key].includes(item.value)) row[key].push(item.value);
            changed(); drawEditor(); message('Kladden er ændret. Kontrollér den før gem.');
        };
        el('catalog').querySelectorAll('[data-add]').forEach(b=>b.onclick=()=>add(Number(b.dataset.add),false));
        el('catalog').querySelectorAll('[data-exclude]').forEach(b=>b.onclick=()=>add(Number(b.dataset.exclude),true));
    }
    function draw() { drawRows(); drawEditor(); drawRoles(); el('legacyRows').checked=definition.legacyPeriodRows!==false; }
    async function load(reportId) {
        if(typeof reportId!=='string')reportId=definition?.reportId||'default';
        if (dirty && !confirm('Kassér ikke-gemte ændringer og genindlæs?')) {drawReports();return;}
        lock(true); message('Henter opsætning og kontoplan…');
        try {
            const [data,list] = await Promise.all([api('/admin/bilancio-definition?reportId='+encodeURIComponent(reportId)),api('/bilancio-trial/reports')]); definition = data.definition; catalog = data.catalog;reportList=list.reports;drawReports();
            dirty=false; selected=0; el('balanceTitle').value=definition.balanceTitle; el('workspace').hidden=false;
            draw();drawCatalog(); el('preview').innerHTML='';el('revision').textContent=`Version ${definition.version}${definition.updatedBy?' · '+definition.updatedBy:''}`; message('Opsætningen er hentet.');
        } catch(error) { message(error.message,true); } finally { lock(false); }
    }
    el('reload').onclick=load;
    el('showHelpers').onchange=drawRows;
    el('showPreview').onclick=()=>{el('previewPanel').open=true;el('previewPanel').scrollIntoView({behavior:'smooth',block:'start'});el('previewBtn').click();};
    el('openPassiver').onclick=()=>{el('passiverError').textContent='';drawRoles();el('passiverDialog').showModal();};
    el('cancelPassiver').onclick=()=>el('passiverDialog').close();
    el('reportSelect').onchange=e=>load(e.target.value);
    el('reportName').oninput=e=>{definition.reportName=e.target.value;changed();};
    el('showPnl').onchange=e=>{definition.showPnl=e.target.checked;changed();};
    async function startReport(duplicate) {
        const name=await askText('Navn på den nye rapport',duplicate?(definition.reportName||'Økonomirapport')+' — kopi':'Ny balance');
        if(!name?.trim())return;
        if(!duplicate && dirty && !confirm('Kassér ikke-gemte ændringer?'))return;
        const scope=definition.scope;
        definition=duplicate?structuredClone(definition):{schema:1,scope,pnl:[{id:'base',name:'Grundlag',type:'formula',formula:'0',visibility:'hidden'}],revenueRow:'base',beforeTaxRow:'base',balanceTitle:'Balance',balance:[{id:'aktiver',name:'Aktiver',type:'heading'}],showPnl:false,legacyPeriodRows:false};
        Object.assign(definition,{reportId:uid('report_'),reportName:name.trim(),version:0});delete definition.updatedBy;delete definition.updatedAt;
        el('section').value='balance';selected=0;el('balanceTitle').value=definition.balanceTitle;changed();drawReports();draw();message('Ny kladde. Gem opsætning opretter rapporten.');
    }
    el('newReport').onclick=()=>startReport(false);el('duplicateReport').onclick=()=>startReport(true);
    el('newSection').onclick=async()=>{const name=await askText('Sektionsnavn (f.eks. Passiver)');if(!name?.trim())return;rows().push({id:uid('section_'),name:name.trim(),type:'heading'});selected=rows().length-1;changed();draw();};
    el('addPassiver').onclick=()=>{
        const list=definition.balance;let assetId=el('assetTotal').value;
        if(!assetId){
            const accounts=list.filter(r=>r.type==='accounts').map(r=>r.id);
            if(!accounts.length){el('passiverError').textContent='Tilføj først Aktiver-konti, eller vælg en eksisterende Aktiver-total.';return;}
            assetId=uid('assets_');list.push({id:assetId,name:'Aktiver i alt',type:'sum',sources:accounts,bold:true});
        }
        const accountId=uid('liabilities_'),totalId=uid('liabilityTotal_'),negative=el('liabilitySign').value==='negative';
        list.push({id:uid('section_'),name:'Passiver',type:'heading'},
            {id:accountId,name:'Passiver — vælg konti',type:'accounts',accounts:[],groups:[],exclude:[],children:false,sign:negative?1:-1},
            {id:totalId,name:'Passiver i alt',type:'sum',sources:[accountId],bold:true},
            {id:uid('check_'),name:'Kontrol · Aktiver / Passiver',type:'formula',formula:'['+assetId+'] '+(negative?'+':'-')+' ['+totalId+']',bold:true,checkZero:true});
        definition.balanceTitle='Balance';el('balanceTitle').value='Balance';el('section').value='balance';selected=list.length-3;changed();draw();el('passiverDialog').close();message('Søg og tilføj dine passivkonti. Total og kontrol er oprettet.');el('catalogSearch').focus();
    };
    el('section').onchange=()=>{selected=0;draw();};
    el('catalogSearch').oninput=drawCatalog;
    el('balanceTitle').oninput=e=>{definition.balanceTitle=e.target.value;changed();};
    el('revenue').onchange=e=>{definition.revenueRow=e.target.value;changed();};
    el('beforeTax').onchange=e=>{definition.beforeTaxRow=e.target.value;changed();};
    el('legacyRows').onchange=e=>{definition.legacyPeriodRows=e.target.checked;changed();};
    el('convertVia').onclick=()=>{
        if(definition.legacyPeriodRows===false) return message('De automatiske rækker er allerede slået fra. Du kan tilføje flere rækker med + Ny række.');
        const suffix=Date.now().toString(36), sales='sales_'+suffix, cost='cost_'+suffix, before='before_'+suffix, margin='margin_'+suffix;
        definition.pnl.push(
            {id:sales,name:'SO salgpris',type:'lager',metric:'sales',period:'selected',currentSales:true,visibility:'hidden'},
            {id:cost,name:'SO kostpris',type:'lager',metric:'cost',period:'selected',visibility:'hidden'},
            {id:before,name:'Periodens resultat før skat',type:'formula',formula:'['+definition.beforeTaxRow+']',visibility:'period',bold:true},
            {id:margin,name:'Regulering fortjeneste ej realiseret (VIA)',type:'formula',formula:'['+sales+'] - ['+cost+']',visibility:'period'},
            {id:'total_'+suffix,name:'Periodens resultat i alt',type:'formula',formula:'['+before+'] + ['+margin+']',visibility:'period',bold:true}
        );
        definition.legacyPeriodRows=false;el('section').value='pnl';selected=definition.pnl.length-3;changed();draw();message('VIA-rækkerne er nu redigerbare i kladden. Kontrollér og gem.');
    };
    el('add').onclick=()=>{rows().push({id:'r'+crypto.randomUUID().replaceAll('-',''),name:'Ny række',type:'accounts',accounts:[],groups:[],exclude:[],children:false,sign:el('section').value==='pnl'?-1:1});selected=rows().length-1;changed();draw();};
    function move(delta) { const list=rows(), next=selected+delta; if(next<0||next>=list.length)return; [list[selected],list[next]]=[list[next],list[selected]];selected=next;changed();draw(); }
    el('up').onclick=()=>move(-1); el('down').onclick=()=>move(1);
    el('save').onclick=async()=>{
        lock(true);message('Kontrollerer og gemmer…');
        try { const data=await api('/admin/bilancio-definition',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(definition)});definition=data.definition;dirty=false;reportList=reportList.filter(r=>r.reportId!==definition.reportId).concat([{reportId:definition.reportId,reportName:definition.reportName}]);drawReports();draw();el('revision').textContent=`Version ${definition.version} · Gemt`;message('Gemt på GOH. Genindlæs rapporten for at bruge opsætningen.'); }
        catch(error){message(error.message,true);}finally{lock(false);}
    };
    const fmt=v=>v==null?'—':new Intl.NumberFormat('da-DK',{maximumFractionDigits:2}).format(v);
    el('previewBtn').onclick=async()=>{
        lock(true);message('Kontrollerer kladden og beregner…');el('preview').innerHTML='';
        try {
            const data=await api('/admin/bilancio-definition/preview',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({definition,year:Number(el('year').value),period:Number(el('period').value)})});
            const table=(title,rows,columns,balance=false)=>`<h3>${esc(title)} · DKK</h3><table><thead><tr><th>Navn</th>${columns.map(c=>`<th>${c}</th>`).join('')}</tr></thead><tbody>${rows.filter(row=>row.visibility!=='hidden' && !(balance && row.visibility==='period')).map(row=>`<tr class="${row.bold?'subtotal':row.type}"><td>${esc(row.name)}</td>${row.amounts.slice(0,columns.length).map((v,i)=>`<td>${row.type==='heading'||(!balance && ((row.visibility==='period' && i>1)||(row.visibility==='ytd' && i<2)))?'':fmt(v)+(row.checkZero?(v==null?' · Mangler data':Math.abs(v)<0.005?' · ✓ OK':' · ⚠ Afvigelse'):'')}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
            el('preview').innerHTML=(data.showPnl===false?'':table('Resultatopgørelse',[...data.rows,...(data.periodOnly||[])],['Måned','Måned sidste år','År til dato','År til dato sidste år']))+table(data.assets.title,data.assets.rows,['År til dato','Sidste år'],true);
            message('Kladden er kontrolleret. Ingen ændringer er gemt.');
        }catch(error){message(error.message,true);}finally{lock(false);}
    };
    const now=new Date();el('year').value=now.getFullYear()-(now.getMonth()<6?1:0);
    ['Juli','August','September','Oktober','November','December','Januar','Februar','Marts','April','Maj','Juni'].forEach((m,i)=>el('period').add(new Option(`${i+1} · ${m}`,i+1)));
    el('period').value=(now.getMonth()+6)%12+1;
    window.addEventListener('beforeunload',event=>{if(dirty){event.preventDefault();event.returnValue='';}});
    load();
})();
