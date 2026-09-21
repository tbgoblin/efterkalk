(() => {
    const el = id => document.getElementById(id);
    const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    let definition, catalog, selected = 0, dirty = false, busy = false;
    const labels = { accounts: 'Konti / grupper', sum: 'Sum af valgte rækker', percent: 'Procent af valgte rækker', heading: 'Overskrift' };
    const rows = () => definition[el('section').value];
    const current = () => rows()[selected];
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
    }
    function drawRoles() {
        const options = definition.pnl.filter(row => row.type !== 'heading').map(row => `<option value="${row.id}">${esc(row.name)}</option>`).join('');
        el('revenue').innerHTML = options; el('revenue').value = definition.revenueRow;
        el('beforeTax').innerHTML = options; el('beforeTax').value = definition.beforeTaxRow;
    }
    function drawRows() {
        el('rows').innerHTML = rows().map((row, index) => `<div class="row"><button type="button" data-index="${index}" class="${index===selected?'selected':''}">${esc(row.name)}<small>${labels[row.type]}</small></button></div>`).join('');
        el('rows').querySelectorAll('[data-index]').forEach(button => button.onclick = () => { selected = Number(button.dataset.index); draw(); });
        el('up').disabled = selected === 0; el('down').disabled = selected === rows().length - 1;
    }
    function drawEditor() {
        const row = current();
        if (!row) { el('editor').innerHTML = '<p>Tilføj en række.</p>'; return; }
        el('editor').innerHTML = `<label>Navn i rapporten<input id="rowName" maxlength="160" value="${esc(row.name)}"></label><label>Type<select id="rowType">${Object.entries(labels).map(([v,l])=>`<option value="${v}" ${row.type===v?'selected':''}>${l}</option>`).join('')}</select></label>` +
            (row.type === 'accounts' ? `<label>Fortegn<select id="sign"><option value="1" ${row.sign===1?'selected':''}>Behold kontofortegn (typisk balance)</option><option value="-1" ${row.sign===-1?'selected':''}>Vend fortegn (typisk resultatopgørelse)</option></select></label><label class="check"><input id="children" type="checkbox" ${row.children?'checked':''}>Medtag undergrupper</label><div>Konti<div class="chips">${row.accounts.map(ac=>`<button class="chip" data-remove="accounts" data-value="${ac}">${ac} ×</button>`).join('')||'<span class="muted">Ingen direkte konti</span>'}</div></div><div>Kontogrupper<div class="chips">${row.groups.map((g,i)=>`<button class="chip" data-remove="groups" data-value="${i}">${esc(g)} ×</button>`).join('')||'<span class="muted">Ingen grupper</span>'}</div></div><div>Udelad konti<div class="chips">${row.exclude.map(ac=>`<button class="chip" data-remove="exclude" data-value="${ac}">${ac} ×</button>`).join('')||'<span class="muted">Ingen undtagelser</span>'}</div></div>` :
            ['sum','percent'].includes(row.type) ? `<div>Medregn disse rækker<div class="source-list">${rows().slice(0,selected).filter(r=>r.type!=='heading').map(r=>`<label><input type="checkbox" data-source="${r.id}" ${row.sources.includes(r.id)?'checked':''}>${esc(r.name)}</label>`).join('')||'<p class="muted">Flyt rækken under de rækker, du vil summere.</p>'}</div></div>${row.type==='percent'?`<label>Procent (f.eks. −22 til skat)<input id="rate" type="number" step="any" min="-1000" max="1000" value="${row.rate}"></label>`:''}` : '<p class="muted">En overskrift har intet beløb og ændrer ikke totalerne.</p>') + '<button id="deleteRow" type="button">Fjern række</button>';
        el('rowName').oninput = event => { row.name = event.target.value; changed(); drawRows(); drawRoles(); };
        el('rowType').onchange = event => {
            const type = event.target.value;
            Object.assign(row, { type });
            if (type === 'accounts') Object.assign(row,{accounts:row.accounts||[],groups:row.groups||[],exclude:row.exclude||[],sign:row.sign||(el('section').value==='pnl'?-1:1),children:row.children||false});
            if (['sum','percent'].includes(type)) Object.assign(row,{sources:row.sources||[],rate:row.rate??-22});
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
            const dependents = rows().filter(r => (r.sources || []).includes(row.id));
            if (dependents.length) return message('Fjern først rækken fra beregningerne: ' + dependents.map(r=>r.name).join(', '),true);
            if (el('section').value === 'pnl' && [definition.revenueRow,definition.beforeTaxRow].includes(row.id)) return message('Vælg først en anden række som omsætning / resultat før skat.',true);
            rows().splice(selected,1); selected = Math.max(0,selected-1); changed(); draw();
        };
    }
    function drawCatalog() {
        if (!catalog) return;
        const query = el('catalogSearch').value.toLocaleLowerCase();
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
    function draw() { drawRows(); drawEditor(); drawRoles(); }
    async function load() {
        if (dirty && !confirm('Kassér ikke-gemte ændringer og genindlæs?')) return;
        lock(true); message('Henter opsætning og kontoplan…');
        try {
            const data = await api('/admin/bilancio-definition'); definition = data.definition; catalog = data.catalog;
            dirty=false; selected=0; el('balanceTitle').value=definition.balanceTitle; el('workspace').hidden=false;
            draw();drawCatalog(); el('preview').innerHTML='';el('revision').textContent=`Version ${definition.version}${definition.updatedBy?' · '+definition.updatedBy:''}`; message('Opsætningen er hentet.');
        } catch(error) { message(error.message,true); } finally { lock(false); }
    }
    el('reload').onclick=load;
    el('section').onchange=()=>{selected=0;draw();};
    el('catalogSearch').oninput=drawCatalog;
    el('balanceTitle').oninput=e=>{definition.balanceTitle=e.target.value;changed();};
    el('revenue').onchange=e=>{definition.revenueRow=e.target.value;changed();};
    el('beforeTax').onchange=e=>{definition.beforeTaxRow=e.target.value;changed();};
    el('add').onclick=()=>{rows().push({id:'r'+crypto.randomUUID().replaceAll('-',''),name:'Ny række',type:'accounts',accounts:[],groups:[],exclude:[],children:false,sign:el('section').value==='pnl'?-1:1});selected=rows().length-1;changed();draw();};
    function move(delta) { const list=rows(), next=selected+delta; if(next<0||next>=list.length)return; [list[selected],list[next]]=[list[next],list[selected]];selected=next;changed();draw(); }
    el('up').onclick=()=>move(-1); el('down').onclick=()=>move(1);
    el('save').onclick=async()=>{
        lock(true);message('Kontrollerer og gemmer…');
        try { const data=await api('/admin/bilancio-definition',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(definition)});definition=data.definition;dirty=false;draw();el('revision').textContent=`Version ${definition.version} · Gemt`;message('Gemt på GOH. Genindlæs rapporten for at bruge opsætningen.'); }
        catch(error){message(error.message,true);}finally{lock(false);}
    };
    const fmt=v=>v==null?'—':new Intl.NumberFormat('da-DK',{maximumFractionDigits:2}).format(v);
    el('previewBtn').onclick=async()=>{
        lock(true);message('Kontrollerer kladden og beregner…');el('preview').innerHTML='';
        try {
            const data=await api('/admin/bilancio-definition/preview',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({definition,year:Number(el('year').value),period:Number(el('period').value)})});
            const table=(title,rows,columns)=>`<h3>${esc(title)} · DKK</h3><table><thead><tr><th>Navn</th>${columns.map(c=>`<th>${c}</th>`).join('')}</tr></thead><tbody>${rows.map(row=>`<tr class="${row.type}"><td>${esc(row.name)}</td>${row.amounts.slice(0,columns.length).map(v=>`<td>${row.type==='heading'?'':fmt(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
            el('preview').innerHTML=table('Resultatopgørelse',data.rows,['Måned','Måned sidste år','År til dato','År til dato sidste år'])+table(data.assets.title,data.assets.rows,['År til dato','Sidste år']);
            message('Kladden er kontrolleret. Ingen ændringer er gemt.');
        }catch(error){message(error.message,true);}finally{lock(false);}
    };
    const now=new Date();el('year').value=now.getFullYear()-(now.getMonth()<6?1:0);
    ['Juli','August','September','Oktober','November','December','Januar','Februar','Marts','April','Maj','Juni'].forEach((m,i)=>el('period').add(new Option(`${i+1} · ${m}`,i+1)));
    el('period').value=(now.getMonth()+6)%12+1;
    window.addEventListener('beforeunload',event=>{if(dirty){event.preventDefault();event.returnValue='';}});
    load();
})();
