(function () {
    'use strict';
    const model=window.GohMonthlySummaryModel;
    const el=id => document.getElementById(id);
    const esc=value => String(value ?? '').replace(/[&<>"']/g,c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    const types={source:'Ordre- eller lagerværdi',accounts:'Konti / kontogrupper',formula:'Formel',manual:'Manuelt beløb pr. måned',heading:'Overskrift'};
    let month='',payload=null,metrics={},requestId=0,draft=null,catalog=null,selected=0,dirty=false,busy=false,uid=0;
    const status=(message,error=false) => { el('kfSummaryStatus').textContent=message; el('kfSummaryStatus').dataset.error=String(error); };
    async function api(url,options) {
        const response=await fetch(url,options);
        const data=await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'Oversigten kunne ikke hentes.');
        return data;
    }
    const json=(method,body) => ({method,headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
    function valueText(row) {
        if (row.value == null) return '—';
        const options=row.unit==='count' ? {maximumFractionDigits:0} : {minimumFractionDigits:2,maximumFractionDigits:2};
        return new Intl.NumberFormat('da-DK',options).format(row.value) + (row.unit==='percent' ? ' %' : row.unit==='dkk' ? ' DKK' : '');
    }
    function renderReport(data,target) {
        const rows=model.evaluate(data.definition,{...data,month,metrics});
        const details=row => row.type==='accounts' ? '<details class="kfs-details"><summary>Bogføring · '+row.accounts.length+' konti</summary>'
            + row.accounts.map(account => '<div>'+esc(account.AcNo)+' '+esc(account.Nm)+' <b>'+esc(valueText({...row,value:Number(account.Amount)*row.sign}))+'</b></div>').join('')+'</details>' : '';
        const cards=rows.filter(row => row.display==='card' && row.type!=='heading');
        const sections=new Map();
        for (const row of rows.filter(row => row.display==='row')) {
            const name=row.section || 'Øvrige tal';
            if (!sections.has(name)) sections.set(name,[]);
            sections.get(name).push(row);
        }
        target.innerHTML='<h3>'+esc(data.definition.title)+'</h3><div class="kfs-cards">'
            +cards.map(row => '<article class="kfs-card"><span>'+esc(row.name)+'</span><strong>'+esc(valueText(row))+'</strong>'+details(row)+'</article>').join('')+'</div><div class="kfs-sections">'
            +[...sections].map(([name,items]) => '<section class="kfs-section"><h4>'+esc(name)+'</h4>'
                +items.map(row => row.type==='heading' ? '<h5>'+esc(row.name)+'</h5>' : '<div class="kfs-row"><div>'+esc(row.name)+details(row)+'</div><strong class="'+(row.value<0 ? 'kfs-negative' : '')+'">'+esc(valueText(row))+'</strong></div>').join('')+'</section>').join('')+'</div>';
    }
    function render() {
        if (!payload || !month) return;
        renderReport(payload,el('kfSummaryReport'));
        el('kfSummaryReport').hidden=false;
        el('kfSummaryLegacy').classList.add('kfs-configured');
    }
    async function load(value) {
        const token=++requestId;
        month=value; payload=null;
        el('kfSummaryLegacy').classList.remove('kfs-configured');
        el('kfSummaryReport').hidden=true;
        el('kfSummaryBuild').hidden=true;
        status('Henter opsætning og bogføring…');
        try {
            const data=await api('/efterkalk/summary?month='+encodeURIComponent(month));
            if (token!==requestId) return;
            payload=data;
            el('kfSummaryBuild').hidden=!data.canEdit;
            render(); status('Bogføring for '+month+' · øvrige tal følger de indlæste ordrer.');
        } catch(error) { if (token===requestId) status(error.message+' Standardvisningen bruges indtil videre.',true); }
    }
    function capture() {
        if (!draft || !draft.rows[selected]) return;
        const row=draft.rows[selected];
        for (const field of ['name','display','unit','section','source','formula']) {
            const input=el('kfs-'+field); if (input) row[field]=input.value;
        }
        if (row.type==='accounts') { row.sign=Number(el('kfs-sign').value); row.children=el('kfs-children').checked; }
        if (row.type==='manual') {
            const values={};
            for (const line of el('kfs-manual').value.split('\n').filter(line => line.trim())) {
                const match=/^\s*(\d{4}-(?:0[1-9]|1[0-2]))\s*=\s*(-?(?:\d+(?:[.,]\d+)?|[.,]\d+))\s*$/.exec(line);
                if (!match) throw new Error('Manuelle beløb: brug fx 2026-10 = 1500,50.');
                if (Object.hasOwn(values,match[1])) throw new Error('Måneden '+match[1]+' står flere gange.');
                values[match[1]]=Number(match[2].replace(',','.'));
            }
            row.monthValues=values;
        }
        draft.title=el('kfs-title').value;
    }
    function errorText(error) { el('kfs-message').textContent=error?.message || String(error || ''); }
    function nav() {
        el('kfs-list').innerHTML=draft.rows.map((row,index) => '<button type="button" data-select="'+index+'" aria-current="'+(index===selected)+'"><span>'+esc(row.name || 'Ny række')+'</span><small>'+esc(types[row.type])+' · '+esc(row.display)+'</small></button>').join('');
    }
    function options(values,current) { return Object.entries(values).map(([value,label]) => '<option value="'+esc(value)+'"'+(value===current?' selected':'')+'>'+esc(label)+'</option>').join(''); }
    function drawCatalog() {
        const host=el('kfs-catalog'); if (!host) return;
        const row=draft.rows[selected],query=el('kfs-search').value.toLocaleLowerCase(),exclude=el('kfs-catalog-mode').value==='exclude';
        const items=[...catalog.accounts.map(account => ({kind:'accounts',value:Number(account.AcNo),label:account.AcNo+' · '+account.Nm,sub:account.AcGr || ''})),
            ...(!exclude ? catalog.groups.map(group => ({kind:'groups',value:String(group.AcGr).trim(),label:String(group.AcGr).trim(),sub:'Kontogruppe'})) : [])]
            .filter(item => (item.label+' '+item.sub).toLocaleLowerCase().includes(query));
        host.innerHTML=items.slice(0,100).map(item => {
            const kind=exclude?'exclude':item.kind,checked=(row[kind] || []).includes(item.value);
            return '<label><input type="checkbox" data-kind="'+kind+'" data-value="'+esc(item.value)+'"'+(checked?' checked':'')+'><span>'+esc(item.label)+'<small>'+esc(item.sub)+'</small></span></label>';
        }).join('')+'<p class="kfs-muted">'+items.length+' fundet'+(items.length>100?' · søg for at indsnævre listen':'')+'</p>';
        el('kfs-selected').innerHTML=['accounts','groups','exclude'].flatMap(kind => (row[kind] || []).map(value => '<button type="button" data-remove-kind="'+kind+'" data-value="'+esc(value)+'">'+(kind==='exclude'?'Udelad ':'')+esc(value)+' ×</button>')).join('');
    }
    function editorRow() {
        const row=draft.rows[selected];
        nav(); el('kfs-preview').innerHTML='';
        const field=(label,id,content) => '<label>'+label+content+'</label>';
        let extra='';
        if (row.type==='source') extra=field('Datakilde','source','<select id="kfs-source">'+options(model.sources,row.source)+'</select>');
        if (row.type==='accounts') extra='<div class="kfs-grid">'+field('Fortegn','sign','<select id="kfs-sign">'+options({'1':'Som bogført (+1)','-1':'Vend fortegn (−1)'},String(row.sign))+'</select>')
            +'<label class="kfs-check"><input id="kfs-children" type="checkbox"'+(row.children?' checked':'')+'>Medtag undergrupper</label></div>'
            +'<div id="kfs-selected" class="kfs-chips"></div><div class="kfs-grid"><input id="kfs-search" type="search" placeholder="Søg kontonummer, navn eller gruppe…" aria-label="Søg konti og grupper">'
            +'<select id="kfs-catalog-mode" aria-label="Medtag eller udelad"><option value="include">Medtag konti / grupper</option><option value="exclude">Udelad enkelte konti</option></select></div><div id="kfs-catalog" class="kfs-catalog"></div>';
        if (row.type==='formula') extra=field('Formel','formula','<textarea id="kfs-formula" rows="3" maxlength="2000" placeholder="[booked] - [material]">'+esc(row.formula || '')+'</textarea>')
            +'<p class="kfs-muted">Brug +, −, *, /, % og parenteser. Klik på en talrække ovenfor for at indsætte den.</p><div class="kfs-chips">'
            +draft.rows.slice(0,selected).filter(item => item.type!=='heading').map(item => '<button type="button" data-insert="'+esc(item.id)+'">'+esc(item.name)+' ['+esc(item.id)+']</button>').join('')+'</div>';
        if (row.type==='manual') extra=field('Beløb pr. måned — én linje pr. måned','manual','<textarea id="kfs-manual" rows="6" placeholder="2026-10 = 1500,50">'+esc(Object.entries(row.monthValues || {}).map(([key,value])=>key+' = '+value).join('\n'))+'</textarea>');
        el('kfs-fields').innerHTML='<p class="kfs-muted">Række-id: '+esc(row.id)+'</p><div class="kfs-grid">'
            +field('Navn','name','<input id="kfs-name" maxlength="160" value="'+esc(row.name)+'">')
            +field('Type','type','<select id="kfs-type">'+options(types,row.type)+'</select>')+'</div><div class="kfs-grid">'
            +field('Visning','display','<select id="kfs-display">'+options({...(row.type!=='heading'?{card:'KPI-kort'}:{}),row:'Række i sektion',hidden:'Skjult (kan bruges i formler)'},row.display)+'</select>')
            +field('Enhed','unit','<select id="kfs-unit">'+options({dkk:'DKK',count:'Antal',percent:'Procent'},row.unit)+'</select>')+'</div>'
            +field('Sektion (for rækker)','section','<input id="kfs-section" maxlength="160" value="'+esc(row.section)+'">')+extra;
        if (row.type==='accounts') drawCatalog();
    }
    function changeType(row,type) {
        const common={id:row.id,name:row.name,type,display:type==='heading'?'row':row.display,unit:row.unit,section:row.section};
        return {...common,...(type==='accounts'?{accounts:[],groups:[],exclude:[],children:false,sign:1}:type==='source'?{source:'invoiced'}:type==='formula'?{formula:''}:type==='manual'?{monthValues:{}}:{})};
    }
    async function openEditor() {
        if (busy) return;
        try {
            const data=await api('/efterkalk/summary-definition');
            draft=structuredClone(data.definition); catalog=data.catalog; selected=0; dirty=false;
            el('kfs-title').value=draft.title; errorText(''); editorRow(); el('kfSummaryDialog').showModal();
        } catch(error) { status(error.message,true); }
    }
    function closeEditor() {
        if (busy || (dirty && !window.confirm('Luk uden at gemme ændringerne?'))) return;
        el('kfSummaryDialog').close(); draft=null;
    }
    async function commit(preview) {
        if (busy) return;
        busy=true; el('kfs-save').disabled=true; el('kfs-preview-button').disabled=true;
        try {
            capture(); const definition=model.validate(draft);
            if (preview) {
                const data=await api('/efterkalk/summary-definition/preview',json('POST',{month,definition}));
                renderReport(data,el('kfs-preview')); errorText('Eksempel for '+month+' · ændringerne er ikke gemt.');
            } else {
                const saved=await api('/efterkalk/summary-definition',json('PUT',definition));
                dirty=false; draft=saved.definition; el('kfSummaryDialog').close();
                await load(month);
            }
        } catch(error) { errorText(error); }
        finally { busy=false; el('kfs-save').disabled=false; el('kfs-preview-button').disabled=false; }
    }
    const guarded=action => { try { capture(); action(); } catch(error) { errorText(error); } };
    el('kfSummaryBuild').addEventListener('click',openEditor);
    el('kfs-close').addEventListener('click',closeEditor);
    el('kfSummaryDialog').addEventListener('cancel',event => {event.preventDefault();closeEditor();});
    el('kfs-save').addEventListener('click',()=>commit(false));
    el('kfs-preview-button').addEventListener('click',()=>commit(true));
    el('kfs-add').addEventListener('click',()=>guarded(()=>{
        draft.rows.push(changeType({id:'r'+Date.now().toString(36)+(uid++),name:'Ny række',display:'row',unit:'dkk',section:'Øvrige tal'},el('kfs-new-type').value));
        selected=draft.rows.length-1; dirty=true;editorRow();
    }));
    el('kfs-up').addEventListener('click',()=>guarded(()=>{if(selected>0){[draft.rows[selected-1],draft.rows[selected]]=[draft.rows[selected],draft.rows[selected-1]];selected--;dirty=true;editorRow();}}));
    el('kfs-down').addEventListener('click',()=>guarded(()=>{if(selected<draft.rows.length-1){[draft.rows[selected+1],draft.rows[selected]]=[draft.rows[selected],draft.rows[selected+1]];selected++;dirty=true;editorRow();}}));
    el('kfs-delete').addEventListener('click',()=>{if(draft.rows.length<=1)return;draft.rows.splice(selected,1);selected=Math.min(selected,draft.rows.length-1);dirty=true;editorRow();});
    el('kfs-defaults').addEventListener('click',()=>{if(!window.confirm('Erstat kladden med standardopsætningen? Den gemmes først, når du klikker Gem.'))return;draft={...model.defaults(),version:draft.version,scope:draft.scope};selected=0;dirty=true;el('kfs-title').value=draft.title;editorRow();});
    el('kfs-list').addEventListener('click',event=>{const button=event.target.closest('[data-select]');if(button)guarded(()=>{selected=Number(button.dataset.select);editorRow();});});
    el('kfs-fields').addEventListener('input',event=>{
        if(event.target.id==='kfs-search'){drawCatalog();return;}
        dirty=true;
    });
    el('kfs-title').addEventListener('input',()=>{dirty=true;});
    el('kfs-fields').addEventListener('change',event=>{
        if(event.target.id==='kfs-catalog-mode'){drawCatalog();return;}
        guarded(()=>{
            const row=draft.rows[selected];
            if(event.target.id==='kfs-type'){draft.rows[selected]=changeType(row,event.target.value);dirty=true;editorRow();return;}
            if(event.target.id==='kfs-source'){
                row.unit=['orders','credit_count','ub_orders'].includes(row.source)?'count':row.source==='dg'?'percent':'dkk';
                el('kfs-unit').value=row.unit;
            }
            if(event.target.dataset.kind){const kind=event.target.dataset.kind,value=kind==='groups'?event.target.dataset.value:Number(event.target.dataset.value);row[kind]=row[kind].filter(item=>item!==value);if(event.target.checked)row[kind].push(value);dirty=true;drawCatalog();}
            nav();
        });
    });
    el('kfs-fields').addEventListener('click',event=>{
        const insert=event.target.closest('[data-insert]');
        if(insert){const input=el('kfs-formula');const start=input.selectionStart,end=input.selectionEnd;input.setRangeText('['+insert.dataset.insert+']',start,end,'end');input.focus();dirty=true;return;}
        const remove=event.target.closest('[data-remove-kind]');
        if(remove){const kind=remove.dataset.removeKind,value=kind==='groups'?remove.dataset.value:Number(remove.dataset.value);draft.rows[selected][kind]=draft.rows[selected][kind].filter(item=>item!==value);dirty=true;drawCatalog();}
    });
    window.GohMonthlySummary={
        load,
        update(values){metrics=values;render();},
        mode(monthly){el('kfSummaryTools').hidden=!monthly;if(!monthly){requestId++;month='';payload=null;el('kfSummaryReport').hidden=true;el('kfSummaryLegacy').classList.remove('kfs-configured');}}
    };
}());
