(function () {
    'use strict';
    const $ = id => document.getElementById(id);
    const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[ch]));
    const fmt = value => Number(value).toLocaleString('da-DK', { minimumFractionDigits:2, maximumFractionDigits:2 });
    const rateFmt = value => Number(value).toLocaleString('da-DK', { minimumFractionDigits:2, maximumFractionDigits:3 });
    const statuses = { draft:'Kladde', submitted:'Afventer', approved:'Godkendt', paid:'Udbetalt', rejected:'Afvist', legacy:'Historik' };
    let data, currentTrip, previewFiles, busy = false;
    let points = [];
    async function api(path, body) {
        const response = await fetch('/mileage' + path, body === undefined ? {} : {
            method:'POST', headers:{ 'Content-Type':'application/json' }, body:JSON.stringify(body)
        });
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || 'Forespørgslen mislykkedes.');
        return result;
    }
    async function perform(action, errorId = 'status') {
        if (busy) return;
        busy = true;
        $(errorId).textContent = '';
        document.querySelectorAll('button').forEach(button => { button.disabled = true; });
        try { await action(); }
        catch (error) {
            const target = $(errorId), dialog = target.closest('dialog');
            const visibleTarget = dialog && !dialog.open ? $('status') : target;
            visibleTarget.textContent = error.message;
            visibleTarget.classList.add('error');
        }
        finally {
            busy = false;
            document.querySelectorAll('button').forEach(button => { button.disabled = false; });
            applyAvailability();
        }
    }
    function applyAvailability() {
        $('newTrip').disabled = !data?.employeeId;
        $('calculateRoute').disabled = !data?.routingConfigured;
        $('tripRoutingNotice').hidden = Boolean(data?.routingConfigured);
        document.querySelectorAll('[data-search]').forEach(button => { button.disabled = !data?.routingConfigured; });
        $('yearPdf').disabled = !$('filterYear').value || !$('filterEmployee').value;
    }
    function options(id, rows, value, label, initial) {
        const previous = $(id).value;
        $(id).innerHTML = initial + rows.map(row => '<option value="'+esc(value(row))+'">'+esc(label(row))+'</option>').join('');
        if ([...$(id).options].some(option => option.value === previous)) $(id).value = previous;
    }
    async function load() {
        data = await api('/data');
        $('identity').textContent = data.user.displayName;
        $('management').hidden = !data.canManage;
        $('setupNotice').hidden = Boolean(data.employeeId);
        $('setupNotice').textContent = 'Din GOH-bruger er ikke tilknyttet en aktiv medarbejder. Bed en kørselsadministrator om at tilknytte dig, før du opretter kørsler.';
        const years = [...new Set(data.trips.map(row => row.date?.slice(0,4)).filter(Boolean).concat(String(new Date().getFullYear())))].sort().reverse();
        options('filterYear', years, x => x, x => x, '<option value="">Alle år / uden dato</option>');
        options('filterEmployee', data.employees, x => x.id, x => x.name, '<option value="">Alle tilgængelige</option>');
        if (!data.canManage && data.employeeId) $('filterEmployee').value = data.employeeId;
        $('addressOptions').innerHTML = data.addresses.map(row => '<option value="'+esc(row.address)+'">'+esc(row.name)+'</option>').join('');
        $('purposeOptions').innerHTML = data.purposes.map(row => '<option value="'+esc(row)+'"></option>').join('');
        if (data.canManage) {
            options('employeeEdit', data.employees, x => x.id, x => x.name, '<option value="">Ny medarbejder</option>');
            options('employeeUser', data.users, x => x.username.toLowerCase(), x => x.name+' ('+x.username+')', '<option value="">Ikke tilknyttet</option>');
            $('rateList').textContent = data.rates.slice().sort((a,b) => b.year-a.year).map(row => row.year+': '+rateFmt(row.high)+' / '+rateFmt(row.low)+' DKK pr. km').join(' · ') || 'Ingen satser oprettet.';
            $('routingStatus').textContent = data.routingConfigured ? 'openrouteservice er konfigureret på serveren.' : 'Automatisk afstand er ikke konfigureret. Brug manuelle km indtil videre.';
        }
        render();
        applyAvailability();
    }
    function render() {
        if (!data) return;
        const year = $('filterYear').value, employee = $('filterEmployee').value, status = $('filterStatus').value;
        const rows = data.trips.filter(row => (!year || row.date?.startsWith(year)) && (!employee || row.employeeId === employee) && (!status || row.status === status))
            .slice().sort((a,b) => (b.date || '').localeCompare(a.date || ''));
        const amount = rows.reduce((sum,row) => sum+(row.allowance?.amount || 0),0);
        $('totals').textContent = rows.length+' kørsler · '+fmt(rows.reduce((sum,row) => sum+row.km,0))+' km · '+fmt(amount)+' DKK dokumenteret. Historiske beløb og afventende kørsler er ikke medregnet i DKK.';
        $('tripRows').innerHTML = rows.map(row => '<tr><td>'+esc(row.date || 'Dato mangler')+'</td><td>'+esc(row.employeeName)+'<small>'+esc(row.purpose)+'</small></td><td><span class="badge '+esc(row.status)+'">'+esc(statuses[row.status])+'</span></td><td class="numeric">'+fmt(row.km)+'</td><td class="numeric">'+(row.allowance ? fmt(row.allowance.amount) : '—')+'</td><td><button type="button" data-detail="'+esc(row.id)+'">Åbn</button></td></tr>').join('') || '<tr><td colspan="6">Ingen kørsler med de valgte filtre.</td></tr>';
        applyAvailability();
    }
    function invalidateDistance() {
        $('tripKm').value = '';
        $('tripBasis').value = '';
    }
    function stopValues() {
        return [...$('stops').querySelectorAll('.stop')].map((row, i) => ({
            address: row.querySelector('[data-address]').value.trim(), note: row.querySelector('[data-note]').value.trim(),
            coordinates: points[i]?.coordinates || null
        }));
    }
    function renderStops() {
        $('stops').innerHTML = points.map((point,index) => '<section class="stop" data-index="'+index+'"><div class="stop-title"><strong>'+esc(index === 0 ? 'Fra' : index === points.length-1 ? 'Til' : 'Mellemstop '+index)+'</strong>'+
            (index > 0 && index < points.length-1 ? '<button type="button" data-remove-stop="'+index+'">Fjern</button>' : '')+'</div><div class="fields"><label>Adresse<input data-address list="addressOptions" value="'+esc(point.address)+'" maxlength="500" required></label><button type="button" data-search="'+index+'"'+(!data.routingConfigured?' disabled':'')+'>Find adresse</button></div><label>Pause / stopnote<input data-note value="'+esc(point.note || '')+'" maxlength="500" placeholder="Fx kundebesøg eller pause"></label><div class="stop-tools" data-results></div></section>').join('');
    }
    function openTrip(trip) {
        $('tripForm').reset();
        $('tripError').textContent = '';
        $('tripId').value = trip?.id || '';
        $('tripTitle').textContent = trip ? 'Rediger kørsel' : 'Ny kørsel';
        $('tripDate').value = trip?.date || new Date().toLocaleDateString('sv-SE');
        $('tripDate').max = new Date().toLocaleDateString('sv-SE');
        $('tripPlate').value = trip?.plate || data.employees.find(row => row.id === data.employeeId)?.plate || '';
        $('tripPurpose').value = trip?.purpose || '';
        $('tripReturn').checked = trip?.roundTrip || false;
        $('tripKm').value = trip?.km || '';
        $('tripBasis').value = trip?.manualReason || '';
        $('tripComment').value = trip?.comment || '';
        points = trip ? trip.stops.map(stop => ({ ...stop })) : [{ address:'',note:'' },{ address:'',note:'' }];
        renderStops();
        $('tripDialog').showModal();
        $('tripDate').focus();
    }
    function detail(id) {
        currentTrip = data.trips.find(row => row.id === id);
        const t = currentTrip;
        if (!t) return;
        $('detailError').textContent = '';
        const own = t.employeeId === data.employeeId;
        const pair = (label,value) => '<dt>'+esc(label)+'</dt><dd>'+esc(value)+'</dd>';
        $('detailBody').innerHTML = '<dl>'+pair('Medarbejder',t.employeeName)+pair('Dato',t.date || 'Mangler')+
            pair('Nummerplade',t.plate)+pair('Status',statuses[t.status])+
            pair('Formål',t.purpose)+pair('Rute / stop',t.stops.map(stop => stop.address+(stop.note?' ('+stop.note+')':'')).join('\n') || 'Se historisk Formål')+
            pair('Retur',t.roundTrip?'Ja':'Nej')+pair('Samlet km',fmt(t.km))+pair('Afstandsgrundlag',t.manualReason)+
            pair('Godtgørelse',t.allowance ? fmt(t.allowance.amount)+' DKK; '+fmt(t.allowance.highKm)+' km × '+rateFmt(t.allowance.highRate)+' + '+fmt(t.allowance.lowKm)+' km × '+rateFmt(t.allowance.lowRate) : t.legacy ? 'Historisk beløb ikke eksporteret; ikke genberegnet.' : t.estimate ? fmt(t.estimate.amount)+' DKK foreløbigt. Endeligt beløb fastlåses ved godkendelse.' : 'Satser eller fuldstændig historik mangler; kan ikke beregnes endnu.')+
            pair('Kommentar',t.comment || '—')+
            (t.legacy ? pair('Historisk betaling',t.legacyPaid?'Markeret udbetalt':'Ikke markeret udbetalt') : '')+
            pair('Hændelser',t.history.map(row => row.at+' · '+row.by+' · '+row.action+(row.note?' · '+row.note:'')).join('\n') || 'Importeret historik')+'</dl>'+
            (data.canManage && t.legacy && !t.date ? '<label>Manglende historisk kørselsdato<input id="legacyDate" type="date"></label><button type="button" data-action="legacy-date">Gem historisk dato</button>' : '')+
            '<label>Begrundelse / kommentar<input id="actionNote" maxlength="1000"></label><div class="actions"><button type="button" data-pdf="'+esc(t.id)+'">Hent PDF</button>'+
            (own && ['draft','rejected'].includes(t.status) ? '<button type="button" data-edit="'+esc(t.id)+'">Rediger</button><button type="button" data-action="submit" class="primary">Send til godkendelse</button>' : '')+
            (data.canManage && !own && t.status === 'submitted' ? '<button type="button" data-action="reject">Afvis med begrundelse</button><button type="button" data-action="approve" class="primary">Godkend og fastlås beløb</button>' : '')+
            (data.canManage && t.status === 'approved' ? '<button type="button" data-action="pay" class="primary">Markér udbetalt</button>' : '')+'</div>';
        if (!$('detailDialog').open) $('detailDialog').showModal();
    }
    async function download(path, filename) {
        const response = await fetch('/mileage'+path);
        if (!response.ok) { const result = await response.json(); throw new Error(result.error || 'PDF kunne ikke dannes.'); }
        const url = URL.createObjectURL(await response.blob());
        const link = document.createElement('a');
        link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove();
        setTimeout(() => URL.revokeObjectURL(url), 30000);
    }
    $('refresh').addEventListener('click', () => perform(load));
    $('newTrip').addEventListener('click', () => openTrip());
    $('closeTrip').addEventListener('click', () => $('tripDialog').close());
    $('closeDetail').addEventListener('click', () => $('detailDialog').close());
    ['filterYear','filterEmployee','filterStatus'].forEach(id => $(id).addEventListener('change',render));
    $('tripRows').addEventListener('click', event => { const button=event.target.closest('[data-detail]'); if(button)detail(button.dataset.detail); });
    $('addStop').addEventListener('click', () => {
        if (points.length >= 12) { $('tripError').textContent='Højst 10 mellemstop.'; return; }
        points = stopValues(); points.splice(points.length-1,0,{address:'',note:''}); renderStops(); invalidateDistance();
    });
    $('tripReturn').addEventListener('change',invalidateDistance);
    $('tripKm').addEventListener('input', () => { $('tripBasis').value=''; });
    $('stops').addEventListener('input', event => {
        if (!event.target.matches('[data-address]')) return;
        const row = event.target.closest('.stop');
        points[Number(row.dataset.index)].coordinates = null;
        row.querySelector('[data-results]').innerHTML = '';
        invalidateDistance();
    });
    $('stops').addEventListener('click', event => {
        const remove=event.target.closest('[data-remove-stop]');
        if (remove) { points=stopValues(); points.splice(Number(remove.dataset.removeStop),1); renderStops(); invalidateDistance(); return; }
        const search=event.target.closest('[data-search]');
        if (!search) return;
        perform(async () => {
            const row = search.closest('.stop'), input = row.querySelector('[data-address]'), query=input.value;
            const result = await api('/addresses/search',{query});
            if (input.value !== query) throw new Error('Adressen blev ændret. Søg igen.');
            const select=document.createElement('select'); select.setAttribute('aria-label','Vælg fundet adresse');
            select.innerHTML='<option value="">Vælg adresse…</option>'+result.addresses.map((item,index)=>'<option value="'+index+'">'+esc(item.address)+'</option>').join('');
            row.querySelector('[data-results]').replaceChildren(select);
            if (!result.addresses.length) throw new Error('Ingen adresser fundet. Præcisér søgningen eller angiv km manuelt.');
            select.addEventListener('change', () => {
                if(select.value==='')return;
                const match=result.addresses[Number(select.value)];
                input.value=match.address; points[Number(row.dataset.index)]={...points[Number(row.dataset.index)],...match}; invalidateDistance();
            });
        },'tripError');
    });
    $('calculateRoute').addEventListener('click', () => perform(async () => {
        const stops=stopValues(), roundTrip=$('tripReturn').checked;
        if (stops.some(stop=>!stop.coordinates)) throw new Error('Find og vælg en adresse for alle stop først.');
        const fingerprint=JSON.stringify(stops);
        const result=await api('/route',{points:stops.map(stop=>stop.coordinates),roundTrip});
        if(fingerprint!==JSON.stringify(stopValues()) || roundTrip!==$('tripReturn').checked) throw new Error('Ruten blev ændret. Beregn igen.');
        $('tripKm').value=result.km;
        $('tripBasis').value='openrouteservice · '+new Date().toISOString()+' · delstrækninger '+result.legs.map(fmt).join(' + ')+' km';
    },'tripError'));
    $('tripForm').addEventListener('submit',event=>{
        event.preventDefault();
        perform(async()=>{
            await api('/trips',{version:data.version,id:$('tripId').value || undefined,date:$('tripDate').value,plate:$('tripPlate').value,
                purpose:$('tripPurpose').value,stops:stopValues(),roundTrip:$('tripReturn').checked,km:Number($('tripKm').value),
                manualReason:$('tripBasis').value,comment:$('tripComment').value});
            $('tripDialog').close(); await load(); $('status').textContent='Kladde gemt i GOH. Åbn kørslen for at sende den til godkendelse.';
        },'tripError');
    });
    $('detailBody').addEventListener('click',event=>{
        const edit=event.target.closest('[data-edit]');
        if(edit){$('detailDialog').close();openTrip(currentTrip);return;}
        const pdf=event.target.closest('[data-pdf]');
        if(pdf){perform(()=>download('/trips/'+encodeURIComponent(pdf.dataset.pdf)+'/document.pdf','Gantech-korsel.pdf'),'detailError');return;}
        const button=event.target.closest('[data-action]');
        if(!button)return;
        perform(async()=>{
            const id=currentTrip.id, action=button.dataset.action;
            await api('/trips/'+encodeURIComponent(id)+(action==='legacy-date'?'/legacy-date':'/action'),
                {version:data.version,action,note:$('actionNote').value,date:$('legacyDate')?.value});
            await load(); detail(id);
        },'detailError');
    });
    $('employeeEdit').addEventListener('change',()=>{
        const row=data.employees.find(item=>item.id===$('employeeEdit').value);
        $('employeeName').value=row?.name || ''; $('employeeUser').value=row?.username || '';
        $('employeePlate').value=row?.plate || ''; $('employeeActive').checked=row?.active ?? true;
    });
    $('employeeForm').addEventListener('submit',event=>{
        event.preventDefault();perform(async()=>{
            await api('/employees',{version:data.version,id:$('employeeEdit').value || undefined,name:$('employeeName').value,
                username:$('employeeUser').value,plate:$('employeePlate').value,active:$('employeeActive').checked});
            await load();$('status').textContent='Medarbejder gemt.';
        });
    });
    $('rateForm').addEventListener('submit',event=>{
        event.preventDefault();perform(async()=>{
            await api('/rates',{version:data.version,year:Number($('rateYear').value),high:Number($('rateHigh').value),low:Number($('rateLow').value)});
            await load();$('status').textContent='Årssatser gemt. Tidligere godkendte beløb er uændrede.';
        });
    });
    $('rateYear').value=new Date().getFullYear();
    $('yearPdf').addEventListener('click',()=>perform(()=>download('/report.pdf?'+new URLSearchParams({year:$('filterYear').value,employeeId:$('filterEmployee').value}),'Gantech-korsel-aarsbilag.pdf')));
    $('importForm').addEventListener('change',()=>{previewFiles=null;$('confirmImport').hidden=true;$('importPreview').textContent='';});
    $('importForm').addEventListener('submit',event=>{
        event.preventDefault();perform(async()=>{
            const files={};
            for(const [key,id] of Object.entries({trips:'importTrips',employees:'importEmployees',addresses:'importAddresses',purposes:'importPurposes'})){
                const file=$(id).files[0];
                if(!file || file.size>2000000)throw new Error('Vælg fire CSV-filer, hver på højst 2 MB.');
                files[key]=await file.text();
            }
            const result=await api('/import',{preview:true,files});
            previewFiles=files;
            $('importPreview').textContent=JSON.stringify(result.counts)+' — '+result.warnings.join(' ');
            $('confirmImport').hidden=false;
        });
    });
    $('confirmImport').addEventListener('click',()=>perform(async()=>{
        if(!previewFiles)throw new Error('Kontrollér CSV-filerne først.');
        const result=await api('/import',{version:data.version,files:previewFiles});
        previewFiles=null;$('confirmImport').hidden=true;await load();
        $('status').textContent='Historikken er gemt i GOH. '+result.warnings.join(' ');
    }));
    perform(load);
}());
