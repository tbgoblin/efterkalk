(function () {
    'use strict';
    const $ = id => document.getElementById(id);
    const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[ch]));
    const fmt = value => Number(value).toLocaleString('da-DK', { minimumFractionDigits:2, maximumFractionDigits:2 });
    const rateFmt = value => Number(value).toLocaleString('da-DK', { minimumFractionDigits:2, maximumFractionDigits:3 });
    const statuses = { draft:'Kladde', submitted:'Afventer', approved:'Godkendt', paid:'Udbetalt', rejected:'Afvist', legacy:'Historik', deleted:'Slettet' };
    let data, currentTrip, previewFiles, addressToSave, vismaActors = [], busy = false, dragging = null;
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
        $(errorId).classList.remove('error');
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
        $('newTrip').disabled = !data?.participants?.length;
        $('calculateRoute').disabled = !data?.routingConfigured;
        $('tripRoutingNotice').hidden = Boolean(data?.routingConfigured);
        document.querySelectorAll('[data-search]').forEach(button => { button.disabled = !data?.routingConfigured; });
        if ($('deletePurpose')) $('deletePurpose').disabled = !$('purposeEdit').value;
        if ($('deleteAddress')) $('deleteAddress').disabled = !$('addressEdit').value;
        $('yearPdf').disabled = !$('filterYear').value || !$('filterEmployee').value;
    }
    function options(id, rows, value, label, initial) {
        const previous = $(id).value;
        $(id).innerHTML = initial + rows.map(row => '<option value="'+esc(value(row))+'">'+esc(label(row))+'</option>').join('');
        if ([...$(id).options].some(option => option.value === previous)) $(id).value = previous;
    }
    async function load() {
        let syncError = '', actorError = '';
        const [employeeSync, actorLookup] = await Promise.allSettled([
            api('/employees/sync', {}), api('/visma-actors')
        ]);
        if (employeeSync.status === 'fulfilled') syncError = (employeeSync.value.warnings || []).join(' ');
        else syncError = employeeSync.reason.message;
        if (actorLookup.status === 'fulfilled' && Array.isArray(actorLookup.value.actors)) {
            vismaActors = actorLookup.value.actors;
        } else {
            vismaActors = [];
            actorError = actorLookup.status === 'rejected' ? actorLookup.reason.message : 'Visma returnerede en ugyldig adressebog.';
        }
        data = await api('/data');
        $('syncNotice').hidden = !syncError;
        $('syncNotice').textContent = syncError ? 'Visma: '+syncError+' Medarbejderlisten kan være ufuldstændig eller forældet.' : '';
        $('vismaAddressNotice').hidden = !actorError;
        $('vismaAddressNotice').textContent = actorError ? 'Visma-adresser: '+actorError : '';
        $('identity').textContent = data.user.displayName;
        $('management').hidden = !data.canManage;
        $('setupNotice').hidden = Boolean(data.employeeId);
        $('setupNotice').textContent = 'Din GOH-bruger er ikke tilknyttet en aktiv medarbejder. Du kan oprette kørsel for en medarbejder, men skal tilknyttes for at se dine egne kørsler.';
        const years = [...new Set(data.trips.map(row => row.date?.slice(0,4)).filter(Boolean).concat(String(new Date().getFullYear())))].sort().reverse();
        options('filterYear', years, x => x, x => x, '<option value="">Alle år / uden dato</option>');
        const visibleEmployees = new Map(data.employees.map(row => [row.id, row.name]));
        data.trips.forEach(row => visibleEmployees.set(row.employeeId, row.employeeName));
        options('filterEmployee', [...visibleEmployees].map(([id,name]) => ({id,name})), x => x.id, x => x.name, '<option value="">Alle tilgængelige</option>');
        if (data.canManage && !$('filterStatus').querySelector('option[value="deleted"]')) $('filterStatus').add(new Option('Slettet', 'deleted'));
        $('purposeOptions').innerHTML = data.purposes.map(row => '<option value="'+esc(row)+'"></option>').join('');
        $('usersAdminLink').hidden = !data.isSuperadmin;
        if (data.canManage) {
            options('employeeEdit', data.employees, x => x.id, x => x.name+(x.vismaEmpNo ? ' (Visma '+x.vismaEmpNo+')' : '')+(!x.active ? ' · inaktiv' : ''), '<option value="">Ny medarbejder</option>');
            options('employeeUser', data.users, x => x.username.toLowerCase(), x => x.name+' ('+x.username+')', '<option value="">Ikke tilknyttet</option>');
            $('purposeEdit').innerHTML = '<option value="">Nyt formål</option>'+data.purposes.map((row,index) => '<option value="'+index+'">'+esc(row)+'</option>').join('');
            $('addressEdit').innerHTML = '<option value="">Ny adresse</option>'+data.addresses.map((row,index) =>
                '<option value="'+index+'">'+esc(row.name || 'Uden navn')+' — '+esc(row.address)+'</option>').join('');
            $('rateList').textContent = data.rates.slice().sort((a,b) => b.year-a.year).map(row => row.year+': '+rateFmt(row.high)+' / '+rateFmt(row.low)+' DKK pr. km').join(' · ') || 'Ingen satser oprettet.';
            $('routingStatus').textContent = data.routingConfigured
                ? 'En API-nøgle er gemt på serveren. Gyldigheden kontrolleres ved adresseopslag.'
                : 'Automatisk afstand er ikke konfigureret. Brug manuelle km indtil videre.';
            $('employeeEdit').dispatchEvent(new Event('change'));
            $('purposeEdit').dispatchEvent(new Event('change'));
            $('addressEdit').dispatchEvent(new Event('change'));
        }
        render();
        applyAvailability();
    }
    function render() {
        if (!data) return;
        const year = $('filterYear').value, employee = $('filterEmployee').value, status = $('filterStatus').value;
        const showHistory = $('filterHistory').checked;
        const rows = data.trips.filter(row => (status === 'deleted' ? row.status === 'deleted' : row.status !== 'deleted')
            && (!year || row.date?.startsWith(year)) && (!employee || row.employeeId === employee)
            && (!status || row.status === status) && (showHistory || row.status !== 'legacy'))
            .slice().sort((a,b) => (b.date || '').localeCompare(a.date || ''));
        const activeRows = rows.filter(row => !['deleted', 'legacy'].includes(row.status));
        const archiveCount = rows.filter(row => row.status === 'legacy').length;
        const amount = activeRows.reduce((sum,row) => sum+(row.allowance?.amount || 0),0);
        const reviews = data.trips.filter(row => row.reconciliationRequired && !['deleted', 'legacy'].includes(row.status)).length;
        $('totals').textContent = activeRows.length+' kørsler · '+fmt(activeRows.reduce((sum,row) => sum+row.km,0))+' km · '+fmt(amount)+' DKK dokumenteret. Historiske beløb og afventende kørsler er ikke medregnet i DKK.'
            +(archiveCount ? ' '+archiveCount+' historiske poster vises kun som arkiv og påvirker ikke satser eller godtgørelse.' : '')
            +(reviews ? ' '+reviews+' godkendte kørsler skal gennemgås: tidligere km er ændret; beløb er ikke genberegnet.' : '');
        $('tripRows').innerHTML = rows.map(row => '<tr><td>'+esc(row.date || 'Dato mangler')+'</td><td>'+esc(row.employeeName)+'<small>'+esc(row.purpose)+'</small></td><td><span class="badge '+esc(row.status)+'">'+esc(statuses[row.status])+'</span>'+(row.reconciliationRequired && row.status !== 'deleted' ? '<small>Kontrollér beregning</small>' : '')+'</td><td class="numeric">'+fmt(row.km)+'</td><td class="numeric">'+(row.status !== 'deleted' && row.allowance ? fmt(row.allowance.amount) : '—')+'</td><td><button type="button" data-detail="'+esc(row.id)+'">Åbn</button></td></tr>').join('') || '<tr><td colspan="6">Ingen kørsler med de valgte filtre.</td></tr>';
        applyAvailability();
    }
    function invalidateDistance() {
        $('tripKm').value = '';
        $('tripBasis').value = '';
    }
    function stopValues() {
        return [...$('stops').querySelectorAll('.stop')].map((row, i) => ({
            address: row.querySelector('[data-address]').value.trim(), note: row.querySelector('[data-note]').value.trim(),
            coordinates: points[i]?.coordinates || null, savedAddress: points[i]?.savedAddress || null, name: points[i]?.name || ''
        }));
    }
    function renderStops() {
        $('stops').innerHTML = points.map((point,index) => '<section class="stop" data-index="'+index+'"><div class="stop-title"><strong>'+esc(index === 0 ? 'Fra' : index === points.length-1 ? 'Til' : 'Mellemstop '+index)+'</strong>'+
            (index > 0 && index < points.length-1 ? '<div class="stop-actions">'+
                (points.length > 3 ? '<button type="button" class="drag-stop" data-drag-stop aria-label="Flyt mellemstop '+index+'" aria-describedby="reorderHelp" aria-keyshortcuts="ArrowUp ArrowDown"><svg aria-hidden="true" width="16" height="16" viewBox="0 0 16 16" fill="currentColor"><circle cx="5" cy="3" r="1.3"/><circle cx="11" cy="3" r="1.3"/><circle cx="5" cy="8" r="1.3"/><circle cx="11" cy="8" r="1.3"/><circle cx="5" cy="13" r="1.3"/><circle cx="11" cy="13" r="1.3"/></svg>Træk for at flytte</button>' : '')+
                '<button type="button" data-remove-stop="'+index+'">Fjern</button></div>' : '')+'</div><div class="fields"><label>Adresse, hotel eller virksomhedsnavn<input data-address autocomplete="off" value="'+esc(point.address)+'" maxlength="500" required></label><button type="button" data-search="'+index+'"'+(!data.routingConfigured?' disabled':'')+'>Find adresse online</button></div><div class="saved-addresses" data-saved-results aria-live="polite"></div>'+(data.canManage?'<div class="stop-tools"><button type="button" data-save-address>Gem fælles adresse</button></div>':'')+'<label>Pause / stopnote<input data-note value="'+esc(point.note || '')+'" maxlength="500" placeholder="Fx kundebesøg eller pause"></label><div class="stop-tools saved-addresses" data-results aria-live="polite"></div></section>').join('');
    }
    function moveStop(index, target) {
        if (target < 1 || target >= points.length - 1 || target === index) return;
        const oldRows=[...$('stops').querySelectorAll('.stop')];
        const oldTops=oldRows.map(row=>row.getBoundingClientRect().top);
        const order=oldRows.map((_,i)=>i);
        order.splice(target,0,order.splice(index,1)[0]);
        points=stopValues();
        const [moved]=points.splice(index,1);
        points.splice(target,0,moved);
        renderStops(); invalidateDistance();
        $('tripNotice').textContent='Mellemstop flyttet til plads '+target+'. Beregn eller angiv afstanden igen.';
        $('stops').querySelector('.stop[data-index="'+target+'"] [data-drag-stop]').focus();
        if (!window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
            $('stops').querySelectorAll('.stop').forEach((row,i)=>{
                const delta=oldTops[order[i]]-row.getBoundingClientRect().top;
                if (delta && Math.abs(delta)<1500) row.animate([{transform:'translateY('+delta+'px)'},{transform:'translateY(0)'}],{duration:220,easing:'ease-out'});
            });
        }
    }
    function clearDrag() {
        if (dragging?.frame) cancelAnimationFrame(dragging.frame);
        dragging?.ghost?.remove();
        $('stops').querySelectorAll('.dragging, .drop-target').forEach(row => row.classList.remove('dragging','drop-target'));
        dragging=null;
    }
    function updateDragTarget(y) {
        const rows=[...$('stops').querySelectorAll('.stop')].slice(1,-1);
        const nearest=rows.reduce((best,row,index) => {
            const rect=row.getBoundingClientRect(), distance=Math.abs(y-(rect.top+rect.height/2));
            return distance < best.distance ? {index:index+1,distance} : best;
        },{index:dragging.index,distance:Infinity});
        if (dragging.target !== nearest.index) {
            dragging.target=nearest.index;
            $('tripNotice').textContent=dragging.target === dragging.index ? 'Træk til en anden plads.' : 'Slip for at flytte mellemstoppet til plads '+dragging.target+'.';
        }
        rows.forEach((row,index) => {
            row.classList.toggle('dragging',index+1 === dragging.index);
            row.classList.toggle('drop-target',index+1 === dragging.target && dragging.target !== dragging.index);
        });
    }
    function scrollDuringDrag() {
        if (!dragging?.active) return;
        const dialog=$('tripDialog'), rect=dialog.getBoundingClientRect();
        const top=Math.max(0,rect.top), bottom=Math.min(window.innerHeight,rect.bottom);
        const step=dragging.y < top+45 ? -12 : dragging.y > bottom-45 ? 12 : 0;
        if (step) { dialog.scrollTop+=step; updateDragTarget(dragging.y); }
        dragging.frame=requestAnimationFrame(scrollDuringDrag);
    }
    function vismaActorAddress(actor) {
        const postalCity=[actor.PNo,actor.PArea].filter(Boolean).join(' ');
        return [actor.Ad1,actor.Ad2,postalCity].filter(Boolean).join(', ');
    }
    function suggestAddresses(row) {
        const query = row.querySelector('[data-address]').value.trim().normalize('NFC').toLocaleLowerCase('da-DK');
        const matches = query.length < 2 ? [] : data.addresses.map((item, index) => ({ item, index }))
            .filter(({ item }) => [item.name, item.address].some(value => String(value || '').normalize('NFC').toLocaleLowerCase('da-DK').includes(query))).slice(0, 10);
        const actors = query.length < 2 ? [] : vismaActors.map((item, index) => ({ item, index }))
            .filter(({ item }) => [item.CustNo,item.SupNo,item.Nm,item.Ad1,item.Ad2,item.PNo,item.PArea]
                .some(value => String(value || '').normalize('NFC').toLocaleLowerCase('da-DK').includes(query))).slice(0, 10);
        row.querySelector('[data-saved-results]').innerHTML = matches.map(({ item, index }) =>
            '<button type="button" data-saved-address="'+index+'">'+esc(item.name || item.address)+'<small>'+esc(item.name ? item.address : '')+'</small></button>').join('')
            +actors.map(({ item, index }) => {
                const address=vismaActorAddress(item), kind=item.ActorType === 'supplier' ? 'Leverandør' : 'Kunde';
                const number=item.ActorType === 'supplier' ? item.SupNo : item.CustNo;
                return '<button type="button" data-visma-actor="'+index+'">'+esc(item.Nm || kind+' '+number)
                    +'<small>Visma · '+kind+' '+esc(number)+(address ? ' · '+esc(address) : '')+'</small></button>';
            }).join('');
    }
    function openTrip(trip) {
        $('tripForm').reset();
        $('tripError').textContent = '';
        $('tripNotice').textContent = '';
        $('tripId').value = trip?.id || '';
        $('tripTitle').textContent = trip ? 'Rediger kørsel' : 'Ny kørsel';
        options('tripEmployee', data.participants, x => x.id, x => x.name+(x.vismaEmpNo ? ' ('+x.vismaEmpNo+')' : ''), '<option value="">Vælg medarbejder</option>');
        if (trip && ![...$('tripEmployee').options].some(option => option.value === trip.employeeId)) {
            $('tripEmployee').add(new Option(trip.employeeName, trip.employeeId));
        }
        $('tripEmployee').value = trip?.employeeId || data.employeeId || '';
        $('tripEmployee').disabled = Boolean(trip);
        $('tripDate').value = trip?.date || new Date().toLocaleDateString('sv-SE');
        $('tripDate').max = new Date().toLocaleDateString('sv-SE');
        $('tripPlate').value = trip?.plate || data.employees.find(row => row.id === $('tripEmployee').value)?.plate || '';
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
        const creator = t.createdBy === data.user.username;
        const pair = (label,value) => '<dt>'+esc(label)+'</dt><dd>'+esc(value)+'</dd>';
        $('detailBody').innerHTML = '<dl>'+pair('Medarbejder',t.employeeName)+(t.createdBy ? pair('Oprettet af',t.createdBy) : '')+pair('Dato',t.date || 'Mangler')+
            pair('Nummerplade',t.plate)+pair('Status',statuses[t.status])+
            pair('Formål',t.purpose)+pair('Rute / stop',t.stops.map(stop => stop.address+(stop.note?' ('+stop.note+')':'')).join('\n') || 'Se historisk Formål')+
            pair('Retur',t.roundTrip?'Ja':'Nej')+pair('Samlet km',fmt(t.km))+pair('Afstandsgrundlag',t.manualReason)+
            pair('Godtgørelse',t.allowance ? fmt(t.allowance.amount)+' DKK; '+fmt(t.allowance.highKm)+' km × '+rateFmt(t.allowance.highRate)+' + '+fmt(t.allowance.lowKm)+' km × '+rateFmt(t.allowance.lowRate) : t.legacy ? 'Arkivpost: historisk beløb er ikke eksporteret eller genberegnet.' : t.estimate ? fmt(t.estimate.amount)+' DKK foreløbigt. Endeligt beløb fastlåses ved godkendelse.' : 'Satser mangler for kørselsåret; beløbet kan ikke beregnes endnu.')+
            pair('Kommentar',t.comment || '—')+
            (t.reconciliationRequired && t.status !== 'deleted' ? pair('Kontrol kræves','Tidligere km er ændret efter godkendelsen. Kontrollér beregningen; beløbet er ikke genberegnet.') : '')+
            (t.status === 'deleted' ? pair('Før sletning',statuses[t.deletedPreviousStatus] || t.deletedPreviousStatus) : '')+
            (t.legacy ? pair('Historisk betaling',t.legacyPaid?'Markeret udbetalt':'Ikke markeret udbetalt') : '')+
            pair('Hændelser',t.history.map(row => row.at+' · '+row.by+' · '+row.action+(row.note?' · '+row.note:'')).join('\n') || 'Importeret historik')+'</dl>'+
            (data.canManage && t.status === 'legacy' && !t.date ? '<label>Manglende historisk kørselsdato<input id="legacyDate" type="date"></label><button type="button" data-action="legacy-date">Gem historisk dato</button>' : '')+
            '<label>Begrundelse / kommentar<input id="actionNote" maxlength="1000"></label><div class="actions">'+
            (t.status !== 'deleted' ? '<button type="button" data-pdf="'+esc(t.id)+'">Hent PDF</button>' : '')+
            ((own || creator) && ['draft','rejected'].includes(t.status) ? '<button type="button" data-edit="'+esc(t.id)+'">Rediger</button><button type="button" data-action="submit" class="primary">Send til godkendelse</button>' : '')+
            (data.canManage && !own && !creator && t.status === 'submitted' ? '<button type="button" data-action="reject">Afvis med begrundelse</button>' : '')+
            (data.canManage && t.status === 'submitted' ? '<button type="button" data-action="approve" class="primary">Godkend og fastlås beløb</button>' : '')+
            (data.canManage && t.status === 'approved' ? '<button type="button" data-action="pay" class="primary">Markér udbetalt</button>' : '')+
            (data.canManage && t.status !== 'deleted' ? '<button type="button" data-action="delete">Slet kørsel</button>' : '')+'</div>';
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
    $('tripDialog').addEventListener('close', clearDrag);
    $('closeAddress').addEventListener('click', () => $('addressDialog').close());
    $('closeDetail').addEventListener('click', () => $('detailDialog').close());
    $('cancelDelete').addEventListener('click', () => $('deleteDialog').close());
    ['filterYear','filterEmployee','filterStatus','filterHistory'].forEach(id => $(id).addEventListener('change',render));
    $('tripRows').addEventListener('click', event => { const button=event.target.closest('[data-detail]'); if(button)detail(button.dataset.detail); });
    $('addStop').addEventListener('click', () => {
        if (points.length >= 12) { $('tripError').textContent='Højst 10 mellemstop.'; return; }
        points = stopValues(); points.splice(points.length-1,0,{address:'',note:''}); renderStops(); invalidateDistance();
    });
    $('tripReturn').addEventListener('change',invalidateDistance);
    $('tripEmployee').addEventListener('change', () => {
        $('tripPlate').value = data.employees.find(row => row.id === $('tripEmployee').value)?.plate || '';
    });
    $('tripKm').addEventListener('input', () => { $('tripBasis').value=''; });
    $('stops').addEventListener('input', event => {
        if (!event.target.matches('[data-address]')) return;
        const row = event.target.closest('.stop');
        points[Number(row.dataset.index)].coordinates = null;
        points[Number(row.dataset.index)].savedAddress = null;
        points[Number(row.dataset.index)].name = '';
        row.querySelector('[data-results]').innerHTML = '';
        suggestAddresses(row);
        invalidateDistance();
    });
    $('stops').addEventListener('keydown', event => {
        const handle=event.target.closest('[data-drag-stop]');
        if (!handle || !['ArrowUp','ArrowDown'].includes(event.key)) return;
        event.preventDefault();
        const index=Number(handle.closest('.stop').dataset.index);
        moveStop(index,index+(event.key === 'ArrowUp' ? -1 : 1));
    });
    $('stops').addEventListener('pointerdown', event => {
        const handle=event.target.closest('[data-drag-stop]');
        if (!handle || !event.isPrimary || event.button !== 0 || dragging) return;
        dragging={pointerId:event.pointerId,index:Number(handle.closest('.stop').dataset.index),x:event.clientX,y:event.clientY,active:false};
        handle.setPointerCapture(event.pointerId);
    });
    $('stops').addEventListener('pointermove', event => {
        if (!dragging || event.pointerId !== dragging.pointerId) return;
        if (!dragging.active && Math.hypot(event.clientX-dragging.x,event.clientY-dragging.y) < 8) return;
        if (!dragging.active) {
            const row=$('stops').querySelector('.stop[data-index="'+dragging.index+'"]');
            dragging.ghost=document.createElement('div');
            dragging.ghost.className='drag-preview';
            dragging.ghost.setAttribute('aria-hidden','true');
            dragging.ghost.textContent='Flytter: '+(row.querySelector('[data-address]').value.trim() || 'Mellemstop '+dragging.index);
            $('tripDialog').appendChild(dragging.ghost);
        }
        dragging.active=true;
        dragging.y=event.clientY;
        const x=Math.max(0,Math.min(event.clientX+12,window.innerWidth-220));
        dragging.ghost.style.transform='translate('+x+'px,'+Math.max(0,event.clientY-30)+'px)';
        updateDragTarget(dragging.y);
        if (!dragging.frame) dragging.frame=requestAnimationFrame(scrollDuringDrag);
    });
    $('stops').addEventListener('pointerup', event => {
        if (!dragging || event.pointerId !== dragging.pointerId) return;
        const {index,target,active}=dragging;
        const handle=event.target.closest('[data-drag-stop]');
        if (active && handle) {
            handle.dataset.dragged='true';
            setTimeout(() => { delete handle.dataset.dragged; },0);
        }
        clearDrag();
        if (active && target !== undefined && target !== index) moveStop(index,target);
        else if (active) $('tripNotice').textContent='Mellemstoppet blev ikke flyttet.';
    });
    $('stops').addEventListener('pointercancel', event => {
        if (dragging && event.pointerId === dragging.pointerId) {
            clearDrag();
            $('tripNotice').textContent='Flytning annulleret.';
        }
    });
    $('stops').addEventListener('click', event => {
        const remove=event.target.closest('[data-remove-stop]');
        if (remove) { points=stopValues(); points.splice(Number(remove.dataset.removeStop),1); renderStops(); invalidateDistance(); return; }
        const handle=event.target.closest('[data-drag-stop]');
        if (handle) {
            if (handle.dataset.dragged) return;
            const row=handle.closest('.stop'), index=Number(row.dataset.index), actions=handle.parentElement;
            let select=actions.querySelector('[data-move-to]');
            if (!select) {
                select=document.createElement('select');
                select.dataset.moveTo='';
                select.setAttribute('aria-label','Flyt mellemstop '+index+' til plads');
                select.innerHTML=points.slice(1,-1).map((_,i) => '<option value="'+(i+1)+'">Plads '+(i+1)+'</option>').join('');
                select.value=String(index);
                select.addEventListener('change',()=>moveStop(index,Number(select.value)));
                actions.appendChild(select);
            }
            select.focus();
            return;
        }
        const vismaActor=event.target.closest('[data-visma-actor]');
        if (vismaActor) {
            const row=vismaActor.closest('.stop'), item=vismaActors[Number(vismaActor.dataset.vismaActor)];
            const kind=item.ActorType === 'supplier' ? 'Leverandør' : 'Kunde';
            const number=item.ActorType === 'supplier' ? item.SupNo : item.CustNo;
            const index=Number(row.dataset.index), address=vismaActorAddress(item) || item.Nm || kind+' '+number;
            row.querySelector('[data-address]').value=address;
            points[index]={...points[index],name:item.Nm || kind+' '+number,address,coordinates:null,savedAddress:null};
            row.querySelector('[data-saved-results]').innerHTML='';
            row.querySelector('[data-results]').textContent='Visma '+kind.toLowerCase()+' valgt. Vælg et bekræftet resultat fra Find adresse online til ruteberegning, eller angiv km manuelt.';
            invalidateDistance();
            return;
        }
        const saved=event.target.closest('[data-saved-address]');
        if (saved) {
            const row=saved.closest('.stop'), item=data.addresses[Number(saved.dataset.savedAddress)], index=Number(row.dataset.index);
            row.querySelector('[data-address]').value=item.address;
            points[index]={...points[index],name:item.name,address:item.address,coordinates:item.coordinates || null,savedAddress:{name:item.name,address:item.address}};
            row.querySelector('[data-saved-results]').innerHTML='';
            row.querySelector('[data-results]').textContent=item.coordinates ? 'Gemte koordinater er klar til ruteberegning.' : 'Find og bekræft adressen online for at gemme koordinater.';
            invalidateDistance();
            return;
        }
        const saveAddress=event.target.closest('[data-save-address]');
        if (saveAddress) {
            const row=saveAddress.closest('.stop'), index=Number(row.dataset.index), address=row.querySelector('[data-address]').value.trim();
            if (!address) { $('tripError').textContent='Skriv en adresse før den gemmes.'; return; }
            addressToSave={address,coordinates:points[index].coordinates || null};
            $('addressName').value=points[index].savedAddress?.name || points[index].name || '';
            $('addressValue').value=address;
            $('addressCoordinates').textContent=addressToSave.coordinates ? 'Bekræftede koordinater gemmes med adressen.' : 'Adressen gemmes uden koordinater; km kan angives manuelt.';
            $('addressError').textContent='';
            $('addressDialog').showModal(); $('addressName').focus();
            return;
        }
        const search=event.target.closest('[data-search]');
        if (!search) return;
        perform(async () => {
            const row = search.closest('.stop'), input = row.querySelector('[data-address]'), query=input.value;
            const result = await api('/addresses/search',{query});
            if (input.value !== query) throw new Error('Adressen blev ændret. Søg igen.');
            if (!result.addresses.length) throw new Error('Ingen adresser fundet. Præcisér søgningen eller angiv km manuelt.');
            const results=row.querySelector('[data-results]');
            results.innerHTML='<strong>Vælg det rigtige resultat:</strong>'+result.addresses.map((item,index)=>
                '<button type="button" data-online-result="'+index+'">'+esc(item.address)+
                (!item.streetAddressKnown ? '<small>Gade og husnummer er ikke særskilt oplyst af søgetjenesten.</small>' : '')+'</button>').join('');
            results.querySelectorAll('[data-online-result]').forEach(button => button.addEventListener('click', () => {
                const match=result.addresses[Number(button.dataset.onlineResult)];
                const index=Number(row.dataset.index), original=points[index].savedAddress;
                input.value=match.address; points[index]={...points[index],...match}; row.querySelector('[data-saved-results]').innerHTML=''; results.textContent='Adresse valgt.'; invalidateDistance();
                if (original?.name) perform(async () => {
                    await api('/addresses',{version:data.version,...original,coordinates:match.coordinates});
                    await load();
                    row.querySelector('[data-results]').textContent='Adressen er bekræftet og koordinaterne gemt til næste gang.';
                },'tripError');
            }));
        },'tripError');
    });
    $('addressForm').addEventListener('submit',event=>{
        event.preventDefault();
        perform(async()=>{
            await api('/addresses',{version:data.version,name:$('addressName').value,...addressToSave});
            $('addressDialog').close(); await load();
            $('tripError').textContent='';
            $('tripNotice').textContent='Adressen er gemt i den fælles adressebog.';
        },'addressError');
    });
    $('deleteReason').addEventListener('input',()=> $('deleteReason').setCustomValidity(''));
    $('deleteForm').addEventListener('submit',event=>{
        event.preventDefault();
        const note=$('deleteReason').value.trim();
        if(!note){
            $('deleteReason').setCustomValidity('Angiv en begrundelse før sletning.');
            $('deleteReason').reportValidity();
            $('deleteReason').focus();
            return;
        }
        perform(async()=>{
            await api('/trips/'+encodeURIComponent(currentTrip.id)+'/delete',{version:data.version,note});
            $('deleteDialog').close();
            $('detailDialog').close();
            await load();
            $('status').textContent='Kørslen er slettet. Tidligere godkendte beløb er uændrede; berørte kørsler er markeret til kontrol.';
        },'deleteError');
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
            await api('/trips',{version:data.version,id:$('tripId').value || undefined,employeeId:$('tripEmployee').value,date:$('tripDate').value,plate:$('tripPlate').value,
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
        if (button.dataset.action === 'delete') {
            $('deleteError').textContent='';
            $('deleteReason').value='';
            $('deleteDialog').showModal();
            $('deleteReason').focus();
            return;
        }
        perform(async()=>{
            const id=currentTrip.id, action=button.dataset.action;
            await api('/trips/'+encodeURIComponent(id)+(action==='legacy-date'?'/legacy-date':action==='delete'?'/delete':'/action'),
                {version:data.version,action,note:$('actionNote').value,date:$('legacyDate')?.value});
            await load();
            detail(id);
        },'detailError');
    });
    $('employeeEdit').addEventListener('change',()=>{
        const row=data.employees.find(item=>item.id===$('employeeEdit').value);
        $('employeeName').value=row?.name || ''; $('employeeUser').value=row?.username || '';
        $('employeePlate').value=row?.plate || ''; $('employeeActive').checked=row?.active ?? true;
        $('employeeName').disabled=Boolean(row?.vismaEmpNo); $('employeeActive').disabled=Boolean(row?.vismaEmpNo);
    });
    $('employeeForm').addEventListener('submit',event=>{
        event.preventDefault();perform(async()=>{
            await api('/employees',{version:data.version,id:$('employeeEdit').value || undefined,name:$('employeeName').value,
                username:$('employeeUser').value,plate:$('employeePlate').value,active:$('employeeActive').checked});
            await load();$('status').textContent='Medarbejder gemt.';
        });
    });
    $('purposeEdit').addEventListener('change',()=>{
        const selected=$('purposeEdit').value;
        const row=selected === '' ? null : data.purposes[Number(selected)];
        $('purposeOriginal').value=row || '';
        $('purposeValue').value=row || '';
        $('deletePurpose').disabled=!row;
    });
    $('purposeForm').addEventListener('submit',event=>{
        event.preventDefault();perform(async()=>{
            await api('/purposes',{version:data.version,purpose:$('purposeValue').value,original:$('purposeOriginal').value || undefined});
            await load();$('purposeStatus').textContent='Formålet er gemt.';
        },'purposeStatus');
    });
    $('deletePurpose').addEventListener('click',()=>{
        const purpose=$('purposeOriginal').value;
        if(!purpose || !confirm('Fjern formålet "'+purpose+'" fra listen? Historiske kørsler ændres ikke.'))return;
        perform(async()=>{
            await api('/purposes/delete',{version:data.version,purpose});
            await load();$('purposeStatus').textContent='Formålet er fjernet. Historiske kørsler er uændrede.';
        },'purposeStatus');
    });
    $('addressEdit').addEventListener('change',()=>{
        const selected=$('addressEdit').value;
        const row=selected === '' ? null : data.addresses[Number(selected)];
        $('addressOriginalName').value=row?.name || '';
        $('addressOriginalValue').value=row?.address || '';
        $('sharedAddressName').value=row?.name || '';
        $('sharedAddressValue').value=row?.address || '';
        $('sharedAddressCoordinates').textContent=row?.coordinates
            ? 'Bekræftede koordinater er gemt.' : row ? 'Ingen koordinater gemt; online opslag kræves til automatisk rute.' : '';
        $('deleteAddress').disabled=!row;
    });
    $('sharedAddressForm').addEventListener('submit',event=>{
        event.preventDefault();perform(async()=>{
            const originalName=$('addressOriginalName').value, originalAddress=$('addressOriginalValue').value;
            const editing=$('addressEdit').value !== '';
            await api('/addresses',{version:data.version,name:$('sharedAddressName').value,address:$('sharedAddressValue').value,
                originalName:editing?originalName:undefined,originalAddress:editing?originalAddress:undefined});
            await load();$('sharedAddressStatus').textContent='Den fælles adresse er gemt.';
        },'sharedAddressStatus');
    });
    $('deleteAddress').addEventListener('click',()=>{
        const name=$('addressOriginalName').value,address=$('addressOriginalValue').value;
        if($('addressEdit').value === '' || !address || !confirm('Fjern "'+(name || 'Adresse uden navn')+' — '+address+'" fra den fælles adressebog?'))return;
        perform(async()=>{
            await api('/addresses/delete',{version:data.version,name,address});
            await load();$('sharedAddressStatus').textContent='Adressen er fjernet fra den fælles adressebog.';
        },'sharedAddressStatus');
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
