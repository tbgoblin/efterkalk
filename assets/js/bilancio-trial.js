(() => {
    const el = id => document.getElementById(id);
    const months = ['Juli','August','September','Oktober','November','December','Januar','Februar','Marts','April','Maj','Juni'];
    const esc = v => String(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', { timeZone:'Europe/Copenhagen',year:'numeric',month:'numeric' }).formatToParts(new Date()).map(p=>[p.type,p.value]));
    // Start with the most recently completed month.
    const previous = new Date(Date.UTC(Number(parts.year), Number(parts.month)-2, 1));
    const m = previous.getUTCMonth();
    el('year').value = previous.getUTCFullYear() - (m < 6 ? 1 : 0);
    months.forEach((name,i)=>el('period').add(new Option(`${i+1} · ${name}`,i+1)));
    el('period').value = (m+6)%12+1;
    function printReport(orientation) {
        const styleEl = el('printOrientation');
        if (styleEl) styleEl.textContent = `@page{size:A4 ${orientation};margin:10mm}`;
        window.print();
    }
    el('printPortraitBtn').addEventListener('click', () => printReport('portrait'));
    el('printLandscapeBtn').addEventListener('click', () => printReport('landscape'));
    let report = null;
    function render() {
        if (!report) return;
        const divisor = Number(el('unit').value), unit = divisor===1000?'t.kr':'DKK';
        const fmt = v => v===null?'—':new Intl.NumberFormat('da-DK',{minimumFractionDigits:divisor===1?2:0,maximumFractionDigits:divisor===1?2:0}).format(Object.is(v,-0)?0:v);
        const pct = v => v===null?'—':new Intl.NumberFormat('da-DK',{maximumFractionDigits:0}).format(v)+'%';
        const y = report.year, p = report.period;
        const calendarYear = y+(p>6?1:0);
        const abbr = i => months[p-1].slice(0,3).toLowerCase()+'-'+String(i).slice(-2);
        const labels = [abbr(calendarYear), abbr(calendarYear-1), `Juli ${y} - ${months[p-1].toLowerCase()} ${calendarYear}`, `Juli ${y-1} - ${months[p-1].toLowerCase()} ${calendarYear-1}`];
        el('pageTitle').textContent = `Gantech A/S – Økonomirapport ${months[p-1].toLowerCase()} ${calendarYear}`;
        const revenueAmounts = report.rows[0].amounts;
        function buildTable(title, colIndices, showLabels, extraRows) {
            const cells = (amounts, percentages) => colIndices.map((ci,li)=>`<td class="${li===0?'current ':''}${amounts[ci]<0?'negative':''}">${fmt(amounts[ci]/divisor)}</td><td class="${li===0?'current':''}">${pct(percentages[ci])}</td>`).join('');
            const head1 = colIndices.map(ci=>`<th colspan="2">Regnskabsåret ${y-ci%2}-${String(y+1-ci%2).slice(-2)}<br>${esc(labels[ci])}</th>`).join('');
            const head2 = colIndices.map(()=>`<th>${unit}</th><th>%</th>`).join('');
            const labelHead1 = showLabels ? '<th class="label-col">Kontogruppe / konto</th>' : '';
            const labelHead2 = showLabels ? '<th class="label-col"></th>' : '';
            const body = report.rows.map((r,index)=>{
                if (r.type === 'subtotal') return `<tr class="summary">${showLabels?`<td class="label-col">${esc(r.name)}</td>`:''}${cells(r.amounts,r.percentages)}</tr>`;
                if (r.type === 'computed') return `<tr class="computed">${showLabels?`<td class="label-col">${esc(r.name)}</td>`:''}${cells(r.amounts,r.percentages)}</tr>`;
                const groupLabel = showLabels ? `<td class="label-col"><button type="button" data-group="${index}" aria-expanded="false">▸ ${esc(r.name)}</button></td>` : '';
                return `<tr class="group">${groupLabel}${cells(r.amounts,r.percentages)}</tr>${r.accounts.map(a=>`<tr class="detail" data-detail="${index}" hidden>${showLabels?`<td class="label-col">${a.account} · ${esc(a.name)}</td>`:''}${cells(a.amounts,a.amounts.map((v,i)=>revenueAmounts[i]===0?null:v/revenueAmounts[i]*100))}</tr>`).join('')}`;
            }).join('');
            const extraHtml = (extraRows||[]).map((r,i)=>`<tr class="${r.type==='computed'?'computed':'summary'}${i===0?' section-start':''}">${showLabels?`<td class="label-col">${esc(r.name)}</td>`:''}${cells(r.amounts,r.percentages)}</tr>`).join('');
            return `<div class="sheet"><h2 class="table-title">${esc(title)}</h2><table><thead><tr>${labelHead1}${head1}</tr><tr>${labelHead2}${head2}</tr></thead><tbody>${body}${extraHtml}</tbody></table></div>`;
        }
        el('report').innerHTML = `<div class="tables">${buildTable('Periodens resultat',[0,1],true,report.periodOnly)}${buildTable('År til dato',[2,3],false)}</div>`;
        el('report').querySelectorAll('[data-group]').forEach(button=>button.addEventListener('click',()=>{
            const open = button.getAttribute('aria-expanded')!=='true';button.setAttribute('aria-expanded',String(open));
            el('report').querySelectorAll(`[data-detail="${button.dataset.group}"]`).forEach(row=>row.hidden=!open);
        }));
    }
    el('unit').addEventListener('change',render);
    el('filters').addEventListener('submit',async event=>{
        event.preventDefault();el('load').disabled=true;el('error').textContent='';el('status').textContent='Henter bogførte beløb…';report=null;el('report').innerHTML='';
        try {
            const response=await fetch(`/bilancio-trial/data?year=${encodeURIComponent(el('year').value)}&period=${encodeURIComponent(el('period').value)}`,{cache:'no-store'});
            const data=await response.json();
            if(!response.ok)throw Error(response.status===401||response.status===403?'Log ind i Operations Hub med adgang til Økonomirapport for at åbne rapporten.':data.error||'Kunne ikke hente rapporten.');
            report=data;render();el('status').textContent=`Opdateret ${new Date(data.generatedAt).toLocaleString('da-DK')} · ${data.rows.reduce((n,r)=>n+(r.accounts?r.accounts.length:0),0)} konti`;
        } catch(error){el('status').textContent='';el('error').textContent=error.message;}
        finally{el('load').disabled=false;}
    });
    el('filters').requestSubmit();
})();
