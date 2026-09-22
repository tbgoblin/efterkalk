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
    let report = null;
    function printReport(orientation) {
        const styleEl = el('printOrientation');
        if (styleEl) styleEl.textContent = `@page{size:A4 ${orientation};margin:10mm}`;
        // A saved admin layout places the two report blocks freely on the A4 page instead of
        // the default stacked flow; --page-w/h and each block's --x/y/w/h drive the print-only CSS.
        const container = el('report'), layout = report?.printLayout?.[orientation];
        container.classList.toggle('custom-layout', !!layout);
        if (layout) {
            const [pageW, pageH] = orientation === 'landscape' ? [277, 190] : [190, 277];
            container.style.setProperty('--page-w', pageW + 'mm');
            container.style.setProperty('--page-h', pageH + 'mm');
            for (const key of ['pnl', 'balance']) {
                const block = container.querySelector(`[data-block="${key}"]`), rect = layout[key];
                if (!block || !rect) continue;
                block.style.setProperty('--x', rect.x + '%');
                block.style.setProperty('--y', rect.y + '%');
                block.style.setProperty('--w', rect.width + '%');
                block.style.setProperty('--h', rect.height + '%');
            }
        }
        // Give the browser a full reflow for the new @page orientation before printing,
        // otherwise the print pipeline can capture the page mid-reflow and rotate/squeeze it to fit.
        requestAnimationFrame(() => requestAnimationFrame(() => window.print()));
    }
    el('printPortraitBtn').addEventListener('click', () => printReport('portrait'));
    el('printLandscapeBtn').addEventListener('click', () => printReport('landscape'));
    window.addEventListener('afterprint', () => el('report').classList.remove('custom-layout'));
    function render() {
        if (!report) return;
        const divisor = Number(el('unit').value), unit = divisor===1000?'t.kr':'DKK';
        const fmt = v => v===null?'—':new Intl.NumberFormat('da-DK',{minimumFractionDigits:divisor===1?2:0,maximumFractionDigits:divisor===1?2:0}).format(Object.is(v,-0)?0:v);
        const pct = v => v===null?'—':new Intl.NumberFormat('da-DK',{maximumFractionDigits:0}).format(v)+'%';
        const y = report.year, p = report.period;
        const calendarYear = y+(p>6?1:0);
        const abbr = i => months[p-1].slice(0,3).toLowerCase()+'-'+String(i).slice(-2);
        const labels = [abbr(calendarYear), abbr(calendarYear-1), `Juli ${y} - ${months[p-1].toLowerCase()} ${calendarYear}`, `Juli ${y-1} - ${months[p-1].toLowerCase()} ${calendarYear-1}`];
        el('pageTitle').textContent = `Gantech A/S – ${report.reportName || 'Økonomirapport'} ${months[p-1].toLowerCase()} ${calendarYear}`;
        const zeroStatus = (row, value) => !row.checkZero ? '' : value == null ? ' · Mangler data' : Math.abs(value) < 0.005 ? ' · ✓ OK' : ' · ⚠ Afvigelse';
        const revenueAmounts = report.revenueAmounts || report.rows[0].amounts;
        // A single table with one <tr> per report line, holding both period and year-to-date
        // columns together. Two independent side-by-side <table> elements can never be guaranteed
        // to keep matching row heights (different classes on filler rows, border rendering, …),
        // so any drift compounds down the page. One shared row per line rules that out entirely.
        function buildTable(extraRows) {
            const colIndices = [0,1,2,3];
            const cls = (...names) => names.filter(Boolean).join(' ');
            const divider = li => li===2?'group-divider':li>0?'period-divider':'';
            const cells = (amounts, percentages, row={}) => colIndices.map((ci,li)=>{
                const blank = (row.visibility==='period' && ci>=2) || (row.visibility==='ytd' && ci<2);
                return `<td class="${cls(li===0&&'current', divider(li), !blank && amounts[ci]<0 && 'negative')}">${blank?'':fmt(amounts[ci] == null ? null : amounts[ci]/divisor)+zeroStatus(row,amounts[ci])}</td><td class="${cls(li===0&&'current')}">${blank?'':pct(percentages[ci])}</td>`;
            }).join('');
            const groupHead = `<th colspan="4">Periodens resultat</th><th colspan="4" class="group-divider">År til dato</th>`;
            const head1 = colIndices.map((ci,li)=>`<th colspan="2" class="${divider(li)}">Regnskabsåret ${y-ci%2}-${String(y+1-ci%2).slice(-2)}<br>${esc(labels[ci])}</th>`).join('');
            const head2 = colIndices.map((ci,li)=>`<th class="${divider(li)}">${unit}</th><th>%</th>`).join('');
            const body = report.rows.map((r,index)=>{
                if (r.visibility === 'hidden') return '';
                if (r.type === 'heading') return `<tr class="summary"><td class="label-col" title="${esc(r.description || r.name)}">${esc(r.name)}</td>${colIndices.map((ci,li)=>`<td class="${divider(li)}"></td><td></td>`).join('')}</tr>`;
                if (r.type === 'subtotal') return `<tr class="summary"><td class="label-col" title="${esc(r.description || r.name)}">${esc(r.name)}</td>${cells(r.amounts,r.percentages,r)}</tr>`;
                if (r.type === 'computed') return `<tr class="${r.bold?'summary':'computed'}"><td class="label-col" title="${esc(r.description || r.name)}">${esc(r.name)}</td>${cells(r.amounts,r.percentages,r)}</tr>`;
                const groupLabel = `<td class="label-col"><button type="button" data-group="${index}" aria-expanded="false">▸ ${esc(r.name)}</button></td>`;
                return `<tr class="group" style="${r.bold?'font-weight:bold':''}">${groupLabel}${cells(r.amounts,r.percentages,r)}</tr>${r.accounts.map(a=>`<tr class="detail" data-detail="${index}" hidden><td class="label-col">${a.account} · ${esc(a.name)}</td>${cells(a.amounts,a.amounts.map((v,i)=>revenueAmounts[i]===0?null:v/revenueAmounts[i]*100))}</tr>`).join('')}`;
            }).join('');
            const extraHtml = (extraRows||[]).map((r,i)=>`<tr class="${r.type==='computed'?'computed':'summary'}${i===0?' section-start':''}"><td class="label-col" title="${esc(r.description || r.name)}">${esc(r.name)}</td>${cells(r.amounts,r.percentages,r)}</tr>`).join('');
            return `<div class="sheet" data-block="pnl"><table><thead><tr><th class="label-col"></th>${groupHead}</tr><tr><th class="label-col">Kontogruppe / konto</th>${head1}</tr><tr><th class="label-col"></th>${head2}</tr></thead><tbody>${body}${extraHtml}</tbody></table></div>`;
        }
        const balanceFmt = value => value == null ? '—' : new Intl.NumberFormat('da-DK', { minimumFractionDigits: divisor===1?2:0, maximumFractionDigits: divisor===1?2:0 }).format(Object.is(value,-0)?0:value/divisor);
        const assetsHtml = report.assets ? `<section class="sheet assets-sheet" data-block="balance" aria-label="${esc(report.assets.title)}"><h2 class="table-title">${esc(report.assets.title)}</h2><table><thead><tr><th class="label-col">År til dato</th><th>${esc(months[p-1])} ${calendarYear}<br>${unit}</th><th>${esc(months[p-1])} ${calendarYear-1}<br>${unit}</th></tr></thead><tbody>${report.assets.rows.filter(row => !['hidden','period'].includes(row.visibility)).map(row => `<tr class="${(row.bold || ['subtotal', 'heading'].includes(row.type)) ? 'summary' : 'asset-row'}"><td class="label-col">${esc(row.name)}</td>${row.amounts.map((amount,i) => `<td class="${i===0?'current':''}">${row.type === 'heading' ? '' : balanceFmt(amount) + zeroStatus(row,amount)}</td>`).join('')}</tr>`).join('')}</tbody></table></section>` : '';
        el('report').innerHTML = `${report.showPnl===false?'':buildTable(report.periodOnly)}${assetsHtml}`;
        el('report').querySelectorAll('[data-group]').forEach(button=>button.addEventListener('click',()=>{
            const open = button.getAttribute('aria-expanded')!=='true';button.setAttribute('aria-expanded',String(open));
            el('report').querySelectorAll(`[data-detail="${button.dataset.group}"]`).forEach(row=>row.hidden=!open);
        }));
    }
    el('unit').addEventListener('change',render);
    el('filters').addEventListener('submit',async event=>{
        event.preventDefault();el('load').disabled=true;el('error').textContent='';el('status').textContent='Henter bogførte beløb…';report=null;el('report').innerHTML='';
        try {
            const response=await fetch(`/bilancio-trial/data?year=${encodeURIComponent(el('year').value)}&period=${encodeURIComponent(el('period').value)}&reportId=${encodeURIComponent(el('reportSelect').value)}`,{cache:'no-store'});
            const data=await response.json();
            if(!response.ok)throw Error(response.status===401||response.status===403?'Log ind i Operations Hub med adgang til Økonomirapport for at åbne rapporten.':data.error||'Kunne ikke hente rapporten.');
            report=data;render();el('status').textContent=`Opdateret ${new Date(data.generatedAt).toLocaleString('da-DK')} · ${data.rows.reduce((n,r)=>n+(r.accounts?r.accounts.length:0),0)} konti`;
        } catch(error){el('status').textContent='';el('error').textContent=error.message;}
        finally{el('load').disabled=false;}
    });
    el('reportSelect').onchange=()=>el('filters').requestSubmit();
    (async()=>{
        try {
            const response=await fetch('/bilancio-trial/reports',{cache:'no-store'});const data=await response.json();
            if(!response.ok)throw Error(data.error||'Log ind med adgang til Økonomirapport.');
            el('reportSelect').innerHTML=data.reports.map(r=>`<option value="${esc(r.reportId)}">${esc(r.reportName)}</option>`).join('');
            const id=new URLSearchParams(location.search).get('reportId')||'default';
            if(!data.reports.some(r=>r.reportId===id))throw Error('Rapporten er ikke gemt eller findes ikke.');
            el('reportSelect').value=id;el('filters').requestSubmit();
        }catch(error){el('error').textContent=error.message;}
    })();
})();
