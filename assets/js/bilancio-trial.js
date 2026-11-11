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
    let extras = { monthly: null, yearly: null, error: '' };
    let extrasToken = 0;
    let activeBundle = null;
    async function api(url, options = {}) {
        const response = await fetch(url, { cache: 'no-store', ...options });
        const data = await response.json();
        if (!response.ok) throw Error(data.error || 'Rapporten kunne ikke hentes.');
        return data;
    }
    function archiveUrl(id, reportId) {
        return `/bilancio-trial/archive/${encodeURIComponent(id)}?reportId=${encodeURIComponent(reportId)}`;
    }
    function applyBundle(bundle) {
        activeBundle = bundle;
        report = bundle.report;
        extras = { monthly: bundle.monthly, yearly: bundle.yearly, error: '' };
        el('year').value = report.year;
        el('period').value = report.period;
        if (bundle.id) el('unit').value = bundle.unit;
        el('unit').disabled = !!bundle.id;
        el('saveReportBtn').disabled = !!bundle.id;
        el('downloadPdfBtn').disabled = !bundle.id;
        el('downloadJsonBtn').disabled = !bundle.id;
        el('archiveNote').disabled = !!bundle.id;
        el('archiveNote').value = bundle.note || '';
        const revision = bundle.configurationRevision;
        el('revisionInfo').textContent = (bundle.id
            ? `GEMT · Rapport Rev. ${bundle.revision} · ${bundle.savedBy} · ${new Date(bundle.savedAt).toLocaleString('da-DK')}`
            : `KLADDE · Næste rapportrevision: ${bundle.baseRevision + 1}`)
            + ` · Konfiguration Rev. ${revision.number} (${revision.id.slice(0, 12)}) · App ${bundle.appVersion}`
            + (bundle.templateLockedAt ? ' · Månedens skabelon er låst' : '')
            + (bundle.note ? ' · ' + bundle.note : '');
        render();
    }
    async function refreshArchive(reportId) {
        const data = await api('/bilancio-trial/archive?reportId=' + encodeURIComponent(reportId));
        if (el('reportSelect').value !== reportId) return;
        el('archiveSelect').innerHTML = '<option value="">Vælg en gemt rapport</option>' + data.reports.map(item =>
            `<option value="${esc(item.id)}">${item.year}-${String(item.year + 1).slice(-2)} · P${item.period} · Rev. ${item.revision} · ${esc(new Date(item.savedAt).toLocaleString('da-DK'))}</option>`
        ).join('');
    }
    el('saveReportBtn').addEventListener('click', async () => {
        if (!activeBundle?.draftId || activeBundle.id) return;
        if (!confirm('Gem rapporten permanent? Den kan ikke overskrives. Rettelser oprettes som en ny revision.')) return;
        const token = extrasToken;
        el('saveReportBtn').disabled = true;
        el('error').textContent = '';
        try {
            const saved = await api('/bilancio-trial/archive', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ draftId: activeBundle.draftId, unit: Number(el('unit').value), note: el('archiveNote').value })
            });
            if (token !== extrasToken) return;
            applyBundle(saved);
            el('status').textContent = 'Rapport og PDF er gemt permanent på GOH.';
            await refreshArchive(saved.report.reportId);
            if (token === extrasToken) el('archiveSelect').value = saved.id;
        } catch (error) {
            if (token === extrasToken) {
                el('error').textContent = error.message;
                el('saveReportBtn').disabled = !!activeBundle?.id;
            }
        }
    });
    el('openArchiveBtn').addEventListener('click', async () => {
        const id = el('archiveSelect').value, reportId = el('reportSelect').value;
        if (!id) { el('error').textContent = 'Vælg en gemt rapportrevision.'; return; }
        const token = ++extrasToken;
        el('load').disabled = true;
        el('saveReportBtn').disabled = true;
        el('error').textContent = '';
        el('status').textContent = 'Henter arkiveret rapport…';
        try {
            const saved = await api(archiveUrl(id, reportId));
            if (token !== extrasToken) return;
            applyBundle(saved);
            el('status').textContent = 'Arkiveret rapport · ingen genberegning fra Visma.';
        } catch (error) {
            if (token === extrasToken) {
                el('error').textContent = error.message; el('status').textContent = '';
                el('saveReportBtn').disabled = !activeBundle || !!activeBundle.id;
            }
        } finally { if (token === extrasToken) el('load').disabled = false; }
    });
    async function download(format) {
        if (!activeBundle?.id) return;
        const { id, report: currentReport, revision } = activeBundle;
        el('error').textContent = '';
        try {
            const url = archiveUrl(id, currentReport.reportId).replace('?reportId=', '/download?reportId=') + '&format=' + format;
            const response = await fetch(url, { cache: 'no-store' });
            if (!response.ok) { const data = await response.json(); throw Error(data.error || 'Download mislykkedes.'); }
            const objectUrl = URL.createObjectURL(await response.blob());
            const link = document.createElement('a');
            link.href = objectUrl;
            link.download = `${currentReport.reportId}-${currentReport.year}-P${currentReport.period}-Rev${revision}.${format}`;
            document.body.appendChild(link);
            link.click();
            link.remove();
            setTimeout(() => URL.revokeObjectURL(objectUrl), 60000);
        } catch (error) { el('error').textContent = error.message; }
    }
    el('downloadPdfBtn').addEventListener('click', () => download('pdf'));
    el('downloadJsonBtn').addEventListener('click', () => download('json'));
    function printReport(orientation) {
        const styleEl = el('printOrientation');
        if (styleEl) styleEl.textContent = `@page{size:A4 ${orientation};margin:10mm}`;
        // A saved admin layout places the two report blocks freely on the A4 page instead of
        // the default stacked flow; --page-w/h and each block's --x/y/w/h drive the (unconditional,
        // not print-only) positioning CSS, so we can measure and fit content before printing.
        // Without a saved layout, "fit-width" just narrows #report to the page's printable width
        // (normal stacked flow, no absolute positioning) so the same measurement still works.
        const container = el('report'), layout = report?.printLayout?.[orientation];
        const [pageW, pageH] = orientation === 'landscape' ? [277, 190] : [190, 277];
        container.classList.toggle('custom-layout', !!layout);
        container.classList.toggle('fit-width', !layout);
        container.style.setProperty('--page-w', pageW + 'mm');
        container.style.setProperty('--page-h', pageH + 'mm');
        const fitBlocks = [...container.querySelectorAll('[data-block], .extra-sheet')];
        for (const block of fitBlocks) {
            const inner = block.querySelector('.block-inner');
            if (!block || !inner) continue;
            inner.style.transform = '';
            if (block.matches('.extra-sheet')) continue;
            const key = block.dataset.block;
            const rect = layout && layout[key];
            if (!rect) continue;
            block.style.setProperty('--x', rect.x + '%');
            block.style.setProperty('--y', rect.y + '%');
            block.style.setProperty('--w', rect.width + '%');
            block.style.setProperty('--h', rect.height + '%');
        }
        // Give the browser a full reflow for the new @page orientation before printing,
        // otherwise the print pipeline can capture the page mid-reflow and rotate/squeeze it to fit.
        requestAnimationFrame(() => {
            for (const block of fitBlocks) {
                const inner = block.querySelector('.block-inner');
                if (!block || !inner || !inner.scrollWidth || !inner.scrollHeight) continue;
                if (block.matches('.extra-sheet')) continue;
                const scaleX = block.clientWidth / inner.scrollWidth, scaleY = block.clientHeight / inner.scrollHeight;
                // With a saved layout, fill the chosen box on both axes (may enlarge). Without one,
                // only ever shrink to the page width and let height flow/paginate naturally.
                const scale = layout ? Math.min(scaleX, scaleY) : Math.min(1, scaleX);
                if (Number.isFinite(scale) && scale > 0) inner.style.transform = `scale(${scale})`;
            }
            requestAnimationFrame(() => window.print());
        });
    }
    el('printPortraitBtn').addEventListener('click', () => printReport('portrait'));
    el('printLandscapeBtn').addEventListener('click', () => printReport('landscape'));
    function printFourPageReport() {
        if (!report || !extras.monthly || !extras.yearly) {
            el('status').textContent = extras.error || 'Attendi il caricamento delle quattro tabelle prima di stampare.';
            return;
        }
        const styleEl = el('printOrientation');
        if (styleEl) styleEl.textContent = '@page{margin:10mm}';
        const container = el('report');
        container.classList.remove('custom-layout', 'fit-width');
        container.classList.add('mixed-print');
        container.closest('main').classList.add('mixed-print');
        container.querySelectorAll('.block-inner').forEach(inner => { inner.style.transform = ''; });
        requestAnimationFrame(() => requestAnimationFrame(() => window.print()));
    }
    el('printFourPageBtn').addEventListener('click', printFourPageReport);
    window.addEventListener('afterprint', () => {
        const styleEl = el('printOrientation');
        if (styleEl) styleEl.textContent = '@page{size:A4 portrait;margin:10mm}';
        const container = el('report');
        container.classList.remove('custom-layout', 'fit-width');
        container.classList.remove('mixed-print');
        container.closest('main').classList.remove('mixed-print');
        container.querySelectorAll('.block-inner').forEach(inner => { inner.style.transform = ''; });
    });
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
            return `<div class="sheet" data-block="pnl"><div class="block-inner"><table><thead><tr><th class="label-col"></th>${groupHead}</tr><tr><th class="label-col">Kontogruppe / konto</th>${head1}</tr><tr><th class="label-col"></th>${head2}</tr></thead><tbody>${body}${extraHtml}</tbody></table></div></div>`;
        }
        const balanceFmt = value => value == null ? '—' : new Intl.NumberFormat('da-DK', { minimumFractionDigits: divisor===1?2:0, maximumFractionDigits: divisor===1?2:0 }).format(Object.is(value,-0)?0:value/divisor);
        const assetsHtml = report.assets ? `<section class="sheet assets-sheet" data-block="balance" aria-label="${esc(report.assets.title)}"><div class="block-inner"><h2 class="table-title">${esc(report.assets.title)}</h2><table><thead><tr><th class="label-col">År til dato</th><th>${esc(months[p-1])} ${calendarYear}<br>${unit}</th><th>${esc(months[p-1])} ${calendarYear-1}<br>${unit}</th></tr></thead><tbody>${report.assets.rows.filter(row => !['hidden','period'].includes(row.visibility)).map(row => `<tr class="${(row.bold || ['subtotal', 'heading'].includes(row.type)) ? 'summary' : 'asset-row'}"><td class="label-col">${esc(row.name)}</td>${row.amounts.map((amount,i) => `<td class="${i===0?'current':''}">${row.type === 'heading' ? '' : balanceFmt(amount) + zeroStatus(row,amount)}</td>`).join('')}</tr>`).join('')}</tbody></table></div></section>` : '';
        // Fiscal periods start in July: period 1 = juli. "okt-26"-style labels for month columns.
        const monthLabel = c => months[c.period-1].slice(0,3).toLowerCase()+'-'+String(c.year+(c.period>6?1:0)).slice(-2);
        const yearLabel = c => `${c.year}-${String(c.year+1).slice(-2)}` + (c.year===report.year && report.period<12 ? ' (ÅTD)' : '');
        // One t.kr column per month/year — no comparison columns and no % columns by design.
        function extraTable(data, key, title, columnLabel, skipVisibility) {
            if (extras.error) return `<section class="sheet extra-sheet"><div class="empty">${esc(title)}: ${esc(extras.error)}</div></section>`;
            if (!data) return `<section class="sheet extra-sheet"><div class="empty">Henter ${esc(title.toLowerCase())}…</div></section>`;
            const blank = data.columns.map(()=>'<td></td>').join('');
            const cells = (amounts,row={}) => amounts.map(v=>`<td class="${v<0?'negative':''}">${fmt(v==null?null:v/divisor)+zeroStatus(row,v)}</td>`).join('');
            const body = data.rows.map((r,index)=>{
                if (r.visibility==='hidden'||r.visibility===skipVisibility) return '';
                if (r.type==='heading') return `<tr class="summary"><td class="label-col">${esc(r.name)}</td>${blank}</tr>`;
                if (r.type==='subtotal') return `<tr class="summary"><td class="label-col">${esc(r.name)}</td>${cells(r.amounts,r)}</tr>`;
                if (r.type==='computed') return `<tr class="${r.bold?'summary':'computed'}"><td class="label-col">${esc(r.name)}</td>${cells(r.amounts,r)}</tr>`;
                return `<tr class="group" style="${r.bold?'font-weight:bold':''}"><td class="label-col"><button type="button" data-group="${key}${index}" aria-expanded="false">▸ ${esc(r.name)}</button></td>${cells(r.amounts,r)}</tr>`+
                    r.accounts.map(a=>`<tr class="detail" data-detail="${key}${index}" hidden><td class="label-col">${a.account} · ${esc(a.name)}</td>${cells(a.amounts)}</tr>`).join('');
            }).join('');
            return `<section class="sheet extra-sheet" aria-label="${esc(title)}"><div class="block-inner"><h2 class="table-title">${esc(title)}</h2><table><thead><tr><th class="label-col">Kontogruppe / konto</th>${data.columns.map(c=>`<th>${esc(columnLabel(c))}</th>`).join('')}</tr><tr><th class="label-col"></th>${data.columns.map(()=>`<th>${unit}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table></div></section>`;
        }
        const monthlyHtml = extraTable(extras.monthly, 'm', 'Seneste 12 måneder', monthLabel, 'ytd');
        const yearlyHtml = extraTable(extras.yearly, 'y', 'Seneste 12 regnskabsår', yearLabel, 'period');
        el('report').innerHTML = `${report.showPnl===false?'':buildTable(report.periodOnly)}${assetsHtml}${monthlyHtml}${yearlyHtml}`;
        el('report').querySelectorAll('[data-group]').forEach(button=>button.addEventListener('click',()=>{
            const open = button.getAttribute('aria-expanded')!=='true';button.setAttribute('aria-expanded',String(open));
            el('report').querySelectorAll(`[data-detail="${button.dataset.group}"]`).forEach(row=>row.hidden=!open);
        }));
    }
    el('unit').addEventListener('change',render);
    el('filters').addEventListener('submit',async event=>{
        event.preventDefault();
        const token = ++extrasToken;
        const reportId = el('reportSelect').value;
        el('load').disabled=true;el('error').textContent='';el('status').textContent='Henter bogførte beløb…';report=null;activeBundle=null;el('report').innerHTML='';extras={monthly:null,yearly:null,error:''};
        el('revisionInfo').textContent = '';
        el('saveReportBtn').disabled = true;
        el('downloadPdfBtn').disabled = true;
        el('downloadJsonBtn').disabled = true;
        el('unit').disabled = false;
        el('archiveNote').disabled = false;
        try {
            await refreshArchive(reportId);
            if (token !== extrasToken) return;
            const data=await api(`/bilancio-trial/draft?year=${encodeURIComponent(el('year').value)}&period=${encodeURIComponent(el('period').value)}&reportId=${encodeURIComponent(reportId)}`);
            if (token !== extrasToken) return;
            applyBundle(data);
            el('status').textContent=`Opdateret ${new Date(report.generatedAt).toLocaleString('da-DK')} · ${report.rows.reduce((n,r)=>n+(r.accounts?r.accounts.length:0),0)} konti`;
        } catch(error){if (token === extrasToken) { el('status').textContent='';el('error').textContent=error.message; }}
        finally{if (token === extrasToken) el('load').disabled=false;}
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
