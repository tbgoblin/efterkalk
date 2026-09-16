(function () {
    'use strict';

    const colors = ['#2c6aa3', '#38a3a5', '#4f8a5b', '#d49a35', '#c45b4b', '#725b9a', '#6b879d', '#a66c3f'];
    const dialog = document.getElementById('configDialog');
    const form = document.getElementById('configForm');
    const report = document.getElementById('report');
    const status = document.getElementById('status');
    const number = value => Number.isFinite(Number(value)) ? Number(value) : 0;
    const esc = value => String(value == null ? '' : value).replace(/[&<>"']/g, char => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[char]));
    const dkk = value => number(value).toLocaleString('da-DK', { maximumFractionDigits:0 }) + ' kr.';
    const mio = value => number(value).toLocaleString('da-DK', { minimumFractionDigits:1, maximumFractionDigits:1 }) + ' mio.';
    const short = value => Math.abs(number(value)) >= 1000 ? (number(value) / 1000).toLocaleString('da-DK', { maximumFractionDigits:1 }) + 'k' : Math.round(number(value)).toLocaleString('da-DK');

    function monthKey(value) {
        const date = new Date(value);
        return Number.isNaN(date.getTime()) ? String(value || '').slice(0, 7) : date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
    }

    function monthLabel(key) {
        const parts = String(key).split('-');
        return parts.length === 2 ? parts[1] + '/' + parts[0].slice(2) : key;
    }

    function readOrderViewSettings() {
        const read = key => {
            try { return JSON.parse(localStorage.getItem(key)) || {}; }
            catch { return {}; }
        };
        return { budget:read('afterkalk_ordreindgang_budget_v1'), holidays:read('afterkalk_ordreindgang_holiday_settings_v1') };
    }

    function orderReportView(orders, settings) {
        const charts = window.GohReportCharts;
        const holidays = settings.holidays || {};
        const holidaySet = charts.orderHolidayWeeks(holidays.holidayWeeksText);
        const ignoreHolidays = holidays.ignoreHolidayWeeks !== false;
        const rows = charts.orderRowsForView(orders.weeklyRows, settings.budget, holidaySet, ignoreHolidays);
        const holidayMask = rows.map(row => holidaySet.has(String(row.weekKey || '')) && number(row.totalOrd) === 0);
        const moving = charts.orderMovingAverage(rows.map(row => number(row.totalOrd)), 3, ignoreHolidays ? holidayMask : null);
        const average = orders.kpis?.avgSumOrd;
        return {
            average:average != null && Number.isFinite(Number(average)) ? Number(average) : null,
            rows:rows.map((row,index) => ({...row, ma3:moving[index], skipOrdLine:ignoreHolidays && holidayMask[index],
                isAnomaly:!holidayMask[index] && moving[index] > 0 && Math.abs((number(row.totalOrd) - moving[index]) / moving[index] * 100) >= 20 }))
        };
    }

    function lineChart(rows, average) {
        const dense = rows.length > 20;
        const width = 1000, height = dense ? 400 : 320, left = 58, top = 32, bottom = dense ? 146 : 66, innerWidth = width - left - 16, innerHeight = height - top - bottom;
        const slot = innerWidth / Math.max(1, rows.length);
        const labelSize = Math.min(11, slot * 0.6), valueSize = Math.min(12, slot * 0.65);
        const markerRadius = Math.min(2.7, slot / 5);
        const series = [
            { key:'totalOrd', label:'Ordre', color:'#0b3f88' },
            { key:'ma3', label:'Gns. ordre (3 uger)', color:'#ff6f00' },
            { key:'totalBudget', label:'Budget', color:'#1b8f3b' }
        ];
        const values = rows.flatMap(row => series.map(item => number(row[item.key])));
        if (average !== null && average !== undefined) values.push(average);
        const min = Math.min(0, ...values), max = Math.max(1000, Math.ceil(Math.max(0,...values) / 1000) * 1000);
        const x = index => left + (index + 0.5) * innerWidth / Math.max(1, rows.length);
        const y = value => top + (max - number(value)) / (max - min) * innerHeight;
        let svg = '<svg class="weekly-chart" style="aspect-ratio:'+width+'/'+height+'" viewBox="0 0 '+width+' '+height+'" role="img" aria-label="Ordre og budget pr. uge">';
        for (let tick = 0; tick <= 6; tick += 1) {
            const value = max - (max - min) * tick / 6, yy = top + innerHeight * tick / 6;
            svg += '<line x1="'+left+'" y1="'+yy+'" x2="'+(width-16)+'" y2="'+yy+'" stroke="#dce5ed"/><text x="'+(left-7)+'" y="'+(yy+3)+'" text-anchor="end" font-size="10" fill="#60758a">'+esc(Math.round(value).toLocaleString('da-DK'))+'</text>';
        }
        const barWidth = Math.min(18, innerWidth / Math.max(1, rows.length) * 0.4);
        rows.forEach((row,index) => {
            const barY = y(row.totalOrd), zeroY = y(0);
            svg += '<rect data-series="ordreBars" x="'+(x(index)-barWidth/2)+'" y="'+Math.min(barY,zeroY)+'" width="'+barWidth+'" height="'+Math.abs(zeroY-barY)+'" fill="#2f5ea5" opacity="0.8"/>';
        });
        series.forEach(item => {
            let path = '', connected = false;
            rows.forEach((row,index) => {
                if (row[item.key] == null || (item.key === 'totalOrd' && row.skipOrdLine)) { connected = false; return; }
                path += (connected ? ' L' : ' M')+x(index)+' '+y(row[item.key]);
                connected = true;
            });
            svg += '<path data-series="'+item.key+'" d="'+path+'" fill="none" stroke="'+item.color+'" stroke-width="'+(item.key === 'ma3' ? 4.5 : 2.5)+'"'+(item.key === 'totalBudget' ? ' stroke-dasharray="8 5"' : '')+'/>';
            rows.forEach((row,index) => {
                if (row[item.key] == null || (item.key === 'totalOrd' && row.skipOrdLine)) return;
                svg += '<circle cx="'+x(index)+'" cy="'+y(row[item.key])+'" r="'+markerRadius+'" fill="'+item.color+'"><title>'+esc(item.key === 'totalOrd' ? 'Ordre: '+number(row.totalOrd).toLocaleString('da-DK')+' t.kr.' : item.label)+'</title></circle>';
            });
        });
        rows.forEach((row,index) => {
            if (row.isAnomaly) svg += '<circle data-series="anomaly" cx="'+x(index)+'" cy="'+(y(row.totalOrd)-7)+'" r="'+Math.min(3.8,slot/4)+'" fill="#fff" stroke="#e65100" stroke-width="2"><title>Anomali (dev fra MA3)</title></circle>';
        });
        if (average !== null && average !== undefined) {
            const averageY = y(average);
            svg += '<line data-series="periodAverage" x1="'+left+'" y1="'+averageY+'" x2="'+(width-16)+'" y2="'+averageY+'" stroke="#7b1fa2" stroke-width="2.6"/>';
            svg += '<text x="'+(width-20)+'" y="'+Math.max(top+14,averageY-8)+'" text-anchor="end" font-size="11" font-weight="700" fill="#7b1fa2" paint-order="stroke" stroke="#fff" stroke-width="3">Gns. ordre i perioden: '+esc(average.toLocaleString('da-DK',{minimumFractionDigits:2,maximumFractionDigits:2}))+'</text>';
        }
        rows.forEach((row,index) => {
            const key = String(row.weekKey || '');
            const labelY = dense ? top + innerHeight + 12 : height - 42;
            const valueY = dense ? top + innerHeight + 78 : height - 20;
            const alignment = position => dense ? 'text-anchor="end" transform="rotate(-90 '+x(index)+' '+position+')"' : 'text-anchor="middle"';
            svg += '<text data-week-label="'+esc(key)+'" x="'+x(index)+'" y="'+labelY+'" '+alignment(labelY)+' font-size="'+labelSize+'" fill="#50677c">'+esc(key.slice(0,4)+'-'+key.slice(-2))+'</text>';
            svg += '<text data-week-value="ordre" x="'+x(index)+'" y="'+valueY+'" '+alignment(valueY)+' font-size="'+valueSize+'" font-weight="700" fill="#2c6aa3">'+esc(number(row.totalOrd).toLocaleString('da-DK', {maximumFractionDigits:0}))+'</text>';
        });
        return svg + '</svg>';
    }

    function stackedRevenue(data) {
        const months = [];
        const cursor = new Date(data.filters.from + '-01T12:00:00');
        while (monthKey(cursor) <= data.filters.to && months.length < 36) {
            months.push(monthKey(cursor) + '-01');
            cursor.setMonth(cursor.getMonth() + 1);
        }
        const chart = window.GohReportCharts.revenueStacked(data.revenue.rows, months);
        return '<section class="panel wide revenue-panel"><div class="panel-head"><h2>Omsætning pr. måned</h2><span class="unit">Mio DKK · '+esc(monthLabel(data.filters.from)+' - '+monthLabel(data.filters.to))+'</span></div><svg class="omsaetning-chart-svg" viewBox="'+chart.viewBox+'" role="img" aria-label="Omsætning stacked pr. konto">'+chart.html+'</svg><div class="omsaetning-legend">'+chart.legend+'</div></section>';
    }

    function chunks(items, size) {
        const result = [];
        for (let index = 0; index < items.length; index += size) result.push(items.slice(index, index + size));
        return result;
    }

    function horizontalBars(items, valueKey, label, formatter) {
        const width=700,rowHeight=25,left=190,right=70,height=Math.max(90,items.length*rowHeight+20),max=Math.max(.01,...items.map(item=>number(item[valueKey]))),inner=width-left-right;
        let svg='<svg viewBox="0 0 '+width+' '+height+'" style="height:'+height+'px" aria-label="'+esc(label)+'">';
        items.forEach((item,index)=>{const yy=10+index*rowHeight,w=Math.max(1,number(item[valueKey])/max*inner);svg+='<text x="'+(left-7)+'" y="'+(yy+13)+'" text-anchor="end" font-size="10" fill="#334f68">'+esc(item.name||item.resGr||item.custNo)+'</text><rect x="'+left+'" y="'+yy+'" width="'+w+'" height="16" fill="'+colors[index%colors.length]+'" rx="2"/><text x="'+(left+w+5)+'" y="'+(yy+12)+'" font-size="9" font-weight="700" fill="#17324d">'+esc(formatter(item[valueKey]))+'</text>';});
        return svg+'</svg>';
    }

    function resourceLoadPanels(data) {
        const grouped = new Map();
        data.load.rows.forEach(row => {
            const key = String(row.ResGr || '').trim();
            if (!key) return;
            if (!grouped.has(key)) grouped.set(key, []);
            grouped.get(key).push(row);
        });
        const cards = [];
        const today = data.filters.loadFrom || data.generatedAt.slice(0,10);
        const resources = [...grouped.entries()].sort(([left],[right]) => left.localeCompare(right,'da',{numeric:true}))
            .map(([key, rows]) => ({key, rows, days:window.GohReportCharts.loadDays(rows, {today})}));
        const wide = resources.some(resource => resource.days.length > 32);
        resources.forEach(({key, rows, days}) => {
            const sum = field => rows.reduce((total,row) => total + number(row[field]),0);
            const capacity = sum('Kap'), reserved = sum('Resv'), evening = sum('Aften');
            const percent = capacity > 0 ? Math.min(160, reserved / capacity * 100).toLocaleString('da-DK',{maximumFractionDigits:1,minimumFractionDigits:1})+'%' : '0,0%';
            const minutes = value => number(value).toLocaleString('da-DK',{maximumFractionDigits:0});
            cards.push('<article class="belastning-resource-chart" data-resource="'+esc(key)+'"><h3>Kapacitetsbelastning: '+esc(key+' '+(rows[0].Nm || ''))+'</h3>'+window.GohReportCharts.belastningCluster(days,{prepared:true,fit:true,viewportWidth:wide ? 2400 : 1200})+'<div class="belastning-mini-meta"><span>Belastning: '+esc(percent)+'</span><span>Resv: '+esc(minutes(reserved))+'</span><span>Kap: '+esc(minutes(capacity))+'</span><span>Aften: '+esc(minutes(evening))+'</span></div></article>');
        });
        return chunks(cards, wide ? 2 : 4).map((page,index) => '<section class="load-page"><div class="panel-head"><h2>Belastning pr. ressource'+(index ? ' (fortsat)' : '')+'</h2><span class="unit">Aktuel plan · '+esc(today)+(data.filters.loadTo ? ' til '+esc(data.filters.loadTo) : ' · '+esc(data.filters.loadDays)+' dage frem')+' · minutter · rest før startdato</span></div><div class="resource-grid'+(wide ? ' single-column' : '')+'">'+page.join('')+'</div></section>').join('') || '<section class="load-page"><h2>Belastning pr. ressource</h2><p class="note">Ingen data i valgt periode.</p></section>';
    }

    function viaComposition(via) {
        const entries=[['Materialer',number(via.costs.MaterialCost)],['Stænger',number(via.costs.StangCost)],['Købedele',number(via.costs.PurchasedPartCost)],['Tid',number(via.costs.TimeCost)]];
        const total=Math.max(1,entries.reduce((sum,item)=>sum+item[1],0));let x=10,svg='<svg viewBox="0 0 1000 150" style="height:150px" aria-label="VIA kostfordeling"><rect x="10" y="32" width="980" height="46" fill="#edf3f7" rx="4"/>';
        entries.forEach((entry,index)=>{const width=entry[1]/total*980;if(width>0)svg+='<rect x="'+x+'" y="32" width="'+width+'" height="46" fill="'+colors[index]+'"><title>'+esc(entry[0]+': '+dkk(entry[1]))+'</title></rect>';x+=width;});
        svg+='</svg><div class="legend">'+entries.map((entry,index)=>'<span><i class="swatch" style="background:'+colors[index]+'"></i>'+esc(entry[0])+': <b>'+esc(dkk(entry[1]))+'</b></span>').join('')+'</div>';return svg;
    }

    function render(data, settings = readOrderViewSettings()) {
        const orderView = orderReportView(data.orders, settings);
        const orderPanels = orderView.rows.length ? '<section class="panel wide order-panel"><div class="panel-head"><h2>Ordreindgang</h2><span class="unit">'+esc(data.filters.orderFrom)+' til '+esc(data.filters.orderTo)+' · t.DKK</span></div>'+lineChart(orderView.rows,orderView.average)+'<div class="legend"><span><i class="swatch" style="background:#2f5ea5"></i>Ordre (søjle)</span><span><i class="swatch" style="background:#0b3f88"></i>Ordre (linje)</span><span><i class="swatch" style="background:#ff6f00"></i>Gns. ordre (3 uger)</span><span><i class="swatch" style="background:#1b8f3b"></i>Budget</span><span><i class="swatch" style="background:#7b1fa2"></i>Gns. ordre i perioden</span></div></section>' : '<section class="panel wide"><h2>Ordreindgang</h2><p class="note">Ingen data i valgt ugeperiode.</p></section>';
        const unknownNote=(data.via.unknownCostCount?data.via.unknownCostCount+' ordrer mangler komplet kostgrundlag.':'Alle åbne ordrer har komplet kostgrundlag.')+(data.via.excludedUndatedCount ? ' '+data.via.excludedUndatedCount+' ordrer uden ordredato er ikke med i perioden.' : '');
        const viaPeriod = data.filters.viaFrom ? 'Ordredato '+data.filters.viaFrom+' til '+data.filters.viaTo : 'Alle ordredatoer';
        report.innerHTML='<header class="report-head"><div><h1>Ledelsesrapport</h1><p class="subtitle">Økonomi, ordreindgang, belastning og arbejde i gang</p></div><div class="meta"><b>'+esc(data.filters.from)+' til '+esc(data.filters.to)+'</b><br>Dannet '+esc(new Date(data.generatedAt).toLocaleString('da-DK'))+'<br>VIA aktuel pr. '+esc(new Date(data.via.asOf).toLocaleString('da-DK'))+'</div></header>'+
            '<section class="kpis"><div class="kpi"><div class="label">Omsætning</div><div class="value">'+esc(mio(data.revenue.totalRevenueMio))+'</div></div><div class="kpi"><div class="label">VIA kost</div><div class="value">'+esc(dkk(data.via.totalCost))+'</div></div><div class="kpi"><div class="label">Åbne ordrers salg</div><div class="value">'+esc(dkk(data.via.totalSales))+'</div></div><div class="kpi"><div class="label">Resterende salg</div><div class="value">'+esc(dkk(data.via.remainingSales))+'</div></div><div class="kpi"><div class="label">Åbne ordrer</div><div class="value">'+esc(data.via.orderCount)+'</div></div></section>'+
            '<div class="grid">'+stackedRevenue(data)+
            '<section class="panel"><div class="panel-head"><h2>Største kunder</h2><span class="unit">'+esc(data.filters.customerFrom || data.filters.from)+' til '+esc(data.filters.customerTo || data.filters.to)+' · mio. DKK</span></div>'+horizontalBars(data.revenue.topCustomers,'revenueMio','Største kunder',mio)+'</section>'+
            '<section class="panel via-panel"><div class="panel-head"><h2>VIA kostfordeling</h2><span class="unit">Aktuel kost · '+esc(viaPeriod)+'</span></div><dl class="via-totals"><div><dt>'+ (data.via.unknownCostCount ? 'Kendt kost (ufuldstændig)' : 'Samlet kost') +'</dt><dd>'+esc(dkk(data.via.totalCost))+'</dd></div><div><dt>Forventet salg (åbne ordrer)</dt><dd>'+esc(dkk(data.via.totalSales))+'</dd></div></dl>'+viaComposition(data.via)+'<p class="note">'+esc(unknownNote)+'</p></section>'+orderPanels+'</div>'+resourceLoadPanels(data);
    }

    async function loadConfig() {
        const response=await fetch('/ledelsesrapport/config');
        const data=await response.json();
        if(!response.ok) throw new Error(data.error||'Ingen adgang til ledelsesrapporten.');
        document.getElementById('accounts').innerHTML=data.accounts.map(account=>'<label><input type="checkbox" value="'+esc(account.acNo)+'" checked> '+esc(account.acNo+' · '+account.name)+'</label>').join('');
    }

    window.setAllAccounts = checked => document.querySelectorAll('#accounts input').forEach(input => { input.checked=checked; });
    window.openConfig = () => dialog.showModal();

    form.addEventListener('submit', async event => {
        event.preventDefault();
        const accounts=[...document.querySelectorAll('#accounts input:checked')].map(input=>input.value);
        if(!accounts.length){status.textContent='Vælg mindst én konto.';return;}
        status.className='';status.textContent='Danner rapport...';document.getElementById('generateBtn').disabled=true;
        try{
            const params=new URLSearchParams({from:document.getElementById('fromMonth').value,to:document.getElementById('toMonth').value,customerFrom:document.getElementById('customerFromMonth').value,customerTo:document.getElementById('customerToMonth').value,orderFrom:document.getElementById('orderFromWeek').value,orderTo:document.getElementById('orderToWeek').value,accounts:accounts.join(','),topCustomers:document.getElementById('topCustomers').value,loadFrom:document.getElementById('loadFromDate').value,loadTo:document.getElementById('loadToDate').value});
            if(document.getElementById('viaPeriod').value==='dates') {
                params.set('viaFrom',document.getElementById('viaFromDate').value);
                params.set('viaTo',document.getElementById('viaToDate').value);
            }
            for (const [start,end] of [['from','to'],['customerFrom','customerTo'],['loadFrom','loadTo'],['viaFrom','viaTo']]) {
                if(params.has(start) && params.get(start)>params.get(end)) throw new Error('Til-dato skal være samme eller efter fra-dato.');
            }
            if(params.get('orderFrom')>params.get('orderTo')) throw new Error('Til-uge skal være samme eller efter fra-uge.');
            const [response, holidayData] = await Promise.all([
                fetch('/ledelsesrapport/data?'+params),
                fetch('/ordreindgang/holiday-settings', {signal:AbortSignal.timeout(5000)})
                    .then(result => result.ok ? result.json() : null).catch(() => null)
            ]);
            const data=await response.json();
            if(!response.ok) throw new Error(data.error||'Rapporten kunne ikke dannes.');
            if(!Array.isArray(data.load?.rows) || ['orderFrom','orderTo','customerFrom','customerTo','loadFrom','loadTo'].some(key=>data.filters?.[key]!==params.get(key)) || ['viaFrom','viaTo'].some(key=>(data.filters?.[key]||'')!==(params.get(key)||''))) throw new Error('Genstart serveren for at aktivere rapportens separate perioder.');
            const settings = readOrderViewSettings();
            if(holidayData?.ok && holidayData.settings) settings.holidays = holidayData.settings;
            render(data,settings);status.textContent='';dialog.close();
        }catch(error){status.textContent=error.message;status.className='error';}
        finally{document.getElementById('generateBtn').disabled=false;}
    });

    const now=new Date();document.getElementById('toMonth').value=now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0');
    const from=new Date(now.getFullYear(),now.getMonth()-11,1);document.getElementById('fromMonth').value=from.getFullYear()+'-'+String(from.getMonth()+1).padStart(2,'0');
    function isoWeek(date) {
        const value = new Date(Date.UTC(date.getFullYear(),date.getMonth(),date.getDate()));
        value.setUTCDate(value.getUTCDate()+4-(value.getUTCDay()||7));
        const year=value.getUTCFullYear();
        const week=Math.ceil(((value-new Date(Date.UTC(year,0,1)))/86400000+1)/7);
        return year+'-W'+String(week).padStart(2,'0');
    }
    document.getElementById('orderFromWeek').value=isoWeek(from);
    document.getElementById('orderToWeek').value=isoWeek(now);
    document.getElementById('customerFromMonth').value=monthKey(from);
    document.getElementById('customerToMonth').value=monthKey(now);
    const dayKey = date => date.getFullYear()+'-'+String(date.getMonth()+1).padStart(2,'0')+'-'+String(date.getDate()).padStart(2,'0');
    document.getElementById('loadFromDate').value=dayKey(now);
    document.getElementById('loadToDate').value=dayKey(new Date(now.getFullYear(),now.getMonth(),now.getDate()+29));
    document.getElementById('viaFromDate').value=dayKey(from);
    document.getElementById('viaToDate').value=dayKey(now);
    document.getElementById('viaPeriod').addEventListener('change',event=>{
        for(const id of ['viaFromDate','viaToDate']) {
            const input=document.getElementById(id);
            input.disabled=event.target.value!=='dates';input.required=!input.disabled;
        }
    });
    loadConfig().then(()=>dialog.showModal()).catch(error=>{report.innerHTML='<div class="empty error">'+esc(error.message)+'</div>';});
}());