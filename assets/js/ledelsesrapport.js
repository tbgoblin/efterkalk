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
        const zeroWeekMask = rows.map(row => number(row.totalOrd) === 0);
        const moving = charts.orderMovingAverage(rows.map(row => number(row.totalOrd)), 3, zeroWeekMask);
        const average = orders.kpis?.avgSumOrd;
        return {
            average:average != null && Number.isFinite(Number(average)) ? Number(average) : null,
            rows:rows.map((row,index) => ({...row, ma3:moving[index], skipOrdLine:ignoreHolidays && holidayMask[index],
            isAnomaly:!zeroWeekMask[index] && moving[index] > 0 && Math.abs((number(row.totalOrd) - moving[index]) / moving[index] * 100) >= 20 }))
        };
    }

    function lineChart(rows, average, selectedLines = ['totalOrd', 'ma3', 'totalBudget', 'periodAverage']) {
        const dense = rows.length > 20;
        const width = 1000, height = dense ? 400 : 320, left = 58, top = 32, bottom = dense ? 146 : 66, innerWidth = width - left - 16, innerHeight = height - top - bottom;
        const slot = innerWidth / Math.max(1, rows.length);
        const labelSize = Math.min(11, slot * 0.6), valueSize = Math.min(12, slot * 0.65);
        const markerRadius = Math.min(2.7, slot / 5);
        const selected = new Set(selectedLines);
        const series = [
            { key:'totalOrd', label:'Ordre', color:'#0b3f88' },
            { key:'ma3', label:'Gns. ordre (3 uger)', color:'#ff6f00' },
            { key:'totalBudget', label:'Budget', color:'#1b8f3b' }
        ].filter(item => selected.has(item.key));
        const values = rows.flatMap(row => [number(row.totalOrd), ...series.map(item => number(row[item.key]))]);
        if (selected.has('periodAverage') && average !== null && average !== undefined) values.push(average);
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
            if (selected.has('ma3') && row.isAnomaly) svg += '<circle data-series="anomaly" cx="'+x(index)+'" cy="'+(y(row.totalOrd)-7)+'" r="'+Math.min(3.8,slot/4)+'" fill="#fff" stroke="#e65100" stroke-width="2"><title>Anomali (dev fra MA3)</title></circle>';
        });
        if (selected.has('periodAverage') && average !== null && average !== undefined) {
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

    function horizontalBars(items, valueKey, label, formatter, rowLabel = item => item.name||item.resGr||item.custNo) {
        const width=780,rowHeight=25,left=190,right=190,height=Math.max(90,items.length*rowHeight+20),max=Math.max(.01,...items.map(item=>number(item[valueKey]))),inner=width-left-right;
        let svg='<svg viewBox="0 0 '+width+' '+height+'" style="height:'+height+'px" aria-label="'+esc(label)+'">';
        items.forEach((item,index)=>{const yy=10+index*rowHeight,w=Math.max(1,number(item[valueKey])/max*inner);svg+='<text x="'+(left-7)+'" y="'+(yy+13)+'" text-anchor="end" font-size="10" fill="#334f68">'+esc(rowLabel(item))+'</text><rect x="'+left+'" y="'+yy+'" width="'+w+'" height="16" fill="'+colors[index%colors.length]+'" rx="2"/><text x="'+(left+w+5)+'" y="'+(yy+12)+'" font-size="9" font-weight="700" fill="#17324d">'+esc(formatter(item[valueKey],item))+'</text>';});
        return svg+'</svg>';
    }

    function customerBars(items) {
        const width=900,rowHeight=34,left=180,right=330,height=Math.max(100,items.length*rowHeight+44),inner=width-left-right;
        const max=Math.max(.01,...items.map(item=>Math.max(number(item.revenueMio),number(item.costMio))));
        let svg='<svg viewBox="0 0 '+width+' '+height+'" style="height:'+height+'px" aria-label="Største kunder med omsætning, kost og dækningsbidrag">';
        items.forEach((item,index)=>{
            const yy=10+index*rowHeight,revenue=Math.max(0,number(item.revenueMio)),known=item.costMio!=null&&Number.isFinite(Number(item.costMio));
            const cost=known?Math.max(0,number(item.costMio)):0,db=known?revenue-cost:null,scale=inner/max;
            const revenueWidth=Math.max(1,revenue*scale),costInside=Math.min(cost,revenue)*scale,positiveDb=Math.max(0,db||0)*scale,loss=Math.max(0,-(db||0))*scale;
            const customer=item.name||item.custNo;
            svg+='<text x="'+(left-7)+'" y="'+(yy+13)+'" text-anchor="end" font-size="10" fill="#334f68">'+esc(customer)+'</text>';
            if(!known) svg+='<rect x="'+left+'" y="'+yy+'" width="'+revenueWidth+'" height="16" fill="#8fa8c2" rx="2"><title>'+esc(customer+': DB kan ikke beregnes')+'</title></rect>';
            else {
                if(costInside>0) svg+='<rect x="'+left+'" y="'+yy+'" width="'+Math.max(1,costInside)+'" height="16" fill="#78909c" rx="2"><title>'+esc('Kost: '+dkk(cost*1000000))+'</title></rect>';
                if(positiveDb>0) svg+='<rect x="'+(left+costInside)+'" y="'+yy+'" width="'+Math.max(1,positiveDb)+'" height="16" fill="#2e7d32" rx="2"><title>'+esc('DB: '+dkk(db*1000000))+'</title></rect>';
                if(loss>0) svg+='<rect x="'+(left+revenueWidth)+'" y="'+yy+'" width="'+Math.max(1,loss)+'" height="16" fill="#c62828" rx="2"><title>'+esc('Negativ DB: '+dkk(db*1000000))+'</title></rect>';
            }
            const summary=known?'Oms. '+dkk(revenue*1000000)+' · DB '+dkk(db*1000000)+' ('+number(item.dbPct).toLocaleString('da-DK',{minimumFractionDigits:1,maximumFractionDigits:1})+'%)':'Oms. '+dkk(revenue*1000000)+' · DB kan ikke beregnes';
            svg+='<text x="'+(left+Math.max(revenueWidth,cost*scale)+7)+'" y="'+(yy+12)+'" font-size="9" font-weight="700" fill="#17324d">'+esc(summary)+'</text>';
        });
        return svg+'</svg><div class="legend"><span><i class="swatch" style="background:#78909c"></i>Kost</span><span><i class="swatch" style="background:#2e7d32"></i>DB</span><span><i class="swatch" style="background:#c62828"></i>Negativ DB</span></div>';
    }

    function resourceLoadPanels(data, selectedResources) {
        const selected = Array.isArray(selectedResources) ? new Set(selectedResources.map(String)) : null;
        const selectedOrder = new Map((selectedResources || []).map((key,index) => [String(key),index]));
        const grouped = new Map();
        data.load.rows.forEach(row => {
            const key = String(row.ResGr || '').trim();
            if (!key || (selected && !selected.has(key))) return;
            if (!grouped.has(key)) grouped.set(key, []);
            grouped.get(key).push(row);
        });
        const cards = [];
        const today = data.filters.loadFrom || data.generatedAt.slice(0,10);
        const resources = [...grouped.entries()].sort(([left],[right]) => selected
            ? (selectedOrder.get(left) ?? Number.MAX_SAFE_INTEGER) - (selectedOrder.get(right) ?? Number.MAX_SAFE_INTEGER)
            : left.localeCompare(right,'da',{numeric:true}))
            .map(([key, rows]) => ({key, rows, days:window.GohReportCharts.loadDays(rows, {today})}));
        const loadFrom = new Date(String(data.filters.loadFrom || '') + 'T00:00:00');
        const loadTo = new Date(String(data.filters.loadTo || '') + 'T00:00:00');
        const configuredDays = Number.isFinite(loadFrom.getTime()) && Number.isFinite(loadTo.getTime())
            ? Math.floor((loadTo - loadFrom) / 86400000) + 1 : number(data.filters.loadDays);
        const wide = configuredDays > 32 || resources.some(resource => resource.days.length > 32);
        resources.forEach(({key, rows, days}) => {
            const sum = field => rows.reduce((total,row) => total + number(row[field]),0);
            const capacity = sum('Kap'), reserved = sum('Resv'), evening = sum('Aften');
            const percent = capacity > 0 ? Math.min(160, reserved / capacity * 100).toLocaleString('da-DK',{maximumFractionDigits:1,minimumFractionDigits:1})+'%' : '0,0%';
            const minutes = value => number(value).toLocaleString('da-DK',{maximumFractionDigits:0});
            cards.push('<article class="belastning-resource-chart" data-resource="'+esc(key)+'"><h3>Kapacitetsbelastning: '+esc(key+' '+(rows[0].Nm || ''))+'</h3>'+window.GohReportCharts.belastningCluster(days,{prepared:true,fit:true,viewportWidth:wide ? 2400 : 1200})+'<div class="belastning-mini-meta"><span>Belastning: '+esc(percent)+'</span><span>Resv: '+esc(minutes(reserved))+'</span><span>Kap: '+esc(minutes(capacity))+'</span><span>Aften: '+esc(minutes(evening))+'</span></div></article>');
        });
        return chunks(cards, wide ? 2 : 4).map((page,index) => '<section class="load-page"><div class="panel-head"><h2>Belastning pr. ressource'+(index ? ' (fortsat)' : '')+'</h2><span class="unit">Aktuel plan · '+esc(today)+(data.filters.loadTo ? ' til '+esc(data.filters.loadTo) : ' · '+esc(data.filters.loadDays)+' dage frem')+' · minutter · rest før startdato</span></div><div class="resource-grid'+(wide ? ' single-column' : '')+'">'+page.join('')+'</div></section>').join('') || '<section class="load-page"><h2>Belastning pr. ressource</h2><p class="note">Ingen data i valgt periode.</p></section>';
    }

    function orderFlowPanel(flow) {
        if(!flow||!flow.total)return '<section class="panel"><h2>Ordrebeholdning</h2><p class="note">Ordrebeholdning kunne ikke hentes. Genstart serveren.</p></section>';
        const total=flow.total,prior=flow.prior||{},received=flow.received||{};
        const columns=[['opening','Primo'],['incoming','Tilgang'],['invoiced','Faktureret'],['closing','Ultimo']];
        const max=Math.max(1,...columns.map(([key])=>Math.abs(number(prior[key]))+Math.abs(number(received[key]))));
        const shortMio=value=>(number(value)/1000000).toLocaleString('da-DK',{minimumFractionDigits:1,maximumFractionDigits:2})+' mio.';
        const asOf=String(flow.asOf||''),date=/^\d{8}$/.test(asOf)?asOf.slice(6,8)+'.'+asOf.slice(4,6)+'.'+asOf.slice(0,4):asOf;
        const bars=columns.map(([key,label])=>{
            const previous=Math.abs(number(prior[key])),fresh=Math.abs(number(received[key]));
            return '<div class="report-flow-column"><strong>'+esc(shortMio(total[key]))+'</strong><div class="report-flow-plot"><span class="report-flow-segment report-flow-prior" style="height:'+(previous/max*100).toFixed(3)+'%"></span><span class="report-flow-segment report-flow-new" style="height:'+(fresh/max*100).toFixed(3)+'%"></span></div><span class="report-flow-label">'+label+'</span></div>';
        }).join('');
        const adjustment=Math.abs(number(total.adjustment))>.01?' + regulering '+dkk(total.adjustment):'';
        const warning=flow.unknownCount?'<p class="note">'+esc(flow.unknownCount+' ordrer uden afstemt fakturahistorik er udeladt.')+'</p>':'';
        return '<section class="panel"><div class="report-flow-head"><div><span>Ordrebeholdning pr. '+esc(date)+'</span><strong>'+esc(shortMio(total.closing))+' <small>DKK</small></strong></div><div><span>Færdigfaktureret i måneden</span><b>'+esc(total.completed)+' <small>ordrer</small></b><span>Heraf '+esc(received.completed||0)+' nye ordrer</span></div></div><div class="legend"><span><i class="swatch report-flow-prior"></i>Ordrer fra tidligere måneder</span><span><i class="swatch report-flow-new"></i>Ordrer modtaget i måneden</span></div><div class="report-flow-chart">'+bars+'</div><div class="report-flow-foot"><span>Primo + tilgang − faktureret'+esc(adjustment)+' = ultimo</span><span>Rest i dag: <b>'+esc(dkk(total.current))+'</b></span></div>'+warning+'</section>';
    }

    function render(data, settings = readOrderViewSettings(), display = {}) {
        const orderView = orderReportView(data.orders, settings);
        const selectedLines = Array.isArray(display.orderLines) ? display.orderLines : ['totalOrd', 'ma3', 'totalBudget', 'periodAverage'];
        const lineLegend = {totalOrd:['#0b3f88','Ordre (linje)'],ma3:['#ff6f00','Gns. ordre (3 uger)'],totalBudget:['#1b8f3b','Budget'],periodAverage:['#7b1fa2','Gns. ordre i perioden']};
        const orderLegend = '<span><i class="swatch" style="background:#2f5ea5"></i>Ordre (søjle)</span>'+selectedLines.filter(key=>lineLegend[key]).map(key=>'<span><i class="swatch" style="background:'+lineLegend[key][0]+'"></i>'+lineLegend[key][1]+'</span>').join('');
        const orderPanels = orderView.rows.length ? '<section class="panel wide order-panel"><div class="panel-head"><h2>Ordreindgang</h2><span class="unit">'+esc(data.filters.orderFrom)+' til '+esc(data.filters.orderTo)+' · t.DKK</span></div>'+lineChart(orderView.rows,orderView.average,selectedLines)+'<div class="legend">'+orderLegend+'</div></section>' : '<section class="panel wide"><h2>Ordreindgang</h2><p class="note">Ingen data i valgt ugeperiode.</p></section>';
        report.innerHTML='<header class="report-head"><div><h1>Ledelsesrapport</h1><p class="subtitle">Økonomi, ordreindgang, belastning og arbejde i gang</p></div><div class="meta"><b>'+esc(data.filters.from)+' til '+esc(data.filters.to)+'</b><br>Dannet '+esc(new Date(data.generatedAt).toLocaleString('da-DK'))+'<br>VIA aktuel pr. '+esc(new Date(data.via.asOf).toLocaleString('da-DK'))+'</div></header>'+
            '<section class="kpis"><div class="kpi"><div class="label">Omsætning</div><div class="value">'+esc(mio(data.revenue.totalRevenueMio))+'</div></div><div class="kpi"><div class="label">VIA kost</div><div class="value">'+esc(dkk(data.via.totalCost))+'</div></div><div class="kpi"><div class="label">Åbne ordrers salg</div><div class="value">'+esc(dkk(data.via.totalSales))+'</div></div><div class="kpi"><div class="label">Resterende salg</div><div class="value">'+esc(dkk(data.via.remainingSales))+'</div></div><div class="kpi"><div class="label">Åbne ordrer</div><div class="value">'+esc(data.via.orderCount)+'</div></div></section>'+
            '<div class="grid">'+stackedRevenue(data)+
            '<section class="panel"><div class="panel-head"><h2>Største kunder</h2><span class="unit">'+esc(data.filters.customerFrom || data.filters.from)+' til '+esc(data.filters.customerTo || data.filters.to)+' · omsætning, kost og DB</span></div>'+customerBars(data.revenue.topCustomers)+'</section>'+
            orderFlowPanel(data.orderFlow)+orderPanels+'</div>'+resourceLoadPanels(data,display.loadResources);
    }

    async function loadConfig() {
        const [response,resourcesResponse]=await Promise.all([fetch('/ledelsesrapport/config'),fetch('/belastning/resources')]);
        const [data,resourcesData]=await Promise.all([response.json(),resourcesResponse.json()]);
        if(!response.ok) throw new Error(data.error||'Ingen adgang til ledelsesrapporten.');
        if(!resourcesResponse.ok) throw new Error(resourcesData.error||'Ressourcerne kunne ikke hentes.');
        document.getElementById('accounts').innerHTML=data.accounts.map(account=>'<label><input type="checkbox" value="'+esc(account.acNo)+'" checked disabled> '+esc(account.acNo+' · '+account.name)+'</label>').join('');
        const resources=[...new Map((resourcesData.resources||[]).map(item=>[String(item.MainR7||'').trim(),item])).values()].filter(item=>String(item.MainR7||'').trim());
        document.getElementById('loadResources').innerHTML=resources.map(item=>'<div class="resource-option" data-resource="'+esc(item.MainR7)+'"><span class="resource-drag-handle" draggable="true" title="Træk for at flytte" aria-label="Træk '+esc(item.R7Nm||item.MainR7)+'">⋮⋮</span><label><input type="checkbox" value="'+esc(item.MainR7)+'" checked> '+esc(item.R7Nm||item.MainR7)+'</label><span class="resource-page-badge"></span><span class="resource-order"><button type="button" data-move="-1" title="Flyt op" aria-label="Flyt '+esc(item.R7Nm||item.MainR7)+' op">↑</button><button type="button" data-move="1" title="Flyt ned" aria-label="Flyt '+esc(item.R7Nm||item.MainR7)+' ned">↓</button></span></div>').join('');
        return data;
    }

    function refreshResourcePagePreview() {
        const from=new Date(document.getElementById('loadFromDate').value+'T00:00:00');
        const to=new Date(document.getElementById('loadToDate').value+'T00:00:00');
        const days=Number.isFinite(from.getTime())&&Number.isFinite(to.getTime())?Math.floor((to-from)/86400000)+1:1;
        const pageSize=days>32?2:4;
        let selectedIndex=0;
        document.querySelectorAll('#loadResources .resource-option').forEach(row=>{
            const checked=row.querySelector('input').checked;
            row.classList.toggle('page-start',checked&&selectedIndex>0&&selectedIndex%pageSize===0);
            row.querySelector('.resource-page-badge').textContent=checked?'Side '+(Math.floor(selectedIndex/pageSize)+1)+' · plads '+(selectedIndex%pageSize+1)+'/'+pageSize:'Ikke med';
            if(checked)selectedIndex+=1;
        });
    }

    function reorderResourceOptions(container, orderedIds) {
        if(!container||!Array.isArray(orderedIds)) return;
        const rows=new Map([...container.querySelectorAll('.resource-option')].map(row=>[String(row.dataset.resource),row]));
        orderedIds.forEach(id=>{const row=rows.get(String(id));if(row){container.appendChild(row);rows.delete(String(id));}});
        rows.forEach(row=>container.appendChild(row));
    }

    document.getElementById('loadResources').addEventListener('click',event=>{
        const button=event.target.closest('[data-move]');
        if(!button)return;
        const row=button.closest('.resource-option'),direction=Number(button.dataset.move);
        const sibling=direction<0?row.previousElementSibling:row.nextElementSibling;
        if(!sibling)return;
        if(direction<0)row.parentElement.insertBefore(row,sibling);else row.parentElement.insertBefore(sibling,row);
        button.focus();
        refreshResourcePagePreview();
    });
    document.getElementById('loadResources').addEventListener('change',refreshResourcePagePreview);
    document.getElementById('loadResources').addEventListener('dragstart',event=>{
        const row=event.target.closest('.resource-option');if(!row)return;
        row.classList.add('dragging');event.dataTransfer.effectAllowed='move';event.dataTransfer.setData('text/plain',row.dataset.resource);
    });
    document.getElementById('loadResources').addEventListener('dragover',event=>{
        const target=event.target.closest('.resource-option'),dragging=document.querySelector('#loadResources .resource-option.dragging');
        if(!target||!dragging||target===dragging)return;
        event.preventDefault();
        const before=event.clientY<target.getBoundingClientRect().top+target.getBoundingClientRect().height/2;
        target.parentElement.insertBefore(dragging,before?target:target.nextElementSibling);
        refreshResourcePagePreview();
    });
    document.getElementById('loadResources').addEventListener('drop',event=>{event.preventDefault();refreshResourcePagePreview();});
    document.getElementById('loadResources').addEventListener('dragend',event=>{event.target.closest('.resource-option')?.classList.remove('dragging');refreshResourcePagePreview();});

    function applyDefaults(defaults) {
        if(!defaults) return;
        const fields={from:'fromMonth',to:'toMonth',customerFrom:'customerFromMonth',customerTo:'customerToMonth',topCustomers:'topCustomers',orderFrom:'orderFromWeek',orderTo:'orderToWeek',loadFrom:'loadFromDate',loadTo:'loadToDate',viaPeriod:'viaPeriod',viaFrom:'viaFromDate',viaTo:'viaToDate'};
        Object.entries(fields).forEach(([key,id])=>{if(defaults[key]!=null) document.getElementById(id).value=String(defaults[key]);});
        const selectedResources=Array.isArray(defaults.loadResources)?new Set(defaults.loadResources.map(String)):null;
        reorderResourceOptions(document.getElementById('loadResources'),defaults.loadResources);
        document.querySelectorAll('#loadResources input').forEach(input=>{input.checked=!selectedResources||selectedResources.has(input.value);});
        const selectedLines=new Set(Array.isArray(defaults.orderLines)?defaults.orderLines:['totalOrd','ma3','totalBudget','periodAverage']);
        document.querySelectorAll('#orderLines input').forEach(input=>{input.checked=selectedLines.has(input.value);});
        document.getElementById('viaPeriod').dispatchEvent(new Event('change'));
        refreshResourcePagePreview();
    }

    window.openConfig = () => dialog.showModal();

    form.addEventListener('submit', async event => {
        event.preventDefault();
        const accounts=[...document.querySelectorAll('#accounts input:checked')].map(input=>input.value);
        if(!accounts.length){status.textContent='Vælg mindst én konto.';return;}
        const display={loadResources:[...document.querySelectorAll('#loadResources input:checked')].map(input=>input.value),orderLines:[...document.querySelectorAll('#orderLines input:checked')].map(input=>input.value)};
        if(!display.loadResources.length){status.textContent='Vælg mindst én ressource til Belastning.';status.className='error';return;}
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
            render(data,settings,display);status.textContent='';dialog.close();
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
    document.getElementById('loadFromDate').addEventListener('change',refreshResourcePagePreview);
    document.getElementById('loadToDate').addEventListener('change',refreshResourcePagePreview);
    document.getElementById('viaFromDate').value=dayKey(from);
    document.getElementById('viaToDate').value=dayKey(now);
    document.getElementById('viaPeriod').addEventListener('change',event=>{
        for(const id of ['viaFromDate','viaToDate']) {
            const input=document.getElementById(id);
            input.disabled=event.target.value!=='dates';input.required=!input.disabled;
        }
    });
    loadConfig().then(config=>{
        applyDefaults(config.defaults);
        if (config.accounts.length) document.querySelectorAll('#accounts input').forEach(input => { input.checked=true; });
        dialog.showModal();
    }).catch(error=>{report.innerHTML='<div class="empty error">'+esc(error.message)+'</div>';});
}());