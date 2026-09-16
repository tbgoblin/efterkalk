(function (root) {
    'use strict';
    const escapeHtml = value => String(value == null ? '' : value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
    const numeric = value => Number.isFinite(Number(value)) ? Number(value) : 0;
    const format = (value, decimals = 0) => numeric(value).toLocaleString('da-DK', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
    const palette = ['#1565c0', '#00acc1', '#00897b', '#7b1fa2', '#ef6c00', '#5e35b1', '#43a047', '#c62828'];

    function revenueStacked(rows, monthKeys) {
        const months = new Map(monthKeys.map(key => [key.slice(0, 7), new Map()]));
        const accounts = new Map();
        for (const row of rows) {
            const key = String(row.acNo || '');
            if (!accounts.has(key)) accounts.set(key, String(row.name || ''));
            const month = row.date instanceof Date ? row.date.toISOString().slice(0, 7) : String(row.date || '').slice(0, 7);
            const values = months.get(month);
            if (values) values.set(key, numeric(values.get(key)) + numeric(row.revenueMio));
        }
        const totals = [...months.values()].map(values => [...values.values()].reduce((sum, value) => sum + value, 0));
        const stacks = [...months.values()].map(values => [...values.values()].reduce((sum, value) => {
            sum[value >= 0 ? 'positive' : 'negative'] += value;
            return sum;
        }, { positive: 0, negative: 0 }));
        const min = Math.min(0, ...stacks.map(stack => stack.negative));
        let max = Math.max(0, ...stacks.map(stack => stack.positive));
        if (min === 0 && max === 0) max = 0.1;
        const left = 48, top = 30, height = 214, innerHeight = 142, barWidth = 34, slot = 52;
        const innerWidth = Math.max(560, monthKeys.length * slot);
        const toY = value => top + (max - value) / (max - min) * innerHeight;
        let html = '<g>';
        for (let tick = 0; tick <= 4; tick += 1) {
            const value = max - tick / 4 * (max - min), yPosition = top + innerHeight * tick / 4;
            html += '<line x1="' + left + '" y1="' + yPosition + '" x2="' + (left + innerWidth) + '" y2="' + yPosition + '" stroke="#d9e6f8" stroke-width="1" />';
            html += '<text x="' + (left - 6) + '" y="' + (yPosition + 4) + '" text-anchor="end" font-size="10" fill="#5f7892">' + format(value, 3) + '</text>';
        }
        html += '<line x1="' + left + '" y1="' + toY(0) + '" x2="' + (left + innerWidth) + '" y2="' + toY(0) + '" stroke="#8fa8c2" stroke-width="1.2" />';
        [...months.entries()].forEach(([month, values], index) => {
            const xPosition = left + index * slot;
            const label = new Date(month + '-01T12:00:00').toLocaleDateString('da-DK', { month: 'short', year: 'numeric' });
            let positive = 0, negative = 0;
            [...accounts.entries()].forEach(([account, name], accountIndex) => {
                const value = numeric(values.get(account));
                if (value === 0) return;
                const start = value > 0 ? positive : negative;
                const end = start + value;
                if (value > 0) positive = end;
                else negative = end;
                html += '<rect x="' + xPosition + '" y="' + Math.min(toY(start), toY(end)) + '" width="' + barWidth + '" height="' + Math.max(1, Math.abs(toY(end) - toY(start))) + '" fill="' + palette[accountIndex % palette.length] + '" rx="2"><title>' + escapeHtml(label + ' - ' + account + ' ' + name + ': ' + format(value, 3) + ' Mio DKK (' + format(value * 1000000, 2) + ' DKK)') + '</title></rect>';
            });
            if (Math.abs(positive) < 0.000001 && Math.abs(negative) < 0.000001) {
                html += '<line x1="' + xPosition + '" y1="' + toY(0) + '" x2="' + (xPosition + barWidth) + '" y2="' + toY(0) + '" stroke="#cddced" stroke-width="1" />';
            }
            const total = totals[index], labelY = Math.max(12, toY(positive) - 8);
            html += '<text class="omsaetning-month-total" x="' + (xPosition + barWidth / 2) + '" y="' + labelY + '" text-anchor="middle" font-size="10" font-weight="700" fill="#173452" paint-order="stroke" stroke="#fff" stroke-width="3" stroke-linejoin="round"><title>' + escapeHtml(label + ': ' + format(total, 3) + ' Mio DKK') + '</title>' + format(total, 3) + ' Mio</text>';
            html += '<text x="' + (xPosition + barWidth / 2) + '" y="' + (top + innerHeight + 14) + '" text-anchor="middle" font-size="10" fill="#47617c">' + escapeHtml(label) + '</text>';
        });
        const legend = [...accounts.entries()].map(([account, name], index) => '<span class="omsaetning-legend-item"><span class="omsaetning-legend-swatch" style="background:' + palette[index % palette.length] + ';"></span>' + escapeHtml(account + ' ' + name) + '</span>').join('');
        return { html: html + '</g>', viewBox: '0 0 ' + (left + innerWidth + 20) + ' ' + height, legend };
    }

    function dateKey(rawDate, dateLabel) {
        const match = String(dateLabel || '').trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
        if (match) return match[3] + '-' + match[2].padStart(2, '0') + '-' + match[1].padStart(2, '0');
        const date = rawDate ? new Date(rawDate) : null;
        return date && Number.isFinite(date.getTime()) ? date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0') : '';
    }

    function displayDate(rawDate, dateLabel) {
        const label = String(dateLabel || '').trim();
        if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(label)) return label;
        const date = rawDate ? new Date(rawDate) : null;
        return date && Number.isFinite(date.getTime()) ? date.toLocaleDateString('da-DK') : label || '-';
    }

    function loadDays(sourceRows, options = {}) {
        const today = options.today || new Date().toISOString().slice(0, 10);
        const todayCut = new Date(today + 'T00:00:00').getTime();
        const groups = new Map();
        sourceRows.forEach((row, index) => {
            const key = dateKey(row.Dato, row.DatoX);
            const sort = key ? new Date(key + 'T00:00:00').getTime() : Number.MAX_SAFE_INTEGER;
            const before = options.collapseBeforeToday !== false && ((!row.Dato && !String(row.DatoX || '').trim()) || sort < todayCut);
            const bucketKey = before ? 'before' : options.collapseBeforeToday === false ? String(index) : key || '__unknown__' + index;
            if (!groups.has(bucketKey)) groups.set(bucketKey, {
                Kap: 0, Resv: 0, Aften: 0, __dayKey: before ? 'before' : key || '__unknown__' + index,
                __dateLabel: before ? '-' : displayDate(row.Dato, row.DatoX), __sort: before ? Number.MIN_SAFE_INTEGER : sort
            });
            const bucket = groups.get(bucketKey);
            bucket.Kap += before ? 0 : numeric(row.Kap);
            bucket.Resv += numeric(row.Resv);
            bucket.Aften += numeric(row.Aften);
        });
        return [...groups.values()].sort((left, right) => left.__sort - right.__sort);
    }

    function belastningCluster(sourceRows, options = {}) {
        const rows = options.prepared ? sourceRows : loadDays(sourceRows, options);
        if (!rows.length) return '';
        const left = 40, right = 12, top = 22, innerHeight = 180, height = 288;
        const parentWidth = Math.max(420, (options.viewportWidth || 1200) * 0.42);
        const groupWidth = options.fit ? Math.min(30, parentWidth / rows.length)
            : Math.max(16, Math.min(30, Math.floor(parentWidth / Math.max(12, Math.min(30, rows.length)))));
        const width = left + Math.max(parentWidth, rows.length * groupWidth) + right;
        const max = rows.reduce((value, row) => Math.max(value, numeric(row.Kap), numeric(row.Resv), numeric(row.Aften)), 1);
        const toY = value => top + innerHeight - Math.max(0, numeric(value)) / max * innerHeight;
        const toX = index => left + index * groupWidth + groupWidth / 2;
        const barWidth = Math.max(options.fit ? 0.5 : 3, Math.min(7, groupWidth / 3.4));
        const labelStep = options.fit ? Math.max(1, Math.ceil(rows.length / Math.floor(parentWidth / 14))) : 1;
        let html = '<svg class="belastning-svg"' + (options.fit ? '' : ' style="min-width:' + width + 'px"') + ' viewBox="0 0 ' + width + ' ' + height + '" preserveAspectRatio="xMinYMin meet" role="img" aria-label="Kapacitetsbelastning">';
        html += '<line class="axis" x1="' + left + '" y1="' + top + '" x2="' + left + '" y2="' + (top + innerHeight) + '"></line><line class="axis" x1="' + left + '" y1="' + (top + innerHeight) + '" x2="' + (width - right) + '" y2="' + (top + innerHeight) + '"></line>';
        for (let tick = 0; tick <= 4; tick += 1) {
            const yPosition = top + innerHeight * tick / 4;
            html += '<line class="grid" x1="' + left + '" y1="' + yPosition + '" x2="' + (width - right) + '" y2="' + yPosition + '"></line><text class="label" x="' + (left - 4) + '" y="' + (yPosition + 3) + '" text-anchor="end">' + format(Math.round(max * (1 - tick / 4))) + '</text>';
        }
        rows.forEach((row, index) => {
            const center = toX(index);
            const clickable = options.clickable && options.resGr && row.__dayKey;
            const click = clickable ? ' onclick="' + escapeHtml('onBelastningDayColumnClick(' + JSON.stringify(String(options.resGr)) + ',' + (options.parity === 0 ? 0 : 1) + ',' + JSON.stringify(row.__dayKey) + ', event)') + '"' : '';
            html += '<rect class="belastning-day-band' + (options.activeDayKey === row.__dayKey ? ' active' : '') + '" x="' + (center - barWidth * 1.9) + '" y="' + top + '" width="' + Math.max(barWidth * 3.8, 10) + '" height="' + innerHeight + '"' + click + '></rect>';
            [['Kap', 'kap', 'Kapacitet'], ['Resv', 'resv', 'Reservationer'], ['Aften', 'aften', 'Rest Aften']].forEach(([key, className, label], seriesIndex) => {
                const yPosition = toY(row[key]);
                html += '<rect class="belastning-series-' + className + '" x="' + (center + barWidth * (seriesIndex - 1.5)) + '" y="' + yPosition + '" width="' + barWidth + '" height="' + (top + innerHeight - yPosition) + '"><title>' + escapeHtml(row.__dateLabel + ' ' + label + ': ' + format(row[key])) + '</title></rect>';
            });
            const labelY = top + innerHeight + 44;
            if (index % labelStep === 0 || index === rows.length - 1) {
                html += '<text class="label" x="' + center + '" y="' + labelY + '" text-anchor="middle" transform="rotate(-90 ' + center + ' ' + labelY + ')">' + escapeHtml(row.__dateLabel) + '</text>';
            }
        });
        [['kap', 'Kapacitet', 0], ['resv', 'Reservationer', 88], ['aften', 'Rest Aften', 196]].forEach(([className, label, offset]) => {
            html += '<rect class="belastning-series-' + className + '" x="' + (left + 4 + offset) + '" y="10" width="10" height="10"></rect><text class="label" x="' + (left + 18 + offset) + '" y="19">' + label + '</text>';
        });
        return html + '</svg>';
    }

    function orderBudgetTargets(config) {
        const safe = config && typeof config === 'object' ? config : {};
        const workDaysRaw = Math.round(Number(safe.workDaysPerYear));
        const dailyRaw = Number(safe.dailyBudget);
        const workDaysPerYear = Number.isFinite(workDaysRaw) ? Math.max(1, Math.min(366, workDaysRaw)) : 235;
        const daily = (Number.isFinite(dailyRaw) ? Math.max(0, dailyRaw) : 280851) / 1000;
        const annual = daily * workDaysPerYear;
        return { daily, weekly: annual / 52, monthly: annual / 12, annual, workDaysPerYear, useManualBudget: safe.useManualBudget !== false };
    }

    function orderRowsForView(rows, config, holidaySet = new Set(), ignoreHolidays = true) {
        const safeRows = Array.isArray(rows) ? rows : [];
        const targets = orderBudgetTargets(config);
        if (!targets.useManualBudget) return safeRows;
        return safeRows.map(row => ({
            ...row,
            totalBudget: ignoreHolidays && holidaySet.has(String(row.weekKey || '')) && numeric(row.totalOrd) === 0 ? 0 : targets.weekly
        }));
    }

    function orderMovingAverage(values, windowSize, skipMask) {
        const safeValues = Array.isArray(values) ? values : [];
        const safeWindow = Math.max(1, Number(windowSize) || 1);
        const safeSkip = Array.isArray(skipMask) ? skipMask : [];
        return safeValues.map((_value, index) => {
            let sum = 0, count = 0;
            for (let previous = index; previous >= 0 && count < safeWindow; previous -= 1) {
                if (safeSkip[previous]) continue;
                sum += Number(safeValues[previous] || 0);
                count += 1;
            }
            return count > 0 ? sum / count : null;
        });
    }

    function orderHolidayWeeks(value) {
        const normalized = String(value || '').split(',').map(part => {
            const cleaned = String(part || '').trim().replace(/[^0-9\-]/g, '');
            const range = /^([0-9]{6})-([0-9]{6})$/.exec(cleaned);
            if (range && range[2] >= range[1]) return range[1] + '-' + range[2];
            const single = cleaned.replace(/-/g, '').slice(0, 6);
            return /^[0-9]{6}$/.test(single) ? single : '';
        }).filter(Boolean);
        const result = new Set();
        for (const part of normalized) {
            const range = /^([0-9]{6})-([0-9]{6})$/.exec(part);
            if (!range) { result.add(part); continue; }
            let year = Number(range[1].slice(0, 4)), week = Number(range[1].slice(4, 6));
            for (let guard = 0; guard < 160; guard += 1) {
                const key = String(year) + String(week).padStart(2, '0');
                result.add(key);
                if (key >= range[2]) break;
                const december28 = new Date(Date.UTC(year, 11, 28));
                december28.setUTCDate(december28.getUTCDate() + 4 - (december28.getUTCDay() || 7));
                const weeks = Math.ceil(((december28 - new Date(Date.UTC(year, 0, 1))) / 86400000 + 1) / 7);
                week += 1;
                if (week > weeks) { year += 1; week = 1; }
            }
        }
        return result;
    }

    const api = { revenueStacked, belastningCluster, loadDays, orderBudgetTargets, orderRowsForView, orderMovingAverage, orderHolidayWeeks };
    if (typeof module === 'object' && module.exports) module.exports = api;
    else root.GohReportCharts = api;
}(typeof globalThis === 'object' ? globalThis : this));