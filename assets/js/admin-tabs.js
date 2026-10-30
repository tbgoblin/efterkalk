// ── Admin tabs (working days, report defaults, user management) · client ────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml (dichiarato nell'inline script),
// getOmsaetningThresholdInputs/scheduleOmsaetningAutoReload (assets/js/omsaetning.js),
// getOrdreindgangHolidayWeeksSet/shouldIgnoreOrdreindgangHolidayWeeks
// (assets/js/ordreindgang.js) e OmsaetningDailyThresholds
// (assets/js/omsaetning-daily-thresholds.js).
const ADMIN_MODULES = [
    ['efterkalk', 'Efterkalkulation'],
    ['salgordreVia', 'SalgOrdre VIA'],
    ['omsaetning', 'Omsætning'],
    ['ordreindgang', 'Ordreindgang'],
    ['ordreoversigt', 'Ordreoversigt'],
    ['belastning', 'Belastning'],
    ['ledelsesrapport', 'Ledelsesrapport'],
    ['personalehandbog', 'Personalehåndbog']
    ,['lagerliste', 'Lagerliste'],
    ['bilancio', 'Økonomirapport (Prøve)'],
    ['bomOverview', 'BOM · Oversigt'],
    ['bomStykliste', 'BOM · Stykliste'],
    ['bomComponents', 'BOM · Komponenter'],
    ['bomResources', 'BOM · Ressourcer'],
    ['bomMaterials', 'BOM · Materialer'],
    ['bomParameters', 'BOM · Parametre'],
    ['bomCalculator', 'BOM · Beregner'],
    ['bomSuppliers', 'BOM · Leverandører'],
    ['bomVismaPreview', 'BOM · Visma-preview/query']
];

function adminHeaders() {
    return { 'Content-Type': 'application/json', Authorization: 'Bearer ' + String(authToken || '') };
}

function setAdminStatus(text, isError) {
    const status = document.getElementById('adminStatus');
    if (!status) return;
    status.textContent = text || '';
    status.style.color = isError ? '#b71c1c' : '#1b5e20';
}

const ADMIN_MONTH_NAMES = ['Januar', 'Februar', 'Marts', 'April', 'Maj', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'December'];

function showAdminSection(sectionKey) {
    const sections = {
        users: ['adminTabUsers', 'adminSectionUsers'],
        diverse: ['adminTabDiverse', 'adminSectionDiverse'],
        'working-days': ['adminTabWorkingDays', 'adminSectionWorkingDays'],
        ledelsesrapport: ['adminTabLedelsesrapport', 'adminSectionLedelsesrapport']
    };
    const selectedKey = Object.prototype.hasOwnProperty.call(sections, sectionKey) ? sectionKey : 'users';
    try { sessionStorage.setItem('adminSection', selectedKey); } catch (_) {}
    Object.entries(sections).forEach(([key, ids]) => {
        const tab = document.getElementById(ids[0]);
        const panel = document.getElementById(ids[1]);
        const selected = key === selectedKey;
        if (tab) {
            tab.classList.toggle('active', selected);
            tab.setAttribute('aria-selected', selected ? 'true' : 'false');
            tab.tabIndex = selected ? 0 : -1;
        }
        if (panel) {
            panel.classList.toggle('active', selected);
            panel.hidden = !selected;
        }
    });
}

function handleAdminTabKeydown(event) {
    if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
    const tabs = Array.from(event.currentTarget.querySelectorAll('[role="tab"]'));
    const currentIndex = tabs.indexOf(document.activeElement);
    if (currentIndex < 0) return;
    event.preventDefault();
    let nextIndex = event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 : currentIndex + (event.key === 'ArrowRight' ? 1 : -1);
    nextIndex = (nextIndex + tabs.length) % tabs.length;
    tabs[nextIndex].focus();
    tabs[nextIndex].click();
}

function getAdminWorkingDaysYear() {
    const input = document.getElementById('adminWorkingDaysYear');
    const year = Number(input && input.value);
    return Number.isInteger(year) && year >= 2000 && year <= 2200 ? year : new Date().getFullYear();
}

function getCalendarWorkingDays(monthKey) {
    if (typeof OmsaetningDailyThresholds === 'undefined') return 0;
    return OmsaetningDailyThresholds.countWorkingDaysInMonth(monthKey, {
        holidayWeekKeys: getOrdreindgangHolidayWeeksSet(),
        excludeCompanyHolidayWeeks: shouldIgnoreOrdreindgangHolidayWeeks()
    });
}

function renderAdminWorkingDays(year, months) {
    const input = document.getElementById('adminWorkingDaysYear');
    const grid = document.getElementById('adminWorkingDaysGrid');
    if (input) input.value = String(year);
    if (!grid) return;
    grid.innerHTML = ADMIN_MONTH_NAMES.map((name, index) => {
        const monthKey = String(year) + '-' + String(index + 1).padStart(2, '0');
        const calendarDays = getCalendarWorkingDays(monthKey);
        const storedValue = months && Object.prototype.hasOwnProperty.call(months, monthKey) ? Number(months[monthKey]) : calendarDays;
        return '<label class="admin-working-day">' + name
            + '<input class="admin-working-day-input" type="number" min="0" max="31" step="1" data-month-key="' + monthKey + '" value="' + String(storedValue) + '">'
            + '<small>Kalender: ' + String(calendarDays) + '</small></label>';
    }).join('');
}

function setAdminWorkingDaysStatus(text, isError) {
    const status = document.getElementById('adminWorkingDaysStatus');
    if (!status) return;
    status.textContent = text || '';
    status.style.color = isError ? '#b71c1c' : '#1b5e20';
}

async function loadAdminWorkingDays() {
    const year = getAdminWorkingDaysYear();
    renderAdminWorkingDays(year, {});
    setAdminWorkingDaysStatus('Indlæser...');
    try {
        const response = await fetch('/admin/working-days?year=' + encodeURIComponent(String(year)), { headers: adminHeaders() });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'Kunne ikke hente arbejdsdage');
        renderAdminWorkingDays(year, data.months || {});
        setAdminWorkingDaysStatus(data.updatedAt ? 'Senest gemt: ' + new Date(data.updatedAt).toLocaleString('da-DK') : 'Kalenderforslag vises.');
    } catch (error) {
        setAdminWorkingDaysStatus(String(error.message || error), true);
    }
}

function resetAdminWorkingDaysToCalendar() {
    document.querySelectorAll('.admin-working-day-input').forEach(input => {
        input.value = String(getCalendarWorkingDays(String(input.dataset.monthKey || '')));
    });
    setAdminWorkingDaysStatus('Kalenderforslag er indsat. Klik Gem arbejdsdage.');
}

async function saveAdminWorkingDays() {
    const year = getAdminWorkingDaysYear();
    const months = {};
    let invalid = false;
    document.querySelectorAll('.admin-working-day-input').forEach(input => {
        const value = Number(input.value);
        if (!Number.isInteger(value) || value < 0 || value > 31) invalid = true;
        else months[String(input.dataset.monthKey || '')] = value;
    });
    if (invalid || Object.keys(months).length !== 12) {
        setAdminWorkingDaysStatus('Alle måneder skal have et helt tal fra 0 til 31.', true);
        return;
    }
    setAdminWorkingDaysStatus('Gemmer...');
    try {
        const response = await fetch('/admin/working-days', { method: 'POST', headers: adminHeaders(), body: JSON.stringify({ year, months }) });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'Kunne ikke gemme arbejdsdage');
        Object.assign(omsaetningWorkingDaysByMonth, data.months || {});
        setAdminWorkingDaysStatus('Arbejdsdage for ' + String(year) + ' er gemt.');
        if (omsaetningInitialized && getOmsaetningThresholdInputs().useDailyBudget) scheduleOmsaetningAutoReload();
    } catch (error) {
        setAdminWorkingDaysStatus(String(error.message || error), true);
    }
}

const ADMIN_REPORT_FIELDS = {
    from: 'adminReportFrom', to: 'adminReportTo', customerFrom: 'adminReportCustomerFrom', customerTo: 'adminReportCustomerTo',
    topCustomers: 'adminReportTopCustomers', orderFrom: 'adminReportOrderFrom', orderTo: 'adminReportOrderTo',
    loadFrom: 'adminReportLoadFrom', loadTo: 'adminReportLoadTo', viaPeriod: 'adminReportViaPeriod',
    viaFrom: 'adminReportViaFrom', viaTo: 'adminReportViaTo'
};

function setAdminReportDefaultsStatus(text, isError) {
    const status = document.getElementById('adminReportDefaultsStatus');
    if (!status) return;
    status.textContent = text || '';
    status.style.color = isError ? '#b71c1c' : '#1b5e20';
}

function updateAdminReportViaDates() {
    const useDates = document.getElementById('adminReportViaPeriod').value === 'dates';
    ['adminReportViaFrom', 'adminReportViaTo'].forEach(id => {
        const input = document.getElementById(id);
        input.disabled = !useDates;
        input.required = useDates;
    });
}

function setAdminReportDefaults(defaults) {
    Object.entries(ADMIN_REPORT_FIELDS).forEach(([key, id]) => {
        if (defaults && defaults[key] != null) document.getElementById(id).value = String(defaults[key]);
    });
    const selectedResources = defaults && Array.isArray(defaults.loadResources) ? new Set(defaults.loadResources.map(String)) : null;
    reorderAdminReportResources(defaults && defaults.loadResources);
    document.querySelectorAll('#adminReportLoadResources input').forEach(input => { input.checked = !selectedResources || selectedResources.has(input.value); });
    const selectedLines = new Set(defaults && Array.isArray(defaults.orderLines) ? defaults.orderLines : ['totalOrd', 'ma3', 'totalBudget', 'periodAverage']);
    document.querySelectorAll('#adminReportOrderLines input').forEach(input => { input.checked = selectedLines.has(input.value); });
    updateAdminReportViaDates();
    refreshAdminReportResourcePagePreview();
}

function reorderAdminReportResources(orderedIds) {
    const container = document.getElementById('adminReportLoadResources');
    if (!container || !Array.isArray(orderedIds)) return;
    const rows = new Map(Array.from(container.querySelectorAll('.admin-report-resource-row')).map(row => [String(row.dataset.resource), row]));
    orderedIds.forEach(id => {
        const row = rows.get(String(id));
        if (!row) return;
        container.appendChild(row);
        rows.delete(String(id));
    });
    rows.forEach(row => container.appendChild(row));
}

function moveAdminReportResource(button, direction) {
    const row = button && button.closest('.admin-report-resource-row');
    if (!row) return;
    const sibling = direction < 0 ? row.previousElementSibling : row.nextElementSibling;
    if (!sibling) return;
    if (direction < 0) row.parentElement.insertBefore(row, sibling);
    else row.parentElement.insertBefore(sibling, row);
    button.focus();
    refreshAdminReportResourcePagePreview();
    setAdminReportDefaultsStatus('Rækkefølgen er ændret. Klik Gem standard.');
}

function refreshAdminReportResourcePagePreview() {
    const from = new Date(String(document.getElementById('adminReportLoadFrom').value || '') + 'T00:00:00');
    const to = new Date(String(document.getElementById('adminReportLoadTo').value || '') + 'T00:00:00');
    const days = Number.isFinite(from.getTime()) && Number.isFinite(to.getTime()) ? Math.floor((to - from) / 86400000) + 1 : 1;
    const pageSize = days > 32 ? 2 : 4;
    let selectedIndex = 0;
    document.querySelectorAll('#adminReportLoadResources .admin-report-resource-row').forEach(row => {
        const checked = row.querySelector('input').checked;
        row.classList.toggle('page-start', checked && selectedIndex > 0 && selectedIndex % pageSize === 0);
        row.querySelector('.admin-report-resource-page').textContent = checked
            ? 'Side ' + (Math.floor(selectedIndex / pageSize) + 1) + ' · plads ' + (selectedIndex % pageSize + 1) + '/' + pageSize
            : 'Ikke med';
        if (checked) selectedIndex += 1;
    });
}

function startAdminReportResourceDrag(event) {
    const row = event.target.closest('.admin-report-resource-row');
    if (!row) return;
    row.classList.add('dragging');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', row.dataset.resource);
}

function dragAdminReportResource(event) {
    const target = event.target.closest('.admin-report-resource-row');
    const dragging = document.querySelector('#adminReportLoadResources .admin-report-resource-row.dragging');
    if (!target || !dragging || target === dragging) return;
    event.preventDefault();
    const before = event.clientY < target.getBoundingClientRect().top + target.getBoundingClientRect().height / 2;
    target.parentElement.insertBefore(dragging, before ? target : target.nextElementSibling);
    refreshAdminReportResourcePagePreview();
}

function endAdminReportResourceDrag(event) {
    event.target.closest('.admin-report-resource-row')?.classList.remove('dragging');
    refreshAdminReportResourcePagePreview();
    setAdminReportDefaultsStatus('Rækkefølgen er ændret. Klik Gem standard.');
}

function getAdminReportSuggestedDefaults() {
    const now = new Date();
    const from = new Date(now.getFullYear(), now.getMonth() - 11, 1);
    const monthKey = date => date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
    const dayKey = date => monthKey(date) + '-' + String(date.getDate()).padStart(2, '0');
    const weekKey = date => {
        const value = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
        value.setUTCDate(value.getUTCDate() + 4 - (value.getUTCDay() || 7));
        const year = value.getUTCFullYear();
        const week = Math.ceil((((value - new Date(Date.UTC(year, 0, 1))) / 86400000) + 1) / 7);
        return year + '-W' + String(week).padStart(2, '0');
    };
    return { from: monthKey(from), to: monthKey(now), customerFrom: monthKey(from), customerTo: monthKey(now), topCustomers: 10,
        orderFrom: weekKey(from), orderTo: weekKey(now), loadFrom: dayKey(now), loadTo: dayKey(new Date(now.getFullYear(), now.getMonth(), now.getDate() + 29)),
        viaPeriod: 'all', viaFrom: dayKey(from), viaTo: dayKey(now) };
}

async function loadAdminLedelsesrapportDefaults() {
    setAdminReportDefaultsStatus('Indlæser...');
    try {
        const [response, resourcesResponse] = await Promise.all([
            fetch('/admin/ledelsesrapport-defaults', { headers: adminHeaders() }),
            fetch('/belastning/resources', { headers: adminHeaders() })
        ]);
        const [data, resourcesData] = await Promise.all([response.json(), resourcesResponse.json()]);
        if (!response.ok || !data.ok) throw new Error(data.error || 'Kunne ikke hente standarden');
        if (!resourcesResponse.ok || !resourcesData.ok) throw new Error(resourcesData.error || 'Kunne ikke hente ressourcer');
        const resources = [...new Map((resourcesData.resources || []).map(item => [String(item.MainR7 || '').trim(), item])).values()].filter(item => String(item.MainR7 || '').trim());
        document.getElementById('adminReportLoadResources').innerHTML = resources.map(item => '<div class="admin-report-resource-row" data-resource="' + escapeHtml(item.MainR7) + '" ondragover="dragAdminReportResource(event)"><span class="admin-report-resource-handle" draggable="true" ondragstart="startAdminReportResourceDrag(event)" ondragend="endAdminReportResourceDrag(event)" title="Træk for at flytte" aria-label="Træk ' + escapeHtml(item.R7Nm || item.MainR7) + '">⋮⋮</span><label><input type="checkbox" value="' + escapeHtml(item.MainR7) + '" onchange="refreshAdminReportResourcePagePreview()"> ' + escapeHtml(item.R7Nm || item.MainR7) + '</label><span class="admin-report-resource-page"></span><span class="admin-report-resource-actions"><button type="button" onclick="moveAdminReportResource(this,-1)" title="Flyt op" aria-label="Flyt ' + escapeHtml(item.R7Nm || item.MainR7) + ' op">↑</button><button type="button" onclick="moveAdminReportResource(this,1)" title="Flyt ned" aria-label="Flyt ' + escapeHtml(item.R7Nm || item.MainR7) + ' ned">↓</button></span></div>').join('');
        setAdminReportDefaults(data.defaults || getAdminReportSuggestedDefaults());
        setAdminReportDefaultsStatus(data.defaults ? 'Standarden er indlæst.' : 'Ingen gemt standard. Rapportens dynamiske forslag bruges.');
    } catch (error) {
        setAdminReportDefaultsStatus(String(error.message || error), true);
    }
}

async function saveAdminLedelsesrapportDefaults() {
    const defaults = Object.fromEntries(Object.entries(ADMIN_REPORT_FIELDS).map(([key, id]) => [key, document.getElementById(id).value]));
    defaults.loadResources = Array.from(document.querySelectorAll('#adminReportLoadResources input:checked')).map(input => input.value);
    defaults.orderLines = Array.from(document.querySelectorAll('#adminReportOrderLines input:checked')).map(input => input.value);
    if (Object.entries(defaults).some(([key, value]) => !value && !(['viaFrom', 'viaTo'].includes(key) && defaults.viaPeriod === 'all'))) {
        setAdminReportDefaultsStatus('Udfyld alle relevante standardfelter.', true);
        return;
    }
    if (!defaults.loadResources.length) {
        setAdminReportDefaultsStatus('Vælg mindst én ressource til Belastning.', true);
        return;
    }
    setAdminReportDefaultsStatus('Gemmer...');
    try {
        const response = await fetch('/admin/ledelsesrapport-defaults', { method: 'POST', headers: adminHeaders(), body: JSON.stringify(defaults) });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || 'Kunne ikke gemme standarden');
        setAdminReportDefaults(data.defaults);
        setAdminReportDefaultsStatus('Standarden er gemt i GOH-databasen.');
    } catch (error) {
        setAdminReportDefaultsStatus(String(error.message || error), true);
    }
}

async function loadAdminUsers() {
    const target = document.getElementById('adminUsers');
    if (!target) return;
    target.textContent = 'Indlæser brugere...';
    try {
        const response = await fetch('/admin/users', { headers: adminHeaders() });
        const data = await response.json();
        if (!response.ok || data.error) throw new Error(data.error || 'Kunne ikke hente brugere');
        const users = Array.isArray(data.users) ? data.users : [];
        target.innerHTML = users.map(user => {
            const permissions = user.permissions || {};
            const checks = user.role === 'superadmin'
                ? '<span>Superadmin har adgang til alle moduler.</span>'
                : ADMIN_MODULES.map(([key, label]) => '<label><input type="checkbox" data-permission="' + key + '" ' + (permissions[key] ? 'checked' : '') + '> ' + label + '</label>').join('');
            return '<div class="admin-user" data-username="' + escapeHtml(String(user.username)) + '">'
                + '<div class="admin-user-head"><div><span class="admin-user-name">' + escapeHtml(String(user.displayName || user.username)) + '</span> <small>(' + escapeHtml(String(user.username)) + ')</small></div>'
                + '<div style="display:flex;gap:12px;align-items:center;">'
                + (user.username === 'admin' ? '' : '<label><input type="checkbox" data-superadmin ' + (user.role === 'superadmin' ? 'checked' : '') + '> Superadmin</label>')
                + '<label><input type="checkbox" data-active ' + (user.active ? 'checked' : '') + '> Aktiv</label></div></div>'
                + '<div class="admin-permissions">' + checks + '</div>'
                + '<div class="admin-user-actions">'
                + (user.username === 'admin' ? '<span>Bootstrap superadmin</span>' : '<input type="password" data-new-password placeholder="Ny kode (valgfri)" style="padding:5px 7px;border:1px solid #c7d7ea;border-radius:5px;">'
                    + '<button type="button" data-admin-username="' + escapeHtml(String(user.username)) + '" onclick="saveAdminUserFromElement(this)">Gem</button>'
                    + '<button type="button" class="danger" data-admin-username="' + escapeHtml(String(user.username)) + '" onclick="deleteAdminUserFromElement(this)">Slet bruger</button>')
                + '</div></div>';
        }).join('') || '<div>Ingen brugere.</div>';
    } catch (error) {
        target.innerHTML = '<div class="error">' + escapeHtml(String(error.message || error)) + '</div>';
    }
}

async function createAdminUser() {
    const username = String((document.getElementById('adminNewUsername') || {}).value || '').trim();
    const displayName = String((document.getElementById('adminNewDisplayName') || {}).value || '').trim();
    const password = String((document.getElementById('adminNewPassword') || {}).value || '');
    try {
        const response = await fetch('/admin/users', { method: 'POST', headers: adminHeaders(), body: JSON.stringify({ username, displayName, password }) });
        const data = await response.json();
        if (!response.ok || data.error) throw new Error(data.error || 'Kunne ikke oprette bruger');
        document.getElementById('adminNewUsername').value = '';
        document.getElementById('adminNewDisplayName').value = '';
        document.getElementById('adminNewPassword').value = '';
        setAdminStatus('Bruger oprettet. Vælg moduler og gem rettigheder.');
        loadAdminUsers();
    } catch (error) {
        setAdminStatus(String(error.message || error), true);
    }
}

function saveAdminUserFromElement(element) {
    saveAdminUser(String(element && element.dataset && element.dataset.adminUsername || ''));
}

async function saveAdminUser(username) {
    const row = Array.from(document.querySelectorAll('.admin-user')).find(item => item.dataset.username === username);
    if (!row) return;
    const permissions = {};
    row.querySelectorAll('[data-permission]').forEach(input => { permissions[input.dataset.permission] = input.checked; });
    const activeInput = row.querySelector('[data-active]');
    const passwordInput = row.querySelector('[data-new-password]');
    const superadminInput = row.querySelector('[data-superadmin]');
    const role = superadminInput && superadminInput.checked ? 'superadmin' : 'user';
    try {
        const response = await fetch('/admin/users/' + encodeURIComponent(username), {
            method: 'PUT', headers: adminHeaders(), body: JSON.stringify({ active: !!(activeInput && activeInput.checked), role, permissions, password: String(passwordInput && passwordInput.value || '') })
        });
        const data = await response.json();
        if (!response.ok || data.error) throw new Error(data.error || 'Kunne ikke gemme bruger');
        setAdminStatus('Rettigheder gemt for ' + username + '. Rolleændringer gælder efter næste login.');
        loadAdminUsers();
    } catch (error) {
        setAdminStatus(String(error.message || error), true);
    }
}

function deleteAdminUserFromElement(element) {
    deleteAdminUser(String(element && element.dataset && element.dataset.adminUsername || ''));
}

async function deleteAdminUser(username) {
    if (!username || !confirm('Slet bruger ' + username + '?')) return;
    try {
        const response = await fetch('/admin/users/' + encodeURIComponent(username) + '/delete', { method: 'POST', headers: adminHeaders() });
        const raw = await response.text();
        let data = {};
        try { data = raw ? JSON.parse(raw) : {}; } catch (_) {
            throw new Error('Serverfejl HTTP ' + response.status + ': ' + raw.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 180));
        }
        if (!response.ok || data.error) throw new Error(data.error || 'Kunne ikke slette bruger');
        setAdminStatus('Bruger ' + username + ' er slettet.');
        loadAdminUsers();
    } catch (error) {
        setAdminStatus(String(error.message || error), true);
    }
}
