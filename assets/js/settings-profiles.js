// ── Settings: database profiler · client ─────────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, pushModalStack/removeModalStack, renderDashboardWidgets
// (dichiarati nell'inline script) e resetSalgordreVia (dichiarato in assets/js/via.js).

// ── Settings: database profiler ──────────────────────────────
let _settingsProfiles = [];
let _settingsActiveId = '';

function _updateDbBadge(profile) {
    renderDashboardWidgets();
    const badge = document.getElementById('dbProfileBadge');
    if (!badge) return;
    const id = String((profile && profile.id) || '');
    const label = String((profile && profile.label) || id || '?');
    const isDefault = id === 'production';
    badge.className = 'db-profile-badge ' + (isDefault ? 'prod' : (id === 'test' ? 'test' : 'other'));
    badge.title = (isDefault ? '⚠️ PRODUKTION' : '✅ TEST/ALT') + ' — ' + String((profile && profile.server) || '') + '/' + String((profile && profile.database) || '') + ' — klik for at ændre';
    badge.textContent = (isDefault ? '🔴' : '🟢') + ' ' + escapeHtml(label);
}

async function loadSettingsProfiles() {
    try {
        const r = await fetch('/settings/db-profiles');
        const d = await r.json();
        _settingsProfiles = d.profiles || [];
        _settingsActiveId = d.activeProfileId || 'production';
        _updateDbBadge(d.activeProfile);
        renderRestPrices(d.restPrices || {});
        renderSettingsProfileList();
        renderSettingsWarning(d.activeProfile);
    } catch (err) {
        console.warn('[settings] load failed:', err.message);
    }
}

function renderRestPrices(prices) {
    const form = document.getElementById('restPricesForm');
    if (!form) return;
    const labels = { '301': '301 Stålplader', '311': '311 Rustfri', '321': '321 Aluminium', '331': '331 Galvaniseret', '381': '381 Kobber/Messing', '302': '302 Andet' };
    form.innerHTML = Object.keys(labels).map(key => '<div><label>' + labels[key] + '</label><input type="number" min="0" step="0.01" data-rest-price="' + key + '" value="' + Number(prices[key] || 0) + '" /></div>').join('');
}

async function saveRestPrices() {
    const values = {};
    document.querySelectorAll('[data-rest-price]').forEach(input => { values[input.dataset.restPrice] = Number(input.value || 0); });
    const status = document.getElementById('restPricesStatus');
    try {
        const response = await fetch('/settings/rest-prices', { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + String(authToken || '') }, body: JSON.stringify({ restPrices: values }) });
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || ('HTTP ' + response.status));
        renderRestPrices(data.restPrices || values);
        if (status) { status.style.color = '#1b5e20'; status.textContent = 'Rest-Preise gemt.'; }
    } catch (err) {
        if (status) { status.style.color = '#b71c1c'; status.textContent = 'Fejl: ' + String(err.message || err); }
    }
}

function renderSettingsProfileList() {
    const list = document.getElementById('settingsProfileList');
    if (!list) return;
    if (!_settingsProfiles.length) { list.innerHTML = '<div style="color:#888;font-size:13px;">Ingen profiler</div>'; return; }
    list.innerHTML = _settingsProfiles.map(p => {
        const isActive = p.id === _settingsActiveId;
        return '<div class="settings-profile-card' + (isActive ? ' active' : '') + '">' +
            '<div class="pinfo">' +
            '<div class="pname">' + escapeHtml(p.label) + (p.readOnly ? ' <span style="font-size:10px;color:#e65100;font-weight:700;">[KUN LÆSNING]</span>' : '') + '</div>' +
            '<div class="pconn">' + escapeHtml(p.server) + ' / ' + escapeHtml(p.database) + '</div>' +
            '</div>' +
            (isActive ? '<span class="pactive">AKTIV</span>' : '') +
            '<button class="btn-select" ' + (isActive ? 'disabled' : 'onclick="switchDbProfile(\'' + p.id + '\')"') + '>' +
            (isActive ? 'Valgt' : 'Vælg') + '</button>' +
            (p.id !== 'production' ? '<button class="btn-del" onclick="editSettingsProfile(' + JSON.stringify(p).replace(/"/g, '&quot;') + ')">Rediger</button>' : '') +
            (p.id !== 'production' && !isActive ? '<button class="btn-del" onclick="deleteSettingsProfile(\'' + p.id + '\')">Slet</button>' : '') +
            '</div>';
    }).join('');
}

function renderSettingsWarning(profile) {
    const el = document.getElementById('settingsDbWarning');
    if (!el) return;
    const id = String((profile && profile.id) || '');
    if (id === 'production') {
        el.className = 'settings-db-warning show prod';
        el.textContent = '⚠️ PRODUKTION AKTIV — Alle skrivninger (opret produkt i Visma) går til det rigtige produktionssystem!';
    } else {
        el.className = 'settings-db-warning show test';
        el.textContent = '✅ TEST/ALTERNATIV DATABASE AKTIV (' + escapeHtml((profile && profile.label) || '') + ') — Sikkert at teste oprettelse af produkter.';
    }
}

async function switchDbProfile(profileId) {
    try {
        const r = await fetch('/settings/db-profiles/active', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ profileId })
        });
        const d = await r.json();
        if (!d.ok) throw new Error(d.error || 'Fejl');
        _settingsProfiles = d.profiles || [];
        _settingsActiveId = d.activeProfileId || profileId;
        resetSalgordreVia();
        _updateDbBadge(d.activeProfile);
        renderSettingsProfileList();
        renderSettingsWarning(d.activeProfile);
    } catch (err) {
        alert('Fejl ved skift: ' + err.message);
    }
}

function editSettingsProfile(profile) {
    const sfId       = document.getElementById('sfId');
    const sfLabel    = document.getElementById('sfLabel');
    const sfServer   = document.getElementById('sfServer');
    const sfDatabase = document.getElementById('sfDatabase');
    const sfReadOnly = document.getElementById('sfReadOnly');
    if (sfId)       sfId.value       = profile.id || '';
    if (sfLabel)    sfLabel.value    = profile.label || '';
    if (sfServer)   sfServer.value   = profile.server || '';
    if (sfDatabase) sfDatabase.value = profile.database || '';
    if (sfReadOnly) sfReadOnly.checked = Boolean(profile.readOnly);
}

async function deleteSettingsProfile(profileId) {
    if (!confirm('Slet profil "' + escapeHtml(profileId) + '"?')) return;
    try {
        const r = await fetch('/settings/db-profiles/' + encodeURIComponent(profileId), { method: 'DELETE' });
        const d = await r.json();
        if (!d.ok) throw new Error(d.error || 'Fejl');
        _settingsProfiles = d.profiles || [];
        _settingsActiveId = d.activeProfileId || _settingsActiveId;
        renderSettingsProfileList();
    } catch (err) {
        alert('Fejl ved sletning: ' + err.message);
    }
}

async function saveSettingsProfile() {
    const sfId       = document.getElementById('sfId');
    const sfLabel    = document.getElementById('sfLabel');
    const sfServer   = document.getElementById('sfServer');
    const sfDatabase = document.getElementById('sfDatabase');
    const sfReadOnly = document.getElementById('sfReadOnly');
    const status     = document.getElementById('settingsFormStatus');
    const profile = {
        id:       sfId ? sfId.value.trim().replace(/\s+/g,'_') : '',
        label:    sfLabel ? sfLabel.value.trim() : '',
        server:   sfServer ? sfServer.value.trim() : '',
        database: sfDatabase ? sfDatabase.value.trim() : '',
        readOnly: sfReadOnly ? sfReadOnly.checked : false
    };
    if (!profile.id || !profile.label || !profile.server || !profile.database) {
        if (status) status.textContent = 'Alle felter er påkrævet.';
        return;
    }
    try {
        const r = await fetch('/settings/db-profiles/upsert', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(profile)
        });
        const d = await r.json();
        if (!d.ok) throw new Error(d.error || 'Fejl');
        _settingsProfiles = d.profiles || [];
        _settingsActiveId = d.activeProfileId || _settingsActiveId;
        renderSettingsProfileList();
        clearSettingsForm();
        if (status) status.textContent = '';
    } catch (err) {
        if (status) status.textContent = 'Fejl: ' + err.message;
    }
}

function clearSettingsForm() {
    ['sfId','sfLabel','sfServer','sfDatabase'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    const sfReadOnly = document.getElementById('sfReadOnly');
    if (sfReadOnly) sfReadOnly.checked = false;
    const status = document.getElementById('settingsFormStatus');
    if (status) status.textContent = '';
}

function openSettingsModal() {
    const overlay = document.getElementById('settingsModalOverlay');
    if (overlay) overlay.classList.add('open');
    loadSettingsProfiles();
    pushModalStack('settingsModalOverlay');
}

function closeSettingsModal() {
    const overlay = document.getElementById('settingsModalOverlay');
    if (overlay) overlay.classList.remove('open');
    removeModalStack('settingsModalOverlay');
}
