// ── Personalehåndbog + QMS · client ─────────────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Usa i globali condivisi
// della pagina: escapeHtml, pushModalStack, removeModalStack.
const PH_BASE_URL = 'http://apv/GHB/';

function openPersonalehåndbog() {
    const modal = document.getElementById('personalehåndbogsModal');
    const iframe = document.getElementById('personalehåndbogsIframe');
    const input = document.getElementById('personalehåndbogsSearchInput');
    if (!modal || !iframe) return;
    if (!iframe.src || iframe.src === 'about:blank' || iframe.src === window.location.href) {
        iframe.src = PH_BASE_URL;
    }
    modal.classList.add('open');
    phCheckStatus();
    pushModalStack('personalehåndbogsModal');
    if (input) setTimeout(() => input.focus(), 150);
}

function closePersonalehåndbog() {
    const modal = document.getElementById('personalehåndbogsModal');
    const iframe = document.getElementById('personalehåndbogsIframe');
    if (modal) modal.classList.remove('open');
    if (iframe) iframe.src = '';
    removeModalStack('personalehåndbogsModal');
}

let qmsDataset = null;
let qmsFlatDocs = [];
let qmsSelectedDocId = null;
let qmsEditMode = false;

function makeQmsId(prefix) {
    return String(prefix || 'id') + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
}

async function loadQmsDataset(force = false) {
    if (qmsDataset && !force) return qmsDataset;
    const r = await fetch('/qms/dataset');
    const data = await r.json();
    if (!data.ok || !data.dataset) throw new Error(data.error || 'QMS dataset fejl');
    qmsDataset = data.dataset;
    return qmsDataset;
}

async function saveQmsDataset() {
    const r = await fetch('/qms/dataset', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ dataset: qmsDataset })
    });
    const data = await r.json();
    if (!data.ok) throw new Error(data.error || 'Kunne ikke gemme dataset');
    qmsDataset = data.dataset;
    return qmsDataset;
}

function flattenQmsDataset() {
    const out = [];
    if (!qmsDataset || !Array.isArray(qmsDataset.folders)) return out;
    for (const folder of qmsDataset.folders) {
        const docs = Array.isArray(folder.documents) ? folder.documents : [];
        for (const doc of docs) {
            out.push({
                folderId: folder.id,
                folderName: folder.name,
                folderDescription: folder.description || '',
                id: doc.id,
                title: doc.title,
                url: doc.url || '',
                content: doc.content || '',
                tags: Array.isArray(doc.tags) ? doc.tags : []
            });
        }
    }
    qmsFlatDocs = out;
    return out;
}

function getSelectedQmsDoc() {
    return qmsFlatDocs.find(d => d.id === qmsSelectedDocId) || null;
}

function renderQmsView(doc) {
    const view = document.getElementById('qmsView');
    if (!view) return;
    if (!doc) {
        view.innerHTML = '<h3>Kvalitetsledelsessystem</h3><div class="qms-view-meta">Vælg et dokument i venstre side.</div>';
        return;
    }
    if (qmsEditMode) {
        view.innerHTML = ''
            + '<h3>Rediger dokument</h3>'
            + '<div class="qms-view-meta">' + escapeHtml(doc.folderName) + '</div>'
            + '<div class="qms-editor">'
            + '<label>Titel</label><input id="qmsEditTitle" value="' + escapeHtml(doc.title) + '" />'
            + '<label>URL (valgfri)</label><input id="qmsEditUrl" value="' + escapeHtml(doc.url) + '" />'
            + '<label>Indhold</label><textarea id="qmsEditContent">' + escapeHtml(doc.content) + '</textarea>'
            + '<div class="qms-editor-actions">'
            + '<button class="save" onclick="qmsSaveCurrentDoc()">Gem dokument</button>'
            + '<button class="delete" onclick="qmsDeleteCurrentDoc()">Slet dokument</button>'
            + '<button class="cancel" onclick="toggleQmsEditMode(false)">Afslut redigering</button>'
            + '</div>'
            + '</div>';
        return;
    }
    view.innerHTML = ''
        + '<h3>' + escapeHtml(doc.title) + '</h3>'
        + '<div class="qms-view-meta">' + escapeHtml(doc.folderName) + '</div>'
        + '<div class="qms-view-content">' + escapeHtml(doc.content) + '</div>'
        + (doc.url ? '<div class="qms-view-link"><a href="' + escapeHtml(doc.url) + '" target="_blank" rel="noopener noreferrer">Åbn original reference</a></div>' : '');
}

function renderQmsList(query = '') {
    const list = document.getElementById('qmsList');
    const label = document.getElementById('qmsListLabel');
    if (!list || !label) return;
    const q = String(query || '').trim().toLowerCase();
    const docs = flattenQmsDataset().filter(doc => {
        if (!q) return true;
        return (doc.title + ' ' + doc.folderName + ' ' + doc.content + ' ' + doc.tags.join(' ')).toLowerCase().includes(q);
    });
    label.textContent = docs.length + ' dokumenter';
    if (docs.length === 0) {
        list.innerHTML = '<div class="qms-empty">Ingen dokumenter matcher din søgning.</div>';
        renderQmsView(null);
        return;
    }
    list.innerHTML = docs.map(doc => (
        '<div class="qms-item" data-doc-id="' + escapeHtml(doc.id) + '" onclick="openQmsPage(this)">' +
        '<div class="qms-item-title">' + escapeHtml(doc.title) + '</div>' +
        '<div class="qms-item-meta">' + escapeHtml(doc.folderName) + '</div>' +
        '</div>'
    )).join('');
    if (!qmsSelectedDocId || !docs.some(d => d.id === qmsSelectedDocId)) {
        qmsSelectedDocId = docs[0].id;
    }
    const active = list.querySelector('.qms-item[data-doc-id="' + CSS.escape(qmsSelectedDocId) + '"]') || list.querySelector('.qms-item');
    if (active) openQmsPage(active);
}

async function openKvalitetsledelsessystem() {
    const modal = document.getElementById('qmsModal');
    const input = document.getElementById('qmsSearchInput');
    if (!modal) return;
    modal.classList.add('open');
    pushModalStack('qmsModal');
    try {
        await loadQmsDataset(false);
        renderQmsList('');
    } catch (err) {
        const list = document.getElementById('qmsList');
        if (list) list.innerHTML = '<div class="qms-empty">Kunne ikke læse QMS dataset: ' + escapeHtml(err.message || '') + '</div>';
        renderQmsView(null);
    }
    if (input) {
        input.value = '';
        setTimeout(() => input.focus(), 120);
    }
}

function closeQmsModal() {
    const modal = document.getElementById('qmsModal');
    if (modal) modal.classList.remove('open');
    removeModalStack('qmsModal');
}

function searchQmsPages() {
    const input = document.getElementById('qmsSearchInput');
    renderQmsList(input ? input.value : '');
}

function openQmsPage(el) {
    const docId = el && el.getAttribute ? el.getAttribute('data-doc-id') : '';
    if (!docId) return;
    qmsSelectedDocId = docId;
    document.querySelectorAll('#qmsList .qms-item').forEach(x => x.classList.remove('active'));
    el.classList.add('active');
    renderQmsView(getSelectedQmsDoc());
}

function toggleQmsEditMode(force) {
    if (typeof force === 'boolean') {
        qmsEditMode = force;
    } else {
        qmsEditMode = !qmsEditMode;
    }
    const btn = document.getElementById('qmsEditToggleBtn');
    if (btn) btn.textContent = qmsEditMode ? 'Visning' : 'Rediger';
    renderQmsView(getSelectedQmsDoc());
}

async function qmsSaveCurrentDoc() {
    const doc = getSelectedQmsDoc();
    if (!doc) return;
    const title = document.getElementById('qmsEditTitle');
    const url = document.getElementById('qmsEditUrl');
    const content = document.getElementById('qmsEditContent');
    const folder = (qmsDataset.folders || []).find(f => f.id === doc.folderId);
    if (!folder) return;
    const target = (folder.documents || []).find(d => d.id === doc.id);
    if (!target) return;
    target.title = String(title && title.value || '').trim() || target.title;
    target.url = String(url && url.value || '').trim();
    target.content = String(content && content.value || '').trim();
    try {
        await saveQmsDataset();
        renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
    } catch (err) {
        alert('Kunne ikke gemme: ' + (err.message || err));
    }
}

async function qmsDeleteCurrentDoc() {
    const doc = getSelectedQmsDoc();
    if (!doc) return;
    if (!confirm('Slet dokumentet "' + doc.title + '"?')) return;
    const folder = (qmsDataset.folders || []).find(f => f.id === doc.folderId);
    if (!folder) return;
    folder.documents = (folder.documents || []).filter(d => d.id !== doc.id);
    qmsSelectedDocId = null;
    try {
        await saveQmsDataset();
        renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
    } catch (err) {
        alert('Kunne ikke slette: ' + (err.message || err));
    }
}

async function qmsCreateFolder() {
    try {
        await loadQmsDataset(false);
        const name = prompt('Navn på ny mappe:');
        if (!name || !name.trim()) return;
        qmsDataset.folders.push({
            id: makeQmsId('folder'),
            name: name.trim(),
            description: '',
            documents: []
        });
        await saveQmsDataset();
        renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
    } catch (err) {
        alert('Kunne ikke oprette mappe: ' + (err.message || err));
    }
}

async function qmsCreateDocument() {
    try {
        await loadQmsDataset(false);
        if (!Array.isArray(qmsDataset.folders) || qmsDataset.folders.length === 0) {
            alert('Opret først en mappe.');
            return;
        }
        const title = prompt('Titel på nyt dokument:');
        if (!title || !title.trim()) return;
        let folder = (qmsDataset.folders || []).find(f => f.id === qmsSelectedDocId) || null;
        const selected = getSelectedQmsDoc();
        if (selected) {
            folder = (qmsDataset.folders || []).find(f => f.id === selected.folderId) || null;
        }
        if (!folder) folder = qmsDataset.folders[0];
        folder.documents = Array.isArray(folder.documents) ? folder.documents : [];
        const docId = makeQmsId('doc');
        folder.documents.push({
            id: docId,
            title: title.trim(),
            url: '',
            content: '',
            tags: []
        });
        qmsSelectedDocId = docId;
        await saveQmsDataset();
        renderQmsList(document.getElementById('qmsSearchInput')?.value || '');
        qmsEditMode = true;
        toggleQmsEditMode(true);
    } catch (err) {
        alert('Kunne ikke oprette dokument: ' + (err.message || err));
    }
}

function phSetStatus(msg) {
    const lbl = document.getElementById('phResultsLabel');
    const msgEl = document.getElementById('phStatusMsg');
    const list = document.getElementById('phResultsList');
    if (list) list.innerHTML = '<div class="ph-status-msg" id="phStatusMsg">' + msg + '</div>';
    if (lbl) lbl.textContent = 'Resultater';
}

async function phCheckStatus() {
    try {
        const r = await fetch('/ph/status');
        const d = await r.json();
        if (d.status === 'indexing') {
            phSetStatus('⏳ Indekserer sitet, vent venligst…');
            setTimeout(phCheckStatus, 2000);
        } else if (d.status === 'ready') {
            phSetStatus('Skriv en søgning og tryk Søg.<br><small style="color:#9aabcc">' + d.count + ' sider indekseret</small>');
        } else if (d.status === 'idle') {
            phSetStatus('Indeks ikke klar. Tryk ↺ Genindekser.');
        } else if (d.status === 'error') {
            phSetStatus('⚠️ Fejl ved indeksering: ' + (d.error || ''));
        }
    } catch { phSetStatus('Kunne ikke kontakte serveren.'); }
}

async function phReindex() {
    phSetStatus('⏳ Indekserer sitet, vent venligst…');
    try {
        await fetch('/ph/reindex', { method: 'POST' });
        setTimeout(phCheckStatus, 1500);
    } catch { phSetStatus('⚠️ Fejl ved genindeksering.'); }
}

function phEscapeHtml(s) {
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function phHighlight(text, terms) {
    let out = phEscapeHtml(text);
    for (const t of terms) {
        if (!t) continue;
        const lower = out.toLowerCase();
        const tl = t.toLowerCase();
        let i = 0, result = '', pos;
        while ((pos = lower.indexOf(tl, i)) !== -1) {
            result += out.slice(i, pos) + '<mark>' + out.slice(pos, pos + t.length) + '</mark>';
            i = pos + t.length;
        }
        out = result + out.slice(i);
    }
    return out;
}

async function searchPersonalehåndbog() {
    const input = document.getElementById('personalehåndbogsSearchInput');
    const list  = document.getElementById('phResultsList');
    const lbl   = document.getElementById('phResultsLabel');
    if (!input || !list) return;
    const q = input.value.trim();
    if (!q) { phCheckStatus(); return; }
    phSetStatus('🔍 Søger…');
    try {
        const r = await fetch('/ph/search?q=' + encodeURIComponent(q));
        const d = await r.json();
        if (d.status === 'indexing') { phSetStatus('⏳ Indekserer endnu, prøv igen om lidt…'); return; }
        if (!d.results || d.results.length === 0) {
            phSetStatus('Ingen resultater for <strong>' + phEscapeHtml(q) + '</strong>.');
            if (lbl) lbl.textContent = '0 resultater';
            return;
        }
        if (lbl) lbl.textContent = d.results.length + ' resultat' + (d.results.length !== 1 ? 'er' : '');
        const terms = q.toLowerCase().split(/\s+/).filter(Boolean);
        list.innerHTML = d.results.map((res, i) => {
            const title = phHighlight(res.title || res.url, terms);
            const snip  = phHighlight(res.snippet, terms);
            const safeUrl = phEscapeHtml(res.url);
            return '<div class="ph-result-item" data-url="' + safeUrl + '" onclick="phOpenResult(this)" title="' + safeUrl + '">'
                + '<div class="ph-result-title">' + title + '</div>'
                + '<div class="ph-result-url">' + safeUrl + '</div>'
                + '<div class="ph-result-snippet">' + snip + '</div>'
                + '</div>';
        }).join('');
        // Auto-load first result
        const first = list.querySelector('.ph-result-item');
        if (first) phOpenResult(first);
    } catch { phSetStatus('⚠️ Søgefejl. Prøv igen.'); }
}

function phOpenResult(el) {
    const url = el.getAttribute('data-url');
    if (!url) return;
    const iframe = document.getElementById('personalehåndbogsIframe');
    if (iframe) iframe.src = url;
    document.querySelectorAll('.ph-result-item').forEach(e => e.classList.remove('ph-active'));
    el.classList.add('ph-active');
}
