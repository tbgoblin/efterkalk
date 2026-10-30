// ── Note ordine / "ORDER NOTES" · client ─────────────────────────────────────
// Estratto verbatim dall'inline script di server.js. Definisce anche
// escapeHtmlFE, usata come globale condiviso da molte altre sezioni estratte
// (Omsaetning, Ordreindgang, Belastning, ecc.) — per questo va caricata prima
// di quelle. Usa a sua volta i globali: pushModalStack/removeModalStack,
// isOrderMarkedUB (dichiarati nell'inline script).
// ── ORDER NOTES ────────────────────────────────────────────────
let orderNotesCache = {};  // ordNo(string) -> { status, text, updatedAt }

async function loadAllNotes() {
    try {
        const r = await fetch('/order-notes-all');
        if (r.ok) {
            orderNotesCache = await r.json();
            scheduleDashboardWidgets();
        }
    } catch {}
}

async function loadOrderNote(ordNo) {
    const numericOrdNo = Number(ordNo || 0);
    if (!numericOrdNo) return null;
    try {
        const r = await fetch('/order-note/' + numericOrdNo);
        if (!r.ok) return null;
        const note = await r.json();
        orderNotesCache[String(numericOrdNo)] = note || { status: '', text: '', updatedAt: null };
        scheduleDashboardWidgets();
        renderOrderNoteBanner(numericOrdNo);
        updateOrderNoteCell(numericOrdNo);
        return note;
    } catch {
        return null;
    }
}

function getOrderNoteHtml(ordNo) {
    const note = orderNotesCache[String(ordNo)];
    if (!note || (!note.status && !note.text && !note.isCreditNote && !note.isUB)) return '<span style="color:#bbb;font-size:12px;">-</span>';
    const icons = { ok: '✅', error: '❌', check: '⚠️', credit: '🧾' };
    const icon = note.isCreditNote ? icons.credit : (icons[note.status] || '📝');
    const cls = note.isCreditNote ? 'credit' : (note.status || 'text');
    const creditLabel = note.isUB ? ' U/B' : (note.isCreditNote ? ' Kreditnota' : '');
    const preview = note.text ? escapeHtmlFE(note.text.slice(0, 40)) + (note.text.length > 40 ? '…' : '') : '';
    return '<span class="note-badge ' + cls + '" onclick="event.stopPropagation();openNotePopup(' + Number(ordNo) + ')">'
        + icon + creditLabel + (preview ? ' ' + preview : '') + '</span>';
}

function isOrderMarkedCreditNote(ordNo) {
    const note = orderNotesCache[String(ordNo)];
    return Boolean(note && note.isCreditNote === true);
}

function isOrderMarkedUB(ordNo) {
    const note = orderNotesCache[String(ordNo)];
    return Boolean(note && note.isUB === true);
}

function escapeHtmlFE(s) {
    return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function openNotePopup(ordNo, fromOrderDetail = false) {
    const note = orderNotesCache[String(ordNo)] || { status: '', text: '' };
    const existing = document.getElementById('notePopupOverlay');
    if (existing) existing.remove();

    const overlay = document.createElement('div');
    overlay.id = 'notePopupOverlay';
    overlay.className = 'note-popup-overlay';
    overlay.innerHTML =
        '<div class="note-popup">' +
        '<h3>📝 Note for ordre <strong>' + ordNo + '</strong></h3>' +
        '<label>Status</label>' +
        '<select id="noteStatusSel">' +
        '<option value="">— ingen status —</option>' +
        '<option value="ok">✅ OK</option>' +
        '<option value="error">❌ Fejl</option>' +
        '<option value="check">⚠️ Tjek</option>' +
        '</select>' +
        '<label style="display:flex;align-items:center;gap:8px;margin:-2px 0 10px 0;font-weight:600;">' +
        '<input id="noteCreditChk" type="checkbox" ' + (note.isCreditNote ? 'checked' : '') + ' style="width:16px;height:16px;" />' +
        'Kreditnota' +
        '</label>' +
        '<label id="noteUBRow" style="display:flex;align-items:center;gap:8px;margin:0 0 12px;font-weight:600;">' +
        '<input id="noteUBChk" type="checkbox" ' + (note.isUB ? 'checked' : '') + ' style="width:15px;height:15px;" />' +
        'U/B — uden beregning (vises separat fra faktureret)' +
        '</label>' +
        '<label>Note</label>' +
        '<textarea id="noteTextArea" placeholder="Skriv en note til denne ordre...">' + escapeHtmlFE(note.text || '') + '</textarea>' +
        (note.updatedAt ? '<div style="font-size:11px;color:#888;margin-bottom:10px;">Sidst opdateret: ' + note.updatedAt.slice(0,16).replace('T',' ') + '</div>' : '') +
        '<div class="note-popup-actions">' +
        '<button class="btn-note-delete" onclick="deleteOrderNote(' + ordNo + ',' + fromOrderDetail + ')">Slet</button>' +
        '<button class="btn-note-cancel" onclick="document.getElementById(\'notePopupOverlay\').remove()">Annuller</button>' +
        '<button class="btn-note-save" onclick="saveOrderNote(' + ordNo + ',' + fromOrderDetail + ')">Gem</button>' +
        '</div></div>';

    document.body.appendChild(overlay);
    overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
    document.getElementById('noteStatusSel').value = note.status || '';
}

async function saveOrderNote(ordNo, fromOrderDetail) {
    const status = document.getElementById('noteStatusSel').value;
    const text = document.getElementById('noteTextArea').value.trim();
    const isCreditNote = Boolean(document.getElementById('noteCreditChk') && document.getElementById('noteCreditChk').checked);
    const isUB = Boolean(document.getElementById('noteUBChk') && document.getElementById('noteUBChk').checked);
    try {
        const r = await fetch('/order-note/' + ordNo, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ status, text, isCreditNote, isUB })
        });
        if (r.ok) {
            const note = await r.json();
            orderNotesCache[String(ordNo)] = note;
            scheduleDashboardWidgets();
            if (typeof refreshKfOrderClassification === 'function') refreshKfOrderClassification(ordNo, note);
        }
    } catch {}
    document.getElementById('notePopupOverlay').remove();
    updateOrderNoteCell(ordNo);
    if (fromOrderDetail) renderOrderNoteBanner(ordNo);
    if (orderListVisible) renderOrderList();
}

async function deleteOrderNote(ordNo, fromOrderDetail) {
    try {
        await fetch('/order-note/' + ordNo, { method: 'DELETE' });
        delete orderNotesCache[String(ordNo)];
        scheduleDashboardWidgets();
        if (typeof refreshKfOrderClassification === 'function') refreshKfOrderClassification(ordNo, null);
    } catch {}
    document.getElementById('notePopupOverlay').remove();
    updateOrderNoteCell(ordNo);
    if (fromOrderDetail) renderOrderNoteBanner(ordNo);
    if (orderListVisible) renderOrderList();
}

function updateOrderNoteCell(ordNo) {
    const listEl = document.getElementById('orderList');
    if (!listEl) return;
    const cells = listEl.querySelectorAll('.order-note-cell[data-ordno="' + ordNo + '"]');
    const html = getOrderNoteHtml(ordNo);
    for (const cell of cells) { cell.innerHTML = html; }
    updateOrderListSummaryPanel();
}

function renderOrderNoteBanner(ordNo) {
    const el = document.getElementById('order-note-banner-' + ordNo);
    if (!el) return;
    const note = orderNotesCache[String(ordNo)];
    if (!note || (!note.status && !note.text && !note.isCreditNote && !note.isUB)) { el.style.display = 'none'; return; }
    const icons = { ok: '✅', error: '❌', check: '⚠️', credit: '🧾' };
    const icon = note.isCreditNote ? icons.credit : (icons[note.status] || '📝');
    const cls = note.isCreditNote ? 'credit' : (note.status || 'text');
    el.className = 'order-note-banner ' + cls;
    el.style.display = 'flex';
    const label = note.isUB ? 'U/B — uden beregning' : note.isCreditNote
        ? 'Kreditnota'
        : (note.status === 'ok' ? 'OK' : note.status === 'error' ? 'Fejl' : note.status === 'check' ? 'Tjek' : 'Note');
    el.innerHTML = '<span class="note-icon">' + icon + '</span><div class="note-body"><strong>' +
        label +
        '</strong>' + (note.text ? ': ' + escapeHtmlFE(note.text) : '') + '</div>' +
        '<span style="font-size:11px;opacity:0.7;margin-left:auto;cursor:pointer;" onclick="openNotePopup(' + ordNo + ',true)">✏️ Rediger</span>';
}
