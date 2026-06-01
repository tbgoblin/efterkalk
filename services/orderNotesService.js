/**
 * orderNotesService.js
 * Persistent notes per order. Stored in order_notes.json next to the executable.
 * Never deleted by "Ryd cache" — survives all cache clears.
 *
 * Schema: { "403668": { status: "ok"|"error"|"check"|"", text: "...", isCreditNote: bool, updatedAt: "ISO8601" } }
 */
const fs = require('fs');
const path = require('path');

const NOTES_FILE = path.join(require('process').env.PORTABLE_EXECUTABLE_DIR || __dirname, '..', 'order_notes.json');

let _notes = null;

function _load() {
    if (_notes !== null) return;
    try {
        if (fs.existsSync(NOTES_FILE)) {
            const raw = fs.readFileSync(NOTES_FILE, 'utf8');
            _notes = JSON.parse(raw);
        } else {
            _notes = {};
        }
    } catch {
        _notes = {};
    }
}

function _save() {
    try {
        fs.writeFileSync(NOTES_FILE, JSON.stringify(_notes, null, 2), 'utf8');
    } catch (err) {
        console.error('[orderNotes] save error:', err.message);
    }
}

function getNote(ordNo) {
    _load();
    return _notes[String(ordNo)] || null;
}

function getAllNotes() {
    _load();
    return { ..._notes };
}

function setNote(ordNo, { status = '', text = '', isCreditNote = false } = {}) {
    _load();
    const key = String(ordNo);
    const creditFlag = Boolean(isCreditNote);
    if (!status && !text.trim() && !creditFlag) {
        delete _notes[key];
    } else {
        _notes[key] = {
            status: status || '',
            text: String(text || '').slice(0, 2000),
            isCreditNote: creditFlag,
            updatedAt: new Date().toISOString()
        };
    }
    _save();
    return _notes[key] || null;
}

function deleteNote(ordNo) {
    _load();
    delete _notes[String(ordNo)];
    _save();
}

module.exports = { getNote, getAllNotes, setNote, deleteNote };
