/**
 * orderNotesService.js
 * Persistent notes per order. Stored in order_notes.json in a stable user-writable path
 * (GANTECH_NOTES_DIR or LOCALAPPDATA\Gantech Efterkalk) with legacy migration.
 * Never deleted by "Ryd cache" — survives all cache clears.
 *
 * Schema: { "403668": { status: "ok"|"error"|"check"|"", text: "...", isCreditNote: bool, updatedAt: "ISO8601" } }
 */
const fs = require('fs');
const path = require('path');

function resolveLegacyNotesFile() {
    return path.join(require('process').env.PORTABLE_EXECUTABLE_DIR || __dirname, '..', 'order_notes.json');
}

function resolveNotesBaseDir() {
    const explicitDir = String(process.env.GANTECH_NOTES_DIR || '').trim();
    if (explicitDir) return explicitDir;

    const localAppData = String(process.env.LOCALAPPDATA || '').trim();
    if (localAppData) return path.join(localAppData, 'Gantech Efterkalk');

    const portableDir = String(process.env.PORTABLE_EXECUTABLE_DIR || '').trim();
    if (portableDir) return portableDir;

    return path.join(__dirname, '..');
}

function resolveNotesFile() {
    return path.join(resolveNotesBaseDir(), 'order_notes.json');
}

function ensureNotesDir(notesFile) {
    try {
        fs.mkdirSync(path.dirname(notesFile), { recursive: true });
    } catch {
        // Ignore directory create errors; save/load will handle failures.
    }
}

function migrateLegacyNotesIfNeeded(notesFile) {
    const legacyFile = resolveLegacyNotesFile();
    if (path.resolve(legacyFile) === path.resolve(notesFile)) return;
    if (!fs.existsSync(legacyFile) || fs.existsSync(notesFile)) return;
    try {
        ensureNotesDir(notesFile);
        fs.copyFileSync(legacyFile, notesFile);
    } catch {
        // Ignore migration failures and continue with normal load behavior.
    }
}

let _notes = null;

function _load() {
    if (_notes !== null) return;
    const notesFile = resolveNotesFile();
    migrateLegacyNotesIfNeeded(notesFile);
    try {
        if (fs.existsSync(notesFile)) {
            const raw = fs.readFileSync(notesFile, 'utf8');
            _notes = JSON.parse(raw);
        } else {
            _notes = {};
        }
    } catch {
        _notes = {};
    }
}

function _save() {
    const notesFile = resolveNotesFile();
    try {
        ensureNotesDir(notesFile);
        fs.writeFileSync(notesFile, JSON.stringify(_notes, null, 2), 'utf8');
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
