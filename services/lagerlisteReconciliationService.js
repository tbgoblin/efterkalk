const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function resolveBaseDir() {
    const explicitDir = String(process.env.GANTECH_NOTES_DIR || '').trim();
    if (explicitDir) return explicitDir;
    const localAppData = String(process.env.LOCALAPPDATA || '').trim();
    if (localAppData) return path.join(localAppData, 'Gantech Efterkalk');
    const portableDir = String(process.env.PORTABLE_EXECUTABLE_DIR || '').trim();
    return portableDir || path.join(__dirname, '..');
}

const file = path.join(resolveBaseDir(), 'lagerliste_reconciliations.json');
let entries = null;

function load() {
    if (entries !== null) return;
    try {
        entries = fs.existsSync(file) ? JSON.parse(fs.readFileSync(file, 'utf8')) : [];
        if (!Array.isArray(entries)) entries = [];
    } catch {
        entries = [];
    }
}

function save() {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(entries, null, 2), 'utf8');
}

function list(periodKey) {
    load();
    return entries.filter(entry => entry.periodKey === String(periodKey || ''));
}

function add({ periodKey, productKey, from, to, amount, note, createdBy }) {
    load();
    const numericAmount = Number(amount || 0);
    const safePeriodKey = String(periodKey || '').trim();
    const safeNote = String(note || '').trim().slice(0, 2000);
    if (!safePeriodKey || !Number.isFinite(numericAmount) || numericAmount <= 0 || !safeNote) {
        throw new Error('Periode, positivt beløb og forklarende note kræves');
    }
    const entry = {
        id: crypto.randomUUID(),
        periodKey: safePeriodKey,
        productKey: String(productKey || '').trim(),
        from: String(from || '').trim().slice(0, 300),
        to: String(to || '').trim().slice(0, 300),
        amount: Math.round(numericAmount * 100) / 100,
        note: safeNote,
        createdBy: String(createdBy || '').trim().slice(0, 100),
        createdAt: new Date().toISOString()
    };
    entries.push(entry);
    save();
    return entry;
}

function remove(id) {
    load();
    const before = entries.length;
    entries = entries.filter(entry => entry.id !== String(id || ''));
    if (entries.length === before) return false;
    save();
    return true;
}

module.exports = { list, add, remove };
