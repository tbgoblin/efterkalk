/**
 * settingsService.js
 * Persistent app settings stored in LOCALAPPDATA\Gantech Efterkalk\settings.json.
 * Manages database connection profiles and the active profile selection.
 * Never deleted by "Ryd cache" — survives all cache clears.
 *
 * Schema:
 * {
 *   activeProfileId: "production",
 *   profiles: [
 *     { id: "production", label: "Produktion", server: "10.2.0.3\\VISMA", database: "F0001", readOnly: false },
 *     { id: "test",       label: "Test",       server: "10.2.0.3\\VISMA", database: "F0001_TEST", readOnly: false }
 *   ]
 * }
 */
const fs   = require('fs');
const path = require('path');

function resolveSettingsBaseDir() {
    const explicit = String(process.env.GANTECH_NOTES_DIR || '').trim();
    if (explicit) return explicit;
    const localAppData = String(process.env.LOCALAPPDATA || '').trim();
    if (localAppData) return path.join(localAppData, 'Gantech Efterkalk');
    const portableDir = String(process.env.PORTABLE_EXECUTABLE_DIR || '').trim();
    if (portableDir) return portableDir;
    return path.join(__dirname, '..');
}

function resolveSettingsFile() {
    return path.join(resolveSettingsBaseDir(), 'settings.json');
}

function ensureSettingsDir(settingsFile) {
    try { fs.mkdirSync(path.dirname(settingsFile), { recursive: true }); } catch { /* ignore */ }
}

// Default built-in profiles (always present; user can add more)
const DEFAULT_PROFILES = [
    {
        id:       'production',
        label:    'Produktion',
        server:   '10.2.0.3\\VISMA',
        database: 'F0001',
        readOnly: false,
        isDefault: true
    }
];

let _settings = null;

function _load() {
    if (_settings !== null) return;
    const file = resolveSettingsFile();
    try {
        if (fs.existsSync(file)) {
            const raw = fs.readFileSync(file, 'utf8');
            _settings = JSON.parse(raw);
        } else {
            _settings = {};
        }
    } catch {
        _settings = {};
    }
    // Ensure minimal shape
    if (!Array.isArray(_settings.profiles)) _settings.profiles = [];
    if (!_settings.activeProfileId)         _settings.activeProfileId = 'production';
}

function _save() {
    const file = resolveSettingsFile();
    try {
        ensureSettingsDir(file);
        fs.writeFileSync(file, JSON.stringify(_settings, null, 2), 'utf8');
    } catch (err) {
        console.error('[settings] save error:', err.message);
    }
}

// ── Public API ────────────────────────────────────────────────────────────────

/** Returns all profiles (defaults merged with user-defined). */
function getAllProfiles() {
    _load();
    // Merge: start with defaults, then add or overwrite with user-saved profiles
    const map = new Map(DEFAULT_PROFILES.map(p => [p.id, { ...p }]));
    for (const p of _settings.profiles) {
        if (p && p.id) map.set(p.id, { ...p });
    }
    return Array.from(map.values());
}

/** Returns the currently active profile object. */
function getActiveProfile() {
    _load();
    const profiles = getAllProfiles();
    const active = profiles.find(p => p.id === _settings.activeProfileId);
    return active || profiles.find(p => p.id === 'production') || profiles[0] || DEFAULT_PROFILES[0];
}

/** Switches the active profile by id. Returns the new active profile. */
function setActiveProfile(profileId) {
    _load();
    const profiles = getAllProfiles();
    const found = profiles.find(p => p.id === String(profileId || '').trim());
    if (!found) throw new Error('Profil ikke fundet: ' + profileId);
    _settings.activeProfileId = found.id;
    _save();
    return found;
}

/** Creates or updates a user-defined profile. `id` must not be 'production'. */
function upsertProfile(profile) {
    _load();
    if (!profile || !profile.id) throw new Error('Profil mangler id');
    const id = String(profile.id).trim();
    if (id === 'production') throw new Error('Produktions-profilen kan ikke ændres');
    const validated = {
        id,
        label:    String(profile.label    || id).trim().slice(0, 60),
        server:   String(profile.server   || '').trim(),
        database: String(profile.database || '').trim(),
        readOnly: Boolean(profile.readOnly)
    };
    if (!validated.server || !validated.database) throw new Error('server og database er påkrævet');
    const idx = _settings.profiles.findIndex(p => p.id === id);
    if (idx >= 0) _settings.profiles[idx] = validated;
    else          _settings.profiles.push(validated);
    _save();
    return validated;
}

/** Deletes a user-defined profile. Cannot delete the active or built-in profiles. */
function deleteProfile(profileId) {
    _load();
    const id = String(profileId || '').trim();
    if (id === 'production') throw new Error('Produktions-profilen kan ikke slettes');
    if (id === _settings.activeProfileId) throw new Error('Kan ikke slette den aktive profil');
    _settings.profiles = _settings.profiles.filter(p => p.id !== id);
    _save();
}

/** Returns settings summary for the API. */
function getSettingsSummary() {
    _load();
    return {
        activeProfileId: _settings.activeProfileId,
        activeProfile:   getActiveProfile(),
        profiles:        getAllProfiles()
    };
}

module.exports = { getAllProfiles, getActiveProfile, setActiveProfile, upsertProfile, deleteProfile, getSettingsSummary };
