function createDesktopCredentialStore({ fs, safeStorage, filePath, enabled }) {
    const available = () => enabled && safeStorage.isEncryptionAvailable();
    function read() {
        if (!available() || !fs.existsSync(filePath)) return null;
        const data = JSON.parse(safeStorage.decryptString(fs.readFileSync(filePath)));
        if (!data || typeof data.username !== 'string' || typeof data.password !== 'string') throw new Error('Ugyldigt gemt login');
        return data;
    }
    return {
        info() {
            if (!available()) return { available: false, saved: false };
            try {
                const data = read();
                return { available: true, saved: !!data, username: data?.username || '' };
            } catch (_) {
                return { available: true, saved: false, damaged: true };
            }
        },
        read,
        save(data) {
            if (!available()) throw new Error('Windows-kryptering er ikke tilgængelig');
            if (!data || typeof data.username !== 'string' || !data.username.trim() || data.username.length > 100
                || typeof data.password !== 'string' || !data.password || data.password.length > 1024) throw new Error('Ugyldigt login');
            const encrypted = safeStorage.encryptString(JSON.stringify({ username: data.username.trim(), password: data.password }));
            fs.writeFileSync(filePath, encrypted, { mode: 0o600 });
        },
        forget() {
            if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
        }
    };
}

module.exports = { createDesktopCredentialStore };