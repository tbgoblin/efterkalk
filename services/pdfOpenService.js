const path = require('path');

function normalizeFileUrl(value) {
    let parsed;
    try {
        parsed = new URL(value);
    } catch (_) {
        throw Object.assign(new Error('Ugyldig PDF-sti.'), { statusCode: 400 });
    }

    let pathname;
    try {
        pathname = decodeURIComponent(parsed.pathname || '');
    } catch (_) {
        throw Object.assign(new Error('Ugyldig PDF-sti.'), { statusCode: 400 });
    }

    const windowsPath = pathname.replace(/\//g, '\\');
    if (parsed.hostname) return '\\\\' + parsed.hostname + windowsPath;
    return /^\\[A-Za-z]:\\/.test(windowsPath) ? windowsPath.slice(1) : windowsPath;
}

function normalizePdfTarget(rawTarget) {
    const value = String(rawTarget || '').trim();
    if (!value || /[\0\r\n]/.test(value)) {
        throw Object.assign(new Error('Ugyldig PDF-sti.'), { statusCode: 400 });
    }

    if (/^https?:\/\//i.test(value)) {
        let parsed;
        try {
            parsed = new URL(value);
        } catch (_) {
            throw Object.assign(new Error('Ugyldig PDF-sti.'), { statusCode: 400 });
        }
        if (!/^https?:$/.test(parsed.protocol) || !parsed.hostname || !/\.pdf$/i.test(parsed.pathname)) {
            throw Object.assign(new Error('Kun PDF er tilladt.'), { statusCode: 400 });
        }
        return { kind: 'url', value: parsed.href };
    }

    const localPath = /^file:\/\//i.test(value)
        ? normalizeFileUrl(value)
        : value.replace(/\//g, '\\');
    const isDrivePath = /^[A-Za-z]:\\/.test(localPath);
    const isUncPath = /^\\\\[^\\]+\\[^\\]+/.test(localPath);
    if ((!isDrivePath && !isUncPath) || path.win32.extname(localPath).toLowerCase() !== '.pdf') {
        throw Object.assign(new Error('Kun en absolut PDF-sti er tilladt.'), { statusCode: 400 });
    }
    return { kind: 'path', value: localPath };
}

async function openPdfTarget(rawTarget, options = {}) {
    const target = normalizePdfTarget(rawTarget);
    if (target.kind === 'url' && typeof options.openExternal === 'function') {
        await options.openExternal(target.value);
        return target;
    }
    if (target.kind === 'path' && typeof options.openPath === 'function') {
        const errorMessage = await options.openPath(target.value);
        if (errorMessage) throw new Error(String(errorMessage));
        return target;
    }

    if (typeof options.spawn !== 'function') {
        throw new Error('Ingen PDF-fremviser er tilgængelig.');
    }
    const child = options.spawn('explorer.exe', [target.value], {
        windowsHide: true,
        detached: true,
        stdio: 'ignore',
        shell: false
    });
    if (child && typeof child.once === 'function') child.once('error', () => {});
    if (child && typeof child.unref === 'function') child.unref();
    return target;
}

module.exports = { normalizePdfTarget, openPdfTarget };
