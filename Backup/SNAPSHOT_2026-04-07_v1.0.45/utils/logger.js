const fs = require('fs');
const path = require('path');

function resolveWritableLogFile() {
    const candidates = [
        process.env.GANTECH_LOG_DIR,
        process.env.LOCALAPPDATA ? path.join(process.env.LOCALAPPDATA, 'Gantech Efterkalk') : null,
        process.env.APPDATA ? path.join(process.env.APPDATA, 'Gantech Efterkalk') : null,
        process.cwd(),
        __dirname
    ].filter(Boolean);

    for (const dir of candidates) {
        try {
            fs.mkdirSync(dir, { recursive: true });
            const testPath = path.join(dir, 'gantech.log');
            fs.appendFileSync(testPath, '');
            return testPath;
        } catch (_) {
            // Try next candidate path.
        }
    }

    return path.join(process.cwd(), 'gantech.log');
}

function createLogger(appVersion) {
    const logFile = resolveWritableLogFile();

    function logEvent(message) {
        const timestamp = new Date().toISOString();
        const logMessage = `[${timestamp}] ${message}\n`;
        try {
            fs.appendFileSync(logFile, logMessage);
        } catch (_) {
            // Keep app alive even if file logging is temporarily unavailable.
        }
        console.log(logMessage.trim());
    }

    if (appVersion) {
        logEvent('=== SERVER STARTED - ' + appVersion + ' ===');
    }

    return {
        logFile,
        logEvent
    };
}

module.exports = {
    resolveWritableLogFile,
    createLogger
};
