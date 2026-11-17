const test = require('node:test');
const assert = require('node:assert/strict');
const { createCustomerSummaryPdfService } = require('../services/customerSummaryPdfService');

test('customer PDF is saved directly and its hidden window is destroyed', async () => {
    const calls = {};
    class Window {
        constructor(options) {
            calls.options = options;
            this.webContents = { printToPDF:async options => { calls.printOptions = options; return Buffer.from('%PDF'); } };
        }
        async loadURL(url) { calls.url = url; }
        destroy() { calls.destroyed = true; }
    }
    const save = createCustomerSummaryPdfService({
        BrowserWindow:Window, sandbox:true,
        dialog:{ showSaveDialog:async (_parent, options) => { calls.dialog = options; return { filePath:'report.pdf' }; } },
        fs:{ promises:{ writeFile:async (path, data) => { calls.path = path; calls.pdf = data; } } }
    });
    assert.deepEqual(await save({}, { month:'2026-10', html:'<!doctype html><html><head></head><body>Lille Industri</body></html>' }), { canceled:false });
    assert.equal(calls.dialog.defaultPath, 'kunder-maanedsanalyse-2026-10.pdf');
    assert.equal(calls.options.webPreferences.javascript, false);
    assert.match(Buffer.from(calls.url.split(',')[1], 'base64').toString(), /Content-Security-Policy/);
    assert.deepEqual(calls.printOptions, { printBackground:true, preferCSSPageSize:true });
    assert.equal(calls.path, 'report.pdf');
    assert.equal(calls.pdf.toString(), '%PDF');
    assert.equal(calls.destroyed, true);
});

test('cancelled or failed customer PDF saves do not report success', async () => {
    let destroyed = false;
    class Window {
        constructor() { this.webContents = { printToPDF:async () => { throw new Error('PDF generation failed'); } }; }
        async loadURL() {}
        destroy() { destroyed = true; }
    }
    const data = { month:'2026-10', html:'<html><head></head><body></body></html>' };
    const options = { BrowserWindow:Window, sandbox:true, fs:{ promises:{ writeFile:async () => {} } } };
    const cancel = createCustomerSummaryPdfService({ ...options, dialog:{ showSaveDialog:async () => ({ canceled:true }) } });
    assert.deepEqual(await cancel({}, data), { canceled:true });
    assert.equal(destroyed, false);
    const fail = createCustomerSummaryPdfService({ ...options, dialog:{ showSaveDialog:async () => ({ filePath:'report.pdf' }) } });
    await assert.rejects(fail({}, data), /PDF generation failed/);
    assert.equal(destroyed, true);
    await assert.rejects(fail({}, { ...data, month:'2026-99' }), /Ugyldige data/);
});
