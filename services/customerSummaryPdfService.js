function createCustomerSummaryPdfService({ BrowserWindow, dialog, fs, sandbox }) {
    return async function saveCustomerSummaryPdf(parent, data) {
        if (!data || !/^\d{4}-(0[1-9]|1[0-2])$/.test(data.month || '')
            || typeof data.html !== 'string' || Buffer.byteLength(data.html, 'utf8') > 10 * 1024 * 1024) {
            throw new Error('Ugyldige data til kundeanalysens PDF.');
        }

        const result = await dialog.showSaveDialog(parent, {
            title: 'Gem kundeanalysen som PDF',
            defaultPath: 'kunder-maanedsanalyse-' + data.month + '.pdf',
            filters: [{ name: 'PDF', extensions: ['pdf'] }]
        });
        if (result.canceled) return { canceled: true };
        if (!result.filePath) throw new Error('PDF-filstien mangler.');

        const reportWindow = new BrowserWindow({
            show: false,
            webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox, javascript: false }
        });
        try {
            const csp = `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'">`;
            const html = data.html.replace('<head>', '<head>' + csp);
            if (html === data.html) throw new Error('Ugyldig PDF-visning.');
            await reportWindow.loadURL('data:text/html;charset=utf-8;base64,' + Buffer.from(html, 'utf8').toString('base64'));
            const pdf = await reportWindow.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true });
            await fs.promises.writeFile(result.filePath, pdf);
            return { canceled: false };
        } finally {
            reportWindow.destroy();
        }
    };
}

module.exports = { createCustomerSummaryPdfService };
