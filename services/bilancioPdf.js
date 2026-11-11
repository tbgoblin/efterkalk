const PDFDocument = require('pdfkit');
const months = ['Juli', 'August', 'September', 'Oktober', 'November', 'December', 'Januar', 'Februar', 'Marts', 'April', 'Maj', 'Juni'];

function generateBilancioPdf(saved) {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocument({ autoFirstPage: false, margin: 28, info: {
            Title: saved.report.reportName + ' - Rev. ' + saved.revision,
            Author: saved.savedBy, CreationDate: new Date(saved.savedAt)
        } });
        const buffers = [];
        doc.on('data', chunk => buffers.push(chunk));
        doc.on('end', () => resolve(Buffer.concat(buffers)));
        doc.on('error', reject);
        const fmt = value => value == null ? '-' : new Intl.NumberFormat('da-DK', {
            minimumFractionDigits: saved.unit === 1 ? 2 : 0, maximumFractionDigits: saved.unit === 1 ? 2 : 0
        }).format(value / saved.unit);
        const pct = value => value == null ? '-' : new Intl.NumberFormat('da-DK', { maximumFractionDigits: 0 }).format(value) + '%';
        const { report, monthly, yearly } = saved;
        const y = report.year, p = report.period;
        const calendarYear = y + (p > 6 ? 1 : 0);
        const monthLabel = column => months[column.period - 1].slice(0, 3).toLowerCase() + '-' + String(column.year + (column.period > 6 ? 1 : 0)).slice(-2);
        const unit = saved.unit === 1 ? 'DKK' : 't.kr';
        function table(title, layout, headers, rows, columns, percentages = false) {
            const fontSize = columns.length === 12 ? 7 : 9;
            let cursor;
            const newPage = () => {
                doc.addPage({ size: 'A4', layout });
                const width = doc.page.width - 56;
                doc.font('Helvetica-Bold').fontSize(13).fillColor('#173653').text('Gantech A/S - ' + report.reportName, 28, 28, { width });
                let top = doc.y + 6;
                doc.font('Helvetica').fontSize(8).text(
                    `${months[p - 1]} ${calendarYear} | Rapport Rev. ${saved.revision} | Konfiguration Rev. ${saved.configurationRevision.number} (${saved.configurationRevision.id.slice(0, 12)}) | App ${saved.appVersion}`,
                    28, top, { width }
                );
                top = doc.y + 4;
                doc.text(`Gemt ${saved.savedAt} | ${saved.savedBy} | ${unit}`, 28, top, { width });
                if (saved.note) doc.text(saved.note, 28, doc.y + 4, { width });
                doc.font('Helvetica-Bold').fontSize(10).text(title, 28, doc.y + 8, { width });
                cursor = doc.y + 8;
                drawRow(['Kontogruppe / konto', ...headers], true);
            };
            function drawRow(texts, header = false, bold = false) {
                const width = doc.page.width - 56;
                const labelWidth = width * (columns.length === 12 ? 0.26 : 0.34);
                const valueWidth = (width - labelWidth) / columns.length;
                const widths = [labelWidth, ...columns.map(() => valueWidth)];
                doc.font(header || bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(fontSize);
                const height = Math.max(19, ...texts.map((text, i) => doc.heightOfString(String(text), { width: widths[i] - 8 }) + 8));
                if (!header && cursor + height > doc.page.height - 30) newPage();
                if (header || bold) doc.rect(28, cursor, width, height).fill(header ? '#173e64' : '#f3f7fb');
                let x = 28;
                doc.font(header || bold ? 'Helvetica-Bold' : 'Helvetica').fontSize(fontSize);
                texts.forEach((text, i) => {
                    doc.fillColor(header ? '#ffffff' : i && String(text).startsWith('-') && text !== '-' ? '#a33835' : '#173653')
                        .text(String(text), x + 4, cursor + 4, { width: widths[i] - 8, align: i ? 'right' : 'left' });
                    x += widths[i];
                });
                doc.moveTo(28, cursor + height).lineTo(28 + width, cursor + height).strokeColor('#d5e0eb').lineWidth(0.5).stroke();
                cursor += height;
            }
            newPage();
            for (const row of rows) {
                if (row.visibility === 'hidden') continue;
                const texts = columns.map(index => {
                    if (row.type === 'heading' || (row.visibility === 'period' && index >= 2 && percentages)
                        || (row.visibility === 'ytd' && index < 2 && percentages)) return '';
                    const value = row.amounts[index];
                    return fmt(value) + (percentages ? ' / ' + pct(row.percentages?.[index]) : '')
                        + (row.checkZero ? value == null ? ' (Mangler data)' : Math.abs(value) < 0.005 ? ' (OK)' : ' (Afvigelse)' : '');
                });
                drawRow([row.name, ...texts], false, row.bold || ['subtotal', 'heading'].includes(row.type));
            }
        }
        try {
            if (report.showPnl !== false) table('Resultatopgørelse', 'landscape', [
                monthLabel({ year: y, period: p }) + ' / %', monthLabel({ year: y - 1, period: p }) + ' / %',
                'ÅTD ' + y + '-' + String(y + 1).slice(-2) + ' / %', 'ÅTD ' + (y - 1) + '-' + String(y).slice(-2) + ' / %'
            ], [...report.rows, ...(report.periodOnly || [])], [0, 1, 2, 3], true);
            if (report.assets) table(report.assets.title, 'portrait', [
                months[p - 1] + ' ' + calendarYear, months[p - 1] + ' ' + (calendarYear - 1)
            ], report.assets.rows.filter(row => row.visibility !== 'period'), [0, 1]);
            table('Seneste 12 måneder', 'landscape', monthly.columns.map(monthLabel),
                monthly.rows.filter(row => row.visibility !== 'ytd'), monthly.columns.map((_, i) => i));
            table('Seneste 12 regnskabsår', 'landscape', yearly.columns.map(c => c.year + '-' + String(c.year + 1).slice(-2) + (c.period < 12 ? ' (ÅTD)' : '')),
                yearly.rows.filter(row => row.visibility !== 'period'), yearly.columns.map((_, i) => i));
            doc.end();
        } catch (error) {
            doc.destroy();
            reject(error);
        }
    });
}
module.exports = { generateBilancioPdf };
