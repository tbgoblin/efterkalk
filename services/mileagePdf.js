const PDFDocument = require('pdfkit');

const money = value => value.toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const rate = value => value.toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 3 });
const statusNames = { draft: 'Kladde', submitted: 'Afventer godkendelse', approved: 'Godkendt', paid: 'Udbetalt', rejected: 'Afvist', legacy: 'Importeret historik' };
const ink = '#17324d';
const accent = '#176d64';
const lineColor = '#d8e2eb';

function generateMileagePdf(trips, { includeSummary = true, includeAmounts = false } = {}) {
    return new Promise((resolve, reject) => {
        const margin = includeSummary ? 36 : 40;
        const doc = new PDFDocument({ size: 'A4', layout: includeSummary ? 'landscape' : 'portrait', margin,
            autoFirstPage: false, info: { Title: 'Gantech - Kørselsgodtgørelse', Author: 'Gantech Operations Hub' } });
        const buffers = [];
        doc.on('data', chunk => buffers.push(chunk));
        doc.on('end', () => resolve(Buffer.concat(buffers)));
        doc.on('error', reject);
        const line = (label, value) => {
            doc.font('Helvetica-Bold').fontSize(10).fillColor(ink).text(label);
            doc.font('Helvetica').fontSize(10).fillColor(ink).text(String(value ?? '-'), { paragraphGap: 7 });
        };
        const operationalTrips = trips.filter(trip => !trip.legacy);
        const archiveTrips = trips.filter(trip => trip.legacy);
        const totalKm = operationalTrips.reduce((sum, trip) => sum + (trip.km || 0), 0);
        const totalAmount = operationalTrips.reduce((sum, trip) => sum + (trip.allowance?.amount || 0), 0);

        if (includeSummary) {
            doc.addPage();
            doc.font('Helvetica-Bold').fontSize(25).fillColor(ink).text('Gantech');
            doc.fontSize(17).fillColor(accent).text('Kørselsgodtgørelse · Årsrapport');
            doc.moveDown(0.35);
            doc.font('Helvetica').fontSize(11).fillColor(ink)
                .text((trips[0]?.employeeName || 'Medarbejder') + ' · ' + (trips[0]?.date?.slice(0, 4) || 'Valgt år'));
            doc.moveDown(1.1);

            const cards = [
                ['Kørsler', String(operationalTrips.length)],
                ['Km i alt', money(totalKm) + ' km'],
                ['Historiske poster', String(archiveTrips.length)]
            ];
            if (includeAmounts) cards.push(['Godkendt beløb', money(totalAmount) + ' DKK']);
            const gap = 12;
            const cardWidth = (doc.page.width - margin * 2 - gap * (cards.length - 1)) / cards.length;
            const cardY = doc.y;
            cards.forEach(([label, value], index) => {
                const x = margin + index * (cardWidth + gap);
                doc.roundedRect(x, cardY, cardWidth, 58, 6).fill('#edf3f7');
                doc.font('Helvetica-Bold').fontSize(9).fillColor(accent).text(label.toUpperCase(), x + 12, cardY + 10, { width: cardWidth - 24 });
                doc.font('Helvetica-Bold').fontSize(15).fillColor(ink).text(value, x + 12, cardY + 27, { width: cardWidth - 24 });
            });
            doc.y = cardY + 76;
            line('Dannet', new Date().toLocaleString('da-DK'));
            doc.font('Helvetica').fontSize(9).fillColor('#60758a')
                .text('Arkivposter er kun til dokumentation og indgår ikke i km- eller godtgørelsesberegningen. Beløb vises kun for kørsler med fastlåst beregning. Dette bilag udfører ingen bankbetaling.');
            drawAnnualTable();
        } else {
            for (const trip of trips) {
                doc.addPage();
                doc.font('Helvetica-Bold').fontSize(16).fillColor(ink).text('Gantech - Kørselsbilag');
                doc.moveDown();
                line('Bilags-ID', trip.id);
                if (trip.legacy) line('Power Apps-reference', trip.sourceId);
                line('Medarbejder / nummerplade', trip.employeeName + ' / ' + trip.plate);
                if (trip.createdBy) line('Oprettet af', trip.createdBy);
                line('Kørselsdato / status', (trip.date || 'Dato mangler') + ' / ' + statusNames[trip.status]);
                if (trip.reconciliationRequired) line('Kontrol kræves', 'Tidligere km er ændret efter godkendelsen. Kontrollér beregningen; beløbet er ikke genberegnet.');
                line('Formål', trip.purpose);
                line('Rute og stop', (trip.stops || []).length ? trip.stops.map((stop, index) =>
                    (index + 1) + '. ' + stop.address + (stop.note ? ' (' + stop.note + ')' : '')).join('\n') : 'Historisk rutebeskrivelse: se Formål ovenfor.');
                line('Retur / samlet antal km', (trip.roundTrip ? 'Ja' : 'Nej') + ' / ' + money(trip.km) + ' km');
                line('Afstandsgrundlag', trip.manualReason);
                if (includeAmounts && trip.allowance) {
                    const a = trip.allowance;
                    line('Årets tidligere godkendte km ved godkendelse', money(a.previousKm));
                    line('Km til og med 20.000 / sats', money(a.highKm) + ' km × ' + rate(a.highRate) + ' DKK');
                    line('Km over 20.000 / sats', money(a.lowKm) + ' km × ' + rate(a.lowRate) + ' DKK');
                    line('Godkendt godtgørelse', money(a.amount) + ' DKK (satsår ' + a.year + ')');
                } else {
                    line('Godtgørelse', includeAmounts
                        ? trip.legacy ? 'Historisk beløb ikke tilgængeligt i eksporten; ikke genberegnet.' : 'Ikke godkendt - intet endeligt beløb.'
                        : 'Beløb er kun synlige for administratorer.');
                }
                if (includeAmounts && trip.allowance) line('Godkendt af / tidspunkt', trip.approvedBy + ' / ' + trip.approvedAt);
                if (trip.comment) line('Kommentar', trip.comment);
                if (trip.legacy) line('Historisk betalingsmarkering', trip.legacyPaid ? 'Udbetalt i Power Apps' : 'Ikke markeret udbetalt i Power Apps');
                line('Hændelser', (trip.history || []).map(item => `${item.at} | ${item.by} | ${item.action}${item.note ? ' | ' + item.note : ''}`).join('\n') || 'Importeret historik');
            }
        }

        const pages = doc.bufferedPageRange();
        for (let index = pages.start; index < pages.start + pages.count; index += 1) {
            doc.switchToPage(index);
            doc.font('Helvetica').fontSize(8).fillColor('#60758a')
                .text('Gantech · Kørselgodtgørelse', margin, doc.page.height - margin - 10, { width: doc.page.width - margin * 2, align: 'right' });
        }
        doc.end();

        function drawAnnualTable() {
            const availableWidth = doc.page.width - margin * 2;
            const columns = [
                { label: 'Dato', width: 58, value: trip => trip.date || '—' },
                { label: 'Formål', width: 105, value: trip => trip.purpose || '—' },
                { label: 'Steder og noter', width: 0, value: trip => (trip.stops || []).length
                    ? trip.stops.map((stop, index) => (index + 1) + '. ' + stop.address + (stop.note ? '\nNote: ' + stop.note : '')).join('\n')
                    : trip.purpose || 'Historisk rutebeskrivelse ikke tilgængelig.' },
                { label: 'Km', width: 48, value: trip => money(trip.km || 0), align: 'right' },
                { label: 'Status', width: 98, value: trip => statusNames[trip.status] || trip.status || '—' }
            ];
            if (includeAmounts) columns.push({ label: 'Godkendt DKK', width: 82, value: trip => trip.allowance ? money(trip.allowance.amount) : '—', align: 'right' });
            const fixedWidth = columns.reduce((sum, column) => sum + column.width, 0);
            const placesColumn = columns.find(column => column.label === 'Steder og noter');
            placesColumn.width = availableWidth - fixedWidth;

            function tablePage() {
                doc.addPage();
                doc.font('Helvetica-Bold').fontSize(15).fillColor(ink).text('Kørsler og besøgte steder');
                doc.moveDown(0.65);
                const headerY = doc.y;
                doc.rect(margin, headerY, availableWidth, 25).fill(ink);
                let x = margin;
                for (const column of columns) {
                    doc.font('Helvetica-Bold').fontSize(8).fillColor('white')
                        .text(column.label, x + 5, headerY + 8, { width: column.width - 10, align: column.align || 'left' });
                    x += column.width;
                }
                return headerY + 25;
            }

            let y = tablePage();
            for (const [index, trip] of trips.entries()) {
                const values = columns.map(column => String(column.value(trip) ?? '—'));
                const heights = columns.map((column, columnIndex) => {
                    doc.font('Helvetica').fontSize(8);
                    return doc.heightOfString(values[columnIndex], { width: column.width - 10, lineGap: 1 });
                });
                const rowHeight = Math.max(26, Math.max(...heights) + 10);
                if (y + rowHeight > doc.page.height - margin - 10) y = tablePage();
                if (index % 2 === 1) doc.rect(margin, y, availableWidth, rowHeight).fill('#f2f7f8');
                doc.rect(margin, y, availableWidth, rowHeight).lineWidth(0.5).stroke(lineColor);
                let x = margin;
                columns.forEach((column, columnIndex) => {
                    doc.font('Helvetica').fontSize(8).fillColor(ink)
                        .text(values[columnIndex], x + 5, y + 5, {
                            width: column.width - 10,
                            align: column.align || 'left',
                            lineGap: 1
                        });
                    x += column.width;
                });
                y += rowHeight;
            }
        }
    });
}
module.exports = { generateMileagePdf };
