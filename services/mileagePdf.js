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
            for (const trip of trips) drawSingleTrip(trip);
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
                    ? trip.stops.map((stop, index) => {
                        const name = String(stop.name || '').trim();
                        const address = String(stop.address || '').trim();
                        const location = name && address && name.toLocaleLowerCase('da-DK') !== address.toLocaleLowerCase('da-DK')
                            ? name + '\n' + address : address || name || '—';
                        return (index + 1) + '. ' + location + (stop.note ? '\nNote: ' + stop.note : '');
                    }).join('\n')
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

        function drawSingleTrip(trip) {
            let contentWidth;
            let bottom;
            let y = margin;

            function startPage(continuation = false) {
                doc.addPage();
                contentWidth = doc.page.width - margin * 2;
                bottom = doc.page.height - margin - 18;
                doc.rect(0, 0, doc.page.width, 82).fill(ink);
                doc.font('Helvetica-Bold').fontSize(8).fillColor('#b8ded7')
                    .text('GANTECH', margin, 15, { characterSpacing: 1.2 });
                doc.font('Helvetica-Bold').fontSize(19).fillColor('#ffffff')
                    .text(continuation ? 'Kørselsbilag · fortsat' : 'Kørselsbilag', margin, 33);
                doc.font('Helvetica').fontSize(8).fillColor('#d8e2eb')
                    .text('Bilags-ID  ' + (trip.id || '—'), margin, 63, { width: contentWidth, align: 'right' });
                y = 88;
            }

            function ensureSpace(height) {
                if (y + height > bottom) startPage(true);
            }

            function measureCard(field, width) {
                const value = String(field.value ?? '—');
                doc.font('Helvetica-Bold').fontSize(field.valueSize || 10);
                const valueHeight = doc.heightOfString(value, { width: width - 22, lineGap: 1 });
                let subtitleHeight = 0;
                if (field.subtitle) {
                    doc.font('Helvetica').fontSize(8.5);
                    subtitleHeight = doc.heightOfString(String(field.subtitle), { width: width - 22, lineGap: 1 }) + 2;
                }
                return Math.max(40, 25 + valueHeight + subtitleHeight);
            }

            function paintCard(field, x, top, width, height) {
                doc.roundedRect(x, top, width, height, 5).fill(field.fill || '#f4f7f9');
                doc.roundedRect(x, top, width, height, 5).lineWidth(0.7).stroke(lineColor);
                doc.font('Helvetica-Bold').fontSize(7.5).fillColor(accent)
                    .text(field.label.toUpperCase(), x + 11, top + 6, { width: width - 22 });
                const value = String(field.value ?? '—');
                doc.font('Helvetica-Bold').fontSize(field.valueSize || 10).fillColor(ink)
                    .text(value, x + 11, top + 17, { width: width - 22, lineGap: 1 });
                if (field.subtitle) {
                    doc.font('Helvetica-Bold').fontSize(field.valueSize || 10);
                    const valueHeight = doc.heightOfString(value, { width: width - 22, lineGap: 1 });
                    doc.font('Helvetica').fontSize(8.5).fillColor('#60758a')
                        .text(String(field.subtitle), x + 11, top + 17 + valueHeight + 2, { width: width - 22, lineGap: 1 });
                }
            }

            function drawCardRow(fields, widths) {
                const gap = 6;
                const height = Math.max(...fields.map((field, index) => measureCard(field, widths[index])));
                ensureSpace(height);
                let x = margin;
                fields.forEach((field, index) => {
                    paintCard(field, x, y, widths[index], height);
                    x += widths[index] + gap;
                });
                y += height + 5;
            }

            function drawField(label, value, options = {}) {
                const field = { label, value: String(value ?? '—'), ...options };
                const height = measureCard(field, contentWidth);
                ensureSpace(height);
                paintCard(field, margin, y, contentWidth, height);
                y += height + 5;
            }

            function sectionHeading(label) {
                ensureSpace(20);
                doc.font('Helvetica-Bold').fontSize(10).fillColor(ink).text(label.toUpperCase(), margin, y);
                doc.moveTo(margin, y + 13).lineTo(margin + contentWidth, y + 13).lineWidth(1).stroke(accent);
                y += 20;
            }

            function drawStop(stop, index) {
                const name = String(stop.name || '').trim();
                const address = String(stop.address || '').trim();
                const separateName = name && address && name.toLocaleLowerCase('da-DK') !== address.toLocaleLowerCase('da-DK');
                const location = address || name || '—';
                const note = String(stop.note || '').trim();
                const textX = margin + 47;
                const textWidth = contentWidth - 59;
                let textHeight = 0;
                if (separateName) {
                    doc.font('Helvetica-Bold').fontSize(10);
                    textHeight += doc.heightOfString(name, { width: textWidth, lineGap: 1 }) + 2;
                }
                doc.font('Helvetica').fontSize(9.5);
                textHeight += doc.heightOfString(location, { width: textWidth, lineGap: 1 });
                if (note) {
                    doc.font('Helvetica-Bold').fontSize(7.5);
                    textHeight += 11 + doc.heightOfString(note, { width: textWidth, lineGap: 1 }) + 2;
                }
                const height = Math.max(40, textHeight + 14);
                ensureSpace(height);
                doc.roundedRect(margin, y, contentWidth, height, 5).fill('#f8fafb');
                doc.roundedRect(margin, y, contentWidth, height, 5).lineWidth(0.7).stroke(lineColor);
                doc.circle(margin + 20, y + 20, 10).fill(accent);
                doc.font('Helvetica-Bold').fontSize(8).fillColor('#ffffff')
                    .text(String(index + 1), margin + 12, y + 16, { width: 16, align: 'center' });
                let cursor = y + 7;
                if (separateName) {
                    doc.font('Helvetica-Bold').fontSize(10).fillColor(ink)
                        .text(name, textX, cursor, { width: textWidth, lineGap: 1 });
                    cursor += doc.heightOfString(name, { width: textWidth, lineGap: 1 }) + 2;
                }
                doc.font('Helvetica').fontSize(9.5).fillColor(ink)
                    .text(location, textX, cursor, { width: textWidth, lineGap: 1 });
                cursor += doc.heightOfString(location, { width: textWidth, lineGap: 1 });
                if (note) {
                    cursor += 2;
                    doc.font('Helvetica-Bold').fontSize(7.5).fillColor(accent).text('NOTE', textX, cursor);
                    cursor += 9;
                    doc.font('Helvetica').fontSize(8.5).fillColor('#42596d')
                        .text(note, textX, cursor, { width: textWidth, lineGap: 1 });
                }
                y += height + 4;
            }

            startPage();
            drawCardRow([
                { label: 'Medarbejder', value: trip.employeeName || '—', subtitle: trip.plate ? 'Nummerplade ' + trip.plate : '' },
                { label: 'Kørselsdato', value: trip.date || 'Dato mangler' },
                { label: 'Status', value: statusNames[trip.status] || trip.status || '—' }
            ], [238, 117, contentWidth - 367]);
            if (trip.createdByName || trip.createdBy) drawField('Oprettet af', trip.createdByName || trip.createdBy);
            drawField('Formål', trip.purpose || '—');
            if (trip.legacy) drawField('Power Apps-reference', trip.sourceId || '—');
            if (trip.reconciliationRequired) {
                drawField('Kontrol kræves', 'Tidligere km er ændret efter godkendelsen. Kontrollér beregningen; beløbet er ikke genberegnet.', { fill: '#fff7e8' });
            }

            sectionHeading('Rute og stop');
            if ((trip.stops || []).length) trip.stops.forEach(drawStop);
            else drawField('Historisk rutebeskrivelse', 'Se formål ovenfor.');

            drawCardRow([
                { label: 'Retur', value: trip.roundTrip ? 'Ja' : 'Nej' },
                { label: 'Samlet antal km', value: money(trip.km || 0) + ' km' }
            ], [(contentWidth - 10) / 2, (contentWidth - 10) / 2]);
            drawField('Afstandsgrundlag', trip.manualReason || '—');

            sectionHeading('Godtgørelse');
            if (includeAmounts && trip.allowance) {
                const allowance = trip.allowance;
                drawField('Godkendt godtgørelse', money(allowance.amount) + ' DKK · satsår ' + allowance.year, {
                    fill: '#eaf4f1',
                    valueSize: 13
                });
                drawField('Beregning', [
                    'Tidligere godkendte km: ' + money(allowance.previousKm),
                    'Km til og med 20.000: ' + money(allowance.highKm) + ' km × ' + rate(allowance.highRate) + ' DKK',
                    'Km over 20.000: ' + money(allowance.lowKm) + ' km × ' + rate(allowance.lowRate) + ' DKK'
                ].join('\n'));
                drawField('Godkendt af / tidspunkt', (trip.approvedByName || trip.approvedBy || '—') + ' / ' + (trip.approvedAt || '—'));
            } else {
                drawField('Godtgørelse', includeAmounts
                    ? trip.legacy ? 'Historisk beløb ikke tilgængeligt i eksporten; ikke genberegnet.' : 'Ikke godkendt - intet endeligt beløb.'
                    : 'Beløb er kun synlige for administratorer.');
            }

            if (trip.comment) drawField('Kommentar', trip.comment);
            if (trip.legacy) drawField('Historisk betalingsmarkering', trip.legacyPaid ? 'Udbetalt i Power Apps' : 'Ikke markeret udbetalt i Power Apps');
            sectionHeading('Hændelser');
            drawField('Historik', (trip.history || []).map(item =>
                `${item.at} | ${item.byName || item.by} | ${item.action}${item.note ? ' | ' + item.note : ''}`).join('\n') || 'Importeret historik');
        }
    });
}
module.exports = { generateMileagePdf };
