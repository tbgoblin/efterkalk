const PDFDocument = require('pdfkit');

const money = value => value.toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const rate = value => value.toLocaleString('da-DK', { minimumFractionDigits: 2, maximumFractionDigits: 3 });
const statusNames = { draft: 'Kladde', submitted: 'Afventer godkendelse', approved: 'Godkendt', paid: 'Udbetalt', rejected: 'Afvist', legacy: 'Importeret historik' };

function generateMileagePdf(trips) {
    return new Promise((resolve, reject) => {
        const doc = new PDFDocument({ size: 'A4', margin: 40, autoFirstPage: false,
            info: { Title: 'Gantech - Kørselsgodtgørelse', Author: 'Gantech Operations Hub' } });
        const buffers = [];
        doc.on('data', chunk => buffers.push(chunk));
        doc.on('end', () => resolve(Buffer.concat(buffers)));
        doc.on('error', reject);
        const line = (label, value) => {
            doc.font('Helvetica-Bold').fontSize(10).fillColor('#17324d').text(label);
            doc.font('Helvetica').fontSize(10).text(String(value ?? '-'), { paragraphGap: 7 });
        };
        doc.addPage();
        doc.font('Helvetica-Bold').fontSize(22).text('Gantech');
        doc.fontSize(16).text('Kørselsgodtgørelse - internt bilag');
        doc.moveDown();
        line('Dannet', new Date().toLocaleString('da-DK'));
        line('Antal kørsler', trips.length);
        line('Km i alt', money(trips.reduce((sum, trip) => sum + trip.km, 0)));
        line('Dokumenteret beløb i alt (DKK)', money(trips.reduce((sum, trip) => sum + (trip.allowance?.amount || 0), 0)));
        doc.fontSize(10).text('Beløbet omfatter kun godkendte/udbetalte kørsler med fastlåst beregning. Historiske beløb, kladder og afventende kørsler er ikke medregnet. Dette bilag udfører ingen bankbetaling.');
        for (const trip of trips) {
            doc.addPage();
            doc.font('Helvetica-Bold').fontSize(16).text('Gantech - Kørselsbilag');
            doc.moveDown();
            line('Bilags-ID', trip.id);
            if (trip.legacy) line('Power Apps-reference', trip.sourceId);
            line('Medarbejder / nummerplade', trip.employeeName + ' / ' + trip.plate);
            line('Kørselsdato / status', (trip.date || 'Dato mangler') + ' / ' + statusNames[trip.status]);
            line('Formål', trip.purpose);
            line('Rute og stop', trip.stops.length ? trip.stops.map((stop, index) =>
                (index + 1) + '. ' + stop.address + (stop.note ? ' (' + stop.note + ')' : '')).join('\n') : 'Historisk rutebeskrivelse: se Formål ovenfor.');
            line('Retur / samlet antal km', (trip.roundTrip ? 'Ja' : 'Nej') + ' / ' + money(trip.km) + ' km');
            line('Afstandsgrundlag', trip.manualReason);
            if (trip.allowance) {
                const a = trip.allowance;
                line('Årets tidligere godkendte km ved godkendelse', money(a.previousKm));
                line('Km til og med 20.000 / sats', money(a.highKm) + ' km × ' + rate(a.highRate) + ' DKK');
                line('Km over 20.000 / sats', money(a.lowKm) + ' km × ' + rate(a.lowRate) + ' DKK');
                line('Godkendt godtgørelse', money(a.amount) + ' DKK (satsår ' + a.year + ')');
                line('Godkendt af / tidspunkt', trip.approvedBy + ' / ' + trip.approvedAt);
            } else line('Godtgørelse', trip.legacy ? 'Historisk beløb ikke tilgængeligt i eksporten; ikke genberegnet.' : 'Ikke godkendt - intet endeligt beløb.');
            if (trip.comment) line('Kommentar', trip.comment);
            if (trip.legacy) line('Historisk betalingsmarkering', trip.legacyPaid ? 'Udbetalt i Power Apps' : 'Ikke markeret udbetalt i Power Apps');
            line('Hændelser', trip.history.map(item => `${item.at} | ${item.by} | ${item.action}${item.note ? ' | ' + item.note : ''}`).join('\n') || 'Importeret historik');
        }
        doc.end();
    });
}
module.exports = { generateMileagePdf };
