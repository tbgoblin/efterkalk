const test = require('node:test');
const assert = require('node:assert/strict');

const { parseWorksheetXml, readLaserTechnicalParameters, normalizeLaserTechnicalRows } = require('../services/bomExcelImportService');

test('reads shared strings, numbers and cached formula values from worksheet XML', () => {
    const xml = '<worksheet><sheetData><row r="2">'
        + '<c r="A2" t="s"><v>0</v></c>'
        + '<c r="B2"><v>8</v></c>'
        + '<c r="C2"><f>B2/2</f><v>4</v></c>'
        + '</row></sheetData></worksheet>';
    const rows = parseWorksheetXml(xml, ['R1100']);
    assert.deepEqual(rows.get(2), ['R1100', 8, 4]);
});

test('reads Laserberegner2 technical parameters including gas pressure', () => {
    const rows = readLaserTechnicalParameters(require('path').join(__dirname, '..', 'BOM.xlsm'));
    const aluminum = rows.find(row => row.technology === 'AL000-00.50M-MIX-S0');
    assert.ok(rows.length >= 30);
    assert.deepEqual(aluminum, {
        technology: 'AL000-00.50M-MIX-S0', material: 'AL', thickness: 0.5, lens: '30 S',
        piercingMilliseconds: 20, vaporPowerW: 500, reducedPowerW: 4000,
        feedrateLargeMmMin: 120000, feedrateMediumMmMin: 10000, feedrateSmallMmMin: 5000,
        feedrateEngravingMmMin: 10000, gasPressureBar: 9, nozzleSizeMm: 3
    });
});

test('normalizes technical rows from the Excel table range', () => {
    const rows = normalizeLaserTechnicalRows([
        ['BR000-01.00M-N2-S0', 'ME', 1, '40 S', 10, 300, 4000, 40000, 10000, 2000, 5000, 35000, 11, 4, null],
        [' ', null, null, null]
    ]);
    assert.deepEqual(rows, [{
        technology: 'BR000-01.00M-N2-S0', material: 'ME', thickness: 1, lens: '40 S',
        piercingMilliseconds: 10, vaporPowerW: 300, reducedPowerW: 4000,
        feedrateLargeMmMin: 40000, feedrateMediumMmMin: 10000, feedrateSmallMmMin: 2000,
        feedrateEngravingMmMin: 5000, gasPressureBar: 11, nozzleSizeMm: 4
    }]);
});

test('normalizes blank and non-numeric optional Excel values to zero', () => {
    const rows = normalizeLaserTechnicalRows([
        ['FE000-01.00M-O2-G', 'SORT', 1, null, null, 'n/a', 4000, 10000, null, 2000, 5000, null, '', undefined]
    ]);
    assert.equal(rows[0].lens, '');
    assert.equal(rows[0].piercingMilliseconds, 0);
    assert.equal(rows[0].vaporPowerW, 0);
    assert.equal(rows[0].gasPressureBar, 0);
    assert.equal(rows[0].nozzleSizeMm, 0);
});