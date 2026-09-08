const test = require('node:test');
const assert = require('node:assert/strict');

const { selectBendingHandlingBand, calculateBendingOperation, calculateFlatOperationMinutes, calculateNitrogenFlowNm3Hour, calculateAssistGasFlowNm3Hour, detectLaserGasType, detectLaserMachine, detectLaserTechnicalMaterial, selectCheapestLaserTechnology } = require('../services/bomService');

const machine = {
    MachineCode: 'R2100', MaxBendLengthMm: 3000, MaxForceKn: 1000,
    BaseCycleSeconds: 8, SecondsPerDegree: 0.02, BackGaugeSeconds: 2,
    SetupMinutes: 10, SafetyFactor: 0.8
};
const bands = [{
    BandName: 'Mellem', MinWeightKg: 0, MaxWeightKg: 20,
    MinLongestSideMm: 0, MaxLongestSideMm: 1500,
    LoadSeconds: 5, UnloadSeconds: 4, Rotate90Seconds: 3, FlipSeconds: 6
}];

test('selects a handling band by both weight and size', () => {
    assert.equal(selectBendingHandlingBand(bands, 5, 800).BandName, 'Mellem');
    assert.equal(selectBendingHandlingBand(bands, 25, 800), null);
});

test('calculates bending cycle, handling and setup per piece', () => {
    const result = calculateBendingOperation({
        bendCount: 2, totalBendLengthMm: 1600, averageAngleDeg: 90,
        thicknessMm: 2, dieOpeningMm: 16, tensileStrengthMpa: 450,
        pieceWeightKg: 5, longestSideMm: 1000, rotate90Count: 1, flipCount: 0
    }, machine, bands);
    assert.equal(result.minutes, 0.59);
    assert.equal(result.setupMinutes, 10);
    assert.equal(result.requiredForceKn, 127.8);
    assert.equal(result.handlingBand, 'Mellem');
});

test('rejects a bend that exceeds machine force', () => {
    assert.throws(() => calculateBendingOperation({
        bendCount: 1, totalBendLengthMm: 2500, averageAngleDeg: 90,
        thicknessMm: 10, dieOpeningMm: 40, tensileStrengthMpa: 450,
        pieceWeightKg: 5, longestSideMm: 1000
    }, machine, bands), /bukkekraft/);
});

test('calculates nitrogen flow from pressure and nozzle size', () => {
    assert.equal(calculateNitrogenFlowNm3Hour(6, 4), 64.99);
    assert.equal(calculateNitrogenFlowNm3Hour(6, 5), 101.55);
    assert.equal(Math.round((calculateNitrogenFlowNm3Hour(6, 5) / 0.862) * 100) / 100, 117.81);
    assert.equal(Math.round((calculateNitrogenFlowNm3Hour(6, 5) / 0.862 * 2.25) * 100) / 100, 265.07);
    assert.equal(calculateNitrogenFlowNm3Hour(0, 3), 0);
    assert.equal(calculateNitrogenFlowNm3Hour(9, 0), 0);
});

test('detects nitrogen and oxygen technologies and calculates both flows', () => {
    assert.equal(detectLaserGasType('AL000-01.00M-N2-S0'), 'nitrogen');
    assert.equal(detectLaserGasType('FE000-03.00M-O2-G'), 'oxygen');
    assert.equal(detectLaserGasType('AL000-03.00M-MIX-S0'), 'mixline');
    assert.ok(calculateAssistGasFlowNm3Hour(9, 3, 'oxygen') > 0);
});

test('maps Visma material families used by Laserberegner2', () => {
    assert.equal(detectLaserTechnicalMaterial('30100100001001', 'DC01'), 'SORT');
    assert.equal(detectLaserTechnicalMaterial('31100100310015', 'Rustfri'), 'RF');
    assert.equal(detectLaserTechnicalMaterial('32100100110151', 'Aluplade'), 'AL');
    assert.equal(detectLaserTechnicalMaterial('33100100800301', 'Aluzink'), 'GAL');
    assert.equal(detectLaserTechnicalMaterial('38100100100101', 'Kobberplade'), 'CO');
    assert.equal(detectLaserTechnicalMaterial('38100100200101', 'Messingplade'), 'ME');
});

test('selects the cheapest complete technology using speed, piercing and gas', () => {
    const base = { Material: 'SORT', Thickness: 1, GasPressureBar: 10, NozzleSizeMm: 2, PiercingMilliseconds: 50 };
    const result = selectCheapestLaserTechnology([
        { ...base, Technology: 'ST000-01.00M-N2-G', FeedrateLargeMmMin: 50000 },
        { ...base, Technology: 'ST000-01.00M-O2-G', FeedrateLargeMmMin: 80000 }
    ], { cutLengthM: 10, piercings: 5, machineRate: 15, nitrogenPricePerKg: 4, oxygenPricePerKg: 0.5 });
    assert.equal(result.selected.row.Technology, 'ST000-01.00M-O2-G');
    assert.ok(result.alternatives[0].minutes < result.alternatives[1].minutes);
});

test('includes MixLine with the configured N2/O2 blend and excludes incomplete rows', () => {
    const result = selectCheapestLaserTechnology([
        { Technology: 'ST000-01.00M-MIX-S0', FeedrateLargeMmMin: 100000, GasPressureBar: 10, NozzleSizeMm: 2 },
        { Technology: 'ST000-01.00M-O2-G', FeedrateLargeMmMin: 80000, GasPressureBar: 0, NozzleSizeMm: 0 }
    ], { cutLengthM: 10, piercings: 1, machineRate: 15,
        nitrogenPricePerKg: 10, oxygenPricePerKg: 20, mixLineOxygenPercent: 22 });
    assert.equal(result.selected.row.Technology, 'ST000-01.00M-MIX-S0');
    assert.equal(result.selected.gasPricePerKg, 12.2);
    assert.equal(result.alternatives.length, 1);
});

test('manual estimation can still use technical speed when gas data is unavailable', () => {
    const row = { Technology: 'ST000-01.00M-MIX-S0', FeedrateLargeMmMin: 60000,
        PiercingMilliseconds: 50, GasPressureBar: 0, NozzleSizeMm: 0 };
    const result = require('../services/bomService').estimateLaserTechnology(row,
        { cutLengthM: 6, piercings: 2, machineRate: 15, requireCompleteGas: false });
    assert.ok(result.minutes > 0);
    assert.equal(result.feedrateMmMin, 60000);
    assert.equal(result.gasCostPerMinute, 0);
});

test('maps Eagle lines and Genius technologies to their machines', () => {
    assert.deepEqual(detectLaserMachine('SS000-05.00M-N2-S0'), { code: 'R1100', name: 'Eagle', line: 'Standard' });
    assert.deepEqual(detectLaserMachine('SS000-05.00M-N2-C0'), { code: 'R1100', name: 'Eagle', line: 'CutLine' });
    assert.deepEqual(detectLaserMachine('SS000-05.00M-N2-F0'), { code: 'R1100', name: 'Eagle', line: 'FastLine' });
    assert.deepEqual(detectLaserMachine('SS000-05.00M-N2-X0'), { code: 'R1100', name: 'Eagle', line: '15 kW' });
    assert.deepEqual(detectLaserMachine('SS000-05.00M-N2-G'), { code: 'R1102', name: 'Laser Genius', line: 'Genius' });
});

test('uses the Genius rate only for -G technologies', () => {
    const base = { Material: 'RF', Thickness: 5, GasPressureBar: 10, NozzleSizeMm: 2,
        PiercingMilliseconds: 30, FeedrateLargeMmMin: 10000 };
    const result = selectCheapestLaserTechnology([
        { ...base, Technology: 'SS000-05.00M-N2-F0' },
        { ...base, Technology: 'SS000-05.00M-N2-G' }
    ], { cutLengthM: 1, machineRate: 16.5, geniusMachineRate: 15 });
    const eagle = result.alternatives.find(row => row.row.Technology.endsWith('-F0'));
    const genius = result.alternatives.find(row => row.row.Technology.endsWith('-G'));
    assert.equal(eagle.machineRate, 16.5);
    assert.equal(genius.machineRate, 15);
});

test('calculates Flad minutes with the Excel type multipliers', () => {
    const input = { widthMm: 400, lengthMm: 100, speed: 1, factor: 10 };
    assert.equal(calculateFlatOperationMinutes({ ...input, type: 'R05' }), 1);
    assert.equal(calculateFlatOperationMinutes({ ...input, type: 'R10' }), 1.33);
    assert.equal(calculateFlatOperationMinutes({ ...input, type: 'R15' }), 1.66);
    assert.equal(calculateFlatOperationMinutes({ ...input, type: 'R20' }), 2);
    assert.equal(calculateFlatOperationMinutes({ ...input, type: 'B05' }), 0.588235);
});