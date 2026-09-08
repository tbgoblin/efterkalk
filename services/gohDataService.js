// ── GOH data hub ─────────────────────────────────────────────────────────────
// Persistenza storica su SQL Server GantechOperationHub: snapshot efterkalk
// (Orders + AfterkalkReadings) e import grezzi (RawImports). Fail-soft come
// gohCacheService: se il DB non risponde si disabilita 5 min e l'app prosegue.
const sql = require('mssql/msnodesqlv8');

const GOH_SERVER = process.env.GOH_CACHE_SERVER || '192.168.17.2\\GOH';
const GOH_DATABASE = process.env.GOH_DATA_DB || 'GantechOperationHub';
const DISABLE_AFTER_ERROR_MS = 5 * 60 * 1000;

const COMPANY_NAME = 'Gantech';
const PLANT_NAME = 'Hovedfabrik';
const MEASUREMENT_TYPES = [
    { name: 'Omsaetning', unit: 'DKK', description: 'Efterkalk: samlet omsætning for ordren' },
    { name: 'Kostpris', unit: 'DKK', description: 'Efterkalk: samlet kostpris for ordren' }
];

let poolPromise = null;
let disabledUntil = 0;
let seedPromise = null;
let bomLaserSchemaPromise = null;
let bomLaserTechnicalSchemaPromise = null;
let bomBendingSchemaPromise = null;
let lastUnavailableError = null;
let logEvent = () => {};
const lastSnapshotByOrder = new Map(); // OrdNo → 'revenue|cost' per dedup in-process

function configure({ logEvent: logger } = {}) {
    if (typeof logger === 'function') logEvent = logger;
}

function isEnabled() {
    return Date.now() >= disabledUntil;
}

function markUnavailable(err) {
    disabledUntil = Date.now() + DISABLE_AFTER_ERROR_MS;
    poolPromise = null;
    seedPromise = null;
    bomLaserSchemaPromise = null;
    bomLaserTechnicalSchemaPromise = null;
    bomBendingSchemaPromise = null;
    const code = err && (err.code || err.number);
    const message = err && err.message ? err.message : String(err || 'Ukendt GOH-fejl');
    lastUnavailableError = (code ? String(code) + ': ' : '') + message;
    logEvent('GOH-DATA UNAVAILABLE (retry in 5 min): ' + lastUnavailableError);
}

function getStatus() {
    return {
        enabled: isEnabled(),
        disabledUntil: disabledUntil ? new Date(disabledUntil).toISOString() : null,
        lastError: lastUnavailableError
    };
}

function handleSchemaError(label, err) {
    const code = err && err.code ? String(err.code) : '';
    if (['ETIMEOUT', 'ESOCKET', 'ECONNCLOSED', 'ENOTOPEN', 'ELOGIN'].includes(code)) {
        markUnavailable(err);
        return;
    }
    const number = err && err.number;
    const message = err && err.message ? err.message : String(err || 'Ukendt schemafejl');
    const migrationHint = /CREATE TABLE permission denied/i.test(message)
        ? ' Kør docs/sql/2026-09-04-bom-calculator-parameters.sql én gang med en GOH-databaseadministrator.'
        : '';
    lastUnavailableError = label + (code || number ? ' [' + (code || number) + ']' : '') + ': ' + message + migrationHint;
    logEvent('GOH-DATA SCHEMA ERROR: ' + lastUnavailableError);
}

async function ensureBomBendingTables() {
    if (bomBendingSchemaPromise) return bomBendingSchemaPromise;
    bomBendingSchemaPromise = (async () => {
        const pool = await getPool();
        if (!pool) return false;
        await pool.request().query(`
            IF OBJECT_ID('dbo.BomBendingMachines', 'U') IS NULL
            BEGIN
                CREATE TABLE dbo.BomBendingMachines (
                    MachineCode nvarchar(100) NOT NULL PRIMARY KEY,
                    Description nvarchar(500) NULL,
                    MaxBendLengthMm decimal(18,3) NOT NULL,
                    MaxForceKn decimal(18,3) NOT NULL,
                    BaseCycleSeconds decimal(18,3) NOT NULL,
                    SecondsPerDegree decimal(18,6) NOT NULL CONSTRAINT DF_BomBendingMachines_Degree DEFAULT (0),
                    BackGaugeSeconds decimal(18,3) NOT NULL CONSTRAINT DF_BomBendingMachines_Gauge DEFAULT (0),
                    SetupMinutes decimal(18,3) NOT NULL CONSTRAINT DF_BomBendingMachines_Setup DEFAULT (0),
                    SafetyFactor decimal(9,4) NOT NULL CONSTRAINT DF_BomBendingMachines_Safety DEFAULT (0.8),
                    IsActive bit NOT NULL CONSTRAINT DF_BomBendingMachines_Active DEFAULT (1),
                    UpdatedAt datetime2(0) NOT NULL CONSTRAINT DF_BomBendingMachines_UpdatedAt DEFAULT (SYSUTCDATETIME()),
                    UpdatedBy nvarchar(100) NULL,
                    CONSTRAINT CK_BomBendingMachines_Positive CHECK (MaxBendLengthMm > 0 AND MaxForceKn > 0 AND BaseCycleSeconds >= 0 AND SafetyFactor > 0 AND SafetyFactor <= 1)
                );
            END;
            IF OBJECT_ID('dbo.BomBendingHandlingBands', 'U') IS NULL
            BEGIN
                CREATE TABLE dbo.BomBendingHandlingBands (
                    HandlingBandId int IDENTITY(1,1) NOT NULL PRIMARY KEY,
                    BandName nvarchar(100) NOT NULL,
                    MinWeightKg decimal(18,3) NOT NULL CONSTRAINT DF_BomBendingHandlingBands_MinWeight DEFAULT (0),
                    MaxWeightKg decimal(18,3) NULL,
                    MinLongestSideMm decimal(18,3) NOT NULL CONSTRAINT DF_BomBendingHandlingBands_MinSide DEFAULT (0),
                    MaxLongestSideMm decimal(18,3) NULL,
                    LoadSeconds decimal(18,3) NOT NULL,
                    UnloadSeconds decimal(18,3) NOT NULL,
                    Rotate90Seconds decimal(18,3) NOT NULL,
                    FlipSeconds decimal(18,3) NOT NULL,
                    IsActive bit NOT NULL CONSTRAINT DF_BomBendingHandlingBands_Active DEFAULT (1),
                    UpdatedAt datetime2(0) NOT NULL CONSTRAINT DF_BomBendingHandlingBands_UpdatedAt DEFAULT (SYSUTCDATETIME()),
                    UpdatedBy nvarchar(100) NULL,
                    CONSTRAINT CK_BomBendingHandlingBands_Valid CHECK (MinWeightKg >= 0 AND MinLongestSideMm >= 0 AND LoadSeconds >= 0 AND UnloadSeconds >= 0 AND Rotate90Seconds >= 0 AND FlipSeconds >= 0)
                );
            END;
            IF OBJECT_ID('dbo.BomBendingActualSamples', 'U') IS NULL
            BEGIN
                CREATE TABLE dbo.BomBendingActualSamples (
                    SampleId bigint IDENTITY(1,1) NOT NULL PRIMARY KEY,
                    MachineCode nvarchar(100) NOT NULL,
                    ProdNo nvarchar(100) NULL,
                    OrderNo nvarchar(100) NULL,
                    Quantity int NOT NULL,
                    PieceWidthMm decimal(18,3) NOT NULL,
                    PieceLengthMm decimal(18,3) NOT NULL,
                    ThicknessMm decimal(18,3) NOT NULL,
                    PieceWeightKg decimal(18,3) NOT NULL,
                    BendCount int NOT NULL,
                    TotalBendLengthMm decimal(18,3) NOT NULL,
                    AverageAngleDeg decimal(18,3) NOT NULL,
                    Rotate90Count int NOT NULL CONSTRAINT DF_BomBendingActualSamples_Rotate DEFAULT (0),
                    FlipCount int NOT NULL CONSTRAINT DF_BomBendingActualSamples_Flip DEFAULT (0),
                    SetupMinutes decimal(18,3) NOT NULL CONSTRAINT DF_BomBendingActualSamples_Setup DEFAULT (0),
                    ActualRunMinutes decimal(18,3) NOT NULL,
                    RecordedAt datetime2(0) NOT NULL CONSTRAINT DF_BomBendingActualSamples_RecordedAt DEFAULT (SYSUTCDATETIME()),
                    RecordedBy nvarchar(100) NULL,
                    CONSTRAINT FK_BomBendingActualSamples_Machine FOREIGN KEY (MachineCode) REFERENCES dbo.BomBendingMachines(MachineCode),
                    CONSTRAINT CK_BomBendingActualSamples_Valid CHECK (Quantity > 0 AND BendCount >= 0 AND ActualRunMinutes >= 0)
                );
            END;
        `);
        return true;
    })().catch(err => {
        bomBendingSchemaPromise = null;
        handleSchemaError('BOM buk-tabeller', err);
        return false;
    });
    return bomBendingSchemaPromise;
}

async function getBomBendingParameters() {
    if (!await ensureBomBendingTables()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const [machines, bands] = await Promise.all([
            pool.request().query(`SELECT MachineCode, Description, MaxBendLengthMm, MaxForceKn, BaseCycleSeconds,
                                         SecondsPerDegree, BackGaugeSeconds, SetupMinutes, SafetyFactor, UpdatedAt, UpdatedBy
                                  FROM dbo.BomBendingMachines WHERE IsActive = 1 ORDER BY MachineCode`),
            pool.request().query(`SELECT HandlingBandId, BandName, MinWeightKg, MaxWeightKg, MinLongestSideMm,
                                         MaxLongestSideMm, LoadSeconds, UnloadSeconds, Rotate90Seconds, FlipSeconds,
                                         UpdatedAt, UpdatedBy
                                  FROM dbo.BomBendingHandlingBands WHERE IsActive = 1
                                  ORDER BY MinWeightKg, MinLongestSideMm`)
        ]);
        return { machines: machines.recordset || [], handlingBands: bands.recordset || [] };
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function upsertBomBendingMachine(input, updatedByValue) {
    if (!await ensureBomBendingTables()) return false;
    const machineCode = String(input && input.machineCode || '').trim();
    const maxLength = Number(input && input.maxBendLengthMm);
    const maxForce = Number(input && input.maxForceKn);
    const baseCycle = Number(input && input.baseCycleSeconds);
    const secondsPerDegree = Number(input && input.secondsPerDegree || 0);
    const backGauge = Number(input && input.backGaugeSeconds || 0);
    const setup = Number(input && input.setupMinutes || 0);
    const safety = Number(input && input.safetyFactor || 0.8);
    if (!machineCode || ![maxLength, maxForce, baseCycle, secondsPerDegree, backGauge, setup, safety].every(Number.isFinite)
        || maxLength <= 0 || maxForce <= 0 || baseCycle < 0 || secondsPerDegree < 0 || backGauge < 0 || setup < 0 || safety <= 0 || safety > 1) return false;
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('machineCode', sql.NVarChar(100), machineCode)
            .input('description', sql.NVarChar(500), String(input.description || '').trim())
            .input('maxLength', sql.Decimal(18, 3), maxLength)
            .input('maxForce', sql.Decimal(18, 3), maxForce)
            .input('baseCycle', sql.Decimal(18, 3), baseCycle)
            .input('secondsPerDegree', sql.Decimal(18, 6), secondsPerDegree)
            .input('backGauge', sql.Decimal(18, 3), backGauge)
            .input('setup', sql.Decimal(18, 3), setup)
            .input('safety', sql.Decimal(9, 4), safety)
            .input('updatedBy', sql.NVarChar(100), String(updatedByValue || '').trim())
            .query(`MERGE dbo.BomBendingMachines AS target USING (SELECT @machineCode AS MachineCode) AS source
                    ON target.MachineCode = source.MachineCode
                    WHEN MATCHED THEN UPDATE SET Description=@description, MaxBendLengthMm=@maxLength, MaxForceKn=@maxForce,
                        BaseCycleSeconds=@baseCycle, SecondsPerDegree=@secondsPerDegree, BackGaugeSeconds=@backGauge,
                        SetupMinutes=@setup, SafetyFactor=@safety, IsActive=1, UpdatedAt=SYSUTCDATETIME(), UpdatedBy=@updatedBy
                    WHEN NOT MATCHED THEN INSERT (MachineCode, Description, MaxBendLengthMm, MaxForceKn, BaseCycleSeconds,
                        SecondsPerDegree, BackGaugeSeconds, SetupMinutes, SafetyFactor, UpdatedBy)
                        VALUES (@machineCode, @description, @maxLength, @maxForce, @baseCycle, @secondsPerDegree,
                        @backGauge, @setup, @safety, @updatedBy);`);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

async function upsertBomBendingHandlingBand(input, updatedByValue) {
    if (!await ensureBomBendingTables()) return false;
    const id = Number(input && input.handlingBandId || 0);
    const values = {
        name: String(input && input.bandName || '').trim(), minWeight: Number(input && input.minWeightKg || 0),
        maxWeight: input && input.maxWeightKg !== '' && input.maxWeightKg != null ? Number(input.maxWeightKg) : null,
        minSide: Number(input && input.minLongestSideMm || 0),
        maxSide: input && input.maxLongestSideMm !== '' && input.maxLongestSideMm != null ? Number(input.maxLongestSideMm) : null,
        load: Number(input && input.loadSeconds), unload: Number(input && input.unloadSeconds),
        rotate: Number(input && input.rotate90Seconds), flip: Number(input && input.flipSeconds)
    };
    if (!values.name || ![values.minWeight, values.minSide, values.load, values.unload, values.rotate, values.flip].every(Number.isFinite)
        || (values.maxWeight != null && !Number.isFinite(values.maxWeight)) || (values.maxSide != null && !Number.isFinite(values.maxSide))) return false;
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('id', sql.Int, id)
            .input('name', sql.NVarChar(100), values.name)
            .input('minWeight', sql.Decimal(18, 3), values.minWeight)
            .input('maxWeight', sql.Decimal(18, 3), values.maxWeight)
            .input('minSide', sql.Decimal(18, 3), values.minSide)
            .input('maxSide', sql.Decimal(18, 3), values.maxSide)
            .input('load', sql.Decimal(18, 3), values.load)
            .input('unload', sql.Decimal(18, 3), values.unload)
            .input('rotate', sql.Decimal(18, 3), values.rotate)
            .input('flip', sql.Decimal(18, 3), values.flip)
            .input('updatedBy', sql.NVarChar(100), String(updatedByValue || '').trim())
            .query(`IF @id > 0 AND EXISTS (SELECT 1 FROM dbo.BomBendingHandlingBands WHERE HandlingBandId=@id)
                        UPDATE dbo.BomBendingHandlingBands SET BandName=@name, MinWeightKg=@minWeight, MaxWeightKg=@maxWeight,
                            MinLongestSideMm=@minSide, MaxLongestSideMm=@maxSide, LoadSeconds=@load, UnloadSeconds=@unload,
                            Rotate90Seconds=@rotate, FlipSeconds=@flip, IsActive=1, UpdatedAt=SYSUTCDATETIME(), UpdatedBy=@updatedBy
                        WHERE HandlingBandId=@id
                    ELSE
                        INSERT dbo.BomBendingHandlingBands (BandName, MinWeightKg, MaxWeightKg, MinLongestSideMm,
                            MaxLongestSideMm, LoadSeconds, UnloadSeconds, Rotate90Seconds, FlipSeconds, UpdatedBy)
                        VALUES (@name, @minWeight, @maxWeight, @minSide, @maxSide, @load, @unload, @rotate, @flip, @updatedBy);`);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

async function addBomBendingActualSample(input, recordedByValue) {
    if (!await ensureBomBendingTables()) return false;
    const machineCode = String(input && input.machineCode || '').trim();
    const numeric = {
        quantity: Number(input && input.quantity), width: Number(input && input.pieceWidthMm), length: Number(input && input.pieceLengthMm),
        thickness: Number(input && input.thicknessMm), weight: Number(input && input.pieceWeightKg), bendCount: Number(input && input.bendCount),
        bendLength: Number(input && input.totalBendLengthMm), angle: Number(input && input.averageAngleDeg),
        rotate: Number(input && input.rotate90Count || 0), flip: Number(input && input.flipCount || 0),
        setup: Number(input && input.setupMinutes || 0), actual: Number(input && input.actualRunMinutes)
    };
    if (!machineCode || !Object.values(numeric).every(Number.isFinite) || numeric.quantity <= 0 || numeric.bendCount < 0 || numeric.actual < 0) return false;
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('machineCode', sql.NVarChar(100), machineCode)
            .input('prodNo', sql.NVarChar(100), String(input.prodNo || '').trim())
            .input('orderNo', sql.NVarChar(100), String(input.orderNo || '').trim())
            .input('quantity', sql.Int, numeric.quantity)
            .input('width', sql.Decimal(18, 3), numeric.width)
            .input('length', sql.Decimal(18, 3), numeric.length)
            .input('thickness', sql.Decimal(18, 3), numeric.thickness)
            .input('weight', sql.Decimal(18, 3), numeric.weight)
            .input('bendCount', sql.Int, numeric.bendCount)
            .input('bendLength', sql.Decimal(18, 3), numeric.bendLength)
            .input('angle', sql.Decimal(18, 3), numeric.angle)
            .input('rotate', sql.Int, numeric.rotate)
            .input('flip', sql.Int, numeric.flip)
            .input('setup', sql.Decimal(18, 3), numeric.setup)
            .input('actual', sql.Decimal(18, 3), numeric.actual)
            .input('recordedBy', sql.NVarChar(100), String(recordedByValue || '').trim())
            .query(`INSERT dbo.BomBendingActualSamples (MachineCode, ProdNo, OrderNo, Quantity, PieceWidthMm,
                        PieceLengthMm, ThicknessMm, PieceWeightKg, BendCount, TotalBendLengthMm, AverageAngleDeg,
                        Rotate90Count, FlipCount, SetupMinutes, ActualRunMinutes, RecordedBy)
                    VALUES (@machineCode, NULLIF(@prodNo,''), NULLIF(@orderNo,''), @quantity, @width, @length,
                        @thickness, @weight, @bendCount, @bendLength, @angle, @rotate, @flip, @setup, @actual, @recordedBy)`);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

async function ensureBomLaserParameterTable() {
    if (bomLaserSchemaPromise) return bomLaserSchemaPromise;
    bomLaserSchemaPromise = (async () => {
        const pool = await getPool();
        if (!pool) return false;
        await pool.request().query(`
            IF OBJECT_ID('dbo.BomLaserParameters', 'U') IS NULL
            BEGIN
                CREATE TABLE dbo.BomLaserParameters (
                    LaserParameterId int IDENTITY(1,1) NOT NULL PRIMARY KEY,
                    ProdNo nvarchar(100) NOT NULL,
                    Description nvarchar(500) NULL,
                    Thickness decimal(18,4) NULL,
                    Machine nvarchar(100) NOT NULL,
                    CutSpeedMPerMin decimal(18,6) NOT NULL,
                    PiercingMinutes decimal(18,6) NOT NULL CONSTRAINT DF_BomLaserParameters_Piercing DEFAULT (0),
                    SurchargePercent decimal(18,6) NOT NULL CONSTRAINT DF_BomLaserParameters_Surcharge DEFAULT (0),
                    Lens nvarchar(100) NULL,
                    IsActive bit NOT NULL CONSTRAINT DF_BomLaserParameters_IsActive DEFAULT (1),
                    Source nvarchar(50) NOT NULL CONSTRAINT DF_BomLaserParameters_Source DEFAULT ('program'),
                    UpdatedAt datetime2(0) NOT NULL CONSTRAINT DF_BomLaserParameters_UpdatedAt DEFAULT (SYSUTCDATETIME()),
                    UpdatedBy nvarchar(100) NULL,
                    CONSTRAINT UQ_BomLaserParameters_ProdNo_Machine UNIQUE (ProdNo, Machine),
                    CONSTRAINT CK_BomLaserParameters_NonNegative CHECK (CutSpeedMPerMin >= 0 AND PiercingMinutes >= 0)
                );
            END
        `);
        return true;
    })().catch(err => {
        bomLaserSchemaPromise = null;
        handleSchemaError('BOM laser-tabel', err);
        return false;
    });
    return bomLaserSchemaPromise;
}

async function getBomLaserParameters(machineValue) {
    if (!await ensureBomLaserParameterTable()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const machine = String(machineValue || '').trim();
        const result = await pool.request()
            .input('machine', sql.NVarChar(100), machine)
            .query(`SELECT ProdNo, Description AS Descr, Thickness AS Tykkelse, Machine AS Maskine,
                           CutSpeedMPerMin AS [Skærehast.], PiercingMinutes AS Pircing,
                           SurchargePercent AS [Tillæg], Lens AS Linse, IsActive, Source, UpdatedAt, UpdatedBy
                    FROM dbo.BomLaserParameters
                    WHERE IsActive = 1 AND Source <> 'visma-seed' AND (@machine = '' OR Machine = @machine)
                    ORDER BY ProdNo, Machine`);
        return result.recordset || [];
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function getBomLaserParameter(prodNoValue, machineValue) {
    if (!await ensureBomLaserParameterTable()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request()
            .input('prodNo', sql.NVarChar(100), String(prodNoValue || '').trim())
            .input('machine', sql.NVarChar(100), String(machineValue || '').trim())
            .query(`SELECT TOP 1 ProdNo, Description AS Descr, Thickness AS Tykkelse, Machine AS Maskine,
                           CutSpeedMPerMin AS Skaerehast, PiercingMinutes AS Piercing,
                           SurchargePercent AS Tillaeg, Lens AS Linse, IsActive, Source, UpdatedAt, UpdatedBy
                    FROM dbo.BomLaserParameters
                    WHERE ProdNo = @prodNo AND Machine = @machine AND IsActive = 1 AND Source <> 'visma-seed'`);
        return (result.recordset || [])[0] || null;
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function upsertBomLaserParameter(input, updatedByValue) {
    if (!await ensureBomLaserParameterTable()) return null;
    const prodNo = String(input && input.prodNo || '').trim();
    const machine = String(input && input.machine || '').trim();
    const cutSpeed = Number(input && input.cutSpeedMPerMin);
    const piercing = Number(input && input.piercingMinutes);
    const surcharge = Number(input && input.surchargePercent);
    if (!prodNo || !machine || !Number.isFinite(cutSpeed) || cutSpeed < 0 || !Number.isFinite(piercing) || piercing < 0 || !Number.isFinite(surcharge)) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        await pool.request()
            .input('prodNo', sql.NVarChar(100), prodNo)
            .input('description', sql.NVarChar(500), String(input.description || '').trim())
            .input('thickness', sql.Decimal(18, 4), Number(input.thickness || 0))
            .input('machine', sql.NVarChar(100), machine)
            .input('cutSpeed', sql.Decimal(18, 6), cutSpeed)
            .input('piercing', sql.Decimal(18, 6), piercing)
            .input('surcharge', sql.Decimal(18, 6), surcharge)
            .input('lens', sql.NVarChar(100), String(input.lens || '').trim())
            .input('updatedBy', sql.NVarChar(100), String(updatedByValue || '').trim())
            .query(`MERGE dbo.BomLaserParameters AS target
                    USING (SELECT @prodNo AS ProdNo, @machine AS Machine) AS source
                    ON target.ProdNo = source.ProdNo AND target.Machine = source.Machine
                    WHEN MATCHED THEN UPDATE SET Description=@description, Thickness=@thickness,
                        CutSpeedMPerMin=@cutSpeed, PiercingMinutes=@piercing, SurchargePercent=@surcharge,
                        Lens=@lens, IsActive=1, Source='program', UpdatedAt=SYSUTCDATETIME(), UpdatedBy=@updatedBy
                    WHEN NOT MATCHED THEN INSERT
                        (ProdNo, Description, Thickness, Machine, CutSpeedMPerMin, PiercingMinutes, SurchargePercent, Lens, Source, UpdatedBy)
                        VALUES (@prodNo, @description, @thickness, @machine, @cutSpeed, @piercing, @surcharge, @lens, 'program', @updatedBy);`);
        return getBomLaserParameter(prodNo, machine);
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function importBomLaserParameters(rows, updatedByValue, overwriteExisting = false) {
    if (!await ensureBomLaserParameterTable()) return null;
    const sourceRows = Array.isArray(rows) ? rows : [];
    const pool = await getPool();
    if (!pool) return null;
    let inserted = 0;
    let updated = 0;
    try {
        for (let offset = 0; offset < sourceRows.length; offset += 400) {
            const chunk = sourceRows.slice(offset, offset + 400);
            const result = await pool.request()
                .input('payload', sql.NVarChar(sql.MAX), JSON.stringify(chunk))
                .input('overwrite', sql.Bit, overwriteExisting === true)
                .input('updatedBy', sql.NVarChar(100), String(updatedByValue || '').trim())
                .query(`MERGE dbo.BomLaserParameters AS target
                        USING OPENJSON(@payload) WITH (
                            ProdNo nvarchar(100) '$.prodNo', Description nvarchar(500) '$.description',
                            Thickness decimal(18,4) '$.thickness', Machine nvarchar(100) '$.machine',
                            CutSpeed decimal(18,6) '$.cutSpeedMPerMin', Piercing decimal(18,6) '$.piercingMinutes',
                            Surcharge decimal(18,6) '$.surchargePercent', Lens nvarchar(100) '$.lens'
                        ) AS source
                        ON target.ProdNo = source.ProdNo AND target.Machine = source.Machine
                        WHEN MATCHED AND @overwrite = 1 THEN UPDATE SET Description=source.Description,
                            Thickness=source.Thickness, CutSpeedMPerMin=source.CutSpeed, PiercingMinutes=source.Piercing,
                            SurchargePercent=source.Surcharge, Lens=source.Lens, IsActive=1, Source='excel-migration',
                            UpdatedAt=SYSUTCDATETIME(), UpdatedBy=@updatedBy
                        WHEN NOT MATCHED THEN INSERT (ProdNo, Description, Thickness, Machine, CutSpeedMPerMin,
                            PiercingMinutes, SurchargePercent, Lens, Source, UpdatedBy)
                            VALUES (source.ProdNo, source.Description, source.Thickness, source.Machine, source.CutSpeed,
                                source.Piercing, source.Surcharge, source.Lens, 'excel-migration', @updatedBy)
                        OUTPUT $action AS MergeAction;`);
            for (const action of result.recordset || []) {
                if (action.MergeAction === 'INSERT') inserted += 1;
                if (action.MergeAction === 'UPDATE') updated += 1;
            }
        }
        return { sourceRows: sourceRows.length, inserted, updated, preserved: sourceRows.length - inserted - updated };
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function ensureBomLaserTechnicalParameterTable() {
    if (bomLaserTechnicalSchemaPromise) return bomLaserTechnicalSchemaPromise;
    bomLaserTechnicalSchemaPromise = (async () => {
        const pool = await getPool();
        if (!pool) return false;
        await pool.request().query(`
            IF OBJECT_ID('dbo.BomLaserTechnicalParameters', 'U') IS NULL
            BEGIN
                CREATE TABLE dbo.BomLaserTechnicalParameters (
                    Technology nvarchar(100) NOT NULL PRIMARY KEY,
                    Material nvarchar(50) NOT NULL,
                    Thickness decimal(18,4) NOT NULL,
                    Lens nvarchar(100) NULL,
                    PiercingMilliseconds decimal(18,3) NOT NULL,
                    VaporPowerW decimal(18,3) NOT NULL,
                    ReducedPowerW decimal(18,3) NOT NULL,
                    FeedrateLargeMmMin decimal(18,3) NOT NULL,
                    FeedrateMediumMmMin decimal(18,3) NOT NULL,
                    FeedrateSmallMmMin decimal(18,3) NOT NULL,
                    FeedrateEngravingMmMin decimal(18,3) NOT NULL,
                    GasPressureBar decimal(18,4) NOT NULL,
                    NozzleSizeMm decimal(18,4) NOT NULL,
                    IsActive bit NOT NULL CONSTRAINT DF_BomLaserTechnicalParameters_Active DEFAULT (1),
                    Source nvarchar(50) NOT NULL CONSTRAINT DF_BomLaserTechnicalParameters_Source DEFAULT ('program'),
                    UpdatedAt datetime2(0) NOT NULL CONSTRAINT DF_BomLaserTechnicalParameters_UpdatedAt DEFAULT (SYSUTCDATETIME()),
                    UpdatedBy nvarchar(100) NULL,
                    CONSTRAINT CK_BomLaserTechnicalParameters_NonNegative CHECK
                        (Thickness > 0 AND PiercingMilliseconds >= 0 AND VaporPowerW >= 0 AND ReducedPowerW >= 0
                         AND FeedrateLargeMmMin >= 0 AND FeedrateMediumMmMin >= 0 AND FeedrateSmallMmMin >= 0
                         AND FeedrateEngravingMmMin >= 0 AND GasPressureBar >= 0 AND NozzleSizeMm >= 0)
                );
            END
        `);
        return true;
    })().catch(err => {
        bomLaserTechnicalSchemaPromise = null;
        handleSchemaError('BOM laser-teknologitabel', err);
        return false;
    });
    return bomLaserTechnicalSchemaPromise;
}

async function getBomLaserTechnicalParameters() {
    if (!await ensureBomLaserTechnicalParameterTable()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request().query(`SELECT Technology, Material, Thickness, Lens,
                    PiercingMilliseconds, VaporPowerW, ReducedPowerW, FeedrateLargeMmMin, FeedrateMediumMmMin,
                    FeedrateSmallMmMin, FeedrateEngravingMmMin, GasPressureBar, NozzleSizeMm,
                    Source, UpdatedAt, UpdatedBy
                FROM dbo.BomLaserTechnicalParameters WHERE IsActive = 1 ORDER BY Material, Thickness, Technology`);
        return result.recordset || [];
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function getBomLaserTechnicalParameter(technologyValue) {
    if (!await ensureBomLaserTechnicalParameterTable()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request()
            .input('technology', sql.NVarChar(100), String(technologyValue || '').trim())
            .query(`SELECT TOP 1 Technology, Material, Thickness, Lens, PiercingMilliseconds,
                    VaporPowerW, ReducedPowerW, FeedrateLargeMmMin, FeedrateMediumMmMin,
                    FeedrateSmallMmMin, FeedrateEngravingMmMin, GasPressureBar, NozzleSizeMm
                FROM dbo.BomLaserTechnicalParameters WHERE Technology = @technology AND IsActive = 1`);
        return (result.recordset || [])[0] || null;
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function getBomLaserTechnicalParametersForMaterial(materialValue, thicknessValue) {
    if (!await ensureBomLaserTechnicalParameterTable()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request()
            .input('material', sql.NVarChar(50), String(materialValue || '').trim())
            .input('thickness', sql.Decimal(18, 3), Number(thicknessValue))
            .query(`SELECT Technology, Material, Thickness, Lens, PiercingMilliseconds,
                           VaporPowerW, ReducedPowerW, FeedrateLargeMmMin, FeedrateMediumMmMin,
                           FeedrateSmallMmMin, FeedrateEngravingMmMin, GasPressureBar, NozzleSizeMm
                    FROM dbo.BomLaserTechnicalParameters
                    WHERE Material = @material AND ABS(Thickness - @thickness) < 0.001 AND IsActive = 1
                    ORDER BY Technology`);
        return result.recordset || [];
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function importBomLaserTechnicalParameters(rows, updatedByValue, overwriteExisting = false) {
    if (!await ensureBomLaserTechnicalParameterTable()) return null;
    const sourceRows = Array.isArray(rows) ? rows : [];
    const pool = await getPool();
    if (!pool) return null;
    let inserted = 0;
    let updated = 0;
    try {
        for (let offset = 0; offset < sourceRows.length; offset += 400) {
            const chunk = sourceRows.slice(offset, offset + 400);
            const result = await pool.request()
                .input('payload', sql.NVarChar(sql.MAX), JSON.stringify(chunk))
                .input('overwrite', sql.Bit, overwriteExisting === true)
                .input('updatedBy', sql.NVarChar(100), String(updatedByValue || '').trim())
                .query(`MERGE dbo.BomLaserTechnicalParameters AS target
                    USING OPENJSON(@payload) WITH (
                        Technology nvarchar(100) '$.technology', Material nvarchar(50) '$.material',
                        Thickness decimal(18,4) '$.thickness', Lens nvarchar(100) '$.lens',
                        PiercingMilliseconds decimal(18,3) '$.piercingMilliseconds', VaporPowerW decimal(18,3) '$.vaporPowerW',
                        ReducedPowerW decimal(18,3) '$.reducedPowerW', FeedrateLargeMmMin decimal(18,3) '$.feedrateLargeMmMin',
                        FeedrateMediumMmMin decimal(18,3) '$.feedrateMediumMmMin', FeedrateSmallMmMin decimal(18,3) '$.feedrateSmallMmMin',
                        FeedrateEngravingMmMin decimal(18,3) '$.feedrateEngravingMmMin', GasPressureBar decimal(18,4) '$.gasPressureBar',
                        NozzleSizeMm decimal(18,4) '$.nozzleSizeMm'
                    ) AS source ON target.Technology = source.Technology
                    WHEN MATCHED AND @overwrite = 1 THEN UPDATE SET Material=source.Material, Thickness=source.Thickness,
                        Lens=source.Lens, PiercingMilliseconds=source.PiercingMilliseconds, VaporPowerW=source.VaporPowerW,
                        ReducedPowerW=source.ReducedPowerW, FeedrateLargeMmMin=source.FeedrateLargeMmMin,
                        FeedrateMediumMmMin=source.FeedrateMediumMmMin, FeedrateSmallMmMin=source.FeedrateSmallMmMin,
                        FeedrateEngravingMmMin=source.FeedrateEngravingMmMin, GasPressureBar=source.GasPressureBar,
                        NozzleSizeMm=source.NozzleSizeMm, IsActive=1, Source='excel-migration',
                        UpdatedAt=SYSUTCDATETIME(), UpdatedBy=@updatedBy
                    WHEN NOT MATCHED THEN INSERT (Technology, Material, Thickness, Lens, PiercingMilliseconds,
                        VaporPowerW, ReducedPowerW, FeedrateLargeMmMin, FeedrateMediumMmMin, FeedrateSmallMmMin,
                        FeedrateEngravingMmMin, GasPressureBar, NozzleSizeMm, Source, UpdatedBy)
                    VALUES (source.Technology, source.Material, source.Thickness, source.Lens, source.PiercingMilliseconds,
                        source.VaporPowerW, source.ReducedPowerW, source.FeedrateLargeMmMin, source.FeedrateMediumMmMin,
                        source.FeedrateSmallMmMin, source.FeedrateEngravingMmMin, source.GasPressureBar, source.NozzleSizeMm,
                        'excel-migration', @updatedBy)
                    OUTPUT $action AS MergeAction;`);
            for (const action of result.recordset || []) {
                if (action.MergeAction === 'INSERT') inserted += 1;
                if (action.MergeAction === 'UPDATE') updated += 1;
            }
        }
        return { sourceRows: sourceRows.length, inserted, updated, preserved: sourceRows.length - inserted - updated };
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function getPool() {
    if (!isEnabled()) return null;
    if (!poolPromise) {
        poolPromise = new sql.ConnectionPool({
            server: GOH_SERVER,
            database: GOH_DATABASE,
            driver: 'msnodesqlv8',
            connectionTimeout: 8000,
            requestTimeout: 30000,
            pool: { max: 4, min: 0, idleTimeoutMillis: 30000 },
            options: { trustedConnection: true, trustServerCertificate: true }
        }).connect()
            .then(pool => {
                pool.on('error', err => markUnavailable(err));
                lastUnavailableError = null;
                logEvent('GOH-DATA CONNECTED: ' + GOH_SERVER + '/' + GOH_DATABASE);
                return pool;
            })
            .catch(err => {
                markUnavailable(err);
                throw err;
            });
    }
    try {
        return await poolPromise;
    } catch {
        return null;
    }
}

// Seed idempotente di Company/Plant/MeasurementTypes; ritorna gli id (cache in-process).
async function ensureSeed() {
    if (seedPromise) return seedPromise;
    seedPromise = (async () => {
        const pool = await getPool();
        if (!pool) throw new Error('GOH data pool unavailable');

        await pool.request()
            .input('name', sql.NVarChar(200), COMPANY_NAME)
            .query(`IF NOT EXISTS (SELECT 1 FROM dbo.Companies WHERE CompanyName = @name)
                    INSERT INTO dbo.Companies (CompanyName, CompanyCode) VALUES (@name, 'GT')`);
        const companyId = (await pool.request()
            .input('name', sql.NVarChar(200), COMPANY_NAME)
            .query('SELECT TOP 1 CompanyId FROM dbo.Companies WHERE CompanyName = @name')).recordset[0].CompanyId;

        await pool.request()
            .input('companyId', sql.Int, companyId)
            .input('name', sql.NVarChar(200), PLANT_NAME)
            .query(`IF NOT EXISTS (SELECT 1 FROM dbo.Plants WHERE PlantName = @name AND CompanyId = @companyId)
                    INSERT INTO dbo.Plants (CompanyId, PlantName) VALUES (@companyId, @name)`);
        const plantId = (await pool.request()
            .input('companyId', sql.Int, companyId)
            .input('name', sql.NVarChar(200), PLANT_NAME)
            .query('SELECT TOP 1 PlantId FROM dbo.Plants WHERE PlantName = @name AND CompanyId = @companyId')).recordset[0].PlantId;

        const typeIds = {};
        for (const mt of MEASUREMENT_TYPES) {
            await pool.request()
                .input('name', sql.NVarChar(200), mt.name)
                .input('unit', sql.NVarChar(50), mt.unit)
                .input('descr', sql.NVarChar(500), mt.description)
                .query(`IF NOT EXISTS (SELECT 1 FROM dbo.MeasurementTypes WHERE TypeName = @name)
                        INSERT INTO dbo.MeasurementTypes (TypeName, Unit, Description) VALUES (@name, @unit, @descr)`);
            typeIds[mt.name] = (await pool.request()
                .input('name', sql.NVarChar(200), mt.name)
                .query('SELECT TOP 1 MeasurementTypeId FROM dbo.MeasurementTypes WHERE TypeName = @name')).recordset[0].MeasurementTypeId;
        }

        return { companyId, plantId, typeIds };
    })().catch(err => {
        seedPromise = null;
        markUnavailable(err);
        return null;
    });
    return seedPromise;
}

// Snapshot efterkalk: upsert Orders per OrderNumber + 2 righe AfterkalkReadings
// (Omsaetning, Kostpris). Dedup: salta se identico all'ultimo snapshot salvato.
async function recordAftercalcSnapshot(ordNo, summary) {
    if (!isEnabled() || !summary) return false;
    const orderNumber = String(ordNo);
    const revenue = Math.round(Number(summary.totalRevenue || 0) * 100) / 100;
    const cost = Math.round(Number(summary.totalCost || 0) * 100) / 100;
    const fingerprint = revenue + '|' + cost;
    if (lastSnapshotByOrder.get(orderNumber) === fingerprint) return false;

    const seed = await ensureSeed();
    if (!seed) return false;
    const pool = await getPool();
    if (!pool) return false;

    try {
        await pool.request()
            .input('companyId', sql.Int, seed.companyId)
            .input('orderNumber', sql.NVarChar(100), orderNumber)
            .query(`IF NOT EXISTS (SELECT 1 FROM dbo.Orders WHERE OrderNumber = @orderNumber)
                    INSERT INTO dbo.Orders (CompanyId, OrderNumber, Status) VALUES (@companyId, @orderNumber, 'efterkalk')`);
        const orderId = (await pool.request()
            .input('orderNumber', sql.NVarChar(100), orderNumber)
            .query('SELECT TOP 1 OrderId FROM dbo.Orders WHERE OrderNumber = @orderNumber')).recordset[0].OrderId;

        // Dedup persistente: confronta con l'ultimo snapshot su DB
        const lastResult = await pool.request()
            .input('orderId', sql.Int, orderId)
            .query(`SELECT mt.TypeName, r.Value
                    FROM dbo.AfterkalkReadings r
                    JOIN dbo.MeasurementTypes mt ON mt.MeasurementTypeId = r.MeasurementTypeId
                    WHERE r.OrderId = @orderId
                      AND r.ReadingTime = (SELECT MAX(ReadingTime) FROM dbo.AfterkalkReadings WHERE OrderId = @orderId)`);
        const lastValues = {};
        for (const row of (lastResult.recordset || [])) {
            lastValues[row.TypeName] = Math.round(Number(row.Value || 0) * 100) / 100;
        }
        if (lastValues.Omsaetning === revenue && lastValues.Kostpris === cost) {
            lastSnapshotByOrder.set(orderNumber, fingerprint);
            return false;
        }

        const readingTime = new Date();
        const resultStatus = summary.hasInvoiceWarning ? 'NoInvoWarning' : 'OK';
        for (const [typeName, value] of [['Omsaetning', revenue], ['Kostpris', cost]]) {
            await pool.request()
                .input('plantId', sql.Int, seed.plantId)
                .input('orderId', sql.Int, orderId)
                .input('readingTime', sql.DateTime2, readingTime)
                .input('typeId', sql.Int, seed.typeIds[typeName])
                .input('value', sql.Decimal(18, 6), value)
                .input('status', sql.NVarChar(50), resultStatus)
                .query(`INSERT INTO dbo.AfterkalkReadings (PlantId, OrderId, ReadingTime, MeasurementTypeId, Value, ResultStatus)
                        VALUES (@plantId, @orderId, @readingTime, @typeId, @value, @status)`);
        }
        lastSnapshotByOrder.set(orderNumber, fingerprint);
        logEvent('GOH-DATA SNAPSHOT: ordNo=' + orderNumber + ' revenue=' + revenue + ' cost=' + cost);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

// Storico efterkalk di un ordine: righe { readingTime, typeName, value, resultStatus }.
async function getOrderTrend(ordNo) {
    if (!isEnabled()) return [];
    const pool = await getPool();
    if (!pool) return [];
    try {
        const result = await pool.request()
            .input('orderNumber', sql.NVarChar(100), String(ordNo))
            .query(`SELECT r.ReadingTime, mt.TypeName, r.Value, r.ResultStatus
                    FROM dbo.AfterkalkReadings r
                    JOIN dbo.Orders o ON o.OrderId = r.OrderId
                    JOIN dbo.MeasurementTypes mt ON mt.MeasurementTypeId = r.MeasurementTypeId
                    WHERE o.OrderNumber = @orderNumber
                    ORDER BY r.ReadingTime ASC`);
        return (result.recordset || []).map(row => ({
            readingTime: row.ReadingTime,
            typeName: row.TypeName,
            value: Number(row.Value || 0),
            resultStatus: row.ResultStatus || null
        }));
    } catch (err) {
        markUnavailable(err);
        return [];
    }
}

// Import grezzo (es. snapshot lagerliste) su RawImports.
async function saveRawImport(sourceName, sourceType, payload) {
    if (!isEnabled()) return false;
    const seed = await ensureSeed();
    if (!seed) return false;
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('companyId', sql.Int, seed.companyId)
            .input('plantId', sql.Int, seed.plantId)
            .input('sourceName', sql.NVarChar(200), String(sourceName))
            .input('sourceType', sql.NVarChar(100), String(sourceType))
            .input('payload', sql.NVarChar(sql.MAX), JSON.stringify(payload))
            .query(`INSERT INTO dbo.RawImports (CompanyId, PlantId, SourceName, SourceType, RawPayload)
                    VALUES (@companyId, @plantId, @sourceName, @sourceType, @payload)`);
        logEvent('GOH-DATA RAWIMPORT: ' + sourceType + '/' + sourceName);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

// Documenti di stato condivisi (users, note ordini, soglie, ...) su dbo.AppState.
async function getAppState(key) {
    if (!isEnabled()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const result = await pool.request()
            .input('key', sql.NVarChar(100), String(key))
            .query('SELECT Payload, UpdatedAt FROM dbo.AppState WHERE StateKey = @key');
        const row = result.recordset && result.recordset[0];
        if (!row) return null;
        return { payload: JSON.parse(row.Payload), updatedAt: row.UpdatedAt };
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function getAppStatesByPrefix(prefix) {
    if (!isEnabled()) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        const normalizedPrefix = String(prefix || '').slice(0, 100);
        const result = await pool.request()
            .input('prefix', sql.NVarChar(100), normalizedPrefix)
            .query(`SELECT StateKey, Payload, UpdatedAt
                    FROM dbo.AppState
                    WHERE LEFT(StateKey, LEN(@prefix)) = @prefix
                    ORDER BY StateKey`);
        const rows = [];
        for (const row of (result.recordset || [])) {
            try {
                rows.push({
                    key: String(row.StateKey || ''),
                    payload: JSON.parse(row.Payload),
                    updatedAt: row.UpdatedAt
                });
            } catch (parseError) {
                if (logEvent) logEvent('GOH APPSTATE JSON ERROR (' + String(row.StateKey || '?') + '): ' + parseError.message);
            }
        }
        return rows;
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}

async function setAppState(key, payload) {
    if (!isEnabled()) return false;
    const pool = await getPool();
    if (!pool) return false;
    try {
        await pool.request()
            .input('key', sql.NVarChar(100), String(key))
            .input('payload', sql.NVarChar(sql.MAX), JSON.stringify(payload))
            .query(`MERGE dbo.AppState AS t
                USING (SELECT @key AS StateKey) AS s ON t.StateKey = s.StateKey
                WHEN MATCHED THEN UPDATE SET Payload = @payload, UpdatedAt = SYSUTCDATETIME()
                WHEN NOT MATCHED THEN INSERT (StateKey, Payload) VALUES (@key, @payload);`);
        return true;
    } catch (err) {
        markUnavailable(err);
        return false;
    }
}

async function upsertBomLaserTechnicalParameter(input, updatedByValue) {
    if (!await ensureBomLaserTechnicalParameterTable()) return null;
    const technology = String(input && input.technology || '').trim();
    const material = String(input && input.material || '').trim();
    const numeric = {
        thickness: Number(input && input.thickness), piercing: Number(input && input.piercingMilliseconds || 0),
        vapor: Number(input && input.vaporPowerW || 0), reduced: Number(input && input.reducedPowerW || 0),
        large: Number(input && input.feedrateLargeMmMin || 0), medium: Number(input && input.feedrateMediumMmMin || 0),
        small: Number(input && input.feedrateSmallMmMin || 0), engraving: Number(input && input.feedrateEngravingMmMin || 0),
        pressure: Number(input && input.gasPressureBar || 0), nozzle: Number(input && input.nozzleSizeMm || 0)
    };
    if (!technology || !material || !Object.values(numeric).every(Number.isFinite) || numeric.thickness <= 0
        || Object.entries(numeric).some(([key, value]) => key !== 'thickness' && value < 0)) return null;
    const pool = await getPool();
    if (!pool) return null;
    try {
        await pool.request()
            .input('technology', sql.NVarChar(100), technology).input('material', sql.NVarChar(50), material)
            .input('thickness', sql.Decimal(18, 4), numeric.thickness).input('lens', sql.NVarChar(100), String(input.lens || '').trim())
            .input('piercing', sql.Decimal(18, 3), numeric.piercing).input('vapor', sql.Decimal(18, 3), numeric.vapor)
            .input('reduced', sql.Decimal(18, 3), numeric.reduced).input('large', sql.Decimal(18, 3), numeric.large)
            .input('medium', sql.Decimal(18, 3), numeric.medium).input('small', sql.Decimal(18, 3), numeric.small)
            .input('engraving', sql.Decimal(18, 3), numeric.engraving).input('pressure', sql.Decimal(18, 4), numeric.pressure)
            .input('nozzle', sql.Decimal(18, 4), numeric.nozzle).input('updatedBy', sql.NVarChar(100), String(updatedByValue || '').trim())
            .query(`MERGE dbo.BomLaserTechnicalParameters AS target
                USING (SELECT @technology AS Technology) AS source ON target.Technology = source.Technology
                WHEN MATCHED THEN UPDATE SET Material=@material, Thickness=@thickness, Lens=@lens,
                    PiercingMilliseconds=@piercing, VaporPowerW=@vapor, ReducedPowerW=@reduced,
                    FeedrateLargeMmMin=@large, FeedrateMediumMmMin=@medium, FeedrateSmallMmMin=@small,
                    FeedrateEngravingMmMin=@engraving, GasPressureBar=@pressure, NozzleSizeMm=@nozzle,
                    IsActive=1, Source='program', UpdatedAt=SYSUTCDATETIME(), UpdatedBy=@updatedBy
                WHEN NOT MATCHED THEN INSERT (Technology, Material, Thickness, Lens, PiercingMilliseconds,
                    VaporPowerW, ReducedPowerW, FeedrateLargeMmMin, FeedrateMediumMmMin, FeedrateSmallMmMin,
                    FeedrateEngravingMmMin, GasPressureBar, NozzleSizeMm, Source, UpdatedBy)
                VALUES (@technology, @material, @thickness, @lens, @piercing, @vapor, @reduced, @large,
                    @medium, @small, @engraving, @pressure, @nozzle, 'program', @updatedBy);`);
        return true;
    } catch (err) {
        markUnavailable(err);
        return null;
    }
}
module.exports = { configure, isEnabled, getStatus, recordAftercalcSnapshot, getOrderTrend, saveRawImport, getAppState, getAppStatesByPrefix, setAppState, getBomLaserParameters, getBomLaserParameter, upsertBomLaserParameter, importBomLaserParameters, getBomLaserTechnicalParameters, getBomLaserTechnicalParameter, getBomLaserTechnicalParametersForMaterial, upsertBomLaserTechnicalParameter, importBomLaserTechnicalParameters, getBomBendingParameters, upsertBomBendingMachine, upsertBomBendingHandlingBand, addBomBendingActualSample, serverLabel: GOH_SERVER + '/' + GOH_DATABASE };
