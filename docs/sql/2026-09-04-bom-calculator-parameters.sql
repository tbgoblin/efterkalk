USE [GantechOperationHub];
GO

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
END;
GO

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
END;
GO

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
GO

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
GO

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
GO