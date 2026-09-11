/*
  GOH - storico Efterkalk / MaanedsDB
  SQL Server. Lo script e idempotente: puo essere eseguito piu volte.

  Modello:
  - EfterkalkSnapshotRun: una revisione di un periodo mensile.
  - EfterkalkOrderSnapshot: dettaglio ordine/fattura della revisione.
  - Le viste espongono i totali correnti per cliente e mese.
  - CLOSED e uno stato informativo: fatture/costi mancanti possono essere aggiornati.
*/

SET XACT_ABORT ON;
BEGIN TRANSACTION;

IF OBJECT_ID(N'dbo.EfterkalkSnapshotRun', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.EfterkalkSnapshotRun (
        SnapshotId       bigint IDENTITY(1,1) NOT NULL,
        PeriodStart      date NOT NULL,
        PeriodEnd        date NOT NULL,
        RevisionNo       int NOT NULL CONSTRAINT DF_EfterkalkSnapshotRun_Revision DEFAULT (1),
        SnapshotStatus   varchar(10) NOT NULL CONSTRAINT DF_EfterkalkSnapshotRun_Status DEFAULT ('OPEN'),
        IsCurrent        bit NOT NULL CONSTRAINT DF_EfterkalkSnapshotRun_Current DEFAULT (1),
        CalculationVersion nvarchar(50) NULL,
        StartedAt        datetime2(0) NOT NULL CONSTRAINT DF_EfterkalkSnapshotRun_Started DEFAULT (SYSUTCDATETIME()),
        CompletedAt      datetime2(0) NULL,
        CreatedAt        datetime2(0) NOT NULL CONSTRAINT DF_EfterkalkSnapshotRun_Created DEFAULT (SYSUTCDATETIME()),
        UpdatedAt        datetime2(0) NOT NULL CONSTRAINT DF_EfterkalkSnapshotRun_Updated DEFAULT (SYSUTCDATETIME()),
        UpdatedBy        nvarchar(100) NULL,
        CONSTRAINT PK_EfterkalkSnapshotRun PRIMARY KEY (SnapshotId),
        CONSTRAINT UQ_EfterkalkSnapshotRun_PeriodRevision UNIQUE (PeriodStart, RevisionNo),
        CONSTRAINT CK_EfterkalkSnapshotRun_Month CHECK (
            PeriodStart = DATEFROMPARTS(YEAR(PeriodStart), MONTH(PeriodStart), 1)
            AND PeriodEnd = EOMONTH(PeriodStart)
        ),
        CONSTRAINT CK_EfterkalkSnapshotRun_Status CHECK (SnapshotStatus IN ('OPEN', 'CLOSED', 'FAILED'))
    );

    CREATE UNIQUE INDEX UX_EfterkalkSnapshotRun_CurrentPeriod
        ON dbo.EfterkalkSnapshotRun (PeriodStart)
        WHERE IsCurrent = 1;
END;

IF OBJECT_ID(N'dbo.EfterkalkOrderSnapshot', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.EfterkalkOrderSnapshot (
        SnapshotOrderId bigint IDENTITY(1,1) NOT NULL,
        SnapshotId      bigint NOT NULL,
        OrdNo           bigint NOT NULL,
        OrderDate       date NULL,
        InvoiceNo       nvarchar(50) NOT NULL CONSTRAINT DF_EfterkalkOrderSnapshot_InvoiceNo DEFAULT (N''),
        InvoiceDate     date NOT NULL,
        CustNo          nvarchar(50) NOT NULL,
        CustomerName    nvarchar(250) NULL,
        Seller          nvarchar(100) NULL,
        Revenue         decimal(19,4) NOT NULL,
        Cost            decimal(19,4) NULL,
        ContributionMargin AS (
            CASE WHEN Cost IS NULL THEN NULL ELSE Revenue - Cost END
        ) PERSISTED,
        MarginPct AS (
            CASE
                WHEN Cost IS NULL OR Revenue = 0 THEN NULL
                ELSE ((Revenue - Cost) * CONVERT(decimal(19,6), 100)) / Revenue
            END
        ) PERSISTED,
        CostComplete    bit NOT NULL CONSTRAINT DF_EfterkalkOrderSnapshot_CostComplete DEFAULT (0),
        IsActive        bit NOT NULL CONSTRAINT DF_EfterkalkOrderSnapshot_Active DEFAULT (1),
        SourceHash      varbinary(32) NULL,
        CalculatedAt    datetime2(0) NOT NULL CONSTRAINT DF_EfterkalkOrderSnapshot_Calculated DEFAULT (SYSUTCDATETIME()),
        CreatedAt       datetime2(0) NOT NULL CONSTRAINT DF_EfterkalkOrderSnapshot_Created DEFAULT (SYSUTCDATETIME()),
        UpdatedAt       datetime2(0) NOT NULL CONSTRAINT DF_EfterkalkOrderSnapshot_Updated DEFAULT (SYSUTCDATETIME()),
        CONSTRAINT PK_EfterkalkOrderSnapshot PRIMARY KEY (SnapshotOrderId),
        CONSTRAINT FK_EfterkalkOrderSnapshot_Run FOREIGN KEY (SnapshotId)
            REFERENCES dbo.EfterkalkSnapshotRun (SnapshotId),
        CONSTRAINT UQ_EfterkalkOrderSnapshot_Source UNIQUE
            (SnapshotId, OrdNo, InvoiceNo, InvoiceDate),
        CONSTRAINT CK_EfterkalkOrderSnapshot_CostState CHECK (
            (CostComplete = 0 AND Cost IS NULL) OR
            (CostComplete = 1 AND Cost IS NOT NULL)
        )
    );

    CREATE INDEX IX_EfterkalkOrderSnapshot_Customer
        ON dbo.EfterkalkOrderSnapshot (SnapshotId, CustNo)
        INCLUDE (Revenue, Cost, CostComplete, IsActive);

    CREATE INDEX IX_EfterkalkOrderSnapshot_Order
        ON dbo.EfterkalkOrderSnapshot (OrdNo, InvoiceDate)
        INCLUDE (SnapshotId, CustNo, Revenue, Cost, IsActive);
END;

EXEC(N'
CREATE OR ALTER VIEW dbo.vw_EfterkalkCustomerCurrent
AS
    SELECT
        r.SnapshotId,
        r.PeriodStart,
        r.PeriodEnd,
        r.RevisionNo,
        o.CustNo,
        MAX(o.CustomerName) AS CustomerName,
        COUNT_BIG(*) AS InvoiceOrderCount,
        SUM(o.Revenue) AS Revenue,
        CASE WHEN SUM(CASE WHEN o.CostComplete = 1 THEN 1 ELSE 0 END) = COUNT_BIG(*)
             THEN SUM(o.Cost) END AS Cost,
        CASE WHEN SUM(CASE WHEN o.CostComplete = 1 THEN 1 ELSE 0 END) = COUNT_BIG(*)
             THEN SUM(o.Revenue) - SUM(o.Cost) END AS ContributionMargin,
        CASE WHEN SUM(CASE WHEN o.CostComplete = 1 THEN 1 ELSE 0 END) = COUNT_BIG(*)
                  AND SUM(o.Revenue) <> 0
             THEN ((SUM(o.Revenue) - SUM(o.Cost)) * CONVERT(decimal(19,6), 100)) / SUM(o.Revenue)
        END AS MarginPct,
        CONVERT(bit, CASE WHEN SUM(CASE WHEN o.CostComplete = 1 THEN 1 ELSE 0 END) = COUNT_BIG(*) THEN 1 ELSE 0 END) AS CostComplete
    FROM dbo.EfterkalkSnapshotRun r
    JOIN dbo.EfterkalkOrderSnapshot o ON o.SnapshotId = r.SnapshotId
    WHERE r.IsCurrent = 1 AND o.IsActive = 1
    GROUP BY r.SnapshotId, r.PeriodStart, r.PeriodEnd, r.RevisionNo, o.CustNo;
');

EXEC(N'
CREATE OR ALTER VIEW dbo.vw_EfterkalkMonthCurrent
AS
    SELECT
        r.SnapshotId,
        r.PeriodStart,
        r.PeriodEnd,
        r.RevisionNo,
        COUNT_BIG(*) AS InvoiceOrderCount,
        COUNT_BIG(DISTINCT o.CustNo) AS CustomerCount,
        SUM(o.Revenue) AS Revenue,
        CASE WHEN SUM(CASE WHEN o.CostComplete = 1 THEN 1 ELSE 0 END) = COUNT_BIG(*)
             THEN SUM(o.Cost) END AS Cost,
        CASE WHEN SUM(CASE WHEN o.CostComplete = 1 THEN 1 ELSE 0 END) = COUNT_BIG(*)
             THEN SUM(o.Revenue) - SUM(o.Cost) END AS ContributionMargin,
        CASE WHEN SUM(CASE WHEN o.CostComplete = 1 THEN 1 ELSE 0 END) = COUNT_BIG(*)
                  AND SUM(o.Revenue) <> 0
             THEN ((SUM(o.Revenue) - SUM(o.Cost)) * CONVERT(decimal(19,6), 100)) / SUM(o.Revenue)
        END AS MarginPct,
        CONVERT(bit, CASE WHEN SUM(CASE WHEN o.CostComplete = 1 THEN 1 ELSE 0 END) = COUNT_BIG(*) THEN 1 ELSE 0 END) AS CostComplete
    FROM dbo.EfterkalkSnapshotRun r
    JOIN dbo.EfterkalkOrderSnapshot o ON o.SnapshotId = r.SnapshotId
    WHERE r.IsCurrent = 1 AND o.IsActive = 1
    GROUP BY r.SnapshotId, r.PeriodStart, r.PeriodEnd, r.RevisionNo;
');

COMMIT TRANSACTION;

SELECT N'GOH Efterkalk snapshot schema ready' AS Result;
