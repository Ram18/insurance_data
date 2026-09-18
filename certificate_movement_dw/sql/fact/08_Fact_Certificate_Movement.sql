/* =============================================================================
   Fact_Certificate_Movement  — periodic snapshot
   Grain: one row per certificate per movement period.
   Source: dbo.CERTIFICATE_MOVEMENT
   Loads AFTER all dimensions. Full reload (snapshot); re-runnable.
   ============================================================================= */

/* ---- DDL --------------------------------------------------------------------*/
IF OBJECT_ID(N'star.Fact_Certificate_Movement', N'U') IS NULL
BEGIN
    CREATE TABLE star.Fact_Certificate_Movement
    (
        FactKey             BIGINT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Fact_Certificate_Movement PRIMARY KEY,
        -- Surrogate foreign keys (default to -1 Unknown, never NULL)
        PeriodKey           INT     NOT NULL DEFAULT -1,
        CertificateKey      INT     NOT NULL DEFAULT -1,
        MovementKey         INT     NOT NULL DEFAULT -1,
        StartStatusKey      INT     NOT NULL DEFAULT -1,
        EndStatusKey        INT     NOT NULL DEFAULT -1,
        StartProductKey     INT     NOT NULL DEFAULT -1,
        EndProductKey       INT     NOT NULL DEFAULT -1,
        StartGeographyKey   INT     NOT NULL DEFAULT -1,
        EndGeographyKey     INT     NOT NULL DEFAULT -1,
        StartDateKey        INT     NOT NULL DEFAULT -1,
        EndDateKey          INT     NOT NULL DEFAULT -1,
        -- Degenerate dimension
        PolicyNo            NVARCHAR(50) NULL,
        -- Measures (additive)
        CertificateCount    INT     NOT NULL DEFAULT 1,
        DeathBenefitStart   FLOAT   NULL,
        DeathBenefitEnd     FLOAT   NULL,
        AnnualPaymentStart  FLOAT   NULL,
        AnnualPaymentEnd    FLOAT   NULL,
        FundValueStart      FLOAT   NULL,
        FundValueEnd        FLOAT   NULL,
        ReserveTotalAxisStart FLOAT NULL,
        ReserveTotalAxisEnd   FLOAT NULL,
        CONSTRAINT FK_Fact_Period      FOREIGN KEY (PeriodKey)         REFERENCES star.Dim_Period(PeriodKey),
        CONSTRAINT FK_Fact_Certificate FOREIGN KEY (CertificateKey)    REFERENCES star.Dim_Certificate(CertificateKey),
        CONSTRAINT FK_Fact_Movement    FOREIGN KEY (MovementKey)       REFERENCES star.Dim_Movement(MovementKey),
        CONSTRAINT FK_Fact_StatusStart FOREIGN KEY (StartStatusKey)    REFERENCES star.Dim_Status(StatusKey),
        CONSTRAINT FK_Fact_StatusEnd   FOREIGN KEY (EndStatusKey)      REFERENCES star.Dim_Status(StatusKey),
        CONSTRAINT FK_Fact_ProductStart FOREIGN KEY (StartProductKey)  REFERENCES star.Dim_Product(ProductKey),
        CONSTRAINT FK_Fact_ProductEnd   FOREIGN KEY (EndProductKey)    REFERENCES star.Dim_Product(ProductKey),
        CONSTRAINT FK_Fact_GeoStart    FOREIGN KEY (StartGeographyKey) REFERENCES star.Dim_Geography(GeographyKey),
        CONSTRAINT FK_Fact_GeoEnd      FOREIGN KEY (EndGeographyKey)   REFERENCES star.Dim_Geography(GeographyKey),
        CONSTRAINT FK_Fact_DateStart   FOREIGN KEY (StartDateKey)      REFERENCES star.Dim_Date(DateKey),
        CONSTRAINT FK_Fact_DateEnd     FOREIGN KEY (EndDateKey)        REFERENCES star.Dim_Date(DateKey)
    );
    CREATE INDEX IX_Fact_CM_Period      ON star.Fact_Certificate_Movement(PeriodKey);
    CREATE INDEX IX_Fact_CM_Movement    ON star.Fact_Certificate_Movement(MovementKey);
    CREATE INDEX IX_Fact_CM_Certificate ON star.Fact_Certificate_Movement(CertificateKey);
END
GO

/* ---- Staging view: resolve every natural key to its surrogate ---------------*/
CREATE OR ALTER VIEW star.stg_Fact_Certificate_Movement AS
SELECT
    PeriodKey         = ISNULL(per.PeriodKey, -1),
    CertificateKey    = ISNULL(dc.CertificateKey, -1),
    MovementKey       = ISNULL(dm.MovementKey, -1),
    StartStatusKey    = ISNULL(ss.StatusKey, -1),
    EndStatusKey      = ISNULL(se.StatusKey, -1),
    StartProductKey   = ISNULL(ps.ProductKey, -1),
    EndProductKey     = ISNULL(pe.ProductKey, -1),
    StartGeographyKey = ISNULL(gs.GeographyKey, -1),
    EndGeographyKey   = ISNULL(ge.GeographyKey, -1),
    StartDateKey      = ISNULL(ds.DateKey, -1),
    EndDateKey        = ISNULL(de.DateKey, -1),
    PolicyNo          = LTRIM(RTRIM(cm.POLICY_NO)),
    CertificateCount  = 1,
    DeathBenefitStart     = cm.BENEFIT_DEATH_TOTAL_START,
    DeathBenefitEnd       = cm.BENEFIT_DEATH_TOTAL_END,
    AnnualPaymentStart    = cm.BENEFIT_INCOME_ANNUAL_START,
    AnnualPaymentEnd      = cm.BENEFIT_INCOME_ANNUALISED_END,
    FundValueStart        = cm.FUND_VALUE_START,
    FundValueEnd          = cm.FUND_VALUE_END,
    ReserveTotalAxisStart = cm.USS_RESERVE_TOTAL_AXIS_START,
    ReserveTotalAxisEnd   = cm.USS_RESERVE_TOTAL_AXIS_END
FROM dbo.CERTIFICATE_MOVEMENT AS cm
LEFT JOIN star.Dim_Period      AS per ON per.PERIOD_KEY          = cm.PERIOD_KEY
LEFT JOIN star.Dim_Certificate AS dc  ON dc.ADW_CERTIFICATE_KEY  = LTRIM(RTRIM(cm.ADW_CERTIFICATE_KEY))
LEFT JOIN star.Dim_Movement    AS dm  ON dm.ADW_STATUS_KEY_START = cm.ADW_STATUS_KEY_START
                                     AND dm.ADW_STATUS_KEY_END   = cm.ADW_STATUS_KEY_END
LEFT JOIN star.Dim_Status      AS ss  ON ss.ADW_STATUS_KEY       = cm.ADW_STATUS_KEY_START
LEFT JOIN star.Dim_Status      AS se  ON se.ADW_STATUS_KEY       = cm.ADW_STATUS_KEY_END
LEFT JOIN star.Dim_Product     AS ps  ON ps.PRODUCT_KEY          = cm.PRODUCT_KEY_START
LEFT JOIN star.Dim_Product     AS pe  ON pe.PRODUCT_KEY          = cm.PRODUCT_KEY_END
LEFT JOIN star.Dim_Geography   AS gs  ON gs.JURISDICTION_KEY     = LTRIM(RTRIM(cm.JURISDICTION_KEY_START))
LEFT JOIN star.Dim_Geography   AS ge  ON ge.JURISDICTION_KEY     = LTRIM(RTRIM(cm.JURISDICTION_KEY_END))
LEFT JOIN star.Dim_Date        AS ds  ON ds.TIME_KEY             = per.TimeKeyStart
LEFT JOIN star.Dim_Date        AS de  ON de.TIME_KEY             = per.TimeKeyEnd;
GO

/* ---- Full reload ------------------------------------------------------------*/
TRUNCATE TABLE star.Fact_Certificate_Movement;
GO
INSERT INTO star.Fact_Certificate_Movement
    (PeriodKey, CertificateKey, MovementKey, StartStatusKey, EndStatusKey, StartProductKey, EndProductKey,
     StartGeographyKey, EndGeographyKey, StartDateKey, EndDateKey, PolicyNo, CertificateCount,
     DeathBenefitStart, DeathBenefitEnd, AnnualPaymentStart, AnnualPaymentEnd, FundValueStart, FundValueEnd,
     ReserveTotalAxisStart, ReserveTotalAxisEnd)
SELECT
     PeriodKey, CertificateKey, MovementKey, StartStatusKey, EndStatusKey, StartProductKey, EndProductKey,
     StartGeographyKey, EndGeographyKey, StartDateKey, EndDateKey, PolicyNo, CertificateCount,
     DeathBenefitStart, DeathBenefitEnd, AnnualPaymentStart, AnnualPaymentEnd, FundValueStart, FundValueEnd,
     ReserveTotalAxisStart, ReserveTotalAxisEnd
FROM star.stg_Fact_Certificate_Movement;
GO
