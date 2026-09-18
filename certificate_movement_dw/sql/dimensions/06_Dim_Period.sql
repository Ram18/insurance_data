/* =============================================================================
   Dim_Period
   Source: dbo.MOVEMENT_PERIODS   Business key: PERIOD_KEY   SCD: Type 1
   Carries the reporting-period metadata (QoQ/YoY/YTD/MoM flags, length) and the
   start/end TIME_KEY used to resolve the fact's Start/End date roles.
   ============================================================================= */

IF OBJECT_ID(N'star.Dim_Period', N'U') IS NULL
BEGIN
    CREATE TABLE star.Dim_Period
    (
        PeriodKey            INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Dim_Period PRIMARY KEY,
        PERIOD_KEY           INT            NOT NULL,   -- business key
        PeriodCode           NVARCHAR(50)   NULL,
        PeriodStartCode      NVARCHAR(50)   NULL,
        PeriodEndCode        NVARCHAR(50)   NULL,
        TimeKeyStart         INT            NULL,       -- -> Dim_Date (Start role) via ETL
        TimeKeyEnd           INT            NULL,       -- -> Dim_Date (End role) via ETL
        YoyPeriods           NVARCHAR(20)   NULL,
        QoqPeriods           NVARCHAR(20)   NULL,
        YtdPeriods           NVARCHAR(20)   NULL,
        MomPeriods           NVARCHAR(20)   NULL,
        PeriodLoadFlag       NVARCHAR(10)   NULL,
        CurrentReportInd     NVARCHAR(10)   NULL,
        CurrentAnalysisInd   NVARCHAR(10)   NULL,
        PeriodLength         FLOAT          NULL,
        CONSTRAINT UQ_Dim_Period_PERIOD_KEY UNIQUE (PERIOD_KEY)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM star.Dim_Period WHERE PeriodKey = -1)
BEGIN
    SET IDENTITY_INSERT star.Dim_Period ON;
    INSERT INTO star.Dim_Period (PeriodKey, PERIOD_KEY, PeriodCode)
    VALUES (-1, -1, N'Unknown');
    SET IDENTITY_INSERT star.Dim_Period OFF;
END
GO

CREATE OR ALTER VIEW star.stg_Dim_Period AS
SELECT
    PERIOD_KEY         = p.PERIOD_KEY,
    PeriodCode         = LTRIM(RTRIM(p.PERIOD_CODE)),
    PeriodStartCode    = LTRIM(RTRIM(p.PERIOD_START_CODE)),
    PeriodEndCode      = LTRIM(RTRIM(p.PERIOD_END_CODE)),
    TimeKeyStart       = p.TIME_KEY_START,
    TimeKeyEnd         = p.TIME_KEY_END,
    YoyPeriods         = LTRIM(RTRIM(p.YOY_PERIODS)),
    QoqPeriods         = LTRIM(RTRIM(p.QOQ_PERIODS)),
    YtdPeriods         = LTRIM(RTRIM(p.YTD_PERIODS)),
    MomPeriods         = LTRIM(RTRIM(p.MOM_PERIODS)),
    PeriodLoadFlag     = LTRIM(RTRIM(p.PERIOD_LOAD_FLAG)),
    CurrentReportInd   = LTRIM(RTRIM(p.CURRENT_REPORT_IND)),
    CurrentAnalysisInd = LTRIM(RTRIM(p.CURRENT_ANALYSIS_IND)),
    PeriodLength       = p.PERIOD_LENGTH
FROM dbo.MOVEMENT_PERIODS AS p;
GO

MERGE star.Dim_Period AS tgt
USING star.stg_Dim_Period AS src
    ON tgt.PERIOD_KEY = src.PERIOD_KEY
WHEN MATCHED THEN UPDATE SET
    tgt.PeriodCode = src.PeriodCode, tgt.PeriodStartCode = src.PeriodStartCode, tgt.PeriodEndCode = src.PeriodEndCode,
    tgt.TimeKeyStart = src.TimeKeyStart, tgt.TimeKeyEnd = src.TimeKeyEnd,
    tgt.YoyPeriods = src.YoyPeriods, tgt.QoqPeriods = src.QoqPeriods, tgt.YtdPeriods = src.YtdPeriods,
    tgt.MomPeriods = src.MomPeriods, tgt.PeriodLoadFlag = src.PeriodLoadFlag,
    tgt.CurrentReportInd = src.CurrentReportInd, tgt.CurrentAnalysisInd = src.CurrentAnalysisInd,
    tgt.PeriodLength = src.PeriodLength
WHEN NOT MATCHED BY TARGET THEN
    INSERT (PERIOD_KEY, PeriodCode, PeriodStartCode, PeriodEndCode, TimeKeyStart, TimeKeyEnd, YoyPeriods,
            QoqPeriods, YtdPeriods, MomPeriods, PeriodLoadFlag, CurrentReportInd, CurrentAnalysisInd, PeriodLength)
    VALUES (src.PERIOD_KEY, src.PeriodCode, src.PeriodStartCode, src.PeriodEndCode, src.TimeKeyStart, src.TimeKeyEnd,
            src.YoyPeriods, src.QoqPeriods, src.YtdPeriods, src.MomPeriods, src.PeriodLoadFlag, src.CurrentReportInd,
            src.CurrentAnalysisInd, src.PeriodLength);
GO
