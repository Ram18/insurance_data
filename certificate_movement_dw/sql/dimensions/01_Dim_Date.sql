/* =============================================================================
   Dim_Date  (role-playing: Period Start Date / Period End Date)
   Source: dbo.TIME_LOOKUP        Business key: TIME_KEY        SCD: Type 1
   ============================================================================= */

/* ---- DDL --------------------------------------------------------------------*/
IF OBJECT_ID(N'star.Dim_Date', N'U') IS NULL
BEGIN
    CREATE TABLE star.Dim_Date
    (
        DateKey            INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Dim_Date PRIMARY KEY,
        TIME_KEY           INT            NOT NULL,     -- business (natural) key
        AsOfDate           DATE           NULL,
        DayNumber          INT            NULL,
        DayName            NVARCHAR(20)   NULL,
        MonthNumber        INT            NULL,
        MonthName          NVARCHAR(20)   NULL,
        MonthAbbr          NVARCHAR(10)   NULL,
        QuarterNumber      INT            NULL,
        YearNumber         INT            NULL,
        MonthYearName      NVARCHAR(30)   NULL,
        QuarterYearName    NVARCHAR(30)   NULL,
        JulianDate         INT            NULL,
        CurrentDateInd     NVARCHAR(5)    NULL,
        CONSTRAINT UQ_Dim_Date_TIME_KEY UNIQUE (TIME_KEY)
    );
END
GO

/* ---- Unknown member (-1) ----------------------------------------------------*/
IF NOT EXISTS (SELECT 1 FROM star.Dim_Date WHERE DateKey = -1)
BEGIN
    SET IDENTITY_INSERT star.Dim_Date ON;
    INSERT INTO star.Dim_Date (DateKey, TIME_KEY, AsOfDate, DayName, MonthName, MonthYearName, QuarterYearName)
    VALUES (-1, -1, NULL, N'Unknown', N'Unknown', N'Unknown', N'Unknown');
    SET IDENTITY_INSERT star.Dim_Date OFF;
END
GO

/* ---- Staging view (trim text, matches source Power Query Text.Trim) ---------*/
CREATE OR ALTER VIEW star.stg_Dim_Date AS
SELECT
    TIME_KEY        = t.TIME_KEY,
    AsOfDate        = TRY_CONVERT(DATE, t.AS_OF_DATE),
    DayNumber       = t.DAY_NUMBER,
    DayName         = LTRIM(RTRIM(t.DAY_NAME)),
    MonthNumber     = t.MONTH_NUMBER,
    MonthName       = LTRIM(RTRIM(t.MONTH_NAME)),
    MonthAbbr       = LTRIM(RTRIM(t.MONTH_ABBR)),
    QuarterNumber   = t.QUARTER_NUMBER,
    YearNumber      = t.YEAR_NUMBER,
    MonthYearName   = LTRIM(RTRIM(t.MONTH_YEAR_NAME)),
    QuarterYearName = LTRIM(RTRIM(t.QUARTER_YEAR_NAME)),
    JulianDate      = t.JULIAN_DATE,
    CurrentDateInd  = LTRIM(RTRIM(t.CURRENT_DATE_IND))
FROM dbo.TIME_LOOKUP AS t;
GO

/* ---- Type 1 load (idempotent MERGE) -----------------------------------------*/
MERGE star.Dim_Date AS tgt
USING star.stg_Dim_Date AS src
    ON tgt.TIME_KEY = src.TIME_KEY
WHEN MATCHED THEN UPDATE SET
    tgt.AsOfDate = src.AsOfDate, tgt.DayNumber = src.DayNumber, tgt.DayName = src.DayName,
    tgt.MonthNumber = src.MonthNumber, tgt.MonthName = src.MonthName, tgt.MonthAbbr = src.MonthAbbr,
    tgt.QuarterNumber = src.QuarterNumber, tgt.YearNumber = src.YearNumber,
    tgt.MonthYearName = src.MonthYearName, tgt.QuarterYearName = src.QuarterYearName,
    tgt.JulianDate = src.JulianDate, tgt.CurrentDateInd = src.CurrentDateInd
WHEN NOT MATCHED BY TARGET THEN
    INSERT (TIME_KEY, AsOfDate, DayNumber, DayName, MonthNumber, MonthName, MonthAbbr,
            QuarterNumber, YearNumber, MonthYearName, QuarterYearName, JulianDate, CurrentDateInd)
    VALUES (src.TIME_KEY, src.AsOfDate, src.DayNumber, src.DayName, src.MonthNumber, src.MonthName,
            src.MonthAbbr, src.QuarterNumber, src.YearNumber, src.MonthYearName, src.QuarterYearName,
            src.JulianDate, src.CurrentDateInd);
GO
