/* =============================================================================
   Dim_Geography  (role-playing: Start Jurisdiction / End Jurisdiction)
   Source: dbo.GEOGRAPHIC_LOOKUP   Business key: JURISDICTION_KEY   SCD: Type 1
   Note: JURISDICTION_KEY is text in the source.
   ============================================================================= */

IF OBJECT_ID(N'star.Dim_Geography', N'U') IS NULL
BEGIN
    CREATE TABLE star.Dim_Geography
    (
        GeographyKey            INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Dim_Geography PRIMARY KEY,
        JURISDICTION_KEY        NVARCHAR(50)   NOT NULL,   -- business key (text)
        JurisdictionCode        NVARCHAR(50)   NULL,
        JurisdictionDescription NVARCHAR(255)  NULL,
        StateCode               NVARCHAR(20)   NULL,       -- drives Relocate State logic
        StateName               NVARCHAR(100)  NULL,
        RegionCode              NVARCHAR(20)   NULL,
        RegionName              NVARCHAR(100)  NULL,
        CountryCode             NVARCHAR(20)   NULL,
        CountryDescription      NVARCHAR(100)  NULL,
        ContinentCode           NVARCHAR(20)   NULL,
        ContinentDescription    NVARCHAR(100)  NULL,
        CONSTRAINT UQ_Dim_Geography_JURISDICTION_KEY UNIQUE (JURISDICTION_KEY)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM star.Dim_Geography WHERE GeographyKey = -1)
BEGIN
    SET IDENTITY_INSERT star.Dim_Geography ON;
    INSERT INTO star.Dim_Geography (GeographyKey, JURISDICTION_KEY, JurisdictionDescription, StateCode, CountryCode)
    VALUES (-1, N'-1', N'Unknown', N'Unknown', N'Unknown');
    SET IDENTITY_INSERT star.Dim_Geography OFF;
END
GO

CREATE OR ALTER VIEW star.stg_Dim_Geography AS
SELECT
    JURISDICTION_KEY        = LTRIM(RTRIM(g.JURISDICTION_KEY)),
    JurisdictionCode        = LTRIM(RTRIM(g.JURISDICTION_CODE)),
    JurisdictionDescription = LTRIM(RTRIM(g.JURISDICTION_DESCRIPTION)),
    StateCode               = LTRIM(RTRIM(g.STATE_CODE)),
    StateName               = LTRIM(RTRIM(g.STATE_NAME)),
    RegionCode              = LTRIM(RTRIM(g.REGION_CODE)),
    RegionName              = LTRIM(RTRIM(g.REGION_NAME)),
    CountryCode             = LTRIM(RTRIM(g.COUNTRY_CODE)),
    CountryDescription      = LTRIM(RTRIM(g.COUNTRY_DESCRIPTION)),
    ContinentCode           = LTRIM(RTRIM(g.CONTINENT_CODE)),
    ContinentDescription    = LTRIM(RTRIM(g.CONTINENT_DESCRIPTION))
FROM dbo.GEOGRAPHIC_LOOKUP AS g;
GO

MERGE star.Dim_Geography AS tgt
USING star.stg_Dim_Geography AS src
    ON tgt.JURISDICTION_KEY = src.JURISDICTION_KEY
WHEN MATCHED THEN UPDATE SET
    tgt.JurisdictionCode = src.JurisdictionCode, tgt.JurisdictionDescription = src.JurisdictionDescription,
    tgt.StateCode = src.StateCode, tgt.StateName = src.StateName, tgt.RegionCode = src.RegionCode,
    tgt.RegionName = src.RegionName, tgt.CountryCode = src.CountryCode,
    tgt.CountryDescription = src.CountryDescription, tgt.ContinentCode = src.ContinentCode,
    tgt.ContinentDescription = src.ContinentDescription
WHEN NOT MATCHED BY TARGET THEN
    INSERT (JURISDICTION_KEY, JurisdictionCode, JurisdictionDescription, StateCode, StateName,
            RegionCode, RegionName, CountryCode, CountryDescription, ContinentCode, ContinentDescription)
    VALUES (src.JURISDICTION_KEY, src.JurisdictionCode, src.JurisdictionDescription, src.StateCode, src.StateName,
            src.RegionCode, src.RegionName, src.CountryCode, src.CountryDescription, src.ContinentCode,
            src.ContinentDescription);
GO
