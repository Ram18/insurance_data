/* =============================================================================
   Dim_Product  (role-playing: Start Product / End Product)
   Source: dbo.PRODUCT_LOOKUP    Business key: PRODUCT_KEY    SCD: Type 1
   ============================================================================= */

IF OBJECT_ID(N'star.Dim_Product', N'U') IS NULL
BEGIN
    CREATE TABLE star.Dim_Product
    (
        ProductKey               INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Dim_Product PRIMARY KEY,
        PRODUCT_KEY              INT            NOT NULL,   -- business key
        ProductCode              NVARCHAR(50)   NULL,
        ProductName              NVARCHAR(255)  NULL,
        ProductType              NVARCHAR(100)  NULL,
        BusinessLine             NVARCHAR(100)  NULL,       -- 'Life' drives Active Increase/Decrease logic
        BusinessSubLine          NVARCHAR(100)  NULL,
        BusinessFund             NVARCHAR(100)  NULL,
        ProductSeries            NVARCHAR(100)  NULL,
        Series                   NVARCHAR(100)  NULL,
        PremiumType              NVARCHAR(50)   NULL,
        IssueType                NVARCHAR(50)   NULL,
        IssuingCompany           NVARCHAR(100)  NULL,
        ReportingBusinessGroup   NVARCHAR(100)  NULL,
        ReportingEntity          NVARCHAR(100)  NULL,
        CONSTRAINT UQ_Dim_Product_PRODUCT_KEY UNIQUE (PRODUCT_KEY)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM star.Dim_Product WHERE ProductKey = -1)
BEGIN
    SET IDENTITY_INSERT star.Dim_Product ON;
    INSERT INTO star.Dim_Product (ProductKey, PRODUCT_KEY, ProductName, BusinessLine, ProductSeries, ReportingEntity)
    VALUES (-1, -1, N'Unknown', N'Unknown', N'Unknown', N'Unknown');
    SET IDENTITY_INSERT star.Dim_Product OFF;
END
GO

CREATE OR ALTER VIEW star.stg_Dim_Product AS
SELECT
    PRODUCT_KEY            = p.PRODUCT_KEY,
    ProductCode            = LTRIM(RTRIM(p.PRODUCT_CODE)),
    ProductName            = LTRIM(RTRIM(p.PRODUCT_NAME)),
    ProductType            = LTRIM(RTRIM(p.PRODUCT_TYPE)),
    BusinessLine           = LTRIM(RTRIM(p.BUSINESS_LINE)),
    BusinessSubLine        = LTRIM(RTRIM(p.BUSINESS_SUB_LINE)),
    BusinessFund           = LTRIM(RTRIM(p.BUSINESS_FUND)),
    ProductSeries          = LTRIM(RTRIM(p.PRODUCT_SERIES)),
    Series                 = LTRIM(RTRIM(p.SERIES)),
    PremiumType            = LTRIM(RTRIM(p.PREMIUM_TYPE)),
    IssueType              = LTRIM(RTRIM(p.ISSUE_TYPE)),
    IssuingCompany         = LTRIM(RTRIM(p.ISSUING_COMPANY)),
    ReportingBusinessGroup = LTRIM(RTRIM(p.REPORTING_BUSINESS_GROUP)),
    ReportingEntity        = LTRIM(RTRIM(p.REPORTING_ENTITY))
FROM dbo.PRODUCT_LOOKUP AS p;
GO

MERGE star.Dim_Product AS tgt
USING star.stg_Dim_Product AS src
    ON tgt.PRODUCT_KEY = src.PRODUCT_KEY
WHEN MATCHED THEN UPDATE SET
    tgt.ProductCode = src.ProductCode, tgt.ProductName = src.ProductName, tgt.ProductType = src.ProductType,
    tgt.BusinessLine = src.BusinessLine, tgt.BusinessSubLine = src.BusinessSubLine, tgt.BusinessFund = src.BusinessFund,
    tgt.ProductSeries = src.ProductSeries, tgt.Series = src.Series, tgt.PremiumType = src.PremiumType,
    tgt.IssueType = src.IssueType, tgt.IssuingCompany = src.IssuingCompany,
    tgt.ReportingBusinessGroup = src.ReportingBusinessGroup, tgt.ReportingEntity = src.ReportingEntity
WHEN NOT MATCHED BY TARGET THEN
    INSERT (PRODUCT_KEY, ProductCode, ProductName, ProductType, BusinessLine, BusinessSubLine, BusinessFund,
            ProductSeries, Series, PremiumType, IssueType, IssuingCompany, ReportingBusinessGroup, ReportingEntity)
    VALUES (src.PRODUCT_KEY, src.ProductCode, src.ProductName, src.ProductType, src.BusinessLine, src.BusinessSubLine,
            src.BusinessFund, src.ProductSeries, src.Series, src.PremiumType, src.IssueType, src.IssuingCompany,
            src.ReportingBusinessGroup, src.ReportingEntity);
GO
