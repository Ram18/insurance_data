/* =============================================================================
   Dim_Certificate
   Source: dbo.CERTIFICATE   Business key: ADW_CERTIFICATE_KEY   SCD: Type 1
   Curated subset of the 109 source columns (add more attributes as reports need
   them; keep it to descriptive, non-measure fields).
   ============================================================================= */

IF OBJECT_ID(N'star.Dim_Certificate', N'U') IS NULL
BEGIN
    CREATE TABLE star.Dim_Certificate
    (
        CertificateKey        INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Dim_Certificate PRIMARY KEY,
        ADW_CERTIFICATE_KEY   NVARCHAR(50)   NOT NULL,   -- business key (text)
        CertificateNumber     NVARCHAR(50)   NULL,
        PolicyNo              NVARCHAR(50)   NULL,
        OriginalCertificateNo NVARCHAR(50)   NULL,
        PlanCode              NVARCHAR(50)   NULL,
        ProductTypeCode       NVARCHAR(50)   NULL,
        CertificateContractType NVARCHAR(50) NULL,
        IssueDate             DATE           NULL,
        IssueAge              INT            NULL,
        Gender                NVARCHAR(20)   NULL,
        CurrentSmokerStatus   NVARCHAR(20)   NULL,
        UnderwritingClass     NVARCHAR(50)   NULL,
        CurrentStatus         NVARCHAR(50)   NULL,
        CurrentRiskStatusGroup NVARCHAR(50)  NULL,
        DeathStatus           NVARCHAR(50)   NULL,
        RescindedIndicator    NVARCHAR(10)   NULL,
        MecIndicator          NVARCHAR(10)   NULL,
        QualifiedIndicator    NVARCHAR(10)   NULL,
        AgentName             NVARCHAR(150)  NULL,
        AgentNumber           NVARCHAR(50)   NULL,
        AgentStatus           NVARCHAR(50)   NULL,
        MarketingGroupCode    NVARCHAR(50)   NULL,
        SourceSystem          NVARCHAR(50)   NULL,
        CountryCode           NVARCHAR(20)   NULL,
        ZipPostalCode         NVARCHAR(20)   NULL,
        CurrentIndicator      NVARCHAR(10)   NULL,
        CONSTRAINT UQ_Dim_Certificate_BK UNIQUE (ADW_CERTIFICATE_KEY)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM star.Dim_Certificate WHERE CertificateKey = -1)
BEGIN
    SET IDENTITY_INSERT star.Dim_Certificate ON;
    INSERT INTO star.Dim_Certificate (CertificateKey, ADW_CERTIFICATE_KEY, CertificateNumber, CurrentStatus)
    VALUES (-1, N'-1', N'Unknown', N'Unknown');
    SET IDENTITY_INSERT star.Dim_Certificate OFF;
END
GO

CREATE OR ALTER VIEW star.stg_Dim_Certificate AS
SELECT
    ADW_CERTIFICATE_KEY      = LTRIM(RTRIM(c.ADW_CERTIFICATE_KEY)),
    CertificateNumber        = LTRIM(RTRIM(c.CERTIFICATE_NUMBER)),
    PolicyNo                 = LTRIM(RTRIM(c.POLICY_NO)),
    OriginalCertificateNo    = LTRIM(RTRIM(c.ORIGINAL_CERTIFICATE_NUMBER)),
    PlanCode                 = LTRIM(RTRIM(c.PLAN_CODE)),
    ProductTypeCode          = LTRIM(RTRIM(c.PRODUCT_TYPE_CODE)),
    CertificateContractType  = LTRIM(RTRIM(c.CERTIFICATE_CONTRACT_TYPE)),
    IssueDate                = TRY_CONVERT(DATE, c.ISSUE_DATE),
    IssueAge                 = c.ISSUE_AGE,
    Gender                   = LTRIM(RTRIM(c.GENDER)),
    CurrentSmokerStatus      = LTRIM(RTRIM(c.CURRENT_SMOKER_STATUS)),
    UnderwritingClass        = LTRIM(RTRIM(c.UNDERWRITING_CLASS)),
    CurrentStatus            = LTRIM(RTRIM(c.CURRENT_STATUS)),
    CurrentRiskStatusGroup   = LTRIM(RTRIM(c.CURRENT_RISK_STATUS_GROUP)),
    DeathStatus              = LTRIM(RTRIM(c.DEATH_STATUS)),
    RescindedIndicator       = LTRIM(RTRIM(c.Rescinded_Indicator)),
    MecIndicator             = LTRIM(RTRIM(c.MEC_INDICATOR)),
    QualifiedIndicator       = LTRIM(RTRIM(c.QUALIFIED_INDICATOR)),
    AgentName                = LTRIM(RTRIM(c.AGENT_NAME)),
    AgentNumber              = LTRIM(RTRIM(c.AGENT_NUMBER)),
    AgentStatus              = LTRIM(RTRIM(c.AGENT_STATUS)),
    MarketingGroupCode       = LTRIM(RTRIM(c.MARKETING_GROUP_CODE)),
    SourceSystem             = LTRIM(RTRIM(c.SOURCE_SYSTEM)),
    CountryCode              = LTRIM(RTRIM(c.COUNTRY_CODE)),
    ZipPostalCode            = LTRIM(RTRIM(c.ZIP_POSTAL_CODE)),
    CurrentIndicator         = LTRIM(RTRIM(c.CURRENT_INDICATOR))
FROM dbo.CERTIFICATE AS c;
GO

MERGE star.Dim_Certificate AS tgt
USING star.stg_Dim_Certificate AS src
    ON tgt.ADW_CERTIFICATE_KEY = src.ADW_CERTIFICATE_KEY
WHEN MATCHED THEN UPDATE SET
    tgt.CertificateNumber = src.CertificateNumber, tgt.PolicyNo = src.PolicyNo,
    tgt.OriginalCertificateNo = src.OriginalCertificateNo, tgt.PlanCode = src.PlanCode,
    tgt.ProductTypeCode = src.ProductTypeCode, tgt.CertificateContractType = src.CertificateContractType,
    tgt.IssueDate = src.IssueDate, tgt.IssueAge = src.IssueAge, tgt.Gender = src.Gender,
    tgt.CurrentSmokerStatus = src.CurrentSmokerStatus, tgt.UnderwritingClass = src.UnderwritingClass,
    tgt.CurrentStatus = src.CurrentStatus, tgt.CurrentRiskStatusGroup = src.CurrentRiskStatusGroup,
    tgt.DeathStatus = src.DeathStatus, tgt.RescindedIndicator = src.RescindedIndicator,
    tgt.MecIndicator = src.MecIndicator, tgt.QualifiedIndicator = src.QualifiedIndicator,
    tgt.AgentName = src.AgentName, tgt.AgentNumber = src.AgentNumber, tgt.AgentStatus = src.AgentStatus,
    tgt.MarketingGroupCode = src.MarketingGroupCode, tgt.SourceSystem = src.SourceSystem,
    tgt.CountryCode = src.CountryCode, tgt.ZipPostalCode = src.ZipPostalCode,
    tgt.CurrentIndicator = src.CurrentIndicator
WHEN NOT MATCHED BY TARGET THEN
    INSERT (ADW_CERTIFICATE_KEY, CertificateNumber, PolicyNo, OriginalCertificateNo, PlanCode, ProductTypeCode,
            CertificateContractType, IssueDate, IssueAge, Gender, CurrentSmokerStatus, UnderwritingClass,
            CurrentStatus, CurrentRiskStatusGroup, DeathStatus, RescindedIndicator, MecIndicator, QualifiedIndicator,
            AgentName, AgentNumber, AgentStatus, MarketingGroupCode, SourceSystem, CountryCode, ZipPostalCode,
            CurrentIndicator)
    VALUES (src.ADW_CERTIFICATE_KEY, src.CertificateNumber, src.PolicyNo, src.OriginalCertificateNo, src.PlanCode,
            src.ProductTypeCode, src.CertificateContractType, src.IssueDate, src.IssueAge, src.Gender,
            src.CurrentSmokerStatus, src.UnderwritingClass, src.CurrentStatus, src.CurrentRiskStatusGroup,
            src.DeathStatus, src.RescindedIndicator, src.MecIndicator, src.QualifiedIndicator, src.AgentName,
            src.AgentNumber, src.AgentStatus, src.MarketingGroupCode, src.SourceSystem, src.CountryCode,
            src.ZipPostalCode, src.CurrentIndicator);
GO
