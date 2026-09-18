/* =============================================================================
   Dim_Status  (role-playing: Start Status / End Status)
   Source: dbo.STATUS_LOOKUP     Business key: ADW_STATUS_KEY     SCD: Type 1
   ============================================================================= */

IF OBJECT_ID(N'star.Dim_Status', N'U') IS NULL
BEGIN
    CREATE TABLE star.Dim_Status
    (
        StatusKey                    INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_Dim_Status PRIMARY KEY,
        ADW_STATUS_KEY               INT            NOT NULL,   -- business key
        CertificateStatus            NVARCHAR(100)  NULL,
        ClassCode                    NVARCHAR(50)   NULL,       -- 'Active' / ... drives movement logic
        PremiumStatus                NVARCHAR(50)   NULL,
        StudyStatusCode              NVARCHAR(50)   NULL,
        StatusDescription            NVARCHAR(255)  NULL,
        DisabilityStatus             NVARCHAR(50)   NULL,
        AdminCertificateStatus       NVARCHAR(100)  NULL,
        AdminCertificateStatusGroup  NVARCHAR(100)  NULL,
        CONSTRAINT UQ_Dim_Status_ADW_STATUS_KEY UNIQUE (ADW_STATUS_KEY)
    );
END
GO

IF NOT EXISTS (SELECT 1 FROM star.Dim_Status WHERE StatusKey = -1)
BEGIN
    SET IDENTITY_INSERT star.Dim_Status ON;
    INSERT INTO star.Dim_Status (StatusKey, ADW_STATUS_KEY, CertificateStatus, ClassCode, StatusDescription)
    VALUES (-1, -1, N'Unknown', N'Unknown', N'Unknown');
    SET IDENTITY_INSERT star.Dim_Status OFF;
END
GO

CREATE OR ALTER VIEW star.stg_Dim_Status AS
SELECT
    ADW_STATUS_KEY              = s.ADW_STATUS_KEY,
    CertificateStatus           = LTRIM(RTRIM(s.CERTIFICATE_STATUS)),
    ClassCode                   = LTRIM(RTRIM(s.CLASS_CODE)),
    PremiumStatus               = LTRIM(RTRIM(s.PREMIUM_STATUS)),
    StudyStatusCode             = LTRIM(RTRIM(s.STUDY_STATUS_CODE)),
    StatusDescription           = LTRIM(RTRIM(s.STATUS_DESCRIPTION)),
    DisabilityStatus            = LTRIM(RTRIM(s.DISABILITY_STATUS)),
    AdminCertificateStatus      = LTRIM(RTRIM(s.ADMIN_CERTIFICATE_STATUS)),
    AdminCertificateStatusGroup = LTRIM(RTRIM(s.ADMIN_CERTIFICATE_STATUS_GROUP))
FROM dbo.STATUS_LOOKUP AS s;
GO

MERGE star.Dim_Status AS tgt
USING star.stg_Dim_Status AS src
    ON tgt.ADW_STATUS_KEY = src.ADW_STATUS_KEY
WHEN MATCHED THEN UPDATE SET
    tgt.CertificateStatus = src.CertificateStatus, tgt.ClassCode = src.ClassCode,
    tgt.PremiumStatus = src.PremiumStatus, tgt.StudyStatusCode = src.StudyStatusCode,
    tgt.StatusDescription = src.StatusDescription, tgt.DisabilityStatus = src.DisabilityStatus,
    tgt.AdminCertificateStatus = src.AdminCertificateStatus,
    tgt.AdminCertificateStatusGroup = src.AdminCertificateStatusGroup
WHEN NOT MATCHED BY TARGET THEN
    INSERT (ADW_STATUS_KEY, CertificateStatus, ClassCode, PremiumStatus, StudyStatusCode,
            StatusDescription, DisabilityStatus, AdminCertificateStatus, AdminCertificateStatusGroup)
    VALUES (src.ADW_STATUS_KEY, src.CertificateStatus, src.ClassCode, src.PremiumStatus,
            src.StudyStatusCode, src.StatusDescription, src.DisabilityStatus,
            src.AdminCertificateStatus, src.AdminCertificateStatusGroup);
GO
