-- =============================================================================
-- Fact_Certificate_Movement  — periodic snapshot (grain: certificate x period)
-- Source: adw_prod10.dbo.CERTIFICATE_MOVEMENT.  Load AFTER all dimensions.
-- Full reload via INSERT OVERWRITE (idempotent).
-- =============================================================================

CREATE TABLE IF NOT EXISTS adw_prod10.star.Fact_Certificate_Movement (
    PeriodKey             BIGINT NOT NULL,
    CertificateKey        BIGINT NOT NULL,
    MovementKey           BIGINT NOT NULL,
    StartStatusKey        BIGINT NOT NULL,
    EndStatusKey          BIGINT NOT NULL,
    StartProductKey       BIGINT NOT NULL,
    EndProductKey         BIGINT NOT NULL,
    StartGeographyKey     BIGINT NOT NULL,
    EndGeographyKey       BIGINT NOT NULL,
    StartDateKey          BIGINT NOT NULL,
    EndDateKey            BIGINT NOT NULL,
    PolicyNo              STRING COMMENT 'Degenerate dimension',
    CertificateCount      INT,
    DeathBenefitStart     DOUBLE,
    DeathBenefitEnd       DOUBLE,
    AnnualPaymentStart    DOUBLE,
    AnnualPaymentEnd      DOUBLE,
    FundValueStart        DOUBLE,
    FundValueEnd          DOUBLE,
    ReserveTotalAxisStart DOUBLE,
    ReserveTotalAxisEnd   DOUBLE,
    CONSTRAINT fk_Fact_Period      FOREIGN KEY (PeriodKey)         REFERENCES adw_prod10.star.Dim_Period(PeriodKey)      NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_Certificate FOREIGN KEY (CertificateKey)    REFERENCES adw_prod10.star.Dim_Certificate(CertificateKey) NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_Movement    FOREIGN KEY (MovementKey)       REFERENCES adw_prod10.star.Dim_Movement(MovementKey)  NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_StatusStart FOREIGN KEY (StartStatusKey)    REFERENCES adw_prod10.star.Dim_Status(StatusKey)      NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_StatusEnd   FOREIGN KEY (EndStatusKey)      REFERENCES adw_prod10.star.Dim_Status(StatusKey)      NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_ProductStart FOREIGN KEY (StartProductKey)  REFERENCES adw_prod10.star.Dim_Product(ProductKey)    NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_ProductEnd   FOREIGN KEY (EndProductKey)    REFERENCES adw_prod10.star.Dim_Product(ProductKey)    NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_GeoStart    FOREIGN KEY (StartGeographyKey) REFERENCES adw_prod10.star.Dim_Geography(GeographyKey) NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_GeoEnd      FOREIGN KEY (EndGeographyKey)   REFERENCES adw_prod10.star.Dim_Geography(GeographyKey) NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_DateStart   FOREIGN KEY (StartDateKey)      REFERENCES adw_prod10.star.Dim_Date(DateKey)          NOT ENFORCED RELY,
    CONSTRAINT fk_Fact_DateEnd     FOREIGN KEY (EndDateKey)        REFERENCES adw_prod10.star.Dim_Date(DateKey)          NOT ENFORCED RELY
) USING DELTA
  COMMENT 'Periodic-snapshot fact: one row per certificate per movement period'
  TBLPROPERTIES (delta.feature.allowColumnDefaults = 'supported');

-- Staging: resolve every natural key to its surrogate (Unknown = -1)
CREATE OR REPLACE VIEW adw_prod10.star.stg_Fact_Certificate_Movement AS
SELECT
    coalesce(per.PeriodKey, -1)      AS PeriodKey,
    coalesce(dc.CertificateKey, -1)  AS CertificateKey,
    coalesce(dm.MovementKey, -1)     AS MovementKey,
    coalesce(ss.StatusKey, -1)       AS StartStatusKey,
    coalesce(se.StatusKey, -1)       AS EndStatusKey,
    coalesce(ps.ProductKey, -1)      AS StartProductKey,
    coalesce(pe.ProductKey, -1)      AS EndProductKey,
    coalesce(gs.GeographyKey, -1)    AS StartGeographyKey,
    coalesce(ge.GeographyKey, -1)    AS EndGeographyKey,
    coalesce(ds.DateKey, -1)         AS StartDateKey,
    coalesce(de.DateKey, -1)         AS EndDateKey,
    trim(cm.POLICY_NO)               AS PolicyNo,
    1                                AS CertificateCount,
    cm.BENEFIT_DEATH_TOTAL_START     AS DeathBenefitStart,
    cm.BENEFIT_DEATH_TOTAL_END       AS DeathBenefitEnd,
    cm.BENEFIT_INCOME_ANNUAL_START   AS AnnualPaymentStart,
    cm.BENEFIT_INCOME_ANNUALISED_END AS AnnualPaymentEnd,
    cm.FUND_VALUE_START              AS FundValueStart,
    cm.FUND_VALUE_END                AS FundValueEnd,
    cm.USS_RESERVE_TOTAL_AXIS_START  AS ReserveTotalAxisStart,
    cm.USS_RESERVE_TOTAL_AXIS_END    AS ReserveTotalAxisEnd
FROM adw_prod10.dbo.CERTIFICATE_MOVEMENT AS cm
LEFT JOIN adw_prod10.star.Dim_Period      AS per ON per.PERIOD_KEY          = cm.PERIOD_KEY
LEFT JOIN adw_prod10.star.Dim_Certificate AS dc  ON dc.ADW_CERTIFICATE_KEY  = trim(cm.ADW_CERTIFICATE_KEY)
LEFT JOIN adw_prod10.star.Dim_Movement    AS dm  ON dm.ADW_STATUS_KEY_START = cm.ADW_STATUS_KEY_START
                                               AND dm.ADW_STATUS_KEY_END   = cm.ADW_STATUS_KEY_END
LEFT JOIN adw_prod10.star.Dim_Status      AS ss  ON ss.ADW_STATUS_KEY       = cm.ADW_STATUS_KEY_START
LEFT JOIN adw_prod10.star.Dim_Status      AS se  ON se.ADW_STATUS_KEY       = cm.ADW_STATUS_KEY_END
LEFT JOIN adw_prod10.star.Dim_Product     AS ps  ON ps.PRODUCT_KEY          = cm.PRODUCT_KEY_START
LEFT JOIN adw_prod10.star.Dim_Product     AS pe  ON pe.PRODUCT_KEY          = cm.PRODUCT_KEY_END
LEFT JOIN adw_prod10.star.Dim_Geography   AS gs  ON gs.JURISDICTION_KEY     = trim(cm.JURISDICTION_KEY_START)
LEFT JOIN adw_prod10.star.Dim_Geography   AS ge  ON ge.JURISDICTION_KEY     = trim(cm.JURISDICTION_KEY_END)
LEFT JOIN adw_prod10.star.Dim_Date        AS ds  ON ds.TIME_KEY             = per.TimeKeyStart
LEFT JOIN adw_prod10.star.Dim_Date        AS de  ON de.TIME_KEY             = per.TimeKeyEnd;

-- Full reload
INSERT OVERWRITE adw_prod10.star.Fact_Certificate_Movement
    (PeriodKey, CertificateKey, MovementKey, StartStatusKey, EndStatusKey, StartProductKey, EndProductKey,
     StartGeographyKey, EndGeographyKey, StartDateKey, EndDateKey, PolicyNo, CertificateCount,
     DeathBenefitStart, DeathBenefitEnd, AnnualPaymentStart, AnnualPaymentEnd, FundValueStart, FundValueEnd,
     ReserveTotalAxisStart, ReserveTotalAxisEnd)
SELECT
     PeriodKey, CertificateKey, MovementKey, StartStatusKey, EndStatusKey, StartProductKey, EndProductKey,
     StartGeographyKey, EndGeographyKey, StartDateKey, EndDateKey, PolicyNo, CertificateCount,
     DeathBenefitStart, DeathBenefitEnd, AnnualPaymentStart, AnnualPaymentEnd, FundValueStart, FundValueEnd,
     ReserveTotalAxisStart, ReserveTotalAxisEnd
FROM adw_prod10.star.stg_Fact_Certificate_Movement;
