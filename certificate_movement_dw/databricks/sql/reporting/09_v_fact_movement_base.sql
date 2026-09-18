-- =============================================================================
-- v_fact_movement_base
-- Fact + dimensions joined, with the reporting-layer 5-period scope applied
-- (the 5 most recent QoQ/YoY periods -> equivalent of the Power Query
-- Table.LastN(QoQ|YoY, 5) filter). Feeds the metric view.
-- Ordering assumption: most recent = highest period end TIME_KEY.
-- =============================================================================

CREATE OR REPLACE VIEW adw_prod10.star.v_fact_movement_base AS
WITH scoped_periods AS (
    SELECT
        PeriodKey,
        ROW_NUMBER() OVER (ORDER BY TimeKeyEnd DESC, PERIOD_KEY DESC) AS rn
    FROM adw_prod10.star.Dim_Period
    WHERE QoqPeriods = 'QoQ' OR YoyPeriods = 'YoY'
)
SELECT
    -- period / date
    per.PeriodKey,
    per.PeriodCode,
    per.PeriodLength,
    per.QoqPeriods,
    per.YoyPeriods,
    de.MonthYearName        AS EndMonthYearName,
    de.QuarterYearName      AS EndQuarterYearName,
    de.YearNumber           AS EndYearNumber,
    -- movement classification
    dm.Movement,
    dm.MovementGroup,
    dm.StartFlag,
    dm.EndFlag,
    dm.CorrectionFlag,
    -- role-playing status
    ss.CertificateStatus    AS StartCertificateStatus,
    se.CertificateStatus    AS EndCertificateStatus,
    ss.ClassCode            AS StartClassCode,
    se.ClassCode            AS EndClassCode,
    -- role-playing geography
    gs.StateCode            AS StartStateCode,
    ge.StateCode            AS EndStateCode,
    ge.RegionName           AS EndRegionName,
    ge.CountryDescription   AS EndCountryDescription,
    -- role-playing product
    ps.BusinessLine         AS StartBusinessLine,
    pe.BusinessLine         AS EndBusinessLine,
    ps.ProductSeries        AS StartProductSeries,
    pe.ProductSeries        AS EndProductSeries,
    pe.ReportingEntity      AS EndReportingEntity,
    pe.ProductName          AS EndProductName,
    -- certificate
    dc.PolicyNo,
    dc.PlanCode,
    dc.AgentName,
    dc.Gender,
    -- measures (raw amounts)
    f.CertificateCount,
    f.DeathBenefitStart,
    f.DeathBenefitEnd,
    f.AnnualPaymentStart,
    f.AnnualPaymentEnd,
    f.FundValueStart,
    f.FundValueEnd,
    f.ReserveTotalAxisStart,
    f.ReserveTotalAxisEnd
FROM adw_prod10.star.Fact_Certificate_Movement AS f
JOIN scoped_periods                       AS sp  ON sp.PeriodKey     = f.PeriodKey AND sp.rn <= 5
JOIN adw_prod10.star.Dim_Period           AS per ON per.PeriodKey    = f.PeriodKey
JOIN adw_prod10.star.Dim_Movement         AS dm  ON dm.MovementKey   = f.MovementKey
JOIN adw_prod10.star.Dim_Status           AS ss  ON ss.StatusKey     = f.StartStatusKey
JOIN adw_prod10.star.Dim_Status           AS se  ON se.StatusKey     = f.EndStatusKey
JOIN adw_prod10.star.Dim_Geography        AS gs  ON gs.GeographyKey  = f.StartGeographyKey
JOIN adw_prod10.star.Dim_Geography        AS ge  ON ge.GeographyKey  = f.EndGeographyKey
JOIN adw_prod10.star.Dim_Product          AS ps  ON ps.ProductKey    = f.StartProductKey
JOIN adw_prod10.star.Dim_Product          AS pe  ON pe.ProductKey    = f.EndProductKey
JOIN adw_prod10.star.Dim_Certificate      AS dc  ON dc.CertificateKey = f.CertificateKey
JOIN adw_prod10.star.Dim_Date             AS de  ON de.DateKey       = f.EndDateKey;
