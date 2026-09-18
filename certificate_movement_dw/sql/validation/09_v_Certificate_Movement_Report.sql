/* =============================================================================
   v_Certificate_Movement_Report
   Reproduces the original ~90 movement metrics (Count / Death Benefit /
   Annual Payment / Fund Value) directly from the star, at the fact grain.
   Each column is a per-row contribution; SUM() them to get the report figures.
   This is the SQL equivalent of the current model's unpivoted Attribute/Value
   output and is the reconciliation anchor. It also mirrors the DAX measures the
   Power BI layer will use.
   ============================================================================= */

CREATE OR ALTER VIEW star.v_Certificate_Movement_Report AS
SELECT
    -- Grouping / slicing attributes (from the role-playing dimensions)
    per.PeriodCode,
    per.PeriodLength,
    ss.CertificateStatus  AS StartCertificateStatus,
    se.CertificateStatus  AS EndCertificateStatus,
    ss.ClassCode          AS StartClassCode,
    se.ClassCode          AS EndClassCode,
    gs.StateCode          AS StartStateCode,
    ge.StateCode          AS EndStateCode,
    ps.BusinessLine       AS StartBusinessLine,
    pe.BusinessLine       AS EndBusinessLine,
    ps.ProductSeries      AS StartProductSeries,
    pe.ProductSeries      AS EndProductSeries,
    dm.Movement,
    dm.MovementGroup,

    -- ============================ COUNT ==================================
    CAST(dm.StartFlag AS INT)                                             AS StartCount,
    CAST(dm.EndFlag   AS INT)                                             AS EndCount,
    IIF(dm.MovementGroup = 'New',     1, 0)                               AS NewCount,
    IIF(dm.MovementGroup = 'Revived', 1, 0)                               AS RevivedCount,
    IIF(dm.MovementGroup = 'Lost',    1, 0)                               AS LostCount,
    IIF(dm.MovementGroup = 'Correction', dm.CorrectionFlag, 0)            AS CorrectionCount,
    IIF(dm.Movement = 'Death',       1, 0)                                AS DeathCount,
    IIF(dm.Movement = 'Surrendered', 1, 0)                                AS SurrenderedCount,
    IIF(dm.Movement = 'Matured',     1, 0)                                AS MaturedCount,
    IIF(dm.Movement = 'Expired',     1, 0)                                AS ExpiredCount,
    IIF(dm.Movement = 'Lapsed',      1, 0)                                AS LapsedCount,
    IIF(dm.Movement = 'Cancelled',   1, 0)                                AS CancelledCount,
    IIF(dm.Movement = 'Rescinded',   1, 0)                                AS RescindedCount,
    IIF(dm.Movement = 'Internal Replacement', 1, 0)                       AS InternalReplacementCount,
    IIF(dm.Movement = 'Miscellaneous', 1, 0)                              AS MiscellaneousCount,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND gs.StateCode <> ge.StateCode, 1, 0)
                                                                          AS RelocateStateIncreaseCount,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND gs.StateCode <> ge.StateCode, 1, 0)
                                                                          AS RelocateStateDecreaseCount,

    -- ========================= DEATH BENEFIT =============================
    IIF(dm.StartFlag = 1, f.DeathBenefitStart, 0)                         AS StartDB,
    IIF(dm.EndFlag   = 1, f.DeathBenefitEnd,   0)                         AS EndDB,
    IIF(dm.MovementGroup = 'New',     f.DeathBenefitEnd, 0)               AS NewDB,
    IIF(dm.MovementGroup = 'Revived', f.DeathBenefitEnd, 0)               AS RevivedDB,
    IIF(dm.MovementGroup = 'Correction', dm.EndFlag*f.DeathBenefitEnd - dm.StartFlag*f.DeathBenefitStart, 0)
                                                                          AS CorrectionDB,
    IIF(dm.MovementGroup = 'Lost',    f.DeathBenefitStart, 0)             AS LostDB,
    IIF(dm.Movement = 'Death',       f.DeathBenefitStart, 0)              AS DeathDB,
    IIF(dm.Movement = 'Surrendered', f.DeathBenefitStart, 0)              AS SurrenderedDB,
    IIF(dm.Movement = 'Expired',     f.DeathBenefitStart, 0)              AS ExpiredDB,
    IIF(dm.Movement = 'Lapsed',      f.DeathBenefitStart, 0)              AS LapsedDB,
    IIF(dm.Movement = 'Matured',     f.DeathBenefitStart, 0)              AS MaturedDB,
    IIF(dm.Movement = 'Rescinded',   f.DeathBenefitStart, 0)             AS RescindedDB,
    IIF(dm.Movement = 'Cancelled',   f.DeathBenefitStart, 0)              AS CancelledDB,
    IIF(dm.Movement = 'Internal Replacement', f.DeathBenefitStart, 0)     AS InternalReplacementDB,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND pe.BusinessLine = 'Life' AND f.DeathBenefitEnd > f.DeathBenefitStart,
        f.DeathBenefitEnd - f.DeathBenefitStart, 0)                       AS ActiveIncreaseDB,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND pe.BusinessLine = 'Life' AND f.DeathBenefitEnd < f.DeathBenefitStart,
        f.DeathBenefitStart - f.DeathBenefitEnd, 0)                       AS ActiveDecreaseDB,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND gs.StateCode <> ge.StateCode, f.DeathBenefitEnd, 0)
                                                                          AS RelocateStateIncreaseDB,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND gs.StateCode <> ge.StateCode, f.DeathBenefitEnd, 0)
                                                                          AS RelocateStateDecreaseDB,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active', f.DeathBenefitEnd - f.DeathBenefitStart, 0)
                                                                          AS NetIncreaseDB,

    -- ========================= ANNUAL PAYMENT ============================
    IIF(dm.StartFlag = 1, f.AnnualPaymentStart, 0)                       AS StartAnnPayt,
    IIF(dm.EndFlag   = 1, f.AnnualPaymentEnd,   0)                       AS EndAnnPayt,
    IIF(dm.MovementGroup = 'New',     f.AnnualPaymentEnd, 0)             AS NewAnnPayt,
    IIF(dm.MovementGroup = 'Revived', f.AnnualPaymentEnd, 0)             AS RevivedAnnPayt,
    IIF(dm.MovementGroup = 'Correction', dm.EndFlag*f.AnnualPaymentEnd - dm.StartFlag*f.AnnualPaymentStart, 0)
                                                                          AS CorrectionAnnPayt,
    IIF(dm.MovementGroup = 'Lost',    f.AnnualPaymentStart, 0)           AS LostAnnPayt,
    IIF(dm.Movement = 'Death',       f.AnnualPaymentStart, 0)            AS DeathAnnPayt,
    IIF(dm.Movement = 'Surrendered', f.AnnualPaymentStart, 0)            AS SurrenderedAnnPayt,
    IIF(dm.Movement = 'Expired',     f.AnnualPaymentStart, 0)            AS ExpiredAnnPayt,
    IIF(dm.Movement = 'Lapsed',      f.AnnualPaymentStart, 0)            AS LapsedAnnPayt,
    IIF(dm.Movement = 'Matured',     f.AnnualPaymentStart, 0)            AS MaturedAnnPayt,
    IIF(dm.Movement = 'Rescinded',   f.AnnualPaymentStart, 0)           AS RescindedAnnPayt,
    IIF(dm.Movement = 'Cancelled',   f.AnnualPaymentStart, 0)            AS CancelledAnnPayt,
    IIF(dm.Movement = 'Internal Replacement', f.AnnualPaymentStart, 0)   AS InternalReplacementAnnPayt,
    IIF(dm.Movement = 'Miscellaneous', f.AnnualPaymentStart, 0)          AS MiscellaneousAnnPayt,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND gs.StateCode <> ge.StateCode, f.AnnualPaymentEnd, 0)
                                                                          AS RelocateStateIncreaseAnnPayt,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND gs.StateCode <> ge.StateCode, f.AnnualPaymentEnd, 0)
                                                                          AS RelocateStateDecreaseAnnPayt,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active', f.AnnualPaymentEnd - f.AnnualPaymentStart, 0)
                                                                          AS NetIncreaseAnnPayt,

    -- =========================== FUND VALUE ==============================
    IIF(dm.StartFlag = 1, f.FundValueStart, 0)                          AS StartFV,
    IIF(dm.EndFlag   = 1, f.FundValueEnd,   0)                          AS EndFV,
    IIF(dm.MovementGroup = 'New',     f.FundValueEnd, 0)                AS NewFV,
    IIF(dm.MovementGroup = 'Revived', f.FundValueEnd, 0)                AS RevivedFV,
    IIF(dm.MovementGroup = 'Correction', dm.EndFlag*f.FundValueEnd - dm.StartFlag*f.FundValueStart, 0)
                                                                          AS CorrectionFV,
    IIF(dm.MovementGroup = 'Lost',    f.FundValueStart, 0)              AS LostFV,
    IIF(dm.Movement = 'Death',       f.FundValueStart, 0)               AS DeathFV,
    IIF(dm.Movement = 'Surrendered', f.FundValueStart, 0)               AS SurrenderedFV,
    IIF(dm.Movement = 'Expired',     f.FundValueStart, 0)               AS ExpiredFV,
    IIF(dm.Movement = 'Lapsed',      f.FundValueStart, 0)               AS LapsedFV,
    IIF(dm.Movement = 'Matured',     f.FundValueStart, 0)               AS MaturedFV,
    IIF(dm.Movement = 'Rescinded',   f.FundValueStart, 0)              AS RescindedFV,
    IIF(dm.Movement = 'Cancelled',   f.FundValueStart, 0)               AS CancelledFV,
    IIF(dm.Movement = 'Internal Replacement', f.FundValueStart, 0)      AS InternalReplacementFV,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND gs.StateCode <> ge.StateCode, f.FundValueEnd, 0)
                                                                          AS RelocateStateIncreaseFV,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active' AND gs.StateCode <> ge.StateCode, f.FundValueEnd, 0)
                                                                          AS RelocateStateDecreaseFV,
    IIF(ss.ClassCode = 'Active' AND se.ClassCode = 'Active', f.FundValueEnd - f.FundValueStart, 0)
                                                                          AS NetIncreaseFV
FROM star.Fact_Certificate_Movement AS f
JOIN star.Dim_Period      AS per ON per.PeriodKey       = f.PeriodKey
JOIN star.Dim_Movement    AS dm  ON dm.MovementKey      = f.MovementKey
JOIN star.Dim_Status      AS ss  ON ss.StatusKey        = f.StartStatusKey
JOIN star.Dim_Status      AS se  ON se.StatusKey        = f.EndStatusKey
JOIN star.Dim_Product     AS ps  ON ps.ProductKey       = f.StartProductKey
JOIN star.Dim_Product     AS pe  ON pe.ProductKey       = f.EndProductKey
JOIN star.Dim_Geography   AS gs  ON gs.GeographyKey     = f.StartGeographyKey
JOIN star.Dim_Geography   AS ge  ON ge.GeographyKey     = f.EndGeographyKey;
GO
