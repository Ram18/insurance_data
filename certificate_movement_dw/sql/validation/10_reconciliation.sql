/* =============================================================================
   Reconciliation — prove the star reproduces the source exactly.
   Run after the fact load. Two checks:
     (A) Internal balancing identities (the model's own "Check (…)" columns) = 0.
     (B) Star totals vs. a direct recompute from dbo.CERTIFICATE_MOVEMENT.
   ============================================================================= */

/* ---- (A) Balancing identities: every Check must be 0 ------------------------*/
;WITH r AS (SELECT * FROM star.v_Certificate_Movement_Report)
SELECT
    CheckCount = SUM(StartCount + NewCount + RevivedCount + RelocateStateIncreaseCount + CorrectionCount
                     - DeathCount - SurrenderedCount - MaturedCount - LostCount - ExpiredCount - LapsedCount
                     - RelocateStateDecreaseCount - CancelledCount - RescindedCount
                     - InternalReplacementCount - MiscellaneousCount - EndCount),
    CheckDB    = SUM(StartDB + NewDB + ActiveIncreaseDB + RevivedDB + RelocateStateIncreaseDB + CorrectionDB
                     - LostDB - CancelledDB - RescindedDB - LapsedDB - SurrenderedDB - DeathDB - MaturedDB
                     - ExpiredDB - RelocateStateDecreaseDB - ActiveDecreaseDB - InternalReplacementDB - EndDB),
    CheckAnn   = SUM(StartAnnPayt + NewAnnPayt + RevivedAnnPayt + NetIncreaseAnnPayt + RelocateStateIncreaseAnnPayt
                     + CorrectionAnnPayt - LostAnnPayt - RescindedAnnPayt - CancelledAnnPayt - LapsedAnnPayt
                     - SurrenderedAnnPayt - DeathAnnPayt - MaturedAnnPayt - ExpiredAnnPayt
                     - RelocateStateDecreaseAnnPayt - InternalReplacementAnnPayt - MiscellaneousAnnPayt - EndAnnPayt),
    CheckFV    = SUM(StartFV + NewFV + RevivedFV + NetIncreaseFV + RelocateStateIncreaseFV + CorrectionFV
                     - LostFV - RescindedFV - CancelledFV - LapsedFV - SurrenderedFV - DeathFV - MaturedFV
                     - ExpiredFV - RelocateStateDecreaseFV - InternalReplacementFV - EndFV)
FROM r;
-- Expect all four ~ 0 (allow tiny float rounding).
GO

/* ---- (B) Star vs. raw source: Start/End totals should tie exactly -----------*/
SELECT src = 'SOURCE',
       StartDB  = SUM(CASE WHEN mml.START_FLAG = 1 THEN cm.BENEFIT_DEATH_TOTAL_START END),
       EndDB    = SUM(CASE WHEN mml.END_FLAG   = 1 THEN cm.BENEFIT_DEATH_TOTAL_END   END),
       StartFV  = SUM(CASE WHEN mml.START_FLAG = 1 THEN cm.FUND_VALUE_START END),
       EndFV    = SUM(CASE WHEN mml.END_FLAG   = 1 THEN cm.FUND_VALUE_END   END),
       Rows_    = COUNT(*)
FROM dbo.CERTIFICATE_MOVEMENT cm
JOIN dbo.MOVEMENT_MEASURE_LOOKUP mml
     ON mml.ADW_STATUS_KEY_START = cm.ADW_STATUS_KEY_START
    AND mml.ADW_STATUS_KEY_END   = cm.ADW_STATUS_KEY_END
UNION ALL
SELECT src = 'STAR',
       StartDB = SUM(StartDB), EndDB = SUM(EndDB),
       StartFV = SUM(StartFV), EndFV = SUM(EndFV),
       Rows_   = COUNT(*)
FROM star.v_Certificate_Movement_Report;
-- The SOURCE and STAR rows should match (subject to the same period scope).
GO

/* ---- Row-count sanity: fact grain must equal source fact row count ----------*/
SELECT
    SourceRows = (SELECT COUNT(*) FROM dbo.CERTIFICATE_MOVEMENT),
    FactRows   = (SELECT COUNT(*) FROM star.Fact_Certificate_Movement),
    UnknownFKs = (SELECT COUNT(*) FROM star.Fact_Certificate_Movement
                  WHERE -1 IN (PeriodKey, CertificateKey, MovementKey, StartStatusKey, EndStatusKey));
-- SourceRows = FactRows; investigate any UnknownFKs (natural keys missing from a dimension).
GO
