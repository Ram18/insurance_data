-- =============================================================================
-- certificate_movement_metrics — Unity Catalog METRIC VIEW (semantic model)
-- Replaces the Power BI dataset + DAX. Dimensions = slicing attributes,
-- measures = the movement metrics (Count / Death Benefit / Annual Payment /
-- Fund Value). Query with MEASURE(), e.g.
--   SELECT `Movement Group`, MEASURE(`End Count`), MEASURE(`End DB`)
--   FROM adw_prod10.star.certificate_movement_metrics
--   GROUP BY `Movement Group`;
-- Scope (5 most recent QoQ/YoY periods) comes from the source reporting view.
-- =============================================================================

CREATE OR REPLACE VIEW adw_prod10.star.certificate_movement_metrics
WITH METRICS
LANGUAGE YAML
AS $$
version: 0.1
source: adw_prod10.star.v_fact_movement_base

dimensions:
  - name: Period Code
    expr: PeriodCode
  - name: End Month-Year
    expr: EndMonthYearName
  - name: End Quarter-Year
    expr: EndQuarterYearName
  - name: End Year
    expr: EndYearNumber
  - name: Movement
    expr: Movement
  - name: Movement Group
    expr: MovementGroup
  - name: Start Certificate Status
    expr: StartCertificateStatus
  - name: End Certificate Status
    expr: EndCertificateStatus
  - name: Start Class Code
    expr: StartClassCode
  - name: End Class Code
    expr: EndClassCode
  - name: Start State
    expr: StartStateCode
  - name: End State
    expr: EndStateCode
  - name: End Region
    expr: EndRegionName
  - name: End Country
    expr: EndCountryDescription
  - name: Start Business Line
    expr: StartBusinessLine
  - name: End Business Line
    expr: EndBusinessLine
  - name: Start Product Series
    expr: StartProductSeries
  - name: End Product Series
    expr: EndProductSeries
  - name: End Reporting Entity
    expr: EndReportingEntity
  - name: End Product Name
    expr: EndProductName
  - name: Plan Code
    expr: PlanCode
  - name: Agent Name
    expr: AgentName
  - name: Gender
    expr: Gender

measures:
  # ------------------------------ COUNT ------------------------------
  - name: Start Count
    expr: "SUM(StartFlag)"
  - name: End Count
    expr: "SUM(EndFlag)"
  - name: New Count
    expr: "SUM(CASE WHEN MovementGroup = 'New' THEN 1 ELSE 0 END)"
  - name: Revived Count
    expr: "SUM(CASE WHEN MovementGroup = 'Revived' THEN 1 ELSE 0 END)"
  - name: Lost Count
    expr: "SUM(CASE WHEN MovementGroup = 'Lost' THEN 1 ELSE 0 END)"
  - name: Correction Count
    expr: "SUM(CASE WHEN MovementGroup = 'Correction' THEN CorrectionFlag ELSE 0 END)"
  - name: Death Count
    expr: "SUM(CASE WHEN Movement = 'Death' THEN 1 ELSE 0 END)"
  - name: Surrendered Count
    expr: "SUM(CASE WHEN Movement = 'Surrendered' THEN 1 ELSE 0 END)"
  - name: Matured Count
    expr: "SUM(CASE WHEN Movement = 'Matured' THEN 1 ELSE 0 END)"
  - name: Expired Count
    expr: "SUM(CASE WHEN Movement = 'Expired' THEN 1 ELSE 0 END)"
  - name: Lapsed Count
    expr: "SUM(CASE WHEN Movement = 'Lapsed' THEN 1 ELSE 0 END)"
  - name: Cancelled Count
    expr: "SUM(CASE WHEN Movement = 'Cancelled' THEN 1 ELSE 0 END)"
  - name: Rescinded Count
    expr: "SUM(CASE WHEN Movement = 'Rescinded' THEN 1 ELSE 0 END)"
  - name: Internal Replacement Count
    expr: "SUM(CASE WHEN Movement = 'Internal Replacement' THEN 1 ELSE 0 END)"
  - name: Miscellaneous Count
    expr: "SUM(CASE WHEN Movement = 'Miscellaneous' THEN 1 ELSE 0 END)"
  - name: Relocate State Increase Count
    expr: "SUM(CASE WHEN StartClassCode = 'Active' AND EndClassCode = 'Active' AND StartStateCode <> EndStateCode THEN 1 ELSE 0 END)"
  - name: Relocate State Decrease Count
    expr: "SUM(CASE WHEN StartClassCode = 'Active' AND EndClassCode = 'Active' AND StartStateCode <> EndStateCode THEN 1 ELSE 0 END)"
  - name: Check (Count)
    expr: "SUM( StartFlag + (CASE WHEN MovementGroup='New' THEN 1 ELSE 0 END) + (CASE WHEN MovementGroup='Revived' THEN 1 ELSE 0 END) + (CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND StartStateCode<>EndStateCode THEN 1 ELSE 0 END) + (CASE WHEN MovementGroup='Correction' THEN CorrectionFlag ELSE 0 END) - (CASE WHEN Movement='Death' THEN 1 ELSE 0 END) - (CASE WHEN Movement='Surrendered' THEN 1 ELSE 0 END) - (CASE WHEN Movement='Matured' THEN 1 ELSE 0 END) - (CASE WHEN MovementGroup='Lost' THEN 1 ELSE 0 END) - (CASE WHEN Movement='Expired' THEN 1 ELSE 0 END) - (CASE WHEN Movement='Lapsed' THEN 1 ELSE 0 END) - (CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND StartStateCode<>EndStateCode THEN 1 ELSE 0 END) - (CASE WHEN Movement='Cancelled' THEN 1 ELSE 0 END) - (CASE WHEN Movement='Rescinded' THEN 1 ELSE 0 END) - (CASE WHEN Movement='Internal Replacement' THEN 1 ELSE 0 END) - (CASE WHEN Movement='Miscellaneous' THEN 1 ELSE 0 END) - EndFlag )"

  # --------------------------- DEATH BENEFIT -------------------------
  - name: Start DB
    expr: "SUM(CASE WHEN StartFlag = 1 THEN DeathBenefitStart ELSE 0 END)"
  - name: End DB
    expr: "SUM(CASE WHEN EndFlag = 1 THEN DeathBenefitEnd ELSE 0 END)"
  - name: New DB
    expr: "SUM(CASE WHEN MovementGroup = 'New' THEN DeathBenefitEnd ELSE 0 END)"
  - name: Revived DB
    expr: "SUM(CASE WHEN MovementGroup = 'Revived' THEN DeathBenefitEnd ELSE 0 END)"
  - name: Correction DB
    expr: "SUM(CASE WHEN MovementGroup = 'Correction' THEN EndFlag*DeathBenefitEnd - StartFlag*DeathBenefitStart ELSE 0 END)"
  - name: Lost DB
    expr: "SUM(CASE WHEN MovementGroup = 'Lost' THEN DeathBenefitStart ELSE 0 END)"
  - name: Death DB
    expr: "SUM(CASE WHEN Movement = 'Death' THEN DeathBenefitStart ELSE 0 END)"
  - name: Surrendered DB
    expr: "SUM(CASE WHEN Movement = 'Surrendered' THEN DeathBenefitStart ELSE 0 END)"
  - name: Expired DB
    expr: "SUM(CASE WHEN Movement = 'Expired' THEN DeathBenefitStart ELSE 0 END)"
  - name: Lapsed DB
    expr: "SUM(CASE WHEN Movement = 'Lapsed' THEN DeathBenefitStart ELSE 0 END)"
  - name: Matured DB
    expr: "SUM(CASE WHEN Movement = 'Matured' THEN DeathBenefitStart ELSE 0 END)"
  - name: Rescinded DB
    expr: "SUM(CASE WHEN Movement = 'Rescinded' THEN DeathBenefitStart ELSE 0 END)"
  - name: Cancelled DB
    expr: "SUM(CASE WHEN Movement = 'Cancelled' THEN DeathBenefitStart ELSE 0 END)"
  - name: Internal Replacement DB
    expr: "SUM(CASE WHEN Movement = 'Internal Replacement' THEN DeathBenefitStart ELSE 0 END)"
  - name: Active Increase DB
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND EndBusinessLine='Life' AND DeathBenefitEnd > DeathBenefitStart THEN DeathBenefitEnd - DeathBenefitStart ELSE 0 END)"
  - name: Active Decrease DB
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND EndBusinessLine='Life' AND DeathBenefitEnd < DeathBenefitStart THEN DeathBenefitStart - DeathBenefitEnd ELSE 0 END)"
  - name: Relocate State Increase DB
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND StartStateCode<>EndStateCode THEN DeathBenefitEnd ELSE 0 END)"
  - name: Relocate State Decrease DB
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND StartStateCode<>EndStateCode THEN DeathBenefitEnd ELSE 0 END)"
  - name: Net Increase DB
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' THEN DeathBenefitEnd - DeathBenefitStart ELSE 0 END)"

  # --------------------------- ANNUAL PAYMENT ------------------------
  - name: Start AnnPayt
    expr: "SUM(CASE WHEN StartFlag = 1 THEN AnnualPaymentStart ELSE 0 END)"
  - name: End AnnPayt
    expr: "SUM(CASE WHEN EndFlag = 1 THEN AnnualPaymentEnd ELSE 0 END)"
  - name: New AnnPayt
    expr: "SUM(CASE WHEN MovementGroup = 'New' THEN AnnualPaymentEnd ELSE 0 END)"
  - name: Revived AnnPayt
    expr: "SUM(CASE WHEN MovementGroup = 'Revived' THEN AnnualPaymentEnd ELSE 0 END)"
  - name: Correction AnnPayt
    expr: "SUM(CASE WHEN MovementGroup = 'Correction' THEN EndFlag*AnnualPaymentEnd - StartFlag*AnnualPaymentStart ELSE 0 END)"
  - name: Lost AnnPayt
    expr: "SUM(CASE WHEN MovementGroup = 'Lost' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Death AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Death' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Surrendered AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Surrendered' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Expired AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Expired' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Lapsed AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Lapsed' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Matured AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Matured' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Rescinded AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Rescinded' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Cancelled AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Cancelled' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Internal Replacement AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Internal Replacement' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Miscellaneous AnnPayt
    expr: "SUM(CASE WHEN Movement = 'Miscellaneous' THEN AnnualPaymentStart ELSE 0 END)"
  - name: Relocate State Increase AnnPayt
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND StartStateCode<>EndStateCode THEN AnnualPaymentEnd ELSE 0 END)"
  - name: Relocate State Decrease AnnPayt
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND StartStateCode<>EndStateCode THEN AnnualPaymentEnd ELSE 0 END)"
  - name: Net Increase AnnPayt
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' THEN AnnualPaymentEnd - AnnualPaymentStart ELSE 0 END)"

  # ----------------------------- FUND VALUE --------------------------
  - name: Start FV
    expr: "SUM(CASE WHEN StartFlag = 1 THEN FundValueStart ELSE 0 END)"
  - name: End FV
    expr: "SUM(CASE WHEN EndFlag = 1 THEN FundValueEnd ELSE 0 END)"
  - name: New FV
    expr: "SUM(CASE WHEN MovementGroup = 'New' THEN FundValueEnd ELSE 0 END)"
  - name: Revived FV
    expr: "SUM(CASE WHEN MovementGroup = 'Revived' THEN FundValueEnd ELSE 0 END)"
  - name: Correction FV
    expr: "SUM(CASE WHEN MovementGroup = 'Correction' THEN EndFlag*FundValueEnd - StartFlag*FundValueStart ELSE 0 END)"
  - name: Lost FV
    expr: "SUM(CASE WHEN MovementGroup = 'Lost' THEN FundValueStart ELSE 0 END)"
  - name: Death FV
    expr: "SUM(CASE WHEN Movement = 'Death' THEN FundValueStart ELSE 0 END)"
  - name: Surrendered FV
    expr: "SUM(CASE WHEN Movement = 'Surrendered' THEN FundValueStart ELSE 0 END)"
  - name: Expired FV
    expr: "SUM(CASE WHEN Movement = 'Expired' THEN FundValueStart ELSE 0 END)"
  - name: Lapsed FV
    expr: "SUM(CASE WHEN Movement = 'Lapsed' THEN FundValueStart ELSE 0 END)"
  - name: Matured FV
    expr: "SUM(CASE WHEN Movement = 'Matured' THEN FundValueStart ELSE 0 END)"
  - name: Rescinded FV
    expr: "SUM(CASE WHEN Movement = 'Rescinded' THEN FundValueStart ELSE 0 END)"
  - name: Cancelled FV
    expr: "SUM(CASE WHEN Movement = 'Cancelled' THEN FundValueStart ELSE 0 END)"
  - name: Internal Replacement FV
    expr: "SUM(CASE WHEN Movement = 'Internal Replacement' THEN FundValueStart ELSE 0 END)"
  - name: Relocate State Increase FV
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND StartStateCode<>EndStateCode THEN FundValueEnd ELSE 0 END)"
  - name: Relocate State Decrease FV
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' AND StartStateCode<>EndStateCode THEN FundValueEnd ELSE 0 END)"
  - name: Net Increase FV
    expr: "SUM(CASE WHEN StartClassCode='Active' AND EndClassCode='Active' THEN FundValueEnd - FundValueStart ELSE 0 END)"
$$;
