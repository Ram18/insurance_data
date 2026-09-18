# USS Certificate Movement — Dimensional Model (Kimball)

A ground-up rebuild of the **USS Certificate Movement** Power BI model as a proper
Kimball star schema, built **in the SQL warehouse first** (`ADW_PROD10`) so the
BI layer sits cleanly on top.

> **Migration target: Databricks.** The report is being migrated off Power BI to
> Databricks. The `databricks/` folder holds the Unity Catalog build — Delta gold
> star, a **metric view** (semantic model + all measures, replacing the dataset +
> DAX), and an **AI/BI dashboard**. The `sql/` folder below is the equivalent
> SQL Server implementation of the same star (kept for reference / parallel run).

**Design decisions (agreed):**

- **Fact type:** periodic **snapshot** — one row per certificate per movement period.
- **SCD:** **Type 1** on every dimension (latest value; reproduces today's numbers).
- **Delivery:** SQL warehouse star schema first (surrogate keys, conformed dimensions),
  then the Power BI model.
- **Goal:** identical results for business users vs. the current `RS|USS Certificate Movement`
  table, with a maintainable, query-foldable, self-documenting model.

> All objects are created in a new **`star`** schema in `ADW_PROD10`, sourced from `dbo`.
> The existing `dbo` tables are left untouched.

---

## 1. What the source model does

The current model compares a book of life-insurance certificates at the **start** and
**end** of a period and classifies each certificate's change (New, Death, Lapsed,
Surrendered, Matured, Correction, Relocate State, …) across four measure families:
**Certificate Count**, **Death Benefit (DB)**, **Annual Payment (AnnPayt)** and
**Fund Value (FV)**.

The classification is data-driven: `dbo.MOVEMENT_MEASURE_LOOKUP` maps every
`(start status, end status)` pair to a `MOVEMENT`, a `MOVEMENT_GROUP` and a set of
flags (`START_FLAG`, `END_FLAG`, `NEW_FLAG`, `REVIVED_FLAG`, `EXIT_FLAG`, `LOST_FLAG`,
`CORRECTION_FLAG`). This is the heart of the model and becomes a **transition dimension**.

---

## 2. Grain

> **One row per certificate, per movement period** — the certificate's start-of-period
> and end-of-period position plus its movement classification for that period.

This is the grain of `dbo.CERTIFICATE_MOVEMENT` and is preserved exactly, so totals
reconcile 1:1 with the current model.

---

## 3. Bus matrix

| Business process | Date (Start/End) | Period | Certificate | Status (Start/End) | Product (Start/End) | Geography (Start/End) | Movement |
| --- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| Certificate movement snapshot | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |

`Status`, `Product`, `Geography` and `Date` are **role-playing** (a Start role and an
End role each), modelled as one physical dimension referenced by two surrogate FKs.

---

## 4. Star schema

```
                 Dim_Date (Start / End)
                        |
 Dim_Period ---- Fact_Certificate_Movement ---- Dim_Certificate
                   |    |    |    |
     Dim_Status ---+    |    |    +--- Dim_Geography  (Start / End roles)
   (Start / End)        |    |
        Dim_Product ----+    +---- Dim_Movement
        (Start / End)
```

---

## 5. Dimension specifications (all Type 1)

Each dimension gets an **integer surrogate key** (`IDENTITY`), retains its **business
(natural) key** for lookups and ETL, and carries a `-1` **Unknown** member so the fact
never has a null FK. Descriptive attributes are trimmed at load.

| Dimension | Source | Surrogate key | Business key | Role-playing | Key attributes kept |
| --- | --- | --- | --- | --- | --- |
| `Dim_Date` | `dbo.TIME_LOOKUP` | `DateKey` | `TIME_KEY` | Start / End | AS_OF_DATE, YEAR/QUARTER/MONTH numbers & names, MONTH_YEAR_NAME, QUARTER_YEAR_NAME, CURRENT_DATE_IND |
| `Dim_Status` | `dbo.STATUS_LOOKUP` | `StatusKey` | `ADW_STATUS_KEY` | Start / End | CERTIFICATE_STATUS, CLASS_CODE, PREMIUM_STATUS, STATUS_DESCRIPTION, ADMIN_CERTIFICATE_STATUS(_GROUP), DISABILITY_STATUS |
| `Dim_Product` | `dbo.PRODUCT_LOOKUP` | `ProductKey` | `PRODUCT_KEY` | Start / End | PRODUCT_CODE/NAME/TYPE, BUSINESS_LINE, BUSINESS_SUB_LINE, PRODUCT_SERIES, SERIES, REPORTING_ENTITY, REPORTING_BUSINESS_GROUP, ISSUING_COMPANY |
| `Dim_Geography` | `dbo.GEOGRAPHIC_LOOKUP` | `GeographyKey` | `JURISDICTION_KEY` | Start / End | JURISDICTION_CODE/DESCRIPTION, STATE_CODE/NAME, REGION_CODE/NAME, COUNTRY_CODE/DESCRIPTION, CONTINENT_CODE |
| `Dim_Movement` | `dbo.MOVEMENT_MEASURE_LOOKUP` | `MovementKey` | `(ADW_STATUS_KEY_START, ADW_STATUS_KEY_END)` | — | MOVEMENT, MOVEMENT_GROUP, START_FLAG, END_FLAG, NEW_FLAG, REVIVED_FLAG, EXIT_FLAG, LOST_FLAG, CORRECTION_FLAG, PREMIUM_PAYING_AT_START |
| `Dim_Period` | `dbo.MOVEMENT_PERIODS` | `PeriodKey` | `PERIOD_KEY` | — | PERIOD_CODE, PERIOD_START/END_CODE, YOY/QOQ/YTD/MOM flags, PERIOD_LENGTH, CURRENT_REPORT_IND, CURRENT_ANALYSIS_IND |
| `Dim_Certificate` | `dbo.CERTIFICATE` | `CertificateKey` | `ADW_CERTIFICATE_KEY` | — | CERTIFICATE_NUMBER, POLICY_NO, PLAN_CODE, ISSUE_DATE, ISSUE_AGE, GENDER, SMOKER_STATUS, PRODUCT_TYPE_CODE, DEATH_STATUS, CURRENT_STATUS, AGENT_NAME/NUMBER, SOURCE_SYSTEM (curated subset of 109) |

**Why `Dim_Movement` is a dimension, not fact columns:** the flags and labels are pure
attributes of a `(start status → end status)` transition. Modelling them once, keyed by
the pair, removes ~90 conditional columns from the pipeline and lets the movement metrics
become clean measures.

---

## 6. Fact specification — `Fact_Certificate_Movement`

**Grain:** certificate × period (periodic snapshot).

**Foreign keys (surrogate):**

| FK column | → Dimension | Role | Source join key |
| --- | --- | --- | --- |
| `PeriodKey` | `Dim_Period` | — | `PERIOD_KEY` |
| `CertificateKey` | `Dim_Certificate` | — | `ADW_CERTIFICATE_KEY` |
| `MovementKey` | `Dim_Movement` | — | `(ADW_STATUS_KEY_START, ADW_STATUS_KEY_END)` |
| `StartStatusKey` / `EndStatusKey` | `Dim_Status` | Start / End | `ADW_STATUS_KEY_START` / `_END` |
| `StartProductKey` / `EndProductKey` | `Dim_Product` | Start / End | `PRODUCT_KEY_START` / `_END` |
| `StartGeographyKey` / `EndGeographyKey` | `Dim_Geography` | Start / End | `JURISDICTION_KEY_START` / `_END` |
| `StartDateKey` / `EndDateKey` | `Dim_Date` | Start / End | via `Dim_Period.TIME_KEY_START/END` |

**Degenerate dimension:** `POLICY_NO` (carried on the fact for drill-through).

**Measures (additive):** the base amounts the movement metrics are built from —

| Measure | Source column |
| --- | --- |
| `DeathBenefitStart` / `DeathBenefitEnd` | `BENEFIT_DEATH_TOTAL_START` / `_END` |
| `AnnualPaymentStart` / `AnnualPaymentEnd` | `BENEFIT_INCOME_ANNUAL_START` / `BENEFIT_INCOME_ANNUALISED_END` |
| `FundValueStart` / `FundValueEnd` | `FUND_VALUE_START` / `_END` |
| `ReserveTotalAxisStart` / `ReserveTotalAxisEnd` | `USS_RESERVE_TOTAL_AXIS_START` / `_END` |
| `CertificateCount` | constant `1` (each row is one certificate) |

> Additional reserve/benefit measures from the 114-column source fact can be added later
> the same way; this set reproduces every metric the current report surfaces.

**Movement metrics are NOT stored** — they are derived from the flags in `Dim_Movement`
and the Start/End dimension attributes (e.g. `Active`/`Active` + state change → Relocate),
exactly as the current Power Query did, but as measures. See `sql/validation` for the
reconciliation views that reproduce the Count / DB / AnnPayt / FV families.

---

## 7. Naming & conventions

- Schema `star`; dimensions `Dim_*`, fact `Fact_*`, staging views `stg_*`, reconciliation `v_recon_*`.
- Surrogate keys: `IDENTITY(1,1)`, `<Entity>Key`; Unknown member = `-1`.
- All loads are **Type 1 `MERGE`** (idempotent upsert), safe to re-run.
- Descriptive text trimmed (`LTRIM(RTRIM(...))`) to match the source Power Query `Text.Trim`.

---

## 8. Build order

1. `sql/00_create_schema.sql` — the `star` schema.
2. `sql/dimensions/*` — dimension DDL + Type 1 loads (this increment).
3. `sql/fact/*` — fact DDL + load (next).
4. `sql/validation/*` — reconciliation vs. the current model (next).
5. Power BI: point the dataset at the `star` schema, add measures, retire the old queries.
