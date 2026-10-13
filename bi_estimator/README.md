# BI Migration Estimator v2

Quotable fixed-price estimator for BI report and semantic-model migrations. It replaces the
"19-Mar-2026 BI Estimate – Calculator" workbook and follows the approved requirements document.

| File | Purpose |
| --- | --- |
| `BI_Migration_Presales_Estimator.xlsx` | **Pre-sales ballpark** (3 tabs). About a dozen inputs, a price range and a sprint plan like the original PoP sheet. |
| `BI_Migration_Estimator_v2.xlsx` | **Full estimator** for the quotable fixed price (17 tabs). Every figure is a live formula. |
| `build_bi_presales.py` / `build_bi_estimator.py` | Generators that build each workbook from code, so every formula can be reviewed and rebuilt. |

## Pre-sales ballpark (simple version)

- **Estimate**: project settings (sprint length, dev team, automation %, contingency, PM / AI readiness)
  and one scope row per source → target pair (reports, pages, complexity profile, % retired, semantic
  models). Results: expected price, ballpark range (-10% / +25%), hours, duration, sprints, peak and average
  team, blended rate, phase table and checks.
- **Plan**: sprint and week grid with coloured phase bars (FTE per week) and the team by role per week.
  Team cells are pre-filled from the phases; overwrite any of them to hand-plan, and the cost follows the grid.
- **Reference**: the numbers behind it, using the same baseline as the full estimator (hours per page,
  model hours, pair factors, automation caps, rate card, role mix).

Use the pre-sales file for early conversations, then move to the full estimator for the quote.

## Using the workbook

1. **Setup**: engagement details, iteration length, team and ramp settings, SIT/UAT/deployment cadence,
   automation % by phase, cost mode.
2. **Scope**: one row per migration pair (source/target BI tool and data platform, inventory,
   complexity, rationalisation, semantic models, security, AI readiness). Two example rows are included;
   overwrite them.
3. **Risks**: named risks with impact hours and probability.
4. **Roles**: rate-card row per plan role and role mix per phase.
5. **Checks**: every check must PASS. The Summary banner turns green only then.
6. **Summary** and **Assumptions**: print for the proposal.

Yellow cells with blue text are inputs. All other cells are locked; sheets are protected **without a
password**, so the BI team can unprotect them to change baseline values.

## Rebuilding

```bash
python build_bi_estimator.py                                  # writes BI_Migration_Estimator_v2.xlsx
python build_bi_estimator.py regression.xlsx --regression     # old base case, for the regression test
python build_bi_presales.py                                   # writes BI_Migration_Presales_Estimator.xlsx
```

openpyxl writes formulas without cached values. Recalculate once (open and save in Excel, or LibreOffice
headless) before reading values programmatically.

## Verification done for this version

- Recalculated in LibreOffice: 5,927 formulas, 0 errors. All 15 checks pass on the example scope.
- **Regression test.** The `--regression` build is configured as the old base case: Tableau → QuickSight,
  5,000 single-page reports, mix 40/40/20, 60% report automation with a 100% cap, and the old
  effective learning gain of 19%. It reproduces the old calculator's report development hours exactly
  (10,558.35 h).
- The example pairs' report, semantic, SIT, UAT, deployment, hypercare and schedule figures were
  re-derived by hand and match the workbook.
