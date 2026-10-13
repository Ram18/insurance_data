# BI Migration Estimator v2

Quotable fixed-price estimator for BI report and semantic-model migrations. It replaces the
"19-Mar-2026 BI Estimate – Calculator" workbook and follows the approved requirements document.

| File | Purpose |
| --- | --- |
| `BI_Migration_Estimator_v2.xlsx` | The workbook (Excel 365). Every figure is a live formula. |
| `build_bi_estimator.py` | Generator that builds the workbook from code, so every formula can be reviewed and rebuilt. |

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
