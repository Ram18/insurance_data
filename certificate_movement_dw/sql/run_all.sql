/* =============================================================================
   Master runner — execute in this order with sqlcmd (:r requires SQLCMD mode)
   against ADW_PROD10, e.g.:
       sqlcmd -S FHVMWEDWSQLP1 -d ADW_PROD10 -i run_all.sql
   Or run each file manually in the order below.
   ============================================================================= */
:r 00_create_schema.sql
:r dimensions/01_Dim_Date.sql
:r dimensions/02_Dim_Status.sql
:r dimensions/03_Dim_Product.sql
:r dimensions/04_Dim_Geography.sql
:r dimensions/05_Dim_Movement.sql
:r dimensions/06_Dim_Period.sql
:r dimensions/07_Dim_Certificate.sql
:r fact/08_Fact_Certificate_Movement.sql
:r validation/09_v_Certificate_Movement_Report.sql
:r validation/10_reconciliation.sql
