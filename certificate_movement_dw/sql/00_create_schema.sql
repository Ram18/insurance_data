/* =============================================================================
   USS Certificate Movement — Dimensional Model
   00 · Create the `star` schema
   Target: SQL Server, database ADW_PROD10
   Idempotent: safe to re-run.
   ============================================================================= */

IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = N'star')
    EXEC (N'CREATE SCHEMA star AUTHORIZATION dbo;');
GO
