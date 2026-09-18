-- =============================================================================
-- USS Certificate Movement — Databricks gold layer
-- 00 · Create the gold schema in Unity Catalog
-- Assumes catalog `adw_prod10` exists and source tables live in `adw_prod10.dbo`.
-- =============================================================================

CREATE SCHEMA IF NOT EXISTS adw_prod10.star
  COMMENT 'Gold dimensional model (Kimball star) for USS Certificate Movement';
