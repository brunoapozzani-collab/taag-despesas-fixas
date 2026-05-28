-- =============================================================================
-- Taag Despesas — Initial Schema
-- =============================================================================

-- Enable UUID generation
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- -----------------------------------------------------------------------------
-- 1. companies — static seed list from COMPANIES constant in expense_engine.py
-- -----------------------------------------------------------------------------
CREATE TABLE companies (
  id   uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  slug text UNIQUE NOT NULL
);

INSERT INTO companies (name, slug) VALUES
  ('Rio de Janeiro',   'rio-de-janeiro'),
  ('Alameda 470',      'alameda-470'),
  ('Artur de Azevedo', 'artur-de-azevedo'),
  ('Mazzini',          'mazzini'),
  ('Alameda 334',      'alameda-334');

-- -----------------------------------------------------------------------------
-- 2. ceo_categoria_lookup — auto-synced from Python CEO_CATEGORY_MAP on start
-- Never edited manually. Python is canonical.
-- -----------------------------------------------------------------------------
CREATE TABLE ceo_categoria_lookup (
  name       text    PRIMARY KEY,
  keywords   text[]  NOT NULL DEFAULT '{}',
  sort_order int     NOT NULL DEFAULT 0
);

-- -----------------------------------------------------------------------------
-- 3. ledger_uploads — index of every Excel upload
-- The .xlsx file lives in Storage at storage_path.
-- The DB is fully rebuildable from Storage files + presets.
-- -----------------------------------------------------------------------------
CREATE TABLE ledger_uploads (
  id           uuid        PRIMARY KEY DEFAULT gen_random_uuid(),
  uploaded_by  uuid        REFERENCES auth.users ON DELETE SET NULL,
  uploaded_at  timestamptz NOT NULL DEFAULT now(),
  period_start date,
  period_end   date,
  filename     text        NOT NULL,
  storage_path text        NOT NULL,
  row_count    int
);

-- -----------------------------------------------------------------------------
-- 4. transactions — one row per ledger transaction post-classification
-- ceo_categoria always written by Python assign_ceo_category(). Never by Next.js.
-- -----------------------------------------------------------------------------
CREATE TABLE transactions (
  id             uuid    PRIMARY KEY DEFAULT gen_random_uuid(),
  upload_id      uuid    REFERENCES ledger_uploads ON DELETE CASCADE,
  pagto          date,
  empresa        text,
  favorecido     text,
  descricao      text,
  cod_despesa    int,
  despesas       text,
  valor          numeric NOT NULL,
  row_hash       text    NOT NULL,
  is_fixed_auto  bool    NOT NULL,
  is_fixed       bool    NOT NULL,
  ceo_categoria  text    NOT NULL
);

CREATE INDEX idx_transactions_upload_id ON transactions (upload_id);
CREATE INDEX idx_transactions_empresa   ON transactions (empresa);
CREATE INDEX idx_transactions_pagto     ON transactions (pagto);
CREATE INDEX idx_transactions_ceo_cat   ON transactions (ceo_categoria);

-- -----------------------------------------------------------------------------
-- 5. presets — Supabase copy of presets.json
-- Local presets.json is the unconditional source of truth.
-- Supabase copy is write-only from Operator (no timestamp comparison).
-- -----------------------------------------------------------------------------
CREATE TABLE presets (
  id                 uuid        PRIMARY KEY DEFAULT gen_random_uuid(),
  operator_id        uuid        UNIQUE NOT NULL REFERENCES auth.users ON DELETE CASCADE,
  fixed_keywords     text[]      NOT NULL DEFAULT '{}',
  fixed_codes        int[]       NOT NULL DEFAULT '{}',
  manual_overrides   jsonb       NOT NULL DEFAULT '{}'::jsonb,
  vendor_company_map jsonb       NOT NULL DEFAULT '{}'::jsonb,
  updated_at         timestamptz NOT NULL DEFAULT now()
);

-- -----------------------------------------------------------------------------
-- 6. monthly_rollups — pre-aggregated CEO reads
-- Written by Operator after confirm. Rebuildable from transactions.
-- -----------------------------------------------------------------------------
CREATE TABLE monthly_rollups (
  id            uuid        PRIMARY KEY DEFAULT gen_random_uuid(),
  company_id    uuid        REFERENCES companies ON DELETE CASCADE,
  year          int         NOT NULL,
  month         int         NOT NULL CHECK (month BETWEEN 1 AND 12),
  ceo_categoria text        NOT NULL,
  total         numeric     NOT NULL DEFAULT 0,
  computed_at   timestamptz NOT NULL DEFAULT now(),
  UNIQUE (company_id, year, month, ceo_categoria)
);

CREATE INDEX idx_rollups_company_ym ON monthly_rollups (company_id, year, month);
CREATE INDEX idx_rollups_ym         ON monthly_rollups (year, month);

-- =============================================================================
-- STORAGE BUCKET
-- =============================================================================
INSERT INTO storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
VALUES (
  'uploads',
  'uploads',
  false,
  52428800,
  ARRAY[
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel'
  ]
);

-- =============================================================================
-- ROW LEVEL SECURITY
-- =============================================================================
ALTER TABLE companies            ENABLE ROW LEVEL SECURITY;
ALTER TABLE ceo_categoria_lookup ENABLE ROW LEVEL SECURITY;
ALTER TABLE ledger_uploads       ENABLE ROW LEVEL SECURITY;
ALTER TABLE transactions         ENABLE ROW LEVEL SECURITY;
ALTER TABLE presets              ENABLE ROW LEVEL SECURITY;
ALTER TABLE monthly_rollups      ENABLE ROW LEVEL SECURITY;

-- Role helpers (JWT app_metadata.app_role claim) — in public schema
CREATE OR REPLACE FUNCTION public.is_operator()
RETURNS bool LANGUAGE sql STABLE SECURITY DEFINER AS $$
  SELECT coalesce(
    (auth.jwt() -> 'app_metadata' ->> 'app_role') = 'operator', false
  )
$$;

CREATE OR REPLACE FUNCTION public.is_ceo()
RETURNS bool LANGUAGE sql STABLE SECURITY DEFINER AS $$
  SELECT coalesce(
    (auth.jwt() -> 'app_metadata' ->> 'app_role') = 'ceo', false
  )
$$;

-- companies: all authenticated can read; seeded via migration only
CREATE POLICY "companies_read"
  ON companies FOR SELECT TO authenticated USING (true);

-- ceo_categoria_lookup: all read; operators can upsert (for _sync_ceo_categories)
CREATE POLICY "ceo_cat_read"
  ON ceo_categoria_lookup FOR SELECT TO authenticated USING (true);
CREATE POLICY "ceo_cat_insert"
  ON ceo_categoria_lookup FOR INSERT TO authenticated
  WITH CHECK (public.is_operator());
CREATE POLICY "ceo_cat_update"
  ON ceo_categoria_lookup FOR UPDATE TO authenticated
  USING (public.is_operator());

-- ledger_uploads: operators only
CREATE POLICY "uploads_insert"
  ON ledger_uploads FOR INSERT TO authenticated
  WITH CHECK (public.is_operator() AND uploaded_by = auth.uid());
CREATE POLICY "uploads_read"
  ON ledger_uploads FOR SELECT TO authenticated
  USING (public.is_operator());

-- transactions: operators write; operators + CEO read
CREATE POLICY "txn_insert"
  ON transactions FOR INSERT TO authenticated
  WITH CHECK (public.is_operator());
CREATE POLICY "txn_update"
  ON transactions FOR UPDATE TO authenticated
  USING (public.is_operator());
CREATE POLICY "txn_delete"
  ON transactions FOR DELETE TO authenticated
  USING (public.is_operator());
CREATE POLICY "txn_read"
  ON transactions FOR SELECT TO authenticated
  USING (public.is_operator() OR public.is_ceo());

-- presets: owner only
CREATE POLICY "presets_owner"
  ON presets FOR ALL TO authenticated
  USING (operator_id = auth.uid())
  WITH CHECK (operator_id = auth.uid());

-- monthly_rollups: operators write; operators + CEO read
CREATE POLICY "rollups_insert"
  ON monthly_rollups FOR INSERT TO authenticated
  WITH CHECK (public.is_operator());
CREATE POLICY "rollups_update"
  ON monthly_rollups FOR UPDATE TO authenticated
  USING (public.is_operator());
CREATE POLICY "rollups_delete"
  ON monthly_rollups FOR DELETE TO authenticated
  USING (public.is_operator());
CREATE POLICY "rollups_read"
  ON monthly_rollups FOR SELECT TO authenticated
  USING (public.is_operator() OR public.is_ceo());

-- Storage RLS (operators upload + read; CEO has no access)
CREATE POLICY "storage_insert"
  ON storage.objects FOR INSERT TO authenticated
  WITH CHECK (bucket_id = 'uploads' AND public.is_operator());
CREATE POLICY "storage_read"
  ON storage.objects FOR SELECT TO authenticated
  USING (bucket_id = 'uploads' AND public.is_operator());
