-- Run this once in your Supabase project's SQL editor, after
-- supabase_migration.sql has already been applied.
--
-- Adds one column to the existing pptx_meta table: `builds` - the
-- per-slide bullet/object build-step data extracted client-side in
-- pptxParse.ts (see extractBuildsForSlide), stored the same way
-- notes/transitions already are.
--
-- Safe to run even if pptx_meta doesn't exist yet, or already has this
-- column - both branches are idempotent.

alter table pptx_meta
  add column if not exists builds jsonb not null default '{}'::jsonb;
