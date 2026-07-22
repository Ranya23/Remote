-- Run this once in your Supabase project's SQL editor. Independent of
-- supabase_migration_auth.sql - can be run before, after, or without it.
--
-- Adds the two small pieces of state the build-animation feature needs:
--
--   1. pptx_meta.builds - per-slide build/animation info extracted client-
--      side at upload time (see pptxParse.ts). Same "public access" policy
--      as the rest of pptx_meta, since it's non-sensitive file metadata
--      anyone with the fileId (e.g. via a remote/audience QR code) already
--      needs to read regardless of login state.
--   2. sessions.build_step - which bullet/object within the *current* slide
--      is showing, synced the same way current_slide already is.

alter table pptx_meta add column if not exists builds jsonb not null default '{}'::jsonb;

alter table sessions add column if not exists build_step int not null default 0;
