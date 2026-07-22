-- Run this once in your Supabase project's SQL editor, AFTER the original
-- supabase_migration.sql has already been run at least once.
-- (Project -> SQL Editor -> New query -> paste -> Run)
--
-- This migration upgrades the old "unique code + name" pseudo-login to real
-- Supabase Auth (email/password + Google OAuth). It:
--
--   1. Adds `profiles` - one row per authenticated user, created automatically
--      on sign-up.
--   2. Adds `user_id` + `size_bytes` to the existing `saved_items` table, so
--      each saved lesson/quiz is owned by a real auth user instead of a
--      free-text code, and so storage usage can be metered.
--   3. Locks `saved_items` down with real row-level security: a user can only
--      ever see/insert/update/delete their own rows (auth.uid() = user_id).
--   4. Adds a trigger that enforces, server-side, the two account limits -
--      5 presentations ("lesson" rows) and 100MB of total storage - so the
--      limits hold even if the app's own pre-upload check is ever bypassed.
--
-- The old `accounts` table and its `code`-based rows are left in place (no
-- data is deleted) so you can migrate off it at your own pace, but nothing
-- in the updated app reads or writes it anymore. Drop it whenever you're
-- ready:   drop table if exists accounts;

-- ---------------------------------------------------------------------------
-- 1. profiles
-- ---------------------------------------------------------------------------
create table if not exists profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  display_name text,
  created_at timestamptz not null default now()
);
alter table profiles enable row level security;

drop policy if exists "users can view own profile" on profiles;
create policy "users can view own profile" on profiles
  for select using (auth.uid() = id);

drop policy if exists "users can update own profile" on profiles;
create policy "users can update own profile" on profiles
  for update using (auth.uid() = id) with check (auth.uid() = id);

-- Auto-create a profile row the moment someone signs up (email/password OR
-- Google - both land here, since both create a row in auth.users).
create or replace function handle_new_user()
returns trigger
language plpgsql
security definer set search_path = public
as $$
begin
  insert into public.profiles (id, display_name)
  values (
    new.id,
    coalesce(new.raw_user_meta_data ->> 'full_name', new.raw_user_meta_data ->> 'name', split_part(new.email, '@', 1))
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function handle_new_user();

-- ---------------------------------------------------------------------------
-- 2. saved_items: move ownership from account_code (text) to user_id (uuid)
-- ---------------------------------------------------------------------------
alter table saved_items add column if not exists user_id uuid references auth.users(id) on delete cascade;
alter table saved_items add column if not exists size_bytes bigint not null default 0;
alter table saved_items alter column account_code drop not null;

create index if not exists saved_items_user_id_idx on saved_items(user_id);

-- Replace the old "wide open" policy with real per-user isolation.
drop policy if exists "public access" on saved_items;

drop policy if exists "users can view own items" on saved_items;
create policy "users can view own items" on saved_items
  for select using (auth.uid() = user_id);

drop policy if exists "users can insert own items" on saved_items;
create policy "users can insert own items" on saved_items
  for insert with check (auth.uid() = user_id);

drop policy if exists "users can update own items" on saved_items;
create policy "users can update own items" on saved_items
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);

drop policy if exists "users can delete own items" on saved_items;
create policy "users can delete own items" on saved_items
  for delete using (auth.uid() = user_id);

-- ---------------------------------------------------------------------------
-- 3. Server-side limit enforcement: 5 presentations, 100MB total storage.
--    ("Presentations" = saved_items rows where kind = 'lesson'.)
-- ---------------------------------------------------------------------------
create or replace function enforce_presentation_limits()
returns trigger
language plpgsql
security definer set search_path = public
as $$
declare
  existing_count int;
  existing_bytes bigint;
begin
  if new.kind <> 'lesson' then
    return new;
  end if;

  select count(*), coalesce(sum(size_bytes), 0)
    into existing_count, existing_bytes
    from saved_items
    where user_id = new.user_id and kind = 'lesson';

  if existing_count >= 5 then
    raise exception 'PRESENTATION_LIMIT_REACHED: you can only have 5 saved presentations at a time'
      using errcode = 'P0001';
  end if;

  if existing_bytes + coalesce(new.size_bytes, 0) > 100 * 1024 * 1024 then
    raise exception 'STORAGE_LIMIT_REACHED: this would exceed your 100MB storage limit'
      using errcode = 'P0001';
  end if;

  return new;
end;
$$;

drop trigger if exists enforce_presentation_limits_trigger on saved_items;
create trigger enforce_presentation_limits_trigger
  before insert on saved_items
  for each row execute function enforce_presentation_limits();
