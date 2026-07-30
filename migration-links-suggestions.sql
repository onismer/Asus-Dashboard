-- ============================================================
-- MIGRATION: PPT link columns + Feature Suggestion / Query wall
-- Run ONCE in Supabase Dashboard > SQL Editor.
-- ============================================================

-- 1 ── PPT hyperlink columns (extracted from the Excel cells' links)
alter table public.tickets
  add column if not exists issue_ppt_link     text,
  add column if not exists rectified_ppt_link text;

-- 2 ── suggestions / queries wall
create table if not exists public.suggestions (
  id          bigint generated always as identity primary key,
  created_at  timestamptz not null default now(),
  email       text not null,
  content     text not null,
  resolved    boolean not null default false,
  resolved_by text,
  resolved_at timestamptz
);

alter table public.suggestions enable row level security;
create policy "read suggestions"   on public.suggestions for select to authenticated using (true);
create policy "post suggestions"   on public.suggestions for insert to authenticated with check (true);
create policy "update suggestions" on public.suggestions for update to authenticated using (true);

revoke all on public.suggestions from anon;
