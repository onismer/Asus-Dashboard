-- ============================================================
-- MIGRATION: allow ANY signed-in user to upload data.
-- Run once in Supabase Dashboard > SQL Editor.
-- Every upload is still recorded in upload_logs (who/when/what),
-- visible in the dashboard under Upload Data > Upload History.
-- ============================================================
create or replace function public.can_write()
returns boolean language sql stable security definer set search_path = public as $$
  select auth.uid() is not null;   -- previously: role in ('uploader','admin')
$$;
