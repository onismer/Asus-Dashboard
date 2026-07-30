-- ============================================================
-- MIGRATION: upload rollback (remove incorrect uploads)
-- Run ONCE in Supabase Dashboard > SQL Editor.
--
-- Every upload now stores a before-image (snapshot) of the rows it
-- inserts / overwrites / deletes. An admin can then remove an upload
-- from Upload History: inserted tickets are deleted and overwritten
-- tickets are restored to their exact previous values.
-- Historical (frozen) uploads are excluded. Removed uploads stay in
-- the history, marked with who removed them and when.
-- ============================================================

create table if not exists public.upload_snapshots (
  id         bigint generated always as identity primary key,
  upload_id  bigint not null references public.upload_logs(id) on delete cascade,
  ticket_id  text   not null,
  action     text   not null check (action in ('insert','update','delete')),
  prev       jsonb,                    -- full previous row (null for 'insert')
  created_at timestamptz not null default now()
);
create index if not exists snapshots_upload_idx on public.upload_snapshots (upload_id);

alter table public.upload_logs
  add column if not exists snapshot_rows int,          -- null = pre-feature upload (not removable)
  add column if not exists removed_by    text,
  add column if not exists removed_at    timestamptz;

-- RLS
alter table public.upload_snapshots enable row level security;
create policy "read snapshots"  on public.upload_snapshots for select to authenticated using (true);
create policy "write snapshots" on public.upload_snapshots for insert to authenticated with check (public.can_write());
create policy "del snapshots"   on public.upload_snapshots for delete to authenticated using (public.can_write());

-- upload_logs now needs UPDATE (final note + removed marking)
drop policy if exists "update logs" on public.upload_logs;
create policy "update logs" on public.upload_logs for update to authenticated using (public.can_write());

revoke all on public.upload_snapshots from anon;
