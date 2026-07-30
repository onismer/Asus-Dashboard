-- ============================================================
-- MIGRATION: Historical-2025 freeze + new tracker format + data protection
-- Run ONCE in Supabase Dashboard > SQL Editor.
--
-- What this does:
--  1. REMOVES all previously uploaded data (tickets, stores, upload logs)
--  2. Adds data_source ('Historical: 2025' / 'Live') + frozen flag
--  3. Adds a trigger that makes frozen rows physically un-modifiable
--  4. Adds columns for the new tracker's data (budget, approval days,
--     store code, macro category, logo logistics)
--  5. Upgrades the store master for the tracker's "Master WOD" sheet
--  6. Locks the anon (not-signed-in) role out of all data
-- ============================================================

-- 1 ── remove all previously uploaded data
truncate public.tickets;
truncate public.stores cascade;
truncate public.upload_logs restart identity;

-- 2 ── data source tagging
alter table public.tickets
  add column if not exists data_source text not null default 'Live',
  add column if not exists frozen      boolean not null default false;

-- 3 ── freeze protection: frozen rows cannot be updated or deleted, by anyone
create or replace function public.protect_frozen()
returns trigger language plpgsql as $$
begin
  if old.frozen then
    raise exception 'Ticket % belongs to the protected historical dataset (%) and cannot be modified or deleted.',
      old.ticket_id, old.data_source;
  end if;
  if tg_op = 'DELETE' then return old; end if;
  return new;
end; $$;

drop trigger if exists tickets_protect_frozen on public.tickets;
create trigger tickets_protect_frozen
  before update or delete on public.tickets
  for each row execute function public.protect_frozen();

-- 4 ── new tracker data points
alter table public.tickets
  add column if not exists store_code          text,
  add column if not exists macro_category      text,
  add column if not exists issue_type          text,     -- Production / Non Production
  add column if not exists total_budget        numeric,  -- spend per ticket
  add column if not exists approval_days       numeric,  -- CMKT approval days
  add column if not exists logo_delivery_status text,
  add column if not exists logo_courier        text,
  add column if not exists logo_vendor         text,
  add column if not exists logo_dispatch_date  date,
  add column if not exists logo_delivery_date  date;
create index if not exists tickets_source_idx on public.tickets (data_source);

-- 5 ── richer store master (from tracker's "Master WOD" sheet)
alter table public.stores
  add column if not exists store_name text,
  add column if not exists city       text,
  add column if not exists state      text,
  add column if not exists branch     text,
  add column if not exists territory  text,
  add column if not exists district   text,
  add column if not exists tm_name    text;

-- 6 ── protection hardening: anonymous (not-signed-in) visitors get nothing.
--     (RLS already restricts to authenticated; this is belt-and-braces.)
revoke all on all tables in schema public from anon;
alter default privileges in schema public revoke all on tables from anon;

-- ============================================================
-- ALSO DO (in the Supabase Dashboard, not SQL):
--   Authentication > Sign In / Providers > Email > disable "Allow new users
--   to sign up" — so only accounts YOU create can ever log in.
-- ============================================================
