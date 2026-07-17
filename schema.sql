-- ============================================================
-- ASUS Maintenance Dashboard — Supabase schema
-- Run this ONCE in Supabase Dashboard > SQL Editor
-- ============================================================

-- ---------- TICKETS (raw data from "Details Sheet") ----------
create table if not exists public.tickets (
  ticket_id                 text primary key,   -- alphanumeric (e.g. 12563, L12417)
  new_ticket_no             text,
  freshdesk_id              text,
  region                    text,
  branch                    text,
  store_name                text,
  store_type                text,
  city                      text,
  address                   text,
  state                     text,
  state_code                text,
  issue_raised_by           text,
  designation               text,
  cmkt_name                 text,
  logo_flag                 text,          -- Logo / Non Logo
  city_classification       text,          -- Tier 1 / Tier 2 / ROI
  asset_installation_date   date,
  issue_raised_date         date,
  issue_raised_year         int,
  quarter_raised            text,          -- e.g. 'Q1 2025'
  problem_reported          text,
  issue_category            text,
  budget_category           text,          -- Electrical / Fixtures / POSM / Branding / Civil
  status                    text,          -- stage of ticket
  approval_date             date,
  approval_tat              numeric,
  tat_city_type             numeric,       -- allowed TAT days per city type
  execution_tentative_date  date,
  rectification_date        date,
  rectification_time        numeric,       -- days to rectify
  quarter_rectified         text,
  final_status              text,          -- Open / Closed
  half_yearly               text,
  ageing_closure_bucket     text,
  rectified_year            int,
  responsibility            text,          -- Channelplay / Asus
  tat_follow                text,          -- InTAT / OutTAT
  material_deployed         text,
  qty                       text,
  extra                     jsonb,         -- all remaining columns preserved as-is
  updated_at                timestamptz not null default now()
);

create index if not exists tickets_region_idx   on public.tickets (region);
create index if not exists tickets_final_idx    on public.tickets (final_status);
create index if not exists tickets_raised_idx   on public.tickets (issue_raised_date);
create index if not exists tickets_budget_idx   on public.tickets (budget_category);

-- ---------- STORE MASTER (from "Total Store Covered" sheet) ----------
create table if not exists public.stores (
  store_code  text primary key,
  created_at  timestamptz not null default now()
);

-- ---------- UPLOAD AUDIT LOG ----------
create table if not exists public.upload_logs (
  id             bigint generated always as identity primary key,
  uploaded_at    timestamptz not null default now(),
  uploaded_by    text,
  file_name      text,
  as_on_date     date,
  total_rows     int,
  inserted_rows  int,
  updated_rows   int,
  deleted_rows   int default 0,
  store_count    int,
  warnings       int default 0,
  note           text
);

-- ---------- USER PROFILES & ROLES ----------
-- role: 'viewer' (read-only, default) | 'uploader' | 'admin'
create table if not exists public.profiles (
  id     uuid primary key references auth.users on delete cascade,
  email  text,
  role   text not null default 'viewer'
);

-- Auto-create a profile whenever a user is created
create or replace function public.handle_new_user()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  insert into public.profiles (id, email) values (new.id, new.email)
  on conflict (id) do nothing;
  return new;
end; $$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- Helper: does the current user have write rights?
create or replace function public.can_write()
returns boolean language sql stable security definer set search_path = public as $$
  select exists (
    select 1 from public.profiles
    where id = auth.uid() and role in ('uploader','admin')
  );
$$;

-- ---------- ROW LEVEL SECURITY ----------
alter table public.tickets      enable row level security;
alter table public.stores       enable row level security;
alter table public.upload_logs  enable row level security;
alter table public.profiles     enable row level security;

-- Any signed-in user can read
create policy "read tickets"     on public.tickets     for select to authenticated using (true);
create policy "read stores"      on public.stores      for select to authenticated using (true);
create policy "read upload_logs" on public.upload_logs for select to authenticated using (true);
create policy "read own profile" on public.profiles    for select to authenticated using (id = auth.uid());

-- Only uploader/admin can write
create policy "write tickets ins" on public.tickets     for insert to authenticated with check (public.can_write());
create policy "write tickets upd" on public.tickets     for update to authenticated using (public.can_write());
create policy "write tickets del" on public.tickets     for delete to authenticated using (public.can_write());
create policy "write stores ins"  on public.stores      for insert to authenticated with check (public.can_write());
create policy "write stores del"  on public.stores      for delete to authenticated using (public.can_write());
create policy "write logs ins"    on public.upload_logs for insert to authenticated with check (public.can_write());

-- ============================================================
-- AFTER RUNNING THIS SCRIPT:
-- 1. Authentication > Users > Add user  → create logins for yourself & the client
--    (disable public signups in Authentication > Providers > Email if desired)
-- 2. Give yourself upload rights (run in SQL editor):
--      update public.profiles set role = 'admin' where email = 'you@example.com';
--    The client stays 'viewer' (read-only) automatically.
-- 3. OPTIONAL (file backup): Storage > New bucket → name it  raw-uploads  (private).
--    Then add policies allowing authenticated INSERT/SELECT on that bucket,
--    or simply skip — the app works without it.
-- ============================================================
