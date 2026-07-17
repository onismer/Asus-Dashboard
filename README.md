# ASUS Store Maintenance Dashboard

Static web dashboard (GitHub Pages) + Supabase database. Upload the daily cumulative Excel report; all tables and charts update automatically. Every chart and table drills down to the underlying tickets, and filtered data can be exported to Excel/CSV.

## Architecture

```
GitHub Pages (this repo — HTML/JS/CSS, no build step)
        │  supabase-js (HTTPS)
        ▼
Supabase ── Postgres (tickets, stores, upload_logs, profiles)
         ── Auth (email/password login)
         ── Storage (optional backup of original upload files)
```

## One-time setup

### 1. Supabase
1. Create a project at https://supabase.com (or use an existing one).
2. Open **SQL Editor**, paste the contents of `supabase/schema.sql`, run it.
3. **Authentication → Users → Add user** — create logins for yourself and the client.
   Recommended: Authentication → Sign In / Up → disable public sign-ups.
4. Grant yourself upload rights (SQL Editor):
   ```sql
   update public.profiles set role = 'admin' where email = 'your@email.com';
   ```
   Everyone else stays `viewer` (read-only; no Upload tab).
5. *(Optional, for file backups)* **Storage → New bucket** named `raw-uploads` (private),
   and add INSERT policy for authenticated users. Skip freely — uploads work without it.

### 2. Credentials
Open `js/config.js` and paste your **Project URL** and **anon public key**
(Supabase → Project Settings → API). The anon key is safe in a public repo;
data access is controlled by login + Row Level Security.

### 3. GitHub Pages
1. Create a GitHub repo, push this folder's contents to it.
2. Repo **Settings → Pages → Source: Deploy from a branch** → `main` / root.
3. The dashboard is live at `https://<user>.github.io/<repo>/` in ~1 minute.

### 4. First data load
Sign in → **Upload Data** tab → drop the cumulative Excel file → review the
validation report → **Confirm & Upload**.

## Daily workflow
Upload the latest cumulative file. The parser:

1. Locates the "Details Sheet" (and "Total Store Covered" sheet if present — refreshes the store master from it automatically).
2. Validates before writing anything: missing required columns block the upload; row errors (missing Ticket ID / Region / dates, duplicate Ticket IDs) are listed and those rows skipped; warnings (inconsistent casing like "Rectified By RV" → auto-normalized, closed tickets without rectification date, unknown values) are listed but don't block.
3. Shows a diff — how many tickets are new / updated / unchanged, and which DB tickets are missing from the file (optional cleanup checkbox).
4. On confirm: chunked upsert by Ticket ID, optional backup of the original file to Storage, and an entry in the upload audit log.

## Features
- **Views**: country Overview (KPIs, quarterly & half-yearly performance, stage funnel, budget split, monthly trend), Regional (region/branch), Stores & Cities (store universe coverage, repeat-issue stores, tier split), Issues & Bottlenecks (top/open/slowest categories, tickets stuck by stage × ageing), Ageing & TAT (open ageing, closure ageing CP vs ASUS/RV, TAT compliance & trend), Data Explorer (raw rows + full ticket drawer).
- **Drill-down**: click any chart bar/slice, KPI card, or summary-table row → jumps to Data Explorer filtered to those exact tickets.
- **Global filters**: year, quarter, region, branch, city tier, budget category, stage, open/closed, responsibility, free-text search — applied to every view; removable filter chips; saveable presets.
- **Export**: filtered raw data to Excel (with an "Applied Filters" sheet) or CSV; every aggregate table has its own CSV button.
- **Upload safety**: validation-before-write, diff preview, audit history, optional raw-file backup.
- **Extras**: dark mode, "data as on" badge from the filename, sortable tables, mobile-friendly layout.

## Data expectations
- Ticket IDs may be alphanumeric (e.g. `12563`, `L12417`) and must be unique.
- Dates in Excel date format or `dd-mm-yyyy` text.
- Extra/unknown columns are preserved in the database (`extra` JSON) and shown in the ticket drawer + exports — adding new columns to the report won't break anything.

## Troubleshooting
- **"Setup needed" screen** → paste credentials in `js/config.js`.
- **Login works but no data** → run `schema.sql`; check the table has rows (Supabase → Table Editor).
- **Upload button missing** → your profile role is `viewer`; run the role UPDATE from step 4.
- **"backup skipped" note in history** → the `raw-uploads` bucket doesn't exist; optional.
- **New team member** → add the user in Supabase Auth; a viewer profile is created automatically.
