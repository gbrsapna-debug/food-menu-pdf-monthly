-- Run this once in Supabase: SQL Editor -> New query -> paste -> Run.
-- Stores the whole menu (week1..week5) as one JSON document.
-- The website backend uses the SERVICE_ROLE key (bypasses RLS),
-- so no public policies are needed and the browser never sees any key.

create table if not exists app_data (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz default now()
);
