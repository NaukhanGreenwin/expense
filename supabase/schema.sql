-- Enable required extension for UUID generation (if not already enabled)
create extension if not exists pgcrypto;

-- Sessions table
create table if not exists public.sessions (
  id uuid primary key,
  status text not null default 'active',
  created_at timestamp with time zone not null default now()
);

-- Uploads table, stores storage paths in Supabase Storage
create table if not exists public.uploads (
  id uuid primary key default gen_random_uuid(),
  session_id uuid not null references public.sessions(id) on delete cascade,
  storage_path text not null,
  original_name text,
  size bigint,
  created_at timestamp with time zone not null default now()
);
create index if not exists uploads_session_idx on public.uploads(session_id);

-- Expenses table
create table if not exists public.expenses (
  id bigserial primary key,
  session_id uuid,
  date date,
  merchant text,
  amount numeric(12,2),
  tax numeric(12,2),
  gl_code text,
  description text,
  name text,
  department text,
  location text,
  property_code text,
  created_at timestamp with time zone not null default now(),
  updated_at timestamp with time zone not null default now()
);
create index if not exists expenses_session_idx on public.expenses(session_id);

-- Expense splits
create table if not exists public.expense_splits (
  id bigserial primary key,
  expense_id bigint not null references public.expenses(id) on delete cascade,
  gl_code text not null,
  amount numeric(12,2) not null,
  percentage numeric(6,3)
);
create index if not exists expense_splits_expense_idx on public.expense_splits(expense_id);

-- NOTE: In production, enable RLS with policies and use service_role on server only.

