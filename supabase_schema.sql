create table if not exists public.thesis_audit_users (
  id uuid primary key default gen_random_uuid(),
  email text not null unique,
  password_hash text not null,
  submissions_used integer not null default 0,
  submission_quota integer not null default 2,
  account_status text not null default 'active',
  is_admin boolean not null default false,
  invite_code text,
  invited_by uuid references public.thesis_audit_users(id),
  created_at timestamptz not null default now()
);

alter table public.thesis_audit_users
add column if not exists submission_quota integer not null default 2;

alter table public.thesis_audit_users
add column if not exists account_status text not null default 'active';

alter table public.thesis_audit_users
add column if not exists is_admin boolean not null default false;

alter table public.thesis_audit_users
add column if not exists invite_code text;

alter table public.thesis_audit_users
add column if not exists invited_by uuid references public.thesis_audit_users(id);

create unique index if not exists thesis_audit_users_invite_code_key
on public.thesis_audit_users(invite_code)
where invite_code is not null;

create table if not exists public.thesis_audit_admin_logs (
  id uuid primary key default gen_random_uuid(),
  actor_user_id uuid references public.thesis_audit_users(id),
  actor_email text not null default '',
  action text not null,
  target_user_id uuid references public.thesis_audit_users(id),
  target_email text not null default '',
  summary text not null,
  details jsonb,
  created_at timestamptz not null default now()
);

create table if not exists public.thesis_audit_reports (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.thesis_audit_users(id),
  user_email text not null default '',
  original_filename text not null,
  report_filename text not null default '',
  report_storage_path text not null default '',
  status text not null default 'success',
  error_message text not null default '',
  created_at timestamptz not null default now()
);

create index if not exists thesis_audit_reports_user_id_created_at_idx
on public.thesis_audit_reports(user_id, created_at desc);

alter table public.thesis_audit_users enable row level security;

alter table public.thesis_audit_admin_logs enable row level security;

alter table public.thesis_audit_reports enable row level security;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'thesis_audit_users'
      and policyname = 'service role can manage thesis audit users'
  ) then
    create policy "service role can manage thesis audit users"
    on public.thesis_audit_users
    for all
    using (auth.role() = 'service_role')
    with check (auth.role() = 'service_role');
  end if;
end $$;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'thesis_audit_admin_logs'
      and policyname = 'service role can manage thesis audit admin logs'
  ) then
    create policy "service role can manage thesis audit admin logs"
    on public.thesis_audit_admin_logs
    for all
    using (auth.role() = 'service_role')
    with check (auth.role() = 'service_role');
  end if;
end $$;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'thesis_audit_reports'
      and policyname = 'service role can manage thesis audit reports'
  ) then
    create policy "service role can manage thesis audit reports"
    on public.thesis_audit_reports
    for all
    using (auth.role() = 'service_role')
    with check (auth.role() = 'service_role');
  end if;
end $$;

create or replace function public.increment_thesis_audit_submissions(
  target_user_id uuid,
  max_allowed integer
)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare
  updated_count integer;
begin
  update public.thesis_audit_users
  set submissions_used = submissions_used + 1
  where id = target_user_id
    and submissions_used < max_allowed;

  get diagnostics updated_count = row_count;
  return updated_count = 1;
end;
$$;
