create table if not exists public.thesis_audit_users (
  id uuid primary key default gen_random_uuid(),
  email text not null unique,
  password_hash text not null,
  submissions_used integer not null default 0,
  submission_quota integer not null default 100,
  account_status text not null default 'active',
  is_admin boolean not null default false,
  invite_code text,
  invited_by uuid references public.thesis_audit_users(id),
  register_ip text not null default '',
  register_user_agent text not null default '',
  last_login_at timestamptz,
  last_login_ip text not null default '',
  last_login_user_agent text not null default '',
  last_audit_at timestamptz,
  last_audit_ip text not null default '',
  last_audit_user_agent text not null default '',
  created_at timestamptz not null default now()
);

alter table public.thesis_audit_users
add column if not exists submission_quota integer not null default 100;

alter table public.thesis_audit_users
add column if not exists account_status text not null default 'active';

alter table public.thesis_audit_users
add column if not exists is_admin boolean not null default false;

alter table public.thesis_audit_users
add column if not exists invite_code text;

alter table public.thesis_audit_users
add column if not exists invited_by uuid references public.thesis_audit_users(id);

alter table public.thesis_audit_users
add column if not exists register_ip text not null default '';

alter table public.thesis_audit_users
add column if not exists register_user_agent text not null default '';

alter table public.thesis_audit_users
add column if not exists last_login_at timestamptz;

alter table public.thesis_audit_users
add column if not exists last_login_ip text not null default '';

alter table public.thesis_audit_users
add column if not exists last_login_user_agent text not null default '';

alter table public.thesis_audit_users
add column if not exists last_audit_at timestamptz;

alter table public.thesis_audit_users
add column if not exists last_audit_ip text not null default '';

alter table public.thesis_audit_users
add column if not exists last_audit_user_agent text not null default '';

create unique index if not exists thesis_audit_users_invite_code_key
on public.thesis_audit_users(invite_code)
where invite_code is not null;

create table if not exists public.thesis_audit_registration_codes (
  id uuid primary key default gen_random_uuid(),
  code text not null unique,
  note text not null default '',
  max_uses integer not null default 20,
  used_count integer not null default 0,
  is_active boolean not null default true,
  created_by text not null default '',
  created_at timestamptz not null default now(),
  constraint thesis_audit_registration_codes_max_uses_positive check (max_uses > 0),
  constraint thesis_audit_registration_codes_used_count_nonnegative check (used_count >= 0),
  constraint thesis_audit_registration_codes_used_count_within_limit check (used_count <= max_uses)
);

create unique index if not exists thesis_audit_registration_codes_code_key
on public.thesis_audit_registration_codes(code);

create index if not exists thesis_audit_registration_codes_created_at_idx
on public.thesis_audit_registration_codes(created_at desc);

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
  college_name text not null default '',
  college_source text not null default '',
  college_raw_text text not null default '',
  client_ip text not null default '',
  user_agent text not null default '',
  original_storage_backend text not null default '',
  original_storage_path text not null default '',
  original_gcs_path text not null default '',
  original_drive_file_id text not null default '',
  original_drive_path text not null default '',
  original_size_bytes bigint not null default 0,
  original_sha256 text not null default '',
  report_storage_backend text not null default '',
  report_gcs_path text not null default '',
  report_size_bytes bigint not null default 0,
  report_sha256 text not null default '',
  created_at timestamptz not null default now()
);

alter table public.thesis_audit_reports
add column if not exists client_ip text not null default '';

alter table public.thesis_audit_reports
add column if not exists user_agent text not null default '';

alter table public.thesis_audit_reports
add column if not exists original_storage_backend text not null default '';

alter table public.thesis_audit_reports
add column if not exists original_storage_path text not null default '';

alter table public.thesis_audit_reports
add column if not exists original_gcs_path text not null default '';

alter table public.thesis_audit_reports
add column if not exists original_drive_file_id text not null default '';

alter table public.thesis_audit_reports
add column if not exists original_drive_path text not null default '';

alter table public.thesis_audit_reports
add column if not exists original_size_bytes bigint not null default 0;

alter table public.thesis_audit_reports
add column if not exists original_sha256 text not null default '';

alter table public.thesis_audit_reports
add column if not exists report_storage_backend text not null default '';

alter table public.thesis_audit_reports
add column if not exists report_gcs_path text not null default '';

alter table public.thesis_audit_reports
add column if not exists report_size_bytes bigint not null default 0;

alter table public.thesis_audit_reports
add column if not exists report_sha256 text not null default '';

alter table public.thesis_audit_reports
add column if not exists college_name text not null default '';

alter table public.thesis_audit_reports
add column if not exists college_source text not null default '';

alter table public.thesis_audit_reports
add column if not exists college_raw_text text not null default '';

create index if not exists thesis_audit_reports_user_id_created_at_idx
on public.thesis_audit_reports(user_id, created_at desc);

create index if not exists thesis_audit_reports_college_created_at_idx
on public.thesis_audit_reports(college_name, created_at desc);

create table if not exists public.thesis_audit_events (
  id uuid primary key default gen_random_uuid(),
  event_type text not null,
  user_id uuid references public.thesis_audit_users(id),
  user_email text not null default '',
  path text not null default '',
  client_ip text not null default '',
  user_agent text not null default '',
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists thesis_audit_events_type_created_at_idx
on public.thesis_audit_events(event_type, created_at desc);

create index if not exists thesis_audit_events_created_at_idx
on public.thesis_audit_events(created_at desc);

alter table public.thesis_audit_users enable row level security;

alter table public.thesis_audit_admin_logs enable row level security;

alter table public.thesis_audit_registration_codes enable row level security;

alter table public.thesis_audit_reports enable row level security;

alter table public.thesis_audit_events enable row level security;

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
      and tablename = 'thesis_audit_registration_codes'
      and policyname = 'service role can manage thesis audit registration codes'
  ) then
    create policy "service role can manage thesis audit registration codes"
    on public.thesis_audit_registration_codes
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

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'thesis_audit_events'
      and policyname = 'service role can manage thesis audit events'
  ) then
    create policy "service role can manage thesis audit events"
    on public.thesis_audit_events
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

create or replace function public.consume_thesis_audit_registration_code(
  target_code text
)
returns table (
  id uuid,
  code text,
  note text,
  max_uses integer,
  used_count integer,
  is_active boolean,
  created_by text,
  created_at timestamptz
)
language plpgsql
security definer
set search_path = public
as $$
begin
  return query
  update public.thesis_audit_registration_codes c
  set used_count = c.used_count + 1
  where c.code = upper(regexp_replace(coalesce(target_code, ''), '[^A-Za-z0-9]', '', 'g'))
    and c.is_active = true
    and c.used_count < c.max_uses
  returning
    c.id,
    c.code,
    c.note,
    c.max_uses,
    c.used_count,
    c.is_active,
    c.created_by,
    c.created_at;
end;
$$;
