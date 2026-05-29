create table if not exists public.thesis_audit_users (
  id uuid primary key default gen_random_uuid(),
  email text not null unique,
  password_hash text not null,
  submissions_used integer not null default 0,
  submission_quota integer not null default 3,
  created_at timestamptz not null default now()
);

alter table public.thesis_audit_users
add column if not exists submission_quota integer not null default 3;

alter table public.thesis_audit_users enable row level security;

create policy "service role can manage thesis audit users"
on public.thesis_audit_users
for all
using (auth.role() = 'service_role')
with check (auth.role() = 'service_role');

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
