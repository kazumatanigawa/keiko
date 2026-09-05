begin;

create table if not exists public.team_membership_events (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references public.profiles(id) on delete cascade,
  team_id uuid references public.teams(id) on delete set null,
  action text not null check (action in ('join', 'transfer_in', 'transfer_out', 'graduate')),
  created_at timestamptz not null default now()
);

create index if not exists team_membership_events_user_created_idx
  on public.team_membership_events (user_id, created_at desc);

alter table public.team_membership_events enable row level security;
revoke all on table public.team_membership_events from public, anon, authenticated;

create or replace function public.keiko_membership_context_for_user(
  p_user_id uuid,
  p_preferred_team_id uuid default null
)
returns jsonb
language sql
stable
security definer
set search_path = public
as $function$
  with memberships as (
    select
      tm.team_id,
      tm.team_role,
      tm.joined_at,
      t.team_name,
      t.team_type,
      t.category,
      coalesce(t.category, '') = 'personal' as is_personal
    from public.team_members tm
    join public.teams t on t.id = tm.team_id and t.status = 'active'
    where tm.user_id = p_user_id
  ),
  selected as (
    select m.*
    from memberships m
    order by
      case when p_preferred_team_id is not null and m.team_id = p_preferred_team_id then 0 else 1 end,
      case when m.is_personal then 1 else 0 end,
      m.joined_at,
      m.team_name
    limit 1
  )
  select jsonb_build_object(
    'team_id', s.team_id,
    'team_name', s.team_name,
    'team_type', s.team_type,
    'team_role', s.team_role,
    'is_personal', coalesce(s.is_personal, false),
    'teams', coalesce((
      select jsonb_agg(jsonb_build_object(
        'teamId', m.team_id,
        'teamName', m.team_name,
        'teamType', m.team_type,
        'teamRole', m.team_role,
        'isPersonal', m.is_personal
      ) order by m.is_personal, m.joined_at, m.team_name)
      from memberships m
    ), '[]'::jsonb)
  )
  from (select 1) seed
  left join selected s on true;
$function$;

create or replace function public.get_keiko_session_context(p_team_id uuid default null)
returns jsonb
language sql
stable
security definer
set search_path = public
as $function$
  select jsonb_build_object(
    'user_id', p.id,
    'display_name', p.display_name,
    'user_type', p.user_type,
    'legacy_user_id', p.legacy_user_id,
    'grade', sp.grade,
    'role_label', sp.role_label,
    'term', sp.term
  ) || public.keiko_membership_context_for_user(p.id, p_team_id)
  from public.profiles p
  left join public.student_profiles sp on sp.user_id = p.id
  where p.id = auth.uid()
    and p.status = 'active';
$function$;

create or replace function public.manage_keiko_membership(
  p_user_id uuid,
  p_action text,
  p_target_team_id uuid default null,
  p_preferred_team_id uuid default null
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_profile public.profiles%rowtype;
  v_target public.teams%rowtype;
  v_personal_team_id uuid;
begin
  if auth.role() <> 'service_role' then
    raise exception 'service_role_required' using errcode = '42501';
  end if;
  if p_action not in ('join', 'transfer', 'graduate') then
    raise exception 'invalid_membership_action' using errcode = '22023';
  end if;

  perform pg_advisory_xact_lock(hashtextextended(p_user_id::text, 0));

  select * into v_profile
  from public.profiles
  where id = p_user_id and status = 'active'
  for update;
  if not found then raise exception 'active_profile_not_found' using errcode = 'P0001'; end if;

  if p_action in ('join', 'transfer') then
    select * into v_target
    from public.teams
    where id = p_target_team_id
      and status = 'active'
      and coalesce(category, '') <> 'personal';
    if not found then raise exception 'target_team_not_found' using errcode = 'P0001'; end if;
  end if;

  if p_action = 'join' then
    delete from public.team_members tm
    using public.teams t
    where tm.team_id = t.id
      and tm.user_id = p_user_id
      and coalesce(t.category, '') = 'personal';

    if not exists (
      select 1 from public.team_members
      where user_id = p_user_id and team_id = p_target_team_id
    ) then
      insert into public.team_members (team_id, user_id, team_role)
      values (p_target_team_id, p_user_id, 'member');
      insert into public.team_membership_events (user_id, team_id, action)
      values (p_user_id, p_target_team_id, 'join');
    end if;
  elsif p_action = 'transfer' then
    insert into public.team_membership_events (user_id, team_id, action)
    select p_user_id, tm.team_id, 'transfer_out'
    from public.team_members tm
    join public.teams t on t.id = tm.team_id
    where tm.user_id = p_user_id
      and tm.team_id <> p_target_team_id
      and coalesce(t.category, '') <> 'personal';

    delete from public.team_members
    where user_id = p_user_id and team_id <> p_target_team_id;

    if not exists (
      select 1 from public.team_members
      where user_id = p_user_id and team_id = p_target_team_id
    ) then
      insert into public.team_members (team_id, user_id, team_role)
      values (p_target_team_id, p_user_id, 'member');
    end if;
    insert into public.team_membership_events (user_id, team_id, action)
    values (p_user_id, p_target_team_id, 'transfer_in');
  else
    insert into public.team_membership_events (user_id, team_id, action)
    select p_user_id, tm.team_id, 'graduate'
    from public.team_members tm
    join public.teams t on t.id = tm.team_id
    where tm.user_id = p_user_id
      and coalesce(t.category, '') <> 'personal';

    delete from public.team_members where user_id = p_user_id;

    select id into v_personal_team_id
    from public.teams
    where owner_user_id = p_user_id and coalesce(category, '') = 'personal'
    order by created_at
    limit 1;

    if v_personal_team_id is null then
      insert into public.teams (
        team_name, team_type, category, audience_type, owner_user_id,
        status, legacy_team_id
      ) values (
        v_profile.display_name || '（個人）', 'general', 'personal', 'personal', p_user_id,
        'active', 'personal_' || replace(p_user_id::text, '-', '')
      ) returning id into v_personal_team_id;
    else
      update public.teams set status = 'active', updated_at = now()
      where id = v_personal_team_id;
    end if;

    insert into public.team_members (team_id, user_id, team_role)
    values (v_personal_team_id, p_user_id, 'owner_admin');

    p_preferred_team_id := v_personal_team_id;
  end if;

  update public.profiles p
  set user_type = case
        when exists (
          select 1 from public.team_members tm
          join public.teams t on t.id = tm.team_id
          where tm.user_id = p_user_id and t.team_type = 'student'
        ) then 'student'
        else 'general'
      end,
      updated_at = now()
  where p.id = p_user_id;

  if exists (
    select 1 from public.team_members tm
    join public.teams t on t.id = tm.team_id
    where tm.user_id = p_user_id and t.team_type = 'student'
  ) and not exists (select 1 from public.student_profiles where user_id = p_user_id) then
    insert into public.student_profiles (user_id, school_name, grade, role_label, term)
    values (p_user_id, v_target.team_name, '', '', '');
  end if;

  if not exists (select 1 from public.general_profiles where user_id = p_user_id) then
    insert into public.general_profiles (user_id, category, bio)
    values (p_user_id, case when p_action = 'graduate' then '個人' else coalesce(v_target.team_name, '一般') end, '');
  end if;

  return public.keiko_membership_context_for_user(
    p_user_id,
    coalesce(p_preferred_team_id, p_target_team_id)
  ) || jsonb_build_object(
    'user_type', (select user_type from public.profiles where id = p_user_id)
  );
end;
$function$;

revoke all on function public.keiko_membership_context_for_user(uuid, uuid) from public, anon, authenticated;
revoke all on function public.get_keiko_session_context(uuid) from public, anon;
grant execute on function public.get_keiko_session_context(uuid) to authenticated;
revoke all on function public.manage_keiko_membership(uuid, text, uuid, uuid) from public, anon, authenticated;
grant execute on function public.manage_keiko_membership(uuid, text, uuid, uuid) to service_role;

notify pgrst, 'reload schema';

commit;
