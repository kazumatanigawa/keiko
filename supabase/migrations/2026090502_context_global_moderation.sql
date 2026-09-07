begin;

alter table public.profiles
  add column if not exists app_role text not null default 'player',
  add column if not exists global_participation_enabled boolean not null default false,
  add column if not exists global_joined_at timestamptz;

alter table public.profiles drop constraint if exists profiles_app_role_check;
alter table public.profiles
  add constraint profiles_app_role_check check (app_role in ('player', 'operator'));

update public.profiles
set app_role = 'operator', updated_at = now()
where id = '6603fa37-433e-4042-877c-98935ffabba0'::uuid;

alter table public.teams
  add column if not exists global_notes_enabled boolean not null default false,
  add column if not exists global_scope text not null default 'disabled';

alter table public.teams drop constraint if exists teams_global_scope_check;
alter table public.teams
  add constraint teams_global_scope_check check (global_scope in ('disabled', 'global', 'school'));

update public.teams
set global_notes_enabled = false,
    global_scope = 'disabled',
    updated_at = now()
where team_type = 'student' or coalesce(category, '') = 'personal';

alter table public.team_notes
  add column if not exists visibility text not null default 'team',
  add column if not exists author_team_id_snapshot uuid references public.teams(id) on delete set null,
  add column if not exists author_team_name_snapshot text not null default '';

alter table public.team_notes drop constraint if exists team_notes_visibility_check;
alter table public.team_notes
  add constraint team_notes_visibility_check check (visibility in ('private', 'team', 'global'));

update public.team_notes n
set visibility = 'team',
    author_team_id_snapshot = n.team_id,
    author_team_name_snapshot = coalesce(t.team_name, '')
from public.teams t
where t.id = n.team_id
  and (n.author_team_id_snapshot is null or n.author_team_name_snapshot = '');

alter table public.team_note_comments
  add column if not exists author_team_id_snapshot uuid references public.teams(id) on delete set null,
  add column if not exists author_team_name_snapshot text not null default '';

update public.team_note_comments c
set author_team_id_snapshot = c.team_id,
    author_team_name_snapshot = coalesce(t.team_name, '')
from public.teams t
where t.id = c.team_id
  and (c.author_team_id_snapshot is null or c.author_team_name_snapshot = '');

create table if not exists public.content_reports (
  id uuid primary key default gen_random_uuid(),
  reporter_user_id uuid not null references public.profiles(id),
  content_type text not null check (content_type in ('note', 'comment')),
  content_id uuid not null,
  note_id uuid not null,
  reported_author_user_id uuid references public.profiles(id) on delete set null,
  reported_author_name_snapshot text not null default '',
  reported_team_name_snapshot text not null default '',
  content_snapshot text not null default '',
  reason text not null,
  status text not null default 'pending' check (status in ('pending', 'resolved', 'dismissed')),
  email_status text not null default 'pending' check (email_status in ('pending', 'sent', 'failed')),
  email_error text not null default '',
  handled_by uuid references public.profiles(id) on delete set null,
  handled_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (reporter_user_id, content_type, content_id)
);

create index if not exists content_reports_status_created_idx
  on public.content_reports (status, created_at desc);

create table if not exists public.keiko_operator_events (
  id uuid primary key default gen_random_uuid(),
  actor_user_id uuid not null references public.profiles(id),
  action text not null,
  target_type text not null,
  target_id uuid,
  details jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists keiko_operator_events_actor_created_idx
  on public.keiko_operator_events (actor_user_id, created_at desc);

alter table public.content_reports enable row level security;
alter table public.keiko_operator_events enable row level security;
revoke all on table public.content_reports from public, anon, authenticated;
revoke all on table public.keiko_operator_events from public, anon, authenticated;

create index if not exists team_notes_visibility_team_created_idx
  on public.team_notes (visibility, team_id, created_at desc, id desc)
  where status = 'active';

create or replace function public.keiko_ensure_personal_membership(p_user_id uuid)
returns uuid
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_team_id uuid;
  v_name text;
begin
  select p.display_name into v_name
  from public.profiles p
  where p.id = p_user_id and p.status = 'active';
  if not found then raise exception 'active_profile_not_found' using errcode = 'P0001'; end if;

  select t.id into v_team_id
  from public.teams t
  where t.owner_user_id = p_user_id and coalesce(t.category, '') = 'personal'
  order by t.created_at
  limit 1;

  if v_team_id is null then
    insert into public.teams (
      team_name, team_type, category, audience_type, owner_user_id,
      status, legacy_team_id, global_notes_enabled, global_scope
    ) values (
      v_name || '（個人）', 'general', 'personal', 'personal', p_user_id,
      'active', 'personal_' || replace(p_user_id::text, '-', ''), false, 'disabled'
    ) returning id into v_team_id;
  else
    update public.teams
    set status = 'active', global_notes_enabled = false, global_scope = 'disabled', updated_at = now()
    where id = v_team_id;
  end if;

  if not exists (
    select 1 from public.team_members tm
    where tm.team_id = v_team_id and tm.user_id = p_user_id
  ) then
    insert into public.team_members (team_id, user_id, team_role)
    values (v_team_id, p_user_id, 'owner_admin');
  end if;

  return v_team_id;
end;
$function$;

do $backfill$
declare
  v_user record;
begin
  for v_user in select id from public.profiles where status = 'active'
  loop
    perform public.keiko_ensure_personal_membership(v_user.id);
  end loop;
end;
$backfill$;

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
      t.global_notes_enabled,
      t.global_scope,
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
    'global_notes_enabled', coalesce(s.global_notes_enabled, false),
    'global_scope', coalesce(s.global_scope, 'disabled'),
    'teams', coalesce((
      select jsonb_agg(jsonb_build_object(
        'teamId', m.team_id,
        'teamName', m.team_name,
        'teamType', m.team_type,
        'teamRole', m.team_role,
        'isPersonal', m.is_personal,
        'globalNotesEnabled', m.global_notes_enabled,
        'globalScope', m.global_scope
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
    'app_role', p.app_role,
    'global_participation_enabled', p.global_participation_enabled,
    'grade', sp.grade,
    'role_label', sp.role_label,
    'term', sp.term
  ) || public.keiko_membership_context_for_user(p.id, p_team_id)
  from public.profiles p
  left join public.student_profiles sp on sp.user_id = p.id
  where p.id = auth.uid()
    and p.status = 'active';
$function$;

create or replace function public.set_keiko_global_participation(
  p_enabled boolean,
  p_hide_existing boolean default false
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
begin
  if v_user_id is null then raise exception 'auth_required' using errcode = '42501'; end if;
  perform public.keiko_ensure_personal_membership(v_user_id);

  update public.profiles
  set global_participation_enabled = p_enabled,
      global_joined_at = case when p_enabled then coalesce(global_joined_at, now()) else global_joined_at end,
      updated_at = now()
  where id = v_user_id and status = 'active';
  if not found then raise exception 'inactive_profile' using errcode = '42501'; end if;

  if not p_enabled and p_hide_existing then
    update public.team_notes
    set status = 'hidden', updated_at = now()
    where author_user_id = v_user_id and visibility = 'global' and status = 'active';
  end if;

  return jsonb_build_object(
    'status', 'ok',
    'globalParticipationEnabled', p_enabled,
    'teams', public.keiko_membership_context_for_user(v_user_id, null)->'teams'
  );
end;
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
  select * into v_profile from public.profiles
  where id = p_user_id and status = 'active' for update;
  if not found then raise exception 'active_profile_not_found' using errcode = 'P0001'; end if;

  v_personal_team_id := public.keiko_ensure_personal_membership(p_user_id);

  if p_action in ('join', 'transfer') then
    select * into v_target from public.teams
    where id = p_target_team_id and status = 'active' and coalesce(category, '') <> 'personal';
    if not found then raise exception 'target_team_not_found' using errcode = 'P0001'; end if;
  end if;

  if p_action = 'join' then
    if not exists (
      select 1 from public.team_members where user_id = p_user_id and team_id = p_target_team_id
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

    delete from public.team_members tm
    using public.teams t
    where tm.team_id = t.id
      and tm.user_id = p_user_id
      and tm.team_id <> p_target_team_id
      and coalesce(t.category, '') <> 'personal';

    if not exists (
      select 1 from public.team_members where user_id = p_user_id and team_id = p_target_team_id
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
    where tm.user_id = p_user_id and coalesce(t.category, '') <> 'personal';

    delete from public.team_members tm
    using public.teams t
    where tm.team_id = t.id
      and tm.user_id = p_user_id
      and coalesce(t.category, '') <> 'personal';

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
    p_user_id, coalesce(p_preferred_team_id, p_target_team_id)
  ) || jsonb_build_object(
    'user_type', (select user_type from public.profiles where id = p_user_id),
    'app_role', v_profile.app_role,
    'global_participation_enabled', v_profile.global_participation_enabled
  );
end;
$function$;

create or replace function public.keiko_timekeeper_payload(p_assignment_id uuid)
returns jsonb
language sql
stable
security definer
set search_path = public
as $function$
  select jsonb_build_object(
    'assignmentId', a.id,
    'practiceDate', a.practice_date,
    'scheduledUserId', a.scheduled_user_id,
    'scheduledName', coalesce(scheduled.display_name, 'メンバー'),
    'assignedUserId', a.assigned_user_id,
    'assignedName', coalesce(assigned.display_name, 'メンバー'),
    'absentUserIds', to_jsonb(a.absent_user_ids),
    'absentNames', coalesce((
      select jsonb_agg(coalesce(p.display_name, 'メンバー') order by u.ordinality)
      from unnest(a.absent_user_ids) with ordinality as u(user_id, ordinality)
      left join public.profiles p on p.id = u.user_id
    ), '[]'::jsonb),
    'isSubstitute', a.assigned_user_id <> a.scheduled_user_id,
    'cycleNumber', c.cycle_number,
    'cyclePosition', a.cycle_index,
    'cycleSize', cardinality(c.member_order),
    'nextUserId', next_member.user_id,
    'nextName', next_profile.display_name,
    'canReplace', exists (
      select 1
      from unnest(c.member_order) as candidate(user_id)
      join public.team_members tm on tm.team_id = a.team_id and tm.user_id = candidate.user_id
      join public.profiles p on p.id = tm.user_id and p.status = 'active'
      where candidate.user_id <> a.assigned_user_id
        and not (candidate.user_id = any(a.absent_user_ids))
    )
  )
  from public.timekeeper_assignments a
  join public.timekeeper_cycles c on c.id = a.cycle_id
  left join public.profiles scheduled on scheduled.id = a.scheduled_user_id
  left join public.profiles assigned on assigned.id = a.assigned_user_id
  left join lateral (
    select candidate.user_id
    from unnest(c.member_order) with ordinality as candidate(user_id, position)
    join public.team_members tm on tm.team_id = a.team_id and tm.user_id = candidate.user_id
    join public.profiles p on p.id = candidate.user_id and p.status = 'active'
    where candidate.position > a.cycle_index
    order by candidate.position
    limit 1
  ) next_member on true
  left join public.profiles next_profile on next_profile.id = next_member.user_id
  where a.id = p_assignment_id;
$function$;

create or replace function public.replace_keiko_timekeeper(p_team_id uuid, p_date date)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_assignment public.timekeeper_assignments%rowtype;
  v_cycle public.timekeeper_cycles%rowtype;
  v_candidate uuid;
  v_new_absent uuid[];
  v_prefix uuid[];
  v_remaining uuid[];
begin
  perform public.get_keiko_timekeeper(p_team_id, p_date);
  perform pg_advisory_xact_lock(hashtextextended(p_team_id::text, 0));

  select * into v_assignment
  from public.timekeeper_assignments
  where team_id = p_team_id and practice_date = p_date
  for update;

  select * into v_cycle
  from public.timekeeper_cycles
  where id = v_assignment.cycle_id
  for update;

  v_new_absent := array_append(v_assignment.absent_user_ids, v_assignment.assigned_user_id);

  select candidate.user_id into v_candidate
  from unnest(v_cycle.member_order) with ordinality as candidate(user_id, position)
  join public.team_members tm on tm.team_id = p_team_id and tm.user_id = candidate.user_id
  join public.profiles p on p.id = candidate.user_id and p.status = 'active'
  where candidate.position > v_assignment.cycle_index
    and not (candidate.user_id = any(v_new_absent))
  order by candidate.position
  limit 1;

  if v_candidate is not null then
    v_prefix := coalesce(v_cycle.member_order[1:v_assignment.cycle_index - 1], '{}'::uuid[]);
    select coalesce(array_agg(candidate.user_id order by candidate.position), '{}'::uuid[])
    into v_remaining
    from unnest(v_cycle.member_order) with ordinality as candidate(user_id, position)
    where candidate.position > v_assignment.cycle_index
      and candidate.user_id <> v_candidate
      and not (candidate.user_id = any(v_new_absent));

    update public.timekeeper_cycles
    set member_order = v_prefix || array[v_candidate] || v_new_absent || v_remaining,
        updated_at = now()
    where id = v_cycle.id;
  else
    select candidate.user_id into v_candidate
    from unnest(v_cycle.member_order) with ordinality as candidate(user_id, position)
    join public.team_members tm on tm.team_id = p_team_id and tm.user_id = candidate.user_id
    join public.profiles p on p.id = candidate.user_id and p.status = 'active'
    where not (candidate.user_id = any(v_new_absent))
    order by candidate.position
    limit 1;

    if v_candidate is null then
      raise exception 'no_alternate_timekeeper' using errcode = 'P0001';
    end if;

    update public.timekeeper_cycles
    set current_index = cardinality(member_order),
        carryover_order = v_new_absent,
        updated_at = now()
    where id = v_cycle.id;
    v_assignment.cycle_index := cardinality(v_cycle.member_order);
  end if;

  update public.timekeeper_assignments
  set assigned_user_id = v_candidate,
      absent_user_ids = v_new_absent,
      cycle_index = v_assignment.cycle_index,
      updated_at = now()
  where id = v_assignment.id;

  return public.keiko_timekeeper_payload(v_assignment.id);
end;
$function$;

create or replace function public.get_keiko_home_dashboard(p_team_id uuid, p_date date)
returns jsonb
language plpgsql
volatile
security invoker
set search_path = public
as $function$
declare
  v_summary jsonb;
  v_is_personal boolean;
begin
  select coalesce(t.category, '') = 'personal' into v_is_personal
  from public.team_members tm
  join public.teams t on t.id = tm.team_id and t.status = 'active'
  where tm.team_id = p_team_id and tm.user_id = auth.uid();
  if not found then raise exception 'team_access_denied' using errcode = '42501'; end if;

  with summary as (
    select count(*)::integer as total_count
    from public.practice_logs pl
    where pl.user_id = auth.uid() and pl.team_id = p_team_id
  ), latest as (
    select jsonb_build_object(
      'id', pl.id, 'practice_date', pl.practice_date, 'condition', pl.condition,
      'learning', pl.learning, 'next_action', pl.next_action, 'good_new', pl.good_new,
      'achievement_status', pl.achievement_status, 'why_missed', pl.why_missed,
      'retry_plan', pl.retry_plan, 'team_id', pl.team_id, 'team_name', t.team_name,
      'created_at', pl.created_at, 'updated_at', pl.updated_at
    ) as log
    from public.practice_logs pl
    join public.teams t on t.id = pl.team_id
    where pl.user_id = auth.uid() and pl.team_id = p_team_id
    order by pl.practice_date desc, pl.created_at desc, pl.id desc
    limit 1
  )
  select jsonb_build_object('total_count', s.total_count, 'latest_log', l.log)
  into v_summary
  from summary s left join latest l on true;

  if v_is_personal then
    return v_summary || jsonb_build_object('timekeeper', null);
  end if;
  return v_summary || jsonb_build_object('timekeeper', public.get_keiko_timekeeper(p_team_id, p_date));
end;
$function$;

create or replace function public.get_keiko_log_target_summary(p_team_id uuid)
returns jsonb
language sql
stable
security invoker
set search_path = public
as $function$
  select coalesce((
    select jsonb_build_object(
      'total_count', count(*) over (),
      'latest_log', jsonb_build_object(
        'id', pl.id, 'practice_date', pl.practice_date, 'condition', pl.condition,
        'learning', pl.learning, 'next_action', pl.next_action, 'good_new', pl.good_new,
        'achievement_status', pl.achievement_status, 'why_missed', pl.why_missed,
        'retry_plan', pl.retry_plan, 'team_id', pl.team_id, 'team_name', t.team_name,
        'created_at', pl.created_at, 'updated_at', pl.updated_at
      )
    )
    from public.practice_logs pl
    join public.teams t on t.id = pl.team_id
    where pl.user_id = auth.uid()
      and pl.team_id = p_team_id
      and exists (
        select 1 from public.team_members tm
        where tm.team_id = p_team_id and tm.user_id = auth.uid()
      )
    order by pl.practice_date desc, pl.created_at desc, pl.id desc
    limit 1
  ), jsonb_build_object('total_count', 0, 'latest_log', null));
$function$;

create or replace function public.get_keiko_logs_page_v2(
  p_team_id uuid,
  p_cursor_date date default null,
  p_cursor_created_at timestamptz default null,
  p_cursor_id uuid default null,
  p_limit integer default 20
)
returns jsonb
language sql
stable
security invoker
set search_path = public
as $function$
  with params as (
    select greatest(1, least(coalesce(p_limit, 20), 50)) as page_limit
  ), ordered_rows as (
    select pl.*, t.team_name
    from public.practice_logs pl
    join public.teams t on t.id = pl.team_id
    where pl.user_id = auth.uid()
      and pl.team_id = p_team_id
      and exists (
        select 1 from public.team_members tm
        where tm.team_id = p_team_id and tm.user_id = auth.uid()
      )
      and (
        p_cursor_date is null or p_cursor_created_at is null or p_cursor_id is null
        or (pl.practice_date, pl.created_at, pl.id) < (p_cursor_date, p_cursor_created_at, p_cursor_id)
      )
    order by pl.practice_date desc, pl.created_at desc, pl.id desc
    limit (select page_limit + 1 from params)
  ), page_rows as (
    select * from ordered_rows
    order by practice_date desc, created_at desc, id desc
    limit (select page_limit from params)
  ), last_row as (
    select practice_date, created_at, id from page_rows
    order by practice_date asc, created_at asc, id asc limit 1
  )
  select jsonb_build_object(
    'logs', coalesce((select jsonb_agg(jsonb_build_object(
      'id', p.id, 'date', p.practice_date, 'cond', p.condition,
      'learning', coalesce(p.learning, ''), 'next', coalesce(p.next_action, ''),
      'goodNew', coalesce(p.good_new, ''), 'achievementStatus', coalesce(p.achievement_status, ''),
      'whyMissed', coalesce(p.why_missed, ''), 'retryPlan', coalesce(p.retry_plan, ''),
      'teamId', p.team_id, 'teamName', p.team_name,
      'createdAt', p.created_at, 'updatedAt', p.updated_at
    ) order by p.practice_date desc, p.created_at desc, p.id desc) from page_rows p), '[]'::jsonb),
    'hasMore', (select count(*) > (select page_limit from params) from ordered_rows),
    'nextCursor', case
      when (select count(*) > (select page_limit from params) from ordered_rows)
      then (select jsonb_build_object('date', l.practice_date, 'createdAt', l.created_at, 'id', l.id) from last_row l)
      else null
    end
  );
$function$;

create or replace function public.keiko_can_access_note(p_note_id uuid, p_user_id uuid)
returns boolean
language sql
stable
security definer
set search_path = public
as $function$
  select coalesce((
    select case n.visibility
      when 'private' then n.author_user_id = p_user_id
      when 'team' then exists (
        select 1 from public.team_members tm where tm.team_id = n.team_id and tm.user_id = p_user_id
      )
      when 'global' then exists (
        select 1 from public.profiles p
        where p.id = p_user_id and p.status = 'active' and p.global_participation_enabled
      )
      else false
    end
    from public.team_notes n
    join public.teams t on t.id = n.team_id and t.status = 'active'
    where n.id = p_note_id and n.status = 'active'
      and (n.visibility <> 'global'
        or coalesce(t.category, '') = 'personal'
        or (t.team_type <> 'student' and t.global_notes_enabled and t.global_scope = 'global'))
  ), false);
$function$;

create or replace function public.get_keiko_notes_feed(
  p_scope text,
  p_team_id uuid default null,
  p_cursor_created_at timestamptz default null,
  p_cursor_id uuid default null,
  p_limit integer default 20
)
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_result jsonb;
begin
  if v_user_id is null then raise exception 'auth_required' using errcode = '42501'; end if;
  if p_scope not in ('private', 'team', 'global') then raise exception 'invalid_scope' using errcode = '22023'; end if;
  if p_scope = 'team' and not exists (
    select 1 from public.team_members tm where tm.team_id = p_team_id and tm.user_id = v_user_id
  ) then raise exception 'team_forbidden' using errcode = '42501'; end if;
  if p_scope = 'global' and not exists (
    select 1 from public.profiles p
    where p.id = v_user_id and p.status = 'active' and p.global_participation_enabled
  ) then raise exception 'global_participation_required' using errcode = '42501'; end if;

  with params as (
    select greatest(1, least(coalesce(p_limit, 20), 50)) as page_limit
  ), ordered_rows as (
    select n.*
    from public.team_notes n
    join public.teams t on t.id = n.team_id and t.status = 'active'
    where n.status = 'active'
      and (
        (p_scope = 'private' and n.visibility = 'private' and n.author_user_id = v_user_id)
        or (p_scope = 'team' and n.team_id = p_team_id and n.visibility in ('team', 'global'))
        or (p_scope = 'global' and n.visibility = 'global'
          and (coalesce(t.category, '') = 'personal'
            or (t.team_type <> 'student' and t.global_notes_enabled and t.global_scope = 'global')))
      )
      and (
        p_cursor_created_at is null or p_cursor_id is null
        or (n.created_at, n.id) < (p_cursor_created_at, p_cursor_id)
      )
    order by n.created_at desc, n.id desc
    limit (select page_limit + 1 from params)
  ), page_rows as (
    select * from ordered_rows order by created_at desc, id desc
    limit (select page_limit from params)
  ), last_row as (
    select created_at, id from page_rows order by created_at asc, id asc limit 1
  )
  select jsonb_build_object(
    'notes', coalesce((select jsonb_agg(jsonb_build_object(
      'noteId', n.id, 'teamId', n.team_id, 'visibility', n.visibility,
      'authorUserId', n.author_user_id,
      'authorName', coalesce(nullif(n.author_name_snapshot, ''), 'メンバー'),
      'authorTeamName', coalesce(nullif(n.author_team_name_snapshot, ''), '個人'),
      'title', n.title, 'body', n.body, 'createdAt', n.created_at, 'updatedAt', n.updated_at,
      'commentCount', (select count(*) from public.team_note_comments c where c.note_id = n.id and c.status = 'active'),
      'comments', '[]'::jsonb, 'commentsLoaded', false,
      'canReport', n.author_user_id <> v_user_id
    ) order by n.created_at desc, n.id desc) from page_rows n), '[]'::jsonb),
    'hasMore', (select count(*) > (select page_limit from params) from ordered_rows),
    'nextCursor', case
      when (select count(*) > (select page_limit from params) from ordered_rows)
      then (select jsonb_build_object('createdAt', l.created_at, 'id', l.id) from last_row l)
      else null
    end
  ) into v_result;
  return v_result;
end;
$function$;

create or replace function public.get_keiko_note_comments_v2(
  p_note_id uuid,
  p_cursor_created_at timestamptz default null,
  p_cursor_id uuid default null,
  p_limit integer default 20
)
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_result jsonb;
begin
  if v_user_id is null or not public.keiko_can_access_note(p_note_id, v_user_id) then
    raise exception 'note_forbidden' using errcode = '42501';
  end if;
  with params as (
    select greatest(1, least(coalesce(p_limit, 20), 50)) as page_limit
  ), ordered_rows as (
    select c.* from public.team_note_comments c
    where c.note_id = p_note_id and c.status = 'active'
      and (p_cursor_created_at is null or p_cursor_id is null or (c.created_at, c.id) > (p_cursor_created_at, p_cursor_id))
    order by c.created_at asc, c.id asc
    limit (select page_limit + 1 from params)
  ), page_rows as (
    select * from ordered_rows order by created_at asc, id asc limit (select page_limit from params)
  ), last_row as (
    select created_at, id from page_rows order by created_at desc, id desc limit 1
  )
  select jsonb_build_object(
    'comments', coalesce((select jsonb_agg(jsonb_build_object(
      'commentId', c.id, 'noteId', c.note_id, 'teamId', c.team_id,
      'authorUserId', c.author_user_id,
      'authorName', coalesce(nullif(c.author_name_snapshot, ''), 'メンバー'),
      'authorTeamName', coalesce(nullif(c.author_team_name_snapshot, ''), '個人'),
      'body', c.body, 'createdAt', c.created_at, 'updatedAt', c.updated_at,
      'isEdited', c.updated_at is distinct from c.created_at,
      'canEdit', c.author_user_id = v_user_id, 'canDelete', c.author_user_id = v_user_id,
      'canReport', c.author_user_id <> v_user_id
    ) order by c.created_at asc, c.id asc) from page_rows c), '[]'::jsonb),
    'hasMore', (select count(*) > (select page_limit from params) from ordered_rows),
    'nextCursor', case
      when (select count(*) > (select page_limit from params) from ordered_rows)
      then (select jsonb_build_object('createdAt', l.created_at, 'id', l.id) from last_row l)
      else null
    end
  ) into v_result;
  return v_result;
end;
$function$;

create or replace function public.save_keiko_note_v2(
  p_scope text,
  p_team_id uuid,
  p_request_id text,
  p_title text,
  p_body text
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_author_name text;
  v_team public.teams%rowtype;
  v_legacy_id text;
  v_row public.team_notes%rowtype;
  v_duplicate boolean := false;
  v_team_label text;
begin
  if v_user_id is null then raise exception 'auth_required' using errcode = '42501'; end if;
  if p_scope not in ('private', 'team', 'global') then raise exception 'invalid_scope' using errcode = '22023'; end if;
  if p_request_id !~ '^[A-Za-z0-9_-]{16,80}$' then raise exception 'invalid_request_id'; end if;
  if nullif(btrim(p_title), '') is null or length(p_title) > 120 then raise exception 'invalid_title'; end if;
  if nullif(btrim(p_body), '') is null or length(p_body) > 5000 then raise exception 'invalid_body'; end if;

  select p.display_name into v_author_name from public.profiles p
  where p.id = v_user_id and p.status = 'active';
  if not found then raise exception 'inactive_profile' using errcode = '42501'; end if;

  select t.* into v_team
  from public.team_members tm
  join public.teams t on t.id = tm.team_id and t.status = 'active'
  where tm.user_id = v_user_id and tm.team_id = p_team_id;
  if not found then raise exception 'team_forbidden' using errcode = '42501'; end if;

  if p_scope = 'team' and coalesce(v_team.category, '') = 'personal' then
    raise exception 'team_scope_requires_team' using errcode = '22023';
  end if;
  if p_scope = 'global' then
    if not exists (
      select 1 from public.profiles p where p.id = v_user_id and p.global_participation_enabled
    ) then raise exception 'global_participation_required' using errcode = '42501'; end if;
    if coalesce(v_team.category, '') <> 'personal'
      and (v_team.team_type = 'student' or not v_team.global_notes_enabled or v_team.global_scope <> 'global')
    then raise exception 'team_global_disabled' using errcode = '42501'; end if;
    if (select count(*) from public.team_notes n
        where n.author_user_id = v_user_id and n.visibility = 'global'
          and n.created_at > now() - interval '1 minute') >= 3
    then raise exception 'rate_limit_exceeded' using errcode = 'P0001'; end if;
  end if;

  v_team_label := case when coalesce(v_team.category, '') = 'personal' then '個人' else v_team.team_name end;
  v_legacy_id := 'app:' || p_request_id;
  select * into v_row from public.team_notes n
  where n.author_user_id = v_user_id and n.legacy_note_id = v_legacy_id limit 1;
  if found then
    v_duplicate := true;
  else
    begin
      insert into public.team_notes (
        legacy_note_id, team_id, author_user_id, author_name_snapshot,
        author_team_id_snapshot, author_team_name_snapshot,
        title, body, visibility, status
      ) values (
        v_legacy_id, p_team_id, v_user_id, v_author_name,
        p_team_id, v_team_label, btrim(p_title), btrim(p_body), p_scope, 'active'
      ) returning * into v_row;
    exception when unique_violation then
      v_duplicate := true;
      select * into v_row from public.team_notes n
      where n.author_user_id = v_user_id and n.legacy_note_id = v_legacy_id limit 1;
    end;
  end if;

  return jsonb_build_object('status', 'ok', 'duplicate', v_duplicate, 'note', jsonb_build_object(
    'noteId', v_row.id, 'teamId', v_row.team_id, 'visibility', v_row.visibility,
    'authorUserId', v_row.author_user_id,
    'authorName', coalesce(nullif(v_row.author_name_snapshot, ''), 'メンバー'),
    'authorTeamName', coalesce(nullif(v_row.author_team_name_snapshot, ''), '個人'),
    'title', v_row.title, 'body', v_row.body, 'createdAt', v_row.created_at,
    'updatedAt', v_row.updated_at, 'commentCount', 0, 'comments', '[]'::jsonb,
    'commentsLoaded', false, 'canReport', false
  ));
end;
$function$;

create or replace function public.add_keiko_note_comment_v2(
  p_note_id uuid,
  p_author_team_id uuid,
  p_request_id text,
  p_body text
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_note public.team_notes%rowtype;
  v_team public.teams%rowtype;
  v_author_name text;
  v_team_label text;
  v_legacy_id text;
  v_row public.team_note_comments%rowtype;
  v_duplicate boolean := false;
begin
  if v_user_id is null or not public.keiko_can_access_note(p_note_id, v_user_id) then
    raise exception 'note_forbidden' using errcode = '42501';
  end if;
  if p_request_id !~ '^[A-Za-z0-9_-]{16,80}$' then raise exception 'invalid_request_id'; end if;
  if nullif(btrim(p_body), '') is null or length(p_body) > 5000 then raise exception 'invalid_body'; end if;

  select * into v_note from public.team_notes where id = p_note_id and status = 'active';
  select p.display_name into v_author_name from public.profiles p
  where p.id = v_user_id and p.status = 'active';
  if not found then raise exception 'inactive_profile' using errcode = '42501'; end if;

  if v_note.visibility = 'team' then
    p_author_team_id := v_note.team_id;
  elsif p_author_team_id is null then
    p_author_team_id := public.keiko_ensure_personal_membership(v_user_id);
  end if;

  select t.* into v_team
  from public.team_members tm join public.teams t on t.id = tm.team_id and t.status = 'active'
  where tm.user_id = v_user_id and tm.team_id = p_author_team_id;
  if not found then raise exception 'team_forbidden' using errcode = '42501'; end if;

  if v_note.visibility = 'global' and coalesce(v_team.category, '') <> 'personal'
    and (v_team.team_type = 'student' or not v_team.global_notes_enabled or v_team.global_scope <> 'global')
  then raise exception 'team_global_disabled' using errcode = '42501'; end if;
  if v_note.visibility = 'global' and (
    select count(*) from public.team_note_comments c
    where c.author_user_id = v_user_id and c.created_at > now() - interval '1 minute'
  ) >= 10 then raise exception 'rate_limit_exceeded' using errcode = 'P0001'; end if;

  v_team_label := case when coalesce(v_team.category, '') = 'personal' then '個人' else v_team.team_name end;
  v_legacy_id := 'app:' || p_request_id;
  select * into v_row from public.team_note_comments c
  where c.author_user_id = v_user_id and c.legacy_comment_id = v_legacy_id limit 1;
  if found then
    v_duplicate := true;
  else
    begin
      insert into public.team_note_comments (
        legacy_comment_id, note_id, team_id, author_user_id, author_name_snapshot,
        author_team_id_snapshot, author_team_name_snapshot, body, status
      ) values (
        v_legacy_id, p_note_id, v_note.team_id, v_user_id, v_author_name,
        p_author_team_id, v_team_label, btrim(p_body), 'active'
      ) returning * into v_row;
    exception when unique_violation then
      v_duplicate := true;
      select * into v_row from public.team_note_comments c
      where c.author_user_id = v_user_id and c.legacy_comment_id = v_legacy_id limit 1;
    end;
  end if;

  return jsonb_build_object('status', 'ok', 'duplicate', v_duplicate, 'comment', jsonb_build_object(
    'commentId', v_row.id, 'noteId', v_row.note_id, 'teamId', v_row.team_id,
    'authorUserId', v_row.author_user_id,
    'authorName', coalesce(nullif(v_row.author_name_snapshot, ''), 'メンバー'),
    'authorTeamName', coalesce(nullif(v_row.author_team_name_snapshot, ''), '個人'),
    'body', v_row.body, 'createdAt', v_row.created_at, 'updatedAt', v_row.updated_at,
    'isEdited', false, 'canEdit', true, 'canDelete', true, 'canReport', false
  ));
end;
$function$;

create or replace function public.update_keiko_note_comment_v2(
  p_comment_id uuid,
  p_body text
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_row public.team_note_comments%rowtype;
begin
  if v_user_id is null then raise exception 'auth_required' using errcode = '42501'; end if;
  if nullif(btrim(p_body), '') is null or length(p_body) > 5000 then raise exception 'invalid_body'; end if;
  update public.team_note_comments c
  set body = btrim(p_body), updated_at = now()
  where c.id = p_comment_id and c.author_user_id = v_user_id and c.status = 'active'
  returning * into v_row;
  if not found then raise exception 'comment_forbidden' using errcode = '42501'; end if;
  return jsonb_build_object('status', 'ok', 'comment', jsonb_build_object(
    'commentId', v_row.id, 'noteId', v_row.note_id, 'teamId', v_row.team_id,
    'authorUserId', v_row.author_user_id,
    'authorName', coalesce(nullif(v_row.author_name_snapshot, ''), 'メンバー'),
    'authorTeamName', coalesce(nullif(v_row.author_team_name_snapshot, ''), '個人'),
    'body', v_row.body, 'createdAt', v_row.created_at, 'updatedAt', v_row.updated_at,
    'isEdited', true, 'canEdit', true, 'canDelete', true, 'canReport', false
  ));
end;
$function$;

create or replace function public.delete_keiko_note_comment_v2(p_comment_id uuid)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_note_id uuid;
begin
  if v_user_id is null then raise exception 'auth_required' using errcode = '42501'; end if;
  update public.team_note_comments c
  set status = 'deleted', updated_at = now()
  where c.id = p_comment_id and c.author_user_id = v_user_id and c.status = 'active'
  returning c.note_id into v_note_id;
  if not found then raise exception 'comment_forbidden' using errcode = '42501'; end if;
  return jsonb_build_object('status', 'ok', 'commentId', p_comment_id, 'noteId', v_note_id);
end;
$function$;

create or replace function public.get_keiko_operator_teams(p_actor_user_id uuid)
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $function$
begin
  if auth.role() <> 'service_role' or not exists (
    select 1 from public.profiles p where p.id = p_actor_user_id and p.status = 'active' and p.app_role = 'operator'
  ) then raise exception 'operator_required' using errcode = '42501'; end if;
  return jsonb_build_object('teams', coalesce((
    select jsonb_agg(jsonb_build_object(
      'teamId', t.id, 'teamName', t.team_name, 'teamType', t.team_type,
      'globalNotesEnabled', t.global_notes_enabled, 'globalScope', t.global_scope
    ) order by t.team_name)
    from public.teams t
    where t.status = 'active' and coalesce(t.category, '') <> 'personal'
  ), '[]'::jsonb));
end;
$function$;

create or replace function public.update_keiko_team_global_settings(
  p_actor_user_id uuid,
  p_team_id uuid,
  p_enabled boolean
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_team public.teams%rowtype;
begin
  if auth.role() <> 'service_role' or not exists (
    select 1 from public.profiles p where p.id = p_actor_user_id and p.status = 'active' and p.app_role = 'operator'
  ) then raise exception 'operator_required' using errcode = '42501'; end if;

  select * into v_team from public.teams where id = p_team_id and status = 'active' for update;
  if not found or coalesce(v_team.category, '') = 'personal' then raise exception 'team_not_found'; end if;
  if p_enabled and v_team.team_type = 'student' then raise exception 'school_global_forbidden'; end if;

  update public.teams
  set global_notes_enabled = p_enabled,
      global_scope = case when p_enabled then 'global' else 'disabled' end,
      updated_at = now()
  where id = p_team_id;

  insert into public.keiko_operator_events (actor_user_id, action, target_type, target_id, details)
  values (p_actor_user_id, 'update_global_settings', 'team', p_team_id, jsonb_build_object('enabled', p_enabled));

  return jsonb_build_object(
    'status', 'ok', 'teamId', p_team_id, 'globalNotesEnabled', p_enabled,
    'globalScope', case when p_enabled then 'global' else 'disabled' end
  );
end;
$function$;

create or replace function public.create_keiko_content_report(
  p_reporter_user_id uuid,
  p_content_type text,
  p_content_id uuid,
  p_reason text
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_note public.team_notes%rowtype;
  v_comment public.team_note_comments%rowtype;
  v_report public.content_reports%rowtype;
  v_author_id uuid;
  v_author_name text;
  v_team_name text;
  v_snapshot text;
begin
  if auth.role() <> 'service_role' then raise exception 'service_role_required' using errcode = '42501'; end if;
  if p_content_type not in ('note', 'comment') then raise exception 'invalid_content_type'; end if;
  if nullif(btrim(p_reason), '') is null or length(p_reason) > 1000 then raise exception 'invalid_reason'; end if;
  if not exists (select 1 from public.profiles p where p.id = p_reporter_user_id and p.status = 'active') then
    raise exception 'inactive_reporter' using errcode = '42501';
  end if;

  if p_content_type = 'note' then
    select * into v_note from public.team_notes where id = p_content_id and status = 'active';
    if not found or not public.keiko_can_access_note(v_note.id, p_reporter_user_id) then raise exception 'content_forbidden'; end if;
    v_author_id := v_note.author_user_id;
    v_author_name := v_note.author_name_snapshot;
    v_team_name := v_note.author_team_name_snapshot;
    v_snapshot := v_note.title || E'\n\n' || v_note.body;
  else
    select * into v_comment from public.team_note_comments where id = p_content_id and status = 'active';
    if not found or not public.keiko_can_access_note(v_comment.note_id, p_reporter_user_id) then raise exception 'content_forbidden'; end if;
    v_note.id := v_comment.note_id;
    v_author_id := v_comment.author_user_id;
    v_author_name := v_comment.author_name_snapshot;
    v_team_name := v_comment.author_team_name_snapshot;
    v_snapshot := v_comment.body;
  end if;
  if v_author_id = p_reporter_user_id then raise exception 'cannot_report_own_content'; end if;

  insert into public.content_reports (
    reporter_user_id, content_type, content_id, note_id, reported_author_user_id,
    reported_author_name_snapshot, reported_team_name_snapshot, content_snapshot, reason
  ) values (
    p_reporter_user_id, p_content_type, p_content_id, v_note.id, v_author_id,
    coalesce(v_author_name, ''), coalesce(v_team_name, ''), left(v_snapshot, 6000), btrim(p_reason)
  )
  on conflict (reporter_user_id, content_type, content_id) do update
    set reason = excluded.reason,
        status = 'pending',
        email_status = 'pending',
        email_error = '',
        handled_by = null,
        handled_at = null,
        updated_at = now()
  returning * into v_report;

  return jsonb_build_object(
    'status', 'ok', 'reportId', v_report.id,
    'contentType', v_report.content_type,
    'authorName', v_report.reported_author_name_snapshot,
    'teamName', v_report.reported_team_name_snapshot,
    'content', v_report.content_snapshot,
    'reason', v_report.reason,
    'createdAt', v_report.created_at
  );
end;
$function$;

create or replace function public.mark_keiko_report_email(
  p_report_id uuid,
  p_status text,
  p_error text default ''
)
returns void
language plpgsql
volatile
security definer
set search_path = public
as $function$
begin
  if auth.role() <> 'service_role' then raise exception 'service_role_required' using errcode = '42501'; end if;
  if p_status not in ('pending', 'sent', 'failed') then raise exception 'invalid_email_status'; end if;
  update public.content_reports
  set email_status = p_status, email_error = left(coalesce(p_error, ''), 1000), updated_at = now()
  where id = p_report_id;
end;
$function$;

create or replace function public.get_keiko_operator_reports(p_actor_user_id uuid, p_limit integer default 20)
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $function$
begin
  if auth.role() <> 'service_role' or not exists (
    select 1 from public.profiles p where p.id = p_actor_user_id and p.status = 'active' and p.app_role = 'operator'
  ) then raise exception 'operator_required' using errcode = '42501'; end if;
  return jsonb_build_object('reports', coalesce((
    select jsonb_agg(jsonb_build_object(
      'reportId', r.id, 'contentType', r.content_type, 'contentId', r.content_id,
      'authorName', r.reported_author_name_snapshot, 'teamName', r.reported_team_name_snapshot,
      'content', r.content_snapshot, 'reason', r.reason, 'status', r.status,
      'emailStatus', r.email_status, 'createdAt', r.created_at
    ) order by r.created_at desc)
    from (
      select * from public.content_reports
      where status = 'pending'
      order by created_at desc
      limit greatest(1, least(coalesce(p_limit, 20), 100))
    ) r
  ), '[]'::jsonb));
end;
$function$;

create or replace function public.moderate_keiko_report(
  p_actor_user_id uuid,
  p_report_id uuid,
  p_action text
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_report public.content_reports%rowtype;
begin
  if auth.role() <> 'service_role' or not exists (
    select 1 from public.profiles p where p.id = p_actor_user_id and p.status = 'active' and p.app_role = 'operator'
  ) then raise exception 'operator_required' using errcode = '42501'; end if;
  if p_action not in ('resolve', 'dismiss', 'hide_content') then raise exception 'invalid_moderation_action'; end if;

  select * into v_report from public.content_reports where id = p_report_id for update;
  if not found then raise exception 'report_not_found'; end if;

  if p_action = 'hide_content' then
    if v_report.content_type = 'note' then
      update public.team_notes set status = 'hidden_by_operator', updated_at = now() where id = v_report.content_id;
    else
      update public.team_note_comments set status = 'hidden_by_operator', updated_at = now() where id = v_report.content_id;
    end if;
  end if;

  update public.content_reports
  set status = case when p_action = 'dismiss' then 'dismissed' else 'resolved' end,
      handled_by = p_actor_user_id, handled_at = now(), updated_at = now()
  where id = p_report_id;

  insert into public.keiko_operator_events (actor_user_id, action, target_type, target_id, details)
  values (p_actor_user_id, p_action, 'report', p_report_id, jsonb_build_object(
    'contentType', v_report.content_type, 'contentId', v_report.content_id
  ));

  return jsonb_build_object('status', 'ok', 'reportId', p_report_id, 'action', p_action);
end;
$function$;

revoke all on function public.keiko_ensure_personal_membership(uuid) from public, anon, authenticated;
grant execute on function public.keiko_ensure_personal_membership(uuid) to service_role;
revoke all on function public.keiko_membership_context_for_user(uuid, uuid) from public, anon, authenticated;
revoke all on function public.get_keiko_session_context(uuid) from public, anon;
grant execute on function public.get_keiko_session_context(uuid) to authenticated;
revoke all on function public.set_keiko_global_participation(boolean, boolean) from public, anon;
grant execute on function public.set_keiko_global_participation(boolean, boolean) to authenticated;
revoke all on function public.manage_keiko_membership(uuid, text, uuid, uuid) from public, anon, authenticated;
grant execute on function public.manage_keiko_membership(uuid, text, uuid, uuid) to service_role;
revoke all on function public.get_keiko_home_dashboard(uuid, date) from public, anon;
grant execute on function public.get_keiko_home_dashboard(uuid, date) to authenticated;
revoke all on function public.get_keiko_log_target_summary(uuid) from public, anon;
grant execute on function public.get_keiko_log_target_summary(uuid) to authenticated;
revoke all on function public.get_keiko_logs_page_v2(uuid, date, timestamptz, uuid, integer) from public, anon;
grant execute on function public.get_keiko_logs_page_v2(uuid, date, timestamptz, uuid, integer) to authenticated;
revoke all on function public.keiko_can_access_note(uuid, uuid) from public, anon, authenticated;
revoke all on function public.get_keiko_notes_feed(text, uuid, timestamptz, uuid, integer) from public, anon;
grant execute on function public.get_keiko_notes_feed(text, uuid, timestamptz, uuid, integer) to authenticated;
revoke all on function public.get_keiko_note_comments_v2(uuid, timestamptz, uuid, integer) from public, anon;
grant execute on function public.get_keiko_note_comments_v2(uuid, timestamptz, uuid, integer) to authenticated;
revoke all on function public.save_keiko_note_v2(text, uuid, text, text, text) from public, anon;
grant execute on function public.save_keiko_note_v2(text, uuid, text, text, text) to authenticated;
revoke all on function public.add_keiko_note_comment_v2(uuid, uuid, text, text) from public, anon;
grant execute on function public.add_keiko_note_comment_v2(uuid, uuid, text, text) to authenticated;
revoke all on function public.update_keiko_note_comment_v2(uuid, text) from public, anon;
grant execute on function public.update_keiko_note_comment_v2(uuid, text) to authenticated;
revoke all on function public.delete_keiko_note_comment_v2(uuid) from public, anon;
grant execute on function public.delete_keiko_note_comment_v2(uuid) to authenticated;
revoke all on function public.get_keiko_operator_teams(uuid) from public, anon, authenticated;
grant execute on function public.get_keiko_operator_teams(uuid) to service_role;
revoke all on function public.update_keiko_team_global_settings(uuid, uuid, boolean) from public, anon, authenticated;
grant execute on function public.update_keiko_team_global_settings(uuid, uuid, boolean) to service_role;
revoke all on function public.create_keiko_content_report(uuid, text, uuid, text) from public, anon, authenticated;
grant execute on function public.create_keiko_content_report(uuid, text, uuid, text) to service_role;
revoke all on function public.mark_keiko_report_email(uuid, text, text) from public, anon, authenticated;
grant execute on function public.mark_keiko_report_email(uuid, text, text) to service_role;
revoke all on function public.get_keiko_operator_reports(uuid, integer) from public, anon, authenticated;
grant execute on function public.get_keiko_operator_reports(uuid, integer) to service_role;
revoke all on function public.moderate_keiko_report(uuid, uuid, text) from public, anon, authenticated;
grant execute on function public.moderate_keiko_report(uuid, uuid, text) to service_role;

notify pgrst, 'reload schema';

commit;
