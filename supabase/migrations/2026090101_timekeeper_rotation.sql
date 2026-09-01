begin;

create table if not exists public.timekeeper_cycles (
  id uuid primary key default gen_random_uuid(),
  team_id uuid not null references public.teams(id) on delete cascade,
  cycle_number integer not null check (cycle_number > 0),
  member_order uuid[] not null check (cardinality(member_order) > 0),
  current_index integer not null default 1 check (current_index > 0),
  carryover_order uuid[] not null default '{}'::uuid[],
  status text not null default 'active' check (status in ('active', 'completed')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (team_id, cycle_number)
);

create unique index if not exists timekeeper_cycles_active_team_uidx
  on public.timekeeper_cycles (team_id)
  where status = 'active';

create table if not exists public.timekeeper_assignments (
  id uuid primary key default gen_random_uuid(),
  team_id uuid not null references public.teams(id) on delete cascade,
  practice_date date not null,
  cycle_id uuid not null references public.timekeeper_cycles(id) on delete cascade,
  cycle_index integer not null check (cycle_index > 0),
  scheduled_user_id uuid not null references public.profiles(id),
  assigned_user_id uuid not null references public.profiles(id),
  absent_user_ids uuid[] not null default '{}'::uuid[],
  status text not null default 'scheduled' check (status in ('scheduled', 'completed')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (team_id, practice_date)
);

create index if not exists timekeeper_assignments_team_date_idx
  on public.timekeeper_assignments (team_id, practice_date desc);

alter table public.timekeeper_cycles enable row level security;
alter table public.timekeeper_assignments enable row level security;

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
      from public.team_members tm
      join public.profiles p on p.id = tm.user_id and p.status = 'active'
      where tm.team_id = a.team_id
        and tm.user_id <> a.assigned_user_id
        and not (tm.user_id = any(a.absent_user_ids))
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

create or replace function public.get_keiko_timekeeper(
  p_team_id uuid,
  p_date date
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_assignment public.timekeeper_assignments%rowtype;
  v_cycle public.timekeeper_cycles%rowtype;
  v_previous_cycle public.timekeeper_cycles%rowtype;
  v_member_order uuid[];
  v_member_id uuid;
  v_cycle_number integer;
  v_has_practice boolean := false;
begin
  if auth.uid() is null or not exists (
    select 1
    from public.team_members tm
    join public.profiles p on p.id = tm.user_id and p.status = 'active'
    join public.teams t on t.id = tm.team_id and t.status = 'active'
    where tm.team_id = p_team_id and tm.user_id = auth.uid()
  ) then
    raise exception 'team_access_denied' using errcode = '42501';
  end if;

  if p_date is distinct from (timezone('Asia/Tokyo', now()))::date then
    raise exception 'invalid_assignment_date' using errcode = '22023';
  end if;

  perform pg_advisory_xact_lock(hashtextextended(p_team_id::text, 0));

  select * into v_assignment
  from public.timekeeper_assignments
  where team_id = p_team_id and practice_date = p_date;
  if found then
    return public.keiko_timekeeper_payload(v_assignment.id);
  end if;

  select * into v_cycle
  from public.timekeeper_cycles
  where team_id = p_team_id and status = 'active'
  for update;

  if found then
    select exists (
      select 1
      from public.practice_logs pl
      where pl.team_id = p_team_id
        and pl.practice_date < p_date
        and pl.practice_date >= coalesce((
          select min(a.practice_date)
          from public.timekeeper_assignments a
          where a.cycle_id = v_cycle.id
            and a.cycle_index = v_cycle.current_index
        ), p_date)
    ) into v_has_practice;

    if v_has_practice then
      update public.timekeeper_assignments
      set status = 'completed', updated_at = now()
      where cycle_id = v_cycle.id
        and cycle_index = v_cycle.current_index;
      v_cycle.current_index := v_cycle.current_index + 1;
      update public.timekeeper_cycles
      set current_index = v_cycle.current_index, updated_at = now()
      where id = v_cycle.id;
    end if;
  end if;

  loop
    if v_cycle.id is null or v_cycle.current_index > cardinality(v_cycle.member_order) then
      if v_cycle.id is not null then
        update public.timekeeper_cycles
        set status = 'completed', updated_at = now()
        where id = v_cycle.id;
        v_previous_cycle := v_cycle;
      end if;

      select coalesce(max(cycle_number), 0) + 1 into v_cycle_number
      from public.timekeeper_cycles
      where team_id = p_team_id;

      select coalesce(array_agg(m.user_id order by m.sort_group, m.carryover_position, random()), '{}'::uuid[])
      into v_member_order
      from (
        select tm.user_id,
          case when tm.user_id = any(coalesce(v_previous_cycle.carryover_order, '{}'::uuid[])) then 0 else 1 end as sort_group,
          coalesce(array_position(v_previous_cycle.carryover_order, tm.user_id), 2147483647) as carryover_position
        from public.team_members tm
        join public.profiles p on p.id = tm.user_id and p.status = 'active'
        where tm.team_id = p_team_id
      ) m;

      if cardinality(v_member_order) = 0 then
        raise exception 'no_active_team_members' using errcode = 'P0001';
      end if;

      insert into public.timekeeper_cycles (team_id, cycle_number, member_order)
      values (p_team_id, v_cycle_number, v_member_order)
      returning * into v_cycle;
    end if;

    v_member_id := v_cycle.member_order[v_cycle.current_index];
    exit when exists (
      select 1
      from public.team_members tm
      join public.profiles p on p.id = tm.user_id and p.status = 'active'
      where tm.team_id = p_team_id and tm.user_id = v_member_id
    );

    v_cycle.current_index := v_cycle.current_index + 1;
    update public.timekeeper_cycles
    set current_index = v_cycle.current_index, updated_at = now()
    where id = v_cycle.id;
  end loop;

  insert into public.timekeeper_assignments (
    team_id, practice_date, cycle_id, cycle_index, scheduled_user_id, assigned_user_id
  ) values (
    p_team_id, p_date, v_cycle.id, v_cycle.current_index, v_member_id, v_member_id
  ) returning * into v_assignment;

  return public.keiko_timekeeper_payload(v_assignment.id);
end;
$function$;

create or replace function public.replace_keiko_timekeeper(
  p_team_id uuid,
  p_date date
)
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
    select tm.user_id into v_candidate
    from public.team_members tm
    join public.profiles p on p.id = tm.user_id and p.status = 'active'
    where tm.team_id = p_team_id
      and not (tm.user_id = any(v_new_absent))
    order by random()
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

create or replace function public.get_keiko_home_dashboard(
  p_team_id uuid,
  p_date date
)
returns jsonb
language plpgsql
volatile
security invoker
set search_path = public
as $function$
declare
  v_summary jsonb;
begin
  v_summary := public.get_keiko_home_summary();
  return v_summary || jsonb_build_object(
    'timekeeper', public.get_keiko_timekeeper(p_team_id, p_date)
  );
end;
$function$;

revoke all on table public.timekeeper_cycles from public, anon, authenticated;
revoke all on table public.timekeeper_assignments from public, anon, authenticated;

revoke all on function public.keiko_timekeeper_payload(uuid) from public, anon, authenticated;
revoke all on function public.get_keiko_timekeeper(uuid, date) from public, anon;
revoke all on function public.replace_keiko_timekeeper(uuid, date) from public, anon;
revoke all on function public.get_keiko_home_dashboard(uuid, date) from public, anon;
grant execute on function public.get_keiko_timekeeper(uuid, date) to authenticated;
grant execute on function public.replace_keiko_timekeeper(uuid, date) to authenticated;
grant execute on function public.get_keiko_home_dashboard(uuid, date) to authenticated;

notify pgrst, 'reload schema';

commit;
