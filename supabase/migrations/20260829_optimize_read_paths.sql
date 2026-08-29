begin;

create index if not exists practice_logs_user_date_id_idx
  on public.practice_logs (user_id, practice_date desc, created_at desc, id desc);

create index if not exists practice_logs_team_date_id_idx
  on public.practice_logs (team_id, practice_date desc, created_at desc, id desc);

create or replace function public.get_keiko_session_context(p_team_id uuid default null)
returns jsonb
language sql
stable
security invoker
set search_path = public
as $function$
  with active_profile as (
    select
      p.id,
      p.display_name,
      p.user_type,
      p.legacy_user_id
    from public.profiles p
    where p.id = auth.uid()
      and p.status = 'active'
  ),
  available_teams as (
    select
      tm.user_id,
      tm.team_id,
      tm.team_role,
      t.team_name,
      t.team_type
    from public.team_members tm
    join public.teams t on t.id = tm.team_id
    join active_profile p on p.id = tm.user_id
    where t.status = 'active'
  ),
  selected_team as (
    select at.*
    from available_teams at
    order by
      case when p_team_id is not null and at.team_id = p_team_id then 0 else 1 end,
      at.team_name asc
    limit 1
  )
  select jsonb_build_object(
    'user_id', p.id,
    'display_name', p.display_name,
    'user_type', p.user_type,
    'legacy_user_id', p.legacy_user_id,
    'team_id', st.team_id,
    'team_name', st.team_name,
    'team_type', st.team_type,
    'team_role', st.team_role,
    'grade', sp.grade,
    'role_label', sp.role_label,
    'term', sp.term
  )
  from active_profile p
  join selected_team st on st.user_id = p.id
  left join public.student_profiles sp on sp.user_id = p.id;
$function$;

create or replace function public.get_keiko_home_summary()
returns jsonb
language sql
stable
security invoker
set search_path = public
as $function$
  with summary as (
    select
      count(*)::integer as total_count,
      avg(pl.condition)::numeric as average_condition
    from public.practice_logs pl
    where pl.user_id = auth.uid()
  ),
  latest as (
    select jsonb_build_object(
      'id', pl.id,
      'practice_date', pl.practice_date,
      'condition', pl.condition,
      'learning', pl.learning,
      'next_action', pl.next_action,
      'good_new', pl.good_new,
      'achievement_status', pl.achievement_status,
      'why_missed', pl.why_missed,
      'retry_plan', pl.retry_plan,
      'created_at', pl.created_at,
      'updated_at', pl.updated_at
    ) as log
    from public.practice_logs pl
    where pl.user_id = auth.uid()
    order by pl.practice_date desc, pl.created_at desc, pl.id desc
    limit 1
  )
  select jsonb_build_object(
    'total_count', s.total_count,
    'average_condition', s.average_condition,
    'latest_log', l.log
  )
  from summary s
  left join latest l on true;
$function$;

revoke all on function public.get_keiko_session_context(uuid) from public;
revoke all on function public.get_keiko_session_context(uuid) from anon;
grant execute on function public.get_keiko_session_context(uuid) to authenticated;

revoke all on function public.get_keiko_home_summary() from public;
revoke all on function public.get_keiko_home_summary() from anon;
grant execute on function public.get_keiko_home_summary() to authenticated;

commit;
