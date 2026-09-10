begin;

create or replace function public.get_keiko_admin_overview(p_actor_user_id uuid)
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $function$
begin
  if auth.role() <> 'service_role' or not exists (
    select 1 from public.profiles p
    where p.id = p_actor_user_id and p.status = 'active' and p.app_role = 'operator'
  ) then raise exception 'operator_required' using errcode = '42501'; end if;

  return jsonb_build_object(
    'summary', jsonb_build_object(
      'activeUsers', (select count(*) from public.profiles where status = 'active'),
      'activeTeams', (select count(*) from public.teams where status = 'active' and coalesce(category, '') <> 'personal'),
      'practiceLogs', (select count(*) from public.practice_logs),
      'activeNotes', (select count(*) from public.team_notes where status = 'active'),
      'logsLast7Days', (select count(*) from public.practice_logs where created_at >= now() - interval '7 days')
    ),
    'teams', coalesce((
      select jsonb_agg(jsonb_build_object(
        'teamId', t.id,
        'teamName', t.team_name,
        'teamType', t.team_type,
        'memberCount', (select count(*) from public.team_members tm where tm.team_id = t.id),
        'logCount', (select count(*) from public.practice_logs pl where pl.team_id = t.id),
        'noteCount', (select count(*) from public.team_notes n where n.team_id = t.id and n.status = 'active'),
        'latestActivityAt', greatest(
          coalesce((select max(pl.created_at) from public.practice_logs pl where pl.team_id = t.id), '-infinity'::timestamptz),
          coalesce((select max(n.created_at) from public.team_notes n where n.team_id = t.id), '-infinity'::timestamptz)
        )
      ) order by t.team_name)
      from public.teams t
      where t.status = 'active' and coalesce(t.category, '') <> 'personal'
    ), '[]'::jsonb)
  );
end;
$function$;

create or replace function public.get_keiko_admin_users(
  p_actor_user_id uuid,
  p_query text default '',
  p_team_id uuid default null,
  p_status text default '',
  p_limit integer default 50,
  p_offset integer default 0
)
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $function$
declare
  v_limit integer := greatest(1, least(coalesce(p_limit, 50), 100));
  v_offset integer := greatest(0, coalesce(p_offset, 0));
  v_query text := btrim(coalesce(p_query, ''));
begin
  if auth.role() <> 'service_role' or not exists (
    select 1 from public.profiles p
    where p.id = p_actor_user_id and p.status = 'active' and p.app_role = 'operator'
  ) then raise exception 'operator_required' using errcode = '42501'; end if;

  return jsonb_build_object(
    'total', (
      select count(*)
      from public.profiles p
      where (nullif(p_status, '') is null or p.status = p_status)
        and (v_query = '' or p.display_name ilike '%' || v_query || '%' or p.login_id ilike '%' || v_query || '%')
        and (p_team_id is null or exists (
          select 1 from public.team_members tm where tm.user_id = p.id and tm.team_id = p_team_id
        ))
    ),
    'users', coalesce((
      select jsonb_agg(jsonb_build_object(
        'userId', u.id,
        'displayName', u.display_name,
        'lastNameKana', coalesce(u.last_name_kana, ''),
        'firstNameKana', coalesce(u.first_name_kana, ''),
        'userType', u.user_type,
        'appRole', u.app_role,
        'status', u.status,
        'createdAt', u.created_at,
        'practiceLogCount', u.practice_log_count,
        'noteCount', u.note_count,
        'commentCount', u.comment_count,
        'latestActivityAt', u.latest_activity_at,
        'teams', u.teams
      ) order by u.latest_activity_at desc nulls last, u.display_name)
      from (
        select
          p.id, p.display_name, p.last_name_kana, p.first_name_kana,
          p.user_type, p.app_role, p.status, p.created_at,
          (select count(*) from public.practice_logs pl where pl.user_id = p.id) as practice_log_count,
          (select count(*) from public.team_notes n where n.author_user_id = p.id and n.status = 'active') as note_count,
          (select count(*) from public.team_note_comments c where c.author_user_id = p.id and c.status = 'active') as comment_count,
          greatest(
            coalesce((select max(pl.created_at) from public.practice_logs pl where pl.user_id = p.id), '-infinity'::timestamptz),
            coalesce((select max(n.created_at) from public.team_notes n where n.author_user_id = p.id), '-infinity'::timestamptz),
            coalesce((select max(c.created_at) from public.team_note_comments c where c.author_user_id = p.id), '-infinity'::timestamptz)
          ) as latest_activity_at,
          coalesce((
            select jsonb_agg(jsonb_build_object(
              'teamId', t.id, 'teamName', t.team_name, 'teamType', t.team_type,
              'teamRole', tm.team_role, 'isPersonal', coalesce(t.category, '') = 'personal'
            ) order by (coalesce(t.category, '') = 'personal'), t.team_name)
            from public.team_members tm
            join public.teams t on t.id = tm.team_id
            where tm.user_id = p.id and t.status = 'active'
          ), '[]'::jsonb) as teams
        from public.profiles p
        where (nullif(p_status, '') is null or p.status = p_status)
          and (v_query = '' or p.display_name ilike '%' || v_query || '%' or p.login_id ilike '%' || v_query || '%')
          and (p_team_id is null or exists (
            select 1 from public.team_members tm where tm.user_id = p.id and tm.team_id = p_team_id
          ))
        order by latest_activity_at desc nulls last, p.display_name
        limit v_limit offset v_offset
      ) u
    ), '[]'::jsonb)
  );
end;
$function$;

create or replace function public.get_keiko_admin_user_detail(
  p_actor_user_id uuid,
  p_user_id uuid,
  p_limit integer default 100
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_profile public.profiles%rowtype;
  v_limit integer := greatest(1, least(coalesce(p_limit, 100), 200));
  v_result jsonb;
begin
  if auth.role() <> 'service_role' or not exists (
    select 1 from public.profiles p
    where p.id = p_actor_user_id and p.status = 'active' and p.app_role = 'operator'
  ) then raise exception 'operator_required' using errcode = '42501'; end if;

  select * into v_profile from public.profiles where id = p_user_id;
  if not found then raise exception 'user_not_found'; end if;

  v_result := jsonb_build_object(
    'profile', jsonb_build_object(
      'userId', v_profile.id,
      'displayName', v_profile.display_name,
      'name', coalesce(v_profile.name, ''),
      'lastName', coalesce(v_profile.last_name, ''),
      'firstName', coalesce(v_profile.first_name, ''),
      'lastNameKana', coalesce(v_profile.last_name_kana, ''),
      'firstNameKana', coalesce(v_profile.first_name_kana, ''),
      'userType', v_profile.user_type,
      'appRole', v_profile.app_role,
      'status', v_profile.status,
      'globalParticipationEnabled', v_profile.global_participation_enabled,
      'createdAt', v_profile.created_at,
      'updatedAt', v_profile.updated_at
    ),
    'studentProfile', coalesce((
      select jsonb_build_object(
        'schoolName', coalesce(sp.school_name, ''),
        'grade', coalesce(sp.grade, ''),
        'term', coalesce(sp.term, ''),
        'roleLabel', coalesce(sp.role_label, '')
      ) from public.student_profiles sp where sp.user_id = p_user_id
    ), '{}'::jsonb),
    'generalProfile', coalesce((
      select jsonb_build_object('category', coalesce(gp.category, ''), 'bio', coalesce(gp.bio, ''))
      from public.general_profiles gp where gp.user_id = p_user_id
    ), '{}'::jsonb),
    'teams', coalesce((
      select jsonb_agg(jsonb_build_object(
        'teamId', t.id, 'teamName', t.team_name, 'teamType', t.team_type,
        'teamRole', tm.team_role, 'isPersonal', coalesce(t.category, '') = 'personal'
      ) order by (coalesce(t.category, '') = 'personal'), t.team_name)
      from public.team_members tm join public.teams t on t.id = tm.team_id
      where tm.user_id = p_user_id
    ), '[]'::jsonb),
    'logs', coalesce((
      select jsonb_agg(jsonb_build_object(
        'id', x.id, 'teamId', x.team_id, 'teamName', coalesce(t.team_name, '個人'),
        'practiceDate', x.practice_date, 'condition', x.condition,
        'learning', x.learning, 'nextAction', x.next_action, 'goodNew', x.good_new,
        'memo', coalesce(x.memo, ''), 'visibility', x.visibility, 'createdAt', x.created_at
      ) order by x.practice_date desc, x.created_at desc)
      from (select * from public.practice_logs where user_id = p_user_id order by practice_date desc, created_at desc limit v_limit) x
      left join public.teams t on t.id = x.team_id
    ), '[]'::jsonb),
    'notes', coalesce((
      select jsonb_agg(jsonb_build_object(
        'id', x.id, 'teamId', x.team_id, 'teamName', coalesce(t.team_name, '個人'),
        'title', x.title, 'body', x.body, 'visibility', x.visibility,
        'status', x.status, 'createdAt', x.created_at, 'updatedAt', x.updated_at
      ) order by x.created_at desc)
      from (select * from public.team_notes where author_user_id = p_user_id order by created_at desc limit v_limit) x
      left join public.teams t on t.id = x.team_id
    ), '[]'::jsonb),
    'comments', coalesce((
      select jsonb_agg(jsonb_build_object(
        'id', x.id, 'noteId', x.note_id, 'teamId', x.team_id,
        'teamName', coalesce(t.team_name, '個人'), 'body', x.body,
        'status', x.status, 'createdAt', x.created_at, 'updatedAt', x.updated_at
      ) order by x.created_at desc)
      from (select * from public.team_note_comments where author_user_id = p_user_id order by created_at desc limit v_limit) x
      left join public.teams t on t.id = x.team_id
    ), '[]'::jsonb)
  );

  insert into public.keiko_operator_events (actor_user_id, action, target_type, target_id, details)
  values (p_actor_user_id, 'view_user_detail', 'user', p_user_id, jsonb_build_object('source', 'admin_os'));

  return v_result;
end;
$function$;

revoke all on function public.get_keiko_admin_overview(uuid) from public, anon, authenticated;
grant execute on function public.get_keiko_admin_overview(uuid) to service_role;
revoke all on function public.get_keiko_admin_users(uuid, text, uuid, text, integer, integer) from public, anon, authenticated;
grant execute on function public.get_keiko_admin_users(uuid, text, uuid, text, integer, integer) to service_role;
revoke all on function public.get_keiko_admin_user_detail(uuid, uuid, integer) from public, anon, authenticated;
grant execute on function public.get_keiko_admin_user_detail(uuid, uuid, integer) to service_role;

notify pgrst, 'reload schema';

commit;
