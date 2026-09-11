begin;

create or replace function public.get_keiko_admin_team_detail(
  p_actor_user_id uuid,
  p_team_id uuid,
  p_section text default 'logs',
  p_limit integer default 50,
  p_offset integer default 0
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = public
as $function$
declare
  v_team public.teams%rowtype;
  v_section text := lower(btrim(coalesce(p_section, 'logs')));
  v_limit integer := greatest(1, least(coalesce(p_limit, 50), 100));
  v_offset integer := greatest(0, coalesce(p_offset, 0));
  v_total bigint := 0;
  v_items jsonb := '[]'::jsonb;
begin
  if auth.role() <> 'service_role' or not exists (
    select 1 from public.profiles p
    where p.id = p_actor_user_id and p.status = 'active' and p.app_role = 'operator'
  ) then raise exception 'operator_required' using errcode = '42501'; end if;

  if v_section not in ('logs', 'notes', 'members') then
    raise exception 'invalid_section';
  end if;

  select * into v_team from public.teams where id = p_team_id;
  if not found then raise exception 'team_not_found'; end if;

  if v_section = 'logs' then
    select count(*) into v_total from public.practice_logs pl where pl.team_id = p_team_id;
    select coalesce(jsonb_agg(jsonb_build_object(
      'id', x.id,
      'userId', x.user_id,
      'displayName', x.display_name,
      'practiceDate', x.practice_date,
      'condition', x.condition,
      'learning', x.learning,
      'nextAction', x.next_action,
      'goodNew', x.good_new,
      'memo', coalesce(x.memo, ''),
      'visibility', x.visibility,
      'createdAt', x.created_at,
      'updatedAt', x.updated_at
    ) order by x.practice_date desc, x.created_at desc), '[]'::jsonb)
    into v_items
    from (
      select pl.*, coalesce(p.display_name, '名前未設定') as display_name
      from public.practice_logs pl
      left join public.profiles p on p.id = pl.user_id
      where pl.team_id = p_team_id
      order by pl.practice_date desc, pl.created_at desc
      limit v_limit offset v_offset
    ) x;
  elsif v_section = 'notes' then
    select count(*) into v_total from public.team_notes n where n.team_id = p_team_id;
    select coalesce(jsonb_agg(jsonb_build_object(
      'id', x.id,
      'authorUserId', x.author_user_id,
      'authorName', x.author_name,
      'title', x.title,
      'body', x.body,
      'visibility', x.visibility,
      'status', x.status,
      'commentCount', x.comment_count,
      'createdAt', x.created_at,
      'updatedAt', x.updated_at
    ) order by x.created_at desc, x.id desc), '[]'::jsonb)
    into v_items
    from (
      select
        n.*,
        coalesce(p.display_name, n.author_name_snapshot, '名前未設定') as author_name,
        (select count(*) from public.team_note_comments c where c.note_id = n.id and c.status = 'active') as comment_count
      from public.team_notes n
      left join public.profiles p on p.id = n.author_user_id
      where n.team_id = p_team_id
      order by n.created_at desc, n.id desc
      limit v_limit offset v_offset
    ) x;
  else
    select count(*) into v_total from public.team_members tm where tm.team_id = p_team_id;
    select coalesce(jsonb_agg(jsonb_build_object(
      'userId', x.user_id,
      'displayName', x.display_name,
      'userType', x.user_type,
      'status', x.status,
      'teamRole', x.team_role,
      'grade', x.grade,
      'term', x.term,
      'roleLabel', x.role_label,
      'practiceLogCount', x.practice_log_count,
      'noteCount', x.note_count,
      'joinedAt', x.joined_at
    ) order by (x.status = 'active') desc, x.display_name), '[]'::jsonb)
    into v_items
    from (
      select
        tm.user_id,
        tm.team_role,
        tm.joined_at,
        coalesce(p.display_name, '名前未設定') as display_name,
        p.user_type,
        p.status,
        coalesce(sp.grade, '') as grade,
        coalesce(sp.term, '') as term,
        coalesce(sp.role_label, '') as role_label,
        (select count(*) from public.practice_logs pl where pl.team_id = p_team_id and pl.user_id = tm.user_id) as practice_log_count,
        (select count(*) from public.team_notes n where n.team_id = p_team_id and n.author_user_id = tm.user_id and n.status = 'active') as note_count
      from public.team_members tm
      join public.profiles p on p.id = tm.user_id
      left join public.student_profiles sp on sp.user_id = tm.user_id
      where tm.team_id = p_team_id
      order by (p.status = 'active') desc, p.display_name
      limit v_limit offset v_offset
    ) x;
  end if;

  insert into public.keiko_operator_events (actor_user_id, action, target_type, target_id, details)
  values (
    p_actor_user_id,
    'view_team_detail',
    'team',
    p_team_id,
    jsonb_build_object('source', 'admin_os', 'section', v_section, 'offset', v_offset, 'limit', v_limit)
  );

  return jsonb_build_object(
    'team', jsonb_build_object(
      'teamId', v_team.id,
      'teamName', v_team.team_name,
      'teamType', v_team.team_type,
      'category', coalesce(v_team.category, ''),
      'status', v_team.status,
      'createdAt', v_team.created_at
    ),
    'summary', jsonb_build_object(
      'memberCount', (select count(*) from public.team_members tm where tm.team_id = p_team_id),
      'logCount', (select count(*) from public.practice_logs pl where pl.team_id = p_team_id),
      'noteCount', (select count(*) from public.team_notes n where n.team_id = p_team_id),
      'commentCount', (select count(*) from public.team_note_comments c where c.team_id = p_team_id and c.status = 'active')
    ),
    'section', v_section,
    'total', v_total,
    'limit', v_limit,
    'offset', v_offset,
    'items', v_items
  );
end;
$function$;

revoke all on function public.get_keiko_admin_team_detail(uuid, uuid, text, integer, integer) from public, anon, authenticated;
grant execute on function public.get_keiko_admin_team_detail(uuid, uuid, text, integer, integer) to service_role;

notify pgrst, 'reload schema';

commit;
