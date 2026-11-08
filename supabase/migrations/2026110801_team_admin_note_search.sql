begin;

create index if not exists practice_logs_team_visibility_date_idx
  on public.practice_logs (team_id, visibility, practice_date desc, created_at desc);

create index if not exists team_notes_team_visibility_status_created_idx
  on public.team_notes (team_id, visibility, status, created_at desc);

create or replace function public.search_keiko_team_logs(
  p_team_id uuid,
  p_query text default '',
  p_user_id uuid default null,
  p_date_from date default null,
  p_date_to date default null,
  p_condition integer default null,
  p_limit integer default 20,
  p_offset integer default 0
)
returns jsonb
language plpgsql
volatile
security definer
set search_path = ''
as $function$
declare
  v_actor_user_id uuid := auth.uid();
  v_query text := btrim(coalesce(p_query, ''));
  v_limit integer := greatest(1, least(coalesce(p_limit, 20), 50));
  v_offset integer := greatest(0, coalesce(p_offset, 0));
  v_result jsonb;
begin
  if v_actor_user_id is null then
    raise exception 'auth_required' using errcode = '42501';
  end if;
  if not public.keiko_is_team_staff(v_actor_user_id, p_team_id) then
    raise exception 'team_forbidden' using errcode = '42501';
  end if;
  if p_condition is not null and (p_condition < 1 or p_condition > 5) then
    raise exception 'invalid_condition' using errcode = '22023';
  end if;
  if p_date_from is not null and p_date_to is not null and p_date_from > p_date_to then
    raise exception 'invalid_date_range' using errcode = '22023';
  end if;

  with filtered as (
    select
      pl.*,
      coalesce(nullif(p.display_name, ''), nullif(pl.display_name_snapshot, ''), '名前未設定') as display_name
    from public.practice_logs pl
    left join public.profiles p on p.id = pl.user_id
    where pl.team_id = p_team_id
      and pl.visibility = 'team'
      and (p_user_id is null or pl.user_id = p_user_id)
      and (p_date_from is null or pl.practice_date >= p_date_from)
      and (p_date_to is null or pl.practice_date <= p_date_to)
      and (p_condition is null or pl.condition = p_condition)
      and (
        v_query = ''
        or coalesce(p.display_name, '') ilike '%' || v_query || '%'
        or coalesce(pl.display_name_snapshot, '') ilike '%' || v_query || '%'
        or coalesce(pl.learning, '') ilike '%' || v_query || '%'
        or coalesce(pl.next_action, '') ilike '%' || v_query || '%'
        or coalesce(pl.good_new, '') ilike '%' || v_query || '%'
        or coalesce(pl.memo, '') ilike '%' || v_query || '%'
      )
  ), paged as (
    select * from filtered
    order by practice_date desc, created_at desc, id desc
    limit v_limit offset v_offset
  )
  select jsonb_build_object(
    'total', (select count(*) from filtered),
    'logs', coalesce((
      select jsonb_agg(jsonb_build_object(
        'id', x.id,
        'userId', x.user_id,
        'displayName', x.display_name,
        'date', x.practice_date,
        'cond', x.condition,
        'learning', x.learning,
        'next', x.next_action,
        'goodNew', x.good_new,
        'memo', coalesce(x.memo, ''),
        'achievementStatus', coalesce(x.achievement_status, ''),
        'whyMissed', coalesce(x.why_missed, ''),
        'retryPlan', coalesce(x.retry_plan, ''),
        'grade', coalesce(x.grade_snapshot, ''),
        'term', coalesce(x.term_snapshot, ''),
        'visibility', x.visibility,
        'createdAt', x.created_at,
        'updatedAt', x.updated_at
      ) order by x.practice_date desc, x.created_at desc, x.id desc)
      from paged x
    ), '[]'::jsonb),
    'members', coalesce((
      select jsonb_agg(jsonb_build_object(
        'userId', tm.user_id,
        'displayName', coalesce(nullif(member.display_name, ''), '名前未設定')
      ) order by member.display_name)
      from public.team_members tm
      join public.profiles member on member.id = tm.user_id and member.status = 'active'
      where tm.team_id = p_team_id
    ), '[]'::jsonb),
    'limit', v_limit,
    'offset', v_offset
  ) into v_result;

  insert into public.team_admin_audit_events (
    actor_user_id, team_id, action, target_type, target_id, metadata
  ) values (
    v_actor_user_id,
    p_team_id,
    'search_team_logs',
    'team',
    p_team_id::text,
    jsonb_build_object(
      'queryLength', length(v_query),
      'hasUserFilter', p_user_id is not null,
      'dateFrom', p_date_from,
      'dateTo', p_date_to,
      'condition', p_condition,
      'offset', v_offset,
      'limit', v_limit
    )
  );

  return v_result;
end;
$function$;

create or replace function public.search_keiko_notes(
  p_scope text,
  p_team_id uuid default null,
  p_query text default '',
  p_cursor_created_at timestamptz default null,
  p_cursor_id uuid default null,
  p_limit integer default 20
)
returns jsonb
language plpgsql
stable
security definer
set search_path = ''
as $function$
declare
  v_user_id uuid := auth.uid();
  v_scope text := lower(btrim(coalesce(p_scope, '')));
  v_query text := btrim(coalesce(p_query, ''));
  v_limit integer := greatest(1, least(coalesce(p_limit, 20), 50));
  v_result jsonb;
begin
  if v_user_id is null then
    raise exception 'auth_required' using errcode = '42501';
  end if;
  if v_scope not in ('private', 'team', 'global') then
    raise exception 'invalid_scope' using errcode = '22023';
  end if;
  if v_scope = 'team' and not exists (
    select 1 from public.team_members tm
    where tm.team_id = p_team_id and tm.user_id = v_user_id
  ) then
    raise exception 'team_forbidden' using errcode = '42501';
  end if;
  if v_scope = 'global' and not exists (
    select 1 from public.profiles p
    where p.id = v_user_id
      and p.status = 'active'
      and p.user_type = 'general'
      and p.global_participation_enabled
  ) then
    raise exception 'global_participation_required' using errcode = '42501';
  end if;

  with ordered_rows as (
    select n.*
    from public.team_notes n
    join public.teams t on t.id = n.team_id and t.status = 'active'
    where n.status = 'active'
      and (
        (v_scope = 'private' and n.visibility = 'private' and n.author_user_id = v_user_id)
        or (v_scope = 'team' and n.team_id = p_team_id and n.visibility in ('team', 'global'))
        or (v_scope = 'global' and n.visibility = 'global'
          and (coalesce(t.category, '') = 'personal'
            or (t.team_type <> 'student' and t.global_notes_enabled and t.global_scope = 'global')))
      )
      and (
        v_query = ''
        or coalesce(n.title, '') ilike '%' || v_query || '%'
        or coalesce(n.body, '') ilike '%' || v_query || '%'
        or coalesce(n.author_name_snapshot, '') ilike '%' || v_query || '%'
        or coalesce(n.author_team_name_snapshot, '') ilike '%' || v_query || '%'
      )
      and (
        p_cursor_created_at is null or p_cursor_id is null
        or (n.created_at, n.id) < (p_cursor_created_at, p_cursor_id)
      )
    order by n.created_at desc, n.id desc
    limit v_limit + 1
  ), page_rows as (
    select * from ordered_rows
    order by created_at desc, id desc
    limit v_limit
  ), last_row as (
    select created_at, id from page_rows
    order by created_at asc, id asc
    limit 1
  )
  select jsonb_build_object(
    'notes', coalesce((
      select jsonb_agg(jsonb_build_object(
        'noteId', n.id,
        'teamId', n.team_id,
        'visibility', n.visibility,
        'authorUserId', n.author_user_id,
        'authorName', coalesce(nullif(n.author_name_snapshot, ''), 'メンバー'),
        'authorTeamName', coalesce(nullif(n.author_team_name_snapshot, ''), '個人'),
        'title', n.title,
        'body', n.body,
        'createdAt', n.created_at,
        'updatedAt', n.updated_at,
        'commentCount', (select count(*) from public.team_note_comments c where c.note_id = n.id and c.status = 'active'),
        'comments', '[]'::jsonb,
        'commentsLoaded', false,
        'canReport', n.author_user_id <> v_user_id
      ) order by n.created_at desc, n.id desc)
      from page_rows n
    ), '[]'::jsonb),
    'hasMore', (select count(*) > v_limit from ordered_rows),
    'nextCursor', case
      when (select count(*) > v_limit from ordered_rows)
      then (select jsonb_build_object('createdAt', l.created_at, 'id', l.id) from last_row l)
      else null
    end,
    'query', v_query
  ) into v_result;

  return v_result;
end;
$function$;

revoke all on function public.search_keiko_team_logs(uuid, text, uuid, date, date, integer, integer, integer)
  from public, anon;
grant execute on function public.search_keiko_team_logs(uuid, text, uuid, date, date, integer, integer, integer)
  to authenticated;

revoke all on function public.search_keiko_notes(text, uuid, text, timestamptz, uuid, integer)
  from public, anon;
grant execute on function public.search_keiko_notes(text, uuid, text, timestamptz, uuid, integer)
  to authenticated;

notify pgrst, 'reload schema';

commit;
