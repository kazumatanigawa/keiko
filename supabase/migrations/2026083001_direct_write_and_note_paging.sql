begin;

create unique index if not exists practice_logs_keiko_request_uidx
  on public.practice_logs (user_id, source_system, legacy_source_row_id)
  where source_system = 'keiko_app' and legacy_source_row_id is not null;

create unique index if not exists team_notes_keiko_request_uidx
  on public.team_notes (author_user_id, legacy_note_id)
  where legacy_note_id is not null;

create unique index if not exists team_note_comments_keiko_request_uidx
  on public.team_note_comments (author_user_id, legacy_comment_id)
  where legacy_comment_id is not null;

create index if not exists team_notes_team_status_created_idx
  on public.team_notes (team_id, status, created_at desc, id desc);

create index if not exists team_note_comments_note_status_created_idx
  on public.team_note_comments (note_id, status, created_at asc, id asc);

create index if not exists practice_logs_team_good_news_idx
  on public.practice_logs (team_id, practice_date desc, created_at desc, id desc)
  where visibility = 'team' and good_new is not null and good_new <> '';

create or replace function public.save_keiko_log(
  p_team_id uuid,
  p_request_id text,
  p_date date,
  p_condition integer,
  p_learning text,
  p_next_action text,
  p_good_new text,
  p_achievement_status text default null,
  p_why_missed text default '',
  p_retry_plan text default ''
)
returns jsonb
language plpgsql
security invoker
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_source_row_id text;
  v_profile public.profiles%rowtype;
  v_grade text := '';
  v_term text := '';
  v_row public.practice_logs%rowtype;
  v_duplicate boolean := false;
begin
  if v_user_id is null then raise exception 'auth_required'; end if;
  if p_request_id !~ '^[A-Za-z0-9_-]{16,80}$' then raise exception 'invalid_request_id'; end if;
  if p_date is null then raise exception 'invalid_date'; end if;
  if p_condition is null or p_condition < 1 or p_condition > 5 then raise exception 'invalid_condition'; end if;
  if p_learning is null or nullif(btrim(p_learning), '') is null or length(p_learning) > 1000 then raise exception 'invalid_learning'; end if;
  if p_next_action is null or nullif(btrim(p_next_action), '') is null or length(p_next_action) > 1000 then raise exception 'invalid_next_action'; end if;
  if p_good_new is null or nullif(btrim(p_good_new), '') is null or length(p_good_new) > 1000 then raise exception 'invalid_good_new'; end if;
  if coalesce(p_achievement_status, '') not in ('', 'done', 'pending') then raise exception 'invalid_achievement_status'; end if;
  if coalesce(p_achievement_status, '') = 'pending'
    and (nullif(btrim(p_why_missed), '') is null or nullif(btrim(p_retry_plan), '') is null)
  then raise exception 'missing_retry_details'; end if;

  select p.* into v_profile
  from public.profiles p
  where p.id = v_user_id and p.status = 'active';
  if not found then raise exception 'inactive_profile'; end if;

  if not exists (
    select 1 from public.team_members tm
    join public.teams t on t.id = tm.team_id and t.status = 'active'
    where tm.user_id = v_user_id and tm.team_id = p_team_id
  ) then raise exception 'team_forbidden'; end if;

  select
    coalesce((select sp.grade from public.student_profiles sp where sp.user_id = v_user_id), ''),
    coalesce((select sp.term from public.student_profiles sp where sp.user_id = v_user_id), '')
  into v_grade, v_term;

  v_source_row_id := 'app:' || p_request_id;
  select * into v_row
  from public.practice_logs pl
  where pl.user_id = v_user_id
    and pl.source_system = 'keiko_app'
    and pl.legacy_source_row_id = v_source_row_id
  limit 1;

  if found then
    v_duplicate := true;
  else
    begin
      insert into public.practice_logs (
        user_id, team_id, practice_date, condition, learning, next_action,
        good_new, memo, visibility, source_system, legacy_source_row_id,
        achievement_status, why_missed, retry_plan, display_name_snapshot,
        grade_snapshot, term_snapshot
      ) values (
        v_user_id, p_team_id, p_date, p_condition, btrim(p_learning),
        btrim(p_next_action), btrim(p_good_new), '', 'team', 'keiko_app',
        v_source_row_id, nullif(p_achievement_status, ''),
        case when p_achievement_status = 'pending' then btrim(p_why_missed) else '' end,
        case when p_achievement_status = 'pending' then btrim(p_retry_plan) else '' end,
        v_profile.display_name, v_grade, v_term
      ) returning * into v_row;
    exception when unique_violation then
      v_duplicate := true;
      select * into v_row
      from public.practice_logs pl
      where pl.user_id = v_user_id
        and pl.source_system = 'keiko_app'
        and pl.legacy_source_row_id = v_source_row_id
      limit 1;
    end;
  end if;

  return jsonb_build_object(
    'status', 'ok',
    'id', v_row.id,
    'duplicate', v_duplicate,
    'log', jsonb_build_object(
      'id', v_row.id, 'date', v_row.practice_date, 'cond', v_row.condition,
      'learning', coalesce(v_row.learning, ''), 'next', coalesce(v_row.next_action, ''),
      'goodNew', coalesce(v_row.good_new, ''),
      'achievementStatus', coalesce(v_row.achievement_status, ''),
      'whyMissed', coalesce(v_row.why_missed, ''), 'retryPlan', coalesce(v_row.retry_plan, ''),
      'createdAt', v_row.created_at, 'updatedAt', v_row.updated_at
    )
  );
end;
$function$;

create or replace function public.get_keiko_notes_page(
  p_team_id uuid,
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
  ),
  ordered_rows as (
    select n.*
    from public.team_notes n
    where n.team_id = p_team_id
      and n.status = 'active'
      and exists (
        select 1 from public.team_members tm
        where tm.team_id = p_team_id and tm.user_id = auth.uid()
      )
      and (
        p_cursor_created_at is null or p_cursor_id is null
        or (n.created_at, n.id) < (p_cursor_created_at, p_cursor_id)
      )
    order by n.created_at desc, n.id desc
    limit (select page_limit + 1 from params)
  ),
  page_rows as (
    select * from ordered_rows
    order by created_at desc, id desc
    limit (select page_limit from params)
  ),
  last_row as (
    select created_at, id from page_rows
    order by created_at asc, id asc limit 1
  )
  select jsonb_build_object(
    'notes', coalesce((
      select jsonb_agg(jsonb_build_object(
        'noteId', n.id,
        'authorUserId', n.author_user_id,
        'authorName', coalesce(nullif(n.author_name_snapshot, ''), 'メンバー'),
        'title', n.title,
        'body', n.body,
        'createdAt', n.created_at,
        'updatedAt', n.updated_at,
        'commentCount', (select count(*) from public.team_note_comments c where c.note_id = n.id and c.status = 'active'),
        'comments', '[]'::jsonb,
        'commentsLoaded', false
      ) order by n.created_at desc, n.id desc)
      from page_rows n
    ), '[]'::jsonb),
    'hasMore', (select count(*) > (select page_limit from params) from ordered_rows),
    'nextCursor', case
      when (select count(*) > (select page_limit from params) from ordered_rows)
      then (select jsonb_build_object('createdAt', l.created_at, 'id', l.id) from last_row l)
      else null
    end
  );
$function$;

create or replace function public.get_keiko_note_comments(
  p_team_id uuid,
  p_note_id uuid,
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
  ),
  ordered_rows as (
    select c.*
    from public.team_note_comments c
    where c.team_id = p_team_id and c.note_id = p_note_id and c.status = 'active'
      and exists (
        select 1 from public.team_members tm
        where tm.team_id = p_team_id and tm.user_id = auth.uid()
      )
      and exists (
        select 1 from public.team_notes n
        where n.id = p_note_id and n.team_id = p_team_id and n.status = 'active'
      )
      and (
        p_cursor_created_at is null or p_cursor_id is null
        or (c.created_at, c.id) > (p_cursor_created_at, p_cursor_id)
      )
    order by c.created_at asc, c.id asc
    limit (select page_limit + 1 from params)
  ),
  page_rows as (
    select * from ordered_rows order by created_at asc, id asc
    limit (select page_limit from params)
  ),
  last_row as (
    select created_at, id from page_rows order by created_at desc, id desc limit 1
  )
  select jsonb_build_object(
    'comments', coalesce((
      select jsonb_agg(jsonb_build_object(
        'commentId', c.id, 'noteId', c.note_id,
        'authorUserId', c.author_user_id,
        'authorName', coalesce(nullif(c.author_name_snapshot, ''), 'メンバー'),
        'body', c.body, 'createdAt', c.created_at, 'updatedAt', c.updated_at,
        'isEdited', c.updated_at is distinct from c.created_at,
        'canEdit', c.author_user_id = auth.uid(),
        'canDelete', c.author_user_id = auth.uid()
      ) order by c.created_at asc, c.id asc)
      from page_rows c
    ), '[]'::jsonb),
    'hasMore', (select count(*) > (select page_limit from params) from ordered_rows),
    'nextCursor', case
      when (select count(*) > (select page_limit from params) from ordered_rows)
      then (select jsonb_build_object('createdAt', l.created_at, 'id', l.id) from last_row l)
      else null
    end
  );
$function$;

create or replace function public.save_keiko_note(
  p_team_id uuid,
  p_request_id text,
  p_title text,
  p_body text
)
returns jsonb
language plpgsql
security invoker
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_author_name text;
  v_legacy_id text;
  v_row public.team_notes%rowtype;
  v_duplicate boolean := false;
begin
  if v_user_id is null then raise exception 'auth_required'; end if;
  if p_request_id !~ '^[A-Za-z0-9_-]{16,80}$' then raise exception 'invalid_request_id'; end if;
  if p_title is null or nullif(btrim(p_title), '') is null or length(p_title) > 120 then raise exception 'invalid_title'; end if;
  if p_body is null or nullif(btrim(p_body), '') is null or length(p_body) > 5000 then raise exception 'invalid_body'; end if;
  select p.display_name into v_author_name from public.profiles p where p.id = v_user_id and p.status = 'active';
  if not found then raise exception 'inactive_profile'; end if;
  if not exists (select 1 from public.team_members tm where tm.team_id = p_team_id and tm.user_id = v_user_id) then raise exception 'team_forbidden'; end if;
  v_legacy_id := 'app:' || p_request_id;
  select * into v_row from public.team_notes n where n.author_user_id = v_user_id and n.legacy_note_id = v_legacy_id limit 1;
  if found then
    v_duplicate := true;
  else
    begin
      insert into public.team_notes (legacy_note_id, team_id, author_user_id, author_name_snapshot, title, body, status)
      values (v_legacy_id, p_team_id, v_user_id, v_author_name, btrim(p_title), btrim(p_body), 'active')
      returning * into v_row;
    exception when unique_violation then
      v_duplicate := true;
      select * into v_row from public.team_notes n where n.author_user_id = v_user_id and n.legacy_note_id = v_legacy_id limit 1;
    end;
  end if;
  return jsonb_build_object('status', 'ok', 'duplicate', v_duplicate, 'note', jsonb_build_object(
    'noteId', v_row.id, 'authorUserId', v_row.author_user_id,
    'authorName', coalesce(nullif(v_row.author_name_snapshot, ''), 'メンバー'),
    'title', v_row.title, 'body', v_row.body, 'createdAt', v_row.created_at,
    'updatedAt', v_row.updated_at, 'commentCount', 0, 'comments', '[]'::jsonb,
    'commentsLoaded', false
  ));
end;
$function$;

create or replace function public.add_keiko_note_comment(
  p_team_id uuid,
  p_note_id uuid,
  p_request_id text,
  p_body text
)
returns jsonb
language plpgsql
security invoker
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_author_name text;
  v_legacy_id text;
  v_row public.team_note_comments%rowtype;
  v_duplicate boolean := false;
begin
  if v_user_id is null then raise exception 'auth_required'; end if;
  if p_request_id !~ '^[A-Za-z0-9_-]{16,80}$' then raise exception 'invalid_request_id'; end if;
  if p_body is null or nullif(btrim(p_body), '') is null or length(p_body) > 5000 then raise exception 'invalid_body'; end if;
  select p.display_name into v_author_name from public.profiles p where p.id = v_user_id and p.status = 'active';
  if not found then raise exception 'inactive_profile'; end if;
  if not exists (
    select 1 from public.team_members tm
    join public.team_notes n on n.team_id = tm.team_id and n.id = p_note_id and n.status = 'active'
    where tm.team_id = p_team_id and tm.user_id = v_user_id
  ) then raise exception 'note_forbidden'; end if;
  v_legacy_id := 'app:' || p_request_id;
  select * into v_row from public.team_note_comments c where c.author_user_id = v_user_id and c.legacy_comment_id = v_legacy_id limit 1;
  if found then
    v_duplicate := true;
  else
    begin
      insert into public.team_note_comments (legacy_comment_id, note_id, team_id, author_user_id, author_name_snapshot, body, status)
      values (v_legacy_id, p_note_id, p_team_id, v_user_id, v_author_name, btrim(p_body), 'active')
      returning * into v_row;
    exception when unique_violation then
      v_duplicate := true;
      select * into v_row from public.team_note_comments c where c.author_user_id = v_user_id and c.legacy_comment_id = v_legacy_id limit 1;
    end;
  end if;
  return jsonb_build_object('status', 'ok', 'duplicate', v_duplicate, 'comment', jsonb_build_object(
    'commentId', v_row.id, 'noteId', v_row.note_id, 'authorUserId', v_row.author_user_id,
    'authorName', coalesce(nullif(v_row.author_name_snapshot, ''), 'メンバー'),
    'body', v_row.body, 'createdAt', v_row.created_at, 'updatedAt', v_row.updated_at,
    'isEdited', false, 'canEdit', true, 'canDelete', true
  ));
end;
$function$;

create or replace function public.update_keiko_note_comment(
  p_team_id uuid,
  p_comment_id uuid,
  p_body text
)
returns jsonb
language plpgsql
security invoker
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_row public.team_note_comments%rowtype;
begin
  if v_user_id is null then raise exception 'auth_required'; end if;
  if p_body is null or nullif(btrim(p_body), '') is null or length(p_body) > 5000 then raise exception 'invalid_body'; end if;
  update public.team_note_comments c
  set body = btrim(p_body), updated_at = now()
  where c.id = p_comment_id and c.team_id = p_team_id
    and c.author_user_id = v_user_id and c.status = 'active'
  returning * into v_row;
  if not found then raise exception 'comment_forbidden'; end if;
  return jsonb_build_object('status', 'ok', 'comment', jsonb_build_object(
    'commentId', v_row.id, 'noteId', v_row.note_id, 'authorUserId', v_row.author_user_id,
    'authorName', coalesce(nullif(v_row.author_name_snapshot, ''), 'メンバー'),
    'body', v_row.body, 'createdAt', v_row.created_at, 'updatedAt', v_row.updated_at,
    'isEdited', true, 'canEdit', true, 'canDelete', true
  ));
end;
$function$;

create or replace function public.delete_keiko_note_comment(
  p_team_id uuid,
  p_comment_id uuid
)
returns jsonb
language plpgsql
security invoker
set search_path = public
as $function$
declare
  v_user_id uuid := auth.uid();
  v_note_id uuid;
begin
  if v_user_id is null then raise exception 'auth_required'; end if;
  update public.team_note_comments c
  set status = 'deleted', updated_at = now()
  where c.id = p_comment_id and c.team_id = p_team_id
    and c.author_user_id = v_user_id and c.status = 'active'
  returning c.note_id into v_note_id;
  if not found then raise exception 'comment_forbidden'; end if;
  return jsonb_build_object('status', 'ok', 'commentId', p_comment_id, 'noteId', v_note_id);
end;
$function$;

revoke all on function public.save_keiko_log(uuid, text, date, integer, text, text, text, text, text, text) from public, anon;
revoke all on function public.get_keiko_notes_page(uuid, timestamptz, uuid, integer) from public, anon;
revoke all on function public.get_keiko_note_comments(uuid, uuid, timestamptz, uuid, integer) from public, anon;
revoke all on function public.save_keiko_note(uuid, text, text, text) from public, anon;
revoke all on function public.add_keiko_note_comment(uuid, uuid, text, text) from public, anon;
revoke all on function public.update_keiko_note_comment(uuid, uuid, text) from public, anon;
revoke all on function public.delete_keiko_note_comment(uuid, uuid) from public, anon;

grant execute on function public.save_keiko_log(uuid, text, date, integer, text, text, text, text, text, text) to authenticated;
grant execute on function public.get_keiko_notes_page(uuid, timestamptz, uuid, integer) to authenticated;
grant execute on function public.get_keiko_note_comments(uuid, uuid, timestamptz, uuid, integer) to authenticated;
grant execute on function public.save_keiko_note(uuid, text, text, text) to authenticated;
grant execute on function public.add_keiko_note_comment(uuid, uuid, text, text) to authenticated;
grant execute on function public.update_keiko_note_comment(uuid, uuid, text) to authenticated;
grant execute on function public.delete_keiko_note_comment(uuid, uuid) to authenticated;

notify pgrst, 'reload schema';

commit;
