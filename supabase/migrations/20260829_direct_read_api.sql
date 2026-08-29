begin;

create or replace function public.get_keiko_logs_page(
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
  ),
  ordered_rows as (
    select
      pl.id,
      pl.practice_date,
      pl.condition,
      pl.learning,
      pl.next_action,
      pl.good_new,
      pl.achievement_status,
      pl.why_missed,
      pl.retry_plan,
      pl.created_at,
      pl.updated_at
    from public.practice_logs pl
    where pl.user_id = auth.uid()
      and (
        p_cursor_date is null
        or p_cursor_created_at is null
        or p_cursor_id is null
        or (pl.practice_date, pl.created_at, pl.id)
          < (p_cursor_date, p_cursor_created_at, p_cursor_id)
      )
    order by pl.practice_date desc, pl.created_at desc, pl.id desc
    limit (select page_limit + 1 from params)
  ),
  page_rows as (
    select *
    from ordered_rows
    order by practice_date desc, created_at desc, id desc
    limit (select page_limit from params)
  ),
  last_row as (
    select pr.practice_date, pr.created_at, pr.id
    from page_rows pr
    order by pr.practice_date asc, pr.created_at asc, pr.id asc
    limit 1
  ),
  page_meta as (
    select count(*) > (select page_limit from params) as has_more
    from ordered_rows
  )
  select jsonb_build_object(
    'logs', coalesce((
      select jsonb_agg(
        jsonb_build_object(
          'id', pr.id,
          'date', pr.practice_date,
          'cond', pr.condition,
          'learning', coalesce(pr.learning, ''),
          'next', coalesce(pr.next_action, ''),
          'goodNew', coalesce(pr.good_new, ''),
          'achievementStatus', coalesce(pr.achievement_status, ''),
          'whyMissed', coalesce(pr.why_missed, ''),
          'retryPlan', coalesce(pr.retry_plan, ''),
          'createdAt', pr.created_at,
          'updatedAt', pr.updated_at
        )
        order by pr.practice_date desc, pr.created_at desc, pr.id desc
      )
      from page_rows pr
    ), '[]'::jsonb),
    'hasMore', (select has_more from page_meta),
    'nextCursor', case
      when (select has_more from page_meta) then (
        select jsonb_build_object(
          'date', lr.practice_date,
          'createdAt', lr.created_at,
          'id', lr.id
        )
        from last_row lr
      )
      else null
    end
  );
$function$;

create or replace function public.get_keiko_good_news(
  p_team_id uuid,
  p_limit integer default 50
)
returns jsonb
language sql
stable
security invoker
set search_path = public
as $function$
  select jsonb_build_object(
    'items', coalesce(jsonb_agg(
      jsonb_build_object(
        'id', rows.id,
        'userId', rows.user_id,
        'name', coalesce(nullif(rows.display_name_snapshot, ''), 'メンバー'),
        'goodNew', rows.good_new,
        'date', rows.practice_date,
        'createdAt', rows.created_at
      )
      order by rows.practice_date desc, rows.created_at desc, rows.id desc
    ), '[]'::jsonb)
  )
  from (
    select
      pl.id,
      pl.user_id,
      pl.practice_date,
      pl.good_new,
      pl.display_name_snapshot,
      pl.created_at
    from public.practice_logs pl
    where pl.team_id = p_team_id
      and pl.visibility = 'team'
      and nullif(btrim(pl.good_new), '') is not null
    order by pl.practice_date desc, pl.created_at desc, pl.id desc
    limit greatest(1, least(coalesce(p_limit, 50), 100))
  ) rows;
$function$;

create or replace function public.get_keiko_notes(
  p_team_id uuid,
  p_limit integer default 100
)
returns jsonb
language sql
stable
security invoker
set search_path = public
as $function$
  select jsonb_build_object(
    'notes', coalesce(jsonb_agg(
      jsonb_build_object(
        'noteId', notes.id,
        'authorUserId', notes.author_user_id,
        'authorName', coalesce(nullif(notes.author_name_snapshot, ''), 'メンバー'),
        'title', notes.title,
        'body', notes.body,
        'createdAt', notes.created_at,
        'updatedAt', notes.updated_at,
        'comments', coalesce((
          select jsonb_agg(
            jsonb_build_object(
              'commentId', comments.id,
              'noteId', comments.note_id,
              'authorUserId', comments.author_user_id,
              'authorName', coalesce(nullif(comments.author_name_snapshot, ''), 'メンバー'),
              'body', comments.body,
              'createdAt', comments.created_at,
              'updatedAt', comments.updated_at,
              'isEdited', comments.updated_at is distinct from comments.created_at,
              'canEdit', comments.author_user_id = auth.uid(),
              'canDelete', comments.author_user_id = auth.uid()
            )
            order by comments.created_at asc, comments.id asc
          )
          from public.team_note_comments comments
          where comments.note_id = notes.id
            and comments.team_id = p_team_id
            and comments.status = 'active'
        ), '[]'::jsonb)
      )
      order by notes.created_at desc, notes.id desc
    ), '[]'::jsonb)
  )
  from (
    select
      n.id,
      n.author_user_id,
      n.author_name_snapshot,
      n.title,
      n.body,
      n.created_at,
      n.updated_at
    from public.team_notes n
    where n.team_id = p_team_id
      and n.status = 'active'
    order by n.created_at desc, n.id desc
    limit greatest(1, least(coalesce(p_limit, 100), 100))
  ) notes;
$function$;

revoke all on function public.get_keiko_logs_page(date, timestamptz, uuid, integer) from public;
revoke all on function public.get_keiko_logs_page(date, timestamptz, uuid, integer) from anon;
grant execute on function public.get_keiko_logs_page(date, timestamptz, uuid, integer) to authenticated;

revoke all on function public.get_keiko_good_news(uuid, integer) from public;
revoke all on function public.get_keiko_good_news(uuid, integer) from anon;
grant execute on function public.get_keiko_good_news(uuid, integer) to authenticated;

revoke all on function public.get_keiko_notes(uuid, integer) from public;
revoke all on function public.get_keiko_notes(uuid, integer) from anon;
grant execute on function public.get_keiko_notes(uuid, integer) to authenticated;

notify pgrst, 'reload schema';

commit;
