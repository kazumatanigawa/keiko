import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import vm from 'node:vm';

const root = new URL('../', import.meta.url);
const read = (path) => readFile(new URL(path, root), 'utf8');

test('inline application JavaScript parses', async () => {
  const html = await read('index.html');
  const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((match) => match[1]);
  assert.equal(scripts.length, 1);
  assert.doesNotThrow(() => new vm.Script(scripts[0]));
});

test('browser bundle contains no retired backend or privileged keys', async () => {
  const html = await read('index.html');
  assert.doesNotMatch(html, /script\.google\.com\/macros/);
  assert.doesNotMatch(html, /sb_secret_/);
  assert.doesNotMatch(html, /service_role/);
  assert.doesNotMatch(html, /AIza[0-9A-Za-z_-]{20,}/);
  assert.match(html, /sb_publishable_/);
});

test('direct write and paging migration exposes authenticated RPCs', async () => {
  const sql = await read('supabase/migrations/2026083001_direct_write_and_note_paging.sql');
  for (const name of [
    'save_keiko_log',
    'get_keiko_notes_page',
    'get_keiko_note_comments',
    'save_keiko_note',
    'add_keiko_note_comment',
    'update_keiko_note_comment',
    'delete_keiko_note_comment',
  ]) {
    assert.match(sql, new RegExp(`create or replace function public\\.${name}\\b`));
    assert.match(sql, new RegExp(`grant execute on function public\\.${name}\\b[\\s\\S]*?to authenticated`));
  }
  assert.match(sql, /least\(coalesce\(p_limit, 20\), 50\)/);
  assert.match(sql, /practice_logs_keiko_request_uidx/);
});

test('Edge API has a deliberately small public action surface', async () => {
  const source = await read('supabase/functions/keiko-api/index.ts');
  for (const action of ['health', 'getTeams', 'login', 'register']) {
    assert.match(source, new RegExp(`action === '${action}'`));
  }
  for (const retired of ['saveLog', 'getLogs', 'saveNote', 'refreshSession']) {
    assert.doesNotMatch(source, new RegExp(`action === '${retired}'`));
  }
  assert.match(source, /KEIKO_AUTH_PEPPER/);
  assert.match(source, /KEIKO_REGISTRATION_CODE/);
  assert.match(source, /action === 'manageMembership'/);
  for (const action of ['getOperatorDashboard', 'updateTeamSettings', 'reportContent', 'moderateReport']) {
    assert.match(source, new RegExp(`action === '${action}'`));
  }
  assert.match(source, /readBearerToken\(request\)/);
  assert.match(source, /constantTimeEquals\(text\(payload\.registrationCode\), REGISTRATION_CODE\)/);
  const config = await read('supabase/config.toml');
  assert.match(config, /\[functions\.keiko-api\][\s\S]*verify_jwt\s*=\s*false/);
});

test('all application collections use bounded reads', async () => {
  const html = await read('index.html');
  assert.match(html, /get_keiko_logs_page_v2/);
  assert.match(html, /get_keiko_notes_feed/);
  assert.match(html, /get_keiko_note_comments_v2/);
  assert.doesNotMatch(html, /get_keiko_notes['"]/);
  assert.match(html, /p_limit:\s*20/);
  const cleanup = await read('supabase/migrations/2026083002_remove_legacy_notes_rpc.sql');
  assert.match(cleanup, /drop function if exists public\.get_keiko_notes\(uuid, integer\)/);
});

test('timekeeper rotation is server-managed and available on the home dashboard', async () => {
  const sql = await read('supabase/migrations/2026090101_timekeeper_rotation.sql');
  const contextSql = await read('supabase/migrations/2026090502_context_global_moderation.sql');
  const html = await read('index.html');

  assert.match(sql, /create table if not exists public\.timekeeper_cycles/);
  assert.match(sql, /create table if not exists public\.timekeeper_assignments/);
  assert.match(sql, /alter table public\.timekeeper_cycles enable row level security/);
  assert.match(sql, /array_agg\(m\.user_id order by m\.sort_group, m\.carryover_position, random\(\)\)/);
  assert.match(sql, /unique \(team_id, practice_date\)/);
  assert.match(sql, /create or replace function public\.replace_keiko_timekeeper/);
  assert.match(sql, /carryover_order = v_new_absent/);
  assert.match(contextSql, /from unnest\(v_cycle\.member_order\)[\s\S]*join public\.team_members/);
  assert.doesNotMatch(contextSql, /select tm\.user_id into v_candidate[\s\S]*order by random\(\)/);
  assert.match(sql, /create or replace function public\.get_keiko_home_dashboard/);
  assert.match(sql, /grant execute on function public\.get_keiko_home_dashboard\(uuid, date\) to authenticated/);

  assert.match(html, /id="timekeeperName"/);
  assert.match(html, /get_keiko_home_dashboard/);
  assert.match(html, /replace_keiko_timekeeper/);
  assert.match(html, /交代中\.\.\./);
});

test('home emphasizes practice day count without redundant condition summary', async () => {
  const html = await read('index.html');
  assert.match(html, /id="practiceDayCount"/);
  assert.match(html, /稽古<span class="practice-day-number"/);
  assert.doesNotMatch(html, /id="totalCount"/);
  assert.doesNotMatch(html, /id="statAvgCond"/);
  assert.doesNotMatch(html, /平均コンディション/);
});

test('multi-team membership changes are audited and restricted to the server', async () => {
  const sql = await read('supabase/migrations/2026090501_multi_team_memberships.sql');
  for (const name of ['keiko_membership_context_for_user', 'get_keiko_session_context', 'manage_keiko_membership']) {
    assert.match(sql, new RegExp(`create or replace function public\\.${name}\\b`));
  }
  assert.match(sql, /create table if not exists public\.team_membership_events/);
  assert.match(sql, /alter table public\.team_membership_events enable row level security/);
  assert.match(sql, /coalesce\(category, ''\) = 'personal'/);
  assert.match(sql, /action in \('join', 'transfer_in', 'transfer_out', 'graduate'\)/);
  assert.match(sql, /auth\.role\(\) <> 'service_role'/);
  assert.match(sql, /grant execute on function public\.manage_keiko_membership[\s\S]*to service_role/);
  assert.doesNotMatch(sql, /grant execute on function public\.manage_keiko_membership[\s\S]*to authenticated/);
});

test('home settings and log form support multiple teams without cluttering single-team users', async () => {
  const html = await read('index.html');
  const helpPosition = html.indexOf('id="helpSection"');
  const settingsPosition = html.indexOf('id="personalSettings"');
  assert.ok(helpPosition > 0 && settingsPosition > helpPosition);
  assert.match(html, /id="formTeamPicker"/);
  assert.match(html, /picker\.classList\.toggle\('show', teams\.length > 1\)/);
  assert.match(html, /p_team_id: getSelectedLogTeamId\(\)/);
  assert.match(html, /get_keiko_log_target_summary/);
  assert.match(html, /openMembershipEditor\('join'\)/);
  assert.match(html, /openMembershipEditor\('transfer'\)/);
  assert.match(html, /graduateFromTeams\(\)/);
  assert.match(html, /登録コードはアプリ内には表示されません/);
  assert.match(html, /KEIKO運営または参加先チームの責任者/);
  assert.match(html, /class="nav-item active" aria-current="page"/);
  assert.match(html, /\.nav-item\.active\{background:var\(--nav-active-soft\)/);
  assert.doesNotMatch(html, /個人利用ではチームノートは表示されません/);
});

test('global notes are opt-in, explicitly scoped, and operator moderated', async () => {
  const sql = await read('supabase/migrations/2026090502_context_global_moderation.sql');
  const html = await read('index.html');
  const edge = await read('supabase/functions/keiko-api/index.ts');

  for (const name of [
    'set_keiko_global_participation',
    'get_keiko_notes_feed',
    'save_keiko_note_v2',
    'add_keiko_note_comment_v2',
    'create_keiko_content_report',
    'update_keiko_team_global_settings',
    'moderate_keiko_report',
  ]) assert.match(sql, new RegExp(`create or replace function public\\.${name}\\b`));

  assert.match(sql, /app_role in \('player', 'operator'\)/);
  assert.match(sql, /visibility in \('private', 'team', 'global'\)/);
  assert.match(sql, /team_type = 'student'[\s\S]*school_global_forbidden/);
  assert.match(sql, /update public\.team_notes n[\s\S]*set visibility = 'team'/);
  assert.match(sql, /global_participation_enabled/);
  assert.match(sql, /create table if not exists public\.content_reports/);
  assert.match(sql, /unique \(reporter_user_id, content_type, content_id\)/);

  assert.match(html, /id="viewContextButtons"/);
  assert.match(html, /id="viewContextPanel"[\s\S]*id="viewContextCurrent"[\s\S]*<div class="content">/);
  assert.match(html, /id="globalParticipationBtn"/);
  assert.match(html, /id="noteScopeButtons"/);
  assert.match(html, /id="noteDestinationBanner"/);
  assert.match(html, /現在の表示先がグローバルでも、外部には公開されません/);
  assert.match(html, /グローバルに公開されます/);
  assert.match(html, /presentation\.confirmation/);
  assert.match(html, /selectNoteScope\('private'\)/);
  assert.match(html, /selectNoteScope\('global'\)/);
  assert.match(html, /id="operatorSettings"/);
  assert.match(html, /reportContent\('note'/);

  assert.match(edge, /🚨KEIKO OS からの通報🚨/);
  assert.match(edge, /kokongakumeiza@gmail\.com/);
  assert.match(edge, /RESEND_API_KEY/);
});
