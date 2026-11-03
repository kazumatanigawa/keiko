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

test('admin OS parses and never contains privileged keys', async () => {
  const html = await read('admin.html');
  const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((match) => match[1]);
  assert.equal(scripts.length, 1);
  assert.doesNotThrow(() => new vm.Script(scripts[0]));
  assert.match(html, /KEIKO Admin OS/);
  assert.match(html, /getAdminOverview/);
  assert.match(html, /searchAdminUsers/);
  assert.match(html, /getAdminUserDetail/);
  assert.match(html, /getAdminTeamDetail/);
  assert.match(html, /data-team-id/);
  assert.match(html, /稽古日が新しい順/);
  assert.match(html, /投稿日が新しい順/);
  assert.match(html, /data-user-notes-id/);
  assert.match(html, /openUser\(button\.dataset\.userNotesId,'notes'\)/);
  assert.match(html, /private:'個人のみ',team:'チーム内',global:'グローバル'/);
  assert.doesNotMatch(html, /sb_secret_/);
  assert.doesNotMatch(html, /service_role/);
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
  for (const action of ['getAdminOverview', 'searchAdminUsers', 'getAdminUserDetail', 'getAdminTeamDetail']) {
    assert.match(source, new RegExp(`action === '${action}'`));
  }
  assert.match(source, /readBearerToken\(request\)/);
  assert.match(source, /constantTimeEquals\(text\(payload\.registrationCode\), REGISTRATION_CODE\)/);
  const config = await read('supabase/config.toml');
  assert.match(config, /\[functions\.keiko-api\][\s\S]*verify_jwt\s*=\s*false/);
});

test('admin OS database functions are operator-only and audited', async () => {
  const sql = await read('supabase/migrations/2026091001_operator_admin_os.sql');
  for (const name of ['get_keiko_admin_overview', 'get_keiko_admin_users', 'get_keiko_admin_user_detail']) {
    assert.match(sql, new RegExp(`create or replace function public\.${name}\\b`));
    assert.match(sql, new RegExp(`revoke all on function public\.${name}[\\s\\S]*?from public, anon, authenticated`));
    assert.match(sql, new RegExp(`grant execute on function public\.${name}[\\s\\S]*?to service_role`));
  }
  assert.match(sql, /app_role = 'operator'/);
  assert.match(sql, /'view_user_detail', 'user'/);
  assert.match(sql, /least\(coalesce\(p_limit, 100\), 200\)/);
});

test('admin team detail is scoped, paged, sorted, and audited', async () => {
  const sql = await read('supabase/migrations/2026091101_operator_team_detail.sql');
  assert.match(sql, /create or replace function public\.get_keiko_admin_team_detail\b/);
  assert.match(sql, /app_role = 'operator'/);
  assert.match(sql, /where pl\.team_id = p_team_id[\s\S]*order by pl\.practice_date desc, pl\.created_at desc/);
  assert.match(sql, /where n\.team_id = p_team_id[\s\S]*order by n\.created_at desc, n\.id desc/);
  assert.match(sql, /least\(coalesce\(p_limit, 50\), 100\)/);
  assert.match(sql, /limit v_limit offset v_offset/);
  assert.match(sql, /'view_team_detail',[\s\S]*'team'/);
  assert.match(sql, /revoke all on function public\.get_keiko_admin_team_detail[\s\S]*from public, anon, authenticated/);
  assert.match(sql, /grant execute on function public\.get_keiko_admin_team_detail[\s\S]*to service_role/);
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
  const practiceDayPosition = html.indexOf('id="practiceDayCard"');
  const helpPosition = html.indexOf('id="helpSection"');
  assert.match(html, /id="practiceDayCount"/);
  assert.match(html, /稽古<span class="practice-day-number"/);
  assert.ok(practiceDayPosition > 0 && helpPosition > practiceDayPosition);
  assert.doesNotMatch(html, /id="homeRecordBtn"/);
  assert.doesNotMatch(html, /id="homeLogsBtn"/);
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
  const learningPosition = html.indexOf('id="learningHubSection"');
  const settingsPosition = html.indexOf('id="personalSettings"');
  assert.ok(helpPosition > 0 && learningPosition > helpPosition && settingsPosition > learningPosition);
  assert.match(html, /id="formTeamPicker"/);
  assert.match(html, /picker\.classList\.toggle\('show', teams\.length > 1\)/);
  assert.match(html, /const preferred = teams\.find\(team => team\.teamId === getTeamId\(\)\)/);
  assert.match(html, /p_team_id: getSelectedLogTeamId\(\)/);
  assert.match(html, /get_keiko_log_target_summary/);
  assert.match(html, /openMembershipEditor\('join'\)/);
  assert.match(html, /openMembershipEditor\('transfer'\)/);
  assert.match(html, /graduateFromTeams\(\)/);
  assert.doesNotMatch(html, /登録コードはアプリ内には表示されません/);
  assert.match(html, /<details class="help-card" id="helpSection">/);
  assert.match(html, /class="help-summary"/);
  assert.doesNotMatch(html, /<details class="help-card" id="helpSection" open/);
  assert.match(html, /<details class="settings-card" id="personalSettings">/);
  assert.match(html, /class="settings-summary"/);
  assert.match(html, /class="settings-summary-sub">チームの移籍・掛け持ち</);
  assert.doesNotMatch(html, /<details class="settings-card" id="personalSettings" open/);
  assert.match(html, /KEIKO運営または参加先チームの責任者/);
  assert.match(html, /class="nav-item active" aria-current="page"/);
  assert.match(html, /\.nav-item\.active\{background:var\(--nav-active-soft\)/);
  assert.doesNotMatch(html, /個人利用ではチームノートは表示されません/);
});

test('home context only switches between regular teams while notes keep every audience', async () => {
  const html = await read('index.html');
  const switcherStart = html.indexOf('function renderViewContextSwitcher()');
  const switcherEnd = html.indexOf('async function switchViewContext', switcherStart);
  const switcher = html.slice(switcherStart, switcherEnd);
  assert.match(html, /稽古するチームを選ぶ/);
  assert.match(html, /掛け持ちしている場合、今回どのチームで稽古するかを選びます/);
  assert.match(switcher, /const regularTeams = getRegularTeams\(\)/);
  assert.match(switcher, /regularTeams\.map/);
  assert.doesNotMatch(switcher, /getMemberships\(\)\.map/);
  assert.doesNotMatch(switcher, /グローバル|自分/);
  assert.match(html, /function renderNoteScopeButtons\(\)[\s\S]*?自分だけ[\s\S]*?getRegularTeams\(\)\.map[\s\S]*?グローバル/);
});

test('registration always refreshes the active team directory', async () => {
  const html = await read('index.html');
  assert.match(html, /id="teamRefreshButton"[\s\S]*?loadTeamOptions\(true\)/);
  assert.match(html, /function initializeAuthUi\(\)[\s\S]*?loadTeamOptions\(true\)/);
  assert.match(html, /if \(tab === 'register'\) loadTeamOptions\(true\)/);
  assert.match(html, /if \(teamState\.loading\)[\s\S]*?teamState\.refreshRequested = true/);
  assert.match(html, /function setTeamDirectoryLoading\(isLoading\)[\s\S]*?select\.disabled = isLoading \|\| isLaunchMode/);
  assert.match(html, /KEIKO管理者が登録した正式なチーム名/);
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
  assert.match(html, /const regularTeams = getRegularTeams\(\);[\s\S]*panel\.style\.display = regularTeams\.length > 1/);
  assert.doesNotMatch(html, /switchViewContext\('global'\)/);
  assert.doesNotMatch(html, /id="globalHomeCard"/);
  assert.match(html, /ノート画面上部の「グローバル」/);
  const homeStart = html.indexOf('<div class="screen" id="screen-home">');
  const notesStart = html.indexOf('<div class="screen" id="screen-notes">');
  const libraryStart = html.indexOf('<div class="screen" id="screen-library">');
  const homeMarkup = html.slice(homeStart, notesStart);
  const notesMarkup = html.slice(notesStart, libraryStart);
  assert.doesNotMatch(homeMarkup, /id="globalParticipationCard"/);
  assert.match(notesMarkup, /id="globalParticipationCard"/);
  assert.match(html, /card\.classList\.toggle\('show', canUseGlobal\)/);
  assert.match(html, /function updateNotes\(\) \{\s*renderGlobalParticipation\(\)/);
  assert.doesNotMatch(html, /id="operatorSettings"/);
  assert.match(html, /reportContent\('note'/);

  assert.match(edge, /🚨KEIKO OS からの通報🚨/);
  assert.match(edge, /kokongakumeiza@gmail\.com/);
  assert.match(edge, /RESEND_API_KEY/);
});

test('home hub exposes one help entry, vertical resources, announcements, and interactive requests', async () => {
  const html = await read('index.html');
  const edge = await read('supabase/functions/keiko-api/index.ts');
  const mediaMigration = await read('supabase/migrations/2026110102_keiko_media_links.sql');
  const requestMigration = await read('supabase/migrations/2026110103_interactive_service_requests.sql');
  const woodChoicesMigration = await read('supabase/migrations/2026110301_bachi_wood_choices.sql');
  const resourceCleanupMigration = await read('supabase/migrations/2026110302_home_resource_cleanup.sql');
  assert.match(html, /id="announcementCard"/);
  assert.doesNotMatch(html, /id="tutorialHomeCard"/);
  assert.match(html, /id="helpSection"/);
  assert.match(html, /class="help-primary-action"[^>]*>使い方を見る</);
  assert.match(html, /id="learningResourceRow"/);
  assert.match(html, /class="home-library-action"[^>]*>ライブラリを見る →</);
  assert.match(html, /renderResourceRow\('learningResourceRow', selectHomeLearningResources\(resources\)\)/);
  assert.match(html, /renderResourceRow\('connectionResourceRow', selectHomeConnectionResources\(resources\)\)/);
  assert.match(html, /const discord = resources\.find\(isDiscordStudyResource\) \|\| fallbacks\.find\(isDiscordStudyResource\)/);
  assert.match(html, /String\(item\?\.resourceId \|\| ''\) === 'discord-study'/);
  assert.match(html, /return \[findResource\(isGakumeizaScoreResource\)\]\.filter\(Boolean\)/);
  assert.match(html, /if \(isGakumeizaYoutubeResource\(item\)\) return '';/);
  assert.match(html, /id="libraryResourceList"/);
  assert.match(html, /id="libraryYoutubeChannelList"/);
  assert.match(html, /古今樂明座 Official YouTube/);
  assert.doesNotMatch(html, /id="libraryResourceCarousel"/);
  assert.doesNotMatch(html, /指導スライドをチェック/);
  assert.doesNotMatch(html, /drive\.google\.com\/drive\/folders\/1a9iaxd2OKBywSuRZvPib1gyGciL22HpM/);
  assert.doesNotMatch(html, /class="brand-support-card"/);
  assert.match(html, /\.resource-row,\.library-resource-list\{display:grid;grid-template-columns:1fr/);
  assert.match(html, /gakumeiza-score/);
  assert.match(html, /古今樂明座の楽譜・音源/);
  assert.match(html, /id="coachingDialog"/);
  assert.match(html, /古今樂明座に指導を依頼する/);
  assert.doesNotMatch(html, /谷川に指導を依頼する/);
  assert.match(html, /id="bachiDialog"/);
  assert.match(html, /keiko:\/\/bachi/);
  assert.match(html, /scheduleBachiEstimate/);
  assert.match(html, /scheduleCoachingEstimate/);
  assert.match(html, /const BACHI_WOOD_CHOICES = \[/);
  assert.match(html, /\{ woodCode: 'hinoki', displayName: 'ヒノキ' \}/);
  assert.match(html, /\{ woodCode: 'hou', displayName: 'ホオ' \}/);
  assert.match(html, /\{ woodCode: 'maple', displayName: 'カエデ' \}/);
  assert.match(html, /requestCatalogState\.woods = BACHI_WOOD_CHOICES/);
  assert.match(html, /role="button" tabindex="0" data-resource-id/);
  assert.match(html, /card\.addEventListener\('click'/);
  assert.match(html, /card\.addEventListener\('keydown'/);
  assert.doesNotMatch(html, /<button class="resource-link"/);
  assert.match(html, /\.resource-link\.is-disabled\{[^}]*font-size:14px/);
  assert.match(html, /function formatRequestDate/);
  assert.match(html, /function requestEmailStatusText/);
  assert.match(html, /id="coachingDuration"/);
  for (const removedField of ['coachingPhone', 'coachingDate', 'coachingParticipants', 'coachingStart', 'coachingEnd', 'coachingNotes']) {
    assert.doesNotMatch(html, new RegExp(`id="${removedField}"`));
  }
  assert.match(html, /\[60, 90, 120, 150, 180\]/);
  assert.match(html, /durationMinutes/);
  assert.match(html, /estimateCoachingRequest/);
  assert.match(html, /createCoachingRequest/);
  assert.match(html, /safeDestinationUrl/);
  assert.match(html, /supported by 古今樂明座<\/div><\/a>/);
  assert.doesNotMatch(html, /supported by 古今樂明座　公式サイトへ/);
  assert.match(html, /https:\/\/m\.youtube\.com\/watch\?v=awLvvl1AvhU&list=PLmk028LkL7RF9l2NIJBfL07O0b1-pBrRJ/);
  assert.match(html, /gakumeiza-youtube/);
  assert.match(html, /古今樂明座 Official YouTubeチャンネル/);
  assert.match(html, /id="youtubeVideoList"/);
  assert.match(html, /FALLBACK_YOUTUBE_VIDEOS/);
  assert.match(html, /getYoutubeVideos/);
  assert.match(html, /youtube-video-card/);
  assert.match(mediaMigration, /where source_key = 'gakumeiza-radio'/);
  assert.match(mediaMigration, /m\.youtube\.com\/watch\?v=awLvvl1AvhU&list=PLmk028LkL7RF9l2NIJBfL07O0b1-pBrRJ/);
  assert.match(edge, /action === 'estimateCoachingRequest'/);
  assert.match(edge, /action === 'createCoachingRequest'/);
  assert.match(edge, /action === 'getRequestCatalogs'/);
  assert.match(edge, /action === 'estimateBachiOrder'/);
  assert.match(edge, /action === 'createBachiOrder'/);
  assert.match(edge, /action === 'getYoutubeVideos'/);
  assert.match(edge, /youtube\.com\/feeds\/videos\.xml\?channel_id=/);
  assert.match(edge, /YOUTUBE_FEED_CACHE_MS/);
  assert.match(edge, /coaching_rate_options/);
  assert.match(edge, /bachi_wood_price_options/);
  assert.match(edge, /sendBachiEmails/);
  assert.match(edge, /sendCoachingEmails/);
  assert.match(edge, /customer_email_status/);
  assert.match(edge, /GOOGLE_MAPS_ROUTES_API_KEY/);
  assert.match(edge, /routes\.googleapis\.com\/directions\/v2:computeRoutes/);
  assert.match(edge, /routes\.travelAdvisory\.transitFare/);
  assert.match(edge, /KEIKO_COACHING_ORIGIN/);
  assert.match(edge, /requiredInteger\(payload\.durationMinutes, '指導時間'/);
  assert.doesNotMatch(edge, /requiredDate\(payload\.requestedDate/);
  assert.doesNotMatch(html, /GOOGLE_MAPS_DIRECTIONS_API_KEY/);
  for (const table of ['bachi_wood_price_options', 'bachi_quantity_price_tiers', 'bachi_order_settings', 'bachi_order_requests', 'coaching_rate_options']) {
    assert.match(requestMigration, new RegExp(`create table if not exists public\\.${table}\\b`));
  }
  assert.match(requestMigration, /\('hinoki', 'ヒノキ', 1800, 42, 24, 1, 10\)/);
  assert.match(requestMigration, /\('hou', 'ホオ', 1600, 42, 24, 1, 20\)/);
  assert.match(requestMigration, /\('maple', 'カエデ', 1600, 42, 24, 1, 30\)/);
  assert.match(woodChoicesMigration, /where wood_code not in \('hinoki', 'hou', 'maple'\)/);
  assert.match(woodChoicesMigration, /\('hinoki', 'ヒノキ', 1800/);
  assert.match(woodChoicesMigration, /\('hou', 'ホオ', 1600/);
  assert.match(woodChoicesMigration, /\('maple', 'カエデ', 1600/);
  assert.match(woodChoicesMigration, /on conflict \(wood_code\) do update/);
  assert.match(resourceCleanupMigration, /set status = 'archived'/);
  assert.match(resourceCleanupMigration, /where source_key = 'gakumeiza-radio'/);
  assert.match(resourceCleanupMigration, /set title = '古今樂明座に指導を依頼する'/);
  assert.match(requestMigration, /\('half_day', '半日（4時間）', 24000/);
  assert.match(requestMigration, /alter column requested_date drop not null/);
  assert.match(requestMigration, /alter column start_time drop not null/);
  assert.match(requestMigration, /alter column end_time drop not null/);
  assert.match(requestMigration, /destination_url = 'keiko:\/\/bachi'/);
});

test('team admin SaaS migration scopes staff to assigned team-visible data', async () => {
  const sql = await read('supabase/migrations/2026110101_home_hub_team_admin_saas.sql');
  for (const table of [
    'home_resources',
    'announcements',
    'announcement_reads',
    'coaching_requests',
    'team_subscriptions',
    'team_entitlements',
    'team_admin_audit_events',
  ]) assert.match(sql, new RegExp(`create table if not exists public\\.${table}\\b`));

  for (const name of [
    'get_keiko_home_hub',
    'mark_keiko_announcement_read',
    'get_keiko_admin_access',
    'get_keiko_scoped_admin_overview',
    'get_keiko_scoped_admin_users',
    'get_keiko_scoped_admin_user_detail',
    'get_keiko_scoped_admin_team_detail',
    'set_keiko_team_staff_role',
  ]) assert.match(sql, new RegExp(`create or replace function public\\.${name}\\b`));

  assert.match(sql, /team_role in \('owner_admin', 'owner', 'manager', 'coach', 'advisor'\)/);
  assert.match(sql, /pl\.visibility = 'team'/);
  assert.match(sql, /n\.visibility = 'team'/);
  assert.match(sql, /team_forbidden/);
  assert.match(sql, /team_admin_audit_events/);
  assert.match(sql, /a\.scope = 'all'/);
  assert.match(sql, /grant execute on function public\.get_keiko_home_hub\(uuid\) to authenticated/);
  assert.match(sql, /grant execute on function public\.get_keiko_scoped_admin_team_detail[\s\S]*to service_role/);
});

test('Admin OS accepts scoped team administrators without exposing operator tools', async () => {
  const html = await read('admin.html');
  const edge = await read('supabase/functions/keiko-api/index.ts');
  assert.match(html, /getAdminAccess/);
  assert.match(html, /TEAM ADMIN MODE/);
  assert.match(html, /operator-only-nav/);
  assert.match(html, /setTeamStaffRole/);
  assert.match(html, /saveAdminAnnouncement/);
  assert.match(html, /全KEIKOユーザー/);
  assert.match(html, /scope:'all'/);
  assert.match(html, /全てのKEIKOユーザーのホームへ公開します/);
  assert.match(html, /チーム管理者からは配信できません/);
  assert.doesNotMatch(html, /id="announcementScope"/);
  assert.doesNotMatch(html, /id="announcementTeam"/);
  const app = await read('index.html');
  assert.match(app, /loadHomeHub\(\{ force: true \}\)/);
  assert.match(app, /visibilitychange[\s\S]*?homeScreen\?\.classList\.contains\('active'\)/);
  assert.doesNotMatch(html, /session\.appRole!=='operator'/);
  assert.match(edge, /get_keiko_scoped_admin_overview/);
  assert.match(edge, /get_keiko_scoped_admin_users/);
  assert.match(edge, /get_keiko_scoped_admin_user_detail/);
  assert.match(edge, /get_keiko_scoped_admin_team_detail/);
});
