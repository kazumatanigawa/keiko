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
  const config = await read('supabase/config.toml');
  assert.match(config, /\[functions\.keiko-api\][\s\S]*verify_jwt\s*=\s*false/);
});

test('all application collections use bounded reads', async () => {
  const html = await read('index.html');
  assert.match(html, /get_keiko_logs_page/);
  assert.match(html, /get_keiko_notes_page/);
  assert.match(html, /get_keiko_note_comments/);
  assert.doesNotMatch(html, /get_keiko_notes['"]/);
  assert.match(html, /p_limit:\s*20/);
  const cleanup = await read('supabase/migrations/2026083002_remove_legacy_notes_rpc.sql');
  assert.match(cleanup, /drop function if exists public\.get_keiko_notes\(uuid, integer\)/);
});

test('timekeeper rotation is server-managed and available on the home dashboard', async () => {
  const sql = await read('supabase/migrations/2026090101_timekeeper_rotation.sql');
  const html = await read('index.html');

  assert.match(sql, /create table if not exists public\.timekeeper_cycles/);
  assert.match(sql, /create table if not exists public\.timekeeper_assignments/);
  assert.match(sql, /alter table public\.timekeeper_cycles enable row level security/);
  assert.match(sql, /array_agg\(m\.user_id order by m\.sort_group, m\.carryover_position, random\(\)\)/);
  assert.match(sql, /unique \(team_id, practice_date\)/);
  assert.match(sql, /create or replace function public\.replace_keiko_timekeeper/);
  assert.match(sql, /carryover_order = v_new_absent/);
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
