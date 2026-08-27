import { createHash, createHmac } from 'node:crypto';

const SPREADSHEET_ID = '1x1HUq_FuRxEdGc3sNWut5rSlKKU0UIMcfn9VNNjF538';
const LOGS_SHEET = '稽古ログ';
const SOURCE_SYSTEM = 'google_sheets';
const BATCH_SIZE = 100;

const applyChanges = process.argv.includes('--apply');
const supabaseUrl = String(process.env.SUPABASE_URL || '').replace(/\/$/, '');
const adminApiKey = String(
  process.env.SUPABASE_SECRET_KEY ||
    process.env.SUPABASE_SERVICE_ROLE_KEY ||
    '',
);
const authPepper = String(process.env.KEIKO_AUTH_PEPPER || '');
const HISTORICAL_TEAM_OVERRIDES = new Map([
  ['宮下仔々', '府中東高校和太鼓部'],
]);

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (quoted && char === '"' && next === '"') {
      field += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (!quoted && char === ',') {
      row.push(field);
      field = '';
    } else if (!quoted && (char === '\n' || char === '\r')) {
      if (char === '\r' && next === '\n') index += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += char;
    }
  }

  row.push(field);
  if (row.some((value) => value !== '')) rows.push(row);
  return rows;
}

function asText(value) {
  return String(value || '').trim();
}

function asNullableText(value) {
  const text = asText(value);
  return text || null;
}

function normalizeName(value) {
  return asText(value)
    .normalize('NFKC')
    .replace(/[\s\u3000]+/g, '');
}

function normalizeDate(value) {
  const text = asText(value).replaceAll('/', '-');
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : '';
}

function asIsoTimestamp(value, fallback) {
  const text = asText(value);
  if (!text) return fallback;
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? fallback : date.toISOString();
}

function normalizeAchievement(value) {
  const status = asText(value);
  return status === 'done' || status === 'pending' ? status : null;
}

async function loadLegacyLogs() {
  const url = new URL(
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq`,
  );
  url.searchParams.set('tqx', 'out:csv');
  url.searchParams.set('sheet', LOGS_SHEET);
  // The second row contains obsolete headers, so force only row 1 as the header.
  url.searchParams.set('headers', '1');
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to read practice logs sheet: HTTP ${response.status}`);
  }

  const rows = parseCsv(await response.text());
  const headers = rows.shift() || [];
  const indexes = new Map(headers.map((header, index) => [header, index]));
  const requiredHeaders = [
    'date',
    'name',
    'cond',
    'learning',
    'next',
    'goodNew',
    'createdAt',
    'userId',
    'teamId',
  ];
  for (const header of requiredHeaders) {
    if (!indexes.has(header)) throw new Error(`Missing logs sheet header: ${header}`);
  }
  const get = (row, header) => asText(row[indexes.get(header)]);

  return rows.map((row, index) => {
    const sheetRow = index + 2;
    const date = normalizeDate(get(row, 'date') || row[0]);
    const rawUserId = get(row, 'userId');
    const rawTeamId = get(row, 'teamId');
    const createdAt = asIsoTimestamp(
      get(row, 'createdAt'),
      date ? new Date(`${date}T00:00:00+09:00`).toISOString() : null,
    );
    return {
      sheetRow,
      legacySourceRowId: `${SPREADSHEET_ID}:${LOGS_SHEET}:${sheetRow}`,
      legacyUserId: rawUserId.startsWith('usr_') ? rawUserId : '',
      legacyTeamId: rawTeamId.startsWith('team_') ? rawTeamId : '',
      legacyName: get(row, 'name') || asText(row[1]),
      legacyGroup: get(row, 'group'),
      practiceDate: date,
      condition: Number(get(row, 'cond') || row[2]),
      learning: get(row, 'learning'),
      nextAction: get(row, 'next'),
      goodNew: get(row, 'goodNew'),
      achievementStatus: normalizeAchievement(get(row, 'achievementStatus')),
      whyMissed: get(row, 'whyMissed'),
      retryPlan: get(row, 'retryPlan'),
      displayNameSnapshot: get(row, 'displayName') || get(row, 'name'),
      gradeSnapshot: asNullableText(get(row, 'grade')),
      termSnapshot: asNullableText(get(row, 'term')),
      createdAt,
      updatedAt: asIsoTimestamp(get(row, 'updatedAt'), createdAt),
    };
  });
}

function classifyLogs(logs) {
  const valid = [];
  const ignored = [];
  const invalid = [];

  for (const log of logs) {
    if (log.sheetRow === 2) {
      ignored.push(log);
      continue;
    }
    const empty = !log.practiceDate && !log.legacyUserId && !log.legacyTeamId;
    if (empty) {
      ignored.push(log);
      continue;
    }
    const errors = [];
    if (!/^\d{4}-\d{2}-\d{2}$/.test(log.practiceDate)) errors.push('invalid date');
    if (!log.legacyUserId && !log.legacyName) errors.push('missing user reference');
    if (!log.legacyTeamId && !log.legacyGroup && !log.legacyName) {
      errors.push('missing team reference');
    }
    if (!Number.isInteger(log.condition) || log.condition < 1 || log.condition > 5) {
      errors.push('invalid condition');
    }
    if (!log.createdAt || !log.updatedAt) errors.push('invalid timestamp');
    if (errors.length) invalid.push({ row: log.sheetRow, errors });
    else valid.push(log);
  }
  return { valid, ignored, invalid };
}

function inferMissingGroups(logs) {
  const groupsByName = new Map();
  for (const log of logs) {
    const name = normalizeName(log.legacyName);
    if (!name || !log.legacyGroup) continue;
    if (!groupsByName.has(name)) groupsByName.set(name, new Set());
    groupsByName.get(name).add(log.legacyGroup);
  }

  for (const log of logs) {
    if (log.legacyGroup) continue;
    const override = HISTORICAL_TEAM_OVERRIDES.get(normalizeName(log.legacyName));
    if (override) {
      log.legacyGroup = override;
      continue;
    }
    const groups = groupsByName.get(normalizeName(log.legacyName));
    if (groups?.size === 1) log.legacyGroup = [...groups][0];
  }
}

async function request(path, options = {}) {
  const response = await fetch(`${supabaseUrl}${path}`, {
    ...options,
    headers: {
      apikey: adminApiKey,
      'content-type': 'application/json',
      ...(options.headers || {}),
    },
  });
  const text = await response.text();
  if (!response.ok) {
    throw new Error(`${options.method || 'GET'} ${path}: ${response.status} ${text}`);
  }
  return text ? JSON.parse(text) : null;
}

async function loadAll(path, pageSize = 1000) {
  const rows = [];
  for (let offset = 0; ; offset += pageSize) {
    const separator = path.includes('?') ? '&' : '?';
    const page = await request(`${path}${separator}limit=${pageSize}&offset=${offset}`);
    rows.push(...page);
    if (page.length < pageSize) return rows;
  }
}

async function upsert(table, conflicts, payload) {
  return request(`/rest/v1/${table}?on_conflict=${encodeURIComponent(conflicts)}`, {
    method: 'POST',
    headers: { prefer: 'resolution=merge-duplicates,return=minimal' },
    body: JSON.stringify(payload),
  });
}

function historicalIdentity(name, teamName) {
  const digest = createHash('sha256')
    .update(`${normalizeName(name)}:${teamName}`)
    .digest('hex')
    .slice(0, 20);
  return {
    email: `historical_${digest}@auth.keiko.invalid`,
    legacyUserId: `historical_${digest}`,
    loginId: `historical_${digest}`,
  };
}

async function ensureHistoricalProfile(identity, team, authUsers) {
  const details = historicalIdentity(identity.name, team.team_name);
  let authUser = authUsers.get(details.email);
  if (!authUser) {
    if (!authPepper) throw new Error('KEIKO_AUTH_PEPPER is required');
    const password = createHmac('sha256', authPepper)
      .update(details.legacyUserId)
      .digest('hex');
    authUser = await request('/auth/v1/admin/users', {
      method: 'POST',
      body: JSON.stringify({
        email: details.email,
        password,
        email_confirm: true,
        user_metadata: {
          display_name: identity.name,
          legacy_user_id: details.legacyUserId,
          historical: true,
        },
      }),
    });
    authUsers.set(details.email, authUser);
  }

  const userType = team.audience_type === 'student' ? 'student' : 'general';
  await upsert('profiles', 'id', {
    id: authUser.id,
    legacy_user_id: details.legacyUserId,
    login_id: details.loginId,
    name: identity.name,
    display_name: identity.name,
    user_type: userType,
    status: 'inactive',
  });
  await upsert('team_members', 'team_id,user_id', {
    team_id: team.id,
    user_id: authUser.id,
    team_role: 'member',
  });
  if (userType === 'student') {
    await upsert('student_profiles', 'user_id', {
      user_id: authUser.id,
      school_name: team.team_name,
    });
  } else {
    await upsert('general_profiles', 'user_id', {
      user_id: authUser.id,
      category: team.team_name,
    });
  }
  return authUser.id;
}

function chunk(rows, size) {
  const chunks = [];
  for (let index = 0; index < rows.length; index += size) {
    chunks.push(rows.slice(index, index + size));
  }
  return chunks;
}

const sourceRows = await loadLegacyLogs();
inferMissingGroups(sourceRows);
const { valid, ignored, invalid } = classifyLogs(sourceRows);
const fatalInvalid = invalid.filter(
  (entry) =>
    entry.errors.length !== 1 || entry.errors[0] !== 'missing user reference',
);
const sourceSummary = {
  sourceRows: sourceRows.length,
  validLogs: valid.length,
  ignoredRows: ignored.length,
  invalidRows: invalid.length,
  quarantinedRows: invalid.length - fatalInvalid.length,
  namedUsers: new Set(valid.map((log) => normalizeName(log.legacyName))).size,
  linkedById: valid.filter((log) => log.legacyUserId && log.legacyTeamId).length,
  linkedByName: valid.filter((log) => !log.legacyUserId || !log.legacyTeamId).length,
};

if (fatalInvalid.length) {
  console.error(
    JSON.stringify({ ...sourceSummary, invalid: fatalInvalid.slice(0, 20) }, null, 2),
  );
  throw new Error('Practice log source contains invalid rows');
}

if (!applyChanges) {
  console.log(JSON.stringify({ mode: 'dry-run', ...sourceSummary }, null, 2));
  console.log('No Supabase data was changed.');
  process.exit(0);
}
if (!supabaseUrl || !adminApiKey) {
  throw new Error('SUPABASE_URL and SUPABASE_SECRET_KEY are required');
}

const [profiles, teams, memberships, existingLogs, authResult] = await Promise.all([
  loadAll('/rest/v1/profiles?select=id,legacy_user_id,login_id,name'),
  loadAll('/rest/v1/teams?select=id,legacy_team_id,team_name,audience_type'),
  loadAll('/rest/v1/team_members?select=user_id,team_id'),
  loadAll(
    `/rest/v1/practice_logs?select=legacy_source_row_id&source_system=eq.${SOURCE_SYSTEM}`,
  ),
  request('/auth/v1/admin/users?page=1&per_page=1000'),
]);
const profileIds = new Map(profiles.map((profile) => [profile.legacy_user_id, profile.id]));
const teamIds = new Map(teams.map((team) => [team.legacy_team_id, team.id]));
const profileIdsByName = new Map(
  profiles.map((profile) => [profile.login_id || normalizeName(profile.name), profile.id]),
);
const teamIdsByName = new Map(teams.map((team) => [asText(team.team_name), team.id]));
const teamsById = new Map(teams.map((team) => [team.id, team]));
const teamIdsByUser = new Map();
for (const membership of memberships) {
  if (!teamIdsByUser.has(membership.user_id)) {
    teamIdsByUser.set(membership.user_id, new Set());
  }
  teamIdsByUser.get(membership.user_id).add(membership.team_id);
}
const authRows = Array.isArray(authResult) ? authResult : authResult.users || [];
const authUsers = new Map(
  authRows.map((user) => [String(user.email || '').toLowerCase(), user]),
);
const existingIds = new Set(
  existingLogs.map((log) => log.legacy_source_row_id).filter(Boolean),
);

function resolveUserId(log) {
  return (
    profileIds.get(log.legacyUserId) ||
    profileIdsByName.get(normalizeName(log.legacyName))
  );
}

function resolveTeamId(log, userId) {
  const explicit = teamIds.get(log.legacyTeamId) || teamIdsByName.get(log.legacyGroup);
  if (explicit) return explicit;
  const membershipsForUser = teamIdsByUser.get(userId);
  return membershipsForUser?.size === 1 ? [...membershipsForUser][0] : null;
}

const historicalIdentities = new Map();
for (const log of valid) {
  if (resolveUserId(log)) continue;
  const teamId = resolveTeamId(log, null);
  const team = teamsById.get(teamId);
  if (!team) continue;
  const key = `${normalizeName(log.legacyName)}:${team.id}`;
  historicalIdentities.set(key, { name: log.legacyName, team });
}
for (const identity of historicalIdentities.values()) {
  const userId = await ensureHistoricalProfile(identity, identity.team, authUsers);
  profileIdsByName.set(normalizeName(identity.name), userId);
  teamIdsByUser.set(userId, new Set([identity.team.id]));
}

const missingMappings = valid
  .filter((log) => {
    const userId = resolveUserId(log);
    const teamId = resolveTeamId(log, userId);
    return !userId || !teamId;
  })
  .map((log) => ({
    row: log.sheetRow,
    missingUser: !resolveUserId(log),
    missingTeam: !resolveTeamId(log, resolveUserId(log)),
  }));
if (missingMappings.length) {
  console.error(JSON.stringify({ missingMappings: missingMappings.slice(0, 30) }, null, 2));
  throw new Error('Practice logs reference users or teams that were not migrated');
}

const pending = valid.filter((log) => !existingIds.has(log.legacySourceRowId));
const payloads = pending.map((log) => ({
  user_id: resolveUserId(log),
  team_id: resolveTeamId(log, resolveUserId(log)),
  practice_date: log.practiceDate,
  condition: log.condition,
  learning: log.learning,
  next_action: log.nextAction,
  good_new: log.goodNew,
  memo: '',
  visibility: 'team',
  source_system: SOURCE_SYSTEM,
  legacy_source_row_id: log.legacySourceRowId,
  achievement_status: log.achievementStatus,
  why_missed: log.whyMissed,
  retry_plan: log.retryPlan,
  display_name_snapshot: log.displayNameSnapshot || null,
  grade_snapshot: log.gradeSnapshot,
  term_snapshot: log.termSnapshot,
  created_at: log.createdAt,
  updated_at: log.updatedAt,
}));

let inserted = 0;
for (const batch of chunk(payloads, BATCH_SIZE)) {
  await request('/rest/v1/practice_logs', {
    method: 'POST',
    headers: { prefer: 'return=minimal' },
    body: JSON.stringify(batch),
  });
  inserted += batch.length;
  console.log(`Inserted ${inserted}/${payloads.length}`);
}

console.log(
  JSON.stringify(
    {
      mode: 'applied',
      ...sourceSummary,
      alreadyMigrated: existingIds.size,
      historicalProfiles: historicalIdentities.size,
      inserted,
    },
    null,
    2,
  ),
);
