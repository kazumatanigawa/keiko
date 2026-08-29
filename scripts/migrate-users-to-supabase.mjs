import { createHmac } from 'node:crypto';

const SPREADSHEET_ID = '1x1HUq_FuRxEdGc3sNWut5rSlKKU0UIMcfn9VNNjF538';
const USERS_SHEET = 'ユーザー';
const DEFAULT_ADMIN_AUTH_USER_ID = '6603fa37-433e-4042-877c-98935ffabba0';
const DEFAULT_ADMIN_LEGACY_USER_ID = 'usr_c9851540b96d4bf38872';
const DEFAULT_ADMIN_PRIMARY_TEAM_LEGACY_ID = 'team_263f18863f204a5087dd';

const applyChanges = process.argv.includes('--apply');
const supabaseUrl = String(process.env.SUPABASE_URL || '').replace(/\/$/, '');
const adminApiKey = String(
  process.env.SUPABASE_SECRET_KEY ||
    process.env.SUPABASE_SERVICE_ROLE_KEY ||
    '',
);
const authPepper = String(process.env.KEIKO_AUTH_PEPPER || '');
const adminAuthUserId = String(
  process.env.ADMIN_AUTH_USER_ID || DEFAULT_ADMIN_AUTH_USER_ID,
);
const adminLegacyUserId = String(
  process.env.ADMIN_LEGACY_USER_ID || DEFAULT_ADMIN_LEGACY_USER_ID,
);
const adminPrimaryTeamLegacyId = String(
  process.env.ADMIN_PRIMARY_TEAM_LEGACY_ID ||
    DEFAULT_ADMIN_PRIMARY_TEAM_LEGACY_ID,
);

function normalizeName(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

function normalizeRole(value) {
  if (value === 'owner_admin' || value === 'founder') return 'owner_admin';
  if (value === 'admin') return 'admin';
  return 'member';
}

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
      if (row.some((value) => value !== '')) rows.push(row);
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

async function loadLegacyUsers() {
  const url = new URL(
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq`,
  );
  url.searchParams.set('tqx', 'out:csv');
  url.searchParams.set('sheet', USERS_SHEET);
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to read users sheet: HTTP ${response.status}`);
  }

  const rows = parseCsv(await response.text());
  const headers = rows.shift() || [];
  const indexes = new Map(headers.map((header, index) => [header, index]));
  for (const header of ['name', 'pin', 'userId', 'teamId', 'role', 'userType']) {
    if (!indexes.has(header)) throw new Error(`Missing users sheet header: ${header}`);
  }
  const get = (row, header) => String(row[indexes.get(header)] || '').trim();

  return rows
    .map((row) => ({
      name: get(row, 'name'),
      pin: get(row, 'pin').padStart(4, '0'),
      createdAt: get(row, 'createdAt'),
      group: get(row, 'group'),
      legacyUserId: get(row, 'userId'),
      legacyTeamId: get(row, 'teamId'),
      role: normalizeRole(get(row, 'role')),
      updatedAt: get(row, 'updatedAt'),
      userType: get(row, 'userType') || 'general',
      lastName: get(row, 'lastName'),
      firstName: get(row, 'firstName'),
      lastNameKana: get(row, 'lastNameKana'),
      firstNameKana: get(row, 'firstNameKana'),
    }))
    .filter((user) => user.name && user.legacyUserId && user.legacyTeamId);
}

function validateUsers(users) {
  const errors = [];
  const names = new Map();
  const ids = new Set();
  for (const user of users) {
    const loginId = normalizeName(user.name);
    if (!/^\d{4}$/.test(user.pin)) errors.push(`${user.legacyUserId}: invalid PIN`);
    if (!['student', 'general'].includes(user.userType)) {
      errors.push(`${user.legacyUserId}: invalid userType ${user.userType}`);
    }
    if (ids.has(user.legacyUserId)) errors.push(`${user.legacyUserId}: duplicate ID`);
    ids.add(user.legacyUserId);
    if (names.has(loginId)) {
      errors.push(`${user.legacyUserId}: duplicate name with ${names.get(loginId)}`);
    }
    names.set(loginId, user.legacyUserId);
  }
  if (errors.length) throw new Error(errors.join('\n'));
}

function syntheticEmail(user) {
  return `${user.legacyUserId.toLowerCase()}@auth.keiko.invalid`;
}

function deriveAuthPassword(user) {
  return createHmac('sha256', authPepper)
    .update(`${user.legacyUserId}:${user.pin}`)
    .digest('hex');
}

async function request(path, options = {}) {
  const response = await fetch(`${supabaseUrl}${path}`, {
    ...options,
    headers: {
      apikey: adminApiKey,
      // New sb_secret keys must be sent as an API key, not as a bearer JWT.
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

async function upsert(table, conflicts, payload) {
  return request(`/rest/v1/${table}?on_conflict=${encodeURIComponent(conflicts)}`, {
    method: 'POST',
    headers: { prefer: 'resolution=merge-duplicates,return=minimal' },
    body: JSON.stringify(payload),
  });
}

async function createAuthUser(user) {
  return request('/auth/v1/admin/users', {
    method: 'POST',
    body: JSON.stringify({
      email: syntheticEmail(user),
      password: deriveAuthPassword(user),
      email_confirm: true,
      user_metadata: {
        display_name: user.name,
        legacy_user_id: user.legacyUserId,
      },
    }),
  });
}

async function updateAuthPassword(authUserId, user) {
  return request(`/auth/v1/admin/users/${encodeURIComponent(authUserId)}`, {
    method: 'PUT',
    body: JSON.stringify({
      password: deriveAuthPassword(user),
      user_metadata: {
        display_name: user.name,
        legacy_user_id: user.legacyUserId,
      },
    }),
  });
}

async function migrateUser(user, teams, authUsers) {
  const team = teams.get(user.legacyTeamId);
  if (!team) throw new Error(`${user.legacyUserId}: missing team ${user.legacyTeamId}`);

  let authUser;
  if (user.legacyUserId === adminLegacyUserId) {
    authUser = { id: adminAuthUserId };
    await updateAuthPassword(authUser.id, user);
  } else {
    const email = syntheticEmail(user).toLowerCase();
    authUser = authUsers.get(email);
    if (!authUser) {
      authUser = await createAuthUser(user);
      authUsers.set(email, authUser);
    }
  }

  const isGlobalAdmin = user.legacyUserId === adminLegacyUserId;
  await upsert('profiles', 'id', {
    id: authUser.id,
    legacy_user_id: user.legacyUserId,
    login_id: normalizeName(user.name),
    name: user.name,
    display_name: user.name,
    last_name: user.lastName || null,
    first_name: user.firstName || null,
    last_name_kana: user.lastNameKana || null,
    first_name_kana: user.firstNameKana || null,
    user_type: isGlobalAdmin ? 'admin' : user.userType,
    status: 'active',
  });
  await upsert('team_members', 'team_id,user_id', {
    team_id: team.id,
    user_id: authUser.id,
    team_role: user.role,
  });

  if (user.userType === 'student') {
    await upsert('student_profiles', 'user_id', {
      user_id: authUser.id,
      school_name: user.group || null,
    });
  } else if (!isGlobalAdmin) {
    await upsert('general_profiles', 'user_id', {
      user_id: authUser.id,
      category: user.group || null,
    });
  }

  if (user.role === 'owner_admin' && !team.owner_user_id) {
    await request(`/rest/v1/teams?id=eq.${encodeURIComponent(team.id)}`, {
      method: 'PATCH',
      headers: { prefer: 'return=minimal' },
      body: JSON.stringify({ owner_user_id: authUser.id }),
    });
    team.owner_user_id = authUser.id;
  }
}

const users = await loadLegacyUsers();
validateUsers(users);
const summary = {
  total: users.length,
  students: users.filter((user) => user.userType === 'student').length,
  general: users.filter((user) => user.userType === 'general').length,
  ownerAdmins: users.filter((user) => user.role === 'owner_admin').length,
};

if (!applyChanges) {
  console.log(JSON.stringify({ mode: 'dry-run', ...summary }, null, 2));
  console.log('No Supabase data was changed.');
  process.exit(0);
}
if (!supabaseUrl || !adminApiKey || !authPepper) {
  throw new Error(
    'SUPABASE_URL, SUPABASE_SECRET_KEY and KEIKO_AUTH_PEPPER are required',
  );
}

const teamRows = await request('/rest/v1/teams?select=id,legacy_team_id,owner_user_id');
const teams = new Map(teamRows.map((team) => [team.legacy_team_id, team]));
const authResult = await request('/auth/v1/admin/users?page=1&per_page=1000');
const authRows = Array.isArray(authResult) ? authResult : authResult.users || [];
const authUsers = new Map(
  authRows.map((user) => [String(user.email || '').toLowerCase(), user]),
);

for (const user of users) {
  await migrateUser(user, teams, authUsers);
  console.log(`Migrated ${user.legacyUserId}`);
}

const adminPrimaryTeam = teams.get(adminPrimaryTeamLegacyId);
if (!adminPrimaryTeam) {
  throw new Error(`Missing admin primary team ${adminPrimaryTeamLegacyId}`);
}
await upsert('team_members', 'team_id,user_id', {
  team_id: adminPrimaryTeam.id,
  user_id: adminAuthUserId,
  team_role: 'owner_admin',
});
if (!adminPrimaryTeam.owner_user_id) {
  await request(`/rest/v1/teams?id=eq.${encodeURIComponent(adminPrimaryTeam.id)}`, {
    method: 'PATCH',
    headers: { prefer: 'return=minimal' },
    body: JSON.stringify({ owner_user_id: adminAuthUserId }),
  });
}
console.log(JSON.stringify({ mode: 'applied', ...summary }, null, 2));
