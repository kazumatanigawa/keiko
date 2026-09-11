const SUPABASE_URL = requireEnv('SUPABASE_URL').replace(/\/+$/, '');
const SERVICE_KEY = requireEnv('SUPABASE_SERVICE_ROLE_KEY');
const PUBLIC_KEY = Deno.env.get('SUPABASE_ANON_KEY') || requireEnv('KEIKO_PUBLISHABLE_KEY');
const AUTH_PEPPER = requireEnv('KEIKO_AUTH_PEPPER');
const REGISTRATION_CODE = requireEnv('KEIKO_REGISTRATION_CODE');
const REPORT_EMAIL_TO = 'kokongakumeiza@gmail.com';
const REPORT_EMAIL_SUBJECT = '🚨KEIKO OS からの通報🚨';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, apikey, content-type, x-client-info',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

type JsonRecord = Record<string, unknown>;

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });
  if (request.method !== 'POST') return json({ status: 'error', code: 'post_required', message: 'POSTでアクセスしてください。' }, 405);

  try {
    const payload = await request.json() as JsonRecord;
    const action = text(payload.action);
    if (!action) throw userError('invalid_request', '操作が指定されていません。');

    if (action === 'health') return json({ status: 'ok', service: 'keiko-api', version: '2026-09-11-01' });
    if (action === 'getTeams') return json(await getTeams());
    if (action === 'login') return json(await login(payload));
    if (action === 'register') return json(await register(payload));
    if (action === 'manageMembership') return json(await manageMembership(request, payload));
    if (action === 'getOperatorDashboard') return json(await getOperatorDashboard(request));
    if (action === 'getAdminOverview') return json(await getAdminOverview(request));
    if (action === 'searchAdminUsers') return json(await searchAdminUsers(request, payload));
    if (action === 'getAdminUserDetail') return json(await getAdminUserDetail(request, payload));
    if (action === 'getAdminTeamDetail') return json(await getAdminTeamDetail(request, payload));
    if (action === 'updateTeamSettings') return json(await updateTeamSettings(request, payload));
    if (action === 'reportContent') return json(await reportContent(request, payload));
    if (action === 'moderateReport') return json(await moderateReport(request, payload));
    throw userError('unsupported_action', 'この操作には対応していません。');
  } catch (error) {
    console.error(error instanceof Error ? error.message : error);
    const appError = error as { code?: string; publicMessage?: string };
    return json({
      status: 'error',
      code: appError.code || 'server_error',
      message: appError.publicMessage || '処理に失敗しました。時間をおいて再度お試しください。',
    }, appError.code ? 400 : 500);
  }
});

async function getTeams() {
  const teams = await adminRequest('/rest/v1/teams', {
    query: { select: 'id,team_name,team_type,category,global_notes_enabled,global_scope', status: 'eq.active', order: 'team_name.asc' },
  }) as JsonRecord[];
  return {
    status: 'ok',
    teams: teams
      .filter((team) => team.category !== 'personal')
      .map((team) => ({
        teamId: team.id,
        teamName: team.team_name,
        teamType: team.team_type || 'general',
        globalNotesEnabled: Boolean(team.global_notes_enabled),
        globalScope: team.global_scope || 'disabled',
      })),
  };
}

async function login(payload: JsonRecord) {
  const name = requiredText(payload.name, '名前', 80);
  const pin = validatePin(payload.pin);
  const profile = await findProfile(name);
  if (!profile || profile.status !== 'active') throw invalidCredentials();

  const authUser = await adminRequest(`/auth/v1/admin/users/${encodeURIComponent(String(profile.id))}`) as JsonRecord;
  if (!authUser.email) throw invalidCredentials();
  const password = await deriveAuthPassword(String(profile.legacy_user_id || profile.id), pin);
  const session = await passwordLogin(String(authUser.email), password);
  return buildSessionResponse(session);
}

async function register(payload: JsonRecord) {
  if (!constantTimeEquals(text(payload.registrationCode), REGISTRATION_CODE)) {
    throw userError('registration_denied', '登録コードが正しくありません。');
  }

  const lastName = requiredText(payload.lastName, '姓', 40);
  const firstName = requiredText(payload.firstName, '名', 40);
  const lastNameKana = requiredKatakana(payload.lastNameKana, '姓の読み');
  const firstNameKana = requiredKatakana(payload.firstNameKana, '名の読み');
  const displayName = normalizeDisplayName(payload.name || `${lastName}${firstName}`);
  const loginId = normalizeLoginId(displayName);
  const pin = validatePin(payload.pin);
  const teamMode = text(payload.teamMode);
  if (!['existing', 'launch'].includes(teamMode)) throw userError('invalid_request', '所属チームの指定が正しくありません。');
  if (await findProfile(loginId)) return { status: 'exists' };

  let team: JsonRecord | null = null;
  let teamName = '';
  let teamType = 'general';
  if (teamMode === 'existing') {
    const teamId = requireUuid(payload.teamId, 'チーム');
    const teams = await adminRequest('/rest/v1/teams', {
      query: { select: 'id,team_name,team_type,status', id: `eq.${teamId}`, status: 'eq.active', limit: '1' },
    }) as JsonRecord[];
    if (!teams.length) throw userError('team_not_found', '選択したチームが見つかりません。');
    team = teams[0];
    teamName = String(team.team_name);
    teamType = normalizeTeamType(team.team_type);
  } else {
    teamName = requiredText(payload.teamName, 'チーム名', 100);
    teamType = normalizeTeamType(payload.teamType);
    const matches = await adminRequest('/rest/v1/teams', {
      query: { select: 'id', team_name: `eq.${teamName}`, status: 'eq.active', limit: '1' },
    }) as JsonRecord[];
    if (matches.length) return { status: 'team_exists' };
  }

  const legacyUserId = generateLegacyId('usr');
  const email = `${legacyUserId.toLowerCase()}@auth.keiko.invalid`;
  const password = await deriveAuthPassword(legacyUserId, pin);
  let authUserId = '';

  try {
    const authUser = await adminRequest('/auth/v1/admin/users', {
      method: 'POST',
      body: { email, password, email_confirm: true, user_metadata: { display_name: displayName } },
    }) as JsonRecord;
    authUserId = String(authUser.id || '');
    if (!authUserId) throw new Error('Supabase Auth user was not created.');

    const userType = teamType === 'student' ? 'student' : 'general';
    await adminRequest('/rest/v1/profiles', {
      method: 'POST',
      body: {
        id: authUserId, user_type: userType, display_name: displayName, login_id: loginId,
        status: 'active', legacy_user_id: legacyUserId, name: displayName,
        last_name: lastName, first_name: firstName, last_name_kana: lastNameKana, first_name_kana: firstNameKana,
      },
    });

    if (teamMode === 'launch') {
      const createdTeams = await adminRequest('/rest/v1/teams', {
        method: 'POST', prefer: 'return=representation',
        body: {
          team_name: teamName, team_type: teamType, audience_type: teamType,
          owner_user_id: authUserId, status: 'active', legacy_team_id: generateLegacyId('team'),
        },
      }) as JsonRecord[];
      team = createdTeams[0];
    }

    await adminRequest('/rest/v1/team_members', {
      method: 'POST',
      body: { team_id: team?.id, user_id: authUserId, team_role: teamMode === 'launch' ? 'owner_admin' : 'member' },
    });
    const profilePath = userType === 'student' ? '/rest/v1/student_profiles' : '/rest/v1/general_profiles';
    const profileBody = userType === 'student'
      ? { user_id: authUserId, school_name: teamName, grade: '', role_label: '', term: '' }
      : { user_id: authUserId, category: teamName, bio: '' };
    await adminRequest(profilePath, { method: 'POST', body: profileBody });
    await adminRequest('/rest/v1/rpc/keiko_ensure_personal_membership', {
      method: 'POST', body: { p_user_id: authUserId },
    });

    return buildSessionResponse(await passwordLogin(email, password));
  } catch (error) {
    if (authUserId) {
      try { await adminRequest(`/auth/v1/admin/users/${encodeURIComponent(authUserId)}`, { method: 'DELETE' }); } catch (cleanupError) { console.error(cleanupError); }
    }
    throw error;
  }
}

async function buildSessionResponse(session: JsonRecord) {
  const accessToken = String(session.access_token || '');
  const context = await userRequest('/rest/v1/rpc/get_keiko_session_context', accessToken, { method: 'POST', body: { p_team_id: null } }) as JsonRecord;
  if (!context.user_id) throw userError('profile_required', '利用者情報が設定されていません。');
  return {
    status: 'ok', accessToken, refreshToken: session.refresh_token,
    expiresIn: Number(session.expires_in || 3600), expiresAt: Number(session.expires_at || 0),
    userId: context.user_id, name: context.display_name, teamId: context.team_id,
    group: context.team_name, teamType: context.team_type, userType: context.user_type, role: context.team_role,
    appRole: context.app_role || 'player',
    globalParticipationEnabled: Boolean(context.global_participation_enabled),
    isPersonal: Boolean(context.is_personal), teams: Array.isArray(context.teams) ? context.teams : [],
  };
}

async function manageMembership(request: Request, payload: JsonRecord) {
  const operation = text(payload.operation);
  if (!['join', 'transfer', 'graduate'].includes(operation)) {
    throw userError('invalid_request', '所属変更の種類が正しくありません。');
  }
  if (operation !== 'graduate' && !constantTimeEquals(text(payload.registrationCode), REGISTRATION_CODE)) {
    throw userError('registration_denied', '登録コードが正しくありません。');
  }

  const accessToken = readBearerToken(request);
  const authUser = await userRequest('/auth/v1/user', accessToken) as JsonRecord;
  const userId = requireUuid(authUser.id, '利用者');
  const targetTeamId = operation === 'graduate' ? null : requireUuid(payload.teamId, 'チーム');
  const preferredTeamId = text(payload.currentTeamId)
    ? requireUuid(payload.currentTeamId, '現在のチーム')
    : null;

  const context = await adminRequest('/rest/v1/rpc/manage_keiko_membership', {
    method: 'POST',
    body: {
      p_user_id: userId,
      p_action: operation,
      p_target_team_id: targetTeamId,
      p_preferred_team_id: operation === 'transfer' ? targetTeamId : preferredTeamId,
    },
  }) as JsonRecord;

  return {
    status: 'ok',
    teamId: context.team_id,
    group: context.team_name,
    teamType: context.team_type,
    userType: context.user_type,
    role: context.team_role,
    appRole: context.app_role || 'player',
    globalParticipationEnabled: Boolean(context.global_participation_enabled),
    isPersonal: Boolean(context.is_personal),
    teams: Array.isArray(context.teams) ? context.teams : [],
  };
}

async function getOperatorDashboard(request: Request) {
  const userId = await authenticatedUserId(request);
  const [teams, reports] = await Promise.all([
    adminRequest('/rest/v1/rpc/get_keiko_operator_teams', {
      method: 'POST', body: { p_actor_user_id: userId },
    }) as Promise<JsonRecord>,
    adminRequest('/rest/v1/rpc/get_keiko_operator_reports', {
      method: 'POST', body: { p_actor_user_id: userId, p_limit: 20 },
    }) as Promise<JsonRecord>,
  ]);
  return {
    status: 'ok',
    teams: Array.isArray(teams.teams) ? teams.teams : [],
    reports: Array.isArray(reports.reports) ? reports.reports : [],
  };
}

async function getAdminOverview(request: Request) {
  const userId = await authenticatedUserId(request);
  const result = await adminRequest('/rest/v1/rpc/get_keiko_admin_overview', {
    method: 'POST', body: { p_actor_user_id: userId },
  }) as JsonRecord;
  return { status: 'ok', ...result };
}

async function searchAdminUsers(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const query = text(payload.query).slice(0, 100);
  const teamId = text(payload.teamId) ? requireUuid(payload.teamId, 'チーム') : null;
  const status = text(payload.status);
  if (status && !['active', 'inactive', 'suspended'].includes(status)) {
    throw userError('invalid_request', 'ユーザー状態の指定が正しくありません。');
  }
  const limit = boundedInteger(payload.limit, 1, 100, 50);
  const offset = boundedInteger(payload.offset, 0, 100000, 0);
  const result = await adminRequest('/rest/v1/rpc/get_keiko_admin_users', {
    method: 'POST',
    body: {
      p_actor_user_id: userId,
      p_query: query,
      p_team_id: teamId,
      p_status: status,
      p_limit: limit,
      p_offset: offset,
    },
  }) as JsonRecord;
  return { status: 'ok', ...result };
}

async function getAdminUserDetail(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const targetUserId = requireUuid(payload.userId, 'ユーザー');
  const result = await adminRequest('/rest/v1/rpc/get_keiko_admin_user_detail', {
    method: 'POST',
    body: {
      p_actor_user_id: userId,
      p_user_id: targetUserId,
      p_limit: boundedInteger(payload.limit, 1, 200, 100),
    },
  }) as JsonRecord;
  return { status: 'ok', ...result };
}

async function getAdminTeamDetail(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const section = text(payload.section) || 'logs';
  if (!['logs', 'notes', 'members'].includes(section)) {
    throw userError('invalid_request', '表示内容の指定が正しくありません。');
  }
  const result = await adminRequest('/rest/v1/rpc/get_keiko_admin_team_detail', {
    method: 'POST',
    body: {
      p_actor_user_id: userId,
      p_team_id: requireUuid(payload.teamId, 'チーム'),
      p_section: section,
      p_limit: boundedInteger(payload.limit, 1, 100, 50),
      p_offset: boundedInteger(payload.offset, 0, 100000, 0),
    },
  }) as JsonRecord;
  return { status: 'ok', ...result };
}

async function updateTeamSettings(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const teamId = requireUuid(payload.teamId, 'チーム');
  const result = await adminRequest('/rest/v1/rpc/update_keiko_team_global_settings', {
    method: 'POST',
    body: { p_actor_user_id: userId, p_team_id: teamId, p_enabled: Boolean(payload.globalNotesEnabled) },
  }) as JsonRecord;
  return result;
}

async function reportContent(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const contentType = text(payload.contentType);
  if (!['note', 'comment'].includes(contentType)) throw userError('invalid_request', '通報対象が正しくありません。');
  const contentId = requireUuid(payload.contentId, '投稿');
  const reason = requiredText(payload.reason, '通報理由', 1000);
  const report = await adminRequest('/rest/v1/rpc/create_keiko_content_report', {
    method: 'POST',
    body: { p_reporter_user_id: userId, p_content_type: contentType, p_content_id: contentId, p_reason: reason },
  }) as JsonRecord;

  const email = await sendReportEmail(report, userId);
  if (email.status !== 'pending') {
    await adminRequest('/rest/v1/rpc/mark_keiko_report_email', {
      method: 'POST',
      body: { p_report_id: report.reportId, p_status: email.status, p_error: email.error || '' },
    });
  }
  return { status: 'ok', reportId: report.reportId, emailStatus: email.status };
}

async function moderateReport(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const operation = text(payload.operation);
  if (!['resolve', 'dismiss', 'hide_content'].includes(operation)) {
    throw userError('invalid_request', '通報への対応方法が正しくありません。');
  }
  return await adminRequest('/rest/v1/rpc/moderate_keiko_report', {
    method: 'POST',
    body: {
      p_actor_user_id: userId,
      p_report_id: requireUuid(payload.reportId, '通報'),
      p_action: operation,
    },
  }) as JsonRecord;
}

async function sendReportEmail(report: JsonRecord, reporterUserId: string) {
  const apiKey = Deno.env.get('RESEND_API_KEY') || '';
  const from = Deno.env.get('KEIKO_REPORT_FROM_EMAIL') || '';
  if (!apiKey || !from) {
    console.warn(`Report ${String(report.reportId)} stored; email provider is not configured.`);
    return { status: 'pending', error: '' };
  }

  const body = [
    'KEIKO OSで投稿の通報を受け付けました。',
    '',
    `通報ID: ${String(report.reportId || '')}`,
    `対象: ${String(report.contentType || '')}`,
    `投稿者: ${String(report.authorName || '未設定')}`,
    `所属: ${String(report.teamName || '個人')}`,
    `通報者ID: ${reporterUserId}`,
    `通報日時: ${String(report.createdAt || new Date().toISOString())}`,
    '',
    '通報理由:',
    String(report.reason || ''),
    '',
    '投稿内容:',
    String(report.content || ''),
  ].join('\n');

  try {
    const response = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ from, to: [REPORT_EMAIL_TO], subject: REPORT_EMAIL_SUBJECT, text: body }),
    });
    if (!response.ok) {
      const remote = await response.text();
      console.error(`Report email failed (${response.status}): ${remote.slice(0, 300)}`);
      return { status: 'failed', error: `Resend ${response.status}` };
    }
    return { status: 'sent', error: '' };
  } catch (error) {
    console.error(error);
    return { status: 'failed', error: error instanceof Error ? error.message : 'email_failed' };
  }
}

async function authenticatedUserId(request: Request) {
  const accessToken = readBearerToken(request);
  const authUser = await userRequest('/auth/v1/user', accessToken) as JsonRecord;
  return requireUuid(authUser.id, '利用者');
}

function readBearerToken(request: Request) {
  const authorization = request.headers.get('authorization') || '';
  const match = authorization.match(/^Bearer\s+(.+)$/i);
  if (!match?.[1]) throw userError('auth_required', 'ログインの有効期限が切れました。');
  return match[1];
}

async function findProfile(name: unknown) {
  const loginId = normalizeLoginId(name);
  if (!loginId) return null;
  const rows = await adminRequest('/rest/v1/profiles', {
    query: { select: 'id,display_name,login_id,status,legacy_user_id', login_id: `eq.${loginId}`, status: 'eq.active', limit: '2' },
  }) as JsonRecord[];
  return rows.length === 1 ? rows[0] : null;
}

async function passwordLogin(email: string, password: string) {
  const response = await fetch(`${SUPABASE_URL}/auth/v1/token?grant_type=password`, {
    method: 'POST',
    headers: { apikey: PUBLIC_KEY, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password }),
  });
  const result = await response.json();
  if (!response.ok) throw invalidCredentials();
  return result as JsonRecord;
}

async function adminRequest(path: string, options: RequestOptions = {}) {
  return supabaseRequest(path, SERVICE_KEY, options);
}

async function userRequest(path: string, accessToken: string, options: RequestOptions = {}) {
  return supabaseRequest(path, accessToken, options);
}

type RequestOptions = { method?: string; query?: Record<string, string>; body?: unknown; prefer?: string };

async function supabaseRequest(path: string, bearer: string, options: RequestOptions = {}) {
  const url = new URL(`${SUPABASE_URL}${path}`);
  Object.entries(options.query || {}).forEach(([key, value]) => url.searchParams.set(key, value));
  const headers: Record<string, string> = { apikey: SERVICE_KEY, Authorization: `Bearer ${bearer}`, Accept: 'application/json' };
  if (options.prefer) headers.Prefer = options.prefer;
  if (options.body !== undefined) headers['Content-Type'] = 'application/json';
  const response = await fetch(url, { method: options.method || 'GET', headers, body: options.body === undefined ? undefined : JSON.stringify(options.body) });
  const result = response.status === 204 ? {} : await response.json().catch(() => ({}));
  if (!response.ok) {
    const remote = result as JsonRecord;
    console.error(`Supabase ${options.method || 'GET'} ${path}: ${response.status} ${String(remote.code || remote.error_code || '')}`);
    if ([401, 403].includes(response.status)) throw userError('auth_required', 'ログインの有効期限が切れました。');
    if (response.status === 409 || remote.code === '23505') throw userError('conflict', '同じ情報がすでに登録されています。');
    throw new Error(`Supabase request failed (${response.status})`);
  }
  return Array.isArray(result) && path.includes('/rpc/') ? (result[0] || {}) : result;
}

async function deriveAuthPassword(legacyUserId: string, pin: string) {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(AUTH_PEPPER), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const digest = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(`${legacyUserId}:${pin}`));
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function normalizeLoginId(value: unknown) { return text(value).normalize('NFKC').replace(/[\s\u3000]+/g, '').toLowerCase(); }
function normalizeDisplayName(value: unknown) { return requiredText(value, '名前', 80).replace(/[\s\u3000]+/g, ''); }
function normalizeTeamType(value: unknown) { return text(value) === 'student' ? 'student' : 'general'; }
function validatePin(value: unknown) { const pin = text(value); if (!/^\d{4}$/.test(pin)) throw userError('invalid_request', 'PINは4桁の数字で入力してください。'); return pin; }
function requireUuid(value: unknown, label: string) { const valueText = text(value); if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(valueText)) throw userError('invalid_request', `${label}の指定が正しくありません。`); return valueText; }
function requiredKatakana(value: unknown, label: string) { const valueText = requiredText(value, label, 40); if (!/^[ァ-ヶー\s\u3000]+$/.test(valueText)) throw userError('invalid_request', `${label}はカタカナで入力してください。`); return valueText.replace(/[\s\u3000]+/g, ''); }
function requiredText(value: unknown, label: string, max: number) { const valueText = text(value); if (!valueText) throw userError('invalid_request', `${label}を入力してください。`); if (valueText.length > max) throw userError('invalid_request', `${label}は${max}文字以内で入力してください。`); return valueText; }
function boundedInteger(value: unknown, min: number, max: number, fallback: number) { const parsed = Number(value); return Number.isInteger(parsed) ? Math.max(min, Math.min(max, parsed)) : fallback; }
function text(value: unknown) { return value === null || value === undefined ? '' : String(value).trim(); }
function generateLegacyId(prefix: string) { return `${prefix}_${crypto.randomUUID().replaceAll('-', '')}`; }
function constantTimeEquals(left: string, right: string) { let difference = left.length ^ right.length; const max = Math.max(left.length, right.length); for (let index = 0; index < max; index += 1) difference |= (left.charCodeAt(index) || 0) ^ (right.charCodeAt(index) || 0); return difference === 0; }
function requireEnv(name: string) { const value = Deno.env.get(name); if (!value) throw new Error(`Missing environment variable: ${name}`); return value; }
function userError(code: string, publicMessage: string) { return Object.assign(new Error(publicMessage), { code, publicMessage }); }
function invalidCredentials() { return userError('invalid_credentials', '名前またはPINが違います。'); }
function json(payload: unknown, status = 200) { return new Response(JSON.stringify(payload), { status, headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' } }); }
