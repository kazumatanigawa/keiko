const SUPABASE_URL = requireEnv('SUPABASE_URL').replace(/\/+$/, '');
const SERVICE_KEY = requireEnv('SUPABASE_SERVICE_ROLE_KEY');
const PUBLIC_KEY = Deno.env.get('SUPABASE_ANON_KEY') || requireEnv('KEIKO_PUBLISHABLE_KEY');
const AUTH_PEPPER = requireEnv('KEIKO_AUTH_PEPPER');
const REGISTRATION_CODE = requireEnv('KEIKO_REGISTRATION_CODE');
const REPORT_EMAIL_TO = 'kokongakumeiza@gmail.com';
const REPORT_EMAIL_SUBJECT = '🚨KEIKO OS からの通報🚨';
const COACHING_EMAIL_SUBJECT = 'KEIKOから指導依頼が届きました';
const BACHI_EMAIL_SUBJECT = 'KEIKOからカスタムバチの依頼が届きました';
const GAKUMEIZA_YOUTUBE_CHANNEL_ID = 'UCDJlnpw9KKDxQVEo4mRA2hg';
const GAKUMEIZA_YOUTUBE_FEED_URL = `https://www.youtube.com/feeds/videos.xml?channel_id=${GAKUMEIZA_YOUTUBE_CHANNEL_ID}`;
const YOUTUBE_FEED_CACHE_MS = 10 * 60 * 1000;

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, apikey, content-type, x-client-info',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

type JsonRecord = Record<string, unknown>;
let youtubeFeedCache: { expiresAt: number; videos: JsonRecord[] } = { expiresAt: 0, videos: [] };

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });
  if (request.method !== 'POST') return json({ status: 'error', code: 'post_required', message: 'POSTでアクセスしてください。' }, 405);

  try {
    const payload = await request.json() as JsonRecord;
    const action = text(payload.action);
    if (!action) throw userError('invalid_request', '操作が指定されていません。');

    if (action === 'health') return json({ status: 'ok', service: 'keiko-api', version: '2026-11-08-01' });
    if (action === 'getTeams') return json(await getTeams());
    if (action === 'login') return json(await login(payload));
    if (action === 'register') return json(await register(payload));
    if (action === 'changePin') return json(await changePin(request, payload));
    if (action === 'manageMembership') return json(await manageMembership(request, payload));
    if (action === 'getOperatorDashboard') return json(await getOperatorDashboard(request));
    if (action === 'getAdminAccess') return json(await getAdminAccess(request));
    if (action === 'getAdminOverview') return json(await getAdminOverview(request));
    if (action === 'searchAdminUsers') return json(await searchAdminUsers(request, payload));
    if (action === 'getAdminUserDetail') return json(await getAdminUserDetail(request, payload));
    if (action === 'getAdminTeamDetail') return json(await getAdminTeamDetail(request, payload));
    if (action === 'setTeamStaffRole') return json(await setTeamStaffRole(request, payload));
    if (action === 'getAdminAnnouncements') return json(await getAdminAnnouncements(request));
    if (action === 'saveAdminAnnouncement') return json(await saveAdminAnnouncement(request, payload));
    if (action === 'getRequestCatalogs') return json(await getRequestCatalogs(request));
    if (action === 'estimateBachiOrder') return json(await estimateBachiOrder(request, payload));
    if (action === 'createBachiOrder') return json(await createBachiOrder(request, payload));
    if (action === 'estimateCoachingRequest') return json(await estimateCoachingRequest(request, payload));
    if (action === 'createCoachingRequest') return json(await createCoachingRequest(request, payload));
    if (action === 'getYoutubeVideos') return json(await getYoutubeVideos(request, payload));
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

async function getYoutubeVideos(request: Request, payload: JsonRecord) {
  await authenticatedUserId(request);
  const limit = boundedInteger(payload.limit, 1, 15, 10);
  const now = Date.now();
  if (youtubeFeedCache.expiresAt > now && youtubeFeedCache.videos.length) {
    return { status: 'ok', videos: youtubeFeedCache.videos.slice(0, limit), cached: true };
  }

  const response = await fetch(GAKUMEIZA_YOUTUBE_FEED_URL, {
    headers: { Accept: 'application/atom+xml, application/xml;q=0.9, text/xml;q=0.8' },
  });
  if (!response.ok) throw new Error(`YouTube feed request failed (${response.status})`);
  const xml = await response.text();
  const videos = [...xml.matchAll(/<entry>([\s\S]*?)<\/entry>/g)]
    .map((match) => parseYoutubeFeedEntry(match[1]))
    .filter((item): item is JsonRecord => Boolean(item))
    .slice(0, 15);
  if (!videos.length) throw new Error('YouTube feed did not contain videos.');

  youtubeFeedCache = { expiresAt: now + YOUTUBE_FEED_CACHE_MS, videos };
  return { status: 'ok', videos: videos.slice(0, limit), cached: false };
}

function parseYoutubeFeedEntry(entry: string): JsonRecord | null {
  const videoId = xmlTagValue(entry, 'yt:videoId');
  const title = xmlTagValue(entry, 'title');
  const publishedAt = xmlTagValue(entry, 'published');
  if (!/^[A-Za-z0-9_-]{11}$/.test(videoId) || !title) return null;
  return {
    videoId,
    title,
    publishedAt,
    thumbnailUrl: `https://i.ytimg.com/vi/${videoId}/hqdefault.jpg`,
    destinationUrl: `https://m.youtube.com/watch?v=${videoId}`,
  };
}

function xmlTagValue(xml: string, tagName: string) {
  const escapedTag = tagName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = xml.match(new RegExp(`<${escapedTag}>([\\s\\S]*?)<\\/${escapedTag}>`));
  return decodeXmlText(match?.[1] || '');
}

function decodeXmlText(value: string) {
  return value
    .replace(/^<!\[CDATA\[([\s\S]*)\]\]>$/, '$1')
    .replace(/&#(\d+);/g, (_, number) => String.fromCodePoint(Number(number)))
    .replace(/&#x([0-9a-f]+);/gi, (_, number) => String.fromCodePoint(Number.parseInt(number, 16)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .trim();
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

async function changePin(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const currentPin = validatePin(payload.currentPin);
  const newPin = validatePin(payload.newPin);
  const confirmPin = validatePin(payload.confirmPin);
  if (newPin !== confirmPin) throw userError('pin_mismatch', '新しいPINが一致しません。');
  if (currentPin === newPin) throw userError('pin_unchanged', '現在と異なるPINを設定してください。');

  const profiles = await adminRequest('/rest/v1/profiles', {
    query: { select: 'id,status,legacy_user_id', id: `eq.${userId}`, status: 'eq.active', limit: '1' },
  }) as JsonRecord[];
  const profile = profiles[0];
  if (!profile) throw userError('profile_required', '利用者情報が見つかりません。');

  const authUser = await adminRequest(`/auth/v1/admin/users/${encodeURIComponent(userId)}`) as JsonRecord;
  const email = String(authUser.email || '');
  if (!email) throw userError('profile_required', 'ログイン情報が見つかりません。');

  const legacyUserId = String(profile.legacy_user_id || userId);
  const currentPassword = await deriveAuthPassword(legacyUserId, currentPin);
  try {
    await passwordLogin(email, currentPassword);
  } catch (error) {
    throw userError('current_pin_invalid', '現在のPINが正しくありません。');
  }

  const newPassword = await deriveAuthPassword(legacyUserId, newPin);
  await adminRequest(`/auth/v1/admin/users/${encodeURIComponent(userId)}`, {
    method: 'PUT',
    body: { password: newPassword },
  });

  try {
    await userRequest('/auth/v1/logout', readBearerToken(request), {
      method: 'POST',
      query: { scope: 'global' },
    });
  } catch (error) {
    console.error('Global logout after PIN change failed', error);
  }

  return { status: 'ok' };
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
  const result = await adminRequest('/rest/v1/rpc/get_keiko_scoped_admin_overview', {
    method: 'POST', body: { p_actor_user_id: userId },
  }) as JsonRecord;
  return { status: 'ok', ...result };
}

async function getAdminAccess(request: Request) {
  const userId = await authenticatedUserId(request);
  const result = await adminRequest('/rest/v1/rpc/get_keiko_admin_access', {
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
  const result = await adminRequest('/rest/v1/rpc/get_keiko_scoped_admin_users', {
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
  const result = await adminRequest('/rest/v1/rpc/get_keiko_scoped_admin_user_detail', {
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
  const result = await adminRequest('/rest/v1/rpc/get_keiko_scoped_admin_team_detail', {
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

async function setTeamStaffRole(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const role = text(payload.role);
  if (!['member', 'owner_admin', 'owner', 'manager', 'coach', 'advisor'].includes(role)) {
    throw userError('invalid_request', '権限の指定が正しくありません。');
  }
  return await adminRequest('/rest/v1/rpc/set_keiko_team_staff_role', {
    method: 'POST',
    body: {
      p_actor_user_id: userId,
      p_team_id: requireUuid(payload.teamId, 'チーム'),
      p_target_user_id: requireUuid(payload.userId, 'ユーザー'),
      p_role: role,
    },
  }) as JsonRecord;
}

async function getAdminAnnouncements(request: Request) {
  const userId = await authenticatedUserId(request);
  const result = await adminRequest('/rest/v1/rpc/get_keiko_admin_announcements', {
    method: 'POST', body: { p_actor_user_id: userId },
  }) as JsonRecord;
  return { status: 'ok', ...result };
}

async function saveAdminAnnouncement(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const scope = text(payload.scope) || 'all';
  const priority = text(payload.priority) || 'normal';
  const status = text(payload.status) || 'draft';
  if (!['all', 'student', 'general', 'team_admin', 'team'].includes(scope)) throw userError('invalid_request', '配信対象が正しくありません。');
  if (!['normal', 'important', 'urgent'].includes(priority)) throw userError('invalid_request', '重要度が正しくありません。');
  if (!['draft', 'published', 'archived'].includes(status)) throw userError('invalid_request', '公開状態が正しくありません。');
  const result = await adminRequest('/rest/v1/rpc/save_keiko_admin_announcement', {
    method: 'POST',
    body: {
      p_actor_user_id: userId,
      p_announcement_id: optionalUuid(payload.announcementId, 'お知らせ'),
      p_scope: scope,
      p_team_id: scope === 'team' ? requireUuid(payload.teamId, 'チーム') : null,
      p_title: requiredText(payload.title, 'タイトル', 120),
      p_body: requiredText(payload.body, '本文', 3000),
      p_priority: priority,
      p_action_label: optionalText(payload.actionLabel, 40),
      p_action_url: optionalText(payload.actionUrl, 1000),
      p_starts_at: optionalIsoDate(payload.startsAt, '公開開始日時') || new Date().toISOString(),
      p_expires_at: optionalIsoDate(payload.expiresAt, '公開終了日時'),
      p_status: status,
    },
  }) as JsonRecord;
  return { status: 'ok', ...result };
}

async function getRequestCatalogs(request: Request) {
  await authenticatedUserId(request);
  const [woods, rates] = await Promise.all([
    adminRequest('/rest/v1/bachi_wood_price_options', {
      query: {
        select: 'wood_code,display_name,base_pair_price,standard_length_cm,min_length_cm,max_length_cm,standard_diameter_mm,min_diameter_mm,max_diameter_mm,lead_months,sort_order',
        active: 'eq.true', order: 'sort_order.asc,display_name.asc',
      },
    }) as Promise<JsonRecord[]>,
    adminRequest('/rest/v1/coaching_rate_options', {
      query: {
        select: 'rate_code,display_name,duration_minutes,description,sort_order',
        active: 'eq.true', order: 'sort_order.asc,display_name.asc',
      },
    }) as Promise<JsonRecord[]>,
  ]);
  return {
    status: 'ok',
    woods: woods.map((item) => ({
      woodCode: item.wood_code,
      displayName: item.display_name,
      basePairPrice: item.base_pair_price,
      standardLengthCm: item.standard_length_cm,
      minLengthCm: item.min_length_cm,
      maxLengthCm: item.max_length_cm,
      standardDiameterMm: item.standard_diameter_mm,
      minDiameterMm: item.min_diameter_mm,
      maxDiameterMm: item.max_diameter_mm,
      leadMonths: item.lead_months,
    })),
    coachingRates: rates.map((item) => ({
      rateCode: item.rate_code,
      displayName: item.display_name,
      durationMinutes: item.duration_minutes,
      description: item.description,
    })),
  };
}

async function estimateBachiOrder(request: Request, payload: JsonRecord) {
  await authenticatedUserId(request);
  return { status: 'ok', ...(await calculateBachiQuote(payload)) };
}

async function createBachiOrder(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const quote = await calculateBachiQuote(payload);
  const rows = await adminRequest('/rest/v1/bachi_order_requests', {
    method: 'POST',
    prefer: 'return=representation',
    body: {
      requester_user_id: userId,
      requester_name: requiredText(payload.requesterName, '依頼者名', 100),
      contact_email: requiredEmail(payload.contactEmail),
      wood_code: quote.woodCode,
      wood_name_snapshot: quote.woodName,
      length_cm: quote.lengthCm,
      diameter_mm: quote.diameterMm,
      quantity_pairs: quote.quantityPairs,
      pair_price_estimate: quote.pairPriceEstimate,
      subtotal_estimate: quote.subtotalEstimate,
      tax_estimate: quote.taxEstimate,
      shipping_fee_estimate: quote.shippingFeeEstimate,
      total_estimate: quote.totalEstimate,
      estimated_delivery_date: quote.estimatedDeliveryDate,
      estimate_note: quote.estimateNote,
      status: 'submitted',
    },
  }) as JsonRecord[];
  const order = rows[0];
  if (!order?.id) throw new Error('Bachi order request was not created.');

  const email = await sendBachiEmails(order);
  await adminRequest('/rest/v1/bachi_order_requests', {
    method: 'PATCH',
    query: { id: `eq.${String(order.id)}` },
    body: {
      operator_email_status: email.operator.status,
      customer_email_status: email.customer.status,
      updated_at: new Date().toISOString(),
    },
  });
  return {
    status: 'ok',
    orderId: order.id,
    ...quote,
    emailStatus: combinedEmailStatus(email),
  };
}

async function calculateBachiQuote(payload: JsonRecord) {
  const woodCode = requiredCode(payload.woodCode, '木の種類');
  const woods = await adminRequest('/rest/v1/bachi_wood_price_options', {
    query: { select: '*', wood_code: `eq.${woodCode}`, active: 'eq.true', limit: '1' },
  }) as JsonRecord[];
  const wood = woods[0];
  if (!wood) throw userError('invalid_request', '選択した木材は現在受け付けていません。');

  const lengthCm = requiredNumber(payload.lengthCm, '長さ');
  const diameterMm = requiredNumber(payload.diameterMm, '太さ');
  const quantityPairs = requiredInteger(payload.quantityPairs, '数量', 1, 500);
  const minLength = Number(wood.min_length_cm);
  const maxLength = Number(wood.max_length_cm);
  const minDiameter = Number(wood.min_diameter_mm);
  const maxDiameter = Number(wood.max_diameter_mm);
  if (lengthCm < minLength || lengthCm > maxLength) {
    throw userError('invalid_request', `長さは${minLength}〜${maxLength}cmで入力してください。`);
  }
  if (diameterMm < minDiameter || diameterMm > maxDiameter) {
    throw userError('invalid_request', `太さは${minDiameter}〜${maxDiameter}mmで入力してください。`);
  }

  const [tiers, settingsRows] = await Promise.all([
    adminRequest('/rest/v1/bachi_quantity_price_tiers', {
      query: { select: 'min_pairs,discount_basis_points,label', active: 'eq.true', min_pairs: `lte.${quantityPairs}`, order: 'min_pairs.desc', limit: '1' },
    }) as Promise<JsonRecord[]>,
    adminRequest('/rest/v1/bachi_order_settings', {
      query: { select: 'tax_rate,standard_shipping_fee,default_lead_months', singleton: 'eq.true', limit: '1' },
    }) as Promise<JsonRecord[]>,
  ]);
  const tier = tiers[0] || {};
  const settings = settingsRows[0] || {};
  const standardLength = Number(wood.standard_length_cm);
  const standardDiameter = Number(wood.standard_diameter_mm);
  const adjustedPairPrice = Number(wood.base_pair_price)
    + Math.ceil(Math.abs(lengthCm - standardLength)) * Number(wood.length_surcharge_per_cm || 0)
    + Math.ceil(Math.abs(diameterMm - standardDiameter)) * Number(wood.diameter_surcharge_per_mm || 0);
  const discountBasisPoints = Number(tier.discount_basis_points || 0);
  const pairPriceEstimate = Math.round(adjustedPairPrice * (10000 - discountBasisPoints) / 10000);
  const subtotalEstimate = pairPriceEstimate * quantityPairs;
  const taxEstimate = Math.round(subtotalEstimate * Number(settings.tax_rate ?? 0.10));
  const shippingFeeEstimate = Number(settings.standard_shipping_fee ?? 1000);
  const totalEstimate = subtotalEstimate + taxEstimate + shippingFeeEstimate;
  const leadMonths = Number(wood.lead_months || settings.default_lead_months || 1);
  const estimatedDeliveryDate = addMonthsToJapanDate(leadMonths);
  return {
    woodCode,
    woodName: String(wood.display_name || woodCode),
    lengthCm,
    diameterMm,
    quantityPairs,
    pairPriceEstimate,
    subtotalEstimate,
    taxEstimate,
    shippingFeeEstimate,
    totalEstimate,
    estimatedDeliveryDate,
    estimateNote: `税込・標準送料を含む概算です。最終仕様と在庫を確認後に正式金額をお知らせします。${tier.label ? ` ${String(tier.label)}の料金を適用しています。` : ''}`,
  };
}

async function createCoachingRequest(request: Request, payload: JsonRecord) {
  const userId = await authenticatedUserId(request);
  const teamId = optionalUuid(payload.teamId, 'チーム');
  let teamName = requiredText(payload.teamName, 'チーム名', 120);
  if (teamId) {
    const memberships = await adminRequest('/rest/v1/team_members', {
      query: { select: 'team_id', team_id: `eq.${teamId}`, user_id: `eq.${userId}`, limit: '1' },
    }) as JsonRecord[];
    if (!memberships.length) throw userError('team_forbidden', '所属していないチームは指定できません。');
    const teams = await adminRequest('/rest/v1/teams', {
      query: { select: 'team_name', id: `eq.${teamId}`, status: 'eq.active', limit: '1' },
    }) as JsonRecord[];
    if (!teams.length) throw userError('team_not_found', '選択したチームが見つかりません。');
    teamName = String(teams[0].team_name || teamName);
  }

  const quote = await calculateCoachingQuote(payload);
  const rows = await adminRequest('/rest/v1/coaching_requests', {
    method: 'POST',
    prefer: 'return=representation',
    body: {
      requester_user_id: userId,
      team_id: teamId,
      team_name_snapshot: teamName,
      contact_name: requiredText(payload.contactName, '担当者名', 100),
      contact_email: requiredEmail(payload.contactEmail),
      contact_phone: null,
      location_text: quote.location,
      duration_minutes: quote.durationMinutes,
      lesson_fee_estimate: quote.lessonFeeEstimate,
      transport_fee_estimate: quote.transportFeeEstimate,
      total_estimate: quote.totalEstimate,
      estimate_note: quote.estimateNote,
      rate_type: quote.rateType,
      rate_label: quote.rateLabel,
      lesson_topic: optionalText(payload.lessonTopic, 2000) || '未入力（返信時に確認）',
      participant_count: null,
      notes: '',
      status: 'submitted',
      email_status: 'pending',
    },
  }) as JsonRecord[];
  const coachingRequest = rows[0];
  if (!coachingRequest?.id) throw new Error('Coaching request was not created.');

  const email = await sendCoachingEmails(coachingRequest);
  const emailStatus = combinedEmailStatus(email);
  await adminRequest('/rest/v1/coaching_requests', {
    method: 'PATCH',
    query: { id: `eq.${String(coachingRequest.id)}` },
    body: {
      email_status: emailStatus,
      operator_email_status: email.operator.status,
      customer_email_status: email.customer.status,
      updated_at: new Date().toISOString(),
    },
  });

  return {
    status: 'ok',
    requestId: coachingRequest.id,
    ...quote,
    emailStatus,
  };
}

async function estimateCoachingRequest(request: Request, payload: JsonRecord) {
  await authenticatedUserId(request);
  return { status: 'ok', ...(await calculateCoachingQuote(payload)) };
}

async function calculateCoachingQuote(payload: JsonRecord) {
  const rateType = requiredCode(payload.rateType || 'hourly', '指導時間');
  const rates = await adminRequest('/rest/v1/coaching_rate_options', {
    query: { select: '*', rate_code: `eq.${rateType}`, active: 'eq.true', limit: '1' },
  }) as JsonRecord[];
  const rate = rates[0];
  if (!rate) throw userError('invalid_request', '選択した指導時間は現在受け付けていません。');
  let durationMinutes = Number(rate.duration_minutes || 0);
  let lessonFeeEstimate = Number(rate.fixed_fee || 0);
  if (rateType === 'hourly') {
    durationMinutes = requiredInteger(payload.durationMinutes, '指導時間', 60, 180);
    const increment = Number(rate.billing_increment_minutes || 30);
    if (durationMinutes % increment !== 0) throw userError('invalid_request', `指導時間は${increment}分単位で選択してください。`);
    const hourlyRate = Number(rate.hourly_rate || 0);
    lessonFeeEstimate = Math.ceil(durationMinutes / increment) * Math.round(hourlyRate * increment / 60);
  } else {
    if (!durationMinutes || !lessonFeeEstimate) throw userError('invalid_request', '指導料金の設定が不足しています。');
    const requestedDuration = requiredInteger(payload.durationMinutes, '指導時間', 30, 720);
    if (requestedDuration !== durationMinutes) throw userError('invalid_request', '選択した指導時間が料金設定と一致しません。');
  }
  const location = requiredText(payload.location, '指導場所', 300);
  const transport = await estimateTransitFare(location);
  return {
    rateType,
    rateLabel: String(rate.display_name || rateType),
    durationMinutes,
    location,
    lessonFeeEstimate,
    transportFeeEstimate: transport.fee,
    totalEstimate: transport.fee === null ? null : lessonFeeEstimate + transport.fee,
    estimateNote: `${String(rate.description || '')}${rate.description ? '。' : ''}${transport.note}`,
  };
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

async function estimateTransitFare(destination: string) {
  const apiKey = Deno.env.get('GOOGLE_MAPS_ROUTES_API_KEY') || Deno.env.get('GOOGLE_MAPS_DIRECTIONS_API_KEY') || '';
  const origin = Deno.env.get('KEIKO_COACHING_ORIGIN') || '';
  if (!apiKey || !origin) {
    return { fee: null as number | null, note: '交通費は府中からの経路を確認後に正式確定します。' };
  }

  try {
    const response = await fetch('https://routes.googleapis.com/directions/v2:computeRoutes', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Goog-Api-Key': apiKey,
        'X-Goog-FieldMask': 'routes.travelAdvisory.transitFare,routes.localizedValues.transitFare',
      },
      body: JSON.stringify({
        origin: { address: origin },
        destination: { address: destination },
        travelMode: 'TRANSIT',
        languageCode: 'ja',
      }),
    });
    const result = await response.json() as JsonRecord;
    const routes = Array.isArray(result.routes) ? result.routes as JsonRecord[] : [];
    const advisory = routes[0]?.travelAdvisory as JsonRecord | undefined;
    const fare = advisory?.transitFare as JsonRecord | undefined;
    const currency = String(fare?.currencyCode || '');
    const oneWay = Number(fare?.units || 0) + Number(fare?.nanos || 0) / 1_000_000_000;
    if (!response.ok || currency !== 'JPY' || !Number.isFinite(oneWay) || oneWay <= 0) {
      return { fee: null as number | null, note: '公共交通の運賃を自動取得できなかったため、交通費は確認後に正式確定します。' };
    }
    const roundTrip = Math.round(oneWay * 2);
    return { fee: roundTrip, note: '府中の基準地点から公共交通を利用する往復運賃の概算です。' };
  } catch (error) {
    console.error('Transit fare estimate failed', error);
    return { fee: null as number | null, note: '交通費は府中からの経路を確認後に正式確定します。' };
  }
}

async function sendCoachingEmails(requestRow: JsonRecord) {
  const operatorTo = Deno.env.get('KEIKO_COACHING_TO_EMAIL') || REPORT_EMAIL_TO;
  const durationLabel = requestRow.rate_type === 'hourly'
    ? `${String(requestRow.rate_label || '時間制')}（${formatDurationMinutes(Number(requestRow.duration_minutes || 0))}）`
    : String(requestRow.rate_label || formatDurationMinutes(Number(requestRow.duration_minutes || 0)));
  const lessonFee = Number(requestRow.lesson_fee_estimate || 0).toLocaleString('ja-JP');
  const transportFee = requestRow.transport_fee_estimate === null || requestRow.transport_fee_estimate === undefined
    ? '確認後に確定'
    : `${Number(requestRow.transport_fee_estimate).toLocaleString('ja-JP')}円`;
  const total = requestRow.total_estimate === null || requestRow.total_estimate === undefined
    ? '交通費確認後に確定'
    : `${Number(requestRow.total_estimate).toLocaleString('ja-JP')}円`;
  const body = [
    'KEIKOから指導依頼を受け付けました。',
    '',
    `依頼ID: ${String(requestRow.id || '')}`,
    `チーム: ${String(requestRow.team_name_snapshot || '')}`,
    `担当者: ${String(requestRow.contact_name || '')}`,
    `メール: ${String(requestRow.contact_email || '')}`,
    `場所: ${String(requestRow.location_text || '')}`,
    `希望する指導時間: ${durationLabel}`,
    '',
    `指導料概算: ${lessonFee}円`,
    `交通費概算: ${transportFee}`,
    `概算合計: ${total}`,
    `注記: ${String(requestRow.estimate_note || '')}`,
    '',
    '相談したい指導内容:',
    String(requestRow.lesson_topic || ''),
    '',
    '具体的な日程は、返信メールで候補日をご案内して調整します。',
  ].join('\n');
  const customerBody = [
    `${String(requestRow.contact_name || '')} 様`,
    '',
    'KEIKOから指導依頼を受け付けました。',
    '谷川の指導可能日と正式金額は、内容を確認後に古今樂明座からご連絡します。',
    '',
    body,
  ].join('\n');
  const [operator, customer] = await Promise.all([
    sendTransactionalEmail({
      to: operatorTo,
      replyTo: String(requestRow.contact_email || ''),
      subject: COACHING_EMAIL_SUBJECT,
      body,
    }),
    sendTransactionalEmail({
      to: String(requestRow.contact_email || ''),
      subject: '指導相談を受け付けました【KEIKO】',
      body: customerBody,
    }),
  ]);
  return { operator, customer };
}

async function sendBachiEmails(order: JsonRecord) {
  const operatorTo = Deno.env.get('KEIKO_BACHI_TO_EMAIL') || REPORT_EMAIL_TO;
  const total = `${Number(order.total_estimate || 0).toLocaleString('ja-JP')}円`;
  const body = [
    'KEIKOからカスタムバチの依頼を受け付けました。',
    '',
    `受付番号: ${String(order.id || '')}`,
    `依頼者: ${String(order.requester_name || '')}`,
    `メール: ${String(order.contact_email || '')}`,
    `木材: ${String(order.wood_name_snapshot || '')}`,
    `長さ: ${String(order.length_cm || '')}cm`,
    `太さ: ${String(order.diameter_mm || '')}mm`,
    `数量: ${String(order.quantity_pairs || '')}組`,
    `概算金額: ${total}`,
    `お届け目安: ${String(order.estimated_delivery_date || '')}`,
    `注記: ${String(order.estimate_note || '')}`,
  ].join('\n');
  const customerBody = [
    `${String(order.requester_name || '')} 様`,
    '',
    'カスタムバチのご依頼を受け付けました。',
    '仕様と在庫を確認後、古今樂明座から正式金額をご連絡します。',
    '',
    body,
  ].join('\n');
  const [operator, customer] = await Promise.all([
    sendTransactionalEmail({
      to: operatorTo,
      replyTo: String(order.contact_email || ''),
      subject: BACHI_EMAIL_SUBJECT,
      body,
    }),
    sendTransactionalEmail({
      to: String(order.contact_email || ''),
      subject: 'カスタムバチの依頼を受け付けました【KEIKO】',
      body: customerBody,
    }),
  ]);
  return { operator, customer };
}

async function sendTransactionalEmail(input: { to: string; subject: string; body: string; replyTo?: string }) {
  const apiKey = Deno.env.get('RESEND_API_KEY') || '';
  const from = Deno.env.get('KEIKO_REPORT_FROM_EMAIL') || '';
  if (!apiKey || !from) return { status: 'pending', error: '' };
  try {
    const payload: JsonRecord = { from, to: [input.to], subject: input.subject, text: input.body };
    if (input.replyTo) payload.reply_to = input.replyTo;
    const response = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!response.ok) {
      const remote = await response.text();
      console.error(`Transactional email failed (${response.status}): ${remote.slice(0, 300)}`);
      return { status: 'failed', error: `Resend ${response.status}` };
    }
    return { status: 'sent', error: '' };
  } catch (error) {
    console.error(error);
    return { status: 'failed', error: error instanceof Error ? error.message : 'email_failed' };
  }
}

function combinedEmailStatus(result: { operator: { status: string }; customer: { status: string } }) {
  if (result.operator.status === 'sent' && result.customer.status === 'sent') return 'sent';
  if (result.operator.status === 'pending' || result.customer.status === 'pending') return 'pending';
  return 'failed';
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
    const remoteMessage = String(remote.message || remote.details || '');
    const knownError = [
      ['admin_required', 'このアカウントにはチーム管理権限がありません。'],
      ['operator_required', 'この操作にはKEIKO運営者権限が必要です。'],
      ['team_forbidden', '担当していないチームは表示できません。'],
      ['user_forbidden', '担当チーム外のユーザーは表示できません。'],
      ['membership_not_found', '対象ユーザーのチーム所属が見つかりません。'],
    ].find(([code]) => remoteMessage.includes(code));
    if (knownError) throw userError(knownError[0], knownError[1]);
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
function optionalUuid(value: unknown, label: string) { return text(value) ? requireUuid(value, label) : null; }
function requiredKatakana(value: unknown, label: string) { const valueText = requiredText(value, label, 40); if (!/^[ァ-ヶー\s\u3000]+$/.test(valueText)) throw userError('invalid_request', `${label}はカタカナで入力してください。`); return valueText.replace(/[\s\u3000]+/g, ''); }
function requiredText(value: unknown, label: string, max: number) { const valueText = text(value); if (!valueText) throw userError('invalid_request', `${label}を入力してください。`); if (valueText.length > max) throw userError('invalid_request', `${label}は${max}文字以内で入力してください。`); return valueText; }
function optionalText(value: unknown, max: number) { const valueText = text(value); if (valueText.length > max) throw userError('invalid_request', `${max}文字以内で入力してください。`); return valueText; }
function requiredEmail(value: unknown) { const email = requiredText(value, 'メールアドレス', 254).toLowerCase(); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw userError('invalid_request', 'メールアドレスを正しく入力してください。'); return email; }
function requiredCode(value: unknown, label: string) { const code = text(value).toLowerCase(); if (!/^[a-z0-9_]{1,50}$/.test(code)) throw userError('invalid_request', `${label}の指定が正しくありません。`); return code; }
function requiredNumber(value: unknown, label: string) { const numberValue = Number(value); if (!Number.isFinite(numberValue) || numberValue <= 0) throw userError('invalid_request', `${label}を正しく入力してください。`); return Math.round(numberValue * 10) / 10; }
function requiredInteger(value: unknown, label: string, min: number, max: number) { const numberValue = Number(value); if (!Number.isInteger(numberValue) || numberValue < min || numberValue > max) throw userError('invalid_request', `${label}は${min}〜${max}で入力してください。`); return numberValue; }
function formatDurationMinutes(minutes: number) { const hours = Math.floor(minutes / 60); const remainder = minutes % 60; if (!hours) return `${remainder}分`; return remainder ? `${hours}時間${remainder}分` : `${hours}時間`; }
function addMonthsToJapanDate(months: number) { const current = Object.fromEntries(new Intl.DateTimeFormat('ja-JP', { timeZone: 'Asia/Tokyo', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date()).map((part) => [part.type, part.value])); const year = Number(current.year); const monthIndex = Number(current.month) - 1 + Math.max(1, Math.round(months)); const targetYear = year + Math.floor(monthIndex / 12); const targetMonthIndex = monthIndex % 12; const lastDay = new Date(Date.UTC(targetYear, targetMonthIndex + 1, 0)).getUTCDate(); const targetDay = Math.min(Number(current.day), lastDay); return `${targetYear}-${String(targetMonthIndex + 1).padStart(2, '0')}-${String(targetDay).padStart(2, '0')}`; }
function optionalIsoDate(value: unknown, label: string) { const valueText = text(value); if (!valueText) return null; const date = new Date(valueText); if (Number.isNaN(date.getTime())) throw userError('invalid_request', `${label}を正しく入力してください。`); return date.toISOString(); }
function boundedInteger(value: unknown, min: number, max: number, fallback: number) { const parsed = Number(value); return Number.isInteger(parsed) ? Math.max(min, Math.min(max, parsed)) : fallback; }
function text(value: unknown) { return value === null || value === undefined ? '' : String(value).trim(); }
function generateLegacyId(prefix: string) { return `${prefix}_${crypto.randomUUID().replaceAll('-', '')}`; }
function constantTimeEquals(left: string, right: string) { let difference = left.length ^ right.length; const max = Math.max(left.length, right.length); for (let index = 0; index < max; index += 1) difference |= (left.charCodeAt(index) || 0) ^ (right.charCodeAt(index) || 0); return difference === 0; }
function requireEnv(name: string) { const value = Deno.env.get(name); if (!value) throw new Error(`Missing environment variable: ${name}`); return value; }
function userError(code: string, publicMessage: string) { return Object.assign(new Error(publicMessage), { code, publicMessage }); }
function invalidCredentials() { return userError('invalid_credentials', '名前またはPINが違います。'); }
function json(payload: unknown, status = 200) { return new Response(JSON.stringify(payload), { status, headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' } }); }
