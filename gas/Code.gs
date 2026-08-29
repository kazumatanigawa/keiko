const CONFIG = Object.freeze({
  propertyNames: Object.freeze({
    supabaseUrl: 'SUPABASE_URL',
    secretKey: 'SUPABASE_SECRET_KEY',
    publishableKey: 'SUPABASE_PUBLISHABLE_KEY',
    authPepper: 'KEIKO_AUTH_PEPPER',
    registrationCode: 'KEIKO_REGISTRATION_CODE',
  }),
  maxTextLength: 1000,
  maxNoteTitleLength: 120,
  maxNoteBodyLength: 5000,
  goodNewsLimit: 50,
  logsLimit: 500,
  notesLimit: 100,
});

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  if ((params.action || 'health') === 'health') {
    return jsonOutput_({ status: 'ok', service: 'keiko-api' });
  }
  if (params.action === 'getTeams') {
    return routeRequest_(params);
  }
  return jsonOutput_({ status: 'error', code: 'post_required', message: 'POSTでアクセスしてください。' });
}

function doPost(e) {
  const payload = parseJsonSafe_(e && e.postData ? e.postData.contents : '');
  if (!payload || typeof payload !== 'object') {
    return jsonOutput_({ status: 'error', code: 'invalid_request', message: 'リクエスト形式が正しくありません。' });
  }
  return routeRequest_(payload);
}

function testSupabaseConnection() {
  const result = handleGetTeams_();
  console.log(JSON.stringify(result));
  return result;
}

function routeRequest_(params) {
  try {
    const action = cleanText_(params.action);
    if (!action) throw userError_('invalid_request', '操作が指定されていません。');

    switch (action) {
      case 'getTeams': return jsonOutput_(handleGetTeams_());
      case 'register': return jsonOutput_(handleRegister_(params));
      case 'login': return jsonOutput_(handleLogin_(params));
      case 'refreshSession': return jsonOutput_(handleRefreshSession_(params));
      case 'saveLog': return jsonOutput_(handleSaveLog_(params));
      case 'getGoodNews': return jsonOutput_(handleGetGoodNews_(params));
      case 'getLogs': return jsonOutput_(handleGetLogs_(params));
      case 'getNotes': return jsonOutput_(handleGetNotes_(params));
      case 'saveNote': return jsonOutput_(handleSaveNote_(params));
      case 'addNoteComment': return jsonOutput_(handleAddNoteComment_(params));
      case 'updateNoteComment': return jsonOutput_(handleUpdateNoteComment_(params));
      case 'deleteNoteComment': return jsonOutput_(handleDeleteNoteComment_(params));
      default: throw userError_('unsupported_action', 'この操作には対応していません。');
    }
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonOutput_({
      status: 'error',
      code: error && error.userCode ? error.userCode : 'server_error',
      message: error && error.userMessage ? error.userMessage : '処理に失敗しました。時間をおいて再度お試しください。',
    });
  }
}

function handleGetTeams_() {
  const teams = adminRequest_('/rest/v1/teams', {
    query: {
      select: 'id,team_name,team_type',
      status: 'eq.active',
      order: 'team_name.asc',
    },
  });
  return {
    status: 'ok',
    teams: teams.map(function(team) {
      return { teamId: team.id, teamName: team.team_name, teamType: team.team_type || 'general' };
    }),
  };
}

function handleLogin_(params) {
  const name = requireText_(params.name, '名前', 80);
  const pin = validatePin_(params.pin);
  const profile = findProfileForLogin_(name);
  if (!profile || profile.status !== 'active') throw invalidCredentials_();

  const authUser = adminRequest_('/auth/v1/admin/users/' + encodeURIComponent(profile.id));
  if (!authUser || !authUser.email) throw invalidCredentials_();
  const password = deriveAuthPassword_(profile.legacy_user_id || profile.id, pin);
  const session = authPasswordLogin_(authUser.email, password);
  return buildSessionResponse_(session, profile.id);
}

function handleRefreshSession_(params) {
  const refreshToken = requireText_(params.refreshToken, '更新トークン', 4096);
  const session = publicRequest_('/auth/v1/token', {
    method: 'post',
    query: { grant_type: 'refresh_token' },
    body: { refresh_token: refreshToken },
    authMode: 'none',
    authErrorCode: 'session_expired',
  });
  if (!session || !session.user || !session.user.id) {
    throw userError_('session_expired', 'ログインの有効期限が切れました。');
  }
  return buildSessionResponse_(session, session.user.id);
}

function handleRegister_(params) {
  const settings = getSettings_();
  if (!settings.registrationCode || !constantTimeEquals_(cleanText_(params.registrationCode), settings.registrationCode)) {
    throw userError_('registration_denied', '登録コードが正しくありません。');
  }

  const lastName = requireText_(params.lastName, '姓', 40);
  const firstName = requireText_(params.firstName, '名', 40);
  const lastNameKana = requireKatakana_(params.lastNameKana, '姓の読み');
  const firstNameKana = requireKatakana_(params.firstNameKana, '名の読み');
  const displayName = normalizeDisplayName_(params.name || lastName + firstName);
  const loginId = normalizeLoginId_(displayName);
  const pin = validatePin_(params.pin);
  const teamMode = cleanText_(params.teamMode);
  if (teamMode !== 'existing' && teamMode !== 'launch') {
    throw userError_('invalid_request', '所属チームの指定が正しくありません。');
  }
  if (findProfileForLogin_(loginId)) {
    return { status: 'exists' };
  }

  let team = null;
  let teamType = 'general';
  let teamName = '';
  if (teamMode === 'existing') {
    const teamId = requireUuid_(params.teamId, 'チーム');
    const teams = adminRequest_('/rest/v1/teams', {
      query: { select: 'id,team_name,team_type,status', id: 'eq.' + teamId, status: 'eq.active', limit: '1' },
    });
    if (!teams.length) throw userError_('team_not_found', '選択したチームが見つかりません。');
    team = teams[0];
    teamType = normalizeTeamType_(team.team_type);
    teamName = team.team_name;
  } else {
    teamName = requireText_(params.teamName, 'チーム名', 100);
    teamType = normalizeTeamType_(params.teamType);
    const matches = adminRequest_('/rest/v1/teams', {
      query: { select: 'id', team_name: 'eq.' + teamName, status: 'eq.active', limit: '1' },
    });
    if (matches.length) return { status: 'team_exists' };
  }

  const legacyUserId = generateLegacyId_('usr');
  const email = legacyUserId.toLowerCase() + '@auth.keiko.invalid';
  const password = deriveAuthPassword_(legacyUserId, pin);
  let authUserId = '';

  try {
    const authUser = adminRequest_('/auth/v1/admin/users', {
      method: 'post',
      body: { email: email, password: password, email_confirm: true, user_metadata: { display_name: displayName } },
    });
    authUserId = authUser.id;
    if (!authUserId) throw new Error('Supabase Auth user was not created.');

    const userType = teamType === 'student' ? 'student' : 'general';
    adminRequest_('/rest/v1/profiles', {
      method: 'post',
      prefer: 'return=representation',
      body: {
        id: authUserId,
        user_type: userType,
        display_name: displayName,
        login_id: loginId,
        status: 'active',
        legacy_user_id: legacyUserId,
        name: displayName,
        last_name: lastName,
        first_name: firstName,
        last_name_kana: lastNameKana,
        first_name_kana: firstNameKana,
      },
    });

    if (teamMode === 'launch') {
      const createdTeams = adminRequest_('/rest/v1/teams', {
        method: 'post',
        prefer: 'return=representation',
        body: {
          team_name: teamName,
          team_type: teamType,
          audience_type: teamType,
          owner_user_id: authUserId,
          status: 'active',
          legacy_team_id: generateLegacyId_('team'),
        },
      });
      team = createdTeams[0];
    }

    adminRequest_('/rest/v1/team_members', {
      method: 'post',
      body: { team_id: team.id, user_id: authUserId, team_role: teamMode === 'launch' ? 'owner_admin' : 'member' },
    });

    const profilePath = userType === 'student' ? '/rest/v1/student_profiles' : '/rest/v1/general_profiles';
    const profileBody = userType === 'student'
      ? { user_id: authUserId, school_name: teamName, grade: '', role_label: '', term: '' }
      : { user_id: authUserId, category: teamName, bio: '' };
    adminRequest_(profilePath, { method: 'post', body: profileBody });

    const session = authPasswordLogin_(email, password);
    return buildSessionResponse_(session, authUserId);
  } catch (error) {
    if (authUserId) {
      try { adminRequest_('/auth/v1/admin/users/' + encodeURIComponent(authUserId), { method: 'delete' }); } catch (cleanupError) { console.error(cleanupError); }
    }
    throw error;
  }
}

function handleSaveLog_(params) {
  const auth = authenticateRequest_(params);
  const practiceDate = validateDate_(params.date);
  const condition = Number(params.cond);
  if (!Number.isInteger(condition) || condition < 1 || condition > 5) {
    throw userError_('invalid_request', 'コンディションは1から5で指定してください。');
  }
  const learning = requireText_(params.learning, '今日の学び', CONFIG.maxTextLength);
  const nextAction = requireText_(params.next, '次の稽古でやること', CONFIG.maxTextLength);
  const goodNew = requireText_(params.goodNew, 'Good&New', CONFIG.maxTextLength);
  const achievementStatus = cleanText_(params.achievementStatus);
  if (achievementStatus && ['done', 'pending'].indexOf(achievementStatus) === -1) {
    throw userError_('invalid_request', '達成状況が正しくありません。');
  }
  const whyMissed = optionalText_(params.whyMissed, 'できなかった理由', CONFIG.maxTextLength);
  const retryPlan = optionalText_(params.retryPlan, '次どうする', CONFIG.maxTextLength);
  if (achievementStatus === 'pending' && (!whyMissed || !retryPlan)) {
    throw userError_('invalid_request', 'できなかった理由と次の行動を入力してください。');
  }
  const requestId = requireRequestId_(params.requestId);
  const sourceRowId = 'app:' + requestId;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = userRequest_('/rest/v1/practice_logs', auth.accessToken, {
      query: { select: 'id', user_id: 'eq.' + auth.user.id, source_system: 'eq.keiko_app', legacy_source_row_id: 'eq.' + sourceRowId, limit: '1' },
    });
    if (existing.length) return { status: 'ok', id: existing[0].id, duplicate: true };

    const rows = userRequest_('/rest/v1/practice_logs', auth.accessToken, {
      method: 'post',
      prefer: 'return=representation',
      body: {
        user_id: auth.user.id,
        team_id: auth.team.id,
        practice_date: practiceDate,
        condition: condition,
        learning: learning,
        next_action: nextAction,
        good_new: goodNew,
        memo: '',
        visibility: 'team',
        source_system: 'keiko_app',
        legacy_source_row_id: sourceRowId,
        achievement_status: achievementStatus || null,
        why_missed: achievementStatus === 'pending' ? whyMissed : null,
        retry_plan: achievementStatus === 'pending' ? retryPlan : null,
        display_name_snapshot: auth.profile.display_name,
        grade_snapshot: auth.studentProfile ? auth.studentProfile.grade || '' : '',
        term_snapshot: auth.studentProfile ? auth.studentProfile.term || '' : '',
      },
    });
    return { status: 'ok', id: rows[0] ? rows[0].id : '' };
  } finally {
    lock.releaseLock();
  }
}

function handleGetLogs_(params) {
  const auth = authenticateRequest_(params);
  const rows = userRequest_('/rest/v1/practice_logs', auth.accessToken, {
    query: {
      select: 'id,practice_date,condition,learning,next_action,good_new,achievement_status,why_missed,retry_plan,created_at,updated_at',
      user_id: 'eq.' + auth.user.id,
      order: 'practice_date.desc,created_at.desc',
      limit: String(CONFIG.logsLimit),
    },
  });
  const response = sessionContext_(auth);
  response.status = 'ok';
  response.logs = rows.map(mapLog_);
  return response;
}

function handleGetGoodNews_(params) {
  const auth = authenticateRequest_(params);
  const rows = userRequest_('/rest/v1/practice_logs', auth.accessToken, {
    query: {
      select: 'id,user_id,practice_date,good_new,display_name_snapshot,created_at',
      team_id: 'eq.' + auth.team.id,
      visibility: 'eq.team',
      good_new: 'not.is.null',
      order: 'practice_date.desc,created_at.desc',
      limit: String(CONFIG.goodNewsLimit),
    },
  });
  return {
    status: 'ok',
    items: rows.filter(function(row) { return cleanText_(row.good_new); }).map(function(row) {
      return {
        id: row.id,
        userId: row.user_id,
        name: row.display_name_snapshot || 'メンバー',
        goodNew: row.good_new,
        date: row.practice_date,
        createdAt: row.created_at,
      };
    }),
  };
}

function handleGetNotes_(params) {
  const auth = authenticateRequest_(params);
  return { status: 'ok', notes: fetchNotes_(auth) };
}

function handleSaveNote_(params) {
  const auth = authenticateRequest_(params);
  const title = requireText_(params.title, 'タイトル', CONFIG.maxNoteTitleLength);
  const body = requireText_(params.body, '内容', CONFIG.maxNoteBodyLength);
  userRequest_('/rest/v1/team_notes', auth.accessToken, {
    method: 'post',
    body: {
      team_id: auth.team.id,
      author_user_id: auth.user.id,
      author_name_snapshot: auth.profile.display_name,
      title: title,
      body: body,
      status: 'active',
    },
  });
  return { status: 'ok', notes: fetchNotes_(auth) };
}

function handleAddNoteComment_(params) {
  const auth = authenticateRequest_(params);
  const noteId = requireUuid_(params.noteId, 'ノート');
  const body = requireText_(params.body, '補足内容', CONFIG.maxNoteBodyLength);
  assertActiveNote_(auth, noteId);
  userRequest_('/rest/v1/team_note_comments', auth.accessToken, {
    method: 'post',
    body: {
      note_id: noteId,
      team_id: auth.team.id,
      author_user_id: auth.user.id,
      author_name_snapshot: auth.profile.display_name,
      body: body,
      status: 'active',
    },
  });
  return { status: 'ok', notes: fetchNotes_(auth) };
}

function handleUpdateNoteComment_(params) {
  const auth = authenticateRequest_(params);
  const commentId = requireUuid_(params.commentId, '補足メモ');
  const body = requireText_(params.body, '補足内容', CONFIG.maxNoteBodyLength);
  assertOwnedComment_(auth, commentId);
  userRequest_('/rest/v1/team_note_comments', auth.accessToken, {
    method: 'patch',
    query: { id: 'eq.' + commentId, author_user_id: 'eq.' + auth.user.id },
    body: { body: body, updated_at: new Date().toISOString() },
  });
  return { status: 'ok', notes: fetchNotes_(auth) };
}

function handleDeleteNoteComment_(params) {
  const auth = authenticateRequest_(params);
  const commentId = requireUuid_(params.commentId, '補足メモ');
  assertOwnedComment_(auth, commentId);
  userRequest_('/rest/v1/team_note_comments', auth.accessToken, {
    method: 'patch',
    query: { id: 'eq.' + commentId, author_user_id: 'eq.' + auth.user.id },
    body: { status: 'deleted', updated_at: new Date().toISOString() },
  });
  return { status: 'ok', notes: fetchNotes_(auth) };
}

function fetchNotes_(auth) {
  const notes = userRequest_('/rest/v1/team_notes', auth.accessToken, {
    query: {
      select: 'id,author_user_id,author_name_snapshot,title,body,created_at,updated_at',
      team_id: 'eq.' + auth.team.id,
      status: 'eq.active',
      order: 'created_at.desc',
      limit: String(CONFIG.notesLimit),
    },
  });
  if (!notes.length) return [];

  const noteIds = notes.map(function(note) { return note.id; }).join(',');
  const comments = userRequest_('/rest/v1/team_note_comments', auth.accessToken, {
    query: {
      select: 'id,note_id,author_user_id,author_name_snapshot,body,created_at,updated_at',
      note_id: 'in.(' + noteIds + ')',
      status: 'eq.active',
      order: 'created_at.asc',
    },
  });
  const commentsByNote = {};
  comments.forEach(function(comment) {
    if (!commentsByNote[comment.note_id]) commentsByNote[comment.note_id] = [];
    commentsByNote[comment.note_id].push({
      commentId: comment.id,
      noteId: comment.note_id,
      authorUserId: comment.author_user_id,
      authorName: comment.author_name_snapshot || 'メンバー',
      body: comment.body,
      createdAt: comment.created_at,
      updatedAt: comment.updated_at,
      isEdited: comment.updated_at && comment.created_at && comment.updated_at !== comment.created_at,
      canEdit: comment.author_user_id === auth.user.id,
      canDelete: comment.author_user_id === auth.user.id,
    });
  });
  return notes.map(function(note) {
    return {
      noteId: note.id,
      authorUserId: note.author_user_id,
      authorName: note.author_name_snapshot || 'メンバー',
      title: note.title,
      body: note.body,
      createdAt: note.created_at,
      updatedAt: note.updated_at,
      comments: commentsByNote[note.id] || [],
    };
  });
}

function assertActiveNote_(auth, noteId) {
  const notes = userRequest_('/rest/v1/team_notes', auth.accessToken, {
    query: { select: 'id', id: 'eq.' + noteId, team_id: 'eq.' + auth.team.id, status: 'eq.active', limit: '1' },
  });
  if (!notes.length) throw userError_('not_found', '対象のノートが見つかりません。');
}

function assertOwnedComment_(auth, commentId) {
  const comments = userRequest_('/rest/v1/team_note_comments', auth.accessToken, {
    query: { select: 'id', id: 'eq.' + commentId, team_id: 'eq.' + auth.team.id, author_user_id: 'eq.' + auth.user.id, status: 'eq.active', limit: '1' },
  });
  if (!comments.length) throw userError_('forbidden', 'この補足メモは編集できません。');
}

function authenticateRequest_(params) {
  const accessToken = requireText_(params.accessToken, 'アクセストークン', 8192);
  let user;
  try {
    user = publicRequest_('/auth/v1/user', { accessToken: accessToken });
  } catch (error) {
    throw userError_('session_expired', 'ログインの有効期限が切れました。');
  }
  if (!user || !user.id) throw userError_('auth_required', 'ログインしてください。');

  const profiles = userRequest_('/rest/v1/profiles', accessToken, {
    query: { select: 'id,user_type,display_name,login_id,status,legacy_user_id', id: 'eq.' + user.id, status: 'eq.active', limit: '1' },
  });
  if (!profiles.length) throw userError_('auth_required', '利用可能なユーザー情報がありません。');

  const memberships = userRequest_('/rest/v1/team_members', accessToken, {
    query: { select: 'team_id,team_role', user_id: 'eq.' + user.id },
  });
  if (!memberships.length) throw userError_('team_required', '所属チームが設定されていません。');
  const requestedTeamId = cleanText_(params.teamId);
  const membership = memberships.find(function(item) { return item.team_id === requestedTeamId; }) || memberships[0];
  const teams = userRequest_('/rest/v1/teams', accessToken, {
    query: { select: 'id,team_name,team_type,status', id: 'eq.' + membership.team_id, status: 'eq.active', limit: '1' },
  });
  if (!teams.length) throw userError_('team_required', '所属チームを確認できません。');

  let studentProfile = null;
  if (profiles[0].user_type === 'student') {
    const studentProfiles = userRequest_('/rest/v1/student_profiles', accessToken, {
      query: { select: 'grade,role_label,term', user_id: 'eq.' + user.id, limit: '1' },
    });
    studentProfile = studentProfiles[0] || null;
  }
  return { accessToken: accessToken, user: user, profile: profiles[0], membership: membership, team: teams[0], studentProfile: studentProfile };
}

function buildSessionResponse_(session, userId) {
  const context = loadSessionContextAdmin_(userId);
  return Object.assign({
    status: 'ok',
    accessToken: session.access_token,
    refreshToken: session.refresh_token,
    expiresIn: Number(session.expires_in || 3600),
    expiresAt: Number(session.expires_at || Math.floor(Date.now() / 1000) + Number(session.expires_in || 3600)),
  }, context);
}

function loadSessionContextAdmin_(userId) {
  const profiles = adminRequest_('/rest/v1/profiles', {
    query: { select: 'id,user_type,display_name,status', id: 'eq.' + userId, status: 'eq.active', limit: '1' },
  });
  if (!profiles.length) throw userError_('auth_required', '利用可能なユーザー情報がありません。');
  const memberships = adminRequest_('/rest/v1/team_members', {
    query: { select: 'team_id,team_role', user_id: 'eq.' + userId, limit: '1' },
  });
  if (!memberships.length) throw userError_('team_required', '所属チームが設定されていません。');
  const teams = adminRequest_('/rest/v1/teams', {
    query: { select: 'id,team_name,team_type', id: 'eq.' + memberships[0].team_id, limit: '1' },
  });
  if (!teams.length) throw userError_('team_required', '所属チームを確認できません。');
  return {
    userId: userId,
    name: profiles[0].display_name,
    teamId: teams[0].id,
    group: teams[0].team_name,
    teamType: teams[0].team_type,
    userType: profiles[0].user_type,
    role: memberships[0].team_role,
  };
}

function sessionContext_(auth) {
  return {
    userId: auth.user.id,
    name: auth.profile.display_name,
    teamId: auth.team.id,
    group: auth.team.team_name,
    teamType: auth.team.team_type,
    userType: auth.profile.user_type,
    role: auth.membership.team_role,
  };
}

function findProfileForLogin_(name) {
  const loginId = normalizeLoginId_(name);
  if (!loginId) return null;
  const rows = adminRequest_('/rest/v1/profiles', {
    query: { select: 'id,display_name,login_id,status,legacy_user_id', login_id: 'eq.' + loginId, limit: '2' },
  });
  return rows.length === 1 ? rows[0] : null;
}

function authPasswordLogin_(email, password) {
  try {
    return publicRequest_('/auth/v1/token', {
      method: 'post',
      query: { grant_type: 'password' },
      body: { email: email, password: password },
      authMode: 'none',
      authErrorCode: 'invalid_credentials',
    });
  } catch (error) {
    if (error && error.userCode === 'invalid_credentials') throw invalidCredentials_();
    throw error;
  }
}

function adminRequest_(path, options) {
  options = options || {};
  options.authMode = 'admin';
  return supabaseRequest_(path, options);
}

function publicRequest_(path, options) {
  options = options || {};
  options.authMode = options.authMode || 'public';
  return supabaseRequest_(path, options);
}

function userRequest_(path, accessToken, options) {
  options = options || {};
  options.authMode = 'user';
  options.accessToken = accessToken;
  return supabaseRequest_(path, options);
}

function supabaseRequest_(path, options) {
  const settings = getSettings_();
  const method = String(options.method || 'get').toLowerCase();
  const authMode = options.authMode || 'public';
  const key = authMode === 'admin' ? settings.secretKey : settings.publishableKey;
  const headers = { apikey: key, Accept: 'application/json' };
  if (authMode === 'admin' && key.indexOf('eyJ') === 0) headers.Authorization = 'Bearer ' + key;
  if (authMode === 'user' || options.accessToken) headers.Authorization = 'Bearer ' + options.accessToken;
  if (options.prefer) headers.Prefer = options.prefer;

  // buildUrl_ already percent-encodes PostgREST filters, including Japanese names.
  const fetchOptions = { method: method, headers: headers, muteHttpExceptions: true, escaping: false };
  if (options.body !== undefined) {
    headers['Content-Type'] = 'application/json';
    fetchOptions.payload = JSON.stringify(options.body);
  }
  const response = UrlFetchApp.fetch(buildUrl_(settings.supabaseUrl, path, options.query), fetchOptions);
  const statusCode = response.getResponseCode();
  const text = response.getContentText();
  const data = text ? parseJsonSafe_(text) : null;
  if (statusCode < 200 || statusCode >= 300) {
    console.error('Supabase request failed: ' + method.toUpperCase() + ' ' + path + ' [' + statusCode + '] ' + sanitizeRemoteError_(data));
    if (options.authErrorCode && statusCode >= 400 && statusCode < 500) {
      throw userError_(options.authErrorCode, options.authErrorCode === 'invalid_credentials' ? '名前またはPINが違います。' : 'ログインの有効期限が切れました。');
    }
    if (statusCode === 401 || statusCode === 403) {
      throw userError_(options.authErrorCode || 'auth_required', options.authErrorCode === 'invalid_credentials' ? '名前またはPINが違います。' : 'ログインの有効期限が切れました。');
    }
    if (statusCode === 409 || (data && data.code === '23505')) {
      throw userError_('conflict', '同じ情報がすでに登録されています。');
    }
    throw new Error('Supabase request failed with status ' + statusCode + '.');
  }
  return data === null ? {} : data;
}

function getSettings_() {
  const properties = PropertiesService.getScriptProperties();
  const settings = {
    supabaseUrl: cleanText_(properties.getProperty(CONFIG.propertyNames.supabaseUrl)).replace(/\/+$/, ''),
    secretKey: cleanText_(properties.getProperty(CONFIG.propertyNames.secretKey)),
    publishableKey: cleanText_(properties.getProperty(CONFIG.propertyNames.publishableKey)),
    authPepper: cleanText_(properties.getProperty(CONFIG.propertyNames.authPepper)),
    registrationCode: cleanText_(properties.getProperty(CONFIG.propertyNames.registrationCode)),
  };
  ['supabaseUrl', 'secretKey', 'publishableKey', 'authPepper'].forEach(function(key) {
    if (!settings[key]) throw new Error('Missing Script Property: ' + CONFIG.propertyNames[key]);
  });
  return settings;
}

function buildUrl_(baseUrl, path, query) {
  let url = baseUrl + (path.charAt(0) === '/' ? path : '/' + path);
  if (!query) return url;
  const parts = [];
  Object.keys(query).forEach(function(key) {
    if (query[key] === undefined || query[key] === null || query[key] === '') return;
    parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(query[key])));
  });
  return parts.length ? url + '?' + parts.join('&') : url;
}

function deriveAuthPassword_(legacyUserId, pin) {
  const pepper = getSettings_().authPepper;
  const bytes = Utilities.computeHmacSha256Signature(String(legacyUserId) + ':' + pin, pepper);
  return bytes.map(function(value) { return ('0' + ((value + 256) % 256).toString(16)).slice(-2); }).join('');
}

function mapLog_(row) {
  return {
    id: row.id,
    date: row.practice_date,
    cond: row.condition,
    learning: row.learning || '',
    next: row.next_action || '',
    goodNew: row.good_new || '',
    achievementStatus: row.achievement_status || '',
    whyMissed: row.why_missed || '',
    retryPlan: row.retry_plan || '',
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

function normalizeLoginId_(value) {
  return cleanText_(value).replace(/[\s\u3000]+/g, '').toLowerCase();
}

function normalizeDisplayName_(value) {
  return requireText_(value, '名前', 80).replace(/[\s\u3000]+/g, '');
}

function normalizeTeamType_(value) {
  return cleanText_(value) === 'student' ? 'student' : 'general';
}

function validatePin_(value) {
  const pin = cleanText_(value);
  if (!/^\d{4}$/.test(pin)) throw userError_('invalid_request', 'PINは4桁の数字で入力してください。');
  return pin;
}

function validateDate_(value) {
  const date = cleanText_(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || isNaN(new Date(date + 'T00:00:00Z').getTime())) {
    throw userError_('invalid_request', '日付が正しくありません。');
  }
  return date;
}

function requireKatakana_(value, label) {
  const text = requireText_(value, label, 40);
  if (!/^[ァ-ヶー\s\u3000]+$/.test(text)) throw userError_('invalid_request', label + 'はカタカナで入力してください。');
  return text.replace(/[\s\u3000]+/g, '');
}

function requireUuid_(value, label) {
  const text = cleanText_(value);
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(text)) {
    throw userError_('invalid_request', label + 'の指定が正しくありません。');
  }
  return text;
}

function requireRequestId_(value) {
  const text = cleanText_(value);
  if (!/^[A-Za-z0-9_-]{16,80}$/.test(text)) throw userError_('invalid_request', '送信IDが正しくありません。');
  return text;
}

function requireText_(value, label, maxLength) {
  const text = cleanText_(value);
  if (!text) throw userError_('invalid_request', label + 'を入力してください。');
  if (text.length > maxLength) throw userError_('invalid_request', label + 'は' + maxLength + '文字以内で入力してください。');
  return text;
}

function optionalText_(value, label, maxLength) {
  const text = cleanText_(value);
  if (text.length > maxLength) throw userError_('invalid_request', label + 'は' + maxLength + '文字以内で入力してください。');
  return text;
}

function cleanText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function constantTimeEquals_(left, right) {
  left = String(left || '');
  right = String(right || '');
  let difference = left.length ^ right.length;
  const maxLength = Math.max(left.length, right.length);
  for (let index = 0; index < maxLength; index += 1) {
    difference |= (left.charCodeAt(index) || 0) ^ (right.charCodeAt(index) || 0);
  }
  return difference === 0;
}

function generateLegacyId_(prefix) {
  return prefix + '_' + Utilities.getUuid().replace(/-/g, '');
}

function parseJsonSafe_(text) {
  try { return JSON.parse(text); } catch (error) { return null; }
}

function sanitizeRemoteError_(data) {
  if (!data || typeof data !== 'object') return 'unknown error';
  return cleanText_(data.code || data.error_code || data.error || 'remote error').slice(0, 100);
}

function invalidCredentials_() {
  return userError_('invalid_credentials', '名前またはPINが違います。');
}

function userError_(code, message) {
  const error = new Error(message);
  error.userCode = code;
  error.userMessage = message;
  return error;
}

function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
