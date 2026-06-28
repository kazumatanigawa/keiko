const CONFIG = {
  schemaVersion: '2026-06-28-01',
  spreadsheetIdProperty: 'SPREADSHEET_ID',
  setupVersionProperty: 'KEIKO_SETUP_VERSION',
  setupAtProperty: 'KEIKO_SETUP_COMPLETED_AT',
  sheetNames: {
    users: 'ユーザー',
    roster: '生徒一覧',
    logs: '稽古ログ',
    members: '部員一覧',
    teams: 'チーム一覧',
    review: '移行要確認',
  },
  userHeaders: ['name', 'pin', 'userId', 'teamId', 'group', 'role', 'createdAt', 'updatedAt'],
  teamHeaders: ['teamId', 'teamName', 'createdByUserId', 'adminUserIds', 'createdAt', 'status'],
  rosterHeaders: ['userId', 'teamId', 'group', 'name', 'displayName', 'grade', 'term', 'role', 'updatedAt'],
  logHeaders: ['userId', 'teamId', 'group', 'name', 'displayName', 'date', 'cond', 'learning', 'next', 'goodNew', 'achievementStatus', 'whyMissed', 'retryPlan', 'grade', 'term', 'createdAt', 'updatedAt'],
  reviewHeaders: ['timestamp', 'sheetName', 'rowNumber', 'reason', 'name', 'group', 'candidateUserIds', 'candidateNames'],
  memberViewHeaders: ['所属', '学年', '名前', '期', '役職', 'userId', 'teamId'],
  sourceBackupSheets: ['ユーザー', '生徒一覧', '稽古ログ', '部員一覧'],
};

const FIELD_ALIASES = {
  name: ['name', '名前', '氏名'],
  pin: ['pin', 'PIN', '暗証番号', 'パスワード'],
  userId: ['userId', 'ユーザーID', 'user_id', 'id', 'ID'],
  teamId: ['teamId', 'チームID', 'team_id'],
  group: ['group', '所属', '所属チーム', '学校', 'チーム'],
  role: ['role', '役割', '権限', 'teamRole'],
  createdAt: ['createdAt', 'created_at', '作成日時'],
  updatedAt: ['updatedAt', 'updated_at', '更新日時'],
  teamName: ['teamName', 'チーム名', 'group', '所属'],
  createdByUserId: ['createdByUserId', '作成者userId'],
  adminUserIds: ['adminUserIds', '管理者ID一覧'],
  status: ['status', '状態'],
  displayName: ['displayName', '表示名', 'name', '名前', '氏名'],
  grade: ['grade', '学年', '現在学年'],
  term: ['term', '期'],
  date: ['date', '日付'],
  cond: ['cond', 'condition', 'コンディション'],
  learning: ['learning', '学び'],
  next: ['next', 'nextAction', '次の稽古でやること'],
  goodNew: ['goodNew', 'Good&New', 'goodnew', 'GoodNew'],
  achievementStatus: ['achievementStatus', '達成状況', '前回達成チェック'],
  whyMissed: ['whyMissed', 'できなかった理由'],
  retryPlan: ['retryPlan', '次どうする'],
};

function doGet(e) {
  return routeRequest_(e && e.parameter ? e.parameter : {}, 'GET');
}

function doPost(e) {
  const payload = parseJsonSafe_(e && e.postData ? e.postData.contents : '{}');
  return routeRequest_(payload || {}, 'POST');
}

function routeRequest_(params, method) {
  try {
    const context = ensureSystemReady_();
    const action = params.action || inferLegacyAction_(params, method);

    switch (action) {
      case 'getTeams':
        return jsonOutput_({ status: 'ok', teams: listTeams_(context) });
      case 'register':
        return jsonOutput_(handleRegister_(context, params));
      case 'login':
        return jsonOutput_(handleLogin_(context, params));
      case 'saveLog':
        return jsonOutput_(handleSaveLog_(context, params));
      case 'getGoodNews':
        return jsonOutput_(handleGetGoodNews_(context, params));
      case 'getLogs':
        return jsonOutput_(handleGetLogs_(context, params));
      default:
        return jsonOutput_({ status: 'error', message: 'Unsupported action.' });
    }
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonOutput_({
      status: 'error',
      message: error && error.message ? error.message : 'Unexpected error.',
    });
  }
}

function inferLegacyAction_(params, method) {
  if (params.action) return params.action;
  if (method === 'GET' && params.name && params.pin) return 'getLogs';
  if (method === 'GET') return 'getTeams';
  return '';
}

function handleRegister_(context, params) {
  const name = cleanText_(params.name);
  const pin = cleanText_(params.pin);
  const teamMode = cleanText_(params.teamMode);

  if (!name) return { status: 'error', message: 'name is required.' };
  if (!/^\d{4}$/.test(pin)) return { status: 'error', message: 'pin must be 4 digits.' };
  if (!['existing', 'launch'].includes(teamMode)) return { status: 'error', message: 'teamMode is invalid.' };

  const usersTable = readTable_(context.sheets.users);
  const teamsTable = readTable_(context.sheets.teams);

  const existingUser = findUsersByName_(usersTable, name);
  if (existingUser.length) {
    return { status: 'exists' };
  }

  let teamId = '';
  let group = '';
  let role = 'member';
  const now = isoNow_();
  const userId = generateId_('usr');

  if (teamMode === 'existing') {
    teamId = cleanText_(params.teamId);
    const teamRecord = findTeamById_(teamsTable, teamId);
    if (!teamRecord) {
      return { status: 'error', message: 'Selected team was not found.' };
    }
    group = cleanText_(getField_(teamsTable, teamRecord.row, 'teamName'));
  } else {
    group = cleanText_(params.teamName);
    if (!group) return { status: 'error', message: 'teamName is required.' };
    const existingTeam = findTeamByName_(teamsTable, group);
    if (existingTeam) {
      return { status: 'team_exists' };
    }
    teamId = generateId_('team');
    role = 'owner_admin';
    appendRecord_(teamsTable, {
      teamId: teamId,
      teamName: group,
      createdByUserId: userId,
      adminUserIds: userId,
      createdAt: now,
      status: 'active',
    });
  }

  appendRecord_(usersTable, {
    name: name,
    pin: pin,
    userId: userId,
    teamId: teamId,
    group: group,
    role: role,
    createdAt: now,
    updatedAt: now,
  });

  upsertRosterForUser_(context, {
    userId: userId,
    name: name,
    displayName: name,
    teamId: teamId,
    group: group,
    role: publicRole_(role),
  });

  rebuildMemberDirectory_(context);

  return {
    status: 'ok',
    userId: userId,
    teamId: teamId,
    group: group,
  };
}

function handleLogin_(context, params) {
  const auth = authenticateUser_(context, params);
  if (!auth.ok) return { status: 'error', message: auth.message };

  touchUserUpdatedAt_(context, auth.user);

  return {
    status: 'ok',
    userId: auth.user.userId,
    teamId: auth.user.teamId,
    group: auth.user.group,
  };
}

function handleSaveLog_(context, params) {
  const auth = authenticateUser_(context, params);
  if (!auth.ok) return { status: 'error', message: auth.message };

  const learning = cleanText_(params.learning);
  const nextAction = cleanText_(params.next);
  const goodNew = cleanText_(params.goodNew);
  if (!learning || !nextAction || !goodNew) {
    return { status: 'error', message: 'learning, next and goodNew are required.' };
  }

  const rosterProfile = findRosterByUserId_(context, auth.user.userId);
  const displayName = rosterProfile.displayName || auth.user.name;
  const now = isoNow_();

  const logsTable = readTable_(context.sheets.logs);
  appendRecord_(logsTable, {
    userId: auth.user.userId,
    teamId: auth.user.teamId,
    group: auth.user.group,
    name: auth.user.name,
    displayName: displayName,
    date: cleanText_(params.date) || formatDateKey_(new Date()),
    cond: cleanText_(params.cond),
    learning: learning,
    next: nextAction,
    goodNew: goodNew,
    achievementStatus: cleanText_(params.achievementStatus),
    whyMissed: cleanText_(params.whyMissed),
    retryPlan: cleanText_(params.retryPlan),
    grade: rosterProfile.grade || '',
    term: rosterProfile.term || '',
    createdAt: now,
    updatedAt: now,
  });

  return {
    status: 'ok',
    userId: auth.user.userId,
    teamId: auth.user.teamId,
    group: auth.user.group,
  };
}

function handleGetLogs_(context, params) {
  const auth = authenticateUser_(context, params);
  if (!auth.ok) return { status: 'error', message: auth.message };

  const logs = getLogsForUser_(context, auth.user);
  return {
    status: 'ok',
    userId: auth.user.userId,
    teamId: auth.user.teamId,
    group: auth.user.group,
    logs: logs,
  };
}

function handleGetGoodNews_(context, params) {
  const auth = authenticateUser_(context, params);
  if (!auth.ok) return { status: 'error', message: auth.message };

  const logTable = readTable_(context.sheets.logs);
  const rosterMap = buildRosterMap_(context);
  const targetGroup = normalizeText_(auth.user.group);
  const targetTeamId = cleanText_(auth.user.teamId);

  const items = logTable.rows
    .map(function (row, index) {
      return {
        rowNumber: index + 2,
        userId: cleanText_(getField_(logTable, row, 'userId')),
        teamId: cleanText_(getField_(logTable, row, 'teamId')),
        group: cleanText_(getField_(logTable, row, 'group')),
        displayName: cleanText_(getField_(logTable, row, 'displayName')) || cleanText_(getField_(logTable, row, 'name')),
        goodNew: cleanText_(getField_(logTable, row, 'goodNew')),
        createdAt: cleanText_(getField_(logTable, row, 'createdAt')) || cleanText_(getField_(logTable, row, 'date')),
        date: cleanText_(getField_(logTable, row, 'date')),
      };
    })
    .filter(function (item) {
      if (!item.goodNew) return false;
      if (targetTeamId && item.teamId) return item.teamId === targetTeamId;
      if (targetGroup && item.group) return normalizeText_(item.group) === targetGroup;
      if (!item.userId) return false;
      const roster = rosterMap[item.userId] || {};
      return normalizeText_(roster.group) === targetGroup;
    })
    .sort(function (a, b) {
      return sortByDateDesc_(a.createdAt || a.date, b.createdAt || b.date);
    })
    .slice(0, 50)
    .map(function (item) {
      return {
        id: 'gn_' + item.rowNumber,
        userId: item.userId,
        name: item.displayName,
        group: item.group,
        goodNew: item.goodNew,
        createdAt: item.createdAt,
        date: item.date,
      };
    });

  return { status: 'ok', items: items };
}

function listTeams_(context) {
  const teamsTable = readTable_(context.sheets.teams);
  return teamsTable.rows
    .map(function (row) {
      return {
        teamId: cleanText_(getField_(teamsTable, row, 'teamId')),
        teamName: cleanText_(getField_(teamsTable, row, 'teamName')),
        status: cleanText_(getField_(teamsTable, row, 'status')) || 'active',
      };
    })
    .filter(function (team) {
      return team.teamId && team.teamName && (!team.status || team.status === 'active');
    })
    .sort(function (a, b) {
      return a.teamName.localeCompare(b.teamName, 'ja');
    })
    .map(function (team) {
      return {
        teamId: team.teamId,
        teamName: team.teamName,
      };
    });
}

function getLogsForUser_(context, user) {
  const logTable = readTable_(context.sheets.logs);
  const rosterMap = buildRosterMap_(context);
  const fallbackName = normalizeText_(user.name);
  const fallbackGroup = normalizeText_(user.group);
  const strictMatches = logTable.rows
    .map(function (row) {
      const rowUserId = cleanText_(getField_(logTable, row, 'userId'));
      const rowName = cleanText_(getField_(logTable, row, 'name'));
      const rowGroup = cleanText_(getField_(logTable, row, 'group'));
      const roster = rowUserId ? (rosterMap[rowUserId] || {}) : {};
      const belongsToUser = rowUserId
        ? rowUserId === user.userId
        : normalizeText_(rowName) === fallbackName && (!fallbackGroup || !rowGroup || normalizeText_(rowGroup) === fallbackGroup);
      if (!belongsToUser) return null;

      return {
        userId: rowUserId || user.userId,
        teamId: cleanText_(getField_(logTable, row, 'teamId')) || user.teamId,
        group: cleanText_(getField_(logTable, row, 'group')) || roster.group || user.group,
        displayName: cleanText_(getField_(logTable, row, 'displayName')) || roster.displayName || user.name,
        grade: cleanText_(getField_(logTable, row, 'grade')) || roster.grade || '',
        term: cleanText_(getField_(logTable, row, 'term')) || roster.term || '',
        date: cleanText_(getField_(logTable, row, 'date')),
        cond: cleanText_(getField_(logTable, row, 'cond')),
        learning: cleanText_(getField_(logTable, row, 'learning')),
        next: cleanText_(getField_(logTable, row, 'next')),
        goodNew: cleanText_(getField_(logTable, row, 'goodNew')),
        achievementStatus: cleanText_(getField_(logTable, row, 'achievementStatus')),
        whyMissed: cleanText_(getField_(logTable, row, 'whyMissed')),
        retryPlan: cleanText_(getField_(logTable, row, 'retryPlan')),
        createdAt: cleanText_(getField_(logTable, row, 'createdAt')),
      };
    })
    .filter(Boolean)
    .sort(function (a, b) {
      return sortByDateDesc_(a.createdAt || a.date, b.createdAt || b.date);
    });

  if (strictMatches.length) {
    return strictMatches;
  }

  return logTable.rows
    .map(function (row) {
      const rowUserId = cleanText_(getField_(logTable, row, 'userId'));
      const rowName = cleanText_(getField_(logTable, row, 'name'));
      if (rowUserId && rowUserId !== user.userId) return null;
      if (normalizeText_(rowName) !== fallbackName) return null;
      const roster = rowUserId ? (rosterMap[rowUserId] || {}) : {};
      return {
        userId: rowUserId || user.userId,
        teamId: cleanText_(getField_(logTable, row, 'teamId')) || roster.teamId || user.teamId,
        group: cleanText_(getField_(logTable, row, 'group')) || roster.group || user.group,
        displayName: cleanText_(getField_(logTable, row, 'displayName')) || roster.displayName || user.name,
        grade: cleanText_(getField_(logTable, row, 'grade')) || roster.grade || '',
        term: cleanText_(getField_(logTable, row, 'term')) || roster.term || '',
        date: cleanText_(getField_(logTable, row, 'date')),
        cond: cleanText_(getField_(logTable, row, 'cond')),
        learning: cleanText_(getField_(logTable, row, 'learning')),
        next: cleanText_(getField_(logTable, row, 'next')),
        goodNew: cleanText_(getField_(logTable, row, 'goodNew')),
        achievementStatus: cleanText_(getField_(logTable, row, 'achievementStatus')),
        whyMissed: cleanText_(getField_(logTable, row, 'whyMissed')),
        retryPlan: cleanText_(getField_(logTable, row, 'retryPlan')),
        createdAt: cleanText_(getField_(logTable, row, 'createdAt')),
      };
    })
    .filter(Boolean)
    .sort(function (a, b) {
      return sortByDateDesc_(a.createdAt || a.date, b.createdAt || b.date);
    });
}

function authenticateUser_(context, params) {
  const pin = cleanText_(params.pin);
  const userId = cleanText_(params.userId);
  const usersTable = readTable_(context.sheets.users);

  if (userId) {
    const byId = usersTable.rows.filter(function (row) {
      return cleanText_(getField_(usersTable, row, 'userId')) === userId;
    });
    if (byId.length === 1 && cleanText_(getField_(usersTable, byId[0], 'pin')) === pin) {
      return { ok: true, user: enrichUserRecord_(context, usersTable, byId[0]) };
    }
  }

  const name = cleanText_(params.name);
  if (!name || !pin) return { ok: false, message: 'name and pin are required.' };

  const matches = usersTable.rows.filter(function (row) {
    return normalizeText_(getField_(usersTable, row, 'name')) === normalizeText_(name) &&
      cleanText_(getField_(usersTable, row, 'pin')) === pin;
  });

  if (matches.length !== 1) {
    return { ok: false, message: 'User not found.' };
  }

  return { ok: true, user: enrichUserRecord_(context, usersTable, matches[0]) };
}

function userRecordFromRow_(table, row) {
  return {
    name: cleanText_(getField_(table, row, 'name')),
    pin: cleanText_(getField_(table, row, 'pin')),
    userId: cleanText_(getField_(table, row, 'userId')),
    teamId: cleanText_(getField_(table, row, 'teamId')),
    group: cleanText_(getField_(table, row, 'group')),
    role: cleanText_(getField_(table, row, 'role')),
  };
}

function touchUserUpdatedAt_(context, user) {
  const usersTable = readTable_(context.sheets.users);
  const row = usersTable.rows.find(function (record) {
    return cleanText_(getField_(usersTable, record, 'userId')) === user.userId;
  });
  if (!row) return;
  setField_(usersTable, row, 'updatedAt', isoNow_());
  flushChangedRows_(usersTable);
}

function upsertRosterForUser_(context, user) {
  const rosterTable = readTable_(context.sheets.roster);
  const existingRow = rosterTable.rows.find(function (row) {
    return cleanText_(getField_(rosterTable, row, 'userId')) === user.userId;
  });
  const payload = {
    userId: user.userId,
    teamId: user.teamId,
    group: user.group,
    name: user.name,
    displayName: user.displayName || user.name,
    role: publicRole_(user.role || 'member'),
    updatedAt: isoNow_(),
  };
  if (existingRow) {
    Object.keys(payload).forEach(function (field) {
      if (payload[field] !== '') setField_(rosterTable, existingRow, field, payload[field]);
    });
    flushChangedRows_(rosterTable);
    return;
  }
  appendRecord_(rosterTable, payload);
}

function findRosterByUserId_(context, userId) {
  const rosterTable = readTable_(context.sheets.roster);
  const row = rosterTable.rows.find(function (record) {
    return cleanText_(getField_(rosterTable, record, 'userId')) === cleanText_(userId);
  });
  if (!row) {
    return { userId: userId, grade: '', term: '', displayName: '', group: '', role: '' };
  }
  return {
    userId: cleanText_(getField_(rosterTable, row, 'userId')),
    teamId: cleanText_(getField_(rosterTable, row, 'teamId')),
    group: cleanText_(getField_(rosterTable, row, 'group')),
    name: cleanText_(getField_(rosterTable, row, 'name')),
    displayName: cleanText_(getField_(rosterTable, row, 'displayName')),
    grade: cleanText_(getField_(rosterTable, row, 'grade')),
    term: cleanText_(getField_(rosterTable, row, 'term')),
    role: cleanText_(getField_(rosterTable, row, 'role')),
  };
}

function buildRosterMap_(context) {
  const rosterTable = readTable_(context.sheets.roster);
  return rosterTable.rows.reduce(function (acc, row) {
    const userId = cleanText_(getField_(rosterTable, row, 'userId'));
    if (!userId) return acc;
    acc[userId] = {
      userId: userId,
      teamId: cleanText_(getField_(rosterTable, row, 'teamId')),
      group: cleanText_(getField_(rosterTable, row, 'group')),
      name: cleanText_(getField_(rosterTable, row, 'name')),
      displayName: cleanText_(getField_(rosterTable, row, 'displayName')) || cleanText_(getField_(rosterTable, row, 'name')),
      grade: cleanText_(getField_(rosterTable, row, 'grade')),
      term: cleanText_(getField_(rosterTable, row, 'term')),
      role: cleanText_(getField_(rosterTable, row, 'role')),
    };
    return acc;
  }, {});
}

function ensureSystemReady_() {
  const spreadsheet = openSpreadsheet_();
  const sheets = ensureSheets_(spreadsheet);

  ensureHeaders_(sheets.users, CONFIG.userHeaders);
  ensureHeaders_(sheets.teams, CONFIG.teamHeaders);
  ensureHeaders_(sheets.roster, CONFIG.rosterHeaders);
  ensureHeaders_(sheets.logs, CONFIG.logHeaders);
  ensureHeaders_(sheets.review, CONFIG.reviewHeaders);
  ensureHeaders_(sheets.members, CONFIG.memberViewHeaders);

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(CONFIG.setupVersionProperty) !== CONFIG.schemaVersion) {
    runMigration_(spreadsheet, sheets);
    props.setProperty(CONFIG.setupVersionProperty, CONFIG.schemaVersion);
    props.setProperty(CONFIG.setupAtProperty, isoNow_());
  }

  syncTeamsFromUserGroupDropdown_(sheets.users, sheets.teams);

  return {
    spreadsheet: spreadsheet,
    sheets: sheets,
    props: props,
  };
}

function runMigration_(spreadsheet, sheets) {
  backupSourceSheets_(spreadsheet);
  const usersTable = readTable_(sheets.users);
  const rosterTable = readTable_(sheets.roster);
  const logsTable = readTable_(sheets.logs);
  const teamsTable = readTable_(sheets.teams);
  const reviewTable = readTable_(sheets.review);

  const groupSeedMap = collectDeclaredGroupsFromUsersDropdown_(sheets.users, usersTable);
  ensureTeamsFromGroups_(teamsTable, groupSeedMap);
  flushChangedRows_(teamsTable);

  migrateUsers_(usersTable, rosterTable, teamsTable);
  flushChangedRows_(usersTable);

  migrateRoster_(rosterTable, usersTable, teamsTable, reviewTable);
  flushChangedRows_(rosterTable);

  migrateLogs_(logsTable, usersTable, rosterTable, teamsTable, reviewTable);
  flushChangedRows_(logsTable);
  flushChangedRows_(reviewTable);

  rebuildMemberDirectory_({
    spreadsheet: spreadsheet,
    sheets: sheets,
  });
}

function collectDeclaredGroupsFromUsersDropdown_(usersSheet, usersTable) {
  const groups = {};
  getDeclaredTeamOptions_(usersSheet, usersTable).forEach(function (group) {
    groups[normalizeText_(group)] = group;
  });
  return groups;
}

function ensureTeamsFromGroups_(teamsTable, groupSeedMap) {
  Object.keys(groupSeedMap).forEach(function (normalizedGroup) {
    const group = groupSeedMap[normalizedGroup];
    if (!findTeamByName_(teamsTable, group)) {
      appendRecord_(teamsTable, {
        teamId: generateId_('team'),
        teamName: group,
        createdByUserId: '',
        adminUserIds: '',
        createdAt: isoNow_(),
        status: 'active',
      });
    }
  });
}

function migrateUsers_(usersTable, rosterTable, teamsTable) {
  const rosterNameMap = buildNameIndex_(rosterTable, 'name');
  usersTable.rows.forEach(function (row) {
    const name = cleanText_(getField_(usersTable, row, 'name'));
    if (!name) return;

    let group = cleanText_(getField_(usersTable, row, 'group'));
    if (!group) {
      const candidates = findCandidateRows_(rosterTable, rosterNameMap, name);
      if (candidates.length === 1) {
        group = cleanText_(getField_(rosterTable, candidates[0], 'group'));
      }
    }

    if (!cleanText_(getField_(usersTable, row, 'userId'))) {
      setField_(usersTable, row, 'userId', generateId_('usr'));
    }
    if (group) {
      setField_(usersTable, row, 'group', group);
      const team = findTeamByName_(teamsTable, group);
      if (team) {
        setField_(usersTable, row, 'teamId', cleanText_(getField_(teamsTable, team.row, 'teamId')));
      }
    }
    if (!cleanText_(getField_(usersTable, row, 'role'))) {
      setField_(usersTable, row, 'role', 'member');
    }
    if (!cleanText_(getField_(usersTable, row, 'createdAt'))) {
      setField_(usersTable, row, 'createdAt', isoNow_());
    }
    setField_(usersTable, row, 'updatedAt', isoNow_());
  });
}

function migrateRoster_(rosterTable, usersTable, teamsTable, reviewTable) {
  const userNameMap = buildNameIndex_(usersTable, 'name');

  rosterTable.rows.forEach(function (row, index) {
    const name = cleanText_(getField_(rosterTable, row, 'name')) || cleanText_(getField_(rosterTable, row, 'displayName'));
    if (!name) return;
    const group = cleanText_(getField_(rosterTable, row, 'group'));

    if (cleanText_(getField_(rosterTable, row, 'userId'))) {
      syncRosterTeamFields_(rosterTable, row, teamsTable);
      return;
    }

    const matched = matchSingleUserRow_(usersTable, userNameMap, name, group);
    if (!matched.userRow) {
      if (matched.ambiguous) {
        appendReviewRow_(reviewTable, {
          sheetName: CONFIG.sheetNames.roster,
          rowNumber: index + 2,
          reason: '複数ユーザー候補',
          name: name,
          group: group,
          candidates: matched.candidates,
          usersTable: usersTable,
        });
      }
      return;
    }

    const userId = cleanText_(getField_(usersTable, matched.userRow, 'userId'));
    setField_(rosterTable, row, 'userId', userId);
    setField_(rosterTable, row, 'teamId', cleanText_(getField_(usersTable, matched.userRow, 'teamId')));
    setField_(rosterTable, row, 'group', cleanText_(getField_(usersTable, matched.userRow, 'group')) || group);
    setField_(rosterTable, row, 'name', cleanText_(getField_(usersTable, matched.userRow, 'name')) || name);
    if (!cleanText_(getField_(rosterTable, row, 'displayName'))) {
      setField_(rosterTable, row, 'displayName', cleanText_(getField_(rosterTable, row, 'name')));
    }
    if (!cleanText_(getField_(rosterTable, row, 'role'))) {
      setField_(rosterTable, row, 'role', publicRole_(cleanText_(getField_(usersTable, matched.userRow, 'role')) || 'member'));
    }
    setField_(rosterTable, row, 'updatedAt', isoNow_());
  });
}

function migrateLogs_(logsTable, usersTable, rosterTable, teamsTable, reviewTable) {
  const userNameMap = buildNameIndex_(usersTable, 'name');
  const rosterMap = buildRosterMapFromTable_(rosterTable);

  logsTable.rows.forEach(function (row, index) {
    const rowUserId = cleanText_(getField_(logsTable, row, 'userId'));
    const name = cleanText_(getField_(logsTable, row, 'name')) || cleanText_(getField_(logsTable, row, 'displayName'));
    const group = cleanText_(getField_(logsTable, row, 'group'));

    if (rowUserId) {
      hydrateLogProfileFields_(logsTable, row, rosterMap[rowUserId], teamsTable);
      return;
    }
    if (!name) return;

    const matched = matchSingleUserRow_(usersTable, userNameMap, name, group);
    if (!matched.userRow) {
      if (matched.ambiguous) {
        appendReviewRow_(reviewTable, {
          sheetName: CONFIG.sheetNames.logs,
          rowNumber: index + 2,
          reason: '複数ユーザー候補',
          name: name,
          group: group,
          candidates: matched.candidates,
          usersTable: usersTable,
        });
      }
      return;
    }

    const userId = cleanText_(getField_(usersTable, matched.userRow, 'userId'));
    setField_(logsTable, row, 'userId', userId);
    setField_(logsTable, row, 'teamId', cleanText_(getField_(usersTable, matched.userRow, 'teamId')));
    setField_(logsTable, row, 'group', cleanText_(getField_(usersTable, matched.userRow, 'group')) || group);
    setField_(logsTable, row, 'displayName', cleanText_(getField_(usersTable, matched.userRow, 'name')) || name);
    hydrateLogProfileFields_(logsTable, row, rosterMap[userId], teamsTable);
    if (!cleanText_(getField_(logsTable, row, 'createdAt'))) {
      setField_(logsTable, row, 'createdAt', cleanText_(getField_(logsTable, row, 'date')) || isoNow_());
    }
    if (!cleanText_(getField_(logsTable, row, 'updatedAt'))) {
      setField_(logsTable, row, 'updatedAt', isoNow_());
    }
  });
}

function hydrateLogProfileFields_(logsTable, row, roster, teamsTable) {
  if (roster) {
    if (!cleanText_(getField_(logsTable, row, 'displayName'))) {
      setField_(logsTable, row, 'displayName', roster.displayName || roster.name || '');
    }
    if (!cleanText_(getField_(logsTable, row, 'grade'))) {
      setField_(logsTable, row, 'grade', roster.grade || '');
    }
    if (!cleanText_(getField_(logsTable, row, 'term'))) {
      setField_(logsTable, row, 'term', roster.term || '');
    }
    if (!cleanText_(getField_(logsTable, row, 'teamId')) && roster.teamId) {
      setField_(logsTable, row, 'teamId', roster.teamId);
    }
    if (!cleanText_(getField_(logsTable, row, 'group')) && roster.group) {
      setField_(logsTable, row, 'group', roster.group);
    }
  }

  if (!cleanText_(getField_(logsTable, row, 'group'))) {
    const teamId = cleanText_(getField_(logsTable, row, 'teamId'));
    const team = teamId ? findTeamById_(teamsTable, teamId) : null;
    if (team) {
      setField_(logsTable, row, 'group', cleanText_(getField_(teamsTable, team.row, 'teamName')));
    }
  }
}

function syncRosterTeamFields_(rosterTable, row, teamsTable) {
  const teamId = cleanText_(getField_(rosterTable, row, 'teamId'));
  const group = cleanText_(getField_(rosterTable, row, 'group'));
  if (teamId && !group) {
    const team = findTeamById_(teamsTable, teamId);
    if (team) {
      setField_(rosterTable, row, 'group', cleanText_(getField_(teamsTable, team.row, 'teamName')));
    }
  } else if (group && !teamId) {
    const team = findTeamByName_(teamsTable, group);
    if (team) {
      setField_(rosterTable, row, 'teamId', cleanText_(getField_(teamsTable, team.row, 'teamId')));
    }
  }
}

function rebuildMemberDirectory_(context) {
  const rosterTable = readTable_(context.sheets.roster);
  const usersTable = readTable_(context.sheets.users);
  const userRoleMap = usersTable.rows.reduce(function (acc, row) {
    const userId = cleanText_(getField_(usersTable, row, 'userId'));
    if (!userId) return acc;
    acc[userId] = cleanText_(getField_(usersTable, row, 'role'));
    return acc;
  }, {});

  const grouped = {};
  rosterTable.rows.forEach(function (row) {
    const userId = cleanText_(getField_(rosterTable, row, 'userId'));
    const group = cleanText_(getField_(rosterTable, row, 'group')) || '未所属';
    if (!grouped[group]) grouped[group] = [];
    grouped[group].push({
      group: group,
      grade: cleanText_(getField_(rosterTable, row, 'grade')),
      displayName: cleanText_(getField_(rosterTable, row, 'displayName')) || cleanText_(getField_(rosterTable, row, 'name')),
      term: cleanText_(getField_(rosterTable, row, 'term')),
      role: publicRole_(cleanText_(getField_(rosterTable, row, 'role')) || userRoleMap[userId] || 'member'),
      userId: userId,
      teamId: cleanText_(getField_(rosterTable, row, 'teamId')),
    });
  });

  const output = [];
  const groups = Object.keys(grouped).sort(function (a, b) {
    return a.localeCompare(b, 'ja');
  });

  groups.forEach(function (group, index) {
    const members = grouped[group].sort(compareMemberRows_);
    output.push(['所属: ' + group, '人数: ' + members.length, '', '', '', '', '']);
    output.push(CONFIG.memberViewHeaders.slice());
    members.forEach(function (member) {
      output.push([
        member.group,
        member.grade,
        member.displayName,
        member.term,
        member.role,
        member.userId,
        member.teamId,
      ]);
    });
    if (index !== groups.length - 1) {
      output.push(['', '', '', '', '', '', '']);
    }
  });

  if (!output.length) {
    output.push(CONFIG.memberViewHeaders.slice());
  }

  const sheet = context.sheets.members;
  sheet.clearContents();
  sheet.getRange(1, 1, output.length, CONFIG.memberViewHeaders.length).setValues(output);
}

function compareMemberRows_(a, b) {
  const gradeDiff = gradeSortKey_(a.grade) - gradeSortKey_(b.grade);
  if (gradeDiff !== 0) return gradeDiff;
  return a.displayName.localeCompare(b.displayName, 'ja');
}

function gradeSortKey_(grade) {
  const text = cleanText_(grade);
  const match = text.match(/(\d+)/);
  if (match) return Number(match[1]);
  return text ? 99 : 999;
}

function appendReviewRow_(reviewTable, payload) {
  const existingRows = reviewTable.rows.filter(function (row) {
    return cleanText_(getField_(reviewTable, row, 'sheetName')) === payload.sheetName &&
      String(getField_(reviewTable, row, 'rowNumber')) === String(payload.rowNumber) &&
      cleanText_(getField_(reviewTable, row, 'reason')) === payload.reason;
  });
  if (existingRows.length) return;

  const candidateUserIds = payload.candidates.map(function (row) {
    return cleanText_(getField_(payload.usersTable, row, 'userId'));
  }).join(', ');
  const candidateNames = payload.candidates.map(function (row) {
    return cleanText_(getField_(payload.usersTable, row, 'name'));
  }).join(', ');

  appendRecord_(reviewTable, {
    timestamp: isoNow_(),
    sheetName: payload.sheetName,
    rowNumber: payload.rowNumber,
    reason: payload.reason,
    name: payload.name,
    group: payload.group,
    candidateUserIds: candidateUserIds,
    candidateNames: candidateNames,
  });
}

function findUsersByName_(usersTable, name) {
  return usersTable.rows.filter(function (row) {
    return normalizeText_(getField_(usersTable, row, 'name')) === normalizeText_(name);
  });
}

function buildNameIndex_(table, fieldName) {
  return table.rows.reduce(function (acc, row) {
    const key = normalizeText_(getField_(table, row, fieldName));
    if (!key) return acc;
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {});
}

function findCandidateRows_(table, index, name) {
  return index[normalizeText_(name)] || [];
}

function matchSingleUserRow_(usersTable, nameIndex, name, group) {
  const candidates = findCandidateRows_(usersTable, nameIndex, name);
  if (!candidates.length) return { userRow: null, ambiguous: false, candidates: [] };
  if (candidates.length === 1) return { userRow: candidates[0], ambiguous: false, candidates: candidates };

  const normalizedGroup = normalizeText_(group);
  if (normalizedGroup) {
    const narrowed = candidates.filter(function (row) {
      return normalizeText_(getField_(usersTable, row, 'group')) === normalizedGroup;
    });
    if (narrowed.length === 1) {
      return { userRow: narrowed[0], ambiguous: false, candidates: narrowed };
    }
  }
  return { userRow: null, ambiguous: true, candidates: candidates };
}

function findTeamById_(teamsTable, teamId) {
  const row = teamsTable.rows.find(function (record) {
    return cleanText_(getField_(teamsTable, record, 'teamId')) === cleanText_(teamId);
  });
  return row ? { row: row } : null;
}

function findTeamByName_(teamsTable, teamName) {
  const row = teamsTable.rows.find(function (record) {
    return normalizeText_(getField_(teamsTable, record, 'teamName')) === normalizeText_(teamName);
  });
  return row ? { row: row } : null;
}

function buildRosterMapFromTable_(rosterTable) {
  return rosterTable.rows.reduce(function (acc, row) {
    const userId = cleanText_(getField_(rosterTable, row, 'userId'));
    if (!userId) return acc;
    acc[userId] = {
      userId: userId,
      teamId: cleanText_(getField_(rosterTable, row, 'teamId')),
      group: cleanText_(getField_(rosterTable, row, 'group')),
      name: cleanText_(getField_(rosterTable, row, 'name')),
      displayName: cleanText_(getField_(rosterTable, row, 'displayName')) || cleanText_(getField_(rosterTable, row, 'name')),
      grade: cleanText_(getField_(rosterTable, row, 'grade')),
      term: cleanText_(getField_(rosterTable, row, 'term')),
      role: cleanText_(getField_(rosterTable, row, 'role')),
    };
    return acc;
  }, {});
}

function openSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(CONFIG.spreadsheetIdProperty);
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function enrichUserRecord_(context, usersTable, row) {
  const user = userRecordFromRow_(usersTable, row);
  const rosterTable = readTable_(context.sheets.roster);
  const logsTable = readTable_(context.sheets.logs);
  const teamsTable = readTable_(context.sheets.teams);
  let changed = false;

  let rosterRow = null;
  if (user.userId) {
    rosterRow = rosterTable.rows.find(function (record) {
      return cleanText_(getField_(rosterTable, record, 'userId')) === user.userId;
    }) || null;
  }
  if (!rosterRow) {
    const rosterMatches = rosterTable.rows.filter(function (record) {
      return normalizeText_(getField_(rosterTable, record, 'name')) === normalizeText_(user.name);
    });
    if (rosterMatches.length === 1) rosterRow = rosterMatches[0];
  }

  if (!user.group && rosterRow) {
    user.group = cleanText_(getField_(rosterTable, rosterRow, 'group'));
    if (user.group) {
      setField_(usersTable, row, 'group', user.group);
      changed = true;
    }
  }

  if (!user.teamId && rosterRow) {
    user.teamId = cleanText_(getField_(rosterTable, rosterRow, 'teamId'));
    if (user.teamId) {
      setField_(usersTable, row, 'teamId', user.teamId);
      changed = true;
    }
  }

  if (!user.group) {
    const logGroups = uniqueNonEmpty_(
      logsTable.rows
        .filter(function (record) {
          return normalizeText_(getField_(logsTable, record, 'name')) === normalizeText_(user.name);
        })
        .map(function (record) {
          return cleanText_(getField_(logsTable, record, 'group'));
        })
    );
    if (logGroups.length === 1) {
      user.group = logGroups[0];
      setField_(usersTable, row, 'group', user.group);
      changed = true;
    }
  }

  if (!user.teamId && user.group) {
    const team = findTeamByName_(teamsTable, user.group);
    if (team) {
      user.teamId = cleanText_(getField_(teamsTable, team.row, 'teamId'));
      if (user.teamId) {
        setField_(usersTable, row, 'teamId', user.teamId);
        changed = true;
      }
    }
  }

  if (!user.group && user.teamId) {
    const teamById = findTeamById_(teamsTable, user.teamId);
    if (teamById) {
      user.group = cleanText_(getField_(teamsTable, teamById.row, 'teamName'));
      if (user.group) {
        setField_(usersTable, row, 'group', user.group);
        changed = true;
      }
    }
  }

  if (changed) {
    setField_(usersTable, row, 'updatedAt', isoNow_());
    flushChangedRows_(usersTable);
  }

  return user;
}

function syncTeamsFromUserGroupDropdown_(usersSheet, teamsSheet) {
  const usersTable = readTable_(usersSheet);
  const teamsTable = readTable_(teamsSheet);
  const declaredOptions = getDeclaredTeamOptions_(usersSheet, usersTable);
  if (!declaredOptions.length) return;

  const declaredMap = {};
  declaredOptions.forEach(function (teamName) {
    declaredMap[normalizeText_(teamName)] = teamName;
    const existing = findTeamByName_(teamsTable, teamName);
    if (!existing) {
      appendRecord_(teamsTable, {
        teamId: generateId_('team'),
        teamName: teamName,
        createdByUserId: '',
        adminUserIds: '',
        createdAt: isoNow_(),
        status: 'active',
      });
      return;
    }
    if (cleanText_(getField_(teamsTable, existing.row, 'status')) !== 'active') {
      setField_(teamsTable, existing.row, 'status', 'active');
    }
  });

  teamsTable.rows.forEach(function (row) {
    const teamName = cleanText_(getField_(teamsTable, row, 'teamName'));
    const createdByUserId = cleanText_(getField_(teamsTable, row, 'createdByUserId'));
    if (!teamName || createdByUserId) return;
    if (!declaredMap[normalizeText_(teamName)]) {
      setField_(teamsTable, row, 'status', 'inactive');
    }
  });

  flushChangedRows_(teamsTable);
}

function getDeclaredTeamOptions_(usersSheet, usersTable) {
  const groupColumnIndex = resolveExistingColumnIndex_(usersTable, 'group');
  if (groupColumnIndex < 0) return [];
  const rowCount = Math.max(usersSheet.getLastRow() - 1, 1);
  const validationRules = usersSheet.getRange(2, groupColumnIndex + 1, rowCount, 1).getDataValidations();
  const optionsMap = {};
  const assignedGroupMap = {};

  usersTable.rows.forEach(function (row) {
    const group = cleanText_(getField_(usersTable, row, 'group'));
    if (group) assignedGroupMap[normalizeText_(group)] = group;
  });

  validationRules.forEach(function (ruleRow) {
    ruleRow.forEach(function (rule) {
      extractTeamOptionsFromValidationRule_(rule).forEach(function (value) {
        optionsMap[normalizeText_(value)] = value;
      });
    });
  });

  const filteredKeys = Object.keys(optionsMap).filter(function (key) {
    if (!Object.keys(assignedGroupMap).length) return true;
    return !!assignedGroupMap[key];
  });

  return filteredKeys
    .map(function (key) { return optionsMap[key]; })
    .filter(Boolean)
    .sort(function (a, b) {
      return a.localeCompare(b, 'ja');
    });
}

function extractTeamOptionsFromValidationRule_(rule) {
  if (!rule) return [];
  const criteriaType = rule.getCriteriaType();
  const criteriaValues = rule.getCriteriaValues();
  if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return uniqueNonEmpty_(criteriaValues[0] || []);
  }
  if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    const sourceRange = criteriaValues[0];
    if (!sourceRange) return [];
    return uniqueNonEmpty_(sourceRange.getDisplayValues().reduce(function (acc, row) {
      return acc.concat(row);
    }, []));
  }
  return [];
}

function resolveExistingColumnIndex_(table, fieldName) {
  const columnName = table.indexMap[fieldName] !== undefined
    ? fieldName
    : resolveExistingAlias_(table, fieldName);
  if (!columnName) return -1;
  return typeof table.indexMap[columnName] === 'number' ? table.indexMap[columnName] : -1;
}

function uniqueNonEmpty_(values) {
  const map = {};
  (values || []).forEach(function (value) {
    const text = cleanText_(value);
    if (!text) return;
    map[normalizeText_(text)] = text;
  });
  return Object.keys(map).map(function (key) {
    return map[key];
  });
}

function ensureSheets_(spreadsheet) {
  return {
    users: getOrCreateSheet_(spreadsheet, CONFIG.sheetNames.users),
    roster: getOrCreateSheet_(spreadsheet, CONFIG.sheetNames.roster),
    logs: getOrCreateSheet_(spreadsheet, CONFIG.sheetNames.logs),
    members: getOrCreateSheet_(spreadsheet, CONFIG.sheetNames.members),
    teams: getOrCreateSheet_(spreadsheet, CONFIG.sheetNames.teams),
    review: getOrCreateSheet_(spreadsheet, CONFIG.sheetNames.review),
  };
}

function getOrCreateSheet_(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function backupSourceSheets_(spreadsheet) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  CONFIG.sourceBackupSheets.forEach(function (name) {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    const backupName = 'Backup_' + stamp + '_' + name;
    if (spreadsheet.getSheetByName(backupName)) return;
    sheet.copyTo(spreadsheet).setName(backupName);
  });
}

function ensureHeaders_(sheet, requiredHeaders) {
  const lastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  const currentHeaders = lastColumn ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(cleanText_) : [];

  if (!currentHeaders.some(Boolean)) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  const existing = currentHeaders.filter(Boolean);
  const missing = requiredHeaders.filter(function (header) {
    return existing.indexOf(header) === -1;
  });
  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
}

function readTable_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0 || lastColumn === 0) {
    return {
      sheet: sheet,
      headers: [],
      rows: [],
      indexMap: {},
      changedRows: {},
    };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(cleanText_);
  const rows = values.slice(1).map(function (row, index) {
    const cloned = row.slice();
    cloned.__rowNumber = index + 2;
    return cloned;
  });
  return {
    sheet: sheet,
    headers: headers,
    rows: rows,
    indexMap: headers.reduce(function (acc, header, index) {
      if (header) acc[header] = index;
      return acc;
    }, {}),
    changedRows: {},
  };
}

function getField_(table, row, fieldName) {
  const aliases = FIELD_ALIASES[fieldName] || [fieldName];
  for (var i = 0; i < aliases.length; i += 1) {
    const columnIndex = table.indexMap[aliases[i]];
    if (typeof columnIndex === 'number') {
      return row[columnIndex];
    }
  }
  return '';
}

function setField_(table, row, fieldName, value) {
  const columnName = table.indexMap[fieldName] !== undefined
    ? fieldName
    : resolveExistingAlias_(table, fieldName) || fieldName;
  const index = table.indexMap[columnName];
  if (typeof index !== 'number') {
    throw new Error('Missing column: ' + fieldName);
  }
  if (row[index] === value) return;
  row[index] = value;
  table.changedRows[row.__rowNumber] = row;
}

function resolveExistingAlias_(table, fieldName) {
  const aliases = FIELD_ALIASES[fieldName] || [];
  return aliases.find(function (alias) {
    return typeof table.indexMap[alias] === 'number';
  }) || null;
}

function appendRecord_(table, fields) {
  const row = new Array(table.headers.length).fill('');
  row.__rowNumber = table.sheet.getLastRow() + 1;
  Object.keys(fields).forEach(function (field) {
    const index = table.indexMap[field];
    if (typeof index === 'number') {
      row[index] = fields[field];
    }
  });
  table.sheet.appendRow(row);
  return row;
}

function flushChangedRows_(table) {
  const rowNumbers = Object.keys(table.changedRows).map(Number).sort(function (a, b) {
    return a - b;
  });
  rowNumbers.forEach(function (rowNumber) {
    const row = table.changedRows[rowNumber];
    const output = row.slice(0, table.headers.length);
    table.sheet.getRange(rowNumber, 1, 1, table.headers.length).setValues([output]);
  });
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseJsonSafe_(text) {
  try {
    return JSON.parse(text || '{}');
  } catch (error) {
    return {};
  }
}

function cleanText_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function normalizeText_(value) {
  const text = cleanText_(value);
  if (!text) return '';
  return text.normalize('NFKC').replace(/\s+/g, '').toLowerCase();
}

function isoNow_() {
  return new Date().toISOString();
}

function generateId_(prefix) {
  return prefix + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 20);
}

function formatDateKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function sortByDateDesc_(a, b) {
  const aTime = Date.parse(a || '') || 0;
  const bTime = Date.parse(b || '') || 0;
  return bTime - aTime;
}

function publicRole_(role) {
  const value = cleanText_(role);
  if (!value || value === 'owner_admin') return 'member';
  return value;
}

function runInitialMigration() {
  const context = ensureSystemReady_();
  rebuildMemberDirectory_(context);
}

function rebuildMemberDirectory() {
  const context = ensureSystemReady_();
  rebuildMemberDirectory_(context);
}
