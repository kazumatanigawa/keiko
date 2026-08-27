const SPREADSHEET_ID = '1x1HUq_FuRxEdGc3sNWut5rSlKKU0UIMcfn9VNNjF538';
const NOTES_SHEET = 'チームノート';
const COMMENTS_SHEET = 'チームノートコメント';

const applyChanges = process.argv.includes('--apply');
const supabaseUrl = String(process.env.SUPABASE_URL || '').replace(/\/$/, '');
const adminApiKey = String(
  process.env.SUPABASE_SECRET_KEY ||
    process.env.SUPABASE_SERVICE_ROLE_KEY ||
    '',
);

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

async function loadSheet(sheetName) {
  const url = new URL(
    `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq`,
  );
  url.searchParams.set('tqx', 'out:csv');
  url.searchParams.set('sheet', sheetName);
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to read ${sheetName}: HTTP ${response.status}`);
  }

  const rows = parseCsv(await response.text());
  const headers = rows.shift() || [];
  const indexes = new Map(headers.map((header, index) => [header, index]));
  const get = (row, header) => String(row[indexes.get(header)] || '').trim();
  return { rows, indexes, get };
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

async function upsertByLegacyId(table, column, value, payload) {
  const existing = await request(
    `/rest/v1/${table}?select=id&${column}=eq.${encodeURIComponent(value)}&limit=1`,
  );
  if (existing.length) {
    return request(`/rest/v1/${table}?id=eq.${encodeURIComponent(existing[0].id)}`, {
      method: 'PATCH',
      headers: { prefer: 'return=representation' },
      body: JSON.stringify(payload),
    });
  }
  return request(`/rest/v1/${table}`, {
    method: 'POST',
    headers: { prefer: 'return=representation' },
    body: JSON.stringify(payload),
  });
}

function requireHeaders(sheet, headers) {
  for (const header of headers) {
    if (!sheet.indexes.has(header)) throw new Error(`Missing header: ${header}`);
  }
}

if (!supabaseUrl || !adminApiKey) {
  throw new Error('SUPABASE_URL and SUPABASE_SECRET_KEY are required');
}

const [notesSheet, commentsSheet] = await Promise.all([
  loadSheet(NOTES_SHEET),
  loadSheet(COMMENTS_SHEET),
]);
requireHeaders(notesSheet, [
  'noteId', 'teamId', 'authorUserId', 'authorName', 'title', 'body',
  'createdAt', 'updatedAt', 'status',
]);
requireHeaders(commentsSheet, [
  'commentId', 'noteId', 'teamId', 'authorUserId', 'authorName', 'body',
  'createdAt', 'updatedAt', 'status',
]);

const sourceNotes = notesSheet.rows
  .map((row) => ({
    legacyNoteId: notesSheet.get(row, 'noteId'),
    legacyTeamId: notesSheet.get(row, 'teamId'),
    legacyAuthorId: notesSheet.get(row, 'authorUserId'),
    authorName: notesSheet.get(row, 'authorName'),
    title: notesSheet.get(row, 'title'),
    body: notesSheet.get(row, 'body'),
    createdAt: notesSheet.get(row, 'createdAt'),
    updatedAt: notesSheet.get(row, 'updatedAt'),
    status: notesSheet.get(row, 'status') || 'active',
  }))
  .filter((note) => note.legacyNoteId);

const sourceComments = commentsSheet.rows
  .map((row) => ({
    legacyCommentId: commentsSheet.get(row, 'commentId'),
    legacyNoteId: commentsSheet.get(row, 'noteId'),
    legacyTeamId: commentsSheet.get(row, 'teamId'),
    legacyAuthorId: commentsSheet.get(row, 'authorUserId'),
    authorName: commentsSheet.get(row, 'authorName'),
    body: commentsSheet.get(row, 'body'),
    createdAt: commentsSheet.get(row, 'createdAt'),
    updatedAt: commentsSheet.get(row, 'updatedAt'),
    status: commentsSheet.get(row, 'status') || 'active',
  }))
  .filter((comment) => comment.legacyCommentId);

console.log(JSON.stringify({ notes: sourceNotes.length, comments: sourceComments.length, mode: applyChanges ? 'apply' : 'dry-run' }, null, 2));
if (!applyChanges) process.exit(0);

const [teams, profiles] = await Promise.all([
  request('/rest/v1/teams?select=id,legacy_team_id'),
  request('/rest/v1/profiles?select=id,legacy_user_id'),
]);
const teamIds = new Map(teams.map((team) => [team.legacy_team_id, team.id]));
const userIds = new Map(profiles.map((profile) => [profile.legacy_user_id, profile.id]));

for (const note of sourceNotes) {
  const teamId = teamIds.get(note.legacyTeamId);
  const authorUserId = userIds.get(note.legacyAuthorId);
  if (!teamId || !authorUserId) {
    throw new Error(`Cannot resolve note ${note.legacyNoteId}`);
  }
  await upsertByLegacyId('team_notes', 'legacy_note_id', note.legacyNoteId, {
    legacy_note_id: note.legacyNoteId,
    team_id: teamId,
    author_user_id: authorUserId,
    author_name_snapshot: note.authorName,
    title: note.title,
    body: note.body,
    created_at: note.createdAt,
    updated_at: note.updatedAt || note.createdAt,
    status: note.status === 'deleted' ? 'deleted' : 'active',
  });
}

const migratedNotes = await request(
  '/rest/v1/team_notes?select=id,legacy_note_id&legacy_note_id=not.is.null',
);
const noteIds = new Map(
  migratedNotes.map((note) => [note.legacy_note_id, note.id]),
);

for (const comment of sourceComments) {
  const noteId = noteIds.get(comment.legacyNoteId);
  const teamId = teamIds.get(comment.legacyTeamId);
  const authorUserId = userIds.get(comment.legacyAuthorId);
  if (!noteId || !teamId || !authorUserId) {
    throw new Error(`Cannot resolve comment ${comment.legacyCommentId}`);
  }
  await upsertByLegacyId('team_note_comments', 'legacy_comment_id', comment.legacyCommentId, {
    legacy_comment_id: comment.legacyCommentId,
    note_id: noteId,
    team_id: teamId,
    author_user_id: authorUserId,
    author_name_snapshot: comment.authorName,
    body: comment.body,
    created_at: comment.createdAt,
    updated_at: comment.updatedAt || comment.createdAt,
    status: comment.status === 'deleted' ? 'deleted' : 'active',
  });
}

console.log('Team notes migration completed.');
