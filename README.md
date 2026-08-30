# KEIKO

Japanese drum practice logging app.

## Architecture

- `index.html`: static mobile-first application
- `supabase/functions/keiko-api`: login and registration API
- Supabase Auth: PIN-based application login through derived passwords
- Supabase Postgres: profiles, teams, practice logs, notes, and comments
- Supabase RLS: per-user and per-team data access
- Supabase RPC: RLS-protected reads and writes for logs, Good&New, notes, and comments
- Browser cache: user/team-scoped home data, recent logs, and notes
- Cursor pagination: logs, notes, and comments load in pages of 20
- Offline support: local drafts and an idempotent write outbox

Authenticated reads and writes go directly from the browser to Supabase with the
publishable key and the user's short-lived access token. The browser never
receives the Supabase secret key, auth pepper, registration code, or stores a
user's PIN. Login and registration run in a Supabase Edge Function.
Google Sheets are used only as the source for one-time migration scripts.

## Migration scripts

Copy `.env.migration.example` to a local `.env.migration`, set the secrets, and
load those variables before running a script. Every script defaults to dry-run;
add `--apply` only after reviewing its summary.

```sh
node scripts/migrate-users-to-supabase.mjs
node scripts/migrate-practice-logs-to-supabase.mjs
node scripts/migrate-team-notes-to-supabase.mjs
```

Database and Edge Function deployment steps are in `supabase/README.md`.
