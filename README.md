# KEIKO

Japanese drum practice logging app.

## Architecture

- `index.html`: static mobile-first application
- `gas/Code.gs`: login, registration, token refresh, and write API gateway
- Supabase Auth: PIN-based application login through derived passwords
- Supabase Postgres: profiles, teams, practice logs, notes, and comments
- Supabase RLS: per-user and per-team data access
- Supabase RPC: RLS-protected direct reads for home, logs, Good&New, and notes
- Browser cache: user/team-scoped home data and up to 100 recent logs
- Log pagination: cursor-based pages of 20 records

Authenticated reads go directly from the browser to Supabase with the
publishable key and the user's short-lived access token. The browser never
receives the Supabase secret key or stores a user's PIN. Login and writes remain
behind GAS so the auth pepper and privileged key stay server-only.
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

GAS configuration and deployment steps are in `gas/README.md`.
