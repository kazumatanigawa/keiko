# KEIKO

Japanese drum practice logging app.

## Architecture

- `index.html`: static mobile-first application
- `gas/Code.gs`: Apps Script API gateway
- Supabase Auth: PIN-based application login through derived passwords
- Supabase Postgres: profiles, teams, practice logs, notes, and comments
- Supabase RLS: per-user and per-team data access
- Supabase RPC: one-call session context and lightweight home summaries
- Browser cache: user/team-scoped home data and up to 100 recent logs
- Log pagination: cursor-based pages of 20 records

The browser never receives the Supabase secret key or stores a user's PIN.
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
