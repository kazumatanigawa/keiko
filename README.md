# KEIKO

Japanese drum practice logging app.

## Architecture

- `index.html`: static mobile-first application
- `supabase/functions/keiko-api`: login, registration, and authenticated membership management API
- Supabase Auth: PIN-based application login through derived passwords
- Supabase Postgres: profiles, teams, practice logs, scoped notes, comments, and reports
- Supabase RLS: per-user and per-team data access
- Supabase RPC: RLS-protected reads and writes for logs, Good&New, notes, and comments
- Browser cache: user/team-scoped home data, recent logs, and notes
- Cursor pagination: logs, notes, and comments load in pages of 20
- Offline support: local drafts and an idempotent write outbox
- Multi-team support: personal use and multiple team memberships can coexist
- View contexts: switch between personal, each team, and the opt-in global feed
- Note visibility: private, team, or global with explicit team and user opt-in
- Moderation: operator-only team settings, report review, and content hiding
- `admin.html`: role-aware overview of users plus paged, date-sorted team logs, notes, members, and comments
- Home hub: operator announcements, first-log tutorial, official resources, community links, and coaching requests
- Team Admin OS: team-scoped log, note, member, and role views for approved owners, managers, coaches, and advisors
- Coaching intake: duration-and-location estimate at 6,000 JPY/hour plus best-effort round-trip public transport fare from Fuchu; dates are arranged by reply email
- Custom bachi intake: configurable wood, dimensions, quantity pricing, one-month delivery estimate, and email receipts

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

The Admin OS is available at `admin.html` after the latest database migration
and Edge Function are deployed. Active profiles with `app_role = 'operator'`
can load the full console. Approved memberships with `team_role` set to
`owner_admin`, `owner`, `manager`, `coach`, or `advisor` can load only assigned
teams and only team-visible logs and notes. Operator and team-admin detail views
are audited separately.

## New team onboarding

School and partner teams are provisioned before members register:

1. The KEIKO operator creates the team with its official name and an initial
   administrator account named `<official team name>管理者`.
2. Students and other members open `はじめて登録`, refresh the team directory
   if needed, and select the existing official team name.
3. Members must not launch a second team when the intended team already exists.

The registration screen refreshes the active team directory when the app starts
and whenever the registration tab opens. It also includes a manual refresh
button so newly provisioned teams appear without clearing browser data.
