# Supabase backend

Apply SQL files in `migrations` in filename order. The current release adds
direct authenticated writes, idempotency indexes, 20-item note pages, and lazy
comment pages in `2026083001_direct_write_and_note_paging.sql`.

Deploy `functions/keiko-api` with JWT verification disabled because login and
registration do not have a user JWT yet. The function still validates every
action and only exposes login, registration, health, and the active team list.

Set these Edge Function secrets:

| Secret | Purpose |
| --- | --- |
| `KEIKO_AUTH_PEPPER` | Existing value used to derive Auth passwords |
| `KEIKO_REGISTRATION_CODE` | Existing registration code |

`SUPABASE_URL`, `SUPABASE_ANON_KEY`, and `SUPABASE_SERVICE_ROLE_KEY` are supplied
by Supabase. Never place server secrets in `index.html` or commit them to Git.

```sh
supabase db push
supabase secrets set KEIKO_AUTH_PEPPER=... KEIKO_REGISTRATION_CODE=...
supabase functions deploy keiko-api --no-verify-jwt
```
