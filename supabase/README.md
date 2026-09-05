# Supabase backend

Apply SQL files in `migrations` in filename order. The current release includes
direct authenticated writes, paged reads, rotating timekeepers, and audited
multi-team membership changes.

Deploy `functions/keiko-api` with JWT verification disabled because login and
registration do not have a user JWT yet. The function still validates every
action. Membership changes require a valid user access token; joining and
transferring also require the registration code.

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
