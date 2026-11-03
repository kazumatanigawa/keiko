# Supabase backend

Apply SQL files in `migrations` in filename order. The current release includes
direct authenticated writes, paged reads, rotating timekeepers, and audited
multi-team membership changes. Global notes are opt-in, school teams are blocked
from global publishing, and operator-only moderation is stored in Postgres.
The `2026091001_operator_admin_os.sql` migration adds the operator-only user,
activity, log, and note read models used by `admin.html`.
The `2026091101_operator_team_detail.sql` migration adds paged, date-sorted team
logs, notes, and member details for the same operator console.

Deploy `functions/keiko-api` with JWT verification disabled because login and
registration do not have a user JWT yet. The function still validates every
action. Membership changes require a valid user access token; joining and
transferring also require the registration code.

Set these Edge Function secrets:

| Secret | Purpose |
| --- | --- |
| `KEIKO_AUTH_PEPPER` | Existing value used to derive Auth passwords |
| `KEIKO_REGISTRATION_CODE` | Existing registration code |
| `RESEND_API_KEY` | Resend API key used to email content reports |
| `KEIKO_REPORT_FROM_EMAIL` | Verified sender, for example `KEIKO OS <report@example.com>` |
| `KEIKO_COACHING_TO_EMAIL` | Address that receives coaching requests; defaults to the existing operator mailbox |
| `KEIKO_BACHI_TO_EMAIL` | Address that receives custom bachi requests; defaults to `kokongakumeiza@gmail.com` |
| `KEIKO_COACHING_ORIGIN` | Exact Fuchu origin used for transport estimates, for example `府中駅, 東京都府中市` |
| `GOOGLE_MAPS_ROUTES_API_KEY` | Server-only Routes API key for public-transit fare estimates |

`SUPABASE_URL`, `SUPABASE_ANON_KEY`, and `SUPABASE_SERVICE_ROLE_KEY` are supplied
by Supabase. Never place server secrets in `index.html` or commit them to Git.

```sh
supabase db push
supabase secrets set KEIKO_AUTH_PEPPER=... KEIKO_REGISTRATION_CODE=...
supabase secrets set RESEND_API_KEY=... KEIKO_REPORT_FROM_EMAIL='KEIKO OS <report@example.com>'
supabase secrets set KEIKO_COACHING_TO_EMAIL=... KEIKO_COACHING_ORIGIN='府中駅, 東京都府中市' GOOGLE_MAPS_ROUTES_API_KEY=...
supabase secrets set KEIKO_BACHI_TO_EMAIL=kokongakumeiza@gmail.com
supabase functions deploy keiko-api --no-verify-jwt
```

The `2026110101_home_hub_team_admin_saas.sql` migration adds home resources,
operator announcements, coaching requests, team-scoped Admin OS access, audit
events, and subscription entitlement tables. A user becomes a team
administrator only after an operator changes that membership role to
`owner_admin`, `owner`, `manager`, `coach`, or `advisor`. Team administrators
can read only `visibility = 'team'` logs and notes from assigned teams. Missing
Google Maps secrets do not block a coaching request; transport is stored as
pending manual confirmation instead.

`2026110103_interactive_service_requests.sql` adds interactive custom-bachi
quotes and configurable coaching rates. Edit `bachi_wood_price_options`,
`bachi_quantity_price_tiers`, `bachi_order_settings`, and
`coaching_rate_options` in the Supabase Table Editor to change prices without
editing the browser application. Both request types are stored before email is
sent. The operator and requester each receive a separate message when Resend is
configured. Enable Google Maps Platform Routes API for the Maps key. The legacy
Directions API secret name is still accepted temporarily so existing deployments
can migrate without an outage.

Coaching requests intentionally store no requested date or start/end clock time.
The form qualifies interest with a duration-and-location estimate first, then the
operator arranges an available date by reply email. Team and contact names are
prefilled from the signed-in user; email and location remain required.

The initial bachi catalog uses one-month lead times and editable reference
prices for ヒノキ, ホオ, and カエデ. Confirm those prices in the Table Editor
before accepting production requests. Half-day and full-day coaching defaults
are four and eight hours at the same 6,000 JPY hourly basis.
