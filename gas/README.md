# GAS Backend

`gas/Code.gs` is the thin API layer between `index.html` and Supabase. It does
not read or write Google Sheets.

## Script Properties

Set these values in Apps Script under **Project Settings > Script Properties**.

| Property | Purpose |
| --- | --- |
| `SUPABASE_URL` | Project URL such as `https://PROJECT_REF.supabase.co` |
| `SUPABASE_SECRET_KEY` | Legacy `service_role` JWT for Apps Script compatibility |
| `SUPABASE_PUBLISHABLE_KEY` | Browser-safe key used for Auth and RLS requests |
| `KEIKO_AUTH_PEPPER` | Same secret value used by the migration scripts |
| `KEIKO_REGISTRATION_CODE` | Code required for all new registrations |

Never place the secret key, auth pepper, or registration code in `index.html`.

Apps Script's URL Fetch service uses a browser-like `Mozilla/5.0` User-Agent.
Supabase rejects new `sb_secret_` keys from browser User-Agents, so this GAS
deployment must currently use the legacy `service_role` JWT from the Legacy API
Keys tab. Keep it server-only. Replace GAS with an Edge Function before Supabase
retires legacy keys.

## Deploy

1. Replace the Apps Script project's `Code.gs` with `gas/Code.gs`.
2. Apply `gas/appsscript.json` if the manifest is managed manually.
3. Add all five Script Properties above.
4. Deploy a new Web App version.
5. Execute the app as the deploying account and allow access to anyone using the app.
6. Keep the `/exec` URL in `index.html` unchanged when updating an existing deployment.
7. Open the app in a private window and verify login, log save, log reload, notes, and logout.

After switching production to this version, Google Sheets are migration archives
only. New application data is stored in Supabase.
