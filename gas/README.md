# Retired GAS backend

KEIKO no longer uses Google Apps Script for login, reads, or writes. The active
backend is `supabase/functions/keiko-api`, and authenticated application data is
read and written through RLS-protected Supabase RPC functions.

`Code.gs` only returns an `api_moved` response for users who still open an old
deployment URL. Script Properties containing Supabase keys and KEIKO secrets can
be removed after the new Edge Function has been verified in production.
