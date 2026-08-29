# Supabase database changes

Apply SQL migrations in filename order from the Supabase SQL Editor before
deploying a GAS version that depends on them.

## 20260829 read-path optimization

Run `migrations/20260829_optimize_read_paths.sql` first. It adds:

- a single-call authenticated session context function;
- a single-call home summary function that returns only the latest log;
- user/date and team/date indexes for practice-log reads.

After the query succeeds, replace Apps Script `Code.gs` with `gas/Code.gs` and
deploy a new Web App version. Finally reload `index.html` and verify login,
home, the first 20 logs, and the load-more button.
