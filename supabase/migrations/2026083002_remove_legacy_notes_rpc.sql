begin;

drop function if exists public.get_keiko_notes(uuid, integer);

notify pgrst, 'reload schema';

commit;
