begin;

update public.home_resources
set status = 'inactive',
    updated_at = now()
where source_key = 'gakumeiza-radio';

update public.home_resources
set title = '古今樂明座に指導を依頼する',
    updated_at = now()
where source_key = 'coaching-request';

commit;
