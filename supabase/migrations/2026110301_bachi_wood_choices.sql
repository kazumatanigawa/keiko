begin;

update public.bachi_wood_price_options
set active = false,
    updated_at = now()
where wood_code not in ('hinoki', 'hou', 'maple')
  and active = true;

insert into public.bachi_wood_price_options (
  wood_code, display_name, base_pair_price, standard_length_cm,
  standard_diameter_mm, lead_months, active, sort_order, updated_at
) values
  ('hinoki', 'ヒノキ', 1800, 42, 24, 1, true, 10, now()),
  ('hou', 'ホオ', 1600, 42, 24, 1, true, 20, now()),
  ('maple', 'カエデ', 1600, 42, 24, 1, true, 30, now())
on conflict (wood_code) do update
set display_name = excluded.display_name,
    active = true,
    sort_order = excluded.sort_order,
    updated_at = now();

commit;
