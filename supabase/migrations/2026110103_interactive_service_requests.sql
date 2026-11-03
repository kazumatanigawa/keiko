begin;

create table if not exists public.bachi_wood_price_options (
  wood_code text primary key,
  display_name text not null,
  base_pair_price integer not null check (base_pair_price >= 0),
  standard_length_cm numeric(5,1) not null default 42,
  min_length_cm numeric(5,1) not null default 30,
  max_length_cm numeric(5,1) not null default 60,
  length_surcharge_per_cm integer not null default 0 check (length_surcharge_per_cm >= 0),
  standard_diameter_mm numeric(5,1) not null default 24,
  min_diameter_mm numeric(5,1) not null default 15,
  max_diameter_mm numeric(5,1) not null default 40,
  diameter_surcharge_per_mm integer not null default 0 check (diameter_surcharge_per_mm >= 0),
  lead_months integer not null default 1 check (lead_months between 1 and 12),
  active boolean not null default true,
  sort_order integer not null default 100,
  updated_at timestamptz not null default now(),
  check (min_length_cm <= standard_length_cm and standard_length_cm <= max_length_cm),
  check (min_diameter_mm <= standard_diameter_mm and standard_diameter_mm <= max_diameter_mm)
);

create table if not exists public.bachi_quantity_price_tiers (
  min_pairs integer primary key check (min_pairs > 0),
  discount_basis_points integer not null default 0 check (discount_basis_points between 0 and 9000),
  label text not null default '',
  active boolean not null default true,
  updated_at timestamptz not null default now()
);

create table if not exists public.bachi_order_settings (
  singleton boolean primary key default true check (singleton),
  tax_rate numeric(6,5) not null default 0.10 check (tax_rate between 0 and 1),
  standard_shipping_fee integer not null default 1000 check (standard_shipping_fee >= 0),
  default_lead_months integer not null default 1 check (default_lead_months between 1 and 12),
  updated_at timestamptz not null default now()
);

create table if not exists public.bachi_order_requests (
  id uuid primary key default gen_random_uuid(),
  requester_user_id uuid not null references public.profiles(id),
  requester_name text not null,
  contact_email text not null,
  wood_code text not null,
  wood_name_snapshot text not null,
  length_cm numeric(5,1) not null check (length_cm > 0),
  diameter_mm numeric(5,1) not null check (diameter_mm > 0),
  quantity_pairs integer not null check (quantity_pairs > 0 and quantity_pairs <= 500),
  pair_price_estimate integer not null check (pair_price_estimate >= 0),
  subtotal_estimate integer not null check (subtotal_estimate >= 0),
  tax_estimate integer not null check (tax_estimate >= 0),
  shipping_fee_estimate integer not null check (shipping_fee_estimate >= 0),
  total_estimate integer not null check (total_estimate >= 0),
  estimated_delivery_date date not null,
  estimate_note text not null default '',
  status text not null default 'submitted' check (status in ('submitted', 'contacted', 'quoted', 'accepted', 'declined', 'cancelled', 'fulfilled')),
  operator_email_status text not null default 'pending' check (operator_email_status in ('pending', 'sent', 'failed')),
  customer_email_status text not null default 'pending' check (customer_email_status in ('pending', 'sent', 'failed')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.coaching_rate_options (
  rate_code text primary key,
  display_name text not null,
  fixed_fee integer check (fixed_fee is null or fixed_fee >= 0),
  hourly_rate integer check (hourly_rate is null or hourly_rate >= 0),
  duration_minutes integer check (duration_minutes is null or duration_minutes > 0),
  billing_increment_minutes integer not null default 30 check (billing_increment_minutes > 0),
  description text not null default '',
  active boolean not null default true,
  sort_order integer not null default 100,
  updated_at timestamptz not null default now(),
  check (fixed_fee is not null or hourly_rate is not null)
);

alter table public.coaching_requests
  alter column requested_date drop not null,
  alter column start_time drop not null,
  alter column end_time drop not null,
  add column if not exists rate_type text not null default 'hourly',
  add column if not exists rate_label text not null default '時間制',
  add column if not exists operator_email_status text not null default 'pending',
  add column if not exists customer_email_status text not null default 'pending';

alter table public.bachi_wood_price_options enable row level security;
alter table public.bachi_quantity_price_tiers enable row level security;
alter table public.bachi_order_settings enable row level security;
alter table public.bachi_order_requests enable row level security;
alter table public.coaching_rate_options enable row level security;

create index if not exists bachi_order_requests_created_idx
  on public.bachi_order_requests (created_at desc);
create index if not exists bachi_order_requests_requester_idx
  on public.bachi_order_requests (requester_user_id, created_at desc);

insert into public.bachi_wood_price_options (
  wood_code, display_name, base_pair_price, standard_length_cm,
  standard_diameter_mm, lead_months, sort_order
) values
  ('hinoki', 'ヒノキ', 1800, 42, 24, 1, 10),
  ('hou', 'ホオ', 1600, 42, 24, 1, 20),
  ('maple', 'カエデ', 1600, 42, 24, 1, 30)
on conflict (wood_code) do nothing;

insert into public.bachi_quantity_price_tiers (min_pairs, discount_basis_points, label)
values
  (1, 0, '通常'),
  (25, 625, '50本以上'),
  (50, 1250, '100本以上')
on conflict (min_pairs) do nothing;

insert into public.bachi_order_settings (singleton, tax_rate, standard_shipping_fee, default_lead_months)
values (true, 0.10, 1000, 1)
on conflict (singleton) do nothing;

insert into public.coaching_rate_options (
  rate_code, display_name, fixed_fee, hourly_rate, duration_minutes,
  billing_increment_minutes, description, sort_order
) values
  ('hourly', '時間制', null, 6000, null, 30, '30分単位で計算', 10),
  ('half_day', '半日（4時間）', 24000, null, 240, 30, '4時間まで', 20),
  ('full_day', '終日（8時間）', 48000, null, 480, 30, '8時間まで', 30)
on conflict (rate_code) do nothing;

update public.home_resources
set destination_url = 'keiko://bachi',
    description = '木材・長さ・太さ・数量から概算と納期を確認し、そのまま注文相談できます。',
    updated_at = now()
where source_key = 'custom-bachi-order';

update public.home_resources
set description = '場所と指導時間だけで概算を確認し、日程は返信メールで調整できます。',
    updated_at = now()
where source_key = 'coaching-request';

commit;
