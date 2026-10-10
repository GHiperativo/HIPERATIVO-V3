create or replace function public.platform_health_ping()
returns text
language sql
stable
security invoker
as $$
  select 'ok'::text;
$$;

comment on function public.platform_health_ping() is 'Read-only liveness probe for Platform Health. Returns only ok and reads no application data.';
