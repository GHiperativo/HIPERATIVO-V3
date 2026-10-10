create or replace function public.platform_migration_ledger()
returns table(version text, name text, sql_md5 text, sql_chars bigint)
language sql
stable
security definer
set search_path = ''
as $$
  select
    m.version,
    coalesce(m.name, ''),
    md5(coalesce(array_to_string(m.statements, E'\n-- statement boundary --\n'), '')),
    coalesce((select sum(length(x)) from unnest(m.statements) x), 0)::bigint
  from supabase_migrations.schema_migrations as m
  order by m.version;
$$;

revoke all on function public.platform_migration_ledger() from public;
revoke all on function public.platform_migration_ledger() from anon;
revoke all on function public.platform_migration_ledger() from authenticated;
grant execute on function public.platform_migration_ledger() to service_role;
