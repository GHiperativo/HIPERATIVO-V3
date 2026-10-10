create or replace function private.platform_migration_ledger_internal()
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

revoke all on function private.platform_migration_ledger_internal() from public;
revoke all on function private.platform_migration_ledger_internal() from anon;
revoke all on function private.platform_migration_ledger_internal() from authenticated;
grant usage on schema private to service_role;
grant execute on function private.platform_migration_ledger_internal() to service_role;

create or replace function public.platform_migration_ledger()
returns table(version text, name text, sql_md5 text, sql_chars bigint)
language sql
stable
security invoker
set search_path = ''
as $$
  select * from private.platform_migration_ledger_internal();
$$;

revoke all on function public.platform_migration_ledger() from public;
revoke all on function public.platform_migration_ledger() from anon;
revoke all on function public.platform_migration_ledger() from authenticated;
grant execute on function public.platform_migration_ledger() to service_role;
