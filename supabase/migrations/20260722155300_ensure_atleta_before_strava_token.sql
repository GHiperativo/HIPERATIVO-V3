-- Garante que um token Strava nunca falhe por ausência temporária do
-- cadastro-pai. Não altera nem substitui tokens existentes.

create or replace function public.ensure_atleta_before_strava_token()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
begin
  insert into public.atletas (
    ath_id, nome, strava_ok, status, updated_at
  )
  values (
    new.ath_id,
    coalesce(nullif(new.nome, ''), new.ath_id),
    'Conectado',
    'Ativo',
    now()
  )
  on conflict (ath_id) do nothing;
  return new;
end;
$$;

revoke execute on function public.ensure_atleta_before_strava_token()
from public, anon, authenticated;

grant execute on function public.ensure_atleta_before_strava_token()
to service_role;

drop trigger if exists trg_ensure_atleta_before_strava_token
on public.tokens_strava;

create trigger trg_ensure_atleta_before_strava_token
before insert or update of ath_id
on public.tokens_strava
for each row
execute function public.ensure_atleta_before_strava_token();
