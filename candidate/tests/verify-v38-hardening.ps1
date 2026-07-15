$ErrorActionPreference = 'Stop'

$candidate = Join-Path $PSScriptRoot '..\v38-hardened\WebApp.js'
$source = Get-Content -Raw -LiteralPath $candidate

$forbidden = @(
  'p.manutencao',
  'HIPERATIVO2026',
  "p.fn === 'setWebAppUrl'",
  "p.fn === 'debug'",
  'return _doManutencao(p.fn'
)

foreach ($pattern in $forbidden) {
  if ($source.Contains($pattern)) {
    throw "Rota administrativa ainda exposta: $pattern"
  }
}

$required = @(
  'function doGet(e)',
  'if (p.code && p.state)',
  'return _processarCallbackOAuth(e)',
  'if (p.reconectar)',
  'return _processarReconexao',
  'if (p.atleta)',
  'return _paginaAtleta',
  "if (p.salvar === 'true')"
)

foreach ($pattern in $required) {
  if (-not $source.Contains($pattern)) {
    throw "Rota pública necessária ausente: $pattern"
  }
}

Write-Output 'OK: rota administrativa removida e rotas públicas preservadas.'
