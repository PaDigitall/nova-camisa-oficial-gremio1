$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '=== Configuração Google clasp — Loja Digital do Grêmio ===' -ForegroundColor Cyan
Write-Host ''

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  Write-Host 'Node.js não foi encontrado neste computador.' -ForegroundColor Yellow
  Write-Host 'Instale o Node.js LTS e execute este script novamente.'
  Write-Host 'Site oficial: https://nodejs.org/'
  exit 1
}

Write-Host 'Abrindo a autorização Google...' -ForegroundColor Green
Write-Host 'Entre com a conta que possui acesso ao Apps Script oficial.'
Write-Host ''

npx --yes @google/clasp@3 login
if ($LASTEXITCODE -ne 0) {
  throw 'O login do clasp não foi concluído.'
}

$authPath = Join-Path $HOME '.clasprc.json'
if (-not (Test-Path $authPath)) {
  throw "Credencial não encontrada em $authPath"
}

$credential = Get-Content -Raw -Path $authPath
$credential | Set-Clipboard

Write-Host ''
Write-Host 'SUCESSO: a credencial CLASPRC_JSON foi copiada para a área de transferência.' -ForegroundColor Green
Write-Host ''
Write-Host 'Agora no GitHub:' -ForegroundColor Cyan
Write-Host '1. Abra Settings > Secrets and variables > Actions.'
Write-Host '2. Clique em New repository secret.'
Write-Host '3. Nome: CLASPRC_JSON'
Write-Host '4. Cole o conteúdo que já está na área de transferência.'
Write-Host '5. Salve.'
Write-Host ''
Write-Host 'Nunca envie ou publique essa credencial em mensagens ou commits.' -ForegroundColor Yellow
