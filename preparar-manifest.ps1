param(
  [Parameter(Mandatory=$true)]
  [string]$BaseUrl
)

$BaseUrl = $BaseUrl.Trim().TrimEnd('/')
if (-not $BaseUrl.StartsWith('https://')) {
  throw 'A URL de hospedagem precisa começar com https://'
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$token = '__ADDIN_BASE_URL__'

$manifestTemplate = Get-Content (Join-Path $root 'manifest.template.xml') -Raw
$manifest = $manifestTemplate.Replace($token, $BaseUrl)
Set-Content -Path (Join-Path $root 'manifest.xml') -Value $manifest -Encoding UTF8

$config = @"
/* Endereço HTTPS da hospedagem do suplemento. */
window.ADDIN_BASE_URL = "$BaseUrl";
"@
Set-Content -Path (Join-Path $root 'config.js') -Value $config -Encoding UTF8

Write-Host 'manifest.xml e config.js preparados para:' $BaseUrl -ForegroundColor Green
