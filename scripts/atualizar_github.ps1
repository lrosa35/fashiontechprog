$ErrorActionPreference = 'Stop'
param(
  [string]$RepoUrl = 'https://github.com/lrosa35/fashiontechprog.git',
  [string]$TargetDir = 'fashiontechprog',
  [string]$Branch = 'main'
)

Set-Location -Path (Join-Path $PSScriptRoot '..')

function Download-Zip($url, $outPath) {
  Invoke-WebRequest -Uri $url -OutFile $outPath -UseBasicParsing
}

function Expand-Zip-Safe($zipPath, $destDir) {
  if (Test-Path $destDir) { Remove-Item -Recurse -Force $destDir }
  Expand-Archive -Path $zipPath -DestinationPath $destDir -Force
}

$zipUrl = $RepoUrl.TrimEnd('.git') + '/archive/refs/heads/' + $Branch + '.zip'
$tmp = Join-Path $env:TEMP ("repo_" + [guid]::NewGuid().ToString() + '.zip')
$unzipDir = Join-Path $env:TEMP ("unz_" + [guid]::NewGuid().ToString())

try {
  Write-Host "Baixando ZIP: $zipUrl"
  Download-Zip $zipUrl $tmp
  Write-Host "Extraindo ZIP..."
  Expand-Zip-Safe $tmp $unzipDir
  # O ZIP do GitHub cria uma pasta <repo>-<branch>
  $inner = Get-ChildItem -Path $unzipDir | Where-Object { $_.PsIsContainer } | Select-Object -First 1
  if (-not $inner) { throw 'Estrutura de ZIP inesperada.' }
  if (Test-Path $TargetDir) { Remove-Item -Recurse -Force $TargetDir }
  Move-Item -Path $inner.FullName -Destination $TargetDir
  Write-Host "[ok] Repositório pronto em $TargetDir"
}
catch {
  Write-Error $_
  exit 1
}
finally {
  if (Test-Path $tmp) { Remove-Item $tmp -Force }
  if (Test-Path $unzipDir) { Remove-Item $unzipDir -Recurse -Force }
}

