$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dist = Join-Path $root "dist"
$version = (Get-Content (Join-Path $root "manifest.json") -Raw | ConvertFrom-Json).version
$xpi = Join-Path $dist "zotero-translate-$version.xpi"
$zip = Join-Path $dist "zotero-translate-$version.zip"
$temp = Join-Path $dist "package"

if (Test-Path $temp) {
  Remove-Item -LiteralPath $temp -Recurse -Force
}
if (!(Test-Path $dist)) {
  New-Item -ItemType Directory -Path $dist | Out-Null
}
if (Test-Path $xpi) {
  Remove-Item -LiteralPath $xpi -Force
}
if (Test-Path $zip) {
  Remove-Item -LiteralPath $zip -Force
}

New-Item -ItemType Directory -Path $temp | Out-Null

Copy-Item -LiteralPath (Join-Path $root "manifest.json") -Destination $temp
Copy-Item -LiteralPath (Join-Path $root "bootstrap.js") -Destination $temp
Copy-Item -LiteralPath (Join-Path $root "zotero-translate.js") -Destination $temp
Copy-Item -LiteralPath (Join-Path $root "prefs.js") -Destination $temp
Copy-Item -LiteralPath (Join-Path $root "preferences.xhtml") -Destination $temp
Copy-Item -LiteralPath (Join-Path $root "README.md") -Destination $temp
Copy-Item -LiteralPath (Join-Path $root "defaults") -Destination $temp -Recurse

Compress-Archive -Path (Join-Path $temp "*") -DestinationPath $zip -Force
Move-Item -LiteralPath $zip -Destination $xpi
Remove-Item -LiteralPath $temp -Recurse -Force

Write-Host "Built $xpi"
