#
# Lance les serveurs DocFX et ouvre les sites dans le navigateur.
# Double-cliquez ou lancez avec : pwsh -File .\start-docs.ps1
#

$docfx        = "$env:USERPROFILE\.dotnet\tools\docfx.exe"
$siteOutilsTs = "$PSScriptRoot\Classe outils topsolid\_site_outilsts"
$siteIFix     = "$PSScriptRoot\..\iFixInvalidity\_site_ifix"

if (-not (Test-Path $docfx)) {
	Write-Error "docfx.exe introuvable dans $docfx."
	Write-Host "Lancez d'abord : C:\Users\flfa\source\repos\NorinRad39\iFixInvalidity\install-docfx-from-nuget.ps1"
	pause; exit 1
}

# ── OutilsTs ────────────────────────────────────────────────────────────────
if (Test-Path $siteOutilsTs) {
	Write-Host "[DocFX] Démarrage serveur OutilsTs sur http://localhost:8081 ..."
	Start-Process -FilePath $docfx -ArgumentList "serve `"$siteOutilsTs`" --port 8081" -WindowStyle Hidden
	Start-Sleep -Milliseconds 800
	Start-Process "http://localhost:8081"
} else {
	Write-Warning "Site OutilsTs introuvable. Lancez d'abord generate-docs.ps1 -Project OutilsTs."
}

# ── iFixInvalidity ──────────────────────────────────────────────────────────
if (Test-Path $siteIFix) {
	Write-Host "[DocFX] Démarrage serveur iFixInvalidity sur http://localhost:8080 ..."
	Start-Process -FilePath $docfx -ArgumentList "serve `"$siteIFix`" --port 8080" -WindowStyle Hidden
	Start-Sleep -Milliseconds 800
	Start-Process "http://localhost:8080"
} else {
	Write-Warning "Site iFixInvalidity introuvable. Lancez d'abord generate-docs.ps1 -Project iFixInvalidity."
}

Write-Host ""
Write-Host "Sites disponibles :"
Write-Host "  OutilsTs       -> http://localhost:8081"
Write-Host "  iFixInvalidity -> http://localhost:8080"
Write-Host ""
Write-Host "Fermez cette fenêtre pour arrêter les serveurs."
pause
