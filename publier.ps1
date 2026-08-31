<#
.SYNOPSIS
    Incrémente la version d'OutilsTs, reconstruit le paquet et le publie sur NuGet.org.

.DESCRIPTION
    Un seul geste pour livrer une version :
      1. Lit la version actuelle dans OutilsTs.csproj.
      2. Calcule la nouvelle version (patch par défaut, ou -Major / -Minor / -Version).
      3. Vérifie sur NuGet.org que cette version n'est pas déjà publiée — une version en
         ligne ne peut plus être remplacée, remonter dessus par erreur est sans retour.
      4. Demande confirmation (résumé de la version et, si fournies, des notes).
      5. Écrit la nouvelle version dans le .csproj.
      6. Reconstruit le projet en Release : GeneratePackageOnBuild produit le .nupkg.
      7. Publie le .nupkg sur NuGet.org avec la clé API fournie.

    La version n'est jamais publiée sans être passée par le .csproj : le paquet qui part
    sur NuGet.org est toujours celui que git peut retrouver à côté de son numéro.

.PARAMETER Major
    Incrémente le premier chiffre (X.0.0) et remet les deux autres à zéro.
    À réserver aux ruptures de compatibilité (signature ou classe publique qui change).

.PARAMETER Minor
    Incrémente le second chiffre (x.Y.0) et remet le troisième à zéro.
    Pour des ajouts qui ne cassent rien de ce qui existe déjà.

.PARAMETER Version
    Fixe la version exacte (ex. 2.1.0) au lieu de l'incrémenter automatiquement.

.PARAMETER ReleaseNotes
    Remplace les notes de version (PackageReleaseNotes) du .csproj. Sans ce paramètre,
    les notes existantes sont conservées telles quelles — à mettre à jour à la main avant
    publication si la version change les notes conservées valent pour l'ancienne version.

.PARAMETER ApiKey
    Clé API NuGet.org. Sans ce paramètre : la variable d'environnement NUGET_API_KEY si
    elle existe, sinon une saisie masquée est demandée. La clé n'est jamais écrite sur
    le disque ni affichée.

.PARAMETER Source
    Flux de publication. NuGet.org par défaut.

.PARAMETER SkipBuild
    Ne pas reconstruire : publie le .nupkg déjà présent pour la nouvelle version.
    Échoue proprement si ce fichier n'existe pas.

.PARAMETER DryRun
    Fait tout — vérification, .csproj, build, paquet — sauf la publication. Pour un essai
    sans clé API et sans toucher à NuGet.org.

.PARAMETER Force
    Ignore la demande de confirmation.

.EXAMPLE
    .\publier.ps1
    Incrémente le patch (2.0.1 -> 2.0.2), construit et publie.

.EXAMPLE
    .\publier.ps1 -Major -ReleaseNotes "Suppression de ProjetPDM, remplacee par PDM."
    Passe à la version majeure suivante (2.0.1 -> 3.0.0) avec de nouvelles notes.

.EXAMPLE
    .\publier.ps1 -Version 2.0.2 -DryRun
    Construit le paquet 2.0.2 sans le publier, pour vérifier que tout compile.
#>

[CmdletBinding(SupportsShouldProcess = $false)]
param(
    [switch]$Major,
    [switch]$Minor,
    [string]$Version,
    [string]$ReleaseNotes,
    [string]$ApiKey,
    [string]$Source = 'https://api.nuget.org/v3/index.json',
    [switch]$SkipBuild,
    [switch]$DryRun,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$DossierScript = $PSScriptRoot
$CheminCsproj  = Join-Path $DossierScript 'Classe outils topsolid\OutilsTs.csproj'

if (-not (Test-Path -LiteralPath $CheminCsproj)) {
    throw "Projet introuvable : $CheminCsproj"
}

# --------------------------------------------------------------------------------------
# Lecture du .csproj
# --------------------------------------------------------------------------------------

function Get-TexteCsproj {
    # UTF8Encoding($true) : le fichier a un BOM, il faut le lire et le réécrire tel quel,
    # sinon Visual Studio et MSBuild le retraitent différemment au prochain enregistrement.
    [System.IO.File]::ReadAllText($CheminCsproj, [System.Text.UTF8Encoding]::new($true))
}

function Set-TexteCsproj {
    param([string]$Texte)
    [System.IO.File]::WriteAllText($CheminCsproj, $Texte, [System.Text.UTF8Encoding]::new($true))
}

$texte = Get-TexteCsproj

$motifVersion = [regex]'<Version>(\d+)\.(\d+)\.(\d+)</Version>'
$correspondance = $motifVersion.Match($texte)
if (-not $correspondance.Success) {
    throw "Aucune balise <Version>X.Y.Z</Version> trouvée dans $CheminCsproj."
}
$versionActuelle = [version]::new(
    [int]$correspondance.Groups[1].Value,
    [int]$correspondance.Groups[2].Value,
    [int]$correspondance.Groups[3].Value)

$motifPackageId = [regex]'<PackageId>([^<]+)</PackageId>'
$correspondancePackageId = $motifPackageId.Match($texte)
$packageId = if ($correspondancePackageId.Success) { $correspondancePackageId.Groups[1].Value } else { 'OutilsTs' }

# --------------------------------------------------------------------------------------
# Calcul de la nouvelle version
# --------------------------------------------------------------------------------------

if ($Version) {
    if ($Version -notmatch '^\d+\.\d+\.\d+$') {
        throw "Format de version invalide : '$Version' (attendu : X.Y.Z)."
    }
    $nouvelleVersion = [version]$Version
}
elseif ($Major) {
    $nouvelleVersion = [version]::new($versionActuelle.Major + 1, 0, 0)
}
elseif ($Minor) {
    $nouvelleVersion = [version]::new($versionActuelle.Major, $versionActuelle.Minor + 1, 0)
}
else {
    $nouvelleVersion = [version]::new($versionActuelle.Major, $versionActuelle.Minor, $versionActuelle.Build + 1)
}

$nouvelleVersionTexte = "{0}.{1}.{2}" -f $nouvelleVersion.Major, $nouvelleVersion.Minor, $nouvelleVersion.Build

if ($nouvelleVersion -le $versionActuelle) {
    throw "La nouvelle version ($nouvelleVersionTexte) n'avance pas par rapport à l'actuelle ($versionActuelle)."
}

# --------------------------------------------------------------------------------------
# Vérification contre NuGet.org : jamais republier une version déjà en ligne
# --------------------------------------------------------------------------------------

function Test-VersionDejaPublieeSurNuGet {
    param([string]$IdPaquet, [string]$VersionTexte)

    $idMinuscule = $IdPaquet.ToLowerInvariant()
    $url = "https://api.nuget.org/v3-flatcontainer/$idMinuscule/index.json"

    try {
        $reponse = Invoke-RestMethod -Uri $url -TimeoutSec 15
        return $reponse.versions -contains $VersionTexte
    }
    catch {
        # Paquet jamais publié (404), ou NuGet.org injoignable : dans les deux cas on ne
        # bloque pas la publication sur une vérification qui n'a pas pu se faire.
        Write-Host "Vérification NuGet.org impossible ($($_.Exception.Message)) : ignorée." -ForegroundColor DarkYellow
        return $false
    }
}

Write-Host "Vérification sur NuGet.org..." -ForegroundColor DarkGray
if (Test-VersionDejaPublieeSurNuGet -IdPaquet $packageId -VersionTexte $nouvelleVersionTexte) {
    throw "La version $nouvelleVersionTexte de $packageId est déjà publiée sur NuGet.org. Une version en ligne ne peut pas être remplacée : choisissez un autre numéro."
}

# --------------------------------------------------------------------------------------
# Confirmation
# --------------------------------------------------------------------------------------

Write-Host ""
Write-Host "Paquet      : $packageId" -ForegroundColor Cyan
Write-Host "Version     : $versionActuelle  ->  $nouvelleVersionTexte" -ForegroundColor Cyan
if ($ReleaseNotes) {
    Write-Host "Notes       : $ReleaseNotes" -ForegroundColor Cyan
}
Write-Host "Flux        : $Source" -ForegroundColor Cyan
Write-Host "Mode        : $(if ($DryRun) { 'essai, sans publication' } else { 'publication réelle' })" -ForegroundColor Cyan
Write-Host ""

if (-not $Force) {
    $reponse = Read-Host "Continuer ? (o/N)"
    if ($reponse -notmatch '^[oOyY]') {
        Write-Host "Annulé." -ForegroundColor Yellow
        exit 0
    }
}

# --------------------------------------------------------------------------------------
# Mise à jour du .csproj
# --------------------------------------------------------------------------------------

$texte = $motifVersion.Replace($texte, "<Version>$nouvelleVersionTexte</Version>", 1)

if ($ReleaseNotes) {
    # Échappement XML minimal : le contenu prend place entre deux balises d'un fichier XML.
    $notesEchappees = $ReleaseNotes.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
    $motifNotes = [regex]'<PackageReleaseNotes>.*?</PackageReleaseNotes>'
    if ($motifNotes.IsMatch($texte)) {
        $texte = $motifNotes.Replace($texte, "<PackageReleaseNotes>$notesEchappees</PackageReleaseNotes>", 1)
    }
    else {
        Write-Host "Aucune balise <PackageReleaseNotes> existante : notes non écrites." -ForegroundColor DarkYellow
    }
}

Set-TexteCsproj -Texte $texte
Write-Host "Version écrite dans le .csproj." -ForegroundColor Green

# --------------------------------------------------------------------------------------
# Construction du paquet
# --------------------------------------------------------------------------------------

$dossierProjet = Split-Path -Parent $CheminCsproj
$cheminNupkg = Join-Path $dossierProjet "bin\Release\$packageId.$nouvelleVersionTexte.nupkg"

if ($SkipBuild) {
    Write-Host "Construction : ignorée (-SkipBuild)." -ForegroundColor DarkGray
    if (-not (Test-Path -LiteralPath $cheminNupkg)) {
        throw "Paquet introuvable : $cheminNupkg (rien à publier sans construction)."
    }
}
else {
    Write-Host "Construction en cours..." -ForegroundColor DarkGray

    # GenerateDocFXEnabled=false : Directory.Build.targets déclenche sinon une génération
    # de documentation en arrière-plan à chaque build d'OutilsTs. Sans effet sur le paquet,
    # mais inutile pour une publication et ça encombre la sortie de la console.
    & dotnet build $CheminCsproj -c Release -p:GenerateDocFXEnabled=false --nologo
    if ($LASTEXITCODE -ne 0) {
        throw "La construction a échoué (code $LASTEXITCODE)."
    }

    if (-not (Test-Path -LiteralPath $cheminNupkg)) {
        throw "Construction terminée mais paquet introuvable : $cheminNupkg"
    }
}

Write-Host "Paquet      : $cheminNupkg" -ForegroundColor Green

# --------------------------------------------------------------------------------------
# Publication
# --------------------------------------------------------------------------------------

if ($DryRun) {
    Write-Host ""
    Write-Host "Essai terminé (-DryRun) : rien n'a été publié." -ForegroundColor Yellow
    exit 0
}

if (-not $ApiKey) {
    $ApiKey = $env:NUGET_API_KEY
}
if (-not $ApiKey) {
    $cleSecurisee = Read-Host "Clé API NuGet.org" -AsSecureString
    $ApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($cleSecurisee))
}
if (-not $ApiKey) {
    throw "Aucune clé API fournie."
}

Write-Host "Publication en cours..." -ForegroundColor DarkGray

# --skip-duplicate : filet de sécurité en plus de la vérification faite plus haut — si la
# version existe déjà sur le flux, NuGet le signale sans faire échouer la commande.
& dotnet nuget push $cheminNupkg --source $Source --api-key $ApiKey --skip-duplicate
$codeSortie = $LASTEXITCODE

$ApiKey = $null

if ($codeSortie -ne 0) {
    throw "La publication a échoué (code $codeSortie)."
}

Write-Host ""
Write-Host "Publication terminée : $packageId $nouvelleVersionTexte est en ligne." -ForegroundColor Green
Write-Host ""
Write-Host "Reste à faire :" -ForegroundColor Cyan
Write-Host "  git add ""Classe outils topsolid\OutilsTs.csproj""" -ForegroundColor DarkGray
Write-Host "  git commit -m ""Publication OutilsTs $nouvelleVersionTexte""" -ForegroundColor DarkGray
