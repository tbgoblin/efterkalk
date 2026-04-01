$ErrorActionPreference = 'Stop'

$repo = 'tbgoblin/efterkalk'
$packagePath = Join-Path $PSScriptRoot 'package.json'
$distPath = Join-Path $PSScriptRoot 'dist'
$latestYmlPath = Join-Path $distPath 'latest.yml'

$ghPath = $null
$ghCmd = Get-Command gh -ErrorAction SilentlyContinue
if ($ghCmd) {
    $ghPath = $ghCmd.Source
}

if (-not $ghPath) {
    $commonGhPaths = @(
        'C:\Program Files\GitHub CLI\gh.exe',
        "$env:LOCALAPPDATA\Programs\GitHub CLI\gh.exe"
    )
    foreach ($p in $commonGhPaths) {
        if (Test-Path $p) {
            $ghPath = $p
            break
        }
    }
}

if (-not $ghPath) {
    throw 'GitHub CLI non trovato. Installa con: winget install --id GitHub.cli'
}

function Invoke-Gh {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Args
    )

    & $ghPath @Args
    if ($LASTEXITCODE -ne 0) {
        throw "gh command failed: gh $($Args -join ' ')"
    }
}

$authOk = $false
try {
    Invoke-Gh -Args @('auth', 'status')
    $authOk = $true
} catch {
    $authOk = $false
}

if (-not $authOk) {
    throw 'Non autenticato su GitHub CLI. Esegui: gh auth login'
}

if (-not (Test-Path $packagePath)) { throw 'package.json non trovato.' }
if (-not (Test-Path $latestYmlPath)) { throw 'dist/latest.yml non trovato. Esegui prima: npm run build:win' }

$package = Get-Content $packagePath -Raw | ConvertFrom-Json
$version = [string]$package.version
if ([string]::IsNullOrWhiteSpace($version)) { throw 'Versione non valida in package.json' }

$exeName = "Gantech-Efterkalk-Setup-$version.exe"
$blockmapName = "$exeName.blockmap"
$exePath = Join-Path $distPath $exeName
$blockmapPath = Join-Path $distPath $blockmapName

if (-not (Test-Path $exePath)) { throw "File mancante: $exeName. Esegui: npm run build:win" }
if (-not (Test-Path $blockmapPath)) { throw "File mancante: $blockmapName. Esegui: npm run build:win" }

$latestYml = Get-Content $latestYmlPath -Raw
if ($latestYml -notmatch "version:\s*$([regex]::Escape($version))") {
    throw "latest.yml non allineato: versione package.json=$version"
}
if ($latestYml -notmatch [regex]::Escape($exeName)) {
    throw "latest.yml non allineato: non contiene $exeName"
}

$tag = "v$version"
$title = "Gantech Efterkalk $tag"

$releaseExists = $true
try {
    Invoke-Gh -Args @('release', 'view', $tag, '--repo', $repo)
} catch {
    $releaseExists = $false
}

if ($releaseExists) {
    Invoke-Gh -Args @('release', 'upload', $tag, $exePath, $blockmapPath, $latestYmlPath, '--repo', $repo, '--clobber')
    Write-Host "Release $tag aggiornata con gli asset latest." -ForegroundColor Green
} else {
    Invoke-Gh -Args @('release', 'create', $tag, $exePath, $blockmapPath, $latestYmlPath, '--repo', $repo, '--title', $title, '--notes', "Release $tag")
    Write-Host "Release $tag creata." -ForegroundColor Green
}

Write-Host "URL: https://github.com/$repo/releases/tag/$tag" -ForegroundColor Cyan
