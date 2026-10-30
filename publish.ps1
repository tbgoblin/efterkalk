param(
    [string]$Message = "Release update",
    [switch]$UseCurrentVersion
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Push-Location $PSScriptRoot

function Assert-Command {
    param([string]$Action)
    if ($LASTEXITCODE -ne 0) { throw "$Action failed (exit $LASTEXITCODE). Publication stopped." }
}

try {
    Write-Host 'Gantech Operations Hub - Release publisher' -ForegroundColor Cyan
    Write-Host 'Checking repository and GitHub authentication...'
    $status = git status --porcelain
    Assert-Command 'git status'
    if ($status) { $status | Write-Host }
    gh auth status
    Assert-Command 'GitHub authentication'
    # Older local tags may intentionally differ; fetching main must not rewrite them.
    git fetch origin main --no-tags
    Assert-Command 'git fetch'
    $branch = git branch --show-current
    Assert-Command 'Reading current branch'
    if ($branch -ne 'main') { throw 'Publish from main after integrating the intended changes.' }
    $behind = git rev-list --count HEAD..origin/main
    Assert-Command 'Checking origin/main'
    if ([int]$behind -gt 0) { throw 'origin/main contains new commits. Integrate them before publishing.' }

    Write-Host 'Running tests...' -ForegroundColor Yellow
    npm test
    Assert-Command 'Tests'

    if (-not $UseCurrentVersion) {
        npm version patch --no-git-tag-version
        Assert-Command 'Version bump'
    }
    $version = [string](Get-Content package.json -Raw | ConvertFrom-Json).version
    if ($version -notmatch '^\d+\.\d+\.\d+$') { throw 'Invalid package version.' }
    $tag = "v$version"
    git show-ref --verify --quiet "refs/tags/$tag"
    if ($LASTEXITCODE -eq 0) { throw "$tag already exists. Bump the version before publishing." }
    if ($LASTEXITCODE -ne 1) { Assert-Command 'Checking release tag' }

    # Fresh output avoids locks from an already-running win-unpacked application.
    $buildRelative = "dist/release-$version-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    $buildPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot $buildRelative))
    $distPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'dist'))
    if (-not $buildPath.StartsWith($distPath + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Build directory must stay inside this workspace dist directory.'
    }
    Write-Host "Building $tag..." -ForegroundColor Yellow
    # Invoke the installed CLI directly: PowerShell's npm.ps1 can consume the --
    # separator and accidentally turn "never" into a Windows build target.
    node (Join-Path $PSScriptRoot 'node_modules/electron-builder/cli.js') --win nsis --publish never "--config.directories.output=$buildRelative"
    Assert-Command 'Windows installer build (use -UseCurrentVersion to retry the prepared version)'
    node scripts/check-release-package.js $buildPath $version
    Assert-Command 'Packaged application validation'

    $installerName = "Gantech-Operations-Hub-Setup-$version.exe"
    foreach ($name in @($installerName, "$installerName.blockmap", 'latest.yml')) {
        $source = Join-Path $buildPath $name
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing release artifact: $source" }
        Copy-Item -LiteralPath $source -Destination (Join-Path $distPath $name) -Force
    }

    Write-Host 'Committing the verified source and version...' -ForegroundColor Yellow
    # Exclusion pathspecs naming .gitignored folders make git add exit 1, so unstage those paths afterwards instead.
    git add --all -- .
    Assert-Command 'git add'
    git reset -q -- data/lagerliste docs/audits
    Assert-Command 'Unstaging local data'
    git commit -m "$Message ($tag)"
    Assert-Command 'git commit'
    git tag -a $tag -m "Gantech Operations Hub $tag"
    Assert-Command 'git tag'
    git push --atomic origin main "refs/tags/$tag"
    Assert-Command 'git push'

    & (Join-Path $PSScriptRoot 'release-github.ps1')
    Assert-Command 'GitHub release'
    gh release view $tag --repo tbgoblin/efterkalk --json url,tagName,isDraft,assets
    Assert-Command 'Release verification'
    Write-Host "Published $tag successfully." -ForegroundColor Green
} finally {
    Pop-Location
}
