param(
    [string]$Message = "Release update"
)

$ErrorActionPreference = 'Stop'

function Stop-WorkspaceLockingProcesses {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    try {
        $escapedRoot = [Regex]::Escape($RootPath)
        $candidates = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
            $_.ProcessId -ne $PID -and (
                ($_.Name -in @('node.exe', 'electron.exe', 'Gantech Efterkalk.exe')) -and (
                    ($_.CommandLine -and $_.CommandLine -match $escapedRoot) -or
                    ($_.ExecutablePath -and $_.ExecutablePath -match $escapedRoot)
                )
            )
        }

        if ($candidates) {
            Write-Host "🛑 Stopping workspace app/server processes before build..." -ForegroundColor Yellow
            foreach ($proc in $candidates) {
                try {
                    Stop-Process -Id $proc.ProcessId -Force -ErrorAction Stop
                    Write-Host "   Stopped PID $($proc.ProcessId) ($($proc.Name))" -ForegroundColor DarkGray
                } catch {
                    Write-Host "   Could not stop PID $($proc.ProcessId) ($($proc.Name))" -ForegroundColor Yellow
                }
            }
            Start-Sleep -Milliseconds 1200
        }
    } catch {
        Write-Host "⚠️ Process cleanup skipped: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host "🚀 Gantech Efterkalk - Automated Release Publisher" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Check git status
Write-Host "📋 Step 1: Checking git status..." -ForegroundColor Yellow
$status = git status --porcelain
if ($status) {
    Write-Host "✅ Changes detected:"
    git status --short
} else {
    Write-Host "✅ No changes to commit"
}
Write-Host ""

# Step 2: Check package.json for current version
$pkg = Get-Content package.json | ConvertFrom-Json
$currentVersion = $pkg.version
Write-Host "📦 Current version in package.json: $currentVersion" -ForegroundColor Cyan
Write-Host ""

# Step 3: Git add and commit
if ($status) {
    Write-Host "📝 Step 2: Adding and committing changes..." -ForegroundColor Yellow
    git add .
    if ($LASTEXITCODE -ne 0) {
        throw "❌ Git add failed"
    }
    
    git commit -m $Message
    if ($LASTEXITCODE -ne 0) {
        throw "❌ Git commit failed"
    }
    Write-Host "✅ Changes committed" -ForegroundColor Green
    Write-Host ""
}

# Step 4: Git push
Write-Host "🔄 Step 3: Pushing to GitHub..." -ForegroundColor Yellow
git push
if ($LASTEXITCODE -ne 0) {
    throw "❌ Git push failed"
}
Write-Host "✅ Pushed to GitHub" -ForegroundColor Green
Write-Host ""

# Step 5: Version bump
Write-Host "⬆️  Step 4: Bumping version (patch)..." -ForegroundColor Yellow
npm version patch
if ($LASTEXITCODE -ne 0) {
    throw "❌ npm version patch failed"
}

# Get new version
$pkg = Get-Content package.json | ConvertFrom-Json
$newVersion = $pkg.version
Write-Host "✅ Version bumped: $currentVersion → $newVersion" -ForegroundColor Green
Write-Host ""

# Step 6: Push tags
Write-Host "🏷️  Step 5: Pushing tags..." -ForegroundColor Yellow
git push --follow-tags
if ($LASTEXITCODE -ne 0) {
    throw "❌ Git push --follow-tags failed"
}
Write-Host "✅ Tags pushed" -ForegroundColor Green
Write-Host ""

# Step 7: Build
Write-Host "🔨 Step 6: Building Windows installer..." -ForegroundColor Yellow

# Stop local workspace processes that may lock native modules like msnodesqlv8/sqlserver.node
Stop-WorkspaceLockingProcesses -RootPath $PSScriptRoot

# Remove stale artifacts for this target version to avoid NSIS "Can't open output file"
$distDir = Join-Path $PSScriptRoot 'dist'
$installerPath = Join-Path $distDir ("Gantech-Efterkalk-Setup-{0}.exe" -f $newVersion)
$blockmapPath = "$installerPath.blockmap"
$uninstallerPath = Join-Path $distDir '__uninstaller-nsis-efterkalk.exe'

foreach ($p in @($installerPath, $blockmapPath, $uninstallerPath)) {
    if (Test-Path $p) {
        try {
            Remove-Item $p -Force -ErrorAction Stop
            Write-Host "🧹 Removed stale build artifact: $p" -ForegroundColor DarkGray
        } catch {
            Write-Host "⚠️ Could not remove artifact before build: $p" -ForegroundColor Yellow
        }
    }
}

npm run build:win
if ($LASTEXITCODE -ne 0) {
    throw "❌ npm run build:win failed"
}
Write-Host "✅ Build complete" -ForegroundColor Green
Write-Host ""

# Step 8: Release to GitHub
Write-Host "📤 Step 7: Publishing to GitHub..." -ForegroundColor Yellow
$releaseScript = Join-Path $PSScriptRoot 'release-github.ps1'
& powershell -ExecutionPolicy Bypass -File $releaseScript
if ($LASTEXITCODE -ne 0) {
    throw "❌ Release script failed"
}
Write-Host "✅ Release published" -ForegroundColor Green
Write-Host ""

# Step 9: Verify release
Write-Host "✔️  Step 8: Verifying release..." -ForegroundColor Yellow
$verifyPath = "C:\Program Files\GitHub CLI\gh.exe"
if (-not (Test-Path $verifyPath)) {
    $ghCmd = Get-Command gh -ErrorAction SilentlyContinue
    if ($ghCmd) {
        $verifyPath = $ghCmd.Source
    }
}

if (Test-Path $verifyPath) {
    & $verifyPath release list --repo tbgoblin/efterkalk -L 1 | Select-Object -First 5
    Write-Host ""
}

Write-Host "================================================" -ForegroundColor Green
Write-Host "✅ SUCCESS: v$newVersion published!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Check the release on GitHub: https://github.com/tbgoblin/efterkalk/releases" -ForegroundColor Cyan
Write-Host "  2. Install the new version or wait for auto-update" -ForegroundColor Cyan
Write-Host ""
