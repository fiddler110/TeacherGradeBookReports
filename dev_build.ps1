#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Build and start the grade-reports container for local development/testing.

.DESCRIPTION
    - Auto-detects the container runtime: podman (preferred on Windows) or docker (Linux/fallback).
    - Writes a temporary .env with safe local defaults if one does not exist.
    - Runs `<runtime> compose up --build -d` to rebuild the image and (re)start the container.
    - Tails container logs until Ctrl-C is pressed, then offers to stop the container.

.PARAMETER Stop
    Stop and remove the container instead of starting it.

.PARAMETER Logs
    Attach to the running container's log stream without rebuilding.

.PARAMETER Clean
    Stop the container AND remove the named data volume (full reset).

.EXAMPLE
    .\dev_build.ps1            # build + start
    .\dev_build.ps1 -Logs      # stream logs of a running container
    .\dev_build.ps1 -Stop      # stop the container
    .\dev_build.ps1 -Clean     # stop + wipe the data volume
#>

param(
    [switch]$Stop,
    [switch]$Logs,
    [switch]$Clean
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectRoot = $PSScriptRoot
Push-Location $ProjectRoot

# ── Colour helpers ────────────────────────────────────────────────────────────
function Info  ($msg) { Write-Host "  $msg" -ForegroundColor Cyan }
function Ok    ($msg) { Write-Host "  $msg" -ForegroundColor Green }
function Warn  ($msg) { Write-Host "  $msg" -ForegroundColor Yellow }
function Err   ($msg) { Write-Host "  $msg" -ForegroundColor Red }

Write-Host ""
Write-Host "=== Grade Report Generator — Local Dev Build ===" -ForegroundColor Cyan
Write-Host ""

# ── Detect container runtime (podman preferred on Windows, docker on Linux) ───
#    Podman exposes a `podman compose` sub-command (via podman-compose or the
#    built-in compose provider).  Docker Compose V2 ships as `docker compose`.
#    We resolve the CLI and compose invocation separately so every subsequent
#    call just uses $Compose (e.g. & $Compose up --build -d).
$ComposeBin  = $null   # executable name  — "podman" or "docker"
$ComposeArgs = @()     # fixed prefix args — ("compose") for both

if (Get-Command podman -ErrorAction SilentlyContinue) {
    $ComposeBin  = "podman"
    $ComposeArgs = @("compose")
    Info "Runtime: podman  (podman compose)"
} elseif (Get-Command docker -ErrorAction SilentlyContinue) {
    $ComposeBin  = "docker"
    $ComposeArgs = @("compose")
    Info "Runtime: docker  (docker compose)"
} else {
    Err "Neither podman nor docker was found on the PATH."
    Err "Install Podman Desktop (Windows) or Docker Desktop and try again."
    Pop-Location
    exit 1
}

# Convenience wrapper — runs  $ComposeBin compose <args>  and returns exit code
function Compose {
    & $ComposeBin @ComposeArgs @args
    return $LASTEXITCODE
}

# ── Handle -Stop / -Clean ─────────────────────────────────────────────────────
if ($Stop -or $Clean) {
    Info "Stopping container..."
    Compose down | Out-Null
    if ($Clean) {
        Warn "Removing data volume (reportgen_grade-reports-data)..."
        & $ComposeBin volume rm reportgen_grade-reports-data 2>$null
        Ok "Volume removed. Next start will create a fresh database."
    }
    Ok "Done."
    Pop-Location
    exit 0
}

# ── Handle -Logs ──────────────────────────────────────────────────────────────
if ($Logs) {
    Info "Streaming logs (Ctrl-C to stop)..."
    Compose logs -f
    Pop-Location
    exit 0
}

# ── Create a local .env if it doesn't exist ───────────────────────────────────
$EnvFile = Join-Path $ProjectRoot ".env"
if (-not (Test-Path $EnvFile)) {
    Warn ".env not found — creating one with local dev defaults."
    @"
# Local dev overrides — NOT for production use
ADMIN_USERNAME=admin
ADMIN_PASSWORD=changeme
ADMIN_SCHOOL=Westfield Academy
SECRET_KEY=dev-secret-key-not-for-production
COOKIE_SECURE=false
GUNICORN_WORKERS=2
TZ=America/Toronto
"@ | Set-Content $EnvFile
    Ok "Created .env with dev defaults."
} else {
    Info "Using existing .env"
}

# ── Build and start ───────────────────────────────────────────────────────────
Info "Building image and starting container..."
$rc = Compose up --build -d

if ($rc -ne 0) {
    Err "$ComposeBin compose up failed (exit $rc)."
    Pop-Location
    exit $rc
}

Ok "Container started."
Write-Host ""
Ok "Application available at: http://localhost:8000"
Ok "Default login:            admin / changeme"
Write-Host ""
Info "Streaming logs — press Ctrl-C to detach (container keeps running)."
Info "Run  .\dev_build.ps1 -Stop   to stop the container."
Info "Run  .\dev_build.ps1 -Clean  to stop and wipe the database volume."
Write-Host ""

try {
    Compose logs -f
} finally {
    Pop-Location
}
