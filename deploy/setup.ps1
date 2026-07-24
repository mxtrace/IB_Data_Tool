#Requires -Version 5.1
<#
.SYNOPSIS
    IB Data Tool Headless - One-click deployment script
.DESCRIPTION
    Installs Python dependencies, copies Aki skill, and validates environment.
    Run this script on a new machine after cloning the repo.
#>

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not (Test-Path "$ProjectDir\src\main.py")) {
    $ProjectDir = "$env:USERPROFILE\Desktop\IB_Data_Tool"
}

$Python = "C:\Program Files\Python313\python.exe"
$SkillDir = "$env:USERPROFILE\.aki\user_preference\akisa\skills\ib-data-tool"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " IB Data Tool Headless - Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. Check Python
Write-Host "[1/5] Checking Python 3.13..." -ForegroundColor Yellow
if (-not (Test-Path $Python)) {
    Write-Host "  ERROR: Python 3.13 not found at $Python" -ForegroundColor Red
    Write-Host "  Please install Python 3.13 from: https://www.python.org/downloads/" -ForegroundColor Red
    exit 1
}
$ver = & $Python --version 2>&1
Write-Host "  OK: $ver" -ForegroundColor Green

# 2. Install pip dependencies
Write-Host "[2/5] Installing Python dependencies..." -ForegroundColor Yellow
& $Python -m pip install --quiet requests urllib3 openpyxl xlrd pywin32 pyautogui
if ($LASTEXITCODE -ne 0) {
    Write-Host "  ERROR: pip install failed" -ForegroundColor Red
    exit 1
}
Write-Host "  OK: All dependencies installed" -ForegroundColor Green

# 3. Switch to headless branch
Write-Host "[3/5] Checking git branch..." -ForegroundColor Yellow
Push-Location $ProjectDir
$branch = git branch --show-current
if ($branch -ne "headless") {
    git checkout headless 2>&1 | Out-Null
    Write-Host "  Switched to headless branch" -ForegroundColor Green
} else {
    Write-Host "  Already on headless branch" -ForegroundColor Green
}
git pull origin headless 2>&1 | Out-Null
Pop-Location

# 4. Install Aki Skill
Write-Host "[4/5] Installing Aki skill..." -ForegroundColor Yellow
if (-not (Test-Path $SkillDir)) {
    New-Item -ItemType Directory -Force $SkillDir | Out-Null
}
Copy-Item "$ProjectDir\deploy\SKILL.md" "$SkillDir\SKILL.md" -Force
Write-Host "  OK: Skill installed to $SkillDir" -ForegroundColor Green

# 5. Validate environment
Write-Host "[5/5] Validating environment..." -ForegroundColor Yellow

$pendingDir = "$env:USERPROFILE\Desktop\Mars_LCL_Package\BookingFilePack\BCFile\PendingList"
if (Test-Path $pendingDir) {
    Write-Host "  OK: PendingList directory exists" -ForegroundColor Green
} else {
    Write-Host "  WARN: PendingList directory not found: $pendingDir" -ForegroundColor DarkYellow
    Write-Host "        (Mars tool may not be installed yet)" -ForegroundColor DarkYellow
}

$templates = "$ProjectDir\模板"
if (Test-Path $templates) {
    Write-Host "  OK: Templates directory exists" -ForegroundColor Green
} else {
    Write-Host "  WARN: Templates directory missing: $templates" -ForegroundColor DarkYellow
    Write-Host "        (Copy from another deployed machine)" -ForegroundColor DarkYellow
}

if (Test-Path "$ProjectDir\FC_Address.xlsx") {
    Write-Host "  OK: FC_Address.xlsx exists" -ForegroundColor Green
} else {
    Write-Host "  WARN: FC_Address.xlsx missing" -ForegroundColor DarkYellow
}

# Done
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Setup complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Open Outlook (ensure mailboxes are accessible)" -ForegroundColor White
Write-Host "  2. Login Firefox to trans-logistics-cn.amazon.com" -ForegroundColor White
Write-Host "  3. Tell Aki: '跑IB Data'" -ForegroundColor White
Write-Host ""
