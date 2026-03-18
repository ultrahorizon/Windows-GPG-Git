param(
    [Parameter(Mandatory = $true)]
    [string]$GitUserName,

    [Parameter(Mandatory = $true)]
    [string]$GitUserEmail
)

$ErrorActionPreference = "Stop"

$GnuPGConfigDir = "$env:APPDATA\gnupg"
$GpgExe = "C:\Program Files\GnuPG\bin\gpg.exe"
$GpgConfExe = "C:\Program Files\GnuPG\bin\gpgconf.exe"
$SshExe = "C:\Windows\System32\OpenSSH\ssh.exe"

$GpgConfFile = "$GnuPGConfigDir\gpg.conf"
$GpgAgentConfFile = "$GnuPGConfigDir\gpg-agent.conf"

# -- Helpers ------------------------------------------------------------------

function Write-Step {
    param([string]$Message)
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Exit-WithError {
    param([string]$Message)
    Write-Host "`nERROR: $Message" -ForegroundColor Red
    exit 1
}

function Invoke-Gpg {
    param([string[]]$Arguments)
    & $GpgExe @Arguments
}

function Convert-ToGitPath {
    param([string]$Path)
    return $Path -replace "\\", "/"
}

# -- 1. Install or verify dependencies ----------------------------------------

Write-Step "Checking dependencies..."

# Git
$GitInstalled = $null -ne (Get-Command "git" -ErrorAction SilentlyContinue)
if (-not $GitInstalled) {
    Write-Host "  Git not found, installing..."
    winget install -e --id Git.Git
    if ($LASTEXITCODE -ne 0) {
        Exit-WithError "Failed to install Git. Please install it manually: winget install -e --id Git.Git"
    }
} else {
    Write-Host "  Git: found." -ForegroundColor Green
}

# GPG
if (-not (Test-Path $GpgExe)) {
    Write-Host "  Gpg4win not found, installing..."
    winget install -e --id GnuPG.Gpg4win
    if ($LASTEXITCODE -ne 0) {
        Exit-WithError "Failed to install Gpg4win. Please install it manually: winget install -e --id GnuPG.Gpg4win"
    }
} else {
    Write-Host "  Gpg4win: found." -ForegroundColor Green
}

# OpenSSH
if (-not (Test-Path $SshExe)) {
    Exit-WithError "OpenSSH not found at '$SshExe'. Please install it via Windows Optional Features and re-run this script.`nSee: https://learn.microsoft.com/en-us/windows-server/administration/openssh/openssh_install_firstuse"
} else {
    Write-Host "  OpenSSH: found." -ForegroundColor Green
}

# -- 2. Check and disable Windows ssh-agent service ---------------------------

Write-Step "Checking Windows ssh-agent service..."
$SshAgentService = Get-Service -Name "ssh-agent" -ErrorAction SilentlyContinue

if ($SshAgentService -and $SshAgentService.Status -eq "Running") {
    Write-Host "  Windows ssh-agent service is running. Elevating to stop and disable it..." -ForegroundColor Yellow
    Start-Process powershell -Verb RunAs -Wait -ArgumentList @(
        "-NoProfile", "-Command",
        "Stop-Service ssh-agent; Set-Service ssh-agent -StartupType Disabled"
    )
    $SshAgentService.Refresh()
    if ($SshAgentService.Status -ne "Stopped") {
        Exit-WithError "Failed to stop the Windows ssh-agent service. Please stop and disable it manually and re-run."
    }
    Write-Host "  Windows ssh-agent service stopped and disabled." -ForegroundColor Green
} else {
    Write-Host "  Windows ssh-agent service is not running." -ForegroundColor Green
}

# -- 3. Check config files do not already exist -------------------------------

Write-Step "Checking for existing config files..."

$ExistingFiles = @()
if (Test-Path $GpgConfFile) { $ExistingFiles += $GpgConfFile }
if (Test-Path $GpgAgentConfFile) { $ExistingFiles += $GpgAgentConfFile }

if ($ExistingFiles.Count -gt 0) {
    Write-Host "The following config files already exist:" -ForegroundColor Yellow
    $ExistingFiles | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    Exit-WithError "Aborting to avoid overwriting existing config files. Please remove or back them up and re-run."
} else {
    Write-Host "  No existing config files found." -ForegroundColor Green
}

# -- 4. YubiKey insertion prompt ----------------------------------------------

Write-Step "YubiKey setup"
Write-Host "Please insert your YubiKey, then press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host "Continuing..." -ForegroundColor Green

# -- 5. GPG card status + fetch public key ------------------------------------

Write-Step "Reading GPG card status..."
$CardStatus = Invoke-Gpg @("--card-status") | Out-String

$UrlMatch = $CardStatus | Select-String -Pattern "URL of public key\s*:\s*(https?://\S+)"
if (-not $UrlMatch) {
    Exit-WithError "Could not find a public key URL in gpg --card-status output. Is your YubiKey inserted and configured?"
}
$KeyUrl = $UrlMatch.Matches[0].Groups[1].Value
Write-Host "  Found public key URL: $KeyUrl" -ForegroundColor Green

Write-Step "Fetching public key from $KeyUrl..."
Invoke-Gpg @("--fetch-keys", $KeyUrl)
Write-Host "Public key fetched." -ForegroundColor Green

# -- 6. Extract key ID (signing) ----------------------------------------------

Write-Step "Extracting key information..."

$KeyIdOutput = Invoke-Gpg @("-K", "--keyid-format=long") | Out-String
$KeyIdLines = $KeyIdOutput -split "`r?`n"

$SigningKeyId = $null
foreach ($Line in $KeyIdLines) {
    if ($Line -match "ssb.*?\/([0-9A-Fa-f]+)\s.*\[S\]") {
        $SigningKeyId = $Matches[1]
        break
    }
}

if (-not $SigningKeyId) {
    Exit-WithError "Could not find the key ID for the [S] (signing) subkey. Check your GPG key has a signing subkey."
}
Write-Host "  Signing key ID: $SigningKeyId" -ForegroundColor Green

# -- 7. Write GPG config files ------------------------------------------------

Write-Step "Writing GPG config files to $GnuPGConfigDir..."

if (-not (Test-Path $GnuPGConfigDir)) {
    New-Item -ItemType Directory -Path $GnuPGConfigDir | Out-Null
}

Set-Content -Path $GpgConfFile -Value "use-agent"
Write-Host "  Written: gpg.conf"

Set-Content -Path $GpgAgentConfFile -Value "enable-win32-openssh-support`nenable-ssh-support`nenable-putty-support"
Write-Host "  Written: gpg-agent.conf"

# -- 8. Restart GPG agent -----------------------------------------------------

Write-Step "Restarting GPG agent..."
& $GpgConfExe --kill gpg-agent
& $GpgConfExe --launch gpg-agent
Write-Host "GPG agent restarted." -ForegroundColor Green

# -- 9. Configure Git ---------------------------------------------------------

Write-Step "Configuring Git..."

$GitConfigs = @(
    @{ Key = "core.sshCommand";  Value = (Convert-ToGitPath $SshExe) },
    @{ Key = "gpg.program";      Value = (Convert-ToGitPath $GpgExe) },
    @{ Key = "user.name";        Value = $GitUserName },
    @{ Key = "user.email";       Value = $GitUserEmail },
    @{ Key = "commit.gpgsign";   Value = "true" },
    @{ Key = "tag.gpgsign";      Value = "true" },
    @{ Key = "user.signingkey";  Value = $SigningKeyId }
)

foreach ($Config in $GitConfigs) {
    git config --global $Config.Key $Config.Value
    Write-Host "  $($Config.Key) = $($Config.Value)"
}

Write-Host "Git configured." -ForegroundColor Green

# -- 10. Register GPG agent on login ------------------------------------------

Write-Step "Registering GPG agent on login..."

$TaskName = "GPG Agent"
$ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($ExistingTask) {
    Write-Host "  Scheduled task '$TaskName' already exists, skipping." -ForegroundColor Yellow
} else {
    $Action = New-ScheduledTaskAction -Execute $GpgConfExe -Argument "--launch gpg-agent"
    $Trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $Settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit ([TimeSpan]::Zero) `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries
    $Principal = New-ScheduledTaskPrincipal `
        -UserId $env:USERNAME `
        -LogonType Interactive `
        -RunLevel Limited

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $Action `
        -Trigger $Trigger `
        -Settings $Settings `
        -Principal $Principal | Out-Null

    if ($LASTEXITCODE -ne 0) {
        Exit-WithError "Failed to register GPG agent startup task."
    }
    Write-Host "  Scheduled task '$TaskName' registered." -ForegroundColor Green
}

# -- Done ---------------------------------------------------------------------

Write-Host "`nSetup complete!" -ForegroundColor Green

