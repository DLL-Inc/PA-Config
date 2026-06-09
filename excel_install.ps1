# ============================================================
#  Install-Excel-AllUsers.ps1
#  - Installs Microsoft 365 Excel silently via ODT
#  - Applies macro + Protected View settings to ALL user profiles
#
#  Requirements : Run as Administrator
#  Tested on    : Windows 10/11, Windows Server 2019/2022
# ============================================================

#Requires -RunAsAdministrator

# Colours for console output
function Write-Step { param($msg) Write-Host "`n== $msg" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "   [OK] $msg" -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "   [!!] $msg" -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "   [XX] $msg" -ForegroundColor Red }


# ============================================================
# PART 1 - INSTALL EXCEL VIA OFFICE DEPLOYMENT TOOL
# ============================================================

Write-Step "Preparing Office Deployment Tool"

$odtPath   = "C:\ODT"
$odtExe    = "$odtPath\setup.exe"
$configXml = "$odtPath\install.xml"
$odtUrl    = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18129-20030.exe"

New-Item -ItemType Directory -Force -Path $odtPath        | Out-Null
New-Item -ItemType Directory -Force -Path "$odtPath\Logs" | Out-Null

# Download ODT
Write-Host "   Downloading ODT from Microsoft..."
try {
    Invoke-WebRequest -Uri $odtUrl -OutFile "$odtPath\odt.exe" -UseBasicParsing -ErrorAction Stop
    Write-OK "ODT downloaded"
} catch {
    Write-Fail "Failed to download ODT: $_"
    exit 1
}

# Extract ODT
Write-Host "   Extracting ODT..."
Start-Process "$odtPath\odt.exe" -ArgumentList "/quiet /extract:$odtPath" -Wait
Write-OK "ODT extracted"

# Build XML config - Excel only, everything else excluded
$xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="PowerPoint" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Word" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="C:\ODT\Logs" />
</Configuration>
"@
Set-Content -Path $configXml -Value $xmlContent
Write-OK "Install config written"

# Run silent install
Write-Step "Installing Excel (this may take 5-15 min depending on internet speed)"
$install = Start-Process $odtExe -ArgumentList "/configure `"$configXml`"" -Wait -PassThru
if ($install.ExitCode -eq 0) {
    Write-OK "Excel installed successfully (exit code 0)"
} else {
    Write-Warn "Installer finished with exit code $($install.ExitCode) - check C:\ODT\Logs for details"
}


# ============================================================
# PART 2 - APPLY SETTINGS TO ALL USER PROFILES
# ============================================================
# Strategy:
#   - For currently LOADED hives  -> write directly to HKU\<SID>
#   - For NOT YET LOADED hives    -> load NTUSER.DAT, write, then unload
# ============================================================

Write-Step "Applying Excel settings to all user profiles"

$relPath = "Software\Microsoft\Office\16.0\Excel\Security"

$settings = @{
    "VBAWarnings"                = 1
    "DisableInternetFilesInPV"   = 1
    "DisableAttachmentsInPV"     = 1
    "DisableUnsafeLocationsInPV" = 1
}

$excludedSids = @("S-1-5-18", "S-1-5-19", "S-1-5-20")

$profiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" |
    Where-Object { $_.PSChildName -notin $excludedSids -and $_.PSChildName -like "S-1-5-21-*" }

if (-not $profiles) {
    Write-Warn "No user profiles found - nothing to configure."
    exit 0
}

foreach ($profile in $profiles) {
    $sid         = $profile.PSChildName
    $profilePath = $profile.ProfileImagePath
    $ntuser      = Join-Path $profilePath "NTUSER.DAT"
    $username    = Split-Path $profilePath -Leaf

    Write-Host "`n   Profile : $username ($sid)"

    $hiveLoaded = Test-Path "Registry::HKEY_USERS\$sid"
    $weLoadedIt = $false

    if (-not $hiveLoaded) {
        if (Test-Path $ntuser) {
            $null = reg load "HKU\$sid" $ntuser 2>&1
            if ($LASTEXITCODE -eq 0) {
                $hiveLoaded = $true
                $weLoadedIt = $true
                Write-Host "   Hive loaded from NTUSER.DAT"
            } else {
                Write-Warn "Could not load hive for $username - skipping (user may be locked/in use)"
                continue
            }
        } else {
            Write-Warn "NTUSER.DAT not found at $ntuser - skipping"
            continue
        }
    } else {
        Write-Host "   Hive already loaded (user is logged in)"
    }

    try {
        $regBase = "Registry::HKEY_USERS\$sid\$relPath"

        if (!(Test-Path $regBase)) {
            New-Item -Path $regBase -Force | Out-Null
        }

        foreach ($name in $settings.Keys) {
            Set-ItemProperty -Path $regBase -Name $name -Value $settings[$name] -Type DWord -ErrorAction Stop
        }

        Write-OK "Settings applied for $username"

    } catch {
        Write-Fail "Error applying settings for $username : $_"

    } finally {
        if ($weLoadedIt) {
            [GC]::Collect()
            Start-Sleep -Milliseconds 500
            $null = reg unload "HKU\$sid" 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "   Hive unloaded cleanly"
            } else {
                Write-Warn "Could not unload hive for $username - it may remain loaded until reboot"
            }
        }
    }
}


# ============================================================
# DONE
# ============================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  All done! Summary:"                        -ForegroundColor Cyan
Write-Host "  - Excel installed (check C:\ODT\Logs)"    -ForegroundColor Cyan
Write-Host "  - Settings applied to all user profiles"  -ForegroundColor Cyan
Write-Host "  - Settings take effect on next Excel open" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
