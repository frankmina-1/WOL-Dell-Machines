<#
    Wake-on-LAN Enablement Script for Dell Systems
    ------------------------------------------------
    This script performs the following actions:

    1. Installs NuGet (required for PowerShellGet)
    2. Installs the SetBios module
    3. Applies BIOS configuration from CSV
    4. Enables Wake-on-LAN settings for all physical NICs
    5. Waits 5 minutes
    6. Uninstalls SetBios and the NuGet provider

    Requirements:
    - Dell machine with supported BIOS
    - VC++ 2015–2019 x64 Redistributable must be installed
    - BIOS CSV file located at C:\bios\wol-bios-changes.csv
#>


# ------------------------------------------------------------
# Logging Initialization
# ------------------------------------------------------------
$LogDir = "C:\Logs"
$LogFile = Join-Path $LogDir "WOL-Script.log"

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

Start-Transcript -Path $LogFile -Force
$ScriptFailed = $false

try {


# ------------------------------------------------------------
# 1. Install NuGet Package Provider (required for PSGet)
# ------------------------------------------------------------

$packageProvider = "NuGet"

if (-not (Get-PackageProvider -ListAvailable -Name $packageProvider)) {
    Write-Host "Installing Package Provider: $packageProvider ..."
    Install-PackageProvider -Name $packageProvider -MinimumVersion 2.8.5.201 -Force -Scope AllUsers
    Write-Host "$packageProvider installed successfully."
}

# ------------------------------------------------------------
# 2. Install SetBios module if not already installed
# ------------------------------------------------------------

$moduleName = "SetBios"

if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    Write-Host "Installing PowerShell module: $moduleName ..."
    Install-Module -Name $moduleName -Force -Scope AllUsers
    Write-Host "$moduleName installed successfully."
}

# ------------------------------------------------------------
# 3. Validate BIOS CSV path
# ------------------------------------------------------------

$csvPath = "C:\bios\wol-bios-changes.csv"

if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at: $csvPath"
    exit 1
}

# ------------------------------------------------------------
# 4. Apply BIOS Settings
#    (Suppress warnings about unsupported UEFI variables)
# ------------------------------------------------------------

Write-Host "Applying BIOS settings from CSV..."
$oldEA = $ErrorActionPreference
$ErrorActionPreference = "SilentlyContinue"

Set-Bios -CSV $csvPath

$ErrorActionPreference = $oldEA
Write-Host "BIOS settings processed."

# ------------------------------------------------------------
# 5. Enable Wake-on-LAN for all physical adapters
# ------------------------------------------------------------

$adapters = Get-NetAdapter -Physical | Where-Object { $_.Status -ne "Disabled" }

foreach ($adapter in $adapters) {

    Write-Host "`nConfiguring adapter: $($adapter.Name)"

    # Common WoL-related advanced properties across NIC vendors
    $wolProperties = @(
        "Wake on Magic Packet",
        "Wake on Magic Packet From Power Off",
        "Wake on magic packet"
    )

    foreach ($prop in $wolProperties) {
        try {
            Set-NetAdapterAdvancedProperty -Name $adapter.Name `
                -DisplayName $prop `
                -DisplayValue "Enabled" `
                -ErrorAction Stop

            Write-Host "  Enabled: $prop"
        }
        catch {
            # Not all properties exist on all adapters — ignore quietly
        }
    }

    # Power Management tab settings
    try {
        Set-NetAdapterPowerManagement -Name $adapter.Name `
            -WakeOnMagicPacket Enabled `
            -WakeOnPattern Enabled `
            -DeviceSleepOnDisconnect Disabled `
            -ErrorAction Stop

        Write-Host "  Power-management settings updated."
    }
    catch {
        Write-Host "  Could not modify power-management settings: $($_.Exception.Message)"
    }
}

Write-Host "`nWake-on-LAN configuration complete."

# ------------------------------------------------------------
# 6. Pause for 5 minutes before cleanup
# ------------------------------------------------------------

Write-Host "`nWaiting 5 minutes before cleanup..."
#Start-Sleep -Seconds 300
Write-Host "Resuming cleanup operations... `n"

# ------------------------------------------------------------
# 7. Uninstall modules (SetBios)
# ------------------------------------------------------------

$modulesToRemove = @("SetBios")

foreach ($mod in $modulesToRemove) {

    # Remove from memory if currently imported
    if (Get-Module -Name $mod) {
        Write-Host "Removing loaded module from memory: $mod"
        Remove-Module -Name $mod -Force -ErrorAction SilentlyContinue
    }

    # Uninstall if installed on system
    if (Get-Module -ListAvailable -Name $mod) {
        try {
            Write-Host "Uninstalling module: $mod"
            Uninstall-Module -Name $mod -AllVersions -Force -ErrorAction Stop
            Write-Host "Module uninstalled: $mod"
        }
        catch {
            Write-Host "Could not uninstall ${mod}: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "Module not installed: $mod"
    }
}

Write-Host "`nModule cleanup complete."

# ------------------------------------------------------------
# 8. Remove NuGet package provider
# ------------------------------------------------------------

Write-Host "`nRemoving NuGet package provider..."

$providerPaths = @(
    "C:\Program Files\PackageManagement\ProviderAssemblies\nuget",
    "C:\Program Files\PackageManagement\ProviderCache\nuget"
)

foreach ($path in $providerPaths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Removed: $path"
    }
}


Write-Host "`nScript completed successfully."

}
catch {
    Write-Host "`nERROR: Script encountered an unexpected exception."
    Write-Host $_.Exception.Message
    $ScriptFailed = $true
}
finally {

    # Stop transcript first so the file is released
    try {
        Stop-Transcript | Out-Null
    }
    catch {
        # Ignore errors if transcript already stopped
    }

    # Now the file is free to write to
    if ($ScriptFailed -eq $true -or $Error.Count -gt 0) {
        Write-Host "`nLOG RESULT: Script completed **WITH ERRORS**."
        Add-Content -Path $LogFile -Value "`n[$(Get-Date)] Script completed WITH ERRORS."
    }
    else {
        Write-Host "`nLOG RESULT: Script completed **SUCCESSFULLY**."
        Add-Content -Path $LogFile -Value "`n[$(Get-Date)] Script completed SUCCESSFULLY."
    }
}