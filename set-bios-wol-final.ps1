<#
    This script configures Wake-on-LAN (WOL) on supported Dell systems.

    It performs the following actions:
      • Installs the DellBIOSProvider PowerShell module if missing
      • Applies BIOS settings:
            - Disables Deep Sleep Control
            - Enables Wake-on-LAN (LANOnly)
      • Configures all physical network adapters for Wake-on-LAN:
            - Enables Advanced Properties for Magic Packet wake
            - Enables NIC power management WOL settings
      • Logs all activity using Start-Transcript and appends a final result
      • Uses a global try/catch/finally to ensure the transcript is closed cleanly

    Requirements:
      • Dell business-class system with supported BIOS
      • PowerShell run as Administrator
      • VC++ 2015–2019 x64 Redistributable installed (required by DellBIOSProvider)
#>


# ------------------------------------------------------------
# Logging Setup
# ------------------------------------------------------------
$LogDir = "C:\Logs"
$LogFile = Join-Path $LogDir "WOL-Script.log"

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

$ScriptFailed = $false

# Start transcript
Start-Transcript -Path $LogFile -Force

try {

    # ------------------------------------------------------------
    # DellBIOSProvider Installation
    # ------------------------------------------------------------
    Write-Host "Checking for DellBIOSProvider module..."
    $moduleInstalled = Get-Module -ListAvailable -Name DellBIOSProvider

    if (-not $moduleInstalled) {
        Write-Host "Installing DellBIOSProvider..."

        try {
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction Stop
        }
        catch {
            Register-PSRepository -Default
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        }

        if (-not (Get-PackageProvider -ListAvailable | Where-Object Name -eq "NuGet")) {
            Install-PackageProvider -Name NuGet -Force -Confirm:$false
        }

        Install-Module -Name DellBIOSProvider -Force -Scope AllUsers
        Write-Host "DellBIOSProvider installed."
    }
    else {
        Write-Host "DellBIOSProvider already installed."
    }

    Import-Module DellBIOSProvider -ErrorAction Stop
    Write-Host "DellBIOSProvider module imported."

    # ------------------------------------------------------------
    # BIOS Configuration
    # ------------------------------------------------------------
    try {
        Set-Item -Path "DellSmbios:\PowerManagement\DeepSleepCtrl" -Value "Disabled" -ErrorAction Stop
        Set-Item -Path "DellSmbios:\PowerManagement\WakeOnLan" -Value "LANOnly" -ErrorAction Stop
        Write-Host "BIOS settings applied."
    }
    catch {
        Write-Error "BIOS update failed: $($_.Exception.Message)"
        $ScriptFailed = $true
    }

    # ------------------------------------------------------------
    # NIC Wake-on-LAN Configuration
    # ------------------------------------------------------------
    $adapters = Get-NetAdapter -Physical | Where-Object { $_.Status -ne "Disabled" }

    foreach ($adapter in $adapters) {
        Write-Host "`nConfiguring adapter: $($adapter.Name)"

        $wolProperties = @(
            "Wake on Magic Packet",
            "Wake on Magic Packet From Power Off",
            "Wake on magic packet"
        )

        foreach ($prop in $wolProperties) {
            try {
                Set-NetAdapterAdvancedProperty -Name $adapter.Name -DisplayName $prop -DisplayValue "Enabled" -ErrorAction Stop
                Write-Host "  Enabled: $prop"
            }
            catch {}
        }

        try {
            Set-NetAdapterPowerManagement -Name $adapter.Name `
                -WakeOnMagicPacket Enabled `
                -WakeOnPattern Enabled `
                -DeviceSleepOnDisconnect Disabled `
                -ErrorAction Stop

            Write-Host "  Power management updated."
        }
        catch {
            Write-Host "  Power management update failed: $($_.Exception.Message)"
            $ScriptFailed = $true
        }
    }

    Write-Host "`nWake-on-LAN configuration complete."

}
catch {
    Write-Error "GLOBAL ERROR: $($_.Exception.Message)"
    $ScriptFailed = $true
}
finally {

    # Stop transcript before writing to log
    try { Stop-Transcript | Out-Null } catch {}

    # Final log status
    if ($ScriptFailed -eq $true -or $Error.Count -gt 0) {
        Write-Host "`nLOG RESULT: Script completed WITH ERRORS."
        Add-Content -Path $LogFile -Value "`n[$(Get-Date)] Script completed WITH ERRORS."
    }
    else {
        Write-Host "`nLOG RESULT: Script completed SUCCESSFULLY."
        Add-Content -Path $LogFile -Value "`n[$(Get-Date)] Script completed SUCCESSFULLY."
    }
}
