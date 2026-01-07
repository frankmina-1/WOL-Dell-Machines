<#
v1.9
.SYNOPSIS
    Configures Wake-on-LAN (WOL) on supported Dell systems.

.DESCRIPTION
    - Copies required PowerShell components from SYSVOL if missing
    - Imports DellBIOSProvider
    - Applies Dell BIOS WOL-related settings
    - Configures active physical network adapters for Wake-on-LAN
    - Logs only real failures with reason, computer name, and IP
    - Ignores missing settings (avoids false failures)
    - Provides a summary of total failed computers at the end

.NOTES
    Requires local administrator privileges.
#>

# ------------------------------------------------------------
# Paths and Variables
# ------------------------------------------------------------
#Path to DellBiosProvider and NuGet Modules
$SysvolModuleRoot = "\\Your\SYSVOL\\root\path\here"

$NuGetSource      = Join-Path $SysvolModuleRoot "PSModules\NuGet"
$DellModuleSource = Join-Path $SysvolModuleRoot "PSModules\DellBIOSProvider"

$NuGetDest      = "C:\Program Files\PackageManagement\ProviderAssemblies\NuGet"
$DellModuleDest = "C:\Program Files\WindowsPowerShell\Modules\DellBIOSProvider"

$LogFile = Join-Path $SysvolModuleRoot "WOL-Script.log"

$ScriptFailed = $false
$ErrorMessage = ""
$ComputerName = $env:COMPUTERNAME
$global:FailedComputers = @()

# ------------------------------------------------------------
# Get primary IPv4 address
# ------------------------------------------------------------
$IPAddress = (Get-NetIPAddress -AddressFamily IPv4 `
    | Where-Object { $_.IPAddress -notlike '169.254*' -and $_.IPAddress -ne '127.0.0.1' } `
    | Select-Object -First 1 -ExpandProperty IPAddress) 

if (-not $IPAddress) { $IPAddress = "N/A" }

try {
    Write-Host "[$ComputerName | $IPAddress] Starting Wake-on-LAN configuration..."

    # ------------------------------------------------------------
    # Copy NuGet Provider if missing
    # ------------------------------------------------------------
    if (-not (Test-Path $NuGetDest)) {
        if (-not (Test-Path $NuGetSource)) { throw "NuGet source not found: $NuGetSource" }
        New-Item -ItemType Directory -Path $NuGetDest -Force | Out-Null
        Copy-Item -Path "$NuGetSource\*" -Destination $NuGetDest -Recurse -Force
    }

    # ------------------------------------------------------------
    # Copy DellBIOSProvider Module if missing
    # ------------------------------------------------------------
    if (-not (Test-Path $DellModuleDest)) {
        if (-not (Test-Path $DellModuleSource)) { throw "DellBIOSProvider source not found: $DellModuleSource" }
        Copy-Item -Path $DellModuleSource -Destination $DellModuleDest -Recurse -Force
    }

    Import-Module DellBIOSProvider -ErrorAction Stop

    # ------------------------------------------------------------
    # Apply BIOS Settings (ignore missing)
    # ------------------------------------------------------------
    $biosFailures = @()
    $biosSettings = @{
        "DellSmbios:\PowerManagement\DeepSleepCtrl" = "Disabled"
        "DellSmbios:\PowerManagement\WakeOnLan"    = "LANOnly"
    }

    foreach ($path in $biosSettings.Keys) {
        if (Test-Path $path) {
            try {
                Set-Item $path -Value $biosSettings[$path] -ErrorAction Stop
            }
            catch { 
                $ScriptFailed = $true
                $biosFailures += "$path ($($_.Exception.Message))"
            }
        }
        # If path doesn't exist, silently skip
    }

    # ------------------------------------------------------------
    # Configure Network Adapters (ignore missing properties)
    # ------------------------------------------------------------
    $nicFailures = @()
    $adapters = Get-NetAdapter -Physical | Where-Object Status -ne "Disabled"

    foreach ($adapter in $adapters) {
        $adapterFailedProps = @()
        $wolProperties = @(
            "Wake on Magic Packet",
            "Wake on Magic Packet From Power Off",
            "Wake on magic packet"
        )

        foreach ($prop in $wolProperties) {
            $existingProp = Get-NetAdapterAdvancedProperty -Name $adapter.Name -ErrorAction SilentlyContinue | Where-Object DisplayName -eq $prop
            if ($existingProp) {
                try {
                    Set-NetAdapterAdvancedProperty -Name $adapter.Name -DisplayName $prop -DisplayValue "Enabled" -ErrorAction Stop
                }
                catch { $adapterFailedProps += $prop }
            }
            # If property doesn't exist, silently skip
        }

        # Attempt power management settings
        try {
            Set-NetAdapterPowerManagement `
                -Name $adapter.Name `
                -WakeOnMagicPacket Enabled `
                -WakeOnPattern Enabled `
                -DeviceSleepOnDisconnect Disabled `
                -ErrorAction Stop
        }
        catch { $adapterFailedProps += "PowerManagement" }

        if ($adapterFailedProps.Count -gt 0) {
            $ScriptFailed = $true
            $nicFailures += "$($adapter.Name) ($($adapterFailedProps -join ', '))"
        }
    }
}
catch {
    $ScriptFailed = $true
    $ErrorMessage = $_.Exception.Message
}
finally {
    if ($ScriptFailed) {
        $logEntry  = "[$ComputerName | $IPAddress]`n"
        $details   = @()
        if ($biosFailures.Count -gt 0) { $details += "BIOS: $($biosFailures -join ', ')" }
        if ($nicFailures.Count -gt 0)  { $details += "NIC: $($nicFailures -join '; ')" }
        if ($ErrorMessage)              { $details += "Fatal: $ErrorMessage" }
        $logEntry += "Script FAILED: " + ($details -join '; ')
        Add-Content -Path $LogFile -Value $logEntry

        $global:FailedComputers += "$ComputerName | $IPAddress"
        Write-Host "RESULT: Script FAILED. Details logged."
    }
    else {
        Write-Host "RESULT: Script completed successfully."
    }

    # ------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------
    if ($global:FailedComputers.Count -gt 0) {
        $summary = "`nTotal computers failed: $($global:FailedComputers.Count)`nFailed computers:`n" + ($global:FailedComputers -join "`n")
        Add-Content -Path $LogFile -Value $summary
        Write-Host $summary
    }
}
