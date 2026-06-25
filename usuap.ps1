<#
.SYNOPSIS
    Registers the local device with Windows Autopilot via Microsoft Graph.

.DESCRIPTION
    Installs prerequisites, connects to Microsoft Graph interactively, collects
    the local device's hardware hash, and uploads it to the Autopilot import
    service. Optionally applies a group tag and/or assigns a device name.

.PARAMETER GroupTag
    Optional group tag to apply to the Autopilot device record.

.PARAMETER AssignedComputerName
    Optional override for the computer name. If not provided, the name is
    automatically generated as DPCLAS-<SerialNumber> (truncated to 15 chars).

.PARAMETER UseWAM
    Re-enables the new (WAM) sign-in method, which allows use of security keys
    but shows the "all apps" page that interferes with enrollment.

.PARAMETER Reboot
    Reboots the computer after completion. Defaults to a 10 second delay unless
    -RebootDelay is set. Cannot be used with -Shutdown.

.PARAMETER Shutdown
    Shuts down the computer after completion. Defaults to a 10 second delay unless
    -ShutdownDelay is set. Cannot be used with -Reboot.

.PARAMETER RebootDelay
    Optional delay in seconds before reboot (only applies if -Reboot is specified).

.PARAMETER ShutdownDelay
    Optional delay in seconds before shutdown (only applies if -Shutdown is specified).

.PARAMETER AutoRemove
    Delete this script after a successful upload. This also happens automatically
    whenever -Reboot or -Shutdown is used.

.EXAMPLE
    .\usuap.ps1 -GroupTag "DPINFT"
    .\usuap.ps1 -GroupTag "DPINFT" -AssignedComputerName "CUSTOM-PC-01"
    .\usuap.ps1 -GroupTag "DPINFT" -Reboot -RebootDelay 5
    .\usuap.ps1 -GroupTag "DPINFT" -Shutdown
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$GroupTag = "",
    [Parameter(Mandatory = $false)]
    [string]$AssignedComputerName = "",
    [Parameter(Mandatory = $false)]
    [switch]$UseWAM,
    [switch]$Reboot,
    [switch]$Shutdown,
    [Parameter(Mandatory = $false)]
    [int]$RebootDelay,
    [Parameter(Mandatory = $false)]
    [int]$ShutdownDelay,
    [switch]$AutoRemove
)

$CLIENT_ID = "87d8aa30-7d13-4f37-8914-ebe8c7097789"
$TENANT_ID = "ac352f9b-eb63-4ca2-9cf9-f4c40047ceff"

# --- Validate parameters ---------------------------------------------------

if ($GroupTag) {
    if ($GroupTag -notmatch 'DP[A-Z]{3,4}') {
        Write-Error "GroupTag must contain 'DP' followed by 3 or 4 uppercase letters (e.g. DPABC, Finance-DPABCD)."
        exit 1
    }
    if ($GroupTag -match '^-' -or $GroupTag -match '-$') {
        Write-Error "GroupTag must not start or end with a dash."
        exit 1
    }
}

if ($Reboot -and $Shutdown) {
    Write-Error "Cannot specify both -Reboot and -Shutdown."
    exit 1
}

# --- Prerequisites ---------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null

if (-not (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)) {
    Write-Host "Installing Microsoft.Graph.Authentication module..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph.Authentication -Force -Scope CurrentUser
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# --- Connect to Microsoft Graph --------------------------------------------

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Set-MgGraphOption -DisableLoginByWAM (-not $UseWAM)

try {
    Connect-MgGraph -Scopes "DeviceManagementServiceConfig.ReadWrite.All" -NoWelcome -TenantId $TENANT_ID -ClientId $CLIENT_ID
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

if (-not (Get-MgContext)) {
    Write-Error "Microsoft Graph context was not established. Aborting."
    exit 1
}

# --- Collect hardware information ------------------------------------------

Write-Host "Collecting hardware information from this device..." -ForegroundColor Cyan

try {
    $bios   = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
    $serial = $bios.SerialNumber.Trim()
}
catch {
    Write-Error "Failed to retrieve serial number: $_"
    exit 1
}

try {
    $devDetail = Get-CimInstance -Namespace root/cimv2/mdm/dmmap `
                    -ClassName MDM_DevDetail_Ext01 `
                    -Filter "InstanceID='Ext' AND ParentID='./DevDetail'" `
                    -ErrorAction Stop
    $hardwareHash = $devDetail.DeviceHardwareData
}
catch {
    Write-Error "Failed to retrieve hardware hash. Ensure the script is running as Administrator: $_"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($hardwareHash)) {
    Write-Error "Hardware hash is empty. Cannot register device."
    exit 1
}

# Auto-generate computer name from serial if not manually provided
if (-not $AssignedComputerName) {
    $prefix          = "DPCLAS-"
    $maxSerial       = 15 - $prefix.Length   # 8 characters left for the serial
    $truncatedSerial = $serial.Substring(0, [Math]::Min($serial.Length, $maxSerial))
    $AssignedComputerName = "$prefix$truncatedSerial"
    Write-Host "  Auto-generated computer name from serial number." -ForegroundColor Cyan
}

Write-Host "  Serial number : $serial" -ForegroundColor Gray
if ($GroupTag) { Write-Host "  Group tag     : $GroupTag" -ForegroundColor Gray }
Write-Host "  Computer name : $AssignedComputerName" -ForegroundColor Gray

# --- Import device to Autopilot --------------------------------------------

$importUri = "https://graph.microsoft.com/beta/deviceManagement/importedWindowsAutopilotDeviceIdentities"

$importBody = @{
    serialNumber       = $serial
    hardwareIdentifier = $hardwareHash
    groupTag           = $GroupTag
    "@odata.type"      = "#microsoft.graph.importedWindowsAutopilotDeviceIdentity"
} | ConvertTo-Json

Write-Host "Uploading device to Autopilot..." -ForegroundColor Cyan

try {
    $importResult = Invoke-MgGraphRequest -Method POST -Uri $importUri `
                        -Body $importBody -ContentType "application/json" -ErrorAction Stop
}
catch {
    Write-Error "Failed to submit device import: $_"
    exit 1
}

$importedId = $importResult.id
if (-not $importedId) {
    Write-Error "Import request succeeded but returned no record ID. Cannot poll for status."
    exit 1
}
Write-Host "  Import record ID: $importedId" -ForegroundColor Gray

# --- Poll for import completion --------------------------------------------

Write-Host "Waiting for import to complete..." -ForegroundColor Cyan

$statusUri    = "$importUri/$importedId"
$startTime    = [datetime]::UtcNow
$timeout      = $startTime.AddMinutes(10)
$importStatus = "unknown"
$extraMessage = ">"

while ([datetime]::UtcNow -lt $timeout) {
    Start-Sleep -Seconds 1
    $elapsed = [int]([datetime]::UtcNow - $startTime).TotalSeconds
    Write-Host "`r  Elapsed: ${elapsed}s ${extraMessage}" -NoNewline

    # Poll the API every 5 seconds because the default 15 is awful
    if ($elapsed % 5 -eq 0) {
        $extraMessage = ("-" * $extraMessage.Length) + ">"
        try {
            $statusResult = Invoke-MgGraphRequest -Method GET -Uri $statusUri -ErrorAction Stop
        }
        catch {
            Write-Warning "`nStatus check failed: $_"
            continue
        }

        $importStatus = $statusResult.state.deviceImportStatus
        if ($importStatus -ne "unknown") { break }
    }
}

Write-Host ""

# --- Handle result ---------------------------------------------------------

switch ($importStatus) {
    "complete" {
        Write-Host "Device successfully registered with Autopilot." -ForegroundColor Green

        if ($AssignedComputerName) {
            Write-Host "Setting assigned computer name to '$AssignedComputerName'..." -ForegroundColor Cyan

            $autopilotUri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities"
            $filterUri    = "${autopilotUri}?`$filter=contains(serialNumber,'$serial')"

            $device      = $null
            $nameTimeout = [datetime]::UtcNow.AddMinutes(5)

            while (-not $device -and [datetime]::UtcNow -lt $nameTimeout) {
                try {
                    $device = (Invoke-MgGraphRequest -Method GET -Uri $filterUri -ErrorAction Stop).value | Select-Object -First 1
                }
                catch {
                    Write-Warning "Failed to query Autopilot device: $_"
                }

                if (-not $device) {
                    $remaining = [int]($nameTimeout - [datetime]::UtcNow).TotalSeconds
                    Write-Host "  Waiting for device to appear in Autopilot... ($remaining`s remaining)" -ForegroundColor Gray
                    Start-Sleep -Seconds 5
                }
            }

            if ($device) {
                $updateUri  = "$autopilotUri/$($device.id)/updateDeviceProperties"
                $updateBody = @{ displayName = $AssignedComputerName } | ConvertTo-Json

                try {
                    Invoke-MgGraphRequest -Method POST -Uri $updateUri -Body $updateBody -ContentType "application/json" -ErrorAction Stop
                    Write-Host "  Computer name set successfully." -ForegroundColor Green
                }
                catch {
                    Write-Warning "Failed to set computer name: $_`nYou may need to set it manually in Intune."
                }
            }
            else {
                Write-Warning "Timed out waiting for device '$serial' to appear in Autopilot. You may need to set the name manually in Intune."
            }
        }

        if ($AutoRemove -or $Reboot -or $Shutdown) {
            Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
        }
        if ($Reboot) {
            if (-not $RebootDelay) { $RebootDelay = 10 }
            Write-Host "Rebooting in $RebootDelay seconds..."
            Start-Sleep $RebootDelay
            Restart-Computer
        }
        if ($Shutdown) {
            if (-not $ShutdownDelay) { $ShutdownDelay = 10 }
            Write-Host "Shutting down in $ShutdownDelay seconds..."
            Start-Sleep $ShutdownDelay
            Stop-Computer
        }
    }
    "error" {
        $errorCode = $statusResult.state.deviceErrorCode
        $errorName = $statusResult.state.deviceErrorName
        Write-Error "Import failed. Error $errorCode : $errorName"
        exit 1
    }
    "completedWithErrors" {
        Write-Warning "Import completed with errors: $($statusResult.state.deviceErrorName)"
    }
    default {
        Write-Warning "Import timed out or returned unexpected status: $importStatus"
    }
}
