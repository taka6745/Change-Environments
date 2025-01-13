<#
.SYNOPSIS
    23.x Service Setup Script (Graceful + Forceful Cleanup)

.DESCRIPTION
    Cleans up previous CAD application processes/services, forcibly if needed.
    Disables/stops/deletes TriTech services, removes TriTech directories.
    Unmounts/remounts the Q: drive, and launches the Service application.

.PARAMETER CustomerName
    Name of the customer.

.PARAMETER Environment
    Name of the environment (e.g., Production, Test, etc.).

.PARAMETER QDrive
    UNC path (slash or backslash syntax) for the Q drive location.

.EXAMPLE
    .\23_Service.ps1 -CustomerName "CityName" -Environment "Test" -QDrive "//myserver/share"

.NOTES
    Must be run as Administrator.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Name of the customer.")]
    [ValidateNotNullOrEmpty()]
    [string]$CustomerName,

    [Parameter(Mandatory = $true, HelpMessage = "Name of the environment.")]
    [ValidateNotNullOrEmpty()]
    [string]$Environment,

    [Parameter(Mandatory = $true, HelpMessage = "UNC path for Q drive (can be forward slashes).")]
    [ValidateNotNullOrEmpty()]
    [string]$QDrive
)

Write-Host "====================================================="
Write-Host "         23.x Service Setup Script (Enhanced)"
Write-Host "====================================================="
Write-Host ""

# --- Check for administrative privileges ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator. Exiting."
    Exit 1
}

Write-Host "Parameters received:"
Write-Host "  CustomerName : $CustomerName"
Write-Host "  Environment  : $Environment"
Write-Host "  QDrive       : $QDrive"
Write-Host ""
Write-Host "Running 23.x Service Setup for '$CustomerName' - '$Environment'"
Write-Host ""

# --------------------------------------------------
# FUNCTION: StopProcessGracefullyOrKill
# --------------------------------------------------
function StopProcessGracefullyOrKill {
    param (
        [Parameter(Mandatory)]
        [string]$ProcessName,
        [int]$TimeoutInSeconds = 5
    )

    $processList = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($processList) {
        Write-Host "Attempting to close process '$ProcessName' gracefully..."
        foreach ($p in $processList) {
            try {
                $p.CloseMainWindow() | Out-Null
            } catch {
                # If there's no GUI main window, ignore
            }
        }

        # Wait briefly for graceful exit
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while ($sw.Elapsed.TotalSeconds -lt $TimeoutInSeconds) {
            $processList = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
            if (-not $processList) { break }
            Start-Sleep -Seconds 1
        }
        $sw.Stop()

        # If still running, force kill
        $processList = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($processList) {
            Write-Warning "Process '$ProcessName' is still running. Forcing kill..."
            foreach ($p in $processList) {
                try {
                    Stop-Process -Id $p.Id -Force
                } catch {
                    Write-Warning "Unable to forcibly kill PID $($p.Id): $($_.Exception.Message)"
                }
            }
        }
    }
}

# --------------------------------------------------
# FUNCTION: StopDisableDeleteService
# --------------------------------------------------
function StopDisableDeleteService {
    param (
        [Parameter(Mandatory)]
        [string]$ServiceName,  # Accepts actual service Name or DisplayName
        [int]$TimeoutInSeconds = 15
    )

    # Identify the service by Name or DisplayName
    $serviceObj = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $serviceObj) {
        $serviceWmi = Get-WmiObject Win32_Service -Filter "DisplayName='$ServiceName'" -ErrorAction SilentlyContinue
        if ($serviceWmi) {
            $serviceObj = Get-Service -Name $serviceWmi.Name -ErrorAction SilentlyContinue
        }
    }

    if (-not $serviceObj) {
        Write-Host "Service '$ServiceName' not found; skipping stop/disable/delete."
        return
    }

    # 1) Disable the service (so it won't restart automatically)
    Write-Host "Disabling service '$($serviceObj.Name)' (DisplayName: '$ServiceName')..."
    try {
        sc.exe config "$($serviceObj.Name)" start= disabled | Out-Null
    } catch {
        Write-Warning "Could not disable service '$ServiceName': $($_.Exception.Message)"
    }

    $serviceObj.Refresh()

    # 2) If the service is Running/StartPending/StopPending, try to stop it
    if ($serviceObj.Status -in ('Running','StartPending','StopPending')) {
        Write-Host "Stopping service '$($serviceObj.Name)' (Status: $($serviceObj.Status))..."
        try {
            Stop-Service -Name $serviceObj.Name -Force -ErrorAction Stop
        } catch {
            Write-Warning "Service '$ServiceName' did not stop gracefully: $($_.Exception.Message)"
        }

        # Wait up to TimeoutInSeconds
        $sw = [Diagnostics.Stopwatch]::StartNew()
        do {
            Start-Sleep -Seconds 1
            $serviceObj.Refresh()
        } while (($serviceObj.Status -in ('Running','StartPending','StopPending')) -and 
                 ($sw.Elapsed.TotalSeconds -lt $TimeoutInSeconds))
        $sw.Stop()

        # If still running, kill the underlying .exe
        if ($serviceObj.Status -in ('Running','StartPending','StopPending')) {
            Write-Warning "Service '$ServiceName' did not stop after $TimeoutInSeconds seconds. Forcing kill..."
            $pName = (Get-WmiObject Win32_Service -Filter "Name='$($serviceObj.Name)'" -ErrorAction SilentlyContinue).PathName
            if ($pName) {
                $exe = Split-Path $pName -Leaf
                Write-Host "Killing process $exe..."
                try {
                    Stop-Process -Name $exe -Force -ErrorAction SilentlyContinue
                } catch {
                    Write-Warning "Unable to forcibly kill $($exe): $($_.Exception.Message)"
                }
            }
        }
    } else {
        Write-Host "Service '$ServiceName' not running (Status: $($serviceObj.Status))."
    }

    # 3) Delete the service from SCM
    $serviceWmiFinal = Get-WmiObject Win32_Service -Filter "Name='$($serviceObj.Name)'" -ErrorAction SilentlyContinue
    if ($serviceWmiFinal) {
        Write-Host "Deleting service '$($serviceWmiFinal.DisplayName)' from SCM..."
        try {
            $serviceWmiFinal.Delete() | Out-Null
            Write-Host "Service '$($serviceWmiFinal.DisplayName)' deleted successfully."
        } catch {
            Write-Error "Failed to delete service '$($serviceWmiFinal.DisplayName)': $($_.Exception.Message)"
        }
    } else {
        Write-Host "Service '$ServiceName' not found in WMI; possibly already removed."
    }
}

# --------------------------------------------------
# FUNCTION: ForceRemoveDirectory
# --------------------------------------------------
function ForceRemoveDirectory {
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (Test-Path $Path) {
        Write-Host "Attempting to remove directory: $Path"
        try {
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            Write-Host "Directory '$Path' removed successfully."
        }
        catch {
            Write-Warning "Initial removal of '$Path' failed: $($_.Exception.Message)"
            Write-Host "Attempting to take ownership and grant Administrators full control on '$Path'..."

            try {
                takeown /F $Path /R /D Y | Out-Null
                icacls $Path /grant Administrators:F /T | Out-Null
            } catch {
                Write-Warning "takeown/icacls failed on '$Path': $($_.Exception.Message)"
            }

            # Try removing again
            try {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
                Write-Host "Directory '$Path' removed after taking ownership."
            }
            catch {
                Write-Warning "Failed to remove '$Path' even after ownership. Attempting rename fallback..."
                $timestamp = (Get-Date -Format "yyyyMMdd_HHmmss")
                $newName = $Path + "_OLD_" + $timestamp
                try {
                    Rename-Item -LiteralPath $Path -NewName (Split-Path $newName -Leaf)
                    Write-Warning "Renamed '$Path' to '$newName'."
                    Write-Warning "Please remove '$newName' manually (e.g., after a reboot)."
                }
                catch {
                    Write-Error "Failed to rename '$Path': $($_.Exception.Message)"
                }
            }
        }
    }
    else {
        Write-Host "Directory '$Path' does not exist; skipping."
    }
}

# -------------------------- MAIN SCRIPT LOGIC --------------------------

#
# 1) Stop running CAD processes
#
Write-Host "Stopping running CAD processes..."
$cadProcesses = @(
    "sdsi.icems.monitor",
    "Visicad",
    "Visinetmap",
    "rostersystem",
    "roster",
    "networkmanager",
    "CommonFunction",
    "tritech.visicad.app.wpf",
    "tritech.service.agent.host"
)
foreach ($proc in $cadProcesses) {
    StopProcessGracefullyOrKill -ProcessName $proc -TimeoutInSeconds 5
}
Write-Host "CAD processes stopped (or killed)."
Write-Host ""

#
# 2) Stop, disable, and delete relevant services
#
Write-Host "Handling 'visinet service'..."
StopDisableDeleteService -ServiceName "visinet service" -TimeoutInSeconds 15
Write-Host ""

Write-Host "Handling 'TriTech Agent Service'..."
StopDisableDeleteService -ServiceName "TriTech Agent Service" -TimeoutInSeconds 15
Write-Host ""

#
# 3) Delete TriTech directories (with fallback if locked)
#
Write-Host "Deleting TriTech directories if they exist..."
$directories = @(
    "C:\TriTech",
    "C:\Program Files\TriTech Software Systems"
)
foreach ($dir in $directories) {
    ForceRemoveDirectory -Path $dir
    Write-Host ""
}

#
# 4) Unmount Q drive if it exists
#
Write-Host "Attempting to unmount Q drive..."
try {
    net use Q: /delete /yes 2>$null
    Write-Host "Q drive unmounted successfully (if it was mounted)."
} catch {
    Write-Warning "Could not unmount Q drive: $($_.Exception.Message)"
}
Write-Host ""

#
# 5) Format the UNC path properly and mount Q drive
#
Write-Host "Mounting new Q drive from config..."
$QDrivePath = "\\" + $QDrive.TrimStart('/').Replace('/', '\')
Write-Host "UNC Path resolved to: $QDrivePath"
Write-Host ""

try {
    Write-Host "Trying 'net use Q:'..."
    net use Q: $QDrivePath | Out-Null

    if ($LASTEXITCODE -eq 0) {
        Write-Host "Q drive mounted successfully using net use."
    } else {
        Write-Host "Failed to mount Q drive using net use. Trying New-PSDrive..."
        New-PSDrive -Name 'Q' -PSProvider FileSystem -Root $QDrivePath -Persist | Out-Null
        Write-Host "Q drive mounted successfully using New-PSDrive."
    }
} catch {
    Write-Error "Error mounting Q drive: $($_.Exception.Message)"
    exit 1
}
Write-Host ""

#
# 6) Verify Q drive is accessible
#
Write-Host "Verifying Q drive..."
if (-not (Test-Path "Q:\")) {
    Write-Error "Q drive not accessible after mounting. Aborting."
    exit 1
}
Write-Host "Q drive is accessible."
Write-Host ""

#
# 7) Launch Service application
#
$serviceLaunchPath = "Q:\Service Launch.lnk"
if (Test-Path $serviceLaunchPath) {
    Write-Host "Launching Service application..."
    try {
        Start-Process $serviceLaunchPath
        Write-Host "Service application launched successfully."
    } catch {
        Write-Error "Error launching Service application: $($_.Exception.Message)"
        exit 1
    }
} else {
    Write-Error "Service Launch.lnk not found at $serviceLaunchPath. Aborting."
    exit 1
}

Write-Host "`nScript completed successfully."
exit 0
