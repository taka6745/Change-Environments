<#
.SYNOPSIS
    23.x Workstation Setup Script

.DESCRIPTION
    Cleans up previous CAD application processes/services on a workstation.
    Removes TriTech directories (taking ownership if locked).
    Finally opens the specified App Service website in the default browser
    so the user can download and install the workstation application manually.

.PARAMETER CustomerName
    Name of the customer.

.PARAMETER Environment
    Name of the environment (e.g., Production, Test, etc.).

.PARAMETER AppService
    URL of the App Service website from which the workstation installer can be downloaded.

.EXAMPLE
    .\23_Workstation.ps1 -CustomerName "CityName" -Environment "Test" -AppService "https://myappservicewebsite.example.com"

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

    [Parameter(Mandatory = $true, HelpMessage = "URL for the App Service website.")]
    [ValidateNotNullOrEmpty()]
    [string]$AppService
)

Write-Host "====================================================="
Write-Host "         23.x Workstation Setup Script"
Write-Host "====================================================="
Write-Host ""

# --- Check for administrative privileges ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator. Exiting."
    Exit 1
}

Write-Host "Running 23.x Workstation Setup"
Write-Host "  Customer   : $CustomerName"
Write-Host "  Environment: $Environment"
Write-Host "  App Service: $AppService"
Write-Host ""

# ------------------------------ FUNCTIONS ------------------------------

<#
.SYNOPSIS
    Gracefully (then forcibly) stop a process by name.
#>
function StopProcessGracefullyOrKill {
    param (
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
                # If the process has no main window, ignore.
            }
        }

        # Wait briefly for graceful exit
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while ($sw.Elapsed.TotalSeconds -lt $TimeoutInSeconds) {
            $processList = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
            if (-not $processList) {
                break
            }
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

<#
.SYNOPSIS
    Disables, stops, and deletes a service (handling StartPending/StopPending).
#>
function StopDisableDeleteService {
    param (
        [Parameter(Mandatory)]
        [string]$ServiceName,  # Service Name or DisplayName
        [int]$TimeoutInSeconds = 15
    )

    # Identify service by Name or DisplayName
    $serviceObj = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $serviceObj) {
        # Try WMI by DisplayName
        $serviceWmi = Get-WmiObject Win32_Service -Filter "DisplayName='$ServiceName'" -ErrorAction SilentlyContinue
        if ($serviceWmi) {
            $serviceObj = Get-Service -Name $serviceWmi.Name -ErrorAction SilentlyContinue
        }
    }

    if (-not $serviceObj) {
        Write-Host "Service '$ServiceName' not found; skipping stop/disable/delete."
        return
    }

    # 1) Disable the service so it won't auto-restart
    Write-Host "Disabling service '$($serviceObj.Name)' (DisplayName: '$ServiceName')..."
    try {
        sc.exe config "$($serviceObj.Name)" start= disabled | Out-Null
    } catch {
        Write-Warning "Could not disable service '$ServiceName': $($_.Exception.Message)"
    }

    $serviceObj.Refresh()

    # 2) If service is in Running/StartPending/StopPending, try stopping it
    if ($serviceObj.Status -in ('Running','StartPending','StopPending')) {
        Write-Host "Stopping service '$($serviceObj.Name)' (Status: $($serviceObj.Status))..."
        try {
            # -Force helps handle StartPending/StopPending
            Stop-Service -Name $serviceObj.Name -Force -ErrorAction Stop
        } catch {
            Write-Warning "Service '$ServiceName' did not stop gracefully: $($_.Exception.Message)"
        }

        # 3) Wait for it to stop
        $sw = [Diagnostics.Stopwatch]::StartNew()
        do {
            Start-Sleep 1
            $serviceObj.Refresh()
        } while (($serviceObj.Status -in ('Running','StartPending','StopPending')) -and ($sw.Elapsed.TotalSeconds -lt $TimeoutInSeconds))
        $sw.Stop()

        # 4) If still running, forcibly kill the .exe
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

    # 5) Delete the service from SCM
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

<#
.SYNOPSIS
    Forcibly remove a directory, with fallback (takeown + icacls), then rename.
#>
function ForceRemoveDirectory {
    param (
        [Parameter(Mandatory = $true)]
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

            # Try once more
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

Write-Host "Handling 'visinet service'..."
StopDisableDeleteService -ServiceName "visinet service" -TimeoutInSeconds 15
Write-Host ""

Write-Host "Handling 'TriTech Agent Service'..."
StopDisableDeleteService -ServiceName "TriTech Agent Service" -TimeoutInSeconds 15
Write-Host ""

Write-Host "Deleting TriTech directories if they exist..."
$directories = @(
    "C:\TriTech",
    "C:\Program Files\TriTech Software Systems"
)
foreach ($dir in $directories) {
    ForceRemoveDirectory -Path $dir
    Write-Host ""
}

Write-Host "Workstation cleanup steps completed."
Write-Host ""

Write-Host "Launching App Service website for manual workstation install:"
Write-Host "  $AppService"
try {
    # This will open the default browser on Windows
    Start-Process $AppService
    Write-Host "App Service website opened successfully."
} catch {
    Write-Error "Error opening App Service website: $($_.Exception.Message)"
}

Write-Host "`nScript completed successfully."
exit 0
