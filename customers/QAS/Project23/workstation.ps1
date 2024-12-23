# Kill off running CAD processes
Write-Host "Stopping running CAD processes..."
$processes = @(
    "sdsi.icems.monitor.exe",
    "Visicad.exe",
    "Visinetmap.exe",
    "rostersystem.exe",
    "roster.exe",
    "networkmanager.exe",
    "CommonFunction.exe",
    "tritech.visicad.app.wpf.exe",
    "tritech.service.agent.host.exe"
)

foreach ($process in $processes) {
    Write-Host "Killing process: $process"
    Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
    taskkill /im $process /F > $null 2>&1
}

# Stop and delete services with retries
Write-Host "Stopping and deleting services..."
$services = @(
    "visinet service",
    "tritech agent service"
)

foreach ($service in $services) {
    Write-Host "Stopping service: $service"
    try {
        # Stop the service
        sc stop $service > $null 2>&1

        # Wait for the service to stop completely
        $retryCount = 0
        while ((Get-Service -DisplayName $service -ErrorAction SilentlyContinue).Status -eq "StopPending" -or `
               (Get-Service -DisplayName $service -ErrorAction SilentlyContinue).Status -eq "Running") {
            Start-Sleep -Seconds 2
            $retryCount++
            if ($retryCount -ge 10) {
                Write-Host "Failed to stop service $service after multiple attempts."
                break
            }
        }

        # Confirm the service is stopped
        if ((Get-Service -DisplayName $service -ErrorAction SilentlyContinue).Status -eq "Stopped") {
            Write-Host "Service $service stopped successfully."
        } else {
            Write-Host "Service $service could not be stopped. Proceeding with deletion anyway."
        }
    } catch {
        Write-Host "Failed to stop service: $service. Attempting deletion anyway."
    }

    Write-Host "Deleting service: $service"
    try {
        # Attempt to delete the service
        sc delete $service > $null 2>&1
        Write-Host "Service $service deleted successfully."
    } catch {
        Write-Host "Failed to delete service: $service. Retrying force delete..."
        Get-WmiObject -Class Win32_Service | Where-Object { $_.DisplayName -eq $service } | ForEach-Object { $_.Delete() }
        if ((Get-Service -DisplayName $service -ErrorAction SilentlyContinue).Status -eq $null) {
            Write-Host "Service $service deleted successfully via WMI."
        } else {
            Write-Host "Service $service could not be deleted. Please investigate manually."
        }
    }
}

# Function to attempt folder deletion with retry logic
function Force-DeleteFolder {
    param (
        [string]$Path
    )

    if (Test-Path $Path) {
        Write-Host "Checking contents of: $Path"
        $contents = Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue
        if ($contents) {
            Write-Host "Attempting to delete contents of: $Path"
            try {
                Remove-Item -Path "$Path\*" -Recurse -Force -ErrorAction Stop
                Write-Host "Successfully deleted contents of $Path."
            } catch {
                Write-Host "Initial deletion failed. Retrying with administrative privileges..."
                Start-Process "cmd.exe" "/c rmdir /s /q `"$Path`"" -Verb RunAs -Wait
                if (Test-Path $Path) {
                    Write-Host "Failed to delete contents of $Path after retry. Aborting..."
                    exit 1
                } else {
                    Write-Host "Successfully deleted contents of $Path after retry."
                }
            }
        } else {
            Write-Host "Path $Path is already empty."
        }
    } else {
        Write-Host "Path $Path does not exist."
    }
}

# Force delete contents of C:\TriTech
$triTechPath = "C:\TriTech"
Force-DeleteFolder -Path $triTechPath

# Force delete contents of C:\Program Files\TriTech Software Systems
$triTechSoftwarePath = "C:\Program Files\TriTech Software Systems"
Force-DeleteFolder -Path $triTechSoftwarePath

# Open the URL
Write-Host "Opening CAD App Service in browser..."
Start-Process "https://qas23devapp1.sdsi.local/CAD.AppService/#"

Write-Host "Script completed."
exit 0
