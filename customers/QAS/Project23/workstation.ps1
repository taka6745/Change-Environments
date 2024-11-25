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

# Stop and delete services
Write-Host "Stopping and deleting services..."
$services = @(
    "visinet service",
    "tritech agent service"
)

foreach ($service in $services) {
    Write-Host "Stopping service: $service"
    sc stop $service > $null 2>&1
    Write-Host "Deleting service: $service"
    sc delete $service > $null 2>&1
}

# Force delete contents of C:\TriTech
$triTechPath = "C:\TriTech"
if (Test-Path $triTechPath) {
    Write-Host "Deleting contents of: $triTechPath"
    try {
        Remove-Item -Path "$triTechPath\*" -Recurse -Force -ErrorAction Stop
        Write-Host "Successfully deleted contents of $triTechPath."
    } catch {
        Write-Host "Failed to delete contents of $triTechPath. Aborting..."
        exit 1
    }
} else {
    Write-Host "Path $triTechPath does not exist."
}

# Open the URL
Write-Host "Opening CAD App Service in browser..."
Start-Process "https://qas23devapp1.sdsi.local/CAD.AppService/#"

Write-Host "Script completed."
exit 0
