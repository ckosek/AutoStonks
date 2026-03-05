#Requires -Version 5.1

# Dynamically locate the AutoStonks folder
# Handles standard Documents, personal OneDrive, and org OneDrive (e.g. "OneDrive - Company")
$FolderPath = $null

$possiblePaths = @(
    "$env:USERPROFILE\OneDrive\Documents\AutoStonks",
    "$env:USERPROFILE\Documents\AutoStonks"
)

# Also search for any "OneDrive - *" variants (work/school accounts)
Get-ChildItem -Path $env:USERPROFILE -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "OneDrive*" } |
    ForEach-Object {
        $possiblePaths += "$($_.FullName)\Documents\AutoStonks"
    }

foreach ($path in $possiblePaths) {
    if (Test-Path $path) {
        $FolderPath = $path
        break
    }
}

if (-not $FolderPath) {
    Write-Host "ERROR: Could not find the AutoStonks folder." -ForegroundColor Red
    Write-Host "Please make sure the AutoStonks folder is inside your Documents folder." -ForegroundColor Yellow
    pause
    exit 1
}

$ScriptPath = "$FolderPath\Set-SPXWallpaper.ps1"
$PS         = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$Arg        = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Unrestricted -File " + '"' + $ScriptPath + '"'

if (-not (Test-Path $ScriptPath)) {
    Write-Host "ERROR: Could not find Set-SPXWallpaper.ps1 in $FolderPath" -ForegroundColor Red
    pause
    exit 1
}

Write-Host "Found AutoStonks folder at: $FolderPath" -ForegroundColor Cyan

function Register-SPXTask {
    param([string]$TaskName, [string]$Time)

    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

    $trigger  = New-ScheduledTaskTrigger -Weekly `
        -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday `
        -At $Time

    $action   = New-ScheduledTaskAction -Execute $PS -Argument $Arg

    $settings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5) `
        -DontStopIfGoingOnBatteries `
        -AllowStartIfOnBatteries

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Trigger  $trigger `
        -Action   $action `
        -Settings $settings `
        -Force | Out-Null

    Write-Host "Registered: $TaskName at $Time ET Mon-Fri" -ForegroundColor Green
}

Register-SPXTask -TaskName "SPXWallpaper_MarketOpen"  -Time "9:45AM"
Register-SPXTask -TaskName "SPXWallpaper_MarketClose" -Time "4:15PM"

Write-Host ""
Write-Host "Setup complete! Tasks will run Mon-Fri at 9:45 AM and 4:15 PM." -ForegroundColor Cyan
Write-Host "To remove later, run these two lines in PowerShell:" -ForegroundColor Cyan
Write-Host "Unregister-ScheduledTask -TaskName SPXWallpaper_MarketOpen -Confirm:`$false" -ForegroundColor White
Write-Host "Unregister-ScheduledTask -TaskName SPXWallpaper_MarketClose -Confirm:`$false" -ForegroundColor White