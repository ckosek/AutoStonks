#Requires -Version 5.1

# Dynamically locate the AutoStonks-main folder whether OneDrive is present or not
$OneDriveDocs = "$env:USERPROFILE\OneDrive\Documents\AutoStonks-main"
$StandardDocs = "$env:USERPROFILE\Documents\AutoStonks-main"

if (Test-Path $OneDriveDocs) {
    $FolderPath = $OneDriveDocs
} elseif (Test-Path $StandardDocs) {
    $FolderPath = $StandardDocs
} else {
    Write-Host "ERROR: Could not find the AutoStonks-main folder in Documents or OneDrive\Documents." -ForegroundColor Red
    Write-Host "Please make sure this script is inside a folder called AutoStonks-main in your Documents." -ForegroundColor Yellow
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

Write-Host "Found AutoStonks-main folder at: $FolderPath" -ForegroundColor Cyan

function Register-SPXTask {
    param([string]$TaskName, [string]$Time)

    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

    $trigger = New-ScheduledTaskTrigger -Weekly `
        -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday `
        -At $Time

    $action = New-ScheduledTaskAction -Execute $PS -Argument $Arg

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