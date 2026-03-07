#Requires -Version 5.1

# Dynamically locate the AutoStonks-main folder
# Handles standard Documents, personal OneDrive, org OneDrive with spaces/apostrophes
$FolderPath = $null

$possiblePaths = @(
    "$env:USERPROFILE\OneDrive\Documents\AutoStonks-main",
    "$env:USERPROFILE\Documents\AutoStonks-main"
)

# Search for any "OneDrive*" variants (work/school accounts with spaces or special characters)
Get-ChildItem -Path $env:USERPROFILE -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -like "OneDrive*" } |
    ForEach-Object {
        $possiblePaths += "$($_.FullName)\Documents\AutoStonks-main"
    }

foreach ($path in $possiblePaths) {
    if (Test-Path -LiteralPath $path) {
        $FolderPath = $path
        break
    }
}

if (-not $FolderPath) {
    Write-Host "ERROR: Could not find the AutoStonks-main folder." -ForegroundColor Red
    Write-Host "Please make sure the AutoStonks-main folder is inside your Documents folder." -ForegroundColor Yellow
    pause
    exit 1
}

$ScriptPath = "$FolderPath\Set-ASWallpaper.ps1"
$PS         = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$Arg        = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Unrestricted -File " + '"' + $ScriptPath + '"'

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    Write-Host "ERROR: Could not find Set-ASWallpaper.ps1 in $FolderPath" -ForegroundColor Red
    pause
    exit 1
}

Write-Host "Found AutoStonks folder at: $FolderPath" -ForegroundColor Cyan

# --- Ticker Configuration ---
Write-Host ""
Write-Host "Which stock or index would you like to track?" -ForegroundColor Cyan
Write-Host "  Examples: NVDA, GME, ^DJI, ^GSPC" -ForegroundColor White
Write-Host "  NOTE: Indices like S&P 500 or Dow Jones require a caret (^) prefix." -ForegroundColor Yellow
Write-Host "  E.g. S&P 500 = ^GSPC   |   Dow Jones = ^DJI   |   Nasdaq = ^IXIC" -ForegroundColor Yellow
Write-Host "  Press ENTER to use the default (^GSPC / S&P 500)." -ForegroundColor White
Write-Host ""
$tickerInput = Read-Host "Enter ticker"

$configFile = "$FolderPath\ticker.config"

if ($tickerInput.Trim() -eq "") {
    # No input — use default, remove any existing config so script uses built-in default
    if (Test-Path -LiteralPath $configFile) { Remove-Item -LiteralPath $configFile }
    Write-Host "Using default ticker: ^GSPC (S&P 500)" -ForegroundColor Green
} else {
    try {
        $tickerInput.Trim() | Set-Content -Path $configFile -Encoding UTF8
        Write-Host "Ticker set to: $($tickerInput.Trim())" -ForegroundColor Green
        Write-Host "Config saved to: $configFile" -ForegroundColor Gray
    }
    catch {
        Write-Host "ERROR saving ticker config: $_" -ForegroundColor Red
    }
}

# --- Register Scheduled Tasks ---
function Register-AS-Task {
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

Register-AS-Task -TaskName "AutoStonks_MarketOpen"  -Time "9:45AM"
Register-AS-Task -TaskName "AutoStonks_MarketClose" -Time "4:15PM"

Write-Host ""
Write-Host "Setup complete! Tasks will run Mon-Fri at 9:45 AM and 4:15 PM." -ForegroundColor Cyan
Write-Host "To change your ticker, just run InstallStonks.bat again." -ForegroundColor White
Write-Host "To remove tasks, run these two lines in PowerShell:" -ForegroundColor Cyan
Write-Host "Unregister-ScheduledTask -TaskName AutoStonks_MarketOpen -Confirm:`$false" -ForegroundColor White
Write-Host "Unregister-ScheduledTask -TaskName AutoStonks_MarketClose -Confirm:`$false" -ForegroundColor White