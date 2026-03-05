#Requires -Version 5.1

$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Definition
$WallpaperUp   = Join-Path $ScriptDir "stonks.png"
$WallpaperDown = Join-Path $ScriptDir "notstonks.png"
$LogFile       = Join-Path $ScriptDir "spx_wallpaper.log"

function Write-Log {
    param([string]$Message)
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

function Get-SPXChange {
    try {
        $url      = "https://query1.finance.yahoo.com/v8/finance/chart/%5EGSPC?interval=1d&range=1d"
        $headers  = @{ "User-Agent" = "Mozilla/5.0" }
        $response = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 15
        $meta     = $response.chart.result[0].meta

        $current   = [double]$meta.regularMarketPrice
        $prevClose = [double]$meta.chartPreviousClose

        if ($prevClose -eq 0) { Write-Log "ERROR: Previous close is zero."; return $null }

        $pct = (($current - $prevClose) / $prevClose) * 100
        Write-Log "SPX Current: $current | Prev Close: $prevClose | Change: $([math]::Round($pct,2))%"
        return $pct
    }
    catch {
        Write-Log "ERROR fetching SPX data: $_"
        return $null
    }
}

function Set-Wallpaper {
    param([string]$ImagePath)

    if (-not (Test-Path $ImagePath)) {
        Write-Log "ERROR: Image not found: $ImagePath"
        return
    }

    try {
        # Use .NET reflection to call SystemParametersInfo without compiling C#
        $type = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        Add-Type -AssemblyName System.Windows.Forms

        $code = @"
using System;
using System.Runtime.InteropServices;
public class WallpaperSetter {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
        if (-not ([System.Management.Automation.PSTypeName]'WallpaperSetter').Type) {
            Add-Type -TypeDefinition $code
        }

        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value "10"
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -Value "0"
        [WallpaperSetter]::SystemParametersInfo(20, 0, $ImagePath, 3)
        Write-Log "Wallpaper set to: $ImagePath"
    }
    catch {
        Write-Log "ERROR setting wallpaper: $_"
    }
}

# --- Main ---
Write-Log "=== SPX Wallpaper Script Started ==="
Write-Log "Script running from: $ScriptDir"

$change = Get-SPXChange

if ($null -eq $change) {
    Write-Log "Could not determine market direction. Wallpaper unchanged."
    exit 1
}

if ($change -ge 0) {
    Write-Log "Market UP -- setting stonks.png"
    Set-Wallpaper -ImagePath $WallpaperUp
} else {
    Write-Log "Market DOWN -- setting notstonks.png"
    Set-Wallpaper -ImagePath $WallpaperDown
}

Write-Log "=== Done ==="