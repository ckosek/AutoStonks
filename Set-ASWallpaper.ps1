#Requires -Version 5.1

$AS_ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Definition
$AS_WallpaperUp   = Join-Path $AS_ScriptDir "stonks.png"
$AS_WallpaperDown = Join-Path $AS_ScriptDir "notstonks.png"
$AS_LogFile       = Join-Path $AS_ScriptDir "autostonks.log"
$AS_ConfigFile    = Join-Path $AS_ScriptDir "ticker.config"
$AS_DefaultTicker = "^GSPC"

function Write-AS-Log {
    param([string]$Message)
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $line
    Add-Content -Path $AS_LogFile -Value $line
}

function Get-AS-Ticker {
    # Read ticker from config file if it exists, otherwise use default
    if (Test-Path -LiteralPath $AS_ConfigFile) {
        $ticker = (Get-Content $AS_ConfigFile -Raw).Trim()
        if ($ticker -ne "") {
            return $ticker
        }
    }
    return $AS_DefaultTicker
}

function Get-AS-Change {
    param([string]$Ticker)

    try {
        # URL encode ^ as %5E for indices like ^GSPC, ^DJI
        $encodedTicker = $Ticker -replace '\^', '%5E'
        $url      = "https://query1.finance.yahoo.com/v8/finance/chart/$encodedTicker`?interval=1d&range=1d"
        $headers  = @{ "User-Agent" = "Mozilla/5.0" }
        $response = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 15
        $meta     = $response.chart.result[0].meta

        $current   = [double]$meta.regularMarketPrice
        $prevClose = [double]$meta.chartPreviousClose

        if ($prevClose -eq 0) { Write-AS-Log "ERROR: Previous close is zero."; return $null }

        $pct = (($current - $prevClose) / $prevClose) * 100
        Write-AS-Log "$Ticker Current: $current | Prev Close: $prevClose | Change: $([math]::Round($pct,2))%"
        return $pct
    }
    catch {
        Write-AS-Log "ERROR fetching data for $Ticker`: $_"
        return $null
    }
}

function Set-AS-Wallpaper {
    param([string]$ImagePath)

    if (-not (Test-Path $ImagePath)) {
        Write-AS-Log "ERROR: Image not found: $ImagePath"
        return
    }

    try {
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
        Write-AS-Log "Wallpaper set to: $ImagePath"
    }
    catch {
        Write-AS-Log "ERROR setting wallpaper: $_"
    }
}

# --- Main ---
$AS_Ticker = Get-AS-Ticker
Write-AS-Log "=== AutoStonks Started ==="
Write-AS-Log "Script running from: $AS_ScriptDir"
Write-AS-Log "Tracking ticker: $AS_Ticker"

$change = Get-AS-Change -Ticker $AS_Ticker

if ($null -eq $change) {
    Write-AS-Log "Could not determine market direction. Wallpaper unchanged."
    exit 1
}

if ($change -ge 0) {
    Write-AS-Log "Market UP -- setting stonks.png"
    Set-AS-Wallpaper -ImagePath $AS_WallpaperUp
} else {
    Write-AS-Log "Market DOWN -- setting notstonks.png"
    Set-AS-Wallpaper -ImagePath $AS_WallpaperDown
}

Write-AS-Log "=== Done ==="