<#
.SYNOPSIS
    Captures documentation screenshots of the Space Report window.

.DESCRIPTION
    Launches the app with --scan/--view so it reaches a known state without any
    clicking, waits for the window title to report "ready", then copies the
    window's screen rectangle to a PNG.

    Screen capture is used rather than PrintWindow because the window hosts a
    WebView2 control, whose composited content PrintWindow does not reliably
    reproduce. That means the window must be visible and unobstructed while this
    runs - do not cover it.

.PARAMETER Exe
    The SpaceReport.exe to drive. Defaults to the Release build.

.PARAMETER OutDir
    Where to write the PNGs. Defaults to .\docs

.EXAMPLE
    pwsh -File capture-screenshots.ps1
#>

[CmdletBinding()]
param(
    [string]$Exe    = (Join-Path $PSScriptRoot 'SpaceReportApp\bin\Release\net8.0-windows\SpaceReport.exe'),
    [string]$OutDir = (Join-Path $PSScriptRoot 'docs'),
    [string]$Drive  = 'C:\',
    [int]   $MinMB  = 500
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path $Exe)) { throw "Not found: $Exe" }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

Add-Type @'
using System;
using System.Runtime.InteropServices;
public class Win {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(
      IntPtr h, int attr, out RECT val, int size);

  // GetWindowRect includes the invisible drop-shadow margin, so capturing it
  // bleeds in whatever sits behind the window. The DWM extended frame bounds
  // are the actual painted edges.
  public const int EXTENDED_FRAME_BOUNDS = 9;
  public static RECT VisibleBounds(IntPtr h) {
    RECT r;
    if (DwmGetWindowAttribute(h, EXTENDED_FRAME_BOUNDS, out r, Marshal.SizeOf(typeof(RECT))) == 0
        && r.R > r.L && r.B > r.T) return r;
    GetWindowRect(h, out r);
    return r;
  }
}
'@

function Save-Window {
    param([System.Diagnostics.Process]$Proc, [string]$Path)

    $h = $Proc.MainWindowHandle
    if ([Win]::IsIconic($h)) { [Win]::ShowWindow($h, 9) | Out-Null }   # SW_RESTORE
    [Win]::SetForegroundWindow($h) | Out-Null
    Start-Sleep -Milliseconds 700     # let the window settle and repaint

    $r = [Win]::VisibleBounds($h)
    $w = $r.R - $r.L; $ht = $r.B - $r.T
    if ($w -le 0 -or $ht -le 0) { throw 'Could not read the window rectangle.' }

    $bmp = New-Object System.Drawing.Bitmap $w, $ht
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.L, $r.T, 0, 0, (New-Object System.Drawing.Size $w, $ht))
    $g.Dispose()
    $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host ("  saved {0}  ({1}x{2})" -f (Split-Path $Path -Leaf), $w, $ht) -ForegroundColor Green
}

function Wait-Ready {
    param([System.Diagnostics.Process]$Proc, [int]$TimeoutSec = 180)
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        $Proc.Refresh()
        if ($Proc.HasExited) { throw 'The app exited before finishing.' }
        if ($Proc.MainWindowTitle -like '*ready*') { return $true }
        Start-Sleep -Milliseconds 400
    }
    throw "Timed out waiting for the scan to finish."
}

function Capture-View {
    param([string]$View, [string]$File, [switch]$DuringScan)

    Write-Host "capturing $File ..." -ForegroundColor Cyan
    Get-Process SpaceReport -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 600

    $args = @('--scan', $Drive, '--min', "$MinMB")
    if ($View) { $args += @('--view', $View) }
    $p = Start-Process $Exe -ArgumentList $args -PassThru

    # wait for a window to exist at all
    for ($i = 0; $i -lt 60; $i++) {
        $p.Refresh(); if ($p.MainWindowHandle -ne 0) { break }; Start-Sleep -Milliseconds 500
    }
    if ($p.MainWindowHandle -eq 0) { throw 'No window appeared.' }

    if ($DuringScan) {
        Start-Sleep -Seconds 12          # far enough in for the bar to be meaningful
    } else {
        Wait-Ready -Proc $p | Out-Null
        Start-Sleep -Milliseconds 800    # let the final render settle
    }
    Save-Window -Proc $p -Path (Join-Path $OutDir $File)
}

Write-Host "Capturing to $OutDir" -ForegroundColor White
Write-Host "Leave the window alone while this runs." -ForegroundColor DarkYellow
Write-Host ''

Capture-View -View ''       -File '02-scanning.png' -DuringScan
Capture-View -View 'system' -File '03-system-cleanup.png'
Capture-View -View 'SAFE'   -File '01-safe-files.png'

Get-Process SpaceReport -ErrorAction SilentlyContinue | Stop-Process -Force
Write-Host ''
Write-Host 'Done. Review every image before committing - they show real paths.' -ForegroundColor Yellow
