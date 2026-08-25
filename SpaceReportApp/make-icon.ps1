<#
.SYNOPSIS
    Generates app.ico for Space Report.

.DESCRIPTION
    The icon is a donut chart in the app's own verdict palette - green (safe),
    amber (review), red (never) - on a dark rounded tile. Three segments only:
    at 16x16 anything busier turns to mush.

    Drawn programmatically rather than hand-painted so it stays reproducible and
    every size is rendered natively instead of being scaled down from one image.
    Re-run after changing anything here:

        pwsh -File make-icon.ps1
#>

[CmdletBinding()]
param([string]$Out = (Join-Path $PSScriptRoot 'app.ico'))

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# Windows uses these; 256 is the one people actually see in Explorer.
$sizes = 16, 20, 24, 32, 40, 48, 64, 128, 256

function New-Tile {
    param([int]$S)

    $bmp = New-Object System.Drawing.Bitmap $S, $S,
           ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.Clear([System.Drawing.Color]::Transparent)

    # --- dark rounded tile -------------------------------------------------
    $pad = [math]::Max(0, [int][math]::Round($S * 0.02))
    $box = $S - ($pad * 2)
    $r   = [math]::Max(2, [int][math]::Round($S * 0.22))
    $d   = $r * 2

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($pad,             $pad,             $d, $d, 180, 90)
    $path.AddArc($pad + $box - $d, $pad,             $d, $d, 270, 90)
    $path.AddArc($pad + $box - $d, $pad + $box - $d, $d, $d,   0, 90)
    $path.AddArc($pad,             $pad + $box - $d, $d, $d,  90, 90)
    $path.CloseFigure()

    $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            (New-Object System.Drawing.Point $pad, $pad),
            (New-Object System.Drawing.Point ($pad + $box), ($pad + $box)),
            [System.Drawing.Color]::FromArgb(255, 39, 47, 66),
            [System.Drawing.Color]::FromArgb(255, 15, 17, 23))
    $g.FillPath($bg, $path)
    $bg.Dispose()

    # --- donut -------------------------------------------------------------
    # Segment sizes are illustrative, not data: a big reassuring green arc,
    # a smaller amber one, a red sliver.
    $thick = [math]::Max(2.0, $S * 0.135)
    $rad   = $S * 0.29
    $cx    = $S / 2.0
    $rect  = New-Object System.Drawing.RectangleF(
                [float]($cx - $rad), [float]($cx - $rad), [float]($rad * 2), [float]($rad * 2))

    # Gaps only where they will actually be visible.
    $gap = if ($S -ge 32) { 5.0 } elseif ($S -ge 24) { 3.0 } else { 0.0 }

    $segments = @(
        @{ Sweep = 176.0; Colour = [System.Drawing.Color]::FromArgb(255,  46, 204, 113) }  # safe
        @{ Sweep = 106.0; Colour = [System.Drawing.Color]::FromArgb(255, 224, 169,  46) }  # review
        @{ Sweep =  78.0; Colour = [System.Drawing.Color]::FromArgb(255, 224,  82,  82) }  # never
    )

    $angle = -90.0
    foreach ($seg in $segments) {
        $pen = New-Object System.Drawing.Pen($seg.Colour, [float]$thick)
        if ($S -ge 32) {
            $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
            $pen.EndCap   = [System.Drawing.Drawing2D.LineCap]::Round
        }
        $sweep = $seg.Sweep - $gap
        $g.DrawArc($pen, $rect, [float]($angle + $gap / 2.0), [float]$sweep)
        $pen.Dispose()
        $angle += $seg.Sweep
    }

    $g.Dispose()
    $path.Dispose()
    return $bmp
}

# --- assemble a multi-resolution .ico ---------------------------------------
# ICO layout: 6-byte header, then one 16-byte directory entry per image, then
# the image payloads. Each payload here is a PNG, which Windows 10/11 accept at
# every size and which keeps the file small.
$pngs = @()
foreach ($s in $sizes) {
    $bmp = New-Tile -S $s
    $ms  = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $pngs += ,@{ Size = $s; Bytes = $ms.ToArray() }
    $ms.Dispose(); $bmp.Dispose()
}

$fs = [System.IO.File]::Create($Out)
$bw = New-Object System.IO.BinaryWriter($fs)

$bw.Write([uint16]0)             # reserved
$bw.Write([uint16]1)             # 1 = icon
$bw.Write([uint16]$pngs.Count)

$offset = 6 + (16 * $pngs.Count)
foreach ($p in $pngs) {
    $dim = if ($p.Size -ge 256) { 0 } else { $p.Size }   # 0 means 256
    $bw.Write([byte]$dim)                                # width
    $bw.Write([byte]$dim)                                # height
    $bw.Write([byte]0)                                   # palette entries
    $bw.Write([byte]0)                                   # reserved
    $bw.Write([uint16]1)                                 # colour planes
    $bw.Write([uint16]32)                                # bits per pixel
    $bw.Write([uint32]$p.Bytes.Length)
    $bw.Write([uint32]$offset)
    $offset += $p.Bytes.Length
}
foreach ($p in $pngs) { $bw.Write($p.Bytes) }

$bw.Flush(); $bw.Dispose(); $fs.Dispose()

Write-Host ("Wrote {0} ({1:N0} bytes, {2} sizes: {3})" -f `
    $Out, (Get-Item $Out).Length, $pngs.Count, ($sizes -join ', ')) -ForegroundColor Green
