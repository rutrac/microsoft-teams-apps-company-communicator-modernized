# Regenerates color.png (192x192) and outline.png (32x32 white-on-transparent silhouette)
# from a single high-resolution source PNG. Run from anywhere:
#   .\Manifest\_build_icons.ps1 -SourcePath C:\path\to\icon-source.png
param(
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [int]$OutlineThreshold = 150
)

Add-Type -AssemblyName System.Drawing

$srcPath = (Resolve-Path $SourcePath).Path
$colorOut = Join-Path $PSScriptRoot 'color.png'
$outlineOut = Join-Path $PSScriptRoot 'outline.png'

$src = [System.Drawing.Image]::FromFile($srcPath)
Write-Host "Source: $($src.Width)x$($src.Height)"

# 192x192 color
$color = New-Object System.Drawing.Bitmap 192, 192
$g = [System.Drawing.Graphics]::FromImage($color)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g.DrawImage($src, 0, 0, 192, 192)
$g.Dispose()
$color.Save($colorOut, [System.Drawing.Imaging.ImageFormat]::Png)
$color.Dispose()
Write-Host "Wrote $colorOut"

# 32x32 outline: white silhouette on transparent
$small = New-Object System.Drawing.Bitmap 32, 32
$g2 = [System.Drawing.Graphics]::FromImage($small)
$g2.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g2.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
$g2.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g2.DrawImage($src, 0, 0, 32, 32)
$g2.Dispose()

$outline = New-Object System.Drawing.Bitmap 32, 32, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
for ($y = 0; $y -lt 32; $y++) {
    for ($x = 0; $x -lt 32; $x++) {
        $p = $small.GetPixel($x, $y)
        $bright = ($p.R + $p.G + $p.B) / 3
        if ($bright -gt $OutlineThreshold) {
            $a = [Math]::Min(255, [int](($bright - $OutlineThreshold) * 2.55))
            $outline.SetPixel($x, $y, [System.Drawing.Color]::FromArgb($a, 255, 255, 255))
        }
        else {
            $outline.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(0, 0, 0, 0))
        }
    }
}
$outline.Save($outlineOut, [System.Drawing.Imaging.ImageFormat]::Png)
$outline.Dispose()
$small.Dispose()
$src.Dispose()
Write-Host "Wrote $outlineOut"
