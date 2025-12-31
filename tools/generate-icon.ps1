Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

$out = Join-Path $PSScriptRoot '..\assets\icon.png'

$w = 40
$h = 40

$bmp = [System.Drawing.Bitmap]::new($w, $h, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($bmp)

$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
$g.Clear([System.Drawing.Color]::Transparent)

$stroke = [System.Drawing.Color]::FromArgb(255, 30, 30, 30)
$pen = [System.Drawing.Pen]::new($stroke, 2)
$thinPen = [System.Drawing.Pen]::new($stroke, 1)

# Matrix/table grid
$rect = [System.Drawing.Rectangle]::new(5, 8, 26, 24)
$g.DrawRectangle($pen, $rect)
$g.DrawLine($thinPen, 14, 8, 14, 32)
$g.DrawLine($thinPen, 23, 8, 23, 32)
$g.DrawLine($thinPen, 5, 16, 31, 16)
$g.DrawLine($thinPen, 5, 24, 31, 24)

# Currency badge
$circleRect = [System.Drawing.Rectangle]::new(27, 6, 12, 12)
$circleRectF = [System.Drawing.RectangleF]::new(27, 6, 12, 12)
$g.DrawEllipse($pen, $circleRect)

$font = [System.Drawing.Font]::new('Segoe UI', 8, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
$brush = [System.Drawing.SolidBrush]::new($stroke)
$sf = [System.Drawing.StringFormat]::new()
$sf.Alignment = [System.Drawing.StringAlignment]::Center
$sf.LineAlignment = [System.Drawing.StringAlignment]::Center
$g.DrawString('$', $font, $brush, $circleRectF, $sf)

# Trend arrow
$arrowPen = [System.Drawing.Pen]::new($stroke, 2)
$g.DrawLine($arrowPen, 28, 28, 36, 20)
$g.DrawLine($arrowPen, 36, 20, 33, 20)
$g.DrawLine($arrowPen, 36, 20, 36, 23)

$g.Dispose()
$bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

Write-Output "Icon updated: $out"
