Add-Type -AssemblyName System.Drawing
$imgPath = 'C:\Users\hawks\.gemini\antigravity\brain\2e9875ce-6252-47d4-b61d-96a0ac7c9f04\kakeibo_icon_retry_1774623199924.png'
$outPath = 'C:\Users\hawks\Documents\Antigravity\kakeibo\RichMenu_1200x405.png'

$src = [System.Drawing.Image]::FromFile($imgPath)
$bmp = New-Object System.Drawing.Bitmap(1200, 405)
$g = [System.Drawing.Graphics]::FromImage($bmp)
# Get top-left pixel
$bgColor = ([System.Drawing.Bitmap]$src).GetPixel(0,0)
$g.Clear($bgColor)
# 1200x405 canvas, want square to be 405x405 in the center (X = (1200-405)/2 = 397)
$g.DrawImage($src, 397, 0, 405, 405)
$g.Dispose()
$bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
$src.Dispose()
$bmp.Dispose()
Write-Output "Image resized and saved to $outPath"
