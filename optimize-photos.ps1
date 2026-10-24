# Make web-sized copies of the photos listed in photos.json.
#
# Usage (from this folder):  powershell -ExecutionPolicy Bypass -File optimize-photos.ps1
#
# Reads originals from Images\ (never published) and writes small JPEGs to
# assets\photos\. Copies are turned upright, resized, and saved without the
# original's metadata (location, camera, date).

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$root = $PSScriptRoot
$config = Get-Content -Raw -Encoding UTF8 (Join-Path $root 'photos.json') | ConvertFrom-Json
$srcDir = Join-Path $root 'Images'
$outDir = Join-Path $root 'assets\photos'
New-Item -ItemType Directory -Force $outDir | Out-Null

# Must match photo_slug() in build.py.
function Get-Slug([string]$name) {
    (($name -replace '\.[^.]+$', '').ToLower() -replace '[^a-z0-9]+', '-').Trim('-') + '.jpg'
}

$jobs = @{}
$jobs[$config.portrait.src] = 900
foreach ($gallery in $config.galleries.PSObject.Properties) {
    foreach ($item in $gallery.Value) { if (-not $jobs.ContainsKey($item.src)) { $jobs[$item.src] = 1400 } }
}

$codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$params = New-Object System.Drawing.Imaging.EncoderParameters 1
$params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), 82L

foreach ($src in $jobs.Keys) {
    $path = Join-Path $srcDir $src
    if (-not (Test-Path -LiteralPath $path)) { Write-Warning "Not found in Images: $src"; continue }
    $img = [System.Drawing.Image]::FromFile($path)
    if ($img.PropertyIdList -contains 0x0112) {
        switch ($img.GetPropertyItem(0x0112).Value[0]) {
            3 { $img.RotateFlip('Rotate180FlipNone') }
            6 { $img.RotateFlip('Rotate90FlipNone') }
            8 { $img.RotateFlip('Rotate270FlipNone') }
        }
    }
    $max = $jobs[$src]
    $scale = [Math]::Min(1.0, $max / [Math]::Max($img.Width, $img.Height))
    $w = [int]($img.Width * $scale); $h = [int]($img.Height * $scale)
    # Drawing into a fresh bitmap leaves all metadata behind.
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = 'HighQualityBicubic'
    $g.DrawImage($img, 0, 0, $w, $h)
    $out = Join-Path $outDir (Get-Slug $src)
    $bmp.Save($out, $codec, $params)
    $g.Dispose(); $bmp.Dispose(); $img.Dispose()
    "{0,-60} -> {1} ({2}x{3}, {4} KB)" -f $src, (Split-Path $out -Leaf), $w, $h, [int]((Get-Item $out).Length / 1KB)
}
