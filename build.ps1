# Builds WorkTimer.exe with the C# compiler that ships with Windows (.NET Framework 4.x). No installs needed.
param(
    [string]$Out = "$PSScriptRoot\WorkTimer.exe",
    [string]$Version = "1.0.0.0"          # CI passes 1.0.<run number>.0 - the app compares this with the latest GitHub release
)
$ErrorActionPreference = "Stop"
if ($Version -notmatch '^\d+\.\d+\.\d+\.\d+$') { throw "Version must look like 1.2.3.0" }
Add-Type -AssemblyName System.Drawing

# --- app icon (multi-size, PNG-compressed .ico) ---
function New-IconPng([int]$size) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $s = $size / 32.0
    $g.FillEllipse((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(27, 31, 36))), 0.5*$s, 0.5*$s, 31*$s, 31*$s)
    $track = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(58, 66, 76)), (3.2 * $s)
    $g.DrawEllipse($track, 5*$s, 5*$s, 22*$s, 22*$s)
    $arc = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(46, 196, 120)), (3.2 * $s)
    $arc.StartCap = 'Round'; $arc.EndCap = 'Round'
    $g.DrawArc($arc, 5*$s, 5*$s, 22*$s, 22*$s, -90, 250)    # progress ring
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), (2.2 * $s)
    $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
    $g.DrawLine($pen, 16*$s, 16*$s, 16*$s, 10*$s)    # minute hand
    $g.DrawLine($pen, 16*$s, 16*$s, 20*$s, 18*$s)    # hour hand
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    return ,$ms.ToArray()
}
$sizes = 16, 32, 48, 256
$pngs = @(); foreach ($sz in $sizes) { $pngs += ,(New-IconPng $sz) }
$icoPath = Join-Path $PSScriptRoot "app.ico"
$fs = [System.IO.File]::Create($icoPath); $bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $sz = $sizes[$i]; $dim = if ($sz -ge 256) { 0 } else { $sz }
    $bw.Write([byte]$dim); $bw.Write([byte]$dim); $bw.Write([byte]0); $bw.Write([byte]0)
    $bw.Write([uint16]1); $bw.Write([uint16]32)
    $bw.Write([uint32]$pngs[$i].Length); $bw.Write([uint32]$offset)
    $offset += $pngs[$i].Length
}
foreach ($p in $pngs) { $bw.Write($p) }
$bw.Close(); $fs.Close()

# --- compile ---
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
$objDir = Join-Path $PSScriptRoot "obj"
New-Item -ItemType Directory -Force $objDir | Out-Null
$verFile = Join-Path $objDir "AssemblyVersion.g.cs"
@"
using System.Reflection;
[assembly: AssemblyVersion("$Version")]
[assembly: AssemblyFileVersion("$Version")]
"@ | Set-Content $verFile -Encoding UTF8
$tmpOut = Join-Path $objDir "WorkTimer.build.exe"
& $csc /nologo /target:winexe /optimize+ /platform:anycpu /warn:4 /codepage:65001 `
    /out:$tmpOut /win32icon:$icoPath /win32manifest:"$PSScriptRoot\app.manifest" `
    /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.IO.Compression.dll `
    "$PSScriptRoot\WorkTimer.cs" $verFile
if ($LASTEXITCODE -ne 0) { throw "Compile failed" }

# stop a running copy so the file can be replaced
Get-Process WorkTimer -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $Out } | Stop-Process -Force
Start-Sleep -Milliseconds 400
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
Move-Item -Force $tmpOut $Out
Write-Host "Built $Out"
