# assets/icon.svg と同じ図形を GDI+ で各サイズに描き、SnapIndenter.ico を作る
# （SVG を変えたら、ここの座標と色も合わせてから実行する）
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$bg    = [System.Drawing.ColorTranslator]::FromHtml('#1D2A44')
$light = [System.Drawing.ColorTranslator]::FromHtml('#C9D2E3')
$amber = [System.Drawing.ColorTranslator]::FromHtml('#F29A3A')
# 線: x1, x2, y, 色（256 x 256 の座標、線幅 20、端は丸）
$lines = @(
    @(56, 168, 70, $light),
    @(96, 200, 110, $amber),
    @(96, 176, 150, $amber),
    @(56, 136, 190, $light)
)

function New-IconBitmap([int]$size) {
    $bmp = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.PixelOffsetMode = 'HighQuality'
    $g.Clear($bg)
    $s = $size / 256.0
    foreach ($l in $lines) {
        $pen = New-Object System.Drawing.Pen $l[3], ([float](20 * $s))
        $pen.StartCap = 'Round'; $pen.EndCap = 'Round'
        $g.DrawLine($pen, [float]($l[0] * $s), [float]($l[2] * $s), [float]($l[1] * $s), [float]($l[2] * $s))
        $pen.Dispose()
    }
    $g.Dispose()
    $bmp
}

# $png が真なら PNG（サイズを抑える）、偽なら 32bit DIB（.NET Framework の Icon クラスは PNG のフレームを扱えない）
function Get-IconImageBytes([int]$size, [bool]$png) {
    $bmp = New-IconBitmap $size
    try {
        if ($png) {
            $ms = New-Object System.IO.MemoryStream
            $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            return , $ms.ToArray()
        }
        $ms = New-Object System.IO.MemoryStream
        $w = New-Object System.IO.BinaryWriter $ms
        $w.Write([int]40); $w.Write([int]$size); $w.Write([int]($size * 2))
        $w.Write([int16]1); $w.Write([int16]32); $w.Write([int]0)
        $w.Write([int]0); $w.Write([int]0); $w.Write([int]0); $w.Write([int]0); $w.Write([int]0)
        for ($y = $size - 1; $y -ge 0; $y--) {
            for ($x = 0; $x -lt $size; $x++) {
                $c = $bmp.GetPixel($x, $y)
                $w.Write([byte]$c.B); $w.Write([byte]$c.G); $w.Write([byte]$c.R); $w.Write([byte]$c.A)
            }
        }
        $maskStride = [int]([Math]::Ceiling($size / 32.0) * 4)
        $w.Write((New-Object byte[] ($maskStride * $size)))
        $w.Flush()
        return , $ms.ToArray()
    } finally { $bmp.Dispose() }
}

function Write-Ico([string]$name, [int[]]$sizes, [int]$pngFrom) {
    $images = foreach ($sz in $sizes) { , (Get-IconImageBytes $sz ($sz -ge $pngFrom)) }
    $out = Join-Path $PSScriptRoot $name
    $fs = [System.IO.File]::Create($out)
    $w = New-Object System.IO.BinaryWriter $fs
    $w.Write([int16]0); $w.Write([int16]1); $w.Write([int16]$sizes.Count)
    $offset = 6 + 16 * $sizes.Count
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $sz = $sizes[$i]; $dim = if ($sz -ge 256) { 0 } else { $sz }
        $w.Write([byte]$dim); $w.Write([byte]$dim); $w.Write([byte]0); $w.Write([byte]0)
        $w.Write([int16]1); $w.Write([int16]32)
        $w.Write([int]$images[$i].Length); $w.Write([int]$offset)
        $offset += $images[$i].Length
    }
    foreach ($img in $images) { $w.Write($img) }
    $w.Dispose()
    Write-Host "作成しました: $out"
}

# DLL の Win32 アイコン（エクスプローラーや「アプリ」一覧で使う）。32 以上は PNG
Write-Ico 'SnapIndenter.ico' @(16, 20, 24, 32, 40, 48, 64, 256) 32
# 設定画面のタイトルバー用（DLL に埋め込むリソース）。すべて DIB
Write-Ico 'SnapIndenter.form.ico' @(16, 20, 24, 32) 9999
