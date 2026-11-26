# Kanna for VBA アンインストーラー
$ErrorActionPreference = 'Stop'

$clsid   = '{DC6A7B4F-8888-4FA0-BFF5-23674D28B50F}'
$progId  = 'Kanna.Connect'
$destDir = Join-Path $env:LOCALAPPDATA 'Kanna'

Write-Host 'Kanna for VBA をアンインストールします...'

$keys = @(
    "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins\$progId",
    "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$progId",
    "HKCU:\Software\Classes\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsid"
)
foreach ($k in $keys) {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force }
}

if (Test-Path $destDir) {
    try {
        Remove-Item $destDir -Recurse -Force
    } catch {
        Write-Warning "フォルダーを削除できませんでした（Office が起動中の可能性があります）: $destDir"
    }
}

Write-Host 'アンインストールが完了しました。' -ForegroundColor Green
Write-Host "設定ファイルは残しています: $env:APPDATA\Kanna\settings.ini"
