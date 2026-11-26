# VBA Formatter アンインストーラー
$ErrorActionPreference = 'Stop'

$clsid   = '{607D618A-9D3E-4359-8B98-D12F7CDD63BA}'
$progId  = 'VbaFormatter.Connect'
$destDir = Join-Path $env:LOCALAPPDATA 'VbaFormatter'

Write-Host 'VBA Formatter をアンインストールします...'

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
Write-Host "設定ファイルは残しています: $env:APPDATA\VbaFormatter\settings.ini"
