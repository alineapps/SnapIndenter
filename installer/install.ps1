# VBA Formatter インストーラー（ユーザー単位・管理者権限不要）
param(
    [string]$Source = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'

$clsid    = '{607D618A-9D3E-4359-8B98-D12F7CDD63BA}'
$progId   = 'VbaFormatter.Connect'
$class    = 'VbaFormatter.AddIn.Connect'
$destDir  = Join-Path $env:LOCALAPPDATA 'VbaFormatter'
$dll      = Join-Path $destDir 'VbaFormatter.dll'
$srcDll   = Join-Path $Source 'VbaFormatter.dll'

Write-Host 'VBA Formatter をインストールします...'

if (-not (Test-Path $srcDll)) {
    throw "VbaFormatter.dll が見つかりません: $srcDll"
}

$running = Get-Process EXCEL, MSACCESS, WINWORD, POWERPNT, OUTLOOK -ErrorAction SilentlyContinue
if ($running) {
    Write-Warning ('Office が起動中です（' + (($running | Select-Object -ExpandProperty ProcessName -Unique) -join ', ') + '）。インストール後に再起動してください。')
}

# 1. DLL をコピー
New-Item -ItemType Directory -Force -Path $destDir | Out-Null
try {
    Copy-Item $srcDll $dll -Force
} catch {
    throw "DLL をコピーできませんでした。Excel / Access などをすべて終了してから再実行してください。`n$($_.Exception.Message)"
}
Unblock-File $dll

$asm      = [Reflection.AssemblyName]::GetAssemblyName($dll)
$asmName  = $asm.FullName
$version  = $asm.Version.ToString()
# パーセントエンコードすると日本語を含むパスで読み込めないため、regasm と同じ非エンコード形式にする
$codeBase = 'file:///' + $dll.Replace('\', '/')

function Set-Key([string]$path, [string]$default) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    if ($null -ne $default) { Set-Item -Path $path -Value $default }
}

function Set-InprocValues([string]$path) {
    Set-ItemProperty -Path $path -Name 'Class'          -Value $class
    Set-ItemProperty -Path $path -Name 'Assembly'       -Value $asmName
    Set-ItemProperty -Path $path -Name 'RuntimeVersion' -Value 'v4.0.30319'
    Set-ItemProperty -Path $path -Name 'CodeBase'       -Value $codeBase
}

# 2. COM クラス登録（64bit / 32bit Office の両方から見えるように両方のビューへ）
foreach ($root in 'HKCU:\Software\Classes\CLSID', 'HKCU:\Software\Classes\WOW6432Node\CLSID') {
    $key = "$root\$clsid"
    if (Test-Path $key) { Remove-Item $key -Recurse -Force }
    Set-Key $key $progId
    Set-Key "$key\ProgId" $progId
    Set-Key "$key\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" $null
    $ips = "$key\InprocServer32"
    Set-Key $ips 'mscoree.dll'
    Set-ItemProperty -Path $ips -Name 'ThreadingModel' -Value 'Both'
    Set-InprocValues $ips
    Set-Key "$ips\$version" $null
    Set-InprocValues "$ips\$version"
}

Set-Key "HKCU:\Software\Classes\$progId" 'VBA Formatter'
Set-Key "HKCU:\Software\Classes\$progId\CLSID" $clsid

# 3. VBE アドインとして登録（Addins = 32bit Office、Addins64 = 64bit Office）
foreach ($kind in 'Addins', 'Addins64') {
    $key = "HKCU:\Software\Microsoft\VBA\VBE\6.0\$kind\$progId"
    Set-Key $key $null
    Set-ItemProperty -Path $key -Name 'FriendlyName'    -Value 'VBA Formatter'
    Set-ItemProperty -Path $key -Name 'Description'     -Value 'VBA コードのインデント整形'
    Set-ItemProperty -Path $key -Name 'LoadBehavior'    -Value 3 -Type DWord
    Set-ItemProperty -Path $key -Name 'CommandLineSafe' -Value 0 -Type DWord
}

Write-Host ''
Write-Host "インストールが完了しました（$dll）。" -ForegroundColor Green
Write-Host 'Excel / Access を起動し、Alt+F11 で VBE を開くとメニューに「整形」が追加されています。'
Write-Host 'ショートカット: Ctrl+Shift+F（カーソル位置のプロシージャまたは選択範囲を整形）'
