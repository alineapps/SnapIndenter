# SnapIndenter インストーラー（ユーザー単位・管理者権限不要）
param(
    [string]$Source = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'

$clsid    = '{0E25551F-0D2E-42D6-A7AD-1F172CD3E5A7}'
$progId   = 'SnapIndenter.Connect'
$class    = 'SnapIndenter.AddIn.Connect'
$destDir  = Join-Path $env:LOCALAPPDATA 'SnapIndenter'
$dll      = Join-Path $destDir 'SnapIndenter.dll'
$srcDll   = Join-Path $Source 'SnapIndenter.dll'

# 表示言語: Office の表示言語を優先し、取れなければ Windows の表示言語に従う
$uiLang = $null
foreach ($v in '16.0', '15.0', '14.0') {
    $lr = Get-ItemProperty "HKCU:\Software\Microsoft\Office\$v\Common\LanguageResources" -ErrorAction SilentlyContinue
    if ($lr -and $lr.UILanguageTag) { $uiLang = $lr.UILanguageTag; break }
}
if (-not $uiLang) { $uiLang = [Globalization.CultureInfo]::CurrentUICulture.Name }
$ja = $uiLang -like 'ja*'
function Get-Msg([string]$ja_, [string]$en) { if ($ja) { $ja_ } else { $en } }

Write-Host (Get-Msg 'SnapIndenter をインストールします...' 'Installing SnapIndenter...')

if (-not (Test-Path $srcDll)) {
    throw ((Get-Msg 'SnapIndenter.dll が見つかりません: ' 'SnapIndenter.dll not found: ') + $srcDll)
}

$running = Get-Process EXCEL, MSACCESS, WINWORD, POWERPNT, OUTLOOK -ErrorAction SilentlyContinue
if ($running) {
    $names = ($running | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
    Write-Warning (Get-Msg "Office が起動中です（$names）。インストール後に再起動してください。" "Office is running ($names). Restart it after installation.")
}

# 0. 旧名称（VBA Formatter 1.0 / Kanna for VBA 1.1）が入っていれば削除する。設定ファイルは SnapIndenter が引き継ぐので残す
$legacies = @(
    @{ Name = 'VBA Formatter'; ProgId = 'VbaFormatter.Connect'; Clsid = '{607D618A-9D3E-4359-8B98-D12F7CDD63BA}'; Dir = 'VbaFormatter' },
    @{ Name = 'Kanna for VBA'; ProgId = 'Kanna.Connect';        Clsid = '{DC6A7B4F-8888-4FA0-BFF5-23674D28B50F}'; Dir = 'Kanna' }
)
foreach ($l in $legacies) {
    $found = $false
    $legacyKeys = @(
        "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins\$($l.ProgId)",
        "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$($l.ProgId)",
        "HKCU:\Software\Classes\$($l.ProgId)",
        "HKCU:\Software\Classes\CLSID\$($l.Clsid)",
        "HKCU:\Software\Classes\WOW6432Node\CLSID\$($l.Clsid)"
    )
    foreach ($k in $legacyKeys) {
        if (Test-Path $k) { Remove-Item $k -Recurse -Force; $found = $true }
    }
    $legacyDir = Join-Path $env:LOCALAPPDATA $l.Dir
    if (Test-Path $legacyDir) {
        $found = $true
        try { Remove-Item $legacyDir -Recurse -Force } catch { Write-Warning ((Get-Msg '旧版のフォルダーを削除できませんでした: ' 'Could not delete the old version''s folder: ') + $legacyDir) }
    }
    if ($found) { Write-Host (Get-Msg "旧名称（$($l.Name)）の登録を削除しました。" "Removed the old version ($($l.Name)).") }
}

# 1. DLL をコピー
New-Item -ItemType Directory -Force -Path $destDir | Out-Null
try {
    Copy-Item $srcDll $dll -Force
} catch {
    throw ((Get-Msg 'DLL をコピーできませんでした。Excel / Access などをすべて終了してから再実行してください。' 'Could not copy the DLL. Close Excel, Access and other Office apps, then run this again.') + "`n$($_.Exception.Message)")
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

Set-Key "HKCU:\Software\Classes\$progId" 'SnapIndenter'
Set-Key "HKCU:\Software\Classes\$progId\CLSID" $clsid

# 3. VBE アドインとして登録（Addins = 32bit Office、Addins64 = 64bit Office）
foreach ($kind in 'Addins', 'Addins64') {
    $key = "HKCU:\Software\Microsoft\VBA\VBE\6.0\$kind\$progId"
    Set-Key $key $null
    Set-ItemProperty -Path $key -Name 'FriendlyName'    -Value 'SnapIndenter'
    Set-ItemProperty -Path $key -Name 'Description'     -Value 'VBA code indenter'
    Set-ItemProperty -Path $key -Name 'LoadBehavior'    -Value 3 -Type DWord
    Set-ItemProperty -Path $key -Name 'CommandLineSafe' -Value 0 -Type DWord
}

Write-Host ''
Write-Host (Get-Msg "インストールが完了しました（$dll）。" "Installation complete ($dll).") -ForegroundColor Green
Write-Host (Get-Msg 'Excel / Access を起動し、Alt+F11 で VBE を開くとメニューに「整形」が追加されています。' 'Start Excel or Access and press Alt+F11. A new "SnapIndenter" menu appears in the VBE.')
Write-Host (Get-Msg 'ショートカット: Ctrl+Shift+F（カーソル位置のプロシージャまたは選択範囲を整形）' 'Shortcut: Ctrl+Shift+F (indents the procedure at the cursor, or the selected lines)')
