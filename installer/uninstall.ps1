# SnapIndenter アンインストーラー
# -ShowDialog: Windows の「アプリ」一覧から実行されたとき用。結果をメッセージボックスで表示する
param([switch]$ShowDialog)
$ErrorActionPreference = 'Stop'

$clsid   = '{0E25551F-0D2E-42D6-A7AD-1F172CD3E5A7}'
$progId  = 'SnapIndenter.Connect'
$destDir = Join-Path $env:LOCALAPPDATA 'SnapIndenter'

# 表示言語: Office の表示言語を優先し、取れなければ Windows の表示言語に従う
$uiLang = $null
foreach ($v in '16.0', '15.0', '14.0') {
    $lr = Get-ItemProperty "HKCU:\Software\Microsoft\Office\$v\Common\LanguageResources" -ErrorAction SilentlyContinue
    if ($lr -and $lr.UILanguageTag) { $uiLang = $lr.UILanguageTag; break }
}
if (-not $uiLang) { $uiLang = [Globalization.CultureInfo]::CurrentUICulture.Name }
$ja = $uiLang -like 'ja*'
function Get-Msg([string]$ja_, [string]$en) { if ($ja) { $ja_ } else { $en } }

Write-Host (Get-Msg 'SnapIndenter をアンインストールします...' 'Uninstalling SnapIndenter...')

$keys = @(
    "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins\$progId",
    "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$progId",
    "HKCU:\Software\Classes\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid",
    "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsid",
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\SnapIndenter'
)
foreach ($k in $keys) {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force }
}

$warning = $null
if (Test-Path $destDir) {
    try {
        Remove-Item $destDir -Recurse -Force
    } catch {
        $warning = (Get-Msg 'フォルダーを削除できませんでした（Office が起動中の可能性があります）: ' 'Could not delete the folder (Office may still be running): ') + $destDir
        Write-Warning $warning
    }
}

Write-Host (Get-Msg 'アンインストールが完了しました。' 'Uninstall complete.') -ForegroundColor Green
$settingsFile = "$env:APPDATA\SnapIndenter\settings.ini"
$keptMsg = if (Test-Path $settingsFile) { (Get-Msg '設定ファイルは残しています: ' 'Your settings file was kept: ') + $settingsFile } else { $null }
if ($keptMsg) { Write-Host $keptMsg }

if ($ShowDialog) {
    Add-Type -AssemblyName System.Windows.Forms
    $msg = (Get-Msg 'SnapIndenter をアンインストールしました。' 'SnapIndenter has been uninstalled.')
    if ($warning) { $msg += "`n`n" + $warning }
    if ($keptMsg) { $msg += "`n`n" + $keptMsg }
    $icon = if ($warning) { 'Warning' } else { 'Information' }
    # 「設定」アプリなどの裏に隠れないよう、最前面のフォームを親にして表示する
    $owner = New-Object System.Windows.Forms.Form -Property @{ TopMost = $true; ShowInTaskbar = $false }
    [void][System.Windows.Forms.MessageBox]::Show($owner, $msg, 'SnapIndenter', 'OK', $icon)
    $owner.Dispose()
}
