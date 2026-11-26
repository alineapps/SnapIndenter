# テスト → Release ビルド → dist に配布用 zip を作成
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

dotnet test tests\VbaFormatter.Tests
if ($LASTEXITCODE -ne 0) { throw 'テストに失敗しました' }

dotnet build src\VbaFormatter.AddIn -c Release
if ($LASTEXITCODE -ne 0) { throw 'ビルドに失敗しました' }

$dllPath = 'src\VbaFormatter.AddIn\bin\Release\net48\VbaFormatter.dll'
$version = [Reflection.AssemblyName]::GetAssemblyName((Resolve-Path $dllPath)).Version.ToString(3)
$stage   = 'dist\VbaFormatter'
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force $stage | Out-Null

Copy-Item $dllPath $stage
Copy-Item installer\* $stage
Copy-Item README.md $stage
Copy-Item docs $stage -Recurse

$zip = "dist\VbaFormatter-$version.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path "$stage\*" -DestinationPath $zip
Write-Host "作成しました: $zip" -ForegroundColor Green
