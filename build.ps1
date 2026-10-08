[CmdletBinding()]
param(
    [string]$OutputDirectory,
    [string]$Python = (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python38-32\python.exe'),
    [string]$SigningCertificateThumbprint = $env:MATRIX_SIGN_CERT_THUMBPRINT,
    [string]$TimestampUrl = 'http://timestamp.digicert.com',
    # 빌드가 끝난 설치 파일을 함께 복사해 둘 폴더 (예: 내려받기를 제공하는 웹 서버의 폴더). 비워 두면 복사하지 않는다.
    [string]$PublishDirectory
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { $OutputDirectory = Join-Path $PSScriptRoot 'dist' }
if (-not (Test-Path -LiteralPath $Python)) {
    throw "Python 3.8 32비트가 필요합니다 (Windows 7과 32비트 PC 지원용). 찾은 경로 없음: $Python"
}
$bits = & $Python -c "import struct,sys; print(sys.version_info[0], sys.version_info[1], struct.calcsize('P')*8)"
if ($bits -ne '3 8 32') { throw "Python 3.8 32비트로 빌드해야 합니다. 현재: $bits" }
& $Python -c "import PyInstaller" 2>$null
if ($LASTEXITCODE -ne 0) { throw "PyInstaller가 없습니다. 먼저 `"$Python`" -m pip install -r requirements-build.txt 를 실행하세요." }

$vendor = Join-Path $PSScriptRoot 'vendor\ghostscript'
foreach ($name in 'gswin32c.exe', 'gsdll32.dll', 'vcredist_x86.exe') {
    if (-not (Test-Path -LiteralPath (Join-Path $vendor $name))) { throw "vendor\ghostscript\$name 이 없습니다. README의 준비물을 확인하세요." }
}

function Sign-Binary([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($SigningCertificateThumbprint)) { return }
    $signTool = (Get-Command signtool.exe -ErrorAction SilentlyContinue).Source
    if (-not $signTool) { throw 'signtool.exe를 찾을 수 없습니다. Windows SDK Signing Tools를 설치하세요.' }
    & $signTool sign /sha1 $SigningCertificateThumbprint /fd SHA256 /tr $TimestampUrl /td SHA256 $Path
    if ($LASTEXITCODE -ne 0) { throw "코드 서명에 실패했습니다: $Path" }
}

$buildRoot = Join-Path $env:TEMP 'matrix-pdf-driver-build'
if (Test-Path -LiteralPath $buildRoot) { Remove-Item -LiteralPath $buildRoot -Recurse -Force }
$payload = Join-Path $buildRoot 'payload'
New-Item -ItemType Directory -Path $payload -Force | Out-Null
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$OutputDirectory = (Resolve-Path $OutputDirectory).Path

& $Python -m PyInstaller `
    --noconfirm --clean --onefile --noconsole `
    --name MatrixPdfDriver `
    --version-file (Join-Path $PSScriptRoot 'version_info.txt') `
    --icon (Join-Path $PSScriptRoot 'assets\hihon.ico') `
    --add-data ((Join-Path $PSScriptRoot 'assets') + ';assets') `
    --distpath $payload `
    --workpath (Join-Path $buildRoot 'work') `
    --specpath (Join-Path $buildRoot 'spec') `
    (Join-Path $PSScriptRoot 'matrix_pdf_driver.py')
if ($LASTEXITCODE -ne 0) { throw '실행 파일 빌드에 실패했습니다.' }
Sign-Binary (Join-Path $payload 'MatrixPdfDriver.exe')

Copy-Item -LiteralPath (Join-Path $vendor 'gswin32c.exe'), (Join-Path $vendor 'gsdll32.dll'), (Join-Path $vendor 'vcredist_x86.exe') -Destination $payload
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'setup.cmd'), (Join-Path $PSScriptRoot 'LICENSE.txt'), (Join-Path $PSScriptRoot 'THIRD-PARTY.txt') -Destination $payload

$files = Get-ChildItem -LiteralPath $payload -File | Select-Object -ExpandProperty Name
$outputExe = Join-Path $OutputDirectory 'Matrix-PDF-Driver-Setup.exe'
$fileStrings = ($files | ForEach-Object -Begin { $i = 0 } -Process { "FILE$i=`"$_`""; $i++ }) -join "`r`n"
$fileRefs = (0..($files.Count - 1) | ForEach-Object { "%FILE$_%=" }) -join "`r`n"
$sed = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=%InstallPrompt%
DisplayLicense=%DisplayLicense%
FinishMessage=%FinishMessage%
TargetName=%TargetName%
FriendlyName=%FriendlyName%
AppLaunched=%AppLaunched%
PostInstallCmd=%PostInstallCmd%
AdminQuietInstCmd=%AdminQuietInstCmd%
UserQuietInstCmd=%UserQuietInstCmd%
SourceFiles=SourceFiles
[Strings]
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName="$outputExe"
FriendlyName="Matrix PDF-Driver Setup"
AppLaunched="cmd.exe /d /c .\setup.cmd"
PostInstallCmd="<None>"
AdminQuietInstCmd="cmd.exe /d /c .\setup.cmd"
UserQuietInstCmd="cmd.exe /d /c .\setup.cmd"
$fileStrings
[SourceFiles]
SourceFiles0=$payload\
[SourceFiles0]
$fileRefs
"@
$sedPath = Join-Path $buildRoot 'installer.sed'
[IO.File]::WriteAllText($sedPath, $sed, [Text.Encoding]::ASCII)
if (Test-Path -LiteralPath $outputExe) { Remove-Item -LiteralPath $outputExe -Force }

# 64비트 iexpress가 만든 설치 파일은 32비트 Windows에서 실행되지 않아 32비트 쪽을 쓴다.
$iexpress = Join-Path $env:SystemRoot 'SysWOW64\iexpress.exe'
if (-not (Test-Path -LiteralPath $iexpress)) { $iexpress = Join-Path $env:SystemRoot 'System32\iexpress.exe' }
& $iexpress /N /Q $sedPath

# iexpress는 묶기가 끝나기 전에 돌아오는 경우가 있어 파일 크기가 멈출 때까지 기다린다.
$deadline = [DateTime]::UtcNow.AddMinutes(3); $lastLength = -1L; $stable = 0
while ([DateTime]::UtcNow -lt $deadline -and $stable -lt 3) {
    Start-Sleep -Seconds 1
    if (Test-Path -LiteralPath $outputExe) {
        $length = (Get-Item -LiteralPath $outputExe).Length
        if ($length -gt 0 -and $length -eq $lastLength) { $stable++ } else { $stable = 0; $lastLength = $length }
    }
}
if ($stable -lt 3) { throw '설치 파일 생성 완료를 확인하지 못했습니다.' }
Sign-Binary $outputExe

Write-Host "설치 파일: $outputExe ($([Math]::Round((Get-Item $outputExe).Length / 1MB, 1)) MB)"
Write-Host "SHA-256: $((Get-FileHash -Algorithm SHA256 -LiteralPath $outputExe).Hash)"
if (-not [string]::IsNullOrWhiteSpace($PublishDirectory)) {
    if (-not (Test-Path -LiteralPath $PublishDirectory -PathType Container)) { throw "복사할 폴더가 없습니다: $PublishDirectory" }
    Copy-Item -LiteralPath $outputExe -Destination $PublishDirectory -Force
    Write-Host "복사해 둔 곳: $(Join-Path $PublishDirectory (Split-Path $outputExe -Leaf))"
}
if ([string]::IsNullOrWhiteSpace($SigningCertificateThumbprint)) {
    Write-Warning '코드 서명 인증서가 없어 서명되지 않은 설치 파일입니다. Windows가 실행 전에 경고를 띄울 수 있습니다.'
}
