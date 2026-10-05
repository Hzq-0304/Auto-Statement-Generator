param(
    [string]$Python = '',
    [string]$Iscc = '',
    [switch]$SkipFreeze
)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Push-Location $projectRoot
try {
    if (-not $Python) {
        $localPython = Join-Path $projectRoot '.venv/Scripts/python.exe'
        if (Test-Path -LiteralPath $localPython) { $Python = $localPython }
        else { $Python = (Get-Command python -ErrorAction Stop).Source }
    }
    if (-not $Iscc) {
        $candidates = @(
            (Join-Path $projectRoot 'build/installer-tools/InnoSetup/ISCC.exe'),
            "${env:ProgramFiles(x86)}/Inno Setup 7/ISCC.exe",
            "${env:ProgramFiles(x86)}/Inno Setup 6/ISCC.exe",
            "$env:ProgramFiles/Inno Setup 7/ISCC.exe",
            "$env:LOCALAPPDATA/Programs/Inno Setup 7/ISCC.exe"
        )
        $Iscc = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
        if (-not $Iscc) { throw '找不到 Inno Setup 编译器，请使用 -Iscc 指定 ISCC.exe。' }
    }
    $version = (& $Python -c 'from app import __version__; print(__version__)').Trim()
    if ($LASTEXITCODE -ne 0) { throw '读取版本失败。' }
    if (-not $SkipFreeze) {
        & $Python -m PyInstaller --noconfirm --windowed --onedir --name 自动对账工具 --add-data 'assets;assets' --add-data 'docs;docs' main.py
        if ($LASTEXITCODE -ne 0) { throw '程序打包失败。' }
    }
    $appDist = Join-Path $projectRoot 'dist/自动对账工具'
    if (-not (Test-Path -LiteralPath (Join-Path $appDist '自动对账工具.exe'))) { throw '未找到已打包程序。' }
    $releaseDir = Join-Path $projectRoot 'release'
    & $Iscc "/DAppVersion=$version" "/DAppDist=$appDist" "/DOutputFolder=$releaseDir" 'packaging/installer.iss'
    if ($LASTEXITCODE -ne 0) { throw '安装包编译失败。' }
    $installer = Join-Path $releaseDir "AutoStatementGenerator-$version-Setup-x64.exe"
    $hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
    "$hash  $([IO.Path]::GetFileName($installer))" | Set-Content -LiteralPath "$installer.sha256" -Encoding ascii
    Write-Output "安装包：$installer"
    Write-Output "SHA256：$hash"
} finally { Pop-Location }
