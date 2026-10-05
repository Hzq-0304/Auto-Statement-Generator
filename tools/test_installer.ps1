param([string]$Installer = '')
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Push-Location $projectRoot
try {
    $versionLine = Get-Content -LiteralPath 'app/__init__.py' | Select-String '__version__'
    $version = [regex]::Match($versionLine.ToString(), '\d+\.\d+\.\d+').Value
    if (-not $Installer) { $Installer = "release/AutoStatementGenerator-$version-Setup-x64.exe" }
    # 验收目录限定在项目 build 下；已安装的用户版本存在时不覆盖它。
    $testRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot ("build/installer-smoke-$version-" + [guid]::NewGuid().ToString('N').Substring(0,8))))
    $allowedRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot 'build')) + [IO.Path]::DirectorySeparatorChar
    if (-not $testRoot.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw '验收目录不在 build 内。' }
    if (Test-Path -LiteralPath $testRoot) { throw '验收目录已存在，请先检查其中内容，不能覆盖。' }
    $regPath = 'HKCU:/Software/Microsoft/Windows/CurrentVersion/Uninstall/{4093D3C0-6992-45E0-9349-4309A9E96D96}_is1'
    if (Test-Path $regPath) { throw '当前用户已安装正式版本，为避免覆盖，停止安装验收。' }
    $setup = (Resolve-Path -LiteralPath $Installer).Path
    New-Item -ItemType Directory -Force -Path 'analysis','outputs/installer-smoke' | Out-Null
    $log = Join-Path $projectRoot 'analysis/installer-smoke.log'
    $process = Start-Process -FilePath $setup -ArgumentList @('/VERYSILENT','/SUPPRESSMSGBOXES','/NORESTART','/SP-','/NOICONS','/TASKS=""',('/DIR="' + $testRoot + '"'),('/LOG="' + $log + '"')) -WindowStyle Hidden -PassThru
    if (-not $process.WaitForExit(60000)) { throw '安装超时。' }
    if ($process.ExitCode -ne 0) { throw "安装失败：$($process.ExitCode)" }
    $exe = Join-Path $testRoot '自动对账工具.exe'
    $uninstaller = Join-Path $testRoot 'unins000.exe'
    try {
        foreach ($file in (Get-ChildItem -LiteralPath 'dist/自动对账工具' -File -Recurse)) {
            $relative = $file.FullName.Substring((Resolve-Path 'dist/自动对账工具').Path.Length).TrimStart([IO.Path]::DirectorySeparatorChar)
            $installed = Join-Path $testRoot $relative
            if ((Get-FileHash -LiteralPath $file.FullName).Hash -ne (Get-FileHash -LiteralPath $installed).Hash) { throw "安装文件校验失败：$relative" }
        }
        if (-not (Test-Path -LiteralPath (Join-Path $testRoot '_internal/assets/statement-template.xlsx'))) { throw '脱敏模板未安装。' }
        $gui = Start-Process -FilePath $exe -WindowStyle Hidden -PassThru
        Start-Sleep -Seconds 3
        $gui.Refresh()
        if ($gui.HasExited -or $gui.MainWindowTitle -ne '自动对账工具') { throw '安装后的GUI启动失败。' }
        $gui.CloseMainWindow() | Out-Null
        if (-not $gui.WaitForExit(10000)) { throw 'GUI未正常关闭。' }
        $quote = Get-ChildItem -File '*报价*.xlsx' | Select-Object -First 1
        $delivery = Get-ChildItem -File '送货*.xlsx' | Select-Object -First 1
        if ($quote -and $delivery) {
            $output = Join-Path $projectRoot 'outputs/installer-smoke/安装后导出.xlsx'
            $arguments = '--quote "{0}" --delivery "{1}" --output "{2}"' -f $quote.FullName,$delivery.FullName,$output
            $cli = Start-Process -FilePath $exe -ArgumentList $arguments -WindowStyle Hidden -PassThru
            if (-not $cli.WaitForExit(30000) -or $cli.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $output)) { throw '安装后样例导出失败。' }
        }
        $sentinel = Join-Path $testRoot 'user-data-preservation.txt'
        '模拟用户自行保存的数据，应在卸载后保留。' | Set-Content -LiteralPath $sentinel
    } finally {
        # 仅运行刚刚安装到已验证目录的卸载器，不使用通配符删除目录。
        if (Test-Path -LiteralPath $uninstaller) {
            $remove = Start-Process -FilePath $uninstaller -ArgumentList '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -WindowStyle Hidden -PassThru
            if (-not $remove.WaitForExit(60000)) { throw '卸载超时。' }
        }
    }
    Start-Sleep -Seconds 2
    if (Test-Path -LiteralPath $exe) { throw '卸载后主程序仍存在。' }
    if (Test-Path $regPath) { throw '卸载后注册信息仍存在。' }
    if (-not (Test-Path -LiteralPath $sentinel)) { throw '卸载错误删除了用户数据。' }
    Write-Output 'INSTALLER PASS：安装、全部文件SHA256、GUI启动、样例导出、卸载、保留用户数据。'
    Write-Output "验收目录保留模拟用户文件：$sentinel"
} finally { Pop-Location }
