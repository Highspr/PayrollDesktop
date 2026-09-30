param([switch]$Installer)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root
$python = Join-Path $root '.desktop-venv\Scripts\python.exe'
if (!(Test-Path $python)) { throw 'Create .desktop-venv and install requirements-dev.txt first. See README.md.' }
$bundledNode = Join-Path $root '.tools\node-v22.23.3-win-x64'
if (Test-Path $bundledNode) { $env:PATH = "$bundledNode;$env:PATH" }
$env:NEXT_TELEMETRY_DISABLED='1'
Push-Location frontend
try {
    & npm.cmd ci --no-fund
    if ($LASTEXITCODE -ne 0) { throw 'npm dependency installation failed' }
    & npm.cmd run build
    if ($LASTEXITCODE -ne 0) { throw 'Frontend build failed' }
} finally { Pop-Location }
& $python -m pytest -q
if ($LASTEXITCODE -ne 0) { throw 'Payroll tests failed' }
& $python -m PyInstaller --noconfirm PayrollDesk.spec
if ($LASTEXITCODE -ne 0) { throw 'Desktop packaging failed' }
& $python scripts\collect_licenses.py
if ($LASTEXITCODE -ne 0) { throw 'Dependency notices could not be collected' }
if ($Installer) {
    $compiler = Join-Path $root '.tools\inno-free\ISCC.exe'
    if (!(Test-Path $compiler)) { $compiler = 'C:\Program Files (x86)\Inno Setup 6\ISCC.exe' }
    if (!(Test-Path $compiler)) { throw 'Install Inno Setup 6 to create the installer. The portable desktop build is already in dist\PayrollDesk.' }
    & $compiler (Join-Path $PSScriptRoot 'installer.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Installer build failed' }
}
