param(
    [string]$Configuration = "Release",
    [string]$Platform = "AnyCPU"
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Resolve-Path (Join-Path $scriptDir "..")
$projectPath = Join-Path $repoRoot "StockManager\StockManager.csproj"
$projectDir = Join-Path $repoRoot "StockManager"
$outputDir = Join-Path $projectDir "bin\$Configuration"
$stagingDir = Join-Path $scriptDir "staging"
$installerOutputDir = Join-Path $scriptDir "Output"

function Get-IsccPath {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe"
    )

    foreach ($path in $candidates) {
        if ($path -and (Test-Path $path)) {
            return $path
        }
    }

    $cmd = Get-Command iscc -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    throw "Cannot find Inno Setup compiler (ISCC.exe). Please install Inno Setup 6 first."
}

function Get-PythonCommand {
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) { return "py" }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) { return "python" }

    throw "Cannot find Python. Please install Python 3.x (recommended 3.10+)."
}

function Get-MSBuildPath {
    $cmd = Get-Command msbuild -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $vsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vsWhere) {
        $found = & $vsWhere -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
        if ($found -and (Test-Path $found)) {
            return $found
        }
    }

    throw "Cannot find MSBuild. Please run in Developer PowerShell or install Visual Studio Build Tools."
}

Write-Host "[1/5] Building StockManager ($Configuration)..." -ForegroundColor Cyan
$msbuildPath = Get-MSBuildPath
& $msbuildPath $projectPath /t:Build /p:Configuration=$Configuration /p:Platform=$Platform

if (-not (Test-Path (Join-Path $outputDir "StockManager.exe"))) {
    throw "Build failed: StockManager.exe was not found in output folder ($outputDir)."
}

Write-Host "[2/5] Preparing staging folder..." -ForegroundColor Cyan
if (Test-Path $stagingDir) {
    Remove-Item $stagingDir -Recurse -Force
}
New-Item -ItemType Directory -Path $stagingDir | Out-Null

Write-Host "[3/5] Copying app files to staging..." -ForegroundColor Cyan
Copy-Item (Join-Path $outputDir "*") $stagingDir -Recurse -Force

$pythonScriptDir = Join-Path $stagingDir "Python"
New-Item -ItemType Directory -Path $pythonScriptDir -Force | Out-Null

Copy-Item (Join-Path $projectDir "Python\*.py") $pythonScriptDir -Force
if (Test-Path (Join-Path $projectDir "Python\requirements.txt")) {
    Copy-Item (Join-Path $projectDir "Python\requirements.txt") $pythonScriptDir -Force
}

Write-Host "[4/5] Building yfinance_fetcher.exe with PyInstaller..." -ForegroundColor Cyan
$pythonCmd = Get-PythonCommand

$pyBuildDir = Join-Path $scriptDir "pybuild"
if (Test-Path $pyBuildDir) {
    Remove-Item $pyBuildDir -Recurse -Force
}
New-Item -ItemType Directory -Path $pyBuildDir | Out-Null

$pyBuildVenv = Join-Path $pyBuildDir "venv"
if ($pythonCmd -eq "py") {
    & py -3 -m venv $pyBuildVenv
} else {
    & python -m venv $pyBuildVenv
}

$venvPython = Join-Path $pyBuildVenv "Scripts\python.exe"
if (-not (Test-Path $venvPython)) {
    throw "Virtual environment creation failed: $venvPython not found."
}

& $venvPython -m pip install --upgrade pip
if (Test-Path (Join-Path $projectDir "Python\requirements.txt")) {
    & $venvPython -m pip install -r (Join-Path $projectDir "Python\requirements.txt")
}
& $venvPython -m pip install pyinstaller

$entryScript = Join-Path $projectDir "Python\yfinance_fetcher.py"
if (-not (Test-Path $entryScript)) {
    throw "Cannot find yfinance_fetcher.py: $entryScript"
}

& $venvPython -m PyInstaller --noconfirm --clean --onefile --name yfinance_fetcher `
    --distpath $pythonScriptDir `
    --workpath (Join-Path $pyBuildDir "work") `
    --specpath (Join-Path $pyBuildDir "spec") `
    $entryScript

$builtExe = Join-Path $pythonScriptDir "yfinance_fetcher.exe"
if (-not (Test-Path $builtExe)) {
    throw "PyInstaller build failed: $builtExe not found."
}

# 清理建置暫存
if (Test-Path $pyBuildDir) {
    Remove-Item $pyBuildDir -Recurse -Force
}

Write-Host "[5/5] Building installer..." -ForegroundColor Cyan
$isccPath = Get-IsccPath
if (-not (Test-Path $installerOutputDir)) {
    New-Item -ItemType Directory -Path $installerOutputDir | Out-Null
}

& $isccPath (Join-Path $scriptDir "StockManager.iss") "/DPublishDir=$stagingDir" "/DOutputDir=$installerOutputDir"

Write-Host "Done. Installer output: $installerOutputDir" -ForegroundColor Green
