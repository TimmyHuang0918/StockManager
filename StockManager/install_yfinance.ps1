# Python yfinance 一鍵安裝腳本
# 自動安裝 Python 依賴

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Python yfinance 環境檢查與安裝" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 檢查 Python 是否安裝
Write-Host "檢查 Python 安裝..." -ForegroundColor Yellow

try {
    $pythonVersion = & python --version 2>&1
    Write-Host "  ✅ 找到: $pythonVersion" -ForegroundColor Green
    $pythonCmd = "python"
}
catch {
    try {
        $pythonVersion = & python3 --version 2>&1
        Write-Host "  ✅ 找到: $pythonVersion" -ForegroundColor Green
        $pythonCmd = "python3"
    }
    catch {
        Write-Host "  ❌ 未找到 Python" -ForegroundColor Red
        Write-Host ""
        Write-Host "請先安裝 Python:" -ForegroundColor Yellow
        Write-Host "  1. 訪問: https://www.python.org/downloads/" -ForegroundColor White
        Write-Host "  2. 下載並安裝最新版本" -ForegroundColor White
        Write-Host "  3. ⚠️  安裝時勾選 'Add Python to PATH'" -ForegroundColor Yellow
        Write-Host ""
        pause
        exit 1
    }
}

Write-Host ""

# 檢查 pip 是否可用
Write-Host "檢查 pip 安裝..." -ForegroundColor Yellow

try {
    $pipVersion = & pip --version 2>&1
    Write-Host "  ✅ 找到: $pipVersion" -ForegroundColor Green
    $pipCmd = "pip"
}
catch {
    try {
        $pipVersion = & pip3 --version 2>&1
        Write-Host "  ✅ 找到: $pipVersion" -ForegroundColor Green
        $pipCmd = "pip3"
    }
    catch {
        Write-Host "  ❌ 未找到 pip" -ForegroundColor Red
        Write-Host "  嘗試使用 python -m pip..." -ForegroundColor Yellow
        $pipCmd = "$pythonCmd -m pip"
    }
}

Write-Host ""

# 檢查 yfinance 是否已安裝
Write-Host "檢查 yfinance 庫..." -ForegroundColor Yellow

try {
    $yfinanceCheck = & $pythonCmd -c "import yfinance; print(yfinance.__version__)" 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✅ yfinance 已安裝: $yfinanceCheck" -ForegroundColor Green
        Write-Host ""
        $response = Read-Host "是否要升級到最新版本? (y/N)"
        
        if ($response -eq 'y' -or $response -eq 'Y') {
            Write-Host ""
            Write-Host "升級 yfinance..." -ForegroundColor Yellow
            & $pipCmd install --upgrade yfinance
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  ✅ yfinance 升級成功" -ForegroundColor Green
            }
            else {
                Write-Host "  ⚠️  升級失敗，但現有版本仍可使用" -ForegroundColor Yellow
            }
        }
    }
}
catch {
    Write-Host "  ❌ yfinance 未安裝" -ForegroundColor Red
    Write-Host ""
    Write-Host "安裝 yfinance..." -ForegroundColor Yellow
    
    & $pipCmd install yfinance
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✅ yfinance 安裝成功" -ForegroundColor Green
        
        # 再次檢查版本
        $yfinanceVersion = & $pythonCmd -c "import yfinance; print(yfinance.__version__)" 2>&1
        Write-Host "  版本: $yfinanceVersion" -ForegroundColor White
    }
    else {
        Write-Host "  ❌ yfinance 安裝失敗" -ForegroundColor Red
        Write-Host ""
        Write-Host "請手動安裝:" -ForegroundColor Yellow
        Write-Host "  $pipCmd install yfinance" -ForegroundColor White
        Write-Host ""
        pause
        exit 1
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "測試 yfinance 腳本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 測試腳本
$scriptPath = "Python\yfinance_fetcher.py"

if (Test-Path $scriptPath) {
    Write-Host "測試股票代碼: NVDA" -ForegroundColor Yellow
    Write-Host ""
    
    & $pythonCmd $scriptPath NVDA
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "  ✅ 測試成功！" -ForegroundColor Green
    }
    else {
        Write-Host ""
        Write-Host "  ⚠️  測試失敗，但腳本存在" -ForegroundColor Yellow
    }
}
else {
    Write-Host "  ⚠️  找不到腳本: $scriptPath" -ForegroundColor Yellow
    Write-Host "  請確認文件存在" -ForegroundColor White
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "安裝完成！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "下一步:" -ForegroundColor Yellow
Write-Host "  1. 打開 AppConfig.cs" -ForegroundColor White
Write-Host "  2. 設置 UsePythonYFinance = true" -ForegroundColor White
Write-Host "  3. 重新建置並運行應用程式" -ForegroundColor White
Write-Host ""
Write-Host "Python 命令: $pythonCmd" -ForegroundColor Cyan
Write-Host "pip 命令: $pipCmd" -ForegroundColor Cyan
Write-Host ""
pause
