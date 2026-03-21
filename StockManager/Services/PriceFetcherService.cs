using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using StockManager.Config;

namespace StockManager.Services
{
        public class PriceFetcherService : IMarketDataGateway
        {
                private readonly object _lock = new object();
                private Dictionary<string, Tuple<double?, double?>> _latestPrices = new Dictionary<string, Tuple<double?, double?>>();
                private Dictionary<string, Dictionary<string, object>> _latestPriceMeta = new Dictionary<string, Dictionary<string, object>>();

                public Tuple<double?, double?, string> GetRealtimePriceWithSource(string ticker)
                {
                        Console.WriteLine($"[數據源] 僅使用 Python yfinance：{ticker}");

                        if (!AppConfig.UsePythonYFinance)
                        {
                                Console.WriteLine("[數據源] UsePythonYFinance=false，無法獲取價格");
                                return Tuple.Create<double?, double?, string>(null, null, "yfinance-disabled");
                        }

                        var result = TryPythonYFinance(ticker);
                        if (result.Item1.HasValue)
                        {
                                return result;
                        }

                        Console.WriteLine($"[數據源] yfinance 失敗，不再切換其他 API：{ticker}");
                        return Tuple.Create<double?, double?, string>(null, null, "yfinance-failed");
                }

                /// <summary>
                /// 嘗試使用 Python yfinance 獲取股價
                /// 文檔: https://pypi.org/project/yfinance/
                /// </summary>
                private Tuple<double?, double?, string> TryPythonYFinance(string ticker)
                {
                        try
                        {
                                Console.WriteLine($"[Python yfinance] 🐍 嘗試獲取 {ticker} 的股價...");

                                var exePath = ResolveYFinanceExecutablePath();
                                // 構建 Python 腳本的完整路徑
                                var scriptPath = ResolveYFinanceScriptPath();
                                Console.WriteLine($"[Python yfinance] 腳本路徑: {scriptPath}");
                                Console.WriteLine($"[Python yfinance] EXE 路徑: {exePath}");
                                Console.WriteLine($"[Python yfinance] Python 命令: {AppConfig.PythonPath}");

                                var hasExe = File.Exists(exePath);
                                var hasScript = File.Exists(scriptPath);

                                if (!hasExe && !hasScript)
                                {
                                        Console.WriteLine($"[Python yfinance] ❌ 失敗原因: EXE 與腳本都不存在");
                                        Console.WriteLine($"[Python yfinance] 請確認文件存在: {scriptPath}");
                                        return Tuple.Create<double?, double?, string>(null, null, null);
                                }

                                if (hasExe)
                                {
                                        Console.WriteLine($"[Python yfinance] ✅ 使用 PyInstaller EXE");
                                }
                                else
                                {
                                        Console.WriteLine($"[Python yfinance] ✅ 使用 Python 腳本");
                                }

                                // 配置 Python 進程
                                var startInfo = CreateYFinanceProcessStartInfo(ticker, exePath, scriptPath);
                                Console.WriteLine($"[Python yfinance] 執行命令: {startInfo.FileName} {startInfo.Arguments}");

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        var startTime = DateTime.Now;
                                        process.Start();

                                        // 讀取輸出
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();

                                        process.WaitForExit(10000); // 10 秒超時
                                        var duration = (DateTime.Now - startTime).TotalSeconds;

                                        Console.WriteLine($"[Python yfinance] 執行時間: {duration:F2} 秒");
                                        Console.WriteLine($"[Python yfinance] 退出代碼: {process.ExitCode}");

                                        if (!string.IsNullOrEmpty(error))
                                        {
                                                Console.WriteLine($"[Python yfinance] ⚠️ 錯誤輸出:");
                                                Console.WriteLine($"{error}");

                                                // 檢查常見錯誤
                                                if (error.Contains("No module named"))
                                                {
                                                        Console.WriteLine($"[Python yfinance] ❌ 失敗原因: yfinance 模組未安裝");
                                                        Console.WriteLine($"[Python yfinance] 解決方案: 運行 'pip install yfinance'");
                                                }
                                                else if (error.Contains("python") && error.Contains("not found"))
                                                {
                                                        Console.WriteLine($"[Python yfinance] ❌ 失敗原因: Python 未找到");
                                                        Console.WriteLine($"[Python yfinance] 解決方案: 確認 Python 已安裝並添加到 PATH");
                                                }
                                        }

                                        if (!string.IsNullOrEmpty(output))
                                        {
                                                Console.WriteLine($"[Python yfinance] 📄 JSON 輸出長度: {output.Length} 字符");

                                                // 解析 JSON 輸出
                                                var result = ParseYFinanceJson(output, ticker);
                                                if (result.Item1.HasValue)
                                                {
                                                        Console.WriteLine($"[Python yfinance] ✅ 成功解析數據");
                                                        return result;
                                                }
                                                else
                                                {
                                                        Console.WriteLine($"[Python yfinance] ❌ 失敗原因: JSON 解析失敗或無有效數據");
                                                        Console.WriteLine($"[Python yfinance] JSON 內容前 200 字符:");
                                                        Console.WriteLine($"{output.Substring(0, Math.Min(200, output.Length))}");
                                                }
                                        }
                                        else
                                        {
                                                Console.WriteLine($"[Python yfinance] ❌ 失敗原因: 無輸出");
                                                if (process.ExitCode != 0)
                                                {
                                                        Console.WriteLine($"[Python yfinance] Python 進程異常退出，退出代碼: {process.ExitCode}");
                                                }
                                        }
                                }
                        }
                        catch (System.ComponentModel.Win32Exception ex)
                        {
                                Console.WriteLine($"[Python yfinance] ❌ 失敗原因: 無法執行 Python");
                                Console.WriteLine($"[Python yfinance] 錯誤: {ex.Message}");
                                Console.WriteLine($"[Python yfinance] 可能原因:");
                                Console.WriteLine($"  1. Python 未安裝");
                                Console.WriteLine($"  2. Python 未添加到 PATH");
                                Console.WriteLine($"  3. PythonPath 配置錯誤 (當前: {AppConfig.PythonPath})");
                                Console.WriteLine($"[Python yfinance] 解決方案:");
                                Console.WriteLine($"  1. 安裝 Python: https://www.python.org/downloads/");
                                Console.WriteLine($"  2. 確保勾選 'Add Python to PATH'");
                                Console.WriteLine($"  3. 或修改 AppConfig.cs 中的 PythonPath");
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Python yfinance] ❌ 失敗原因: 未預期的錯誤");
                                Console.WriteLine($"[Python yfinance] 錯誤類型: {ex.GetType().Name}");
                                Console.WriteLine($"[Python yfinance] 錯誤信息: {ex.Message}");
                                Console.WriteLine($"[Python yfinance] 堆疊追蹤: {ex.StackTrace}");
                        }

                        Console.WriteLine($"[Python yfinance] ⚠️ 返回 null，將嘗試下一個數據源");
                        return Tuple.Create<double?, double?, string>(null, null, null);
                }

                /// <summary>
                /// 解析 yfinance 返回的 JSON
                /// </summary>
                private Tuple<double?, double?, string> ParseYFinanceJson(string json, string ticker)
                {
                        try
                        {
                                // 檢查是否成功
                                var successMatch = Regex.Match(json, "\"success\"\\s*:\\s*(true|false)");
                                if (!successMatch.Success || successMatch.Groups[1].Value == "false")
                                {
                                        // 提取錯誤信息
                                        var errorMatch = Regex.Match(json, "\"error\"\\s*:\\s*\"([^\"]+)\"");
                                        if (errorMatch.Success)
                                        {
                                                Console.WriteLine($"[Python yfinance] {ticker} API 錯誤: {errorMatch.Groups[1].Value}");
                                        }
                                        return Tuple.Create<double?, double?, string>(null, null, null);
                                }

                                // 提取價格數據
                                var currentPrice = ExtractYFinanceValue(json, "current_price");
                                var previousClose = ExtractYFinanceValue(json, "previous_close");
                                var changePercent = ExtractYFinanceValue(json, "change_percent");

                                // 如果沒有漲跌幅但有價格，計算它
                                if (!changePercent.HasValue && currentPrice.HasValue && previousClose.HasValue && previousClose.Value != 0)
                                {
                                        var change = currentPrice.Value - previousClose.Value;
                                        changePercent = (change / previousClose.Value) * 100;
                                }

                                if (currentPrice.HasValue)
                                {
                                        var marketState = ExtractYFinanceString(json, "market_state");
                                        var stateTag = string.IsNullOrEmpty(marketState) ? "" : $" ({marketState})";

                                        Console.WriteLine($"[Python yfinance] {ticker}: 當前=${currentPrice.Value:F2}, 前收=${previousClose?.ToString("F2") ?? "N/A"}, 漲跌幅={changePercent?.ToString("F2") ?? "N/A"}%{stateTag}");

                                        return Tuple.Create(currentPrice, changePercent, $"yfinance{stateTag}");
                                }
                                else
                                {
                                        Console.WriteLine($"[Python yfinance] {ticker}: 無法解析價格");
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Python yfinance] 解析 JSON 失敗: {ex.Message}");
                        }

                        return Tuple.Create<double?, double?, string>(null, null, null);
                }

                /// <summary>
                /// 從 yfinance JSON 中提取數值
                /// </summary>
                private double? ExtractYFinanceValue(string json, string fieldName)
                {
                        try
                        {
                                var pattern = $"\"{fieldName}\"\\s*:\\s*([0-9.eE+-]+)";
                                var match = Regex.Match(json, pattern);

                                if (match.Success)
                                {
                                        var valueStr = match.Groups[1].Value;
                                        if (double.TryParse(valueStr, System.Globalization.NumberStyles.Any,
                                                System.Globalization.CultureInfo.InvariantCulture, out double value))
                                        {
                                                return value;
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Python yfinance] 解析 {fieldName} 失敗: {ex.Message}");
                        }

                        return null;
                }

                /// <summary>
                /// 從 yfinance JSON 中提取字符串
                /// </summary>
                private string ExtractYFinanceString(string json, string fieldName)
                {
                        try
                        {
                                var pattern = $"\"{fieldName}\"\\s*:\\s*\"([^\"]+)\"";
                                var match = Regex.Match(json, pattern);

                                if (match.Success)
                                {
                                        return match.Groups[1].Value;
                                }
                        }
                        catch
                        {
                                // 忽略錯誤
                        }

                        return null;
                }

                public void UpdatePrice(string ticker)
                {
                        var result = GetRealtimePriceWithSource(ticker);
                        var price = result.Item1;
                        var pct = result.Item2;
                        var source = result.Item3;
                        var updatedAt = price.HasValue ? (DateTime?)DateTime.Now : null;

                        lock (_lock)
                        {
                                _latestPrices[ticker] = Tuple.Create(price, pct);
                                _latestPriceMeta[ticker] = new Dictionary<string, object>
                                {
                                        { "source", source },
                                        { "updated_at", updatedAt },
                                        { "previous_close", null }  // 先設為 null，後面會更新
                                };
                        }
                }

                /// <summary>
                /// 更新價格並返回前收盤價（用於手動計算漲跌幅）
                /// </summary>
                public void UpdatePriceWithPreviousClose(string ticker)
                {
                        try
                        {
                                if (!AppConfig.UsePythonYFinance)
                                {
                                        Console.WriteLine("[更新] UsePythonYFinance=false，已跳過更新");
                                        return;
                                }

                                if (!TryUpdateWithPythonYFinance(ticker))
                                {
                                        Console.WriteLine($"[更新] yfinance 更新失敗（不使用其他 API）：{ticker}");
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[更新錯誤] {ticker}: {ex.Message}");
                        }
                }

                /// <summary>
                /// 嘗試使用 Python yfinance 更新價格
                /// </summary>
                private bool TryUpdateWithPythonYFinance(string ticker)
                {
                        try
                        {
                                var exePath = ResolveYFinanceExecutablePath();
                                var scriptPath = ResolveYFinanceScriptPath();

                                if (!File.Exists(exePath) && !File.Exists(scriptPath))
                                {
                                        return false;
                                }

                                var startInfo = CreateYFinanceProcessStartInfo(ticker, exePath, scriptPath);

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        process.Start();
                                        var output = process.StandardOutput.ReadToEnd();
                                        process.WaitForExit(10000);

                                        if (!string.IsNullOrEmpty(output))
                                        {
                                                var currentPrice = ExtractYFinanceValue(output, "current_price");
                                                var previousClose = ExtractYFinanceValue(output, "previous_close");

                                                if (currentPrice.HasValue)
                                                {
                                                        var marketState = ExtractYFinanceString(output, "market_state");
                                                        var source = string.IsNullOrEmpty(marketState) ? "yfinance" : $"yfinance ({marketState})";

                                                        lock (_lock)
                                                        {
                                                                _latestPrices[ticker] = Tuple.Create(currentPrice, (double?)null);
                                                                _latestPriceMeta[ticker] = new Dictionary<string, object>
                                                                {
                                                                        { "source", source },
                                                                        { "updated_at", DateTime.Now },
                                                                        { "previous_close", previousClose }
                                                                };
                                                        }

                                                        Console.WriteLine($"[yfinance更新] {ticker}: 當前=${currentPrice.Value:F2}, 前收=${previousClose?.ToString("F2") ?? "null"}");
                                                        return true;
                                                }
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[yfinance更新失敗] {ticker}: {ex.Message}");
                        }

                        return false;
                }

                /// <summary>
                /// 建立 yfinance 子程序啟動資訊（優先使用 PyInstaller EXE）
                /// </summary>
                private ProcessStartInfo CreateYFinanceProcessStartInfo(string ticker, string exePath, string scriptPath)
                {
                        return CreateYFinanceCommandProcessStartInfo(exePath, scriptPath, ticker);
                }

                private ProcessStartInfo CreateYFinanceCommandProcessStartInfo(string exePath, string scriptPath, string arguments)
                {
                        if (File.Exists(exePath))
                        {
                                return new ProcessStartInfo
                                {
                                        FileName = exePath,
                                        Arguments = arguments,
                                        UseShellExecute = false,
                                        RedirectStandardOutput = true,
                                        RedirectStandardError = true,
                                        CreateNoWindow = true,
                                        StandardOutputEncoding = System.Text.Encoding.UTF8
                                };
                        }

                        return new ProcessStartInfo
                        {
                                FileName = AppConfig.PythonPath,
                                Arguments = $"\"{scriptPath}\" {arguments}",
                                UseShellExecute = false,
                                RedirectStandardOutput = true,
                                RedirectStandardError = true,
                                CreateNoWindow = true,
                                StandardOutputEncoding = System.Text.Encoding.UTF8
                        };
                }

                /// <summary>
                /// 解析 yfinance EXE 路徑（PyInstaller onefile 產物）。
                /// </summary>
                private string ResolveYFinanceExecutablePath()
                {
                        var baseDir = AppDomain.CurrentDomain.BaseDirectory;

                        // 1) 安裝/輸出目錄：Python\yfinance_fetcher.exe
                        var outputPath = Path.Combine(baseDir, "Python", "yfinance_fetcher.exe");
                        if (File.Exists(outputPath))
                        {
                                return outputPath;
                        }

                        // 2) 專案目錄：..\..\Python\dist\yfinance_fetcher.exe
                        var projectDistPath = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "Python", "dist", "yfinance_fetcher.exe"));
                        if (File.Exists(projectDistPath))
                        {
                                return projectDistPath;
                        }

                        return outputPath;
                }

                /// <summary>
                /// 解析 yfinance 腳本路徑。
                /// 優先使用輸出目錄，若不存在則回退到專案目錄。
                /// </summary>
                private string ResolveYFinanceScriptPath()
                {
                        var baseDir = AppDomain.CurrentDomain.BaseDirectory;

                        // 1) bin\Debug\Python\yfinance_fetcher.py
                        var outputPath = Path.Combine(baseDir, AppConfig.YFinanceScriptPath);
                        if (File.Exists(outputPath))
                        {
                                return outputPath;
                        }

                        // 2) 回退到專案目錄: ..\..\Python\yfinance_fetcher.py
                        var projectPath = Path.GetFullPath(Path.Combine(baseDir, "..", "..", AppConfig.YFinanceScriptPath));
                        if (File.Exists(projectPath))
                        {
                                return projectPath;
                        }

                        // 3) 再回退一層（某些執行環境）: ..\..\..\Python\yfinance_fetcher.py
                        var fallbackPath = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", AppConfig.YFinanceScriptPath));
                        return fallbackPath;
                }

                public bool TryGetHistoricalData(string ticker, string period, string interval, out List<MarketHistoryBar> bars)
                {
                        bars = new List<MarketHistoryBar>();

                        try
                        {
                                var exePath = ResolveYFinanceExecutablePath();
                                var scriptPath = ResolveYFinanceScriptPath();

                                if (!File.Exists(exePath) && !File.Exists(scriptPath))
                                {
                                        return false;
                                }

                                var arguments = $"{ticker} history {period} {interval}";
                                var startInfo = CreateYFinanceCommandProcessStartInfo(exePath, scriptPath, arguments);

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        process.Start();
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();
                                        process.WaitForExit(15000);

                                        if (!string.IsNullOrWhiteSpace(error))
                                        {
                                                Console.WriteLine($"[yfinance history] {ticker} 錯誤: {error}");
                                        }

                                        var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                                        if (lines.Length == 0 || lines[0] != "HISTORY_OK")
                                        {
                                                return false;
                                        }

                                        for (int i = 1; i < lines.Length; i++)
                                        {
                                                var parts = lines[i].Split('|');
                                                if (parts.Length < 6)
                                                {
                                                        continue;
                                                }

                                                double open;
                                                double high;
                                                double low;
                                                double close;
                                                long volume;

                                                if (!double.TryParse(parts[1], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out open) ||
                                                        !double.TryParse(parts[2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out high) ||
                                                        !double.TryParse(parts[3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out low) ||
                                                        !double.TryParse(parts[4], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out close) ||
                                                        !long.TryParse(parts[5], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out volume))
                                                {
                                                        continue;
                                                }

                                                bars.Add(new MarketHistoryBar
                                                {
                                                        DateText = parts[0],
                                                        Open = open,
                                                        High = high,
                                                        Low = low,
                                                        Close = close,
                                                        Volume = volume
                                                });
                                        }
                                }

                                return bars.Count > 0;
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[yfinance history] {ticker} 失敗: {ex.Message}");
                                return false;
                        }
                }

                public bool TryGetFundamentals(string ticker, out Dictionary<string, double> metrics)
                {
                        metrics = new Dictionary<string, double>();

                        try
                        {
                                var exePath = ResolveYFinanceExecutablePath();
                                var scriptPath = ResolveYFinanceScriptPath();

                                if (!File.Exists(exePath) && !File.Exists(scriptPath))
                                {
                                        return false;
                                }

                                var arguments = $"{ticker} fundamentals";
                                var startInfo = CreateYFinanceCommandProcessStartInfo(exePath, scriptPath, arguments);

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        process.Start();
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();
                                        process.WaitForExit(15000);

                                        if (!string.IsNullOrWhiteSpace(error))
                                        {
                                                Console.WriteLine($"[yfinance fundamentals] {ticker} 錯誤: {error}");
                                        }

                                        var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                                        if (lines.Length == 0 || lines[0] != "FUNDAMENTALS_OK")
                                        {
                                                return false;
                                        }

                                        for (int i = 1; i < lines.Length; i++)
                                        {
                                                var parts = lines[i].Split('|');
                                                if (parts.Length != 2)
                                                {
                                                        continue;
                                                }

                                                double value;
                                                if (double.TryParse(parts[1], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out value))
                                                {
                                                        metrics[parts[0]] = value;
                                                }
                                        }
                                }

                                return true;
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[yfinance fundamentals] {ticker} 失敗: {ex.Message}");
                                return false;
                        }
                }

                public Dictionary<string, Tuple<double?, double?>> GetPrices()
                {
                        lock (_lock)
                        {
                                return new Dictionary<string, Tuple<double?, double?>>(_latestPrices);
                        }
                }

                public Dictionary<string, Dictionary<string, object>> GetPriceMeta()
                {
                        lock (_lock)
                        {
                                return new Dictionary<string, Dictionary<string, object>>(_latestPriceMeta);
                        }
                }

                /// <summary>
                /// 從緩存中獲取股票的價格和漲跌幅
                /// </summary>
                public Tuple<double?, double?> GetCachedPrice(string ticker)
                {
                        lock (_lock)
                        {
                                if (_latestPrices.TryGetValue(ticker, out var cached))
                                {
                                        Console.WriteLine($"[緩存] {ticker}: Price={cached.Item1}, Change={cached.Item2}");
                                        return cached;
                                }
                                Console.WriteLine($"[緩存] {ticker}: 無緩存數據");
                                return Tuple.Create<double?, double?>(null, null);
                        }
                }

                public Tuple<double?, double?, DateTime?> GetRealtimeTwQuote(string ticker)
                {
                        return null;
                }
        }
}
