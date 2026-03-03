using System;
using System.Diagnostics;
using System.IO;

namespace StockManager
{
        /// <summary>
        /// Python yfinance 診斷工具
        /// 測試 Python 環境、yfinance 安裝和數據獲取
        /// </summary>
        public static class TestPythonYFinance
        {
                public static void Run()
                {
                        Console.WriteLine("========================================");
                        Console.WriteLine("🐍 Python yfinance 診斷工具");
                        Console.WriteLine("========================================");
                        Console.WriteLine($"時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                        Console.WriteLine("");

                        // 測試 1: 檢查 Python 安裝
                        Console.WriteLine("📌 測試 1: 檢查 Python 安裝");
                        Console.WriteLine("----------------------------------------");
                        TestPythonInstallation();
                        Console.WriteLine("");

                        // 測試 2: 檢查 yfinance 安裝
                        Console.WriteLine("📌 測試 2: 檢查 yfinance 安裝");
                        Console.WriteLine("----------------------------------------");
                        TestYFinanceInstallation();
                        Console.WriteLine("");

                        // 測試 3: 檢查腳本文件
                        Console.WriteLine("📌 測試 3: 檢查腳本文件");
                        Console.WriteLine("----------------------------------------");
                        TestScriptFile();
                        Console.WriteLine("");

                        // 測試 4: 測試單個股票數據獲取
                        Console.WriteLine("📌 測試 4: 測試數據獲取 (AAPL)");
                        Console.WriteLine("----------------------------------------");
                        TestStockData("AAPL");
                        Console.WriteLine("");

                        // 測試 5: 測試台股數據獲取
                        Console.WriteLine("📌 測試 5: 測試台股數據 (2330.TW)");
                        Console.WriteLine("----------------------------------------");
                        TestStockData("2330.TW");
                        Console.WriteLine("");

                        // 測試 6: 直接測試 Python 命令
                        Console.WriteLine("📌 測試 6: 直接測試 Python 命令");
                        Console.WriteLine("----------------------------------------");
                        TestDirectPythonCommand();
                        Console.WriteLine("");

                        Console.WriteLine("========================================");
                        Console.WriteLine("診斷完成！");
                        Console.WriteLine("========================================");
                }

                private static void TestPythonInstallation()
                {
                        try
                        {
                                var pythonPath = Config.AppConfig.PythonPath;
                                Console.WriteLine($"配置的 Python 路徑: {pythonPath}");

                                var startInfo = new ProcessStartInfo
                                {
                                        FileName = pythonPath,
                                        Arguments = "--version",
                                        UseShellExecute = false,
                                        RedirectStandardOutput = true,
                                        RedirectStandardError = true,
                                        CreateNoWindow = true
                                };

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        process.Start();
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();
                                        process.WaitForExit();

                                        if (process.ExitCode == 0)
                                        {
                                                Console.WriteLine($"✅ Python 已安裝: {output.Trim()}");
                                        }
                                        else
                                        {
                                                Console.WriteLine($"❌ Python 執行失敗");
                                                Console.WriteLine($"錯誤: {error}");
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"❌ 無法執行 Python: {ex.Message}");
                                Console.WriteLine($"");
                                Console.WriteLine($"可能的原因:");
                                Console.WriteLine($"1. Python 未安裝");
                                Console.WriteLine($"2. Python 未添加到 PATH");
                                Console.WriteLine($"3. 配置的路徑錯誤");
                                Console.WriteLine($"");
                                Console.WriteLine($"解決方案:");
                                Console.WriteLine($"1. 安裝 Python: https://www.python.org/downloads/");
                                Console.WriteLine($"2. 確保勾選 'Add Python to PATH'");
                                Console.WriteLine($"3. 或修改 AppConfig.cs 中的 PythonPath");
                        }
                }

                private static void TestYFinanceInstallation()
                {
                        try
                        {
                                var pythonPath = Config.AppConfig.PythonPath;
                                var startInfo = new ProcessStartInfo
                                {
                                        FileName = pythonPath,
                                        Arguments = "-c \"import yfinance; print(yfinance.__version__)\"",
                                        UseShellExecute = false,
                                        RedirectStandardOutput = true,
                                        RedirectStandardError = true,
                                        CreateNoWindow = true
                                };

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        process.Start();
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();
                                        process.WaitForExit();

                                        if (process.ExitCode == 0)
                                        {
                                                Console.WriteLine($"✅ yfinance 已安裝: 版本 {output.Trim()}");
                                        }
                                        else
                                        {
                                                Console.WriteLine($"❌ yfinance 未安裝");
                                                Console.WriteLine($"錯誤: {error}");
                                                Console.WriteLine($"");
                                                Console.WriteLine($"解決方案:");
                                                Console.WriteLine($"1. 運行: pip install yfinance");
                                                Console.WriteLine($"2. 或運行: install_yfinance.ps1");
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"❌ 檢查失敗: {ex.Message}");
                        }
                }

                private static void TestScriptFile()
                {
                        try
                        {
                                var scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Config.AppConfig.YFinanceScriptPath);
                                Console.WriteLine($"腳本路徑: {scriptPath}");

                                if (File.Exists(scriptPath))
                                {
                                        Console.WriteLine($"✅ 腳本文件存在");
                                        
                                        var fileInfo = new FileInfo(scriptPath);
                                        Console.WriteLine($"文件大小: {fileInfo.Length} bytes");
                                        Console.WriteLine($"最後修改: {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");

                                        // 讀取前幾行
                                        var lines = File.ReadAllLines(scriptPath);
                                        Console.WriteLine($"總行數: {lines.Length}");
                                        
                                        if (lines.Length > 0)
                                        {
                                                Console.WriteLine($"第一行: {lines[0]}");
                                        }
                                }
                                else
                                {
                                        Console.WriteLine($"❌ 腳本文件不存在");
                                        Console.WriteLine($"");
                                        Console.WriteLine($"解決方案:");
                                        Console.WriteLine($"1. 確認 Python 文件夾存在");
                                        Console.WriteLine($"2. 確認 yfinance_fetcher.py 存在");
                                        Console.WriteLine($"3. 檢查文件路徑配置");
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"❌ 檢查失敗: {ex.Message}");
                        }
                }

                private static void TestStockData(string ticker)
                {
                        try
                        {
                                var pythonPath = Config.AppConfig.PythonPath;
                                var scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Config.AppConfig.YFinanceScriptPath);

                                Console.WriteLine($"測試股票: {ticker}");
                                Console.WriteLine($"Python: {pythonPath}");
                                Console.WriteLine($"腳本: {scriptPath}");
                                Console.WriteLine("");

                                if (!File.Exists(scriptPath))
                                {
                                        Console.WriteLine($"❌ 腳本文件不存在，跳過測試");
                                        return;
                                }

                                var startInfo = new ProcessStartInfo
                                {
                                        FileName = pythonPath,
                                        Arguments = $"\"{scriptPath}\" {ticker}",
                                        UseShellExecute = false,
                                        RedirectStandardOutput = true,
                                        RedirectStandardError = true,
                                        CreateNoWindow = true,
                                        StandardOutputEncoding = System.Text.Encoding.UTF8
                                };

                                Console.WriteLine($"執行命令: {pythonPath} \"{scriptPath}\" {ticker}");
                                Console.WriteLine("");

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        var startTime = DateTime.Now;
                                        process.Start();
                                        
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();
                                        
                                        process.WaitForExit(15000); // 15 秒超時
                                        var endTime = DateTime.Now;
                                        var duration = (endTime - startTime).TotalSeconds;

                                        Console.WriteLine($"執行時間: {duration:F2} 秒");
                                        Console.WriteLine($"退出代碼: {process.ExitCode}");
                                        Console.WriteLine("");

                                        if (!string.IsNullOrEmpty(error))
                                        {
                                                Console.WriteLine($"❌ 錯誤輸出:");
                                                Console.WriteLine(error);
                                                Console.WriteLine("");
                                        }

                                        if (!string.IsNullOrEmpty(output))
                                        {
                                                Console.WriteLine($"📄 JSON 輸出:");
                                                Console.WriteLine(output);
                                                Console.WriteLine("");

                                                // 檢查是否包含 success
                                                if (output.Contains("\"success\": true") || output.Contains("\"success\":true"))
                                                {
                                                        Console.WriteLine($"✅ 成功獲取數據");
                                                        
                                                        // 嘗試提取價格
                                                        var priceMatch = System.Text.RegularExpressions.Regex.Match(output, "\"current_price\"\\s*:\\s*([0-9.]+)");
                                                        if (priceMatch.Success)
                                                        {
                                                                Console.WriteLine($"價格: ${priceMatch.Groups[1].Value}");
                                                        }
                                                }
                                                else if (output.Contains("\"success\": false") || output.Contains("\"success\":false"))
                                                {
                                                        Console.WriteLine($"❌ 獲取數據失敗");
                                                        
                                                        // 提取錯誤信息
                                                        var errorMatch = System.Text.RegularExpressions.Regex.Match(output, "\"error\"\\s*:\\s*\"([^\"]+)\"");
                                                        if (errorMatch.Success)
                                                        {
                                                                Console.WriteLine($"錯誤: {errorMatch.Groups[1].Value}");
                                                        }
                                                }
                                                else
                                                {
                                                        Console.WriteLine($"⚠️ 無法判斷結果");
                                                }
                                        }
                                        else
                                        {
                                                Console.WriteLine($"❌ 無輸出");
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"❌ 測試失敗: {ex.Message}");
                                Console.WriteLine($"堆疊追蹤: {ex.StackTrace}");
                        }
                }

                private static void TestDirectPythonCommand()
                {
                        try
                        {
                                var pythonPath = Config.AppConfig.PythonPath;
                                
                                // 測試簡單的 Python 命令
                                var testCode = "import sys; print(f'Python {sys.version}'); print('UTF-8 Test: 測試中文')";
                                
                                var startInfo = new ProcessStartInfo
                                {
                                        FileName = pythonPath,
                                        Arguments = $"-c \"{testCode}\"",
                                        UseShellExecute = false,
                                        RedirectStandardOutput = true,
                                        RedirectStandardError = true,
                                        CreateNoWindow = true,
                                        StandardOutputEncoding = System.Text.Encoding.UTF8
                                };

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        process.Start();
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();
                                        process.WaitForExit();

                                        if (process.ExitCode == 0)
                                        {
                                                Console.WriteLine($"✅ Python 命令執行成功:");
                                                Console.WriteLine(output);
                                        }
                                        else
                                        {
                                                Console.WriteLine($"❌ Python 命令執行失敗:");
                                                Console.WriteLine(error);
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"❌ 測試失敗: {ex.Message}");
                        }
                }
        }
}
