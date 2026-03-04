using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using StockManager.Config;

namespace StockManager.Services
{
        public class PriceFetcherService
        {
                private readonly object _lock = new object();
                private Dictionary<string, Tuple<double?, double?>> _latestPrices = new Dictionary<string, Tuple<double?, double?>>();
                private Dictionary<string, Dictionary<string, object>> _latestPriceMeta = new Dictionary<string, Dictionary<string, object>>();

                /// <summary>
                /// 獲取美東時間
                /// </summary>
                private DateTime GetEasternTime()
                {
                        try
                        {
                                // UTC-5 (EST) 或 UTC-4 (EDT，夏令時間)
                                var utcNow = DateTime.UtcNow;
                                var easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                                return TimeZoneInfo.ConvertTimeFromUtc(utcNow, easternZone);
                        }
                        catch
                        {
                                // 如果找不到時區，手動計算（簡化版本，使用 UTC-5）
                                return DateTime.UtcNow.AddHours(-5);
                        }
                }

                /// <summary>
                /// 判斷美股當前的交易狀態
                /// </summary>
                private string GetMarketStatus()
                {
                        var et = GetEasternTime();
                        var dayOfWeek = et.DayOfWeek;

                        // 週末休市
                        if (dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday)
                        {
                                return "Closed";
                        }

                        var time = et.TimeOfDay;

                        // 盤前: 4:00 AM - 9:30 AM ET
                        if (time >= new TimeSpan(4, 0, 0) && time < new TimeSpan(9, 30, 0))
                        {
                                return "PreMarket";
                        }
                        // 盤中: 9:30 AM - 4:00 PM ET
                        else if (time >= new TimeSpan(9, 30, 0) && time < new TimeSpan(16, 0, 0))
                        {
                                return "Regular";
                        }
                        // 盤後: 4:00 PM - 8:00 PM ET
                        else if (time >= new TimeSpan(16, 0, 0) && time < new TimeSpan(20, 0, 0))
                        {
                                return "PostMarket";
                        }
                        else
                        {
                                return "Closed";
                        }
                }

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

                /// <summary>
                /// 嘗試使用 Intrinio API 獲取股價
                /// 文檔: https://docs.intrinio.com/documentation/web_api/get_security_realtime_price_v2
                /// </summary>
                private Tuple<double?, double?, string> TryIntrinioAPI(string ticker)
                {
                        try
                        {
                                // Intrinio 使用標準符號格式
                                var intrinioTicker = ticker.Split('.')[0]; // 移除 .TW, .HK 等後綴

                                // 使用實時價格端點
                                // 免費計劃: 500 次請求/天
                                var url = $"{AppConfig.IntrinioBaseUrl}/securities/{intrinioTicker}/prices/realtime?api_key={AppConfig.IntrinioApiKey}";

                                var request = (HttpWebRequest)WebRequest.Create(url);
                                request.Timeout = 10000;
                                request.UserAgent = "StockManager/1.0";
                                request.Accept = "application/json";

                                using (var response = request.GetResponse())
                                using (var stream = response.GetResponseStream())
                                using (var reader = new StreamReader(stream))
                                {
                                        var json = reader.ReadToEnd();

                                        Console.WriteLine($"[Intrinio] {ticker} API 回應長度: {json.Length}");

                                        // 解析 JSON 回應
                                        // Intrinio 格式: {"last_price":151.80,"change":1.25,"percent_change":0.83}
                                        var lastPrice = ExtractIntrinioValue(json, "last_price");
                                        var percentChange = ExtractIntrinioValue(json, "percent_change");
                                        var change = ExtractIntrinioValue(json, "change");

                                        if (lastPrice.HasValue)
                                        {
                                                if (percentChange.HasValue)
                                                {
                                                        Console.WriteLine($"[Intrinio] {ticker}: 價格=${lastPrice.Value:F2}, 變動={change?.ToString("F2") ?? "N/A"}, 漲跌幅={percentChange.Value:F2}%");
                                                }
                                                else if (change.HasValue)
                                                {
                                                        // 如果沒有百分比但有變動金額，嘗試計算
                                                        var previousPrice = lastPrice.Value - change.Value;
                                                        if (previousPrice != 0)
                                                        {
                                                                percentChange = (change.Value / previousPrice) * 100;
                                                                Console.WriteLine($"[Intrinio] {ticker}: 價格=${lastPrice.Value:F2}, 漲跌幅={percentChange.Value:F2}% (計算)");
                                                        }
                                                }
                                                else
                                                {
                                                        Console.WriteLine($"[Intrinio] {ticker}: 價格=${lastPrice.Value:F2} (無漲跌幅數據)");
                                                }

                                                return Tuple.Create(lastPrice, percentChange, "Intrinio");
                                        }
                                        else
                                        {
                                                Console.WriteLine($"[Intrinio] {ticker}: 無法解析價格");
                                        }
                                }
                        }
                        catch (WebException ex)
                        {
                                if (ex.Response != null)
                                {
                                        using (var errorStream = ex.Response.GetResponseStream())
                                        using (var errorReader = new StreamReader(errorStream))
                                        {
                                                var errorContent = errorReader.ReadToEnd();
                                                Console.WriteLine($"[Intrinio] {ticker} API 錯誤: {errorContent}");
                                        }
                                }
                                else
                                {
                                        Console.WriteLine($"[Intrinio] {ticker} 網路錯誤: {ex.Message}");
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Intrinio] {ticker} 失敗: {ex.Message}");
                        }

                        return Tuple.Create<double?, double?, string>(null, null, null);
                }

                /// <summary>
                /// 從 Intrinio JSON 中提取數值
                /// </summary>
                private double? ExtractIntrinioValue(string json, string fieldName)
                {
                        try
                        {
                                // Intrinio 格式: "fieldName":123.45 或 "fieldName":null
                                var pattern = $"\"{fieldName}\":([0-9.eE+-]+)";
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
                                Console.WriteLine($"[Intrinio] 解析 {fieldName} 失敗: {ex.Message}");
                        }

                        return null;
                }

                /// <summary>
                /// 嘗試使用 Marketstack API 獲取股價
                /// 文檔: https://marketstack.com/documentation
                /// </summary>
                private Tuple<double?, double?, string> TryMarketstackAPI(string ticker)
                {
                        try
                        {
                                // Marketstack 使用不同的符號格式（不支持 .TW 等後綴）
                                var marketstackTicker = ticker.Split('.')[0]; // 移除 .TW, .HK 等後綴

                                // 使用 intraday 端點獲取最新價格（實時數據）
                                // 免費計劃: latest 端點，每月 1000 次請求
                                var url = $"{AppConfig.MarketstackBaseUrl}/eod/latest?access_key={AppConfig.MarketstackApiKey}&symbols={marketstackTicker}";

                                var request = (HttpWebRequest)WebRequest.Create(url);
                                request.Timeout = 10000;
                                request.UserAgent = "StockManager/1.0";
                                request.Accept = "application/json";

                                using (var response = request.GetResponse())
                                using (var stream = response.GetResponseStream())
                                using (var reader = new StreamReader(stream))
                                {
                                        var json = reader.ReadToEnd();

                                        Console.WriteLine($"[Marketstack] {ticker} API 回應長度: {json.Length}");

                                        // 解析 JSON 回應
                                        // Marketstack 格式: {"data":[{"open":150.0,"high":152.0,"low":149.0,"close":151.0,"volume":1000000}]}
                                        var closePrice = ExtractMarketstackValue(json, "close");
                                        var openPrice = ExtractMarketstackValue(json, "open");

                                        double? changePercent = null;

                                        // 計算漲跌幅（基於開盤價）
                                        if (closePrice.HasValue && openPrice.HasValue && openPrice.Value != 0)
                                        {
                                                var change = closePrice.Value - openPrice.Value;
                                                changePercent = (change / openPrice.Value) * 100;

                                                Console.WriteLine($"[Marketstack] {ticker}: 開盤=${openPrice.Value:F2}, 收盤=${closePrice.Value:F2}, 漲跌={change:F2} ({changePercent.Value:F2}%)");
                                        }
                                        else if (closePrice.HasValue)
                                        {
                                                Console.WriteLine($"[Marketstack] {ticker}: 收盤=${closePrice.Value:F2} (無開盤價，無法計算漲跌幅)");
                                        }

                                        if (closePrice.HasValue)
                                        {
                                                return Tuple.Create(closePrice, changePercent, "Marketstack");
                                        }
                                        else
                                        {
                                                Console.WriteLine($"[Marketstack] {ticker}: 無法解析價格");
                                        }
                                }
                        }
                        catch (WebException ex)
                        {
                                if (ex.Response != null)
                                {
                                        using (var errorStream = ex.Response.GetResponseStream())
                                        using (var errorReader = new StreamReader(errorStream))
                                        {
                                                var errorContent = errorReader.ReadToEnd();
                                                Console.WriteLine($"[Marketstack] {ticker} API 錯誤: {errorContent}");
                                        }
                                }
                                else
                                {
                                        Console.WriteLine($"[Marketstack] {ticker} 網路錯誤: {ex.Message}");
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Marketstack] {ticker} 失敗: {ex.Message}");
                        }

                        return Tuple.Create<double?, double?, string>(null, null, null);
                }

                /// <summary>
                /// 從 Marketstack JSON 中提取數值
                /// </summary>
                private double? ExtractMarketstackValue(string json, string fieldName)
                {
                        try
                        {
                                // Marketstack 格式: "fieldName":123.45 或 "fieldName":null
                                // 在 data 數組的第一個元素中
                                var pattern = $"\"data\":\\[\\{{[^}}]*\"{fieldName}\":([0-9.eE+-]+)";
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

                                // 簡化版本：直接搜索字段
                                var simplePattern = $"\"{fieldName}\":([0-9.eE+-]+)";
                                var simpleMatch = Regex.Match(json, simplePattern);

                                if (simpleMatch.Success)
                                {
                                        var valueStr = simpleMatch.Groups[1].Value;
                                        if (double.TryParse(valueStr, System.Globalization.NumberStyles.Any,
                                                System.Globalization.CultureInfo.InvariantCulture, out double value))
                                        {
                                                return value;
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Marketstack] 解析 {fieldName} 失敗: {ex.Message}");
                        }

                        return null;
                }

                private Tuple<double?, double?, string> TryYahooFinanceV8(string ticker)
                {
                        try
                        {
                                var url = $"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?interval=1d&range=1d";
                                var request = (HttpWebRequest)WebRequest.Create(url);
                                request.Timeout = 10000;
                                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";
                                request.Accept = "application/json";

                                using (var response = request.GetResponse())
                                using (var stream = response.GetResponseStream())
                                using (var reader = new StreamReader(stream))
                                {
                                        var json = reader.ReadToEnd();

                                        // 判斷當前市場狀態
                                        var marketStatus = GetMarketStatus();
                                        var et = GetEasternTime();

                                        Console.WriteLine($"[時區] 美東時間: {et:yyyy-MM-dd HH:mm:ss} ET, 市場狀態: {marketStatus}");

                                        double? currentPrice = null;
                                        double? changePercent = null;

                                        // 首先獲取前一交易日收盤價（這是計算漲跌幅的基準）
                                        var previousClose = ExtractJsonValue(json, "regularMarketPreviousClose");

                                        // 詳細診斷：顯示所有可能的價格數據
                                        var regularPrice = ExtractJsonValue(json, "regularMarketPrice");
                                        var prePrice = ExtractJsonValue(json, "preMarketPrice");
                                        var postPrice = ExtractJsonValue(json, "postMarketPrice");

                                        Console.WriteLine($"[診斷] {ticker} 所有價格數據:");
                                        Console.WriteLine($"  前收盤 (previousClose): ${previousClose?.ToString("F2") ?? "null"}");
                                        Console.WriteLine($"  盤中價 (regularMarket): ${regularPrice?.ToString("F2") ?? "null"}");
                                        Console.WriteLine($"  盤前價 (preMarket): ${prePrice?.ToString("F2") ?? "null"}");
                                        Console.WriteLine($"  盤後價 (postMarket): ${postPrice?.ToString("F2") ?? "null"}");

                                        // 根據市場狀態選擇當前價格
                                        if (marketStatus == "PreMarket")
                                        {
                                                // 盤前：優先使用盤前價格
                                                currentPrice = ExtractJsonValue(json, "preMarketPrice");
                                                if (!currentPrice.HasValue)
                                                {
                                                        currentPrice = ExtractJsonValue(json, "regularMarketPrice");
                                                }
                                                Console.WriteLine($"[診斷] {ticker} 使用盤前價格: ${currentPrice?.ToString("F2") ?? "null"}");
                                        }
                                        else if (marketStatus == "PostMarket")
                                        {
                                                // 盤後：優先使用盤後價格
                                                currentPrice = ExtractJsonValue(json, "postMarketPrice");
                                                if (!currentPrice.HasValue)
                                                {
                                                        currentPrice = ExtractJsonValue(json, "regularMarketPrice");
                                                }
                                                Console.WriteLine($"[診斷] {ticker} 使用盤後價格: ${currentPrice?.ToString("F2") ?? "null"}");
                                        }
                                        else
                                        {
                                                // 盤中或休市：使用常規市場價格
                                                currentPrice = ExtractJsonValue(json, "regularMarketPrice");
                                                Console.WriteLine($"[診斷] {ticker} 使用盤中/休市價格: ${currentPrice?.ToString("F2") ?? "null"}");
                                        }

                                        // 計算漲跌幅：基於前一交易日收盤價
                                        if (currentPrice.HasValue && previousClose.HasValue && previousClose.Value != 0)
                                        {
                                                var change = currentPrice.Value - previousClose.Value;
                                                changePercent = (change / previousClose.Value) * 100;

                                                var statusTag = marketStatus == "PreMarket" ? "盤前" : 
                                                                        marketStatus == "PostMarket" ? "盤後" : 
                                                                        marketStatus == "Regular" ? "盤中" : "休市";

                                                Console.WriteLine($"[計算] {ticker}: 當前=${currentPrice.Value:F2}, 前收=${previousClose.Value:F2}, 漲跌={change:F2} ({changePercent.Value:F2}%) [{statusTag}]");

                                                // 驗證計算
                                                var expectedChange = ((currentPrice.Value - previousClose.Value) / previousClose.Value) * 100;
                                                Console.WriteLine($"[驗證] {ticker}: 手動計算漲跌幅 = ({currentPrice.Value:F2} - {previousClose.Value:F2}) / {previousClose.Value:F2} × 100 = {expectedChange:F2}%");
                                        }
                                        else if (currentPrice.HasValue)
                                        {
                                                // 有價格但沒有前收盤價，嘗試使用 API 返回的漲跌幅
                                                if (marketStatus == "PreMarket")
                                                {
                                                        changePercent = ExtractJsonValue(json, "preMarketChangePercent");
                                                }
                                                else if (marketStatus == "PostMarket")
                                                {
                                                        changePercent = ExtractJsonValue(json, "postMarketChangePercent");
                                                }
                                                else
                                                {
                                                        changePercent = ExtractJsonValue(json, "regularMarketChangePercent");
                                                }

                                                Console.WriteLine($"[V8] {ticker}: 價格=${currentPrice.Value:F2}, 漲跌={changePercent?.ToString("F2") ?? "N/A"}% (使用API數據，無前收盤價)");
                                        }

                                        if (currentPrice.HasValue)
                                        {
                                                var statusTag = marketStatus == "PreMarket" ? "盤前" : 
                                                                        marketStatus == "PostMarket" ? "盤後" : 
                                                                        marketStatus == "Regular" ? "盤中" : "休市";

                                                return Tuple.Create(currentPrice, changePercent, $"Yahoo V8 ({statusTag})");
                                        }
                                        else
                                        {
                                                Console.WriteLine($"[V8] {ticker}: 無法解析價格");
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[V8] {ticker} 失敗: {ex.Message}");
                        }

                        return Tuple.Create<double?, double?, string>(null, null, null);
                }

                private Tuple<double?, double?, string> TryYahooFinanceV7(string ticker)
                {
                        try
                        {
                                var url = $"https://query2.finance.yahoo.com/v7/finance/quote?symbols={ticker}";
                                var request = (HttpWebRequest)WebRequest.Create(url);
                                request.Timeout = 10000;
                                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";
                                request.Accept = "application/json";

                                using (var response = request.GetResponse())
                                using (var stream = response.GetResponseStream())
                                using (var reader = new StreamReader(stream))
                                {
                                        var json = reader.ReadToEnd();

                                        var currentPrice = ExtractJsonValue(json, "regularMarketPrice");
                                        var previousClose = ExtractJsonValue(json, "regularMarketPreviousClose");
                                        double? changePercent = null;

                                        // 優先自己計算漲跌幅（基於前一交易日收盤價）
                                        if (currentPrice.HasValue && previousClose.HasValue && previousClose.Value != 0)
                                        {
                                                var change = currentPrice.Value - previousClose.Value;
                                                changePercent = (change / previousClose.Value) * 100;
                                                Console.WriteLine($"[V7] {ticker}: 當前=${currentPrice.Value:F2}, 前收=${previousClose.Value:F2}, 漲跌={change:F2} ({changePercent.Value:F2}%)");
                                        }
                                        else if (currentPrice.HasValue)
                                        {
                                                // 如果沒有前收盤價，嘗試使用 API 返回的漲跌幅
                                                changePercent = ExtractJsonValue(json, "regularMarketChangePercent");
                                                Console.WriteLine($"[V7] {ticker}: 價格=${currentPrice.Value:F2}, 漲跌={changePercent?.ToString("F2") ?? "N/A"}% (使用API數據)");
                                        }

                                        if (currentPrice.HasValue)
                                        {
                                                return Tuple.Create(currentPrice, changePercent, "Yahoo V7");
                                        }
                                        else
                                        {
                                                Console.WriteLine($"[V7] {ticker}: 無法解析價格");
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[V7] {ticker} 失敗: {ex.Message}");
                        }

                        return Tuple.Create<double?, double?, string>(null, null, null);
                }

                private double? ExtractJsonValue(string json, string fieldName)
                {
                        try
                        {
                                // 格式 1: "fieldName":{"raw":123.45,"fmt":"123.45"}
                                var pattern1 = $"\"{fieldName}\":\\{{\"raw\":([0-9.eE+-]+)";
                                var match1 = Regex.Match(json, pattern1);
                                if (match1.Success)
                                {
                                        var valueStr = match1.Groups[1].Value;
                                        if (double.TryParse(valueStr, System.Globalization.NumberStyles.Any,
                                                System.Globalization.CultureInfo.InvariantCulture, out double value1))
                                        {
                                                // Console.WriteLine($"  解析 {fieldName} (格式1) = {value1}");
                                                return value1;
                                        }
                                }

                                // 格式 2: "fieldName":123.45
                                var pattern2 = $"\"{fieldName}\":([0-9.eE+-]+)";
                                var match2 = Regex.Match(json, pattern2);
                                if (match2.Success)
                                {
                                        var valueStr = match2.Groups[1].Value;
                                        if (double.TryParse(valueStr, System.Globalization.NumberStyles.Any,
                                                System.Globalization.CultureInfo.InvariantCulture, out double value2))
                                        {
                                                // Console.WriteLine($"  解析 {fieldName} (格式2) = {value2}");
                                                return value2;
                                        }
                                }

                                // 格式 3: "fieldName":{"value":123.45}
                                var pattern3 = $"\"{fieldName}\":\\{{\"value\":([0-9.eE+-]+)";
                                var match3 = Regex.Match(json, pattern3);
                                if (match3.Success)
                                {
                                        var valueStr = match3.Groups[1].Value;
                                        if (double.TryParse(valueStr, System.Globalization.NumberStyles.Any,
                                                System.Globalization.CultureInfo.InvariantCulture, out double value3))
                                        {
                                                // Console.WriteLine($"  解析 {fieldName} (格式3) = {value3}");
                                                return value3;
                                        }
                                }

                                // Console.WriteLine($"  無法找到 {fieldName}");
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"解析 {fieldName} 失敗: {ex.Message}");
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
                        if (File.Exists(exePath))
                        {
                                return new ProcessStartInfo
                                {
                                        FileName = exePath,
                                        Arguments = ticker,
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
                                Arguments = $"\"{scriptPath}\" {ticker}",
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

                /// <summary>
                /// 嘗試使用 Intrinio 更新價格
                /// </summary>
                private bool TryUpdateWithIntrinio(string ticker)
                {
                        try
                        {
                                var intrinioTicker = ticker.Split('.')[0];

                                // 獲取歷史數據（含前收盤價）
                                var url = $"{AppConfig.IntrinioBaseUrl}/securities/{intrinioTicker}/prices?api_key={AppConfig.IntrinioApiKey}&page_size=2";

                                var request = (HttpWebRequest)WebRequest.Create(url);
                                request.Timeout = 10000;
                                request.UserAgent = "StockManager/1.0";

                                using (var response = request.GetResponse())
                                using (var stream = response.GetResponseStream())
                                using (var reader = new StreamReader(stream))
                                {
                                        var json = reader.ReadToEnd();

                                        // 獲取最新的兩天數據
                                        var currentClose = ExtractIntrinioDataByIndex(json, "close", 0);  // 今天
                                        var previousClose = ExtractIntrinioDataByIndex(json, "close", 1); // 昨天

                                        if (currentClose.HasValue)
                                        {
                                                lock (_lock)
                                                {
                                                        _latestPrices[ticker] = Tuple.Create(currentClose, (double?)null);
                                                        _latestPriceMeta[ticker] = new Dictionary<string, object>
                                                        {
                                                                { "source", "Intrinio" },
                                                                { "updated_at", DateTime.Now },
                                                                { "previous_close", previousClose }
                                                        };
                                                }

                                                Console.WriteLine($"[Intrinio更新] {ticker}: 當前=${currentClose.Value:F2}, 前收=${previousClose?.ToString("F2") ?? "null"}");
                                                return true;
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Intrinio更新失敗] {ticker}: {ex.Message}");
                        }

                        return false;
                }

                /// <summary>
                /// 從 Intrinio 數據數組中提取指定索引的值
                /// </summary>
                private double? ExtractIntrinioDataByIndex(string json, string fieldName, int index)
                {
                        try
                        {
                                // 找到 stock_prices 數組
                                var dataPattern = "\"stock_prices\":\\[([^\\]]+)\\]";
                                var dataMatch = Regex.Match(json, dataPattern);

                                if (dataMatch.Success)
                                {
                                        var dataContent = dataMatch.Groups[1].Value;

                                        // 分割成個別物件
                                        var objects = Regex.Split(dataContent, "\\},\\{");

                                        if (index < objects.Length)
                                        {
                                                var targetObject = objects[index];

                                                // 在目標物件中搜索字段
                                                var fieldPattern = $"\"{fieldName}\":([0-9.eE+-]+)";
                                                var fieldMatch = Regex.Match(targetObject, fieldPattern);

                                                if (fieldMatch.Success)
                                                {
                                                        var valueStr = fieldMatch.Groups[1].Value;
                                                        if (double.TryParse(valueStr, System.Globalization.NumberStyles.Any,
                                                                System.Globalization.CultureInfo.InvariantCulture, out double value))
                                                        {
                                                                return value;
                                                        }
                                                }
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Intrinio] 解析索引 {index} 的 {fieldName} 失敗: {ex.Message}");
                        }

                        return null;
                }

                /// <summary>
                /// 嘗試使用 Marketstack 更新價格
                /// </summary>
                private bool TryUpdateWithMarketstack(string ticker)
                {
                        try
                        {
                                var marketstackTicker = ticker.Split('.')[0];
                                var url = $"{AppConfig.MarketstackBaseUrl}/eod?access_key={AppConfig.MarketstackApiKey}&symbols={marketstackTicker}&limit=2";

                                var request = (HttpWebRequest)WebRequest.Create(url);
                                request.Timeout = 10000;
                                request.UserAgent = "StockManager/1.0";

                                using (var response = request.GetResponse())
                                using (var stream = response.GetResponseStream())
                                using (var reader = new StreamReader(stream))
                                {
                                        var json = reader.ReadToEnd();

                                        // 獲取最新的兩天數據
                                        var currentClose = ExtractMarketstackDataByIndex(json, "close", 0);  // 今天
                                        var previousClose = ExtractMarketstackDataByIndex(json, "close", 1); // 昨天

                                        if (currentClose.HasValue)
                                        {
                                                lock (_lock)
                                                {
                                                        _latestPrices[ticker] = Tuple.Create(currentClose, (double?)null);
                                                        _latestPriceMeta[ticker] = new Dictionary<string, object>
                                                        {
                                                                { "source", "Marketstack EOD" },
                                                                { "updated_at", DateTime.Now },
                                                                { "previous_close", previousClose }
                                                        };
                                                }

                                                Console.WriteLine($"[Marketstack更新] {ticker}: 當前=${currentClose.Value:F2}, 前收=${previousClose?.ToString("F2") ?? "null"}");
                                                return true;
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Marketstack更新失敗] {ticker}: {ex.Message}");
                        }

                        return false;
                }

                /// <summary>
                /// 從 Marketstack 數據數組中提取指定索引的值
                /// </summary>
                private double? ExtractMarketstackDataByIndex(string json, string fieldName, int index)
                {
                        try
                        {
                                // 找到 data 數組
                                var dataPattern = "\"data\":\\[([^\\]]+)\\]";
                                var dataMatch = Regex.Match(json, dataPattern);

                                if (dataMatch.Success)
                                {
                                        var dataContent = dataMatch.Groups[1].Value;

                                        // 分割成個別物件
                                        var objects = Regex.Split(dataContent, "\\},\\{");

                                        if (index < objects.Length)
                                        {
                                                var targetObject = objects[index];

                                                // 在目標物件中搜索字段
                                                var fieldPattern = $"\"{fieldName}\":([0-9.eE+-]+)";
                                                var fieldMatch = Regex.Match(targetObject, fieldPattern);

                                                if (fieldMatch.Success)
                                                {
                                                        var valueStr = fieldMatch.Groups[1].Value;
                                                        if (double.TryParse(valueStr, System.Globalization.NumberStyles.Any,
                                                                System.Globalization.CultureInfo.InvariantCulture, out double value))
                                                        {
                                                                return value;
                                                        }
                                                }
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[Marketstack] 解析索引 {index} 的 {fieldName} 失敗: {ex.Message}");
                        }

                        return null;
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
        }
}
