using System;
using System.Collections.Generic;
using System.IO;

namespace StockManager.Config
{
        public static class AppConfig
        {
                public static readonly string UserConfigDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), 
                        ".stock_monitor");

                public static readonly string UserStocksFile = Path.Combine(UserConfigDir, "stocks.json");
                public static readonly string UserTwStocksFile = Path.Combine(UserConfigDir, "tw_stocks.json");

                public static readonly Dictionary<string, string> DefaultStocks = new Dictionary<string, string>
                {
                        { "AAPL", "蘋果" },
                        { "MSFT", "微軟" },
                        { "GOOG", "Alphabet" },
                        { "AMZN", "亞馬遜" },
                        { "TSLA", "特斯拉" },
                        { "MU", "美光" },
                        { "GRAB", "Grab" },
                        { "NVDA", "輝達" },
                        { "AMD", "超微" },
                        { "INTC", "英特爾" },
                        { "TSM", "台積電ADR" },
                        { "BABA", "阿里巴巴" },
                        { "BTC-USD", "比特幣" }
                };

                public static readonly Dictionary<string, string> DefaultTwStocks = new Dictionary<string, string>
                {
                        { "2330.TW", "台積電" },
                        { "2317.TW", "鴻海" },
                        { "2454.TW", "聯發科" },
                        { "2308.TW", "台達電" },
                        { "2382.TW", "廣達" },
                        { "2412.TW", "中華電" },
                        { "2881.TW", "富邦金" },
                        { "2882.TW", "國泰金" },
                        { "2886.TW", "兆豐金" },
                        { "2891.TW", "中信金" },
                        { "2303.TW", "聯電" },
                        { "3711.TW", "日月光投控" },
                        { "2002.TW", "中鋼" },
                        { "1301.TW", "台塑" }
                };

                public static readonly int UpdateInterval = 10;

                // 🔑 Marketstack API 配置
                // 免費註冊: https://marketstack.com/signup/free
                // 免費計劃限制: 1000 次請求/月
                public static readonly string MarketstackApiKey = "YOUR_API_KEY_HERE"; // ⚠️ 請替換為您的 API Key
                public static readonly bool UseMarketstack = false; // ❌ 已禁用（僅使用 yfinance）
                public static readonly string MarketstackBaseUrl = "http://api.marketstack.com/v1";

                // 🔑 Intrinio API 配置
                // 免費註冊: https://intrinio.com/signup
                // 免費計劃限制: 500 次請求/天
                public static readonly string IntrinioApiKey = "YOUR_API_KEY_HERE"; // ⚠️ 請替換為您的 API Key
                public static readonly bool UseIntrinio = false; // ❌ 已禁用（僅使用 yfinance）
                public static readonly string IntrinioBaseUrl = "https://api-v2.intrinio.com";

                // 🐍 Python yfinance 配置
                public static readonly bool UsePythonYFinance = true; // ✅ 已啟用（主要數據源）
                public static readonly string PythonPath = ResolvePythonPath();
                public static readonly string YFinanceScriptPath = "Python\\yfinance_fetcher.py"; // yfinance 腳本路徑

                private static string ResolvePythonPath()
                {
                        // 最高優先：環境變數覆寫（方便不同機器直接切換）
                        var envPython = Environment.GetEnvironmentVariable("STOCKMANAGER_PYTHON");
                        if (!string.IsNullOrWhiteSpace(envPython))
                        {
                                var candidate = envPython.Trim().Trim('"');

                                // 可接受完整路徑，或直接給命令名稱（例如 python / py）
                                if (string.Equals(candidate, "python", StringComparison.OrdinalIgnoreCase) ||
                                        string.Equals(candidate, "py", StringComparison.OrdinalIgnoreCase) ||
                                        File.Exists(candidate))
                                {
                                        return candidate;
                                }
                        }

                        var baseDir = AppDomain.CurrentDomain.BaseDirectory;

                        // 安裝包內建 Python venv
                        var bundledVenvPython = Path.Combine(baseDir, "PythonRuntime", "Scripts", "python.exe");
                        if (File.Exists(bundledVenvPython))
                        {
                                return bundledVenvPython;
                        }

                        // 其他可能的相對路徑（開發環境）
                        var localVenvPython = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "PythonRuntime", "Scripts", "python.exe"));
                        if (File.Exists(localVenvPython))
                        {
                                return localVenvPython;
                        }

                        // 回退到系統 Python
                        return "python";
                }

                public static bool TryLoadTaiwanSectorCsv(out List<string> sectorOrder, out Dictionary<string, string> tickerToSector)
                {
                        return TaiwanSectorConfigData.TryLoad(out sectorOrder, out tickerToSector);
                }
        }
}
