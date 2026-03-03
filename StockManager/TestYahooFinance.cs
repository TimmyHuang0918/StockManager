using System;
using System.IO;
using System.Net;

namespace StockManager
{
        /// <summary>
        /// 測試 Yahoo Finance API 連接
        /// </summary>
        public static class TestYahooFinance
        {
                public static void TestConnection(string ticker = "NVDA")
                {
                        Console.WriteLine("========================================");
                        Console.WriteLine($"🌐 測試 Yahoo Finance API 連接 - {ticker}");
                        Console.WriteLine("========================================\n");

                        try
                        {
                                // 測試 V8 API
                                Console.WriteLine("📡 測試 V8 API...");
                                var url = $"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?interval=1d&range=1d";
                                Console.WriteLine($"URL: {url}\n");

                                var request = (HttpWebRequest)WebRequest.Create(url);
                                request.Timeout = 10000;
                                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";
                                request.Accept = "application/json";

                                using (var response = request.GetResponse())
                                {
                                        Console.WriteLine($"✅ 連接成功！");
                                        Console.WriteLine($"狀態碼: {((HttpWebResponse)response).StatusCode}");
                                        Console.WriteLine($"內容類型: {response.ContentType}");
                                        Console.WriteLine($"內容長度: {response.ContentLength} bytes\n");

                                        using (var stream = response.GetResponseStream())
                                        using (var reader = new StreamReader(stream))
                                        {
                                                var json = reader.ReadToEnd();
                                                
                                                Console.WriteLine($"📄 回應內容長度: {json.Length} 字符");
                                                Console.WriteLine($"\n📋 前 500 字符預覽:");
                                                Console.WriteLine("----------------------------------------");
                                                Console.WriteLine(json.Substring(0, Math.Min(500, json.Length)));
                                                Console.WriteLine("----------------------------------------\n");

                                                // 簡單檢查是否包含關鍵字段
                                                var hasPrice = json.Contains("regularMarketPrice");
                                                var hasPreviousClose = json.Contains("regularMarketPreviousClose");
                                                var hasPreMarket = json.Contains("preMarketPrice");
                                                var hasPostMarket = json.Contains("postMarketPrice");

                                                Console.WriteLine("🔍 關鍵字段檢查:");
                                                Console.WriteLine($"  regularMarketPrice: {(hasPrice ? "✅ 存在" : "❌ 不存在")}");
                                                Console.WriteLine($"  regularMarketPreviousClose: {(hasPreviousClose ? "✅ 存在" : "❌ 不存在")}");
                                                Console.WriteLine($"  preMarketPrice: {(hasPreMarket ? "✅ 存在" : "❌ 不存在")}");
                                                Console.WriteLine($"  postMarketPrice: {(hasPostMarket ? "✅ 存在" : "❌ 不存在")}");
                                        }
                                }

                                Console.WriteLine("\n========================================");
                                Console.WriteLine("✅ Yahoo Finance API 可以正常訪問！");
                                Console.WriteLine("========================================");
                        }
                        catch (WebException ex)
                        {
                                Console.WriteLine($"\n❌ 網路錯誤: {ex.Message}");
                                
                                if (ex.Response != null)
                                {
                                        var response = (HttpWebResponse)ex.Response;
                                        Console.WriteLine($"狀態碼: {response.StatusCode}");
                                        Console.WriteLine($"狀態描述: {response.StatusDescription}");
                                        
                                        using (var stream = response.GetResponseStream())
                                        using (var reader = new StreamReader(stream))
                                        {
                                                var errorContent = reader.ReadToEnd();
                                                Console.WriteLine($"\n錯誤回應內容:");
                                                Console.WriteLine(errorContent);
                                        }
                                }

                                Console.WriteLine("\n可能的原因:");
                                Console.WriteLine("1. 網路連接問題");
                                Console.WriteLine("2. Yahoo Finance API 暫時無法訪問");
                                Console.WriteLine("3. 防火牆阻擋");
                                Console.WriteLine("4. 請求頻率過高");
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"\n❌ 錯誤: {ex.Message}");
                                Console.WriteLine($"類型: {ex.GetType().Name}");
                                Console.WriteLine($"\n堆疊追蹤:");
                                Console.WriteLine(ex.StackTrace);
                        }
                }

                public static void TestMultipleStocks()
                {
                        Console.WriteLine("\n========================================");
                        Console.WriteLine("🧪 測試多個股票");
                        Console.WriteLine("========================================\n");

                        var stocks = new[] { "AAPL", "MSFT", "GOOGL", "NVDA", "TSLA" };

                        foreach (var ticker in stocks)
                        {
                                Console.WriteLine($"\n📊 測試 {ticker}...");
                                
                                try
                                {
                                        var url = $"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?interval=1d&range=1d";
                                        var request = (HttpWebRequest)WebRequest.Create(url);
                                        request.Timeout = 5000;
                                        request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36";

                                        using (var response = request.GetResponse())
                                        {
                                                Console.WriteLine($"  ✅ {ticker} - 連接成功");
                                        }
                                }
                                catch (Exception ex)
                                {
                                        Console.WriteLine($"  ❌ {ticker} - 失敗: {ex.Message}");
                                }

                                System.Threading.Thread.Sleep(500); // 間隔 500ms
                        }

                        Console.WriteLine("\n========================================");
                        Console.WriteLine("測試完成");
                        Console.WriteLine("========================================");
                }
        }
}
