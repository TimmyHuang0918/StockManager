using System;
using StockManager.Services;

namespace StockManager
{
        /// <summary>
        /// NVDA 診斷測試工具
        /// 在 MainWindow 中調用 TestNVDA.Run() 來診斷問題
        /// </summary>
        public static class TestNVDA
        {
                public static void Run()
                {
                        Console.WriteLine("========================================");
                        Console.WriteLine("🔍 NVDA 診斷測試開始");
                        Console.WriteLine("========================================");
                        
                        var fetcher = new PriceFetcherService();
                        var result = fetcher.GetRealtimePriceWithSource("NVDA");
                        
                        var price = result.Item1;
                        var changePercent = result.Item2;
                        var source = result.Item3;
                        
                        Console.WriteLine("");
                        Console.WriteLine("📊 最終結果:");
                        Console.WriteLine($"  當前價格: ${price?.ToString("F2") ?? "null"}");
                        Console.WriteLine($"  漲跌幅: {changePercent?.ToString("F2") ?? "null"}%");
                        Console.WriteLine($"  資料來源: {source ?? "null"}");
                        Console.WriteLine("");
                        
                        // 驗證：如果前收盤是 $195.56
                        if (price.HasValue)
                        {
                                double expectedPreviousClose = 195.56;
                                double actualChange = price.Value - expectedPreviousClose;
                                double actualChangePercent = (actualChange / expectedPreviousClose) * 100;
                                
                                Console.WriteLine("✅ 驗證計算（基於前收盤 $195.56）:");
                                Console.WriteLine($"  當前價格: ${price.Value:F2}");
                                Console.WriteLine($"  前收盤價: ${expectedPreviousClose:F2}");
                                Console.WriteLine($"  漲跌金額: ${actualChange:F2}");
                                Console.WriteLine($"  應該顯示漲跌幅: {actualChangePercent:F2}%");
                                Console.WriteLine("");
                                
                                if (changePercent.HasValue)
                                {
                                        double difference = Math.Abs(changePercent.Value - actualChangePercent);
                                        
                                        if (difference < 0.1) // 允許 0.1% 誤差
                                        {
                                                Console.WriteLine("✅ 漲跌幅計算正確！");
                                        }
                                        else
                                        {
                                                Console.WriteLine("❌ 漲跌幅計算有誤！");
                                                Console.WriteLine($"  系統顯示: {changePercent.Value:F2}%");
                                                Console.WriteLine($"  應該顯示: {actualChangePercent:F2}%");
                                                Console.WriteLine($"  差異: {difference:F2}%");
                                                
                                                // 反推前收盤價
                                                if (changePercent.Value != 0)
                                                {
                                                        double impliedPreviousClose = price.Value / (1 + changePercent.Value / 100);
                                                        Console.WriteLine($"  系統使用的前收盤價: ${impliedPreviousClose:F2}");
                                                        Console.WriteLine($"  正確的前收盤價: ${expectedPreviousClose:F2}");
                                                }
                                        }
                                }
                                else
                                {
                                        Console.WriteLine("❌ 系統未返回漲跌幅！");
                                }
                        }
                        
                        Console.WriteLine("");
                        Console.WriteLine("========================================");
                        Console.WriteLine("🔍 NVDA 診斷測試結束");
                        Console.WriteLine("========================================");
                }
        }
}
