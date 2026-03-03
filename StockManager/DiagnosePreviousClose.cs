using System;
using System.IO;
using System.Net;

namespace StockManager
{
        /// <summary>
        /// 診斷工具：檢查為什麼 previousClose 是 null
        /// </summary>
        public static class DiagnosePreviousClose
        {
                public static void Run(string ticker = "NVDA")
                {
                        Console.WriteLine("========================================");
                        Console.WriteLine($"🔍 診斷 {ticker} 的前收盤價問題");
                        Console.WriteLine("========================================\n");

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

                                        Console.WriteLine("✅ API 連接成功！\n");
                                        Console.WriteLine($"JSON 長度: {json.Length} 字符\n");

                                        // 1. 檢查是否包含 previousClose 關鍵字
                                        Console.WriteLine("🔍 步驟 1: 搜索 'previousClose' 關鍵字");
                                        Console.WriteLine("----------------------------------------");

                                        if (json.Contains("previousClose"))
                                        {
                                                Console.WriteLine("✅ JSON 中包含 'previousClose'");

                                                // 找出所有包含 previousClose 的片段
                                                var index = 0;
                                                var count = 0;
                                                while ((index = json.IndexOf("previousClose", index)) != -1)
                                                {
                                                        count++;
                                                        var start = Math.Max(0, index - 50);
                                                        var length = Math.Min(150, json.Length - start);
                                                        var snippet = json.Substring(start, length);

                                                        Console.WriteLine($"\n匹配 {count}:");
                                                        Console.WriteLine($"位置: {index}");
                                                        Console.WriteLine($"前後文:");
                                                        Console.WriteLine(snippet);
                                                        Console.WriteLine();

                                                        index++;
                                                }

                                                Console.WriteLine($"總共找到 {count} 個 'previousClose'");
                                        }
                                        else
                                        {
                                                Console.WriteLine("❌ JSON 中不包含 'previousClose'");
                                                Console.WriteLine("   可能的原因:");
                                                Console.WriteLine("   1. Yahoo Finance API 改變了字段名稱");
                                                Console.WriteLine("   2. 這個股票沒有前收盤價數據");
                                                Console.WriteLine("   3. API 版本不同");
                                        }

                                        Console.WriteLine("\n");
                                        Console.WriteLine("🔍 步驟 2: 搜索其他可能的字段");
                                        Console.WriteLine("----------------------------------------");

                                        var possibleFields = new[]
                                        {
                                                "regularMarketPreviousClose",
                                                "previousClose",
                                                "prevClose",
                                                "chartPreviousClose",
                                                "close"
                                        };

                                        foreach (var field in possibleFields)
                                        {
                                                if (json.Contains(field))
                                                {
                                                        Console.WriteLine($"✅ 找到字段: {field}");

                                                        // 顯示該字段的值
                                                        var index = json.IndexOf($"\"{field}\"");
                                                        if (index != -1)
                                                        {
                                                                var start = index;
                                                                var end = Math.Min(index + 200, json.Length);
                                                                var snippet = json.Substring(start, end - start);

                                                                Console.WriteLine($"   內容: {snippet.Substring(0, Math.Min(150, snippet.Length))}");
                                                        }
                                                }
                                                else
                                                {
                                                        Console.WriteLine($"❌ 未找到字段: {field}");
                                                }
                                        }

                                        Console.WriteLine("\n");
                                        Console.WriteLine("🔍 步驟 3: 查看完整的 meta 區域");
                                        Console.WriteLine("----------------------------------------");

                                        if (json.Contains("\"meta\""))
                                        {
                                                var metaIndex = json.IndexOf("\"meta\"");
                                                var metaStart = metaIndex;
                                                var metaEnd = json.IndexOf("}", metaIndex + 1000); // 查找接下來的 1000 字符內的結束括號

                                                if (metaEnd != -1)
                                                {
                                                        metaEnd = Math.Min(metaEnd + 500, json.Length); // 再往後取 500 字符
                                                        var metaContent = json.Substring(metaStart, metaEnd - metaStart);

                                                        Console.WriteLine("Meta 區域內容:");
                                                        Console.WriteLine(metaContent.Substring(0, Math.Min(1500, metaContent.Length)));
                                                }
                                        }
                                        else
                                        {
                                                Console.WriteLine("❌ JSON 中不包含 'meta' 區域");
                                        }

                                        Console.WriteLine("\n");
                                        Console.WriteLine("🔍 步驟 4: 保存完整 JSON 到文件");
                                        Console.WriteLine("----------------------------------------");

                                        try
                                        {
                                                var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                                                var filename = Path.Combine(desktop, $"{ticker}_response.json");
                                                File.WriteAllText(filename, json);
                                                Console.WriteLine($"✅ 已保存到: {filename}");
                                                Console.WriteLine($"   請打開此文件查看完整的 JSON 回應");
                                        }
                                        catch (Exception ex)
                                        {
                                                Console.WriteLine($"❌ 無法保存文件: {ex.Message}");
                                        }

                                        Console.WriteLine("\n========================================");
                                        Console.WriteLine("診斷完成");
                                        Console.WriteLine("========================================");
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"❌ 錯誤: {ex.Message}");
                                Console.WriteLine($"堆疊追蹤:\n{ex.StackTrace}");
                        }
                }
        }
}
