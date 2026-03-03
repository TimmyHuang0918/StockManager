using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace StockManager.Services
{
        public class StockManagerService
        {
                private readonly string _configFile;
                private readonly object _lock = new object();
                private Dictionary<string, string> _stocks;
                private string _viewedStock;

                public StockManagerService(Dictionary<string, string> initialStocks, string configFile = null)
                {
                        _configFile = configFile;

                        if (!string.IsNullOrEmpty(configFile) && File.Exists(configFile))
                        {
                                try
                                {
                                        _stocks = LoadFromFile(configFile);
                                        Console.WriteLine($"已載入配置檔案 {_stocks.Count} 檔股票");
                                }
                                catch (Exception ex)
                                {
                                        Console.WriteLine($"載入配置檔案錯誤: {ex.Message}，使用預設股票清單");
                                        _stocks = new Dictionary<string, string>(initialStocks);
                                }
                        }
                        else
                        {
                                _stocks = new Dictionary<string, string>(initialStocks);
                                if (!string.IsNullOrEmpty(configFile))
                                {
                                        SaveToFile();
                                }
                        }
                }

                private Dictionary<string, string> LoadFromFile(string filePath)
                {
                        var result = new Dictionary<string, string>();
                        var json = File.ReadAllText(filePath);

                        json = json.Trim().Trim('{', '}');
                        var pairs = json.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        foreach (var pair in pairs)
                        {
                                var parts = pair.Split(new[] { ':' }, 2);
                                if (parts.Length == 2)
                                {
                                        var key = parts[0].Trim().Trim('"');
                                        var value = parts[1].Trim().Trim('"');
                                        result[key] = value;
                                }
                        }

                        return result;
                }

                public void SaveToFile()
                {
                        if (string.IsNullOrEmpty(_configFile))
                                return;

                        try
                        {
                                var configDir = Path.GetDirectoryName(_configFile);
                                if (!string.IsNullOrEmpty(configDir) && !Directory.Exists(configDir))
                                {
                                        Directory.CreateDirectory(configDir);
                                }

                                lock (_lock)
                                {
                                        var sb = new StringBuilder();
                                        sb.AppendLine("{");

                                        var count = 0;
                                        foreach (var kvp in _stocks)
                                        {
                                                count++;
                                                sb.Append($"  \"{kvp.Key}\": \"{kvp.Value}\"");
                                                if (count < _stocks.Count)
                                                        sb.AppendLine(",");
                                                else
                                                        sb.AppendLine();
                                        }

                                        sb.AppendLine("}");

                                        File.WriteAllText(_configFile, sb.ToString(), Encoding.UTF8);
                                }
                                Console.WriteLine($"已保存股票清單至 {_configFile}");
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"保存配置檔案錯誤: {ex.Message}");
                        }
                }

                public Dictionary<string, string> GetStocks()
                {
                        lock (_lock)
                        {
                                return new Dictionary<string, string>(_stocks);
                        }
                }

                public List<string> GetTickers()
                {
                        lock (_lock)
                        {
                                return _stocks.Keys.ToList();
                        }
                }

                public void AddStock(string ticker, string name)
                {
                        lock (_lock)
                        {
                                _stocks[ticker.ToUpper()] = name;
                        }
                        SaveToFile();
                }

                public void RemoveStock(string ticker)
                {
                        lock (_lock)
                        {
                                if (_stocks.ContainsKey(ticker))
                                {
                                        _stocks.Remove(ticker);
                                }
                        }
                        SaveToFile();
                }

                public void SetViewedStock(string ticker)
                {
                        _viewedStock = ticker;
                }

                public string GetViewedStock()
                {
                        return _viewedStock;
                }
        }
}
