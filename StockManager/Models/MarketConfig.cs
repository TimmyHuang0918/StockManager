using System.Collections.Generic;

namespace StockManager.Models
{
        public class MarketConfig
        {
                public string MarketName { get; set; }
                public Dictionary<string, string> DefaultStocks { get; set; }
                public string ConfigFilePath { get; set; }

                public MarketConfig(string marketName, Dictionary<string, string> defaultStocks, string configFilePath)
                {
                        MarketName = marketName;
                        DefaultStocks = defaultStocks;
                        ConfigFilePath = configFilePath;
                }
        }
}
