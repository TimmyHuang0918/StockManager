using System;
using System.Collections.Generic;

namespace StockManager.Services
{
    public class MarketHistoryBar
    {
        public string DateText { get; set; }
        public double Open { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Close { get; set; }
        public long Volume { get; set; }
    }

    public interface IMarketDataGateway
    {
        Tuple<double?, double?, string> GetRealtimePriceWithSource(string ticker);
        void UpdatePrice(string ticker);
        void UpdatePriceWithPreviousClose(string ticker);
        Dictionary<string, Tuple<double?, double?>> GetPrices();
        Dictionary<string, Dictionary<string, object>> GetPriceMeta();
        Tuple<double?, double?> GetCachedPrice(string ticker);
        Tuple<double?, double?, DateTime?> GetRealtimeTwQuote(string ticker);
        bool TryGetHistoricalData(string ticker, string period, string interval, out List<MarketHistoryBar> bars);
        bool TryGetFundamentals(string ticker, out Dictionary<string, double> metrics);
    }
}
