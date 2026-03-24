using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace StockManager.Services
{
    public class CompositeMarketDataGateway : IMarketDataGateway
    {
        private readonly IMarketDataGateway _fallback;
        private readonly Func<string, Tuple<double?, double?, DateTime?>> _twSkQuoteProvider;
        private readonly Func<bool> _isSkEnabledProvider;
        private readonly Func<string, string, string, List<MarketHistoryBar>> _twSkKLineProvider;

        public CompositeMarketDataGateway(
            IMarketDataGateway fallback,
            Func<string, Tuple<double?, double?, DateTime?>> twSkQuoteProvider = null,
            Func<bool> isSkEnabledProvider = null,
            Func<string, string, string, List<MarketHistoryBar>> twSkKLineProvider = null)
        {
            _fallback = fallback ?? throw new ArgumentNullException(nameof(fallback));
            _twSkQuoteProvider = twSkQuoteProvider;
            _isSkEnabledProvider = isSkEnabledProvider;
            _twSkKLineProvider = twSkKLineProvider;
        }

        public Tuple<double?, double?, string> GetRealtimePriceWithSource(string ticker)
        {
            var skQuote = GetRealtimeTwQuote(ticker);
            if (IsTwTicker(ticker) && skQuote != null && skQuote.Item1.HasValue)
            {
                return Tuple.Create(skQuote.Item1, skQuote.Item2, "SK.OnNotifyQuoteLONG");
            }

            return _fallback.GetRealtimePriceWithSource(ticker);
        }

        public void UpdatePrice(string ticker)
        {
            _fallback.UpdatePrice(ticker);
        }

        public void UpdatePriceWithPreviousClose(string ticker)
        {
            _fallback.UpdatePriceWithPreviousClose(ticker);
        }

        public Dictionary<string, Tuple<double?, double?>> GetPrices()
        {
            return _fallback.GetPrices();
        }

        public Dictionary<string, Dictionary<string, object>> GetPriceMeta()
        {
            return _fallback.GetPriceMeta();
        }

        public Tuple<double?, double?> GetCachedPrice(string ticker)
        {
            return _fallback.GetCachedPrice(ticker);
        }

        public Tuple<double?, double?, DateTime?> GetRealtimeTwQuote(string ticker)
        {
            if (!IsTwTicker(ticker) || _twSkQuoteProvider == null)
            {
                return null;
            }

            if (_isSkEnabledProvider != null && !_isSkEnabledProvider())
            {
                return null;
            }

            try
            {
                return _twSkQuoteProvider(NormalizeTwTicker(ticker));
            }
            catch
            {
                return null;
            }
        }

        public bool TryGetHistoricalData(string ticker, string period, string interval, out List<MarketHistoryBar> bars)
        {
            bars = null;

            if (IsTwTicker(ticker)
                && _twSkKLineProvider != null
                && (_isSkEnabledProvider == null || _isSkEnabledProvider()))
            {
                try
                {
                    var skBars = _twSkKLineProvider(NormalizeTwTicker(ticker), period, interval);
                    if (skBars != null && skBars.Count > 0)
                    {
                        bars = skBars;
                        return true;
                    }
                }
                catch
                {
                }
            }

            return _fallback.TryGetHistoricalData(ticker, period, interval, out bars);
        }

        public bool TryGetFundamentals(string ticker, out Dictionary<string, double> metrics)
        {
            return _fallback.TryGetFundamentals(ticker, out metrics);
        }

        private bool IsTwTicker(string ticker)
        {
            if (string.IsNullOrWhiteSpace(ticker))
            {
                return false;
            }

            var normalized = ticker.Trim().ToUpperInvariant();
            return normalized.EndsWith(".TW", StringComparison.OrdinalIgnoreCase)
                || normalized.EndsWith(".TWO", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(normalized, "^\\d{4}[A-Z]?$");
        }

        private string NormalizeTwTicker(string ticker)
        {
            var normalized = (ticker ?? string.Empty).Trim().ToUpperInvariant();
            if (normalized.EndsWith(".TWO", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(0, normalized.Length - 4);
            }

            if (!normalized.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
            {
                normalized += ".TW";
            }

            return normalized;
        }
    }
}
