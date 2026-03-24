using SKCOMLib;
using StockManager.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace StockManager
{
    public partial class MainWindow
    {
        private readonly SKAPI m_api = SKAPI.Instance;
        private readonly object _skKLineLock = new object();
        private readonly Dictionary<string, List<MarketHistoryBar>> _skKLineBuffer = new Dictionary<string, List<MarketHistoryBar>>(StringComparer.OrdinalIgnoreCase);
        private string _skKLineActiveStockNo;
        private DateTime _skKLineLastEventAt = DateTime.MinValue;

        private Tuple<double?, double?, DateTime?> TryGetLatestTwSkQuote(string ticker)
        {
            if (!_isCapitalLoggedIn)
            {
                return null;
            }

            var key = NormalizeTwTickerForUi(ticker);
            if (string.IsNullOrWhiteSpace(key))
            {
                return null;
            }

            lock (_skTwQuoteLock)
            {
                Tuple<double?, double?, double?, DateTime> quote;
                if (!_twSkQuoteCache.TryGetValue(key, out quote))
                {
                    return null;
                }

                return Tuple.Create(quote.Item1, quote.Item3, (DateTime?)quote.Item4);
            }
        }

        public Tuple<double?, double?, DateTime?> GetLatestTwSkQuoteForTrend(string ticker)
        {
            return TryGetLatestTwSkQuote(ticker);
        }

        private void RegisterSkEventsIfNeeded()
        {
            if (_isSkEventsRegistered)
            {
                return;
            }

            m_api.OnReplyMessage += OnAnnouncement;
            void OnAnnouncement(string strUserID, string bstrMessage, out short nConfirmCode)
            {
                nConfirmCode = -1;
            }

            m_api.OnConnection += (nKind, code) =>
            {
                Console.WriteLine($"[OnConnection] nKind={nKind}, code={code}");

                if (nKind == 3003)
                {
                    _isSkQuoteConnectionReady = true;

                    if (!_hasInitialSkStocksRequested)
                    {
                        _hasInitialSkStocksRequested = true;
                        SubscribeTwStocksFromSk();
                    }
                }
            };

            m_api.OnNotifyQuoteLONG += (nMarketNo, nIndex) =>
            {
                var skStock = new SKSTOCKLONG();
                var code = m_api.SKQuoteLib_GetStockByIndexLONG(nMarketNo, nIndex, ref skStock);
                if (code == 0)
                {
                    var ticker = NormalizeTwTickerForUi(skStock.bstrStockNo);
                    var close = TryGetScaledSkPrice(skStock, skStock.nClose);
                    var prevClose = TryGetScaledSkPrice(skStock, (int)(TryGetSkNumericField(skStock, "nRef") ?? 0));
                    double? changePercent = null;
                    if (close.HasValue && prevClose.HasValue && Math.Abs(prevClose.Value) > 0.000001)
                    {
                        changePercent = (close.Value - prevClose.Value) / prevClose.Value * 100;
                    }

                    lock (_skTwQuoteLock)
                    {
                        _twSkQuoteCache[ticker] = Tuple.Create(close, prevClose, changePercent, DateTime.Now);
                        _lastSkQuoteReceivedAt = DateTime.Now;
                    }

                    Console.WriteLine($"[SK即時] {ticker} 成交={close?.ToString("F2") ?? "N/A"} 漲跌={changePercent?.ToString("F2") ?? "N/A"}%");
                }
            };

            m_api.OnNotifyKLineData += OnNotifyKLineData;
            void OnNotifyKLineData(string bstrStockNo, string bstrData)
            {
                MarketHistoryBar bar;
                if (!TryParseSkKLineBar(bstrData, out bar))
                {
                    return;
                }

                var stockNo = NormalizeTwTickerForSkRequest(bstrStockNo);
                if (string.IsNullOrWhiteSpace(stockNo))
                {
                    stockNo = _skKLineActiveStockNo;
                }

                if (string.IsNullOrWhiteSpace(stockNo))
                {
                    return;
                }

                lock (_skKLineLock)
                {
                    List<MarketHistoryBar> list;
                    if (!_skKLineBuffer.TryGetValue(stockNo, out list))
                    {
                        list = new List<MarketHistoryBar>();
                        _skKLineBuffer[stockNo] = list;
                    }

                    var existed = list.FirstOrDefault(x => x.DateText == bar.DateText);
                    if (existed == null)
                    {
                        list.Add(bar);
                    }
                    else
                    {
                        existed.Open = bar.Open;
                        existed.High = bar.High;
                        existed.Low = bar.Low;
                        existed.Close = bar.Close;
                        existed.Volume = bar.Volume;
                    }

                    _skKLineLastEventAt = DateTime.Now;
                }
            }

            _isSkEventsRegistered = true;
        }

        private List<MarketHistoryBar> TryGetTwSkKLineHistoryForHistoricalData(string ticker, string period, string interval)
        {
            if (!_isCapitalLoggedIn || !_isSkQuoteConnectionReady)
            {
                return null;
            }

            if (!string.Equals(interval, "1d", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var stockNo = NormalizeTwTickerForSkRequest(ticker);
            if (string.IsNullOrWhiteSpace(stockNo))
            {
                return null;
            }

            var endDate = DateTime.Today;
            var startDate = endDate.AddMonths(-3);

            var normalizedPeriod = (period ?? string.Empty).Trim().ToLowerInvariant();
            if (normalizedPeriod.EndsWith("mo"))
            {
                int months;
                if (int.TryParse(normalizedPeriod.Substring(0, normalizedPeriod.Length - 2), out months) && months > 0)
                {
                    startDate = endDate.AddMonths(-months);
                }
            }
            else if (normalizedPeriod.EndsWith("y"))
            {
                int years;
                if (int.TryParse(normalizedPeriod.Substring(0, normalizedPeriod.Length - 1), out years) && years > 0)
                {
                    startDate = endDate.AddYears(-years);
                }
            }

            lock (_skKLineLock)
            {
                _skKLineActiveStockNo = stockNo;
                _skKLineLastEventAt = DateTime.MinValue;
                _skKLineBuffer[stockNo] = new List<MarketHistoryBar>();
            }

	    var requestCode = m_api.SKQuoteLib_RequestKLineAMByDate(
                stockNo,
                4,
                1,
                0,
                startDate.ToString("yyyyMMdd"),
                endDate.ToString("yyyyMMdd"),
                1);

            if (requestCode != 0)
            {
                Console.WriteLine($"[SK KLine請求失敗] {stockNo}, code={requestCode}");
                return null;
            }

            var waitUntil = DateTime.Now.AddSeconds(4);
            while (DateTime.Now < waitUntil)
            {
                DateTime lastEventAt;
                int count;
                lock (_skKLineLock)
                {
                    lastEventAt = _skKLineLastEventAt;
                    List<MarketHistoryBar> list;
                    count = _skKLineBuffer.TryGetValue(stockNo, out list) && list != null ? list.Count : 0;
                }

                if (count > 0 && lastEventAt != DateTime.MinValue && (DateTime.Now - lastEventAt).TotalMilliseconds >= 350)
                {
                    break;
                }

                Thread.Sleep(80);
            }

            lock (_skKLineLock)
            {
                List<MarketHistoryBar> bars;
                if (!_skKLineBuffer.TryGetValue(stockNo, out bars) || bars == null || bars.Count == 0)
                {
                    return null;
                }

                return bars
                    .OrderBy(x => x.DateText)
                    .Select(x => new MarketHistoryBar
                    {
                        DateText = x.DateText,
                        Open = x.Open,
                        High = x.High,
                        Low = x.Low,
                        Close = x.Close,
                        Volume = x.Volume
                    })
                    .ToList();
            }
        }

        private bool TryParseSkKLineBar(string data, out MarketHistoryBar bar)
        {
            bar = null;
            if (string.IsNullOrWhiteSpace(data))
            {
                return false;
            }

            var values = data.Split(',');
            if (values.Length < 6)
            {
                return false;
            }

            DateTime date;
            if (!DateTime.TryParse(values[0].Trim(), out date))
            {
                var text = values[0].Trim();
                var formats = new[] { "yyyy/MM/dd", "yyyy-MM-dd", "yyyyMMdd" };
                if (!DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    return false;
                }
            }

            double open;
            double high;
            double low;
            double close;
            double volumeValue;
            if (!double.TryParse(values[1].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out open)
                || !double.TryParse(values[2].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out high)
                || !double.TryParse(values[3].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out low)
                || !double.TryParse(values[4].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out close)
                || !double.TryParse(values[5].Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out volumeValue))
            {
                return false;
            }

            bar = new MarketHistoryBar
            {
                DateText = date.ToString("yyyy-MM-dd"),
                Open = open,
                High = high,
                Low = low,
                Close = close,
                Volume = (long)Math.Max(0, Math.Round(volumeValue))
            };

            return true;
        }

        private void SubscribeTwStocksFromSk(bool forceResubscribe = false)
        {
            if (!_isCapitalLoggedIn || _twStockList == null)
            {
                return;
            }

            if (!_isSkQuoteConnectionReady)
            {
                Console.WriteLine("[SK訂閱] 尚未收到 OnConnection nKind=3003，暫不送出 RequestStocks。");
                return;
            }

            var targetStocks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var stock in _twStockList)
            {
                var stockNo = NormalizeTwTickerForSkRequest(stock.Ticker);
                if (!string.IsNullOrWhiteSpace(stockNo))
                {
                    targetStocks.Add(stockNo);
                }
            }

            var toCancel = _skSubscribedTwStocks
                .Where(x => !targetStocks.Contains(x))
                .ToList();

            if (toCancel.Count > 0)
            {
                var cancelArg = string.Join(",", toCancel);
                var cancelCode = m_api.SKQuoteLib_CancelRequestStocks(cancelArg);
                Console.WriteLine($"[SK取消訂閱] {cancelArg} => {cancelCode}");

                if (cancelCode == 0)
                {
                    foreach (var stockNo in toCancel)
                    {
                        _skSubscribedTwStocks.Remove(stockNo);

                        var uiTicker = NormalizeTwTickerForUi(stockNo);
                        lock (_skTwQuoteLock)
                        {
                            if (_twSkQuoteCache.ContainsKey(uiTicker))
                            {
                                _twSkQuoteCache.Remove(uiTicker);
                            }
                        }
                    }
                }
            }

            var toSubscribe = forceResubscribe
                ? targetStocks.ToList()
                : targetStocks.Where(x => !_skSubscribedTwStocks.Contains(x)).ToList();

            if (toSubscribe.Count > 0)
            {
                var subscribeArg = string.Join(",", toSubscribe);
                Thread.Sleep(1000);
                var code = m_api.SKQuoteLib_RequestStocks(1, subscribeArg);
                Console.WriteLine($"[{(forceResubscribe ? "SK重送訂閱" : "SK訂閱")}] {subscribeArg} => {code}");
                if (code == 0)
                {
                    foreach (var stockNo in toSubscribe)
                    {
                        _skSubscribedTwStocks.Add(stockNo);
                    }
                }
            }
        }

        private void EnsureSkSubscriptionAlive()
        {
            if (_isSkSectorBatchUpdating || !_isCapitalLoggedIn || _skSubscribedTwStocks.Count == 0)
            {
                return;
            }

            DateTime lastQuoteAt;
            lock (_skTwQuoteLock)
            {
                lastQuoteAt = _lastSkQuoteReceivedAt;
            }

            var now = DateTime.Now;
            var noQuoteSeconds = lastQuoteAt == DateTime.MinValue
                ? double.MaxValue
                : (now - lastQuoteAt).TotalSeconds;

            if (noQuoteSeconds < 45)
            {
                return;
            }

            if ((now - _lastSkResubscribeAt).TotalSeconds < 60)
            {
                return;
            }

            _lastSkResubscribeAt = now;
            Console.WriteLine($"[SK監控] 已 {noQuoteSeconds:F0} 秒未收到即時事件，重送訂閱...");
            SubscribeTwStocksFromSk(true);
        }

        private string NormalizeTwTickerForSkRequest(string ticker)
        {
            var t = (ticker ?? string.Empty).Trim().ToUpperInvariant();
            if (t.EndsWith(".TWO", StringComparison.OrdinalIgnoreCase))
            {
                t = t.Substring(0, t.Length - 4);
            }
            if (t.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
            {
                t = t.Substring(0, t.Length - 3);
            }
            return t;
        }

        private string NormalizeTwTickerForUi(string stockNo)
        {
            var t = (stockNo ?? string.Empty).Trim().ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(t))
            {
                return t;
            }

            if (t.EndsWith(".TWO", StringComparison.OrdinalIgnoreCase))
            {
                t = t.Substring(0, t.Length - 4);
            }

            return t.EndsWith(".TW", StringComparison.OrdinalIgnoreCase) ? t : (t + ".TW");
        }

        private double? TryGetScaledSkPrice(object skStock, int rawPrice)
        {
            if (rawPrice <= 0)
            {
                return null;
            }

            try
            {
                double decimals = TryGetSkNumericField(skStock, "sDecimal") ?? TryGetSkNumericField(skStock, "nDecimal") ?? 0;
                if (decimals > 0 && decimals <= 6)
                {
                    return rawPrice / Math.Pow(10, decimals);
                }
            }
            catch
            {
            }

            if (rawPrice > 10000)
            {
                return rawPrice / 100.0;
            }

            return rawPrice;
        }

        private double? TryGetSkNumericField(object target, string fieldName)
        {
            if (target == null || string.IsNullOrWhiteSpace(fieldName))
            {
                return null;
            }

            try
            {
                var type = target.GetType();
                var field = type.GetField(fieldName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (field != null)
                {
                    var value = field.GetValue(target);
                    if (value != null)
                    {
                        return Convert.ToDouble(value);
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private void ApplySkTwQuotesToStockList()
        {
            Dictionary<string, Tuple<double?, double?, double?, DateTime>> snapshot;
            lock (_skTwQuoteLock)
            {
                snapshot = new Dictionary<string, Tuple<double?, double?, double?, DateTime>>(_twSkQuoteCache, StringComparer.OrdinalIgnoreCase);
            }

            Console.WriteLine($"台股數據(SK): {snapshot.Count} 筆");
            foreach (var stock in _twStockList)
            {
                Tuple<double?, double?, double?, DateTime> quote;
                var quoteKey = NormalizeTwTickerForUi(stock.Ticker);
                if (!snapshot.TryGetValue(quoteKey, out quote))
                {
                    continue;
                }

                stock.Price = quote.Item1;
                stock.PreviousClose = quote.Item2;
                stock.ChangePercent = quote.Item3;
                stock.Source = "SK.OnNotifyQuoteLONG";
                stock.UpdatedAt = quote.Item4;
            }
        }
    }
}
