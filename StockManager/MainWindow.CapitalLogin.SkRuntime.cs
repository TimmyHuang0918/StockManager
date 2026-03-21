using SKDLLCSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace StockManager
{
    public partial class MainWindow
    {
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

            SK.OnReplyMessage += (strLoginID, strMessage) =>
            {
                Console.WriteLine($"[OnReplyMessage] {strLoginID}: {strMessage}");
            };

            SK.OnConnection += (loginID, code) =>
            {
                Console.WriteLine($"[OnConnection] {loginID}, Code={code}");
            };

            SK.OnNotifyQuoteLONG += (nMarketNo, strStockNo) =>
            {
                var pSKStockLONG = SK.SKQuoteLib_GetStockByStockNo(nMarketNo, strStockNo);
                if (pSKStockLONG.nCode == 0)
                {
                    var ticker = NormalizeTwTickerForUi(strStockNo);
                    var close = TryGetScaledSkPrice(pSKStockLONG, pSKStockLONG.nClose);
                    var prevClose = TryGetScaledSkPrice(pSKStockLONG, (int)(TryGetSkNumericField(pSKStockLONG, "nRef") ?? 0));
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

            _isSkEventsRegistered = true;
        }

        private void SubscribeTwStocksFromSk(bool forceResubscribe = false)
        {
            if (!_isCapitalLoggedIn || _twStockList == null)
            {
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
                var cancelCode = SK.SKQuoteLib_CancelRequestStocks(cancelArg);
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
                System.Threading.Thread.Sleep(1000);
                var code = SK.SKQuoteLib_RequestStocks(subscribeArg);
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
