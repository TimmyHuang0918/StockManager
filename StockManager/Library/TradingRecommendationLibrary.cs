using System;
using System.Collections.Generic;
using System.Linq;

namespace StockManager.Library
{
    public class TrendRecommendationResult
    {
        public int Score { get; set; }
        public List<string> Reasons { get; set; }
    }

    public static class TradingRecommendationLibrary
    {
        public static TrendRecommendationResult CalculateAdvancedRecommendation(List<CandlestickData> data, double currentPrice, double? changePercent)
        {
            var reasons = new List<string>();
            if (data == null || data.Count < 2)
            {
                return new TrendRecommendationResult
                {
                    Score = 50,
                    Reasons = reasons
                };
            }

            var closes = data.Select(d => d.Close).ToList();
            var avgPrice = closes.Average();
            var latest = data[data.Count - 1];

            var score = 50;

            var ma5 = CalculateMA(closes, 5);
            var ma20 = CalculateMA(closes, 20);
            if (currentPrice > ma5 && ma5 > ma20)
            {
                score += 12;
                reasons.Add($"MA 多頭排列（現價 {currentPrice:F2} > MA5 {ma5:F2} > MA20 {ma20:F2}）");
            }
            else if (currentPrice < ma5 && ma5 < ma20)
            {
                score -= 12;
                reasons.Add($"MA 空頭排列（現價 {currentPrice:F2} < MA5 {ma5:F2} < MA20 {ma20:F2}）");
            }
            else
            {
                reasons.Add("MA 結構中性，趨勢尚未明確");
            }

            var macdTuple = BuildMACDComponents(closes);
            var macd = macdTuple.Item1[macdTuple.Item1.Count - 1];
            var signal = macdTuple.Item2[macdTuple.Item2.Count - 1];
            var hist = macdTuple.Item3[macdTuple.Item3.Count - 1];
            var prevHist = macdTuple.Item3.Count > 1 ? macdTuple.Item3[macdTuple.Item3.Count - 2] : hist;

            if (macd > signal)
            {
                score += 14;
                reasons.Add($"MACD 位於訊號線上方（MACD {macd:F3} > Signal {signal:F3}）");
            }
            else
            {
                score -= 14;
                reasons.Add($"MACD 位於訊號線下方（MACD {macd:F3} < Signal {signal:F3}）");
            }

            if (hist > prevHist)
            {
                score += 6;
                reasons.Add("MACD 柱狀體擴大，短線動能轉強");
            }
            else if (hist < prevHist)
            {
                score -= 6;
                reasons.Add("MACD 柱狀體縮小，短線動能轉弱");
            }

            var rsi = CalculateRSI(closes, 14);
            if (rsi < 30)
            {
                score += 10;
                reasons.Add($"RSI={rsi:F1} 處於超賣區，具反彈機會");
            }
            else if (rsi > 70)
            {
                score -= 10;
                reasons.Add($"RSI={rsi:F1} 處於超買區，短線回檔風險較高");
            }
            else if (rsi >= 45 && rsi <= 60)
            {
                score += 3;
                reasons.Add($"RSI={rsi:F1} 位於中性偏多區間");
            }
            else
            {
                reasons.Add($"RSI={rsi:F1}，未出現極端訊號");
            }

            var avgVol20 = data.Skip(Math.Max(0, data.Count - 20)).Average(x => (double)x.Volume);
            var volumeRatio = avgVol20 > 0 ? latest.Volume / avgVol20 : 1.0;
            if (volumeRatio >= 1.2)
            {
                if (macd > signal)
                {
                    score += 8;
                    reasons.Add($"成交量放大 {volumeRatio:F2} 倍，且多方訊號成立");
                }
                else
                {
                    score -= 8;
                    reasons.Add($"成交量放大 {volumeRatio:F2} 倍，但空方訊號較強");
                }
            }
            else
            {
                reasons.Add($"成交量為 20 日均量的 {volumeRatio:F2} 倍，量能一般");
            }

            if (currentPrice > avgPrice * 1.08)
            {
                score -= 4;
                reasons.Add("現價偏離均價較高，追價風險上升");
            }
            else if (currentPrice < avgPrice * 0.92)
            {
                score += 4;
                reasons.Add("現價低於均價，具均值回歸空間");
            }

            if (changePercent.HasValue)
            {
                if (changePercent.Value >= 3) score += 2;
                else if (changePercent.Value <= -3) score -= 2;
            }

            score = Math.Max(0, Math.Min(100, score));

            return new TrendRecommendationResult
            {
                Score = score,
                Reasons = reasons
            };
        }

        public static int CalculateSimpleScore(double? price, double? previousClose, double? changePercent)
        {
            var score = 50;

            if (changePercent.HasValue)
            {
                if (changePercent.Value >= 4) score += 20;
                else if (changePercent.Value >= 2) score += 12;
                else if (changePercent.Value > 0) score += 5;
                else if (changePercent.Value <= -4) score -= 20;
                else if (changePercent.Value <= -2) score -= 12;
                else if (changePercent.Value < 0) score -= 5;
            }
            else
            {
                score -= 5;
            }

            if (price.HasValue && previousClose.HasValue)
            {
                var gap = price.Value - previousClose.Value;
                if (gap > 0) score += 3;
                else if (gap < 0) score -= 3;
            }

            return Math.Max(0, Math.Min(100, score));
        }

        public static string GetSimpleSuggestion(int score)
        {
            if (score >= 70) return "偏多（買入）";
            if (score >= 50) return "中性（觀望）";
            return "偏空（賣出）";
        }

        public static List<Tuple<int, string, string>> BuildBacktestSignals(List<CandlestickData> data)
        {
            var signals = new List<Tuple<int, string, string>>();
            if (data == null || data.Count < 30)
            {
                return signals;
            }

            var lastSignal = string.Empty;

            for (int i = 20; i < data.Count; i++)
            {
                var slice = data.Take(i + 1).ToList();
                var closes = slice.Select(x => x.Close).ToList();
                var current = slice[slice.Count - 1];

                var ma5 = CalculateMA(closes, 5);
                var ma20 = CalculateMA(closes, 20);
                var rsi = CalculateRSI(closes, 14);
                var macd = BuildMACDComponents(closes);
                var macdNow = macd.Item1[macd.Item1.Count - 1];
                var signalNow = macd.Item2[macd.Item2.Count - 1];
                var avgVol20 = slice.Skip(Math.Max(0, slice.Count - 20)).Average(x => (double)x.Volume);
                var volRatio = avgVol20 > 0 ? current.Volume / avgVol20 : 1.0;

                var buy = macdNow > signalNow && current.Close > ma5 && ma5 > ma20 && rsi < 70;
                var sell = macdNow < signalNow && current.Close < ma5 && ma5 < ma20 && rsi > 30;

                var strengthScore = 0;
                strengthScore += macdNow > signalNow ? 1 : -1;
                strengthScore += ma5 > ma20 ? 1 : -1;
                if (rsi >= 40 && rsi <= 60) strengthScore += 1;
                if (volRatio >= 1.2) strengthScore += 1;

                if (buy && lastSignal != "BUY")
                {
                    var action = strengthScore >= 3 ? "STRONG_BUY" : "BUY";
                    var levelText = action == "STRONG_BUY" ? "強買" : "買入";
                    var reason = $"{levelText}｜MACD多頭 + MA5>MA20 + RSI={rsi:F1} + 量比{volRatio:F2}";
                    signals.Add(Tuple.Create(i, action, reason));
                    lastSignal = "BUY";
                }
                else if (sell && lastSignal != "SELL")
                {
                    var action = strengthScore <= -3 ? "STRONG_SELL" : "SELL";
                    var levelText = action == "STRONG_SELL" ? "強賣" : "賣出";
                    var reason = $"{levelText}｜MACD空頭 + MA5<MA20 + RSI={rsi:F1} + 量比{volRatio:F2}";
                    signals.Add(Tuple.Create(i, action, reason));
                    lastSignal = "SELL";
                }
            }

            return signals;
        }

        private static double CalculateMA(List<double> prices, int period)
        {
            if (prices.Count < period) return prices[prices.Count - 1];
            return prices.Skip(prices.Count - period).Take(period).Average();
        }

        private static double CalculateRSI(List<double> prices, int period)
        {
            if (prices.Count < period + 1) return 50;

            var gains = new List<double>();
            var losses = new List<double>();

            for (int i = prices.Count - period; i < prices.Count; i++)
            {
                var change = prices[i] - prices[i - 1];
                gains.Add(change > 0 ? change : 0);
                losses.Add(change < 0 ? -change : 0);
            }

            var avgGain = gains.Average();
            var avgLoss = losses.Average();

            if (avgLoss == 0) return 100;
            var rs = avgGain / avgLoss;
            return 100 - (100 / (1 + rs));
        }

        private static Tuple<List<double>, List<double>, List<double>> BuildMACDComponents(List<double> closes)
        {
            var macdSeries = new List<double>();
            var signalSeries = new List<double>();
            var histSeries = new List<double>();

            if (closes == null || closes.Count == 0)
            {
                return Tuple.Create(macdSeries, signalSeries, histSeries);
            }

            var macdLineForSignal = new List<double>();
            for (int i = 0; i < closes.Count; i++)
            {
                var slice = closes.Take(i + 1).ToList();
                if (slice.Count < 26)
                {
                    macdSeries.Add(0);
                    signalSeries.Add(0);
                    histSeries.Add(0);
                    macdLineForSignal.Add(0);
                    continue;
                }

                var ema12 = CalculateEMA(slice, 12);
                var ema26 = CalculateEMA(slice, 26);
                var macd = ema12 - ema26;
                macdSeries.Add(macd);
                macdLineForSignal.Add(macd);

                var signal = CalculateEMA(macdLineForSignal, 9);
                signalSeries.Add(signal);
                histSeries.Add(macd - signal);
            }

            return Tuple.Create(macdSeries, signalSeries, histSeries);
        }

        private static double CalculateEMA(List<double> prices, int period)
        {
            if (prices.Count < period) return prices[prices.Count - 1];

            var multiplier = 2.0 / (period + 1);
            var ema = prices.Take(period).Average();

            for (int i = period; i < prices.Count; i++)
            {
                ema = (prices[i] - ema) * multiplier + ema;
            }

            return ema;
        }
    }
}
