using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using StockManager;

namespace StockManager.Library
{
    public class KLineRenderCandle
    {
        public DateTime Date { get; set; }
        public double Open { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Close { get; set; }
        public long Volume { get; set; }
        public string ToolTipText { get; set; }
    }

    public class KLineRenderResult
    {
        public double ChartLeft { get; set; }
        public double ChartRight { get; set; }
        public double ChartTop { get; set; }
        public double ChartBottom { get; set; }
        public double PriceTop { get; set; }
        public double PriceBottom { get; set; }
        public double MinPrice { get; set; }
        public double MaxPrice { get; set; }
        public double PriceRange { get; set; }
        public double Spacing { get; set; }
        public double BodyWidth { get; set; }
        public List<KLineRenderCandle> DisplayCandles { get; set; }

        public double MapPriceToY(double price)
        {
            return PriceBottom - ((price - MinPrice) / PriceRange) * (PriceBottom - PriceTop);
        }
    }

    public class KLineChartTheme
    {
        public Brush TextBrush { get; set; }
        public Brush AxisBrush { get; set; }
        public Brush GridBrush { get; set; }
        public Brush UpFillBrush { get; set; }
        public Brush UpStrokeBrush { get; set; }
        public Brush DownFillBrush { get; set; }
        public Brush DownStrokeBrush { get; set; }
        public Brush Ma5Brush { get; set; }
        public Brush Ma20Brush { get; set; }

        public static KLineChartTheme Dark => new KLineChartTheme
        {
            TextBrush = Brushes.White,
            AxisBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            GridBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
            UpFillBrush = new SolidColorBrush(Color.FromRgb(102, 187, 106)),
            UpStrokeBrush = new SolidColorBrush(Color.FromRgb(56, 142, 60)),
            DownFillBrush = new SolidColorBrush(Color.FromRgb(239, 83, 80)),
            DownStrokeBrush = new SolidColorBrush(Color.FromRgb(211, 47, 47)),
            Ma5Brush = new SolidColorBrush(Color.FromRgb(255, 193, 7)),
            Ma20Brush = new SolidColorBrush(Color.FromRgb(103, 58, 183))
        };

        public static KLineChartTheme Dim => new KLineChartTheme
        {
            TextBrush = Brushes.DimGray,
            AxisBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            GridBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
            UpFillBrush = new SolidColorBrush(Color.FromRgb(76, 175, 80)),
            UpStrokeBrush = new SolidColorBrush(Color.FromRgb(56, 142, 60)),
            DownFillBrush = new SolidColorBrush(Color.FromRgb(244, 67, 54)),
            DownStrokeBrush = new SolidColorBrush(Color.FromRgb(211, 47, 47)),
            Ma5Brush = new SolidColorBrush(Color.FromRgb(255, 193, 7)),
            Ma20Brush = new SolidColorBrush(Color.FromRgb(103, 58, 183))
        };
    }

    public class KLineChartRenderOptions
    {
        public bool EnableBacktestOverlay { get; set; } = true;
        public bool EnableMacdOverlay { get; set; } = true;
	public bool EnableRsiOverlay { get; set; } = true;
    }

    public class TrendChartModuleFlags
    {
        public bool ShowPriceChart { get; set; } = true;
        public bool ShowVolumeChart { get; set; } = true;
        public bool ShowMacdChart { get; set; } = true;
        public bool ShowRsiChart { get; set; } = true;
        public bool ShowBacktestOverlay { get; set; } = true;

        public static TrendChartModuleFlags All => new TrendChartModuleFlags();
    }

    public static class KLineChartLibrary
    {
        public static void RenderTrendModule(
            Canvas priceCanvas,
            Canvas volumeCanvas,
            Canvas macdCanvas,
            Canvas rsiCanvas,
            IList<KLineRenderCandle> candles,
            string title,
            KLineChartTheme theme,
            int maxBars,
            TrendChartModuleFlags flags = null)
        {
            flags = flags ?? TrendChartModuleFlags.All;
            var recent = (candles ?? new List<KLineRenderCandle>())
                .Skip(Math.Max(0, Math.Max(0, (candles ?? new List<KLineRenderCandle>()).Count) - Math.Max(1, maxBars)))
                .ToList();

            if (flags.ShowPriceChart && priceCanvas != null)
            {
                DrawCandles(
                    priceCanvas,
                    recent,
                    title,
                    theme,
                    maxBars,
                    new KLineChartRenderOptions
                    {
                        EnableBacktestOverlay = flags.ShowBacktestOverlay,
                        EnableMacdOverlay = false,
                        EnableRsiOverlay = false
                    });
            }
            else if (priceCanvas != null)
            {
                priceCanvas.Children.Clear();
            }

            if (flags.ShowVolumeChart && volumeCanvas != null)
            {
                DrawVolumePanel(volumeCanvas, recent, theme);
            }
            else if (volumeCanvas != null)
            {
                volumeCanvas.Children.Clear();
            }

            if (flags.ShowMacdChart && macdCanvas != null)
            {
                DrawMacdPanel(macdCanvas, recent, theme);
            }
            else if (macdCanvas != null)
            {
                macdCanvas.Children.Clear();
            }

            if (flags.ShowRsiChart && rsiCanvas != null)
            {
                DrawRsiPanel(rsiCanvas, recent, theme);
            }
            else if (rsiCanvas != null)
            {
                rsiCanvas.Children.Clear();
            }
        }

        public static KLineRenderResult DrawCandles(
            Canvas canvas,
            IList<KLineRenderCandle> candles,
            string title,
            KLineChartTheme theme,
            int maxBars,
            KLineChartRenderOptions options = null)
        {
            canvas.Children.Clear();
            ToolTipService.SetInitialShowDelay(canvas, 120);
            ToolTipService.SetShowDuration(canvas, 60000);
            if (candles == null || candles.Count == 0)
            {
                return null;
            }

            theme = theme ?? KLineChartTheme.Dark;
            options = options ?? new KLineChartRenderOptions();
            var width = canvas.ActualWidth;
            if (width < 320)
            {
                width = 820;
            }

            var height = canvas.Height > 0 ? canvas.Height : 250;
            var chartLeft = 56.0;
            var chartRight = width - 16.0;
            var priceTop = 26.0;
            var chartBottomPadding = 20.0;
            var panelGap = 6.0;

            var targetVolumeHeight = 42.0;
            var targetIndicatorHeight = 44.0;
            var minPanelHeight = 26.0;

            var panelCount = 1 + (options.EnableMacdOverlay ? 1 : 0) + (options.EnableRsiOverlay ? 1 : 0);
            var totalGapHeight = panelGap * panelCount;

            var availableHeight = Math.Max(120, height - priceTop - chartBottomPadding);
            var minPriceAreaHeight = 70.0;
            var maxPanelArea = Math.Max(40, availableHeight - minPriceAreaHeight);

            var requestedPanelHeight = targetVolumeHeight +
                                       (options.EnableMacdOverlay ? targetIndicatorHeight : 0) +
                                       (options.EnableRsiOverlay ? targetIndicatorHeight : 0);
            var scale = requestedPanelHeight > 0 ? Math.Min(1.0, maxPanelArea / requestedPanelHeight) : 1.0;

            var volumeHeight = Math.Max(minPanelHeight, targetVolumeHeight * scale);
            var indicatorPanelHeight = Math.Max(minPanelHeight, targetIndicatorHeight * scale);

            var currentBottom = height - chartBottomPadding;
            var rsiBottom = currentBottom;
            var rsiTop = currentBottom;
            if (options.EnableRsiOverlay)
            {
                rsiBottom = currentBottom;
                rsiTop = rsiBottom - indicatorPanelHeight;
                currentBottom = rsiTop - panelGap;
            }

            var macdBottom = currentBottom;
            var macdTop = currentBottom;
            if (options.EnableMacdOverlay)
            {
                macdBottom = currentBottom;
                macdTop = macdBottom - indicatorPanelHeight;
                currentBottom = macdTop - panelGap;
            }

            var volumeBottom = currentBottom;
            var volumeTop = volumeBottom - volumeHeight;
            currentBottom = volumeTop - panelGap;

            var priceBottom = Math.Max(priceTop + minPriceAreaHeight, currentBottom);

            var recent = candles.Skip(Math.Max(0, candles.Count - Math.Max(1, maxBars))).ToList();
            var maxPrice = recent.Max(x => x.High);
            var minPrice = recent.Min(x => x.Low);
            var range = Math.Max(0.01, maxPrice - minPrice);
            var spacing = Math.Max(4.0, (chartRight - chartLeft) / Math.Max(1, recent.Count));
            var bodyWidth = Math.Max(2.0, spacing * 0.6);
            var maxVolume = Math.Max(1L, recent.Max(x => x.Volume));

            Func<double, double> mapPriceToY = price =>
                priceBottom - ((price - minPrice) / range) * (priceBottom - priceTop);

            for (int i = 0; i <= 4; i++)
            {
                var ratio = i / 4.0;
                var y = priceTop + ratio * (priceBottom - priceTop);
                var priceLabel = maxPrice - ratio * range;

                canvas.Children.Add(new Line
                {
                    X1 = chartLeft,
                    Y1 = y,
                    X2 = chartRight,
                    Y2 = y,
                    Stroke = theme.GridBrush,
                    StrokeThickness = 1
                });

                var yText = new TextBlock
                {
                    Text = priceLabel.ToString("F2"),
                    FontSize = 10,
                    Foreground = theme.TextBrush
                };
                Canvas.SetLeft(yText, 4);
                Canvas.SetTop(yText, y - 8);
                canvas.Children.Add(yText);
            }

            canvas.Children.Add(new Line { X1 = chartLeft, Y1 = priceTop, X2 = chartLeft, Y2 = priceBottom, Stroke = theme.AxisBrush, StrokeThickness = 1 });
            canvas.Children.Add(new Line { X1 = chartLeft, Y1 = priceBottom, X2 = chartRight, Y2 = priceBottom, Stroke = theme.AxisBrush, StrokeThickness = 1 });
            canvas.Children.Add(new Line { X1 = chartLeft, Y1 = volumeTop, X2 = chartLeft, Y2 = volumeBottom, Stroke = theme.AxisBrush, StrokeThickness = 1 });
            canvas.Children.Add(new Line { X1 = chartLeft, Y1 = volumeBottom, X2 = chartRight, Y2 = volumeBottom, Stroke = theme.AxisBrush, StrokeThickness = 1 });

            var xTickCount = Math.Min(6, recent.Count);
            var hasIntraday = recent.Select(x => x.Date.Date).Distinct().Count() <= 1;
            if (xTickCount > 1)
            {
                for (int i = 0; i < xTickCount; i++)
                {
                    var idx = (int)Math.Round(i * (recent.Count - 1) / (double)(xTickCount - 1));
                    var x = chartLeft + idx * spacing + bodyWidth / 2;

                    canvas.Children.Add(new Line
                    {
                        X1 = x,
                        Y1 = priceBottom,
                        X2 = x,
                        Y2 = priceBottom + 4,
                        Stroke = theme.AxisBrush,
                        StrokeThickness = 1
                    });

                    var dateText = new TextBlock
                    {
                        Text = hasIntraday ? recent[idx].Date.ToString("HH:mm") : recent[idx].Date.ToString("MM/dd"),
                        FontSize = 9,
                        Foreground = theme.TextBrush
                    };
                    Canvas.SetLeft(dateText, x - 16);
                    Canvas.SetTop(dateText, priceBottom + 4);
                    canvas.Children.Add(dateText);
                }
            }

            var volumeText = new TextBlock
            {
                Text = "成交量",
                FontSize = 10,
                Foreground = theme.TextBrush
            };
            Canvas.SetLeft(volumeText, 4);
            Canvas.SetTop(volumeText, volumeTop - 2);
            canvas.Children.Add(volumeText);

            var titleBlock = new TextBlock
            {
                Text = $"K線（近{recent.Count}日） {title}",
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Foreground = theme.TextBrush
            };
            Canvas.SetLeft(titleBlock, chartLeft);
            Canvas.SetTop(titleBlock, 0);
            canvas.Children.Add(titleBlock);

            for (int i = 0; i < recent.Count; i++)
            {
                var c = recent[i];
                var x = chartLeft + i * spacing;
                var openY = mapPriceToY(c.Open);
                var closeY = mapPriceToY(c.Close);
                var highY = mapPriceToY(c.High);
                var lowY = mapPriceToY(c.Low);
                var tooltipText = string.IsNullOrWhiteSpace(c.ToolTipText)
                    ? $"時間: {c.Date:yyyy-MM-dd HH:mm}\n開: {c.Open:F2}\n高: {c.High:F2}\n低: {c.Low:F2}\n收: {c.Close:F2}\n量: {c.Volume:N0}"
                    : c.ToolTipText;

                var isUp = c.Close >= c.Open;
                var stroke = isUp ? theme.UpStrokeBrush : theme.DownStrokeBrush;
                var fill = isUp ? theme.UpFillBrush : theme.DownFillBrush;

                var wick = new Line
                {
                    X1 = x + bodyWidth / 2,
                    Y1 = highY,
                    X2 = x + bodyWidth / 2,
                    Y2 = lowY,
                    Stroke = stroke,
                    StrokeThickness = 1,
                    IsHitTestVisible = true
                };
                ToolTipService.SetToolTip(wick, CreateChartToolTip(tooltipText));
                canvas.Children.Add(wick);

                var body = new Rectangle
                {
                    Width = bodyWidth,
                    Height = Math.Max(1, Math.Abs(closeY - openY)),
                    Fill = fill,
                    Stroke = stroke,
                    StrokeThickness = 1,
                    IsHitTestVisible = true
                };
                ToolTipService.SetToolTip(body, CreateChartToolTip(tooltipText));
                Canvas.SetLeft(body, x);
                Canvas.SetTop(body, Math.Min(openY, closeY));
                canvas.Children.Add(body);

                var volumeBarHeight = (c.Volume / (double)maxVolume) * (volumeBottom - volumeTop);
                var volumeRect = new Rectangle
                {
                    Width = bodyWidth,
                    Height = Math.Max(1, volumeBarHeight),
                    Fill = fill,
                    Opacity = 0.45,
                    IsHitTestVisible = true
                };
                ToolTipService.SetToolTip(volumeRect, CreateChartToolTip(tooltipText));
                Canvas.SetLeft(volumeRect, x);
                Canvas.SetTop(volumeRect, volumeBottom - volumeRect.Height);
                canvas.Children.Add(volumeRect);

                var hitArea = new Rectangle
                {
                    Width = Math.Max(spacing, bodyWidth + 2),
                    Height = volumeBottom - priceTop,
                    Fill = new SolidColorBrush(Color.FromArgb(1, 255, 255, 255)),
                    IsHitTestVisible = true
                };
                ToolTipService.SetToolTip(hitArea, CreateChartToolTip(tooltipText));
                Canvas.SetLeft(hitArea, x - (Math.Max(spacing, bodyWidth + 2) - bodyWidth) / 2);
                Canvas.SetTop(hitArea, priceTop);
                canvas.Children.Add(hitArea);
            }

            var closeSeries = recent.Select(x => x.Close).ToList();
            var ma5 = BuildMASeries(closeSeries, 5);
            var ma20 = BuildMASeries(closeSeries, 20);

            DrawMALine(canvas, ma5, chartLeft, spacing, bodyWidth, mapPriceToY, theme.Ma5Brush);
            DrawMALine(canvas, ma20, chartLeft, spacing, bodyWidth, mapPriceToY, theme.Ma20Brush);

            var legend = new TextBlock
            {
                Text = "MA5(黃)  MA20(紫)",
                FontSize = 10,
                Foreground = theme.TextBrush
            };
            Canvas.SetLeft(legend, chartRight - 100);
            Canvas.SetTop(legend, 4);
            canvas.Children.Add(legend);

            if (options.EnableBacktestOverlay)
            {
                DrawBacktestOverlay(canvas, recent, chartLeft, priceTop, priceBottom, spacing, bodyWidth, minPrice, range, theme);
            }

            var hasMacd = options.EnableMacdOverlay;
            var hasRsi = options.EnableRsiOverlay;

            if (options.EnableMacdOverlay)
            {
                DrawMacdOverlay(canvas, recent, chartLeft, chartRight, macdTop, macdBottom, spacing, theme);
            }

            if (options.EnableRsiOverlay)
            {
                DrawRsiOverlay(canvas, recent, chartLeft, chartRight, rsiTop, rsiBottom, spacing, theme);
            }

            return new KLineRenderResult
            {
                ChartLeft = chartLeft,
                ChartRight = chartRight,
                ChartTop = priceTop,
                ChartBottom = volumeBottom,
                PriceTop = priceTop,
                PriceBottom = priceBottom,
                MinPrice = minPrice,
                MaxPrice = maxPrice,
                PriceRange = range,
                Spacing = spacing,
                BodyWidth = bodyWidth,
                DisplayCandles = recent
            };
        }

        private static List<double?> BuildMASeries(List<double> prices, int period)
        {
            var result = new List<double?>();
            for (int i = 0; i < prices.Count; i++)
            {
                if (i < period - 1)
                {
                    result.Add(null);
                    continue;
                }

                result.Add(prices.Skip(i - period + 1).Take(period).Average());
            }

            return result;
        }

        private static void DrawMALine(
            Canvas canvas,
            List<double?> ma,
            double chartLeft,
            double spacing,
            double bodyWidth,
            Func<double, double> mapPriceToY,
            Brush brush)
        {
            var polyline = new Polyline
            {
                Stroke = brush,
                StrokeThickness = 1.5
            };

            for (int i = 0; i < ma.Count; i++)
            {
                if (!ma[i].HasValue)
                {
                    continue;
                }

                var x = chartLeft + i * spacing + bodyWidth / 2;
                var y = mapPriceToY(ma[i].Value);
                polyline.Points.Add(new Point(x, y));
            }

            if (polyline.Points.Count > 1)
            {
                canvas.Children.Add(polyline);
            }
        }

        private static void DrawBacktestOverlay(
            Canvas canvas,
            List<KLineRenderCandle> candles,
            double chartLeft,
            double chartTop,
            double chartBottom,
            double spacing,
            double bodyWidth,
            double minPrice,
            double range,
            KLineChartTheme theme)
        {
            var data = candles.Select(x => new CandlestickData
            {
                Date = x.Date.ToString("yyyy-MM-dd"),
                Open = x.Open,
                High = x.High,
                Low = x.Low,
                Close = x.Close,
                Volume = x.Volume,
                ChangeAmount = x.Close - x.Open,
                ChangePercent = Math.Abs(x.Open) > 0.000001 ? (x.Close - x.Open) / x.Open * 100 : 0
            }).ToList();

            var signals = TradingRecommendationLibrary.BuildBacktestSignals(data);
            foreach (var signal in signals)
            {
                if (signal.Item1 < 0 || signal.Item1 >= candles.Count)
                {
                    continue;
                }

                var candle = candles[signal.Item1];
                var x = chartLeft + signal.Item1 * spacing + bodyWidth / 2;
                var highY = chartBottom - ((candle.High - minPrice) / range) * (chartBottom - chartTop);
                var lowY = chartBottom - ((candle.Low - minPrice) / range) * (chartBottom - chartTop);
                var isBuy = signal.Item2.Contains("BUY");

                var marker = new TextBlock
                {
                    Text = isBuy ? "B" : "S",
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    Foreground = isBuy ? theme.UpStrokeBrush : theme.DownStrokeBrush,
                    Background = Brushes.White,
                    Padding = new Thickness(2, 0, 2, 0)
                };

                var signalReason = signal.Item3 ?? string.Empty;
                var markerTooltip =
                    $"訊號: {signal.Item2}\n" +
                    $"日期: {candle.Date:yyyy-MM-dd HH:mm}\n" +
                    $"開: {candle.Open:F2} 高: {candle.High:F2} 低: {candle.Low:F2} 收: {candle.Close:F2}\n" +
                    $"量: {candle.Volume:N0}\n" +
                    $"原因: {signalReason}";
                ToolTipService.SetToolTip(marker, CreateChartToolTip(markerTooltip));

                Canvas.SetLeft(marker, x - 6);
                Canvas.SetTop(marker, isBuy ? lowY + 3 : highY - 16);
                canvas.Children.Add(marker);
            }
        }

        private static void DrawMacdOverlay(
            Canvas canvas,
            List<KLineRenderCandle> candles,
            double chartLeft,
            double chartRight,
            double overlayTop,
            double overlayBottom,
            double spacing,
            KLineChartTheme theme)
        {
            var closes = candles.Select(x => x.Close).ToList();
            var macdTuple = BuildMacdComponents(closes);
            var macd = macdTuple.Item1;
            var signal = macdTuple.Item2;
            if (macd.Count < 2 || signal.Count < 2)
            {
                return;
            }

            var all = macd.Concat(signal).ToList();
            var min = all.Min();
            var max = all.Max();
            if (Math.Abs(max - min) < 0.000001)
            {
                max += 1;
                min -= 1;
            }

            Func<double, double> mapY = v => overlayBottom - ((v - min) / (max - min)) * (overlayBottom - overlayTop);

            // MACD 子圖座標軸
            canvas.Children.Add(new Line
            {
                X1 = chartLeft,
                Y1 = overlayTop,
                X2 = chartLeft,
                Y2 = overlayBottom,
                Stroke = theme.AxisBrush,
                StrokeThickness = 1
            });
            canvas.Children.Add(new Line
            {
                X1 = chartLeft,
                Y1 = overlayBottom,
                X2 = chartRight,
                Y2 = overlayBottom,
                Stroke = theme.AxisBrush,
                StrokeThickness = 1
            });

            var zeroY = mapY(0);
            canvas.Children.Add(new Line
            {
                X1 = chartLeft,
                Y1 = zeroY,
                X2 = chartRight,
                Y2 = zeroY,
                Stroke = theme.GridBrush,
                StrokeThickness = 1,
                StrokeDashArray = new DoubleCollection { 3, 2 }
            });

            var maxLabel = new TextBlock { Text = max.ToString("F2"), FontSize = 8, Foreground = theme.TextBrush };
            Canvas.SetLeft(maxLabel, 2);
            Canvas.SetTop(maxLabel, overlayTop - 2);
            canvas.Children.Add(maxLabel);

            var minLabel = new TextBlock { Text = min.ToString("F2"), FontSize = 8, Foreground = theme.TextBrush };
            Canvas.SetLeft(minLabel, 2);
            Canvas.SetTop(minLabel, overlayBottom - 10);
            canvas.Children.Add(minLabel);

            var macdLine = new Polyline
            {
                Stroke = new SolidColorBrush(Color.FromRgb(33, 150, 243)),
                StrokeThickness = 1,
                Opacity = 0.8
            };

            var signalLine = new Polyline
            {
                Stroke = new SolidColorBrush(Color.FromRgb(255, 152, 0)),
                StrokeThickness = 1,
                Opacity = 0.8
            };

            var candleCenterOffset = Math.Max(2.0, spacing * 0.6) / 2.0;

            for (int i = 0; i < macd.Count; i++)
            {
                var x = chartLeft + i * spacing + candleCenterOffset;
                macdLine.Points.Add(new Point(x, mapY(macd[i])));
                signalLine.Points.Add(new Point(x, mapY(signal[i])));
            }

            canvas.Children.Add(macdLine);
            canvas.Children.Add(signalLine);

            var label = new TextBlock
            {
                Text = "MACD（Y軸） / 時間（X軸）",
                FontSize = 9,
                Foreground = theme.TextBrush
            };
            Canvas.SetLeft(label, chartLeft + 2);
            Canvas.SetTop(label, Math.Max(0, overlayTop - 12));
            canvas.Children.Add(label);
        }

        private static void DrawRsiOverlay(
            Canvas canvas,
            List<KLineRenderCandle> candles,
            double chartLeft,
            double chartRight,
            double overlayTop,
            double overlayBottom,
            double spacing,
            KLineChartTheme theme)
        {
            var closes = candles.Select(x => x.Close).ToList();
            var rsiSeries = BuildRsiSeries(closes, 14);
            if (rsiSeries.Count < 2)
            {
                return;
            }

            Func<double, double> mapY = v => overlayBottom - (Math.Max(0, Math.Min(100, v)) / 100.0) * (overlayBottom - overlayTop);

            // RSI 子圖座標軸
            canvas.Children.Add(new Line
            {
                X1 = chartLeft,
                Y1 = overlayTop,
                X2 = chartLeft,
                Y2 = overlayBottom,
                Stroke = theme.AxisBrush,
                StrokeThickness = 1
            });
            canvas.Children.Add(new Line
            {
                X1 = chartLeft,
                Y1 = overlayBottom,
                X2 = chartRight,
                Y2 = overlayBottom,
                Stroke = theme.AxisBrush,
                StrokeThickness = 1
            });

            var rsi70Y = mapY(70);
            var rsi50Y = mapY(50);
            var rsi30Y = mapY(30);

            canvas.Children.Add(new Line { X1 = chartLeft, Y1 = rsi70Y, X2 = chartRight, Y2 = rsi70Y, Stroke = theme.GridBrush, StrokeThickness = 1, StrokeDashArray = new DoubleCollection { 3, 2 } });
            canvas.Children.Add(new Line { X1 = chartLeft, Y1 = rsi50Y, X2 = chartRight, Y2 = rsi50Y, Stroke = theme.GridBrush, StrokeThickness = 1, StrokeDashArray = new DoubleCollection { 2, 2 } });
            canvas.Children.Add(new Line { X1 = chartLeft, Y1 = rsi30Y, X2 = chartRight, Y2 = rsi30Y, Stroke = theme.GridBrush, StrokeThickness = 1, StrokeDashArray = new DoubleCollection { 3, 2 } });

            var label70 = new TextBlock { Text = "70", FontSize = 8, Foreground = theme.TextBrush };
            Canvas.SetLeft(label70, 2);
            Canvas.SetTop(label70, rsi70Y - 8);
            canvas.Children.Add(label70);

            var label50 = new TextBlock { Text = "50", FontSize = 8, Foreground = theme.TextBrush };
            Canvas.SetLeft(label50, 2);
            Canvas.SetTop(label50, rsi50Y - 8);
            canvas.Children.Add(label50);

            var label30 = new TextBlock { Text = "30", FontSize = 8, Foreground = theme.TextBrush };
            Canvas.SetLeft(label30, 2);
            Canvas.SetTop(label30, rsi30Y - 8);
            canvas.Children.Add(label30);

            var rsiLine = new Polyline
            {
                Stroke = new SolidColorBrush(Color.FromRgb(156, 39, 176)),
                StrokeThickness = 1,
                Opacity = 0.8
            };

            var candleCenterOffset = Math.Max(2.0, spacing * 0.6) / 2.0;

            for (int i = 0; i < rsiSeries.Count; i++)
            {
                var x = chartLeft + i * spacing + candleCenterOffset;
                rsiLine.Points.Add(new Point(x, mapY(rsiSeries[i])));
            }

            canvas.Children.Add(rsiLine);

            var label = new TextBlock
            {
                Text = "RSI（Y軸） / 時間（X軸）",
                FontSize = 9,
                Foreground = theme.TextBrush
            };
            Canvas.SetLeft(label, chartLeft + 2);
            Canvas.SetTop(label, Math.Max(0, overlayTop - 12));
            canvas.Children.Add(label);
        }

        private static Tuple<List<double>, List<double>, List<double>> BuildMacdComponents(List<double> closes)
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

                var ema12 = CalculateEma(slice, 12);
                var ema26 = CalculateEma(slice, 26);
                var macd = ema12 - ema26;
                macdSeries.Add(macd);
                macdLineForSignal.Add(macd);

                var signal = CalculateEma(macdLineForSignal, 9);
                signalSeries.Add(signal);
                histSeries.Add(macd - signal);
            }

            return Tuple.Create(macdSeries, signalSeries, histSeries);
        }

        private static List<double> BuildRsiSeries(List<double> closes, int period)
        {
            var result = new List<double>();
            if (closes == null || closes.Count == 0)
            {
                return result;
            }

            for (int i = 0; i < closes.Count; i++)
            {
                var slice = closes.Take(i + 1).ToList();
                if (slice.Count < period + 1)
                {
                    result.Add(50);
                    continue;
                }

                result.Add(CalculateRsi(slice, period));
            }

            return result;
        }

        private static double CalculateRsi(List<double> prices, int period)
        {
            if (prices.Count < period + 1)
            {
                return 50;
            }

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
            if (avgLoss == 0)
            {
                return 100;
            }

            var rs = avgGain / avgLoss;
            return 100 - (100 / (1 + rs));
        }

        private static double CalculateEma(List<double> prices, int period)
        {
            if (prices == null || prices.Count == 0)
            {
                return 0;
            }

            if (prices.Count < period)
            {
                return prices[prices.Count - 1];
            }

            var multiplier = 2.0 / (period + 1);
            var ema = prices.Take(period).Average();
            for (int i = period; i < prices.Count; i++)
            {
                ema = (prices[i] - ema) * multiplier + ema;
            }

            return ema;
        }

        private static ToolTip CreateChartToolTip(string text)
        {
            return new ToolTip
            {
                Background = new SolidColorBrush(Color.FromRgb(245, 245, 245)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(176, 190, 197)),
                BorderThickness = new Thickness(1),
                Content = new TextBlock
                {
                    Text = text,
                    Foreground = Brushes.Black
                }
            };
        }

        private static void DrawVolumePanel(Canvas canvas, List<KLineRenderCandle> candles, KLineChartTheme theme)
        {
            canvas.Children.Clear();
            if (candles == null || candles.Count == 0)
            {
                return;
            }

            var width = canvas.ActualWidth > 0 ? canvas.ActualWidth : 800;
            var height = canvas.ActualHeight > 0 ? canvas.ActualHeight : 110;
            var left = 56.0;
            var right = width - 16.0;
            var top = 10.0;
            var bottom = height - 22;

            canvas.Children.Add(new Line { X1 = left, Y1 = top, X2 = left, Y2 = bottom, Stroke = theme.AxisBrush, StrokeThickness = 1 });
            canvas.Children.Add(new Line { X1 = left, Y1 = bottom, X2 = right, Y2 = bottom, Stroke = theme.AxisBrush, StrokeThickness = 1 });

            var maxVolume = Math.Max(1L, candles.Max(d => d.Volume));
            var spacing = Math.Max(1.0, (right - left) / Math.Max(1, candles.Count));
            var barWidth = Math.Max(1, spacing * 0.6);

            for (int i = 0; i < candles.Count; i++)
            {
                var c = candles[i];
                var barHeight = (c.Volume / (double)maxVolume) * (bottom - top);
                var rect = new Rectangle
                {
                    Width = barWidth,
                    Height = Math.Max(1, barHeight),
                    Fill = c.Close >= c.Open ? theme.UpFillBrush : theme.DownFillBrush,
                    Opacity = 0.85
                };
                Canvas.SetLeft(rect, left + i * spacing);
                Canvas.SetTop(rect, bottom - rect.Height);
                canvas.Children.Add(rect);
            }
        }

        private static void DrawMacdPanel(Canvas canvas, List<KLineRenderCandle> candles, KLineChartTheme theme)
        {
            canvas.Children.Clear();
            if (candles == null || candles.Count < 2)
            {
                return;
            }

            var width = canvas.ActualWidth > 0 ? canvas.ActualWidth : 800;
            var left = 56.0;
            var right = width - 16.0;
            var spacing = Math.Max(1.0, (right - left) / Math.Max(1, candles.Count));
            DrawMacdOverlay(canvas, candles, left, right, 10.0, (canvas.ActualHeight > 0 ? canvas.ActualHeight : 110) - 22.0, spacing, theme);
        }

        private static void DrawRsiPanel(Canvas canvas, List<KLineRenderCandle> candles, KLineChartTheme theme)
        {
            canvas.Children.Clear();
            if (candles == null || candles.Count < 2)
            {
                return;
            }

            var width = canvas.ActualWidth > 0 ? canvas.ActualWidth : 800;
            var left = 56.0;
            var right = width - 16.0;
            var spacing = Math.Max(1.0, (right - left) / Math.Max(1, candles.Count));
            DrawRsiOverlay(canvas, candles, left, right, 10.0, (canvas.ActualHeight > 0 ? canvas.ActualHeight : 110) - 22.0, spacing, theme);
        }
    }
}
