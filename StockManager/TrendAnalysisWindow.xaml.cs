using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Xml.Linq;
using StockManager.Config;
using StockManager.Services;

namespace StockManager
{
        public partial class TrendAnalysisWindow : Window
        {
                private string _ticker;
                private string _stockName;
                private PriceFetcherService _priceFetcher;
                private string _currentPeriod = "3mo";
                private bool _isLoaded = false; // 標記視窗是否已完全載入
                private List<CandlestickData> _historicalData = new List<CandlestickData>();
                private bool _suppressStockSwitchEvent = false;
                private Line _crosshairVertical;
                private Line _crosshairHorizontal;
                private double _chartAreaLeft;
                private double _chartAreaRight;
                private double _chartAreaTop;
                private double _chartAreaBottom;
                private LoadingProgressPopup _loadingWindow;

                public TrendAnalysisWindow(string ticker, string stockName, PriceFetcherService priceFetcher)
                {
                        try
                        {
                                Console.WriteLine($"[趨勢視窗] 開始初始化: {ticker} - {stockName}");

                                InitializeComponent();

                                Console.WriteLine($"[趨勢視窗] InitializeComponent 完成");

                                _ticker = ticker;
                                _stockName = stockName;
                                _priceFetcher = priceFetcher;

                                Console.WriteLine($"[趨勢視窗] 設定參數完成");

                                // 等待視窗載入完成後再繪製圖表
                                this.Loaded += TrendAnalysisWindow_Loaded;

                                Console.WriteLine($"[趨勢視窗] 建構函數完成");
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[趨勢視窗錯誤] 建構函數發生錯誤: {ex.Message}");
                                Console.WriteLine($"[趨勢視窗錯誤] 堆疊追蹤: {ex.StackTrace}");

                                MessageBox.Show(
                                        $"初始化趨勢視窗時發生錯誤:\n\n{ex.Message}\n\n內部異常: {ex.InnerException?.Message}",
                                        "初始化錯誤",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Error
                                );
                                throw;
                        }
                }

                private List<CandlestickData> GetDisplayData(List<CandlestickData> data)
                {
                        if (data == null || data.Count == 0)
                        {
                                return new List<CandlestickData>();
                        }

                        // 依照當前時間範圍顯示完整天數（不再限制最多 50 筆）
                        return data.ToList();
                }

                private async void TrendAnalysisWindow_Loaded(object sender, RoutedEventArgs e)
                {
                        try
                        {
                                Console.WriteLine($"[趨勢視窗] Loaded 事件觸發");

                                // 設置預設的時間週期選項（3個月）
                                if (cboPeriod != null && cboPeriod.Items.Count > 1)
                                {
                                        cboPeriod.SelectedIndex = 1; // 選擇 "3個月"
                                        Console.WriteLine($"[趨勢視窗] 設置預設週期為 3個月");
                                }

                                InitializeStockSwitcher();

                                _isLoaded = true; // 標記為已載入

                                Console.WriteLine($"[趨勢視窗] 開始載入數據");
                                await LoadDataAsync();
                                Console.WriteLine($"[趨勢視窗] 數據載入完成");
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[趨勢視窗錯誤] Loaded 事件中發生錯誤: {ex.Message}");
                                Console.WriteLine($"[趨勢視窗錯誤] 堆疊追蹤: {ex.StackTrace}");
                                Console.WriteLine($"[趨勢視窗錯誤] 內部異常: {ex.InnerException?.Message}");

                                MessageBox.Show(
                                        $"載入數據時發生錯誤:\n\n{ex.Message}\n\n{ex.InnerException?.Message}\n\n請查看調試視窗獲取詳細信息",
                                        "載入錯誤",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Warning
                                );

                                // 設置錯誤狀態
                                if (txtTitle != null) txtTitle.Text = $"錯誤 - {_stockName} ({_ticker})";
                                if (txtSubtitle != null) txtSubtitle.Text = $"載入失敗: {ex.Message}";
                                if (txtCurrentPrice != null) txtCurrentPrice.Text = "錯誤";
                        }
                }

                private void InitializeStockSwitcher()
                {
                        if (cboStockSwitcher == null)
                        {
                                return;
                        }

                        _suppressStockSwitchEvent = true;
                        try
                        {
                                var isTwMarket = IsTwTicker(_ticker);
                                var marketDefaults = isTwMarket ? AppConfig.DefaultTwStocks : AppConfig.DefaultStocks;
                                var itemMap = new Dictionary<string, StockSwitcherItem>(StringComparer.OrdinalIgnoreCase);

                                foreach (var kv in marketDefaults)
                                {
                                        itemMap[kv.Key] = new StockSwitcherItem
                                        {
                                                Ticker = kv.Key,
                                                Name = kv.Value
                                        };
                                }

                                foreach (var t in _priceFetcher.GetPrices().Keys)
                                {
                                        if (IsTwTicker(t) == isTwMarket && !itemMap.ContainsKey(t))
                                        {
                                                itemMap[t] = new StockSwitcherItem
                                                {
                                                        Ticker = t,
                                                        Name = t
                                                };
                                        }
                                }

                                if (!itemMap.ContainsKey(_ticker))
                                {
                                        itemMap[_ticker] = new StockSwitcherItem
                                        {
                                                Ticker = _ticker,
                                                Name = _stockName
                                        };
                                }

                                var items = itemMap.Values.OrderBy(x => x.Name).ThenBy(x => x.Ticker).ToList();
                                cboStockSwitcher.ItemsSource = items;
                                cboStockSwitcher.SelectedValue = _ticker;
                        }
                        finally
                        {
                                _suppressStockSwitchEvent = false;
                        }
                }

                private void CboStockSwitcher_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                        if (_suppressStockSwitchEvent || !_isLoaded)
                        {
                                return;
                        }

                        var selectedTicker = cboStockSwitcher.SelectedValue as string;
                        if (string.IsNullOrWhiteSpace(selectedTicker))
                        {
                                return;
                        }

                        SwitchToTicker(selectedTicker.Trim().ToUpperInvariant());
                }

                private void CboStockSwitcher_KeyDown(object sender, KeyEventArgs e)
                {
                        if (e.Key != Key.Enter || !_isLoaded)
                        {
                                return;
                        }

                        var typedTicker = cboStockSwitcher.Text;
                        if (string.IsNullOrWhiteSpace(typedTicker))
                        {
                                return;
                        }

                        SwitchToTicker(typedTicker.Trim().ToUpperInvariant());
                }

                private void SwitchToTicker(string newTicker)
                {
                        if (string.IsNullOrWhiteSpace(newTicker) || string.Equals(_ticker, newTicker, StringComparison.OrdinalIgnoreCase))
                        {
                                return;
                        }

                        // 僅允許在同一市場切換（美股只能切美股，台股只能切台股）
                        if (IsTwTicker(newTicker) != IsTwTicker(_ticker))
                        {
                                MessageBox.Show("請在同一市場切換股票（美股→美股、台股→台股）。", "切換限制", MessageBoxButton.OK, MessageBoxImage.Information);
                                cboStockSwitcher.Text = _ticker;
                                return;
                        }

                        _ticker = newTicker;
                        if (AppConfig.DefaultStocks.ContainsKey(newTicker))
                        {
                                _stockName = AppConfig.DefaultStocks[newTicker];
                        }
                        else if (AppConfig.DefaultTwStocks.ContainsKey(newTicker))
                        {
                                _stockName = AppConfig.DefaultTwStocks[newTicker];
                        }
                        else
                        {
                                _stockName = newTicker;
                        }

                        Console.WriteLine($"[趨勢視窗] 切換股票: {_ticker} ({_stockName})");
                        _ = LoadDataAsync();
                }

                private bool IsTwTicker(string ticker)
                {
                        return !string.IsNullOrEmpty(ticker) && ticker.EndsWith(".TW", StringComparison.OrdinalIgnoreCase);
                }

                private class StockSwitcherItem
                {
                        public string Ticker { get; set; }
                        public string Name { get; set; }
                        public string DisplayName => $"{Name} ({Ticker})";
                }

                private async Task LoadDataAsync()
                {
                        try
                        {
                                Console.WriteLine($"[趨勢視窗] 開始 LoadData for {_ticker}");

                                ShowLoadingWindow();
                                UpdateLoadingWindow(5, "初始化畫面...");
                                await Task.Yield();

                                if (txtTitle == null || txtSubtitle == null || txtCurrentPrice == null)
                                {
                                        Console.WriteLine($"[趨勢視窗錯誤] UI 控制項為 null: txtTitle={txtTitle != null}, txtSubtitle={txtSubtitle != null}, txtCurrentPrice={txtCurrentPrice != null}");
                                        throw new InvalidOperationException("UI 控制項未正確初始化");
                                }

                                txtTitle.Text = $"{_stockName} ({_ticker}) - 趨勢分析";
                                txtSubtitle.Text = $"資料更新時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | 週期: {GetPeriodDisplay(_currentPeriod)}";
                                UpdateLoadingWindow(15, "更新標題資訊...");

                                Console.WriteLine($"[趨勢視窗] 準備獲取即時價格");

                                if (_priceFetcher == null)
                                {
                                        Console.WriteLine($"[趨勢視窗錯誤] PriceFetcher 為 null");
                                        throw new InvalidOperationException("PriceFetcher 未設置");
                                }

                                // 獲取當前價格
                                var result = await Task.Run(() => _priceFetcher.GetRealtimePriceWithSource(_ticker));
                                var currentPrice = result.Item1;
                                var changePercent = result.Item2;
                                UpdateLoadingWindow(45, "取得即時價格...");

                                Console.WriteLine($"[趨勢視窗] 獲取價格結果: Price={currentPrice}, Change={changePercent}");

                                // 如果 API 沒返回漲跌幅，嘗試從緩存獲取
                                if (!changePercent.HasValue && currentPrice.HasValue)
                                {
                                        Console.WriteLine($"[趨勢視窗] API 未返回漲跌幅，嘗試從緩存獲取");
                                        var cached = _priceFetcher.GetCachedPrice(_ticker);
                                        if (cached.Item2.HasValue)
                                        {
                                                changePercent = cached.Item2;
                                                Console.WriteLine($"[趨勢視窗] 從緩存成功獲取漲跌幅: {changePercent.Value:F2}%");
                                        }
                                        else
                                        {
                                                Console.WriteLine($"[趨勢視窗] 緩存中也沒有漲跌幅數據");
                                        }
                                }

                                Console.WriteLine($"[趨勢視窗] txtChange 狀態: null={txtChange == null}");

                                if (currentPrice.HasValue)
                                {
                                        txtCurrentPrice.Text = $"${currentPrice.Value:F2}";
                                        Console.WriteLine($"[趨勢視窗] 已設置當前價格: {txtCurrentPrice.Text}");

                                        if (changePercent.HasValue && txtChange != null)
                                        {
                                                var changeText = changePercent.Value >= 0 
                                                        ? $"▲ +{changePercent.Value:F2}%" 
                                                        : $"▼ {changePercent.Value:F2}%";

                                                txtChange.Text = changeText;
                                                txtChange.Foreground = changePercent.Value >= 0 
                                                        ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50"))
                                                        : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336"));

                                                Console.WriteLine($"[趨勢視窗] 已設置漲跌幅: {changeText}");
                                        }
                                        else if (txtChange != null)
                                        {
                                                txtChange.Text = "N/A";
                                                txtChange.Foreground = new SolidColorBrush(Colors.Gray);
                                                Console.WriteLine($"[趨勢視窗] changePercent 為 null，設置為 N/A");
                                        }
                                        else
                                        {
                                                Console.WriteLine($"[趨勢視窗錯誤] txtChange 控制項為 null！");
                                        }

                                        Console.WriteLine($"[趨勢視窗] 準備生成分析數據");
                                        UpdateLoadingWindow(60, "抓取歷史資料與計算技術指標...");
                                        await Task.Yield();

                                        // 生成歷史數據和技術指標（背景執行）
                                        await GenerateAnalysisDataAsync(currentPrice.Value, changePercent);
                                        UpdateLoadingWindow(95, "完成，正在整理畫面...");

                                        Console.WriteLine($"[趨勢視窗] LoadData 完成");
                                }
                                else
                                {
                                        Console.WriteLine($"[趨勢視窗] 無法獲取價格，顯示錯誤訊息");
                                        txtCurrentPrice.Text = "無法獲取";
                                        txtSubtitle.Text = "無法獲取股價數據，請檢查網路連線或股票代號";
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[趨勢視窗錯誤] LoadData 發生錯誤: {ex.Message}");
                                Console.WriteLine($"[趨勢視窗錯誤] 錯誤類型: {ex.GetType().Name}");
                                Console.WriteLine($"[趨勢視窗錯誤] 堆疊追蹤: {ex.StackTrace}");
                                if (ex.InnerException != null)
                                {
                                        Console.WriteLine($"[趨勢視窗錯誤] 內部異常: {ex.InnerException.Message}");
                                }

                                if (txtSubtitle != null) txtSubtitle.Text = $"載入數據時發生錯誤: {ex.Message}";
                                if (txtCurrentPrice != null) txtCurrentPrice.Text = "錯誤";

                                throw; // 重新拋出以便在 Loaded 事件中捕獲
                        }
                        finally
                        {
                                UpdateLoadingWindow(100, "完成");
                                HideLoadingWindow();
                        }
                }

                private void ShowLoadingWindow()
                {
                        if (_loadingWindow == null)
                        {
                                _loadingWindow = new LoadingProgressPopup
                                {
                                        Owner = this
                                };
                        }

                        if (!_loadingWindow.IsVisible)
                        {
                                _loadingWindow.Show();
                        }

                        _loadingWindow.Activate();
                }

                private void UpdateLoadingWindow(double progress, string message)
                {
                        if (_loadingWindow == null)
                        {
                                return;
                        }

                        _loadingWindow.UpdateProgress(progress, message);
                }

                private void HideLoadingWindow()
                {
                        if (_loadingWindow != null)
                        {
                                _loadingWindow.Close();
                                _loadingWindow = null;
                        }
                }

                private async Task GenerateAnalysisDataAsync(double currentPrice, double? changePercent)
                {
                        try
                        {
                                Console.WriteLine($"[趨勢視窗] 開始 GenerateAnalysisData (async)");

                                var calc = await Task.Run(() => ComputeAnalysisData(currentPrice));
                                _historicalData = calc.HistoricalData;

                                txtHigh.Text = $"${calc.MaxPrice:F2}";
                                txtLow.Text = $"${calc.MinPrice:F2}";
                                txtVolume.Text = FormatVolume(calc.AvgVolume);

                                // 背景計算後回到 UI 套用技術指標
                                txtMA5.Text = $"${calc.MA5:F2}";
                                txtMA20.Text = $"${calc.MA20:F2}";
                                txtMA60.Text = $"${calc.MA60:F2}";
                                txtMAAnalysis.Text = calc.MAAnalysis;

                                txtRSI.Text = $"{calc.RSI:F1}";
                                pbRSI.Value = calc.RSI;
                                txtRSIAnalysis.Text = calc.RSIAnalysis;
                                var rsiBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(calc.RSIColorHex));
                                pbRSI.Foreground = rsiBrush;
                                txtRSI.Foreground = rsiBrush;

                                txtMACD.Text = $"{calc.MACD:F3}";
                                txtSignal.Text = $"{calc.Signal:F3}";
                                txtHistogram.Text = $"{calc.Histogram:F3}";
                                txtHistogram.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(calc.HistogramColorHex));
                                txtMACDAnalysis.Text = calc.MACDAnalysis;

                                Console.WriteLine($"[趨勢視窗] 準備繪製圖表");

                                // 繪製圖表
                                DrawCandlestickChart(calc.HistoricalData);
                                DrawVolumeChart(calc.HistoricalData);
                                DrawMacdChart(calc.HistoricalData);
                                DrawRsiChart(calc.HistoricalData);

                                Console.WriteLine($"[趨勢視窗] 生成 AI 分析");

                                // AI 分析
                                GenerateAIAnalysis(calc.HistoricalData, currentPrice, changePercent);
                                UpdateFundamentalAnalysis();
                                UpdateMajorNewsTable();

                                Console.WriteLine($"[趨勢視窗] GenerateAnalysisData 完成");
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[趨勢視窗錯誤] GenerateAnalysisData 發生錯誤: {ex.Message}");
                                Console.WriteLine($"[趨勢視窗錯誤] 堆疊追蹤: {ex.StackTrace}");

                                MessageBox.Show(
                                        $"生成分析數據時發生錯誤\n\n{ex.Message}",
                                        "分析錯誤",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Warning
                                );
                        }
                }

                private void UpdateMajorNewsTable()
                {
                        if (dgMajorNews == null)
                        {
                                return;
                        }

                        try
                        {
                                var url = $"https://feeds.finance.yahoo.com/rss/2.0/headline?s={_ticker}&region=US&lang=zh-TW";
                                var news = new List<MajorNewsItem>();

                                using (var client = new WebClient())
                                {
                                        client.Encoding = System.Text.Encoding.UTF8;
                                        var xml = client.DownloadString(url);
                                        var doc = XDocument.Parse(xml);

                                        var items = doc.Descendants("item").Take(8).ToList();
                                        foreach (var item in items)
                                        {
                                                var title = item.Element("title")?.Value ?? string.Empty;
                                                var pubDate = item.Element("pubDate")?.Value;

                                                DateTime parsedTime;
                                                var showTime = DateTime.TryParse(pubDate, out parsedTime)
                                                        ? parsedTime.ToString("MM-dd HH:mm")
                                                        : "--";

                                                var impact = EvaluateNewsImpact(title);

                                                news.Add(new MajorNewsItem
                                                {
                                                        Time = showTime,
                                                        Title = title,
                                                        ImpactLevel = impact
                                                });
                                        }
                                }

                                dgMajorNews.ItemsSource = news;
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[重大新聞] 載入失敗: {ex.Message}");
                                dgMajorNews.ItemsSource = new List<MajorNewsItem>
                                {
                                        new MajorNewsItem { Time = "--", Title = "目前無法取得重大新聞", ImpactLevel = "--" }
                                };
                        }
                }

                private string EvaluateNewsImpact(string title)
                {
                        if (string.IsNullOrWhiteSpace(title))
                        {
                                return "低";
                        }

                        var t = title.ToLowerInvariant();
                        var high = new[] { "earnings", "guidance", "merger", "acquisition", "lawsuit", "fed", "財報", "併購", "訴訟", "降評", "升評" };
                        var medium = new[] { "target", "rating", "forecast", "outlook", "目標價", "展望", "預測" };

                        if (high.Any(k => t.Contains(k.ToLowerInvariant())))
                        {
                                return "高";
                        }

                        if (medium.Any(k => t.Contains(k.ToLowerInvariant())))
                        {
                                return "中";
                        }

                        return "低";
                }

                private AnalysisCalculationResult ComputeAnalysisData(double currentPrice)
                {
                        var historicalData = new List<CandlestickData>();

                        if (!TryLoadHistoricalDataFromYFinance(_ticker, _currentPeriod, historicalData))
                        {
                                Console.WriteLine($"[趨勢視窗] 無法取得 {_ticker} 歷史資料，取消隨機資料回填。");
                        }

                        if (historicalData.Count == 0)
                        {
                                return new AnalysisCalculationResult
                                {
                                        HistoricalData = historicalData,
                                        MinPrice = currentPrice,
                                        MaxPrice = currentPrice,
                                        AvgVolume = 0,
                                        MA5 = currentPrice,
                                        MA20 = currentPrice,
                                        MA60 = currentPrice,
                                        MAAnalysis = "無法取得歷史資料，無法計算均線。",
                                        RSI = 50,
                                        RSIAnalysis = "無法取得歷史資料，無法計算 RSI。",
                                        RSIColorHex = "#9E9E9E",
                                        MACD = 0,
                                        Signal = 0,
                                        Histogram = 0,
                                        MACDAnalysis = "無法取得歷史資料，無法計算 MACD。",
                                        HistogramColorHex = "#9E9E9E"
                                };
                        }

                        var minPrice = historicalData.Min(x => x.Low);
                        var maxPrice = historicalData.Max(x => x.High);
                        var avgVolume = historicalData.Count > 0 ? historicalData.Sum(x => x.Volume) / historicalData.Count : 0;

                        var closes = historicalData.Select(d => d.Close).ToList();
                        var ma5 = CalculateMA(closes, 5);
                        var ma20 = CalculateMA(closes, 20);
                        var ma60 = CalculateMA(closes, 60);

                        string maAnalysis;
                        if (currentPrice > ma5 && ma5 > ma20 && ma20 > ma60)
                                maAnalysis = "✅ 多頭排列：短中長期均線呈多頭排列，趨勢向上，建議持有或買入。";
                        else if (currentPrice < ma5 && ma5 < ma20 && ma20 < ma60)
                                maAnalysis = "❌ 空頭排列：短中長期均線呈空頭排列，趨勢向下，建議觀望或減倉。";
                        else
                                maAnalysis = "⚠️ 均線糾結：均線排列不明確，市場方向不清晰，建議謹慎操作。";

                        var rsi = CalculateRSI(closes, 14);
                        string rsiAnalysis;
                        string rsiColor;
                        if (rsi < 30)
                        {
                                rsiColor = "#4CAF50";
                                rsiAnalysis = "💰 RSI 超賣區（<30）：股票可能被過度賣出，存在反彈機會，可考慮買入。";
                        }
                        else if (rsi > 70)
                        {
                                rsiColor = "#F44336";
                                rsiAnalysis = "⚠️ RSI 超買區（>70）：股票可能被過度買入，存在回調風險，建議獲利了結。";
                        }
                        else
                        {
                                rsiColor = "#2196F3";
                                rsiAnalysis = $"📊 RSI 正常區間（30-70）：當前 RSI = {rsi:F1}，市場處於正常狀態，可持續觀察。";
                        }

                        var macdResult = CalculateMACD(closes);
                        var histogram = macdResult.Item3;
                        string macdAnalysis;
                        string histColor;
                        if (histogram > 0)
                        {
                                histColor = "#4CAF50";
                                macdAnalysis = "✅ MACD 柱狀圖為正，動能偏多。";
                        }
                        else
                        {
                                histColor = "#F44336";
                                macdAnalysis = "📉 MACD 柱狀圖為負，動能偏空。";
                        }

                        return new AnalysisCalculationResult
                        {
                                HistoricalData = historicalData,
                                MinPrice = minPrice,
                                MaxPrice = maxPrice,
                                AvgVolume = avgVolume,
                                MA5 = ma5,
                                MA20 = ma20,
                                MA60 = ma60,
                                MAAnalysis = maAnalysis,
                                RSI = rsi,
                                RSIAnalysis = rsiAnalysis,
                                RSIColorHex = rsiColor,
                                MACD = macdResult.Item1,
                                Signal = macdResult.Item2,
                                Histogram = macdResult.Item3,
                                MACDAnalysis = macdAnalysis,
                                HistogramColorHex = histColor
                        };
                }

                private class AnalysisCalculationResult
                {
                        public List<CandlestickData> HistoricalData { get; set; }
                        public double MinPrice { get; set; }
                        public double MaxPrice { get; set; }
                        public long AvgVolume { get; set; }
                        public double MA5 { get; set; }
                        public double MA20 { get; set; }
                        public double MA60 { get; set; }
                        public string MAAnalysis { get; set; }
                        public double RSI { get; set; }
                        public string RSIAnalysis { get; set; }
                        public string RSIColorHex { get; set; }
                        public double MACD { get; set; }
                        public double Signal { get; set; }
                        public double Histogram { get; set; }
                        public string MACDAnalysis { get; set; }
                        public string HistogramColorHex { get; set; }
                }

                private void CalculateTechnicalIndicators(List<CandlestickData> data, double currentPrice)
                {
                        try
                        {
                                if (data == null || data.Count < 60)
                                {
                                        Console.WriteLine($"數據不足，無法計算技術指標。需要至少 60 筆數據，目前只有 {data?.Count ?? 0} 筆");
                                        return;
                                }

                                var closes = data.Select(d => d.Close).ToList();

                                // 計算移動平均線
                                var ma5 = CalculateMA(closes, 5);
                                var ma20 = CalculateMA(closes, 20);
                                var ma60 = CalculateMA(closes, 60);

                                txtMA5.Text = $"${ma5:F2}";
                                txtMA20.Text = $"${ma20:F2}";
                                txtMA60.Text = $"${ma60:F2}";

                                // MA 分析
                                if (currentPrice > ma5 && ma5 > ma20 && ma20 > ma60)
                                {
                                        txtMAAnalysis.Text = "✅ 多頭排列：短中長期均線呈多頭排列，趨勢向上，建議持有或買入。";
                                }
                                else if (currentPrice < ma5 && ma5 < ma20 && ma20 < ma60)
                                {
                                        txtMAAnalysis.Text = "❌ 空頭排列：短中長期均線呈空頭排列，趨勢向下，建議觀望或減倉。";
                                }
                                else
                                {
                                        txtMAAnalysis.Text = "⚠️ 均線糾結：均線排列不明確，市場方向不清晰，建議謹慎操作。";
                                }

                                // 計算 RSI
                                var rsi = CalculateRSI(closes, 14);
                                txtRSI.Text = $"{rsi:F1}";
                                pbRSI.Value = rsi;

                                if (rsi < 30)
                                {
                                        pbRSI.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50"));
                                        txtRSI.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50"));
                                        txtRSIAnalysis.Text = "💰 RSI 超賣區（<30）：股票可能被過度賣出，存在反彈機會，可考慮買入。";
                                }
                                else if (rsi > 70)
                                {
                                        pbRSI.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336"));
                                        txtRSI.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336"));
                                        txtRSIAnalysis.Text = "⚠️ RSI 超買區（>70）：股票可能被過度買入，存在回調風險，建議獲利了結。";
                                }
                                else
                                {
                                        pbRSI.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2196F3"));
                                        txtRSI.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2196F3"));
                                        txtRSIAnalysis.Text = $"📊 RSI 正常區間（30-70）：當前 RSI = {rsi:F1}，市場處於正常狀態，可持續觀察。";
                                }

                                // 計算 MACD
                                var macdResult = CalculateMACD(closes);
                                txtMACD.Text = $"{macdResult.Item1:F3}";
                                txtSignal.Text = $"{macdResult.Item2:F3}";
                                txtHistogram.Text = $"{macdResult.Item3:F3}";

                                var histogram = macdResult.Item3;
                                if (histogram > 0)
                                {
                                        txtHistogram.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50"));
                                        var prevHistogram = histogram * 0.8; // 簡化比較
                                        txtMACDAnalysis.Text = histogram > prevHistogram
                                                ? "✅ MACD 金叉：柱狀圖為正且擴大，買入信號強烈。" 
                                                : "📈 MACD 正值：柱狀圖為正，趨勢向上，可考慮持有。";
                                }
                                else
                                {
                                        txtHistogram.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336"));
                                        var prevHistogram = histogram * 0.8;
                                        txtMACDAnalysis.Text = histogram < prevHistogram
                                                ? "❌ MACD 死叉：柱狀圖為負且擴大，賣出信號強烈。"
                                                : "📉 MACD 負值：柱狀圖為負，趨勢向下，建議觀望。";
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"計算技術指標時發生錯誤: {ex.Message}");
                                txtMAAnalysis.Text = $"計算錯誤: {ex.Message}";
                        }
                }

                private double CalculateMA(List<double> prices, int period)
                {
                        if (prices.Count < period) return prices.Last();
                        return prices.Skip(prices.Count - period).Take(period).Average();
                }

                private double CalculateRSI(List<double> prices, int period)
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

                private Tuple<double, double, double> CalculateMACD(List<double> prices)
                {
                        if (prices.Count < 26) return Tuple.Create(0.0, 0.0, 0.0);

                        var ema12 = CalculateEMA(prices, 12);
                        var ema26 = CalculateEMA(prices, 26);
                        var macd = ema12 - ema26;

                        var macdLine = new List<double> { macd };
                        var signal = CalculateEMA(macdLine, 9);
                        var histogram = macd - signal;

                        return Tuple.Create(macd, signal, histogram);
                }

                private double CalculateEMA(List<double> prices, int period)
                {
                        if (prices.Count < period) return prices.Last();

                        var multiplier = 2.0 / (period + 1);
                        var ema = prices.Take(period).Average();

                        for (int i = period; i < prices.Count; i++)
                        {
                                ema = (prices[i] - ema) * multiplier + ema;
                        }

                        return ema;
                }

                private void DrawCandlestickChart(List<CandlestickData> data)
                {
                        try
                        {
                                chartCanvas.Children.Clear();

                                if (data == null || data.Count == 0)
                                {
                                        // 顯示無數據提示
                                        var noDataText = new TextBlock
                                        {
                                                Text = "暫無圖表數據",
                                                FontSize = 16,
                                                Foreground = new SolidColorBrush(Colors.Gray),
                                                HorizontalAlignment = HorizontalAlignment.Center,
                                                VerticalAlignment = VerticalAlignment.Center
                                        };
                                        chartCanvas.Children.Add(noDataText);
                                        Canvas.SetLeft(noDataText, 350);
                                        Canvas.SetTop(noDataText, 125);
                                        return;
                                }

                                // 等待 Canvas 實際渲染
                                chartCanvas.UpdateLayout();

                                var width = chartCanvas.ActualWidth > 0 ? chartCanvas.ActualWidth : 800;
                                var height = chartCanvas.ActualHeight > 0 ? chartCanvas.ActualHeight : 250;

                                if (width <= 0) width = 800;
                                if (height <= 0) height = 250;

                                var paddingXLeft = 40.0;
                                var paddingXRight = 15.0;
                                var paddingTop = 10.0;
                                var paddingBottom = 28.0;

                                var minPrice = data.Min(d => d.Low);
                                var maxPrice = data.Max(d => d.High);
                                var priceRange = maxPrice - minPrice;

                                if (priceRange == 0) priceRange = 1; // 避免除以零

                                var displayData = GetDisplayData(data);
                                var displayCount = displayData.Count;
                                var plotWidth = Math.Max(width - paddingXLeft - paddingXRight, 100);
                                var plotHeight = Math.Max(height - paddingTop - paddingBottom, 80);
                                var candleWidth = Math.Max(plotWidth / displayCount * 0.6, 2);
                                var candleSpacing = plotWidth / displayCount;

                                var chartLeft = paddingXLeft;
                                var chartRight = width - paddingXRight;
                                var chartTop = paddingTop;
                                var chartBottom = height - paddingBottom;

                                _chartAreaLeft = chartLeft;
                                _chartAreaRight = chartRight;
                                _chartAreaTop = chartTop;
                                _chartAreaBottom = chartBottom;

                                // 繪製 Y 軸
                                var yAxis = new Line
                                {
                                        X1 = chartLeft,
                                        Y1 = chartTop,
                                        X2 = chartLeft,
                                        Y2 = chartBottom,
                                        Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                        StrokeThickness = 1.5
                                };
                                chartCanvas.Children.Add(yAxis);

                                // 繪製 X 軸
                                var xAxis = new Line
                                {
                                        X1 = chartLeft,
                                        Y1 = chartBottom,
                                        X2 = chartRight,
                                        Y2 = chartBottom,
                                        Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                        StrokeThickness = 1.5
                                };
                                chartCanvas.Children.Add(xAxis);

                                // Y 軸刻度與價格標籤
                                var yTickCount = 5;
                                for (int i = 0; i <= yTickCount; i++)
                                {
                                        var ratio = i / (double)yTickCount;
                                        var y = chartBottom - ratio * (chartBottom - chartTop);
                                        var price = minPrice + ratio * priceRange;

                                        var yTick = new Line
                                        {
                                                X1 = chartLeft - 4,
                                                Y1 = y,
                                                X2 = chartLeft,
                                                Y2 = y,
                                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                                StrokeThickness = 1
                                        };
                                        chartCanvas.Children.Add(yTick);

                                        var yLabel = new TextBlock
                                        {
                                                Text = price.ToString("F2"),
                                                FontSize = 9,
                                                Foreground = new SolidColorBrush(Colors.DimGray)
                                        };
                                        Canvas.SetLeft(yLabel, 2);
                                        Canvas.SetTop(yLabel, y - 8);
                                        chartCanvas.Children.Add(yLabel);
                                }

                                // X 軸刻度與日期標籤
                                var xTickCount = Math.Min(6, displayCount);
                                if (xTickCount > 1)
                                {
                                        for (int i = 0; i < xTickCount; i++)
                                        {
                                                var dataIndex = (int)Math.Round(i * (displayCount - 1) / (double)(xTickCount - 1));
                                                var x = chartLeft + dataIndex * candleSpacing + candleWidth / 2;

                                                var xTick = new Line
                                                {
                                                        X1 = x,
                                                        Y1 = chartBottom,
                                                        X2 = x,
                                                        Y2 = chartBottom + 4,
                                                        Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                                        StrokeThickness = 1
                                                };
                                                chartCanvas.Children.Add(xTick);

                                                var candleData = displayData[dataIndex];
                                                var xLabel = new TextBlock
                                                {
                                                        Text = candleData.Date,
                                                        FontSize = 9,
                                                        Foreground = new SolidColorBrush(Colors.DimGray)
                                                };
                                                Canvas.SetLeft(xLabel, x - 16);
                                                Canvas.SetTop(xLabel, chartBottom + 6);
                                                chartCanvas.Children.Add(xLabel);
                                        }
                                }

                                for (int i = 0; i < displayCount; i++)
                                {
                                        var candle = displayData[i];
                                        var x = chartLeft + i * candleSpacing;
                                        var tooltipText =
                                                $"日期: {candle.Date}\n" +
                                                $"開盤: ${candle.Open:F2}\n" +
                                                $"最高: ${candle.High:F2}\n" +
                                                $"最低: ${candle.Low:F2}\n" +
                                                $"收盤: ${candle.Close:F2}\n" +
                                                $"成交量: {candle.Volume:N0}";

                                        var openY = chartBottom - ((candle.Open - minPrice) / priceRange) * (chartBottom - chartTop);
                                        var closeY = chartBottom - ((candle.Close - minPrice) / priceRange) * (chartBottom - chartTop);
                                        var highY = chartBottom - ((candle.High - minPrice) / priceRange) * (chartBottom - chartTop);
                                        var lowY = chartBottom - ((candle.Low - minPrice) / priceRange) * (chartBottom - chartTop);

                                        var isUp = candle.Close >= candle.Open;
                                        var color = isUp ? Colors.Green : Colors.Red;

                                        // 繪製高低線
                                        var wickLine = new Line
                                        {
                                                X1 = x + candleWidth / 2,
                                                Y1 = highY,
                                                X2 = x + candleWidth / 2,
                                                Y2 = lowY,
                                                Stroke = new SolidColorBrush(color),
                                                StrokeThickness = 1
                                        };
                                        ToolTipService.SetToolTip(wickLine, tooltipText);
                                        chartCanvas.Children.Add(wickLine);

                                        // 繪製實體
                                        var bodyHeight = Math.Max(Math.Abs(closeY - openY), 1);
                                        var bodyRect = new Rectangle
                                        {
                                                Width = candleWidth,
                                                Height = bodyHeight,
                                                Fill = isUp ? new SolidColorBrush(Color.FromRgb(76, 175, 80)) : new SolidColorBrush(Color.FromRgb(244, 67, 54)),
                                                Stroke = new SolidColorBrush(color),
                                                StrokeThickness = 1
                                        };
                                        ToolTipService.SetToolTip(bodyRect, tooltipText);
                                        Canvas.SetLeft(bodyRect, x);
                                        Canvas.SetTop(bodyRect, Math.Min(openY, closeY));
                                        chartCanvas.Children.Add(bodyRect);

                                        // 增加透明命中區，讓滑鼠更容易觸發 Tooltip
                                        var hitArea = new Rectangle
                                        {
                                                Width = candleSpacing,
                                                Height = chartBottom - chartTop,
                                                Fill = Brushes.Transparent
                                        };
                                        ToolTipService.SetToolTip(hitArea, tooltipText);
                                        Canvas.SetLeft(hitArea, x - (candleSpacing - candleWidth) / 2);
                                        Canvas.SetTop(hitArea, chartTop);
                                        chartCanvas.Children.Add(hitArea);
                                }

                                // 繪製 MA5 / MA20 均線（疊加在 K 線主圖）
                                var closeSeries = displayData.Select(d => d.Close).ToList();
                                var ma5Series = BuildMASeries(closeSeries, 5);
                                var ma20Series = BuildMASeries(closeSeries, 20);

                                DrawMovingAverageLine(chartCanvas, ma5Series, chartLeft, chartTop, chartBottom, candleSpacing, candleWidth, minPrice, priceRange,
                                        Color.FromRgb(255, 193, 7), 1.5); // 黃色 MA5
                                DrawMovingAverageLine(chartCanvas, ma20Series, chartLeft, chartTop, chartBottom, candleSpacing, candleWidth, minPrice, priceRange,
                                        Color.FromRgb(103, 58, 183), 1.5); // 紫色 MA20

                                // 回測目前預測方式，並在 K 線上標記買賣點
                                var signals = BuildBacktestSignals(displayData);
                                DrawBacktestSignals(chartCanvas, displayData, signals, chartLeft, chartTop, chartBottom, candleSpacing, candleWidth, minPrice, priceRange);

                                // 添加價格標籤
                                var priceLabel = new TextBlock
                                {
                                        Text = $"最高: ${maxPrice:F2}  最低: ${minPrice:F2}",
                                        FontSize = 10,
                                        Foreground = new SolidColorBrush(Colors.Gray)
                                };
                                Canvas.SetLeft(priceLabel, 10);
                                Canvas.SetTop(priceLabel, 5);
                                chartCanvas.Children.Add(priceLabel);

                                var maLegend = new TextBlock
                                {
                                        Text = "MA5(黃)  MA20(紫)",
                                        FontSize = 10,
                                        Foreground = new SolidColorBrush(Colors.Gray)
                                };
                                Canvas.SetLeft(maLegend, 190);
                                Canvas.SetTop(maLegend, 5);
                                chartCanvas.Children.Add(maLegend);

                                if (signals.Count > 0)
                                {
                                        var buyCount = signals.Count(s => s.Item2.Contains("BUY"));
                                        var sellCount = signals.Count(s => s.Item2.Contains("SELL"));
                                        var strongBuyCount = signals.Count(s => s.Item2 == "STRONG_BUY");
                                        var strongSellCount = signals.Count(s => s.Item2 == "STRONG_SELL");
                                        var signalSummary = new TextBlock
                                        {
                                                Text = $"回測訊號: BUY {buyCount} (強{strongBuyCount}) / SELL {sellCount} (強{strongSellCount})",
                                                FontSize = 10,
                                                Foreground = new SolidColorBrush(Colors.DimGray)
                                        };
                                        Canvas.SetLeft(signalSummary, 330);
                                        Canvas.SetTop(signalSummary, 5);
                                        chartCanvas.Children.Add(signalSummary);
                                }

                                EnsureCrosshairElements();
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"繪製 K 線圖時發生錯誤: {ex.Message}");

                                // 顯示錯誤提示
                                chartCanvas.Children.Clear();
                                var errorText = new TextBlock
                                {
                                        Text = $"圖表繪製錯誤\n{ex.Message}",
                                        FontSize = 12,
                                        Foreground = new SolidColorBrush(Colors.Red),
                                        TextWrapping = TextWrapping.Wrap
                                };
                                chartCanvas.Children.Add(errorText);
                                Canvas.SetLeft(errorText, 10);
                                Canvas.SetTop(errorText, 10);
                        }
                }

                private void ChartCanvas_MouseMove(object sender, MouseEventArgs e)
                {
                        if (_crosshairVertical == null || _crosshairHorizontal == null)
                        {
                                return;
                        }

                        var p = e.GetPosition(chartCanvas);
                        var x = Math.Max(_chartAreaLeft, Math.Min(_chartAreaRight, p.X));
                        var y = Math.Max(_chartAreaTop, Math.Min(_chartAreaBottom, p.Y));

                        _crosshairVertical.X1 = x;
                        _crosshairVertical.Y1 = _chartAreaTop;
                        _crosshairVertical.X2 = x;
                        _crosshairVertical.Y2 = _chartAreaBottom;

                        _crosshairHorizontal.X1 = _chartAreaLeft;
                        _crosshairHorizontal.Y1 = y;
                        _crosshairHorizontal.X2 = _chartAreaRight;
                        _crosshairHorizontal.Y2 = y;

                        _crosshairVertical.Visibility = Visibility.Visible;
                        _crosshairHorizontal.Visibility = Visibility.Visible;
                }

                private void ChartCanvas_MouseLeave(object sender, MouseEventArgs e)
                {
                        if (_crosshairVertical != null)
                        {
                                _crosshairVertical.Visibility = Visibility.Collapsed;
                        }

                        if (_crosshairHorizontal != null)
                        {
                                _crosshairHorizontal.Visibility = Visibility.Collapsed;
                        }
                }

                private void EnsureCrosshairElements()
                {
                        _crosshairVertical = new Line
                        {
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1,
                                StrokeDashArray = new DoubleCollection { 2, 2 },
                                Visibility = Visibility.Collapsed,
                                IsHitTestVisible = false
                        };

                        _crosshairHorizontal = new Line
                        {
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1,
                                StrokeDashArray = new DoubleCollection { 2, 2 },
                                Visibility = Visibility.Collapsed,
                                IsHitTestVisible = false
                        };

                        chartCanvas.Children.Add(_crosshairVertical);
                        chartCanvas.Children.Add(_crosshairHorizontal);
                }

                private void ChartCanvas_SizeChanged(object sender, SizeChangedEventArgs e)
                {
                        if (!_isLoaded)
                        {
                                return;
                        }

                        if (_historicalData != null && _historicalData.Count > 0)
                        {
                                DrawCandlestickChart(_historicalData);
                                DrawVolumeChart(_historicalData);
                                DrawMacdChart(_historicalData);
                                DrawRsiChart(_historicalData);
                        }
                }

                private void IndicatorCanvas_SizeChanged(object sender, SizeChangedEventArgs e)
                {
                        if (!_isLoaded)
                        {
                                return;
                        }

                        if (_historicalData != null && _historicalData.Count > 0)
                        {
                                DrawVolumeChart(_historicalData);
                                DrawMacdChart(_historicalData);
                                DrawRsiChart(_historicalData);
                        }
                }

                private void DrawVolumeChart(List<CandlestickData> data)
                {
                        if (volumeCanvas == null)
                        {
                                return;
                        }

                        volumeCanvas.Children.Clear();
                        if (data == null || data.Count == 0)
                        {
                                return;
                        }

                        var displayData = GetDisplayData(data);
                        var displayCount = displayData.Count;

                        var width = volumeCanvas.ActualWidth > 0 ? volumeCanvas.ActualWidth : 800;
                        var height = volumeCanvas.ActualHeight > 0 ? volumeCanvas.ActualHeight : 120;
                        var left = 40.0;
                        var right = width - 15;
                        var top = 10.0;
                        var bottom = height - 22;

                        var yAxis = new Line
                        {
                                X1 = left,
                                Y1 = top,
                                X2 = left,
                                Y2 = bottom,
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1
                        };
                        volumeCanvas.Children.Add(yAxis);

                        var xAxis = new Line
                        {
                                X1 = left,
                                Y1 = bottom,
                                X2 = right,
                                Y2 = bottom,
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1
                        };
                        volumeCanvas.Children.Add(xAxis);

                        var maxVolume = displayData.Max(d => (double)d.Volume);
                        if (maxVolume <= 0)
                        {
                                maxVolume = 1;
                        }

                        var plotWidth = Math.Max(right - left, 10);
                        var spacing = plotWidth / displayCount;
                        var barWidth = Math.Max(spacing * 0.65, 1);

                        for (int i = 0; i < displayData.Count; i++)
                        {
                                var item = displayData[i];
                                var x = left + i * spacing + (spacing - barWidth) / 2;
                                var barHeight = (item.Volume / maxVolume) * (bottom - top);
                                var y = bottom - barHeight;

                                var isUp = item.Close >= item.Open;
                                var bar = new Rectangle
                                {
                                        Width = barWidth,
                                        Height = Math.Max(barHeight, 1),
                                        Fill = isUp
                                                ? new SolidColorBrush(Color.FromRgb(76, 175, 80))
                                                : new SolidColorBrush(Color.FromRgb(244, 67, 54)),
                                        Opacity = 0.85
                                };

                                Canvas.SetLeft(bar, x);
                                Canvas.SetTop(bar, y);
                                volumeCanvas.Children.Add(bar);
                        }

                        var title = new TextBlock
                        {
                                Text = "Volume (綠=上漲, 紅=下跌)",
                                FontSize = 10,
                                Foreground = new SolidColorBrush(Colors.DimGray)
                        };
                        Canvas.SetLeft(title, 4);
                        Canvas.SetTop(title, 2);
                        volumeCanvas.Children.Add(title);
                }

                private void DrawMacdChart(List<CandlestickData> data)
                {
                        if (macdCanvas == null)
                        {
                                return;
                        }

                        var displayData = GetDisplayData(data);
                        var closeSeries = displayData.Select(d => d.Close).ToList();
                        var macdTuple = BuildMACDComponents(closeSeries);
                        var macd = macdTuple.Item1;
                        var signal = macdTuple.Item2;
                        var histogram = macdTuple.Item3;

                        macdCanvas.Children.Clear();
                        if (macd.Count < 2)
                        {
                                return;
                        }

                        var width = macdCanvas.ActualWidth > 0 ? macdCanvas.ActualWidth : 800;
                        var height = macdCanvas.ActualHeight > 0 ? macdCanvas.ActualHeight : 120;
                        var left = 40.0;
                        var right = width - 15;
                        var top = 10.0;
                        var bottom = height - 22;

                        var all = macd.Concat(signal).Concat(histogram).ToList();
                        var min = all.Min();
                        var max = all.Max();
                        if (Math.Abs(max - min) < 0.000001)
                        {
                                max += 1;
                                min -= 1;
                        }

                        var yAxis = new Line
                        {
                                X1 = left,
                                Y1 = top,
                                X2 = left,
                                Y2 = bottom,
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1
                        };
                        macdCanvas.Children.Add(yAxis);

                        var xAxis = new Line
                        {
                                X1 = left,
                                Y1 = bottom,
                                X2 = right,
                                Y2 = bottom,
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1
                        };
                        macdCanvas.Children.Add(xAxis);

                        var zeroY = bottom - ((0 - min) / (max - min)) * (bottom - top);
                        var zeroLine = new Line
                        {
                                X1 = left,
                                Y1 = zeroY,
                                X2 = right,
                                Y2 = zeroY,
                                Stroke = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                                StrokeThickness = 1,
                                StrokeDashArray = new DoubleCollection { 3, 2 }
                        };
                        macdCanvas.Children.Add(zeroLine);

                        var plotWidth = Math.Max(right - left, 10);
                        var spacing = plotWidth / macd.Count;
                        var barWidth = Math.Max(spacing * 0.7, 1);

                        for (int i = 0; i < histogram.Count; i++)
                        {
                                var x = left + i * spacing + (spacing - barWidth) / 2;
                                var valueY = bottom - ((histogram[i] - min) / (max - min)) * (bottom - top);
                                var barTop = Math.Min(valueY, zeroY);
                                var barHeight = Math.Max(Math.Abs(valueY - zeroY), 1);

                                var bar = new Rectangle
                                {
                                        Width = barWidth,
                                        Height = barHeight,
                                        Fill = histogram[i] >= 0
                                                ? new SolidColorBrush(Color.FromRgb(76, 175, 80))
                                                : new SolidColorBrush(Color.FromRgb(244, 67, 54)),
                                        Opacity = 0.7
                                };
                                Canvas.SetLeft(bar, x);
                                Canvas.SetTop(bar, barTop);
                                macdCanvas.Children.Add(bar);
                        }

                        var macdLine = new Polyline
                        {
                                Stroke = new SolidColorBrush(Color.FromRgb(33, 150, 243)),
                                StrokeThickness = 1.6
                        };

                        var signalLine = new Polyline
                        {
                                Stroke = new SolidColorBrush(Color.FromRgb(255, 152, 0)),
                                StrokeThickness = 1.4
                        };

                        for (int i = 0; i < macd.Count; i++)
                        {
                                var x = left + i * spacing + spacing / 2;
                                var macdY = bottom - ((macd[i] - min) / (max - min)) * (bottom - top);
                                var signalY = bottom - ((signal[i] - min) / (max - min)) * (bottom - top);
                                macdLine.Points.Add(new Point(x, macdY));
                                signalLine.Points.Add(new Point(x, signalY));
                        }

                        macdCanvas.Children.Add(macdLine);
                        macdCanvas.Children.Add(signalLine);

                        var title = new TextBlock
                        {
                                Text = "MACD / Signal / Histogram",
                                FontSize = 10,
                                Foreground = new SolidColorBrush(Colors.DimGray)
                        };
                        Canvas.SetLeft(title, 4);
                        Canvas.SetTop(title, 2);
                        macdCanvas.Children.Add(title);
                }

                private void DrawRsiChart(List<CandlestickData> data)
                {
                        if (rsiCanvas == null)
                        {
                                return;
                        }

                        var displayData = GetDisplayData(data);
                        var closeSeries = displayData.Select(d => d.Close).ToList();
                        var series = BuildRSISeries(closeSeries, 14);

                        rsiCanvas.Children.Clear();
                        if (series == null || series.Count < 2)
                        {
                                return;
                        }

                        var width = rsiCanvas.ActualWidth > 0 ? rsiCanvas.ActualWidth : 800;
                        var height = rsiCanvas.ActualHeight > 0 ? rsiCanvas.ActualHeight : 120;
                        var left = 40.0;
                        var right = width - 15;
                        var top = 10.0;
                        var bottom = height - 22;

                        // 座標軸
                        var yAxis = new Line
                        {
                                X1 = left,
                                Y1 = top,
                                X2 = left,
                                Y2 = bottom,
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1
                        };
                        rsiCanvas.Children.Add(yAxis);

                        var xAxis = new Line
                        {
                                X1 = left,
                                Y1 = bottom,
                                X2 = right,
                                Y2 = bottom,
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1
                        };
                        rsiCanvas.Children.Add(xAxis);

                        Func<double, double> rsiToY = rsi => bottom - (rsi / 100.0) * (bottom - top);

                        // 超買(70)虛線
                        var overboughtY = rsiToY(70);
                        var overboughtLine = new Line
                        {
                                X1 = left,
                                Y1 = overboughtY,
                                X2 = right,
                                Y2 = overboughtY,
                                Stroke = new SolidColorBrush(Color.FromRgb(244, 67, 54)),
                                StrokeThickness = 1,
                                StrokeDashArray = new DoubleCollection { 4, 3 }
                        };
                        rsiCanvas.Children.Add(overboughtLine);

                        // 超賣(30)虛線
                        var oversoldY = rsiToY(30);
                        var oversoldLine = new Line
                        {
                                X1 = left,
                                Y1 = oversoldY,
                                X2 = right,
                                Y2 = oversoldY,
                                Stroke = new SolidColorBrush(Color.FromRgb(76, 175, 80)),
                                StrokeThickness = 1,
                                StrokeDashArray = new DoubleCollection { 4, 3 }
                        };
                        rsiCanvas.Children.Add(oversoldLine);

                        // RSI 主線
                        var plotWidth = Math.Max(right - left, 10);
                        var spacing = plotWidth / series.Count;
                        var rsiLine = new Polyline
                        {
                                Stroke = new SolidColorBrush(Color.FromRgb(33, 150, 243)),
                                StrokeThickness = 1.6
                        };

                        for (int i = 0; i < series.Count; i++)
                        {
                                var x = left + i * spacing + spacing / 2;
                                var y = rsiToY(Math.Max(0, Math.Min(100, series[i])));
                                rsiLine.Points.Add(new Point(x, y));
                        }
                        rsiCanvas.Children.Add(rsiLine);

                        // 標籤
                        var title = new TextBlock
                        {
                                Text = "RSI (14) / 超買70 / 超賣30",
                                FontSize = 10,
                                Foreground = new SolidColorBrush(Colors.DimGray)
                        };
                        Canvas.SetLeft(title, 4);
                        Canvas.SetTop(title, 2);
                        rsiCanvas.Children.Add(title);

                        var label70 = new TextBlock
                        {
                                Text = "70",
                                FontSize = 9,
                                Foreground = new SolidColorBrush(Color.FromRgb(244, 67, 54))
                        };
                        Canvas.SetLeft(label70, 4);
                        Canvas.SetTop(label70, overboughtY - 8);
                        rsiCanvas.Children.Add(label70);

                        var label30 = new TextBlock
                        {
                                Text = "30",
                                FontSize = 9,
                                Foreground = new SolidColorBrush(Color.FromRgb(76, 175, 80))
                        };
                        Canvas.SetLeft(label30, 4);
                        Canvas.SetTop(label30, oversoldY - 8);
                        rsiCanvas.Children.Add(label30);
                }

                private void DrawIndicatorChart(Canvas canvas, List<double> values, string label, Color lineColor)
                {
                        canvas.Children.Clear();

                        if (values == null || values.Count < 2)
                        {
                                return;
                        }

                        var width = canvas.ActualWidth > 0 ? canvas.ActualWidth : 800;
                        var height = canvas.ActualHeight > 0 ? canvas.ActualHeight : 120;
                        var left = 40.0;
                        var right = width - 15;
                        var top = 10.0;
                        var bottom = height - 22;

                        var yAxis = new Line
                        {
                                X1 = left,
                                Y1 = top,
                                X2 = left,
                                Y2 = bottom,
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1
                        };
                        canvas.Children.Add(yAxis);

                        var xAxis = new Line
                        {
                                X1 = left,
                                Y1 = bottom,
                                X2 = right,
                                Y2 = bottom,
                                Stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
                                StrokeThickness = 1
                        };
                        canvas.Children.Add(xAxis);

                        var plotWidth = Math.Max(right - left, 10);
                        var candleSpacing = plotWidth / values.Count;
                        var candleWidth = Math.Max(candleSpacing * 0.6, 1);

                        DrawCurve(canvas, values, left, top, right, bottom, candleSpacing, candleWidth, lineColor, 1.4);

                        var title = new TextBlock
                        {
                                Text = label,
                                FontSize = 10,
                                Foreground = new SolidColorBrush(Colors.DimGray)
                        };
                        Canvas.SetLeft(title, 4);
                        Canvas.SetTop(title, 2);
                        canvas.Children.Add(title);
                }

                private void DrawCurve(Canvas canvas, List<double> values, double chartLeft, double chartTop, double chartRight, double chartBottom,
                        double candleSpacing, double candleWidth, Color color, double thickness)
                {
                        if (values == null || values.Count < 2)
                        {
                                return;
                        }

                        var min = values.Min();
                        var max = values.Max();
                        var range = max - min;
                        if (range == 0)
                        {
                                range = 1;
                        }

                        var line = new Polyline
                        {
                                Stroke = new SolidColorBrush(color),
                                StrokeThickness = thickness,
                                Opacity = 0.95
                        };

                        for (int i = 0; i < values.Count; i++)
                        {
                                var x = chartLeft + i * candleSpacing + candleWidth / 2;
                                var normalized = (values[i] - min) / range;
                                var y = chartBottom - normalized * (chartBottom - chartTop);
                                line.Points.Add(new Point(x, y));
                        }

                        canvas.Children.Add(line);
                }

                private Tuple<List<double>, List<double>, List<double>> BuildMACDComponents(List<double> closes)
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

                private List<double> BuildRSISeries(List<double> closes, int period)
                {
                        var rsiSeries = new List<double>();

                        if (closes == null || closes.Count == 0)
                        {
                                return rsiSeries;
                        }

                        for (int i = 0; i < closes.Count; i++)
                        {
                                var slice = closes.Take(i + 1).ToList();
                                if (slice.Count < period + 1)
                                {
                                        rsiSeries.Add(50);
                                        continue;
                                }

                                rsiSeries.Add(CalculateRSI(slice, period));
                        }

                        return rsiSeries;
                }

                private List<double> BuildMASeries(List<double> prices, int period)
                {
                        var series = new List<double>();
                        if (prices == null || prices.Count == 0)
                        {
                                return series;
                        }

                        for (int i = 0; i < prices.Count; i++)
                        {
                                if (i < period - 1)
                                {
                                        series.Add(double.NaN);
                                        continue;
                                }

                                var avg = prices.Skip(i - period + 1).Take(period).Average();
                                series.Add(avg);
                        }

                        return series;
                }

                private void DrawMovingAverageLine(Canvas canvas, List<double> values, double chartLeft, double chartTop, double chartBottom,
                        double candleSpacing, double candleWidth, double minPrice, double priceRange, Color color, double thickness)
                {
                        if (values == null || values.Count < 2)
                        {
                                return;
                        }

                        Polyline segment = null;

                        for (int i = 0; i < values.Count; i++)
                        {
                                if (double.IsNaN(values[i]))
                                {
                                        if (segment != null)
                                        {
                                                canvas.Children.Add(segment);
                                                segment = null;
                                        }
                                        continue;
                                }

                                if (segment == null)
                                {
                                        segment = new Polyline
                                        {
                                                Stroke = new SolidColorBrush(color),
                                                StrokeThickness = thickness,
                                                Opacity = 0.95
                                        };
                                }

                                var x = chartLeft + i * candleSpacing + candleWidth / 2;
                                var y = chartBottom - ((values[i] - minPrice) / priceRange) * (chartBottom - chartTop);
                                segment.Points.Add(new Point(x, y));
                        }

                        if (segment != null)
                        {
                                canvas.Children.Add(segment);
                        }
                }

                private List<Tuple<int, string, string>> BuildBacktestSignals(List<CandlestickData> data)
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
                                var current = slice.Last();

                                var ma5 = CalculateMA(closes, 5);
                                var ma20 = CalculateMA(closes, 20);
                                var rsi = CalculateRSI(closes, 14);
                                var macd = BuildMACDComponents(closes);
                                var macdNow = macd.Item1.Last();
                                var signalNow = macd.Item2.Last();
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

                private void DrawBacktestSignals(Canvas canvas, List<CandlestickData> data, List<Tuple<int, string, string>> signals,
                        double chartLeft, double chartTop, double chartBottom, double candleSpacing, double candleWidth,
                        double minPrice, double priceRange)
                {
                        if (signals == null || signals.Count == 0)
                        {
                                return;
                        }

                        foreach (var signal in signals)
                        {
                                var index = signal.Item1;
                                if (index < 0 || index >= data.Count)
                                {
                                        continue;
                                }

                                var candle = data[index];
                                var x = chartLeft + index * candleSpacing + candleWidth / 2;
                                var highY = chartBottom - ((candle.High - minPrice) / priceRange) * (chartBottom - chartTop);
                                var lowY = chartBottom - ((candle.Low - minPrice) / priceRange) * (chartBottom - chartTop);

                                var isBuy = signal.Item2.Contains("BUY");
                                var isStrong = signal.Item2.StartsWith("STRONG_");
                                var marker = new TextBlock
                                {
                                        Text = isBuy ? (isStrong ? "B++" : "B") : (isStrong ? "S++" : "S"),
                                        FontSize = isStrong ? 12 : 11,
                                        FontWeight = FontWeights.Bold,
                                        Foreground = isBuy
                                                ? new SolidColorBrush(isStrong ? Color.FromRgb(27, 94, 32) : Color.FromRgb(46, 125, 50))
                                                : new SolidColorBrush(isStrong ? Color.FromRgb(183, 28, 28) : Color.FromRgb(198, 40, 40)),
                                        Background = isBuy
                                                ? new SolidColorBrush(isStrong ? Color.FromRgb(200, 230, 201) : Color.FromRgb(232, 245, 233))
                                                : new SolidColorBrush(isStrong ? Color.FromRgb(255, 205, 210) : Color.FromRgb(255, 235, 238)),
                                        Padding = new Thickness(3, 1, 3, 1)
                                };

                                ToolTipService.SetToolTip(marker,
                                        $"{signal.Item2} 建議\n日期: {candle.Date}\n收盤: ${candle.Close:F2}\n強度: {(isStrong ? "高" : "一般")}\n原因: {signal.Item3}");

                                Canvas.SetLeft(marker, x - 7);
                                Canvas.SetTop(marker, isBuy ? lowY + 4 : highY - 20);
                                canvas.Children.Add(marker);
                        }
                }

                private bool TryLoadHistoricalDataFromYFinance(string ticker, string period, List<CandlestickData> target)
                {
                        try
                        {
                                var scriptPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AppConfig.YFinanceScriptPath);
                                if (!File.Exists(scriptPath))
                                {
                                        scriptPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", AppConfig.YFinanceScriptPath));
                                }

                                if (!File.Exists(scriptPath))
                                {
                                        Console.WriteLine($"[趨勢視窗] 找不到 yfinance 腳本: {scriptPath}");
                                        return false;
                                }

                                var startInfo = new ProcessStartInfo
                                {
                                        FileName = AppConfig.PythonPath,
                                        Arguments = $"\"{scriptPath}\" {ticker} history {period}",
                                        UseShellExecute = false,
                                        RedirectStandardOutput = true,
                                        RedirectStandardError = true,
                                        CreateNoWindow = true,
                                        StandardOutputEncoding = System.Text.Encoding.UTF8
                                };

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        process.Start();
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();
                                        process.WaitForExit(15000);

                                        if (!string.IsNullOrWhiteSpace(error))
                                        {
                                                Console.WriteLine($"[趨勢視窗] yfinance history 錯誤: {error}");
                                        }

                                        var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                                        if (lines.Length == 0 || lines[0] != "HISTORY_OK")
                                        {
                                                Console.WriteLine("[趨勢視窗] yfinance history 無有效輸出");
                                                return false;
                                        }

                                        target.Clear();
                                        for (int i = 1; i < lines.Length; i++)
                                        {
                                                var parts = lines[i].Split('|');
                                                if (parts.Length < 6)
                                                {
                                                        continue;
                                                }

                                                DateTime date;
                                                double open, high, low, close;
                                                long volume;

                                                if (!DateTime.TryParse(parts[0], out date) ||
                                                    !double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out open) ||
                                                    !double.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out high) ||
                                                    !double.TryParse(parts[3], NumberStyles.Any, CultureInfo.InvariantCulture, out low) ||
                                                    !double.TryParse(parts[4], NumberStyles.Any, CultureInfo.InvariantCulture, out close) ||
                                                    !long.TryParse(parts[5], NumberStyles.Any, CultureInfo.InvariantCulture, out volume))
                                                {
                                                        continue;
                                                }

                                                var changeAmount = close - open;
                                                var changePercent = open != 0 ? (changeAmount / open) * 100 : 0;

                                                target.Add(new CandlestickData
                                                {
                                                        Date = date.ToString("MM/dd"),
                                                        Open = open,
                                                        High = high,
                                                        Low = low,
                                                        Close = close,
                                                        Volume = volume,
                                                        ChangeAmount = changeAmount,
                                                        ChangePercent = changePercent
                                                });
                                        }

                                        return target.Count > 0;
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[趨勢視窗] 讀取 yfinance 歷史資料失敗: {ex.Message}");
                                return false;
                        }
                }

                private void UpdateFundamentalAnalysis()
                {
                        if (txtFundamentalAnalysis == null)
                        {
                                return;
                        }

                        try
                        {
                                var scriptPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AppConfig.YFinanceScriptPath);
                                if (!File.Exists(scriptPath))
                                {
                                        scriptPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", AppConfig.YFinanceScriptPath));
                                }

                                if (!File.Exists(scriptPath))
                                {
                                        txtFundamentalAnalysis.Text = "無法載入財報分析：找不到 yfinance 腳本。";
                                        return;
                                }

                                var startInfo = new ProcessStartInfo
                                {
                                        FileName = AppConfig.PythonPath,
                                        Arguments = $"\"{scriptPath}\" {_ticker} fundamentals",
                                        UseShellExecute = false,
                                        RedirectStandardOutput = true,
                                        RedirectStandardError = true,
                                        CreateNoWindow = true,
                                        StandardOutputEncoding = System.Text.Encoding.UTF8
                                };

                                using (var process = new Process { StartInfo = startInfo })
                                {
                                        process.Start();
                                        var output = process.StandardOutput.ReadToEnd();
                                        var error = process.StandardError.ReadToEnd();
                                        process.WaitForExit(15000);

                                        if (!string.IsNullOrWhiteSpace(error))
                                        {
                                                Console.WriteLine($"[財報分析] yfinance 錯誤: {error}");
                                        }

                                        var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                                        if (lines.Length == 0 || lines[0] != "FUNDAMENTALS_OK")
                                        {
                                                txtFundamentalAnalysis.Text = "目前無法取得完整財報資料，請稍後再試。";
                                                return;
                                        }

                                        var metrics = new Dictionary<string, double>();
                                        for (int i = 1; i < lines.Length; i++)
                                        {
                                                var parts = lines[i].Split('|');
                                                if (parts.Length != 2)
                                                {
                                                        continue;
                                                }

                                                double value;
                                                if (double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                                                {
                                                        metrics[parts[0]] = value;
                                                }
                                        }

                                        var reasons = new List<string>();
                                        var metricDetails = new List<string>();
                                        var fScore = 50;

                                        double val;
                                        if (metrics.TryGetValue("trailingPE", out val))
                                        {
                                                metricDetails.Add($"PE(TTM): {val:F2}");
                                                if (val > 0 && val <= 25)
                                                {
                                                        fScore += 8;
                                                        reasons.Add($"本益比 PE={val:F1} 處於合理區間");
                                                }
                                                else if (val > 40)
                                                {
                                                        fScore -= 8;
                                                        reasons.Add($"本益比 PE={val:F1} 偏高，估值壓力較大");
                                                }
                                        }

                                        if (metrics.TryGetValue("earningsGrowth", out val))
                                        {
                                                metricDetails.Add($"獲利成長率: {val * 100:F2}%");
                                                if (val > 0.10)
                                                {
                                                        fScore += 10;
                                                        reasons.Add($"獲利成長率 {val * 100:F1}% 表現良好");
                                                }
                                                else if (val < 0)
                                                {
                                                        fScore -= 10;
                                                        reasons.Add($"獲利成長率 {val * 100:F1}% 為負，需留意");
                                                }
                                        }

                                        if (metrics.TryGetValue("revenueGrowth", out val))
                                        {
                                                metricDetails.Add($"營收成長率: {val * 100:F2}%");
                                                if (val > 0.08)
                                                {
                                                        fScore += 8;
                                                        reasons.Add($"營收成長率 {val * 100:F1}% 具支撐");
                                                }
                                                else if (val < 0)
                                                {
                                                        fScore -= 8;
                                                        reasons.Add($"營收成長率 {val * 100:F1}% 轉弱");
                                                }
                                        }

                                        if (metrics.TryGetValue("profitMargins", out val))
                                        {
                                                metricDetails.Add($"淨利率: {val * 100:F2}%");
                                                if (val > 0.15)
                                                {
                                                        fScore += 6;
                                                        reasons.Add($"淨利率 {val * 100:F1}% 健康");
                                                }
                                                else if (val < 0.05)
                                                {
                                                        fScore -= 6;
                                                        reasons.Add($"淨利率 {val * 100:F1}% 偏低");
                                                }
                                        }

                                        if (metrics.TryGetValue("debtToEquity", out val))
                                        {
                                                metricDetails.Add($"負債權益比(D/E): {val:F2}");
                                                if (val <= 100)
                                                {
                                                        fScore += 4;
                                                        reasons.Add($"負債權益比 {val:F1} 風險可控");
                                                }
                                                else
                                                {
                                                        fScore -= 4;
                                                        reasons.Add($"負債權益比 {val:F1} 偏高");
                                                }
                                        }

                                        if (metrics.TryGetValue("forwardPE", out val))
                                        {
                                                metricDetails.Add($"Forward PE: {val:F2}");
                                                if (val > 0 && val <= 22)
                                                {
                                                        fScore += 4;
                                                        reasons.Add($"預估本益比 {val:F1} 合理");
                                                }
                                                else if (val > 35)
                                                {
                                                        fScore -= 4;
                                                        reasons.Add($"預估本益比 {val:F1} 偏高");
                                                }
                                        }

                                        if (metrics.TryGetValue("returnOnEquity", out val))
                                        {
                                                metricDetails.Add($"ROE: {val * 100:F2}%");
                                                if (val >= 0.12)
                                                {
                                                        fScore += 6;
                                                        reasons.Add($"ROE {val * 100:F1}% 顯示資本效率良好");
                                                }
                                                else if (val < 0.05)
                                                {
                                                        fScore -= 6;
                                                        reasons.Add($"ROE {val * 100:F1}% 偏低");
                                                }
                                        }

                                        if (metrics.TryGetValue("currentRatio", out val))
                                        {
                                                metricDetails.Add($"流動比率: {val:F2}");
                                                if (val >= 1.2)
                                                {
                                                        fScore += 3;
                                                        reasons.Add($"流動比率 {val:F2}，短期償債能力尚可");
                                                }
                                                else if (val < 1.0)
                                                {
                                                        fScore -= 3;
                                                        reasons.Add($"流動比率 {val:F2} 偏低，需留意資金壓力");
                                                }
                                        }

                                        if (metrics.TryGetValue("marketCap", out val))
                                        {
                                                metricDetails.Add($"市值: {FormatVolume((long)val)}");
                                        }

                                        fScore = Math.Max(0, Math.Min(100, fScore));

                                        if (reasons.Count == 0)
                                        {
                                                txtFundamentalAnalysis.Text = "財報欄位不足，暫無法給出有效基本面結論。";
                                                return;
                                        }

                                        var level = fScore >= 70 ? "偏正向" : (fScore >= 50 ? "中性" : "偏保守");
                                        var action = fScore >= 75 ? "🟢 財報指標：偏買入" :
                                                     fScore >= 55 ? "🟡 財報指標：觀望/持有" :
                                                     "🔴 財報指標：偏賣出或保守";

                                        if (fScore >= 55)
                                        {
                                                txtFundamentalAnalysis.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1B5E20"));
                                        }
                                        else
                                        {
                                                txtFundamentalAnalysis.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B71C1C"));
                                        }

                                        txtFundamentalAnalysis.Text =
                                                $"財報評分：{fScore}/100（{level}）\n" +
                                                $"{action}\n\n" +
                                                "【核心財報數據】\n" +
                                                string.Join("\n", metricDetails.Select(m => $"• {m}")) +
                                                "\n\n【建議原因】\n" +
                                                string.Join("\n", reasons.Select(r => $"• {r}"));
                                }
                        }
                        catch (Exception ex)
                        {
                                txtFundamentalAnalysis.Text = $"財報分析失敗：{ex.Message}";
                        }
                }

                private void GenerateAIAnalysis(List<CandlestickData> data, double currentPrice, double? changePercent)
                {
                        if (data == null || data.Count < 2)
                        {
                                txtCurrentRecommendation.Text = "資料不足，無法產生交易建議";
                                return;
                        }

                        var closes = data.Select(d => d.Close).ToList();
                        var avgPrice = closes.Average();
                        var latest = data.Last();

                        // 綜合評分 (0-100)
                        int score = 50;
                        var reasons = new List<string>();

                        // 1) MA 結構
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

                        // 2) MACD（快慢線 + 柱狀體動能）
                        var macdTuple = BuildMACDComponents(closes);
                        var macd = macdTuple.Item1.Last();
                        var signal = macdTuple.Item2.Last();
                        var hist = macdTuple.Item3.Last();
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

                        // 3) RSI
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

                        // 4) 成交量確認
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

                        // 5) 價格相對平均位階
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

                        // 當日波動加權
                        if (changePercent.HasValue)
                        {
                                if (changePercent.Value >= 3) score += 2;
                                else if (changePercent.Value <= -3) score -= 2;
                        }

                        score = Math.Max(0, Math.Min(100, score));

                        txtScore.Text = score.ToString();
                        pbScore.Value = score;

                        if (score >= 70)
                        {
                                txtScoreLevel.Text = "★★★★★ 強烈買入";
                                txtScoreLevel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50"));
                                pbScore.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50"));

                                if (txtCurrentRecommendation != null)
                                {
                                        txtCurrentRecommendation.Text = $"偏多（買入）｜評分 {score}/100｜建議分批布局，止損參考 ${currentPrice * 0.95:F2}";
                                        txtCurrentRecommendation.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1B5E20"));
                                }

                                txtSuggestionTitle.Text = "💰 強烈買入建議";
                                suggestionBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F5E9"));
                                suggestionBorder.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4CAF50"));
                                txtSuggestion.Text = $"根據技術分析，該股票綜合評分達到 {score} 分（滿分100），處於強勢區間。\n\n" +
                                        $"建議理由：\n" +
                                        string.Join("\n", reasons.Select(r => $"• {r}")) +
                                        $"\n\n💡 操作建議：可考慮分批買入，設定止損點在 ${currentPrice * 0.95:F2}";
                        }
                        else if (score >= 50)
                        {
                                txtScoreLevel.Text = "★★★☆☆ 持有觀望";
                                txtScoreLevel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF9800"));
                                pbScore.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF9800"));

                                if (txtCurrentRecommendation != null)
                                {
                                        txtCurrentRecommendation.Text = $"中性（觀望）｜評分 {score}/100｜等待突破訊號再加碼";
                                        txtCurrentRecommendation.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E65100"));
                                }

                                txtSuggestionTitle.Text = "📊 持有觀望";
                                suggestionBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF3E0"));
                                suggestionBorder.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF9800"));
                                txtSuggestion.Text = $"根據技術分析，該股票綜合評分為 {score} 分，處於中性區間。\n\n" +
                                        $"市場狀況：\n" +
                                        string.Join("\n", reasons.Select(r => $"• {r}")) +
                                        $"\n\n💡 操作建議：已持有者建議持有觀望，未持有者等待更明確信號。";
                        }
                        else
                        {
                                txtScoreLevel.Text = "★☆☆☆☆ 建議賣出";
                                txtScoreLevel.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336"));
                                pbScore.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336"));

                                if (txtCurrentRecommendation != null)
                                {
                                        txtCurrentRecommendation.Text = $"偏空（減碼/賣出）｜評分 {score}/100｜控制風險，等待回穩再評估";
                                        txtCurrentRecommendation.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B71C1C"));
                                }

                                txtSuggestionTitle.Text = "⚠️ 建議賣出";
                                suggestionBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFEBEE"));
                                suggestionBorder.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F44336"));
                                txtSuggestion.Text = $"根據技術分析，該股票綜合評分僅 {score} 分，處於弱勢區間。\n\n" +
                                        $"風險因素：\n" +
                                        string.Join("\n", reasons.Select(r => $"• {r}")) +
                                        $"\n\n💡 操作建議：建議減倉或清倉，等待更好的進場時機。";
                        }

                        // 風險提示
                        txtRiskWarning.Text = "本分析僅供參考，不構成投資建議。股市有風險，投資需謹慎。\n\n" +
                                "⚠️ 重要提醒：\n" +
                                "• 優先使用 yfinance 歷史資料，若資料抓取失敗會回退到模擬資料\n" +
                                "• 請以實際市場數據和專業投資顧問建議為準\n" +
                                "• 投資前請充分評估自身風險承受能力\n" +
                                "• 建議設定止損點，控制投資風險";
                }

                private string GetPeriodDisplay(string period)
                {
                        switch (period)
                        {
                                case "1mo": return "1個月";
                                case "3mo": return "3個月";
                                case "6mo": return "6個月";
                                case "1y": return "1年";
                                case "2y": return "2年";
                                default: return period;
                        }
                }

                private int GetDaysForPeriod(string period)
                {
                        switch (period)
                        {
                                case "1mo": return 30;
                                case "3mo": return 90;
                                case "6mo": return 180;
                                case "1y": return 365;
                                case "2y": return 730;
                                default: return 90;
                        }
                }

                private string FormatVolume(long volume)
                {
                        if (volume >= 1000000000)
                                return $"{volume / 1000000000.0:F2}B";
                        else if (volume >= 1000000)
                                return $"{volume / 1000000.0:F2}M";
                        else if (volume >= 1000)
                                return $"{volume / 1000.0:F2}K";
                        else
                                return volume.ToString();
                }

                private async void CboPeriod_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                        try
                        {
                                // 只有在視窗完全載入後才處理週期變更
                                if (!_isLoaded)
                                {
                                        Console.WriteLine($"[趨勢視窗] ComboBox SelectionChanged 觸發，但視窗未完全載入，忽略");
                                        return;
                                }

                                if (cboPeriod?.SelectedItem is ComboBoxItem item)
                                {
                                        var newPeriod = item.Tag?.ToString();
                                        if (!string.IsNullOrEmpty(newPeriod) && newPeriod != _currentPeriod)
                                        {
                                                Console.WriteLine($"[趨勢視窗] 週期變更: {_currentPeriod} → {newPeriod}");
                                                _currentPeriod = newPeriod;
                                                await LoadDataAsync();
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[趨勢視窗錯誤] SelectionChanged 發生錯誤: {ex.Message}");
                        }
                }

                private void BtnTabChart_Click(object sender, RoutedEventArgs e)
                {
                        chartTab.Visibility = Visibility.Visible;
                        technicalTab.Visibility = Visibility.Collapsed;
                        analysisTab.Visibility = Visibility.Collapsed;

                        btnTabChart.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3F2FD"));
                        btnTabTechnical.Background = new SolidColorBrush(Colors.White);
                        btnTabAnalysis.Background = new SolidColorBrush(Colors.White);
                }

                private void BtnTabTechnical_Click(object sender, RoutedEventArgs e)
                {
                        chartTab.Visibility = Visibility.Collapsed;
                        technicalTab.Visibility = Visibility.Visible;
                        analysisTab.Visibility = Visibility.Collapsed;

                        btnTabChart.Background = new SolidColorBrush(Colors.White);
                        btnTabTechnical.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3F2FD"));
                        btnTabAnalysis.Background = new SolidColorBrush(Colors.White);
                }

                private void BtnTabAnalysis_Click(object sender, RoutedEventArgs e)
                {
                        chartTab.Visibility = Visibility.Collapsed;
                        technicalTab.Visibility = Visibility.Collapsed;
                        analysisTab.Visibility = Visibility.Visible;

                        btnTabChart.Background = new SolidColorBrush(Colors.White);
                        btnTabTechnical.Background = new SolidColorBrush(Colors.White);
                        btnTabAnalysis.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E3F2FD"));
                }

                private async void BtnRefresh_Click(object sender, RoutedEventArgs e)
                {
                        await LoadDataAsync();
                }

                private void BtnClose_Click(object sender, RoutedEventArgs e)
                {
                        HideLoadingWindow();
                        Close();
                }
        }

        internal class LoadingProgressPopup : Window
        {
                private readonly ProgressBar _progressBar;
                private readonly TextBlock _statusText;

                public LoadingProgressPopup()
                {
                        Title = "載入中";
                        Width = 360;
                        Height = 120;
                        WindowStartupLocation = WindowStartupLocation.CenterOwner;
                        ResizeMode = ResizeMode.NoResize;
                        WindowStyle = WindowStyle.ToolWindow;
                        ShowInTaskbar = false;
                        Topmost = true;

                        var panel = new StackPanel
                        {
                                Margin = new Thickness(16)
                        };

                        _statusText = new TextBlock
                        {
                                Text = "準備中...",
                                FontSize = 13,
                                Margin = new Thickness(0, 0, 0, 10),
                                Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2F3B45"))
                        };

                        _progressBar = new ProgressBar
                        {
                                Minimum = 0,
                                Maximum = 100,
                                Height = 16,
                                Value = 0
                        };

                        panel.Children.Add(_statusText);
                        panel.Children.Add(_progressBar);
                        Content = panel;
                }

                public void UpdateProgress(double progress, string message)
                {
                        _progressBar.Value = Math.Max(0, Math.Min(100, progress));
                        _statusText.Text = message;
                }
        }

        public class CandlestickData
        {
                public string Date { get; set; }
                public double Open { get; set; }
                public double High { get; set; }
                public double Low { get; set; }
                public double Close { get; set; }
                public long Volume { get; set; }

                // 漲跌幅相關屬性
                public double ChangeAmount { get; set; }  // 漲跌金額 (收盤 - 開盤)
                public double ChangePercent { get; set; } // 漲跌幅百分比
                public bool IsPositive => ChangeAmount >= 0; // 是否上漲
        }

        public class MajorNewsItem
        {
                public string Time { get; set; }
                public string Title { get; set; }
                public string ImpactLevel { get; set; }
        }

        public class PriceHistoryItem
        {
                public string Date { get; set; }
                public double Price { get; set; }
                public double BarWidth { get; set; }
                public bool IsPositive { get; set; }
        }
}
