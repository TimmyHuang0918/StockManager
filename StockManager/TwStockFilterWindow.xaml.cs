using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Web.Script.Serialization;
using StockManager.Config;
using StockManager.Converters;
using StockManager.Models;
using StockManager.Services;
using IOPath = System.IO.Path;

namespace StockManager
{
    public partial class TwStockFilterWindow : Window
    {
        private static readonly string[] YahooSectorFallback = new[]
        {
            "水泥", "食品", "塑膠", "紡織", "電機機械", "電器電纜", "化學", "生技", "玻璃", "造紙", "鋼鐵", "橡膠",
            "汽車", "半導體", "電腦週邊", "光電", "通訊網路", "電子零組件", "電子通路", "資訊服務", "其他電子",
            "營建", "航運", "觀光餐旅", "金融業", "貿易百貨", "油電燃氣", "運動休閒", "數位雲端", "綠能環保", "其他"
        };

        private readonly List<StockInfo> _sourceStocks = new List<StockInfo>();
        private readonly List<SectorSampleItem> _sectorSamples = new List<SectorSampleItem>();
        private readonly ObservableCollection<TwFilterStockItem> _filteredStocks = new ObservableCollection<TwFilterStockItem>();
        private readonly List<string> _csvSectorOrder = new List<string>();
        private readonly Dictionary<string, string> _tickerToCsvSector = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> _yahooListedSectors = new List<string>();
        private readonly Dictionary<string, string> _tickerToYahooSector = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly string _twYFinanceSectorCacheFile = IOPath.Combine(AppConfig.UserConfigDir, "tw_sector_yfinance_cache.json");
        private System.Windows.Threading.DispatcherTimer _autoYFinanceUpdateTimer;
        private bool _isAutoYFinanceUpdating;
        private DateTime _lastAutoYFinanceUpdateDate = DateTime.MinValue;
        private DateTime? _lastYFinanceCacheUpdatedAt;
        private bool _loadedFromCsv;
        private string _noPriceSummary = string.Empty;

        public TwStockFilterWindow()
        {
            InitializeComponent();
            dgFilteredTwStocks.ItemsSource = _filteredStocks;
            Loaded += TwStockFilterWindow_Loaded;
            Closed += TwStockFilterWindow_Closed;
        }

        private async void TwStockFilterWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await System.Threading.Tasks.Task.Run(() => LoadAllListedTwStocks());
            ApplyFilter();
        }

        private void TwStockFilterWindow_Closed(object sender, EventArgs e)
        {
            if (_autoYFinanceUpdateTimer != null)
            {
                _autoYFinanceUpdateTimer.Stop();
                _autoYFinanceUpdateTimer = null;
            }
        }

        private void BtnApplyFilter_Click(object sender, RoutedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            var hasMinPriceInput = !string.IsNullOrWhiteSpace(txtMinPrice.Text);
            var hasMinChangeInput = !string.IsNullOrWhiteSpace(txtMinChange.Text);
            var hasMaxChangeInput = !string.IsNullOrWhiteSpace(txtMaxChange.Text);

            var minPrice = ParseNumeric(txtMinPrice.Text);
            var minChange = ParseNumeric(txtMinChange.Text);
            var maxChange = ParseNumeric(txtMaxChange.Text);

            if (hasMinPriceInput && !minPrice.HasValue)
            {
                MessageBox.Show("最低價格格式錯誤，請輸入數字。", "篩選條件錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtMinPrice.Focus();
                return;
            }

            if (hasMinChangeInput && !minChange.HasValue)
            {
                MessageBox.Show("最小漲跌幅格式錯誤，請輸入數字。", "篩選條件錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtMinChange.Focus();
                return;
            }

            if (hasMaxChangeInput && !maxChange.HasValue)
            {
                MessageBox.Show("最大漲跌幅格式錯誤，請輸入數字。", "篩選條件錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtMaxChange.Focus();
                return;
            }

            if (minChange.HasValue && maxChange.HasValue && minChange.Value > maxChange.Value)
            {
                MessageBox.Show("最小漲跌幅不能大於最大漲跌幅。", "篩選條件錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtMinChange.Focus();
                return;
            }

            var onlyRising = chkOnlyRising.IsChecked == true;

            IEnumerable<StockInfo> query = _sourceStocks;

            if (minPrice.HasValue)
            {
                query = query.Where(x => x.Price.HasValue && x.Price.Value >= minPrice.Value);
            }

            if (minChange.HasValue)
            {
                query = query.Where(x => x.ChangePercent.HasValue && x.ChangePercent.Value >= minChange.Value);
            }

            if (maxChange.HasValue)
            {
                query = query.Where(x => x.ChangePercent.HasValue && x.ChangePercent.Value <= maxChange.Value);
            }

            if (onlyRising)
            {
                query = query.Where(x => x.ChangePercent.HasValue && x.ChangePercent.Value > 0);
            }

            var result = query
                .Select(x =>
                {
                    var score = CalculateTradingScore(x);
                    return new TwFilterStockItem
                    {
                        Ticker = x.Ticker,
                        Name = x.Name,
                        Price = x.Price,
                        ChangePercent = x.ChangePercent,
                        PreviousClose = x.PreviousClose,
                        UpdatedAt = x.UpdatedAt,
                        TradingScore = score,
                        TradingSuggestion = GetTradingSuggestion(score)
                    };
                })
                .OrderByDescending(x => x.TradingScore)
                .ThenByDescending(x => x.ChangePercent ?? double.MinValue)
                .ToList();

            _filteredStocks.Clear();
            foreach (var stock in result)
            {
                _filteredStocks.Add(stock);
            }

            Title = $"台股篩選（{_filteredStocks.Count} 筆）";
        }

        private async void BtnSectorOverview_Click(object sender, RoutedEventArgs e)
        {
            if (_sectorSamples.Count == 0)
            {
                MessageBox.Show("尚未載入族群資料。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            await EnsureSectorYFinanceDataAsync();

            var groupedAll = _sectorSamples
                .GroupBy(x => MapSectorCategory(x.Industry, x.Name, x.Ticker))
                .ToDictionary(g => g.Key, g => g.ToList());

            var grouped = _sectorSamples
                .Where(x => x.ChangePercent.HasValue)
                .GroupBy(x => MapSectorCategory(x.Industry, x.Name, x.Ticker))
                .ToDictionary(g => g.Key, g => g.ToList());

            var summary = new List<SectorTrendItem>();
            var sectorFramework = _csvSectorOrder.Count > 0
                ? _csvSectorOrder
                : (_yahooListedSectors.Count > 0 ? _yahooListedSectors : YahooSectorFallback.ToList());
            if (_sectorSamples.Any(x => string.Equals(x.Industry, "ETF", StringComparison.OrdinalIgnoreCase))
                && !sectorFramework.Contains("ETF"))
            {
                sectorFramework = sectorFramework.Concat(new[] { "ETF" }).ToList();
            }
            foreach (var sector in sectorFramework)
            {
                List<SectorSampleItem> bucket;
                if (!grouped.TryGetValue(sector, out bucket) || bucket.Count == 0)
                {
                    summary.Add(new SectorTrendItem
                    {
                        Industry = sector,
                        AverageChangePercent = 0,
                        RisingCount = 0,
                        FallingCount = 0,
                        Trend = "無資料"
                    });
                    continue;
                }

                var avg = bucket.Average(x => x.ChangePercent ?? 0);
                summary.Add(new SectorTrendItem
                {
                    Industry = sector,
                    AverageChangePercent = avg,
                    RisingCount = bucket.Count(x => (x.ChangePercent ?? 0) > 0),
                    FallingCount = bucket.Count(x => (x.ChangePercent ?? 0) < 0),
                    Trend = avg >= 0 ? "上漲" : "下跌"
                });
            }

            var grid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserAddRows = false,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                AlternatingRowBackground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#F7FAFF"),
                RowBackground = System.Windows.Media.Brushes.White,
                GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
                HorizontalGridLinesBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#E6ECF2"),
                VerticalGridLinesBrush = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Background = System.Windows.Media.Brushes.White,
                ItemsSource = summary,
                Margin = new Thickness(0)
            };

            grid.ColumnHeaderStyle = new Style(typeof(Control))
            {
                Setters =
                {
                    new Setter(Control.BackgroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2C3E50")),
                    new Setter(Control.ForegroundProperty, System.Windows.Media.Brushes.White),
                    new Setter(Control.FontWeightProperty, FontWeights.Bold),
                    new Setter(Control.BorderBrushProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2C3E50")),
                    new Setter(Control.PaddingProperty, new Thickness(8, 6, 8, 6))
                }
            };

            grid.CellStyle = new Style(typeof(DataGridCell))
            {
                Setters =
                {
                    new Setter(Control.BorderThicknessProperty, new Thickness(0)),
                    new Setter(Control.PaddingProperty, new Thickness(8, 5, 8, 5))
                }
            };

            grid.Columns.Add(new DataGridTextColumn { Header = "族群", Binding = new System.Windows.Data.Binding("Industry"), Width = 180 });
            grid.Columns.Add(new DataGridTextColumn { Header = "平均漲跌幅(%)", Binding = new System.Windows.Data.Binding("AverageChangePercent") { StringFormat = "{0:F2}" }, Width = 120 });
            grid.Columns.Add(new DataGridTextColumn { Header = "上漲家數", Binding = new System.Windows.Data.Binding("RisingCount"), Width = 90 });
            grid.Columns.Add(new DataGridTextColumn { Header = "下跌家數", Binding = new System.Windows.Data.Binding("FallingCount"), Width = 90 });
            var trendColumn = new DataGridTemplateColumn { Header = "整體方向", Width = 90 };
            var trendText = new FrameworkElementFactory(typeof(TextBlock));
            trendText.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("Trend"));
            trendText.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold);
            trendText.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);

            var trendStyle = new Style(typeof(TextBlock));
            trendStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Trend"),
                Value = "上漲",
                Setters = { new Setter(TextBlock.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#E74C3C")) }
            });
            trendStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Trend"),
                Value = "下跌",
                Setters = { new Setter(TextBlock.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#27AE60")) }
            });
            trendStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Trend"),
                Value = "無資料",
                Setters = { new Setter(TextBlock.ForegroundProperty, System.Windows.Media.Brushes.Gray) }
            });
            trendText.SetValue(TextBlock.StyleProperty, trendStyle);

            trendColumn.CellTemplate = new DataTemplate { VisualTree = trendText };
            grid.Columns.Add(trendColumn);

            grid.SelectionChanged += (s, args) =>
            {
                var selected = grid.SelectedItem as SectorTrendItem;
                if (selected == null)
                {
                    return;
                }

                List<SectorSampleItem> bucket;
                if (!groupedAll.TryGetValue(selected.Industry, out bucket))
                {
                    bucket = new List<SectorSampleItem>();
                }

                ShowSectorStockDetails(selected.Industry, bucket);
                grid.SelectedItem = null;
            };

            var win = new Window
            {
                Title = "台股族群漲跌總覽（點擊族群看成分股）",
                Owner = this,
                Width = 760,
                Height = 560,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#F3F6FA")
            };

            var layoutRoot = new Grid { Margin = new Thickness(12) };
            layoutRoot.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layoutRoot.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            var headerPanel = new Border
            {
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2C3E50"),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 0, 0, 10),
                Child = new TextBlock
                {
                    Text = "📊 台股族群漲跌總覽（點擊族群看成分股）",
                    Foreground = System.Windows.Media.Brushes.White,
                    FontWeight = FontWeights.Bold,
                    FontSize = 14
                }
            };

            var contentBorder = new Border
            {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#D9E2EC"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(8),
                Child = grid
            };

            Grid.SetRow(headerPanel, 0);
            Grid.SetRow(contentBorder, 1);
            layoutRoot.Children.Add(headerPanel);
            layoutRoot.Children.Add(contentBorder);
            win.Content = layoutRoot;

            win.ShowDialog();
        }

        private void ShowSectorStockDetails(string sectorName, List<SectorSampleItem> sectorStocks)
        {
            var stockLookup = _sourceStocks.ToDictionary(x => x.Ticker, StringComparer.OrdinalIgnoreCase);
            var detailItems = new List<SectorStockDetailItem>();

            foreach (var sample in sectorStocks)
            {
                StockInfo stock;
                if (!stockLookup.TryGetValue(sample.Ticker, out stock))
                {
                    continue;
                }

                double? changeAmount = null;
                if (stock.Price.HasValue && stock.PreviousClose.HasValue)
                {
                    changeAmount = stock.Price.Value - stock.PreviousClose.Value;
                }

                detailItems.Add(new SectorStockDetailItem
                {
                    Ticker = stock.Ticker,
                    Name = stock.Name,
                    Price = stock.Price,
                    ChangeAmount = changeAmount,
                    ChangePercent = stock.ChangePercent
                });
            }

            var sorted = detailItems
                .OrderByDescending(x => x.ChangePercent ?? double.MinValue)
                .ThenBy(x => x.Ticker)
                .ToList();

            var grid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserAddRows = false,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                AlternatingRowBackground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#F7FAFF"),
                RowBackground = System.Windows.Media.Brushes.White,
                GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
                HorizontalGridLinesBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#E6ECF2"),
                VerticalGridLinesBrush = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Background = System.Windows.Media.Brushes.White,
                ItemsSource = sorted,
                Margin = new Thickness(0)
            };

            grid.ColumnHeaderStyle = new Style(typeof(Control))
            {
                Setters =
                {
                    new Setter(Control.BackgroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2C3E50")),
                    new Setter(Control.ForegroundProperty, System.Windows.Media.Brushes.White),
                    new Setter(Control.FontWeightProperty, FontWeights.Bold),
                    new Setter(Control.BorderBrushProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2C3E50")),
                    new Setter(Control.PaddingProperty, new Thickness(8, 6, 8, 6))
                }
            };

            grid.CellStyle = new Style(typeof(DataGridCell))
            {
                Setters =
                {
                    new Setter(Control.BorderThicknessProperty, new Thickness(0)),
                    new Setter(Control.PaddingProperty, new Thickness(8, 5, 8, 5))
                }
            };

            grid.Columns.Add(new DataGridTextColumn { Header = "股票代號", Binding = new System.Windows.Data.Binding("Ticker"), Width = 110 });
            grid.Columns.Add(new DataGridTextColumn { Header = "公司名稱", Binding = new System.Windows.Data.Binding("Name"), Width = 180 });
            grid.Columns.Add(new DataGridTextColumn { Header = "股價", Binding = new System.Windows.Data.Binding("Price") { StringFormat = "{0:F2}" }, Width = 110 });

            var positiveNegativeConverter = new PositiveNegativeConverter();

            var changeAmountColumn = new DataGridTemplateColumn { Header = "漲跌", Width = 100 };
            var changeAmountText = new FrameworkElementFactory(typeof(TextBlock));
            changeAmountText.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("ChangeAmount") { StringFormat = "{0:F2}" });
            changeAmountText.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold);
            changeAmountText.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);

            var changeAmountStyle = new Style(typeof(TextBlock));
            changeAmountStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("ChangeAmount") { Converter = positiveNegativeConverter },
                Value = "Positive",
                Setters =
                {
                    new Setter(TextBlock.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#E74C3C")),
                    new Setter(TextBlock.TextProperty, new System.Windows.Data.Binding("ChangeAmount") { StringFormat = "+{0:F2}" })
                }
            });
            changeAmountStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("ChangeAmount") { Converter = positiveNegativeConverter },
                Value = "Negative",
                Setters =
                {
                    new Setter(TextBlock.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#27AE60")),
                    new Setter(TextBlock.TextProperty, new System.Windows.Data.Binding("ChangeAmount") { StringFormat = "{0:F2}" })
                }
            });
            changeAmountStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("ChangeAmount"),
                Value = null,
                Setters =
                {
                    new Setter(TextBlock.ForegroundProperty, System.Windows.Media.Brushes.Gray),
                    new Setter(TextBlock.TextProperty, "--"),
                    new Setter(TextBlock.FontWeightProperty, FontWeights.Normal)
                }
            });
            changeAmountText.SetValue(TextBlock.StyleProperty, changeAmountStyle);
            changeAmountColumn.CellTemplate = new DataTemplate { VisualTree = changeAmountText };
            grid.Columns.Add(changeAmountColumn);

            var changePercentColumn = new DataGridTemplateColumn { Header = "漲跌幅(%)", Width = 110 };
            var changePercentText = new FrameworkElementFactory(typeof(TextBlock));
            changePercentText.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("ChangePercent") { StringFormat = "{0:F2}" });
            changePercentText.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold);
            changePercentText.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);

            var changePercentStyle = new Style(typeof(TextBlock));
            changePercentStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("ChangePercent") { Converter = positiveNegativeConverter },
                Value = "Positive",
                Setters =
                {
                    new Setter(TextBlock.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#E74C3C")),
                    new Setter(TextBlock.TextProperty, new System.Windows.Data.Binding("ChangePercent") { StringFormat = "▲ {0:F2}" })
                }
            });
            changePercentStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("ChangePercent") { Converter = positiveNegativeConverter },
                Value = "Negative",
                Setters =
                {
                    new Setter(TextBlock.ForegroundProperty, (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#27AE60")),
                    new Setter(TextBlock.TextProperty, new System.Windows.Data.Binding("ChangePercent") { StringFormat = "▼ {0:F2}" })
                }
            });
            changePercentStyle.Triggers.Add(new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("ChangePercent"),
                Value = null,
                Setters =
                {
                    new Setter(TextBlock.ForegroundProperty, System.Windows.Media.Brushes.Gray),
                    new Setter(TextBlock.TextProperty, "--"),
                    new Setter(TextBlock.FontWeightProperty, FontWeights.Normal)
                }
            });
            changePercentText.SetValue(TextBlock.StyleProperty, changePercentStyle);
            changePercentColumn.CellTemplate = new DataTemplate { VisualTree = changePercentText };
            grid.Columns.Add(changePercentColumn);

            grid.PreviewMouseRightButtonDown += (s, e) =>
            {
                var source = e.OriginalSource as DependencyObject;
                while (source != null && !(source is DataGridRow))
                {
                    source = VisualTreeHelper.GetParent(source);
                }

                var row = source as DataGridRow;
                if (row != null)
                {
                    row.IsSelected = true;
                    row.Focus();
                }
            };

            var contextMenu = new ContextMenu();
            var addToMainListMenuItem = new MenuItem { Header = "加入主頁面台股清單" };
            addToMainListMenuItem.Click += (s, e) =>
            {
                var selectedStock = grid.SelectedItem as SectorStockDetailItem;
                if (selectedStock == null)
                {
                    return;
                }

                var mainWindow = Owner as MainWindow;
                if (mainWindow == null)
                {
                    MessageBox.Show("找不到主頁面視窗，無法加入清單。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string message;
                var success = mainWindow.TryAddTwStockFromFilter(selectedStock.Ticker, selectedStock.Name, out message);
                MessageBox.Show(message, success ? "成功" : "提示", MessageBoxButton.OK,
                    success ? MessageBoxImage.Information : MessageBoxImage.Warning);
            };

            contextMenu.Items.Add(addToMainListMenuItem);
            grid.ContextMenu = contextMenu;

            var detailWindow = new Window
            {
                Title = $"{sectorName} 成分股（{sorted.Count} 檔）",
                Owner = this,
                Width = 900,
                Height = 760,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#F3F6FA")
            };

            var layoutRoot = new Grid { Margin = new Thickness(12) };
            layoutRoot.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            layoutRoot.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            layoutRoot.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var headerPanel = new Border
            {
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2C3E50"),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 0, 0, 10)
            };

            var headerTextBlock = new TextBlock
            {
                Text = $"📈 {sectorName} 成分股（{sorted.Count} 檔）｜載入 yfinance 中...",
                Foreground = System.Windows.Media.Brushes.White,
                FontWeight = FontWeights.Bold,
                FontSize = 14
            };

            var progressTextBlock = new TextBlock
            {
                Text = "yfinance 更新進度：0/0",
                Foreground = System.Windows.Media.Brushes.White,
                FontSize = 11,
                Margin = new Thickness(0, 6, 0, 4),
                Opacity = 0.9
            };

            var progressBar = new ProgressBar
            {
                Height = 8,
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#4FC3F7"),
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#455A64")
            };

            var headerStack = new StackPanel();
            headerStack.Children.Add(headerTextBlock);
            headerStack.Children.Add(progressTextBlock);
            headerStack.Children.Add(progressBar);
            headerPanel.Child = headerStack;

            var contentBorder = new Border
            {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#D9E2EC"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(8),
                Child = grid
            };

            var realtimeToggleButton = new Button
            {
                Content = "即時模式：關閉",
                Width = 130,
                Height = 30,
                Margin = new Thickness(0, 0, 10, 0),
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#546E7A"),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0)
            };

            var realtimeStatusText = new TextBlock
            {
                Text = "K線：請在列表中選擇股票",
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#455A64")
            };

            var periodTodayButton = new Button
            {
                Content = "當日",
                Width = 56,
                Height = 28,
                Margin = new Thickness(0, 0, 6, 0),
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#90A4AE"),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0)
            };

            var period1MButton = new Button
            {
                Content = "1月",
                Width = 56,
                Height = 28,
                Margin = new Thickness(0, 0, 6, 0),
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#90A4AE"),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0)
            };

            var period3MButton = new Button
            {
                Content = "3月",
                Width = 56,
                Height = 28,
                Margin = new Thickness(0, 0, 6, 0),
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#1976D2"),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0)
            };

            var period6MButton = new Button
            {
                Content = "6月",
                Width = 56,
                Height = 28,
                Margin = new Thickness(0, 0, 6, 0),
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#90A4AE"),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0)
            };

            var period1YButton = new Button
            {
                Content = "1年",
                Width = 56,
                Height = 28,
                Margin = new Thickness(0, 0, 10, 0),
                Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#90A4AE"),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0)
            };

            var controlPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 10, 0, 6)
            };
            controlPanel.Children.Add(realtimeToggleButton);
            controlPanel.Children.Add(periodTodayButton);
            controlPanel.Children.Add(period1MButton);
            controlPanel.Children.Add(period3MButton);
            controlPanel.Children.Add(period6MButton);
            controlPanel.Children.Add(period1YButton);
            controlPanel.Children.Add(realtimeStatusText);

            var klineCanvas = new Canvas
            {
                Height = 250,
                Background = Brushes.White,
                ClipToBounds = true
            };

            var chartBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#D9E2EC"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(8),
                Child = new StackPanel
                {
                    Children =
                    {
                        controlPanel,
                        klineCanvas
                    }
                }
            };

            var isRealtimeMode = false;
            var isRealtimeRefreshing = false;
            var realtimeTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(20) };
            var currentKLinePeriod = "3mo";
            var currentKLineInterval = "1d";

            Action updatePeriodButtonStyles = () =>
            {
                var normalBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#90A4AE");
                var activeBrush = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#1976D2");

                periodTodayButton.Background = currentKLinePeriod == "1d" ? activeBrush : normalBrush;
                period1MButton.Background = currentKLinePeriod == "1mo" ? activeBrush : normalBrush;
                period3MButton.Background = currentKLinePeriod == "3mo" ? activeBrush : normalBrush;
                period6MButton.Background = currentKLinePeriod == "6mo" ? activeBrush : normalBrush;
                period1YButton.Background = currentKLinePeriod == "1y" ? activeBrush : normalBrush;
            };

            Action refreshKLine = () =>
            {
                var selected = grid.SelectedItem as SectorStockDetailItem ?? sorted.FirstOrDefault();
                if (selected == null)
                {
                    return;
                }

                var history = new List<KLinePoint>();
                var ok = TryLoadHistoricalDataFromYFinanceForKLine(NormalizeTwTicker(selected.Ticker), currentKLinePeriod, currentKLineInterval, history);
                if (!ok)
                {
                    klineCanvas.Children.Clear();
                    var msg = new TextBlock
                    {
                        Text = $"{selected.Ticker} 無法載入 K 線資料",
                        Foreground = Brushes.Gray
                    };
                    Canvas.SetLeft(msg, 10);
                    Canvas.SetTop(msg, 10);
                    klineCanvas.Children.Add(msg);
                    return;
                }

                DrawKLineChart(klineCanvas, history, $"{selected.Ticker} {selected.Name}");
                realtimeStatusText.Text = isRealtimeMode
                    ? $"即時模式中（20秒）｜K線：{selected.Ticker}"
                    : $"K線：{selected.Ticker}";
            };

            grid.SelectionChanged += (s, e) => refreshKLine();

            periodTodayButton.Click += (s, e) =>
            {
                currentKLinePeriod = "1d";
                currentKLineInterval = "5m";
                updatePeriodButtonStyles();
                refreshKLine();
            };
            period1MButton.Click += (s, e) =>
            {
                currentKLinePeriod = "1mo";
                currentKLineInterval = "1d";
                updatePeriodButtonStyles();
                refreshKLine();
            };
            period3MButton.Click += (s, e) =>
            {
                currentKLinePeriod = "3mo";
                currentKLineInterval = "1d";
                updatePeriodButtonStyles();
                refreshKLine();
            };
            period6MButton.Click += (s, e) =>
            {
                currentKLinePeriod = "6mo";
                currentKLineInterval = "1d";
                updatePeriodButtonStyles();
                refreshKLine();
            };
            period1YButton.Click += (s, e) =>
            {
                currentKLinePeriod = "1y";
                currentKLineInterval = "1d";
                updatePeriodButtonStyles();
                refreshKLine();
            };

            realtimeTimer.Tick += async (s, e) =>
            {
                if (!isRealtimeMode || isRealtimeRefreshing || !detailWindow.IsLoaded)
                {
                    return;
                }

                isRealtimeRefreshing = true;
                try
                {
                    await System.Threading.Tasks.Task.Run(() => RefreshSectorStockRealtimePrices(sorted, null));
                    sorted.Sort((a, b) =>
                    {
                        var changeCompare = Nullable.Compare(b.ChangePercent, a.ChangePercent);
                        if (changeCompare != 0)
                        {
                            return changeCompare;
                        }

                        return string.Compare(a.Ticker, b.Ticker, StringComparison.OrdinalIgnoreCase);
                    });
                    grid.Items.Refresh();
                    refreshKLine();
                }
                finally
                {
                    isRealtimeRefreshing = false;
                }
            };

            realtimeToggleButton.Click += (s, e) =>
            {
                isRealtimeMode = !isRealtimeMode;
                if (isRealtimeMode)
                {
                    realtimeToggleButton.Content = "即時模式：開啟";
                    realtimeToggleButton.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#2E7D32");
                    realtimeTimer.Start();
                    refreshKLine();
                }
                else
                {
                    realtimeToggleButton.Content = "即時模式：關閉";
                    realtimeToggleButton.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFromString("#546E7A");
                    realtimeStatusText.Text = "即時模式已關閉";
                    realtimeTimer.Stop();
                }
            };

            Grid.SetRow(headerPanel, 0);
            Grid.SetRow(contentBorder, 1);
            Grid.SetRow(chartBorder, 2);
            layoutRoot.Children.Add(headerPanel);
            layoutRoot.Children.Add(contentBorder);
            layoutRoot.Children.Add(chartBorder);
            detailWindow.Content = layoutRoot;

            detailWindow.Loaded += async (s, e) =>
            {
                await System.Threading.Tasks.Task.Run(() => RefreshSectorStockRealtimePrices(sorted, (current, total) =>
                {
                    if (!detailWindow.IsLoaded)
                    {
                        return;
                    }

                    detailWindow.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        var safeTotal = total <= 0 ? 1 : total;
                        var percent = current * 100.0 / safeTotal;
                        progressBar.Value = percent;
                        progressTextBlock.Text = $"yfinance 更新進度：{current}/{total}";
                    }));
                }));

                if (!detailWindow.IsLoaded)
                {
                    return;
                }

                detailWindow.Dispatcher.Invoke(() =>
                {
                    sorted.Sort((a, b) =>
                    {
                        var changeCompare = Nullable.Compare(b.ChangePercent, a.ChangePercent);
                        if (changeCompare != 0)
                        {
                            return changeCompare;
                        }

                        return string.Compare(a.Ticker, b.Ticker, StringComparison.OrdinalIgnoreCase);
                    });

                    grid.Items.Refresh();
                    headerTextBlock.Text = $"📈 {sectorName} 成分股（{sorted.Count} 檔）｜yfinance 已更新";
                    progressBar.Value = 100;
                    progressTextBlock.Text = "yfinance 更新完成";
                    if (grid.SelectedItem == null && sorted.Count > 0)
                    {
                        grid.SelectedItem = sorted[0];
                    }
                    updatePeriodButtonStyles();
                    refreshKLine();
                });
            };

            detailWindow.Closed += (s, e) => realtimeTimer.Stop();

            detailWindow.ShowDialog();
        }

        private bool TryLoadHistoricalDataFromYFinanceForKLine(string ticker, string period, string interval, List<KLinePoint> target)
        {
            try
            {
                var exePath = ResolveYFinanceExecutablePathForKLine();
                var scriptPath = ResolveYFinanceScriptPathForKLine();
                var hasExe = File.Exists(exePath);
                var hasScript = File.Exists(scriptPath);
                if (!hasExe && !hasScript)
                {
                    return false;
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = hasExe ? exePath : AppConfig.PythonPath,
                    Arguments = hasExe ? $"{ticker} history {period} {interval}" : $"\"{scriptPath}\" {ticker} history {period} {interval}",
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
                    process.WaitForExit(15000);

                    var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    if (lines.Length == 0 || lines[0] != "HISTORY_OK")
                    {
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
                        if (!DateTime.TryParse(parts[0], out date)
                            || !double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out open)
                            || !double.TryParse(parts[2], NumberStyles.Any, CultureInfo.InvariantCulture, out high)
                            || !double.TryParse(parts[3], NumberStyles.Any, CultureInfo.InvariantCulture, out low)
                            || !double.TryParse(parts[4], NumberStyles.Any, CultureInfo.InvariantCulture, out close)
                            || !long.TryParse(parts[5], NumberStyles.Any, CultureInfo.InvariantCulture, out volume))
                        {
                            continue;
                        }

                        target.Add(new KLinePoint
                        {
                            Date = date,
                            Open = open,
                            High = high,
                            Low = low,
                            Close = close,
                            Volume = volume
                        });
                    }

                    return target.Count > 0;
                }
            }
            catch
            {
                return false;
            }
        }

        private string ResolveYFinanceExecutablePathForKLine()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var outputPath = IOPath.Combine(baseDir, "Python", "yfinance_fetcher.exe");
            if (File.Exists(outputPath))
            {
                return outputPath;
            }

            var projectDistPath = IOPath.GetFullPath(IOPath.Combine(baseDir, "..", "..", "Python", "dist", "yfinance_fetcher.exe"));
            if (File.Exists(projectDistPath))
            {
                return projectDistPath;
            }

            return outputPath;
        }

        private string ResolveYFinanceScriptPathForKLine()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var outputPath = IOPath.Combine(baseDir, AppConfig.YFinanceScriptPath);
            if (File.Exists(outputPath))
            {
                return outputPath;
            }

            var projectPath = IOPath.GetFullPath(IOPath.Combine(baseDir, "..", "..", AppConfig.YFinanceScriptPath));
            if (File.Exists(projectPath))
            {
                return projectPath;
            }

            return outputPath;
        }

        private void DrawKLineChart(Canvas chartCanvas, List<KLinePoint> candles, string title)
        {
            chartCanvas.Children.Clear();
            if (candles == null || candles.Count == 0)
            {
                return;
            }

            var width = chartCanvas.ActualWidth;
            if (width < 320)
            {
                width = 820;
            }

            var height = chartCanvas.Height;
            var chartLeft = 56.0;
            var chartRight = width - 16.0;
            var priceTop = 26.0;
            var priceBottom = height * 0.65;
            var volumeTop = priceBottom + 20.0;
            var volumeBottom = height - 20.0;

            var recent = candles.Skip(Math.Max(0, candles.Count - 45)).ToList();
            var maxPrice = recent.Max(x => x.High);
            var minPrice = recent.Min(x => x.Low);
            var range = Math.Max(0.01, maxPrice - minPrice);
            var spacing = Math.Max(4.0, (chartRight - chartLeft) / Math.Max(1, recent.Count));
            var bodyWidth = Math.Max(2.0, spacing * 0.6);
            var maxVolume = Math.Max(1L, recent.Max(x => x.Volume));

            Func<double, double> mapPriceToY = price =>
                priceBottom - ((price - minPrice) / range) * (priceBottom - priceTop);

            var gridBrush = new SolidColorBrush(Color.FromRgb(230, 230, 230));
            var axisBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120));

            for (int i = 0; i <= 4; i++)
            {
                var ratio = i / 4.0;
                var y = priceTop + ratio * (priceBottom - priceTop);
                var priceLabel = maxPrice - ratio * range;

                chartCanvas.Children.Add(new Line
                {
                    X1 = chartLeft,
                    Y1 = y,
                    X2 = chartRight,
                    Y2 = y,
                    Stroke = gridBrush,
                    StrokeThickness = 1
                });

                var yText = new TextBlock
                {
                    Text = priceLabel.ToString("F2"),
                    FontSize = 10,
                    Foreground = Brushes.DimGray
                };
                Canvas.SetLeft(yText, 4);
                Canvas.SetTop(yText, y - 8);
                chartCanvas.Children.Add(yText);
            }

            chartCanvas.Children.Add(new Line { X1 = chartLeft, Y1 = priceTop, X2 = chartLeft, Y2 = priceBottom, Stroke = axisBrush, StrokeThickness = 1 });
            chartCanvas.Children.Add(new Line { X1 = chartLeft, Y1 = priceBottom, X2 = chartRight, Y2 = priceBottom, Stroke = axisBrush, StrokeThickness = 1 });
            chartCanvas.Children.Add(new Line { X1 = chartLeft, Y1 = volumeTop, X2 = chartLeft, Y2 = volumeBottom, Stroke = axisBrush, StrokeThickness = 1 });
            chartCanvas.Children.Add(new Line { X1 = chartLeft, Y1 = volumeBottom, X2 = chartRight, Y2 = volumeBottom, Stroke = axisBrush, StrokeThickness = 1 });

            var xTickCount = Math.Min(6, recent.Count);
            var hasIntraday = recent.Select(x => x.Date.Date).Distinct().Count() <= 1;
            if (xTickCount > 1)
            {
                for (int i = 0; i < xTickCount; i++)
                {
                    var idx = (int)Math.Round(i * (recent.Count - 1) / (double)(xTickCount - 1));
                    var x = chartLeft + idx * spacing + bodyWidth / 2;

                    chartCanvas.Children.Add(new Line
                    {
                        X1 = x,
                        Y1 = priceBottom,
                        X2 = x,
                        Y2 = priceBottom + 4,
                        Stroke = axisBrush,
                        StrokeThickness = 1
                    });

                    var dateText = new TextBlock
                    {
                        Text = hasIntraday ? recent[idx].Date.ToString("HH:mm") : recent[idx].Date.ToString("MM/dd"),
                        FontSize = 9,
                        Foreground = Brushes.DimGray
                    };
                    Canvas.SetLeft(dateText, x - 16);
                    Canvas.SetTop(dateText, priceBottom + 4);
                    chartCanvas.Children.Add(dateText);
                }
            }

            chartCanvas.Children.Add(new TextBlock
            {
                Text = "成交量",
                FontSize = 10,
                Foreground = Brushes.DimGray
            });
            Canvas.SetLeft(chartCanvas.Children[chartCanvas.Children.Count - 1], 4);
            Canvas.SetTop(chartCanvas.Children[chartCanvas.Children.Count - 1], volumeTop - 2);

            var titleBlock = new TextBlock
            {
                Text = $"K線（近{recent.Count}日） {title}",
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.DimGray
            };
            Canvas.SetLeft(titleBlock, chartLeft);
            Canvas.SetTop(titleBlock, 0);
            chartCanvas.Children.Add(titleBlock);

            for (int i = 0; i < recent.Count; i++)
            {
                var c = recent[i];
                var x = chartLeft + i * spacing;
                var openY = mapPriceToY(c.Open);
                var closeY = mapPriceToY(c.Close);
                var highY = mapPriceToY(c.High);
                var lowY = mapPriceToY(c.Low);
                var tooltipText =
                    $"時間: {c.Date:yyyy-MM-dd HH:mm}\n" +
                    $"開: {c.Open:F2}\n" +
                    $"高: {c.High:F2}\n" +
                    $"低: {c.Low:F2}\n" +
                    $"收: {c.Close:F2}\n" +
                    $"量: {c.Volume:N0}";

                var isUp = c.Close >= c.Open;
                var stroke = isUp ? new SolidColorBrush(Color.FromRgb(56, 142, 60)) : new SolidColorBrush(Color.FromRgb(211, 47, 47));
                var fill = isUp ? new SolidColorBrush(Color.FromRgb(102, 187, 106)) : new SolidColorBrush(Color.FromRgb(239, 83, 80));

                var wick = new Line
                {
                    X1 = x + bodyWidth / 2,
                    Y1 = highY,
                    X2 = x + bodyWidth / 2,
                    Y2 = lowY,
                    Stroke = stroke,
                    StrokeThickness = 1
                };
                ToolTipService.SetToolTip(wick, tooltipText);
                chartCanvas.Children.Add(wick);

                var body = new Rectangle
                {
                    Width = bodyWidth,
                    Height = Math.Max(1, Math.Abs(closeY - openY)),
                    Fill = fill,
                    Stroke = stroke,
                    StrokeThickness = 1
                };
                ToolTipService.SetToolTip(body, tooltipText);
                Canvas.SetLeft(body, x);
                Canvas.SetTop(body, Math.Min(openY, closeY));
                chartCanvas.Children.Add(body);

                var volumeHeight = (c.Volume / (double)maxVolume) * (volumeBottom - volumeTop);
                var volumeRect = new Rectangle
                {
                    Width = bodyWidth,
                    Height = Math.Max(1, volumeHeight),
                    Fill = fill,
                    Opacity = 0.45
                };
                ToolTipService.SetToolTip(volumeRect, tooltipText);
                Canvas.SetLeft(volumeRect, x);
                Canvas.SetTop(volumeRect, volumeBottom - volumeRect.Height);
                chartCanvas.Children.Add(volumeRect);

                var hitArea = new Rectangle
                {
                    Width = Math.Max(spacing, bodyWidth + 2),
                    Height = volumeBottom - priceTop,
                    Fill = Brushes.Transparent
                };
                ToolTipService.SetToolTip(hitArea, tooltipText);
                Canvas.SetLeft(hitArea, x - (Math.Max(spacing, bodyWidth + 2) - bodyWidth) / 2);
                Canvas.SetTop(hitArea, priceTop);
                chartCanvas.Children.Add(hitArea);
            }

            var ma5 = BuildMASeries(recent, 5);
            var ma20 = BuildMASeries(recent, 20);

            DrawMALine(chartCanvas, ma5, chartLeft, spacing, bodyWidth, mapPriceToY, Color.FromRgb(255, 193, 7));
            DrawMALine(chartCanvas, ma20, chartLeft, spacing, bodyWidth, mapPriceToY, Color.FromRgb(103, 58, 183));

            var legend = new TextBlock
            {
                Text = "MA5(黃)  MA20(紫)",
                FontSize = 10,
                Foreground = Brushes.DimGray
            };
            Canvas.SetLeft(legend, chartRight - 100);
            Canvas.SetTop(legend, 4);
            chartCanvas.Children.Add(legend);
        }

        private List<double?> BuildMASeries(List<KLinePoint> data, int period)
        {
            var result = new List<double?>();
            for (int i = 0; i < data.Count; i++)
            {
                if (i < period - 1)
                {
                    result.Add(null);
                    continue;
                }

                var avg = data.Skip(i - period + 1).Take(period).Average(x => x.Close);
                result.Add(avg);
            }
            return result;
        }

        private void DrawMALine(
            Canvas canvas,
            List<double?> ma,
            double chartLeft,
            double spacing,
            double bodyWidth,
            Func<double, double> mapPriceToY,
            Color color)
        {
            var polyline = new Polyline
            {
                Stroke = new SolidColorBrush(color),
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

        private async System.Threading.Tasks.Task EnsureSectorYFinanceDataAsync()
        {
            if (TryApplyTodaySectorYFinanceCache())
            {
                UpdateNoPriceSummary();
                UpdateDataStatusText();
                return;
            }

            await Dispatcher.BeginInvoke(new Action(() =>
            {
                MessageBox.Show("尚未找到今日台股 yfinance 快取，請回主頁面點擊「更新台股快取」。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }));
        }

        private bool TryApplyTodaySectorYFinanceCache()
        {
            try
            {
                if (!File.Exists(_twYFinanceSectorCacheFile))
                {
                    return false;
                }

                var json = File.ReadAllText(_twYFinanceSectorCacheFile);
                var serializer = new JavaScriptSerializer();
                var cache = serializer.Deserialize<SectorYFinanceCacheEnvelope>(json);
                if (cache == null || cache.Items == null || cache.Items.Count == 0)
                {
                    return false;
                }

                if (!string.Equals(cache.Date, DateTime.Today.ToString("yyyy-MM-dd"), StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                var latestUpdatedAt = cache.Items
                    .Where(x => x.UpdatedAt.HasValue)
                    .Select(x => x.UpdatedAt.Value)
                    .OrderByDescending(x => x)
                    .FirstOrDefault();
                _lastYFinanceCacheUpdatedAt = latestUpdatedAt == default(DateTime) ? (DateTime?)null : latestUpdatedAt;

                var map = cache.Items.ToDictionary(x => x.Ticker, StringComparer.OrdinalIgnoreCase);
                foreach (var stock in _sourceStocks)
                {
                    SectorYFinanceCacheItem item;
                    if (!map.TryGetValue(stock.Ticker, out item))
                    {
                        continue;
                    }

                    stock.Price = item.Price;
                    stock.PreviousClose = item.PreviousClose;
                    stock.ChangePercent = item.ChangePercent;
                    stock.Source = "yfinance-cache";
                    stock.UpdatedAt = item.UpdatedAt;
                }

                foreach (var sample in _sectorSamples)
                {
                    StockInfo stock;
                    if ((stock = _sourceStocks.FirstOrDefault(x => string.Equals(x.Ticker, sample.Ticker, StringComparison.OrdinalIgnoreCase))) != null)
                    {
                        sample.ChangePercent = stock.ChangePercent;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private void SaveSectorYFinanceCache(List<SectorYFinanceCacheItem> items)
        {
            try
            {
                if (!Directory.Exists(AppConfig.UserConfigDir))
                {
                    Directory.CreateDirectory(AppConfig.UserConfigDir);
                }

                var serializer = new JavaScriptSerializer();
                var payload = new SectorYFinanceCacheEnvelope
                {
                    Date = DateTime.Today.ToString("yyyy-MM-dd"),
                    Items = items ?? new List<SectorYFinanceCacheItem>()
                };

                File.WriteAllText(_twYFinanceSectorCacheFile, serializer.Serialize(payload), System.Text.Encoding.UTF8);
            }
            catch
            {
            }
        }

        private string NormalizeTwTicker(string ticker)
        {
            var normalized = (ticker ?? string.Empty).Trim().ToUpperInvariant();
            if (!normalized.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
            {
                normalized += ".TW";
            }
            return normalized;
        }

        private void RefreshSectorStockRealtimePrices(List<SectorStockDetailItem> stocks, Action<int, int> onProgress)
        {
            try
            {
                if (stocks == null || stocks.Count == 0)
                {
                    onProgress?.Invoke(0, 0);
                    return;
                }

                var priceFetcher = new PriceFetcherService();
                var tickerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var stock in stocks)
                {
                    if (string.IsNullOrWhiteSpace(stock.Ticker))
                    {
                        continue;
                    }

                    var normalizedTicker = stock.Ticker.Trim().ToUpperInvariant();
                    if (!normalizedTicker.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
                    {
                        normalizedTicker += ".TW";
                    }

                    tickerMap[stock.Ticker] = normalizedTicker;
                }

                var distinctTickers = tickerMap.Values.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                var total = distinctTickers.Count;
                var current = 0;

                foreach (var ticker in distinctTickers)
                {
                    priceFetcher.UpdatePriceWithPreviousClose(ticker);
                    current++;
                    onProgress?.Invoke(current, total);
                    System.Threading.Thread.Sleep(80);
                }

                var latestPrices = priceFetcher.GetPrices();
                var latestMeta = priceFetcher.GetPriceMeta();

                foreach (var stock in stocks)
                {
                    string normalizedTicker;
                    if (!tickerMap.TryGetValue(stock.Ticker, out normalizedTicker))
                    {
                        continue;
                    }

                    Tuple<double?, double?> realtime;
                    if (!latestPrices.TryGetValue(normalizedTicker, out realtime))
                    {
                        continue;
                    }

                    var currentPrice = realtime.Item1;
                    if (!currentPrice.HasValue)
                    {
                        continue;
                    }

                    stock.Price = currentPrice;

                    double? previousClose = null;
                    Dictionary<string, object> meta;
                    if (latestMeta.TryGetValue(normalizedTicker, out meta) && meta != null && meta.ContainsKey("previous_close"))
                    {
                        previousClose = meta["previous_close"] as double?;
                    }

                    if (previousClose.HasValue && Math.Abs(previousClose.Value) > 0.000001)
                    {
                        stock.ChangeAmount = currentPrice.Value - previousClose.Value;
                        stock.ChangePercent = (stock.ChangeAmount.Value / previousClose.Value) * 100;
                    }
                    else
                    {
                        stock.ChangeAmount = null;
                        stock.ChangePercent = realtime.Item2;
                    }
                }
            }
            catch
            {
            }
        }

        private void LoadAllListedTwStocks()
        {
            try
            {
                Dispatcher.Invoke(() => txtDataStatus.Text = "資料來源：TWSE 全部上市股票 + Yahoo類股架構載入中...");
                var loadedFromCsv = LoadSectorMappingFromCsv();
                _loadedFromCsv = loadedFromCsv;
                if (!loadedFromCsv)
                {
                    var yahooSectorLinks = LoadYahooListedSectorFramework();
                    LoadYahooSectorConstituents(yahooSectorLinks);
                }

                _sourceStocks.Clear();
                _sectorSamples.Clear();

                var stockMap = new Dictionary<string, StockInfo>(StringComparer.OrdinalIgnoreCase);
                var sectorSampleMap = new Dictionary<string, SectorSampleItem>(StringComparer.OrdinalIgnoreCase);

                if (loadedFromCsv)
                {
                    foreach (var csvTicker in _tickerToCsvSector.Keys)
                    {
                        var ticker = NormalizeCsvTicker(csvTicker);
                        if (string.IsNullOrWhiteSpace(ticker) || stockMap.ContainsKey(ticker))
                        {
                            continue;
                        }

                        var stock = new StockInfo(ticker, ticker)
                        {
                            Source = "CSV族群"
                        };
                        _sourceStocks.Add(stock);
                        stockMap[ticker] = stock;

                        var sample = new SectorSampleItem
                        {
                            Ticker = ticker,
                            Industry = GetCsvSectorByTicker(ticker),
                            Name = ticker
                        };
                        _sectorSamples.Add(sample);
                        sectorSampleMap[ticker] = sample;
                    }
                }

                var url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL";
                string json;
                using (var client = new WebClient())
                {
                    client.Encoding = System.Text.Encoding.UTF8;
                    json = client.DownloadString(url);
                }

                var serializer = new JavaScriptSerializer();
                var rows = serializer.Deserialize<List<Dictionary<string, string>>>(json) ?? new List<Dictionary<string, string>>();

                foreach (var row in rows)
                {
                    string ticker;
                    string name;
                    string closingPriceText;
                    string changeText;

                    if (!row.TryGetValue("Code", out ticker) || !row.TryGetValue("Name", out name))
                    {
                        continue;
                    }

                    string industry;
                    row.TryGetValue("Industry", out industry);

                    row.TryGetValue("ClosingPrice", out closingPriceText);
                    row.TryGetValue("Change", out changeText);

                    var isEtf = IsEtfTicker(ticker);
                    if (loadedFromCsv && !stockMap.ContainsKey(ticker) && !isEtf)
                    {
                        continue;
                    }

                    var close = ParseNumeric(closingPriceText);
                    var changeAmount = ParseNumeric(changeText);
                    double? previousClose = null;
                    double? changePercent = null;

                    if (close.HasValue && changeAmount.HasValue)
                    {
                        previousClose = close.Value - changeAmount.Value;
                        if (previousClose.HasValue && Math.Abs(previousClose.Value) > 0.000001)
                        {
                            changePercent = (changeAmount.Value / previousClose.Value) * 100;
                        }
                    }

                    StockInfo stock;
                    if (!stockMap.TryGetValue(ticker, out stock))
                    {
                        stock = new StockInfo(ticker, name);
                        _sourceStocks.Add(stock);
                        stockMap[ticker] = stock;
                    }

                    stock.Name = name;
                    stock.Price = close;
                    stock.PreviousClose = previousClose;
                    stock.ChangePercent = changePercent;
                    stock.Source = "TWSE上市";
                    stock.UpdatedAt = DateTime.Now;

                    SectorSampleItem sample;
                    if (!sectorSampleMap.TryGetValue(ticker, out sample))
                    {
                        sample = new SectorSampleItem { Ticker = ticker };
                        _sectorSamples.Add(sample);
                        sectorSampleMap[ticker] = sample;
                    }

                    var csvSector = loadedFromCsv ? GetCsvSectorByTicker(ticker) : null;
                    sample.Industry = !string.IsNullOrWhiteSpace(csvSector)
                        ? csvSector
                        : (isEtf ? "ETF" : industry);
                    sample.Name = name;
                    sample.ChangePercent = changePercent;
                }

                var cacheApplied = TryApplyTodaySectorYFinanceCache();
                if (cacheApplied && _lastYFinanceCacheUpdatedAt.HasValue)
                {
                    _lastAutoYFinanceUpdateDate = _lastYFinanceCacheUpdatedAt.Value.Date;
                }
                UpdateNoPriceSummary();
                UpdateDataStatusText();
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => txtDataStatus.Text = "資料來源：載入失敗");
                Dispatcher.Invoke(() => MessageBox.Show($"載入全部上市台股失敗：{ex.Message}", "載入失敗", MessageBoxButton.OK, MessageBoxImage.Warning));
            }

        }

        private string NormalizeCsvTicker(string ticker)
        {
            var normalized = (ticker ?? string.Empty).Trim().ToUpperInvariant();
            if (normalized.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(0, normalized.Length - 3);
            }

            return normalized;
        }

        private string GetCsvSectorByTicker(string ticker)
        {
            if (string.IsNullOrWhiteSpace(ticker))
            {
                return null;
            }

            string sector;
            if (_tickerToCsvSector.TryGetValue(ticker, out sector))
            {
                return sector;
            }

            var twTicker = ticker.EndsWith(".TW", StringComparison.OrdinalIgnoreCase)
                ? ticker
                : ticker + ".TW";

            if (_tickerToCsvSector.TryGetValue(twTicker, out sector))
            {
                return sector;
            }

            return null;
        }

        private bool IsEtfTicker(string ticker)
        {
            if (string.IsNullOrWhiteSpace(ticker))
            {
                return false;
            }

            var normalized = ticker.Trim().ToUpperInvariant();
            if (normalized.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
            {
                normalized = normalized.Substring(0, normalized.Length - 3);
            }

            return Regex.IsMatch(normalized, "^00\\d+[A-Z]?$");
        }

        private void RefreshAllStocksFromYFinance(Action<int, int> onProgress)
        {
            var fetcher = new PriceFetcherService();
            var tickers = _sourceStocks
                .Select(x => NormalizeTwTicker(x.Ticker))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var total = tickers.Count;

            var completed = 0;
            System.Threading.Tasks.Parallel.ForEach(
                tickers,
                new System.Threading.Tasks.ParallelOptions { MaxDegreeOfParallelism = 4 },
                ticker =>
                {
                    fetcher.UpdatePriceWithPreviousClose(ticker);
                    var current = System.Threading.Interlocked.Increment(ref completed);
                    onProgress?.Invoke(current, total);
                });

            var prices = fetcher.GetPrices();
            var meta = fetcher.GetPriceMeta();
            var cacheItems = new List<SectorYFinanceCacheItem>();

            foreach (var stock in _sourceStocks)
            {
                var ticker = NormalizeTwTicker(stock.Ticker);
                Tuple<double?, double?> priceTuple;
                if (!prices.TryGetValue(ticker, out priceTuple))
                {
                    continue;
                }

                var price = priceTuple.Item1;
                double? previousClose = null;
                Dictionary<string, object> itemMeta;
                if (meta.TryGetValue(ticker, out itemMeta) && itemMeta != null && itemMeta.ContainsKey("previous_close"))
                {
                    previousClose = itemMeta["previous_close"] as double?;
                }

                double? changePercent;
                if (price.HasValue && previousClose.HasValue && Math.Abs(previousClose.Value) > 0.000001)
                {
                    changePercent = (price.Value - previousClose.Value) / previousClose.Value * 100;
                }
                else
                {
                    changePercent = priceTuple.Item2;
                }

                stock.Price = price;
                stock.PreviousClose = previousClose;
                stock.ChangePercent = changePercent;
                stock.Source = "yfinance";
                stock.UpdatedAt = DateTime.Now;

                cacheItems.Add(new SectorYFinanceCacheItem
                {
                    Ticker = stock.Ticker,
                    Price = stock.Price,
                    PreviousClose = stock.PreviousClose,
                    ChangePercent = stock.ChangePercent,
                    UpdatedAt = stock.UpdatedAt
                });
            }

            foreach (var sample in _sectorSamples)
            {
                var stock = _sourceStocks.FirstOrDefault(x => string.Equals(x.Ticker, sample.Ticker, StringComparison.OrdinalIgnoreCase));
                sample.ChangePercent = stock?.ChangePercent;
            }

            SaveSectorYFinanceCache(cacheItems);
            _lastYFinanceCacheUpdatedAt = DateTime.Now;
        }

        private void StartAutoYFinanceBackgroundUpdate()
        {
            if (_autoYFinanceUpdateTimer != null)
            {
                return;
            }

            _autoYFinanceUpdateTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(5)
            };
            _autoYFinanceUpdateTimer.Tick += async (s, e) => await TryRunAutoYFinanceUpdateAsync();
            _autoYFinanceUpdateTimer.Start();
        }

        private async System.Threading.Tasks.Task TryRunAutoYFinanceUpdateAsync()
        {
            if (_isAutoYFinanceUpdating || _sourceStocks.Count == 0)
            {
                return;
            }

            var now = DateTime.Now;
            var updateStartTime = new TimeSpan(14, 0, 0);
            if (now.TimeOfDay < updateStartTime)
            {
                return;
            }

            if (_lastAutoYFinanceUpdateDate.Date == now.Date)
            {
                return;
            }

            if (_lastYFinanceCacheUpdatedAt.HasValue
                && _lastYFinanceCacheUpdatedAt.Value.Date == now.Date
                && _lastYFinanceCacheUpdatedAt.Value.TimeOfDay >= updateStartTime)
            {
                _lastAutoYFinanceUpdateDate = now.Date;
                UpdateDataStatusText();
                return;
            }

            _isAutoYFinanceUpdating = true;
            try
            {
                await System.Threading.Tasks.Task.Run(() => RefreshAllStocksFromYFinance(null));
                _lastAutoYFinanceUpdateDate = now.Date;

                UpdateNoPriceSummary();
                ApplyFilter();
                UpdateDataStatusText();
            }
            catch
            {
            }
            finally
            {
                _isAutoYFinanceUpdating = false;
            }
        }

        private void UpdateNoPriceSummary()
        {
            var noPriceStocks = _sourceStocks
                .Where(x => !x.Price.HasValue)
                .Select(x => x.Ticker)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            _noPriceSummary = noPriceStocks.Count > 0
                ? $"｜無股價資料：{noPriceStocks.Count} 檔（例：{string.Join(",", noPriceStocks.Take(8))}）"
                : string.Empty;
        }

        private void UpdateDataStatusText()
        {
            var yfinanceText = _lastYFinanceCacheUpdatedAt.HasValue
                ? $"今日{_lastYFinanceCacheUpdatedAt.Value:HH:mm}已更新"
                : "無";

            var text = _loadedFromCsv
                ? $"資料來源：CSV族群股票（{_sourceStocks.Count} 檔）｜TWSE補齊名稱｜yfinance暫存：{yfinanceText}{_noPriceSummary}"
                : $"資料來源：TWSE 全部上市股票（{_sourceStocks.Count} 筆）｜Yahoo上市類股（{_yahooListedSectors.Count} 類，{_tickerToYahooSector.Count} 檔已對應）｜yfinance暫存：{yfinanceText}{_noPriceSummary}";

            if (Dispatcher.CheckAccess())
            {
                txtDataStatus.Text = text;
            }
            else
            {
                Dispatcher.Invoke(() => txtDataStatus.Text = text);
            }
        }

        private bool LoadSectorMappingFromCsv()
        {
            List<string> sectorOrder;
            Dictionary<string, string> tickerToSector;
            if (!AppConfig.TryLoadTaiwanSectorCsv(out sectorOrder, out tickerToSector))
            {
                return false;
            }

            _csvSectorOrder.Clear();
            _tickerToCsvSector.Clear();

            foreach (var category in sectorOrder)
            {
                _csvSectorOrder.Add(category);
            }

            foreach (var kv in tickerToSector)
            {
                _tickerToCsvSector[kv.Key] = kv.Value;
            }

            return _tickerToCsvSector.Count > 0;
        }

        private List<YahooSectorLink> LoadYahooListedSectorFramework()
        {
            var result = new List<YahooSectorLink>();
            try
            {
                var url = "https://tw.stock.yahoo.com/class/";
                string html;
                using (var client = new WebClient())
                {
                    client.Encoding = System.Text.Encoding.UTF8;
                    html = client.DownloadString(url);
                }

                var matches = Regex.Matches(html,
                    "href=[\"'](?<href>/class-quote\\?sectorId=\\d+&exchange=TAI[^\"']*)[\"'][^>]*>(?<name>[^<]+)</a>",
                    RegexOptions.IgnoreCase);

                var sectors = new List<string>();
                foreach (Match m in matches)
                {
                    var name = WebUtility.HtmlDecode(m.Groups["name"].Value).Trim();
                    var href = WebUtility.HtmlDecode(m.Groups["href"].Value).Trim();
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    // 排除非產業分析類別
                    if (name.Contains("認購") || name.Contains("認售") || name.Contains("牛證") || name.Contains("熊證") || name.Contains("指數類"))
                    {
                        continue;
                    }

                    if (!sectors.Contains(name))
                    {
                        sectors.Add(name);
                        result.Add(new YahooSectorLink
                        {
                            Name = name,
                            Url = "https://tw.stock.yahoo.com" + href
                        });
                    }
                }

                _yahooListedSectors.Clear();
                foreach (var s in sectors)
                {
                    _yahooListedSectors.Add(s);
                }

                return result;
            }
            catch
            {
                _yahooListedSectors.Clear();
                _yahooListedSectors.AddRange(YahooSectorFallback);
                return result;
            }
        }

        private void LoadYahooSectorConstituents(List<YahooSectorLink> sectorLinks)
        {
            _tickerToYahooSector.Clear();

            foreach (var sector in sectorLinks)
            {
                try
                {
                    string html;
                    using (var client = new WebClient())
                    {
                        client.Encoding = System.Text.Encoding.UTF8;
                        html = client.DownloadString(sector.Url);
                    }

                    var quoteMatches = Regex.Matches(html, "/quote/(?<ticker>\\d{4})(?:\\.TW)?", RegexOptions.IgnoreCase);
                    foreach (Match m in quoteMatches)
                    {
                        var ticker = m.Groups["ticker"].Value;
                        if (string.IsNullOrWhiteSpace(ticker))
                        {
                            continue;
                        }

                        if (!_tickerToYahooSector.ContainsKey(ticker))
                        {
                            _tickerToYahooSector[ticker] = sector.Name;
                        }
                    }

                    var symbolMatches = Regex.Matches(html, "\"symbol\"\\s*:\\s*\"(?<ticker>\\d{4})\\.TW\"", RegexOptions.IgnoreCase);
                    foreach (Match m in symbolMatches)
                    {
                        var ticker = m.Groups["ticker"].Value;
                        if (string.IsNullOrWhiteSpace(ticker))
                        {
                            continue;
                        }

                        if (!_tickerToYahooSector.ContainsKey(ticker))
                        {
                            _tickerToYahooSector[ticker] = sector.Name;
                        }
                    }
                }
                catch
                {
                }
            }
        }

        private double? ParseNumeric(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var cleaned = text.Replace(",", "").Trim();
            cleaned = Regex.Replace(cleaned, "[^0-9+\\-.]", "");

            double value;
            if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            {
                return value;
            }

            if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
            {
                return value;
            }

            return null;
        }

        private int CalculateTradingScore(StockInfo stock)
        {
            var score = 50;
            var change = stock.ChangePercent;

            if (change.HasValue)
            {
                if (change.Value >= 4) score += 20;
                else if (change.Value >= 2) score += 12;
                else if (change.Value > 0) score += 5;
                else if (change.Value <= -4) score -= 20;
                else if (change.Value <= -2) score -= 12;
                else if (change.Value < 0) score -= 5;
            }
            else
            {
                score -= 5;
            }

            if (stock.Price.HasValue && stock.PreviousClose.HasValue)
            {
                var gap = stock.Price.Value - stock.PreviousClose.Value;
                if (gap > 0) score += 3;
                else if (gap < 0) score -= 3;
            }

            return Math.Max(0, Math.Min(100, score));
        }

        private string GetTradingSuggestion(int score)
        {
            if (score >= 70) return "偏多（買入）";
            if (score >= 50) return "中性（觀望）";
            return "偏空（賣出）";
        }

        private class TwFilterStockItem
        {
            public string Ticker { get; set; }
            public string Name { get; set; }
            public double? Price { get; set; }
            public double? ChangePercent { get; set; }
            public double? PreviousClose { get; set; }
            public DateTime? UpdatedAt { get; set; }
            public int TradingScore { get; set; }
            public string TradingSuggestion { get; set; }
        }

        private class SectorSampleItem
        {
            public string Ticker { get; set; }
            public string Industry { get; set; }
            public string Name { get; set; }
            public double? ChangePercent { get; set; }
        }

        private class SectorTrendItem
        {
            public string Industry { get; set; }
            public double AverageChangePercent { get; set; }
            public int RisingCount { get; set; }
            public int FallingCount { get; set; }
            public string Trend { get; set; }
        }

        private class SectorStockDetailItem
        {
            public string Ticker { get; set; }
            public string Name { get; set; }
            public double? Price { get; set; }
            public double? ChangeAmount { get; set; }
            public double? ChangePercent { get; set; }
        }

        private class KLinePoint
        {
            public DateTime Date { get; set; }
            public double Open { get; set; }
            public double High { get; set; }
            public double Low { get; set; }
            public double Close { get; set; }
            public long Volume { get; set; }
        }

        private class YahooSectorLink
        {
            public string Name { get; set; }
            public string Url { get; set; }
        }

        private class SectorYFinanceCacheEnvelope
        {
            public string Date { get; set; }
            public List<SectorYFinanceCacheItem> Items { get; set; }
        }

        private class SectorYFinanceCacheItem
        {
            public string Ticker { get; set; }
            public double? Price { get; set; }
            public double? PreviousClose { get; set; }
            public double? ChangePercent { get; set; }
            public DateTime? UpdatedAt { get; set; }
        }

        private string MapSectorCategory(string industry, string name, string ticker)
        {
            var i = (industry ?? string.Empty).ToLowerInvariant();
            var n = (name ?? string.Empty).ToLowerInvariant();

            string csvSector;
            if (!string.IsNullOrWhiteSpace(ticker) && _tickerToCsvSector.TryGetValue(ticker, out csvSector))
            {
                return csvSector;
            }

            string yahooSector;
            if (!string.IsNullOrWhiteSpace(ticker) && _tickerToYahooSector.TryGetValue(ticker, out yahooSector))
            {
                return yahooSector;
            }

            if (i.Contains("水泥")) return "水泥";
            if (i.Contains("食品")) return "食品";
            if (i.Contains("塑膠")) return "塑膠";
            if (i.Contains("紡織")) return "紡織";
            if (i.Contains("電機") || i.Contains("機械")) return "電機機械";
            if (i.Contains("電器") || i.Contains("電纜")) return "電器電纜";
            if (i.Contains("化學") || i.Contains("化工")) return "化學";
            if (i.Contains("生技") || n.Contains("醫療")) return "生技";
            if (i.Contains("玻璃") || i.Contains("陶瓷")) return "玻璃";
            if (i.Contains("造紙") || i.Contains("紙")) return "造紙";
            if (i.Contains("鋼鐵")) return "鋼鐵";
            if (i.Contains("橡膠")) return "橡膠";
            if (i.Contains("汽車") || n.Contains("車用")) return "汽車";
            if (i.Contains("半導體")) return "半導體";
            if (i.Contains("電腦") || i.Contains("週邊") || n.Contains("伺服器")) return "電腦週邊";
            if (i.Contains("光電") || n.Contains("led") || n.Contains("光學")) return "光電";
            if (i.Contains("通訊") || i.Contains("通信") || i.Contains("網路")) return "通訊網路";
            if (i.Contains("電子零組件") || n.Contains("pcb") || n.Contains("連接器")) return "電子零組件";
            if (i.Contains("電子通路")) return "電子通路";
            if (i.Contains("資訊服務")) return "資訊服務";
            if (i.Contains("其他電子")) return "其他電子";
            if (i.Contains("建材") || i.Contains("營造") || i.Contains("建設")) return "營建";
            if (i.Contains("航運") || i.Contains("航空")) return "航運";
            if (i.Contains("觀光") || i.Contains("餐旅")) return "觀光餐旅";
            if (i.Contains("金融") || i.Contains("保險") || i.Contains("證券") || i.Contains("銀行")) return "金融業";
            if (i.Contains("貿易") || i.Contains("百貨")) return "貿易百貨";
            if (i.Contains("油電") || i.Contains("燃氣")) return "油電燃氣";
            if (i.Contains("運動休閒") || n.Contains("運動")) return "運動休閒";
            if (i.Contains("數位雲端") || n.Contains("雲端")) return "數位雲端";
            if (i.Contains("綠能") || n.Contains("風電") || n.Contains("太陽能") || n.Contains("儲能")) return "綠能環保";

            return _yahooListedSectors.Contains("其他") ? "其他" : "其他電子";
        }
    }
}
