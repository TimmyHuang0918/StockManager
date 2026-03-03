using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Web.Script.Serialization;
using StockManager.Config;
using StockManager.Converters;
using StockManager.Models;
using StockManager.Services;

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

        public TwStockFilterWindow()
        {
            InitializeComponent();
            dgFilteredTwStocks.ItemsSource = _filteredStocks;
            Loaded += TwStockFilterWindow_Loaded;
        }

        private async void TwStockFilterWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await System.Threading.Tasks.Task.Run(() => LoadAllListedTwStocks());
            ApplyFilter();
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

        private void BtnSectorOverview_Click(object sender, RoutedEventArgs e)
        {
            if (_sectorSamples.Count == 0)
            {
                MessageBox.Show("尚未載入族群資料。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

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

            Grid.SetRow(headerPanel, 0);
            Grid.SetRow(contentBorder, 1);
            layoutRoot.Children.Add(headerPanel);
            layoutRoot.Children.Add(contentBorder);
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
                });
            };

            detailWindow.ShowDialog();
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
                if (!loadedFromCsv)
                {
                    var yahooSectorLinks = LoadYahooListedSectorFramework();
                    LoadYahooSectorConstituents(yahooSectorLinks);
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

                _sourceStocks.Clear();
                _sectorSamples.Clear();

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

                    var stock = new StockInfo(ticker, name)
                    {
                        Price = close,
                        PreviousClose = previousClose,
                        ChangePercent = changePercent,
                        Source = "TWSE上市",
                        UpdatedAt = DateTime.Now
                    };

                    _sourceStocks.Add(stock);

                    _sectorSamples.Add(new SectorSampleItem
                    {
                        Ticker = ticker,
                        Industry = industry,
                        Name = name,
                        ChangePercent = changePercent
                    });
                }

                Dispatcher.Invoke(() => txtDataStatus.Text = loadedFromCsv
                    ? $"資料來源：TWSE 全部上市股票（{_sourceStocks.Count} 筆）｜CSV族群（{_csvSectorOrder.Count} 類，{_tickerToCsvSector.Count} 檔已對應）"
                    : $"資料來源：TWSE 全部上市股票（{_sourceStocks.Count} 筆）｜Yahoo上市類股（{_yahooListedSectors.Count} 類，{_tickerToYahooSector.Count} 檔已對應）");
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => txtDataStatus.Text = "資料來源：載入失敗");
                Dispatcher.Invoke(() => MessageBox.Show($"載入全部上市台股失敗：{ex.Message}", "載入失敗", MessageBoxButton.OK, MessageBoxImage.Warning));
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

        private class YahooSectorLink
        {
            public string Name { get; set; }
            public string Url { get; set; }
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
