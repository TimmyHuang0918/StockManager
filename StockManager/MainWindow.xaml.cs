using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.IO;
using System.Net;
using System.Web.Script.Serialization;
using System.Windows.Threading;
using System.Xml.Linq;
using System.Diagnostics;
using StockManager.Models;
using StockManager.Services;
using StockManager.Config;

namespace StockManager
{
        /// <summary>
        /// MainWindow.xaml 的互動邏輯
        /// </summary>
        public partial class MainWindow : Window
        {
                private StockManagerService _usStockManager;
                private StockManagerService _twStockManager;
                private PriceFetcherService _usPriceFetcher;
                private PriceFetcherService _twPriceFetcher;
                private MonitorService _usMonitor;
                private MonitorService _twMonitor;

                private ObservableCollection<StockInfo> _usStockList;
                private ObservableCollection<StockInfo> _twStockList;
                private ObservableCollection<StockInfo> _usFilteredStockList;
                private ObservableCollection<StockInfo> _twFilteredStockList;
                private ObservableCollection<HoldingInfo> _usHoldingList;
                private ObservableCollection<HoldingInfo> _twHoldingList;
                private ObservableCollection<NewsImpactItem> _newsImpactList;
                private readonly string _holdingUsersFile = System.IO.Path.Combine(AppConfig.UserConfigDir, "holding_users.json");
                private List<string> _holdingUsers = new List<string>();
                private string _currentHoldingUser = "default";
                private double _usRealizedPnL = 0;
                private double _twRealizedPnL = 0;
                private DateTime _lastNewsUpdate = DateTime.MinValue;
                private bool _isNewsUpdating = false;
                private Dictionary<string, string> _translationCache = new Dictionary<string, string>();

                private DispatcherTimer _updateTimer;
                private DispatcherTimer _countdownTimer;
                private int _countdownSeconds = 0;
                private int _updateInterval = 10;

                // 手動刷新防抖動
                private DateTime _lastManualRefresh = DateTime.MinValue;
                private const int MANUAL_REFRESH_COOLDOWN = 5; // 5 秒冷卻時間

                private DebugWindow _debugWindow;

                public MainWindow()
                {
                        InitializeComponent();
                        InitializeServices();
                        InitializeUI();
                        StartMonitoring();

                        // 重定向 Console 輸出到調試視窗
                        Console.SetOut(new DebugTextWriter(this));
                }

                private void CboHoldingUser_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                        if (cboHoldingUser == null)
                        {
                                return;
                        }

                        var selected = cboHoldingUser.SelectedItem as string;
                        if (string.IsNullOrWhiteSpace(selected))
                        {
                                return;
                        }

                        if (string.Equals(_currentHoldingUser, selected, StringComparison.OrdinalIgnoreCase))
                        {
                                return;
                        }

                        _currentHoldingUser = selected.Trim();
                        LoadHoldings();
                        statusText.Text = $"已切換庫存使用者：{_currentHoldingUser}";
                }

                private void BtnAddHoldingUser_Click(object sender, RoutedEventArgs e)
                {
                        var input = ShowUserNameInputDialog();
                        if (string.IsNullOrWhiteSpace(input)) return;

                        // 允許中文等名稱，只過濾檔名非法字元
                        var invalidChars = System.IO.Path.GetInvalidFileNameChars();
                        var user = new string(input.Select(ch => invalidChars.Contains(ch) ? '_' : ch).ToArray()).Trim();
                        if (string.IsNullOrWhiteSpace(user))
                        {
                                MessageBox.Show("使用者名稱格式無效", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        var existed = _holdingUsers.Contains(user, StringComparer.OrdinalIgnoreCase);
                        if (!existed)
                        {
                                _holdingUsers.Add(user);
                                _holdingUsers = _holdingUsers.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
                                SaveHoldingUsersToFile();
                        }

                        cboHoldingUser.ItemsSource = null;
                        cboHoldingUser.ItemsSource = _holdingUsers;
                        cboHoldingUser.SelectedItem = _holdingUsers.FirstOrDefault(x => string.Equals(x, user, StringComparison.OrdinalIgnoreCase));
                        statusText.Text = existed ? $"已切換庫存使用者：{user}" : $"已新增並切換庫存使用者：{user}";
                }

                private void BtnDeleteHoldingUser_Click(object sender, RoutedEventArgs e)
                {
                        var selected = cboHoldingUser?.SelectedItem as string;
                        if (string.IsNullOrWhiteSpace(selected))
                        {
                                MessageBox.Show("請先選擇要刪除的使用者", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        if (_holdingUsers.Count <= 1)
                        {
                                MessageBox.Show("至少要保留一個使用者", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        var confirm = MessageBox.Show(
                                $"確定要刪除使用者「{selected}」嗎？\n\n此動作會刪除該使用者的庫存與損益資料。",
                                "刪除使用者警告",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Warning);

                        if (confirm != MessageBoxResult.Yes)
                        {
                                return;
                        }

                        var userDir = System.IO.Path.Combine(AppConfig.UserConfigDir, "users", selected);
                        if (Directory.Exists(userDir))
                        {
                                Directory.Delete(userDir, true);
                        }

                        _holdingUsers = _holdingUsers
                                .Where(x => !string.Equals(x, selected, StringComparison.OrdinalIgnoreCase))
                                .ToList();
                        SaveHoldingUsersToFile();

                        var next = _holdingUsers.FirstOrDefault() ?? "default";
                        if (!_holdingUsers.Any())
                        {
                                _holdingUsers.Add(next);
                                SaveHoldingUsersToFile();
                        }

                        cboHoldingUser.ItemsSource = null;
                        cboHoldingUser.ItemsSource = _holdingUsers;
                        cboHoldingUser.SelectedItem = next;
                        statusText.Text = $"已刪除使用者：{selected}";
                }

                private string ShowUserNameInputDialog()
                {
                        var dialog = new Window
                        {
                                Title = "新增使用者",
                                Width = 340,
                                Height = 200,
                                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                                ResizeMode = ResizeMode.NoResize,
                                Owner = this,
                                Background = new SolidColorBrush(Color.FromRgb(245, 247, 250))
                        };

                        var root = new Grid { Margin = new Thickness(14) };
                        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

                        var label = new TextBlock { Text = "請輸入使用者名稱", FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(Color.FromRgb(44, 62, 80)) };
                        Grid.SetRow(label, 0);
                        root.Children.Add(label);

                        var textBox = new TextBox { Height = 30, Margin = new Thickness(0, 8, 0, 0), VerticalContentAlignment = VerticalAlignment.Center };
                        Grid.SetRow(textBox, 1);
                        root.Children.Add(textBox);

                        var hint = new TextBlock { Text = "可使用中文，系統會自動過濾不合法字元", Margin = new Thickness(0, 8, 0, 0), Foreground = new SolidColorBrush(Color.FromRgb(96, 125, 139)), FontSize = 11 };
                        Grid.SetRow(hint, 2);
                        root.Children.Add(hint);

                        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
                        Grid.SetRow(buttons, 3);

                        var btnOk = new Button { Content = "確定", Width = 80, Height = 30, Margin = new Thickness(0, 0, 8, 0), Background = new SolidColorBrush(Color.FromRgb(39, 174, 96)), Foreground = Brushes.White, BorderThickness = new Thickness(0) };
                        var btnCancel = new Button { Content = "取消", Width = 80, Height = 30, Background = new SolidColorBrush(Color.FromRgb(149, 165, 166)), Foreground = Brushes.White, BorderThickness = new Thickness(0) };

                        btnOk.Click += (s, e) => dialog.DialogResult = true;
                        btnCancel.Click += (s, e) => dialog.DialogResult = false;

                        buttons.Children.Add(btnOk);
                        buttons.Children.Add(btnCancel);
                        root.Children.Add(buttons);

                        dialog.Content = root;
                        dialog.Loaded += (s, e) => textBox.Focus();

                        var result = dialog.ShowDialog();
                        if (result != true)
                        {
                                return null;
                        }

                        return (textBox.Text ?? string.Empty).Trim();
                }

                private void InitializeServices()
                {
                        // 初始化美股服務
                        _usStockManager = new StockManagerService(AppConfig.DefaultStocks, AppConfig.UserStocksFile);
                        _usPriceFetcher = new PriceFetcherService();
                        _usMonitor = new MonitorService(_usStockManager, _usPriceFetcher);

                        // 初始化台股服務
                        _twStockManager = new StockManagerService(AppConfig.DefaultTwStocks, AppConfig.UserTwStocksFile);
                        _twPriceFetcher = new PriceFetcherService();
                        _twMonitor = new MonitorService(_twStockManager, _twPriceFetcher);
                }

                private void InitializeUI()
                {
                        // 初始化股票列表
                        _usStockList = new ObservableCollection<StockInfo>();
                        _twStockList = new ObservableCollection<StockInfo>();
                        _usFilteredStockList = new ObservableCollection<StockInfo>();
                        _twFilteredStockList = new ObservableCollection<StockInfo>();
                        _usHoldingList = new ObservableCollection<HoldingInfo>();
                        _twHoldingList = new ObservableCollection<HoldingInfo>();
                        _newsImpactList = new ObservableCollection<NewsImpactItem>();

                        dgUsStocks.ItemsSource = _usFilteredStockList;
                        dgTwStocks.ItemsSource = _twFilteredStockList;
                        dgUsHoldings.ItemsSource = _usHoldingList;
                        dgTwHoldings.ItemsSource = _twHoldingList;
                        dgNewsImpact.ItemsSource = _newsImpactList;

                        InitializeHoldingUsers();

                        // 載入股票數據
                        LoadStockData();
                        LoadHoldings();

                        // 設置定時器
                        _updateTimer = new DispatcherTimer();
                        _updateTimer.Interval = TimeSpan.FromSeconds(1);
                        _updateTimer.Tick += UpdateTimer_Tick;
                        _updateTimer.Start();

                        _countdownTimer = new DispatcherTimer();
                        _countdownTimer.Interval = TimeSpan.FromSeconds(1);
                        _countdownTimer.Tick += CountdownTimer_Tick;
                        _countdownTimer.Start();
                }

                private void LoadStockData()
                {
                        // 載入美股
                        var usStocks = _usStockManager.GetStocks();
                        _usStockList.Clear();
                        foreach (var stock in usStocks)
                        {
                                _usStockList.Add(new StockInfo(stock.Key, stock.Value));
                        }
                        ApplyFilter("US");

                        // 載入台股
                        var twStocks = _twStockManager.GetStocks();
                        _twStockList.Clear();
                        foreach (var stock in twStocks)
                        {
                                _twStockList.Add(new StockInfo(stock.Key, stock.Value));
                        }
                        ApplyFilter("TW");
                }

                private void LoadHoldings()
                {
                        _usRealizedPnL = ReadRealizedPnL(GetRealizedFile("US"));
                        _twRealizedPnL = ReadRealizedPnL(GetRealizedFile("TW"));

                        _usHoldingList.Clear();
                        foreach (var h in ReadHoldingsFromFile(GetHoldingFile("US")))
                        {
                                _usHoldingList.Add(h);
                        }

                        _twHoldingList.Clear();
                        foreach (var h in ReadHoldingsFromFile(GetHoldingFile("TW")))
                        {
                                _twHoldingList.Add(h);
                        }

                        UpdateHoldingSummary("US");
                        UpdateHoldingSummary("TW");
                }

                private double ReadRealizedPnL(string filePath)
                {
                        try
                        {
                                if (!File.Exists(filePath)) return 0;
                                var text = File.ReadAllText(filePath).Trim();
                                if (double.TryParse(text, out double value)) return value;
                        }
                        catch { }

                        return 0;
                }

                private void SaveRealizedPnL(string market)
                {
                        try
                        {
                                if (!Directory.Exists(AppConfig.UserConfigDir))
                                {
                                        Directory.CreateDirectory(AppConfig.UserConfigDir);
                                }

                                var filePath = GetRealizedFile(market);
                                var dir = System.IO.Path.GetDirectoryName(filePath);
                                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                                {
                                        Directory.CreateDirectory(dir);
                                }

                                var value = market == "US" ? _usRealizedPnL : _twRealizedPnL;
                                File.WriteAllText(filePath, value.ToString("F4"), Encoding.UTF8);
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[已實現損益儲存失敗] {market}: {ex.Message}");
                        }
                }

                private List<HoldingInfo> ReadHoldingsFromFile(string filePath)
                {
                        try
                        {
                                if (!File.Exists(filePath))
                                {
                                        return new List<HoldingInfo>();
                                }

                                var json = File.ReadAllText(filePath);
                                if (string.IsNullOrWhiteSpace(json))
                                {
                                        return new List<HoldingInfo>();
                                }

                                var serializer = new JavaScriptSerializer();
                                var dtoList = serializer.Deserialize<List<HoldingDto>>(json) ?? new List<HoldingDto>();

                                return dtoList.Select(x => new HoldingInfo
                                {
                                        Ticker = x.Ticker,
                                        Name = x.Name,
                                        Quantity = x.Quantity,
                                        AverageCost = x.AverageCost,
                                        CurrentPrice = x.CurrentPrice,
                                        UpdatedAt = x.UpdatedAt
                                }).ToList();
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[庫存讀取失敗] {filePath}: {ex.Message}");
                                return new List<HoldingInfo>();
                        }
                }

                private void SaveHoldings(string market)
                {
                        try
                        {
                                if (!Directory.Exists(AppConfig.UserConfigDir))
                                {
                                        Directory.CreateDirectory(AppConfig.UserConfigDir);
                                }

                                var filePath = GetHoldingFile(market);
                                var dir = System.IO.Path.GetDirectoryName(filePath);
                                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                                {
                                        Directory.CreateDirectory(dir);
                                }

                                var source = market == "US" ? _usHoldingList : _twHoldingList;

                                var data = source.Select(x => new HoldingDto
                                {
                                        Ticker = x.Ticker,
                                        Name = x.Name,
                                        Quantity = x.Quantity,
                                        AverageCost = x.AverageCost,
                                        CurrentPrice = x.CurrentPrice,
                                        UpdatedAt = x.UpdatedAt
                                }).ToList();

                                var serializer = new JavaScriptSerializer();
                                var json = serializer.Serialize(data);
                                File.WriteAllText(filePath, json, Encoding.UTF8);
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[庫存儲存失敗] {market}: {ex.Message}");
                        }
                }

                private string GetUserStorageDir()
                {
                        var userId = string.IsNullOrWhiteSpace(_currentHoldingUser) ? "default" : _currentHoldingUser.Trim();
                        return System.IO.Path.Combine(AppConfig.UserConfigDir, "users", userId);
                }

                private string GetHoldingFile(string market)
                {
                        var fileName = market == "US" ? "us_holdings.json" : "tw_holdings.json";
                        return System.IO.Path.Combine(GetUserStorageDir(), fileName);
                }

                private string GetRealizedFile(string market)
                {
                        var fileName = market == "US" ? "us_realized_pnl.txt" : "tw_realized_pnl.txt";
                        return System.IO.Path.Combine(GetUserStorageDir(), fileName);
                }

                private void InitializeHoldingUsers()
                {
                        _holdingUsers = LoadHoldingUsersFromFile();
                        if (_holdingUsers.Count == 0)
                        {
                                _holdingUsers.Add("default");
                        }

                        if (!_holdingUsers.Contains(_currentHoldingUser, StringComparer.OrdinalIgnoreCase))
                        {
                                _currentHoldingUser = _holdingUsers[0];
                        }

                        cboHoldingUser.ItemsSource = _holdingUsers;
                        cboHoldingUser.SelectedItem = _currentHoldingUser;
                }

                private List<string> LoadHoldingUsersFromFile()
                {
                        try
                        {
                                if (!File.Exists(_holdingUsersFile))
                                {
                                        return new List<string>();
                                }

                                var json = File.ReadAllText(_holdingUsersFile);
                                var serializer = new JavaScriptSerializer();
                                var users = serializer.Deserialize<List<string>>(json) ?? new List<string>();
                                return users.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                        }
                        catch
                        {
                                return new List<string>();
                        }
                }

                private void SaveHoldingUsersToFile()
                {
                        try
                        {
                                if (!Directory.Exists(AppConfig.UserConfigDir))
                                {
                                        Directory.CreateDirectory(AppConfig.UserConfigDir);
                                }

                                var serializer = new JavaScriptSerializer();
                                var json = serializer.Serialize(_holdingUsers);
                                File.WriteAllText(_holdingUsersFile, json, Encoding.UTF8);
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[使用者清單儲存失敗] {ex.Message}");
                        }
                }

                private void StartMonitoring()
                {
                        _usMonitor.StartThreads(_updateInterval);
                        _twMonitor.StartThreads(_updateInterval);
                        statusText.Text = "監控已啟動";

                        // 🐍 啟動時立即使用 yfinance 抓取一次數據
                        Task.Run(() =>
                        {
                                PreloadStockData();
                        });
                }

                /// <summary>
                /// 🐍 預載股票數據（使用 Python yfinance）
                /// 在應用啟動時執行一次，立即獲取所有股票的最新價格
                /// </summary>
                private void PreloadStockData()
                {
                        try
                        {
                                Dispatcher.Invoke(() =>
                                {
                                        statusText.Text = "🐍 正在使用 yfinance 載入股價數據...";
                                });

                                Console.WriteLine("========================================");
                                Console.WriteLine("🐍 預載股票數據（Python yfinance）");
                                Console.WriteLine("========================================");
                                Console.WriteLine($"時間: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                                Console.WriteLine("");

                                // 預載美股數據
                                var usTickers = _usStockManager.GetTickers();
                                Console.WriteLine($"📊 美股: 開始載入 {usTickers.Count} 支股票");

                                int usSuccessCount = 0;
                                int usFailCount = 0;

                                for (int i = 0; i < usTickers.Count; i++)
                                {
                                        var ticker = usTickers[i];
                                        Console.WriteLine($"  [{i + 1}/{usTickers.Count}] 正在載入 {ticker}...");

                                        try
                                        {
                                                _usPriceFetcher.UpdatePriceWithPreviousClose(ticker);

                                                // 檢查是否成功
                                                var prices = _usPriceFetcher.GetPrices();
                                                if (prices.ContainsKey(ticker) && prices[ticker].Item1.HasValue)
                                                {
                                                        Console.WriteLine($"    ✅ 成功: ${prices[ticker].Item1.Value:F2}");
                                                        usSuccessCount++;
                                                }
                                                else
                                                {
                                                        Console.WriteLine($"    ⚠️  無數據");
                                                        usFailCount++;
                                                }
                                        }
                                        catch (Exception ex)
                                        {
                                                Console.WriteLine($"    ❌ 失敗: {ex.Message}");
                                                usFailCount++;
                                        }

                                        // 添加間隔避免請求過快
                                        if (i < usTickers.Count - 1)
                                        {
                                                System.Threading.Thread.Sleep(300); // 300ms 間隔
                                        }

                                        // 更新進度
                                        Dispatcher.Invoke(() =>
                                        {
                                                var progress = (i + 1) * 100.0 / usTickers.Count;
                                                progressBar.Value = progress;
                                                statusText.Text = $"🐍 載入美股 [{i + 1}/{usTickers.Count}] {ticker}...";
                                        });
                                }

                                Console.WriteLine($"📊 美股載入完成: 成功 {usSuccessCount}/{usTickers.Count}, 失敗 {usFailCount}");
                                Console.WriteLine("");

                                // 預載台股數據
                                var twTickers = _twStockManager.GetTickers();
                                Console.WriteLine($"📊 台股: 開始載入 {twTickers.Count} 支股票");

                                int twSuccessCount = 0;
                                int twFailCount = 0;

                                for (int i = 0; i < twTickers.Count; i++)
                                {
                                        var ticker = twTickers[i];
                                        Console.WriteLine($"  [{i + 1}/{twTickers.Count}] 正在載入 {ticker}...");

                                        try
                                        {
                                                _twPriceFetcher.UpdatePriceWithPreviousClose(ticker);

                                                // 檢查是否成功
                                                var prices = _twPriceFetcher.GetPrices();
                                                if (prices.ContainsKey(ticker) && prices[ticker].Item1.HasValue)
                                                {
                                                        Console.WriteLine($"    ✅ 成功: ${prices[ticker].Item1.Value:F2}");
                                                        twSuccessCount++;
                                                }
                                                else
                                                {
                                                        Console.WriteLine($"    ⚠️  無數據");
                                                        twFailCount++;
                                                }
                                        }
                                        catch (Exception ex)
                                        {
                                                Console.WriteLine($"    ❌ 失敗: {ex.Message}");
                                                twFailCount++;
                                        }

                                        // 添加間隔
                                        if (i < twTickers.Count - 1)
                                        {
                                                System.Threading.Thread.Sleep(300);
                                        }

                                        // 更新進度
                                        Dispatcher.Invoke(() =>
                                        {
                                                var progress = (i + 1) * 100.0 / twTickers.Count;
                                                progressBar.Value = progress;
                                                statusText.Text = $"🐍 載入台股 [{i + 1}/{twTickers.Count}] {ticker}...";
                                        });
                                }

                                Console.WriteLine($"📊 台股載入完成: 成功 {twSuccessCount}/{twTickers.Count}, 失敗 {twFailCount}");
                                Console.WriteLine("");
                                Console.WriteLine("========================================");
                                Console.WriteLine($"🎉 預載完成！總計: 美股 {usSuccessCount}/{usTickers.Count}, 台股 {twSuccessCount}/{twTickers.Count}");
                                Console.WriteLine("========================================");
                                Console.WriteLine("");

                                // 更新 UI 顯示
                                Dispatcher.Invoke(() =>
                                {
                                        UpdatePriceDisplay();
                                        progressBar.Value = 0;
                                        statusText.Text = $"✅ 預載完成 - 美股 {usSuccessCount}/{usTickers.Count}, 台股 {twSuccessCount}/{twTickers.Count}";
                                });
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"❌ 預載數據失敗: {ex.Message}");
                                Console.WriteLine($"堆疊追蹤: {ex.StackTrace}");

                                Dispatcher.Invoke(() =>
                                {
                                        statusText.Text = "⚠️ 預載失敗，將使用定時更新";
                                        progressBar.Value = 0;
                                });
                        }
                }

                private void UpdateTimer_Tick(object sender, EventArgs e)
                {
                        UpdatePriceDisplay();
                }

                private void CountdownTimer_Tick(object sender, EventArgs e)
                {
                        _countdownSeconds++;
                        if (_countdownSeconds >= _updateInterval)
                        {
                                _countdownSeconds = 0;
                        }

                        int remaining = _updateInterval - _countdownSeconds;
                        double progress = (_countdownSeconds / (double)_updateInterval) * 100;

                        progressBar.Value = progress;

                        var currentTab = marketTabControl.SelectedItem as TabItem;
                        var tabHeader = currentTab?.Header?.ToString() ?? "";
                        statusText.Text = $"下次更新: {remaining} 秒 | 當前市場: {tabHeader}";
                }

                private void UpdatePriceDisplay()
                {
                        // 更新美股價格顯示
                        var usPrices = _usPriceFetcher.GetPrices();
                        var usPriceMeta = _usPriceFetcher.GetPriceMeta();

                        Console.WriteLine($"=== 更新價格顯示 ({DateTime.Now:HH:mm:ss}) ===");
                        Console.WriteLine($"美股數據: {usPrices.Count} 筆");

                        foreach (var stock in _usStockList)
                        {
                                if (usPrices.ContainsKey(stock.Ticker))
                                {
                                        var priceData = usPrices[stock.Ticker];
                                        stock.Price = priceData.Item1;

                                        // 🎯 自己計算漲跌幅：從 Meta 獲取前收盤價
                                        if (usPriceMeta.ContainsKey(stock.Ticker))
                                        {
                                                var meta = usPriceMeta[stock.Ticker];

                                                // 獲取前收盤價
                                                double? previousClose = null;
                                                if (meta.ContainsKey("previous_close") && meta["previous_close"] != null)
                                                {
                                                        previousClose = meta["previous_close"] as double?;
                                                }

                                                // 設置前收盤價到 UI
                                                stock.PreviousClose = previousClose;

                                                // 手動計算漲跌幅
                                                if (stock.Price.HasValue && previousClose.HasValue && previousClose.Value != 0)
                                                {
                                                        var change = stock.Price.Value - previousClose.Value;
                                                        stock.ChangePercent = (change / previousClose.Value) * 100;

                                                        Console.WriteLine($"  [手動計算] {stock.Ticker}: 當前=${stock.Price.Value:F2}, 前收=${previousClose.Value:F2}, 漲跌幅={stock.ChangePercent.Value:F2}%");
                                                }
                                                else
                                                {
                                                        // 如果沒有前收盤價，使用 API 返回的漲跌幅
                                                        stock.ChangePercent = priceData.Item2;

                                                        if (stock.ChangePercent.HasValue)
                                                        {
                                                                Console.WriteLine($"  [API數據] {stock.Ticker}: Price={stock.Price?.ToString("F2") ?? "null"}, Change={stock.ChangePercent.Value:F2}%");
                                                        }
                                                }

                                                // 更新來源和時間
                                                stock.Source = meta.ContainsKey("source") ? meta["source"]?.ToString() : "N/A";
                                                stock.UpdatedAt = meta.ContainsKey("updated_at") ? meta["updated_at"] as DateTime? : null;
                                        }
                                        else
                                        {
                                                // 沒有 Meta 數據，使用 API 返回的漲跌幅
                                                stock.ChangePercent = priceData.Item2;
                                        }

                                        // 輝達特殊診斷
                                        if (stock.Name == "輝達")
                                        {
                                                Console.WriteLine($"\n========== 輝達 (NVDA) 詳細診斷 ==========");
                                                Console.WriteLine($"當前價格: ${stock.Price?.ToString("F2") ?? "null"}");
                                                Console.WriteLine($"漲跌幅: {stock.ChangePercent?.ToString("F2") ?? "null"}%");

                                                if (stock.Price.HasValue && stock.ChangePercent.HasValue)
                                                {
                                                        // 反推系統使用的前收盤價
                                                        double impliedPreviousClose = stock.Price.Value / (1 + stock.ChangePercent.Value / 100);
                                                        Console.WriteLine($"系統使用的前收盤價: ${impliedPreviousClose:F2}");

                                                        // 與正確值比較
                                                        double correctPreviousClose = 195.56;
                                                        Console.WriteLine($"正確的前收盤價: ${correctPreviousClose:F2}");
                                                        Console.WriteLine($"差異: ${Math.Abs(impliedPreviousClose - correctPreviousClose):F2}");

                                                        if (Math.Abs(impliedPreviousClose - correctPreviousClose) < 0.10)
                                                        {
                                                                Console.WriteLine($"✅ 前收盤價正確！");
                                                        }
                                                        else
                                                        {
                                                                Console.WriteLine($"❌ 前收盤價錯誤！");
                                                        }
                                                }
                                                Console.WriteLine($"==========================================\n");
                                        }
                                }
                        }

                        // 更新台股價格顯示
                        var twPrices = _twPriceFetcher.GetPrices();
                        var twPriceMeta = _twPriceFetcher.GetPriceMeta();

                        Console.WriteLine($"台股數據: {twPrices.Count} 筆");

                        foreach (var stock in _twStockList)
                        {
                                if (twPrices.ContainsKey(stock.Ticker))
                                {
                                        var priceData = twPrices[stock.Ticker];
                                        stock.Price = priceData.Item1;
                                        stock.ChangePercent = priceData.Item2;

                                        // 從 Meta 獲取前收盤價（台股）
                                        if (twPriceMeta.ContainsKey(stock.Ticker))
                                        {
                                                var meta = twPriceMeta[stock.Ticker];

                                                double? previousClose = null;
                                                if (meta.ContainsKey("previous_close") && meta["previous_close"] != null)
                                                {
                                                        previousClose = meta["previous_close"] as double?;
                                                }

                                                // 設置前收盤價到 UI
                                                stock.PreviousClose = previousClose;

                                                // 優先用前收盤價自行計算漲跌幅
                                                if (stock.Price.HasValue && previousClose.HasValue && previousClose.Value != 0)
                                                {
                                                        var change = stock.Price.Value - previousClose.Value;
                                                        stock.ChangePercent = (change / previousClose.Value) * 100;
                                                }
                                        }

                                        // 詳細診斷日誌
                                        var priceStr = priceData.Item1?.ToString("F2") ?? "null";
                                        var changeStr = stock.ChangePercent?.ToString("F2") ?? "null";
                                        var hasChange = stock.ChangePercent.HasValue ? "✅" : "❌";
                                        Console.WriteLine($"  {stock.Ticker}: Price={priceStr}, Change={changeStr}% {hasChange}");
                                }
                                if (twPriceMeta.ContainsKey(stock.Ticker))
                                {
                                        var meta = twPriceMeta[stock.Ticker];
                                        stock.Source = meta.ContainsKey("source") ? meta["source"]?.ToString() : "N/A";
                                        stock.UpdatedAt = meta.ContainsKey("updated_at") ? meta["updated_at"] as DateTime? : null;
                                }
                        }

                        UpdateHoldingPrices("US");
                        UpdateHoldingPrices("TW");

                        dgUsStocks.Items.Refresh();
                        dgTwStocks.Items.Refresh();
                        dgUsHoldings.Items.Refresh();
                        dgTwHoldings.Items.Refresh();
                        _ = UpdateNewsImpactAsync();
                        Console.WriteLine("=== UI 已刷新 ===\n");
                }

                private void UpdateHoldingPrices(string market)
                {
                        var stocks = market == "US" ? _usStockList : _twStockList;
                        var holdings = market == "US" ? _usHoldingList : _twHoldingList;

                        foreach (var holding in holdings)
                        {
                                var stock = stocks.FirstOrDefault(x => string.Equals(x.Ticker, holding.Ticker, StringComparison.OrdinalIgnoreCase));
                                if (stock != null)
                                {
                                        holding.Name = stock.Name;
                                        holding.CurrentPrice = stock.Price;
                                        holding.UpdatedAt = stock.UpdatedAt;
                                }
                        }

                        UpdateHoldingSummary(market);
                }

                private void UpdateHoldingSummary(string market)
                {
                        var holdings = market == "US" ? _usHoldingList : _twHoldingList;
                        var totalCost = holdings.Sum(x => x.CostAmount);
                        var totalPnl = holdings.Sum(x => x.UnrealizedPnL);
                        var totalPnlPct = totalCost > 0 ? (totalPnl / totalCost) * 100 : 0;
                        var realized = market == "US" ? _usRealizedPnL : _twRealizedPnL;

                        var text = market == "US"
                                ? $"美股未實現：{totalPnl:N2} ({totalPnlPct:F2}%)"
                                : $"台股未實現：{totalPnl:N2} ({totalPnlPct:F2}%)";

                        var realizedText = market == "US"
                                ? $"美股已實現：{realized:N2}"
                                : $"台股已實現：{realized:N2}";

                        if (market == "US")
                        {
                                txtUsTotalPnL.Text = text;
                                txtUsTotalPnL.Foreground = totalPnl >= 0 ? new SolidColorBrush(Color.FromRgb(39, 174, 96)) : new SolidColorBrush(Color.FromRgb(231, 76, 60));
                                txtUsRealizedPnL.Text = realizedText;
                                txtUsRealizedPnL.Foreground = realized >= 0 ? new SolidColorBrush(Color.FromRgb(39, 174, 96)) : new SolidColorBrush(Color.FromRgb(231, 76, 60));
                        }
                        else
                        {
                                txtTwTotalPnL.Text = text;
                                txtTwTotalPnL.Foreground = totalPnl >= 0 ? new SolidColorBrush(Color.FromRgb(39, 174, 96)) : new SolidColorBrush(Color.FromRgb(231, 76, 60));
                                txtTwRealizedPnL.Text = realizedText;
                                txtTwRealizedPnL.Foreground = realized >= 0 ? new SolidColorBrush(Color.FromRgb(39, 174, 96)) : new SolidColorBrush(Color.FromRgb(231, 76, 60));
                        }
                }

                private async Task UpdateNewsImpactAsync()
                {
                        if (_isNewsUpdating)
                        {
                                return;
                        }

                        if ((DateTime.Now - _lastNewsUpdate).TotalSeconds < 60)
                        {
                                return;
                        }

                        _isNewsUpdating = true;
                        try
                        {
                                var currentTab = marketTabControl.SelectedIndex;
                                var isUsMarket = currentTab == 0;
                                var keyword = isUsMarket ? "US stock market finance" : "台股 財經";

                                var newsItems = await Task.Run(() => GetLatestNews(keyword, isUsMarket));

                                txtMaxImpactStock.Text = $"{(isUsMarket ? "美股" : "台股")}新聞｜更新時間：{DateTime.Now:HH:mm:ss}";

                                _newsImpactList.Clear();
                                foreach (var item in newsItems)
                                {
                                        _newsImpactList.Add(item);
                                }

                                _lastNewsUpdate = DateTime.Now;
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[新聞影響] 更新失敗: {ex.Message}");
                        }
                        finally
                        {
                                _isNewsUpdating = false;
                        }
                }

                private List<NewsImpactItem> GetLatestNews(string keyword, bool isUsMarket)
                {
                        var result = new List<NewsImpactItem>();
                        try
                        {
                                var q = Uri.EscapeDataString(keyword);
                                var url = isUsMarket
                                        ? $"https://news.google.com/rss/search?q={q}&hl=en-US&gl=US&ceid=US:en"
                                        : $"https://news.google.com/rss/search?q={q}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant";

                                using (var client = new WebClient())
                                {
                                        client.Encoding = Encoding.UTF8;
                                        var xml = client.DownloadString(url);
                                        var doc = XDocument.Parse(xml);
                                        var items = doc.Descendants("item").Take(12);

                                        foreach (var item in items)
                                        {
                                                var originalTitle = item.Element("title")?.Value ?? "(無標題)";
                                                var title = originalTitle;
                                                if (isUsMarket)
                                                {
                                                        title = TranslateHeadlineToTraditionalChinese(title);
                                                }

                                                var source = item.Element("source")?.Value ?? "新聞來源";
                                                DateTime published;
                                                var pubDateText = item.Element("pubDate")?.Value;
                                                var showTime = DateTime.TryParse(pubDateText, out published)
                                                        ? published.ToString("MM-dd HH:mm")
                                                        : "--";

                                                var importanceScore = CalculateNewsImportance(originalTitle, source, showTime == "--" ? (DateTime?)null : published);

                                                result.Add(new NewsImpactItem
                                                {
                                                        Time = showTime,
                                                        Headline = title,
                                                        Source = source,
                                                        ImportanceScore = importanceScore,
                                                        ImportanceLevel = GetImportanceLevel(importanceScore)
                                                });
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[最新新聞] 載入失敗: {ex.Message}");
                        }

                        if (result.Count == 0)
                        {
                                result.Add(new NewsImpactItem
                                {
                                        Time = "--",
                                        Headline = "目前無法取得最新新聞",
                                        Source = "系統",
                                        ImportanceScore = 0,
                                        ImportanceLevel = "--"
                                });
                        }

                        return result;
                }

                private string TranslateHeadlineToTraditionalChinese(string text)
                {
                        if (string.IsNullOrWhiteSpace(text))
                        {
                                return text;
                        }

                        // 已含中文就不翻譯
                        if (Regex.IsMatch(text, "[\u4e00-\u9fff]"))
                        {
                                return text;
                        }

                        if (_translationCache.ContainsKey(text))
                        {
                                return _translationCache[text];
                        }

                        try
                        {
                                var q = Uri.EscapeDataString(text);
                                var url = $"https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=zh-TW&dt=t&q={q}";

                                using (var client = new WebClient())
                                {
                                        client.Encoding = Encoding.UTF8;
                                        var response = client.DownloadString(url);

                                        var match = Regex.Match(response, "^\\[\\[\\[\"(?<translated>.*?)\"");
                                        if (match.Success)
                                        {
                                                var translated = Regex.Unescape(match.Groups["translated"].Value);
                                                translated = translated.Replace("\\n", " ");
                                                _translationCache[text] = translated;
                                                return translated;
                                        }
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[新聞翻譯] 失敗: {ex.Message}");
                        }

                        _translationCache[text] = text;
                        return text;
                }

                private int CalculateNewsImportance(string title, string source, DateTime? publishedAt)
                {
                        if (string.IsNullOrWhiteSpace(title))
                        {
                                return 20;
                        }

                        var t = title.ToLowerInvariant();
                        var score = 10;

                        var highImpactKeywords = new[]
                        {
                                "earnings", "guidance", "fed", "rate", "inflation", "bankruptcy", "lawsuit", "merger", "acquisition",
                                "財報", "升息", "降息", "通膨", "併購", "裁員", "訴訟", "破產", "停止交易", "停牌"
                        };

                        var mediumImpactKeywords = new[]
                        {
                                "analyst", "target", "rating", "forecast", "outlook", "estimate",
                                "目標價", "評級", "展望", "預測", "估值", "庫存", "買入", "賣出"
                        };

                        var highHitCount = highImpactKeywords.Count(k => t.Contains(k));
                        var mediumHitCount = mediumImpactKeywords.Count(k => t.Contains(k));

                        score += Math.Min(highHitCount * 18, 54);
                        score += Math.Min(mediumHitCount * 8, 24);

                        if (Regex.IsMatch(t, "\\d+(\\.\\d+)?%"))
                        {
                                score += 8;
                        }

                        if (t.Contains("breaking") || t.Contains("urgent") || t.Contains("突發") || t.Contains("快訊"))
                        {
                                score += 10;
                        }

                        if (t.Contains("surge") || t.Contains("plunge") || t.Contains("rally") ||
                                t.Contains("暴跌") || t.Contains("重挫") || t.Contains("大漲") || t.Contains("飆升"))
                        {
                                score += 10;
                        }

                        if (t.Contains("rumor") || t.Contains("傳聞") || t.Contains("市場傳言"))
                        {
                                score -= 8;
                        }

                        var s = (source ?? string.Empty).ToLowerInvariant();
                        if (s.Contains("reuters") || s.Contains("bloomberg") || s.Contains("wsj") ||
                                s.Contains("cnbc") || s.Contains("financial times") || s.Contains("associated press"))
                        {
                                score += 8;
                        }

                        if (publishedAt.HasValue)
                        {
                                var ageMinutes = Math.Max(0, (DateTime.Now - publishedAt.Value).TotalMinutes);
                                if (ageMinutes <= 30) score += 10;
                                else if (ageMinutes <= 120) score += 6;
                                else if (ageMinutes <= 1440) score += 3;
                        }

                        // 若關鍵字都沒命中，給保守基準分
                        if (highHitCount == 0 && mediumHitCount == 0)
                        {
                                score += 20;
                        }

                        return Math.Max(0, Math.Min(100, score));
                }

                private string GetImportanceLevel(int score)
                {
                        if (score >= 75) return "高";
                        if (score >= 50) return "中";
                        return "低";
                }

                private void BtnAddStock_Click(object sender, RoutedEventArgs e)
                {
                        var dialog = new AddStockDialog();
                        if (dialog.ShowDialog() == true)
                        {
                                var ticker = dialog.Ticker.ToUpper();
                                var name = dialog.StockName;
                                var market = dialog.Market;

                                if (market == "US")
                                {
                                        _usStockManager.AddStock(ticker, name);
                                }
                                else
                                {
                                        _twStockManager.AddStock(ticker, name);
                                }

                                LoadStockData();
                                MessageBox.Show($"已新增股票: {ticker} - {name}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                }

                private void BtnRemoveStock_Click(object sender, RoutedEventArgs e)
                {
                        var currentTab = marketTabControl.SelectedIndex;
                        DataGrid currentGrid = currentTab == 0 ? dgUsStocks : dgTwStocks;

                        if (currentGrid.SelectedItem is StockInfo selectedStock)
                        {
                                var result = MessageBox.Show(
                                        $"確定要移除股票 {selectedStock.Ticker} - {selectedStock.Name} 嗎？",
                                        "確認移除",
                                        MessageBoxButton.YesNo,
                                        MessageBoxImage.Question);

                                if (result == MessageBoxResult.Yes)
                                {
                                        if (currentTab == 0)
                                        {
                                                _usStockManager.RemoveStock(selectedStock.Ticker);
                                        }
                                        else
                                        {
                                                _twStockManager.RemoveStock(selectedStock.Ticker);
                                        }

                                        LoadStockData();
                                        MessageBox.Show("已移除股票", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                                }
                        }
                        else
                        {
                                MessageBox.Show("請先選擇要移除的股票", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                }

                private void BtnTwFilter_Click(object sender, RoutedEventArgs e)
                {
                        var filterWindow = new TwStockFilterWindow();
                        filterWindow.Owner = this;
                        filterWindow.Show();
                }

                private async void BtnTwYFinanceCacheUpdate_Click(object sender, RoutedEventArgs e)
                {
                        var button = sender as Button;
                        var cacheProgressBar = FindName("pbTwSectorCacheUpdate") as ProgressBar;
                        var cachePercentText = FindName("txtTwSectorCachePercent") as TextBlock;
                        var cacheEtaText = FindName("txtTwSectorCacheEta") as TextBlock;
                        if (button != null)
                        {
                                button.IsEnabled = false;
                        }

                        if (cacheProgressBar != null)
                        {
                                cacheProgressBar.Value = 0;
                                cacheProgressBar.Visibility = Visibility.Visible;
                        }
                        if (cachePercentText != null)
                        {
                                cachePercentText.Text = "0%";
                                cachePercentText.Visibility = Visibility.Visible;
                        }
                        if (cacheEtaText != null)
                        {
                                cacheEtaText.Text = "預估剩餘 --:--";
                                cacheEtaText.Visibility = Visibility.Visible;
                        }

                        statusText.Text = "台股族群 yfinance 快取更新中...";

                        try
                        {
                                var tickers = LoadTwSectorUniverseTickersForCache();
                                if (tickers.Count == 0)
                                {
                                        throw new InvalidOperationException("找不到可更新的台股族群股票清單。");
                                }

                                await Task.Run(() =>
                                {
                                        var total = Math.Max(1, tickers.Count);
                                        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                                        for (int i = 0; i < tickers.Count; i++)
                                        {
                                                var t = tickers[i];
                                                var normalized = t.EndsWith(".TW", StringComparison.OrdinalIgnoreCase) ? t : (t + ".TW");
                                                _twPriceFetcher.UpdatePriceWithPreviousClose(normalized);

                                                var percent = (i + 1) * 100.0 / total;
                                                var elapsed = stopwatch.Elapsed.TotalSeconds;
                                                var avgPerItem = elapsed / Math.Max(1, i + 1);
                                                var remainingSeconds = Math.Max(0, (total - (i + 1)) * avgPerItem);
                                                var eta = TimeSpan.FromSeconds(remainingSeconds);
                                                Dispatcher.BeginInvoke(new Action(() =>
                                                {
                                                        if (cacheProgressBar != null)
                                                        {
                                                                cacheProgressBar.Value = percent;
                                                        }
                                                        if (cachePercentText != null)
                                                        {
                                                                cachePercentText.Text = $"{percent:F0}%";
                                                        }
                                                        if (cacheEtaText != null)
                                                        {
                                                                cacheEtaText.Text = $"預估剩餘 {eta:mm\\:ss}";
                                                        }
                                                }));

                                                System.Threading.Thread.Sleep(80);
                                        }
                                });

                                var prices = _twPriceFetcher.GetPrices();
                                var meta = _twPriceFetcher.GetPriceMeta();
                                var cacheItems = new List<Dictionary<string, object>>();

                                foreach (var t in tickers)
                                {
                                        var normalized = t.EndsWith(".TW", StringComparison.OrdinalIgnoreCase) ? t : (t + ".TW");
                                        Tuple<double?, double?> priceTuple;
                                        if (!prices.TryGetValue(normalized, out priceTuple))
                                        {
                                                continue;
                                        }

                                        var currentPrice = priceTuple.Item1;
                                        double? previousClose = null;
                                        Dictionary<string, object> itemMeta;
                                        if (meta.TryGetValue(normalized, out itemMeta) && itemMeta != null && itemMeta.ContainsKey("previous_close"))
                                        {
                                                previousClose = itemMeta["previous_close"] as double?;
                                        }

                                        double? changePercent;
                                        if (currentPrice.HasValue && previousClose.HasValue && Math.Abs(previousClose.Value) > 0.000001)
                                        {
                                                changePercent = (currentPrice.Value - previousClose.Value) / previousClose.Value * 100;
                                        }
                                        else
                                        {
                                                changePercent = priceTuple.Item2;
                                        }

                                        cacheItems.Add(new Dictionary<string, object>
                                        {
                                                { "Ticker", t.EndsWith(".TW", StringComparison.OrdinalIgnoreCase) ? t.Substring(0, t.Length - 3) : t },
                                                { "Price", currentPrice },
                                                { "PreviousClose", previousClose },
                                                { "ChangePercent", changePercent },
                                                { "UpdatedAt", DateTime.Now }
                                        });
                                }

                                if (!Directory.Exists(AppConfig.UserConfigDir))
                                {
                                        Directory.CreateDirectory(AppConfig.UserConfigDir);
                                }

                                var payload = new Dictionary<string, object>
                                {
                                        { "Date", DateTime.Today.ToString("yyyy-MM-dd") },
                                        { "Items", cacheItems }
                                };

                                var serializer = new JavaScriptSerializer();
                                var cacheFile = System.IO.Path.Combine(AppConfig.UserConfigDir, "tw_sector_yfinance_cache.json");
                                File.WriteAllText(cacheFile, serializer.Serialize(payload), Encoding.UTF8);

                                statusText.Text = $"✅ 台股族群 yfinance 快取已更新（{cacheItems.Count} 檔）";
                                if (cachePercentText != null)
                                {
                                        cachePercentText.Text = "100%";
                                }
                                if (cacheEtaText != null)
                                {
                                        cacheEtaText.Text = "預估剩餘 00:00";
                                }
                                MessageBox.Show($"台股族群 yfinance 快取更新完成，共 {cacheItems.Count} 檔。", "更新完成", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        catch (Exception ex)
                        {
                                statusText.Text = "❌ 台股族群 yfinance 快取更新失敗";
                                MessageBox.Show($"更新台股族群 yfinance 快取失敗：{ex.Message}", "更新失敗", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        finally
                        {
                                if (cacheProgressBar != null)
                                {
                                        cacheProgressBar.Visibility = Visibility.Collapsed;
                                }
                                if (cachePercentText != null)
                                {
                                        cachePercentText.Visibility = Visibility.Collapsed;
                                }
                                if (cacheEtaText != null)
                                {
                                        cacheEtaText.Visibility = Visibility.Collapsed;
                                }

                                if (button != null)
                                {
                                        button.IsEnabled = true;
                                }
                        }
                }

                private List<string> LoadTwSectorUniverseTickersForCache()
                {
                        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        List<string> sectorOrder;
                        Dictionary<string, string> tickerToSector;
                        if (AppConfig.TryLoadTaiwanSectorCsv(out sectorOrder, out tickerToSector) && tickerToSector != null)
                        {
                                foreach (var t in tickerToSector.Keys)
                                {
                                        var normalized = (t ?? string.Empty).Trim().ToUpperInvariant();
                                        if (normalized.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
                                        {
                                                normalized = normalized.Substring(0, normalized.Length - 3);
                                        }

                                        if (!string.IsNullOrWhiteSpace(normalized))
                                        {
                                                result.Add(normalized);
                                        }
                                }
                        }

                        try
                        {
                                var url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL";
                                using (var client = new WebClient())
                                {
                                        client.Encoding = Encoding.UTF8;
                                        var json = client.DownloadString(url);
                                        var serializer = new JavaScriptSerializer();
                                        var rows = serializer.Deserialize<List<Dictionary<string, string>>>(json) ?? new List<Dictionary<string, string>>();

                                        foreach (var row in rows)
                                        {
                                                string code;
                                                if (row.TryGetValue("Code", out code) && !string.IsNullOrWhiteSpace(code))
                                                {
                                                        result.Add(code.Trim().ToUpperInvariant());
                                                }
                                        }
                                }
                        }
                        catch
                        {
                        }

                        return result.OrderBy(x => x).ToList();
                }

                private void BtnUsAddHolding_Click(object sender, RoutedEventArgs e)
                {
                        AddHolding("US", txtUsHoldingTicker, txtUsHoldingQty, txtUsHoldingCost);
                }

                private void BtnTwAddHolding_Click(object sender, RoutedEventArgs e)
                {
                        AddHolding("TW", txtTwHoldingTicker, txtTwHoldingQty, txtTwHoldingCost);
                }

                private void BtnUsRemoveHolding_Click(object sender, RoutedEventArgs e)
                {
                        RemoveHolding("US", dgUsHoldings);
                }

                private void BtnTwRemoveHolding_Click(object sender, RoutedEventArgs e)
                {
                        RemoveHolding("TW", dgTwHoldings);
                }

                private void BtnUsResetHoldings_Click(object sender, RoutedEventArgs e)
                {
                        ResetHoldings("US");
                }

                private void BtnTwResetHoldings_Click(object sender, RoutedEventArgs e)
                {
                        ResetHoldings("TW");
                }

                private void BtnUsUpdateHolding_Click(object sender, RoutedEventArgs e)
                {
                        UpdateHolding("US", dgUsHoldings, txtUsHoldingTicker, txtUsHoldingQty, txtUsHoldingCost);
                }

                private void BtnTwUpdateHolding_Click(object sender, RoutedEventArgs e)
                {
                        UpdateHolding("TW", dgTwHoldings, txtTwHoldingTicker, txtTwHoldingQty, txtTwHoldingCost);
                }

                private void BtnUsSellHolding_Click(object sender, RoutedEventArgs e)
                {
                        SellHolding("US", dgUsHoldings, txtUsSellQty, txtUsSellPrice);
                }

                private void BtnTwSellHolding_Click(object sender, RoutedEventArgs e)
                {
                        SellHolding("TW", dgTwHoldings, txtTwSellQty, txtTwSellPrice);
                }

                private void DgUsHoldings_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                        FillHoldingEditor("US", dgUsHoldings, txtUsHoldingTicker, txtUsHoldingQty, txtUsHoldingCost, txtUsSellPrice);
                }

                private void DgTwHoldings_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                        FillHoldingEditor("TW", dgTwHoldings, txtTwHoldingTicker, txtTwHoldingQty, txtTwHoldingCost, txtTwSellPrice);
                }

                private void AddHolding(string market, TextBox tickerTextBox, TextBox qtyTextBox, TextBox costTextBox)
                {
                        var ticker = (tickerTextBox.Text ?? string.Empty).Trim().ToUpperInvariant();
                        if (market == "TW" && !ticker.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
                        {
                                ticker += ".TW";
                        }

                        if (string.IsNullOrWhiteSpace(ticker))
                        {
                                MessageBox.Show("請輸入股票代號", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        if (!double.TryParse((qtyTextBox.Text ?? string.Empty).Trim(), out double qty) || qty <= 0)
                        {
                                MessageBox.Show("請輸入正確股數", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        if (!double.TryParse((costTextBox.Text ?? string.Empty).Trim(), out double avgCost) || avgCost <= 0)
                        {
                                MessageBox.Show("請輸入正確均價", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        var stockList = market == "US" ? _usStockList : _twStockList;
                        var holdingList = market == "US" ? _usHoldingList : _twHoldingList;

                        var stock = stockList.FirstOrDefault(x => string.Equals(x.Ticker, ticker, StringComparison.OrdinalIgnoreCase));
                        var name = stock?.Name ?? ticker;
                        var currentPrice = stock?.Price;

                        var existing = holdingList.FirstOrDefault(x => string.Equals(x.Ticker, ticker, StringComparison.OrdinalIgnoreCase));
                        if (existing != null)
                        {
                                var totalCostAmount = existing.CostAmount + (avgCost * qty);
                                var totalQty = existing.Quantity + qty;

                                existing.Quantity = totalQty;
                                existing.AverageCost = totalQty > 0 ? totalCostAmount / totalQty : 0;
                                existing.CurrentPrice = currentPrice;
                                existing.UpdatedAt = DateTime.Now;
                        }
                        else
                        {
                                holdingList.Add(new HoldingInfo
                                {
                                        Ticker = ticker,
                                        Name = name,
                                        Quantity = qty,
                                        AverageCost = avgCost,
                                        CurrentPrice = currentPrice,
                                        UpdatedAt = DateTime.Now
                                });
                        }

                        SaveHoldings(market);
                        UpdateHoldingSummary(market);

                        tickerTextBox.Clear();
                        qtyTextBox.Clear();
                        costTextBox.Clear();
                }

                private void RemoveHolding(string market, DataGrid grid)
                {
                        if (!(grid.SelectedItem is HoldingInfo selected))
                        {
                                MessageBox.Show("請先選擇要移除的庫存", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        var result = MessageBox.Show($"確定移除庫存 {selected.Ticker} 嗎？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (result != MessageBoxResult.Yes)
                        {
                                return;
                        }

                        var holdingList = market == "US" ? _usHoldingList : _twHoldingList;
                        holdingList.Remove(selected);
                        SaveHoldings(market);
                        UpdateHoldingSummary(market);
                }

                private void UpdateHolding(string market, DataGrid grid, TextBox tickerTextBox, TextBox qtyTextBox, TextBox costTextBox)
                {
                        if (!(grid.SelectedItem is HoldingInfo selected))
                        {
                                MessageBox.Show("請先選擇要修改的庫存", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        var ticker = (tickerTextBox.Text ?? string.Empty).Trim().ToUpperInvariant();
                        if (market == "TW" && !ticker.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
                        {
                                ticker += ".TW";
                        }

                        if (string.IsNullOrWhiteSpace(ticker))
                        {
                                MessageBox.Show("請輸入股票代號", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        if (!double.TryParse((qtyTextBox.Text ?? string.Empty).Trim(), out double qty) || qty <= 0)
                        {
                                MessageBox.Show("請輸入正確股數", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        if (!double.TryParse((costTextBox.Text ?? string.Empty).Trim(), out double avgCost) || avgCost <= 0)
                        {
                                MessageBox.Show("請輸入正確均價", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        selected.Ticker = ticker;
                        selected.Quantity = qty;
                        selected.AverageCost = avgCost;
                        selected.UpdatedAt = DateTime.Now;

                        SaveHoldings(market);
                        UpdateHoldingPrices(market);
                }

                private void SellHolding(string market, DataGrid grid, TextBox sellQtyTextBox, TextBox sellPriceTextBox)
                {
                        if (!(grid.SelectedItem is HoldingInfo selected))
                        {
                                MessageBox.Show("請先選擇要賣出的庫存", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        if (!double.TryParse((sellQtyTextBox.Text ?? string.Empty).Trim(), out double sellQty) || sellQty <= 0)
                        {
                                MessageBox.Show("請輸入正確賣出股數", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        if (sellQty > selected.Quantity)
                        {
                                MessageBox.Show("賣出股數不可超過持有股數", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        var sellPriceText = (sellPriceTextBox.Text ?? string.Empty).Trim();
                        double sellPrice;
                        if (!string.IsNullOrWhiteSpace(sellPriceText))
                        {
                                if (!double.TryParse(sellPriceText, out sellPrice) || sellPrice <= 0)
                                {
                                        MessageBox.Show("請輸入正確賣出價格", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                        return;
                                }
                        }
                        else if (selected.CurrentPrice.HasValue && selected.CurrentPrice.Value > 0)
                        {
                                sellPrice = selected.CurrentPrice.Value;
                        }
                        else
                        {
                                MessageBox.Show("請輸入賣出價格（目前無即時價格）", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                                return;
                        }

                        var realized = (sellPrice - selected.AverageCost) * sellQty;
                        if (market == "US")
                        {
                                _usRealizedPnL += realized;
                        }
                        else
                        {
                                _twRealizedPnL += realized;
                        }

                        selected.Quantity -= sellQty;
                        selected.UpdatedAt = DateTime.Now;

                        var holdings = market == "US" ? _usHoldingList : _twHoldingList;
                        if (selected.Quantity <= 0.000001)
                        {
                                holdings.Remove(selected);
                        }

                        SaveHoldings(market);
                        SaveRealizedPnL(market);
                        UpdateHoldingSummary(market);

                        sellQtyTextBox.Clear();
                        sellPriceTextBox.Clear();
                }

                private void FillHoldingEditor(string market, DataGrid grid, TextBox tickerTextBox, TextBox qtyTextBox, TextBox costTextBox, TextBox sellPriceTextBox)
                {
                        if (!(grid.SelectedItem is HoldingInfo selected))
                        {
                                return;
                        }

                        tickerTextBox.Text = selected.Ticker;
                        qtyTextBox.Text = selected.Quantity.ToString("F2");
                        costTextBox.Text = selected.AverageCost.ToString("F2");
                        sellPriceTextBox.Text = selected.CurrentPrice.HasValue ? selected.CurrentPrice.Value.ToString("F2") : string.Empty;
                }

                private void ResetHoldings(string market)
                {
                        var marketLabel = market == "US" ? "美股" : "台股";
                        var result = MessageBox.Show(
                                $"確定要重置{marketLabel}庫存嗎？\n\n這會清空庫存與已實現損益。",
                                "確認重置",
                                MessageBoxButton.YesNo,
                                MessageBoxImage.Warning);

                        if (result != MessageBoxResult.Yes)
                        {
                                return;
                        }

                        if (market == "US")
                        {
                                _usHoldingList.Clear();
                                _usRealizedPnL = 0;

                                txtUsHoldingTicker.Clear();
                                txtUsHoldingQty.Clear();
                                txtUsHoldingCost.Clear();
                                txtUsSellQty.Clear();
                                txtUsSellPrice.Clear();
                        }
                        else
                        {
                                _twHoldingList.Clear();
                                _twRealizedPnL = 0;

                                txtTwHoldingTicker.Clear();
                                txtTwHoldingQty.Clear();
                                txtTwHoldingCost.Clear();
                                txtTwSellQty.Clear();
                                txtTwSellPrice.Clear();
                        }

                        SaveHoldings(market);
                        SaveRealizedPnL(market);
                        UpdateHoldingSummary(market);
                }

                public bool TryAddTwStockFromFilter(string ticker, string name, out string message)
                {
                        if (string.IsNullOrWhiteSpace(ticker))
                        {
                                message = "股票代號無效，無法加入。";
                                return false;
                        }

                        var normalizedTicker = ticker.Trim().ToUpperInvariant();
                        if (!normalizedTicker.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
                        {
                                normalizedTicker += ".TW";
                        }

                        var twStocks = _twStockManager.GetStocks();
                        if (twStocks.ContainsKey(normalizedTicker))
                        {
                                message = $"{normalizedTicker} 已在主頁面台股清單中。";
                                return false;
                        }

                        var stockName = string.IsNullOrWhiteSpace(name) ? normalizedTicker : name.Trim();
                        _twStockManager.AddStock(normalizedTicker, stockName);
                        LoadStockData();

                        message = $"已加入主頁面清單：{normalizedTicker} - {stockName}";
                        statusText.Text = message;
                        return true;
                }

                private void BtnRefresh_Click(object sender, RoutedEventArgs e)
                {
                        // 檢查冷卻時間
                        var now = DateTime.Now;
                        var timeSinceLastRefresh = (now - _lastManualRefresh).TotalSeconds;

                        if (timeSinceLastRefresh < MANUAL_REFRESH_COOLDOWN)
                        {
                                var remaining = MANUAL_REFRESH_COOLDOWN - (int)timeSinceLastRefresh;
                                statusText.Text = $"⏰ 請稍後 {remaining} 秒再刷新";
                                Console.WriteLine($"[手動刷新] 冷卻中，剩餘 {remaining} 秒");
                                return;
                        }

                        // 更新最後刷新時間
                        _lastManualRefresh = now;
                        _countdownSeconds = 0;
                        statusText.Text = "🔄 手動刷新中...";
                        Console.WriteLine($"[手動刷新] 開始刷新，下次可刷新時間: {_lastManualRefresh.AddSeconds(MANUAL_REFRESH_COOLDOWN):HH:mm:ss}");

                        Task.Run(() =>
                        {
                                var currentTab = marketTabControl.Dispatcher.Invoke(() => marketTabControl.SelectedIndex);

                                if (currentTab == 0)
                                {
                                        var tickers = _usStockManager.GetTickers();
                                        Console.WriteLine($"[手動刷新] 美股：開始更新 {tickers.Count} 個股票");

                                        for (int i = 0; i < tickers.Count; i++)
                                        {
                                                var ticker = tickers[i];
                                                Console.WriteLine($"[手動刷新] 美股 [{i+1}/{tickers.Count}] 更新 {ticker}");

                                                // 使用新方法：獲取價格和前收盤價
                                                _usPriceFetcher.UpdatePriceWithPreviousClose(ticker);

                                                // 添加請求間隔，避免頻率過高
                                                if (i < tickers.Count - 1) // 最後一個不需要等待
                                                {
                                                        System.Threading.Thread.Sleep(600);
                                                }
                                        }

                                        Console.WriteLine($"[手動刷新] 美股：完成更新 {tickers.Count} 個股票");
                                }
                                else
                                {
                                        var tickers = _twStockManager.GetTickers();
                                        Console.WriteLine($"[手動刷新] 台股：開始更新 {tickers.Count} 個股票");

                                        for (int i = 0; i < tickers.Count; i++)
                                        {
                                                var ticker = tickers[i];
                                                Console.WriteLine($"[手動刷新] 台股 [{i+1}/{tickers.Count}] 更新 {ticker}");

                                                // 使用新方法：獲取價格和前收盤價
                                                _twPriceFetcher.UpdatePriceWithPreviousClose(ticker);

                                                // 添加請求間隔，避免頻率過高
                                                if (i < tickers.Count - 1)
                                                {
                                                        System.Threading.Thread.Sleep(600);
                                                }
                                        }

                                        Console.WriteLine($"[手動刷新] 台股：完成更新 {tickers.Count} 個股票");
                                }

                                Dispatcher.Invoke(() =>
                                {
                                        UpdatePriceDisplay();
                                        statusText.Text = "✅ 刷新完成";
                                        Console.WriteLine($"[手動刷新] 刷新完成，下次可刷新時間: {DateTime.Now.AddSeconds(MANUAL_REFRESH_COOLDOWN):HH:mm:ss}");
                                });
                        });
                }

                private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
                {
                        var textBox = sender as TextBox;
                        var market = textBox?.Tag?.ToString();
                        ApplyFilter(market);
                }

                private void ApplyFilter(string market)
                {
                        if (market == "US")
                        {
                                var searchText = txtUsSearch?.Text?.ToLower() ?? "";
                                _usFilteredStockList.Clear();

                                foreach (var stock in _usStockList)
                                {
                                        if (string.IsNullOrEmpty(searchText) ||
                                                stock.Ticker.ToLower().Contains(searchText) ||
                                                stock.Name.ToLower().Contains(searchText))
                                        {
                                                _usFilteredStockList.Add(stock);
                                        }
                                }
                        }
                        else if (market == "TW")
                        {
                                var searchText = txtTwSearch?.Text?.ToLower() ?? "";
                                _twFilteredStockList.Clear();

                                foreach (var stock in _twStockList)
                                {
                                        if (string.IsNullOrEmpty(searchText) ||
                                                stock.Ticker.ToLower().Contains(searchText) ||
                                                stock.Name.ToLower().Contains(searchText))
                                        {
                                                _twFilteredStockList.Add(stock);
                                        }
                                }
                        }
                }

                private void MarketTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                        _countdownSeconds = 0;
                        _lastNewsUpdate = DateTime.MinValue;
                        _ = UpdateNewsImpactAsync();
                }

                private void BtnTestNVDA_Click(object sender, RoutedEventArgs e)
                {
                        Console.WriteLine("\n");
                        TestNVDA.Run();
                        Console.WriteLine("\n");

                        if (_debugWindow != null && _debugWindow.IsVisible)
                        {
                                MessageBox.Show("NVDA 診斷測試已完成！\n請查看調試視窗中的詳細輸出。", "測試完成", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                        {
                                MessageBox.Show("NVDA 診斷測試已完成！\n請開啟調試視窗（點擊 🐛 調試）查看結果。", "測試完成", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                }

                private void BtnTestAPI_Click(object sender, RoutedEventArgs e)
                {
                        Console.WriteLine("\n");

                        // 測試單個股票（NVDA）
                        TestYahooFinance.TestConnection("NVDA");

                        Console.WriteLine("\n");

                        // 測試多個股票
                        TestYahooFinance.TestMultipleStocks();

                        Console.WriteLine("\n");

                        // 🎯 新增：診斷 previousClose 為什麼是 null
                        Console.WriteLine("\n");
                        DiagnosePreviousClose.Run("NVDA");
                        Console.WriteLine("\n");

                        // 🐍 新增：測試 Python yfinance
                        Console.WriteLine("\n");
                        TestPythonYFinance.Run();
                        Console.WriteLine("\n");

                        if (_debugWindow != null && _debugWindow.IsVisible)
                        {
                                MessageBox.Show("API 測試已完成！\n\n包括:\n- Yahoo Finance V8/V7\n- Python yfinance\n- previousClose 診斷\n\n請查看調試視窗中的詳細輸出。", "測試完成", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                        {
                                MessageBox.Show("API 測試已完成！\n請開啟調試視窗（點擊 🐛 調試）查看結果。", "測試完成", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                }

                private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
                {
                        try
                        {
                                var dataGrid = sender as DataGrid;
                                if (dataGrid?.SelectedItem is StockInfo selectedStock)
                                {
                                        Console.WriteLine($"[雙擊] 選中股票: {selectedStock.Ticker} - {selectedStock.Name}");

                                        // 確定使用哪個市場的 PriceFetcher
                                        var currentTab = marketTabControl.SelectedIndex;
                                        var priceFetcher = currentTab == 0 ? _usPriceFetcher : _twPriceFetcher;

                                        Console.WriteLine($"[雙擊] 使用 {(currentTab == 0 ? "美股" : "台股")} PriceFetcher");

                                        // 打開趨勢分析視窗
                                        try
                                        {
                                                var trendWindow = new TrendAnalysisWindow(
                                                        selectedStock.Ticker,
                                                        selectedStock.Name,
                                                        priceFetcher
                                                );
                                                trendWindow.Owner = this;
                                                Console.WriteLine($"[雙擊] 準備顯示趨勢視窗");
                                                trendWindow.ShowDialog();
                                        }
                                        catch (Exception ex)
                                        {
                                                Console.WriteLine($"[雙擊錯誤] 創建或顯示趨勢視窗時發生錯誤: {ex.Message}");
                                                Console.WriteLine($"[雙擊錯誤] 堆疊追蹤: {ex.StackTrace}");

                                                MessageBox.Show(
                                                        $"無法開啟趨勢分析視窗\n\n錯誤信息：{ex.Message}\n\n請檢查調試視窗查看詳細信息",
                                                        "錯誤",
                                                        MessageBoxButton.OK,
                                                        MessageBoxImage.Error
                                                );
                                        }
                                }
                                else
                                {
                                        Console.WriteLine("[雙擊] 未選中任何股票或選中項目不是 StockInfo");
                                }
                        }
                        catch (Exception ex)
                        {
                                Console.WriteLine($"[雙擊外層錯誤] {ex.Message}");
                                Console.WriteLine($"[雙擊外層錯誤] 堆疊追蹤: {ex.StackTrace}");

                                MessageBox.Show(
                                        $"處理雙擊事件時發生錯誤\n\n{ex.Message}",
                                        "錯誤",
                                        MessageBoxButton.OK,
                                        MessageBoxImage.Error
                                );
                        }
                }

                private void BtnDebug_Click(object sender, RoutedEventArgs e)
                {
                        if (_debugWindow == null || !_debugWindow.IsLoaded)
                        {
                                _debugWindow = new DebugWindow();
                                _debugWindow.Show();
                        }
                        else
                        {
                                _debugWindow.Activate();
                        }
                }

                public void LogToDebugWindow(string message)
                {
                        try
                        {
                                // 確保在 UI 線程上執行
                                if (!Dispatcher.CheckAccess())
                                {
                                        // 如果不在 UI 線程，使用 Dispatcher.BeginInvoke
                                        Dispatcher.BeginInvoke(new Action(() => LogToDebugWindow(message)));
                                        return;
                                }

                                // 在 UI 線程上執行
                                if (_debugWindow != null && _debugWindow.IsLoaded)
                                {
                                        _debugWindow.AppendLog(message);
                                }
                        }
                        catch (Exception ex)
                        {
                                // 靜默處理錯誤，避免影響主程式
                                System.Diagnostics.Debug.WriteLine($"[LogToDebugWindow錯誤] {ex.Message}");
                        }
                }

                protected override void OnClosed(EventArgs e)
                {
                        SaveHoldings("US");
                        SaveHoldings("TW");
                        SaveRealizedPnL("US");
                        SaveRealizedPnL("TW");
                        _usMonitor?.StopThreads();
                        _twMonitor?.StopThreads();
                        _updateTimer?.Stop();
                        _countdownTimer?.Stop();
                        _debugWindow?.Close();
                        base.OnClosed(e);
                }
        }

        public class HoldingInfo : System.ComponentModel.INotifyPropertyChanged
        {
                private string _ticker;
                private string _name;
                private double _quantity;
                private double _averageCost;
                private double? _currentPrice;
                private DateTime? _updatedAt;

                public string Ticker
                {
                        get => _ticker;
                        set { _ticker = value; OnPropertyChanged(nameof(Ticker)); }
                }

                public string Name
                {
                        get => _name;
                        set { _name = value; OnPropertyChanged(nameof(Name)); }
                }

                public double Quantity
                {
                        get => _quantity;
                        set
                        {
                                _quantity = value;
                                NotifyCalculatedChanged();
                                OnPropertyChanged(nameof(Quantity));
                        }
                }

                public double AverageCost
                {
                        get => _averageCost;
                        set
                        {
                                _averageCost = value;
                                NotifyCalculatedChanged();
                                OnPropertyChanged(nameof(AverageCost));
                        }
                }

                public double? CurrentPrice
                {
                        get => _currentPrice;
                        set
                        {
                                _currentPrice = value;
                                NotifyCalculatedChanged();
                                OnPropertyChanged(nameof(CurrentPrice));
                        }
                }

                public DateTime? UpdatedAt
                {
                        get => _updatedAt;
                        set { _updatedAt = value; OnPropertyChanged(nameof(UpdatedAt)); }
                }

                public double CostAmount => Quantity * AverageCost;
                public double MarketValue => Quantity * (CurrentPrice ?? 0);
                public double UnrealizedPnL => MarketValue - CostAmount;
                public double PnLPercent => CostAmount > 0 ? (UnrealizedPnL / CostAmount) * 100 : 0;

                public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

                private void NotifyCalculatedChanged()
                {
                        OnPropertyChanged(nameof(CostAmount));
                        OnPropertyChanged(nameof(MarketValue));
                        OnPropertyChanged(nameof(UnrealizedPnL));
                        OnPropertyChanged(nameof(PnLPercent));
                }

                private void OnPropertyChanged(string propertyName)
                {
                        PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
                }
        }

        public class HoldingDto
        {
                public string Ticker { get; set; }
                public string Name { get; set; }
                public double Quantity { get; set; }
                public double AverageCost { get; set; }
                public double? CurrentPrice { get; set; }
                public DateTime? UpdatedAt { get; set; }
        }

        public class NewsImpactItem
        {
                public string Time { get; set; }
                public string Headline { get; set; }
                public string Source { get; set; }
                public int ImportanceScore { get; set; }
                public string ImportanceLevel { get; set; }
        }

        // 用於重定向 Console 輸出到調試視窗的 TextWriter
        public class DebugTextWriter : System.IO.TextWriter
        {
                private MainWindow _mainWindow;
                private System.Text.StringBuilder _buffer;
                private System.Threading.Timer _flushTimer;

                public DebugTextWriter(MainWindow mainWindow)
                {
                        _mainWindow = mainWindow;
                        _buffer = new System.Text.StringBuilder();

                        // 使用定時器批次處理輸出，減少 Dispatcher 調用
                        _flushTimer = new System.Threading.Timer(
                                _ => FlushBuffer(),
                                null,
                                100, // 首次延遲 100ms
                                100  // 之後每 100ms
                        );
                }

                public override void WriteLine(string value)
                {
                        try
                        {
                                if (!string.IsNullOrEmpty(value))
                                {
                                        lock (_buffer)
                                        {
                                                _buffer.AppendLine(value);
                                        }
                                }
                        }
                        catch
                        {
                                // 靜默處理錯誤
                        }
                }

                public override void Write(string value)
                {
                        // 暫存寫入，等待換行時一起輸出
                        try
                        {
                                if (!string.IsNullOrEmpty(value))
                                {
                                        lock (_buffer)
                                        {
                                                _buffer.Append(value);
                                        }
                                }
                        }
                        catch
                        {
                                // 靜默處理錯誤
                        }
                }

                private void FlushBuffer()
                {
                        try
                        {
                                string content = null;

                                lock (_buffer)
                                {
                                        if (_buffer.Length > 0)
                                        {
                                                content = _buffer.ToString();
                                                _buffer.Clear();
                                        }
                                }

                                if (!string.IsNullOrEmpty(content))
                                {
                                        _mainWindow?.LogToDebugWindow(content.TrimEnd('\r', '\n'));
                                }
                        }
                        catch
                        {
                                // 靜默處理錯誤
                        }
                }

                public override System.Text.Encoding Encoding
                {
                        get { return System.Text.Encoding.UTF8; }
                }

                protected override void Dispose(bool disposing)
                {
                        if (disposing)
                        {
                                _flushTimer?.Dispose();
                                FlushBuffer(); // 最後刷新一次
                        }
                        base.Dispose(disposing);
                }
        }
}
