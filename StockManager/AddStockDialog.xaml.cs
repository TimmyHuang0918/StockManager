using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using StockManager.Config;

namespace StockManager
{
        public partial class AddStockDialog : Window
        {
                private static Dictionary<string, string> _twListedStocksCache;

                public string Ticker { get; private set; }
                public string StockName { get; private set; }
                public string Market { get; private set; }

                public AddStockDialog()
                {
                        InitializeComponent();
                        LoadSuggestionList();
                }

                private async void CmbMarket_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                        if (cmbMarket == null || txtTicker == null || lstSuggestions == null)
                        {
                                return;
                        }

                        await EnsureTwListedStocksLoadedAsync();
                        LoadSuggestionList();
                }

                private void TxtTicker_TextChanged(object sender, TextChangedEventArgs e)
                {
                        if (txtTicker == null || lstSuggestions == null)
                        {
                                return;
                        }

                        LoadSuggestionList();
                }

                private void LstSuggestions_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {
                        var selected = lstSuggestions.SelectedItem as SuggestionItem;
                        if (selected == null)
                        {
                                return;
                        }

                        txtTicker.Text = selected.Ticker;
                        txtName.Text = selected.Name;
                }

                private void BtnOk_Click(object sender, RoutedEventArgs e)
                {
                        var ticker = txtTicker.Text?.Trim();
                        var name = txtName.Text?.Trim();

                        if (string.IsNullOrEmpty(ticker))
                        {
                                MessageBox.Show("請輸入股票代號", "錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                                txtTicker.Focus();
                                return;
                        }

                        if (string.IsNullOrEmpty(name))
                        {
                                MessageBox.Show("請輸入公司名稱", "錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                                txtName.Focus();
                                return;
                        }

                        var selectedItem = cmbMarket.SelectedItem as System.Windows.Controls.ComboBoxItem;
                        Market = selectedItem?.Tag?.ToString() ?? "US";

                        if (Market == "TW" && !ticker.EndsWith(".TW", StringComparison.OrdinalIgnoreCase))
                        {
                                ticker += ".TW";
                        }

                        Ticker = ticker;
                        StockName = name;

                        DialogResult = true;
                        Close();
                }

                private void BtnCancel_Click(object sender, RoutedEventArgs e)
                {
                        DialogResult = false;
                        Close();
                }

                private async System.Threading.Tasks.Task EnsureTwListedStocksLoadedAsync()
                {
                        var selectedItem = cmbMarket.SelectedItem as ComboBoxItem;
                        var market = selectedItem?.Tag?.ToString() ?? "US";
                        if (market != "TW" || _twListedStocksCache != null)
                        {
                                return;
                        }

                        await System.Threading.Tasks.Task.Run(() =>
                        {
                                try
                                {
                                        var url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL";
                                        string json;
                                        using (var client = new WebClient())
                                        {
                                                client.Encoding = System.Text.Encoding.UTF8;
                                                json = client.DownloadString(url);
                                        }

                                        var serializer = new JavaScriptSerializer();
                                        var rows = serializer.Deserialize<List<Dictionary<string, string>>>(json) ?? new List<Dictionary<string, string>>();

                                        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                        foreach (var row in rows)
                                        {
                                                string code;
                                                string name;
                                                if (!row.TryGetValue("Code", out code) || !row.TryGetValue("Name", out name))
                                                {
                                                        continue;
                                                }

                                                dict[code + ".TW"] = name;
                                        }

                                        _twListedStocksCache = dict;
                                }
                                catch
                                {
                                        _twListedStocksCache = new Dictionary<string, string>();
                                }
                        });
                }

                private void LoadSuggestionList()
                {
                        if (cmbMarket == null || txtTicker == null || lstSuggestions == null)
                        {
                                return;
                        }

                        var selectedItem = cmbMarket.SelectedItem as ComboBoxItem;
                        var market = selectedItem?.Tag?.ToString() ?? "US";
                        var keyword = (txtTicker.Text ?? string.Empty).Trim().ToLowerInvariant();

                        Dictionary<string, string> source;
                        if (market == "TW")
                        {
                                source = _twListedStocksCache ?? AppConfig.DefaultTwStocks;
                        }
                        else
                        {
                                source = AppConfig.DefaultStocks;
                        }

                        var suggestions = source
                                .Where(x => string.IsNullOrEmpty(keyword)
                                            || x.Key.ToLowerInvariant().Contains(keyword)
                                            || x.Value.ToLowerInvariant().Contains(keyword))
                                .Take(60)
                                .Select(x => new SuggestionItem
                                {
                                        Ticker = x.Key,
                                        Name = x.Value
                                })
                                .ToList();

                        lstSuggestions.ItemsSource = suggestions;
                        lstSuggestions.DisplayMemberPath = "DisplayName";
                }

                private class SuggestionItem
                {
                        public string Ticker { get; set; }
                        public string Name { get; set; }
                        public string DisplayName => $"{Ticker} - {Name}";
                }
        }
}
