using SKCOMLib;
using StockManager.Services;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace StockManager
{
    public partial class MainWindow
    {
        private volatile bool _isSkQuoteConnectionReady;
        private bool _hasInitialSkStocksRequested;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string loginId;
            string password;
            bool rememberCredential;
            if (!TryShowCapitalLoginDialog(out loginId, out password, out rememberCredential))
            {
                return;
            }

            RegisterSkEventsIfNeeded();

            var resultCode = m_api.SKCenterLib_Login(loginId, password);
            Console.WriteLine($"登入結果: {resultCode}");
            if (resultCode != 0)
            {
                MessageBox.Show($"登入失敗，代碼: {resultCode}", "群益登入", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var loginButton = sender as Button ?? FindName("btnTwCapitalLogin") as Button;
            if (loginButton != null)
            {
                loginButton.Visibility = Visibility.Collapsed;
            }

            var loginInfoText = FindName("txtTwCapitalLoginInfo") as TextBlock;
            if (loginInfoText != null)
            {
                loginInfoText.Text = $"群益帳號: {loginId}";
                loginInfoText.Visibility = Visibility.Visible;
            }

            statusText.Text = $"群益登入成功：{loginId}";
            if (rememberCredential)
            {
                SaveCapitalCredential(loginId, password);
            }
            else
            {
                ClearCapitalCredential();
            }

            _isCapitalLoggedIn = true;
            _isSkQuoteConnectionReady = false;
            _hasInitialSkStocksRequested = false;

	    // 與報價伺服器連線
	    int nCode = m_api.SKQuoteLib_EnterMonitorLONG();

	    //事件回傳查詢主機時間的結果
	    {
		m_api.OnNotifyServerTime += OnNotifyServerTime;
		void OnNotifyServerTime(short sHour, short sMinute, short sSecond, int nTotal)
		{
		    
		}
	    }
	    // 事件回傳證券市場－整零價差即時行情
	    {
		m_api.OnNotifyOddLotSpreadDeal += OnNotifyOddLotSpreadDeal;
		void OnNotifyOddLotSpreadDeal(short sMarketNo, string bstrStockNo, int nDealPrice, short sDigit)
		{
		   
		}
	    }
	    
	    //事件最佳五檔(國內報價)
	    {
		m_api.OnNotifyBest5LONG += OnNotifyBest5LONG;
		void OnNotifyBest5LONG(short sMarketNo, int nStockidx, int nBestBid1, int nBestBidQty1, int nBestBid2, int nBestBidQty2, int nBestBid3, int nBestBidQty3, int nBestBid4, int nBestBidQty4, int nBestBid5, int nBestBidQty5, int nExtendBid, int nExtendBidQty, int nBestAsk1, int nBestAskQty1, int nBestAsk2, int nBestAskQty2, int nBestAsk3, int nBestAskQty3, int nBestAsk4, int nBestAskQty4, int nBestAsk5, int nBestAskQty5, int nExtendAsk, int nExtendAskQty, int nSimulate)
		{
		}
	    }
	    //(LONG index)當有索取的個股成交明細有所異動，即透過向此註冊事件回傳所異動的個股成交明細
	    {
		m_api.OnNotifyTicksLONG += OnNotifyTicksLONG;
		void OnNotifyTicksLONG(short sMarketNo, int nIndex, int nPtr, int nDate, int nTimehms, int nTimemillismicros, int nBid, int nAsk, int nClose, int nQty, int nSimulate)
		{
		   
		}
	    }
	    //(LONG index)當首次索取個股成交明細，此事件會回補當天Tick 【沒用到喔】
	    {
		//m_pSKQuote.OnNotifyHistoryTicksLONG += new _ISKQuoteLibEvents_OnNotifyHistoryTicksLONGEventHandler(OnNotifyHistoryTicksLONG);
		void OnNotifyHistoryTicksLONG(short sMarketNo, int nIndex, int nPtr, int nDate, int nTimehms, int nTimemillismicros, int nBid, int nAsk, int nClose, int nQty, int nSimulate)
		{
		    
		}
	    }
	    //20241120 最後一筆歷史KLine
	    m_api.OnKLineComplete += OnKLineComplete;
	    void OnKLineComplete(string bstrEndString)
	    {
		
	    }
	    //事件報價通知
	    {
		m_api.OnNotifyQuoteLONG += OnNotifyQuoteLONG;
		void OnNotifyQuoteLONG(short sMarketNo, int nIndex)
		{
		    
		}
	    }
	    //透過呼叫 SKQuoteLib_GetMarketBuySellUpDown 後，事件回傳大盤成交張筆資料
	    {
		m_api.OnNotifyMarketTot += OnNotifyMarketTot;
		void OnNotifyMarketTot(short sMarketNo, short sPtr, int nTime, int nTotv, int nTots, int nTotc)
		{
		    
		}
	    }
	    //透過呼叫 SKQuoteLib_GetMarketBuySellUpDown 後，事件回傳大盤成交買賣張筆數資料
	    {
		m_api.OnNotifyMarketBuySell += OnNotifyMarketBuySell;
		void OnNotifyMarketBuySell(short sMarketNo, short sPtr, int nTime, int nBc, int nSc, int nBs, int nSs)
		{

		}
	    }
	    //透過呼叫 SKQuoteLib_GetMarketBuySellUpDown 後，事件回傳大盤成交上漲下跌家數資料(包含『含權證家數』、 『不含權證家數』)
	    {
		m_api.OnNotifyMarketHighLowNoWarrant += OnNotifyMarketHighLowNoWarrant;
		void OnNotifyMarketHighLowNoWarrant(short sMarketNo, int nPtr, int nTime, int nUp, int nDown, int nHigh, int nLow, int nNoChange, int nUpNoW, int nDownNoW, int nHighNoW, int nLowNoW, int nNoChangeNoW)
		{
		    
		}
	    }
	    //(LONG index)事件回傳證券市場－技術分析平滑異同平均線MACD數值。（日線－完整）
	    {
		m_api.OnNotifyMACDLONG += OnNotifyMACDLONG;
		void OnNotifyMACDLONG(short sMarketNo, int nStockidx, string bstrMACD, string bstrDIF, string bstrOSC)
		{
		   
		}
	    }
	    //(LONG index)事件回傳技術分析－布林通道。（日線－完整）
	    {
		m_api.OnNotifyBoolTunelLONG += OnNotifyBoolTunelLONG;
		void OnNotifyBoolTunelLONG(short sMarketNo, int nStockidx, string bstrAVG, string bstrUBT, string bstrLBT)
		{
		    
		}
	    }
	    //事件回傳技術分析資訊
	    {
		m_api.OnNotifyKLineData += OnNotifyKLineData;
		void OnNotifyKLineData(string bstrStockNo, string bstrData)
		{
		    string[] values = new string[6];
		    values = bstrData.Split(',');
		}
	    }
	    //事件回傳指定國內市場－各類股商品清單
	    {
		m_api.OnNotifyCommodityListWithTypeNo += OnNotifyCommodityListWithTypeNo;
		void OnNotifyCommodityListWithTypeNo(short sMarketNo, string bstrStockData)
		{
		   
		}
	    }

	    string msg = "【SKQuoteLib_EnterMonitorLONG】" + m_api.SKCenterLib_GetReturnCodeMessage(nCode);
	    Console.WriteLine(msg);
        }

        private bool TryShowCapitalLoginDialog(out string loginId, out string password, out bool rememberCredential)
        {
            loginId = string.Empty;
            password = string.Empty;
            rememberCredential = false;
            var tempLoginId = string.Empty;
            var tempPassword = string.Empty;
            var tempRememberCredential = false;
            var savedCredential = LoadCapitalCredential();

            var dialog = new Window
            {
                Title = "群益登入",
                Width = 360,
                Height = 450,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Owner = this,
                Background = (Brush)new BrushConverter().ConvertFromString("#F5F7FA")
            };

            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var lblId = new TextBlock { Text = "帳號", FontWeight = FontWeights.Bold, Foreground = (Brush)new BrushConverter().ConvertFromString("#2C3E50") };
            Grid.SetRow(lblId, 0);
            root.Children.Add(lblId);

            var txtId = new TextBox { Height = 30, Margin = new Thickness(0, 6, 0, 0), VerticalContentAlignment = VerticalAlignment.Center };
            Grid.SetRow(txtId, 1);
            root.Children.Add(txtId);

            var lblPwd = new TextBlock { Text = "密碼", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 10, 0, 0), Foreground = (Brush)new BrushConverter().ConvertFromString("#2C3E50") };
            Grid.SetRow(lblPwd, 2);
            root.Children.Add(lblPwd);

            var txtPwd = new PasswordBox { Height = 30, Margin = new Thickness(0, 6, 0, 0) };
            Grid.SetRow(txtPwd, 3);
            root.Children.Add(txtPwd);

            var chkRemember = new CheckBox
            {
                Content = "記住帳號密碼",
                Margin = new Thickness(0, 10, 0, 0),
                IsChecked = savedCredential != null,
                Foreground = Brushes.Black
            };
            Grid.SetRow(chkRemember, 4);
            root.Children.Add(chkRemember);

            var hint = new TextBlock
            {
                Text = "請輸入群益帳號與密碼",
                FontSize = 11,
                Foreground = (Brush)new BrushConverter().ConvertFromString("#607D8B"),
                Margin = new Thickness(0, 10, 0, 0)
            };
            Grid.SetRow(hint, 5);
            root.Children.Add(hint);

            var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
            var btnOk = new Button { Content = "登入", Width = 80, Height = 30, Margin = new Thickness(0, 0, 8, 0), Background = (Brush)new BrushConverter().ConvertFromString("#27AE60"), Foreground = Brushes.White, BorderThickness = new Thickness(0), IsDefault = true };
            var btnCancel = new Button { Content = "取消", Width = 80, Height = 30, Background = (Brush)new BrushConverter().ConvertFromString("#95A5A6"), Foreground = Brushes.White, BorderThickness = new Thickness(0), IsCancel = true };

            btnOk.Click += (s, e) =>
            {
                var id = (txtId.Text ?? string.Empty).Trim();
                var pwd = txtPwd.Password ?? string.Empty;
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(pwd))
                {
                    MessageBox.Show("請輸入帳號與密碼", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                tempLoginId = id;
                tempPassword = pwd;
                tempRememberCredential = chkRemember.IsChecked == true;
                dialog.DialogResult = true;
            };

            btnCancel.Click += (s, e) => dialog.DialogResult = false;

            btnPanel.Children.Add(btnOk);
            btnPanel.Children.Add(btnCancel);
            Grid.SetRow(btnPanel, 6);
            root.Children.Add(btnPanel);

            dialog.Content = root;
            dialog.Loaded += (s, e) =>
            {
                if (savedCredential != null)
                {
                    txtId.Text = savedCredential.Item1;
                    txtPwd.Password = savedCredential.Item2;
                }
                txtId.Focus();
            };

            if (dialog.ShowDialog() == true)
            {
                loginId = tempLoginId;
                password = tempPassword;
                rememberCredential = tempRememberCredential;
                return true;
            }

            return false;
        }
    }
}
