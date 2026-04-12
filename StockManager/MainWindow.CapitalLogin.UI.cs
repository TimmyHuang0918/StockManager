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

            var nCode = m_api.SKQuoteLib_EnterMonitorLONG();
            var msg = "【SKQuoteLib_EnterMonitorLONG】" + m_api.SKCenterLib_GetReturnCodeMessage(nCode);
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
