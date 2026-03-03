using System;
using System.Windows;

namespace StockManager
{
        public partial class DebugWindow : Window
        {
                public DebugWindow()
                {
                        InitializeComponent();
                }

                public void AppendLog(string message)
                {
                        if (Dispatcher.CheckAccess())
                        {
                                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                                txtLog.ScrollToEnd();
                        }
                        else
                        {
                                Dispatcher.Invoke(() => AppendLog(message));
                        }
                }

                private void BtnClear_Click(object sender, RoutedEventArgs e)
                {
                        txtLog.Clear();
                }

                private void BtnClose_Click(object sender, RoutedEventArgs e)
                {
                        Close();
                }
        }
}
