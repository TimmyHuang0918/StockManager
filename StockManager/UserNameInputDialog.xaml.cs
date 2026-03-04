using System.Windows;

namespace StockManager
{
    public partial class UserNameInputDialog : Window
    {
        public string UserName => txtUserName.Text?.Trim();

        public UserNameInputDialog()
        {
            InitializeComponent();
            Loaded += (s, e) => txtUserName.Focus();
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(UserName))
            {
                MessageBox.Show("請輸入使用者名稱", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
