using System.Globalization;
using System.Windows;

namespace StockManager
{
    public partial class CashInputDialog : Window
    {
        public double CashAmount { get; private set; }

        public CashInputDialog(string title, double currentAmount)
        {
            InitializeComponent();
            Title = title;
            txtCashAmount.Text = currentAmount.ToString("F2", CultureInfo.InvariantCulture);
            Loaded += (s, e) =>
            {
                txtCashAmount.Focus();
                txtCashAmount.SelectAll();
            };
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            var text = (txtCashAmount.Text ?? string.Empty).Trim();
            double amount;
            if (!double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out amount)
                && !double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out amount))
            {
                MessageBox.Show("Please enter a valid amount", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (amount < 0)
            {
                MessageBox.Show("Amount must be >= 0", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            CashAmount = amount;
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
