using System;
using System.Globalization;
using System.Windows.Data;

namespace StockManager.Converters
{
        public class PositiveNegativeConverter : IValueConverter
        {
                public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
                {
                        var change = value as double?;
                        if (change.HasValue)
                        {
                                if (change.Value > 0)
                                        return "Positive";
                                else if (change.Value < 0)
                                        return "Negative";
                                else
                                        return "Neutral";
                        }
                        return "Neutral";
                }

                public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
                {
                        throw new NotImplementedException();
                }
        }
}
