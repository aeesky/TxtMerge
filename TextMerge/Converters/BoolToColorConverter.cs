using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace TextMerge.Converters
{
    public class BoolToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            Color color = Colors.DarkOrange;
            if ((bool) value)
            {
                return Colors.Green;
            }
            return new SolidColorBrush(color);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
