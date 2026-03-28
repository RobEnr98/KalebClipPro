using System;
using System.Globalization;
using System.Windows.Data;

namespace KalebClipPro
{
    public class StringContainsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string content && parameter is string searchTerm)
            {
                return content.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}