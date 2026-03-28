using System;
using System.Globalization;
using System.Windows.Data;

namespace KalebClipPro.Infrastructure
{
    public class ClipCategoryConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not string texto || string.IsNullOrWhiteSpace(texto)) return "Text";

            texto = texto.Trim();

            if (texto.Contains("xmlns=\"http") || texto.StartsWith("<") || texto.Contains("public class") || 
                (texto.Contains("{") && texto.Contains("}") && texto.Contains(";"))) 
                return "Code";

            if ((texto.StartsWith("http://") || texto.StartsWith("https://")) && !texto.Contains(" "))
                return "Url";

            return "Text";
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }
}