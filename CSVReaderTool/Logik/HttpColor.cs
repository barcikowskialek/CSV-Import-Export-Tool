using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace CSVReaderTool.Logik
{
    public class HttpColor : IValueConverter
    {
        public object Convert(object value, Type t, object p, CultureInfo c)
        {
            var s = value?.ToString() ?? "";
            return s.StartsWith("http:", StringComparison.OrdinalIgnoreCase)
                ? Brushes.Orange
                : Brushes.Black;
        }

        public object ConvertBack(object value, Type t, object p, CultureInfo c) => value;
    }
}