using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace nrot.T590.Gui
{
    [ValueConversion(typeof(DateTime), typeof(string))]
    public class DateConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                return string.Empty;
            }

            var date = (DateTime)value;
            return date.ToShortDateString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var strValue = value as string;
            return DateTime.TryParse(strValue, out DateTime resultDateTime) ? resultDateTime : DependencyProperty.UnsetValue;
        }
    }
}
