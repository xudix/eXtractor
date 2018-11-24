using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace eXtractor
{
    /// <summary>
    /// This class is used to connect the tag/file array with the input boxes. The text of input boxes are strings.
    /// </summary>
    [ValueConversion(typeof(string[]), typeof(string))]
    public class StringArrayConverter : IValueConverter
    {
        private char[] tagSeparators = { ' ', ',', '\t', '\n', '\r', ';', '|', '\"' };
        private char[] fileSeparators = { '\r', '\n', '\t', '\"', '|' };
        // Convert method is from Source to Target. Source is string[] and target is string
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
            (value != null) ? String.Join("\r\n", (string[])value) + "\r\n" : "";

        // ConvertBack method is from Target to Source
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            switch ((string)parameter)
            {
                case "tag":
                    return ((string)value).Split(tagSeparators, StringSplitOptions.RemoveEmptyEntries);
                case "file":
                    return ((string)value).Split(fileSeparators, StringSplitOptions.RemoveEmptyEntries);
                default:
                    return null;
            }

        }

    }

    /// <summary>
    /// This class connects a input box with a DateTime object via the ExtactedData.ParseDate method
    /// </summary>
    [ValueConversion(typeof(DateTime), typeof(string))]
    public class StringDateConverter : IValueConverter
    {
        // Convert method is from Source to Target. Source is DateTime and target is string
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
            (value != null) ? ((DateTime)value).ToString(@"yyyy/MM/dd") : "";

        // ConvertBack method is from Target to Source
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            try
            {
                return ExtractedData.ParseDate((string)value);
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// This class connects a input box with a TimeSpan object via the ExtactedData.ParseTime method
    /// </summary>
    [ValueConversion(typeof(TimeSpan), typeof(string))]
    public class StringTimeConverter : IValueConverter
    {
        // Convert method is from Source to Target. Source is DateTime and target is string
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
            (value != null) ? ((TimeSpan)value).ToString(@"h\:mm\:ss") : "";

        // ConvertBack method is from Target to Source
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            try
            {
                return ExtractedData.ParseTime((string)value);
            }
            catch
            {
                return null;
            }
        }
    }

}
