using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;


namespace eXtractor.PlotWindow
{
    /// <summary>
    /// 
    /// </summary>
    /// <seealso cref="System.Windows.Data.IMultiValueConverter" />
    public class OrientationConverter : IMultiValueConverter
    {
        /// <summary>
        /// Converts source values to a value for the binding target. The data binding engine calls this method when it propagates the values from source bindings to the binding target.
        /// </summary>
        /// <param name="values">The array of values that the source bindings in the <see cref="T:System.Windows.Data.MultiBinding" /> produces. The value <see cref="F:System.Windows.DependencyProperty.UnsetValue" /> indicates that the source binding has no value to provide for conversion.</param>
        /// <param name="targetType">The type of the binding target property.</param>
        /// <param name="parameter">The converter parameter to use.</param>
        /// <param name="culture">The culture to use in the converter.</param>
        /// <returns>
        /// A converted value.If the method returns null, the valid null value is used.A return value of <see cref="T:System.Windows.DependencyProperty" />.<see cref="F:System.Windows.DependencyProperty.UnsetValue" /> indicates that the converter did not produce a value, and that the binding will use the <see cref="P:System.Windows.Data.BindingBase.FallbackValue" /> if it is available, or else will use the default value.A return value of <see cref="T:System.Windows.Data.Binding" />.<see cref="F:System.Windows.Data.Binding.DoNothing" /> indicates that the binding does not transfer the value or use the <see cref="P:System.Windows.Data.BindingBase.FallbackValue" /> or the default value.
        /// </returns>
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values[0] == DependencyProperty.UnsetValue) return null;

            return (Orientation?)values[0] ?? (Orientation)values[1];
        }

        /// <summary>
        /// Converts a binding target value to the source binding values.
        /// </summary>
        /// <param name="value">The value that the binding target produces.</param>
        /// <param name="targetTypes">The array of types to convert to. The array length indicates the number and types of values that are suggested for the method to return.</param>
        /// <param name="parameter">The converter parameter to use.</param>
        /// <param name="culture">The culture to use in the converter.</param>
        /// <returns>
        /// An array of values that have been converted from the target value back to the source values.
        /// </returns>
        /// <exception cref="System.NotImplementedException"></exception>
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// This class connects the values shown in the legend to corresponding properties
    /// The main purpose of this class is to handle Single.NaN.
    /// When the value is NaN, the textblock will show empty string
    /// </summary>
    [ValueConversion(typeof(float), typeof(string))]
    public class StrFloatConverter : IValueConverter
    {
        // Convert method is from Source to Target. Source is Double and target is string
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
            (!Single.IsNaN((float)value)) ? ((float)value).ToString("f") : "";

        // ConvertBack method is from Target to Source
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((string)value == "")
                return Single.NaN;
            try
            {
                return Single.Parse((string)value);
            }
            catch
            {
                return Single.NaN;
            }
        }
    }

    /// <summary>
    /// This class connects a input box with a DateTime object via the ExtactedData.ParseDateTime method
    /// </summary>
    [ValueConversion(typeof(DateTime), typeof(string))]
    public class StringDateTimeConverter : IValueConverter
    {
        // Convert method is from Source to Target. Source is DateTime and target is string
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
            (value != null) ? ((DateTime)value).ToString(@"yyyy/M/d H:mm:ss") : "";

        // ConvertBack method is from Target to Source
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            try
            {
                return ExtractedData.ParseDateTime((string)value);
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>
    /// This class connects the Y axis min and max input box to corresponding properties
    /// The main purpose of this class is to handle Double.NaN.
    /// When the min/max property is set to NaN, the input box will show empty string
    /// </summary>
    [ValueConversion(typeof(double), typeof(string))]
    public class StrDoubleConverter : IValueConverter
    {
        // Convert method is from Source to Target. Source is Double and target is string
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
            (!Double.IsNaN((double)value)) ? ((double)value).ToString("f") : "";

        // ConvertBack method is from Target to Source
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((string)value == "")
                return Double.NaN;
            try
            {
                return Double.Parse((string)value);
            }
            catch
            {
                return Double.NaN;
            }
        }
    }

    // Calculate the height of the cursor line based on the height of the Chart and height of the XAxis
    // The cursor height is the difference between the Chart.ActualHeight and XAxis.ActualHeight
    public class CursorHeightCalc : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            return (double)values[0] - (double)values[1];
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
