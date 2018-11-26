using LiveCharts.Wpf;
using System;
using System.Collections.Generic;

namespace eXtractor.PlotWindow
{
    public class SeriesWithValueViewModel
    {
        public SeriesWithValueViewModel(SeriesViewModel series, float value)
        {
            _series = series;
            Value = value;
        }


        /// <summary>
        /// Create a List of SeriesWithValueViewModel from a List of SeriesViewModel and a IList of values.
        /// </summary>
        /// <param name="series">The List of SeriesViewModel to be included in the SeriesWithValueViewModel</param>
        /// <param name="values">The IList of values to be included in the SeriesWithValueViewModel</param>
        /// <returns></returns>
        internal static List<SeriesWithValueViewModel> CreateList(List<SeriesViewModel> series, IList<float> values)
        {
            if (series == null)
                return null;
            var result = new List<SeriesWithValueViewModel>(series.Count);
            for (int i = 0; i < series.Count; i++)
            {
                if (values != null && values.Count > i)
                    result.Add(new SeriesWithValueViewModel(series[i], values[i]));
                else
                    result.Add(new SeriesWithValueViewModel(series[i], Single.NaN));
            }
            return result;
        }

        private SeriesViewModel _series;
        //
        // Summary:
        //     Series Title
        public string Title
        {
            get => _series.Title;
            set
            {
                _series.Title = value;
            }
        }
        //
        // Summary:
        //     Series stroke
        public System.Windows.Media.Brush Stroke
        {
            get => _series.Stroke;
            set
            {
                _series.Stroke = value;
            }
        }
        //
        // Summary:
        //     Series Stroke thickness
        public double StrokeThickness
        {
            get => _series.StrokeThickness;
            set
            {
                _series.StrokeThickness = value;
            }
        }
        //
        // Summary:
        //     Series Fill
        public System.Windows.Media.Brush Fill
        {
            get => _series.Fill;
            set
            {
                _series.Fill = value;
            }
        }
        //
        // Summary:
        //     Series point Geometry
        public System.Windows.Media.Geometry PointGeometry
        {
            get => _series.PointGeometry;
            set
            {
                _series.PointGeometry = value;
            }
        }
        //
        // Summary:
        //     Value to be shown
        public float Value { get; set; }
    }
}
