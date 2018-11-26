using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using LiveCharts.Wpf;

namespace eXtractor.PlotWindow
{
    /// <summary>
    /// This is a variation of the DefaultLegend Class of LiveCharts Library
    /// It provides the capability to display values under each legend item
    /// To display the values, bind the Values property to an IList<float> object e.g. an array.
    /// Note: You need to make a new IList<> object everytime you want to update the display
    /// Otherwise, the values will not be updated in the UI.
    /// </summary>
    public partial class LegendWithValues : UserControl, IChartLegend
    {

        private List<SeriesViewModel> _series;

        private List<SeriesWithValueViewModel> _seriesWithValue;

        /// <summary>
        /// Initializes a new instance of DefaultLegend class
        /// </summary>
        public LegendWithValues()
        {
            InitializeComponent();
            SeriesWithValue = SeriesWithValueViewModel.CreateList(Series, Values);
            DataContext = this;
        }

        /// <summary>
        /// Property changed event
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Gets the series displayed in the legend.
        /// </summary>
        /// Chart of LiveCharts library will put inject the Series here. 
        public List<SeriesViewModel> Series
        {
            get { return _series; }
            set
            {
                _series = value;
                SeriesWithValue = SeriesWithValueViewModel.CreateList(Series, Values);
                OnPropertyChanged("Series");
            }
        }

        /// <summary>
        /// Gets the series With Values displayed in the legend.
        /// </summary>
        public List<SeriesWithValueViewModel> SeriesWithValue
        {
            get { return _seriesWithValue; }
            set
            {
                _seriesWithValue = value;
                OnPropertyChanged("SeriesWithValue");
            }
        }

        /// <summary>
        /// The Values Property
        /// </summary>
        public static readonly DependencyProperty ValuesProperty = DependencyProperty.Register(
            "Values", typeof(IList<float>), typeof(LegendWithValues), new FrameworkPropertyMetadata(null, new PropertyChangedCallback(OnValuesPropertyChange)));

        private static void OnValuesPropertyChange(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var source = d as LegendWithValues;
            source.SeriesWithValue = SeriesWithValueViewModel.CreateList(source.Series, (IList<float>)e.NewValue);
        }

        /// <summary>
        /// Gets or sets the Values of the legend.
        /// </summary>
        public IList<float> Values
        {
            get { return (IList<float>)GetValue(ValuesProperty); }
            set => SetValue(ValuesProperty, value);
        }
        /// <summary>
        /// The XValue property. XValue is shown on top of the Legend
        /// </summary>
        public static readonly DependencyProperty XValueProperty = DependencyProperty.Register(
            "XValue", typeof(String), typeof(LegendWithValues), new FrameworkPropertyMetadata("2018/01/01 00:00:00", FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsParentMeasure | FrameworkPropertyMetadataOptions.AffectsRender));
        /// <summary>
        /// Gets or sets the text to be displayed on top of the legend
        /// </summary>
        public string XValue
        {
            get { return (String)GetValue(XValueProperty); }
            set { SetValue(XValueProperty, value); }
        }


        /// <summary>
        /// The orientation property
        /// </summary>
        public static readonly DependencyProperty OrientationProperty = DependencyProperty.Register(
            "Orientation", typeof(Orientation?), typeof(LegendWithValues), new PropertyMetadata(null));
        /// <summary>
        /// Gets or sets the orientation of the legend, default is null, if null LiveCharts will decide which orientation to use, based on the Chart.Legend location property.
        /// </summary>
        public Orientation? Orientation
        {
            get { return (Orientation)GetValue(OrientationProperty); }
            set { SetValue(OrientationProperty, value); }
        }

        /// <summary>
        /// The internal orientation property
        /// </summary>
        public static readonly DependencyProperty InternalOrientationProperty = DependencyProperty.Register(
            "InternalOrientation", typeof(Orientation), typeof(LegendWithValues),
            new PropertyMetadata(default(Orientation)));

        /// <summary>
        /// Gets or sets the internal orientation.
        /// </summary>
        /// <value>
        /// The internal orientation.
        /// </value>
        public Orientation InternalOrientation
        {
            get { return (Orientation)GetValue(InternalOrientationProperty); }
            set { SetValue(InternalOrientationProperty, value); }
        }

        /// <summary>
        /// The bullet size property
        /// </summary>
        public static readonly DependencyProperty BulletSizeProperty = DependencyProperty.Register(
            "BulletSize", typeof(double), typeof(LegendWithValues), new PropertyMetadata(15d));
        /// <summary>
        /// Gets or sets the bullet size, the bullet size modifies the drawn shape size.
        /// </summary>
        public double BulletSize
        {
            get { return (double)GetValue(BulletSizeProperty); }
            set { SetValue(BulletSizeProperty, value); }
        }

        /// <summary>
        /// Called when [property changed].
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            if (PropertyChanged != null) PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
