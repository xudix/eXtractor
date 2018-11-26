using LiveCharts;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace eXtractor.PlotWindow
{
    /// <summary>
    /// Interaction logic for PlotWindowView.xaml
    /// </summary>
    public partial class PlotWindowView : Window, INotifyPropertyChanged
    {
        private PlotWindowViewModel viewModel;

        public PlotWindowView(PlotWindowViewModel viewModel)
        {
            this.viewModel = viewModel;
            InitializeComponent();
            DataContext = this.viewModel;
        }

        #region Properties and fields related to zoom behavior

        /// <summary>
        /// These fields are used to support the draw-to-zoom function
        /// Click the mouse on the chart and draw a rectangular area. 
        /// The chart will zoom in to the rectangular area.
        /// </summary>
        private Point mouseDownPoint, mouseUpPoint;

        private bool isZoomDrawing;
        /// <summary>
        /// Indicate whether the plot is in zooming mode.
        /// </summary>
        /// Currently, the plot goes into zooming  mode when mouse left button is pressed down in the plot.
        /// A rectangle zoom box will be drawn as the mouse moves with left button down
        public bool IsZoomDrawing
        {
            get => isZoomDrawing;
            set
            {
                if (isZoomDrawing != value)
                {
                    isZoomDrawing = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private double zoomBoxWidth;
        /// <summary>
        /// The width of the zoom box
        /// </summary>
        public double ZoomBoxWidth
        {
            get => zoomBoxWidth;
            set
            {
                if (zoomBoxWidth != value)
                {
                    zoomBoxWidth = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private double zoomBoxHeight;
        /// <summary>
        /// The height of the zoom box
        /// </summary>
        public double ZoomBoxHeight
        {
            get => zoomBoxHeight;
            set
            {
                if (zoomBoxHeight != value)
                {
                    zoomBoxHeight = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private Thickness zoomBoxMargin;
        /// <summary>
        /// The location of the zoom box in the grid
        /// </summary>
        public Thickness ZoomBoxMargin
        {
            get => zoomBoxMargin;
            set
            {
                if (zoomBoxMargin != value)
                {
                    zoomBoxMargin = value;
                    NotifyPropertyChanged();
                }
            }
        }

        #endregion

        #region Properties and fiels related to the cursor


        /// <summary>
        ///  Property for the location of CursorLine1
        /// </summary>
        private double cursor1X;
        public double Cursor1X
        {
            get => cursor1X;
            set
            {
                if (cursor1X != value)
                {
                    cursor1X = value;
                    NotifyPropertyChanged();
                }
            }
        }



        #endregion

        #region UI event handlers

        /// <summary>
        /// Reset the plot range when user doubleclick the chart
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Chart_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            viewModel.MoveZoom(0);
            e.Handled = true;
        }

        private void Chart_MouseMove(object sender, MouseEventArgs e)
        {
            // Simply moving the mouse in the chart. Will move the cursor
            Cursor1X = e.GetPosition(Chart).X;
            if (Cursor1X > Chart.ActualWidth - Chart.ChartLegend.ActualWidth)
                Cursor1X = Chart.ActualWidth - Chart.ChartLegend.ActualWidth;
            //Cursor1Margin = new Thickness(position, 0, Chart.ActualWidth - position - 1, XAxis.ActualHeight);
            UpdateLegendValues(Chart.ConvertToChartValues(new Point(Cursor1X, 0)).X);
            //RawPoint = e.GetPosition(Chart);
            //ConvertedPoint = Chart.ConvertToChartValues(RawPoint);

            // If the mouse is moving in the chart, allow the following behavior:
            // IsZoomDrawing means the mouse left button was pressed. Performing zoom
            if (IsZoomDrawing)
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    Point currentPoint = e.GetPosition(Chart);
                    double xmin = mouseDownPoint.X;
                    double xmax = currentPoint.X;
                    double ymin = currentPoint.Y;
                    double ymax = mouseDownPoint.Y;
                    // make sure min is smaller than max
                    if (xmin > xmax)
                    {
                        double temp = xmin;
                        xmin = xmax;
                        xmax = temp;
                    }
                    // Limit the zoombox in the chart

                    if (xmax > Chart.ActualWidth - Chart.ChartLegend.ActualWidth)
                        xmax = Chart.ActualWidth - Chart.ChartLegend.ActualWidth;
                    if (ymin > ymax)
                    {
                        double temp = ymin;
                        ymin = ymax;
                        ymax = temp;
                    }
                    // Limit the zoombox in the chart
                    if (ymin < 0) ymin = 0;
                    if (ymax > Chart.ActualHeight) ymax = Chart.ActualHeight;
                    if (xmax >= xmin && ymax >= ymin)
                    {
                        ZoomBoxWidth = xmax - xmin;
                        ZoomBoxHeight = ymax - ymin;
                        ZoomBoxMargin = new Thickness(xmin, ymin, Chart.ActualWidth - xmax, Chart.ActualHeight - ymax);
                    }
                }
                else // If the user moved the mouse to outside the Grid, release the button, and move it back into the chart, zooming will be canceled
                {
                    IsZoomDrawing = false;
                    ZoomBoxHeight = 0;
                    ZoomBoxWidth = 0;
                }
            }
            else
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    // If IsZoomDrawing is not set but mouse left button is pressed, then the user probably pressed the button outside the chart and moved it inside.
                    // Will start zooming
                    // Previousw MouseLeftButtonDown event handler is moved here
                    // The reason is, if the user click the mouse left button outside the chart and move the mouse inside,
                    // the program won't be able to catch it. 
                    mouseDownPoint = e.GetPosition(Chart);
                    IsZoomDrawing = true;
                }
                else
                {

                }

            }
        }

        private void Chart_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (IsZoomDrawing)
            {
                mouseUpPoint = e.GetPosition(Chart);
                IsZoomDrawing = false;
                ZoomBoxHeight = 0;
                ZoomBoxWidth = 0;
                // If the up and down points are very close to each other, it's not zooming
                if ((mouseDownPoint.X - mouseUpPoint.X) * (mouseDownPoint.X - mouseUpPoint.X) + (mouseDownPoint.Y - mouseUpPoint.Y) * (mouseDownPoint.Y - mouseUpPoint.Y) < 100)
                    return;

                mouseDownPoint = Chart.ConvertToChartValues(mouseDownPoint);
                mouseUpPoint = Chart.ConvertToChartValues(mouseUpPoint);
                // Find the coordinates for the zoom rectangular
                double xmin = mouseDownPoint.X;
                double xmax = mouseUpPoint.X;
                double ymin = mouseUpPoint.Y;
                double ymax = mouseDownPoint.Y;
                // make sure min is smaller than max
                if (xmin > xmax)
                {
                    double temp = xmin;
                    xmin = xmax;
                    xmax = temp;
                }
                if (ymin > ymax)
                {
                    double temp = ymin;
                    ymin = ymax;
                    ymax = temp;
                }
                // If the start point is outside the range. No need to continue.
                int startIndex = (int)xmin;
                if (startIndex < 0) startIndex = 0;
                else if (startIndex >= PointsPerLine) return;
                // If the change is too small, the user is probably not intended to zoom
                if ((ymax - ymin) / (YAxis.ActualMaxValue - YAxis.ActualMinValue) > 0.02)
                {
                    YMax = ymax;
                    YMin = ymin;
                }
                // If the change is too small, the user is probably not intended to zoom
                if ((xmax - xmin) / PointsPerLine > 0.02)
                {

                    int endIndex = (int)Math.Ceiling(xmax);
                    if (endIndex >= PointsPerLine) endIndex = PointsPerLine - 1;
                    // Here we are changing the private field "startDateTime" instead of the property "StartDateTime"
                    // because if we chagne the property, UpdatePoints() method will be invoked and DateTimeStrs will be changed.
                    startDateTime = ExtractedData.ParseDateTime(DateTimeStrs[startIndex]);
                    NotifyPropertyChanged("StartDateTime");
                    // Here we invoke NotifyPropertyChanged and UpdatePoints manually because
                    // when the zoom operation is ended (mouse up) outside the chart, EndDatTime will be the same. 
                    // In this case, the UpdatePoints method will not be invoked. 
                    // However, we may need to invoke it since StartDateTime may have changed.
                    endDateTime = ExtractedData.ParseDateTime(DateTimeStrs[endIndex]);
                    NotifyPropertyChanged("EndDateTime");
                    UpdatePoints();
                    PlotRangeChanged(this, new PlotRangeChangedEventArgs(startDateTime, endDateTime, this));
                }
                RecordRange();
            }
            //else
            //{
            //    // Not zooming. simply clicking
            //    //Will update the values shown in the legend


            //}
        }




        #endregion


        #region INotifyPropertyChanged Implementation

        /// <summary>
        /// This event notifies UI to update content after the back-end data is changed by program
        /// It is required by the INotifyPropertyChanged interface.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// This method is called by the Set accessor of each property to invoke PropertyChanged event.
        /// </summary>
        /// <param name="propertyName"></param>
        /// The CallerMemberName attribute that is applied to the optional propertyName parameter causes the property name of the caller to be substituted as an argument.
        private void NotifyPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] String propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        #endregion



    }
}
