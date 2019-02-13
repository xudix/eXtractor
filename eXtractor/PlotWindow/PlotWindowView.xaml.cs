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
        PlotWindowViewModel viewModel;

        public PlotWindowView(PlotWindowViewModel viewModel)
        {
            this.viewModel = viewModel;
            DataContext = this.viewModel;
            InitializeComponent();
            settingButton.IsChecked = false;
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

        /// <summary>
        /// Move the cursor and change the zoom box when the mouse is moved
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Chart_MouseMove(object sender, MouseEventArgs e)
        {
            // Simply moving the mouse in the chart. Will move the cursor
            Point currentPoint = e.GetPosition(Chart);
            Cursor1X = currentPoint.X;
            if (Cursor1X > Chart.ActualWidth - Chart.ChartLegend.ActualWidth)
                Cursor1X = Chart.ActualWidth - Chart.ChartLegend.ActualWidth;
            viewModel.UpdateLegendValues(Chart.ConvertToChartValues(currentPoint).X);

            // If the mouse is moving in the chart, check if it's zooming. If zooming, will draw zoom box; otherwise, will check if we should enter zooming mode
            // IsZoomDrawing means the mouse left button was pressed. Performing zoom
            if (IsZoomDrawing)
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
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
            else // not zooming
            {
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    // If IsZoomDrawing is not set but mouse left button is pressed, then the user probably pressed the button outside the chart and moved it inside.
                    // Will start zooming
                    // Previousw MouseLeftButtonDown event handler is moved here
                    // The reason is, if the user click the mouse left button outside the chart and move the mouse inside,
                    // the program won't be able to catch it. 
                    mouseDownPoint = currentPoint;
                    IsZoomDrawing = true;
                }
            }
        }

        /// <summary>
        /// Perform zooming when mouse button is released
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                // keep the new range within the previous range
                //xmin = (xmin < XAxis.ActualMinValue ? XAxis.ActualMinValue : xmin);
                //xmax = (xmax > XAxis.ActualMaxValue ? XAxis.ActualMaxValue : xmax);
                // if the new range is not inside the previous range, will do nothing
                if (xmin >= XAxis.ActualMaxValue || xmax <= XAxis.ActualMinValue) return;
                // Zoom in Y axis. If the change is too small, the user is probably not intended to zoom
                if ((ymax - ymin) / (YAxis.ActualMaxValue - YAxis.ActualMinValue) > 0.02)
                {
                    viewModel.YMax = ymax;
                    viewModel.YMin = ymin;
                }
                // Zoom in X axis. If the change is too small, the user is probably not intended to zoom
                if ((xmax - xmin) / (XAxis.ActualMaxValue - XAxis.ActualMinValue) > 0.02)
                {
                    viewModel.SetXRangeAndRecord(new DateTime((long)(xmin * TimeSpan.FromDays(1).Ticks)), new DateTime((long)(xmax * TimeSpan.FromDays(1).Ticks)));
                }
                
            }
            //else
            //{
            //    // Not zooming. simply clicking
            //    //Will update the values shown in the legend


            //}
        }

        /// <summary>
        /// When mouse left button is pressed down, start drawing the zoom box and record the mouse down point
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Chart_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            mouseDownPoint = e.GetPosition(Chart);
            IsZoomDrawing = true;
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
