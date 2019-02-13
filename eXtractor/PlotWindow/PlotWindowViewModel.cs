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
    // The Model for PlotWindow will be the ExtractedData
    public class PlotWindowViewModel: INotifyPropertyChanged
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="extractionRequest"></param>
        /// <param name="parent"></param>
        public PlotWindowViewModel(ExtractionRequestModel extractionRequest, MainWindowViewModel parent)
        {
            model = new ExtractedData(extractionRequest);
            startDateTime = model.DateTimes[0];
            endDateTime = model.DateTimes[model.pointCount - 1];
            resolution = extractionRequest.Resolution;
            UpdatePoints();
            cursor1Values = new float[model.Tags.Length];
            plotRanges.Add(new PlotRange(startDateTime, endDateTime, yMin, yMax));
            currentZoomIndex = 0;

            if (parent != null)
                // Subscribe the plotwindow to the plotRangeChanged event transmitted by the main window
                WeakEventManager<MainWindowViewModel, PlotRangeChangedEventArgs>.AddHandler(parent, "TransmitPlotRangeChanged", OnPlotWindowRangeChanged);

            view = new PlotWindowView(this);
            view.Closed += (object source, EventArgs args) => PlotWindowClosed(this, args);
            view.Show();
        }

        #region Properties exposed to the View

        private SeriesCollection pointsToPlot;
        /// <summary>
        /// Contains the points to be shown in the plot. Binding path for the plot
        /// Each of Series contains pointsPerLine points
        /// </summary>
        public SeriesCollection PointsToPlot
        {
            get => pointsToPlot;
            set
            {
                if (pointsToPlot != value)
                {
                    pointsToPlot = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private DateTime[] dateTimesToPlot;
        /// <summary>
        /// Contains the time stamps of points to be shown in the plot.
        /// The array contains pointsPerLine points
        /// </summary>
        public DateTime[] DateTimesToPlot
        {
            get => dateTimesToPlot;
            set
            {
                if (dateTimesToPlot != value)
                {
                    dateTimesToPlot = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private string[] dateTimeStrs;
        /// <summary>
        /// Contains the label to be shown in the X axis of the plot. Binding path for the Labels property of AxisX
        /// Contains pointsPerLine strings
        /// </summary>
        public string[] DateTimeStrs
        {
            get => dateTimeStrs;
            set
            {
                if (dateTimeStrs != value)
                {
                    dateTimeStrs = value;
                    NotifyPropertyChanged();
                }
            }
        }

        
        private int resolution;
        /// <summary>
        /// The Resolution property corresponds to the parameter resolultion in the constructor
        /// It depicts the rough number of points in each of the line
        /// </summary>
        public int Resolution
        {
            get => resolution;
            set
            {
                if (resolution != value)
                {
                    resolution = value;
                    UpdatePoints();
                }
            }
        }

        /// <summary>
        /// The actual number of points on each line in the plot. 
        /// </summary>
        public int PointsPerLine { get; set; }

        /// <summary>
        /// This property decides whether this plot window will zoom with other synchronized plot windows.
        /// </summary>
        /// It's also binded to the SyncZoomCheckBox
        public bool SyncZoom { get; set; }

        private DateTime startDateTime;
        /// <summary>
        /// Contains start Date and Time from user input
        /// This will be the minimum of X axis in the plot
        /// </summary>
        public DateTime StartDateTime
        {
            get
                => startDateTime;
            set
            {
                //if (startDateTime != value)
                //{
                    startDateTime = value;
                    UpdatePoints();
                    if (SyncZoom)
                        PlotRangeChanged(this, new PlotRangeChangedEventArgs(startDateTime, endDateTime, this));
                    NotifyPropertyChanged();
                //}
            }
        }

        private DateTime endDateTime;
        /// <summary>
        /// Contains end Date and Time from user input
        /// This will be the maximum of X axis in the plot
        /// </summary>
        public DateTime EndDateTime
        {
            get
                => endDateTime;
            set
            {
                //if (endDateTime != value)
                //{
                    endDateTime = value;
                    UpdatePoints();
                    if (SyncZoom)
                        PlotRangeChanged(this, new PlotRangeChangedEventArgs(startDateTime, endDateTime, this));
                    NotifyPropertyChanged();
                //}
            }
        }

        private double yMin = Double.NaN;
        /// <summary>
        /// The min value of Y axis
        /// </summary>
        public double YMin
        {
            get
                => yMin;
            set
            {
                if (yMin != value)
                {
                    yMin = value;
                    NotifyPropertyChanged();
                }
            }
        }

        private double yMax = Double.NaN;
        /// <summary>
        /// The max value of Y axis
        /// </summary>
        public double YMax
        {
            get
                => yMax;
            set
            {
                if (yMax != value)
                {
                    yMax = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// The format to show DateTime values in the plot. Default is "yyyy/M/d H:mm:ss"
        /// </summary>
        public string DateTimeFormat = "yyyy/M/d H:mm:ss";


        private float[] cursor1Values;
        /// <summary>
        /// CursorValues is a list of float that represents the values of all tags at a time.
        /// This property is shown in the Legend with Values.
        /// </summary>
        public float[] Cursor1Values
        {
            get => cursor1Values;
            set
            {
                if (cursor1Values != value)
                {
                    cursor1Values = value;
                    NotifyPropertyChanged();
                }
            }
        }


        private string cursor1Time;
        /// <summary>
        /// Property for the displayed time for cursor 1
        /// </summary>
        public string Cursor1Time
        {
            get => cursor1Time;
            set
            {
                if (cursor1Time != value)
                {
                    cursor1Time = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public Func<double, string> DateTimeFormatter { get; set; }

        #endregion

        #region Commands exposed to the View

        /// <summary>
        /// Command for the ZoomPrevious button
        /// </summary>
        public ICommand ZoomPreviousCommand
        {
            get => new RelayCommand(() => MoveZoom(-1));
        }

        /// <summary>
        /// Command for the ZoomNext button
        /// </summary>
        public ICommand ZoomNextCommand
        {
            get => new RelayCommand(() => MoveZoom(1));
        }

        /// <summary>
        /// Command for the X axis min reset button. Change the StartDateTime to the beginning of all data
        /// </summary>
        public ICommand XAxisMinResetCommand
        {
            get => new RelayCommand(() => StartDateTime = model.DateTimes[0]);
        }

        /// <summary>
        /// Command for the X axis max reset button. Change the EndDateTime to the end of all data
        /// </summary>
        public ICommand XAxisMaxResetCommand
        {
            get => new RelayCommand(() => EndDateTime = model.DateTimes[model.pointCount - 1]);
        }

        /// <summary>
        /// Command for the Y axis min reset button. Change the YMin to NaN
        /// </summary>
        public ICommand YAxisMinResetCommand
        {
            get => new RelayCommand(() => YMin = Double.NaN);
        }

        /// <summary>
        /// Command for the Y axis max reset button. Change the YMax to NaN
        /// </summary>
        public ICommand YAxisMaxResetCommand
        {
            get => new RelayCommand(() => YMax = Double.NaN);
        }

        /// <summary>
        /// Command for the Export button. Will call the WriteToFile method of the model
        /// </summary>
        public ICommand ExportCommand
        {
            get => new RelayCommand(() => model.WriteToFile(StartDateTime, EndDateTime, "csv"));
        }

        #endregion

        #region private methods
        /// <summary>
        /// If the start/end datetime or Resolution is changed, will regenerate the PoitnsToPlot collection as well as the DateTimeStrs
        /// </summary>
        private void UpdatePoints()
        {
            int startIndex = 0, endIndex;
            SeriesCollection newPointsToPlot = new SeriesCollection();
            // Nothing to plot
            if (model.pointCount == 0)
            {
                PointsToPlot = newPointsToPlot;
                DateTimesToPlot = new DateTime[0];
                //DateTimeStrs = new string[0];
                return;
            }
            // Set the formatter that converts the number to a string
            DateTimeFormatter = value => new DateTime((long)(value * TimeSpan.FromDays(1).Ticks)).ToString(DateTimeFormat);
            if (startDateTime > endDateTime)
            {
                DateTime tempDateTime = startDateTime;
                startDateTime = endDateTime;
                endDateTime = tempDateTime;
            }
            // find the first timestamp that is larger than startDateTime. Will start PickPoints from here
            while (startIndex < model.pointCount && model.DateTimes[startIndex] < startDateTime)
                startIndex++;
            // If we reach the end of the data and found no time stamp larger than startDateTime, nothing to plot
            if (startIndex == model.pointCount)
            {
                PointsToPlot = newPointsToPlot;
                DateTimesToPlot = new DateTime[0];
                //DateTimeStrs = new string[0];
                return;
            }
            endIndex = startIndex;
            // find the fist timeStamp that is larger than endDateTime. End PickPoints one point before that.
            while (endIndex < model.pointCount && model.DateTimes[endIndex] <= endDateTime)
                endIndex++;
            endIndex--;

            if (startIndex == -1)
                startIndex = 0;
            if (endIndex == -1)
                endIndex = model.pointCount - 1;
            // The local variable for storing the points to be plotted
            ValWithTimeStamp[] temp;
            // Figure out actual interval to take points from ExtractedData to pointsToPlot
            int interval = (endIndex - startIndex) / (resolution - 1);
            if (interval < 1)
                interval = 1;
            // pointCount is the actual number of points in each line
            int pointCount = (endIndex - startIndex) / interval + 1;
            // oldIndex is the index of a point in the rawData. newIndex is that in the new array
            int oldIndex, newIndex;
            // Copy data from model.RawData
            for (int i = 0; i < model.RawData.Count; i++) // for each tag (each array in RawData)
            {
                // create an empty array to take the data
                temp = new ValWithTimeStamp[pointCount];
                oldIndex = startIndex;
                // copy data from the array of values
                for (newIndex = 0; newIndex < pointCount; newIndex++)
                {
                    //temp[newIndex].Value = model.RawData[i][oldIndex];
                    //temp[newIndex].TimeStamp = model.DateTimes[oldIndex];
                    temp[newIndex] = new ValWithTimeStamp() {
                        Value = model.RawData[i][oldIndex],
                        TimeStamp = model.DateTimes[oldIndex]
                    };
                    oldIndex += interval;
                }
                var series = new LineSeries(dateTimeConfig)
                {
                    Title = model.Tags[i],
                    Values = new ChartValues<ValWithTimeStamp>(temp),
                    LineSmoothness = 0,
                    PointGeometry = null,
                    Fill = Brushes.Transparent,
                };
                newPointsToPlot.Add(series);
            }
            PointsToPlot = newPointsToPlot;
            PointsPerLine = pointCount;
            // copy time stamps from model.DateTimes
            //string[] newDateTimeStrs = new string[pointCount];
            DateTime[] newDateTimesToPlot = new DateTime[pointCount];
            oldIndex = startIndex;
            for (newIndex = 0; newIndex < pointCount; newIndex++)
            {
                newDateTimesToPlot[newIndex] = model.DateTimes[oldIndex];
                //newDateTimeStrs[newIndex] = newDateTimesToPlot[newIndex].ToString(DateTimeFormat);
                oldIndex += interval;
            }
            DateTimesToPlot = newDateTimesToPlot;
            //DateTimeStrs = newDateTimeStrs;
        }

        /// <summary>
        /// Write the current plot rage into the plotRanges list.
        /// This method should be called whenever the plot rnage is changed by zooming or sizing.
        /// </summary>
        private void RecordRange()
        {
            // If currentZoomIndex = plotRanges.Count-1, i.e. currently at the last zoom, then new zooming will add new entry to List plotRanges
            // If currenZoomIndex is anything smaller, then new zooming will erase the following plotRanges and create new item
            if (currentZoomIndex < plotRanges.Count - 1)
                plotRanges.RemoveRange(currentZoomIndex + 1, plotRanges.Count - currentZoomIndex - 1);
            plotRanges.Add(new PlotRange(StartDateTime, EndDateTime, YMin, YMax));
            currentZoomIndex = plotRanges.Count - 1;
        }


        private void SetRange(PlotRange newRange)
        {
            startDateTime = newRange.startDateTime;
            NotifyPropertyChanged("StartDateTime");
            EndDateTime = newRange.endDateTime;
            YMax = newRange.yMax;
            YMin = newRange.yMin;
        }

        /// <summary>
        /// When the range of another window is changed and these windows are all in SyncZoom mode, adjust the X range of this plot
        /// </summary>
        /// <param name="source">The plot that chagned range. If the source is this plot, will do nothing.</param>
        /// <param name="e">This PlotRangeChangedEventArgs contains the new range.</param>
        private void OnPlotWindowRangeChanged(object source, PlotRangeChangedEventArgs e)
        {
            if (e.InitialSource as PlotWindowViewModel != this)
            {
                if (SyncZoom) // Only when SyncZoom == true will this window respond to this event
                {
                    if (e.StartDateTime != StartDateTime || e.EndDateTime != EndDateTime)
                    {
                        // We are modifying the fields to avoid firing the PlotRangeChanged event again.
                        if (e.StartDateTime != StartDateTime)
                        {
                            startDateTime = e.StartDateTime;
                            NotifyPropertyChanged("StartDateTime");
                        }
                        if (e.EndDateTime != EndDateTime)
                        {
                            endDateTime = e.EndDateTime;
                            NotifyPropertyChanged("EndDateTime");
                        }
                        UpdatePoints();
                        RecordRange();
                    }
                }
            }

        }

        #endregion


        #region Public methods exposed to View

        /// <summary>
        /// Change the rance in the X axis and record it to the List 
        /// </summary>
        /// <param name="newStartDatetime"></param>
        /// <param name="newEndDateTime"></param>
        public void SetXRangeAndRecord(DateTime newStartDatetime, DateTime newEndDateTime)
        {
            startDateTime = newStartDatetime;
            NotifyPropertyChanged("StartDateTime");
            EndDateTime = newEndDateTime;
            RecordRange();
        }


        /// <summary>
        /// Move the plot range to previous or next in List plotRanges
        /// </summary>
        /// <param name="zoomIndex">Indicate the direction to move in the List plotRange
        /// zoomIndex = 0 corresponds to the starting range; zoomIndex = -1 will move the range back by one step; zoomIndex = 1 will move the range forward by one step</param>
        public void MoveZoom(int zoomIndex)
        {
            switch (zoomIndex)
            {
                case -1: // zoom previous
                    if (currentZoomIndex > 0)
                        currentZoomIndex--;
                    else // Already at starting range. No need to do anything
                        return;
                    break;
                case 0: // zoom to starting range
                    if (currentZoomIndex > 0)
                        currentZoomIndex = 0;
                    else
                        return;
                    break;
                case 1: // zoom next
                    if (currentZoomIndex < plotRanges.Count - 1)
                        currentZoomIndex++;
                    else
                        return; // already at last range. No need to do anything
                    break;
            }
            SetRange(plotRanges[currentZoomIndex]);
        }



        /// <summary>
        /// Update the tag values based on the X axis positionin the chart
        /// </summary>
        /// <param name="chartValue">The X position of the point. 
        /// Typically, it shoule come from: chartValues = Chart.ConvertToChartValues(e.GetPosition(Chart)).X, 
        /// where e is a MouseEventArg</param>
        public void UpdateLegendValues(double chartValue)
        {
            DateTime cursorDateTime = new DateTime((long)(chartValue * TimeSpan.FromDays(1).Ticks));
            // Find the location of the cursor datetime in the data
            int valueIndex = 0;
            for(;valueIndex < PointsPerLine; valueIndex++)
            {
                if (DateTimesToPlot[valueIndex] >= cursorDateTime) break;
            }
            if (valueIndex > 0 && valueIndex < PointsPerLine) // determine which of the two adjacent points is closest to the cursor datetime
            {
                if (DateTimesToPlot[valueIndex] - cursorDateTime > cursorDateTime - DateTimesToPlot[valueIndex - 1]) // Cursor datetime is closer to the previous point
                    valueIndex--;
            }
            else if (valueIndex == PointsPerLine) // The cursor datetime is after the end time of the plot, use the last point in the plot
                valueIndex--;
            
            // If we change the elements of CursorValues one by one, the setter will not be called, and NotifyPropertyChanged will not be fired
            // In addition, the DependencyProperty in LegendWith Values will be changed but the PropertyChangedCallback will not be triggered
            // This is probably because the array object (reference to the array) is never chagned. 
            // Assigning a new array to it will trigger the PropertyChangedCallback
            float[] temp = new float[model.Tags.Length];
            for (int i = 0; i < PointsToPlot.Count; i++)
            {
                temp[i] = ((ValWithTimeStamp)PointsToPlot[i].Values[valueIndex]).Value;
            }
            Cursor1Values = temp;
            Cursor1Time = DateTimesToPlot[valueIndex].ToString(DateTimeFormat);
        }

        #endregion

        #region Private fields

        /// <summary>
        /// This is the model, which contains all the data
        /// </summary>
        private ExtractedData model;

        /// <summary>
        /// This is the view.
        /// </summary>
        private PlotWindowView view;

        /// <summary>
        /// The List of PlotRange will keep track of all previous zooming activities. Thus, zooming can be reversed
        /// </summary>
        private List<PlotRange> plotRanges = new List<PlotRange>();
       
        /// <summary>
        /// currentZoomIndex indicates where we are at the List of plotRanges.
        /// </summary>
        /// If currentZoomIndex = plotRanges.Count-1, i.e. currently at the last zoom, then new zooming will add new entry to List plotRanges
        /// If currenZoomIndex is anything smaller, then new zooming will erase the following plotRanges and create new item
        private int currentZoomIndex;

        /// <summary>
        /// This mapper is used to show a value with time stamp in a charg
        /// </summary>
        private LiveCharts.Configurations.CartesianMapper<ValWithTimeStamp> dateTimeConfig = LiveCharts.Configurations.Mappers.Xy<ValWithTimeStamp>()
            .X(valWithTimeStamp => (double)valWithTimeStamp.TimeStamp.Ticks/TimeSpan.FromDays(1).Ticks)
            .Y(valWithTimeStamp => valWithTimeStamp.Value);

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

        #region Events
        /// <summary>
        /// This event is fired when SyncZoom is true and the X range of the plot is changed.
        /// The event fired here is subscribed by the MainWindow, which will transmit the event to other PlotWindows
        /// The listeners (PlotWindows) with SyncZoom set to true will update the X range of their plot
        /// </summary>
        public event EventHandler<PlotRangeChangedEventArgs> PlotRangeChanged = delegate { };

        /// <summary>
        /// This event is fired when the view is closed. It's used to relay the Closed event to the main window
        /// </summary>
        public event EventHandler<EventArgs> PlotWindowClosed = delegate { };

        #endregion

        /// <summary>
        /// This struct contains information of the min and max of both X and Y axis in the plot
        /// </summary>
        private struct PlotRange
        {
            public DateTime startDateTime;
            public DateTime endDateTime;
            public double yMin;
            public double yMax;

            public PlotRange(DateTime startDateTime, DateTime endDateTime, double yMin, double yMax)
            {
                this.startDateTime = startDateTime;
                this.endDateTime = endDateTime;
                this.yMax = yMax;
                this.yMin = yMin;
            }
        }
        
        private class ValWithTimeStamp
        {
            public DateTime TimeStamp { get; set; }
            public float Value { get; set; }
        }
    }
}
