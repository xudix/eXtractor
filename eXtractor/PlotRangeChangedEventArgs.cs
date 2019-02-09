using System;

namespace eXtractor
{
    /// <summary>
    /// Contains the new start datetime and end datetime for the plot range changed event.
    /// </summary>
    public class PlotRangeChangedEventArgs : EventArgs
    {
        public PlotRangeChangedEventArgs(DateTime startDateTime, DateTime endDateTime, PlotWindow.PlotWindowViewModel initialSource)
        {
            StartDateTime = startDateTime;
            EndDateTime = endDateTime;
            InitialSource = initialSource;
        }
        public DateTime StartDateTime, EndDateTime;
        public PlotWindow.PlotWindowViewModel InitialSource;
    }
}