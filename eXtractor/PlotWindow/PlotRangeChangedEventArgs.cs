using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eXtractor.PlotWindow
{
    /// <summary>
    /// Contains the nwe start datetime and end datetime for the plot range changed event.
    /// </summary>
    public class PlotRangeChangedEventArgs : EventArgs
    {
        public PlotRangeChangedEventArgs(DateTime startDateTime, DateTime endDateTime, PlotWindowViewModel initialSource)
        {
            StartDateTime = startDateTime;
            EndDateTime = endDateTime;
            InitialSource = initialSource;
        }

        public DateTime StartDateTime, EndDateTime;
        public PlotWindowViewModel InitialSource;

    }
}
