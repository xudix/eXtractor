using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace eXtractor
{
    public class MainWindowViewModel: INotifyPropertyChanged
    {

        public MainWindowViewModel()
        {
            model = ExtractionRequestModel.CreateExtractionRequestModel();
        }

        private ExtractionRequestModel model;

        #region Properties exposed to the view
        /// <summary>
        /// The Date part of StartDateTime
        /// </summary>
        public DateTime StartDate
        {
            get => model.StartDateTime.Date;
            set
            {
                if(value != model.StartDateTime.Date)
                {
                    model.StartDateTime = value + model.StartDateTime.TimeOfDay;
                    NotifyPropertyChanged();
                }
            }
        }
        /// <summary>
        /// The Date part of EndDateTime
        /// </summary>
        public DateTime EndDate
        {
            get => model.EndDateTime.Date;
            set
            {
                if (value != model.EndDateTime.Date)
                {
                    model.EndDateTime = value + model.EndDateTime.TimeOfDay;
                    NotifyPropertyChanged();
                }
            }
        }
        /// <summary>
        /// The Time part of StartDateTime
        /// </summary>
        public TimeSpan StartTime
        {
            get => model.StartDateTime.TimeOfDay;
            set
            {
                if (value != model.StartDateTime.TimeOfDay)
                {
                    model.StartDateTime = model.StartDateTime.Date + value;
                    NotifyPropertyChanged();
                }
            }
        }
        /// <summary>
        /// The Time part of EndDateTime
        /// </summary>
        public TimeSpan EndTime
        {
            get => model.EndDateTime.TimeOfDay;
            set
            {
                if (value != model.EndDateTime.TimeOfDay)
                {
                    model.EndDateTime = model.EndDateTime.Date + value;
                    NotifyPropertyChanged();
                }
            }
        }
        /// <summary>
        /// The program will take a data point every "Interval" time steps. 
        /// When Interval == 1, it takes all data points; when Interval ==2, it skips a data point after taking one.
        /// </summary>
        public int Interval
        {
            get => model.Interval;
            set
            {
                if (value != model.Interval)
                {
                    model.Interval = value;
                    NotifyPropertyChanged();
                }
            }
        }
        /// <summary>
        /// Resolution is the rough number of points shown in each line in the plot.
        /// It does not affect the number of points taken from the data file.
        /// A higher resolution may lower the PlotWindow performance
        /// </summary>
        public int Resolution
        {
            get => model.Resolution;
            set
            {
                if (value != model.Resolution)
                {
                    model.Resolution = value;
                    NotifyPropertyChanged();
                }
            }
        }
        /// <summary>
        /// Selected Tags. It's an array of string. Each item in the array is a requested tag.
        /// </summary>
        public string[] SelectedTags
        {
            get => model.SelectedTags;
            set
            {
                if(value != model.SelectedTags)
                {
                    model.SelectedTags = value;
                    NotifyPropertyChanged();
                }
            }
        }
        /// <summary>
        /// Selected data files, an array of string. Each item in the array is the path of a data file
        /// </summary>
        public string[] SelectedFiles
        {
            get => model.SelectedFiles;
            set
            {
                if (value != model.SelectedFiles)
                {
                    model.SelectedFiles = value;
                    NotifyPropertyChanged();
                }
            }
        }


        #endregion

        #region Commands exposed to the view
        /// <summary>
        /// This command will open a OpenFileDialog and allow the user to select a data file.
        /// Then the first row of the file will be read and converted to a list of tags.
        /// The list will be loaded to a PickTagWindow, which allows the user to select the tags needed.
        /// </summary>
        public ICommand PickTagCommand
        {
            get => new RelayCommand(PickTag_fromDataFile);
        }

        /// <summary>
        /// This method will open a OpenFileDialog and allow the user to select several data files.
        /// The selected files will be appended to the SelectedFiles property
        /// </summary>
        public ICommand PickFileCommand
        {
            get => new RelayCommand(PickFile);
        }

        public ICommand ExportCommand
        {
            get => new RelayCommand(model.Export);
        }

        public ICommand PlotCommand
        {
            get => new RelayCommand(Plot);
        }

        #endregion

        #region private methods
        /// <summary>
        /// This method will open a OpenFileDialog and allow the user to select a data file.
        /// Then the first row of the file will be read and converted to a list of tags.
        /// The list will be loaded to a PickTagWindow, which allows the user to select the tags needed.
        /// The selected tags will be loaded to the SelectedTags property
        /// </summary>
        private void PickTag_fromDataFile()
        {
            string tagList = String.Empty;
            string[] tagArray = new string[0];
            // The dialog to select Tag List
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Select Data File",
                DefaultExt = ".*",
                Filter = "All Files|*.*|Excel Documents (.xlsx)|*.xlsx|CSV files (.csv)|*.csv|Text documents (.txt)|*.txt",
            };
            if (!String.IsNullOrEmpty(model.FilePath)) dialog.InitialDirectory = model.FilePath;

            // If the file is selected, record the file path and open the file
            if (dialog.ShowDialog() == true)
            {
                model.FilePath = Path.GetDirectoryName(dialog.FileName);
                if (Path.GetExtension(dialog.FileName).ToLower() == ".xlsx")
                {
                    try
                    {
                        using (ZipArchive xlsxFile = new ZipArchive(new FileStream(dialog.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                        {
                            tagArray = XlsxTool.GetHeaderWithColReference(xlsxFile).header;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to obtain tags. Please select an XLSX, CSV, or TXT data file. \nOriginal error: " + ex.Message);
                        return;
                    }
                }
                else
                {
                    try
                    {
                        using (StreamReader sr = new StreamReader(new FileStream(dialog.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
                        {
                            // Read the first line from the file. 
                            tagList = sr.ReadLine();
                            // Remove the "Date", "Time", and "Millitm" fields from the first line
                            tagList = Regex.Replace(tagList, @"^(;?date)?[\W_]*time[\W_]*(millitm)?", "", RegexOptions.IgnoreCase);
                            if (!String.IsNullOrEmpty(tagList))
                            {
                                char[] separators = { ' ', ',', '\t', '\n', '\r', ';' };
                                tagArray = tagList.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to obtain tags. Please select an XLSX, CSV, or TXT data file. \nOriginal error: " + ex.Message);
                        return;
                    }
                }
                if (tagArray.Length > 0)
                {
                    PickTagWindowView pickTagDialog = new PickTagWindowView(tagArray);
                    //pickTagDialog.Owner = this;
                    // The ShowDialog() method of Window class will show the window and disable the mian window.
                    if (pickTagDialog.ShowDialog() == true && pickTagDialog.SelectedTags != null)
                    {
                        SelectedTags = (SelectedTags != null) ? SelectedTags.Concat(pickTagDialog.SelectedTags).ToArray() : pickTagDialog.SelectedTags;
                    }
                }
            }
        }

        /// <summary>
        /// This method will open a OpenFileDialog and allow the user to select several data files.
        /// The selected files will be appended to the SelectedFiles property
        /// </summary>
        private void PickFile()
        {
            // The dialog to select data files
            OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "Select Data Files",
                DefaultExt = ".*",
                Filter = "All Files|*.*|Excel Documents (.xlsx)|*.xlsx|CSV files (.csv)|*.csv|Text documents (.txt)|*.txt",
                Multiselect = true
            };
            if (!String.IsNullOrEmpty(model.FilePath)) dialog.InitialDirectory = model.FilePath;

            // If the file is selected, record the file path and open the file
            if (dialog.ShowDialog() == true && dialog.FileNames != null)
            {
                model.FilePath = Path.GetDirectoryName(dialog.FileNames[0]);
                //SelectedFiles = SelectedFiles  + String.Join("\r\n", dialog.FileNames) + "\r\n";
                SelectedFiles = (SelectedFiles != null) ? SelectedFiles.Concat(dialog.FileNames).ToArray() : dialog.FileNames;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void Plot()
        {
            if (SelectedTags != null && SelectedFiles != null)
            {
                try
                {
                    model.SaveSettings();
                    PlotWindow.PlotWindowViewModel plotWindowViewModel = new PlotWindow.PlotWindowViewModel(model, this);
                    //WeakEventManager<PlotWindow, EventArgs>.AddHandler(plotWindow, "Closed", OnPlotWindowClosed);
                    //WeakEventManager<PlotWindow, PlotRangeChangedEventArgs>.AddHandler(plotWindow, "PlotRangeChanged", OnPlotWindowRangeChanged);

                    GC.Collect();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Fail to show plot window. Original error: " + ex.Message + "\n" + ex.StackTrace);
                }
            }
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
