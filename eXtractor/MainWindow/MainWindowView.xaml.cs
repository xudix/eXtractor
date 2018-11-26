using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace eXtractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindowView : Window
    {
        public MainWindowView()
        {
            InitializeComponent();
        }

        #region The following method is used to control the double click behavior of the date and time input TextBox
        // When user click it the first time or double click it, it will select all text.
        // The XAML file register GotKeyboardFocus and MouseDoubleClick event to Textbox_GotFocus method, 
        // and PreviewMouseLeftButtonDown event to SelectivelyIgnoreMouseButton. 
        // It seems that when the mouse click on the textbox, it fires GotKeyboardFocus then PreviewMouseLeftButtonDown. By ignoring the second event, the mouse click will not turn "select all" into the curser at a point.

        // Method comes from https://social.msdn.microsoft.com/Forums/vstudio/en-US/564b5731-af8a-49bf-b297-6d179615819f/how-to-selectall-in-textbox-when-textbox-gets-focus-by-mouse-click?forum=wpf

        private void SelectivelyIgnoreMouseButton(object sender, MouseButtonEventArgs e)
        {
            if (sender != null && !(sender as TextBox).IsKeyboardFocusWithin)
            {
                e.Handled = true;
                (sender as TextBox).Focus();
                //Console.Write("SelIgnor Sender name: " + (sender as TextBox).Name + "; Event: " + e.RoutedEvent + "\r\n");
            }
        }

        private void Textbox_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as TextBox).SelectAll();
            e.Handled = true;
            //Console.Write("GotFocus Sender name: " + (sender as TextBox).Name + "; Event: " + e.RoutedEvent + "\r\n");
        }

        #endregion
    }


}
