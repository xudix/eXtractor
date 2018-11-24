using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Shapes;

namespace eXtractor
{
    /// <summary>
    /// Interaction logic for PickTagWindowView.xaml
    /// </summary>
    public partial class PickTagWindowView : Window
    {

        public ObservableCollection<string> TagCollection { get; set; }

        public string[] SelectedTags => (tagListBox.SelectedItems != null) ? tagListBox.SelectedItems.Cast<string>().ToArray() : null;

        public PickTagWindowView(string[] tagArray)
        {
            InitializeComponent();
            DataContext = this;
            TagCollection = new ObservableCollection<string>(tagArray);
        }

        private void DoneButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;

            //this.Close();
        }
    }
}
