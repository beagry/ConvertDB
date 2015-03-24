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
using Converter;
using Telerik.Windows.Controls;

namespace WpfApplication2
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        private CompareViewModel view;
        public Window1(Dictionary<string, ObservableCollection<string>> dict, ICollection<WorksheetInfo> wsInfos)
        {
            InitializeComponent();
            view = new CompareViewModel(dict,wsInfos);
//            this.DataContext = view;
            UnbindexListBox.ItemsSource = view.UnbindedColumns;
            BindedColumnsListBox.ItemsSource = view.BindedColumnsDictionary;
        }
    }
}
