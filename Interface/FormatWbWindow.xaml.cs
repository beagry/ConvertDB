using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
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
using Converter.Properties;

namespace UI
{
    /// <summary>
    /// Interaction logic for FormatWbWindow.xaml
    /// </summary>
    public partial class FormatWbWindow : Window
    {
        public FormatWbWindow()
        {
            InitializeComponent();
        }
    }

    class FormatDbViewModel:INotifyPropertyChanged
    {
        public string Path { get; set; }

        public ICommand FormatWorkbook()
        {
            throw new NotImplementedException();
        }


        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged == null) return;
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
