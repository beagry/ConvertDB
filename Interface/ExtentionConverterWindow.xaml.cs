using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using ExcelRLibrary;
using UI.Annotations;

namespace UI
{
    /// <summary>
    /// Interaction logic for ExtentionConverterWindow.xaml
    /// </summary>
    public partial class ExtentionConverterWindow : Window
    {
        private ExtentionConverterViewModel viewModel;
        public ExtentionConverterWindow()
        {
            InitializeComponent();

            ExtentionsComboBox.ItemsSource = Enum.GetValues(typeof(XlSimpleSaveType)).Cast<XlSimpleSaveType>();
            viewModel = new ExtentionConverterViewModel();
            DataContext = viewModel;
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {

            var saver = new ExcelSaver {SaveType = viewModel.SaveType,OverWrite = viewModel.OverWrite};
            await Task.Run(() =>
            {
                foreach (var path in viewModel.WorkbooksToResave.Select(w => w.Path))
                    saver.ResaveWorkbook(path);
            });
            viewModel.ResetModel();
        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files != null)
                files.ToList().ForEach(s =>
                {
                    if (viewModel.WorkbooksToResave.All(w => w.Path != s))
                    {
                        viewModel.WorkbooksToResave.Add(new SelectedWorkbook(s));
                    }
                });
        }
    }

    sealed class ExtentionConverterViewModel:INotifyPropertyChanged
    {
        private XlSimpleSaveType saveType;
        private bool overWrite;

        public XlSimpleSaveType SaveType
        {
            get { return saveType; }
            set
            {
                if (value == saveType) return;
                saveType = value;
                OnPropertyChanged("Type");
            }
        }

        public bool OverWrite
        {
            get { return overWrite; }
            set
            {
                if (overWrite == value) return;
                overWrite = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<SelectedWorkbook> WorkbooksToResave { get; set; }

        public ExtentionConverterViewModel()
        {
            WorkbooksToResave = new ObservableCollection<SelectedWorkbook>();
            SaveType = XlSimpleSaveType.Xlsx;
            OverWrite = false;
        }

        public void ResetModel()
        {
            SaveType = XlSimpleSaveType.Xlsx;
            WorkbooksToResave.Clear();
            OverWrite = false;
        }

        #region PropChanged
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
