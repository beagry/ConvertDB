using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
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
using Converter.Template_workbooks;
using ExcelRLibrary;
using Formater;
using Microsoft.Win32;

namespace UI
{
    /// <summary>
    /// Interaction logic for FormatWbWindow.xaml
    /// </summary>
    public partial class FormatWbWindow : Window
    {
        private readonly FormatDbViewModel viewModel;
        public FormatWbWindow()
        {
            InitializeComponent();
            viewModel = new FormatDbViewModel();
            DataContext = viewModel;
        }

        private void SelectWbButton_Click(object sender, RoutedEventArgs e)
        {
            viewModel.Path = SelectFile();
        }

        private void CatalogButton_Click(object sender, RoutedEventArgs e)
        {
            viewModel.CatalogSupportWorkbook.Path = SelectFile();
        }

        private string SelectFile(string msg = "")
        {
            var fd = new OpenFileDialog
            {
                Multiselect = false,
                Filter = "Excel Files (*.xlsx, *.csv)| *.xlsx; *.csv",
                Title = msg
            };

            return fd.ShowDialog() == true ? fd.FileName : "";
        }

        private void OKTMOButton_Click(object sender, RoutedEventArgs e)
        {
            viewModel.OktmoSupportWorkbook.Path = SelectFile();
        }

        private void VGTCatalogButton_Click(object sender, RoutedEventArgs e)
        {
            viewModel.VgtCatalogSupportWorkbook.Path = SelectFile();
        }

        private void SubjSourceButton_Click(object sender, RoutedEventArgs e)
        {
            viewModel.SubjectSourceSupportWorkbook.Path = SelectFile();
        }

        private async   void StartButton_Click(object sender, RoutedEventArgs e)
        {
            var convert = new DbToConvert(viewModel)
            {
                ColumnsToReserve = new List<string> { "SUBJECT", "REGION", "NEAR_CITY", "SYSTEM_GAS", "SYSTEM_WATER", "SYSTEM_SEWERAGE", "SYSTEM_ELECTRICITY" },
                DoDescription =  viewModel.DoDescription
            };
            var button = sender as Button;
            if (button == null) return;
            viewModel.StartWork();
            var checkHeadResult =  await Task.Run(() => convert.ColumnHeadIsOk());
            if (!checkHeadResult) return;

            //Запусть обработки в новом потоке
            await Task.Run(() => convert.FormatWorksheet());

            convert.ExcelPackage.SaveWithDialog();
            viewModel.EndWork();
        }
    }

    sealed class FormatDbViewModel:ViewModelAbstract, IFormatDbParams
    {
        private string path;

        public FormatDbViewModel():base()
        {
            Enums = new ObservableCollection<EnumView<XlTemplateWorkbookType>>();
            foreach (XlTemplateWorkbookType e in Enum.GetValues(typeof(XlTemplateWorkbookType)))
                Enums.Add(new EnumView<XlTemplateWorkbookType>(e));

            CatalogSupportWorkbook = new SupportWorkbookViewModel();
            OktmoSupportWorkbook = new SupportWorkbookViewModel();
            VgtCatalogSupportWorkbook = new SupportWorkbookViewModel();
            SubjectSourceSupportWorkbook = new SupportWorkbookViewModel();
        }

        public string Path  
        {
            get { return path; }
            set
            {
                if (path == value) return;
                path = value;
                OnPropertyChanged();
            }
        }


        public ISupportWorkbookViewModel CatalogSupportWorkbook { get; set; }
        public ISupportWorkbookViewModel OktmoSupportWorkbook { get; set; }
        public ISupportWorkbookViewModel VgtCatalogSupportWorkbook { get; set; }
        public ISupportWorkbookViewModel SubjectSourceSupportWorkbook { get; set; }


        public XlTemplateWorkbookType WorkbookType { get; set; }
        public bool DoDescription { get; set; }
        public ObservableCollection<EnumView<XlTemplateWorkbookType>> Enums { get; set; }

//        public new event PropertyChangedEventHandler PropertyChanged;

//        [NotifyPropertyChangedInvocator]
//        protected override void OnPropertyChanged([CallerMemberName] string propertyName = null)
//        {
//            if (PropertyChanged == null) return;
//            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
//        }
    }


    internal sealed class SupportWorkbookViewModel : ISupportWorkbookViewModel
    {
        private Task<List<string>> initialWsNamesTask;

        private string path;
        private bool taskInProgress;
        private bool hasWorksheets;

        public string Path
        {
            get { return path; }
            set
            {
                if (path == value) return;
                path = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<string> Worksheets { get; set; }

        public bool HasWorksheets
        {
            get { return hasWorksheets; }
            set
            {
                if(value==hasWorksheets) return;
                hasWorksheets = value; 
                OnPropertyChanged();
            }
        }

        public string SelectedWorksheet { get; set; }


        public SupportWorkbookViewModel()
        {
            Worksheets = new ObservableCollection<string>();
            PropertyChanged += OnPathchanged;
            Worksheets.CollectionChanged += (sender, args) => CheckWorksheetsCount();
        }

        private void OnPathchanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "Path") return;
            InitialWorksheetsList();
        }

        private async void InitialWorksheetsList()
        {
            if (!File.Exists(Path)) return;
            if (initialWsNamesTask != null)
                await initialWsNamesTask;

            TaskInProgress = true;

            var reader = new ExcelReader();
            Worksheets.Clear();

            initialWsNamesTask = Task.Run(() => reader.GetWorksheetsNames(Path));
            var wss = await initialWsNamesTask;
            wss.ForEach(s => Worksheets.Add(s));
//            CheckWorksheetsCount();

            TaskInProgress = false;
        }

        private void CheckWorksheetsCount()
        {
            HasWorksheets = Worksheets.Any();
        }

    public bool TaskInProgress
        {
            get { return taskInProgress; }
            set
            {
                if (taskInProgress == value) return;
                taskInProgress = value; 
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged == null) return;
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
