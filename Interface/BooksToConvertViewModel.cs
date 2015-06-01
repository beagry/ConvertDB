using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Converter.Properties;
using Converter.Template_workbooks;
using ExcelRLibrary;

namespace UI
{
    public sealed class BooksToConvertViewModel:INotifyPropertyChanged
    {
        private XlTemplateWorkbookType workbooksType;
        private bool editMode;
        private string status;
        private bool workInProgress;

        public BooksToConvertViewModel()
        {
            EditMode = true;
            WorkbooksType = XlTemplateWorkbookType.LandProperty;
            Workbooks = new ObservableCollection<SelectedWorkbook>();
            Status = "Готово к работе";
        }
        public BooksToConvertViewModel(IEnumerable<SelectedWorkbook> workbooksPaths, XlTemplateWorkbookType workbooksType):this()
        {
            Workbooks = new ObservableCollection<SelectedWorkbook>(workbooksPaths);
            WorkbooksType = workbooksType;
        }

        public bool EditMode
        {
            get { return editMode; }
            set
            {
                if (editMode == value) return;
                editMode = value;
                OnPropertyChanged();
            }
        }

        public bool WorkInProgress
        {
            get { return workInProgress; }
            set
            {
                if (workInProgress == value)return;
                workInProgress = value;
                OnPropertyChanged();
            }
        }

        public string Status
        {
            get
            {
                return status;
            }
            set
            {
                if (status == value) return;
                status = value;
                OnPropertyChanged();
            }
        }

        public void StartWork()
        {
            var message = "В процессе...";
            StartWork(message);
        }

        public void StartWork(string message)
        {
            EditMode = false;
            WorkInProgress = true;
            Status = message;
        }


        public void EndWork()
        {
            EndWork("Готово");
        }

        public void EndWork(string message)
        {
            EditMode = true;
            WorkInProgress = false;
            Status = message;
        }


        public ObservableCollection<SelectedWorkbook> Workbooks{ get; set; }

        public XlTemplateWorkbookType WorkbooksType
        {
            get { return workbooksType; }
            set
            {
                if (value == workbooksType) return;
                workbooksType = value;
                OnPropertyChanged();
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
