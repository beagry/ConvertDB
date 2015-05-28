using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Converter;
using Converter.Properties;
using Converter.Template_workbooks;
using ExcelRLibrary;

namespace UI
{
    public sealed class BooksToConvertViewModel:INotifyPropertyChanged
    {
        private XlTemplateWorkbookType workbooksType;
        private bool editMode;

        public BooksToConvertViewModel()
        {
            EditMode = true;
            workbooksType = XlTemplateWorkbookType.LandProperty;
            Workbooks = new ObservableCollection<SelectedWorkbook>();
        }
        public BooksToConvertViewModel(IEnumerable<SelectedWorkbook> workbooksPaths, XlTemplateWorkbookType workbooksType)
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

        public ObservableCollection<SelectedWorkbook> Workbooks{ get; set; }

        public XlTemplateWorkbookType WorkbooksType
        {
            get { return workbooksType; }
            set
            {
                if (value == workbooksType) return;
                workbooksType = value;
                OnPropertyChanged("WorkbooksType");
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
