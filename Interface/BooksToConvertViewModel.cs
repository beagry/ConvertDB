using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Converter;
using Converter.Annotations;

namespace UI
{
    public sealed class BooksToConvertViewModel:INotifyPropertyChanged
    {
        private XlTemplateWorkbookTypes workbooksType;

        public BooksToConvertViewModel()
        {
            workbooksType = XlTemplateWorkbookTypes.LandProperty;
            Workbooks = new ObservableCollection<SelectedWorkbook>();
        }
        public BooksToConvertViewModel(IEnumerable<SelectedWorkbook> workbooksPaths, XlTemplateWorkbookTypes workbooksType)
        {
            Workbooks = new ObservableCollection<SelectedWorkbook>(workbooksPaths);
            WorkbooksType = workbooksType;
        }

        public ObservableCollection<SelectedWorkbook> Workbooks{ get; set; }

        public XlTemplateWorkbookTypes WorkbooksType
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
