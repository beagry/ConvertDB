using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Converter.Annotations;

namespace UnionWorkbooks
{
    sealed class ViewModel:INotifyPropertyChanged
    {
        public int MaxRequiredItems { get; set; }

        public long TotalItemsQuantity { get { return Workbooks.Sum(w => w.MaxRowsInWorkbook); } }

        public ObservableCollection<WorkbookWithItemsQnt> Workbooks { get; set; }

        public ObservableCollection<string> WorksheetsToCopy { get; set; }

        public List<string> AllWorksheetsCollection
        {
            get { return Workbooks.SelectMany(w => w.WorksheetsNamesList).Distinct().ToList(); }
        }

        public byte HeadSize { get; set; }
        

        public ViewModel()
        {
            MaxRequiredItems = 500000;
            Workbooks = new ObservableCollection<WorkbookWithItemsQnt>();
            WorksheetsToCopy = new ObservableCollection<string>();
            HeadSize = 8;
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