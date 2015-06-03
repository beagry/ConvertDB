using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Converter.Properties;

namespace UnionWorkbooks
{
    internal sealed class ViewModel : INotifyPropertyChanged
    {
        private bool allSheetsInOne;
        private bool editMode;
        private bool workInProgress;

        public ViewModel()
        {
            EditMode = true;
            MaxRequiredItems = 500000;
            Workbooks = new ObservableCollection<WorkbookWithItemsQnt>();
            WorksheetsToCopy = new ObservableCollection<string>();
            HeadSize = 8;

            UpdaWorkbooksDepends();

            Workbooks.CollectionChanged += (sender, args) => { UpdaWorkbooksDepends(); };

            WorksheetsToCopy.CollectionChanged += WorksheetsToCopy_CollectionChanged;
        }

        public int MaxRequiredItems { get; set; }

        public long TotalItemsQuantity
        {
            get { return Workbooks.Sum(w => w.MaxRowsInWorkbook); }
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
                if (workInProgress == value) return;
                workInProgress = value;
                OnPropertyChanged();
            }
        }

        public void StartWork()
        {
            EditMode = false;
            WorkInProgress = true;
        }

        public void EndWork()
        {
            EditMode = true;
            WorkInProgress = false;
        }

        public ObservableCollection<WorkbookWithItemsQnt> Workbooks { get; set; }
        public ObservableCollection<string> WorksheetsToCopy { get; set; }

        public bool WorksheetsToCopyAreEmpty
        {
            get { return WorksheetsToCopy.Count == 0 && AllWorksheetsCollection.Count != 0; }
        }

        public bool TotalObjectsQntTooHigh
        {
            get { return TotalItemsQuantity > 50000; }
        }

        public List<string> AllWorksheetsCollection
        {
            get { return Workbooks.SelectMany(w => w.WorksheetsNamesList).Distinct().ToList(); }
        }

        public byte HeadSize { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;

        private void UpdaWorkbooksDepends()
        {
            OnPropertyChanged("AllWorksheetsCollection");

            OnPropertyChanged("WorksheetsToCopyAreEmpty");

            OnPropertyChanged("TotalItemsQuantity");
            OnPropertyChanged("TotalObjectsQntTooHigh");
        }

        private void WorksheetsToCopy_CollectionChanged(object sender,
            NotifyCollectionChangedEventArgs e)
        {
            var coll = sender as ObservableCollection<string>;
            if (coll != null)
                foreach (var workbook in Workbooks)
                    workbook.WorksheetsForCountMaxRows = new List<string>(coll);

            OnPropertyChanged("WorksheetsToCopyAreEmpty");

            OnPropertyChanged("TotalItemsQuantity");
            OnPropertyChanged("TotalObjectsQntTooHigh");
        }

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}