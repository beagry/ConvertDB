using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Converter.Models;
using Converter.Properties;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Microsoft.Win32;

namespace UI
{
    public sealed class BooksToConvertViewModel : INotifyPropertyChanged
    {
        private bool editMode;
        private string status;
        private XlTemplateWorkbookType workbooksType;
        private bool workInProgress;
        private string mainBasePath;

        public BooksToConvertViewModel()
        {
            EditMode = true;
            WorkbooksType = XlTemplateWorkbookType.LandProperty;
            Workbooks = new ObservableCollection<SelectWorkbookViewModel>();
            Status = "Готово к работе";
            Workbooks.CollectionChanged += OnWorkbookAdd;
        }

        private void OnWorkbookAdd(object sender, NotifyCollectionChangedEventArgs e)
        {
            var newWbs = e.NewItems;
            if (newWbs == null || newWbs.Count == 0) return;
            foreach (var newWb in newWbs.Cast<SelectWorkbookViewModel>())
            {
                newWb.WorkbookChecked += GroupWb;
            }
        }

        private readonly ConcurrentDictionary<int,List<ColumnInfo>> workbooksGourps = new ConcurrentDictionary<int, List<ColumnInfo>>();
        private async void GroupWb(object sender, EventArgs e)
        {
            var wb = (SelectWorkbookViewModel) sender;
            var groupNum = 0; 
            await Task.Run(() =>
            {
                var columns = wb.Columns.OrderBy(c => c.Index).ToList();
                                
                foreach (
                    var pair in
                        workbooksGourps.Where(pair => columns.SequenceEqual(pair.Value, new ColumnInfoComparer())))
                {
                    groupNum = pair.Key;
                }
                if (groupNum != 0) return;

                groupNum = workbooksGourps.Any() ? workbooksGourps.Keys.Max() + 1 : 1;
                workbooksGourps.GetOrAdd(groupNum, columns);
            });
            wb.GroupNum = groupNum;
        }

        public bool UseMainBase { get; set; }

        public string MainBasePath
        {
            get { return mainBasePath; }
            set
            {
                if (mainBasePath == value)
                    return;
                mainBasePath = value;
                OnPropertyChanged();
            }
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

        public string Status
        {
            get { return status; }
            set
            {
                if (status == value) return;
                status = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<SelectWorkbookViewModel> Workbooks { get; set; }

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

        public void SelectMainBasePath()
        {
            var fd = new OpenFileDialog();
            if (fd.ShowDialog() == true)
            {
                MainBasePath = fd.FileName;
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

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    internal class ColumnInfoComparer : IEqualityComparer<ColumnInfo>
    {
        public bool Equals(ColumnInfo x, ColumnInfo y)
        {
            if (object.ReferenceEquals(x, y)) return true;

            return (x.Index.Equals(y.Index) && x.Name.Equals(y.Name));
        }



        public int GetHashCode(ColumnInfo obj)
        {
            var nameHash = obj.Name == null ? 0 : obj.Name.GetHashCode();
            var indexHash = obj.Index.GetHashCode();

            return nameHash ^ indexHash;
        }
    }


    public class SelectWorkbookViewModel : SelectedWorkbook, INotifyPropertyChanged
    {
        public static int SamplesCount = 50;
        private int groupNum;


        public SelectWorkbookViewModel(string s): base(s)
        {
            groupNum = 0;
            Columns = new ObservableCollection<ColumnInfo>();
            CheckWorkbook();
        }

        public int GroupNum
        {
            get { return groupNum; }
            set
            {
                if (value == groupNum) return;
                groupNum = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<ColumnInfo> Columns { get; set; }

        public async void CheckWorkbook()
        {
            await CheckWorkbookAsync();
            OnWorkbookChecked();
        }

        private Task CheckWorkbookAsync()
        {
            return Task.Run(() =>
            {
                if (!File.Exists(Path)) return;
                var reader = new ExcelReader();
                var ds = reader.GetWsStructs(Path);
                if (ds == null) return;
                var dt = ds.Tables.Cast<DataTable>().First();
                Columns =
                    new ObservableCollection<ColumnInfo>(
                        dt.Columns.Cast<DataColumn>()
                            .Select(cl => new ColumnInfo(dt.Columns.IndexOf(cl) + 1, cl.ColumnName))
                            .ToList());
            });
        }

        public event EventHandler WorkbookChecked;

        private void OnWorkbookChecked()
        {
            if (WorkbookChecked != null) WorkbookChecked(this, new EventArgs());
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [Annotations.NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}