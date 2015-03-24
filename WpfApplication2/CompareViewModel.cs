using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Converter;
using UI.Annotations;

namespace WpfApplication2
{
    public sealed class CompareViewModel:INotifyPropertyChanged
    {
        private ICollection<WorksheetInfo> worksheets;

//        private ObservableCollection<string> unbindedColumns;

        private Dictionary<string, ObservableCollection<string>> bindedColumnsDictionary;

        public CompareViewModel(Dictionary<string, ObservableCollection<string>> bindedColumns, ICollection<WorksheetInfo> worksheetsSamples)
        {
            worksheets = worksheetsSamples;

            bindedColumnsDictionary = bindedColumns;

            UnbindedColumns = new ObservableCollection<string>(
                worksheets.SelectMany(w => w.Columns) //Единый список колонок
                    .Select(c => c.Name) //Взять их имена
                    .Distinct()
                    .Except(bindedColumns.SelectMany(kp => kp.Value)) //исключить уже выбранные 
                    .ToList());
        }
        public CompareViewModel()
        {
            worksheets = new List<WorksheetInfo>();
            UnbindedColumns = new ObservableCollection<string>();
            bindedColumnsDictionary = new Dictionary<string, ObservableCollection<string>>();
        }


        public Dictionary<string, ObservableCollection<string>> BindedColumnsDictionary
        {
            get { return bindedColumnsDictionary; }
            set
            {
                if (Equals(value, bindedColumnsDictionary)) return;
                bindedColumnsDictionary = value;
                OnPropertyChanged("BindedColumnsDictionary");
            }
        }

        public ObservableCollection<string> UnbindedColumns { get; set; }

        public IEnumerable<string> GetColumnValuesExamples(string columnName, byte examplesQnt = 10)
        {
            return
                worksheets.Where(
                    w => w.Columns.Any(c => string.Equals(c.Name, columnName, StringComparison.OrdinalIgnoreCase)))
                    .SelectMany(w => w.Columns)
                    .SelectMany(c => c.ValuesExamples)
                    .OrderBy(s => Guid.NewGuid())
                    .Take(examplesQnt);
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