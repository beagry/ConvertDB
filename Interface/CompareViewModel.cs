using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Converter;
using Converter.Models;
using Telerik.Windows.Controls;
using UI.Annotations;

namespace UI
{
    public sealed class CompareViewModel:INotifyPropertyChanged
    {
        private readonly ICollection<WorksheetInfo> worksheets;
        private Dictionary<string, ObservableCollection<string>> bindedColumnsDictionary;
        private string lastSelectedItem;



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

        public string LastSelectedItem
        {
            get { return lastSelectedItem; }
            set
            {
                if (Equals(value, lastSelectedItem)) return;
                lastSelectedItem = value;
                OnPropertyChanged();
            }
        }

        public IEnumerable<string> LastSelectedColumnValuesExamples
        {
            get
            {
                return GetColumnValuesExamples(LastSelectedItem);
            }
        }



        public CompareViewModel(Dictionary<string, ObservableCollection<string>> bindedColumns,
            ICollection<WorksheetInfo> worksheetsSamples):this()
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

        private CompareViewModel()
        {
            StyleManager.ApplicationTheme = new ModernTheme();
            worksheets = new List<WorksheetInfo>();
            UnbindedColumns = new ObservableCollection<string>();
            bindedColumnsDictionary = new Dictionary<string, ObservableCollection<string>>();
        }

        private IEnumerable<string> GetColumnValuesExamples(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) return new List<string>();

            var suitWorksheets =
                worksheets.Where(
                    w =>
                        w.Columns.Any(
                            c => string.Equals(c.Name, LastSelectedItem, StringComparison.OrdinalIgnoreCase) && c.ValuesExamples != null)).ToList();

            if (!suitWorksheets.Any()) return new List<string>();

            return suitWorksheets.SelectMany(w => w.Columns)
                    .Where(c => string.Equals(c.Name, LastSelectedItem, StringComparison.OrdinalIgnoreCase))
                    .SelectMany(c => c.ValuesExamples)
                    .OrderBy(s => Guid.NewGuid())
                    .ToList();
        }

        #region inotifyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        public void UpdateValuesExamples()
        {
            OnPropertyChanged("LastSelectedColumnValuesExamples");
        }
    }
}