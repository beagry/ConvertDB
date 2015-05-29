using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using Converter;
using Converter.Models;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using ExcelRLibrary;
using ExcelRLibrary.TemplateWorkbooks;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using Telerik.Windows.Controls;
using UI.Annotations;
using LandPropertyTemplateWorkbook = Converter.Template_workbooks.LandPropertyTemplateWorkbook;

namespace UI
{
    public sealed class CompareViewModel:INotifyPropertyChanged
    {
        private readonly ICollection<WorksheetInfo> worksheets;
        private string lastSelectedItem;
        private TemplateWbsRespository repository;
        private XlTemplateWorkbookType wbType;

        public ObservableCollection<JustColumnViewModel> BindedColumns { get; set; }

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



        public CompareViewModel(Dictionary<JustColumn, ObservableCollection<string>> bindedColumns,
            ICollection<WorksheetInfo> worksheetsSamples, XlTemplateWorkbookType wbType):this()
        {
            this.wbType = wbType;
            worksheets = worksheetsSamples;

            BindedColumns = new ObservableCollection<JustColumnViewModel>();
            BindedColumns.CollectionChanged += OnBindedCollectionChange;
            bindedColumns.Select(c => new JustColumnViewModel(c.Key) {SuitedColumns = c.Value}).ForEach(jc =>BindedColumns.Add(jc));

            UnbindedColumns = new ObservableCollection<string>(
                worksheets.SelectMany(w => w.Columns) //Единый список колонок
                    .Select(c => c.Name) //Взять их имена
                    .Distinct()
                    .Except(bindedColumns.SelectMany(kp => kp.Value)) //исключить уже выбранные 
                    .ToList());

            UpdateListsFromDb();
        }

        private CompareViewModel()
        {
            StyleManager.ApplicationTheme = new ModernTheme();
            worksheets = new List<WorksheetInfo>();
            UnbindedColumns = new ObservableCollection<string>();

            repository =  UnitOfWorkSingleton.UnitOfWork.TemplateWbsRespository;
        }

        private void OnBindedCollectionChange(object sender, NotifyCollectionChangedEventArgs args)
        {
            if (args.NewItems != null)
            {
                foreach (var newItem in args.NewItems.Cast<JustColumnViewModel>())
                {
                    var item = newItem;
                    newItem.SuitedColumns.CollectionChanged += (o, eventArgs) => OnColumnNameCollectionChange(item, eventArgs);
                }
            }
        }

        private void OnColumnNameCollectionChange(JustColumnViewModel owner, NotifyCollectionChangedEventArgs args)
        {
            if (args.NewItems != null)
            {
                foreach (var newBindedColumnName in args.NewItems.Cast<string>())
                {
                    repository.AddColumnColumnPair(owner.CodeName,newBindedColumnName);
                }
                repository.Save();
            }

            if (args.OldItems != null)
            {
                foreach (var removeBindedColumn in args.OldItems.Cast<string>())
                {
                    repository.RemoveColumnColumnPair(owner.CodeName,removeBindedColumn);
                }
                repository.Save();
            }
        }


        public void AddNewcolumn(string columnName)
        {
            var newColumnIndex = BindedColumns.Max(j => j.Index) + 1;
            BindedColumns.Add(new JustColumnViewModel(columnName,columnName,newColumnIndex));
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


        #region INotifyChanged
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

        public void CombineWorkbooks()
        {
            var dict = BindedColumns.ToDictionary(k => k.CodeName, v => v.SuitedColumns.ToList());

            var typifer = new WorkbookTypifier<LandPropertyTemplateWorkbook>()
            {
                RulesDictionary = dict,
                WorkbooksPaths = worksheets.Select(w => w.Workbook.Path).Distinct().ToList()
            };

            var result = typifer.CombineToSingleWorkbook();
            if (result == null) return;

            result.SaveWithDialog("Обработанная выгрузка");
        }

        private void UpdateListsFromDb()
        {
            //Нам нужны попытаться найти неразнесённые колонки
            //В базе

        }

        private void SaveBindedColumnToDb()
        {


        }


        public static Dictionary<JustColumn, ObservableCollection<string>> DitctToObservDict(
            Dictionary<JustColumn, List<string>> sourceDict)
        {
            return sourceDict.ToDictionary(k => k.Key, v => new ObservableCollection<string>(v.Value));
        }

        public static Dictionary<JustColumn, List<string>> ObservDictToDict(
            Dictionary<JustColumn, ObservableCollection<string>> sourceDict)
        {
            return sourceDict.ToDictionary(k => k.Key, v => v.Value.ToList());
        }
    }

    public class JustColumnViewModel:JustColumn
    {
        public JustColumnViewModel(JustColumn column):base(column.CodeName,column.Description,column.Index)
        {
            SuitedColumns = new ObservableCollection<string>();
        }

        public JustColumnViewModel(string codename, string description, int index) : base(codename, description, index)
        {
            SuitedColumns = new ObservableCollection<string>();
        }

        public ObservableCollection<string> SuitedColumns { get; set; }
    }
}