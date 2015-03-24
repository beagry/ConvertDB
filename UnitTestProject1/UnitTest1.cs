using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using Converter;
using Converter.Template_workbooks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UI;

namespace UnitTestProject1
{
    [TestClass]
    public class WindowsTests
    {
        [TestMethod]
        public void CompareWindow()
        {
            var wb = new LandPropertyTemplateWorkbook();

            var binded = new Dictionary<string, ObservableCollection<string>>()
            {
                {"SUBJECT",new ObservableCollection<string>(){"COLUMN1","COLUMN2","ADDITIONAL_1","ONE_MORE"}},
                {"REGION",new ObservableCollection<string>(){"REGION1","SOME_OTHER_REGION"}},
                {"NEAR_CITY",new ObservableCollection<string>(){"CITY","SUPER_CITY"}},
                {"DESCTIPTION",new ObservableCollection<string>(){"MY_BLA_BLA_BLA","TELL_SOME_PURE"}},
                {"COMMENTS",new ObservableCollection<string>()}
            };

            var allColumns = new List<WorksheetInfo>()
            {
                new WorksheetInfo("WS1",new List<WorksheetColumnInfo>(){new WorksheetColumnInfo(1,"COLUMN1"),new WorksheetColumnInfo(2,"COLUMN2"),new WorksheetColumnInfo(3,"ADDITIONAL_1")}),
                new WorksheetInfo("WS2",new List<WorksheetColumnInfo>(){new WorksheetColumnInfo(1,"REGION1"),new WorksheetColumnInfo(2,"SOME_OTHER_REGION"){ValuesExamples = new List<string>{"регион1","Несустветная фигня", "Лорем инспун"}},new WorksheetColumnInfo(3,"ANOTHER_ADDITIONAL"){ValuesExamples = new List<string>{"Когда мы быыылии молодыыыыми","Несустветная фигня", "Xgkdfl;s"}}}),
                new WorksheetInfo("WS3",new List<WorksheetColumnInfo>(){new WorksheetColumnInfo(1,"CITY"),new WorksheetColumnInfo(2,"SUPER_CITY"){ValuesExamples = new List<string>{"Дом","Не дом", "Полудом"}},new WorksheetColumnInfo(3,"ONE_MORE")}),
                new WorksheetInfo("WS4",new List<WorksheetColumnInfo>(){new WorksheetColumnInfo(1,"MY_BLA_BLA_BLA"),new WorksheetColumnInfo(2,"TELL_SOME_PURE"),new WorksheetColumnInfo(3,"AND_MORE_ONE_COLUMN")}),
            };

            var w = new ColumnsCompareWindow(binded,allColumns);
            w.ShowDialog();
        }
    }
}
