﻿using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using Converter;
using Converter.Models;
using Converter.Template_workbooks;
using Converter.Template_workbooks.EFModels;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using UI;
using TemplateWorkbook = Converter.Template_workbooks.EFModels.TemplateWorkbook;

namespace UnitTestProject1
{
    [TestClass]
    public class WindowsTests
    {
        [TestMethod]
        public void CompareWindow()
        {
            var wb = new LandPropertyTemplateWorkbook();

            var binded = new Dictionary<string, List<string>>()
            {
                {"SUBJECT",new List<string>(){"COLUMN1","COLUMN2","ADDITIONAL_1","ONE_MORE"}},
                {"REGION",new List<string>(){"REGION1","SOME_OTHER_REGION"}},
                {"NEAR_CITY",new List<string>(){"CITY","SUPER_CITY"}},
                {"DESCTIPTION",new List<string>(){"MY_BLA_BLA_BLA","TELL_SOME_PURE"}},
                {"COMMENTS",new List<string>()}
            };

            var allColumns = new List<WorksheetInfo>()
            {
                new WorksheetInfo("WS1",new List<ColumnInfo>(){new ColumnInfo(1,"COLUMN1"),new ColumnInfo(2,"COLUMN2"),new ColumnInfo(3,"ADDITIONAL_1")}),
                new WorksheetInfo("WS2",new List<ColumnInfo>(){new ColumnInfo(1,"REGION1"),new ColumnInfo(2,"SOME_OTHER_REGION"){ValuesExamples = new List<string>{"регион1","Несустветная фигня", "Лорем инспун"}},new ColumnInfo(3,"ANOTHER_ADDITIONAL"){ValuesExamples = new List<string>{"Когда мы быыылии молодыыыыми","Несустветная фигня", "Xgkdfl;s"}}}),
                new WorksheetInfo("WS3",new List<ColumnInfo>(){new ColumnInfo(1,"CITY"),new ColumnInfo(2,"SUPER_CITY"){ValuesExamples = new List<string>{"Дом","Не дом", "Полудом"}},new ColumnInfo(3,"ONE_MORE")}),
                new WorksheetInfo("WS4",new List<ColumnInfo>(){new ColumnInfo(1,"MY_BLA_BLA_BLA"),new ColumnInfo(2,"TELL_SOME_PURE"),new ColumnInfo(3,"AND_MORE_ONE_COLUMN")}),
            };

            var w = new ColumnsCompareWindow(binded,allColumns);
            w.ShowDialog();
        }

    }


    [TestClass]
    public class DbTests
    {

        [TestMethod]
        public void TryToCreateWbsDataBase()
        {
            var db = new TemplateWbsContext();

            var books = db.TemplateWorkbooks;

            Assert.IsTrue(books.First().Columns.Count == 60, "В первой книге слишком мало колонок");
        }
    }
}
