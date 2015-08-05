using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Formater.SupportWorksheetsClasses;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace Formater
{
    static class ExcelExtensions
    {
        public static readonly Color BadColor = Color.Crimson;
        public static readonly Color GoodColor = Color.Aquamarine;
        public static readonly Color Clear = Color.Transparent;
    }
}