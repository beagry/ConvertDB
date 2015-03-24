using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Converter
{
    /// <summary>
    /// Упрощённый список форматов для сохранения Excel книги
    /// </summary>
    public enum XlSaveType
    {
        Xls,
        Xlsx,
        Xlsb
    }

    class WorkbookSaver:IDisposable
    {
        private const string OutFolderName = "Out";
        private readonly string workFolder;
        private readonly Excel.Application app;
        private static readonly List<WBExtentionInfo> Extentions;

        public string SaveFolder { get; set; }
        public bool CreateSaveFolderIfMissing { get; set; }

        private struct WBExtentionInfo
        {
            public string Extention { get; set; }
            public Excel.XlFileFormat SaveFormatNum { get; set; }
            public XlSaveType SimpleSaveType { get; set; }
        }

        static WorkbookSaver()
        {
            Extentions = new List<WBExtentionInfo>
            {
                new WBExtentionInfo(){Extention = ".xls",SaveFormatNum = Excel.XlFileFormat.xlWorkbookNormal},
                new WBExtentionInfo(){Extention = ".xlsx",SaveFormatNum = Excel.XlFileFormat.xlOpenXMLWorkbook},
                new WBExtentionInfo(){Extention = ".xlsb",SaveFormatNum = Excel.XlFileFormat.xlExcel12}
            };
        }

        public WorkbookSaver(string workPath)
        {
            workFolder = workPath;

            SaveFolder = SetSaveFolderWithCreate(workPath);

            app = ExcelApp.GetExcelApplication();
            app.Visible = false;
            app.DisplayAlerts = false;
        }

        public void SaveWorkbookAs(Excel.Workbook wb, XlSaveType saveType)
        {
            var folderToSave = SetSaveFolderWithCreate(wb.Path);
            SaveWorkbookAs(wb,saveType,SaveFolder);
        }

        public void SaveWorkbookAs(Excel.Workbook wb, XlSaveType saveType, string saveFolder)
        {
            if (wb == null) return;
            if (saveFolder == null) throw new ArgumentNullException("saveFolder");

            var extParams = Extentions.Find(e => e.SimpleSaveType == saveType);

            saveFolder = SetSaveFolderWithCreate(wb.Path);
            var fileNameToSave = wb.Name + extParams.Extention;

            app.DisplayAlerts = false;
            wb.SaveAs(saveFolder + "\\" + fileNameToSave,
                extParams.SaveFormatNum,
                ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
            app.DisplayAlerts = true;
        }

        public void ResaveFilesAsXlsx()
        {
            var files = Directory.GetFiles(workFolder, "*.xlsb");
            var ends = files.Count() == 1 ? "a" : files.Count() < 5 ? "и" : "";
            Console.WriteLine(@"Всего найдено {0} книг", files.Count());
            
            foreach (string s in files)
                ResaveWbAsXlsx(s);
        }

        private void ResaveWbAsXlsx(string wbPath)
        {
            try
            {
                var wb = app.Workbooks.Open(wbPath);
                var wbName = Path.GetFileNameWithoutExtension(wbPath);
                try
                {
                    app.DisplayAlerts = false;
                    wb.SaveAs(SaveFolder + wbName, 
                        Excel.XlFileFormat.xlOpenXMLWorkbook,
                        ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);
                    app.DisplayAlerts = true;
                    wb.Close();
                }
                catch (COMException)
                {
                    throw new Exception("Ошибка при сохранении" + wbName);
                }
            }
            catch (COMException)
            {
                Debug.Print("Ошибка при открытии.");
            }
        }

        private static string SetSaveFolderWithCreate(string currentPath)
        {
            if (string.IsNullOrEmpty(currentPath)) return null;
            var s = currentPath + "\\" + OutFolderName + "\\";

            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            return s;
        }

        public void Dispose()
        {
            if (app == null) return;
            app.Quit();
            Marshal.ReleaseComObject(app);
        }
        
    }
}
