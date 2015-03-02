using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Converter.Template_workbooks;
using TestWSSourceClass;
using Action = System.Action;
using Application = System.Windows.Forms.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace Converter
{
    //todo
    public static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            
            Excel.Worksheet worksheet;
            string[] workbooksList = null;
            Excel.Worksheet targetWorksheet = null;
            bool newExcel = false;
            bool okPressed = false;
            List<string> uniqWorkbooks = new List<string>();

            TemplateWorkbook templateWorkbook = SetTemplateWorkbook();
            if (templateWorkbook == null)
                return;

            //Выбираем фАЙЛ(Ы)
            FileDialog openFileDialog = new OpenFileDialog()
            {
                Multiselect = true,
                Filter = @"Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv",
                Title = @"Выберите файл(ы) для обработки"
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                workbooksList = openFileDialog.FileNames;

            Excel.Application xlApplication = ExcelApp.GetExcelApplication();
            if (xlApplication == null) throw new Exception("Не удалось запустить Excel. Проверьте, установлен ли он на Вашем компьютере");

            if (workbooksList == null || !workbooksList.Any())
            {
                if (!xlApplication.Visible)
                    xlApplication.Quit();
                return;
            }

            Console.WriteLine("Поиск книг с уникальными шапками..");
            var groupedWorkbook = GroupWorkBooksByHead(workbooksList, xlApplication);
            uniqWorkbooks = groupedWorkbook.DistinctBy(x => x.Value).Select(x => x.Key).ToList();
            
            Console.WriteLine("Найдено {0} различных видов книг",uniqWorkbooks.Count);
            Console.WriteLine("Анализ книг...");

            var resultColumns = new List<JustColumn>();
            //Формируем список из уникальных колонок всех источников
            foreach (var s in uniqWorkbooks)
            {
                Process.Start(s);
//                Process process = Marshal.GetActiveObject();
//                xlApplication.WindowState = Excel.XlWindowState.xlMinimized;

                string workbookName = Path.GetFileName(s);
                Console.WriteLine("Обрабатываю книгу \"" + workbookName + "\"");
                Excel.Workbook workbook = xlApplication.Workbooks.Cast<Excel.Workbook>().First(x => x.Name == workbookName);
                if (workbook == null) continue;

                worksheet = workbook.Worksheets[1];
                var sourceWs = new SourceWs(worksheet, xlApplication, templateWorkbook);
                sourceWs.CheckColumns();

                resultColumns.AddRange(sourceWs.SourceColumns.Where(x => !resultColumns.Select(x2 => x2.Description.ToLower()).Contains(x.Description.ToLower())));
                workbook.Close(0);
            }

            var form = new CompareColumnsForm(ref resultColumns, templateWorkbook);
            Console.WriteLine("Загружаю форму сравнения столбцов");
            form.button1.Click += (sender, eventArgs) => okPressed = true;
            form.TopMost = true;
            keybd_event(0x5B, 0, 0, 0);
            keybd_event(0x4D, 0, 0, 0);
            keybd_event(0x5B, 0, 0x2, 0);
            Application.Run(form);
            if (!okPressed) //Если форму просто закрыли
            {
                xlApplication.Visible = true;
                xlApplication.ScreenUpdating = true;
                return;
            }

            ProgressForm progressForm = new ProgressForm();
            progressForm.progressBar1.Maximum = workbooksList.Count();
            progressForm.progressBar1.Step = 1;
            progressForm.progressBar1.Value = 0;
            progressForm.TopMost = true;
            new Thread(progressForm.Show).Start();
            //Цикл по книгам
            foreach (string sourceWorkbookPath in workbooksList)
            {
                Process.Start(sourceWorkbookPath);
//                xlApplication.WindowState  = Excel.XlWindowState.xlMinimized;
                string workbookName = Path.GetFileName(sourceWorkbookPath);
                Console.WriteLine("Копирую книгу \"{0}\"", workbookName);

                Excel.Workbook workbook = xlApplication.Workbooks.Cast<Excel.Workbook>().FirstOrDefault(x => x.Name == workbookName);
                progressForm.UpdateStatus(workbookName);
                if (workbook == null) continue;
                       
                worksheet = workbook.Worksheets[1];
                var sourceWs = new SourceWs(worksheet, xlApplication,templateWorkbook);

                if (targetWorksheet == null)
                {
                    targetWorksheet = xlApplication.Workbooks.Add().Worksheets[1];
                    //Пишем заголовки столбцов
                    foreach (var c in templateWorkbook.TemplateColumns)
                    {
                        ((Excel.Range)(targetWorksheet.Cells[1, c.Index])).Value2 = c.CodeName;
                    }
                }

                sourceWs.FillWorksheet(ref targetWorksheet,resultColumns.Where(x => !String.IsNullOrEmpty(x.CodeName)));
                workbook.Close(false);
            }
            progressForm.Invoke(new Action(progressForm.Close));
            Console.WriteLine("Готово");
            xlApplication.Visible = true;
            xlApplication.ScreenUpdating = true;
            var saveFileDialog = new SaveFileDialog
            {
                FileName = "Упорядоченная выгрузка",
                DefaultExt = "*.xlsx",
                Title = @"Выберите место для сохранения",
                Filter = @"Excel Files|*.xlsx"
            };

            xlApplication.WindowState = Excel.XlWindowState.xlNormal;
            if (saveFileDialog.ShowDialog() != DialogResult.OK) return;
            string path = saveFileDialog.FileName;
            if (targetWorksheet != null) targetWorksheet.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook);
        }


        private static TemplateWorkbook SetTemplateWorkbook()
        {
            List<TemplateWorkbook> templateWorkbooks = new List<TemplateWorkbook>(); //Объект List для передачи его в форму по ссылке. TemplateWorkbook по ссылке почему то не передаётся
            SelectWorkbookTypeForm selectWorkbookTypeForm = new SelectWorkbookTypeForm(ref templateWorkbooks);
            //Форма для выбора тип обрабатываемых книг
            Application.Run(selectWorkbookTypeForm); //не продолжает код до заркытия

            if (templateWorkbooks.Count == 0) return null;
            //selectWorkbookTypeForm.Show(); //Запускает и продолжает код
            return templateWorkbooks[0];
        }

        public static Dictionary<string, int> GroupWorkBooksByHead(IEnumerable<string> workbooksPaths, Excel.Application xlApplication)
        {
//            var excel.XlApplication = GetExcelApplication();
            var resultDictionary = new Dictionary<string, int>();

            //Группировка
            var wsTypes = new List<WSType>();
            var n = 1;
            foreach (var s in workbooksPaths)
            {
                //prepare 
                Excel.Workbook workbook;
                Process.Start(s);
                var wbName = Path.GetFileName(s);
                try
                {
                    workbook = Enumerable.Cast<Excel.Workbook>(xlApplication.Workbooks)
                    .First(x => x.Name == wbName );
                }
                catch (InvalidOperationException e)
                {
                    throw new Exception(String.Format("Не удаётся получить доступ к книге {0}, необходим доступ с возможнстью редактирования \nПпопробуйте скопировать книги к себе на компьютер и объединить их там"));
                }
                

#if DEBUG
                Debug.Assert(workbook != null);
#endif

                Excel.Worksheet worksheet = workbook.Worksheets[1];

//                var head = workbookTable.Columns.Cast<DataColumn>().Select(cl => cl.ColumnName).ToList(); //new List<string>();
                var head = new List<string>();

                var lastUsedColumn = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
                var headRow = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, lastUsedColumn]];

                foreach (Excel.Range cell in headRow)
                    if (!String.IsNullOrEmpty(cell.Value2))
                        head.Add(cell.Value2.ToString());

                //Сравниваем колонки из книги с уже имеющимися колонками
                if (wsTypes.Any(x => x.Heads.SequenceEqual(head)))
                {
                    //находим такую же послежовательность колонок в списке
                    resultDictionary.Add(s, wsTypes.First(x => x.Heads.SequenceEqual(head)).GroupNumber);
                }
                else
                {
                    wsTypes.Add(new WSType { Heads = head, GroupNumber = n });
                    resultDictionary.Add(s, n);
                    n++;
                }
                workbook.Close(false);
            }
            return resultDictionary;
        }

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
    }

}
