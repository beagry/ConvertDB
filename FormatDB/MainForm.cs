using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Converter.Template_workbooks;
using ExcelRLibrary;
using Formater.Properties;
using Microsoft.Office.Interop.Excel;
using Button = System.Windows.Forms.Button;
using DataTable = System.Data.DataTable;
using TextBox = System.Windows.Forms.TextBox;

namespace Formater
{
    public sealed partial class MainForm : Form
    {
        private DbToConvert convert;
        private Color badColor = Color.Crimson;
        private Color goodColor = Color.Aquamarine;
        private const string ExcelFilter = @"Excel Files|*.xls;*.xlsx;*.xlsm;*.xlsb";
        private Task task;
        public string WorkbookPath{get { return workbookPathTextBox.Text; }}
        public string OKTMOPath {get { return OKTMOPathTextBox.Text; }}
        public string OKTMOWsName {get { return OKTMOWorksheetCBox.Text; }}
        public string CatalogPath {get { return CatalogPathTextBox.Text; }}
        public string CatalogWsName { get { return CatalogWorksheetCBox.Text; } }
        public string SubjectLinkPath { get { return SubjectSourcePathTextBox.Text; } }
        public string SubjectLinkWsName { get { return SubjectSourceWorksheetCBox.Text; } }
        public string VGTPath { get { return VGTPathTextBox.Text; } }
        public string VGTWsName { get { return VGTWorksheetCBox.Text; } }


        //Bug не реагирует на изменение существущего файла на существующий
        //Bug При Изменении существующего файла в выпадающем списке остаются листы предыдущего файла

        public MainForm()
        {
            
            InitializeComponent();
            tabControl1.TabPages.Remove(tabPage2);
            InitialFormAsync();
        }
        private async void InitialFormAsync()
        {
            //
            //Save Settings textboxes
            //При зменени пути к книге, записываем в настройки новый путь
            OKTMOPathTextBox.TextChanged += (sender, args) =>
            {
                var textBox = sender as TextBox;
                if (textBox != null) Settings.Default.OKTMOWorkbookPath = textBox.Text;
                Settings.Default.Save();
            };

            CatalogPathTextBox.TextChanged += (sender, args) =>
            {
                var textBox = sender as TextBox;
                if (textBox != null) Settings.Default.CatalogWorkbookPath = textBox.Text;
                Settings.Default.Save();
            };

            SubjectSourcePathTextBox.TextChanged += (sender, args) =>
            {
                var textBox = sender as TextBox;
                if (textBox != null) Settings.Default.SourceListBySubjectWorkbookPath = textBox.Text;
                Settings.Default.Save();
            };

            VGTPathTextBox.TextChanged += (sender, args) =>
            {
                var textBox = sender as TextBox;
                if (textBox != null) Settings.Default.VGTWorkbookPath = textBox.Text;
                Settings.Default.Save();
            };

            //
            //Save Settings CBOXES
            OKTMOWorksheetCBox.TextChanged += (sender, args) => 
            {
                var cbox = sender as ComboBox;
                if (cbox != null) Settings.Default.OKTMOWorksheetName = cbox.Text;
                Settings.Default.Save();
            };

            CatalogWorksheetCBox.TextChanged += (sender, args) =>
            {
                var cbox = sender as ComboBox;
                if (cbox != null) Settings.Default.CatalogWorksheetName = cbox.Text;
                Settings.Default.Save();
            };

            SubjectSourceWorksheetCBox.TextChanged += (sender, args) =>
            {
                var cbox = sender as ComboBox;
                if (cbox != null) Settings.Default.SubjectSourceWorksheetName = cbox.Text;
                Settings.Default.Save();
            };

            VGTWorksheetCBox.TextChanged += (sender, args) =>
            {
                var cbox = sender as ComboBox;
                if (cbox != null) Settings.Default.VGTWorksheetName = cbox.Text;
                Settings.Default.Save();
            };

            //
            //Set Events ColorChange
            OKTMOPathTextBox.BackColorChanged += TextBox_BackColorChanged;
            CatalogPathTextBox.BackColorChanged += TextBox_BackColorChanged;
            SubjectSourcePathTextBox.BackColorChanged += TextBox_BackColorChanged;
            VGTPathTextBox.BackColorChanged += TextBox_BackColorChanged;

            //
            //Set Events ChePath
            OKTMOPathTextBox.TextChanged += CheckPath;
            CatalogPathTextBox.TextChanged += CheckPath;
            SubjectSourcePathTextBox.TextChanged += CheckPath;
            VGTPathTextBox.TextChanged += CheckPath;
            workbookPathTextBox.TextChanged += CheckPath;

            //Mouse Hover
            OpenButton.MouseEnter += OnMouseEnter;
            OpenButton.MouseLeave += OnMouseLeave;
            

            await Task.Run(() =>
            {
                //
                //Set Paths
                OKTMOPathTextBox.Text = Settings.Default.OKTMOWorkbookPath ?? String.Empty;
                CatalogPathTextBox.Text = Settings.Default.CatalogWorkbookPath ?? String.Empty;
                SubjectSourcePathTextBox.Text = Settings.Default.SourceListBySubjectWorkbookPath ?? String.Empty;
                VGTPathTextBox.Text = Settings.Default.VGTWorkbookPath ?? String.Empty;

                //Set worksheets
                OKTMOWorksheetCBox.Text = Settings.Default.OKTMOWorksheetName ?? String.Empty;
                CatalogWorksheetCBox.Text = Settings.Default.CatalogWorksheetName ?? String.Empty;
                SubjectSourceWorksheetCBox.Text = Settings.Default.SubjectSourceWorksheetName ?? String.Empty;
                VGTWorksheetCBox.Text = Settings.Default.VGTWorksheetName ?? String.Empty;

                //
                //Check Paths
                foreach (
                    var textBox in
                        tableLayoutPanel1.Controls.OfType<TextBox>())
                {
                    CheckPath(textBox, new EventArgs());
                }

                //
                foreach (var cbox in tableLayoutPanel1.Controls.OfType<ComboBox>())
                {
                    //Check
                    CheckCbox(cbox);

                    //add events
                    cbox.TextChanged += CBoxOnIndexChange;
                }
            });
            
            tabControl1.TabPages.Insert(1, tabPage2);
            StartButton.Enabled = true;

            //MouseHover
            StartButton.MouseEnter += OnMouseEnter;
            StartButton.MouseLeave += OnMouseLeave;
        }

        private void OnMouseLeave(object sender, EventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;
            btn.BackColor = Color.SlateGray;
        }


        private void OnMouseEnter(object sender, EventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;
            btn.BackColor = Color.RoyalBlue;
        }

        //
        //Buttons Clicks
        //
        private async void StartButton_Click(object sender, EventArgs e)
        {
            convert = new DbToConvert(this, XlTemplateWorkbookType.LandProperty) { ColumnsToReserve = new List<string> { "SUBJECT", "REGION", "NEAR_CITY", "SYSTEM_GAS", "SYSTEM_WATER", "SYSTEM_SEWERAGE", "SYSTEM_ELECTRICITY" } };
            var button = sender as Button;
            if (button == null) return;

#if !DEBUG
            if (FormHasInvalidControl())
            {
                MessageBox.Show(
                    @"Не все поля заполнены",
                    @"Операция прервана", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
#endif

            tabPage1.Enabled = false;
            tabPage2.Enabled = false;

            WarningLabel.Visible = true;
            if (!convert.ColumnHeadIsOk()) return;

            //Запусть обработки в новом потоке
            await Task.Run(() => convert.FormatWorksheet());

            convert.ExcelPackage.SaveWithDialog();
            WarningLabel.Visible = false;
            tabPage1.Enabled = true;
            tabPage1.Enabled = true;

        }

        private void OpenButton_Click(object sender, EventArgs e)
        {
            if (progressBar.Value > 0) progressBar.Value = 0;

            using (var fd = new OpenFileDialog())
            {
                fd.Filter = ExcelFilter;
                fd.Title = @"Выберите рабочий файл";

                if (fd.ShowDialog() == DialogResult.OK)
                {
                    workbookPathTextBox.Text = fd.FileName;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ChangeWorkbookPath(CatalogPathTextBox);
        }

        private void SelectOKTMOFileButton_Click(object sender, EventArgs e)
        {
            ChangeWorkbookPath(OKTMOPathTextBox);
        }

        private void SelectSubjectSourceFileButton_Click(object sender, EventArgs e)
        {
            ChangeWorkbookPath(SubjectSourcePathTextBox);
        }

        private void SelectVGTFileButton_Click(object sender, EventArgs e)
        {
            ChangeWorkbookPath(VGTPathTextBox);
        }


        //
        //Additional Methods
        //
        private void CBoxOnIndexChange(object sender, EventArgs e)
        {
            CheckCbox(sender as ComboBox);
        }

        private void CheckCbox(ComboBox box)
        {
            if (box == null) return;
            box.BackColor = badColor;
            if (box.SelectedIndex != -1)
                box.BackColor = goodColor;
        }

        private void CheckPath(object sender, EventArgs e)
        {
            var textBox = sender as TextBox;
            if (textBox != null)
            {
                textBox.BackColor = badColor;
                if (File.Exists(textBox.Text))
                    textBox.BackColor =   goodColor ;
            }
        }

        private static void ChangeWorkbookPath(TextBox textBox)
        {
            using (var fileDialog = new OpenFileDialog())
            {
                fileDialog.Filter = ExcelFilter;
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox.Text = string.Empty;
                    textBox.Text = fileDialog.FileName;
                }
            }
        }

        /// <summary>
        /// Метод заполняет выпадающий список из книги, путь которой указан в вызывающем объекте
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="eventArgs"></param>
        private void TextBox_BackColorChanged(object sender, EventArgs eventArgs)
        {

            TextBox textBox = sender as TextBox;
            if (textBox == null) return;
            if (!File.Exists(textBox.Text)) return;

            //Открываем книгу
            var path = textBox.Text;
            var reader = new ExcelReader();
            var wsNames = reader.GetWorksheetsNames(path);

            //Выявляем нужный выпадающий список
            var cbox = (ComboBox) tableLayoutPanel1.GetControlFromPosition(3,
                tableLayoutPanel1.GetPositionFromControl(textBox).Row);

            var backUpValue = cbox.Text;
            cbox.Items.Clear(); //Очищаем
            //Заполняем выпадающий список
            if (wsNames != null)
            {
                cbox.Items.AddRange(wsNames.ToArray());
            }

            cbox.Text = backUpValue;
            CheckCbox(cbox); //Проверяем текущее значение выпадающего списка
        }

        private bool FormHasInvalidControl()
        {
            //False если хоть один контроллер с текстовым полен не верено заполнен
            return tableLayoutPanel1.Controls.Cast<Control>()
                .Where(control => control is TextBox || control is ComboBox)
                .Any(control => control.BackColor == badColor) || workbookPathTextBox.BackColor != goodColor;
        }

    }
}
