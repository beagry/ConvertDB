using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Converter.Template_workbooks;
using Microsoft.Office.Interop.Excel;
using Font = System.Drawing.Font;
using Label = System.Windows.Forms.Label;
using ListBox = System.Windows.Forms.ListBox;
using Point = System.Drawing.Point;

//TODO: Проверить выпадающие списки на предмет урезания вариантов
//TODO: Проверить выгрузку результатов по нажатию клавиши "ОК"

namespace Converter
{
    public partial class CompareColumnsForm : Form
    {
        private SourceWs sourceWs;
        private Worksheet targetWorksheet;
        private string[] comboboxItems;
        private List<JustColumn> resultColumns;
        private const int columnWidth = 200;
        private Dictionary<int, int> resultDictionary;
        private ComboBox lastUsedComboBox;
        private bool okPressed ;
        private TemplateWorkbook templateWorkbook;
        public CompareColumnsForm(ref List<JustColumn> resultColumns,TemplateWorkbook workbook)
        {
            InitializeComponent();
            //resultDictionary.Add(1,1);
            this.resultColumns = resultColumns;
            comboboxItems = resultColumns.Select(x => x.Description).ToArray();
            this.okPressed = okPressed;
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.UserPaint |
                ControlStyles.DoubleBuffer,
                true);
            templateWorkbook = workbook;

            //Create Column
            int i = 1;
            foreach (JustColumn item in templateWorkbook.TemplateColumns)
            {
                var row = new RowStyle(SizeType.AutoSize);
                tableLayoutPanel1.RowStyles.Add(row);
                CreateNewLabel(item.Description,item.CodeName, i, tableLayoutPanel1.RowStyles.Count - 1);
                CreateNewComboBox(item.CodeName,i, 1);
                i++;
            }

            //Берём все столбцы, что смогли идентифицироваться
            foreach (JustColumn column in resultColumns.Where(x => !String.IsNullOrEmpty(x.CodeName)))
            {
                //Находим строку
                foreach (Control label in tableLayoutPanel1.Controls.Cast<Control>().Where(x=> x is Label))
                {
                    if ((string) label.Tag == "Descr" ) continue;
                    if ((string) label.Tag != column.CodeName) continue;
                    int i2 = 0;
                    int row = tableLayoutPanel1.GetRow(label);

                    do{
                        i2++;
                        if (tableLayoutPanel1.ColumnCount - 1 < i2) AddColumn();

                    } while (!String.IsNullOrEmpty(tableLayoutPanel1.GetControlFromPosition(i2, row).Text));

                    tableLayoutPanel1.GetControlFromPosition(i2, row).Text =column.Description;
                }
            }
            ChangeListBox(this, new EventArgs());

            Shown += (obj, arg) =>
            {
                TopLevel = TopMost;
//                button1.Select();
            };
        }

        private void CreateNewLabel(string text,string tag, int index,int row)
        {
            var label = new Label
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204))),
                Location = new Point(3, 0),
                Margin = new Padding(5),
                Name = "label" + index,
                Size = new Size(276, 26),
                TabIndex = 0,
                Text = text,
                Tag = tag
            };
            tableLayoutPanel1.Controls.Add(label, 0, row);    
        }

        private void CreateNewComboBox(string tag, int row, int column)
        {
            var comboBox = new ComboBox
            {
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                AutoSize = true,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,    
                Font =
                    new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular,
                        GraphicsUnit.Point, ((byte) (204))),
                FormattingEnabled = true,
                Location = new Point(278, 3),
                Name = "comboBox" + row,
                Size = new Size(columnWidth, 24),
                TabIndex = 1,
                Tag =  tag
            };
            comboBox.TextChanged += ChangeListBox;
            comboBox.Click += (sender, args) => lastUsedComboBox = sender as ComboBox;
            comboBox.Click += UnSuedColumnListBox_SelectedIndexChanged;
            comboBox.MouseWheel += richTextBox_MouseWheel;
            comboBox.MouseWheel += ToolStripComboBox_MouseWheel;
            comboBox.Items.AddRange(comboboxItems);
            tableLayoutPanel1.Controls.Add(comboBox, column, row);
            return;
        }

        private void ToolStripComboBox_MouseWheel(object o, MouseEventArgs e)
        {
            //Cast the MouseEventArgs to HandledMouseEventArgs
            HandledMouseEventArgs mwe = (HandledMouseEventArgs)e;

            //Indicate that this event was handled
            //(prevents the event from being sent to its parent control)
            mwe.Handled = true;
        }

        void richTextBox_MouseWheel(object sender, MouseEventArgs e)
        {
            //if (this.Parent == null || (this.Parent.GetType() != typeof (TableLayoutPanel))) return;
            TableLayoutPanel parentPanel = tableLayoutPanel1; //(TableLayoutPanel)this.Parent;
            if (e.Delta == 0) return;
            int newVerticalScrollValue = parentPanel.VerticalScroll.Value - e.Delta;
            if (newVerticalScrollValue > parentPanel.VerticalScroll.Maximum)
                newVerticalScrollValue = parentPanel.VerticalScroll.Maximum;

            if (newVerticalScrollValue < parentPanel.VerticalScroll.Minimum)
                newVerticalScrollValue = parentPanel.VerticalScroll.Minimum;

            parentPanel.VerticalScroll.Enabled = true;
            parentPanel.VerticalScroll.Value = newVerticalScrollValue;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Лучше ничего не трогать до окончания операции","",MessageBoxButtons.OK,MessageBoxIcon.Information);
            //Сохраняем разнсенные столбцы
            foreach (ComboBox comboBox in tableLayoutPanel1.Controls.Cast<Control>()
                .Where(x => x is ComboBox && !String.IsNullOrEmpty(x.Text)).Cast<ComboBox>())
            {
                resultColumns.First(x => x.Description == comboBox.Text).CodeName = (string) comboBox.Tag;
            }

            //Сохраняем неразнесенные столбцы
            foreach (string unusedColumnName in this.UnUsedColumnListBox.Items.Cast<string>())
            {
                resultColumns.First(x => x.Description == unusedColumnName).CodeName = templateWorkbook.UnUsedColumnCode;
            }


            //Dictionary<int, int> resultDictionary = new Dictionary<int, int>();
            //Dictionary<string, string> dictionary = tableLayoutPanel1.Controls.Cast<Control>()
            //    .Where(x => x is ComboBox && !String.IsNullOrEmpty(x.Text)).Cast<ComboBox>()
            //    .ToDictionary(x => x.Text, x => x.Tag.ToString());
            
            //foreach (KeyValuePair<string, string> keyValuePair in dictionary)
            //{
            //    JustColumn srColumn = sourceColumns.First(x => x.Name == keyValuePair.Key);
            //    JustColumn targetColumn = LandPropertyTemplateWorkbook.TemplateColumns.First(x => x.Code == keyValuePair.Value);

            //    resultDictionary.Add(srColumn.Index + 1, targetColumn.Index);
            //}
            //this.resultDictionary = resultDictionary;
            Close();
        }

        private void addColumnButton_Click(object sender, EventArgs e)
        {
            AddColumn();
        }

        private void AddColumn()
        {
            this.Enabled = false;
            tableLayoutPanel1.ColumnCount ++;
            tableLayoutPanel1.ColumnStyles.Insert(tableLayoutPanel1.ColumnCount - 1,new ColumnStyle(SizeType.Absolute,columnWidth));
            //tableLayoutPanel1.Size = new Size(tableLayoutPanel1.Size.Width + columnWidth, tableLayoutPanel1.Size.Height);
            Size = new Size(Size.Width + columnWidth, Size.Height);
            tableLayoutPanel1.SetColumnSpan(SourceDescrLabel, tableLayoutPanel1.ColumnCount - 1);

            int i = 1;
            foreach (JustColumn item in templateWorkbook.TemplateColumns)
            {
                //insetr to last column
                CreateNewComboBox(item.CodeName, i, tableLayoutPanel1.ColumnCount - 1);
                i++;
            }
            this.Enabled = true;
        }

        private void CutComboboxItems(object o, EventArgs e)
        {
            List<string> usedItems =
                tableLayoutPanel1.Controls.Cast<Control>()
                    .Where(x => x is ComboBox && !String.IsNullOrEmpty(x.Text))
                    .Select(x => x.Text).Distinct()
                    .ToList();
            foreach (ComboBox combobox in tableLayoutPanel1.Controls.Cast<Control>().Where(x=> x is ComboBox).Cast<ComboBox>())
            {
                combobox.Items.Clear();
                var newList = resultColumns.Select(x => x.Description).ToList().Except(usedItems).ToArray();
                combobox.Items.AddRange(newList);
                if (!String.IsNullOrEmpty(combobox.Text))
                {
                    combobox.Items.Add(combobox.Text);
                }
            }
        }

        private void ChangeListBox(object o, EventArgs eventArgs)
        {
            List<string> usedItems =
                tableLayoutPanel1.Controls.Cast<Control>()
                    .Where(x => x is ComboBox && !String.IsNullOrEmpty(x.Text))
                    .Select(x => x.Text).Distinct()
                    .ToList();
            UnUsedColumnListBox.Items.Clear();
            var newList = resultColumns.Select(x => x.Description).ToList().Except(usedItems).ToArray();
            UnUsedColumnListBox.Items.AddRange(newList);


            //Изменение выпадающих списков
//            foreach (Control cbox in tableLayoutPanel1.Controls)
//            {
//                if (!(cbox is ComboBox)) continue;
//                var box = cbox as ComboBox;
//                box.Items.Clear();
//                box.Items.AddRange(newList);
//                box.Items.Add(cbox.Text);  
//            } 
        }

        private void UnSuedColumnListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            const int examplesQuantity =15;
            if (sender is ListBox)
            {
                if (UnUsedColumnListBox.SelectedItem == null) return;
            }
            else
            {
                Debug.Assert(sender is ComboBox);
                if ((sender as ComboBox).SelectedIndex == -1) return;
            }
            
            string selectedColumName = sender is ListBox? UnUsedColumnListBox.SelectedItem.ToString() : (sender as ComboBox).Text;
            var examples = resultColumns.First(x => x.Description == selectedColumName).Examples;

            ColumnExampleListBox.Items.Clear();
            ColumnExampleListBox.Items.AddRange(Enumerable.Take(examples, examplesQuantity).ToArray());
        }

        private void UnUsedColumnListBox_DoubleClick(object sender, EventArgs e)
        {
            if (lastUsedComboBox == null) return;
            if (String.IsNullOrEmpty(lastUsedComboBox.Text))
                lastUsedComboBox.Text = UnUsedColumnListBox.SelectedItem.ToString();
        }
    }
}
