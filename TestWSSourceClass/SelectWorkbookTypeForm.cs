using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Converter.Template_workbooks;
using TestWSSourceClass;

namespace Converter
{
    public partial class SelectWorkbookTypeForm : Form
    {
        private List<TemplateWorkbook> _workbook;
        private Dictionary<XlTemplateWorkbooks, string> templateWorkbookNamesDictionary; 
        public SelectWorkbookTypeForm(ref List<TemplateWorkbook> workbook)
        {
            InitializeComponent();
            _workbook = workbook;
            
            //Соответствия коду шаблоны и его описания
            templateWorkbookNamesDictionary = new Dictionary<XlTemplateWorkbooks, string>();
            templateWorkbookNamesDictionary.Add(XlTemplateWorkbooks.LandProperty, "Земельные участки");
            templateWorkbookNamesDictionary.Add(XlTemplateWorkbooks.CommerceProperty, "Коммерческая недвижимость");
            templateWorkbookNamesDictionary.Add(XlTemplateWorkbooks.CityLivaArea, "Городское жильё");
            templateWorkbookNamesDictionary.Add(XlTemplateWorkbooks.CountyLiveArea, "Загородное жильё");


            //Заполняем выпадающий список всеми видами шаблонов
            foreach (KeyValuePair<XlTemplateWorkbooks, string> keyValuePair in templateWorkbookNamesDictionary)
            {
                comboBox1.Items.Add(keyValuePair.Value);
            }

//            this.Activated += (sender, args) =>
//            {
//                comboBox1.DroppedDown = true;
//            };
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == String.Empty)
            {
                MessageBox.Show(@"Не выбран тип книг",@"Операция прервана",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            }

            XlTemplateWorkbooks wbType = templateWorkbookNamesDictionary.First(x => x.Value == comboBox1.Text).Key;

            switch (wbType)
            {
                case XlTemplateWorkbooks.CommerceProperty:
                    _workbook.Add(new CommercePropertyTemplateWorkbook());
                    break;
                case XlTemplateWorkbooks.CityLivaArea:
                    _workbook.Add(new CityLivaAreaTemplateWorkbook());
                    break;
                case XlTemplateWorkbooks.CountyLiveArea:
                    _workbook.Add(new CountryLiveAreaTemplateWorkbook());
                    break;
                case XlTemplateWorkbooks.LandProperty:
                    _workbook.Add(new LandPropertyTemplateWorkbook());
                    break;
            }
            Close();
        }

        private void SelectWorkbookTypeForm_Load(object sender, EventArgs e)
        {

        }
    }
}
