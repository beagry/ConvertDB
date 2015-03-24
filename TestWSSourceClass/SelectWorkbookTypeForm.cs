using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Converter.Template_workbooks;

namespace Converter
{
    public partial class SelectWorkbookTypeForm : Form
    {
        private List<TemplateWorkbook> _workbook;
        private Dictionary<XlTemplateWorkbookTypes, string> templateWorkbookNamesDictionary; 
        public SelectWorkbookTypeForm(ref List<TemplateWorkbook> workbook)
        {
            InitializeComponent();
            _workbook = workbook;
            
            //Соответствия коду шаблоны и его описания
            templateWorkbookNamesDictionary = new Dictionary<XlTemplateWorkbookTypes, string>();
            templateWorkbookNamesDictionary.Add(XlTemplateWorkbookTypes.LandProperty, "Земельные участки");
            templateWorkbookNamesDictionary.Add(XlTemplateWorkbookTypes.CommerceProperty, "Коммерческая недвижимость");
            templateWorkbookNamesDictionary.Add(XlTemplateWorkbookTypes.CityLivaArea, "Городское жильё");
            templateWorkbookNamesDictionary.Add(XlTemplateWorkbookTypes.CountyLiveArea, "Загородное жильё");


            //Заполняем выпадающий список всеми видами шаблонов
            foreach (KeyValuePair<XlTemplateWorkbookTypes, string> keyValuePair in templateWorkbookNamesDictionary)
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

            XlTemplateWorkbookTypes wbType = templateWorkbookNamesDictionary.First(x => x.Value == comboBox1.Text).Key;

            switch (wbType)
            {
                case XlTemplateWorkbookTypes.CommerceProperty:
                    _workbook.Add(new CommercePropertyTemplateWorkbook());
                    break;
                case XlTemplateWorkbookTypes.CityLivaArea:
                    _workbook.Add(new CityLivaAreaTemplateWorkbook());
                    break;
                case XlTemplateWorkbookTypes.CountyLiveArea:
                    _workbook.Add(new CountryLiveAreaTemplateWorkbook());
                    break;
                case XlTemplateWorkbookTypes.LandProperty:
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
