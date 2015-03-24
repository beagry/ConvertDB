using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Converter;
using UI;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            var binded = new Dictionary<string, ICollection<string>>()
            {
                {"SUBJECT",new List<string>(){"COLUMN1","COLUMN2","ADDITIONAL_1","ONE_MORE"}},
                {"REGION",new List<string>(){"REGION1","SOME_OTHER_REGION"}},
                {"NEAR_CITY",new List<string>(){"CITY","SUPER_CITY"}},
                {"DESCTIPTION",new List<string>(){"MY_BLA_BLA_BLA","TELL_SOME_PURE"}},
            };

            var allColumns = new List<WorksheetInfo>
            {
                new WorksheetInfo("WS1",new List<WorksheetColumnInfo>(){new WorksheetColumnInfo(1,"COLUMN1"),new WorksheetColumnInfo(2,"COLUMN2"),new WorksheetColumnInfo(3,"ADDITIONAL_1")}),
                new WorksheetInfo("WS2",new List<WorksheetColumnInfo>(){new WorksheetColumnInfo(1,"REGION1"),new WorksheetColumnInfo(2,"SOME_OTHER_REGION"),new WorksheetColumnInfo(3,"ANOTHER_ADDITIONAL")}),
                new WorksheetInfo("WS3",new List<WorksheetColumnInfo>(){new WorksheetColumnInfo(1,"CITY"),new WorksheetColumnInfo(2,"SUPER_CITY"),new WorksheetColumnInfo(3,"ONE_MORE")}),
                new WorksheetInfo("WS4",new List<WorksheetColumnInfo>(){new WorksheetColumnInfo(1,"MY_BLA_BLA_BLA"),new WorksheetColumnInfo(2,"TELL_SOME_PURE"),new WorksheetColumnInfo(3,"AND_MORE_ONE_COLUMN")}),
            };

            DataContext = new CompareViewModel(binded, allColumns);
        }
    }
}
