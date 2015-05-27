using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converter.Template_workbooks.EFModels
{
    class TemplateWbsContext:DbContext
    {
        public TemplateWbsContext():base("TWBsContext")
        {
            Database.SetInitializer(new TemplateWbsInitializer());
        }
        public DbSet<TemplateWorkbook> TemplateWorkbooks { get; set; }

    }

    class TemplateWbsInitializer: DropCreateDatabaseAlways<TemplateWbsContext>
    {
        protected override void Seed(TemplateWbsContext context)
        {
            base.Seed(context);
            InitializeLandWorkbook(context);

        }

        private void InitializeLandWorkbook(TemplateWbsContext context)
        {
            var s11 = new SearchCritetia() { Text = "1" };
            var s21 = new SearchCritetia() { Text = "2" };
            var s31 = new SearchCritetia() { Text = "3" };

            var lc1 = new BindedColumn() { Name = "Clmn21" };
            var lc2 = new BindedColumn() { Name = "Clmn454" };
            var lc3 = new BindedColumn() { Name = "SubjColumn" };


            var c1 = new TemplateColumn() { Name = "Колонка 1", CodeName = "Column1" };
            c1.BindedColumns.Add(lc1);
            c1.SearchCritetias.Add(s11);

            var c2 = new TemplateColumn() { Name = "Колонка 2", CodeName = "Column3" };
            c2.BindedColumns.Add(lc2);
            c2.SearchCritetias.Add(s21);

            var c3 = new TemplateColumn() { Name = "Колонка 3", CodeName = "Column3" };
            c3.BindedColumns.Add(lc3);
            c3.SearchCritetias.Add(s31);

            var landWb = new TemplateWorkbook() { WorkbookType = XlTemplateWorkbookType.LandProperty };
            landWb.Columns.AddRange(new List<TemplateColumn>() { c1, c2, c3 });


            context.SaveChanges();
        }
    }
}
