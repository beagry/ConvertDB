using System;
using System.Data.Entity;
using System.Text;
using System.Threading.Tasks;
using ExcelRLibrary;

namespace Converter.Template_workbooks.EFModels
{
    public class TemplateWbsContext:DbContext
    {
        public TemplateWbsContext():base("TWBsContext")
        {
            Database.SetInitializer(new TemplateWbsInitializer());
        }
        public DbSet<TemplateWorkbook> TemplateWorkbooks { get; set; }
        public DbSet<TemplateColumn> TemplateColumns { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<TemplateWorkbook>()
                .HasMany(p => p.Columns)
                .WithMany(p => p.TemplateWorkbooks);

        }
    }
}
