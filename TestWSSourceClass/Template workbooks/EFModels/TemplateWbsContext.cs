using System;
using System.Data.Entity;
using System.IO;

namespace Converter.Template_workbooks.EFModels
{
    public class TemplateWbsContext : DbContext
    {
        public TemplateWbsContext() : base("TWBsContext")
        {
            AppDomain.CurrentDomain.SetData("DataDirectory", Directory.GetCurrentDirectory());
            Database.SetInitializer(new TemplateWbsInitializer());
        }

        public DbSet<TemplateWorkbook> TemplateWorkbooks { get; set; }
        public DbSet<TemplateColumn> TemplateColumns { get; set; }

        public DbSet<BindedColumn> BindedColumns { get; set; } 

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<TemplateWorkbook>()
                .HasMany(p => p.Columns)
                .WithMany(p => p.TemplateWorkbooks);
        }
    }
}