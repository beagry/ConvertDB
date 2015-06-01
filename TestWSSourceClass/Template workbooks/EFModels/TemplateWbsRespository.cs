using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;

namespace Converter.Template_workbooks.EFModels
{
    internal interface IRepository<T> : IDisposable where T : class
    {
        IEnumerable<T> GetObjectsList();
        T GetObject(int id);
        void Create(T item);
        void Update(T item);
        void Delete(int id);
        void Save();
    }


    public class UnitOfWork : IDisposable
    {
        private readonly TemplateWbsContext db = new TemplateWbsContext();
        private bool disposed;
        private TemplateColumnRepository templateColumnRepository;
        private TemplateWbsRespository templateWbsRespository;

        public TemplateWbsRespository TemplateWbsRespository
        {
            get { return templateWbsRespository ?? (templateWbsRespository = new TemplateWbsRespository(db)); }
        }

        public TemplateColumnRepository TemplateColumnRepository
        {
            get { return templateColumnRepository ?? (templateColumnRepository = new TemplateColumnRepository(db)); }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Save()
        {
            db.SaveChanges();
        }

        public virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    db.Dispose();
                }
            }
            disposed = true;
        }
    }

    public sealed class TemplateColumnRepository : IRepository<TemplateColumn>
    {
        private bool disposed;

        public TemplateColumnRepository(TemplateWbsContext db)
        {
            Context = db;
        }

        public TemplateColumnRepository()
        {
            Context = UnitOfWorkSingleton.Context;
        }

        public TemplateWbsContext Context { get; private set; }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public IEnumerable<TemplateColumn> GetObjectsList()
        {
            return Context.TemplateColumns;
        }

        public TemplateColumn GetObject(int id)
        {
            return Context.TemplateColumns.Find(id);
        }

        public void Create(TemplateColumn item)
        {
            Context.TemplateColumns.Add(item);
        }

        public void Update(TemplateColumn item)
        {
            Context.Entry(item).State = EntityState.Modified;
        }

        public void Delete(int id)
        {
            var column = Context.TemplateColumns.Find(id);
            if (column == null) return;
            Context.TemplateColumns.Remove(column);
        }

        public void Save()
        {
            Context.SaveChanges();
        }

        public void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    Context.Dispose();
                }
            }
            disposed = true;
        }
    }

    public sealed class TemplateWbsRespository : IRepository<TemplateWorkbook>
    {
        private bool disposed;

        public TemplateWbsRespository(TemplateWbsContext db)
        {
            Context = db;
        }

        public TemplateWbsRespository()
        {
            Context = UnitOfWorkSingleton.Context;
        }

        public TemplateWbsContext Context { get; private set; }

        public IEnumerable<TemplateWorkbook> GetObjectsList()
        {
            return Context.TemplateWorkbooks;
        }

        public TemplateWorkbook GetObject(int id)
        {
            return Context.TemplateWorkbooks.Find(id);
        }

        public void Create(TemplateWorkbook item)
        {
            Context.TemplateWorkbooks.Add(item);
        }

        public void Update(TemplateWorkbook item)
        {
            Context.Entry(item).State = EntityState.Modified;
        }

        public void Delete(int id)
        {
            var wb = Context.TemplateWorkbooks.Find(id);
            if (wb == null) return;

            Context.TemplateWorkbooks.Remove(wb);
        }

        public void Save()
        {
            Context.SaveChanges();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void AddColumnColumnPair(string templateColumName, string bindingColumnName)
        {
            //add new reference
            var newOnwerColumn = Context.TemplateColumns.FirstOrDefault(c => c.CodeName.Equals(templateColumName));
            
            //NULL when column was handle created
            if (newOnwerColumn == null) return;
            newOnwerColumn.BindedColumns.Add(new BindedColumn {Name = bindingColumnName});
            Context.Entry(newOnwerColumn).State = EntityState.Modified;
        }

        public void RemoveColumnColumnPair(string templateColumName, string bindingColumnName)
        {
            //Remove old bindingColumn reference

            var col = Context.TemplateColumns.FirstOrDefault(c => c.CodeName.Equals(templateColumName));
            if (col == null) return;
            var bindedColumn = col.BindedColumns.FirstOrDefault(b => b.Name.Equals(bindingColumnName));
            if (bindedColumn == null) return;

            col.BindedColumns.Remove(bindedColumn);
            Context.BindedColumns.Remove(bindedColumn);
        }

        public IEnumerable<TemplateColumn> TemplateColumnsOfWorkbook(XlTemplateWorkbookType wbType)
        {
            return Context.TemplateColumns.Where(c => c.TemplateWorkbooks.Any(w => w.WorkbookType == wbType));
        }

        public TemplateWorkbook GetTypedWorkbook(XlTemplateWorkbookType wbType)
        {
            return Context.TemplateWorkbooks.FirstOrDefault(w => w.WorkbookType == wbType);
        }

        public void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    Context.Dispose();
                }
            }
            disposed = true;
        }
    }
}