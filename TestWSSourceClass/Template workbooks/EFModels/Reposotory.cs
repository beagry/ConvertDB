using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Converter.Template_workbooks.EFModels
{
//    public interface IRepository<T> : IDisposable where T : class
//    {
//        /// <summary>
//        /// Gets all objects from database
//        /// </summary>
//        IQueryable<T> All();
//
//        /// <summary>
//        /// Gets objects from database by filter.
//        /// </summary>
//        /// <param name="predicate">Specified a filter</param>
//        IQueryable<T> Filter(Expression<Func<T, bool>> predicate);
//
//        /// <summary>
//        /// Gets objects from database with filting and paging.
//        /// </summary>
//        /// <typeparam name="Key"></typeparam>
//        /// <param name="filter">Specified a filter</param>
//        /// <param name="total">Returns the total records count of the filter.</param>
//        /// <param name="index">Specified the page index.</param>
//        /// <param name="size">Specified the page size</param>
//        IQueryable<T> Filter<Key>(Expression<Func<T, bool>> filter,
//            out int total, int index = 0, int size = 50);
//
//        /// <summary>
//        /// Gets the object(s) is exists in database by specified filter.
//        /// </summary>
//        /// <param name="predicate">Specified the filter expression</param>
//        bool Contains(Expression<Func<T, bool>> predicate);
//
//        /// <summary>
//        /// Find object by keys.
//        /// </summary>
//        /// <param name="keys">Specified the search keys.</param>
//        T Find(params object[] keys);
//
//        /// <summary>
//        /// Find object by specified expression.
//        /// </summary>
//        /// <param name="predicate"></param>
//        T Find(Expression<Func<T, bool>> predicate);
//
//        /// <summary>
//        /// Create a new object to database.
//        /// </summary>
//        /// <param name="t">Specified a new object to create.</param>
//        T Create(T t);
//
//        /// <summary>
//        /// Delete the object from database.
//        /// </summary>
//        /// <param name="t">Specified a existing object to delete.</param>        
//        void Delete(T t);
//
//        /// <summary>
//        /// Delete objects from database by specified filter expression.
//        /// </summary>
//        /// <param name="predicate"></param>
//        int Delete(Expression<Func<T, bool>> predicate);
//
//        /// <summary>
//        /// Update object changes and save to database.
//        /// </summary>
//        /// <param name="t">Specified the object to save.</param>
//        int Update(T t);
//
//        /// <summary>
//        /// Get the total objects count.
//        /// </summary>
//        int Count { get; }
//    }
//
//    public class Repository<TObject> : IRepository<TObject>
//        where TObject : class
//    {
//        protected TemplateWbsContext Context = null;
//        private bool shareContext = false;
//
//        public Repository()
//        {
//            System.Runtime.Remoting.Contexts.Context = new TemplateWbsContext();
//        }
//
//        public Repository(TemplateWbsContext context)
//        {
//            System.Runtime.Remoting.Contexts.Context = context;
//            shareContext = true;
//        }
//
//        protected DbSet<TObject> DbSet
//        {
//            get
//            {
//                return System.Runtime.Remoting.Contexts.Context.Set<TObject>();
//            }
//        }
//
//        public void Dispose()
//        {
//            if (shareContext && (System.Runtime.Remoting.Contexts.Context != null))
//                System.Runtime.Remoting.Contexts.Context.Dispose();
//        }
//
//        public virtual IQueryable<TObject> All()
//        {
//            return DbSet.AsQueryable();
//        }
//
//        public virtual IQueryable<TObject>
//        Filter(Expression<Func<TObject, bool>> predicate)
//        {
//            return DbSet.Where(predicate).AsQueryable<TObject>();
//        }
//
//        public virtual IQueryable<TObject> Filter(Expression<Func<TObject, bool>> filter,
//         out int total, int index = 0, int size = 50)
//        {
//            int skipCount = index * size;
//            var _resetSet = filter != null ? DbSet.Where(filter).AsQueryable() :
//                DbSet.AsQueryable();
//            _resetSet = skipCount == 0 ? _resetSet.Take(size) :
//                _resetSet.Skip(skipCount).Take(size);
//            total = _resetSet.Count();
//            return _resetSet.AsQueryable();
//        }
//
//        public bool Contains(Expression<Func<TObject, bool>> predicate)
//        {
//            return DbSet.Count(predicate) > 0;
//        }
//
//        public virtual TObject Find(params object[] keys)
//        {
//            return DbSet.Find(keys);
//        }
//
//        public virtual TObject Find(Expression<Func<TObject, bool>> predicate)
//        {
//            return DbSet.FirstOrDefault(predicate);
//        }
//
//        public virtual TObject Create(TObject TObject)
//        {
//            var newEntry = DbSet.Add(TObject);
//            if (!shareContext)
//                System.Runtime.Remoting.Contexts.Context.SaveChanges();
//            return newEntry;
//        }
//
//        public virtual int Count
//        {
//            get
//            {
//                return DbSet.Count();
//            }
//        }
//
//        public virtual int Delete(TObject TObject)
//        {
//            DbSet.Remove(TObject);
//            if (!shareContext)
//                return System.Runtime.Remoting.Contexts.Context.SaveChanges();
//            return 0;
//        }
//
//        public virtual int Update(TObject TObject)
//        {
//            var entry = System.Runtime.Remoting.Contexts.Context.Entry(TObject);
//            DbSet.Attach(TObject);
//            entry.State = EntityState.Modified;
//            if (!shareContext)
//                return System.Runtime.Remoting.Contexts.Context.SaveChanges();
//            return 0;
//        }
//
//        public virtual int Delete(Expression<Func<TObject, bool>> predicate)
//        {
//            var objects = Filter(predicate);
//            foreach (var obj in objects)
//                DbSet.Remove(obj);
//            if (!shareContext)
//                return System.Runtime.Remoting.Contexts.Context.SaveChanges();
//            return 0;
//        }
//    }
}
