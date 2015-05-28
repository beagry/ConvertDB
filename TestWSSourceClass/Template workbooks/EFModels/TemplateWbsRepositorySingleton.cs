using System.Collections.Generic;

namespace Converter.Template_workbooks.EFModels
{
    public class UnitOfWorkSingleton
    {
        private static TemplateWbsContext _db;

        private static UnitOfWork _unitOfWork;

        public static TemplateWbsContext Context
        {
            get { return _db ?? (_db = new TemplateWbsContext()); }
        }

        public static UnitOfWork UnitOfWork
        {
            get { return _unitOfWork ?? (_unitOfWork = new UnitOfWork()); }
        }

        protected UnitOfWorkSingleton()
        {
        }
    }
}
