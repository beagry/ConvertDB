namespace Converter.Template_workbooks.EFModels
{
    public class UnitOfWorkSingleton
    {
        private static TemplateWbsContext _db;
        private static UnitOfWork _unitOfWork;

        protected UnitOfWorkSingleton()
        {
        }

        public static TemplateWbsContext Context
        {
            get { return _db ?? (_db = new TemplateWbsContext()); }
        }

        public static UnitOfWork UnitOfWork
        {
            get { return _unitOfWork ?? (_unitOfWork = new UnitOfWork()); }
        }
    }
}