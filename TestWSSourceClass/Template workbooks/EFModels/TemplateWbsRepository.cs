namespace Converter.Template_workbooks.EFModels
{
    class TemplateWbsRepository 
    {
        private static TemplateWbsContext db;

        public static TemplateWbsContext Context { get { return db ?? (db = new TemplateWbsContext()); } }

        protected TemplateWbsRepository()
        {
            
        }
    }
}
