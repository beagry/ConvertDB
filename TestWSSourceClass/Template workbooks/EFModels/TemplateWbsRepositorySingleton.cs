using System.Collections.Generic;

namespace Converter.Template_workbooks.EFModels
{
    public class TemplateWbsRepositorySingleton 
    {
        private static TemplateWbsContext db;

        private static TemplateWbsRespository respository;

        public static TemplateWbsContext Context { get { return db ?? (db = new TemplateWbsContext()); } }

        public static TemplateWbsRespository Respository { get { return respository ?? (respository = new TemplateWbsRespository()); } }

        protected TemplateWbsRepositorySingleton()
        {
            
        }
    }
}
