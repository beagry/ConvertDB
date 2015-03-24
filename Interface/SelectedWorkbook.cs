using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace UI
{
    public static class EnumerableExtentions
    {
        public static IEnumerable<string> GetWorkbooksPaths(this IEnumerable<SelectedWorkbook> workbooks)
        {
            return workbooks == null ? null : workbooks.Select(w => w.Path);
        }
    }
    public struct SelectedWorkbook
    {
        public SelectedWorkbook(string path) : this()
        {
            Path = path;
            if (File.Exists(path))
                Name = System.IO.Path.GetFileName(path);
        }
        public string Path { get; set; }
        public string Name { get; private set ; }
    }
}
