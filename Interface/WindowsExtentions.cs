using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace UI
{
    public static class WindowsExtentions
    {
        public static void ListBox_DropWorkbook(object sender, DragEventArgs e)
        {
            var list = sender as ListBox;
            if (list == null) return;

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (files != null)
                files.ToList().ForEach(s =>
                {
                    if (list.Items.Cast<SelectedWorkbook>().All(w => w.Path != s))
                        list.Items.Add(new SelectedWorkbook(s));
                });
        }
    }
}
