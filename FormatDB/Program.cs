using System;
using System.Windows.Forms;

namespace Formater
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            MainForm form = new MainForm();
            Application.Run(form);
        }
    }
}
