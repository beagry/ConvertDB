using System;
using System.Windows.Forms;

namespace Converter
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
        }

        public void UpdateStatus(string currentBook)
        {
            progressBar1.Invoke(new Action(()=>
            {
                progressBar1.PerformStep();
                cunnrentIndexLabel.Text = String.Format("{0} из {1}", progressBar1.Value, progressBar1.Maximum);
                currentBookLabel.Text = currentBook;
            }));
        }
    }
}
