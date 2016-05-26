using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SwenBusinessTools
{
    public partial class SwenUserControl : UserControl
    {
        public SwenUserControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var app = Globals.ThisAddIn.Application;

            string fileName = @"c:\temp\prova.pdf";

           // app.ActiveWindow.Document.SaveAs2(ref fileName, )
        }
    }
}
