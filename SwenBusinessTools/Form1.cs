using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace SwenBusinessTools
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;

            var properties = doc.CustomDocumentProperties;
            //((Microsoft.Office.Core.DocumentProperties)properties).Add("NOME_AZIENDA", false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, "SWEN S.R.L");


           // textBox1.Text = ReadDocumentProperty(doc, "NOME_AZIENDA");

        }

        private string ReadDocumentProperty(Word.Document doc, string propertyName)
        {
            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Microsoft.Office.Core.DocumentProperties)doc.CustomDocumentProperties;

            foreach (Microsoft.Office.Core.DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return null;
        }
    }
}
