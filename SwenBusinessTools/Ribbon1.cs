using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace SwenBusinessTools
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new Form1();
            form.Show();
        }

        private void grpImpostazioni_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            var form = new Form1();
            form.Show();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            var templateOfferta = SwenBusinessTools.Properties.Resources.template_offerta;
            var template =  System.IO.Path.Combine(System.IO.Path.GetTempPath(), "template_offerta.dotx");
            System.IO.File.WriteAllBytes(template, templateOfferta);

            var doc = Globals.ThisAddIn.Application.Documents.Add(template);
            doc.Activate();

            var properties = doc.CustomDocumentProperties;
            ((Microsoft.Office.Core.DocumentProperties)properties).Add("NOME_AZIENDA", false, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, "SWEN S.R.L");

            //foreach (Word.Section wordSection in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            //{
            //    Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            //    footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
            //    footerRange.Font.Size = 20;
            //    footerRange.Text = "Confidential";
            //}

            //foreach (Word.Section section in Globals.ThisAddIn.Application.ActiveDocument.Sections)
            //{
            //    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

            //    //headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldAuthor);
            //    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldDocProperty, "NOME_AZIENDA");
            //    // headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);

            //    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            //}

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


        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Close();
        }

        private void btnChiudi_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.Application.Quit();
            Globals.ThisAddIn.Application.ActiveDocument.Close();
        }
    }
}
