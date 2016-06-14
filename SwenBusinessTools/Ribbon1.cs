using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel; 


namespace SwenBusinessTools
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();

            //ribbonDropDownItemImpl1.OfficeImageId = Word.
            //ribbonDropDownItemImpl1.Label = "Offerta Economica";
            //ribbonDropDownItemImpl1.ScreenTip = "Offerta Economica";
            //ribbonDropDownItemImpl1.SuperTip = "Template word di offerta economica";

            //this.gryTemplates.Items.Add(ribbonDropDownItemImpl1);

            //gallery1.Items.

            

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void grpImpostazioni_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
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

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = false;
            
            var workBook = excelApp.Workbooks.Open(@"c:\temp\doc1.xls");

            var sheet = (Excel.Worksheet)workBook.Sheets["Foglio1"];

            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            var pArea = sheet.PageSetup.PrintArea;
            sheet.Range[pArea].Copy();
            
            var doc = Globals.ThisAddIn.Application.ActiveDocument;

            doc.Paragraphs[1].Range.PasteExcelTable(true, false, false);

            var WordTable = doc.Tables[1];
            WordTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

            workBook.Close();
            excelApp.Quit();
        }

        private void gryTemplates_Click(object sender, RibbonControlEventArgs e)
        {
            var selectedItem = ((Microsoft.Office.Tools.Ribbon.RibbonGallery)sender).SelectedItem;

            switch (selectedItem.Id)
            {
                case "__id3":
                    var temp = new frmGenerazioneTemplateFattura();
                    temp.ShowDialog();
                    break;

            }
        }

        private void btnAggiungiProgetti_Click(object sender, RibbonControlEventArgs e)
        {
            var progetto = new frmAggiungiProgetto();
            progetto.Show();
        }
    }
}
