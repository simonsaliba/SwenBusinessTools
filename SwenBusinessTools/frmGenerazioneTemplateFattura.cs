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
    public partial class frmGenerazioneTemplateFattura : Form
    {
        public frmGenerazioneTemplateFattura()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void chkAggiungiNotaTecnica_CheckedChanged(object sender, EventArgs e)
        {
            txtNotatecnica.Enabled = chkAggiungiNotaTecnica.Checked;
        }

        private void chkNotaCommerciale_CheckedChanged(object sender, EventArgs e)
        {
            txtNotaCommerciale.Enabled = chkNotaCommerciale.Checked;
        }

        private void chkNotaInterpretativa_CheckedChanged(object sender, EventArgs e)
        {
            txtNotaInterpretativa.Enabled = chkNotaInterpretativa.Checked;
        }

        private void chkACC_CheckedChanged(object sender, EventArgs e)
        {
            txtACC.Enabled = chkACC.Checked;
        }

        private void chkSER_CheckedChanged(object sender, EventArgs e)
        {
            txtSER.Enabled = chkSER.Checked;
        }

        private void wizardOffertaEconomica_Finished(object sender, EventArgs e)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            Word.Paragraph notaTecnica = document.Paragraphs.Add(System.Reflection.Missing.Value);
            notaTecnica.Range.Text = "Nota tecnica";
            notaTecnica.Range.Font.Bold = 1;
            notaTecnica.Range.Font.Size = 13.0f;
            notaTecnica.Range.Font.Name = "Trebuchet MS";
            notaTecnica.Range.Font.Color = Word.WdColor.wdColorBlue;
            notaTecnica.Range.InsertParagraphAfter();

            Word.Paragraph testoNotaTecnica = document.Paragraphs.Add(notaTecnica.Range);
            testoNotaTecnica.Range.Text = txtNotatecnica.Text;
            testoNotaTecnica.Range.Font.Bold = 0;
            testoNotaTecnica.Range.Font.Size = 10.0f;
            testoNotaTecnica.Range.Font.Name = "Trebuchet MS";
            testoNotaTecnica.Range.Font.Color = Word.WdColor.wdColorBlack;
            testoNotaTecnica.Range.InsertParagraphAfter();


            Word.Paragraph notaCommerciale = document.Paragraphs.Add(testoNotaTecnica.Range);
            notaCommerciale.Range.Text = "Nota commerciale";
            notaCommerciale.Range.Font.Bold = 1;
            notaCommerciale.Range.Font.Size = 13.0f;
            notaCommerciale.Range.Font.Name = "Trebuchet MS";
            notaCommerciale.Range.Font.Color = Word.WdColor.wdColorDarkBlue;
            notaCommerciale.Range.InsertParagraphAfter();


            Word.Paragraph testoNotaCommerciale = document.Paragraphs.Add(notaCommerciale.Range);
            testoNotaCommerciale.Range.Text = txtNotaCommerciale.Text;
            testoNotaCommerciale.Range.Font.Bold = 0;
            testoNotaCommerciale.Range.Font.Size = 10.0f;
            testoNotaCommerciale.Range.Font.Name = "Trebuchet MS";
            testoNotaCommerciale.Range.Font.Color = Word.WdColor.wdColorBlack;
            testoNotaCommerciale.Range.InsertParagraphAfter();

        }
    }
}
