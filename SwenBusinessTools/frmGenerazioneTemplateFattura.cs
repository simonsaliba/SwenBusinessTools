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
using static SwenBusinessTools.CustomStyle;

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
           
        }

        private void chkSER_CheckedChanged(object sender, EventArgs e)
        {
            txtSER.Enabled = chkSER.Checked;
        }

        private void wizardOffertaEconomica_Finished(object sender, EventArgs e)
        {
            var document = Globals.ThisAddIn.Application.ActiveDocument;

            SetCustomStyle(document);

            Word.Paragraph notaTecnica = document.Paragraphs.Add(System.Reflection.Missing.Value);
            notaTecnica.Range.Text = "Nota tecnica";
            notaTecnica.Range.set_Style("titolo 3");
            notaTecnica.Range.InsertParagraphAfter();
            

            Word.Paragraph testoNotaTecnica = document.Paragraphs.Add(notaTecnica.Range);
            testoNotaTecnica.Range.Text = txtNotatecnica.Text;
            testoNotaTecnica.Range.set_Style("normale");
            testoNotaTecnica.Range.InsertParagraphAfter();


            Word.Paragraph notaCommerciale = document.Paragraphs.Add(testoNotaTecnica.Range);
            notaCommerciale.Range.Text = "Nota commerciale";
            notaCommerciale.Range.set_Style("titolo 3");
            notaCommerciale.Range.InsertParagraphAfter();


            Word.Paragraph testoNotaCommerciale = document.Paragraphs.Add(notaCommerciale.Range);
            testoNotaCommerciale.Range.Text = txtNotaCommerciale.Text;
            testoNotaCommerciale.Range.set_Style("normale");
            testoNotaCommerciale.Range.InsertParagraphAfter();


            Word.Paragraph condizioniGenerali = document.Paragraphs.Add(testoNotaCommerciale.Range);
            condizioniGenerali.Range.Text = "Condizioni generali";
            condizioniGenerali.Range.set_Style("titolo 3");
            condizioniGenerali.Range.InsertParagraphAfter();


            Word.Paragraph testoCondizioniGenerali = document.Paragraphs.Add(condizioniGenerali.Range);
            testoCondizioniGenerali.Range.Text = "La presente offerta si intende valida per 7 gg. alle seguenti condizioni di fornitura (individuare le voci che interessano sulla base della categoria indicata per ciascun articolo in offerta); l’accettazione della presente offerta implica la tacita accettazione di tutte le condizioni applicabili a ciascuna categoria merceologica offerta.";
            testoCondizioniGenerali.Range.set_Style("normale");
            testoCondizioniGenerali.Range.InsertParagraphAfter();


            Word.Paragraph testoGenerali = document.Paragraphs.Add(testoCondizioniGenerali.Range);
            testoGenerali.Range.Text = "Generali:";
            testoGenerali.Range.set_Style("condizionigenerali");
            testoGenerali.Range.Bold = 1;
            testoGenerali.Range.InsertParagraphAfter();


            Word.Paragraph testoElencoGenerale = document.Paragraphs.Add(testoGenerali.Range);
            testoElencoGenerale.Range.ListFormat.ApplyBulletDefault(Word.WdListType.wdListBullet);
            //testoElencoGenerale.Range.Text = 
            testoElencoGenerale.Range.set_Style("condizionigenerali");

            testoElencoGenerale.Range.InsertAfter("I prezzi sono da considerarsi al netto dell'IVA.");
            testoElencoGenerale.Range.InsertAfter("Il pagamento non potrà comunque essere differito oltre i limiti indicati in fattura a qualunque titolo, incluse eventuali contestazioni o malfunzionamenti anche parziali sia su prodotti SWEN o di terzi, (che vanno comunque disciplinati come interventi in garanzia da ");


            Word.Table table = document.Tables.Add(testoNotaCommerciale.Range, 5, 2);
            SetTableBolders(table);
            table.Range.set_Style("TabellaFirma");

            //first row
            Word.Cell cellOrdineAcquisto = table.Cell(1, 1);
            cellOrdineAcquisto.Range.Text = "Ordine di acquisto";
            cellOrdineAcquisto.Range.Font.Bold = 1;
            cellOrdineAcquisto.Range.Font.Size = 13.0f;
            cellOrdineAcquisto.Range.Font.Color = (Word.WdColor)(128 + 0x100 * 128 + 0x10000 * 128);

            Word.Cell cellRiferimento = table.Cell(1, 2);
            cellRiferimento.Range.Text = "Riferimento (sarà citato in fattura)";

            //2° row
            Word.Cell cellRrdianante= table.Cell(2, 1);
            cellRrdianante.Range.Text = "Ordinante (si prega specificare il nome e cognome del responsabile per l’ordine)";

            Word.Cell cellDataApprovazione = table.Cell(2, 2);
            cellDataApprovazione.Range.Text = "Data approvazione";

            //3° row
            Word.Cell cellOpzioniScelte = table.Cell(3, 1);
            cellOpzioniScelte.Range.Text = "Opzioni scelte (indicare lista dei riferimenti in caso di opzioni o alternative)";

            Word.Cell cellImportoTotale = table.Cell(3, 2);
            cellImportoTotale.Range.Text = "IMPORTO TOTALE escl. IVA (incluse opzioni desiderate))";

            //4° row
            Word.Cell firmaClienteAccettazione = table.Cell(4, 1);
            firmaClienteAccettazione.Merge(table.Cell(4, 2));
            firmaClienteAccettazione.Range.Text = "Firma del Cliente per accettazione della presente offerta quale ordine di acquisto";

            //5° row
            Word.Cell firmaCliente = table.Cell(5, 1);
            firmaCliente.Merge(table.Cell(5, 2));
            firmaCliente.Range.Text = "Firma del Cliente ai sensi degli art. 1341 e 1342 del Codice Civile e successive modificazioni, per approvazione esplicita di tutti i paragrafi della presente offerta, in particolare ogni singolo comma del paragrafo “condizioni generali”. ";
        }
    }
}
