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
using Office = Microsoft.Office.Core;
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
            testoGenerali.Range.InsertParagraphAfter();

            Word.Paragraph testoElencoGenerale = document.Paragraphs.Add(testoGenerali.Range);
            testoElencoGenerale.Range.ListFormat.ApplyBulletDefault();
            
            testoElencoGenerale.Range.ParagraphFormat.SpaceAfterAuto = 0;
            testoElencoGenerale.Range.ParagraphFormat.SpaceBeforeAuto = 0;
            testoElencoGenerale.Range.ParagraphFormat.FirstLineIndent = -7f;
            testoElencoGenerale.Range.ParagraphFormat.LeftIndent = 7f;
            testoElencoGenerale.Range.ParagraphFormat.SpaceAfter = 0F;
            testoElencoGenerale.Range.ParagraphFormat.SpaceBefore = 0F;
            testoElencoGenerale.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
            testoElencoGenerale.Range.ParagraphFormat.LineSpacing = 13.8f;
            testoElencoGenerale.Range.Font.Bold = 0;
            testoElencoGenerale.Range.Font.Color = Word.WdColor.wdColorBlack;

            testoElencoGenerale.Range.InsertBefore("I prezzi sono da considerarsi al netto dell'IVA \n");
            testoElencoGenerale.Range.InsertBefore("Il pagamento non potrà comunque essere differito oltre i limiti indicati in fattura a qualunque titolo, incluse eventuali contestazioni o malfunzionamenti anche parziali sia su prodotti Swen o di terzi, (che vanno comunque disciplinati come \"interventi in garanzia\", da espletarsi come specificato nel seguito) che sui servizi resi da Swen, per i quali valgono le penali stabilite nei SLA concordati con il Cliente \n");
            testoElencoGenerale.Range.InsertBefore("Il mancato o ritardato pagamento delle fatture emesse a qualunque titolo nei confronti di un Cliente comporta l’immediata sospensione di servizi e forniture servizio fino a regolarizzazione, fatti salvi eventuali interessi, risarcimenti e danni subiti.\n");
            testoElencoGenerale.Range.InsertBefore("Eventuali supplementi richiesti dal Cliente (diversi da  quanto specificato nella presente offerta) devono essere comunque disciplinati in separata sede\n");
            testoElencoGenerale.Range.InsertBefore("La consegna di tutti i beni elencati e comunque il completamento della fornitura è subordinato alla permanenza di disponibilità commerciale dei prodotti offerti. In caso di documentata indisponibilità sul mercato nei tempi stabiliti per la consegna, la SWEN si riserva il diritto di escludere alcuni degli articoli offerti dalla fornitura senza alcuna penale (se non la restituzione di eventuali acconti già versati in proporzione al valore degli articoli non disponibili); qualora l’indisponibilità sia temporanea il Cliente potrà scegliere se attendere la nuova disponibilità o rinunciare agli articoli non disponibili; in caso si indisponibilità permanente è implicita la rinuncia del Cliente\n");
            testoElencoGenerale.Range.InsertBefore("In presenza di articoli indisponibili, la SWEN si impegna a fornire prodotti alternativi di pari requisiti con offerta separata, a prezzi e condizioni da rinegoziare\n");
            testoElencoGenerale.Range.InsertBefore("I beni materiali sono coperti da garanzie a norma di legge ed in particolare:\n");
            testoElencoGenerale.Range.ListFormat.ListIndent();
            testoElencoGenerale.Range.InsertBefore("per i consumatori, cioè coloro che acquistano per scopi estranei alla propria attività professionale o imprenditoriale, il venditore applicherà il Decreto Legislativo 2 febbraio 2002, n.24. - artt. 1519-bis e seguenti c.c. - (due anni dalla consegna alle condizioni di legge); \n");
            testoElencoGenerale.Range.InsertBefore("per gli altri acquirenti, che solitamente acquistano con partita IVA, varranno le garanzie di legge di cui agli articoli 1490 e seguenti c.c. (un anno dalla consegna alle condizioni di legge). \n");
            testoElencoGenerale.Range.InsertBefore("Restano in ogni caso fatte salve eventuali deroghe specifiche per prodotto o categoria di prodotti (come di seguito indicato) e le garanzie contrattuali rilasciate direttamente dal produttore.");



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
