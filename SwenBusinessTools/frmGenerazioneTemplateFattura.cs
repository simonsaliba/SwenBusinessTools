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
using static SwenBusinessTools.Common;
using System.Reflection;

namespace SwenBusinessTools
{
    public partial class frmGenerazioneTemplateFattura : Form
    {
        Word.Document document = null;

        Word.Selection Selection = null;

        object Missing = System.Reflection.Missing.Value;

        public frmGenerazioneTemplateFattura()
        {
            InitializeComponent();
            document = Globals.ThisAddIn.Application.ActiveDocument;
            Selection = Globals.ThisAddIn.Application.Selection;
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

        private void AddLogo(Word.Document document)
        {
            var logoSwen = document.InlineShapes.AddPicture(@"./Resources/Image1.png", false, true);
            logoSwen.Height = CentimetersToPoints(3.66f);
            logoSwen.Width = CentimetersToPoints(6.48f);

            var logoShape = logoSwen.ConvertToShape();
            logoShape.Left = 0;
            logoShape.Top = 0;
        }

        private void SetDocumentMargin()
        {
            //Imposta margini
            document.PageSetup.TopMargin = CentimetersToPoints(1.37f);
            document.PageSetup.BottomMargin = CentimetersToPoints(2.75f);
            document.PageSetup.LeftMargin = CentimetersToPoints(1.5f);
            document.PageSetup.RightMargin = CentimetersToPoints(1.5f);
            document.PageSetup.Gutter = CentimetersToPoints(0.5f);
            document.PageSetup.HeaderDistance = CentimetersToPoints(1.3f);
            document.PageSetup.FooterDistance = CentimetersToPoints(1.25f);

            //Orientamento
            document.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
        }

        private void wizardOffertaEconomica_Finished(object sender, EventArgs e)
        {

            SetDocumentMargin();

            SetCustomStyle(document);

            //object noReset = false;
            //object password = "simon";
            //object useIRM = false;
            //object enforceStyleLock = false;

            
            //Aggiungi logo al documento
            var logoSwen = document.InlineShapes.AddPicture(@"./Resources/Image1.png", false, true);
            logoSwen.Height = CentimetersToPoints(3.66f);
            logoSwen.Width = CentimetersToPoints(6.48f);

            var logoShape = logoSwen.ConvertToShape();
            //logoShape.WrapFormat. = 0;
            //logoShape.WrapFormat.Side = Word.WdWrapSideType.
            logoShape.Left = 0;
            logoShape.Top = 0;
            

            #region Tabella Tipo Documento
            Word.Table t0 = document.Tables.Add(document.Range(0,0), 7, 9);

            t0.Range.set_Style("TabellaTipoDoc");
            t0.Range.Font.Size = 9;

            //imposta le colonne
            t0.Columns[1].Width = CentimetersToPoints(4.01f);

            //imposta le righe
            t0.Rows[1].HeightRule = Word.WdRowHeightRule.wdRowHeightAtLeast;
            t0.Rows[1].Height = CentimetersToPoints(0.29f);
            t0.Rows[2].HeightRule = Word.WdRowHeightRule.wdRowHeightAtLeast;
            t0.Rows[2].Height = CentimetersToPoints(1.22f);
            t0.Rows[3].HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            t0.Rows[3].Height = CentimetersToPoints(0.45f);
            t0.Rows[4].HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            t0.Rows[4].Height = CentimetersToPoints(0.45f);
            t0.Rows[5].HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            t0.Rows[5].Height = CentimetersToPoints(0.45f);
            t0.Rows[6].HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            t0.Rows[6].Height = CentimetersToPoints(0.45f);
            t0.Rows[7].HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
            t0.Rows[7].Height = CentimetersToPoints(0.45f);


            t0.Cell(1, 2).Merge(t0.Cell(1, 7));
            t0.Cell(1, 3).Merge(t0.Cell(1, 4));
            t0.Cell(2, 1).Merge(t0.Cell(2, 9));
            t0.Cell(3, 2).Merge(t0.Cell(3, 9));
            t0.Cell(4, 2).Merge(t0.Cell(4, 9));
            t0.Cell(6, 2).Merge(t0.Cell(6, 9));
            t0.Cell(7, 2).Merge(t0.Cell(7, 9));

            t0.PreferredWidth = CentimetersToPoints(10.64f);
            t0.TableDirection = Word.WdTableDirection.wdTableDirectionLtr;
            t0.Rows.Alignment = Word.WdRowAlignment.wdAlignRowRight;
            t0.Rows.WrapAroundText = 1;

            SetTableBolders(t0);

            t0.Cell(1, 1).Width = CentimetersToPoints(4.01f);
            t0.Cell(1, 2).Width = CentimetersToPoints(5.40f);
            t0.Cell(1, 3).Width = CentimetersToPoints(1.24f);
            t0.Cell(2, 1).Width = CentimetersToPoints(10.65f);
            t0.Cell(3, 1).Width = CentimetersToPoints(4.01f);
            t0.Cell(3, 2).Width = CentimetersToPoints(6.64f);
            t0.Cell(4, 1).Width = CentimetersToPoints(4.01f);
            t0.Cell(4, 2).Width = CentimetersToPoints(6.64f);
            t0.Cell(6, 1).Width = CentimetersToPoints(4.01f);
            t0.Cell(6, 2).Width = CentimetersToPoints(6.64f);
            t0.Cell(7, 1).Width = CentimetersToPoints(4.01f);
            t0.Cell(7, 2).Width = CentimetersToPoints(6.64f);
            t0.Cell(5, 1).Width = CentimetersToPoints(4.01f);
            t0.Cell(5, 2).Width = CentimetersToPoints(0.83f);
            t0.Cell(5, 3).Width = CentimetersToPoints(0.83f);
            t0.Cell(5, 4).Width = CentimetersToPoints(0.83f);
            t0.Cell(5, 5).Width = CentimetersToPoints(0.83f);
            t0.Cell(5, 6).Width = CentimetersToPoints(0.83f);
            t0.Cell(5, 7).Width = CentimetersToPoints(0.83f);
            t0.Cell(5, 8).Width = CentimetersToPoints(0.83f);
            t0.Cell(5, 9).Width = CentimetersToPoints(0.83f);

            t0.Cell(1, 1).Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            t0.Cell(1, 1).Shading.BackgroundPatternColor = (Word.WdColor)(219 + 0x100 * 229 + 0x10000 * 241);
            t0.Cell(1, 2).Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            t0.Cell(1, 2).Shading.BackgroundPatternColor = (Word.WdColor)(219 + 0x100 * 229 + 0x10000 * 241);
            t0.Cell(1, 3).Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            t0.Cell(1, 3).Shading.BackgroundPatternColor = (Word.WdColor)(219 + 0x100 * 229 + 0x10000 * 241);
            t0.Cell(3, 1).Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            t0.Cell(3, 1).Shading.BackgroundPatternColor = (Word.WdColor)(219 + 0x100 * 229 + 0x10000 * 241);
            t0.Cell(4, 1).Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            t0.Cell(4, 1).Shading.BackgroundPatternColor = (Word.WdColor)(219 + 0x100 * 229 + 0x10000 * 241);
            t0.Cell(5, 1).Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            t0.Cell(5, 1).Shading.BackgroundPatternColor = (Word.WdColor)(219 + 0x100 * 229 + 0x10000 * 241);
            t0.Cell(6, 1).Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            t0.Cell(6, 1).Shading.BackgroundPatternColor = (Word.WdColor)(219 + 0x100 * 229 + 0x10000 * 241);
            t0.Cell(7, 1).Shading.Texture = Word.WdTextureIndex.wdTextureNone;
            t0.Cell(7, 1).Shading.BackgroundPatternColor = (Word.WdColor)(219 + 0x100 * 229 + 0x10000 * 241);

            t0.Cell(1, 1).Range.Text = "Tipo documento";
            t0.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            t0.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0;
            t0.Cell(1, 1).Range.ParagraphFormat.SpaceBefore = 0;

            t0.Cell(1, 2).Range.Text = "Offerta Economica";
            t0.Cell(1, 2).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            t0.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0;
            t0.Cell(1, 2).Range.ParagraphFormat.SpaceBefore = 0;

            t0.Cell(1, 3).Range.Text = "OFC";
            t0.Cell(1, 3).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            t0.Cell(1, 3).Range.ParagraphFormat.SpaceAfter = 0;
            t0.Cell(1, 3).Range.ParagraphFormat.SpaceBefore = 0;

            t0.Cell(3, 1).Range.Text = "Data";
            t0.Cell(3, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            t0.Cell(3, 1).Range.ParagraphFormat.SpaceAfter = 0;
            t0.Cell(3, 1).Range.ParagraphFormat.SpaceBefore = 0;

            t0.Cell(4, 1).Range.Text = "Invio per";
            t0.Cell(4, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            t0.Cell(4, 1).Range.ParagraphFormat.SpaceAfter = 0;
            t0.Cell(4, 1).Range.ParagraphFormat.SpaceBefore = 0;

            t0.Cell(5, 1).Range.Text = "Classificazione";
            t0.Cell(5, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            t0.Cell(5, 1).Range.ParagraphFormat.SpaceAfter = 0;
            t0.Cell(5, 1).Range.ParagraphFormat.SpaceBefore = 0;

            t0.Cell(6, 1).Range.Text = "Versione";
            t0.Cell(6, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            t0.Cell(6, 1).Range.ParagraphFormat.SpaceAfter = 0;
            t0.Cell(6, 1).Range.ParagraphFormat.SpaceBefore = 0;

            t0.Cell(7, 1).Range.Text = "Id documento";
            t0.Cell(7, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            t0.Cell(7, 1).Range.ParagraphFormat.SpaceAfter = 0;
            t0.Cell(7, 1).Range.ParagraphFormat.SpaceBefore = 0;

            #endregion



            /*
            object start = 0, end = 0;
            Word.Range rng = document.Range(ref start, ref end);
            rng.SetRange(rng.End, rng.End);

            #region Tabella riferimenti
            Word.Table t1 = document.Tables.Add(rng, 8, 3);
            t1.Range.set_Style("TabellaFirma");
            t1.Range.Font.Size = 9;

            t1.PreferredWidth = CentimetersToPoints(17.5f);
            t1.TableDirection = Word.WdTableDirection.wdTableDirectionLtr;
            
            t1.Columns[1].Width = CentimetersToPoints(7.64f) ;
            t1.Columns[2].Width = CentimetersToPoints(1.15f);
            t1.Columns[3].Width = CentimetersToPoints(8.72f);

            foreach (var row  in t1.Rows)
            {
                ((Word.Row)row).HeightRule = Word.WdRowHeightRule.wdRowHeightExactly;
                ((Word.Row)row).Height = CentimetersToPoints(0.4f);
            }

            t1.Cell(1, 1).Range.Text = "Rif. Vs. richiesta" ;
            t1.Cell(2, 1).Range.Text = "Vs. prot. N°";
            t1.Cell(3, 1).Range.Text = "Del";
            document.Range(t1.Cell(1, 1).Range.Start, t1.Cell(3, 1).Range.End).Select();
            Selection.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            Selection.Range.Borders.OutsideColor = (Word.WdColor)(51 + 0x100 * 51 + 0x10000 * 153);

            t1.Cell(4, 1).Range.Text = "Rif. Ns. Off. Tecnica";
            t1.Cell(5, 1).Range.Text = "Ns. protocollo N°";
            t1.Cell(6, 1).Range.Text = "Del";
            document.Range(t1.Cell(4, 1).Range.Start, t1.Cell(6, 1).Range.End).Select();
            Selection.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            Selection.Range.Borders.OutsideColor = (Word.WdColor)(51 + 0x100 * 51 + 0x10000 * 153);

            t1.Cell(7, 1).Range.Text = "Rif. Persona";
            t1.Cell(8, 1).Range.Text = "Progetto";
            document.Range(t1.Cell(7, 1).Range.Start, t1.Cell(8, 1).Range.End).Select();
            Selection.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            Selection.Range.Borders.OutsideColor = (Word.WdColor)(51 + 0x100 * 51 + 0x10000 * 153);

            t1.Cell(1, 3).Range.Text = "Destinatario";
            t1.Cell(2, 3).Range.Text = "Spett.";
            t1.Cell(3, 3).Range.Text = "";
            t1.Cell(4, 3).Range.Text = "";
            t1.Cell(5, 3).Range.Text = "";
            t1.Cell(6, 3).Range.Text = "";
            document.Range(t1.Cell(1, 3).Range.Start, t1.Cell(6, 3).Range.End).Select();
            Selection.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            Selection.Range.Borders.OutsideColor = (Word.WdColor)(51 + 0x100 * 51 + 0x10000 * 153);

            t1.Cell(7, 3).Range.Text = "C.A.";
            t1.Cell(8, 3).Range.Text = "P.C.";
            document.Range(t1.Cell(7, 3).Range.Start, t1.Cell(8, 3).Range.End).Select();
            Selection.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            Selection.Range.Borders.OutsideColor = (Word.WdColor)(51 + 0x100 * 51 + 0x10000 * 153);
            #endregion

            rng.SetRange(t1.Range.End, t1.Range.End);

            rng.InsertParagraphAfter();

            rng.SetRange(rng.End, rng.End);

            Word.Paragraph p1 = document.Paragraphs.Add(rng);
            p1.Range.Text = "Con la presente Vi rimettiamo offerta economica per i prodotti / servizi elencati nella sezione “Configurazione offerta”, eventualmente descritti dettagliatamente nell’apposita offerta tecnica indicata in calce.";
            p1.Range.set_Style("normale");
            p1.Range.InsertParagraphAfter();


            
            //Word.Paragraph notaTecnica = document.Paragraphs.Add(System.Reflection.Missing.Value);
            Word.Paragraph notaTecnica = document.Paragraphs.Add(p1.Range);
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

            testoElencoGenerale.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            testoElencoGenerale.Range.ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
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

            Word.Paragraph w1 = document.Paragraphs.Add(testoElencoGenerale.Range);
            w1.Range.Text = "\n";
            w1.Range.InsertParagraphAfter();

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

            //document.Protect(Word.WdProtectionType.wdAllowOnlyReading, ref noReset, ref password, ref useIRM, ref enforceStyleLock);

            //notaTecnica.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);
            */
        }
    }
}
