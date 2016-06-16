using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


namespace SwenBusinessTools
{
    public static class CustomStyle
    {
        public static void SetCustomStyle(Word.Document document)
        {
            var styles = document.Styles;
            try
            {
                styles.Add("CondizioniGenerali");
                styles.Add("CondizioniGeneraliNoBold");
                styles.Add("TabellaFirma");
                styles.Add("TabellaTipoDoc");
            }
            catch { }

            foreach (var style in styles)
            {
                if (((Word.Style)style).NameLocal.ToLower() == "titolo 3")
                {
                    ((Word.Style)style).Font.Bold = 1;
                    ((Word.Style)style).Font.Size = 13f;
                    ((Word.Style)style).Font.Name = "Trebuchet MS";
                    ((Word.Style)style).Font.Color = (Word.WdColor)(54 + 0x100 * 95 + 0x10000 * 145);
                }

                if (((Word.Style)style).NameLocal.ToLower() == "normale")
                {
                    ((Word.Style)style).Font.Bold = 0;
                    ((Word.Style)style).Font.Size = 10f;
                    ((Word.Style)style).Font.Name = "Trebuchet MS";
                    ((Word.Style)style).ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    ((Word.Style)style).ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
                    //test
                    //((Word.Style)style).ParagraphFormat.SpaceAfterAuto = 0;
                    //((Word.Style)style).ParagraphFormat.SpaceBeforeAuto = 0;
                    //((Word.Style)style).ParagraphFormat.FirstLineIndent = -7f;
                    //((Word.Style)style).ParagraphFormat.LeftIndent = 7f;
                    //fine test

                    ((Word.Style)style).Font.Color = Word.WdColor.wdColorBlack;
                }

                if (((Word.Style)style).NameLocal.ToLower() == "condizionigenerali")
                {
                    ((Word.Style)style).Font.Bold = 1;
                    ((Word.Style)style).Font.Size = 8f;
                    ((Word.Style)style).Font.Name = "Trebuchet MS";
                    ((Word.Style)style).ParagraphFormat.SpaceAfter = 0F;
                    ((Word.Style)style).ParagraphFormat.SpaceBefore = 0F;
                    ((Word.Style)style).Font.Color = Word.WdColor.wdColorBlack;
                }

                if (((Word.Style)style).NameLocal.ToLower() == "condizionigeneralinobold")
                {
                    ((Word.Style)style).Font.Bold = 0;
                    ((Word.Style)style).Font.Size = 8f;
                    ((Word.Style)style).Font.Name = "Trebuchet MS";

                    //General
                    ((Word.Style)style).ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    ((Word.Style)style).ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;

                    //Rientri
                    ((Word.Style)style).ParagraphFormat.SpaceAfterAuto = 0;
                    ((Word.Style)style).ParagraphFormat.SpaceBeforeAuto = 0;
                    ((Word.Style)style).ParagraphFormat.FirstLineIndent = -7f;
                    ((Word.Style)style).ParagraphFormat.LeftIndent = 7f;

                    //Spaziatura
                    ((Word.Style)style).ParagraphFormat.SpaceAfter = 0F;
                    ((Word.Style)style).ParagraphFormat.SpaceBefore = 0F;
                    ((Word.Style)style).ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                    ((Word.Style)style).ParagraphFormat.LineSpacing = 13.8f;

                    //Color
                    ((Word.Style)style).Font.Color = Word.WdColor.wdColorBlack;
                }

                if (((Word.Style)style).NameLocal.ToLower() == "tabellatipodoc")
                {
                    ((Word.Style)style).Font.Bold = 1;
                    ((Word.Style)style).Font.Size = 8f;
                    ((Word.Style)style).Font.Name = "Trebuchet MS";
                    ((Word.Style)style).Font.Color = (Word.WdColor)(51 + 0x100 * 51 + 0x10000 * 153);
                }

            }
        }

        public static void SetTableBolders(Word.Table table)
        {
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideColor = (Word.WdColor)(51 + 0x100 * 51 + 0x10000 * 153);
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideColor = (Word.WdColor)(51 + 0x100 * 51 + 0x10000 * 153);
        }

        public static void SetTableBolders(Word.Table table, Word.WdColor color)
        {
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideColor = color;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideColor = color;
        }

        public static Word.WdColor ColorRGB(int red, int green, int blue)
        {
            return (Word.WdColor)(red + 0x100 * green + 0x10000 * blue);
        }
    }
}
