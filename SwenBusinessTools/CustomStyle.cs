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

            styles.Add("CondizioniGenerali");
            styles.Add("TabellaFirma");

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
                    ((Word.Style)style).Font.Color = Word.WdColor.wdColorBlack;
                }

                if (((Word.Style)style).NameLocal.ToLower() == "condizionigenerali")
                {
                    ((Word.Style)style).Font.Bold = 0;
                    ((Word.Style)style).Font.Size = 8f;
                    ((Word.Style)style).Font.Name = "Trebuchet MS";
                    ((Word.Style)style).Font.Color = Word.WdColor.wdColorBlack;
                }

                if (((Word.Style)style).NameLocal.ToLower() == "tabellafirma")
                {
                    ((Word.Style)style).Font.Bold = 1;
                    ((Word.Style)style).Font.Size = 6f;
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
    }
}
