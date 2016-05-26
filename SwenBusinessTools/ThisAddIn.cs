using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace SwenBusinessTools
{
    public partial class ThisAddIn
    {
        private SwenUserControl myUserControl1;

        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //myUserControl1 = new SwenUserControl();
            //myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "SWEN s.r.l.");
            //myCustomTaskPane.Visible = true;           
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Codice generato da VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
