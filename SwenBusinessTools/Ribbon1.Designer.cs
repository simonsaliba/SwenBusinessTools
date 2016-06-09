namespace SwenBusinessTools
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione componenti

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpGenerale = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.grpImpostazioni = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.gryTemplates = this.Factory.CreateRibbonGallery();
            this.btnApri = this.Factory.CreateRibbonButton();
            this.btnSalvaCopia = this.Factory.CreateRibbonButton();
            this.btnSalvaVersione = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.btnChiudiDocumentoAttivo = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpGenerale.SuspendLayout();
            this.grpImpostazioni.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpGenerale);
            this.tab1.Groups.Add(this.grpImpostazioni);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "SWEN Tools";
            this.tab1.Name = "tab1";
            // 
            // grpGenerale
            // 
            this.grpGenerale.Items.Add(this.gryTemplates);
            this.grpGenerale.Items.Add(this.separator1);
            this.grpGenerale.Items.Add(this.btnApri);
            this.grpGenerale.Items.Add(this.separator2);
            this.grpGenerale.Items.Add(this.btnSalvaCopia);
            this.grpGenerale.Items.Add(this.btnSalvaVersione);
            this.grpGenerale.Label = "Generale";
            this.grpGenerale.Name = "grpGenerale";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // grpImpostazioni
            // 
            this.grpImpostazioni.DialogLauncher = ribbonDialogLauncherImpl1;
            this.grpImpostazioni.Items.Add(this.button1);
            this.grpImpostazioni.Items.Add(this.button3);
            this.grpImpostazioni.Label = "Impostazioni";
            this.grpImpostazioni.Name = "grpImpostazioni";
            this.grpImpostazioni.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.grpImpostazioni_DialogLauncherClick);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnChiudiDocumentoAttivo);
            this.group1.Label = "Chiudi";
            this.group1.Name = "group1";
            // 
            // gryTemplates
            // 
            this.gryTemplates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            ribbonDropDownItemImpl1.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl1.Label = "Offerta Economica";
            ribbonDropDownItemImpl1.ScreenTip = "Offerta Economica";
            ribbonDropDownItemImpl1.SuperTip = "Template word di offerta economica";
            this.gryTemplates.Items.Add(ribbonDropDownItemImpl1);
            this.gryTemplates.Label = "New Template";
            this.gryTemplates.Name = "gryTemplates";
            this.gryTemplates.OfficeImageId = "FileNew";
            this.gryTemplates.ScreenTip = "Business Templates";
            this.gryTemplates.ShowImage = true;
            this.gryTemplates.ShowItemLabel = false;
            this.gryTemplates.SuperTip = "Creazione di nuovi documenti di business";
            this.gryTemplates.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.gryTemplates_Click);
            // 
            // btnApri
            // 
            this.btnApri.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnApri.Label = "Apri";
            this.btnApri.Name = "btnApri";
            this.btnApri.OfficeImageId = "FileOpen";
            this.btnApri.ShowImage = true;
            // 
            // btnSalvaCopia
            // 
            this.btnSalvaCopia.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSalvaCopia.Label = "Salva Copia";
            this.btnSalvaCopia.Name = "btnSalvaCopia";
            this.btnSalvaCopia.OfficeImageId = "FileSave";
            this.btnSalvaCopia.ShowImage = true;
            this.btnSalvaCopia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // btnSalvaVersione
            // 
            this.btnSalvaVersione.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSalvaVersione.Label = "Salva Versione";
            this.btnSalvaVersione.Name = "btnSalvaVersione";
            this.btnSalvaVersione.OfficeImageId = "FileSave";
            this.btnSalvaVersione.ShowImage = true;
            this.btnSalvaVersione.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "button1";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Label = "button3";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // btnChiudiDocumentoAttivo
            // 
            this.btnChiudiDocumentoAttivo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnChiudiDocumentoAttivo.Label = "Chiudi Documento Attivo";
            this.btnChiudiDocumentoAttivo.Name = "btnChiudiDocumentoAttivo";
            this.btnChiudiDocumentoAttivo.OfficeImageId = "FileClose";
            this.btnChiudiDocumentoAttivo.ShowImage = true;
            this.btnChiudiDocumentoAttivo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChiudi_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpGenerale.ResumeLayout(false);
            this.grpGenerale.PerformLayout();
            this.grpImpostazioni.ResumeLayout(false);
            this.grpImpostazioni.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpImpostazioni;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpGenerale;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSalvaCopia;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApri;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSalvaVersione;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChiudiDocumentoAttivo;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gryTemplates;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
