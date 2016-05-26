﻿namespace SwenBusinessTools
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpGenerale = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.grpImpostazioni = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.gallery1 = this.Factory.CreateRibbonGallery();
            this.button12 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.btnApri = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
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
            this.grpGenerale.Items.Add(this.gallery1);
            this.grpGenerale.Items.Add(this.separator1);
            this.grpGenerale.Items.Add(this.btnApri);
            this.grpGenerale.Items.Add(this.separator2);
            this.grpGenerale.Items.Add(this.button2);
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
            // gallery1
            // 
            this.gallery1.Buttons.Add(this.button12);
            this.gallery1.Buttons.Add(this.button13);
            this.gallery1.ColumnCount = 5;
            this.gallery1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.gallery1.Description = "Swen Templates";
            ribbonDropDownItemImpl1.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl1.Label = "Offerta Economica";
            ribbonDropDownItemImpl1.ScreenTip = "Offerta Economica";
            ribbonDropDownItemImpl1.SuperTip = "Template dell\'offerta economica Offerta Economica";
            ribbonDropDownItemImpl1.Tag = "";
            ribbonDropDownItemImpl2.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl2.Label = "Fattura 1";
            ribbonDropDownItemImpl2.ScreenTip = "Fattura";
            ribbonDropDownItemImpl2.SuperTip = "Fattura 1";
            ribbonDropDownItemImpl3.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl3.Label = "Fattura 2";
            ribbonDropDownItemImpl4.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl4.Label = "Item3";
            ribbonDropDownItemImpl5.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl5.Label = "Item4";
            ribbonDropDownItemImpl6.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl6.Label = "Item5";
            ribbonDropDownItemImpl7.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl7.Label = "Item6";
            ribbonDropDownItemImpl8.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl8.Label = "Item7";
            ribbonDropDownItemImpl9.Image = global::SwenBusinessTools.Properties.Resources.preview_offerta_economica;
            ribbonDropDownItemImpl9.Label = "Item8";
            this.gallery1.Items.Add(ribbonDropDownItemImpl1);
            this.gallery1.Items.Add(ribbonDropDownItemImpl2);
            this.gallery1.Items.Add(ribbonDropDownItemImpl3);
            this.gallery1.Items.Add(ribbonDropDownItemImpl4);
            this.gallery1.Items.Add(ribbonDropDownItemImpl5);
            this.gallery1.Items.Add(ribbonDropDownItemImpl6);
            this.gallery1.Items.Add(ribbonDropDownItemImpl7);
            this.gallery1.Items.Add(ribbonDropDownItemImpl8);
            this.gallery1.Items.Add(ribbonDropDownItemImpl9);
            this.gallery1.Label = "Nuovo Template";
            this.gallery1.Name = "gallery1";
            this.gallery1.OfficeImageId = "FileNew";
            this.gallery1.RowCount = 2;
            this.gallery1.ScreenTip = "Templates";
            this.gallery1.ShowImage = true;
            this.gallery1.ShowItemLabel = false;
            this.gallery1.ShowItemSelection = true;
            this.gallery1.SuperTip = "Swen Templates per la creazione di nuovi documenti di business";
            // 
            // button12
            // 
            this.button12.Label = "button12";
            this.button12.Name = "button12";
            // 
            // button13
            // 
            this.button13.Label = "button13";
            this.button13.Name = "button13";
            // 
            // btnApri
            // 
            this.btnApri.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnApri.Label = "Apri";
            this.btnApri.Name = "btnApri";
            this.btnApri.OfficeImageId = "FileOpen";
            this.btnApri.ShowImage = true;
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Label = "Salva Copia";
            this.button2.Name = "button2";
            this.button2.OfficeImageId = "FileSave";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        private Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApri;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSalvaVersione;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChiudiDocumentoAttivo;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}