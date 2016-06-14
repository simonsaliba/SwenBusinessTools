namespace SwenBusinessTools
{
    partial class UserControlProgetto
    {
        /// <summary> 
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare 
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSalva = new System.Windows.Forms.Button();
            this.btnSelezionaCliente = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.txtCliente = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtResponsabileProgetto = new System.Windows.Forms.TextBox();
            this.cboBusinessUnit = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cboTipoProgetto = new System.Windows.Forms.ComboBox();
            this.lblDIP = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtCodiceProgetto = new System.Windows.Forms.TextBox();
            this.txtDescrizione = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dtpDataFine = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpDataInizio = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnSalva
            // 
            this.btnSalva.Location = new System.Drawing.Point(217, 265);
            this.btnSalva.Name = "btnSalva";
            this.btnSalva.Size = new System.Drawing.Size(83, 29);
            this.btnSalva.TabIndex = 37;
            this.btnSalva.Text = "Salva";
            this.btnSalva.UseVisualStyleBackColor = true;
            // 
            // btnSelezionaCliente
            // 
            this.btnSelezionaCliente.Location = new System.Drawing.Point(362, 226);
            this.btnSelezionaCliente.Name = "btnSelezionaCliente";
            this.btnSelezionaCliente.Size = new System.Drawing.Size(27, 23);
            this.btnSelezionaCliente.TabIndex = 36;
            this.btnSelezionaCliente.Text = "...";
            this.btnSelezionaCliente.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(18, 230);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(39, 13);
            this.label7.TabIndex = 35;
            this.label7.Text = "Cliente";
            // 
            // txtCliente
            // 
            this.txtCliente.Location = new System.Drawing.Point(82, 227);
            this.txtCliente.Name = "txtCliente";
            this.txtCliente.Size = new System.Drawing.Size(273, 20);
            this.txtCliente.TabIndex = 34;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(18, 204);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(130, 13);
            this.label6.TabIndex = 33;
            this.label6.Text = "Responsabile del progetto";
            // 
            // txtResponsabileProgetto
            // 
            this.txtResponsabileProgetto.Location = new System.Drawing.Point(154, 201);
            this.txtResponsabileProgetto.Name = "txtResponsabileProgetto";
            this.txtResponsabileProgetto.Size = new System.Drawing.Size(235, 20);
            this.txtResponsabileProgetto.TabIndex = 32;
            // 
            // cboBusinessUnit
            // 
            this.cboBusinessUnit.FormattingEnabled = true;
            this.cboBusinessUnit.Items.AddRange(new object[] {
            "",
            "ACS",
            "DEV",
            "ENT",
            "FOR",
            "RES",
            "INT"});
            this.cboBusinessUnit.Location = new System.Drawing.Point(106, 174);
            this.cboBusinessUnit.Name = "cboBusinessUnit";
            this.cboBusinessUnit.Size = new System.Drawing.Size(283, 21);
            this.cboBusinessUnit.TabIndex = 31;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(18, 177);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(69, 13);
            this.label5.TabIndex = 30;
            this.label5.Text = "Business unit";
            // 
            // cboTipoProgetto
            // 
            this.cboTipoProgetto.FormattingEnabled = true;
            this.cboTipoProgetto.Location = new System.Drawing.Point(106, 147);
            this.cboTipoProgetto.Name = "cboTipoProgetto";
            this.cboTipoProgetto.Size = new System.Drawing.Size(283, 21);
            this.cboTipoProgetto.TabIndex = 29;
            // 
            // lblDIP
            // 
            this.lblDIP.AutoSize = true;
            this.lblDIP.Location = new System.Drawing.Point(18, 150);
            this.lblDIP.Name = "lblDIP";
            this.lblDIP.Size = new System.Drawing.Size(70, 13);
            this.lblDIP.TabIndex = 28;
            this.lblDIP.Text = "Tipo progetto";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(18, 124);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(82, 13);
            this.label4.TabIndex = 27;
            this.label4.Text = "Codice progetto";
            // 
            // txtCodiceProgetto
            // 
            this.txtCodiceProgetto.Location = new System.Drawing.Point(106, 121);
            this.txtCodiceProgetto.Name = "txtCodiceProgetto";
            this.txtCodiceProgetto.Size = new System.Drawing.Size(283, 20);
            this.txtCodiceProgetto.TabIndex = 26;
            // 
            // txtDescrizione
            // 
            this.txtDescrizione.Location = new System.Drawing.Point(82, 46);
            this.txtDescrizione.Multiline = true;
            this.txtDescrizione.Name = "txtDescrizione";
            this.txtDescrizione.Size = new System.Drawing.Size(307, 68);
            this.txtDescrizione.TabIndex = 25;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(18, 46);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 13);
            this.label3.TabIndex = 24;
            this.label3.Text = "Descrizione";
            // 
            // dtpDataFine
            // 
            this.dtpDataFine.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataFine.Location = new System.Drawing.Point(288, 18);
            this.dtpDataFine.Name = "dtpDataFine";
            this.dtpDataFine.Size = new System.Drawing.Size(101, 20);
            this.dtpDataFine.TabIndex = 23;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(226, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 13);
            this.label2.TabIndex = 22;
            this.label2.Text = "Data fine";
            // 
            // dtpDataInizio
            // 
            this.dtpDataInizio.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataInizio.Location = new System.Drawing.Point(82, 16);
            this.dtpDataInizio.Name = "dtpDataInizio";
            this.dtpDataInizio.Size = new System.Drawing.Size(101, 20);
            this.dtpDataInizio.TabIndex = 21;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 20;
            this.label1.Text = "Data inizio";
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.Location = new System.Drawing.Point(306, 265);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(83, 29);
            this.btnAnnulla.TabIndex = 19;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // UserControlProgetto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnSalva);
            this.Controls.Add(this.btnSelezionaCliente);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtCliente);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtResponsabileProgetto);
            this.Controls.Add(this.cboBusinessUnit);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cboTipoProgetto);
            this.Controls.Add(this.lblDIP);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtCodiceProgetto);
            this.Controls.Add(this.txtDescrizione);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dtpDataFine);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dtpDataInizio);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAnnulla);
            this.Name = "UserControlProgetto";
            this.Size = new System.Drawing.Size(411, 307);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSalva;
        private System.Windows.Forms.Button btnSelezionaCliente;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtCliente;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtResponsabileProgetto;
        private System.Windows.Forms.ComboBox cboBusinessUnit;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboTipoProgetto;
        private System.Windows.Forms.Label lblDIP;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtCodiceProgetto;
        private System.Windows.Forms.TextBox txtDescrizione;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtpDataFine;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpDataInizio;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnAnnulla;
    }
}
