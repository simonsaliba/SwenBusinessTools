namespace SwenBusinessTools
{
    partial class frmAggiungiProgetto
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.userControlProgetto1 = new SwenBusinessTools.UserControlProgetto();
            this.SuspendLayout();
            // 
            // userControlProgetto1
            // 
            this.userControlProgetto1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.userControlProgetto1.Location = new System.Drawing.Point(0, 0);
            this.userControlProgetto1.Name = "userControlProgetto1";
            this.userControlProgetto1.Size = new System.Drawing.Size(416, 305);
            this.userControlProgetto1.TabIndex = 0;
            // 
            // frmAggiungiProgetto
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(416, 305);
            this.Controls.Add(this.userControlProgetto1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmAggiungiProgetto";
            this.ShowInTaskbar = false;
            this.Text = "Dati Progetto";
            this.ResumeLayout(false);

        }

        #endregion

        private UserControlProgetto userControlProgetto1;
    }
}