using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            txtACC.Enabled = chkACC.Checked;
        }

        private void chkSER_CheckedChanged(object sender, EventArgs e)
        {
            txtSER.Enabled = chkSER.Checked;
        }

        private void wizardOffertaEconomica_Finished(object sender, EventArgs e)
        {
            //todo: aggiundere il codice per la generazione del template
            MessageBox.Show("Genrazione tdocumento è avvenuta con successo", this.Name);
        }
    }
}
