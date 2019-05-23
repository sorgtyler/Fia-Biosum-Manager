using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_processor_opcost_settings : UserControl
    {
        private frmDialog _frmDialog = null;
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultProcessorXPSFile;

        public uc_processor_opcost_settings()
        {
            InitializeComponent();


            if (frmMain.g_strOPCOSTDirectory.Trim().Length == 0)
            {
                txtOpcost.Text = frmSettings.GetDefaultOpcostPath();
            }
            else
            {
                if (System.IO.File.Exists(frmMain.g_strOPCOSTDirectory) == true)
                    txtOpcost.Text = frmMain.g_strOPCOSTDirectory;
            }

            if (frmMain.g_strRDirectory.Trim().Length > 0 &&
                System.IO.File.Exists(frmMain.g_strRDirectory) == true) txtRdir.Text = frmMain.g_strRDirectory;

            this.m_oEnv = new env();
        }

        private void btnRdir_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog oDialog = new OpenFileDialog();
            oDialog.Title = "32-bit version of RScript.exe File";
            oDialog.Filter = "RScript File (RScript.EXE) |RScript.EXE";
            DialogResult result = oDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtRdir.Text = oDialog.FileName;
            }
        }
        public frmDialog ReferenceDialog
        {
            get { return _frmDialog; }
            set { _frmDialog = value; }
        }

        private void txtRdir_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            e.Handled = true;
        }

        private void txtOpcost_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            e.Handled = true;
        }

        private void btnOpcost_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog oDialog = new OpenFileDialog();
            oDialog.Title = "OPCOST R File";
            oDialog.Filter = "OPCOST File (*.R) |*.r";
            DialogResult result = oDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.txtOpcost.Text = oDialog.FileName;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (txtRdir.Text.Trim().Length == 0)
            {
                MessageBox.Show("Specify RScript.EXE file", "FIA Biosum");
                return;
            }
            if (txtOpcost.Text.Trim().Length == 0)
            {
                MessageBox.Show("Specify OPCOST R file", "FIA Biosum");
                return;
            }

            if (System.IO.File.Exists(txtRdir.Text) == false)
            {
                MessageBox.Show("Specified RScript.EXE file not found", "FIA Biosum");
                return;
            }

            if (System.IO.File.Exists(txtOpcost.Text) == false)
            {
                MessageBox.Show("Specified OPCOST R file not found", "FIA Biosum");
                return;
            }
            frmMain.g_strOPCOSTDirectory = txtOpcost.Text.Trim();
            frmMain.g_strRDirectory = txtRdir.Text.Trim();

            ReferenceDialog.DialogResult = DialogResult.OK;
            ReferenceDialog.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            ReferenceDialog.DialogResult = DialogResult.Cancel;
            ReferenceDialog.Close();
        }

        private void txtRdir_Enter(object sender, EventArgs e)
        {
            
        }

        private void txtRdir_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete) e.SuppressKeyPress = true;
           
        }

        private void txtOpcost_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete) e.SuppressKeyPress = true;
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "PROCESSOR", "OPCOSTSETTINGS" });
        }

    }
}
