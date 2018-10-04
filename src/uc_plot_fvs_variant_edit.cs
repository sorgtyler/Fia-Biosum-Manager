using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_plot_fvs_variant_edit.
	/// </summary>
	public class uc_plot_fvs_variant_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.ComboBox cmbVariant;
		private int m_intError=0;
		private System.Windows.Forms.Label lblBiosumPlotId;
		private System.Windows.Forms.Label lblStateCd;
		private System.Windows.Forms.Label lblCountyCd;
		private System.Windows.Forms.Label lblPlot;
		private System.Windows.Forms.Label lblHalfState;
        private ComboBox cmbFvsLocation;
        private Label label7;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

	    private Dictionary<string, List<string>> dictVariantLocations;

	    public ado_data_access m_ado;

	    public uc_plot_fvs_variant_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
		    InitializeVariantLocationDictionary();
		    if (dictVariantLocations.ContainsKey(strVariant))
		    {
		        this.cmbFvsLocation.Items.AddRange(this.dictVariantLocations[this.strVariant].ToArray());
		    }
		}

	    private void InitializeVariantLocationDictionary()
	    {
            m_ado = new ado_data_access();
            dictVariantLocations = new Dictionary<string, List<string>>();
	        using (var conn = new OleDbConnection(m_ado.getMDBConnString(
	            frmMain.g_oEnv.strApplicationDataDirectory.Trim() + frmMain.g_strBiosumDataDir + "\\" +
	            Tables.Reference.DefaultBiosumReferenceDbFile, "", "")))
	        {
                conn.Open();
                m_ado.SqlQueryReader(conn, "SELECT DISTINCT fvs_variant, fvsloccode FROM fiadb_fvs_variant;");
	            if (this.m_ado.m_intError==0)
	            {
	                if (this.m_ado.m_OleDbDataReader.HasRows)
	                {
	                    string variant = "";
	                    string location = "";
	                    while (this.m_ado.m_OleDbDataReader.Read())
	                    {
                            variant = m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim();
	                        if (!dictVariantLocations.ContainsKey(variant))
	                        {
	                            dictVariantLocations.Add(variant, new List<string>());
	                        }
	                        location = m_ado.m_OleDbDataReader["fvsloccode"].ToString().Trim();
	                        dictVariantLocations[variant].Add(location);
	                    }
	                }
	            }
                m_ado.m_OleDbDataReader.Dispose();
	        }
	        m_ado = null;
	    }

	    /// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cmbFvsLocation = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.lblHalfState = new System.Windows.Forms.Label();
            this.lblPlot = new System.Windows.Forms.Label();
            this.lblCountyCd = new System.Windows.Forms.Label();
            this.lblStateCd = new System.Windows.Forms.Label();
            this.cmbVariant = new System.Windows.Forms.ComboBox();
            this.lblBiosumPlotId = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cmbFvsLocation);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.lblHalfState);
            this.groupBox1.Controls.Add(this.lblPlot);
            this.groupBox1.Controls.Add(this.lblCountyCd);
            this.groupBox1.Controls.Add(this.lblStateCd);
            this.groupBox1.Controls.Add(this.cmbVariant);
            this.groupBox1.Controls.Add(this.lblBiosumPlotId);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnOK);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.groupBox1.Size = new System.Drawing.Size(624, 352);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // cmbFvsLocation
            // 
            this.cmbFvsLocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbFvsLocation.Location = new System.Drawing.Point(248, 245);
            this.cmbFvsLocation.Name = "cmbFvsLocation";
            this.cmbFvsLocation.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmbFvsLocation.Size = new System.Drawing.Size(296, 24);
            this.cmbFvsLocation.TabIndex = 18;
            this.cmbFvsLocation.Validating += new System.ComponentModel.CancelEventHandler(this.cmbFvsLocation_Validating);
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(34, 249);
            this.label7.Name = "label7";
            this.label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label7.Size = new System.Drawing.Size(198, 16);
            this.label7.TabIndex = 19;
            this.label7.Text = "FVS Location Code";
            // 
            // lblHalfState
            // 
            this.lblHalfState.BackColor = System.Drawing.Color.White;
            this.lblHalfState.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHalfState.Location = new System.Drawing.Point(248, 180);
            this.lblHalfState.Name = "lblHalfState";
            this.lblHalfState.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblHalfState.Size = new System.Drawing.Size(56, 24);
            this.lblHalfState.TabIndex = 17;
            this.lblHalfState.Text = "0";
            // 
            // lblPlot
            // 
            this.lblPlot.BackColor = System.Drawing.Color.White;
            this.lblPlot.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPlot.Location = new System.Drawing.Point(248, 149);
            this.lblPlot.Name = "lblPlot";
            this.lblPlot.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblPlot.Size = new System.Drawing.Size(104, 24);
            this.lblPlot.TabIndex = 16;
            this.lblPlot.Text = "0";
            // 
            // lblCountyCd
            // 
            this.lblCountyCd.BackColor = System.Drawing.Color.White;
            this.lblCountyCd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCountyCd.Location = new System.Drawing.Point(248, 118);
            this.lblCountyCd.Name = "lblCountyCd";
            this.lblCountyCd.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblCountyCd.Size = new System.Drawing.Size(56, 24);
            this.lblCountyCd.TabIndex = 15;
            this.lblCountyCd.Text = "0";
            // 
            // lblStateCd
            // 
            this.lblStateCd.BackColor = System.Drawing.Color.White;
            this.lblStateCd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStateCd.Location = new System.Drawing.Point(248, 87);
            this.lblStateCd.Name = "lblStateCd";
            this.lblStateCd.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblStateCd.Size = new System.Drawing.Size(40, 24);
            this.lblStateCd.TabIndex = 14;
            this.lblStateCd.Text = "0";
            // 
            // cmbVariant
            // 
            this.cmbVariant.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbVariant.Items.AddRange(new object[] {
            "AK - SouthEast Alaska, Coastal BC",
            "BM - Blue Mountains",
            "CA - Inland CA, Southern Cascades ",
            "CI - Central Idaho",
            "CR - Central Rockies",
            "CS - Central States",
            "EC - Eastside Cascades",
            "EM - Eastern Montana",
            "KT - Kootenai/Kaniksu/Tally Lake",
            "LS - Lake States",
            "NC - Northern California (Klamath Mountains)",
            "NE - Northeast",
            "NI - Northern Idaho (Inland Empire)",
            "PN - Pacific Northwest Coast",
            "SE - Southeast",
            "SN - Southhern",
            "SO - South Central Oregon, Northeast California",
            "TT - Tetons",
            "UT - Utah",
            "WC - Weside Cascades",
            "WS - Westside Sierra Nevada",
            ""});
            this.cmbVariant.Location = new System.Drawing.Point(248, 211);
            this.cmbVariant.Name = "cmbVariant";
            this.cmbVariant.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmbVariant.Size = new System.Drawing.Size(296, 24);
            this.cmbVariant.TabIndex = 0;
            this.cmbVariant.TextChanged += new System.EventHandler(this.cmbVariant_TextChanged);
            this.cmbVariant.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVariant_KeyDown);
            this.cmbVariant.Validating += new System.ComponentModel.CancelEventHandler(this.cmbVariant_Validating);
            // 
            // lblBiosumPlotId
            // 
            this.lblBiosumPlotId.BackColor = System.Drawing.Color.White;
            this.lblBiosumPlotId.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBiosumPlotId.Location = new System.Drawing.Point(248, 56);
            this.lblBiosumPlotId.Name = "lblBiosumPlotId";
            this.lblBiosumPlotId.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblBiosumPlotId.Size = new System.Drawing.Size(248, 24);
            this.lblBiosumPlotId.TabIndex = 13;
            this.lblBiosumPlotId.Text = "0";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(305, 280);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 48);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(217, 280);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(88, 48);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(72, 215);
            this.label6.Name = "label6";
            this.label6.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label6.Size = new System.Drawing.Size(160, 16);
            this.label6.TabIndex = 8;
            this.label6.Text = "FVS Variant";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(72, 183);
            this.label5.Name = "label5";
            this.label5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label5.Size = new System.Drawing.Size(160, 16);
            this.label5.TabIndex = 7;
            this.label5.Text = "Half State";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(72, 151);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label4.Size = new System.Drawing.Size(160, 16);
            this.label4.TabIndex = 6;
            this.label4.Text = "Plot";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(80, 118);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label3.Size = new System.Drawing.Size(152, 16);
            this.label3.TabIndex = 5;
            this.label3.Text = "County";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(72, 91);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(160, 16);
            this.label2.TabIndex = 4;
            this.label2.Text = "State";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(72, 64);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(160, 24);
            this.label1.TabIndex = 3;
            this.label1.Text = "Biosum Plot Id";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 19);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblTitle.Size = new System.Drawing.Size(618, 24);
            this.lblTitle.TabIndex = 2;
            this.lblTitle.Text = "Plot FVS Variant Edit";
            // 
            // uc_plot_fvs_variant_edit
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_plot_fvs_variant_edit";
            this.Size = new System.Drawing.Size(624, 352);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		

		

		private void cmbVariant_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=true;
		}

		

		
		private void val_data()
		{
            this.m_intError = 0;
			if (this.cmbVariant.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Select A Variant!!","FIA Biosum",
					             System.Windows.Forms.MessageBoxButtons.OK,
					             System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				this.cmbVariant.Focus();
				return;
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.Cancel;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.val_data();
			if (this.m_intError==0)
				((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}
		
	
		public string strBiosumPlotId
		{
			set	{ this.lblBiosumPlotId.Text = value; }
		}
		public string strStateCd
		{
			set	{ this.lblStateCd.Text = value; }
		}
		public string strCountyCd
		{
			set	{ this.lblCountyCd.Text = value; }
		}
		public string strPlot
		{
			set	{ this.lblPlot.Text = value; }
		}
		public string strVariant
		{
			set	{ this.cmbVariant.Text = value; }
			get { return cmbVariant.Text.Trim().Length>=2 ? this.cmbVariant.Text.Trim().Substring(0,2).ToString(): ""; }
		}

		public string strFvsLocation
		{
			set	{ this.cmbFvsLocation.Text = value; }
			get { return this.cmbFvsLocation.Text.Trim(); }
		}
		public string strHalfState
		{
			set	{ this.lblHalfState.Text = value; }
		}

        private void cmbVariant_TextChanged(object sender, EventArgs e)
        {
            this.cmbFvsLocation.Items.Clear();
            if (dictVariantLocations.ContainsKey(strVariant))
            {
		        this.cmbFvsLocation.Items.AddRange(this.dictVariantLocations[this.strVariant].ToArray());
            }
        }

        private void cmbFvsLocation_Validating(object sender, CancelEventArgs e)
        {
            if (!String.IsNullOrEmpty(cmbFvsLocation.Text.Trim()))
            {
                int temp;
                //if not integer value
                if (!int.TryParse(cmbFvsLocation.Text, out temp))
                {
                    MessageBox.Show(
                        "Error: Invalid FVS Location Code.\r\n\r\nEnter a positive integer between three and four digits long, or select a value from the list if available.",
                        "FIA Biosum");
                    cmbFvsLocation.Text = "";
                    strFvsLocation = "";
                }
                //If integer value, but not positive and three to four digits long
                else if (int.TryParse(cmbFvsLocation.Text, out temp) && (temp < 100 || temp > 9999))
                {
                    MessageBox.Show(
                        "Error: Invalid FVS Location Code.\r\n\r\nEnter a positive integer code between three and four digits long, or select a value from the list if available.",
                        "FIA Biosum");
                    cmbFvsLocation.Text = "";
                    strFvsLocation = "";
                }
            }
        }

        private void cmbVariant_Validating(object sender, CancelEventArgs e)
        {
            if (cmbVariant.Text.Trim().Length >= 2)
            {
                strVariant = cmbVariant.Text.Trim().Substring(0, 2);
            }
            else
            {
                strVariant = "";
            }
        }
	}
}
