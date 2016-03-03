using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_fvs_tree_spc_conversion_edit.
	/// </summary>
	public class uc_fvs_tree_spc_conversion_edit : System.Windows.Forms.UserControl
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
		private System.Windows.Forms.Label lblID;
		private System.Windows.Forms.ComboBox cmbVariant;
		private System.Windows.Forms.TextBox txtSpCd;
		private System.Windows.Forms.TextBox txtCommon;
		private System.Windows.Forms.TextBox txtConvert;
		private System.Windows.Forms.TextBox txtComments;
		private int m_intError=0;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_fvs_tree_spc_conversion_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call

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
			this.lblTitle = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOK = new System.Windows.Forms.Button();
			this.lblID = new System.Windows.Forms.Label();
			this.cmbVariant = new System.Windows.Forms.ComboBox();
			this.txtSpCd = new System.Windows.Forms.TextBox();
			this.txtCommon = new System.Windows.Forms.TextBox();
			this.txtConvert = new System.Windows.Forms.TextBox();
			this.txtComments = new System.Windows.Forms.TextBox();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtComments);
			this.groupBox1.Controls.Add(this.txtConvert);
			this.groupBox1.Controls.Add(this.txtCommon);
			this.groupBox1.Controls.Add(this.txtSpCd);
			this.groupBox1.Controls.Add(this.cmbVariant);
			this.groupBox1.Controls.Add(this.lblID);
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
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(624, 352);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 19);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(618, 24);
			this.lblTitle.TabIndex = 2;
			this.lblTitle.Text = "FVS Tree Species Conversion Edit";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(104, 64);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(32, 24);
			this.label1.TabIndex = 3;
			this.label1.Text = "ID:";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(99, 128);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(160, 16);
			this.label2.TabIndex = 4;
			this.label2.Text = "FIA Tree Species Code";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(99, 156);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(152, 16);
			this.label3.TabIndex = 5;
			this.label3.Text = "Common Name";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(98, 189);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(160, 16);
			this.label4.TabIndex = 6;
			this.label4.Text = "Convert To FIA Species Code";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(100, 228);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(160, 16);
			this.label5.TabIndex = 7;
			this.label5.Text = "Comments";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(100, 96);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(160, 16);
			this.label6.TabIndex = 8;
			this.label6.Text = "FVS Variant";
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
			// lblID
			// 
			this.lblID.BackColor = System.Drawing.Color.White;
			this.lblID.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblID.Location = new System.Drawing.Point(280, 56);
			this.lblID.Name = "lblID";
			this.lblID.Size = new System.Drawing.Size(40, 24);
			this.lblID.TabIndex = 13;
			this.lblID.Text = "0";
			// 
			// cmbVariant
			// 
			this.cmbVariant.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
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
															"SN - Southern",
															"SO - South Central Oregon, Northeast California",
															"TT - Tetons",
															"UT - Utah",
															"WC - Weside Cascades",
															"WS - Westside Sierra Nevada",
															""});
			this.cmbVariant.Location = new System.Drawing.Point(280, 88);
			this.cmbVariant.Name = "cmbVariant";
			this.cmbVariant.Size = new System.Drawing.Size(296, 24);
			this.cmbVariant.TabIndex = 0;
			this.cmbVariant.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVariant_KeyDown);
			// 
			// txtSpCd
			// 
			this.txtSpCd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSpCd.Location = new System.Drawing.Point(280, 120);
			this.txtSpCd.MaxLength = 3;
			this.txtSpCd.Name = "txtSpCd";
			this.txtSpCd.Size = new System.Drawing.Size(56, 23);
			this.txtSpCd.TabIndex = 1;
			this.txtSpCd.Text = "";
			this.txtSpCd.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSpCd_KeyDown);
			this.txtSpCd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSpCd_KeyPress);
			// 
			// txtCommon
			// 
			this.txtCommon.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCommon.Location = new System.Drawing.Point(280, 152);
			this.txtCommon.MaxLength = 50;
			this.txtCommon.Name = "txtCommon";
			this.txtCommon.Size = new System.Drawing.Size(296, 23);
			this.txtCommon.TabIndex = 2;
			this.txtCommon.Text = "";
			// 
			// txtConvert
			// 
			this.txtConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtConvert.Location = new System.Drawing.Point(280, 187);
			this.txtConvert.MaxLength = 3;
			this.txtConvert.Name = "txtConvert";
			this.txtConvert.Size = new System.Drawing.Size(56, 23);
			this.txtConvert.TabIndex = 3;
			this.txtConvert.Text = "";
			this.txtConvert.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtConvert_KeyDown);
			this.txtConvert.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtConvert_KeyPress);
			// 
			// txtComments
			// 
			this.txtComments.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtComments.Location = new System.Drawing.Point(280, 223);
			this.txtComments.MaxLength = 50;
			this.txtComments.Name = "txtComments";
			this.txtComments.Size = new System.Drawing.Size(296, 23);
			this.txtComments.TabIndex = 4;
			this.txtComments.Text = "";
			// 
			// uc_fvs_tree_spc_conversion_edit
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_fvs_tree_spc_conversion_edit";
			this.Size = new System.Drawing.Size(624, 352);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void txtSpCd_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (Char.IsDigit((char)e.KeyValue))
			{
			}
			else
			{
				if (e.KeyCode == Keys.Back)
				{
				}
				else
				{
					e.Handled=true;	
				}
			}
		}

		private void txtConvert_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (Char.IsDigit((char)e.KeyValue))
			{
			}
			else
			{
				if (e.KeyCode == Keys.Back)
				{
				}
				else
				{
					e.Handled=true;	
				}
			}
		}

		private void cmbVariant_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=true;
		}

		private void txtSpCd_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
			
		}

		private void txtConvert_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
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

			if (this.txtSpCd.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Enter A Tree Species!!","FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				this.txtSpCd.Focus();
				return;
			}

			if (this.txtConvert.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Enter A Converted Tree Species!!","FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				this.txtConvert.Focus();
				return;
			}

			if (Convert.ToInt32(this.txtSpCd.Text) < 300)
			{
				
			}


			if (Convert.ToInt32(this.txtSpCd.Text) > 299)
			{
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
	
		public string strId
		{
			set	{ this.lblID.Text = value; }
			get { return this.lblID.Text.ToString().Trim(); }
		}
		public string strSpCd
		{
			set	{ this.txtSpCd.Text = value; }
			get { return this.txtSpCd.Text.ToString(); }
		}
		public string strConvertedSpCd
		{
			set	{ this.txtConvert.Text = value; }
			get { return this.txtConvert.Text.ToString(); }
		}
		public string strCommonName
		{
			set	{ this.txtCommon.Text = value; }
			get { return this.txtCommon.Text.ToString(); }
		}
		public string strVariant
		{
			set	{ this.cmbVariant.Text = value; }
			get { return this.cmbVariant.Text.Trim().Substring(0,2).ToString(); }
		}
		public string strComments
		{
			set	{ this.txtComments.Text = value; }
			get { return this.txtComments.Text.ToString(); }
		}

	}
}
