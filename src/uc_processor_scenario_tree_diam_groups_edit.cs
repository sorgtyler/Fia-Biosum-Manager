using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_tree_diam_groups_edit.
	/// </summary>
	public class uc_processor_scenario_tree_diam_groups_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Label lblID;
		private System.Windows.Forms.TextBox txtMin;
		private System.Windows.Forms.TextBox txtMax;
		private FIA_Biosum_Manager.txtNumeric m_txtMinDiam;
		private FIA_Biosum_Manager.txtNumeric m_txtMaxDiam;
		private int m_intError=0;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        /// <summary>
        /// user class for adding and editing tree diameter groups
        /// </summary>
		public uc_processor_scenario_tree_diam_groups_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_txtMinDiam = new txtNumeric(3,1);
			this.groupBox1.Controls.Add(this.m_txtMinDiam);
			this.m_txtMinDiam.Size = this.txtMin.Size;
			this.m_txtMinDiam.Location = this.txtMin.Location;
			this.m_txtMinDiam.Font = this.txtMin.Font;
			this.m_txtMinDiam.TabIndex = this.txtMin.TabIndex;
			this.m_txtMinDiam.Enabled=true;
			this.m_txtMinDiam.ReadOnly=false;
			this.m_txtMinDiam.bEdit = true;
			this.txtMin.Visible=false;
			this.m_txtMinDiam.Visible=true;


			
			this.m_txtMaxDiam = new txtNumeric(3,1);
			this.groupBox1.Controls.Add(this.m_txtMaxDiam);
			this.m_txtMaxDiam.Size = this.txtMax.Size;
			this.m_txtMaxDiam.Location = this.txtMax.Location;
			this.m_txtMaxDiam.Font = this.txtMax.Font;
			this.m_txtMaxDiam.TabIndex = this.txtMax.TabIndex;
			this.m_txtMaxDiam.Enabled=true;
			this.m_txtMaxDiam.ReadOnly=false;
			this.m_txtMaxDiam.bEdit = true;
			this.txtMax.Visible=false;
			this.m_txtMaxDiam.Visible=true;



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
		public string txtMinimumDiameter
		{
			
		   
		   get {return m_txtMinDiam.Text.Trim();}
		   set
		   {   
			   m_txtMinDiam.Text = value;
		   }
		   
			
		}
		public string txtMaximumDiameter
		{
			
		   
			get {return m_txtMaxDiam.Text.Trim();}
			set 
			{
				m_txtMaxDiam.Text = value;
			   
			}
		   
			
		}
		public string txtGroupId
		{
			get {return this.lblID.Text.Trim();}
			set {this.lblID.Text = value;}
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtMax = new System.Windows.Forms.TextBox();
			this.txtMin = new System.Windows.Forms.TextBox();
			this.lblID = new System.Windows.Forms.Label();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOK = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtMax);
			this.groupBox1.Controls.Add(this.txtMin);
			this.groupBox1.Controls.Add(this.lblID);
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.btnOK);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(480, 304);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// txtMax
			// 
			this.txtMax.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtMax.Location = new System.Drawing.Point(280, 168);
			this.txtMax.Name = "txtMax";
			this.txtMax.Size = new System.Drawing.Size(160, 26);
			this.txtMax.TabIndex = 6;
			this.txtMax.Text = "";
			// 
			// txtMin
			// 
			this.txtMin.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtMin.Location = new System.Drawing.Point(280, 120);
			this.txtMin.Name = "txtMin";
			this.txtMin.Size = new System.Drawing.Size(160, 26);
			this.txtMin.TabIndex = 4;
			this.txtMin.Text = "";
			// 
			// lblID
			// 
			this.lblID.BackColor = System.Drawing.Color.White;
			this.lblID.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblID.Location = new System.Drawing.Point(280, 81);
			this.lblID.Name = "lblID";
			this.lblID.Size = new System.Drawing.Size(32, 24);
			this.lblID.TabIndex = 2;
			this.lblID.Text = "1";
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(232, 224);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(88, 48);
			this.btnCancel.TabIndex = 8;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(144, 224);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(88, 48);
			this.btnOK.TabIndex = 7;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(81, 176);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(184, 32);
			this.label3.TabIndex = 5;
			this.label3.Text = "Maximum Diameter:";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(81, 128);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(184, 32);
			this.label2.TabIndex = 3;
			this.label2.Text = "Minimum Diameter:";
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(161, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(104, 32);
			this.label1.TabIndex = 1;
			this.label1.Text = "Group ID:";
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(474, 32);
			this.lblTitle.TabIndex = 0;
			this.lblTitle.Text = "Tree Diameter Groups Edit";
			// 
            // uc_processor_scenario_tree_diam_groups_edit
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_processor_scenario_tree_diam_groups_edit";
			this.Size = new System.Drawing.Size(480, 304);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.val_input();
			if (this.m_intError==0)
    			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.Cancel;
		}
		private void val_input()
		{
			this.m_intError=0;
			if (this.m_txtMinDiam.Text.Trim().Length < 2)
			{
				this.m_intError=-1;
				MessageBox.Show("Enter a value for minimum diameter","Tree Diameter Groups", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_txtMinDiam.Focus();
				return;
			}

			if (this.m_txtMaxDiam.Text.Trim().Length < 2)
			{
				this.m_intError=-1;
				MessageBox.Show("Enter a value for maximum diameter","Tree Diameter Groups", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_txtMaxDiam.Focus();
				return;
			}
			if (Convert.ToDouble(this.m_txtMaxDiam.Text.Trim()) < Convert.ToDouble(this.m_txtMinDiam.Text.Trim()))
			{
				this.m_intError=-1;
				MessageBox.Show("Minimum value cannot be greater than maximum value","Tree Diameter Groups", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_txtMinDiam.Focus();
				return;
			}

			if (this.m_txtMaxDiam.Text.Trim() == this.m_txtMinDiam.Text.Trim())
			{
				this.m_intError=-1;
				MessageBox.Show("Minimum and maximum values cannot be equal","Tree Diameter Groups", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_txtMinDiam.Focus();
				return;
			}


		}
	}
}
