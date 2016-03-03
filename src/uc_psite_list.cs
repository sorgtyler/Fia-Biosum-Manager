using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_psite_list.
	/// </summary>
	public class uc_psite_list : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.ListView lstPSite;
		private System.Windows.Forms.Button btnDelete;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnClear;
		private System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnNew;
		private System.Windows.Forms.Button btnEdit;
		private System.Windows.Forms.Button btnClose;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_psite_list()
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
			this.lstPSite = new System.Windows.Forms.ListView();
			this.btnDelete = new System.Windows.Forms.Button();
			this.btnHelp = new System.Windows.Forms.Button();
			this.btnClear = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnSave = new System.Windows.Forms.Button();
			this.btnNew = new System.Windows.Forms.Button();
			this.btnEdit = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnDelete);
			this.groupBox1.Controls.Add(this.btnHelp);
			this.groupBox1.Controls.Add(this.btnClear);
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.btnSave);
			this.groupBox1.Controls.Add(this.btnNew);
			this.groupBox1.Controls.Add(this.btnEdit);
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.lstPSite);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(720, 512);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(714, 24);
			this.lblTitle.TabIndex = 27;
			this.lblTitle.Text = "Processing Sites List";
			// 
			// lstPSite
			// 
			this.lstPSite.FullRowSelect = true;
			this.lstPSite.GridLines = true;
			this.lstPSite.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lstPSite.HideSelection = false;
			this.lstPSite.Location = new System.Drawing.Point(16, 48);
			this.lstPSite.MultiSelect = false;
			this.lstPSite.Name = "lstPSite";
			this.lstPSite.Size = new System.Drawing.Size(688, 352);
			this.lstPSite.TabIndex = 28;
			this.lstPSite.View = System.Windows.Forms.View.Details;
			// 
			// btnDelete
			// 
			this.btnDelete.Enabled = false;
			this.btnDelete.Location = new System.Drawing.Point(304, 416);
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(64, 32);
			this.btnDelete.TabIndex = 36;
			this.btnDelete.Text = "Delete";
			// 
			// btnHelp
			// 
			this.btnHelp.Location = new System.Drawing.Point(8, 464);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(64, 32);
			this.btnHelp.TabIndex = 34;
			this.btnHelp.Text = "Help";
			// 
			// btnClear
			// 
			this.btnClear.Location = new System.Drawing.Point(368, 416);
			this.btnClear.Name = "btnClear";
			this.btnClear.Size = new System.Drawing.Size(64, 32);
			this.btnClear.TabIndex = 33;
			this.btnClear.Text = "Clear All";
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(496, 416);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 32);
			this.btnCancel.TabIndex = 32;
			this.btnCancel.Text = "Cancel";
			// 
			// btnSave
			// 
			this.btnSave.Enabled = false;
			this.btnSave.Location = new System.Drawing.Point(432, 416);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(64, 32);
			this.btnSave.TabIndex = 31;
			this.btnSave.Text = "Save";
			// 
			// btnNew
			// 
			this.btnNew.Location = new System.Drawing.Point(176, 416);
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(64, 32);
			this.btnNew.TabIndex = 29;
			this.btnNew.Text = "New";
			// 
			// btnEdit
			// 
			this.btnEdit.Location = new System.Drawing.Point(240, 416);
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(64, 32);
			this.btnEdit.TabIndex = 30;
			this.btnEdit.Text = "Edit";
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(616, 464);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 35;
			this.btnClose.Text = "Close";
			// 
			// uc_psite_list
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_psite_list";
			this.Size = new System.Drawing.Size(720, 512);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
	}
}
