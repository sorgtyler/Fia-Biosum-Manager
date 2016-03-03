using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_project_document_links_edit.
	/// </summary>
	public class uc_project_document_links_edit : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox grpboxSharedPrivate;
		public System.Windows.Forms.CheckBox chkShared;
		public System.Windows.Forms.CheckBox chkPrivate;
		private System.Windows.Forms.GroupBox grpboxDocument;
		private System.Windows.Forms.GroupBox grpboxCategory;
		public System.Windows.Forms.ComboBox cmbCategory;
		private System.Windows.Forms.GroupBox grpboxDescription;
		public System.Windows.Forms.TextBox txtDescription;
		public System.Windows.Forms.Label lblMsg;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnDocument;
		public string m_strSharedPrivate;
		public System.Windows.Forms.TextBox txtDocument;
	
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_project_document_links_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.grpboxSharedPrivate.Left = 4;
			this.grpboxDescription.Left = 4;
			this.grpboxDocument.Left = 4;
			this.grpboxCategory.Left = 4;
			this.lblMsg.Left = 4;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_project_document_links_edit));
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOK = new System.Windows.Forms.Button();
			this.lblMsg = new System.Windows.Forms.Label();
			this.grpboxDescription = new System.Windows.Forms.GroupBox();
			this.txtDescription = new System.Windows.Forms.TextBox();
			this.grpboxCategory = new System.Windows.Forms.GroupBox();
			this.cmbCategory = new System.Windows.Forms.ComboBox();
			this.grpboxDocument = new System.Windows.Forms.GroupBox();
			this.btnDocument = new System.Windows.Forms.Button();
			this.grpboxSharedPrivate = new System.Windows.Forms.GroupBox();
			this.chkPrivate = new System.Windows.Forms.CheckBox();
			this.chkShared = new System.Windows.Forms.CheckBox();
			this.txtDocument = new System.Windows.Forms.TextBox();
			this.groupBox1.SuspendLayout();
			this.grpboxDescription.SuspendLayout();
			this.grpboxCategory.SuspendLayout();
			this.grpboxDocument.SuspendLayout();
			this.grpboxSharedPrivate.SuspendLayout();
			this.SuspendLayout();
			// 
			// lblTitle
			// 
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(8, 14);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(448, 24);
			this.lblTitle.TabIndex = 31;
			this.lblTitle.Text = "Project Document Links Depository Edit";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.btnOK);
			this.groupBox1.Controls.Add(this.lblMsg);
			this.groupBox1.Controls.Add(this.grpboxDescription);
			this.groupBox1.Controls.Add(this.grpboxCategory);
			this.groupBox1.Controls.Add(this.grpboxDocument);
			this.groupBox1.Controls.Add(this.grpboxSharedPrivate);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(544, 464);
			this.groupBox1.TabIndex = 32;
			this.groupBox1.TabStop = false;
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			// 
			// btnCancel
			// 
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(280, 400);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(80, 48);
			this.btnCancel.TabIndex = 38;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(168, 400);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(80, 48);
			this.btnOK.TabIndex = 37;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// lblMsg
			// 
			this.lblMsg.Location = new System.Drawing.Point(24, 368);
			this.lblMsg.Name = "lblMsg";
			this.lblMsg.Size = new System.Drawing.Size(504, 16);
			this.lblMsg.TabIndex = 36;
			// 
			// grpboxDescription
			// 
			this.grpboxDescription.Controls.Add(this.txtDescription);
			this.grpboxDescription.Location = new System.Drawing.Point(24, 240);
			this.grpboxDescription.Name = "grpboxDescription";
			this.grpboxDescription.Size = new System.Drawing.Size(504, 120);
			this.grpboxDescription.TabIndex = 35;
			this.grpboxDescription.TabStop = false;
			this.grpboxDescription.Text = "Description";
			// 
			// txtDescription
			// 
			this.txtDescription.Location = new System.Drawing.Point(8, 24);
			this.txtDescription.Multiline = true;
			this.txtDescription.Name = "txtDescription";
			this.txtDescription.Size = new System.Drawing.Size(488, 80);
			this.txtDescription.TabIndex = 0;
			this.txtDescription.Text = "";
			// 
			// grpboxCategory
			// 
			this.grpboxCategory.Controls.Add(this.cmbCategory);
			this.grpboxCategory.Location = new System.Drawing.Point(16, 184);
			this.grpboxCategory.Name = "grpboxCategory";
			this.grpboxCategory.Size = new System.Drawing.Size(512, 48);
			this.grpboxCategory.TabIndex = 34;
			this.grpboxCategory.TabStop = false;
			this.grpboxCategory.Text = "Category";
			// 
			// cmbCategory
			// 
			this.cmbCategory.Items.AddRange(new object[] {
															 "GIS",
															 "FVS",
															 "Core Analysis",
															 "Processor",
															 "FRCS",
															 "Other"});
			this.cmbCategory.Location = new System.Drawing.Point(8, 17);
			this.cmbCategory.Name = "cmbCategory";
			this.cmbCategory.Size = new System.Drawing.Size(496, 21);
			this.cmbCategory.TabIndex = 0;
			this.cmbCategory.Text = "Core Analysis";
			// 
			// grpboxDocument
			// 
			this.grpboxDocument.Controls.Add(this.txtDocument);
			this.grpboxDocument.Controls.Add(this.btnDocument);
			this.grpboxDocument.Location = new System.Drawing.Point(16, 112);
			this.grpboxDocument.Name = "grpboxDocument";
			this.grpboxDocument.Size = new System.Drawing.Size(512, 64);
			this.grpboxDocument.TabIndex = 33;
			this.grpboxDocument.TabStop = false;
			this.grpboxDocument.Text = "Document";
			// 
			// btnDocument
			// 
			this.btnDocument.Image = ((System.Drawing.Image)(resources.GetObject("btnDocument.Image")));
			this.btnDocument.Location = new System.Drawing.Point(471, 19);
			this.btnDocument.Name = "btnDocument";
			this.btnDocument.Size = new System.Drawing.Size(32, 32);
			this.btnDocument.TabIndex = 3;
			this.btnDocument.Click += new System.EventHandler(this.btnDocument_Click);
			// 
			// grpboxSharedPrivate
			// 
			this.grpboxSharedPrivate.Controls.Add(this.chkPrivate);
			this.grpboxSharedPrivate.Controls.Add(this.chkShared);
			this.grpboxSharedPrivate.Location = new System.Drawing.Point(16, 56);
			this.grpboxSharedPrivate.Name = "grpboxSharedPrivate";
			this.grpboxSharedPrivate.Size = new System.Drawing.Size(512, 48);
			this.grpboxSharedPrivate.TabIndex = 32;
			this.grpboxSharedPrivate.TabStop = false;
			this.grpboxSharedPrivate.Text = "Save As";
			// 
			// chkPrivate
			// 
			this.chkPrivate.Location = new System.Drawing.Point(294, 20);
			this.chkPrivate.Name = "chkPrivate";
			this.chkPrivate.Size = new System.Drawing.Size(114, 16);
			this.chkPrivate.TabIndex = 1;
			this.chkPrivate.Text = "Private Document";
			// 
			// chkShared
			// 
			this.chkShared.Location = new System.Drawing.Point(79, 20);
			this.chkShared.Name = "chkShared";
			this.chkShared.Size = new System.Drawing.Size(121, 16);
			this.chkShared.TabIndex = 0;
			this.chkShared.Text = "Shared Document";
			// 
			// txtDocument
			// 
			this.txtDocument.Location = new System.Drawing.Point(8, 24);
			this.txtDocument.Name = "txtDocument";
			this.txtDocument.Size = new System.Drawing.Size(456, 20);
			this.txtDocument.TabIndex = 4;
			this.txtDocument.Text = "";
			// 
			// uc_project_document_links_edit
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_project_document_links_edit";
			this.Size = new System.Drawing.Size(544, 464);
			this.groupBox1.ResumeLayout(false);
			this.grpboxDescription.ResumeLayout(false);
			this.grpboxCategory.ResumeLayout(false);
			this.grpboxDocument.ResumeLayout(false);
			this.grpboxSharedPrivate.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			this.grpboxSharedPrivate.Width = this.groupBox1.Width - 8;
			this.grpboxCategory.Width = this.grpboxSharedPrivate.Width;
			this.grpboxDocument.Width = this.grpboxSharedPrivate.Width;
			this.grpboxDescription.Width = this.grpboxSharedPrivate.Width;

			this.chkPrivate.Left = (int)(this.grpboxSharedPrivate.Width * .60); 
			this.chkShared.Left = (int)(this.grpboxSharedPrivate.Width * .40) - this.chkShared.Width ; 

            this.btnDocument.Left = this.grpboxDocument.Width - this.btnDocument.Width - 8;
			this.txtDocument.Width = this.btnDocument.Left - this.txtDocument.Left - 8;

			this.cmbCategory.Width = this.grpboxCategory.Width - this.cmbCategory.Left - 8;
			this.txtDescription.Width = this.grpboxDescription.Width - this.txtDescription.Left - 8;
            this.lblMsg.Width = this.txtDescription.Width;

			this.btnOK.Left = (int)(this.groupBox1.Width * .50) - this.btnOK.Width;
			this.btnCancel.Left = this.btnOK.Left + this.btnOK.Width;

			this.btnOK.Top = this.groupBox1.Height - this.btnOK.Height - 8;
			this.btnCancel.Top = this.btnOK.Top;

			this.lblMsg.Top = this.btnOK.Top - this.lblMsg.Height - 4;
			this.grpboxDescription.Height = this.lblMsg.Top - this.grpboxDescription.Top - 4;
			this.txtDescription.Height = this.grpboxDescription.Height - this.txtDescription.Top - 8;
			


		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
		   
		}

		private void btnDocument_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Open FIA Biosum Project Access File";
			OpenFileDialog1.Filter = "All Files (*.*) |*.*";
			//OpenFileDialog1.Reset();
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.txtDocument.Text = OpenFileDialog1.FileName.Trim();
				}
			}
			OpenFileDialog1 = null;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			int x=0;
			//validate document file
			if (this.txtDocument.Text.Trim().Length == 0) 
			{
                MessageBox.Show("Enter A Document","Document Link Edit",MessageBoxButtons.OK,MessageBoxIcon.Error);    
				this.txtDocument.Focus();
				return;
			}

			//validate whether this is  a private or shared document
			if (this.chkShared.Checked==false && this.chkPrivate.Checked==false) 
			{
			    MessageBox.Show("Check Private and/or Shared Document","Document Link Edit",MessageBoxButtons.OK,MessageBoxIcon.Error);    
				this.chkShared.Focus();
				return;
			}

			//validate combo text containing category value
			for (x=0;x<= this.cmbCategory.Items.Count - 1; x++)
			{
				if (this.cmbCategory.Text.Trim().ToUpper() ==
					this.cmbCategory.Items[x].ToString().Trim().ToUpper())
				{
					break;
				}
			}
			if ((x > this.cmbCategory.Items.Count - 1))
			{
				MessageBox.Show("Select A Valid Category","Document Link Edit",MessageBoxButtons.OK,MessageBoxIcon.Error);    
				this.cmbCategory.Focus();
				return;
			}

			if (this.chkShared.Checked==true && this.chkPrivate.Checked==true)
				this.m_strSharedPrivate="B";
			else if (this.chkShared.Checked==true)
				this.m_strSharedPrivate="S";
			else if (this.chkPrivate.Checked==true)
				this.m_strSharedPrivate="P";
            
			this.ParentForm.DialogResult = DialogResult.OK;



		}
		public string strOKButtonText
		{
			set	{ this.btnOK.Text  = value; }
			get { return this.btnOK.Text.ToString(); }
		}
		public string strCancelButtonText
		{
			set	{ this.btnCancel.Text  = value; }
			get { return this.btnCancel.Text.ToString(); }
		}
		
	}
}
