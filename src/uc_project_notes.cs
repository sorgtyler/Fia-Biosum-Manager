using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_project_notes.
	/// </summary>
	public class uc_project_notes : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.ToolBar tbrProjectNotes;
		private System.Windows.Forms.ToolBarButton btnGlobal;
		private System.Windows.Forms.ToolBarButton btnPrivate;
		private System.Windows.Forms.ToolBarButton btnBoth;
		public System.Windows.Forms.ToolBarButton btnSave;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.Label lblShared;
		public System.Windows.Forms.TextBox txtNotesShared;
		private System.Windows.Forms.Label lblPersonal;
		public System.Windows.Forms.TextBox txtNotesPersonal;
		private System.Windows.Forms.Button btnClose;
		private System.ComponentModel.IContainer components;
		private bool m_bLoadNotes=true;
		private int m_intError=0;

		public uc_project_notes()
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_project_notes));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnClose = new System.Windows.Forms.Button();
			this.lblPersonal = new System.Windows.Forms.Label();
			this.txtNotesPersonal = new System.Windows.Forms.TextBox();
			this.lblShared = new System.Windows.Forms.Label();
			this.txtNotesShared = new System.Windows.Forms.TextBox();
			this.tbrProjectNotes = new System.Windows.Forms.ToolBar();
			this.btnGlobal = new System.Windows.Forms.ToolBarButton();
			this.btnPrivate = new System.Windows.Forms.ToolBarButton();
			this.btnBoth = new System.Windows.Forms.ToolBarButton();
			this.btnSave = new System.Windows.Forms.ToolBarButton();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.lblPersonal);
			this.groupBox1.Controls.Add(this.txtNotesPersonal);
			this.groupBox1.Controls.Add(this.lblShared);
			this.groupBox1.Controls.Add(this.txtNotesShared);
			this.groupBox1.Controls.Add(this.tbrProjectNotes);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(576, 680);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(464, 632);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 37;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// lblPersonal
			// 
			this.lblPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPersonal.Location = new System.Drawing.Point(8, 368);
			this.lblPersonal.Name = "lblPersonal";
			this.lblPersonal.Size = new System.Drawing.Size(80, 16);
			this.lblPersonal.TabIndex = 36;
			this.lblPersonal.Text = "Private";
			// 
			// txtNotesPersonal
			// 
			this.txtNotesPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNotesPersonal.Location = new System.Drawing.Point(5, 392);
			this.txtNotesPersonal.Multiline = true;
			this.txtNotesPersonal.Name = "txtNotesPersonal";
			this.txtNotesPersonal.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtNotesPersonal.Size = new System.Drawing.Size(560, 224);
			this.txtNotesPersonal.TabIndex = 35;
			this.txtNotesPersonal.Text = "";
			this.txtNotesPersonal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotesPersonal_KeyPress);
			// 
			// lblShared
			// 
			this.lblShared.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblShared.Location = new System.Drawing.Point(8, 104);
			this.lblShared.Name = "lblShared";
			this.lblShared.Size = new System.Drawing.Size(80, 16);
			this.lblShared.TabIndex = 34;
			this.lblShared.Text = "Shared";
			// 
			// txtNotesShared
			// 
			this.txtNotesShared.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNotesShared.Location = new System.Drawing.Point(5, 128);
			this.txtNotesShared.Multiline = true;
			this.txtNotesShared.Name = "txtNotesShared";
			this.txtNotesShared.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtNotesShared.Size = new System.Drawing.Size(560, 224);
			this.txtNotesShared.TabIndex = 33;
			this.txtNotesShared.Text = "";
			this.txtNotesShared.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotesShared_KeyPress);
			this.txtNotesShared.TextChanged += new System.EventHandler(this.txtNotes_TextChanged);
			// 
			// tbrProjectNotes
			// 
			this.tbrProjectNotes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.tbrProjectNotes.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
																							   this.btnGlobal,
																							   this.btnPrivate,
																							   this.btnBoth,
																							   this.btnSave});
			this.tbrProjectNotes.ButtonSize = new System.Drawing.Size(50, 36);
			this.tbrProjectNotes.Divider = false;
			this.tbrProjectNotes.DropDownArrows = true;
			this.tbrProjectNotes.ImageList = this.imageList1;
			this.tbrProjectNotes.Location = new System.Drawing.Point(3, 56);
			this.tbrProjectNotes.Name = "tbrProjectNotes";
			this.tbrProjectNotes.ShowToolTips = true;
			this.tbrProjectNotes.Size = new System.Drawing.Size(570, 41);
			this.tbrProjectNotes.TabIndex = 32;
			this.tbrProjectNotes.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tbrProjectNotes_ButtonClick);
			// 
			// btnGlobal
			// 
			this.btnGlobal.ImageIndex = 0;
			this.btnGlobal.Text = "Shared";
			// 
			// btnPrivate
			// 
			this.btnPrivate.ImageIndex = 1;
			this.btnPrivate.Text = "Private";
			// 
			// btnBoth
			// 
			this.btnBoth.ImageIndex = 2;
			this.btnBoth.Pushed = true;
			this.btnBoth.Text = "Both";
			// 
			// btnSave
			// 
			this.btnSave.DropDownMenu = this.contextMenu1;
			this.btnSave.Enabled = false;
			this.btnSave.ImageIndex = 3;
			this.btnSave.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton;
			this.btnSave.Text = "Save";
			// 
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.menuItem1,
																						 this.menuItem2,
																						 this.menuItem3});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "Shared";
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 1;
			this.menuItem2.Text = "Private";
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 2;
			this.menuItem3.Text = "Both";
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(570, 40);
			this.lblTitle.TabIndex = 31;
			this.lblTitle.Text = "Project Notes";
			// 
			// uc_project_notes
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_project_notes";
			this.Size = new System.Drawing.Size(576, 680);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void txtNotes_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;

				if (this.ParentForm.WindowState != System.Windows.Forms.FormWindowState.Minimized) // && 
				{
					if (this.tbrProjectNotes.Buttons[0].Pushed ==true)
					{
						this.txtNotesPersonal.Visible=false;
						this.lblPersonal.Visible=false;
						this.txtNotesShared.Visible=true;
						this.lblShared.Visible=true;
						this.txtNotesShared.Height = this.groupBox1.ClientSize.Height - this.txtNotesShared.Top - this.btnClose.Height - 10; //this.groupBox1.Height - ;
						this.txtNotesShared.Width = this.groupBox1.Width - 10;

					}
					else if (this.tbrProjectNotes.Buttons[1].Pushed==true)
					{
						this.txtNotesShared.Visible=false;
						this.lblShared.Visible=false;
						this.txtNotesPersonal.Visible=true;
						this.lblPersonal.Visible=true;
						this.txtNotesPersonal.Top = this.txtNotesShared.Top;
						this.lblPersonal.Top = this.lblShared.Top;
						this.txtNotesPersonal.Height = this.groupBox1.ClientSize.Height - this.txtNotesShared.Top - this.btnClose.Height - 10; //this.groupBox1.Height - ;
						this.txtNotesPersonal.Width = this.groupBox1.Width - 10;
					}
					else if (this.tbrProjectNotes.Buttons[2].Pushed==true)
					{
						this.txtNotesPersonal.Visible=true;
						this.lblPersonal.Visible=true;
						this.txtNotesShared.Visible=true;
						this.lblShared.Visible=true;

						this.lblShared.Top = this.tbrProjectNotes.Top + 
							this.tbrProjectNotes.Height + 10;
						this.txtNotesShared.Top = this.lblShared.Top + this.lblShared.Height + 10;
						this.txtNotesShared.Height= (int)(this.btnClose.Top / 2) - 
							this.tbrProjectNotes.Height - 
							this.lblPersonal.Height -  this.lblShared.Height -  15;
						this.lblPersonal.Top = this.txtNotesShared.Top + this.txtNotesShared.Height + 10;

						this.txtNotesShared.Height= (int)(this.btnClose.Top / 2) - 
							this.tbrProjectNotes.Height - 
							this.lblPersonal.Height -  this.lblShared.Height - 15; 
						this.txtNotesPersonal.Top = this.lblPersonal.Top + this.lblPersonal.Height + 10;
						this.txtNotesPersonal.Height = this.txtNotesShared.Height;
						this.txtNotesPersonal.Width = this.groupBox1.Width - 10;
						this.txtNotesShared.Width = this.txtNotesPersonal.Width;
					
					}
				}
			}
			catch
			{
			}
		}

		private void tbrProjectNotes_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			try
			{
				switch (e.Button.Text.Trim().ToUpper())
				{
					case "SHARED":
						this.tbrProjectNotes.Buttons[0].Pushed=true;	    
						this.tbrProjectNotes.Buttons[1].Pushed=false;	    
						this.tbrProjectNotes.Buttons[2].Pushed=false;	
					    this.groupBox1_Resize(sender, null);
						break;
					case "PRIVATE":
						this.tbrProjectNotes.Buttons[0].Pushed=false;	    
						this.tbrProjectNotes.Buttons[1].Pushed=true;	    
						this.tbrProjectNotes.Buttons[2].Pushed=false;	 
						 this.groupBox1_Resize(sender, null);

						break;
					case "BOTH":
						this.tbrProjectNotes.Buttons[0].Pushed=false;	    
						this.tbrProjectNotes.Buttons[1].Pushed=false;	    
						this.tbrProjectNotes.Buttons[2].Pushed=true;	
						 this.groupBox1_Resize(sender, null);

						break;
					case "SAVE":
						this.m_intError=0;
						this.savevalues();
						if (this.m_intError==0)
						{
							this.btnSave.Enabled=false;
						}

						break;
				}
			}
			catch
			{
			}
		}
		public void loadnotes()
		{

			//set the buttons and menu options
			if (((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text.Trim().Length > 0 &&
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text.Trim().Length > 0)
			{
				this.tbrProjectNotes.Buttons[0].Enabled=true;
				this.tbrProjectNotes.Buttons[1].Enabled=true;
				this.tbrProjectNotes.Buttons[2].Enabled=true;
				this.menuItem1.Enabled=true;
				this.menuItem2.Enabled=true;
				this.menuItem3.Enabled=true;

			}
			else if (((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text.Trim().Length > 0)
			{
				this.tbrProjectNotes.Buttons[0].Enabled=true;
				this.tbrProjectNotes.Buttons[1].Enabled=false;
				this.tbrProjectNotes.Buttons[2].Enabled=false;
				this.menuItem1.Enabled=true;
				this.menuItem2.Enabled=false;
				this.menuItem3.Enabled=false;
				this.tbrProjectNotes.Buttons[0].Pushed=true;
				this.tbrProjectNotes.Buttons[1].Pushed=false;
				this.tbrProjectNotes.Buttons[2].Pushed=false;

			}
			else if (((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text.Trim().Length > 0)
			{
				this.tbrProjectNotes.Buttons[0].Enabled=false;
				this.tbrProjectNotes.Buttons[1].Enabled=true;
				this.tbrProjectNotes.Buttons[2].Enabled=false;
				this.menuItem1.Enabled=false;
				this.menuItem2.Enabled=true;
				this.menuItem3.Enabled=false;
				this.tbrProjectNotes.Buttons[0].Pushed=false;
				this.tbrProjectNotes.Buttons[1].Pushed=true;
				this.tbrProjectNotes.Buttons[2].Pushed=false;

			}
			else
			{
				this.tbrProjectNotes.Buttons[0].Enabled=false;
				this.tbrProjectNotes.Buttons[1].Enabled=false;
				this.tbrProjectNotes.Buttons[2].Enabled=false;
				this.menuItem1.Enabled=false;
				this.menuItem2.Enabled=false;
				this.menuItem3.Enabled=false;
				
			}

			if (this.m_bLoadNotes==true)
			{
				string strSQL="";
				string strConn="";
				//string str="";
				ado_data_access p_ado = new ado_data_access();
				string strMDB="";
				
				if (((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text.Trim().Length > 0)
				{
					strMDB=((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text.Trim() + "\\personal_project_links_and_notes.mdb";
					strConn=p_ado.getMDBConnString(strMDB,"admin","");
					//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDB + ";User Id=admin;Password=;";
					p_ado.OpenConnection(strConn);	
					if (p_ado.m_intError == 0)
					{
						strSQL = "SELECT TOP 1 * FROM notes;";
						p_ado.SqlQueryReader(p_ado.m_OleDbConnection,strSQL);

						if (p_ado.m_intError != 0)
						{
							p_ado.m_OleDbConnection.Close();
							
						}
					}
					if (p_ado.m_intError == 0)
					{
						while (p_ado.m_OleDbDataReader.Read())
						{
							if (p_ado.m_OleDbDataReader["notes"] != System.DBNull.Value)
							{
								if (p_ado.m_OleDbDataReader["notes"].ToString().Trim().Length > 0)
								{
									this.txtNotesPersonal.Text =  p_ado.m_OleDbDataReader["notes"].ToString();
								}
							}
					
						}
						p_ado.m_OleDbDataReader.Close();
						p_ado.m_OleDbConnection.Close();
					}
				}

			    if (((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text.Trim().Length > 0)
				{
					strMDB=((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text.Trim() + "\\shared_project_links_and_notes.mdb";
					strConn=p_ado.getMDBConnString(strMDB,"admin","");
					//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDB + ";User Id=admin;Password=;";
					p_ado.OpenConnection(strConn);	
					if (p_ado.m_intError == 0)
					{
						strSQL = "SELECT TOP 1 * FROM notes;";
						p_ado.SqlQueryReader(p_ado.m_OleDbConnection,strSQL);

						if (p_ado.m_intError != 0)
						{
							p_ado.m_OleDbConnection.Close();
							
						}
					}
					if (p_ado.m_intError == 0)
					{
						while (p_ado.m_OleDbDataReader.Read())
						{
							if (p_ado.m_OleDbDataReader["notes"] != System.DBNull.Value)
							{
								if (p_ado.m_OleDbDataReader["notes"].ToString().Trim().Length > 0)
								{
									this.txtNotesShared.Text =  p_ado.m_OleDbDataReader["notes"].ToString();
								}
							}
					
						}
						p_ado.m_OleDbDataReader.Close();
						p_ado.m_OleDbConnection.Close();
					}

				}
				this.m_bLoadNotes=false;
				p_ado = null;
			}
			

		}
		public void savevalues()
		{
			if (this.menuItem1.Enabled==true)
			{
				this.saveshared();
			}
			if (this.menuItem2.Enabled==true)
			{
				this.saveprivate();
			}

		}
		private void saveshared()
		{
			ado_data_access p_ado = new ado_data_access();
			string strMDB=((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtShared.Text.Trim() + "\\shared_project_links_and_notes.mdb";
			string strConn=p_ado.getMDBConnString(strMDB,"admin","");
			//string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDB + ";User Id=admin;Password=;";
			string strSQL="";
			p_ado.OpenConnection(strConn);	
			if (p_ado.m_intError == 0)
			{
				try
				{
					string strNotes = this.txtNotesShared.Text;
					strNotes = p_ado.FixString(strNotes,"'","''");
					if ((int)p_ado.getRecordCount(p_ado.m_OleDbConnection,"SELECT COUNT(*) FROM NOTES","NOTES") > 0)
					{
						strSQL = "UPDATE NOTES SET NOTES = '" + strNotes + "'";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
					}
					else
					{
						strSQL = "INSERT INTO NOTES (NOTES) VALUES ('" + strNotes + "')";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
					}
				}
				catch
				{
				}
			}
			p_ado = null;

		}
		private void saveprivate()
		{
			ado_data_access p_ado = new ado_data_access();
			string strMDB=((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtPersonal.Text.Trim() + "\\personal_project_links_and_notes.mdb";
			string strConn = p_ado.getMDBConnString(strMDB,"admin","");
			//string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDB + ";User Id=admin;Password=;";
			string strSQL="";
			p_ado.OpenConnection(strConn);	
			if (p_ado.m_intError == 0)
			{
				try
				{
					string strNotes = this.txtNotesPersonal.Text;
					strNotes =  p_ado.FixString(strNotes,"'","''");
					if ((int)p_ado.getRecordCount(p_ado.m_OleDbConnection,"SELECT COUNT(*) FROM NOTES","NOTES") > 0)
					{
						strSQL = "UPDATE NOTES SET NOTES = '" + strNotes + "'";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
					}
					else
					{
						strSQL = "INSERT INTO NOTES (NOTES) VALUES ('" + strNotes + "')";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
					}
				}
				catch
				{
				}
			}
			p_ado = null;
		}
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			 this.ParentForm.Close();
		}

		private void txtNotesShared_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			 if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void txtNotesPersonal_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			 if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}
	}
}
