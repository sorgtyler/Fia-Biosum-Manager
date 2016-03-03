using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_datasource.
	/// </summary>
	public class uc_scenario_datasource : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.Label lblTablesRequired;
		public System.Windows.Forms.ListView lstRequiredTables;
		public System.Windows.Forms.GroupBox groupBox1;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public int intError;
		private System.Windows.Forms.Button btnClose;
		public System.Windows.Forms.Button btnCopy;
		public System.Windows.Forms.Label lblTitle;
		public string strError;
		private System.Windows.Forms.Button btnEdit;
		public string strTable;
		const int TABLETYPE = 0;
		const int PATH = 1;
		const int MDBFILE = 2;
		const int FILESTATUS = 3;
		const int TABLE = 4;
		const int TABLESTATUS = 5;
		const int RECORDCOUNT = 6;
		public string m_strRandomPathAndFile = "";
		public System.Windows.Forms.Button btnCheckAll;
		public System.Windows.Forms.Button btnClearAll;
		public int m_intNumberOfValidTables=0;  //MDB file is FOUND and table is FOUND

      

		public uc_scenario_datasource()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.lstRequiredTables.View = System.Windows.Forms.View.Details;
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
		public void InitialSize()
		{
		}
		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.lblTablesRequired = new System.Windows.Forms.Label();
			this.lstRequiredTables = new System.Windows.Forms.ListView();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnEdit = new System.Windows.Forms.Button();
			this.lblTitle = new System.Windows.Forms.Label();
			this.btnCopy = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.btnCheckAll = new System.Windows.Forms.Button();
			this.btnClearAll = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// lblTablesRequired
			// 
			this.lblTablesRequired.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTablesRequired.Location = new System.Drawing.Point(16, 72);
			this.lblTablesRequired.Name = "lblTablesRequired";
			this.lblTablesRequired.Size = new System.Drawing.Size(168, 24);
			this.lblTablesRequired.TabIndex = 0;
			this.lblTablesRequired.Text = "Required Tables";
			// 
			// lstRequiredTables
			// 
			this.lstRequiredTables.CheckBoxes = true;
			this.lstRequiredTables.FullRowSelect = true;
			this.lstRequiredTables.GridLines = true;
			this.lstRequiredTables.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lstRequiredTables.HideSelection = false;
			this.lstRequiredTables.Location = new System.Drawing.Point(24, 104);
			this.lstRequiredTables.MultiSelect = false;
			this.lstRequiredTables.Name = "lstRequiredTables";
			this.lstRequiredTables.Size = new System.Drawing.Size(688, 272);
			this.lstRequiredTables.TabIndex = 1;
			this.lstRequiredTables.View = System.Windows.Forms.View.Details;
			this.lstRequiredTables.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lstRequiredTables_MouseDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnClearAll);
			this.groupBox1.Controls.Add(this.btnCheckAll);
			this.groupBox1.Controls.Add(this.btnEdit);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Controls.Add(this.btnCopy);
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.lblTablesRequired);
			this.groupBox1.Controls.Add(this.lstRequiredTables);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(736, 512);
			this.groupBox1.TabIndex = 2;
			this.groupBox1.TabStop = false;
			// 
			// btnEdit
			// 
			this.btnEdit.Location = new System.Drawing.Point(328, 384);
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(96, 32);
			this.btnEdit.TabIndex = 25;
			this.btnEdit.Text = "Edit";
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(730, 32);
			this.lblTitle.TabIndex = 24;
			this.lblTitle.Text = "Scenario Data Source";
			// 
			// btnCopy
			// 
			this.btnCopy.Location = new System.Drawing.Point(16, 448);
			this.btnCopy.Name = "btnCopy";
			this.btnCopy.Size = new System.Drawing.Size(264, 32);
			this.btnCopy.TabIndex = 5;
			this.btnCopy.Text = "Copy Checked Tables To The Scenario Directory";
			this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(632, 464);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 3;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// btnCheckAll
			// 
			this.btnCheckAll.Location = new System.Drawing.Point(16, 416);
			this.btnCheckAll.Name = "btnCheckAll";
			this.btnCheckAll.Size = new System.Drawing.Size(264, 32);
			this.btnCheckAll.TabIndex = 26;
			this.btnCheckAll.Text = "Check All";
			this.btnCheckAll.Click += new System.EventHandler(this.btnCheckAll_Click);
			// 
			// btnClearAll
			// 
			this.btnClearAll.Location = new System.Drawing.Point(16, 384);
			this.btnClearAll.Name = "btnClearAll";
			this.btnClearAll.Size = new System.Drawing.Size(264, 32);
			this.btnClearAll.TabIndex = 27;
			this.btnClearAll.Text = "Clear All";
			this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
			// 
			// uc_scenario_datasource
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_datasource";
			this.Size = new System.Drawing.Size(736, 512);
			this.Resize += new System.EventHandler(this.uc_scenario_datasource_Resize);
			this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.uc_scenario_datasource_MouseDown);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show(this.lstRequiredTables.Columns[0].Width.ToString());
		}
		
		public void populate_listview_grid()
		{
             		
			string strPathAndFile="";
			string strSQL="";
			string strConn="";

			this.lstRequiredTables.Clear();

            ado_data_access p_ado = new ado_data_access();
 
			this.lstRequiredTables.Columns.Add("Table Type", 50, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Path", 150, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("MDB File", 60, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("File Status", 80, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Table Name", 50, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Table Status", 80, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Record Count", 80, HorizontalAlignment.Left);

			System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
			string strDir = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\core\\db";
			string strFile = "scenario_core_rule_definitions.mdb"; 
			StringBuilder strFullPath = new StringBuilder(strDir);
			strFullPath.Append("\\");
			strFullPath.Append(strFile);
			strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFullPath.ToString() + ";User Id=admin;Password=;";
			oConn.ConnectionString = strConn;
			intError = 0;
			strError = "";
			try
			{
				oConn.Open();
			}
			catch (System.Data.OleDb.OleDbException oleException)
			{
				strError = "Failed to make an oleDb connection with " + strConn;
				MessageBox.Show (strError);
				intError = -1;
				return;
			}
			System.Data.OleDb.OleDbCommand oCommand = oConn.CreateCommand();

			oCommand.CommandText = "select table_type,path,file,table_name from scenario_datasource" + 
				                   " where scenario_id = '" + 
				                           ((frmCoreScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.ToString() + "';";
			try
			{
				System.Data.OleDb.OleDbDataReader oDataReader = oCommand.ExecuteReader();
				int x = 0;
				
				while (oDataReader.Read())
				{
					// Add a ListItem object to the ListView.
					System.Windows.Forms.ListViewItem entryListItem =
						  lstRequiredTables.Items.Add(oDataReader["table_type"].ToString());
					 entryListItem.UseItemStyleForSubItems=false;

					this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["path"].ToString());
					this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["file"].ToString());
					strPathAndFile = oDataReader["path"].ToString().Trim() + "\\" + oDataReader["file"].ToString().Trim();
					if (System.IO.File.Exists(strPathAndFile) == true) 
					{
						ListViewItem.ListViewSubItem FileStatusSubItem = 
							entryListItem.SubItems.Add("Found");
						FileStatusSubItem.ForeColor = System.Drawing.Color.Black;
						FileStatusSubItem.BackColor = System.Drawing.Color.White;
						FileStatusSubItem.Font = new System.Drawing.Font(
							"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_name"].ToString());

						//see if the table exists in the mdb database container
						dao_data_access p_dao = new dao_data_access();
						if (p_dao.TableExists(strPathAndFile,oDataReader["table_name"].ToString().Trim()) == true)
						{
							this.lstRequiredTables.Items[x].SubItems.Add("Found");
                            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strPathAndFile + ";User Id=admin;Password=;";
							strSQL = "select count(*) from " + oDataReader["table_name"].ToString();
							this.lstRequiredTables.Items[x].SubItems.Add(Convert.ToString(p_ado.getRecordCount(strConn,strSQL,oDataReader["table_name"].ToString())));

						}
						else 
						{
							ListViewItem.ListViewSubItem TableStatusSubItem = 
								entryListItem.SubItems.Add("Not Found");
							TableStatusSubItem.ForeColor = System.Drawing.Color.White;
							TableStatusSubItem.BackColor = System.Drawing.Color.Red;
							TableStatusSubItem.Font = new System.Drawing.Font(
								"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
							this.lstRequiredTables.Items[x].SubItems.Add("0");
						}
						p_dao = null;
					}
					else 
					{
						ListViewItem.ListViewSubItem FileStatusSubItem = 
							entryListItem.SubItems.Add("Not Found");
						FileStatusSubItem.ForeColor = System.Drawing.Color.White;
						FileStatusSubItem.BackColor = System.Drawing.Color.Red;
						FileStatusSubItem.Font = new System.Drawing.Font(
							"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_name"].ToString());
						ListViewItem.ListViewSubItem TableStatusSubItem = 
							entryListItem.SubItems.Add("Not Found");
						TableStatusSubItem.ForeColor = System.Drawing.Color.White;
						TableStatusSubItem.BackColor = System.Drawing.Color.Red;
						TableStatusSubItem.Font = new System.Drawing.Font(
							"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
						this.lstRequiredTables.Items[x].SubItems.Add("0");
					}
					x++;
				}
				oDataReader.Close();
			}
			catch
			{
				intError = -1;
				strError = "The Query Command " + oCommand.CommandText.ToString() + " Failed";
				MessageBox.Show(strError);
				oConn.Close();
				p_ado= null;
				return;
			}
			
			oConn.Close();
			

			this.lstRequiredTables.Columns[TABLETYPE].Width = -1;
			this.lstRequiredTables.Columns[PATH].Width = -1;
			this.lstRequiredTables.Columns[MDBFILE].Width = -1;

			this.lstRequiredTables.Columns[TABLE].Width = -1;

			this.lstRequiredTables.Height = 
				(int)this.CreateGraphics().MeasureString(this.lstRequiredTables.Items[0].Text,this.Font).Height * 11;
			this.lstRequiredTables.Height = (int)(this.lstRequiredTables.Height * 1.5);

			p_ado = null;
		}
		public void LoadValues(string p_strScenarioId)
		{
			string strPathAndFile="";
			string strSQL="";
			string strConn="";

			this.lstRequiredTables.Clear();

			ado_data_access p_ado = new ado_data_access();
 
			this.lstRequiredTables.Columns.Add("Table Type", 50, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Path", 150, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("MDB File", 60, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("File Status", 80, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Table Name", 50, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Table Status", 80, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Record Count", 80, HorizontalAlignment.Left);

			System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
			string strDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\core\\db";
			string strFile = "scenario_core_rule_definitions.mdb"; 
			StringBuilder strFullPath = new StringBuilder(strDir);
			strFullPath.Append("\\");
			strFullPath.Append(strFile);
			strConn= p_ado.getMDBConnString(strFullPath.ToString().Trim(),"admin","");
			oConn.ConnectionString = strConn;
			intError = 0;
			strError = "";
			try
			{
				oConn.Open();
			}
			catch (System.Data.OleDb.OleDbException oleException)
			{
				strError = "Failed to make an oleDb connection with " + strConn;
				MessageBox.Show (strError);
				intError = -1;
				return;
			}
			System.Data.OleDb.OleDbCommand oCommand = oConn.CreateCommand();

			oCommand.CommandText = "select table_type,path,file,table_name from scenario_datasource" + 
				" where scenario_id = '" + p_strScenarioId + "';";
			try
			{
				System.Data.OleDb.OleDbDataReader oDataReader = oCommand.ExecuteReader();
				int x = 0;
				
				while (oDataReader.Read())
				{
					// Add a ListItem object to the ListView.
					System.Windows.Forms.ListViewItem entryListItem =
						lstRequiredTables.Items.Add(oDataReader["table_type"].ToString());
					entryListItem.UseItemStyleForSubItems=false;

					this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["path"].ToString());
					this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["file"].ToString());
					strPathAndFile = oDataReader["path"].ToString().Trim() + "\\" + oDataReader["file"].ToString().Trim();
					if (System.IO.File.Exists(strPathAndFile) == true) 
					{
						ListViewItem.ListViewSubItem FileStatusSubItem = 
							entryListItem.SubItems.Add("Found");
						FileStatusSubItem.ForeColor = System.Drawing.Color.Black;
						FileStatusSubItem.BackColor = System.Drawing.Color.White;
						FileStatusSubItem.Font = new System.Drawing.Font(
							"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_name"].ToString());

						//see if the table exists in the mdb database container
						dao_data_access p_dao = new dao_data_access();
						if (p_dao.TableExists(strPathAndFile,oDataReader["table_name"].ToString().Trim()) == true)
						{
							this.lstRequiredTables.Items[x].SubItems.Add("Found");
							strConn = p_ado.getMDBConnString(strPathAndFile,"admin","");
							strSQL = "select count(*) from " + oDataReader["table_name"].ToString();
							this.lstRequiredTables.Items[x].SubItems.Add(Convert.ToString(p_ado.getRecordCount(strConn,strSQL,oDataReader["table_name"].ToString())));

						}
						else 
						{
							ListViewItem.ListViewSubItem TableStatusSubItem = 
								entryListItem.SubItems.Add("Not Found");
							TableStatusSubItem.ForeColor = System.Drawing.Color.White;
							TableStatusSubItem.BackColor = System.Drawing.Color.Red;
							TableStatusSubItem.Font = new System.Drawing.Font(
								"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
							this.lstRequiredTables.Items[x].SubItems.Add("0");
						}
						p_dao = null;
					}
					else 
					{
						ListViewItem.ListViewSubItem FileStatusSubItem = 
							entryListItem.SubItems.Add("Not Found");
						FileStatusSubItem.ForeColor = System.Drawing.Color.White;
						FileStatusSubItem.BackColor = System.Drawing.Color.Red;
						FileStatusSubItem.Font = new System.Drawing.Font(
							"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_name"].ToString());
						ListViewItem.ListViewSubItem TableStatusSubItem = 
							entryListItem.SubItems.Add("Not Found");
						TableStatusSubItem.ForeColor = System.Drawing.Color.White;
						TableStatusSubItem.BackColor = System.Drawing.Color.Red;
						TableStatusSubItem.Font = new System.Drawing.Font(
							"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
						this.lstRequiredTables.Items[x].SubItems.Add("0");
					}
					x++;
				}
				oDataReader.Close();
			}
			catch
			{
				intError = -1;
				strError = "The Query Command " + oCommand.CommandText.ToString() + " Failed";
				MessageBox.Show(strError);
				oConn.Close();
				p_ado= null;
				return;
			}
			
			oConn.Close();
			

			this.lstRequiredTables.Columns[TABLETYPE].Width = -1;
			this.lstRequiredTables.Columns[PATH].Width = -1;
			this.lstRequiredTables.Columns[MDBFILE].Width = -1;
			this.lstRequiredTables.Columns[TABLE].Width = -1;
			this.lstRequiredTables.Height = 
				(int)this.CreateGraphics().MeasureString(this.lstRequiredTables.Items[0].Text,this.Font).Height * 11;
			this.lstRequiredTables.Height = (int)(this.lstRequiredTables.Height * 1.5);
			p_ado = null;

		}

		private void uc_scenario_datasource_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			((frmCoreScenario)this.ParentForm).m_bPopup = false;
		}

		private void lstRequiredTables_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			((frmCoreScenario)this.ParentForm).m_bPopup = false;
		}

		private void uc_scenario_datasource_Resize(object sender, System.EventArgs e)
		{

		}
		public void resize_uc_scenario_datasource()
		{
			this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
			this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
			this.lstRequiredTables.Left = 5;
			this.lstRequiredTables.Width = this.Width - 10;
			
			this.btnEdit.Top = this.lstRequiredTables.Top + this.lstRequiredTables.Height + 5;
			this.btnEdit.Left = (int) (this.groupBox1.Width * .50) - (int) (this.btnEdit.Width / 2);
			this.btnClearAll.Left = this.lstRequiredTables.Left + 5;
			this.btnCheckAll.Left = this.btnClearAll.Left;
			this.btnCopy.Left = this.btnClearAll.Left;
			this.btnClearAll.Top = this.btnEdit.Top + 5;
			this.btnCheckAll.Top = this.btnClearAll.Top + this.btnClearAll.Height;
			this.btnCopy.Top = this.btnCheckAll.Top + this.btnCheckAll.Height;
		}

		private void btnUndo_Click(object sender, System.EventArgs e)
		{
			DialogResult result = MessageBox.Show("All your changes will be undone and replaced with data source values in the database?(y/n)", "Data Sources", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result)
			{
				case DialogResult.Yes :
					this.populate_listview_grid();
				    break;
			}
		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			if (this.lstRequiredTables.SelectedItems.Count == 0)
			{
				return;
			}
            string strConn="";
			string strSQL = "";
			int y;
			ado_data_access p_ado = new ado_data_access();

            string strMDBFullPath = this.lstRequiredTables.SelectedItems[0].SubItems[PATH].Text + "\\" + 
				             this.lstRequiredTables.SelectedItems[0].SubItems[MDBFILE].Text;
			string strTable = this.lstRequiredTables.SelectedItems[0].SubItems[TABLE].Text;
           

			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog((frmMain)this.ParentForm.ParentForm);
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Text = "Core Analysis: Edit " + this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text + " Data Source";

			FIA_Biosum_Manager.uc_datasource_edit p_uc = new uc_datasource_edit(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\core\\db\\scenario_core_rule_definitions.mdb",
				                                                                ((frmCoreScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.Trim().ToLower());
			frmTemp.Controls.Add(p_uc);
           
			

			frmTemp.Height=0;
			frmTemp.Width=0;
			if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
			{
				for (y=1;;y++)
				{
					frmTemp.Height = y;
					if (p_uc.Top + 
						p_uc.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (p_uc.Left + p_uc.Width > frmTemp.ClientSize.Width + 2)
			{
				for (y=1;;y++)
				{
					frmTemp.Width = y;
					if (p_uc.Left + 
						p_uc.Width < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.Left = 0;
			frmTemp.Top = 0;

			p_uc.lblMDBFile.Text = strMDBFullPath.Trim();
			p_uc.lblTable.Text = strTable.Trim();
			p_uc.lblTableType.Text = 
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim();
			
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
									
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				utils p_utils = new utils();
				string strDir = p_utils.getDirectory(p_uc.lblNewMDBFile.Text);
				string strFile = p_utils.getFileName(p_uc.lblNewMDBFile.Text);
				this.lstRequiredTables.SelectedItems[0].SubItems[PATH].Text = strDir;
				this.lstRequiredTables.SelectedItems[0].SubItems[MDBFILE].Text = strFile;
				ListViewItem.ListViewSubItem FileSubItem = 
					this.lstRequiredTables.SelectedItems[0].SubItems[FILESTATUS];

				// Change the expenseItem object's color and font.
				FileSubItem.ForeColor = System.Drawing.Color.Black;
				FileSubItem.BackColor = System.Drawing.Color.White;
				FileSubItem.Font = new System.Drawing.Font(
					"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
				this.lstRequiredTables.SelectedItems[0].SubItems[FILESTATUS].Text = "Found";
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLE].Text = p_uc.lblNewTable.Text;
				ListViewItem.ListViewSubItem TableSubItem = 
					this.lstRequiredTables.SelectedItems[0].SubItems[TABLESTATUS];

				// Change the expenseItem object's color and font.
				TableSubItem.ForeColor = System.Drawing.Color.Black;
				TableSubItem.BackColor = System.Drawing.Color.White;
				FileSubItem.Font = new System.Drawing.Font(
					"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLESTATUS].Text = "Found";
				p_utils = null;

				strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + p_uc.lblNewMDBFile.Text.Trim() + ";User Id=admin;Password=;";
				strSQL = "select count(*) from " + p_uc.lblNewTable.Text.Trim();
				this.lstRequiredTables.SelectedItems[0].SubItems[RECORDCOUNT].Text =
					Convert.ToString(p_ado.getRecordCount(strConn,strSQL,p_uc.lblNewTable.Text.Trim()));

			}
			frmTemp.Close();
			frmTemp.Dispose();
			frmTemp = null;
			p_ado = null;
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Visible=false;
			((frmCoreScenario)this.ParentForm).Height = 0 ;
		}

		private void btnCopy_Click(object sender, System.EventArgs e)
		{
		   /***********************************************************
		    **since we are copying all the mdb files to the same
			**scenario directory, ensure that there are not 2 of the same
			**mdb file names that reside in 2 different directories
			************************************************************/
            int x;
			int y;
			string strMDBFile;
			string strDir;
			string strCopyFrom;
			string strCopyTo;
			string strCopyToDir;
			string strConn;
			string strTable="";
			System.Data.DataTable p_dt;

			if (this.lstRequiredTables.CheckedItems.Count==0) return;

			
			
			Datasource p_datasource = new Datasource(((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim(),((frmCoreScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.ToString());
            string strTempMDBFile = p_datasource.CreateMDBAndTableDataSourceLinks();
			try
			{
				
				string strScenarioConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
					((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + 
					"\\core\\db\\scenario_core_rule_definitions.mdb;User Id=admin;Password=;";
			   /*******************************
			    ** copy the files
			    *******************************/
				FIA_Biosum_Manager.frmTherm p_frmTherm;
				p_frmTherm = new FIA_Biosum_Manager.frmTherm();
				p_frmTherm.btnCancel.Visible=false;
				p_frmTherm.Show();
				p_frmTherm.Focus();
			
				
				p_frmTherm.Text = "Copy Files To Scenario Directory";
				p_frmTherm.Refresh();
				p_frmTherm.progressBar1.Minimum = 1;
				p_frmTherm.AbortProcess = false;
				p_frmTherm.progressBar1.Maximum = this.lstRequiredTables.Items.Count;
				p_frmTherm.lblMsg.Text = "";
				p_frmTherm.lblMsg.Visible=true;
				p_frmTherm.lblMsg.Refresh();

			    strCopyToDir = ((frmCoreScenario)this.ParentForm).uc_scenario1.txtScenarioPath.Text.Trim() + "\\db";
			    for (x=0; x <= this.lstRequiredTables.Items.Count-1; x++)
			    {
					if (this.lstRequiredTables.Items[x].Checked==true)
					{

						
						//get structure of the table
						ado_data_access p_ado = new ado_data_access();
						strConn = p_ado.getMDBConnString(strTempMDBFile,"","");
						p_dt = p_ado.getTableSchema(strConn,"select * from " + this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
						

						//get the table structure

						strMDBFile = this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim();
						strDir = this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim();
						strCopyFrom = strDir + "\\" + strMDBFile;
						strCopyTo = strCopyToDir + "\\" + strMDBFile;
						p_frmTherm.Increment(x+1);
						p_frmTherm.lblMsg.Text = strCopyTo;
						p_frmTherm.lblMsg.Refresh();

						dao_data_access p_dao = new dao_data_access();
						strTable=this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim();
						//check to see if the file is already in the scenario directory
						if (strCopyFrom.ToUpper().Trim() != strCopyTo.ToUpper().Trim())
						{
							
							//create the MDB  File
							if (System.IO.File.Exists(strCopyTo) == false)
							{
								p_dao.CreateMDB(strCopyTo);
								p_dao.CreateMDBTableFromDataSetTable(strCopyTo,strTable,p_dt,true);
							}
							else
							{
								//check to see if the table is in the mdb file
								if (p_dao.TableExists(strCopyTo,strTable) == true)
								{
									FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
									frmTemp.MaximizeBox = false;
									frmTemp.BackColor = System.Drawing.SystemColors.Control;
									frmTemp.Text = "Copy Table:" + strTable + " Exists";

									FIA_Biosum_Manager.uc_table_exists_dialog p_uc = new uc_table_exists_dialog();
									frmTemp.Controls.Add(p_uc);

									frmTemp.Height=0;
									frmTemp.Width=0;
									if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
									{
										for (y=1;;y++)
										{
											frmTemp.Height = y;
											if (p_uc.Top + 
												p_uc.Height < 
												frmTemp.ClientSize.Height)
											{
												break;
											}
										}

									}
									if (p_uc.Left + p_uc.Width > frmTemp.ClientSize.Width + 2)
									{
										for (y=1;;y++)
										{
											frmTemp.Width = y;
											if (p_uc.Left + 
												p_uc.Width < 
												frmTemp.ClientSize.Width)
											{
												break;
											}
										}

									}
									frmTemp.Left = 0;
									frmTemp.Top = 0;
									p_uc.strMDBFileLabel = "Table Contents Of " + strCopyTo.Trim();
									p_uc.listBox1.Items.Clear();
									p_dao.LoadTablesIntoListBox(strCopyTo,p_uc.listBox1);
									p_uc.strTable=strTable;
									frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
									frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
									for (;;)
									{
									   
										System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
										if (result==System.Windows.Forms.DialogResult.OK)
										{
											if (p_uc.rdoOverwrite.Checked==true)
											{
												p_dao.DeleteTableFromMDB(strCopyTo,this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
												p_dao.CreateMDBTableFromDataSetTable(strCopyTo,strTable,p_dt,true);
												break;
											}
											else if (p_dao.TableExists(strCopyTo,p_uc.strTable) == false)
											{
												strTable = p_uc.strTable;
												p_dao.CreateMDBTableFromDataSetTable(strCopyTo,strTable,p_dt,true);
												break;
											}
											else
											{
												frmTemp.Text = "Copy Table:" + p_uc.strTable + " Exists";
											}
										}
										else
										{
											strTable="";
											break;
										}
									}
									frmTemp.Dispose();
								}
							}
							//create primary indexes and autonumbers
							p_datasource.SetPrimaryIndexesAndAutoNumbers(strCopyTo,this.lstRequiredTables.Items[x].Text.Trim(),strTable,p_dao);
							if (strTable.Trim().Length > 0)
							{

								//append the contents of the source table to the target table
								p_dao.CreateTableLink(strTempMDBFile,"temp_link",strCopyTo,strTable,true);
								p_dao=null;
								p_ado.m_strSQL = "INSERT INTO temp_link SELECT * FROM " + this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim() + ";";
								p_ado.SqlNonQuery(strConn,p_ado.m_strSQL);
								if (p_ado.m_intError==0)
								{
									p_ado.m_strSQL = "UPDATE scenario_datasource SET path = '" + strCopyToDir + "',table_name = '" + strTable.Trim() + "' " + 
										" WHERE trim(scenario_id) = '" + ((frmCoreScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.Trim() + "' AND " +
										"trim(table_type) = '" + this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim() + "';";
									p_ado.SqlNonQuery(strScenarioConn, p_ado.m_strSQL);
									if (p_ado.m_intError == 0)
									{
										this.lstRequiredTables.Items[x].SubItems[PATH].Text = strCopyToDir;
										this.lstRequiredTables.Items[x].SubItems[TABLE].Text = strTable.Trim();
									}
								}
							}
							
						}

						p_dao=null;
						p_ado = null;
					}
			    }
				this.lstRequiredTables.Columns[PATH].Width = -1;
				p_frmTherm.Close();
				p_frmTherm = null;

			}
			catch (Exception e2)
			{
				MessageBox.Show(e2.Message);

			}
			p_datasource = null;

		}
		public bool ScenarioDataSourceTableExist(string strTableName)
		{
			int x;
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim().ToUpper()==strTableName.Trim().ToUpper())
				{
					if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="NOT FOUND")
					{
						return false;
					}
					if (this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="NOT FOUND")
					{
						return false;
					}
					return true;
				}   
			}
      
			return false;



		}
		/*****************************************************
		 ** create a mdb table in the users temporary dir
		 ** and create a link to each of the scenario 
		 ** data source tables.  Return the name of the 
		 ** temporary mdb file to the calling function
		 *****************************************************/
		public string CreateMDBAndScenarioTableDataSourceLinks(string strDestinationLinkDir)
		{
			string strTempMDB="";
			int x;
            this.m_intNumberOfValidTables=0;
			//used to get the temporary random file name
			utils p_utils = new utils();
			//used to create a link to the table
			dao_data_access p_dao = new dao_data_access();
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
					if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
						this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
					{
						if (strTempMDB.Trim().Length == 0)
						{
							//get temporary mdb file
							strTempMDB = 
								p_utils.getRandomFile(strDestinationLinkDir,"mdb");

							//create a temporary mdb that will contain all 
							//the links to the scenario datasource tables
							p_dao.CreateMDB(strTempMDB);

						}
						p_dao.CreateTableLink(strTempMDB,
							this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim(),
							this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim() + "\\" +
							     this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim(),
							this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
						this.m_intNumberOfValidTables++;
						

					}
			}
			p_utils = null;
			p_dao = null;
            if (strTempMDB.Trim().Length == 0)
				MessageBox.Show("!!None of the scenario data source tables are found!!");
			return strTempMDB;
		}

		public void getScenarioDataSourceMDBAndTableName(string strTableID,ref string strMDBPathAndFile,ref string strTable)
		{
			int x;
			for (x=0; x<= this.lstRequiredTables.Items.Count-1;x++)
			{
				if (strTableID.Trim().ToUpper() == 
					this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper()
					&&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper() =="FOUND" 
					&&
					this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper() == "FOUND")
				{
					  strMDBPathAndFile = this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim() + "\\" + this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim();
					  strTable = this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim();
					  break;
				}
			}
		}

		public void getNumberOfValidTables()
		{
			int x;
			this.m_intNumberOfValidTables=0;
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
				{
					this.m_intNumberOfValidTables++;
				}
			}

		}

		public void LoadScenarioDataSourceTablesIntoListBox(System.Windows.Forms.ListBox listbox1)
		{
			int x;
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
				{
					listbox1.Items.Add(this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
				}
			}
		}
		

		
		 
		/********************************************************
		 ** return the row associated with the table type
		 ********************************************************/
		public int getDataSourceTableNameRow(string pcTableId)
		{
			int x;
			for (x=0; x<= this.lstRequiredTables.Items.Count-1;x++)
			{
				if (pcTableId.Trim().ToUpper() == 
					this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper())
				{
					return x;
				}
			}
			return -1;
		}
		/********************************************************
		 ** return the table name associated with the table type
		 ********************************************************/
		public string getDataSourceTableName(string pcTableId)
		{
			int x;
			for (x=0; x<= this.lstRequiredTables.Items.Count-1;x++)
			{
				if (pcTableId.Trim().ToUpper() == 
					this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper()
					&&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper() =="FOUND" 
					&&
					this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper() == "FOUND")
				{
					return this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim();
				}
			}
			return "";
		}
		public int val_datasources()
		{
			

            int x=0;
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="NOT FOUND")
				{
					MessageBox.Show("Run Scenario Failed: Scenario data source file " + this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim() + "\\" + 
						            this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim() + " is not found");
					return -1;
				}
				if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="NOT FOUND")
				{
					MessageBox.Show("Run Scenario Failed: Scenario data source table " + this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim() + 
						 " is not found");
					return -1;
				}
				if (this.lstRequiredTables.Items[x].SubItems[RECORDCOUNT].Text.Trim().ToUpper()=="0")
				{
					MessageBox.Show("Run Scenario Failed: Scenario data source table " + this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim() + 
						" has 0 records");
					return -1;
				}
			}
      
			return 0;
		}

		private void btnClearAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstRequiredTables.Items.Count-1;x++)
			{
				this.lstRequiredTables.Items[x].Checked=false;
			}
																	
		}

		private void btnCheckAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstRequiredTables.Items.Count-1;x++)
			{
				this.lstRequiredTables.Items[x].Checked=true;
			}

		}
		
		
      
	}
}
