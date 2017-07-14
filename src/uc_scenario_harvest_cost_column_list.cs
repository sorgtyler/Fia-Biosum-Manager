using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_list.
	/// </summary>
	public class uc_scenario_harvest_cost_column_list : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnEdit;
		private System.Windows.Forms.Button btnNew;
		public int m_intDialogHt;
		public int m_intDialogWd;
		private int m_intError=0;
		private string m_strError="";
		private System.Windows.Forms.Button btnDefault;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnClear;
		private System.Windows.Forms.Button btnHelp;
		private FIA_Biosum_Manager.ado_data_access m_ado;
		private FIA_Biosum_Manager.Datasource m_datasource;
		private string m_strHarvestCostDbFile;
		private string m_strHarvestCostTableName;
		private string m_strHarvestCostConn;
		private System.Windows.Forms.Button btnDelete;
		private FIA_Biosum_Manager.frmCoreScenario _frmScenario;
		private string _strScenarioId="";
		private System.Windows.Forms.ListView lstCol;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		private System.Data.OleDb.OleDbConnection m_OleDbConnectionHarvestCostDbFile;
		public string m_strScenarioColumnNameList="";       //list from current scenario
		public string m_strAllScenarioColumnNameList="";    //list from all the scenarios
		public string m_strHarvestTableColumnNameList="";   //harvest table column name list
		private FIA_Biosum_Manager.utils m_oUtils = new utils();
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvAlternateRowColors=new ListViewAlternateBackgroundColors();
		const int COLUMN_NULL=0;
		const int COLUMN_SCENARIO=1;
		const int COLUMN_FIELD=2;
		const int COLUMN_DESC=3;
		private FIA_Biosum_Manager.frmDialog _frmDialog;
		public bool m_bFirstTime=true;


		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_scenario_harvest_cost_column_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			
			this.m_intDialogHt = this.groupBox1.Top + this.btnClose.Top + this.btnClose.Height + 20;
			this.m_intDialogWd = this.groupBox1.Left + this.btnClose.Left + this.btnClose.Width + 20;

			m_oLvAlternateRowColors.ReferenceAlternateBackgroundColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvAlternateRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvAlternateRowColors.ReferenceListView=this.lstCol;
			this.m_oLvAlternateRowColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvAlternateRowColors.CustomFullRowSelect=true;
			if (frmMain.g_oGridViewFont != null) lstCol.Font = frmMain.g_oGridViewFont;

			// TODO: Add any initialization after the InitializeComponent call

		}
		~uc_scenario_harvest_cost_column_list()
		{

			try
			{
				if (this.m_OleDbConnectionHarvestCostDbFile != null)
				{
					while (this.m_OleDbConnectionHarvestCostDbFile.State != ConnectionState.Closed)
					{
						this.m_OleDbConnectionHarvestCostDbFile.Close();
						System.Threading.Thread.Sleep(1000);
					}
				}
				if (this.m_OleDbConnectionScenario != null)
				{
					while (this.m_OleDbConnectionScenario.State != ConnectionState.Closed)
					{
						this.m_OleDbConnectionScenario.Close();
						System.Threading.Thread.Sleep(1000);
					}
				}
			}
			catch
			{
			}
			this.m_oUtils=null;
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
		
		public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		public string ScenarioId
		{
			get {return _strScenarioId;}
			set {_strScenarioId=value;}
		}
		public void loadvalues()
		{
			string strConn;
			      
			this.lstCol.Clear();
			
			this.m_oLvAlternateRowColors.InitializeRowCollection();
			this.lstCol.Columns.Add("",2,HorizontalAlignment.Left);
			this.lstCol.Columns.Add("Scenario",60,HorizontalAlignment.Left);
			this.lstCol.Columns.Add("Harvest Cost Column", 200, HorizontalAlignment.Left);
			this.lstCol.Columns.Add("Description", 300, HorizontalAlignment.Left);

			this.m_intError=0;
			this.m_strError="";

			//
			//OPEN CONNECTION TO DB FILE CONTAINING HARVEST COST TABLE
			//
			this.m_ado = new ado_data_access();

			ScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

			this.m_datasource = new Datasource();
			m_datasource.m_strDataSourceMDBFile=frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +  "\\core\\db\\scenario_core_rule_definitions.mdb";
			m_datasource.m_strDataSourceTableName = "scenario_datasource";
			m_datasource.m_strScenarioId=ScenarioId.Trim();
			m_datasource.LoadTableColumnNamesAndDataTypes=false;
			m_datasource.LoadTableRecordCount=false;
			m_datasource.populate_datasource_array();

			this.m_strHarvestCostDbFile = m_datasource.getFullPathAndFile("Harvest Costs");

			this.m_OleDbConnectionHarvestCostDbFile = new System.Data.OleDb.OleDbConnection();
			this.m_strHarvestCostConn = m_ado.getMDBConnString(this.m_strHarvestCostDbFile,"admin","");
			m_ado.OpenConnection(this.m_strHarvestCostConn,ref this.m_OleDbConnectionHarvestCostDbFile);
			if (m_ado.m_intError !=0)
			{
				m_ado=null;
				m_intError=m_ado.m_intError;
				m_strError=m_ado.m_strError;
				return;
			}
			

			this.m_strHarvestCostTableName=m_datasource.getValidDataSourceTableName("Harvest Costs");

			LoadHarvestCostTableColumns();
			//
			//OPEN CONNECTION TO DB FILE CONTAINING HARVEST COST SCENARIO TABLE
			//
			//scenario mdb connection
			string strScenarioMDB = 
			  frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
			  "\\core\\db\\scenario_core_rule_definitions.mdb";

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn = m_ado.getMDBConnString(strScenarioMDB,"admin","");
			m_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (m_ado.m_intError != 0)
			{
				m_intError=m_ado.m_intError;
				m_strError=m_ado.m_strError;
				m_ado = null;
				return;
			}

			LoadHarvestCostScenarioData();;
		}
		private void LoadHarvestCostTableColumns()
		{
			int x;

			string strFieldsList= m_ado.getFieldNames(this.m_OleDbConnectionHarvestCostDbFile,"SELECT * FROM " + this.m_strHarvestCostTableName);

			string[] strFieldsArray = this.m_oUtils.ConvertListToArray(strFieldsList,",");


			strFieldsList="";
			for (x=0;x<=strFieldsArray.Length - 1;x++)
			{
				switch (strFieldsArray[x].Trim().ToUpper())
				{
					case "BIOSUM_COND_ID":
						break;
					case "RXPACKAGE":
						break;
					case "RXCYCLE":
						break;
					case "RX":
						break;
					case "COMPLETE_CPA":
						break;
					case "HARVEST_CPA":
						break;
					case "HARVEST_CPA_WARNING_MSG":
						break;
					default:
						strFieldsList = strFieldsList + strFieldsArray[x].Trim() + ",";
						break;
				}
			}

			if (strFieldsList.Trim().Length > 0)
			{
				strFieldsList=strFieldsList.Substring(0,strFieldsList.Length - 1);
				strFieldsArray=null;
				strFieldsArray = this.m_oUtils.ConvertListToArray(strFieldsList,",");
			}

			this.m_strHarvestTableColumnNameList = strFieldsList;

		}
		private void LoadHarvestCostScenarioData()
		{
			string strColumn="";
			string strDesc="";
			string strSQL = "SELECT * FROM scenario_harvest_cost_columns WHERE " + 
				" TRIM(scenario_id) = '" + ScenarioId.Trim() + "';";
			m_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
			
			if (m_ado.m_intError==0)
			{
				try
				{
					//load up each row in the FIADB plot input table
					if (m_ado.m_OleDbDataReader.HasRows)
					{
						while (m_ado.m_OleDbDataReader.Read())
						{
							strColumn="";
							strDesc="";
							
							//make sure the row is not null values
							if (m_ado.m_OleDbDataReader[0] != System.DBNull.Value &&
								m_ado.m_OleDbDataReader[0].ToString().Trim().Length > 0)
							{
								strColumn =m_ado.m_OleDbDataReader["ColumnName"].ToString().Trim();
								strDesc = m_ado.m_OleDbDataReader["description"].ToString();
								this.lstCol.BeginUpdate();
								System.Windows.Forms.ListViewItem listItem = new ListViewItem();
								listItem.UseItemStyleForSubItems=false;
								listItem.Text = "";
								listItem.SubItems.Add(ScenarioId);
								listItem.SubItems.Add(strColumn);
								listItem.SubItems.Add(strDesc);
								this.lstCol.Items.Add(listItem);
								this.lstCol.EndUpdate();
								this.m_strScenarioColumnNameList = this.m_strScenarioColumnNameList + strColumn + ",";
								this.m_oLvAlternateRowColors.AddRow();
								this.m_oLvAlternateRowColors.AddColumns(lstCol.Items.Count-1,lstCol.Columns.Count);
								
							}

						}
						if (m_strScenarioColumnNameList.Trim().Length > 0)
							this.m_strScenarioColumnNameList=this.m_strScenarioColumnNameList.Substring(0,this.m_strScenarioColumnNameList.Length - 1);
					}
					m_ado.m_OleDbDataReader.Close();
					if (this.lstCol.Items.Count > 0)
					{
						this.lstCol.Columns[COLUMN_FIELD].Width = -1;
						this.lstCol.Columns[COLUMN_SCENARIO].Width=-1;
					}
					if (this.lstCol.Items.Count > 0)
					{                                                       
						if (this.lstCol.SelectedItems.Count == 0)
						{
							this.lstCol.Items[this.lstCol.Items.Count-1].Selected=true;
							this.btnEdit.Enabled=true;
							this.btnDelete.Enabled=true;
						}
					}
					//get all the columns
					strSQL = "SELECT DISTINCT ColumnName FROM scenario_harvest_cost_columns";
					m_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
					//load up each row in the FIADB plot input table
					if (m_ado.m_OleDbDataReader.HasRows)
					{
						while (m_ado.m_OleDbDataReader.Read())
						{
							strColumn="";
							
							//make sure the row is not null values
							if (m_ado.m_OleDbDataReader[0] != System.DBNull.Value &&
								m_ado.m_OleDbDataReader[0].ToString().Trim().Length > 0)
							{
								strColumn =m_ado.m_OleDbDataReader["ColumnName"].ToString().Trim();
								m_strAllScenarioColumnNameList = this.m_strAllScenarioColumnNameList + strColumn + ",";
								
							}

						}
						if (m_strAllScenarioColumnNameList.Trim().Length > 0)
							this.m_strAllScenarioColumnNameList=this.m_strAllScenarioColumnNameList.Substring(0,this.m_strAllScenarioColumnNameList.Length - 1);

					}
					m_ado.m_OleDbDataReader.Close();

					((frmDialog)this.ParentForm).Enabled=true;
				}
				catch (Exception caught)
				{
					this.m_intError=-1;
					this.m_strError=caught.Message;
					MessageBox.Show(caught.Message);
				}
				

			}
			else
			{
				this.m_intError=m_ado.m_intError;
				this.m_strError=m_ado.m_strError;
			}
		}
		private void LoadCurrentList()
		{
			int x;
			this.m_strScenarioColumnNameList="";
			for (x=0;x<=this.lstCol.Items.Count-1;x++)
			{
				m_strScenarioColumnNameList = this.lstCol.Items[x].SubItems[COLUMN_FIELD].Text;
			}
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnDelete = new System.Windows.Forms.Button();
			this.btnHelp = new System.Windows.Forms.Button();
			this.btnClear = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnSave = new System.Windows.Forms.Button();
			this.btnDefault = new System.Windows.Forms.Button();
			this.btnNew = new System.Windows.Forms.Button();
			this.btnEdit = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.lstCol = new System.Windows.Forms.ListView();
			this.lblTitle = new System.Windows.Forms.Label();
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
			this.groupBox1.Controls.Add(this.btnDefault);
			this.groupBox1.Controls.Add(this.btnNew);
			this.groupBox1.Controls.Add(this.btnEdit);
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.lstCol);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(624, 480);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// btnDelete
			// 
			this.btnDelete.Enabled = false;
			this.btnDelete.Location = new System.Drawing.Point(310, 392);
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(64, 32);
			this.btnDelete.TabIndex = 10;
			this.btnDelete.Text = "Delete";
			this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
			// 
			// btnHelp
			// 
			this.btnHelp.Location = new System.Drawing.Point(16, 432);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(96, 32);
			this.btnHelp.TabIndex = 8;
			this.btnHelp.Text = "Help";
			// 
			// btnClear
			// 
			this.btnClear.Location = new System.Drawing.Point(374, 392);
			this.btnClear.Name = "btnClear";
			this.btnClear.Size = new System.Drawing.Size(64, 32);
			this.btnClear.TabIndex = 7;
			this.btnClear.Text = "Clear All";
			this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(502, 392);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 32);
			this.btnCancel.TabIndex = 6;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnSave
			// 
			this.btnSave.Enabled = false;
			this.btnSave.Location = new System.Drawing.Point(438, 392);
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(64, 32);
			this.btnSave.TabIndex = 5;
			this.btnSave.Text = "Save";
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// btnDefault
			// 
			this.btnDefault.Location = new System.Drawing.Point(68, 392);
			this.btnDefault.Name = "btnDefault";
			this.btnDefault.Size = new System.Drawing.Size(114, 32);
			this.btnDefault.TabIndex = 2;
			this.btnDefault.Text = "Use Default Values";
			//this.btnDefault.Click += new System.EventHandler(this.btnDefault_Click);
			// 
			// btnNew
			// 
			this.btnNew.Location = new System.Drawing.Point(182, 392);
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(64, 32);
			this.btnNew.TabIndex = 3;
			this.btnNew.Text = "New";
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnEdit
			// 
			this.btnEdit.Enabled = false;
			this.btnEdit.Location = new System.Drawing.Point(246, 392);
			this.btnEdit.Name = "btnEdit";
			this.btnEdit.Size = new System.Drawing.Size(64, 32);
			this.btnEdit.TabIndex = 4;
			this.btnEdit.Text = "Edit";
			this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(513, 432);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 9;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// lstCol
			// 
			this.lstCol.GridLines = true;
			this.lstCol.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lstCol.HideSelection = false;
			this.lstCol.Location = new System.Drawing.Point(16, 48);
			this.lstCol.MultiSelect = false;
			this.lstCol.Name = "lstCol";
			this.lstCol.Size = new System.Drawing.Size(592, 336);
			this.lstCol.TabIndex = 1;
			this.lstCol.View = System.Windows.Forms.View.Details;
			this.lstCol.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstCol_MouseUp);
			this.lstCol.SelectedIndexChanged += new System.EventHandler(this.lstCol_SelectedIndexChanged);
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(618, 24);
			this.lblTitle.TabIndex = 0;
			this.lblTitle.Text = "Harvest Cost Columns";
			// 
			// uc_scenario_harvest_cost_column_list
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_harvest_cost_column_list";
			this.Size = new System.Drawing.Size(624, 480);
			this.Resize += new System.EventHandler(this.uc_scenario_harvest_cost_column_list_Resize);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			if (this.btnSave.Enabled==true)
			{
				DialogResult result = MessageBox.Show("Save Changes Y/N","Plot Treatments",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
				if (result == System.Windows.Forms.DialogResult.Yes)
				{
					this.savevalues();
				}
			}
			this.ParentForm.Close();
		}

		private void btnNew_Click(object sender, System.EventArgs e)
		{

			this.Edit("New");

		}
		private void Edit(string strType)
		{
			int y;
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Initialize_Scenario_Harvest_Costs_Column_Edit_Control();


			frmTemp.Height=0;
			frmTemp.Width=0;
			if (frmTemp.Top + frmTemp.uc_scenario_harvest_cost_column_edit1.Height > frmTemp.ClientSize.Height + 2)
			{
				for (y=1;;y++)
				{
					frmTemp.Height = y;
					if (frmTemp.uc_scenario_harvest_cost_column_edit1.Top + 
						frmTemp.uc_scenario_harvest_cost_column_edit1.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (frmTemp.uc_scenario_harvest_cost_column_edit1.Left + frmTemp.uc_scenario_harvest_cost_column_edit1.Width > frmTemp.ClientSize.Width + 2)
			{
				for (y=1;;y++)
				{
					frmTemp.Width = y;
					if (frmTemp.uc_scenario_harvest_cost_column_edit1.Left + 
						frmTemp.uc_scenario_harvest_cost_column_edit1.Width < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.Left = 0;
			frmTemp.Top = 0;
      
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			frmTemp.uc_scenario_harvest_cost_column_edit1.Dock = System.Windows.Forms.DockStyle.Fill;		
			
			frmTemp.uc_scenario_harvest_cost_column_edit1.EditType=strType;
			
			if (strType.Trim().ToUpper() == "NEW")
			{
				frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnList = this.m_strAllScenarioColumnNameList;
				this.m_strScenarioColumnNameList="";
				for (y=0;y<=this.lstCol.Items.Count-1;y++)
				{
					this.m_strScenarioColumnNameList=this.m_strScenarioColumnNameList + this.lstCol.Items[y].SubItems[COLUMN_FIELD].Text.Trim() + ",";
				}
				if (this.m_strScenarioColumnNameList.Trim().Length > 0)
					this.m_strScenarioColumnNameList=this.m_strScenarioColumnNameList.Substring(0,this.m_strScenarioColumnNameList.Length - 1);
				frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText = "";
				frmTemp.uc_scenario_harvest_cost_column_edit1.CurrentSelectedColumnList=this.m_strScenarioColumnNameList;
				frmTemp.uc_scenario_harvest_cost_column_edit1.HarvestCostTableColumnList = this.m_strHarvestTableColumnNameList;
				frmTemp.uc_scenario_harvest_cost_column_edit1.loadvalues();
				frmTemp.uc_scenario_harvest_cost_column_edit1.lblEdit.Show();
			}
			else
			{
				frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText=this.lstCol.SelectedItems[0].SubItems[COLUMN_FIELD].Text;
				frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription = this.lstCol.SelectedItems[0].SubItems[COLUMN_DESC].Text;
				frmTemp.uc_scenario_harvest_cost_column_edit1.cmbCol.Enabled=false;
				frmTemp.uc_scenario_harvest_cost_column_edit1.lblEdit.Hide();
			}
			
			frmTemp.Text = strType + " Harvest Cost Column";
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				if (strType.Trim().ToUpper()=="NEW")
				{
					this.lstCol.BeginUpdate();
					this.lstCol.Items.Add("");
				    this.lstCol.Items[this.lstCol.Items.Count-1].SubItems.Add(this.ScenarioId);
					this.lstCol.Items[this.lstCol.Items.Count-1].UseItemStyleForSubItems=false;
					this.lstCol.Items[this.lstCol.Items.Count-1].SubItems.Add(frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText);
					this.lstCol.Items[this.lstCol.Items.Count-1].SubItems.Add(frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription);
					this.lstCol.EndUpdate();
					this.m_oLvAlternateRowColors.AddRow();
					this.m_oLvAlternateRowColors.AddColumns(lstCol.Items.Count-1,lstCol.Columns.Count);
					this.lstCol.Items[this.lstCol.Items.Count-1].Selected= true;
					this.LoadCurrentList();
					
				}
				else
				{
					this.lstCol.SelectedItems[0].SubItems[COLUMN_DESC].Text = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription;
				}
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
			frmTemp.Dispose();
		}

        //private void btnDefault_Click(object sender, System.EventArgs e)
        //{
        //    this.lstCol.Clear();
			
        //    this.lstCol.Columns.Add("",2,HorizontalAlignment.Left);
        //    this.lstCol.Columns.Add("Scenario",60,HorizontalAlignment.Left);
        //    this.lstCol.Columns.Add("Harvest Cost Component", 200, HorizontalAlignment.Left);
        //    this.lstCol.Columns.Add("Description", 300, HorizontalAlignment.Left);
			
        //    this.m_oLvAlternateRowColors.InitializeRowCollection();
        //    this.m_intError=0;
        //    this.lstCol.BeginUpdate();

			//this.lstCol.Items.Add("");
			//this.lstCol.Items[0].SubItems.Add(ScenarioId);
			//lstCol.Items[0].UseItemStyleForSubItems=false;
			//this.lstCol.Items[0].SubItems.Add("water_barring_roads_cpa");
			//this.lstCol.Items[0].SubItems.Add(" ");
			//this.m_oLvAlternateRowColors.AddRow();
			//this.m_oLvAlternateRowColors.AddColumns(0,lstCol.Columns.Count);
            
			//this.lstCol.Items.Add("");
			//this.lstCol.Items[1].SubItems.Add(ScenarioId);
			//lstCol.Items[1].UseItemStyleForSubItems=false;
			//this.lstCol.Items[1].SubItems.Add("brush_cutting_cpa");
			//this.lstCol.Items[1].SubItems.Add(" ");
			//this.m_oLvAlternateRowColors.AddRow();
			//this.m_oLvAlternateRowColors.AddColumns(1,lstCol.Columns.Count);


        //    this.lstCol.EndUpdate();
        //    if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;

        //    this.lstCol.Columns[COLUMN_FIELD].Width = -1;
        //}

		private void btnClear_Click(object sender, System.EventArgs e)
		{
			if (this.lstCol.Items.Count > 0) this.btnSave.Enabled=true;
			this.btnDelete.Enabled=false;
			this.btnEdit.Enabled=false;
			this.lstCol.Items.Clear();
			this.m_oLvAlternateRowColors.InitializeRowCollection();
//			this.lstCol.Columns.Add("Harvest Cost Column", 60, HorizontalAlignment.Left);
//			this.lstCol.Columns.Add("Description", 150, HorizontalAlignment.Left);
		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			
			
			if (this.lstCol.SelectedItems.Count == 0)
				return;

			this.Edit("Modify");

		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
		  this.savevalues();
		}
		private void val_data()
		{
			//int x;
			//int y;
			

			//if (this.lstCol.Items.Count==0)
			//{
			//	MessageBox.Show("No plot treatments to save","Plot Treatments",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			//	this.m_intError=-1;
			//	return;
			//}

			//should have a connection initialized in loadvalues module so check if okay
			//if (m_strConn.Trim().Length == 0)
			//{
			//	MessageBox.Show("Error: No connection string to plot treatment table","Plot Treatments",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			//	this.m_intError=-1;
			//	return;
			//}
            

		}
		public void savevalues()
		{

			int x;
			int y;

            m_intError=0;
			m_strError="";
			string strCol;
			string strDesc;

			//
			//delete harvest cost scenario
			//
			//delete the current records from the core analysis scenario
			this.m_ado.m_strSQL = "DELETE FROM scenario_harvest_cost_columns WHERE " + 
				" TRIM(scenario_id) = '" + ScenarioId.Trim() + "';";
			this.m_ado.SqlNonQuery(m_OleDbConnectionScenario,this.m_ado.m_strSQL);

			
			if (lstCol.Items.Count > 0)
			{
				
				//
				//create the column in the table
				//
				//load existing columns that are in the harvest cost table
				LoadHarvestCostTableColumns();
				string[] strArray = this.m_oUtils.ConvertListToArray(this.m_strHarvestTableColumnNameList,",");
				for (x=0;x<=lstCol.Items.Count-1;x++)
				{
					strCol = lstCol.Items[x].SubItems[COLUMN_FIELD].Text.Trim();
					strDesc = lstCol.Items[x].SubItems[COLUMN_DESC].Text.Trim();
					strDesc = m_ado.FixString(strDesc,"'","''");

					//make sure column does not already exist
					for (y=0;y<=strArray.Length-1;y++)
					{
						if (strCol.ToUpper() == strArray[y].Trim().ToUpper())
							break;
					}
					if (y > strArray.Length-1)
					{
					   m_ado.m_strSQL = "ALTER TABLE " + this.m_strHarvestCostTableName + " ADD COLUMN " + strCol + " SINGLE;";	
					   m_ado.SqlNonQuery(this.m_OleDbConnectionHarvestCostDbFile,m_ado.m_strSQL);
					}
					//add the column info to the core analysis scenario table
					m_ado.m_strSQL = "INSERT INTO scenario_harvest_cost_columns " + 
						"(scenario_id,columnname,description) VALUES " + 
						"('" + ScenarioId + "','" + strCol.Trim() + "','" + strDesc.Trim() + "');";
					this.m_ado.SqlNonQuery(this.m_OleDbConnectionScenario,m_ado.m_strSQL);
				}
			}
			this.btnSave.Enabled=false;




		}
		private void getHarvestCostColumnDatasource()
		{

		    m_ado = new ado_data_access();


		}

		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			if (this.lstCol.SelectedItems.Count==0) return;
			int intSelected = this.lstCol.SelectedItems[0].Index;
			this.m_oLvAlternateRowColors.m_oRowCollection.Remove(intSelected);
			this.m_oLvAlternateRowColors.m_intSelectedRow=-1;
			this.lstCol.SelectedItems[0].Remove();
			if (this.lstCol.Items.Count > 0)
			{
				if (intSelected == 0)
				{
					this.lstCol.Items[intSelected].Selected=true;
				}
				else if (intSelected == this.lstCol.Items.Count-1)
				{
					this.lstCol.Items[intSelected].Selected=true;
				}
				else
				{
					this.lstCol.Items[intSelected-1].Selected=true;
				}
				
			}
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			this.lstCol.Focus();
		}

		private void lstCol_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.btnDelete.Enabled==false) this.btnDelete.Enabled=true;
			if (this.btnEdit.Enabled==false) this.btnEdit.Enabled=true;

			if (this.lstCol.SelectedItems.Count > 0)
				this.m_oLvAlternateRowColors.DelegateListViewItem(lstCol.SelectedItems[0]);
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{

			this.ParentForm.Close();

			
		}

		private void uc_scenario_harvest_cost_column_list_Resize(object sender, System.EventArgs e)
		{
			try
			{

				if (this.m_bFirstTime==false)
				{
					this.lstCol.Width = this.ClientSize.Width - this.lstCol.Left * 2;

					this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
					this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;

				
				
					this.lstCol.Height = this.btnClose.Top - this.lstCol.Top - (this.btnEdit.Height * 2);
				
			
					this.btnEdit.Top = this.lstCol.Top + this.lstCol.Height + 5;
					this.btnEdit.Left = (int)(this.Width * .50) - this.btnEdit.Width;
					this.btnNew.Left = this.btnEdit.Left - this.btnNew.Width;
					this.btnDefault.Left = this.btnNew.Left - this.btnDefault.Width;

					this.btnDelete.Left = this.btnEdit.Left + this.btnDelete.Width;
					this.btnClear.Left = this.btnDelete.Left + this.btnClear.Width;
					this.btnSave.Left = this.btnClear.Left + this.btnSave.Width;
					this.btnCancel.Left = this.btnSave.Left + this.btnCancel.Width;

				
					this.btnCancel.Top = this.btnEdit.Top;
					this.btnSave.Top = this.btnEdit.Top;
					this.btnClear.Top = this.btnEdit.Top;
					this.btnDelete.Top = this.btnEdit.Top;
					this.btnNew.Top =this.btnEdit.Top;
					this.btnDefault.Top = this.btnEdit.Top;
				
					this.btnHelp.Top = this.btnClose.Top;
					this.btnHelp.Left = this.lstCol.Left;
				}
			}
			catch
			{
			}
		}

		private void lstCol_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = lstCol.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lstCol.Items[lstCol.TopItem.Index + (int)dblRow-1].Selected=true;
				}
			}
			catch 
			{
			}
		}

		private void uc_scenario_harvest_cost_column_list_Resize2(object sender, System.EventArgs e)
		{
			
			
		}
	
		public FIA_Biosum_Manager.frmDialog ReferenceDialog
		{
			get {return this._frmDialog;}
			set {this._frmDialog=value;}
		}
	}
}
