using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Threading;

/*
 private void button1_Click(object sender, System.EventArgs e)
{
   uc_PostFvsForeFrcs upfff=new uc_PostFvsForeFrcs();
   upfff.Bounds=new Rectangle(122, 80, 632, 416);
   this.Controls.Add(upfff);
   upfff.Show();
}
*/

namespace FIA_Biosum_Manager
{
	/// <summary> </summary>
	public class uc_PostFvsForeFrcs : System.Windows.Forms.UserControl
	{
		public run_PostFvsForeFrcs procPreFrcs=null;
		public Hashtable htSelectedRxFile=null;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnAppend;
		private System.Windows.Forms.Button btnRefresh;
		private System.Windows.Forms.Button btnClearAll;
		private System.Windows.Forms.Button btnChkAll;
		
		private System.Windows.Forms.TextBox txtOutDir;
		private System.Windows.Forms.Label label4;
		public int m_intError=0;
		public string m_strError="";
		public int m_intWarning=0;
		public string m_strWarning="";
		public System.Windows.Forms.ListView lstFvsOutput;

		private const int COL_CHECKBOX=0;
		private const int COL_VARIANT = 1;
		private const int COL_RX = 2;
		private const int COL_RUNSTATUS = 3;
		private const int COL_MDBOUT = 4;
		private const int COL_FOUND=5;
		private const int COL_SUMMARYCOUNT = 6;
		//private const int COL_SUMMARYPREYEAR=6;
		//private const int COL_SUMMARYPOSTYEAR=7;
		private const int COL_CUTCOUNT = 7;
		private const int COL_LEFTCOUNT = 8;
		private const int COL_POTFIRECOUNT = 9;


		private string m_strFFETable="";
		//private string m_strVariant;
		private string m_strOutMDBFile="";
		private string m_strRxTable="";
		private string m_strPlotTable="";
		private string m_strCondTable="";
		private string m_strFvsTreeTable="";
		private string m_strFVSSummaryAuditTable="audit_fvs_summary_year_table_counts";
		private string[] m_strFVSPrePostYearAuditTablesArray = null;
		private Datasource m_DataSource;
		private ado_data_access m_ado;
		private dao_data_access m_dao;
		private string m_strConn="";
		private string m_strTempMDBFile="";
		private string m_strProjDir="";

		string m_strLogFile;
		string m_strLogDate;

		private System.Threading.Thread m_thread;
		private FIA_Biosum_Manager.frmTherm m_frmTherm;
		private bool m_bDebug=true;
		private System.Windows.Forms.Label lblRunStatus;
		private utils m_oUtils = new utils();
		private System.Windows.Forms.Button btnAudit;
		private System.Windows.Forms.GroupBox grpboxAudit;
		private System.Windows.Forms.GroupBox grpboxAppend;
		private bool _bDisplayAuditMsg=true;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnViewLogFile;
		private System.Windows.Forms.Button btnAuditDb;
		private ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();
		int m_intCheckedCount=0;
		int m_intTotalSteps=75;
		int m_intCurrentStep=0;
		int m_intCurrentCount=0;


		


		



		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_PostFvsForeFrcs(string p_strProjDir)
		{

			InitializeComponent();
			this.m_strProjDir = p_strProjDir;

			this.txtOutDir.Text = this.m_strProjDir + "\\fvs\\db\\out";

			m_DataSource = new Datasource();
			m_DataSource.LoadTableColumnNamesAndDataTypes=false;
			m_DataSource.LoadTableRecordCount=false;
			m_DataSource.m_strDataSourceMDBFile = p_strProjDir.Trim() + "\\db\\project.mdb";
			m_DataSource.m_strDataSourceTableName = "datasource";
			m_DataSource.m_strScenarioId="";
			m_DataSource.populate_datasource_array();

			this.m_strPlotTable = this.m_DataSource.getValidDataSourceTableName("PLOT");
			this.m_strRxTable = this.m_DataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS");
			this.m_strCondTable= this.m_DataSource.getValidDataSourceTableName("CONDITION");
			this.m_strFvsTreeTable = this.m_DataSource.getValidDataSourceTableName("FVS TREE LIST FOR PROCESSOR");
			this.m_strTempMDBFile = this.m_DataSource.CreateMDBAndTableDataSourceLinks();



			if (this.m_strPlotTable.Trim().Length == 0) 
			{
				MessageBox.Show("!!Could Not Locate Plot Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strRxTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate Treatment Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strCondTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate Condition Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}
			if (this.m_strFvsTreeTable.Trim().Length == 0)
			{
				MessageBox.Show("!!Could Not Locate FVS Out Processor In Tree Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}

			this.m_ado = new ado_data_access();
			this.m_dao = new dao_data_access();
			this.m_strConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");
			this.m_ado.OpenConnection(this.m_strConn);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_ado = null;
				return ;

			}
			



			htSelectedRxFile=new Hashtable();
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
			this.lblTitle = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAuditDb = new System.Windows.Forms.Button();
			this.btnViewLogFile = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.grpboxAppend = new System.Windows.Forms.GroupBox();
			this.btnAppend = new System.Windows.Forms.Button();
			this.grpboxAudit = new System.Windows.Forms.GroupBox();
			this.btnAudit = new System.Windows.Forms.Button();
			this.lblRunStatus = new System.Windows.Forms.Label();
			this.lstFvsOutput = new System.Windows.Forms.ListView();
			this.txtOutDir = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.btnRefresh = new System.Windows.Forms.Button();
			this.btnClearAll = new System.Windows.Forms.Button();
			this.btnChkAll = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.btnHelp = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.grpboxAppend.SuspendLayout();
			this.grpboxAudit.SuspendLayout();
			this.SuspendLayout();
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(626, 32);
			this.lblTitle.TabIndex = 2;
			this.lblTitle.Text = "Join And Append FVS Out Data ";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(192, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Available FVS Out MDB Tables";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnAuditDb);
			this.groupBox1.Controls.Add(this.btnViewLogFile);
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.grpboxAppend);
			this.groupBox1.Controls.Add(this.grpboxAudit);
			this.groupBox1.Controls.Add(this.lblRunStatus);
			this.groupBox1.Controls.Add(this.lstFvsOutput);
			this.groupBox1.Controls.Add(this.txtOutDir);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.btnRefresh);
			this.groupBox1.Controls.Add(this.btnClearAll);
			this.groupBox1.Controls.Add(this.btnChkAll);
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.btnHelp);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(632, 504);
			this.groupBox1.TabIndex = 5;
			this.groupBox1.TabStop = false;
			// 
			// btnAuditDb
			// 
			this.btnAuditDb.Location = new System.Drawing.Point(360, 288);
			this.btnAuditDb.Name = "btnAuditDb";
			this.btnAuditDb.Size = new System.Drawing.Size(152, 32);
			this.btnAuditDb.TabIndex = 60;
			this.btnAuditDb.Text = "Open Audit Pre/Post Table";
			this.btnAuditDb.Click += new System.EventHandler(this.btnAuditDb_Click);
			// 
			// btnViewLogFile
			// 
			this.btnViewLogFile.Location = new System.Drawing.Point(512, 288);
			this.btnViewLogFile.Name = "btnViewLogFile";
			this.btnViewLogFile.Size = new System.Drawing.Size(112, 32);
			this.btnViewLogFile.TabIndex = 59;
			this.btnViewLogFile.Text = "Open Audit Log File";
			this.btnViewLogFile.Click += new System.EventHandler(this.btnViewLogFile_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(264, 288);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(80, 32);
			this.btnCancel.TabIndex = 58;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Visible = false;
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// grpboxAppend
			// 
			this.grpboxAppend.Controls.Add(this.btnAppend);
			this.grpboxAppend.Location = new System.Drawing.Point(8, 392);
			this.grpboxAppend.Name = "grpboxAppend";
			this.grpboxAppend.Size = new System.Drawing.Size(608, 56);
			this.grpboxAppend.TabIndex = 57;
			this.grpboxAppend.TabStop = false;
			this.grpboxAppend.Text = "Step 2: Append";
			// 
			// btnAppend
			// 
			this.btnAppend.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnAppend.Location = new System.Drawing.Point(3, 16);
			this.btnAppend.Name = "btnAppend";
			this.btnAppend.Size = new System.Drawing.Size(602, 37);
			this.btnAppend.TabIndex = 50;
			this.btnAppend.Text = "Append  FVS Output Data";
			this.btnAppend.Click += new System.EventHandler(this.btnAppend_Click);
			// 
			// grpboxAudit
			// 
			this.grpboxAudit.Controls.Add(this.btnAudit);
			this.grpboxAudit.Location = new System.Drawing.Point(8, 328);
			this.grpboxAudit.Name = "grpboxAudit";
			this.grpboxAudit.Size = new System.Drawing.Size(608, 56);
			this.grpboxAudit.TabIndex = 56;
			this.grpboxAudit.TabStop = false;
			this.grpboxAudit.Text = "Step 1: Audit ";
			// 
			// btnAudit
			// 
			this.btnAudit.Dock = System.Windows.Forms.DockStyle.Fill;
			this.btnAudit.Location = new System.Drawing.Point(3, 16);
			this.btnAudit.Name = "btnAudit";
			this.btnAudit.Size = new System.Drawing.Size(602, 37);
			this.btnAudit.TabIndex = 55;
			this.btnAudit.Text = "Audit Check";
			this.btnAudit.Click += new System.EventHandler(this.btnAudit_Click);
			// 
			// lblRunStatus
			// 
			this.lblRunStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRunStatus.Location = new System.Drawing.Point(144, 464);
			this.lblRunStatus.Name = "lblRunStatus";
			this.lblRunStatus.Size = new System.Drawing.Size(352, 32);
			this.lblRunStatus.TabIndex = 54;
			this.lblRunStatus.Text = "Run Status";
			this.lblRunStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblRunStatus.Visible = false;
			// 
			// lstFvsOutput
			// 
			this.lstFvsOutput.CheckBoxes = true;
			this.lstFvsOutput.GridLines = true;
			this.lstFvsOutput.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.lstFvsOutput.HideSelection = false;
			this.lstFvsOutput.Location = new System.Drawing.Point(8, 88);
			this.lstFvsOutput.MultiSelect = false;
			this.lstFvsOutput.Name = "lstFvsOutput";
			this.lstFvsOutput.Size = new System.Drawing.Size(616, 192);
			this.lstFvsOutput.TabIndex = 53;
			this.lstFvsOutput.View = System.Windows.Forms.View.Details;
			this.lstFvsOutput.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFvsOutput_MouseUp);
			this.lstFvsOutput.SelectedIndexChanged += new System.EventHandler(this.lstFvsOutput_SelectedIndexChanged);
			this.lstFvsOutput.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstFvsOutput_ItemCheck);
			// 
			// txtOutDir
			// 
			this.txtOutDir.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtOutDir.Location = new System.Drawing.Point(152, 58);
			this.txtOutDir.Name = "txtOutDir";
			this.txtOutDir.Size = new System.Drawing.Size(464, 20);
			this.txtOutDir.TabIndex = 52;
			this.txtOutDir.Text = "";
			this.txtOutDir.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtOutDir_KeyPress);
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(16, 58);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(128, 16);
			this.label4.TabIndex = 51;
			this.label4.Text = "FVS Output Directory";
			// 
			// btnRefresh
			// 
			this.btnRefresh.Location = new System.Drawing.Point(136, 288);
			this.btnRefresh.Name = "btnRefresh";
			this.btnRefresh.Size = new System.Drawing.Size(64, 32);
			this.btnRefresh.TabIndex = 49;
			this.btnRefresh.Text = "Refresh";
			this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
			// 
			// btnClearAll
			// 
			this.btnClearAll.Location = new System.Drawing.Point(72, 288);
			this.btnClearAll.Name = "btnClearAll";
			this.btnClearAll.Size = new System.Drawing.Size(64, 32);
			this.btnClearAll.TabIndex = 48;
			this.btnClearAll.Text = "Clear All";
			this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
			// 
			// btnChkAll
			// 
			this.btnChkAll.Location = new System.Drawing.Point(8, 288);
			this.btnChkAll.Name = "btnChkAll";
			this.btnChkAll.Size = new System.Drawing.Size(64, 32);
			this.btnChkAll.TabIndex = 47;
			this.btnChkAll.Text = "Check All";
			this.btnChkAll.Click += new System.EventHandler(this.btnChkAll_Click);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(520, 464);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 45;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// btnHelp
			// 
			this.btnHelp.Location = new System.Drawing.Point(16, 464);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(96, 32);
			this.btnHelp.TabIndex = 44;
			this.btnHelp.Text = "Help";
			// 
			// uc_PostFvsForeFrcs
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_PostFvsForeFrcs";
			this.Size = new System.Drawing.Size(632, 504);
			this.Resize += new System.EventHandler(this.uc_PostFvsForeFrcs_Resize);
			this.groupBox1.ResumeLayout(false);
			this.grpboxAppend.ResumeLayout(false);
			this.grpboxAudit.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		#region Private Methods
		//((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory
		private void HandleDisplayRxSelection(object sender, EventArgs e)
		{                         
			if(htSelectedRxFile!=null)
			{
				htSelectedRxFile.Clear();
				htSelectedRxFile=procPreFrcs.FioIsRxExist(@"C:\FiaBiosumProjects\fia_biosum\orca_converted_data\fvs\db\out");          
			}
			Array aryT=Array.CreateInstance(Type.GetType("System.String"), htSelectedRxFile.Keys.Count);
			htSelectedRxFile.Keys.CopyTo(aryT, 0);
			Array.Sort(aryT);

		}
	  
      
		private void HandleSelectedRx(object sender, EventArgs e)
		{
			if(htSelectedRxFile!=null && htSelectedRxFile.Count >1)
			{
				//foreach(string item in listBox1.SelectedItems)
				//{
				// procPreFrcs.WfSelectedRxList.Add((string)htSelectedRxFile[item.Trim()]); //adds to the selected list of RX 
				//}
				//UnitTest();
			}
			else
				return;
		}

		#region UNIT TEST
		private void UnitTest()
		{
			try
			{
				/*//1) UpdateCondPreValues
				string strCIDpotFireStandId="4047C01";
				string strCIDmaster="2001006000000700000075061";
				procPreFrcs.DbRecsGetCondTabPreValues(strCIDmaster, strCIDpotFireStandId, procPreFrcs.WfSelectedRxList[0].ToString());                 
				*/
            
				//2) DbRecsTdgTsg2ProcDbCopy
				//procPreFrcs.DbRecsTdgTsg2ProcDbCopy(true);   

				//3) DbRecsFinalizeFvsTree
				//procPreFrcs.DbRecsFinalizeFvsTree(procPreFrcs.WfSelectedRxList[0].ToString());

				//4) DbRecsUpdateFfeFromPotFire
				procPreFrcs.DbRecsUpdateFfeFromPotFire(procPreFrcs.WfSelectedRxList[0].ToString());
            
			}
			catch(Exception e)
			{
				MessageBox.Show(e.Message,"UnitTest");
			}
		}



		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();
		}

		#endregion
		public void loadvalues()
		{

			string strOutDirAndFile;
			string strConn;
			int x;
			int intCount=0;
			try
			{
				System.Windows.Forms.ListViewItem entryListItem=null;
				this.m_oLvAlternateColors.InitializeRowCollection();
				this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
				this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceForegroundColor=frmMain.g_oGridViewRowForegroundColor;
				this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceListView = lstFvsOutput;
				this.m_oLvAlternateColors.CustomFullRowSelect=true;
				if (frmMain.g_oGridViewFont!=null) this.lstFvsOutput.Font = frmMain.g_oGridViewFont;
				this.lstFvsOutput.Clear();
				this.lstFvsOutput.Columns.Add("",2,HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("FVS Variant", 80, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Rx", 30, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Run Status",250,HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Output File", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("File Found", 80, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Summary Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Tree Cut List Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Tree Standing Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Potential Fire Count", 100, HorizontalAlignment.Left);

				this.lstFvsOutput.Columns[COL_CHECKBOX].Width = -2;

				string strLinkTableName="";
				string[] strTableNames=null;
				string[] strLinkedTables = new string[1000];
				string strCurVariantRx="";
				string strVariantRx="";
				

				string strTempDbFile="";
				dao_data_access oDao = new dao_data_access();
				dao_data_access oDao2 = new dao_data_access();
				strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
				oDao.CreateMDB(strTempDbFile);
				oDao.OpenDb(strTempDbFile);


				this.m_ado.m_strSQL = "SELECT DISTINCT  a.fvs_variant,  b.rx  " + 
					"FROM " + this.m_strPlotTable + " a, " + 
					"(SELECT rx " + 
					"FROM " + this.m_strRxTable +  ") b " + 
					"WHERE a.fvs_variant IS NOT NULL AND LEN(TRIM(a.fvs_variant)) > 0 AND b.rx IS NOT NULL AND LEN(TRIM(b.rx)) > 0 ORDER BY a.fvs_variant, b.rx;";
				this.m_ado.SqlQueryReader(this.m_strConn,this.m_ado.m_strSQL);

				while (this.m_ado.m_OleDbDataReader.Read())
				{
					
					
					this.m_strOutMDBFile = "FVSOUT_" + this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "_" + this.m_ado.m_OleDbDataReader["rx"].ToString().Trim() + ".MDB";
					strOutDirAndFile=this.txtOutDir.Text.Trim() + "\\" + this.m_strOutMDBFile.Trim();
					


					if (System.IO.File.Exists(strOutDirAndFile) == true)
					{

						strTableNames = new string[100];						
						oDao2.getTableNames(strOutDirAndFile,ref strTableNames);
						for (x=0;x<=strTableNames.Length-1;x++)
						{
							if (strTableNames[x]==null || strTableNames[x].Trim().Length==0) break;
							if (strTableNames[x].Trim().ToUpper()=="FVS_SUMMARY" ||
								strTableNames[x].Trim().ToUpper()=="FVS_TREELIST" ||
								strTableNames[x].Trim().ToUpper()=="FVS_CUTLIST" ||
								strTableNames[x].Trim().ToUpper()=="FVS_POTFIRE")
							{
								strLinkTableName=strTableNames[x].Trim() + "_" + m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "_" + m_ado.m_OleDbDataReader["rx"].ToString().Trim();
								oDao.CreateTableLink(oDao.m_DaoDatabase,strLinkTableName,strOutDirAndFile,strTableNames[x]);
								intCount++;
								strLinkedTables[intCount-1]=strLinkTableName;

							}
						}
					}
					else
					{
						intCount++;
						strLinkedTables[intCount-1] = "FILENOTFOUND_" + m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "_" + m_ado.m_OleDbDataReader["rx"].ToString().Trim();
						
					}
				}
				oDao2.m_DaoWorkspace.Close();
				oDao.m_DaoDatabase.Close();
				oDao.m_DaoWorkspace.Close();
				m_ado.m_OleDbDataReader.Close();

				ado_data_access oAdo = new ado_data_access();

				oAdo.OpenConnection(oAdo.getMDBConnString(strTempDbFile,"",""));


				for (x=0;x<=intCount-1;x++)
				{
					strVariantRx=strLinkedTables[x].Substring(strLinkedTables[x].Length-1,1);
					strVariantRx=strLinkedTables[x].Substring(strLinkedTables[x].Length-4,2) + strVariantRx;
					if (strVariantRx.Trim() != strCurVariantRx.Trim())
					{
						
						this.m_strOutMDBFile = "FVSOUT_" + strVariantRx.Substring(0,2) + "_" + strVariantRx.Substring(2,1) + ".MDB";
						frmMain.g_sbpInfo.Text = "Loading FVS Output file " + this.m_strOutMDBFile + "...Stand By";
						frmMain.g_sbpInfo.Parent.Refresh();

						// Add a ListItem object to the ListView.
						entryListItem =
							this.lstFvsOutput.Items.Add("");
					
						entryListItem.UseItemStyleForSubItems=false;
						this.m_oLvAlternateColors.AddRow();
						this.m_oLvAlternateColors.AddColumns(lstFvsOutput.Items.Count-1,lstFvsOutput.Columns.Count);
						entryListItem.SubItems.Add(strVariantRx.Substring(0,2));
						entryListItem.SubItems.Add(strVariantRx.Substring(2,1));
						entryListItem.SubItems.Add(" ");  //out mdb file
						entryListItem.SubItems.Add(" ");  //file found
						entryListItem.SubItems.Add(" ");  //summary record count
						entryListItem.SubItems.Add(" ");  //tree cut list record count
						entryListItem.SubItems.Add(" ");  //tree standing record count
						entryListItem.SubItems.Add(" ");  //potential fire record count
						entryListItem.SubItems.Add(" ");  //run status
						strCurVariantRx=strVariantRx;

						if (strLinkedTables[x].IndexOf("FILENOTFOUND",0)==0)
						{
							entryListItem.SubItems[COL_FOUND].ForeColor = System.Drawing.Color.White;
							entryListItem.SubItems[COL_FOUND].BackColor = System.Drawing.Color.Red;
							this.m_oLvAlternateColors.m_oRowCollection.Item(this.lstFvsOutput.Items.Count-1).m_oColumnCollection.Item(COL_FOUND).UpdateColumn=false;
							entryListItem.SubItems[COL_FOUND].Text = "No";
						}
						else
						{
							entryListItem.SubItems[COL_FOUND].Text = "Yes";
						}
						entryListItem.SubItems[COL_MDBOUT].Text = this.m_strOutMDBFile;
						this.m_oLvAlternateColors.ListViewItem(lstFvsOutput.Items[lstFvsOutput.Items.Count-1]);
						
					}
					if (strLinkedTables[x].IndexOf("FILENOTFOUND",0)!=0)
					{
						if (strLinkedTables[x].ToUpper().IndexOf("FVS_SUMMARY",0)==0)
						{
							entryListItem.SubItems[COL_SUMMARYCOUNT].Text = Convert.ToString(Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from " + strLinkedTables[x],"fvs_summary")));
						}
						else if (strLinkedTables[x].ToUpper().IndexOf("FVS_CUTLIST",0)==0)
						{
							entryListItem.SubItems[COL_CUTCOUNT].Text = Convert.ToString(Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from " + strLinkedTables[x],"fvs_cutlist")));
						}
						else if (strLinkedTables[x].ToUpper().IndexOf("FVS_TREELIST",0)==0)
						{
							entryListItem.SubItems[COL_LEFTCOUNT].Text = Convert.ToString(Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from " + strLinkedTables[x],"fvs_treelist")));
						}
						else if (strLinkedTables[x].ToUpper().IndexOf("FVS_POTFIRE",0)==0)
						{
							entryListItem.SubItems[COL_POTFIRECOUNT].Text = Convert.ToString(Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from " + strLinkedTables[x],"fvs_potfire")));
						}
					}
				
				}
			    oAdo.CloseConnection(oAdo.m_OleDbConnection);

			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_PostFvsForeFrcs:loadvalues() \n" + 
					"Err Msg - " + e.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
			this.Refresh();



		}

		public void loadvalues_old()
		{

			string strOutDirAndFile;
			string strConn;
			int x;
			int intCount=0;
			try
			{

				this.m_ado.m_strSQL = "SELECT DISTINCT  a.fvs_variant,  b.rx  " + 
					"FROM " + this.m_strPlotTable + " a, " + 
					"(SELECT rx " + 
					"FROM " + this.m_strRxTable +  ") b " + 
					"WHERE a.fvs_variant IS NOT NULL AND LEN(TRIM(a.fvs_variant)) > 0 AND b.rx IS NOT NULL AND LEN(TRIM(b.rx)) > 0 ORDER BY a.fvs_variant, b.rx;";
				this.m_ado.SqlQueryReader(this.m_strConn,this.m_ado.m_strSQL);

				this.m_oLvAlternateColors.InitializeRowCollection();
				this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
				this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceForegroundColor=frmMain.g_oGridViewRowForegroundColor;
				this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceListView = lstFvsOutput;
				this.m_oLvAlternateColors.CustomFullRowSelect=true;
				//lstFvsOutput.ForeColor = frmMain.g_oGridViewRowForegroundColor;
				if (frmMain.g_oGridViewFont!=null) this.lstFvsOutput.Font = frmMain.g_oGridViewFont;
				this.lstFvsOutput.Clear();
				this.lstFvsOutput.Columns.Add("",2,HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("FVS Variant", 80, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Rx", 30, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Run Status",250,HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Output File", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("File Found", 80, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Summary Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Tree Cut List Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Tree Standing Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Potential Fire Count", 100, HorizontalAlignment.Left);

				this.lstFvsOutput.Columns[COL_CHECKBOX].Width = -2;
				

				while (this.m_ado.m_OleDbDataReader.Read())
				{
					// Add a ListItem object to the ListView.
					System.Windows.Forms.ListViewItem entryListItem =
						this.lstFvsOutput.Items.Add("");
					
					//System.Windows.Forms.ListViewItem entryListItem =
					//	this.lstFvsOutput.Items.Add(this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim());
					entryListItem.UseItemStyleForSubItems=false;
					this.m_oLvAlternateColors.AddRow();
					this.m_oLvAlternateColors.AddColumns(lstFvsOutput.Items.Count-1,lstFvsOutput.Columns.Count);
					entryListItem.SubItems.Add(this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim());
					entryListItem.SubItems.Add(this.m_ado.m_OleDbDataReader["rx"].ToString().Trim());
					entryListItem.SubItems.Add(" ");  //out mdb file
					entryListItem.SubItems.Add(" ");  //file found
					entryListItem.SubItems.Add(" ");  //summary record count
					//entryListItem.SubItems.Add(" ");  //summary pre treatment year
					//entryListItem.SubItems.Add(" ");  //summary post treatment year
					entryListItem.SubItems.Add(" ");  //tree cut list record count
					entryListItem.SubItems.Add(" ");  //tree standing record count
					entryListItem.SubItems.Add(" ");  //potential fire record count
					entryListItem.SubItems.Add(" ");  //run status
					


					this.m_strOutMDBFile = "FVSOUT_" + this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "_" + this.m_ado.m_OleDbDataReader["rx"].ToString().Trim() + ".MDB";
					strOutDirAndFile=this.txtOutDir.Text.Trim() + "\\" + this.m_strOutMDBFile.Trim();

					if (System.IO.File.Exists(strOutDirAndFile) == true)
					{
						frmMain.g_sbpInfo.Text = "Loading FVS Output file " + this.m_strOutMDBFile + "...Stand By";
						frmMain.g_sbpInfo.Parent.Refresh();
						//entryListItem.SubItems[COL_FOUND].ForeColor = System.Drawing.Color.Black;
						//if ((entryListItem.Index % 2)==0)
						//	entryListItem.SubItems[COL_FOUND].BackColor = frmMain.g_oGridViewRowBackgroundColor;
						//else
						//	entryListItem.SubItems[COL_FOUND].BackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;

						entryListItem.SubItems[COL_FOUND].Text = "Yes";
						//entryListItem.SubItems[COL_FOUND].ForeColor = frmMain.g_oGridViewRowForegroundColor;
						
						strConn = this.m_ado.getMDBConnString(strOutDirAndFile,"","");
						FIA_Biosum_Manager.ado_data_access oAdo = new ado_data_access();
						oAdo.OpenConnection(strConn);
						entryListItem.SubItems[COL_MDBOUT].Text = this.m_strOutMDBFile;
						if (oAdo.TableExist(oAdo.m_OleDbConnection,"fvs_summary"))
						{
							entryListItem.SubItems[COL_SUMMARYCOUNT].Text = Convert.ToString(Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from fvs_summary","fvs_summary")));
							//int intPreRxYear=-1;
							//int intPostRxYear=-1;
							//int intError = Validate_FvsSummaryPrePostTreatmentYear(oAdo,oAdo.m_OleDbConnection,"FVS_SUMMARY",ref intPreRxYear,ref intPostRxYear);
							//if (intPreRxYear==-1)
							//{
							//}
							//else
							//{
							//	entryListItem.SubItems[COL_SUMMARYPREYEAR].Text = intPreRxYear.ToString().Trim();
							//}
							//if (intPostRxYear==-1)
							//{
							//}
							//else
							//{
							//	entryListItem.SubItems[COL_SUMMARYPOSTYEAR].Text = intPostRxYear.ToString().Trim();
							//}

							
						}
						if (this.m_dao.TableExists(strOutDirAndFile,"fvs_cutlist") == true)
						{
							//intCount=Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from (SELECT TOP 1 * FROM fvs_cutlist)","fvs_cutlist"));
							//if (intCount > 0) entryListItem.SubItems[COL_CUTCOUNT].Text = "Has Records";
							//else entryListItem.SubItems[COL_CUTCOUNT].Text = "0";
							entryListItem.SubItems[COL_CUTCOUNT].Text = Convert.ToString(Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from fvs_cutlist","fvs_cutlist")));
						}
						if (this.m_dao.TableExists(strOutDirAndFile,"fvs_treelist") == true)
						{
							//intCount=Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from (SELECT TOP 1 * FROM fvs_treelist)","fvs_treelist"));
							//if (intCount > 0) entryListItem.SubItems[COL_LEFTCOUNT].Text = "Has Records";
							//else entryListItem.SubItems[COL_LEFTCOUNT].Text = "0";

							entryListItem.SubItems[COL_LEFTCOUNT].Text = Convert.ToString(Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from fvs_treelist","fvs_treelist")));
						}
						if (this.m_dao.TableExists(strOutDirAndFile,"fvs_potfire") == true)
						{
							//intCount=Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from (SELECT TOP 1 * FROM fvs_potfire)","fvs_potfire"));
							//if (intCount > 0) entryListItem.SubItems[COL_POTFIRECOUNT].Text = "Has Records";
							//else entryListItem.SubItems[COL_POTFIRECOUNT].Text = "0";

							entryListItem.SubItems[COL_POTFIRECOUNT].Text = Convert.ToString(Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from fvs_potfire","fvs_potfire")));
						}
						oAdo.CloseConnection(oAdo.m_OleDbConnection);
					}
					else
					{
						entryListItem.SubItems[COL_FOUND].ForeColor = System.Drawing.Color.White;
						entryListItem.SubItems[COL_FOUND].BackColor = System.Drawing.Color.Red;
						this.m_oLvAlternateColors.m_oRowCollection.Item(this.lstFvsOutput.Items.Count-1).m_oColumnCollection.Item(COL_FOUND).UpdateColumn=false;
						entryListItem.SubItems[COL_FOUND].Text = "No";
					}
					this.m_oLvAlternateColors.ListViewItem(lstFvsOutput.Items[lstFvsOutput.Items.Count-1]);

				}
				this.m_ado.m_OleDbDataReader.Close();



				//check if condition table has been updated with the pretreatment fire potential values
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_PostFvsForeFrcs:loadvalues() \n" + 
					"Err Msg - " + e.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
			this.Refresh();



		}

		private void txtOutDir_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void textBox1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void btnRefresh_Click(object sender, System.EventArgs e)
		{
			loadvalues();
		}

		private void btnChkAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstFvsOutput.Items.Count-1;x++)
			{
				this.lstFvsOutput.Items[x].Checked=true;
			}
		}

		private void btnClearAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstFvsOutput.Items.Count-1;x++)
			{
				this.lstFvsOutput.Items[x].Checked=false;
			}

		}

		private void uc_PostFvsForeFrcs_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.btnHelp.Top = this.btnClose.Top;
				this.lstFvsOutput.Left = 5;
				this.lstFvsOutput.Width = this.Width - 10;
				this.grpboxAppend.Top = this.btnClose.Top - this.grpboxAppend.Height - 5;
				this.grpboxAudit.Top = this.grpboxAppend.Top - this.grpboxAudit.Height - 5;
				this.btnChkAll.Top = this.grpboxAudit.Top - btnChkAll.Height - 5;
				this.btnClearAll.Top = btnChkAll.Top;
				this.btnRefresh.Top = btnChkAll.Top;
				this.lstFvsOutput.Height = this.btnChkAll.Top - this.lstFvsOutput.Top - 5;
				this.grpboxAudit.Width = this.lstFvsOutput.Width;
				this.grpboxAppend.Width = this.lstFvsOutput.Width;
				this.btnCancel.Top = this.btnChkAll.Top;
				this.btnCancel.Left = this.ClientSize.Width / 2 - (int)(btnCancel.Width * .5);
				this.btnViewLogFile.Top = this.btnChkAll.Top;
				this.btnViewLogFile.Left = this.lstFvsOutput.ClientSize.Width - (this.lstFvsOutput.Left * 2) - this.btnViewLogFile.Width;
				btnAuditDb.Top = this.btnChkAll.Top;
				this.btnAuditDb.Left = this.btnViewLogFile.Left - this.btnAuditDb.Width;
				this.txtOutDir.Width = this.Width - this.txtOutDir.Left - 10;

				this.lblRunStatus.Top = this.btnClose.Top;
				this.lblRunStatus.Left = (int)(this.groupBox1.Width / 2) - (int)(this.lblRunStatus.Width / 2);


			}
			catch
			{
			}
		}

		private void btnAppend_Click(object sender, System.EventArgs e)
		{
			if (this.lstFvsOutput.CheckedItems.Count ==0)
			{
				MessageBox.Show("No Boxes Are Checked","FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			this.val_data();
			if (this.m_intError==0)
			{
				

				this.m_frmTherm = new frmTherm(((frmDialog)ParentForm),"FVS OUT DATA",
					"FVS Output","2");
				this.m_frmTherm.lblMsg.Text="";
				this.Enabled=false;
				this.grpboxAppend.Enabled=false;
				this.grpboxAudit.Enabled=false;
				this.btnChkAll.Enabled=false;
				this.btnClearAll.Enabled=false;
				this.btnRefresh.Enabled=false;
				this.btnClose.Enabled=false;
				this.btnHelp.Enabled=false;
				this.btnCancel.Visible=false;
				this.btnViewLogFile.Enabled=false;
				this.btnAuditDb.Enabled=false;

				
				frmMain.g_oDelegate.CurrentThreadProcessAborted=false;
				frmMain.g_oDelegate.CurrentThreadProcessDone=false;
				frmMain.g_oDelegate.CurrentThreadProcessStarted=false;
				frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(AppendAndUpdateRecords));
				frmMain.g_oDelegate.InitializeThreadEvents();
				frmMain.g_oDelegate.m_oThread.IsBackground=true;
				frmMain.g_oDelegate.CurrentThreadProcessIdle=false;
				frmMain.g_oDelegate.m_oThread.Start();
				

			}
			
		}
		private void val_data()
		{
			this.m_intError=0;
			for (int x=0; x<=this.lstFvsOutput.Items.Count-1;x++)
			{
				if (this.lstFvsOutput.Items[x].Checked==true)
				{
					if (this.lstFvsOutput.Items[x].SubItems[COL_FOUND].Text.Trim() == "No")
					{
						MessageBox.Show("!!File " + 
							this.lstFvsOutput.Items[x].SubItems[COL_MDBOUT].Text.Trim() + " Not Found!!",
							"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						this.m_intError=-1;
						break;
					}
					if (this.lstFvsOutput.Items[x].SubItems[COL_SUMMARYCOUNT].Text.Trim().Length  == 0 ||
						this.lstFvsOutput.Items[x].SubItems[COL_SUMMARYCOUNT].Text.Trim() == "0")
					{
						MessageBox.Show("!!Summary Table In File  " + this.lstFvsOutput.Items[x].SubItems[COL_MDBOUT].Text.Trim() + " " + 
							" Does Not Exist Or Has 0 Records!!",
							"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						this.m_intError=-1;
						break;
					}
					
					if (this.lstFvsOutput.Items[x].SubItems[COL_CUTCOUNT].Text.Trim().Length  == 0)
					{
						MessageBox.Show("!!Cut Tree Table In File  " + this.lstFvsOutput.Items[x].SubItems[COL_MDBOUT].Text.Trim() + " " + 
							" Does Not Exist!!",
							"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						this.m_intError=-1;
						break;
					}
					if (this.lstFvsOutput.Items[x].SubItems[COL_LEFTCOUNT].Text.Trim().Length  == 0)
					{
						MessageBox.Show("!!Tree NOT Cut Table In File  " + this.lstFvsOutput.Items[x].SubItems[COL_MDBOUT].Text.Trim() + " " + 
							" Does Not Exist!!",
							"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						this.m_intError=-1;
						break;
					}

					if (this.lstFvsOutput.Items[x].SubItems[COL_POTFIRECOUNT].Text.Trim().Length  == 0)
					{
						MessageBox.Show("!!Potential Fire Table In File  " + this.lstFvsOutput.Items[x].SubItems[COL_MDBOUT].Text.Trim() + " " + 
							" Does Not Exist!!",
							"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						this.m_intError=-1;
						break;
					}
                  

				}
			}
			
		}
		private void Validation(ado_data_access p_oAdo, 
								System.Data.OleDb.OleDbConnection p_oConn,
								string[] p_strFVSTableIdArray,
								string[] p_strFVSTableNameArray,
			                    int p_intListViewItem,
								string p_strVariant,string p_strRx,bool p_bAudit,
								ref int p_intError,ref string p_strError,ref int p_intWarning,ref string p_strWarning)
		{
			int x,y;
			bool bBadVariant=false;
			string strTableId="";
			string strTableName="";
			string strTableName2="";
			string strSummaryTableName="";
			//
			//fvs_cases
			//
			//make sure table exists
			for (x=0;x<=p_strFVSTableIdArray.Length - 1;x++)
			{
				if (p_strFVSTableIdArray[x].Trim().ToUpper() == "FVS_CASES")
					break;
			}
			if (x>p_strFVSTableIdArray.Length - 1)
			{
				p_intError=-1;
				p_strError = "FVS_Cases table missing";
				return;
			}
			//make sure only one fvs variant variable exists
			strTableName = p_strFVSTableNameArray[x].Trim();
			bBadVariant = p_oAdo.ValuesExistNotEqualToTargetValue(p_oConn,
				strTableName,
				"variant",
				p_strVariant.Trim(),
				false);

			if (p_oAdo.m_intError==p_oAdo.ErrorCodeNoErrors && bBadVariant==true)
			{
				p_intError=-1;
				p_strError = "Incorrect variant found in FVS_Cases.variant column";
			}
			else if (p_oAdo.m_intError==p_oAdo.ErrorCodeTableNotFound)
			{
				p_intError=-1;
				p_strError = "FVS_Cases table missing";
			}
			else if (p_oAdo.m_intError==p_oAdo.ErrorCodeColumnNotFound)
			{
				p_intError=-1;
				p_strError="FVS_Cases.variant column not found";
							
			}
			if (p_intError<0) return;
			//
			//fvs summary
			//
			//make sure table exists
			for (x=0;x<=p_strFVSTableIdArray.Length - 1;x++)
			{
				if (p_strFVSTableIdArray[x].Trim().ToUpper() == "FVS_SUMMARY")
					break;
			}
			if (x>p_strFVSTableIdArray.Length - 1)
			{
				p_intError=-1;
				p_strError = "FVS_Summary table missing";
				return;
			}
			strSummaryTableName = p_strFVSTableNameArray[x].Trim();
			if (!p_bAudit) frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,p_intListViewItem,COL_RUNSTATUS,"Text","Processing...FVS_SUMMARY");
			if (p_oAdo.TableExist(p_oConn,strSummaryTableName))
			{
				this.CreateSummaryTableFVSPrePostYearWorkTables(p_oAdo,p_oConn,strSummaryTableName);
				this.Validate_FvsSummaryPrePostTreatmentYear(p_oAdo,p_oConn,ref p_intError,ref p_strError,ref p_intWarning,ref p_strWarning,false);
				if (p_intError !=0)
				{
					switch (p_intError)
					{
						case -2:
							p_strError="FVS_Summary table has no records";
							break;
						case -3:
							p_strError="FVS_Summary table has pre-treatment year null value detected";
							break;
						case -4:
							p_strError="FVS_Summary table has post-treatment year null value detected";
							break;
					}

				}
			}
			else
			{
				p_intError=-1;
				p_strError="FVS_Summary table missing";
			}
			if (p_intError<0) return;

			GetSummaryPrePostCounts(p_oAdo,p_oConn,p_strFVSTableIdArray,p_strFVSTableNameArray);

			strTableName="";
			strTableName2="";
			for (x=0;x<=p_strFVSTableIdArray.Length - 1;x++)
			{
				if (p_strFVSTableIdArray[x].Trim().ToUpper() == "FVS_TREELIST")
				{
					strTableName = p_strFVSTableNameArray[x];
					break;
				}
			}
			
			for (x=0;x<=p_strFVSTableIdArray.Length - 1;x++)
			{
				if (p_strFVSTableIdArray[x].Trim().ToUpper() == "FVS_CUTLIST")
				{
					strTableName2 = p_strFVSTableNameArray[x];
					break;
				}
			}
			
			this.CreateTreeListTableFVSPrePostYearWorkTables(p_oAdo,p_oConn,strTableName,strTableName2);
			for (y=0;y<=p_strFVSTableIdArray.Length-1;y++)
			{
				if (p_strFVSTableIdArray[y] == null) break;

				if (p_strFVSTableIdArray[y].Trim().ToUpper() == "FVS_SUMMARY" ||
					p_strFVSTableIdArray[y].Trim().ToUpper() == "FVS_CASES")
				{
				}
				else if (p_strFVSTableIdArray[y].Trim().ToUpper() == "FVS_TREELIST" && p_bAudit)
				{
					if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,"-----FVS_TREELIST-----\r\n");
					if (p_bAudit) frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,p_intListViewItem,COL_RUNSTATUS,"Text","Processing Audit...FVS_Treelist");

									

					p_intWarning=0;
					p_strWarning="";

					this.Validate_PreTreatmentYearForTreeList(p_oAdo,p_oConn,ref p_intError,ref p_strError,ref p_intWarning,ref p_strWarning,true);

					if (p_intError==0 && p_intWarning==0)
					{
						if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
					}
					else if (p_intWarning !=0)
					{
						if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,p_strWarning + "\r\n");
					}

									

					if (p_intError !=0)
					{
						if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,p_strError + "\r\n");
						if (p_intError==-3)
						{
							p_strError = "FVS_Treelist Minimum treatment year not found in FVS_Summary table (See Log File)";
							//p_strItemDialogMsg = strItemDialogMsg  + strDbFile + ":FVS_Treelist Minimum treatment year not found in FVS_Summary table\r\n";
						}
					}

									
				}
				else if (p_strFVSTableIdArray[y].Trim().ToUpper() == "FVS_CUTLIST")
				{
					frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,p_intListViewItem,COL_RUNSTATUS,"Text","Processing...FVS_Cutlist");

					p_intWarning=0;
					p_strWarning="";

					this.Validate_PostTreatmentYearForCutList(p_oAdo,p_oConn,p_strFVSTableNameArray[y],strSummaryTableName,ref p_intError,ref p_strError,ref p_intWarning,ref p_strWarning,false);

					if (p_intError !=0)
					{
						if (p_intError==-3)
						{
							p_strError = "FVS_Treelist Minimum treatment year not found in FVS_Summary table (See Log File)";
							//strItemDialogMsg = strItemDialogMsg  + strDbFile + ":FVS_Treelist Minimum treatment year not found in FVS_Summary table\r\n";
						}
						else if (p_intError==-4)
						{
							p_strError = "FVS_Cutlist Standid and year not found in the fvs_summary table";
							//strItemDialogMsg = strItemDialogMsg  + strDbFile + ":FVS_Cutlist standid and year not found in FVS_Summary table\r\n";

						}
					}

									
				}
				else if (p_strFVSTableIdArray[y].Trim().ToUpper() == "FVS_POTFIRE")
				{
					if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,"-----FVS_POTFIRE-----\r\n");
					if (p_bAudit) frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing Audit...FVS_Potfire");

					CreatePotFireTablePrePostYearWorkTables(p_oAdo,p_oConn,p_strFVSTableNameArray[y]);

					p_intWarning=0;
					p_strWarning="";

					this.Validate_PotFire(p_oAdo,p_oConn,p_strFVSTableNameArray[y],ref p_intError,ref p_strError,ref p_intWarning,ref p_strWarning,p_bAudit);

					if (p_intError==0 && p_intWarning==0)
					{
						if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
					}
					else if (p_intWarning !=0)
					{
						if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,p_strWarning + "\r\n");
					}

									

					if (p_intError !=0)
					{
						if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,p_strError + "\r\n");
					}
				}

					//else if (p_strFVSTableIdArray[y].Trim().ToUpper() == "FVS_CUTLIST")
					//{
					//	p_intError = this.Validate_PostTreatmentYearForCutList(p_oAdo,p_oConn);
					//	if (p_intError==-1)
					//	{
					//		this.m_strError=this.m_strError +  strDbFile + "." + p_strFVSTableIdArray[y].Trim() + ": post-treatment year not found\r\n";
					//		frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"BackColor",Color.Red);
					//		frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Audit Error: post-treatment year not found in " + p_strFVSTableIdArray[y].Trim() + " table");
					//	}
					//}
				else
				{
					if (p_strFVSTableIdArray[y].Substring(0,4) == "FVS_")
					{
						if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,"-----" + p_strFVSTableIdArray[y].Trim().ToUpper() + "-----\r\n");
						if (p_bAudit) frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing Audit..." + p_strFVSTableIdArray[y].Trim());

						CreateGenericTablePrePostYearWorkTables(p_oAdo,p_oConn,p_strFVSTableNameArray[y]);

						p_intWarning=0;
						p_strWarning="";

						this.Validate_FVSGenericTable(p_oAdo,p_oConn,p_strFVSTableNameArray[y].Trim(),ref p_intError,ref p_strError,ref p_intWarning,ref p_strWarning,true);

						if (p_intError==0 && p_intWarning==0)
						{
							if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
						}
						else if (p_intWarning !=0)
						{
							if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,p_strWarning + "\r\n");
						}

									

						if (p_intError !=0)
						{
							if (p_bAudit) frmMain.g_oUtils.WriteText(m_strLogFile,p_strError + "\r\n");
						}					}
				}
				if (p_intError !=0) break;
			}

		}
		private void PrePostTableStructureEdits(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strWorkDbFile, 
			                                    string p_strVariant, string p_strRx, bool p_bUpdatePreTableWithVariant,
												int p_intListViewItem,string[] p_strTableIdArray,string[] p_strTableNameArray,ref int p_intError,ref string p_strError)
		{
			    int x,y,z,zz;
				string strFVSOutPath="";
				string strSourceColumnsList="";
				string strSourceColumnsReservedWordFormattedList="";
				string[] strSourceColumnsArray=null;
				string[] strSourceTableArray=null;
				string[] strSourceTableNameArray=null;
				string strDestColumnsList="";
				string[] strDestColumnsArray=null;
				string[] strDestTableArray=null;
				System.Data.DataTable oDataTableSchema;
				bool bFound;
				//
				//make sure all the tables and columns exist
				//
				p_oAdo.m_strSQL="";
				for (y=0;y<=p_strTableIdArray.Length-1;y++)
				{
					if (p_strTableIdArray[y] != null)
					{
						if (p_strTableIdArray[y].Trim().ToUpper() != "FVS_CASES" &&
							p_strTableIdArray[y].Trim().ToUpper() != "FVS_TREELIST" &&
							p_strTableIdArray[y].Trim().ToUpper() != "FVS_CUTLIST" &&
							p_strTableIdArray[y].Trim().ToUpper() != "FVS_ATRTLIST")
						{
								if (!p_oAdo.TableExist(p_oConn,"PRE_" + p_strTableIdArray[y].Trim()))
								{
									//create the table
									oDataTableSchema=p_oAdo.getTableSchema(p_oConn,"SELECT * FROM " + p_strTableNameArray[y].Trim());
									strSourceColumnsList = p_oAdo.getFieldNames(p_oConn,"SELECT * FROM " + p_strTableNameArray[y].Trim());
									strSourceColumnsReservedWordFormattedList = p_oAdo.FormatReservedWordsInColumnNameList(strSourceColumnsList,",");
									strSourceColumnsArray = m_oUtils.ConvertListToArray(strSourceColumnsList,",");
									p_oAdo.m_strSQL="";
									for (z=0;z<=oDataTableSchema.Rows.Count-1;z++)
									{
										p_oAdo.m_strSQL = p_oAdo.m_strSQL + 
											  p_oAdo.FormatCreateTableSqlFieldItem(oDataTableSchema.Rows[z]) + ",";
									}
									if (p_oAdo.m_strSQL.Trim().Length > 0)
									{
										p_oAdo.m_strSQL = p_oAdo.m_strSQL.Substring(0,p_oAdo.m_strSQL.Length - 1);
										p_oAdo.m_strSQL = p_strTableIdArray[y].Trim() + " (biosum_cond_id text(25),rx text(1),fvs_variant text(2)," + 
											p_oAdo.m_strSQL + ")";

										if (m_bDebug)
											this.WriteText("c:\\tmp\\biosum_debug.txt",p_oAdo.m_strSQL + "\r\n\r\n");
										p_oAdo.SqlNonQuery(p_oConn,"CREATE TABLE PRE_" + p_oAdo.m_strSQL);
										p_oAdo.SqlNonQuery(p_oConn,"CREATE TABLE POST_" + p_oAdo.m_strSQL);


										p_oAdo.m_strSQL = "ALTER TABLE PRE_" + p_strTableIdArray[y].Trim() + " DROP COLUMN RX";
										p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

										p_oAdo.m_strSQL = "CREATE INDEX biosumcondididx_pre ON PRE_" + p_strTableIdArray[y].Trim() + " (biosum_cond_id)";
										p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

										p_oAdo.m_strSQL = "CREATE INDEX biosumcondididx_post ON POST_" + p_strTableIdArray[y].Trim() + " (biosum_cond_id)";
										p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

										p_oAdo.m_strSQL = "CREATE INDEX biosumcondidrxidx_post ON POST_" + p_strTableIdArray[y].Trim() + " (biosum_cond_id,rx)";
										p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

									}
								}
								else
								{
									//see if columns are the same
									//p_strTableNameArray = m_strOutMDBFile.Trim() + "_" + p_strTableIdArray[y].Trim();
									//p_strTableNameArray = p_strTableNameArray.Replace(".","_");
									oDataTableSchema=p_oAdo.getTableSchema(p_oConn,"SELECT * FROM " + p_strTableNameArray[y].Trim());
									strSourceColumnsList = p_oAdo.getFieldNames(p_oConn,"SELECT * FROM " + p_strTableNameArray[y].Trim());
									strSourceColumnsReservedWordFormattedList = p_oAdo.FormatReservedWordsInColumnNameList(strSourceColumnsList,",");
									strSourceColumnsArray = m_oUtils.ConvertListToArray(strSourceColumnsList,",");
									strDestColumnsList = p_oAdo.getFieldNames(p_oConn,"SELECT * FROM PRE_" + p_strTableIdArray[y].Trim());
									strDestColumnsArray = m_oUtils.ConvertListToArray(strDestColumnsList,",");

									p_oAdo.m_strSQL="";
									for (z=0;z<=oDataTableSchema.Rows.Count-1;z++)
									{
											
										if (oDataTableSchema.Rows[z]["ColumnName"] != System.DBNull.Value)
										{
											bFound=false;
											for (zz=0;zz<=strDestColumnsArray.Length - 1;zz++)
											{
												if (oDataTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() == 
													strDestColumnsArray[zz].Trim().ToUpper())
												{
													bFound=true;
													break;
												}
											}
											if (!bFound)
											{
												//column not found so let's add it
												p_oAdo.m_strSQL = p_oAdo.FormatCreateTableSqlFieldItem(oDataTableSchema.Rows[z]);
												p_oAdo.SqlNonQuery(p_oConn,"ALTER TABLE PRE_" + p_strTableIdArray[y].Trim() + " " + 
													"ADD COLUMN " + p_oAdo.m_strSQL);

												p_oAdo.SqlNonQuery(p_oConn,"ALTER TABLE POST_" + p_strTableIdArray[y].Trim() + " " + 
													"ADD COLUMN " + p_oAdo.m_strSQL);

											}
										}
											
									}
											

								}
								oDataTableSchema.Dispose();
								frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg,"Text",p_strVariant.Trim() + " " + p_strRx.Trim() + ": Get Pre And Post Treatment Years");
								frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Refresh");

								frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Value",3);

												
								//p_oAdo.m_OleDbDataReader.Close();
							    //p_oAdo.CloseConnection(p_oConn);
							    
								//p_oAdo.OpenConnection(p_oAdo.getMDBConnString(p_strWorkDbFile,"",""));


							    //load pre-post treatment years for a stand
							    if (p_oAdo.TableExist(p_oConn,"temp_stand_pre_post_rx"))
									    p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_stand_pre_post_rx");

							    p_oAdo.m_strSQL = "SELECT  MIN(a.year) AS post_rx_year, b.pre_rx_year,a.standid " + 
									              "INTO temp_stand_pre_post_rx " + 
									              "FROM " + p_strTableNameArray[y].Trim() + " a," + 
													"(SELECT MIN(year) AS pre_rx_year , standid  " + 
									                 "FROM " + p_strTableNameArray[y].Trim() + " " + 
									                 "WHERE standid IS NOT NULL GROUP BY standid) b " + 
												  "WHERE a.standid=b.standid AND " + 
									                    "a.year <> pre_rx_year " + 
									              "GROUP BY a.standid,b.pre_rx_year";

							    p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

							   string strFormattedSelectColumnList="";
							   for (z=0;z<=strSourceColumnsArray.Length-1;z++)
							   {
									strFormattedSelectColumnList = strFormattedSelectColumnList + "a." + strSourceColumnsArray[z].Trim() + ",";
							   }
							   
								strFormattedSelectColumnList=strFormattedSelectColumnList.Substring(0,strFormattedSelectColumnList.Length-1);
								
							    

								//
								//pre-treatment processing
								//
								if (p_bUpdatePreTableWithVariant==true)
								{
									frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text",p_strVariant.Trim() + " " + p_strRx.Trim() + ": Insert Pre-Treatment Values");
									frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Refresh");

									//delete any rows with the current variant
									p_oAdo.m_strSQL = "DELETE FROM PRE_" + p_strTableIdArray[y].Trim() + " " + 
										"WHERE FVS_VARIANT='" + p_strVariant.Trim() + "'";
									p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

                                    
									//insert any rows with the current variant
									p_oAdo.m_strSQL = "INSERT INTO PRE_" + p_strTableIdArray[y].Trim() + " " + 
										"(fvs_variant," + strSourceColumnsReservedWordFormattedList + ") " + 
										"SELECT '" + p_strVariant + "' AS fvs_variant," + strFormattedSelectColumnList + " " + 
										"FROM " + p_strTableNameArray[y] + " a,temp_stand_pre_post_rx b " + 
										"WHERE a.year IS NOT NULL AND " + 
											  "(a.standid=b.standid AND a.year = b.pre_rx_year)";

                                                    
									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt",p_oAdo.m_strSQL + "\r\n\r\n");
									p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

									//update biosum_cond_id column
									p_oAdo.m_strSQL = "UPDATE PRE_" + p_strTableIdArray[y].Trim() + " " + 
										"SET biosum_cond_id = IIF((biosum_cond_id IS NULL OR LEN(TRIM(biosum_cond_id))=0) AND (standid IS NOT NULL AND LEN(TRIM(standid)) = 25),MID(standid,1,25),'') " + 
										"WHERE FVS_VARIANT='" + p_strVariant.Trim() + "'";
														            

									p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
													
								}
								frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Value",4);
								//
								//post-treatment processing
								//
								//delete any rows with the current variant and treatment
								this.m_frmTherm.lblMsg.Text = p_strVariant.Trim() + " " + p_strRx.Trim() + ": Insert Post-Treatment Values";
								this.m_frmTherm.lblMsg.Refresh();
								p_oAdo.m_strSQL = "DELETE FROM POST_" + p_strTableIdArray[y].Trim() + " " + 
									"WHERE FVS_VARIANT='" + p_strVariant.Trim() + "' AND " + 
									"RX='" + p_strRx + "'";
								p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
								//insert new rows with the current variant and treatment
								p_oAdo.m_strSQL = "INSERT INTO POST_" + p_strTableIdArray[y].Trim() + " " + 
									"(fvs_variant,rx," + strSourceColumnsReservedWordFormattedList + ") " + 
									"SELECT '" + p_strVariant + "' AS fvs_variant," + 
									"'" + p_strRx + "' AS rx," + strFormattedSelectColumnList + " " + 
									"FROM " + p_strTableNameArray[y] + " a,temp_stand_pre_post_rx b " + 
									"WHERE a.year IS NOT NULL AND " + 
										  "(a.standid=b.standid AND a.year = b.post_rx_year)";

								p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
								//update biosum_cond_id column
								p_oAdo.m_strSQL = "UPDATE POST_" + p_strTableIdArray[y].Trim() + " " + 
									"SET biosum_cond_id = IIF((biosum_cond_id IS NULL OR LEN(TRIM(biosum_cond_id))=0) AND (standid IS NOT NULL AND LEN(TRIM(standid)) = 25),MID(standid,1,25),'') " + 
									"WHERE FVS_VARIANT='" + p_strVariant.Trim() + "' AND " + 
									"RX='" + p_strRx + "'";
														            

								p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
								frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Value",5);

							}
											
						
						//else p_oAdo.m_OleDbDataReader.Close();
					}
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Value",6);
//					if (bError) break;
				}
	
		}
		private void AppendAndUpdateRecords()
		{
			string strOutDirAndFile;
			string strVariant="";
			string strCurVariant="";
			string strRx="";
			string strFvsTreeFile;
			string strFvsTreeTable;
			string strLink="";
			bool bUpdateCondTable=false;
			int m_intCurrentCount=0;
			int m_intCheckedCount=0;
			string strFVSOutPath="";
			string[] strSourceTableArray=null;
			string[] strSourceTableNameArray=null;
			string[] strDestTableArray=null;
			ado_data_access oAdo = new ado_data_access();
			bool bGoodVariant=false;
			string strWorkDbFile="";
			string strWorkFvsOutProcessorInDbFile="";
			int intCount;
			
			frmMain.g_oDelegate.CurrentThreadProcessStarted=true;
			int x,y,z;

			m_intCheckedCount=0;
			m_intTotalSteps=50000;
			m_intCurrentStep=0;
			m_intCurrentCount=0;

			try
			{
				intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(lstFvsOutput,"Count",false);
				//inititalize the run status column for each row to blank in the list view
				for (x=0;x<=intCount-1;x++)
				{
					//alternate the row color in the list view
					this.m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn=true;
					this.m_oLvAlternateColors.DelegateListViewSubItem(lstFvsOutput.Items[x],COL_RUNSTATUS);

					//inititalize the run status column for each row to blank in the list view
					frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,uc_PostFvsForeFrcs.COL_RUNSTATUS,"Text","");

					//see if checked
					if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(lstFvsOutput,x,"CHECKED",false))
						m_intCheckedCount++;
				}

				m_intCurrentCount=0;

				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Maximum",m_intCheckedCount);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Minimum",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Value",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2,"Text","Overall Progress");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Visible",true);

				if (m_bDebug)
					this.WriteText("c:\\tmp\\biosum_debug.txt","check point 1 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");


				//create a link to each of the selected fvs out files in the temp mdb file
				//close the current ado oledb connection
				if (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
                    this.m_ado.CloseConnection(m_ado.m_OleDbConnection);

				if (m_bDebug)
					this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 2 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");


				//
				//create the biosum_fvsout_prepost_rx.mdb table if it does not exist
				//
				strFVSOutPath = frmMain.g_oFrmMain.frmProject.uc_project1.m_strDBProjectDirectory + "\\fvs\\db\\biosum_fvsout_prepost_rx.mdb";
				//get random file name that will be used as the work db file for biosum_fvsout_prepost_rx.mdb tables
				strWorkDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");

				if (m_bDebug)
					this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 3 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
				

				if (!System.IO.File.Exists(strFVSOutPath))
				{
					this.m_dao.CreateMDB(strFVSOutPath);
				}

				if (m_bDebug)
					this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 4 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Maximum",2);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Minimum",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","Backing up " + strFVSOutPath);
				//copy to work db file to temp folder
				System.IO.File.Copy(strFVSOutPath,strWorkDbFile,true);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Value",1);

				if (m_bDebug)
					this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 5 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

				//create a temporary link to the fvs_tree table
				strFvsTreeFile = m_DataSource.getFullPathAndFile("FVS TREE LIST FOR PROCESSOR");
				strFvsTreeTable = this.m_strFvsTreeTable;
				//get a random temporary file name that will be used as the work db file for the fvs out processor in mdb file.
				strWorkFvsOutProcessorInDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,".accdb");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","Backing up " + strFvsTreeFile);
				//copy the production file to the temp folder which will be used as the work db file.
				System.IO.File.Copy(strFvsTreeFile,strWorkFvsOutProcessorInDbFile,true);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Value",2);
				//create the link to the fvs_tree table in the work db file.
				m_dao.CreateTableLink(strWorkDbFile,strFvsTreeTable,strWorkFvsOutProcessorInDbFile,strFvsTreeTable,true);

                //alternate the row color in the list view
				
				for (x=0;x<=intCount-1;x++)
				{
					this.m_oLvAlternateColors.DelegateListViewSubItem(lstFvsOutput.Items[x],COL_RUNSTATUS);
					frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"Text","");
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Value",m_intCurrentCount);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Maximum",100);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Minimum",0);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Value",0);

					
					if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(this.lstFvsOutput,x,"Checked",false)==true)
					{
						m_intCurrentCount++;
						int intItemError=0;
						string strItemError="";
						int intItemWarning=0;
						string strItemWarning="";
						m_intCurrentStep=0;
						
						
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.lblRunStatus,"Text","Processing " + m_intCurrentCount.ToString().Trim() + " of " + m_intCheckedCount.ToString().Trim() + "     " + Convert.ToString((m_intCurrentCount * 100) /m_intCheckedCount).Trim() + "%");
						frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.lblRunStatus,"Refresh");


						//make sure the list view item is selected and visible to the user
						frmMain.g_oDelegate.ExecuteControlMethodWithParam((System.Windows.Forms.Control)lstFvsOutput,"EnsureVisible",new object[] {x});
						frmMain.g_oDelegate.SetListViewItemPropertyValue(lstFvsOutput,x,"Selected",true);
						frmMain.g_oDelegate.SetListViewItemPropertyValue(lstFvsOutput,x,"Focused",true);
						

						this.m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn=false;
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"BackColor",Color.DarkGoldenrod);
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"ForeColor",Color.White);
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing");

						//get the variant
						strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsOutput,x,COL_VARIANT,"Text",false);
						strVariant=strVariant.Trim();

						//get the treatment
						strRx = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RX,"Text",false);
						strRx=strRx.Trim();

						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 6 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
						//see if this is a different variant
						//only update the pre-treatment tables when the variant changes
						if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
						{
							bGoodVariant=false;
							bUpdateCondTable=true;
							strCurVariant = strVariant;
						}

						
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","Processing " + strVariant.Trim() + " " + strRx.Trim());
						frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Refresh");

						//
						//get the table names found in biosum_fvsout_prepost_rx.mdb
						//
						this.m_dao.getTableNames(strWorkDbFile,ref strDestTableArray);

						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","Create Table Links \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

						//get the fvs output file. 
						strOutDirAndFile = this.txtOutDir.Text.Trim()  + "\\" + 
							Convert.ToString(frmMain.g_oDelegate.GetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_MDBOUT,"Text",false)).Trim();
						this.m_strOutMDBFile = Convert.ToString(frmMain.g_oDelegate.GetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_MDBOUT,"Text",false)).Trim();

						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","strOutDirAndFile=" + strOutDirAndFile + "  \r\n START: " + System.DateTime.Now.ToString() + "\r\n");


						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","m_strOutMDBFile=" + m_strOutMDBFile + "  \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

						frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent,1,"Processing " + m_strOutMDBFile + "...Stand By");
						//
						//get the list of tables in the variant+rx fvsoutput db file
						//
						string[] strTempArray=null;
						string strTemp="";
						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 7 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

						this.m_dao.getTableNames(strOutDirAndFile,ref strTempArray);
						
						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 8 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

						//reload into the array only tables that will be appended to the project PRE/POST tables.
						for (y=0;y<=strTempArray.Length-1;y++)
						{
							if (strTempArray==null) break;
							if (strTempArray[y]!=null && strTempArray[y].Trim().Length > 3)
							{
								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt","strTempArray[" + y.ToString() + "]" + strTempArray[y] + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
								if (strTempArray[y].Substring(0,4)=="FVS_")
								{
									//not loading pre-treatment trees
									if (strTempArray[y].Trim().ToUpper() != "FVS_TREELIST"  &&
									    strTempArray[y].Trim().ToUpper() != "FVS_ATRTLIST")
											strTemp = strTemp + strTempArray[y] + ",";
								}
							}
						}
						strTemp = strTemp.Substring(0,strTemp.Length - 1);
						strSourceTableArray = frmMain.g_oUtils.ConvertListToArray(strTemp,",");
						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 9 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");



						
						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","strWorkDbFile=" + strWorkDbFile + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
						//
						//create links in the work file biosum_fvsout_prepost_rx.mdb to all the fvsoutput variant+rx tables
						//
						z=0;
						strSourceTableNameArray = new string[strSourceTableArray.Length];
						for (y=0;y<=strSourceTableArray.Length-1;y++)
						{
							if (strSourceTableArray[y] == null) break;
							if (strSourceTableArray[y].Trim().ToUpper() != "FVS_TREELIST" &&
								strSourceTableArray[y].Trim().ToUpper() != "FVS_ATRTLIST")
							{
								strLink = m_strOutMDBFile.Trim() + "_" + strSourceTableArray[y].Trim();
								strLink = strLink.Replace(".","_");
								strSourceTableNameArray[y]=strLink;

								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt","strSourceTableArray[" + y.ToString() + "]=" + strSourceTableArray[y] + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

								m_dao.CreateTableLink(strWorkDbFile,strLink,strOutDirAndFile,strSourceTableArray[y].Trim(),true);
								z++;
							}
							else
							{
								strSourceTableNameArray[y]=" ";
							}
						}
						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","checkpoint 10 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

						UpdateTherm(100);
						//
						//open ado connection
						//
						oAdo.OpenConnection(oAdo.getMDBConnString(strWorkDbFile,"",""));

						//
						//validation
						//
						Validation(oAdo,oAdo.m_OleDbConnection,strSourceTableArray,strSourceTableNameArray,x,strVariant,strRx,false,ref intItemError,ref strItemError,ref intItemWarning,ref strItemWarning);
						//
						//Table Structure Checks And Edits
						//
						if (intItemError==0)
						{
							PrePostTableStructureEdits(oAdo,oAdo.m_OleDbConnection,
														strWorkDbFile,
														strVariant,strRx,
														bUpdateCondTable,x,
														strSourceTableArray,strSourceTableNameArray,
														ref intItemError,ref strItemError);
						}
						UpdateTherm(100);
						//
						//update FVSTree table
						//
						if (intItemError==0)
						{
							UpdateFVSTreeTable(oAdo,oAdo.m_OleDbConnection,strVariant,strRx,
								               strFvsTreeTable,
								               strSourceTableArray,strSourceTableNameArray,
								               ref intItemError,ref strItemError);
						}
						UpdateTherm(100);
						//
						//clean up for this list item
						//
						if (intItemError==0)
						{
							//if (Convert.ToString(frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text",false)).Trim().ToUpper()=="PROCESSING")
							//{
							bGoodVariant=true;
							//}

						}
						else
						{
							//variant+rx processing
							//if an error for the variant+rx combination than delete
							//all records with this combination
							DeleteVariantRxFromPrePostTables(oAdo,oAdo.m_OleDbConnection,strSourceTableArray,strVariant,strRx);
							//variant only processing
							//look to see if the next item in the list is the current variant, if not,
							//check if any post-treatment records with the current variant, if not,
							//delete the variant from the pre-treatment tables
							if (!bGoodVariant) DeleteVariantFromPreTables(oAdo,oAdo.m_OleDbConnection,x,strSourceTableArray,strVariant);

						}
						//close the ado connection in order to use dao to delete table links
						oAdo.CloseConnection(oAdo.m_OleDbConnection);
						UpdateTherm(100);
						//delete links
						for (y=0;y<=strSourceTableArray.Length-1;y++)
						{
							if (strSourceTableArray[y] == null) break;
							if (strSourceTableArray[y].Trim().ToUpper() != "FVS_TREELIST" &&
								strSourceTableArray[y].Trim().ToUpper() != "FVS_ATRTLIST")
							{
								strLink = m_strOutMDBFile.Trim() + "_" + strSourceTableArray[y].Trim();
								strLink = strLink.Replace(".","_");
								m_dao.DeleteTableFromMDB(strWorkDbFile,strLink);
							}
							UpdateTherm(1);
						}
						m_dao.getTableNames(strWorkDbFile,ref strSourceTableArray);
						//delete all the temporary work tables for the list view item
						for (y=0;y<=strSourceTableArray.Length-1;y++)
						{
							if (strSourceTableArray[y] == null) break;
							if (strSourceTableArray[y].Trim().Length > 5 && 
								strSourceTableArray[y].Substring(0,4) != "PRE_" && 
								strSourceTableArray[y].Substring(0,5) != "POST_" && 
								strSourceTableArray[y].Trim().ToUpper() != this.m_strFvsTreeTable.Trim().ToUpper())
							{
								m_dao.DeleteTableFromMDB(strWorkDbFile,strSourceTableArray[y]);
							}
							UpdateTherm(1);
						}

						if (intItemError==0)
						{
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"BackColor",Color.DarkGreen);
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"Text","Completed");
						}
						else if (intItemError != 0)
						{
							m_intError=intItemError;
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"BackColor",Color.Red);
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","ERROR:" + strItemError);
							MessageBox.Show("ERROR:" + strItemError,"FIA Biosum");

						}
						UpdateTherm(this.m_intTotalSteps-1);
						System.Threading.Thread.Sleep(2000);
						
					}

				}

				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","Cleaning up");
				//delete the fvs_tree link from the work db file
				m_dao.DeleteTableFromMDB(strWorkDbFile,this.m_strFvsTreeTable);
				//copy the biosum_fvsout_prepost_rx.mdb temp work db file to the production path
				System.IO.File.Copy(strWorkDbFile,strFVSOutPath,true);
				//copy the fvs out processor in temp work db file to the production path
				System.IO.File.Copy(strWorkFvsOutProcessorInDbFile,strFvsTreeFile,true);
				this.FVSRecordsFinished();
			}
			catch (System.Threading.ThreadInterruptedException err)
			{

				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch  (System.Threading.ThreadAbortException err)
			{
				this.ThreadCleanUp();
				
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_PostFvsForeFrcs:AppendAndUpdateRecords  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"FVS Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent,1,"Ready");
			//ThreadCleanUp();

			CleanupThread();

			frmMain.g_oDelegate.m_oEventThreadStopped.Set();
			this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);

		}
		private void DeleteVariantRxFromPrePostTables(FIA_Biosum_Manager.ado_data_access p_oAdo,
														System.Data.OleDb.OleDbConnection p_oConn,
														string[] p_strTableArray,string p_strVariant, string p_strRx)
		{
			for (int y=0;y<=p_strTableArray.Length-1;y++)
			{
				if (p_strTableArray[y] == null) break;
				if (p_strTableArray[y].Trim().ToUpper() != "FVS_CASES" &&
					p_strTableArray[y].Trim().ToUpper() != "FVS_TREELIST" &&
					p_strTableArray[y].Trim().ToUpper() != "FVS_CUTLIST" &&
					p_strTableArray[y].Trim().ToUpper() != "FVS_ATRTLIST")
				{
					//delete any rows with with variant and rx
					if (p_oAdo.TableExist(p_oConn,"POST_" + p_strTableArray[y].Trim()))
					{
						p_oAdo.m_strSQL = "DELETE FROM POST_" + p_strTableArray[y].Trim() + " " + 
							"WHERE FVS_VARIANT='" + p_strVariant.Trim() + "' AND " + 
							"RX='" + p_strRx + "'";
						p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
					}
				}
			}
		}
		private void DeleteVariantFromPreTables(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,int p_intListViewItem,string[] p_strTableArray,string p_strVariant)
		{
		    int y,z;

			//check to see if the current variant has been saved
			bool bDeleteVariant=true;
			//check if the variant changes for the next item
			for (y=p_intListViewItem+1;y<=this.lstFvsOutput.Items.Count-1;y++)
			{

				if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(this.lstFvsOutput,y,"Checked",false)==true)
				{
					//check if the next item is the same variant
					if (Convert.ToString(frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsOutput,y,COL_VARIANT,"Text",false)).Trim().ToUpper() ==
						p_strVariant.Trim().ToUpper())
					{
						bDeleteVariant=false;
						break;
					}
					else break;
				}
			}
				//delete all records for this variant
			if (bDeleteVariant)
			{
				for (z=0;z<=p_strTableArray.Length-1;z++)
				{
					if (p_strTableArray[z] == null) break;
					if (p_strTableArray[z].Trim().ToUpper() != "FVS_CASES" &&
						p_strTableArray[z].Trim().ToUpper() != "FVS_TREELIST" &&
						p_strTableArray[z].Trim().ToUpper() != "FVS_CUTLIST" &&
						p_strTableArray[z].Trim().ToUpper() != "FVS_ATRTLIST")
					{
						//check to see if there are some valid post treatment records.
						//if valid post-treatment records exist for the current variant 
						//then do not delete the pretreatment records for the same variant.
						//post-treatment records are dependent on pre-treatment records
						if (p_oAdo.TableExist(p_oConn,"POST_" + p_strTableArray[z].Trim()))
						{
							if (p_oAdo.ValuesExistEqualToTargetValue(p_oConn,"POST_" + p_strTableArray[z].Trim(),"fvs_variant",p_strVariant.Trim(),false))
								bDeleteVariant=false;
						}
						if (bDeleteVariant)
						{
							if (p_oAdo.TableExist(p_oConn,"PRE_" + p_strTableArray[z].Trim()))
							{
								//check to see if there are other post treatment records

								//delete any rows with the current variant
								p_oAdo.m_strSQL = "DELETE FROM PRE_" + p_strTableArray[z].Trim() + " " + 
									"WHERE FVS_VARIANT='" + p_strVariant.Trim() + "'";
								p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
							}
						}
					}
										
				}
			}
			
		}

		private void UpdateFVSTreeTable(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,
			                            string p_strVariant,string p_strRx,
			                            string p_strFvsTreeTable,
			                            string[] p_strTableIdArray,string[] p_strTableNameArray,
										ref int p_intError,ref string p_strError)
		{
			
			    int x;
			    string strCutListTable="";
			    string strCasesTable="";

				//create a table from the summary audit table that lists the first year
			    //  a stand has trees cut. This will be used as the post-treatment year.
				if (p_oAdo.TableExist(p_oConn,"cutlist_post_rx_year"))
					  p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE cutlist_post_rx_year");

			     p_oAdo.SqlNonQuery(p_oConn,"SELECT MIN(year) AS post_rx_year,standid INTO cutlist_post_rx_year FROM " + this.m_strFVSSummaryAuditTable + " WHERE standid IS NOT NULL AND fvs_cutlist_count > 0 GROUP BY standid");

						
				/**************************************************************
				 **delete records in the fvs_tree table that have the current
				 **rx and variant
				 **************************************************************/
							
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg,"Text",p_strVariant.Trim() + " " + p_strRx.Trim() + ": Deleting Old Records In FVS Tree Table With Current Variant And Rx Values");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)m_frmTherm.lblMsg,"Refresh");
				p_oAdo.m_strSQL = "DELETE FROM " + p_strFvsTreeTable + " f " +
					"WHERE TRIM(UCASE(f.rx)) = '" + p_strRx.Trim().ToUpper() + "' AND " + 
					"TRIM(UCASE(f.fvs_variant)) = '" + p_strVariant.Trim().ToUpper() + "';";
				if (m_bDebug)
					this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

				p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);


				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Value",7);
				if (m_bDebug)
					this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

				p_intError=p_oAdo.m_intError;
				

				if (p_intError == 0)
				{
					for (x=0;x<=p_strTableIdArray.Length -1 ;x++)
					{
						if (p_strTableIdArray[x].Trim().ToUpper()=="FVS_CUTLIST")
							  strCutListTable = p_strTableNameArray[x].Trim();
						if (p_strTableIdArray[x].Trim().ToUpper()=="FVS_CASES")
							strCasesTable = p_strTableNameArray[x].Trim();
						
							  
					}

					

					this.m_frmTherm.lblMsg.Text = p_strVariant.Trim() + " " + p_strRx.Trim() + ": Adding New Records In FVS Tree Table With Current Variant And Rx Values";
					this.m_frmTherm.lblMsg.Refresh();
					//add the cut list tree records to the fvs out processer in tree table
						
					p_oAdo.m_strSQL = "INSERT INTO " + p_strFvsTreeTable + " " +
						"(biosum_cond_id, fvs_variant, fvs_tree_id, rx," + 
						"cut_leave, fvs_species, tpa, dbh, ht) " + 
						"SELECT c.StandID AS biosum_cond_id, c.Variant AS fvs_variant, IIf(Len(Trim(t.treeid))=4," + 
						"c.variant+'000'+Trim(t.treeid),IIf(Len(Trim(t.treeid))=5," + 
						"c.variant+'00'+Trim(t.treeid),IIf(Len(Trim(t.treeid))=6,"  +
						"c.variant+'0'+Trim(t.treeid),c.variant+Trim(t.treeid)))) AS fvs_tree_id," + 
						"'" + p_strRx.Trim().ToUpper() + "' AS rx, 'C' AS cut_leave," + 
						"t.Species AS fvs_species, t.TPA, t.DBH, t.Ht " + 
						"FROM " + strCasesTable + " c," +  strCutListTable + " t,cutlist_post_rx_year p " + 
						"WHERE c.CaseID = t.CaseID AND (t.standid=p.standid AND t.year=p.post_rx_year)";

					if (m_bDebug)
						this.WriteText("c:\\tmp\\biosum_debug.txt",p_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
					p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
					if (m_bDebug)
						this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

					p_intError=p_oAdo.m_intError;
								
				}
			    if (p_oAdo.TableExist(p_oConn,"cutlist_post_rx_year"))
				      p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE cutlist_post_rx_year");
			
		}
        
		private void AppendAndUpdateRecords_SAVE_20110912()
		{
			string strOutDirAndFile;
			string strVariant="";
			string strCurVariant="";
			string strRx="";
			//string strCopyFrom;
			//string strCopyTo;
			string strPotFireTable;
			string strCutListTable;
			string strCasesTable;
			//string strPreRxYear="";
			//string strPostRxYear="";
			string strLink="";
			bool bUpdateCondTable=false;
			bool bSkip;
			bool bBadVariant;
			DialogResult result;
			int intCurrentCount=0;
			int intSelectedCount=0;


			result = DialogResult.Yes;

			//int y;
			int x;
			//bool bResult;
			//bAbort=false;
			try
			{
				bSkip=false;
				
				this.ParentForm.Enabled=false;
				this.m_frmTherm.progressBar1.Maximum = this.lstFvsOutput.Items.Count;
				this.m_frmTherm.progressBar1.Minimum = 0;
				this.m_frmTherm.lblMsg.Text="";
				this.m_frmTherm.lblMsg.Visible=true;


				//create a link to each of the selected fvs out files in the temp mdb file
				//close the current ado oledb connection
				if (this.m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
					this.m_ado.m_OleDbConnection.Close();

				while (this.m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
				{
					System.Threading.Thread.Sleep(1000);
				}


				
				
				//dao_data_access p_dao=new dao_data_access();
				for (x=0;x<=this.lstFvsOutput.Items.Count-1;x++)
				{
					this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].BackColor = Color.White;
					this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].Text = "";
					if (this.lstFvsOutput.Items[x].Checked==true)
					{

						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","Create Table Links \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

						
						strOutDirAndFile = this.txtOutDir.Text.Trim()  + "\\" + this.lstFvsOutput.Items[x].SubItems[COL_MDBOUT].Text.Trim();
						this.m_strOutMDBFile = this.lstFvsOutput.Items[x].SubItems[COL_MDBOUT].Text.Trim();
						frmMain.g_sbpInfo.Text = "Creating Links To Tables In " + m_strOutMDBFile + "...Stand By";
						strLink =this.m_strOutMDBFile.Trim() + "_fvs_potfire";
						strLink = strLink.Replace(".","_");
						this.m_dao.CreateTableLink(this.m_strTempMDBFile,strLink,strOutDirAndFile,"fvs_potfire",true);
						strLink =this.m_strOutMDBFile.Trim() + "_fvs_cutlist";
						strLink = strLink.Replace(".","_");
						this.m_dao.CreateTableLink(this.m_strTempMDBFile,strLink,strOutDirAndFile,"fvs_cutlist",true);
						strLink =this.m_strOutMDBFile.Trim() + "_fvs_cases";
						strLink = strLink.Replace(".","_");
						this.m_dao.CreateTableLink(this.m_strTempMDBFile,strLink,strOutDirAndFile,"fvs_cases",true);
						if (m_bDebug)
							this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

						if (!this.m_dao.TableExists(this.m_strTempMDBFile,"fvs_potfire_work_table"))
						{
							if (m_bDebug)
								this.WriteText("c:\\tmp\\biosum_debug.txt","Get fvs_potfire_work_table Table Structure \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

							ado_data_access p_ado = new ado_data_access();

							p_ado.OpenConnection(this.m_strConn);

							/*********************************************
							 **get fields from the potential fire table
							 *********************************************/
							strPotFireTable = this.m_strOutMDBFile.Trim() + "_fvs_potfire";
							strPotFireTable = strPotFireTable.Replace(".","_");
							p_ado.m_strSQL = "SELECT standid AS biosum_cond_id,year AS pre_treatment_year,torch_index,crown_index," + 
								"tot_flame_sev,tot_flame_mod," + 
								"fire_type_sev,fire_type_mod," +
								"canopy_ht,canopy_density," +
								"mortality_ba_sev,mortality_ba_mod," +
								"mortality_vol_sev,mortality_vol_mod FROM " + strPotFireTable;

							/****************************************************************
							 **get the table structure that results from executing the sql
							 ****************************************************************/
							System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,p_ado.m_strSQL);
							this.m_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"fvs_pre_potfire_work_table",p_dt,true);

							p_ado.m_strSQL = "SELECT standid AS biosum_cond_id,year AS post_treatment_year,torch_index,crown_index," + 
								"tot_flame_sev,tot_flame_mod," + 
								"fire_type_sev,fire_type_mod," +
								"canopy_ht,canopy_density," +
								"mortality_ba_sev,mortality_ba_mod," +
								"mortality_vol_sev,mortality_vol_mod FROM " + strPotFireTable;
							p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,p_ado.m_strSQL);
							this.m_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"fvs_post_potfire_work_table",p_dt,true);
							this.m_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"fvs_post_potfire_work_table2",p_dt,true);
							p_dt.Dispose();
							p_ado.m_OleDbDataReader.Close();
							p_ado.m_OleDbConnection.Close();

							//close the ado connection
							p_ado.m_OleDbConnection.Close();

							while (p_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
							{
							}
							if (m_bDebug)
								this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
						}


					}

				}
				frmMain.g_sbpInfo.Text = "Ready";
				

				//p_dao=null;
				//reopen the ado connection containing the new table links that have just been added
				this.m_ado.OpenConnection(this.m_strConn);
                



				//prompt user to delete any existing records
				string strMsg="";
				int intCount=0;
				this.WriteText("c:\\tmp\\biosum_debug.txt","\r\n" + "Get Variants With Records" + "\r\n " + "START:" + System.DateTime.Now.ToString());

				strMsg = "Appending the variants and treatments you selected will overwrite \r\n " + 
					"the same variants and treatments that exist from a previous append. \r\n " + 
					"Do you wish to check which variants and treatments will be overwritten? (Y/N)\r\n";
				result = MessageBox.Show(strMsg,"FIA Biosum", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
				strMsg="";
				if (result==DialogResult.Yes)
				{
					FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
					frmTemp.Text = "FIA Biosum";
					FIA_Biosum_Manager.TemplateGroupBox oGroupBox = new TemplateGroupBox(frmTemp);
					FIA_Biosum_Manager.TemplateTitle oTitle = new TemplateTitle(oGroupBox,0,0,"Existing Variant + Rx List");
					FIA_Biosum_Manager.TemplateListBox oListBox = new TemplateListBox(oGroupBox,"listbox");
					FIA_Biosum_Manager.TemplateInputLabel oInputLabel = new TemplateInputLabel(oGroupBox,"inputlabel",
						"The listed variant and treatment record(s) exist in the ffe and fvs_tree table. " + 
						"Do you want to delete them and append new records? (Y/N)");
					oInputLabel.RightToLeft=System.Windows.Forms.RightToLeft.No;
					FIA_Biosum_Manager.TemplateOkCancelButtons oButtons = new TemplateOkCancelButtons(frmTemp,oGroupBox);
					for (x=0;x<=this.lstFvsOutput.Items.Count-1;x++)
					{
						if (this.lstFvsOutput.Items[x].Checked==true)
						{
							//get the variant
							strVariant = this.lstFvsOutput.Items[x].SubItems[COL_VARIANT].Text.Trim();

							//only update the condition table when the variant changes
							if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
							{
								bUpdateCondTable=true;
								strCurVariant = strVariant;
							}
							//get the treatment
							strRx = this.lstFvsOutput.Items[x].SubItems[COL_RX].Text.Trim();

							frmMain.g_sbpInfo.Text = "Checking if variant " + strVariant + " and treatment " + strRx + " exist in FFE table";

							if (m_ado.TableExist(m_ado.m_OleDbConnection,"plot_cond_count_work_table"))
								m_ado.SqlNonQuery(m_ado.m_OleDbConnection,"DROP TABLE plot_cond_count_work_table");


							m_ado.m_strSQL = "SELECT c.biosum_plot_id,c.biosum_cond_id " + 
								"INTO plot_cond_count_work_table " + 
								"FROM " + this.m_strPlotTable + " p, " +
								this.m_strCondTable + " c " + 
								"WHERE c.biosum_plot_id = p.biosum_plot_id AND " + 
								"TRIM(p.fvs_variant) = '" + strVariant.Trim() + "'";

							m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);

							if (m_ado.TableExist(m_ado.m_OleDbConnection,"ffe_count_work_table"))
								m_ado.SqlNonQuery(m_ado.m_OleDbConnection,"DROP TABLE ffe_count_work_table");
							m_ado.m_strSQL = "SELECT f.biosum_cond_id,f.rx " + 
								"INTO ffe_count_work_table " + 
								"FROM " + this.m_strFFETable + " f " + 
								"WHERE TRIM(f.rx) = '" + strRx.Trim() + "'";
							m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);

						

							m_ado.m_strSQL="SELECT COUNT(*) " + 
								"FROM (SELECT TOP 1 f.biosum_cond_id FROM ffe_count_work_table f, " +
								"plot_cond_count_work_table c " + 
								"WHERE f.biosum_cond_id=c.biosum_cond_id)";
						


							intCount= (int)this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL,this.m_strFFETable);

							if (intCount > 0)
							{
								oListBox.Items.Add("Variant " + strVariant + " Treatment " + strRx);
								strMsg += " Variant " + strVariant + " Treatment " + strRx + "\n";
							}
						}
					}	
					if (strMsg.Trim().Length > 0)
					{
						
						oButtons.btnOK.Top = frmTemp.ClientSize.Height - oButtons.btnOK.Height - 5;
						oButtons.btnCancel.Top = oButtons.btnOK.Top;
						oButtons.btnOK.Left = (int)(frmTemp.ClientSize.Width / 2) - (int)(oButtons.btnOK.Width - 5);
						oButtons.btnCancel.Left = (int)(frmTemp.ClientSize.Width / 2) + 5;
						oInputLabel.Width =oGroupBox.Width - 50;
						oInputLabel.Top = oButtons.btnOK.Top - oInputLabel.Height - 10;
						oInputLabel.Left = (int)(frmTemp.ClientSize.Width / 2) - (int)(oInputLabel.Width * .5);
						oListBox.Top = oTitle.Top + oTitle.Height + 10;
						oListBox.Left = oTitle.Left;
						oListBox.Width = oGroupBox.Width - (oListBox.Left*2);
						oListBox.Height = oGroupBox.Height - oListBox.Top - oButtons.btnOK.Height - oInputLabel.Height - 30;
						oButtons.btnOK.Text = "Yes";
						oButtons.btnCancel.Text = "No";
						
						frmMain.g_sbpInfo.Text = "Ready";

						result = frmTemp.ShowDialog();
						if (result==DialogResult.OK) result=DialogResult.Yes;
					
					}
					if (m_bDebug)
						this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
				}
				else if (result==DialogResult.No)
				{
					result = DialogResult.Yes;
				}
				if (result==DialogResult.Yes)
				{
					intSelectedCount=lstFvsOutput.CheckedItems.Count;
					intCurrentCount=0;
					this.lblRunStatus.Text ="";
					this.lblRunStatus.Show();

					/***************************************************************
					 **fvs outputs biosum_cond_id (standid) as 255 characters so 
					 **lets change it to standard length 25 and create an index
					 ***************************************************************/
					this.m_ado.m_strSQL = "ALTER TABLE fvs_pre_potfire_work_table ALTER COLUMN biosum_cond_id CHAR(25)";
					this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
					m_ado.m_strSQL = "CREATE INDEX biosumcondididx ON fvs_pre_potfire_work_table (biosum_cond_id)";
					this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
					this.m_ado.m_strSQL = "ALTER TABLE fvs_post_potfire_work_table ALTER COLUMN biosum_cond_id CHAR(25)";
					this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
					m_ado.m_strSQL = "CREATE INDEX biosumcondididx2 ON fvs_post_potfire_work_table (biosum_cond_id)";
					this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
					this.m_ado.m_strSQL = "ALTER TABLE fvs_post_potfire_work_table2 ALTER COLUMN biosum_cond_id CHAR(25)";
					this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
					m_ado.m_strSQL = "CREATE INDEX biosumcondididx3 ON fvs_post_potfire_work_table2 (biosum_cond_id)";
					this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);




					for (x=0;x<=this.lstFvsOutput.Items.Count-1;x++)
					{
						

						this.m_frmTherm.progressBar1.Value = x;
						this.m_frmTherm.progressBar2.Maximum = 6;
						this.m_frmTherm.progressBar2.Minimum = 0;
						this.m_frmTherm.progressBar2.Value = 0;
					
				

						if (this.lstFvsOutput.Items[x].Checked==true)
						{
							intCurrentCount++;
							this.lblRunStatus.Text = "Processing " + intCurrentCount.ToString().Trim() + " of " + intSelectedCount.ToString().Trim() + "     " + Convert.ToString((intCurrentCount * 100) /intSelectedCount).Trim() + "%";
							this.lblRunStatus.Refresh();

							this.lstFvsOutput.EnsureVisible(x);
							this.lstFvsOutput.Items[x].Selected=true;
							this.lstFvsOutput.Items[x].Focused=true;
							this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].ForeColor=Color.White;
							this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].BackColor=Color.DarkGoldenrod;
							this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].Text = "Processing";
							

							bSkip=false;
							this.m_strOutMDBFile = this.lstFvsOutput.Items[x].SubItems[COL_MDBOUT].Text.Trim();
							strPotFireTable = this.m_strOutMDBFile.Trim() + "_fvs_potfire";
							strPotFireTable = strPotFireTable.Replace(".","_");
							strCutListTable = this.m_strOutMDBFile.Trim() + "_fvs_cutlist";
							strCutListTable = strCutListTable.Replace(".","_");
							strCasesTable = this.m_strOutMDBFile.Trim() + "_fvs_cases";
							strCasesTable = strCasesTable.Replace(".","_");

							//get the variant
							strVariant = this.lstFvsOutput.Items[x].SubItems[COL_VARIANT].Text.Trim();

							//only update the condition table when the variant changes
							if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
							{
								bUpdateCondTable=true;
								strCurVariant = strVariant;
							}
							//get the treatment
							strRx = this.lstFvsOutput.Items[x].SubItems[COL_RX].Text.Trim();

							bBadVariant=false;

							//check to ensure the variant in the fvs cases table
							//matches the current variant
							if (this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,
								"SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + strCasesTable + " WHERE TRIM(variant) <> '" + strVariant + "')",
								strCasesTable) > 0)
							{
								bBadVariant=true;
								MessageBox.Show("Error: The " + this.m_strOutMDBFile + " contains a variant other than " + strVariant,"FIA Biosum");
								this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].BackColor=Color.Red;
								this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].Text = "Error:Incorrect Variant in FVS_Cases table";
							}
							if (bBadVariant==false)
							{

								//check if any records in the potential fire table
								if (this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,
									"SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + strPotFireTable + ")",
									strPotFireTable) > 0)
								{
									if (bUpdateCondTable==true)
									{
										this.m_frmTherm.lblMsg.Text = strVariant.Trim() + " " + strRx.Trim() + ": Update Condition Table With New Pre-Treatment Fire Potential Values";
										this.m_frmTherm.lblMsg.Refresh();

										//delete all records from the work table
										if (this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,
											"select COUNT(*) from fvs_pre_potfire_work_table",
											"fvs_pre_potfire_work_table") > 0)
										{
											this.m_ado.m_strSQL = "DELETE FROM fvs_pre_potfire_work_table;";
											this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);

										}

										//load the pretreatment fire potential values into the work table
									
										this.m_ado.m_strSQL = "INSERT INTO fvs_pre_potfire_work_table " + 
											"SELECT a.biosum_cond_id,b.pre_treatment_year," +
											"c.torch_index,c.crown_index," +
											"c.tot_flame_sev,c.tot_flame_mod," + 
											"c.fire_type_sev,c.fire_type_mod," + 
											"c.canopy_ht,c.canopy_density," + 
											"c.mortality_ba_sev,c.mortality_ba_mod," +
											"c.mortality_vol_sev,c.mortality_vol_mod " + 
											"FROM " + this.m_strCondTable.Trim() + " a, " + 
											"(SELECT standid,MIN(year) AS pre_treatment_year " + 
											"FROM " + strPotFireTable.Trim() + " " + 
											"GROUP BY standid) b, " + 
											"(SELECT " + strPotFireTable.Trim() + ".* FROM " + strPotFireTable.Trim() + " ) c " + 
											"WHERE (TRIM(b.standid) = TRIM(a.biosum_cond_id)) AND " + 
											"(TRIM(c.standid) = TRIM(b.standid) AND " + 
											"c.year = b.pre_treatment_year);";
									
										
										if (m_bDebug)
											this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

										this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
										if (m_bDebug)
											this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


										
										this.m_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
											"INNER JOIN fvs_pre_potfire_work_table f " + 
											"ON c.biosum_cond_id = f.biosum_cond_id " + 
											"SET c.pre_not_calc_yn = IIF(f.torch_index < 0,'Y','N')," + 
											"c.pre_tot_flame_sev = f.tot_flame_sev," + 
											"c.pre_tot_flame_mod = f.tot_flame_mod," + 
											"c.pre_fire_type_sev = f.fire_type_sev," + 
											"c.pre_fire_type_mod = f.fire_type_mod," + 
											"c.pre_torch_index = f.torch_index," + 
											"c.pre_crown_index = f.crown_index," + 
											"c.pre_canopy_ht = f.canopy_ht," + 
											"c.pre_canopy_density = f.canopy_density," + 
											"c.pre_mortality_ba_sev = f.mortality_ba_sev," + 
											"c.pre_mortality_ba_mod = f.mortality_ba_mod," + 
											"c.pre_mortality_vol_sev = f.mortality_vol_sev," + 
											"c.pre_mortality_vol_mod = f.mortality_vol_mod;";

										if (m_bDebug)
											this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

										this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
										bUpdateCondTable=false;
										if (m_bDebug)
											this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

									}
								}
								else
								{
									this.m_ado.m_OleDbDataReader.Close();
									MessageBox.Show("!!0 FVS Potential Fire Table Records In The MDB File " + this.m_strOutMDBFile,
										"FIA Biosum",
										System.Windows.Forms.MessageBoxButtons.OK,
										System.Windows.Forms.MessageBoxIcon.Exclamation);
									bSkip=true;
								}
								
								if (this.m_ado.m_intError != 0 || this.m_intError != 0)
								{
									if (this.m_ado.m_intError!=0)
										this.m_intError = this.m_ado.m_intError;
									break;
								}
							
								this.m_frmTherm.progressBar2.Value = 1;

                        
								/**********************************************
									 **delete records in the ffe table whose
									 **plots have the same variant and rx
									 **as the currently processing variant and rx
									 **********************************************/
								this.m_frmTherm.lblMsg.Text = strVariant.Trim() + " " + strRx.Trim() + ": Deleting Old FFE Table Records With Current Variant And Rx Values";
								this.m_frmTherm.lblMsg.Refresh();
								if (m_ado.TableExist(m_ado.m_OleDbConnection,"ffe_temp_work_table"))
									m_ado.SqlNonQuery(m_ado.m_OleDbConnection,"DROP TABLE ffe_temp_work_table");

								if (m_ado.TableExist(m_ado.m_OleDbConnection,"cond_plot_temp_work_table"))
									m_ado.SqlNonQuery(m_ado.m_OleDbConnection,"DROP TABLE cond_plot_temp_work_table");

								if (m_ado.TableExist(m_ado.m_OleDbConnection,"FFE_rows_to_delete_work_table"))
									m_ado.SqlNonQuery(m_ado.m_OleDbConnection,"DROP TABLE FFE_rows_to_delete_work_table");

								//create a temp table containin ffe biosum cond id with the current treatment id
								this.m_ado.m_strSQL = "SELECT DISTINCT biosum_cond_id,rx " + 
									"INTO ffe_temp_work_table " + 
									"FROM " + this.m_strFFETable + " f " + 
									"WHERE rx ='" + strRx.Trim() + "'";
								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
								this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


								//create a temp table with cond and biosum cond id and fvs variant 
								this.m_ado.m_strSQL = "SELECT DISTINCT c.biosum_cond_id,p.fvs_variant " + 
									"INTO cond_plot_temp_work_table " + 
									"FROM " + this.m_strCondTable + " c," + 
									this.m_strPlotTable + " p " + 
									"WHERE c.biosum_plot_id = p.biosum_plot_id AND " + 
									"TRIM(p.fvs_variant)='" + strVariant.Trim() + "'";
								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

								this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

								                         
								                      
								

								this.m_ado.m_strSQL = "SELECT DISTINCT c.biosum_cond_id,f.rx " + 
									"INTO FFE_rows_to_delete_work_table " + 
									"FROM  ffe_temp_work_table f," +
									"cond_plot_temp_work_table c " + 
									"WHERE f.biosum_cond_id = c.biosum_cond_id";

								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
								this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

								m_ado.m_strSQL = "CREATE INDEX tempidx ON FFE_rows_to_delete_work_table (biosum_cond_id, rx)";
								this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);

								
								if (this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,
									"select COUNT(*) from FFE_rows_to_delete_work_table",
									"FFE_rows_to_delete_work_table") > 0)
								{

									this.m_ado.m_strSQL = "DELETE FROM " + this.m_strFFETable + " f " +
										"WHERE EXISTS (SELECT c.biosum_cond_id,c.rx " + 
										"FROM FFE_rows_to_delete_work_table c " +
										"WHERE f.biosum_cond_id=c.biosum_cond_id AND " + 
										"f.rx = c.rx)";

									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
									this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

									if (this.m_ado.m_intError !=0)
									{
										this.m_intError = this.m_ado.m_intError;
										break;
									}


								}



								
						
								this.m_frmTherm.progressBar2.Value = 2;


								/**************************************************************
									 **delete records in the fvs_tree table that have the current
									 **rx and variant
									 **************************************************************/
								this.m_frmTherm.lblMsg.Text = strVariant.Trim() + " " + strRx.Trim() + ": Deleting Old Records In FVS Tree Table With Current Variant And Rx Values";
								this.m_frmTherm.lblMsg.Refresh();
								this.m_ado.m_strSQL = "DELETE FROM " + this.m_strFvsTreeTable + " f " +
									"WHERE TRIM(UCASE(f.rx)) = '" + strRx.Trim().ToUpper() + "' AND " + 
									"TRIM(UCASE(f.fvs_variant)) = '" + strVariant.Trim().ToUpper() + "';";
								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

								this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
								if (m_bDebug)
									this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

								if (this.m_ado.m_intError !=0)
								{
									this.m_intError = this.m_ado.m_intError;
									break;
								}

								this.m_frmTherm.progressBar2.Value = 3;

						

								if (bSkip==false)
								{
									this.m_frmTherm.lblMsg.Text = strVariant.Trim() + " " + strRx.Trim() + ": Adding New FFE Table Records With Current Variant And Rx Values";
									this.m_frmTherm.lblMsg.Refresh();

									//delete all records from the work table
									if (this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,
										"select COUNT(*) from fvs_post_potfire_work_table",
										"fvs_post_potfire_work_table") > 0)
									{
										this.m_ado.m_strSQL = "DELETE FROM fvs_post_potfire_work_table;";
										this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);

									}
									if (this.m_ado.getRecordCount(this.m_ado.m_OleDbConnection,
										"select COUNT(*) from fvs_post_potfire_work_table2",
										"fvs_post_potfire_work_table2") > 0)
									{
										this.m_ado.m_strSQL = "DELETE FROM fvs_post_potfire_work_table2;";
										this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);

									}
									//append every treatment for the stand except the first one (pre-treatment)
									this.m_ado.m_strSQL = "INSERT INTO fvs_post_potfire_work_table " + 
										"SELECT DISTINCT a.standid AS biosum_cond_id,a.year AS post_treatment_year," + 
										"a.torch_index,a.crown_index," + 
										"a.tot_flame_sev,a.tot_flame_mod," + 
										"a.fire_type_sev,a.fire_type_mod," + 
										"a.canopy_ht,a.canopy_density," + 
										"a.mortality_ba_sev,a.mortality_ba_mod," + 
										"a.mortality_vol_sev,a.mortality_vol_mod " + 
										"FROM " + strPotFireTable.Trim() + " a " + 
										"WHERE NOT EXISTS (SELECT biosum_cond_id,pre_treatment_year FROM  fvs_pre_potfire_work_table b " + 
										"WHERE TRIM(a.standid)=b.biosum_cond_id AND a.year=b.pre_treatment_year);";

									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
									this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");



									//get the post treatment record for each stand
									this.m_ado.m_strSQL = "INSERT INTO fvs_post_potfire_work_table2 " + 
										"SELECT a.biosum_cond_id, a.post_treatment_year," +
										"a.torch_index,a.crown_index," +
										"a.tot_flame_sev,a.tot_flame_mod," + 
										"a.fire_type_sev,a.fire_type_mod," + 
										"a.canopy_ht,a.canopy_density," + 
										"a.mortality_ba_sev,a.mortality_ba_mod," +
										"a.mortality_vol_sev,a.mortality_vol_mod " + 
										"FROM fvs_post_potfire_work_table a, " + 
										"(SELECT biosum_cond_id,MIN(post_treatment_year) AS min_year " + 
										"FROM fvs_post_potfire_work_table " + 
										"GROUP BY biosum_cond_id) b " + 
										"WHERE b.biosum_cond_id = a.biosum_cond_id AND " + 
										"b.min_year = a.post_treatment_year;";
									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
									this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


									//insert post treatment fire potential values into the ffe table
									this.m_ado.m_strSQL = "INSERT INTO " + this.m_strFFETable + " " + 
										"(biosum_cond_id,rx,post_not_calc_yn," + 
										"post_tot_flame_sev,post_tot_flame_mod," + 
										"post_fire_type_sev,post_fire_type_mod," + 
										"post_torch_index,post_crown_index," + 
										"post_canopy_ht,post_canopy_density," + 
										"post_mortality_ba_sev,post_mortality_ba_mod," + 
										"post_mortality_vol_sev,post_mortality_vol_mod) " + 
										"SELECT biosum_cond_id,'" + strRx.Trim() + "' AS rx," + 
										"IIF(torch_index < 0,'Y','N') AS post_not_calc_yn," + 
										"tot_flame_sev AS post_tot_flame_sev," + 
										"tot_flame_mod AS post_fot_flame_mod," + 
										"fire_type_sev AS post_fire_type_sev," + 
										"fire_type_mod AS post_fire_type_mod," + 
										"torch_index AS post_torch_index," + 
										"crown_index AS post_crown_index," + 
										"canopy_ht AS post_canopy_ht," + 
										"canopy_density AS post_canopy_density," + 
										"mortality_ba_sev AS post_mortality_ba_sev," + 
										"mortality_ba_mod AS post_mortality_ba_mod," + 
										"mortality_vol_sev AS post_mortality_vol_sev," + 
										"mortality_vol_mod AS post_mortality_vol_mod " +
										"FROM fvs_post_potfire_work_table2;";

									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
									this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
									if (this.m_ado.m_intError !=0)
									{
										this.m_intError = this.m_ado.m_intError;
										break;
									}
						
                        
						
									this.m_frmTherm.progressBar2.Value = 4;
									this.m_frmTherm.lblMsg.Text = strVariant.Trim() + " " + strRx.Trim() + ": Adding New Records In FVS Tree Table With Current Variant And Rx Values";
									this.m_frmTherm.lblMsg.Refresh();
									//add the cut list tree records to the fvs out processer in tree table
						
									this.m_ado.m_strSQL = "INSERT INTO " + this.m_strFvsTreeTable + " " +
										"(biosum_cond_id, fvs_variant, fvs_tree_id, rx," + 
										"cut_leave, fvs_species, tpa, dbh, ht) " + 
										"SELECT c.StandID AS biosum_cond_id, c.Variant AS fvs_variant, IIf(Len(Trim(t.treeid))=4," + 
										"c.variant+'000'+Trim(t.treeid),IIf(Len(Trim(t.treeid))=5," + 
										"c.variant+'00'+Trim(t.treeid),IIf(Len(Trim(t.treeid))=6,"  +
										"c.variant+'0'+Trim(t.treeid),c.variant+Trim(t.treeid)))) AS fvs_tree_id," + 
										"'" + strRx.Trim().ToUpper() + "' AS rx, 'C' AS cut_leave," + 
										"t.Species AS fvs_species, t.TPA, t.DBH, t.Ht " + 
										"FROM " + strCasesTable + " c " + 
										"INNER JOIN " + strCutListTable + " t " + 
										"ON c.CaseID = t.CaseID;";

									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt",m_ado.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
									this.m_ado.SqlNonQuery(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
									if (m_bDebug)
										this.WriteText("c:\\tmp\\biosum_debug.txt","DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

									if (this.m_ado.m_intError !=0)
									{
										this.m_intError = this.m_ado.m_intError;
										break;
									}
									this.m_frmTherm.progressBar2.Value = 5;
								}
								this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].BackColor=Color.DarkGreen;
								this.lstFvsOutput.Items[x].SubItems[uc_PostFvsForeFrcs.COL_RUNSTATUS].Text = "Completed";
							}
						}
						this.m_frmTherm.progressBar2.Value = this.m_frmTherm.progressBar2.Maximum;
						
					}
					this.lblRunStatus.Text = "Completed " + intCurrentCount.ToString().Trim() + " of " + intSelectedCount.ToString().Trim() + "     " + Convert.ToString((intCurrentCount * 100) /intSelectedCount).Trim() + "%";
					this.m_frmTherm.progressBar1.Value=this.m_frmTherm.progressBar1.Maximum;
					
				}
				this.FVSRecordsFinished();
			}
			catch (System.Threading.ThreadInterruptedException err)
			{

				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch  (System.Threading.ThreadAbortException err)
			{
				this.ThreadCleanUp();
				
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_PostFvsForeFrcs:AppendAndUpdateRecords  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"FVS Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			ThreadCleanUp();

		}


		public void StopThread()
		{
			
			string strMsg="";
			
			if (m_frmTherm.Text.Trim() == "FVS Output")
			   strMsg = "Do you wish to cancel appending and updating fvs out data (Y/N)?";
			else
			   strMsg = "Do you wish to cancel audit (Y/N)?";

			frmMain.g_oDelegate.AbortProcessing("FIA Biosum",strMsg);

			if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
			{
				this.m_frmTherm.AbortProcess = true;
				frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,this.lstFvsOutput.SelectedItems[0].Index,COL_RUNSTATUS,"BackColor",Color.Red);
				frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,this.lstFvsOutput.SelectedItems[0].Index,COL_RUNSTATUS,"Text","Cancelled");

				frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent,0,"Ready");
				this.ThreadCleanUp();
			}

			
			
		}
		public void FVSRecordsFinished()
		{

			
			if (this.m_frmTherm != null)
			{
				this.m_frmTherm.Close();
				this.m_frmTherm.Dispose();
				this.m_frmTherm = null;
			}
			

		}
		private void ThreadCleanUp()
		{
			try
			{
				this.ParentForm.Enabled=true;
				this.btnRefresh.Enabled=true;
				this.btnAppend.Enabled=true;
				this.btnChkAll.Enabled=true;
				this.btnClearAll.Enabled=true;
				this.btnClose.Enabled=true;
				this.btnViewLogFile.Enabled=true;
				this.btnAuditDb.Enabled=true;
				this.btnHelp.Enabled=true;
				this.btnAudit.Enabled=true;
				this.grpboxAppend.Enabled=true;
				this.grpboxAudit.Enabled=true;
				
				this.Enabled=true;
				if (this.m_frmTherm != null)
				{
					this.m_frmTherm.Close();
					this.m_frmTherm.Dispose();
					this.m_frmTherm = null;
				}
				this.m_thread = null;
			}
			catch
			{
			}

		}
		private void WriteText(string p_strTextFile,string p_strText)
		{
			System.IO.FileStream oTextFileStream;
			System.IO.StreamWriter oTextStreamWriter;

			if (!System.IO.File.Exists(p_strTextFile))
			{
				oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Create, 
					System.IO.FileAccess.Write);
			}
			else
			{
				oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Append, 
					System.IO.FileAccess.Write);
			}
			
			oTextStreamWriter = new System.IO.StreamWriter(oTextFileStream);
			oTextStreamWriter.Write(p_strText);
			oTextStreamWriter.Close();
			oTextFileStream.Close();
		}

		private void btnAudit_Click(object sender, System.EventArgs e)
		{
			if (this.lstFvsOutput.CheckedItems.Count ==0)
			{
				MessageBox.Show("No Boxes Are Checked","FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			this.DisplayAuditMessage=true;
			this.val_data();
			if (this.m_intError==0)
			{
				this.m_frmTherm = new frmTherm(((frmDialog)ParentForm),"FVS OUT DATA",
					"FVS Output Audit","2");
				this.m_frmTherm.lblMsg.Text="";
				
				this.grpboxAppend.Enabled=false;
				this.grpboxAudit.Enabled=false;
				this.btnChkAll.Enabled=false;
				this.btnClearAll.Enabled=false;
				this.btnRefresh.Enabled=false;
				this.btnClose.Enabled=false;
				this.btnHelp.Enabled=false;
				this.btnViewLogFile.Enabled=false;
				this.btnAuditDb.Enabled=false;



				frmMain.g_oDelegate.CurrentThreadProcessAborted=false;
				frmMain.g_oDelegate.CurrentThreadProcessDone=false;
				frmMain.g_oDelegate.CurrentThreadProcessStarted=false;
				frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(Audit));
				frmMain.g_oDelegate.CurrentThreadProcessIdle=false;
				frmMain.g_oDelegate.InitializeThreadEvents();
				frmMain.g_oDelegate.m_oThread.IsBackground=true;
				
				frmMain.g_oDelegate.m_oThread.Start();
				
				
			}
		}
		
		public void Audit()
		{
			frmMain.g_oDelegate.CurrentThreadProcessStarted=true;
			this.m_intError=0;
			int intCount=0;
			m_intCheckedCount=0;
			m_intTotalSteps=50000;
			m_intCurrentStep=0;
			m_strError="";
			m_strWarning="";
			m_intWarning=0;
			m_intCurrentCount=0;
			

			
			if (DisplayAuditMessage)
			{
				this.m_strError="Audit Results \r\n";
				this.m_strError=m_strError + "-------------\r\n\r\n";
			}

			string strOutDirAndFile;
			string strDbFile;
			string strVariant="";
			string strRx="";
			string strCasesTable;
			

			int intCurrentCount=0;
			string[] strSourceTableArray=null;
			ado_data_access oAdo = new ado_data_access();
			
			int intPreRxYear;
			int intPostRxYear;
			bool bSkip;
			bool bResult=false;
			
			
			System.DateTime oDate = System.DateTime.Now;
			string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
			m_strLogDate = oDate.ToString(strDateFormat);
			m_strLogDate = m_strLogDate.Replace("/","_"); m_strLogDate=m_strLogDate.Replace(":","_");
			
			

			
			int x,y,z,zz;
			
			try
			{
				bSkip=false;
				intPostRxYear=-1;
				intPreRxYear=-1;
				
				
				if (this.m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
					this.m_ado.m_OleDbConnection.Close();

				while (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
				{
					System.Threading.Thread.Sleep(1000);
				}

				
				intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(lstFvsOutput,"Count",false);
				for (x=0;x<=intCount-1;x++)
				{
					this.m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn=true;
					this.m_oLvAlternateColors.DelegateListViewSubItem(lstFvsOutput.Items[x],COL_RUNSTATUS);
					frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,uc_PostFvsForeFrcs.COL_RUNSTATUS,"Text","");
					//see if checked
					if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(lstFvsOutput,x,"Checked",false))
						m_intCheckedCount++;
				}
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Maximum",m_intCheckedCount);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Minimum",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Value",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2,"Text","Overall Progress");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Visible",true);

				for (x=0;x<=intCount-1;x++)
				{
					this.m_oLvAlternateColors.DelegateListViewSubItem(lstFvsOutput.Items[x],COL_RUNSTATUS);
					frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","");
					this.m_oLvAlternateColors.DelegateListViewSubItem(lstFvsOutput.Items[x],COL_RUNSTATUS);
					frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"Text","");

					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Value",m_intCurrentCount);					
					if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(this.lstFvsOutput,x,"Checked",false)==true)
					{

						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Maximum",100);
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Minimum",0);
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1,"Value",0);


						string strItemDialogMsg="";
						m_intCurrentCount++;
						int intItemError=0;
						string strItemError="";
						int intItemWarning=0;
						string strItemWarning="";
						m_intCurrentStep=0;
						
						
						
						this.m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn=false;
						frmMain.g_oDelegate.ExecuteControlMethodWithParam((System.Windows.Forms.Control)lstFvsOutput,"EnsureVisible",new object[] {x});
						frmMain.g_oDelegate.SetListViewItemPropertyValue(lstFvsOutput,x,"Selected",true);
						frmMain.g_oDelegate.SetListViewItemPropertyValue(lstFvsOutput,x,"Focused",true);
						

						
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"BackColor",Color.DarkGoldenrod);
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"ForeColor",Color.White);
						frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing Audit");

						//get the variant
						strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsOutput,x,COL_VARIANT,"Text",false);
						strVariant=strVariant.Trim();

						//get the treatment
						strRx = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RX,"Text",false);
						strRx=strRx.Trim();

						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","Processing " + strVariant.Trim() + " " + strRx.Trim());
						frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Refresh");

						strDbFile = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsOutput,x,COL_MDBOUT,"Text",false);
						strDbFile = strDbFile.Trim();

						strOutDirAndFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir,"Text",false);
						strOutDirAndFile=strOutDirAndFile.Trim();	

						

						strOutDirAndFile = strOutDirAndFile  + "\\" + strDbFile;

						
						



						oAdo.OpenConnection(oAdo.getMDBConnString(strOutDirAndFile,"",""));

						m_strLogFile = oAdo.m_OleDbConnection.DataSource.ToString().Trim() + "_audit_" + m_strLogDate.Replace(" ","_") + ".txt";


						frmMain.g_oUtils.WriteText(m_strLogFile,"AUDIT LOG \r\n");
						frmMain.g_oUtils.WriteText(m_strLogFile,"--------- \r\n\r\n");
						frmMain.g_oUtils.WriteText(m_strLogFile,"Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n");
						frmMain.g_oUtils.WriteText(m_strLogFile,"Database File:" + strDbFile + "\r\n");
						frmMain.g_oUtils.WriteText(m_strLogFile,"Variant:" + strVariant + " \r\n");
						frmMain.g_oUtils.WriteText(m_strLogFile,"Treatment:" + strRx + " \r\n\r\n");
						


						strSourceTableArray=oAdo.getTableNames(oAdo.m_OleDbConnection);
						InitializeAuditLogTableArray(strSourceTableArray);

						frmMain.g_oUtils.WriteText(m_strLogFile,"-----FVS_CASES-----\r\n");

						//check to ensure the variant in the fvs cases table
						//matches the current variant

						
						strCasesTable = "fvs_cases";
						bResult = oAdo.ValuesExistNotEqualToTargetValue(oAdo.m_OleDbConnection,
							strCasesTable,
							"variant",
							strVariant.Trim(),
							false);

						this.UpdateTherm(10);

						if (oAdo.m_intError==oAdo.ErrorCodeNoErrors && bResult==true)
						{
							intItemError=-1;
							strItemError = strItemError + "ERROR:Incorrect variant found in FVS_Cases.variant column";
							strItemDialogMsg=strItemDialogMsg + strDbFile + ": Incorrect variant found in FVS_Cases.variant column\r\n";
							frmMain.g_oUtils.WriteText(m_strLogFile,"ERROR:Incorrect variant found in variant column\r\n\r\n");
							
						}
						else if (oAdo.m_intError==oAdo.ErrorCodeTableNotFound)
						{
							intItemError=-1;
							strItemError = strItemError + "FVS_Cases table missing\r\n";
							strItemDialogMsg=strItemDialogMsg + strDbFile + ":FVS_Cases table missing\r\n";
							frmMain.g_oUtils.WriteText(m_strLogFile,"ERROR: table missing\r\n\r\n");
						}
						else if (oAdo.m_intError==oAdo.ErrorCodeColumnNotFound)
						{
							intItemError=-1;
							strItemError=strItemError + "FVS_Cases.variant column not found\r\n";
							strItemDialogMsg=strItemDialogMsg + strDbFile + ":FVS_Cases.variant column not found\r\n";
							frmMain.g_oUtils.WriteText(m_strLogFile,"ERROR: variant column not found\r\n\r\n");
						}
						else
						{
							frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
						}
						UpdateTherm(10);
						//check if summary table exists
						if (intItemError==0)
						{

							frmMain.g_oUtils.WriteText(m_strLogFile,"-----FVS_SUMMARY-----\r\n");
							if (oAdo.TableExist(oAdo.m_OleDbConnection,"FVS_Summary"))
							{
								frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing Audit...FVS_SUMMARY");
								
								//
								//get the pre-treatment and post-treatment year
								//
								this.CreateSummaryTableFVSPrePostYearWorkTables(oAdo,oAdo.m_OleDbConnection,"fvs_summary");
								//list the pre and post treatment years for each stand

								//check pre and post-treatment years
								this.Validate_FvsSummaryPrePostTreatmentYear(oAdo,oAdo.m_OleDbConnection,ref intItemError,ref strItemError,ref intItemWarning,ref strItemWarning,true);
								if (intItemError==0 && intItemWarning==0)
								{
									frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
								}
								else if (intItemWarning !=0)
								{
									frmMain.g_oUtils.WriteText(m_strLogFile,strItemWarning + "\r\n");
									strItemWarning="See Log File";
								}
								if (intItemError !=0)
								{
									switch (intItemError)
									{
										case -2:
											strItemDialogMsg=strItemDialogMsg + strDbFile + ":FVS_Summary table has no records\r\n";
											break;
										case -3:
											strItemDialogMsg=strItemDialogMsg + strDbFile + ":FVS_Summary table has pre-treatment year null value detected\r\n";
											break;
										case -4:
											strItemDialogMsg=strItemDialogMsg + strDbFile + ":FVS_Summary table has post-treatment year null value detected\r\n";
											break;

									}

								}
							
								
							}
							else
							{
								intItemError=-1;
								strItemDialogMsg=strItemDialogMsg + strDbFile + ":FVS_Summary table missing\r\n";
								frmMain.g_oUtils.WriteText(m_strLogFile,"ERROR: FVS_Summary table missing\r\n\r\n");	
							}
							
						}
						UpdateTherm(10);
						if (intItemError==0)
						{
							strSourceTableArray=oAdo.getTableNames(oAdo.m_OleDbConnection);
							
							//getPrePostCounts
							GetSummaryPrePostCounts(oAdo,oAdo.m_OleDbConnection,strSourceTableArray,strSourceTableArray);
							this.CreateTreeListTableFVSPrePostYearWorkTables(oAdo,oAdo.m_OleDbConnection,"fvs_treelist","fvs_cutlist");

							UpdateTherm(10);
							

							for (y=0;y<=strSourceTableArray.Length-1;y++)
							{
								if (strSourceTableArray[y] == null) break;

								if (strSourceTableArray[y].Trim().ToUpper() == "FVS_SUMMARY" ||
									strSourceTableArray[y].Trim().ToUpper() == "FVS_CASES")
								{
								}
								else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_TREELIST")
								{
									frmMain.g_oUtils.WriteText(m_strLogFile,"-----FVS_TREELIST-----\r\n");
									frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing Audit...FVS_Treelist");

									

									intItemWarning=0;
									strItemWarning="";

									this.Validate_PreTreatmentYearForTreeList(oAdo,oAdo.m_OleDbConnection,ref intItemError,ref strItemError,ref intItemWarning,ref strItemWarning,true);

									if (intItemError==0 && intItemWarning==0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
									}
									else if (intItemWarning !=0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,strItemWarning + "\r\n");
									}

									

									if (intItemError !=0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,strItemError + "\r\n");
										if (intItemError==-3)
										{
											strItemError = "FVS_Treelist Minimum treatment year not found in FVS_Summary table (See Log File)";
											strItemDialogMsg = strItemDialogMsg  + strDbFile + ":FVS_Treelist Minimum treatment year not found in FVS_Summary table\r\n";
										}
									}
									UpdateTherm(10);
									
								}
								else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_CUTLIST")
								{
									frmMain.g_oUtils.WriteText(m_strLogFile,"-----FVS_CUTLIST-----\r\n");
									frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing Audit...FVS_Cutlist");

									intItemWarning=0;
									strItemWarning="";

									this.Validate_PostTreatmentYearForCutList(oAdo,oAdo.m_OleDbConnection,"fvs_cutlist","fvs_summary",ref intItemError,ref strItemError,ref intItemWarning,ref strItemWarning,true);

									if (intItemError==0 && intItemWarning==0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
									}
									else if (intItemWarning !=0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,strItemWarning + "\r\n");
									}

									

									if (intItemError !=0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,strItemError + "\r\n");
										if (intItemError==-3)
										{
											strItemError = "FVS_Treelist Minimum treatment year not found in FVS_Summary table (See Log File)";
											strItemDialogMsg = strItemDialogMsg  + strDbFile + ":FVS_Treelist Minimum treatment year not found in FVS_Summary table\r\n";
										}
										else if (intItemError==-4)
										{
											strItemError = "FVS_Cutlist Standid and year not found in the fvs_summary table";
											strItemDialogMsg = strItemDialogMsg  + strDbFile + ":FVS_Cutlist standid and year not found in FVS_Summary table\r\n";

										}
									}
									UpdateTherm(10);
									
								}
								else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_POTFIRE")
								{
									frmMain.g_oUtils.WriteText(m_strLogFile,"-----FVS_POTFIRE-----\r\n");
									frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing Audit...FVS_Potfire");

									CreatePotFireTablePrePostYearWorkTables(oAdo,oAdo.m_OleDbConnection,"FVS_POTFIRE");

									intItemWarning=0;
									strItemWarning="";

									this.Validate_PotFire(oAdo,oAdo.m_OleDbConnection,"FVS_POTFIRE",ref intItemError,ref strItemError,ref intItemWarning,ref strItemWarning,true);

									if (intItemError==0 && intItemWarning==0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
									}
									else if (intItemWarning !=0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,strItemWarning + "\r\n");
									}

									

									if (intItemError !=0)
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,strItemError + "\r\n");
									}
									UpdateTherm(10);
								}

								
								else
								{
									if (strSourceTableArray[y].Substring(0,4) == "FVS_")
									{
										frmMain.g_oUtils.WriteText(m_strLogFile,"-----" + strSourceTableArray[y].Trim().ToUpper() + "-----\r\n");
										frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","Processing Audit..." + strSourceTableArray[y].Trim());

										CreateGenericTablePrePostYearWorkTables(oAdo,oAdo.m_OleDbConnection,strSourceTableArray[y].Trim());

										intItemWarning=0;
										strItemWarning="";

										this.Validate_FVSGenericTable(oAdo,oAdo.m_OleDbConnection,strSourceTableArray[y].Trim(),ref intItemError,ref strItemError,ref intItemWarning,ref strItemWarning,true);

										if (intItemError==0 && intItemWarning==0)
										{
											frmMain.g_oUtils.WriteText(m_strLogFile,"OK\r\n\r\n");
										}
										else if (intItemWarning !=0)
										{
											frmMain.g_oUtils.WriteText(m_strLogFile,strItemWarning + "\r\n");
										}

									

										if (intItemError !=0)
										{
											frmMain.g_oUtils.WriteText(m_strLogFile,strItemError + "\r\n");
										}
										UpdateTherm(10);
									}
								

								}
								
								if (strSourceTableArray[y].Trim().ToUpper() != "FVS_TREELIST" &&
									strSourceTableArray[y].Trim().ToUpper() != "FVS_CUTLIST" &&
									strSourceTableArray[y].Trim().ToUpper() != "FVS_ATRTLIST")
								{
							
								}
								if (intItemError!=0) break;
								
							}
							
						}
						if (intItemError==0 && intItemWarning==0)
						{
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"BackColor",Color.Green);
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"ForeColor",Color.White);
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","AUDIT: OK");
						}
						else if (intItemError != 0)
						{
							m_intError=intItemError;
							if (strItemError.Trim().Length > 50)
							{
								strItemError = strItemError.Substring(0,45) + "....(See log file)";

							}
							if (strItemDialogMsg.Trim().Length > 0)
							{
								m_strError = m_strError + strItemDialogMsg;
							}
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"BackColor",Color.Red);
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","AUDIT ERROR:" + strItemError.Replace("\r\n"," ").Replace("ERROR:"," "));
						}
						else if (intItemWarning !=0)
						{
							if (strItemWarning.Trim().Length > 50)
							{
								strItemWarning = strItemWarning.Substring(0,45) + "....(See log file)";

							}
							frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,x,COL_RUNSTATUS,"BackColor",Color.DarkOrange);
							if (strItemWarning.Substring(0,8)=="WARNING:")
								frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","AUDIT " + strItemWarning.Replace("\r\n"," ").Replace("WARNING:"," "));
							else
								frmMain.g_oDelegate.SetListViewSubItemPropertyValue(lstFvsOutput,x,COL_RUNSTATUS,"Text","AUDIT WARNING:" + strItemWarning.Replace("\r\n"," ").Replace("WARNING:"," "));

						}
						frmMain.g_oUtils.WriteText(m_strLogFile,"Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n\r\n");
						frmMain.g_oUtils.WriteText(m_strLogFile,"**EOF**");
						
						oAdo.CloseConnection(oAdo.m_OleDbConnection);
						frmMain.g_oDelegate.ExecuteListViewItemsMethod(lstFvsOutput,"Refresh");

						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Value",100);
					    System.Threading.Thread.Sleep(2000);

					}
				}
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar2,"Value",m_intCheckedCount);
				 System.Threading.Thread.Sleep(2000);
				this.FVSRecordsFinished();
			}
			catch (System.Threading.ThreadInterruptedException err)
			{
				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch  (System.Threading.ThreadAbortException err)
			{
				if (oAdo.m_OleDbConnection != null)
				{
					if (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
					{
						oAdo.CloseConnection(oAdo.m_OleDbConnection);
					}
				}
			    this.ThreadCleanUp();
				this.CleanupThread();
				
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_PostFvsForeFrcs:Audit  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"FVS Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}

			if (DisplayAuditMessage)
			{
				if (m_intError==0) this.m_strError=m_strError + "Passed Audit";
				else m_strError = m_strError + "\r\n\r\n" + "Failed Audit";
				MessageBox.Show(m_strError,"FIA Biosum");
			}
			
			CleanupThread();

			frmMain.g_oDelegate.m_oEventThreadStopped.Set();
			this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);

			
			
			
			
		}
		private void InitializeAuditLogTableArray(string[] p_strTableArray)
		{
			int z=0;
			for (int y=0;y<=p_strTableArray.Length-1;y++)
			{
				if (p_strTableArray[y].Substring(0,4)=="FVS_")
				{
					z++;
				}
			}
			if (z==0) return;

			this.m_strFVSPrePostYearAuditTablesArray = new string[z];

			z=0;
			for (int y=0;y<=p_strTableArray.Length-1;y++)
			{
				if (p_strTableArray[y].Substring(0,4)=="FVS_")
				{
					
					this.m_strFVSPrePostYearAuditTablesArray[z]="audit_pre_post_rx_year_" + p_strTableArray[y];
					z++;
				}
			}
		}
		private void GetSummaryPrePostCounts(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string[] p_strTableIdArray,string[] p_strTableNameArray)
		{
			
			for (int y=0;y<=p_strTableIdArray.Length-1;y++)
			{
				if (p_strTableIdArray[y].Substring(0,4)=="FVS_")
				{
					
					switch (p_strTableIdArray[y].Trim().ToUpper())
					{
						case "FVS_SUMMARY":
							break;
						case "FVS_CASES":
							break;
						//case "FVS_STRCLASS":
						//	break;
						default:
							p_oAdo.m_strSQL = "ALTER TABLE " + this.m_strFVSSummaryAuditTable + " ADD COLUMN " + p_strTableIdArray[y].Trim() + "_count INTEGER";
							p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
							p_oAdo.m_strSQL = "UPDATE " + this.m_strFVSSummaryAuditTable + " SET " + p_strTableIdArray[y].Trim() + "_count=0";
							p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
							p_oAdo.m_strSQL = "UPDATE " + this.m_strFVSSummaryAuditTable + " a " + 
								"INNER JOIN " + p_strTableNameArray[y].Trim() + " b " + 
								"ON a.standid=b.standid AND " + 
								"a.year = b.year " + 
								"SET a." + p_strTableIdArray[y].Trim() + "_count = " + 
								"a." + p_strTableIdArray[y].Trim() + "_count + 1";
							p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
							break;
					}

				}
			}

		}
		private void CreateSummaryTableFVSPrePostYearWorkTables(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strSummaryTableName)
		{
			int x;

			//drop tables if they exist
			if (p_oAdo.TableExist(p_oConn,"SUMMARY_PRE_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE summary_pre_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"summary_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE summary_post_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"audit_pre_post_rx_year_fvs_summary"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE audit_pre_post_rx_year_fvs_summary");

			if (p_oAdo.TableExist(p_oConn,this.m_strFVSSummaryAuditTable))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE " + this.m_strFVSSummaryAuditTable);



			//all stands

			p_oAdo.m_strSQL = "SELECT a.standid, b.pre_rx_year, c.year_count," + 
									  "IIF(year_count=1,b.pre_rx_year,0) AS post_rx_year " + 
				              "INTO summary_PRE_RX_YEAR_work " + 
				              "FROM " + p_strSummaryTableName + " a," +
							       "(SELECT MIN(year) AS pre_rx_year,standid " + 
				                    "FROM " + p_strSummaryTableName + " " + 
				                    "GROUP BY standid) b," + 
								   "(SELECT standid, COUNT(YEAR) as year_count " + 
				                    "FROM " + p_strSummaryTableName + " " + 
				                    "GROUP BY standid) c " + 
                             "WHERE b.standid=a.standid AND " + 
				                   "c.standid=a.standid AND " + 
				                   "b.standid=c.standid";





			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

            p_oAdo.m_strSQL = "SELECT a.standid,MIN(a.year) AS post_rx_year " + 
				              "INTO summary_POST_RX_YEAR_work " + 
				              "FROM " + p_strSummaryTableName + " a, summary_PRE_RX_YEAR_work b " + 
				              "WHERE b.post_rx_year=0  AND " + 
				                    "a.standid=b.standid AND " + 
									"a.year <> b.pre_rx_year " + 
				              "GROUP BY a.standid, post_rx_year";


			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);


			p_oAdo.m_strSQL = "UPDATE summary_PRE_RX_YEAR_work a " + 
				              "INNER JOIN summary_POST_RX_YEAR_work b " + 
				              "ON a.standid=b.standid " + 
				              "SET a.post_rx_year = " + 
				                  "IIF(a.year_count=1,a.pre_rx_year,b.post_rx_year)";

			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);


			p_oAdo.m_strSQL = "SELECT STANDID,PRE_RX_YEAR,POST_RX_YEAR " + 
				"INTO audit_pre_post_rx_year_fvs_summary " + 
				"FROM summary_PRE_RX_YEAR_work";
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);


			//pre and post treatment year summary audit table
			p_oAdo.m_strSQL = "CREATE TABLE " + this.m_strFVSSummaryAuditTable + " (standid TEXT(255), " + p_oAdo.FormatReservedWordColumnName("year") + " INTEGER)";
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			p_oAdo.m_strSQL = "INSERT INTO " + this.m_strFVSSummaryAuditTable + " (standid,`year`) SELECT standid,year FROM " + p_strSummaryTableName;
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			if (p_oAdo.TableExist(p_oConn,"summary_PRE_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE summary_pre_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"summary_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE summary_post_rx_year_work");

		}
		private void CreateGenericTablePrePostYearWorkTables(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn, string p_strSourceTable)
		{
			//drop tables if they exist
			if (p_oAdo.TableExist(p_oConn,p_strSourceTable.Trim() + "_PRE_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE " + p_strSourceTable.Trim() + "_pre_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,p_strSourceTable.Trim() + "_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE " + p_strSourceTable.Trim() + "_post_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"audit_pre_post_rx_year_" + p_strSourceTable.Trim()))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE audit_pre_post_rx_year_" + p_strSourceTable.Trim());



			//all stands
			p_oAdo.m_strSQL = "SELECT STANDID, MIN(YEAR) AS PRE_RX_YEAR " + 
				"INTO " + p_strSourceTable.Trim() + "_PRE_RX_YEAR_work " + 
				"FROM " + p_strSourceTable.Trim()  +  " "  + 
				"GROUP BY STANDID";

			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			p_oAdo.m_strSQL = "SELECT a.STANDID, MIN(a.YEAR) AS POST_RX_YEAR " + 
				"INTO " + p_strSourceTable.Trim() + "_POST_RX_YEAR_work  " + 
				"FROM " + p_strSourceTable.Trim() + " a," + p_strSourceTable.Trim() + "_PRE_RX_YEAR_work b " + 
				"WHERE a.standid=b.standid AND " + 
				"a.year<> b.pre_rx_year " + 
				"GROUP BY a.STANDID";

			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			p_oAdo.m_strSQL = "SELECT a.STANDID,a.PRE_RX_YEAR,b.POST_RX_YEAR " + 
				"INTO audit_pre_post_rx_year_" + p_strSourceTable.Trim() +  " " + 
				"FROM " + p_strSourceTable.Trim() + "_POST_RX_YEAR_work b," + p_strSourceTable.Trim() + "_PRE_RX_YEAR_work a " + 
				"WHERE a.standid=b.standid";
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			if (p_oAdo.TableExist(p_oConn,p_strSourceTable.Trim() + "_PRE_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE " + p_strSourceTable.Trim() + "_pre_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,p_strSourceTable.Trim() + "_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE " + p_strSourceTable.Trim() + "_post_rx_year_work");

		}
		private void CreatePotFireTablePrePostYearWorkTables(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strPotFireTableName)
		{
			//drop tables if they exist
			if (p_oAdo.TableExist(p_oConn,"POTFIRE_PRE_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE potfire_pre_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"POTFIRE_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE potfire_post_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"audit_pre_post_rx_year_fvs_potfire"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE audit_pre_post_rx_year_fvs_potfire");



			//all stands
			p_oAdo.m_strSQL = "SELECT STANDID, MIN(YEAR) AS PRE_RX_YEAR " + 
				"INTO potfire_PRE_RX_YEAR_work " + 
				"FROM " + p_strPotFireTableName + " " + 
				"GROUP BY STANDID";

			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			p_oAdo.m_strSQL = "SELECT a.STANDID, MIN(a.YEAR) AS POST_RX_YEAR " + 
				"INTO potfire_POST_RX_YEAR_work  " + 
				"FROM " + p_strPotFireTableName + " a,potfire_PRE_RX_YEAR_work b " + 
				"WHERE a.standid=b.standid AND " + 
				"a.year<> b.pre_rx_year " + 
				"GROUP BY a.STANDID";

			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			p_oAdo.m_strSQL = "SELECT a.STANDID,a.PRE_RX_YEAR,b.POST_RX_YEAR " + 
				"INTO audit_pre_post_rx_year_fvs_potfire " + 
				"FROM potfire_POST_RX_YEAR_work b, potfire_PRE_RX_YEAR_work a " + 
				"WHERE a.standid=b.standid";
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			if (p_oAdo.TableExist(p_oConn,"potfire_PRE_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE potfire_pre_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"potfire_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE potfire_post_rx_year_work");

		}

		private void CreateTreeListTableFVSPrePostYearWorkTables(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTreeListTableName,string p_strCutListTableName)
		{
			if (p_strTreeListTableName.Trim().Length > 0)
			{
				if (p_oAdo.TableExist(p_oConn,p_strTreeListTableName))
				{
					if (p_oAdo.TableExist(p_oConn, "treelist_work"))
						p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE treelist_work");

					p_oAdo.m_strSQL = "SELECT DISTINCT standid,year INTO treelist_work FROM " + p_strTreeListTableName;
					p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
				}
			}
			
			if (p_oAdo.TableExist(p_oConn,p_strCutListTableName))
			{
				if (p_oAdo.TableExist(p_oConn, "cutlist_work"))
					p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE cutlist_work");

				p_oAdo.m_strSQL = "SELECT DISTINCT standid,year INTO cutlist_work FROM " + p_strCutListTableName;
				p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
			}

			






		}
		
		private int Validate_FvsCasesVariant(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName,string p_strVariant)
		{
			int intError=0;

			

			//check to ensure the variant in the fvs cases table
			//matches the current variant
			if (p_oAdo.getRecordCount(p_oConn,
				"SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + p_strTableName + " WHERE TRIM(variant) <> '" + p_strVariant + "')",
				p_strTableName) > 0)
			{
				intError=-1;
			}

			return intError;
			


		}

		private void Validate_FvsSummaryPrePostTreatmentYear(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,ref int p_intItemError,ref string p_strItemError, ref int p_intItemWarning,ref string p_strItemWarning,bool p_bDoWarnings)
		{
			int z=0;
			int intError=0;
			int intPreYear=-1;
			int intPostYear=-1;
			bool bWarningFirstTime=true;
			

			

			p_oAdo.m_strSQL = "SELECT standid,pre_rx_year,post_rx_year FROM audit_pre_post_rx_year_fvs_summary WHERE standid IS NOT NULL";
			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			if (p_oAdo.m_intError==0)
			{
			
				if (p_oAdo.m_OleDbDataReader.HasRows)
				{
					z=0;
					while (p_oAdo.m_OleDbDataReader.Read())
					{
						this.UpdateTherm(1);
						intPreYear=-1;
						intPostYear=-1;

					
						if (p_oAdo.m_OleDbDataReader["pre_rx_year"] != System.DBNull.Value)
						{
							intPreYear = Convert.ToInt32(p_oAdo.m_OleDbDataReader["pre_rx_year"]);
						}
						else
						{
							p_intItemError=-3;
							p_strItemError =  p_strItemError + "ERROR: Stand " + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() +  " has no pre treatment year value\r\n";
							break;
						}
						if (p_oAdo.m_OleDbDataReader["post_rx_year"] != System.DBNull.Value)
						{
							intPostYear = Convert.ToInt32(p_oAdo.m_OleDbDataReader["post_rx_year"]);
						}
						else
						{
							p_intItemError=-4;
							p_strItemError =  p_strItemError + "ERROR: Stand " + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() +  " has no post treatment year value\r\n";
							break;
						}
						if (p_bDoWarnings)
						{
							if (intPreYear+1 != intPostYear)
							{
								z++;
								if (bWarningFirstTime)
								{
									p_strItemWarning=p_strItemWarning + "WARNING: Biosum expects the initial FVS_Summary post-treatment year to equal pre-treatment year + 1 \r\n\r\n"; 
									p_strItemWarning=p_strItemWarning + "FVS_Summary plot list is...\r\n";
									bWarningFirstTime=false;
								}
								p_intItemWarning = -1;
								p_strItemWarning = p_strItemWarning + "WARNING: Stand:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() +  " PRE-TREATMENT YEAR:" + intPreYear.ToString().Trim() + " POST-TREATMENT YEAR:" + intPostYear.ToString().Trim() + "\r\n";

							}
						}
					
					}
					if (p_bDoWarnings) if (z > 0) p_strItemWarning = p_strItemWarning + "COUNT:" + z.ToString().Trim();
				}
				else
				{
					p_intItemError=-2;
					p_strItemError=p_strItemError + "ERROR: No records in table\r\n";
				}
				p_oAdo.m_OleDbDataReader.Close();
			}
			else
			{
				p_intItemError=p_oAdo.m_intError;
				p_strItemError = p_strItemError + "ERROR:" + p_oAdo.m_strError + "\r\n";
			}
			

			
		}

		/// <summary>
		/// make sure the designated table has the pre treatment year represented in the table for each plot
		/// </summary>
		/// <param name="p_oAdo"></param>
		/// <param name="p_oConn"></param>
		/// <param name="p_strTableName"></param>
		/// <returns></returns>
		private void Validate_PreTreatmentYearForTreeList(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, 
			ref int p_intItemError, ref string p_strItemError, 
			ref int p_intItemWarning, ref string p_strItemWarning, bool p_bAudit)
		{
			int z=0;
			int intError=0;
			int intPreYear=-1;
			int intYear=-1;
			string strStandId;
			int intCount1=0;
			int intCount2=0;

			//ERRORS
			//
			//see if any records 
			//
			p_oAdo.m_strSQL = "SELECT TOP 1 COUNT(*) FROM treelist_work";
			if ((int)p_oAdo.getRecordCount(p_oConn,p_oAdo.m_strSQL,"treelist_work") == 0)
			{ 
				p_intItemError=-2;
				p_strItemError=p_strItemError + "ERROR: No trees in tree list table\r\n";
				return;
			}
			//
			//ensure the tree list standid,minimum treatment year exists in the fvs summary table
			//
			if (p_oAdo.TableExist(p_oConn,"treelist_min_rx_year_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE treelist_min_rx_year_work");
			p_oAdo.m_strSQL = "SELECT MIN(YEAR) AS min_rx_year ,standid " + 
				              "INTO treelist_min_rx_year_work " + 
				              "FROM treelist_work " + 
				              "GROUP BY standid";
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
			p_oAdo.AddPrimaryKey(p_oConn,"treelist_min_rx_year_work","standid_pk","standid");
			UpdateTherm(10);
			p_oAdo.m_strSQL = "SELECT a.standid, a.min_rx_year " + 
				              "FROM treelist_min_rx_year_work a " + 
				              "WHERE NOT EXISTS " + 
				                 "(SELECT b.standid " + 
				                  "FROM FVS_summary b " + 
				                  "WHERE " + 
				                  "a.standid=b.standid AND " + 
				                  "a.min_rx_year = b.year)";
				
			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(500);			
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
			   p_intItemError=-3;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					p_strItemError= p_strItemError + "ERROR: STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " TREELIST MINIMUM YEAR:" +  p_oAdo.m_OleDbDataReader["min_rx_year"].ToString().Trim() + " not found in fvs_summary table\r\n";
				}
			  	
			}
			p_oAdo.m_OleDbDataReader.Close();

			if (p_intItemError != 0) return;

			if (p_bAudit==false) return;


			
			//WARNINGS
			//
			//get a list of summary table plots not found in the treelist table
			//
			p_oAdo.m_strSQL = "SELECT DISTINCT a.standid, b.total_count " + 
				              "FROM  " + this.m_strFVSSummaryAuditTable + " a," + 
				                   "(SELECT SUM(fvs_treelist_count) AS total_count, standid " + 
				                    "FROM " + this.m_strFVSSummaryAuditTable + " " + 
				                    "GROUP BY standid)  b " + 
							 "WHERE a.standid = b.standid AND b.total_count=0";
			p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);

			UpdateTherm(200);
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_intItemWarning=-1;
				p_strItemWarning=p_strItemWarning + "\r\nWARNING: There are pre-treatment plots in the FVS_summary table \r\n" + 
					"whose standid is not found in the FVS_Treelist table.\r\n\r\n";


				p_strItemWarning=p_strItemWarning + "FVS_Summary plot List...\r\n";
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
						
					p_strItemWarning=p_strItemWarning + "WARNING: StandId: " +  p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " not found in FVS_Treelist table\r\n";
					z++;
				}
				p_strItemWarning = p_strItemWarning + "COUNT: " + z.ToString().Trim() + "\r\n\r\n";
					
			}
			p_oAdo.m_OleDbDataReader.Close();
			
			//
			//get a list of plots who are in the treelist table but do not have the fvs summary pre-treatment year
			//
			
			//check to see if the plot +  pretreatment year count is 0 
			//If so, then check to see if the plot is in the treelist table with a 
			//different year
            p_oAdo.m_strSQL = "SELECT a.standid,a.year,a.fvs_treelist_count " + 
				              "FROM " + this.m_strFVSSummaryAuditTable + " a," + 
									"(SELECT SUM(a.fvs_treelist_count) as total_count, a.standid " + 
				                     "FROM " + this.m_strFVSSummaryAuditTable + " a " + 
				                     "INNER JOIN audit_pre_post_rx_year_fvs_summary b " + 
				                     "ON a.standid=b.standid  " + 
				                     "WHERE a.year <> b.pre_rx_year AND " + 
										   "a.fvs_treelist_count > 0 " + 
				                     "GROUP BY a.standid) b," + 
										"(SELECT a.standid,a.year,a.fvs_treelist_count " + 
										 "FROM " + this.m_strFVSSummaryAuditTable + " a " + 
										 "INNER JOIN audit_pre_post_rx_year_fvs_summary b " + 
										 "ON a.standid=b.standid " + 
										 "WHERE a.fvs_treelist_count=0 AND " + 
											   "a.year=b.pre_rx_year) c " + 
							 "WHERE a.standid=b.standid AND " + 
								   "a.standid=c.standid AND " + 
						           "b.standid=c.standid AND a.fvs_treelist_count > 0";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(1000);
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_intItemWarning=-1;
				p_strItemWarning=p_strItemWarning + "WARNING: Plots in the treelist table should have the pre-treatment year \r\n" + 
					"assigned to them. However, the plots listed below are in the treelist table but none of them are assigned the pre-treatment year\r\n\r\n";


				p_strItemWarning=p_strItemWarning + "FVS_Summary plot List...\r\n";
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
						
					p_strItemWarning=p_strItemWarning + "WARNING: StandId: " +  p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " Assigned Year: " +  p_oAdo.m_OleDbDataReader["year"].ToString().Trim() + "\r\n";
					z++;
				}
				p_strItemWarning = p_strItemWarning + "COUNT: " + z.ToString().Trim() + "\r\n\r\n";
					
			}
			p_oAdo.m_OleDbDataReader.Close();

			
    


		}
		private void Validate_PostTreatmentYearForCutList(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, 
            string p_strCutlistTableName, string p_strSummaryTableName,
			ref int p_intItemError, ref string p_strItemError, 
			ref int p_intItemWarning, ref string p_strItemWarning,bool p_bDoWarnings)
		{
			int z=0;
			int intError=0;
			int intPreYear=-1;
			int intYear=-1;
			string strStandId;
			int intCount1=0;
			int intCount2=0;

			//ERRORS
			//
			//see if any records 
			//
			p_oAdo.m_strSQL = "SELECT TOP 1 COUNT(*) FROM cutlist_work";
			if ((int)p_oAdo.getRecordCount(p_oConn,p_oAdo.m_strSQL,"cutlist_work") == 0)
			{ 
				p_intItemError=-2;
				p_strItemError=p_strItemError + "ERROR: No trees in tree cut list table\r\n";
				return;
			}
			//
			//ensure the cut list standid,treatment year exists in the fvs summary table
			//
			if (p_oAdo.TableExist(p_oConn,"cutlist_rx_year_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE cutlist_rx_year_work");

			p_oAdo.m_strSQL = "SELECT COUNT(*) AS treecount, standid,year INTO cutlist_rx_year_work FROM " + p_strCutlistTableName + " GROUP BY standid,year";
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			UpdateTherm(100);
			p_oAdo.m_strSQL = "SELECT a.standid,a.year,a.treecount " + 
				              "FROM cutlist_rx_year_work a " + 
				              "WHERE NOT EXISTS (SELECT b.standid,b.year " + 
				                                "FROM " + this.m_strFVSSummaryAuditTable + " b " + 
				                                "WHERE a.standid=b.standid AND a.year=b.year)";



			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);


			UpdateTherm(1000);			

			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_intItemError=-4;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					p_strItemError= p_strItemError + "ERROR: STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " " + 
						                                    "YEAR:" +  p_oAdo.m_OleDbDataReader["year"].ToString().Trim() + " " + 
						                                    "TREECOUNT:" + p_oAdo.m_OleDbDataReader["treecount"].ToString().Trim() + " Standid and year not found in the fvs_summary table\r\n";
				}
			  	
			}
			p_oAdo.m_OleDbDataReader.Close();

			if (p_intItemError != 0) return;

			if (p_oAdo.TableExist(p_oConn,"cutlist_min_rx_year_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE cutlist_min_rx_year_work");
			p_oAdo.m_strSQL = "SELECT MIN(YEAR) AS min_rx_year ,standid " + 
				"INTO cutlist_min_rx_year_work " + 
				"FROM cutlist_work " + 
				"GROUP BY standid";
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
			p_oAdo.AddPrimaryKey(p_oConn,"cutlist_min_rx_year_work","standid_pk","standid");
			UpdateTherm(100);
			p_oAdo.m_strSQL = "SELECT a.standid, a.min_rx_year " + 
				"FROM cutlist_min_rx_year_work a " + 
				"WHERE NOT EXISTS " + 
				"(SELECT b.standid " + 
				"FROM " + p_strSummaryTableName + " b " + 
				"WHERE " + 
				"a.standid=b.standid AND " + 
				"a.min_rx_year = b.year)";
				
			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(1000);			
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_intItemError=-3;
                if (p_bDoWarnings)
				{
					while (p_oAdo.m_OleDbDataReader.Read())
					{
						p_strItemError= p_strItemError + "ERROR: STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " CUTLIST MINIMUM YEAR:" +  p_oAdo.m_OleDbDataReader["min_rx_year"].ToString().Trim() + " not found in fvs_summary table\r\n";
					}
				}
				else 
					p_strItemError="STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " CUTLIST MINIMUM YEAR:" +  p_oAdo.m_OleDbDataReader["min_rx_year"].ToString().Trim() + " not found in fvs_summary table";
			  	
			}
			p_oAdo.m_OleDbDataReader.Close();

			if (p_intItemError != 0) return;

			if (p_bDoWarnings==false) return;

			//WARNINGS
			//
			//check for plots that have plots in the tree list file but none in the cut list table
			//
			p_oAdo.m_strSQL = "SELECT DISTINCT a.standid, c.pre_rx_year,c.fvs_treelist_count,b.total_cutlist_count " + 
				              "FROM " + this.m_strFVSSummaryAuditTable + " a," + 
									"(SELECT SUM(fvs_cutlist_count) AS total_cutlist_count, standid " + 
				                     "FROM " + this.m_strFVSSummaryAuditTable + " " + 
				                     "GROUP BY standid) b," + 
										"(SELECT a.standid,b.pre_rx_year,a.fvs_treelist_count " + 
				                         "FROM " + this.m_strFVSSummaryAuditTable + " a " + 
				                         "INNER JOIN audit_pre_post_rx_year_fvs_summary b " +  
				                         "ON a.standid=b.standid " + 
				                         "WHERE a.standid=b.standid AND " + 
				                               "a.year=b.pre_rx_year AND " + 
				                               "a.fvs_treelist_count > 0) c " +
							  "WHERE a.standid = b.standid AND " + 
				                    "a.standid=c.standid AND " + 
				                    "c.standid=b.standid AND b.total_cutlist_count=0";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(1000);
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_intItemWarning=-1;
				p_strItemWarning=p_strItemWarning + "WARNING: These plots have trees in the treelist table but none in the cutlist table for the assigned pre-treatment year \r\n\r\n";

				p_strItemWarning=p_strItemWarning + "Plot List...\r\n";
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
						
					p_strItemWarning=p_strItemWarning + "WARNING: StandId: " +  p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " " + 
                                                                 "Pre-Treatment Year: " +  p_oAdo.m_OleDbDataReader["pre_rx_year"].ToString().Trim() + " " + 
                                                                 "FVS_Treelist Count: " + p_oAdo.m_OleDbDataReader["fvs_treelist_count"].ToString().Trim() + " " + 
						                                         "FVS_Cutlist Count: " + p_oAdo.m_OleDbDataReader["total_cutlist_count"].ToString().Trim() + "\r\n";
					z++;
				}
				p_strItemWarning = p_strItemWarning + "COUNT: " + z.ToString().Trim() + "\r\n\r\n";
					
			}
			p_oAdo.m_OleDbDataReader.Close();
			//
			//post treatment year check.  Check if the post treatment year exists and if not then check for other years.
			//
			p_oAdo.m_strSQL = "SELECT  a.standid,a.year,a.fvs_cutlist_count,b.pre_rx_year,b.post_rx_year,c.treecount " + 
				              "FROM " + this.m_strFVSSummaryAuditTable + " a, " + 
				                   "audit_pre_post_rx_year_fvs_summary b," + 
								  "cutlist_rx_year_work c " + 
                              "WHERE a.standid=b.standid AND " + 
				                    "a.standid=c.standid AND " + 
				                    "a.year=c.year AND " + 
				                    "a.fvs_cutlist_count > 0  AND b.pre_rx_year + 1 <> a.year";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(500);
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_intItemWarning=-1;
				p_strItemWarning=p_strItemWarning + "WARNING: Post-treatment year should be pre-treatment + 1, however, these plots use a different post-treatment year \r\n\r\n";

				p_strItemWarning=p_strItemWarning + "Plot List...\r\n";
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
						
					p_strItemWarning=p_strItemWarning + "WARNING: StandId: " +  p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " " + 
						"Post-Treatment Year: " +  p_oAdo.m_OleDbDataReader["year"].ToString().Trim() + " " + 
						"FVS_Cutlist Count: " + p_oAdo.m_OleDbDataReader["fvs_cutlist_count"].ToString().Trim() + "\r\n";
					z++;
				}
				p_strItemWarning = p_strItemWarning + "COUNT: " + z.ToString().Trim() + "\r\n\r\n";
					
			}
			p_oAdo.m_OleDbDataReader.Close();

		}
		private void Validate_PotFire(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,
			string p_strPotFireTableName,
			ref int p_intItemError, ref string p_strItemError, 
			ref int p_intItemWarning, ref string p_strItemWarning,bool p_bDoWarnings)
		{
			int z=0;
			int intError=0;
			int intPreYear=-1;
			int intYear=-1;
			string strStandId;
			int intCount1=0;
			int intCount2=0;

			//ERRORS
			//
			//see if any records 
			//
			p_oAdo.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 standid FROM " + p_strPotFireTableName + ")";
			if ((int)p_oAdo.getRecordCount(p_oConn,p_oAdo.m_strSQL,"FVS_POTFIRE") == 0)
			{ 
				p_intItemError=-2;
				p_strItemError=p_strItemError + "ERROR: No fvs potfire records\r\n";
				return;
			}
			//
			//ensure each plot has both a pre and post year
			//
			p_oAdo.m_strSQL = "SELECT standid,pre_rx_year,post_rx_year " + 
				              "FROM audit_pre_post_rx_year_fvs_potfire " + 
				              "WHERE pre_rx_year IS NULL OR post_rx_year IS NULL";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(100);			
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_intItemError=-4;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					UpdateTherm(1);
					if (p_oAdo.m_OleDbDataReader["pre_rx_year"] == System.DBNull.Value)
					{
						p_strItemError= p_strItemError + "ERROR: STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " " + 
							" Pre-Treatment year has a null value.\r\n";
					}
					else if (p_oAdo.m_OleDbDataReader["post_rx_year"] == System.DBNull.Value)
					{
						p_strItemError= p_strItemError + "ERROR: STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " " + 
							" Post-Treatment year has a null value.\r\n";

					}
				}
			}
			p_oAdo.m_OleDbDataReader.Close();

			if (p_intItemError != 0) return;
			if (p_bDoWarnings==false) return;


			//WARNINGS
			//
			//check if pre-treatment year + 1 = post-treatment year
			//
			p_oAdo.m_strSQL = "SELECT standid,pre_rx_year,post_rx_year " + 
				              "FROM audit_pre_post_rx_year_fvs_potfire " + 
				              "WHERE pre_rx_year + 1 <> post_rx_year";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(100);			
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_strItemWarning=p_strItemWarning + "WARNING: Biosum expects the initial post-treatment year to equal pre-treatment year + 1 \r\n\r\n"; 
				p_strItemWarning=p_strItemWarning + "PotFire plot list is...\r\n";
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					UpdateTherm(1);
					
					z++;
					p_intItemWarning = -1;
					p_strItemWarning = p_strItemWarning + "WARNING: Stand:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() +  " PRE-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["pre_rx_year"].ToString().Trim() + " POST-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["post_rx_year"].ToString().Trim() + "\r\n";
				}
				
				if (z > 0) p_strItemWarning = p_strItemWarning + "COUNT:" + z.ToString().Trim()+ "\r\n\r\n";
			}
			p_oAdo.m_OleDbDataReader.Close();
			//
			//check if the pre-treatment and post-treatment years match the fvs_summary pre and post treatment years
			//
			p_oAdo.m_strSQL = "SELECT a.standid,a.pre_rx_year,a.post_rx_year,b.pre_rx_year AS summary_pre_rx_year,b.post_rx_year AS summary_post_rx_year " + 
				              "FROM audit_pre_post_rx_year_fvs_summary b,audit_pre_post_rx_year_fvs_potfire a " + 
				              "WHERE a.standid=b.standid AND (a.pre_rx_year <> b.pre_rx_year OR a.post_rx_year <> b.post_rx_year)";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(100);			
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				
				p_strItemWarning=p_strItemWarning + "WARNING: Either PotFire stand Pre-Treatment year or Post-Treatment year is not equal to the Summary stand Pre-Treatment or Post-Treatment \r\n\r\n"; 
				p_strItemWarning=p_strItemWarning + "PotFire list is...\r\n";
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					UpdateTherm(1);
					
					z++;
					p_intItemWarning = -1;
					p_strItemWarning = p_strItemWarning + "WARNING: Stand:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() +  " " + 
                                                         "POTFIRE.PRE-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["pre_rx_year"].ToString().Trim() + " " + 
						                                 "POTFIRE.POST-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["post_rx_year"].ToString().Trim() + " " + 
														 "SUMMARY.PRE-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["summary_pre_rx_year"].ToString().Trim() + " " + 
						                                 "SUMMARY.POST-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["summary_post_rx_year"].ToString().Trim() + "\r\n";
				}
				if (z > 0) p_strItemWarning = p_strItemWarning + "COUNT:" + z.ToString().Trim() + "\r\n\r\n";
			}
			p_oAdo.m_OleDbDataReader.Close();
		}
		private void Validate_FVSGenericTable(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, 
			string p_strFvsTable, ref int p_intItemError, ref string p_strItemError, 
			ref int p_intItemWarning, ref string p_strItemWarning,bool p_bDoWarnings)
		{
			int z=0;
			int intError=0;
			int intPreYear=-1;
			int intYear=-1;
			string strStandId;
			int intCount1=0;
			int intCount2=0;

			//ERRORS
			//
			//see if any records 
			//
			p_oAdo.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 standid FROM " + p_strFvsTable.Trim() + ")";
			if ((int)p_oAdo.getRecordCount(p_oConn,p_oAdo.m_strSQL,p_strFvsTable) == 0)
			{ 
				p_intItemError=-2;
				p_strItemError=p_strItemError + "ERROR: No " + p_strFvsTable + " records\r\n";
				return;
			}
			UpdateTherm(1);
			//
			//ensure each plot has both a pre and post year
			//
			p_oAdo.m_strSQL = "SELECT standid,pre_rx_year,post_rx_year " + 
				"FROM audit_pre_post_rx_year_" + p_strFvsTable.Trim() + " " + 
				"WHERE pre_rx_year IS NULL OR post_rx_year IS NULL";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(10);			
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_intItemError=-4;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					UpdateTherm(1);
					if (p_oAdo.m_OleDbDataReader["pre_rx_year"] == System.DBNull.Value)
					{
						p_strItemError= p_strItemError + "ERROR: STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " " + 
							" Pre-Treatment year has a null value.\r\n";
					}
					else if (p_oAdo.m_OleDbDataReader["post_rx_year"] == System.DBNull.Value)
					{
						p_strItemError= p_strItemError + "ERROR: STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " " + 
							" Post-Treatment year has a null value.\r\n";

					}
				}
			}
			p_oAdo.m_OleDbDataReader.Close();

			if (p_intItemError != 0) return;
			if (p_bDoWarnings==false) return;


			//WARNINGS
			//
			//check if pre-treatment year + 1 = post-treatment year
			//
			p_oAdo.m_strSQL = "SELECT standid,pre_rx_year,post_rx_year " + 
				"FROM audit_pre_post_rx_year_" + p_strFvsTable.Trim() + " " + 
				"WHERE pre_rx_year + 1 <> post_rx_year";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(10);			
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_strItemWarning=p_strItemWarning + "WARNING: Biosum expects the initial post-treatment year to equal pre-treatment year + 1 \r\n\r\n"; 
				p_strItemWarning=p_strItemWarning + p_strFvsTable.Trim() + " plot list is...\r\n";
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					UpdateTherm(1);
					z++;
					p_intItemWarning = -1;
					p_strItemWarning = p_strItemWarning + "WARNING: Stand:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() +  " PRE-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["pre_rx_year"].ToString().Trim() + " POST-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["post_rx_year"].ToString().Trim() + "\r\n";
				}
				
				if (z > 0) p_strItemWarning = p_strItemWarning + "COUNT:" + z.ToString().Trim()+ "\r\n\r\n";
			}
			p_oAdo.m_OleDbDataReader.Close();
			//
			//check if the pre-treatment and post-treatment years match the fvs_summary pre and post treatment years
			//
			p_oAdo.m_strSQL = "SELECT a.standid,a.pre_rx_year,a.post_rx_year,b.pre_rx_year AS summary_pre_rx_year,b.post_rx_year AS summary_post_rx_year " + 
				"FROM audit_pre_post_rx_year_fvs_summary b,audit_pre_post_rx_year_" + p_strFvsTable.Trim() + " a " + 
				"WHERE a.standid=b.standid AND (a.pre_rx_year <> b.pre_rx_year OR a.post_rx_year <> b.post_rx_year)";

			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			UpdateTherm(10);			
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				p_strItemWarning=p_strItemWarning + "WARNING: Either " + p_strFvsTable.Trim() + " stand Pre-Treatment year or Post-Treatment year is not equal to the Summary stand Pre-Treatment or Post-Treatment \r\n\r\n"; 
				p_strItemWarning=p_strItemWarning + p_strFvsTable.Trim() + " list is...\r\n";
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					UpdateTherm(1);
					z++;
					p_intItemWarning = -1;
					p_strItemWarning = p_strItemWarning + "WARNING: Stand:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() +  " " + 
						"POTFIRE.PRE-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["pre_rx_year"].ToString().Trim() + " " + 
						"POTFIRE.POST-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["post_rx_year"].ToString().Trim() + " " + 
						"SUMMARY.PRE-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["summary_pre_rx_year"].ToString().Trim() + " " + 
						"SUMMARY.POST-TREATMENT YEAR:" + p_oAdo.m_OleDbDataReader["summary_post_rx_year"].ToString().Trim() + "\r\n";
				}
				if (z > 0) p_strItemWarning = p_strItemWarning + "COUNT:" + z.ToString().Trim() + "\r\n\r\n";
			}
			p_oAdo.m_OleDbDataReader.Close();
		}

		

		private int Validate_FvsSummaryPrePostTreatmentYear(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName,ref int p_intPreRxYear,ref int p_intPostRxYear)
		{
			int z=0;
			int intError=0;



			p_oAdo.m_strSQL = "SELECT DISTINCT YEAR FROM " + p_strTableName + " WHERE YEAR IS NOT NULL ORDER BY YEAR";
			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			p_intPreRxYear=-1;
			p_intPostRxYear=-1;
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				z=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					UpdateTherm(1);
					z++;
					switch (z)
					{
						case 1:
							p_intPreRxYear=Convert.ToInt32(p_oAdo.m_OleDbDataReader["year"]);
							break;
						case 2:
							p_intPostRxYear = Convert.ToInt32(p_oAdo.m_OleDbDataReader["year"]);
							break;
					}
					if (z==2) break;
				}
				p_oAdo.m_OleDbDataReader.Close();
				if (p_intPreRxYear > -1 &&
					p_intPostRxYear > -1)
				{
					if (p_intPreRxYear + 1 != p_intPostRxYear)
					{
						intError=-1;
					}
				}
				else
				{
					intError=-2;
				}
			}
			else
				intError=-3;

			return intError;
		}
		private void Validate_FvsSummaryPrePostTreatmentYear(ado_data_access p_oAdo, string p_strConn,string p_strTableName,ref int p_intPreRxYear,ref int p_intPostRxYear)
		{
			int z=0;
			int intError=0;
			p_oAdo.OpenConnection(p_strConn);
			intError = Validate_FvsSummaryPrePostTreatmentYear(p_oAdo,p_oAdo.m_OleDbConnection,p_strTableName,ref p_intPreRxYear,ref p_intPostRxYear);
			p_oAdo.CloseConnection(p_oAdo.m_OleDbConnection);

		}

		


		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			CancelThread();
		}
		private void CancelThread()
		{
			bool bAbort=frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel processing (Y/N)?");
			if (bAbort)
			{
				if (frmMain.g_oDelegate.m_oThread.IsAlive)
				{
					frmMain.g_oDelegate.m_oThread.Join();
				}
				frmMain.g_oDelegate.StopThread();
				CleanupThread();
			}
		}
		private void CleanupThread()
		{
			this.grpboxAppend.Enabled=true;
			this.grpboxAudit.Enabled=true;
			this.btnChkAll.Enabled=true;
			this.btnClearAll.Enabled=true;
			this.btnRefresh.Enabled=true;
			this.btnClose.Enabled=true;
			this.btnHelp.Enabled=true;
			this.btnViewLogFile.Enabled=true;
			this.btnAuditDb.Enabled=true;
			this.btnCancel.Visible=false;

			this.ParentForm.Enabled=true;
			this.Enabled=true;
		}

		private void lstFvsOutput_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			lstFvsOutput.Items[e.Index].Selected=true;
		}

		private void btnViewLogFile_Click(object sender, System.EventArgs e)
		{
			if (this.lstFvsOutput.SelectedItems.Count==0) return;

			

			string strSearch = this.lstFvsOutput.SelectedItems[0].SubItems[COL_MDBOUT].Text.Trim() + "_AUDIT_*.txt";
			
			string strDirectory = this.txtOutDir.Text.Trim();
			
			string[]  strFiles= System.IO.Directory.GetFiles(strDirectory,strSearch);

			FIA_Biosum_Manager.frmDialog oDlg = new frmDialog();

			oDlg.uc_select_list_item1.lblTitle.Text = "Open Audit Log File";
			oDlg.uc_select_list_item1.listBox1.Sorted = true;
			for (int x=0;x<=strFiles.Length - 1;x++)
			{
				oDlg.uc_select_list_item1.listBox1.Items.Add(strFiles[x].Substring(strDirectory.Length+1,strFiles[x].Length - strDirectory.Length - 1));
			}
			if (oDlg.uc_select_list_item1.listBox1.Items.Count > 0) oDlg.uc_select_list_item1.listBox1.SelectedIndex = oDlg.uc_select_list_item1.listBox1.Items.Count-1;
			oDlg.uc_select_list_item1.lblMsg.Text = "Log File Contents of " + strDirectory;
			oDlg.uc_select_list_item1.lblMsg.Show();

			oDlg.uc_select_list_item1.Show();

			DialogResult result = oDlg.ShowDialog();
			if (result==DialogResult.OK)
			{
				string strDirAndFile = strDirectory + "\\" + oDlg.uc_select_list_item1.listBox1.SelectedItems[0].ToString().Trim();
				System.Diagnostics.Process proc = new System.Diagnostics.Process();
				proc.StartInfo.UseShellExecute = true;
				try
				{
					proc.StartInfo.FileName = strDirAndFile;
				}
				catch
				{
				}
				try
				{
					proc.Start();
				}
				catch (Exception err)
				{
					MessageBox.Show("!!Error!! \n" + 
						"Module - uc_PostFvsForeFrcs:btnViewLogFile_Click \n" + 
						"Err Msg - " + err.Message,
						"View Script",System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
				}
				proc=null;
			}

						

			//strOutDirAndFile = strOutDirAndFile  + "\\" + strDbFile;



		}

		private void btnAuditDb_Click(object sender, System.EventArgs e)
		{

			if (this.lstFvsOutput.SelectedItems.Count==0) return;

			
			

		    string strDbFile = this.lstFvsOutput.SelectedItems[0].SubItems[COL_MDBOUT].Text.Trim();
			string strOutDirAndFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir,"Text",false);
			strOutDirAndFile=strOutDirAndFile.Trim();	
			strOutDirAndFile = strOutDirAndFile  + "\\" + strDbFile;
			ado_data_access oAdo = new ado_data_access();
			if (!oAdo.TableExist(oAdo.getMDBConnString(strOutDirAndFile,"",""),this.m_strFVSSummaryAuditTable)) 
			{
				oAdo=null;
				return;
			}

			FIA_Biosum_Manager.frmGridView oFrm = new frmGridView();
			

			
			oAdo.m_strSQL = "SELECT * FROM " + this.m_strFVSSummaryAuditTable;
			oFrm.LoadDataSet(oAdo.getMDBConnString(strOutDirAndFile,"",""),oAdo.m_strSQL,this.m_strFVSSummaryAuditTable);
		
			
			
			oFrm.Text = "Database: Browse (Pre/Post Audit Tables)";
			if (this.m_strFVSPrePostYearAuditTablesArray!=null)
			{
				for (int x=0;x<=this.m_strFVSPrePostYearAuditTablesArray.Length-1;x++)
				{
					if (oAdo.TableExist(oAdo.getMDBConnString(strOutDirAndFile,"",""),this.m_strFVSPrePostYearAuditTablesArray[x].Trim()))
					{
						oFrm.LoadDataSet(oAdo.getMDBConnString(strOutDirAndFile,"",""),"SELECT * FROM " + this.m_strFVSPrePostYearAuditTablesArray[x].Trim(),m_strFVSPrePostYearAuditTablesArray[x].Trim());
					}
				}
			}
			oFrm.TileGridViews();
			oFrm.Show();
			oFrm.Focus();
		}
		private void UpdateTherm(int p_intIncrement)
		{
			if (m_frmTherm != null)
			{
				m_intCurrentStep = m_intCurrentStep + p_intIncrement;
	//			if ((int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Maximum",false) >= m_intCurrentStep &&
	//				this.m_intCurrentStep <=this.m_intTotalSteps)
				if (this.m_intCurrentStep <=this.m_intTotalSteps)
				{
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Value",(int)Math.Round((double)(m_intCurrentStep* 100)/m_intTotalSteps,0));
					//frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)m_frmTherm.progressBar1,"Refresh");
				}
				else
				{
					//auto expand
					m_intTotalSteps = m_intTotalSteps + 1000;
				}
			}
		}

		private void lstFvsOutput_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			int x;
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = this.lstFvsOutput.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lstFvsOutput.Items[this.lstFvsOutput.TopItem.Index + (int)dblRow-1].Selected=true;
					this.m_oLvAlternateColors.DelegateListViewItem(lstFvsOutput.Items[this.lstFvsOutput.TopItem.Index + (int)dblRow-1]);
				}
			}
			catch 
			{
			}
		}

		private void lstFvsOutput_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lstFvsOutput.SelectedItems.Count > 0)
				m_oLvAlternateColors.DelegateListViewItem(lstFvsOutput.SelectedItems[0]);
		}

	
		public bool DisplayAuditMessage
		{
			get {return _bDisplayAuditMsg;}
			set {_bDisplayAuditMsg=value;}
		}


		
	}
}
