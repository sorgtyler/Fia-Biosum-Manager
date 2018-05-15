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
	/// Summary description for uc_scenario_notes.
	/// </summary>
	public class uc_core_scenario_run : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmCoreScenario _frmScenario=null;
		public System.Windows.Forms.Button btnViewLog;
		public System.Windows.Forms.Button btnAccess;
		public System.Windows.Forms.Button btnViewScenarioTables;
        public System.Windows.Forms.Button btnViewAuditTables;
		public System.Windows.Forms.Label lblMsg;
		public System.Windows.Forms.Button btnCancel;

		private int m_intError=0;
		public System.Data.DataSet m_ds;
		public System.Data.OleDb.OleDbConnection m_conn;
		public System.Data.OleDb.OleDbDataAdapter m_da;
		public RunCore m_oRunCore;
		public string m_strCustomPlotSQL="";
		private FIA_Biosum_Manager.frmGridView m_frmGridView;

		public bool m_bUserCancel=false;
		private bool m_bAbortThread=false;
		private System.Threading.Thread m_thread=null;
		public System.Windows.Forms.Label m_lblCurrentProcessStatus;

		private FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.Variables _oFVSPrePostVariables;
		private FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.Variable_Collection _oFVSPrePostOptimization;
		private FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection _oFVSPrePostTieBreaker;

        private System.Windows.Forms.Panel panel1;

        private ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();
        public ListViewEmbeddedControls.ListViewEx listViewEx1;
        private ColumnHeader colNull;
        private ColumnHeader colDesc;
        private ColumnHeader colStatus;
        private Panel pnlFileSizeMonitor;
        public uc_filesize_monitor uc_filesize_monitor3;
        public uc_filesize_monitor uc_filesize_monitor2;
        public uc_filesize_monitor uc_filesize_monitor1;
        public uc_filesize_monitor uc_filesize_monitor4;



        private string[] m_strVariantArray = null;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_core_scenario_run()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
            LoadListView();
           


		}
        private void LoadListView()
        {
            //
            //INITIALIZE LISTVIEW ALTERNATE ROW COLORS
            //
           
            this.m_oLvAlternateColors.InitializeRowCollection();
            this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceListView = this.listViewEx1;
            this.m_oLvAlternateColors.CustomFullRowSelect = true;
            m_oLvAlternateColors.ColumnsToNotUpdate("1");
            if (frmMain.g_oGridViewFont != null) this.listViewEx1.Font = frmMain.g_oGridViewFont;

          
            //
            //Validate Rule Definitions
            //
            this.AddListViewRowItem("Validate Rule Definitions",false);
            //
            //Save Rule Definitions
            //
            this.AddListViewRowItem("Save Rule Definitions",false);
            //
            //Initialize and Load Variables
            //
            this.AddListViewRowItem("Initialize and Load Variables", false);
            //
            //Accessibility
            //
            this.AddListViewRowItem("Determine If Stand And Conditions Are Accessible For Treatment And Harvest",false);
            //
            //Least Expensive Routes
            //
            this.AddListViewRowItem("Get Least Expensive Route From Stand To Wood Processing Facility",false);
            //
            //
            //
            this.AddListViewRowItem("Sum Tree Yields, Volume, And Value For A Stand And Treatment",false);
            //
            //
            //
            this.AddListViewRowItem("Sum Tree Yields, Volume, And Value For A Stand And Treatment Package", false);
            //
            //
            //
            this.AddListViewRowItem("Apply User Defined Filters And Get Valid Stand Combinations", false);
            //
            //
            //
            this.AddListViewRowItem("Populate Valid Combination Audit Data", true);
            //
            //
            //
            this.AddListViewRowItem("Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Treatment", false);
            //
            //
            //
            this.AddListViewRowItem("Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Treatment Package", false);
            //
            //
            //
            this.AddListViewRowItem("Cycle 1: Identify Effective Treatments For Each Stand", false);
            //
            //
            //
            this.AddListViewRowItem("Cycle 1: Optimize the Effective Treatments For Each Stand", false);
            //
            //
            //
            this.AddListViewRowItem("Cycle 1: Load Tie Breaker Tables", false);
            //
            //
            //
            this.AddListViewRowItem("Cycle 1: Identify The Best Effective Treatment For Each Stand", false);
            //
            //
            //
            this.AddListViewRowItem("Cycle 1: Summarize The Best Effective Treatment Yields, Revenue, Costs, And Acres By Stand", false);
            //
            //
            //
            this.AddListViewRowItem("Cycle 1: Summarize The Best Effective Treatment Yields, Revenue, Costs, And Acres By Wood Processing Facility", false);
            //
            //
            //
            this.AddListViewRowItem("Cycle 1: Summarize The Best Effective Treatment Yields, Revenue, Costs, And Acres By Land Ownership Groups", false);
            //
            //
            //
            this.AddListViewRowItem("All Treatments: Summarize Treatment Yields, Revenue, Costs, And Acres By Stand", false);
           

            this.listViewEx1.Columns[2].Width = -1;



        }

        private void AddListViewRowItem(string p_strDescription,bool p_bCheckBox)
        {
            System.Windows.Forms.ListViewItem entryListItem = null;
          
            entryListItem = listViewEx1.Items.Add(" ");


            entryListItem.UseItemStyleForSubItems = false;

            this.m_oLvAlternateColors.AddRow();
            this.m_oLvAlternateColors.AddColumns(listViewEx1.Items.Count - 1, listViewEx1.Columns.Count);

           
           
            if (p_bCheckBox)
            {
                System.Windows.Forms.CheckBox oCheckBox = new CheckBox();

                listViewEx1.AddEmbeddedControl(oCheckBox, 0, listViewEx1.Items.Count - 1, System.Windows.Forms.DockStyle.Fill);
               
                oCheckBox.Text = "";
                oCheckBox = (CheckBox)listViewEx1.GetEmbeddedControl(0, listViewEx1.Items.Count - 1);
                oCheckBox.Visible = p_bCheckBox;
                oCheckBox.Show();
                oCheckBox.Enabled = p_bCheckBox;
                oCheckBox.Checked = true;
            }
           



           
            //ProgressBar
            entryListItem.SubItems.Add(" ");
            ProgressBarBasic.ProgressBarBasic pb = new ProgressBarBasic.ProgressBarBasic();
            pb.Value = 0;
            pb.Name = "ProgressBar" + listViewEx1.Items.Count.ToString().Trim();
            pb.Orientation = System.Windows.Forms.Orientation.Horizontal;
            pb.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
            pb.BackColor = Color.LightGray;
            pb.ForeColor = Color.Gold;
            pb.Visible = true;
            // Embed the ProgressBar in Column 1
            listViewEx1.AddEmbeddedControl(pb, 1, listViewEx1.Items.Count - 1);

            pb = (ProgressBarBasic.ProgressBarBasic)listViewEx1.GetEmbeddedControl(1, listViewEx1.Items.Count - 1);

            frmMain.g_oDelegate.SetControlPropertyValue(pb, "Maximum", 100);

            frmMain.g_oDelegate.SetControlPropertyValue(pb, "Minimum", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(pb, "Value", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(pb, "Text", "0%");

            //Description
            entryListItem.SubItems.Add(p_strDescription);
            this.m_oLvAlternateColors.ListViewItem(listViewEx1.Items[listViewEx1.Items.Count - 1]);
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.pnlFileSizeMonitor = new System.Windows.Forms.Panel();
            this.uc_filesize_monitor3 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor2 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor1 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.listViewEx1 = new ListViewEmbeddedControls.ListViewEx();
            this.colNull = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colStatus = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colDesc = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lblMsg = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnAccess = new System.Windows.Forms.Button();
            this.btnViewScenarioTables = new System.Windows.Forms.Button();
            this.btnViewAuditTables = new System.Windows.Forms.Button();
            this.btnViewLog = new System.Windows.Forms.Button();
            this.uc_filesize_monitor4 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.pnlFileSizeMonitor.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(926, 508);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.pnlFileSizeMonitor);
            this.panel1.Controls.Add(this.listViewEx1);
            this.panel1.Controls.Add(this.lblMsg);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.lblTitle);
            this.panel1.Controls.Add(this.btnAccess);
            this.panel1.Controls.Add(this.btnViewScenarioTables);
            this.panel1.Controls.Add(this.btnViewAuditTables);
            this.panel1.Controls.Add(this.btnViewLog);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(920, 489);
            this.panel1.TabIndex = 40;
            // 
            // pnlFileSizeMonitor
            // 
            this.pnlFileSizeMonitor.AutoScroll = true;
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor4);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor3);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor2);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor1);
            this.pnlFileSizeMonitor.Location = new System.Drawing.Point(12, 380);
            this.pnlFileSizeMonitor.Name = "pnlFileSizeMonitor";
            this.pnlFileSizeMonitor.Size = new System.Drawing.Size(829, 99);
            this.pnlFileSizeMonitor.TabIndex = 69;
            // 
            // uc_filesize_monitor3
            // 
            this.uc_filesize_monitor3.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor3.Information = "";
            this.uc_filesize_monitor3.Location = new System.Drawing.Point(429, 10);
            this.uc_filesize_monitor3.Name = "uc_filesize_monitor3";
            this.uc_filesize_monitor3.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor3.TabIndex = 2;
            this.uc_filesize_monitor3.Visible = false;
            // 
            // uc_filesize_monitor2
            // 
            this.uc_filesize_monitor2.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor2.Information = "";
            this.uc_filesize_monitor2.Location = new System.Drawing.Point(206, 10);
            this.uc_filesize_monitor2.Name = "uc_filesize_monitor2";
            this.uc_filesize_monitor2.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor2.TabIndex = 1;
            this.uc_filesize_monitor2.Visible = false;
            // 
            // uc_filesize_monitor1
            // 
            this.uc_filesize_monitor1.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor1.Information = "";
            this.uc_filesize_monitor1.Location = new System.Drawing.Point(5, 10);
            this.uc_filesize_monitor1.Name = "uc_filesize_monitor1";
            this.uc_filesize_monitor1.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor1.TabIndex = 0;
            this.uc_filesize_monitor1.Visible = false;
            // 
            // listViewEx1
            // 
            this.listViewEx1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colNull,
            this.colStatus,
            this.colDesc});
            this.listViewEx1.FullRowSelect = true;
            this.listViewEx1.GridLines = true;
            this.listViewEx1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listViewEx1.Location = new System.Drawing.Point(12, 33);
            this.listViewEx1.MultiSelect = false;
            this.listViewEx1.Name = "listViewEx1";
            this.listViewEx1.Size = new System.Drawing.Size(896, 324);
            this.listViewEx1.TabIndex = 42;
            this.listViewEx1.UseCompatibleStateImageBehavior = false;
            this.listViewEx1.View = System.Windows.Forms.View.Details;
            // 
            // colNull
            // 
            this.colNull.Text = "Optional";
            this.colNull.Width = 55;
            // 
            // colStatus
            // 
            this.colStatus.Text = "Run Status";
            this.colStatus.Width = 250;
            // 
            // colDesc
            // 
            this.colDesc.Text = "Description";
            this.colDesc.Width = 502;
            // 
            // lblMsg
            // 
            this.lblMsg.Enabled = false;
            this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg.ForeColor = System.Drawing.Color.Black;
            this.lblMsg.Location = new System.Drawing.Point(9, 360);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(660, 16);
            this.lblMsg.TabIndex = 38;
            this.lblMsg.Text = "lblMsg";
            this.lblMsg.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.Black;
            this.btnCancel.Location = new System.Drawing.Point(83, 3);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(72, 24);
            this.btnCancel.TabIndex = 39;
            this.btnCancel.Text = "Start";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(8, 5);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(69, 23);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Run";
            // 
            // btnAccess
            // 
            this.btnAccess.Enabled = false;
            this.btnAccess.ForeColor = System.Drawing.Color.Black;
            this.btnAccess.Location = new System.Drawing.Point(315, 7);
            this.btnAccess.Name = "btnAccess";
            this.btnAccess.Size = new System.Drawing.Size(120, 20);
            this.btnAccess.TabIndex = 33;
            this.btnAccess.Text = "Microsoft Access";
            this.btnAccess.Click += new System.EventHandler(this.btnAccess_Click);
            // 
            // btnViewScenarioTables
            // 
            this.btnViewScenarioTables.ForeColor = System.Drawing.Color.Black;
            this.btnViewScenarioTables.Location = new System.Drawing.Point(189, 7);
            this.btnViewScenarioTables.Name = "btnViewScenarioTables";
            this.btnViewScenarioTables.Size = new System.Drawing.Size(120, 20);
            this.btnViewScenarioTables.TabIndex = 32;
            this.btnViewScenarioTables.Text = "View Results Tables";
            this.btnViewScenarioTables.Click += new System.EventHandler(this.btnViewScenarioTables_Click);
            // 
            // btnViewAuditTables
            // 
            this.btnViewAuditTables.ForeColor = System.Drawing.Color.Black;
            this.btnViewAuditTables.Location = new System.Drawing.Point(439, 7);
            this.btnViewAuditTables.Name = "btnViewAuditTables";
            this.btnViewAuditTables.Size = new System.Drawing.Size(120, 20);
            this.btnViewAuditTables.TabIndex = 31;
            this.btnViewAuditTables.Text = "View Audit Data";
            this.btnViewAuditTables.Click += new System.EventHandler(this.btnViewAuditTables_Click);
            // 
            // btnViewLog
            // 
            this.btnViewLog.ForeColor = System.Drawing.Color.Black;
            this.btnViewLog.Location = new System.Drawing.Point(565, 7);
            this.btnViewLog.Name = "btnViewLog";
            this.btnViewLog.Size = new System.Drawing.Size(96, 20);
            this.btnViewLog.TabIndex = 34;
            this.btnViewLog.Text = "View Log File";
            this.btnViewLog.Click += new System.EventHandler(this.btnViewLog_Click);
            // 
            // uc_filesize_monitor4
            // 
            this.uc_filesize_monitor4.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor4.Information = "";
            this.uc_filesize_monitor4.Location = new System.Drawing.Point(616, 10);
            this.uc_filesize_monitor4.Name = "uc_filesize_monitor4";
            this.uc_filesize_monitor4.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor4.TabIndex = 3;
            this.uc_filesize_monitor4.Visible = false;
            // 
            // uc_core_scenario_run
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_core_scenario_run";
            this.Size = new System.Drawing.Size(926, 508);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.pnlFileSizeMonitor.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Visible=false;
			
			((frmCoreScenario)this.ParentForm).Height = 0 ; //((frmScenario)this.ParentForm).grpboxMenu.Height * 2;
		}

		private void btnAccess_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo.UseShellExecute = true;
			try
			{
                proc.StartInfo.FileName = ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";		
            }
			catch
			{
			}
			try
			{
				proc.Start();
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.Message);
			}
		}

		private void btnViewAuditTables_Click(object sender, System.EventArgs e)
		{
            viewAuditTables();	
           

		}
        private void viewAuditTables()
        {
            string strMDBPathAndFile = "";
            string strConn = "";
            string strTable = "";
            int x;
            ado_data_access oAdo = new ado_data_access();
            if (m_strVariantArray == null)
            {
                FIA_Biosum_Manager.Datasource oDs = new Datasource(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                    this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim());
                m_strVariantArray = oDs.getVariants();
                oDs = null;
            }
            this.m_frmGridView = new frmGridView();
            this.m_frmGridView.Text = "Core Analysis: Audit";
            lblMsg.Text = "";
            lblMsg.Show();
            for (x = 0; x <= m_strVariantArray.Length - 1; x++)
            {
                 strMDBPathAndFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\audit_" + m_strVariantArray[x].Trim() + ".mdb";
                 if (System.IO.File.Exists(strMDBPathAndFile) == true)
                 {
                     oAdo.OpenConnection(oAdo.getMDBConnString(strMDBPathAndFile, "", ""));
                     if (oAdo.TableExist(oAdo.m_OleDbConnection, frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName) == true)
                     {
                         strTable = frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName;
                         strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" +
                                    strMDBPathAndFile + ";User Id=admin;Password=;";

                         this.lblMsg.Text = strTable + "_" + m_strVariantArray[x].Trim();
                         this.lblMsg.Refresh();

                         this.m_frmGridView.LoadDataSet(strConn, "select * from " + strTable + " " + strTable + "_" + m_strVariantArray[x].Trim(), strTable + "_" + m_strVariantArray[x].Trim());
                         
                     }
                     if (oAdo.TableExist(oAdo.m_OleDbConnection, frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName) == true)
                     {
                         strTable = frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName;
                         strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" +
                                    strMDBPathAndFile + ";User Id=admin;Password=;";

                         this.lblMsg.Text = strTable + "_" + m_strVariantArray[x].Trim();
                         this.lblMsg.Refresh();

                         this.m_frmGridView.LoadDataSet(strConn, "select * from " + strTable + " " + strTable + "_" + m_strVariantArray[x].Trim(), strTable + "_" + m_strVariantArray[x].Trim());
                     }
                     oAdo.CloseConnection(oAdo.m_OleDbConnection);
                 }
            }
            this.m_frmGridView.TileGridViews();
            this.m_frmGridView.Show();
            this.m_frmGridView.Focus();
            lblMsg.Text = "";
            lblMsg.Refresh();
            lblMsg.Hide();
            oAdo = null;

            
        }

        private void btnViewScenarioTables_Click(object sender, System.EventArgs e)
		{
			this.viewScenarioTables();
		}

		/// <summary>
		/// every scenario_results.mdb table is viewed in a uc_gridview control
		/// </summary>
		private void viewScenarioTables()
		{
			string strMDBPathAndFile="";
			string strConn="";
			string strSQL="";
			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			
			
			strMDBPathAndFile = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";

			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					
					this.lblMsg.Text = "";
					this.lblMsg.Visible=true;
					
						strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
							strMDBPathAndFile + ";User Id=admin;Password=;";
					
					this.m_frmGridView = new frmGridView();
					this.m_frmGridView.Text = "Core Analysis: Run Scenario Results ("  + this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + ")";
					for (int x=0; x <= intCount-1;x++)
					{
						this.lblMsg.Text = strTableNames[x];
						this.lblMsg.Refresh();
						strSQL = "select * from " + strTableNames[x].Trim();
						this.m_frmGridView.LoadDataSet(strConn,strSQL,strTableNames[x].Trim());
						

					}
					
					this.lblMsg.Text="";
					this.lblMsg.Visible=false;
					if (intCount > 1) this.m_frmGridView.TileGridViews();
					this.m_frmGridView.Show();
					this.m_frmGridView.Focus();


				}
				else
				{
					MessageBox.Show("No Tables Found In " + strMDBPathAndFile);
				}

			}
			p_dao = null;
		}

		private void btnViewLog_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo.UseShellExecute = true;
			try
			{
				proc.StartInfo.FileName = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runlog.txt";
			}
			catch
			{
			}
			try
			{
				proc.Start();
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.Message);
			}
		}

		

		private void btnCancel_Click(object sender, System.EventArgs e)
		{


			DialogResult result;

			if (this.btnCancel.Text.Trim().ToUpper() == "CANCEL")
			{

				if (this.m_bAbortThread==true) return;

				if (frmMain.g_oDelegate.m_oThread == null)
				{
					result =  MessageBox.Show("Cancel Running The Scenario (Y/N)?","Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
					switch (result) 
					{
						case DialogResult.Yes:
							this.btnCancel.Text = "Start";
							this.m_bUserCancel=true;
                            uc_filesize_monitor1.EndMonitoringFile();
                            uc_filesize_monitor2.EndMonitoringFile();
							return;
						case DialogResult.No:
							return;
					}
				}
				else
				{
					if (frmMain.g_oDelegate.m_oThread.IsAlive)
					{  
						frmMain.g_oDelegate.AbortProcessing("Cancel Processing","Cancel Running The Scenario (Y/N)?");
						if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
						{
							frmMain.g_oDelegate.StopThread();
                            if (FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic != null)
                            {
                                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                                frmMain.g_oDelegate.SetControlPropertyValue((Control)FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "TextStyle", ProgressBarBasic.ProgressBarBasic.TextStyleType.Text);
                                frmMain.g_oDelegate.SetControlPropertyValue((Control)FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Text", "!!Cancelled!!");
                            }
                            if (m_lblCurrentProcessStatus != null)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((Control)m_lblCurrentProcessStatus, "ForeColor", Color.Red);
                                frmMain.g_oDelegate.SetControlPropertyValue((Control)m_lblCurrentProcessStatus, "Text", "!!Cancelled!!");

                            }
                            frmMain.g_oDelegate.SetControlPropertyValue((Control)btnCancel, "Text", "Start");
							this.m_bUserCancel=true;
                            uc_filesize_monitor1.EndMonitoringFile();
                            uc_filesize_monitor2.EndMonitoringFile();

						}
					}

				}

				
			}
			else
			{
				RunScenario();
			}
		}
        public static void UpdateThermPercent(ProgressBarBasic.ProgressBarBasic p_oPb, int p_intMin, int p_intMax, int p_intValue)
        {
            p_oPb.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
            int intPercent = (int)(((double)(p_intValue - p_intMin) /
                (double)(p_intMax - p_intMin)) * 100);

            frmMain.g_oDelegate.SetControlPropertyValue(p_oPb, "Value", intPercent);

            frmMain.g_oDelegate.ExecuteControlMethod(p_oPb, "Refresh");

        }
        public static void UpdateThermText(ProgressBarBasic.ProgressBarBasic p_oPb, string p_strText)
        {
           p_oPb.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Text;
           frmMain.g_oDelegate.SetControlPropertyValue(p_oPb,"Text",p_strText);
           frmMain.g_oDelegate.ExecuteControlMethod(p_oPb, "Refresh");
        }
        public static int GetListViewItemIndex(ListViewEmbeddedControls.ListViewEx p_oLv, string p_strDesc)
        {
            
            int intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(p_oLv,"Count",false);
            for (int x = 0; x <= intCount - 1; x++)
            {
                string strDesc = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(p_oLv, x, 2, "Text", false);
                if (strDesc.Trim().ToUpper() == p_strDesc.Trim().ToUpper())
                {
                    return x;
                }
            }
            return -1;
        }

        private void RunScenario()
		{

            int x;

			this.lblMsg.Text = "";
            this.btnCancel.Enabled = false;
            this.btnViewAuditTables.Enabled = false;
            this.btnViewLog.Enabled = false;
            this.btnViewScenarioTables.Enabled = false;
            this.btnAccess.Enabled = false;

            for (x = 0; x <= listViewEx1.Items.Count - 1; x++)
            {
                ProgressBarBasic.ProgressBarBasic pb = new ProgressBarBasic.ProgressBarBasic();
               
                pb = (ProgressBarBasic.ProgressBarBasic)listViewEx1.GetEmbeddedControl(1, x);
                pb.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
                pb.TextColor = Color.Black;
                frmMain.g_oDelegate.SetControlPropertyValue(pb, "Maximum", 100);

                frmMain.g_oDelegate.SetControlPropertyValue(pb, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(pb, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(pb, "Text", "0%");
            }


			this.Refresh();
            this.listViewEx1.Items[0].Selected = true;
            FIA_Biosum_Manager.RunCore.g_bCoreRun = true;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic =(ProgressBarBasic.ProgressBarBasic)listViewEx1.GetEmbeddedControl(1, 0);
            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = 0;

			this.val_CoreRunData();
			if (this.m_intError==0)
			{
                this.btnCancel.Enabled = true;
				this.btnCancel.Text = "Cancel";
				this.btnCancel.Refresh();
				this.btnViewAuditTables.Enabled=false;
				this.btnViewScenarioTables.Enabled=false;
				this.btnAccess.Enabled=false;
				this.btnViewLog.Enabled=false;
				this.m_bUserCancel=false;


				frmMain.g_oDelegate.CurrentThreadProcessIdle=false;
				frmMain.g_oDelegate.CurrentThreadProcessAborted=false;
				frmMain.g_oDelegate.CurrentThreadProcessDone=false;
				frmMain.g_oDelegate.CurrentThreadProcessStarted=false;
				frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(this.StartScenarioRunProcess));
				frmMain.g_oDelegate.InitializeThreadEvents();
				frmMain.g_oDelegate.m_oThread.IsBackground=true;
				frmMain.g_oDelegate.m_oThread.Start();
			}
			else
			{
                this.btnCancel.Enabled = true;
                this.btnViewAuditTables.Enabled = true;
                this.btnViewLog.Enabled = true;
                this.btnViewScenarioTables.Enabled = true;
                this.btnAccess.Enabled = true;
				if (this.ReferenceCoreScenarioForm.WindowState == System.Windows.Forms.FormWindowState.Minimized)
					this.ReferenceCoreScenarioForm.WindowState = System.Windows.Forms.FormWindowState.Normal;
				ReferenceCoreScenarioForm.Focus();

			}

		}
		private void StartScenarioRunProcess()
		{
			frmMain.g_oDelegate.CurrentThreadProcessStarted=true;
			this.m_oRunCore = new RunCore(this);
			if (!this.m_bAbortThread)
			{
				
			}
			System.Threading.Thread.Sleep(1000);
			
		}
		private void StopThread()
		{
			this.m_thread.Suspend();
            DialogResult result =  MessageBox.Show("Cancel Running The Scenario (Y/N)?","Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					this.m_bAbortThread=true;
					this.m_thread.Resume();
					this.m_thread.Abort();
					while (this.m_thread.IsAlive)
					{
						this.lblMsg.Text = "Attempting To Abort Process...Stand By";
						this.m_thread.Join(1000);
						if (this.m_thread.IsAlive) this.m_thread.Abort();
						System.Windows.Forms.Application.DoEvents();
					}
					return;
				case DialogResult.No:
					this.m_thread.Resume();
					return;
			}                
			
		}
		public void chkTreeSumTable()
		{
			string strMDBPathAndFile="";
			string strConn="";
			
			ado_data_access p_ado = new ado_data_access();

           
            int intIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(listViewEx1, "Sum Tree Yields, Volume, And Value For A Stand And Treatment");
            
           
			
			strMDBPathAndFile = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
			strConn=p_ado.getMDBConnString(strMDBPathAndFile,"admin","");
			if (p_ado.TableExist(strConn,"tree_vol_val_sum_by_rx"))
			{
				if (p_ado.getRecordCount(strConn,"select COUNT(*) from tree_vol_val_sum_by_rx","tree_vol_val_sum_by_rx") == 0)
				{
                   
				}
				else
				{
                   
                  
				}
			}
			else
			{
                
			}

		}
		

		/// <summary>
		/// Check to see if  fastest travel times 
		/// from plot to wood processing site exist in the plot table
		/// </summary>
		public void chkPlotTableForTravelTimes()
		{
			string strConn="";
            
            int intIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(listViewEx1, "Get Least Expensive Route From Stand To Wood Processing Facility");
			ado_data_access p_ado = new ado_data_access();

			string strTable = "";
			string strMDBFile = "";
			this.ReferenceCoreScenarioForm.uc_datasource1.getScenarioDataSourceMDBAndTableName("PLOT",ref strMDBFile,ref strTable);
			strConn=p_ado.getMDBConnString(strMDBFile,"admin","");
			if (p_ado.getRecordCount(strConn,"select COUNT(*) from " + strTable + " p WHERE p.merch_haul_cost_id IS NOT NULL AND p.chip_haul_cost_id IS NOT NULL",strTable) == 0)

			{
               
			}
			else
			{
                

			}
			p_ado = null;

		}
        public static void UpdateThermPercent()
        {
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep++;
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent(
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic,
                FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps,
                FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps,
                FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep);
        }

		/// <summary>
		/// validate each component required for running core analysis
		/// </summary>
		private void val_CoreRunData()
		{
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 10;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;

			this.m_intError=0;

            if (this.m_intError == 0)
            {
                this.m_intError = ReferenceCoreScenarioForm.uc_scenario_owner_groups1.ValInput();
               
            }
           
                


            if (this.m_intError == 0)
            {
                UpdateThermPercent();
                this.m_intError = ReferenceCoreScenarioForm.uc_scenario_costs1.val_costs();
            }


            if (this.m_intError == 0)
            {
                UpdateThermPercent();
                this.m_intError = ReferenceCoreScenarioForm.uc_scenario_psite1.val_psites();
            }


            if (this.m_intError == 0)
            {
                UpdateThermPercent();
                this.m_intError = ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.val_processorscenario();
            }
            

			if (this.m_intError==0)  
			{
                UpdateThermPercent();
				this.m_intError= ReferenceCoreScenarioForm.uc_scenario_filter1.Val_PlotFilter(ReferenceCoreScenarioForm.uc_scenario_filter1.txtCurrentSQL.Text.Trim());
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_filter1.m_strError,"FIA Biosum");

			}
            
			if (this.m_intError==0)  
			{
                UpdateThermPercent();
				this.m_intError = ReferenceCoreScenarioForm.uc_scenario_cond_filter1.Val_CondFilter(ReferenceCoreScenarioForm.uc_scenario_cond_filter1.txtCurrentSQL.Text.Trim());
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_cond_filter1.m_strError,"FIA Biosum");

			}
            
			if (this.m_intError==0) 
			{
                UpdateThermPercent();
				this.m_intError = ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.Audit(false);
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_strError,"FIA Biosum");
			}
           
			if (this.m_intError==0) 
			{
                UpdateThermPercent();
				ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.DisplayAuditMessage=false;
				ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.Audit();
				this.m_intError = ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.m_intError;
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.m_strError,"FIA Biosum");
			}

			if (this.m_intError==0) 
			{
                UpdateThermPercent();
				this.m_intError = ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.Audit(false);
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_strError,"FIA Biosum");
			}



            if (this.m_intError == 0)
            {

                /***************************************************************************
                     **make sure all the scenario datasource tables and files are available
                     **and ready for use
                     ***************************************************************************/
                UpdateThermPercent();
                if (ReferenceCoreScenarioForm.m_ldatasourcefirsttime == true)
                {
                    ReferenceCoreScenarioForm.uc_datasource1.populate_listview_grid();
                    ReferenceCoreScenarioForm.m_ldatasourcefirsttime = false;
                }
                this.m_intError = ReferenceCoreScenarioForm.uc_datasource1.val_datasources();
                if (this.m_intError == 0)
                {
                    UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");

                    FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = 1;
                    listViewEx1.Items[1].Selected = true;
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)listViewEx1.GetEmbeddedControl(1, 1);


                    ReferenceCoreScenarioForm.SaveRuleDefinitions();

                    UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
                }
            }
            else
            {
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
            }
				

		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
            
            groupBox1_Resize();
            
		}
        public void groupBox1_Resize()
        {
            this.listViewEx1.Width = this.groupBox1.Width - (int)(this.listViewEx1.Left * 2);
            this.lblMsg.Width = this.listViewEx1.Width;
            this.pnlFileSizeMonitor.Top = this.groupBox1.Height - this.groupBox1.Top - this.pnlFileSizeMonitor.Height - 2;
            this.lblMsg.Top = pnlFileSizeMonitor.Top - this.lblMsg.Height - 2;
            this.listViewEx1.Height = this.lblMsg.Top - this.listViewEx1.Top - 5;

            if (uc_filesize_monitor1.lblMaxSize.Left + uc_filesize_monitor1.lblMaxSize.Width > uc_filesize_monitor1.Width)
            {
                for (; ; )
                {
                    uc_filesize_monitor1.Width = uc_filesize_monitor1.Width + 1;
                    if (uc_filesize_monitor1.lblMaxSize.Left + uc_filesize_monitor1.lblMaxSize.Width < uc_filesize_monitor1.Width)
                        break;

                }
            }
            if (uc_filesize_monitor2.lblMaxSize.Left + uc_filesize_monitor2.lblMaxSize.Width > uc_filesize_monitor2.Width)
            {
                for (; ; )
                {
                    uc_filesize_monitor2.Width = uc_filesize_monitor2.Width + 1;
                    if (uc_filesize_monitor2.lblMaxSize.Left + uc_filesize_monitor2.lblMaxSize.Width < uc_filesize_monitor2.Width)
                        break;

                }
            }
            if (uc_filesize_monitor3.lblMaxSize.Left + uc_filesize_monitor3.lblMaxSize.Width > uc_filesize_monitor3.Width)
            {
                for (; ; )
                {
                    uc_filesize_monitor3.Width = uc_filesize_monitor3.Width + 1;
                    if (uc_filesize_monitor3.lblMaxSize.Left + uc_filesize_monitor3.lblMaxSize.Width < uc_filesize_monitor3.Width)
                        break;

                }
            }

            if (uc_filesize_monitor4.lblMaxSize.Left + uc_filesize_monitor4.lblMaxSize.Width > uc_filesize_monitor4.Width)
            {
                for (; ; )
                {
                    uc_filesize_monitor4.Width = uc_filesize_monitor4.Width + 1;
                    if (uc_filesize_monitor4.lblMaxSize.Left + uc_filesize_monitor4.lblMaxSize.Width < uc_filesize_monitor4.Width)
                        break;

                }
            }






            this.uc_filesize_monitor2.Left = this.uc_filesize_monitor1.Left + uc_filesize_monitor2.Width + 2;
            this.uc_filesize_monitor3.Left = this.uc_filesize_monitor2.Left + uc_filesize_monitor3.Width + 2;
            this.uc_filesize_monitor4.Left = this.uc_filesize_monitor3.Left + uc_filesize_monitor4.Width + 2;
        }



        private void UpdateOptimizationGroupboxText()
        {
            
            FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.Variable_Collection oOptimizationVariableCollection = ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.m_oSavVariableCollection;

            for (int x = 0; x <= oOptimizationVariableCollection.Count - 1; x++)
            {
                if (oOptimizationVariableCollection.Item(x).bSelected)
                {
                   
                    break;
                }
            }
        }
        public void UpdateOptimizationVariableGroupboxText(string p_strOptimizationVariableName)
        {
            
        }
		public FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.Variables ReferenceFVSPrePostVariables
		{
			get {return _oFVSPrePostVariables;}
			set {_oFVSPrePostVariables=value;}
		}
		public FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.Variable_Collection ReferenceFVSPrePostOptimization
		{
			get {return _oFVSPrePostOptimization;}
			set {_oFVSPrePostOptimization=value;}
		}
		public FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection ReferenceFVSPrePostTieBreaker
		{
			get {return this._oFVSPrePostTieBreaker;}
			set {this._oFVSPrePostTieBreaker=value;}
		}
		


		public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }
	}
	/// <summary>
	/// main class used for running the core analysis scenario
	/// </summary>
	public class RunCore
	{
		FIA_Biosum_Manager.uc_core_scenario_run _uc_scenario_run;
		FIA_Biosum_Manager.frmCoreScenario _frmScenario;
		private int m_intError;
		private string m_strSQL;
		private string m_strConn;
		public string m_strTempMDBFile;
		public ado_data_access m_ado;
        private dao_data_access m_oDao;
		public System.Data.OleDb.OleDbConnection m_TempMDBFileConn;
		public string m_strSystemResultsDbPathAndFile="";
        public string m_strFVSPreValidComboDbPathAndFile = "";
        public string m_strFVSPostValidComboDbPathAndFile = "";
        
		private env m_oEnv;
		private utils m_oUtils;
		public string m_strPlotTable;
		public string m_strRxTable;
        public string m_strRxPackageTable;
		public string m_strTravelTimeTable;
		public string m_strCondTable;
		public string m_strFFETable;
		public string m_strHvstCostsTable;
		public string m_strPSiteTable;
		public string m_strTreeVolValBySpcDiamGroupsTable;
		public string m_strTreeVolValSumTable = "tree_vol_val_sum_by_rx";
		public string m_strUserDefinedPlotSQL;
	    public string m_strUserDefinedCondSQL;
		
		private string m_strLine;
		private FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.VariableItem m_oOptimizationVariable = new FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.VariableItem();
		private string m_strOptimizationTableName="";
		private string m_strOptimizationSourceTableName="";
		private string m_strOptimizationTableNameSql="";
		private string m_strOptimizationColumnNameSql="";
		private string m_strOptimizationSourceColumnName="";
		private string m_strOptimizationAggregateSql="";
		private string m_strOptimizationAggregateColumnName="";
        private ProcessorScenarioItem m_oProcessorScenarioItem = null;
        private RxTools m_oRxTools = new RxTools();
        private RxPackageItem_Collection m_oRxPackageItem_Collection = new RxPackageItem_Collection();
		private FIA_Biosum_Manager.macrosubst m_oVarSub = new macrosubst();

        public static bool g_bCoreRun = false;
        public static ProgressBarBasic.ProgressBarBasic g_oCurrentProgressBarBasic = null;
        public static int g_intCurrentProgressBarBasicMinimumSteps = 1;
        public static int g_intCurrentProgressBarBasicMaximumSteps = -1;
        public static int g_intCurrentProgressBarBasicCurrentStep = -1;
        public static int g_intCurrentListViewItem = 0;
        System.Windows.Forms.CheckBox oCheckBox = null;
        int intListViewIndex = -1;
        private string m_strDebugFile = "";
        string[] m_strVariantArray = null;

        

		
		

		public RunCore(FIA_Biosum_Manager.uc_core_scenario_run p_form)
		{
			
			this.m_intError=0;
			_uc_scenario_run = p_form;
			try
			{
                m_strDebugFile = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runlog.txt";


                if (frmMain.g_bDebug)
                {
                    if (System.IO.File.Exists(m_strDebugFile)) System.IO.File.Delete(m_strDebugFile);
                    this.m_strLine = "START: Core Analysis Run Log " + System.DateTime.Now.ToString();
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_strLine + "\r\n\r\n");
                    if (frmMain.g_intDebugLevel > 1)
                    {
                        
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Project: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtProjectId.Text + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Project Directory: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Scenario Directory: " + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------------------------------------------\r\n");
                    }
                    
				}
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.ToString());
				this.m_intError=-1;
				return;
			}
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Create A Temporary MDB File With Links To All The Core Tables And Scenario Result Tables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "------------------------------------------------------------------------------------------------\r\n");
            }
			try
			{
                //
                //INITIALIZE AND LOAD VARIABLES
                //
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic =(ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, 2);
                FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = 2;
                FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 21;
                FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps=1;
                FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep=1;

                frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, 2);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, 2, "Selected", true);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, 2, "focused", true);

				/**************************************************************************
				 **first lets create a temp mdb with links to all the scenario core 
				 **and result tables
				 **************************************************************************/
				this.m_oUtils=new utils();
				this.m_oEnv = new env();

                //get the selected processor scenario item
                if (this.ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem != null)
                    this.m_oProcessorScenarioItem = this.ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem;

                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

                //load the treatment packages
                this.m_oRxTools.LoadAllRxPackageItems(m_oRxPackageItem_Collection);

                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                
				this.m_oVarSub.ReferenceSQLMacroSubstitutionVariableCollection = 
					frmMain.g_oSQLMacroSubstitutionVariable_Collection;

				this.m_strSystemResultsDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"mdb");
				this.CopyScenarioResultsTable(this.m_strSystemResultsDbPathAndFile,ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb");


                this.m_strFVSPreValidComboDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "mdb");
                this.CopyScenarioResultsTable(this.m_strFVSPreValidComboDbPathAndFile, ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo.mdb");

                this.m_strFVSPostValidComboDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "mdb");
                this.CopyScenarioResultsTable(this.m_strFVSPostValidComboDbPathAndFile, ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo.mdb");


                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

				this.m_strTempMDBFile = 
					ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.CreateMDBAndScenarioTableDataSourceLinks(this.m_oEnv.strTempDir);

                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

				this.m_strUserDefinedPlotSQL= 
					this.m_oVarSub.SQLTranslateVariableSubstitution(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_filter1.txtCurrentSQL.Text.Trim());

				this.m_strUserDefinedCondSQL= 
					this.m_oVarSub.SQLTranslateVariableSubstitution(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_cond_filter1.txtCurrentSQL.Text.Trim());

            
				if (this.m_strTempMDBFile.Trim().Length == 0)
				{
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                    if (frmMain.g_bDebug)
            			frmMain.g_oUtils.WriteText(m_strDebugFile,"!!RunCore: Error Creating MDB File Containing Links To All The Tables!!\r\n");
					MessageBox.Show("RunCore: Error Creating MDB File Containing Links To All The Core Tables");
				}
				else
				{

                    if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)_uc_scenario_run.uc_filesize_monitor1, "Visible", false) == false)
                    {
                        _uc_scenario_run.uc_filesize_monitor1.BeginMonitoringFile(
                            m_strTempMDBFile, 2000000000, "2GB");
                        _uc_scenario_run.uc_filesize_monitor1.Information = "Work table containing table links";

                        _uc_scenario_run.uc_filesize_monitor2.BeginMonitoringFile(
                            m_strSystemResultsDbPathAndFile, 2000000000, "2GB");
                        _uc_scenario_run.uc_filesize_monitor2.Information = "Scenario results DB file";

                    }

                    m_oDao = new dao_data_access();
                    //compact the scenario_results.mdb file

                    CompactMDB(m_strSystemResultsDbPathAndFile, null);

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile,"--Delete Scenario Result Table Records--\r\n");
					this.DeleteScenarioResultRecords();

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

                     if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile,"--Initialize Air Destruction Tables--\r\n");
					InitializeAirDestructionTables();

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

                     if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile,"Links MDB File: " + this.m_strTempMDBFile + "\r\n");
					ReferenceUserControlScenarioRun.btnAccess.Enabled=false;

					getTableNames();

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

					//effective variable
					FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.Variables oFvsVar =
						ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar;

					/********************************************************************
					 **get optimization variable
					 ********************************************************************/
					FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.Variable_Collection oOptimizationVariableCollection = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.m_oSavVariableCollection;
						
					for (int x=0;x<=oOptimizationVariableCollection.Count-1;x++)
					{
						if (oOptimizationVariableCollection.Item(x).bSelected)
						{
							m_oOptimizationVariable.Copy(oOptimizationVariableCollection.Item(x),ref m_oOptimizationVariable);
							break;
						}
					}

					if (m_oOptimizationVariable.strMaxYN=="Y")
					{
						this.m_strOptimizationTableName = "Max";
						this.m_strOptimizationAggregateSql="MAX";
						this.m_strOptimizationAggregateColumnName="max_optimization_value";
					}
					else
					{
						this.m_strOptimizationTableName = "Min";
						this.m_strOptimizationAggregateSql="MIN";
						this.m_strOptimizationAggregateColumnName="min_optimization_value";
					}

					if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
					{
						this.m_strOptimizationTableName = this.m_strOptimizationTableName + "NR";
						this.m_strOptimizationSourceTableName="product_yields_net_rev_costs_summary_by_rx";
						this.m_strOptimizationTableNameSql = "cycle1_effective_product_yields_net_rev_costs_summary_by_rx";
						this.m_strOptimizationSourceColumnName="max_nr_dpa";
						this.m_strOptimizationColumnNameSql="post_variable_value";
							
					}
					else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
					{
						this.m_strOptimizationSourceTableName="product_yields_net_rev_costs_summary_by_rx";
						this.m_strOptimizationTableNameSql = "cycle1_effective_product_yields_net_rev_costs_summary_by_rx";
						this.m_strOptimizationColumnNameSql="post_variable_value";
						this.m_strOptimizationSourceColumnName= "merch_yield_cf";
						this.m_strOptimizationTableName = this.m_strOptimizationTableName + "MerchVol";
						if (this.m_oOptimizationVariable.bUseFilter) this.m_strOptimizationTableName = this.m_strOptimizationTableName + "NR";
					}
					else
					{
						
						string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strFVSVariableName,".");
						this.m_strOptimizationSourceTableName=strCol[0];
						this.m_strOptimizationSourceColumnName = strCol[1];
						this.m_strOptimizationTableNameSql="cycle1_effective_optimization_treatments";
						if (this.m_oOptimizationVariable.strValueSource=="POST")
						{
							this.m_strOptimizationColumnNameSql="post_variable_value";
						}
						else if (this.m_oOptimizationVariable.strValueSource=="POST-PRE")
						{
							this.m_strOptimizationColumnNameSql="change_value";
						}


						this.m_strOptimizationTableName = this.m_strOptimizationTableName + strCol[1].Trim();
						if (this.m_oOptimizationVariable.bUseFilter) this.m_strOptimizationTableName = this.m_strOptimizationTableName + "NR";
					}
					//revenue filter
					if (this.m_oOptimizationVariable.bUseFilter)
					{
						switch (this.m_oOptimizationVariable.strFilterOperator)
						{
							case "<":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "lt";
								break;
							case ">":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "gt";
								break;
							case "<=":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "lte";
								break;
							case ">=":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "gte";
								break;
							case "<>":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "ne";
								break;
							case "=":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "e";
								break;
						}
                        if (m_oOptimizationVariable.dblFilterValue>=0)
						    this.m_strOptimizationTableName = this.m_strOptimizationTableName + "P" + Convert.ToString(this.m_oOptimizationVariable.dblFilterValue).Trim();
                        else
                            this.m_strOptimizationTableName = this.m_strOptimizationTableName + "N" + Convert.ToString(this.m_oOptimizationVariable.dblFilterValue*-1).Trim();
					}
					else
					{
						if (oFvsVar.m_bOverallEffectiveNetRevEnabled)
						{
							switch (oFvsVar.m_strOverallEffectiveNetRevOperator.Trim())
							{
								case "<":
									this.m_strOptimizationTableName = this.m_strOptimizationTableName + "lt";
									break;
								case ">":
									this.m_strOptimizationTableName = this.m_strOptimizationTableName + "gt";
									break;
								case "<=":
									this.m_strOptimizationTableName = this.m_strOptimizationTableName + "lte";
									break;
								case ">=":
									this.m_strOptimizationTableName = this.m_strOptimizationTableName + "gte";
									break;
								case "<>":
									this.m_strOptimizationTableName = this.m_strOptimizationTableName + "ne";
									break;
								case "=":
									this.m_strOptimizationTableName = this.m_strOptimizationTableName + "e";
									break;
							}
                            if (Convert.ToDouble(oFvsVar.m_strOverallEffectiveNetRevValue.Trim().Replace("$", "")) >= 0)
                                this.m_strOptimizationTableName = this.m_strOptimizationTableName + "P" + oFvsVar.m_strOverallEffectiveNetRevValue.Trim().Replace("$", "");
                            else
                            {
                                this.m_strOptimizationTableName = this.m_strOptimizationTableName + "N" + Convert.ToString(Convert.ToDouble(oFvsVar.m_strOverallEffectiveNetRevValue.Trim().Replace("$", "")) * -1);
                            }
						}
					}
                    m_strOptimizationTableName = "cycle1_" + m_strOptimizationTableName;

                    getVariants();
                    CreateAuditTables();
					CreateScenarioResultTables();
                    CreateValidComboTables();


                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

					//CREATE TABLE LINKS
					CreateScenarioResultTableLinks();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

                    //CREATE PROCESSOR SCENARIO TABLE LINKS
                    CreateProcessorScenarioResultTableLinks();

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

					this.CreateAuditTableLinks();
					if (this.m_intError != 0) return;

					this.CreateScenarioTableLinks();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }
                    this.CreateValidComboTableLinks();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }
                    

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

                    m_oRxTools.CreateTableLinksToFVSPrePostTables(m_strTempMDBFile);
                    m_intError = m_oRxTools.m_intError;
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

					//CREATE WORK TABLES
					CreateTableStructureOfHarvestCosts();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

					CreateTableStructureForIntensity();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
					
                    this.CreateTableStructureForScenarioProcessingSites();

                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

					this.CreateTableStructureForPSiteSum();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

					this.CreateTableStructureForOwnershipSum();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

					
					this.CreateTableStructureForHaulCosts();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

					this.CreateTableStructureForPlotCondAccessiblity();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }


                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

                    

					this.m_ado = new ado_data_access();

					this.m_strConn=m_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");


			
					this.m_TempMDBFileConn = new System.Data.OleDb.OleDbConnection();
					this.m_ado.OpenConnection(this.m_strConn,ref this.m_TempMDBFileConn);
					if (this.m_ado.m_intError==0)
					{
						/******************************************************************
						 **create table structure of user defined sql plot filter statement
						 ******************************************************************/
						CreateTableStructureOfUserDefinedPlotSQL();

                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

						/********************************************************************
						 **create table structure for condition table filters
						 ********************************************************************/
						this.CreateTableStructureForUserDefinedConditionTable();

                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

						/********************************************************************
						 **filter scenario selected processing sites
						 ********************************************************************/
						FilterPSites();

                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

						


						//close the connection if it is open
						if (this.m_TempMDBFileConn.State == System.Data.ConnectionState.Open)
						{
							this.m_TempMDBFileConn.Close();
						}
						if (this.m_intError==0)
						{
							this.m_strConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
							this.m_TempMDBFileConn = new System.Data.OleDb.OleDbConnection();
							this.m_ado.OpenConnection(this.m_strConn,ref this.m_TempMDBFileConn);
							this.m_intError=this.m_ado.m_intError;
						}

                        if (this.m_intError == 0)
                        {
                            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
                            
                        }
                        else
                        {
                            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                           
                        }


						/***********************************************************
						 **identify the plots that are accessible
						 ***********************************************************/
                        FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem + 1;
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);

						if (this.m_intError==0) 
						{
                           
                           
							this.PlotAccessible();
				
						}
						else
						{
                            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "NA");
						}

                       

                       

						/**************************************************************
						 **get the fastest travel time from plot to processing site
						 **************************************************************/
                        intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                           ReferenceUserControlScenarioRun.listViewEx1, "Get Least Expensive Route From Stand To Wood Processing Facility");
                      
                        FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);

						if (this.m_intError == 0)
						{

                           

							this.getHaulCosts();

						}
						else
						{
							if (ReferenceUserControlScenarioRun.m_bUserCancel==false)
							{
                                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "NA");
							}
						}
                    
						/***************************************************************************
						 **sum up tree volumes and values by plot+condition, treatment and species
						 ***************************************************************************/
                        intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                           ReferenceUserControlScenarioRun.listViewEx1, "Sum Tree Yields, Volume, And Value For A Stand And Treatment");
                        
                        FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);

						if (this.m_intError == 0) //&& (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)oCheckBox,"Checked",false)==true && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.sumTreeVolVal();

						}
						else
						{
							if (ReferenceUserControlScenarioRun.m_bUserCancel == false)
							{
                                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "NA");
							}
						}
                        /***************************************************************************
						 **sum up tree volumes and values by plot+rxpackage
						 ***************************************************************************/
                        intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                          ReferenceUserControlScenarioRun.listViewEx1, "Sum Tree Yields, Volume, And Value For A Stand And Treatment Package");
                       
                        FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);

                        if (this.m_intError == 0) 
                        {
                            this.sumTreeVolValByRxPackage();

                        }
                        else
                        {
                            if (ReferenceUserControlScenarioRun.m_bUserCancel == false)
                            {
                                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "NA");
                            }
                        }
						/***************************************************************************
						 **valid combos
						 ***************************************************************************/
                       
                        
						if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.validcombos();

						}
						/**********************************************************************
						 **wood product yields net revenue and costs summary by treatment table
						 **********************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.product_yields_net_rev_costs_summary_by_rx();

						}
                        /*******************************************************************************
						 **wood product yields net revenue and costs summary by treatment package table 
						 *******************************************************************************/
                        if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                        {
                            this.product_yields_net_rev_costs_summary_by_rxpackage();

                        }
                        //compact
                        if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
                            CompactMDB(m_strSystemResultsDbPathAndFile,m_TempMDBFileConn);
                        }
						/***************************************************************************
						 **effective treatments
						 ***************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Cycle1Effective();

						}
						/**************************************************************************
						 **optimization
						 **************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Cycle1Optimization();
						}
						/**************************************************************************
						 **tie breakers
						 **************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.tiebreaker();
						}
						
						/*********************************************************************
						 **find the best treatments for revenue, torch/crown index improvement,
						 **and merch removal
						 *********************************************************************/ 
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Cycle1_best_rx_summary();
						}
						/*******************************************************************************
						 **expand acreage for best treatments by plot 
						 *******************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Cycle1BestTreatmentsByPlot();
						}
						/**********************************************************************
						 **sum up the values by processing site
						 **********************************************************************/
						if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Cycle1SumPSite();

						}

						/**********************************************************************
						 **sum up the values by ownership
						 **********************************************************************/
						if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Cycle1SumOwnership();

						}
                        
                        /**********************************************************************
						 **all treatment by plot for cost, revenue and volume
						 **********************************************************************/
                        if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                        {
                            TreatmentAcreageExpansionByPlot();

                        }

                        if (m_TempMDBFileConn != null) m_ado.CloseConnection(m_TempMDBFileConn);
                        m_TempMDBFileConn = null;
                       
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
                            CompactMDB(m_strSystemResultsDbPathAndFile, null);
							System.DateTime oDate = System.DateTime.Now;
							string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
							string strFileDate = oDate.ToString(strDateFormat);
							strFileDate = strFileDate.Replace("/","_"); strFileDate=strFileDate.Replace(":","_");
							this.CreateHtml();
							this.CopyScenarioResultsTable(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb",this.m_strSystemResultsDbPathAndFile);
                            this.CopyScenarioResultsTable(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo_fvspre.mdb", this.m_strFVSPreValidComboDbPathAndFile);
                            this.CopyScenarioResultsTable(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo_fvspost.mdb", this.m_strFVSPostValidComboDbPathAndFile);
							this.m_strSystemResultsDbPathAndFile=ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
                            this.m_strFVSPreValidComboDbPathAndFile = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo_fvspre.mdb";
                            
							this.CopyScenarioResultsTable(
								ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results_" + this.m_strOptimizationTableName + "_" + strFileDate.Trim() + ".mdb",
								ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb");


						}

						

                        if (frmMain.g_bDebug)
                        {
                            this.m_strLine = "END: Core Analysis Run Log " + System.DateTime.Now.ToString();
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_strLine + "\r\n\r\n");
                        }

						
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",false);
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnCancel,"Text","Start");
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewScenarioTables,"Enabled",true);
						if (this.m_intError == 0) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnAccess,"Enabled",true);

                        

						if (this.m_intError != 0)
						{
							MessageBox.Show("Completed With Errors");
						}
						else
						{
                            if (ReferenceUserControlScenarioRun.m_bUserCancel == false)
                                MessageBox.Show("Successfully Completed");
                            else
                                MessageBox.Show("Process Cancelled");
						}
					}
				}
				
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewLog,"Enabled",true);

                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


                
                //audit
                intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                           ReferenceUserControlScenarioRun.listViewEx1, "Populate Valid Combination Audit Data");
                oCheckBox = (CheckBox)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(
                    0, intListViewIndex);
                FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
                if ((bool)frmMain.g_oDelegate.GetControlPropertyValue(oCheckBox, "Enabled", false) == true)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewAuditTables, "Enabled", true);
                }

              
			}
			catch (Exception err)
			{

                if (err.Message.Trim().ToUpper() != "THREAD WAS BEING ABORTED.")
                {
                    MessageBox.Show("Run Scenario " + err.Message, "FIA Biosum");
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                    frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.btnCancel, "Text", "Start");
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error!!\r\n");

                }
                if (m_TempMDBFileConn != null) m_ado.CloseConnection(m_TempMDBFileConn);
                m_TempMDBFileConn = null;
			}
            _uc_scenario_run.uc_filesize_monitor1.EndMonitoringFile();
            _uc_scenario_run.uc_filesize_monitor2.EndMonitoringFile();
            if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)_uc_scenario_run.uc_filesize_monitor3,"Visible",false))
                _uc_scenario_run.uc_filesize_monitor3.EndMonitoringFile();
            if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)_uc_scenario_run.uc_filesize_monitor4, "Visible", false))
                _uc_scenario_run.uc_filesize_monitor4.EndMonitoringFile();

			frmMain.g_oDelegate.m_oEventThreadStopped.Set();
			this.ReferenceUserControlScenarioRun.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
		}
			
		public RunCore()
		{

		}
		~RunCore()
		{
			
		}

        private void getVariants()
        {
            //
            //get all the variants
            //
            Datasource oDs = new Datasource(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(), this._uc_scenario_run.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim());
            m_strVariantArray = oDs.getVariants();
            
        }
        private void CreateAuditTables()
        {
            int x;
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();
            string strMDBPathAndFile = "";
            //
            //create an audit DB file for every variant
            //
            for (x = 0; x <= m_strVariantArray.Length - 1; x++)
            {
                strMDBPathAndFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\audit_" + m_strVariantArray[x].Trim() + ".mdb";
                if (System.IO.File.Exists(strMDBPathAndFile) == false)
                {
                    oDao.CreateMDB(strMDBPathAndFile);
                }
                oAdo.OpenConnection(oAdo.getMDBConnString(strMDBPathAndFile, "", ""));
                if (oAdo.TableExist(oAdo.m_OleDbConnection, frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName) == false)
                {
                    frmMain.g_oTables.m_oAudit.CreatePlotCondAuditTable(oAdo, oAdo.m_OleDbConnection, frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName);
                }
                if (oAdo.TableExist(oAdo.m_OleDbConnection, frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName) == false)
                {
                    frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTable(oAdo, oAdo.m_OleDbConnection, frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName);
                }
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
            }
            oDao = null;
            oAdo = null;

        }
		/// <summary>
		/// create a table structure that will hold
		/// the plot data that results when running the user 
		/// defined sql
		/// </summary>
		/// <param name="strUserDefinedSQL"></param>
		private void CreateTableStructureOfUserDefinedPlotSQL()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureOfUserDefinedPlotSQL\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			
			
			/********************************************************
			 **get the user defined PLOT filter sql
			 ********************************************************/
			this.m_strSQL = this.m_strUserDefinedPlotSQL;
			/****************************************************************
			 **get the table structure that results from executing the sql
			 ****************************************************************/
			System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,this.m_strSQL);

			/*****************************************************************
			 **create the table structure in the scenario_results.mdb file
			 **and give it the name of userdefinedplotfilter
			 *****************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Create userdefinedplotfilter Table Schema From User Defined Plot Filter SQL\r\n");
			dao_data_access p_dao = new dao_data_access();
			p_dao.CreateMDBTableFromDataSetTable(this.m_strSystemResultsDbPathAndFile,"userdefinedplotfilter",p_dt,true);
			p_dt.Dispose();
			this.m_ado.m_OleDbDataReader.Close();
			if (p_dao.m_intError!=0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Creating Table Schema!!\r\n");
				this.m_intError=p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the user defined plot filter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Create link in " + this.m_strTempMDBFile + "\r\n");

			p_dao.CreateTableLink(this.m_strTempMDBFile,"userdefinedplotfilter",m_strSystemResultsDbPathAndFile,"userdefinedplotfilter",true);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,m_strSystemResultsDbPathAndFile + "\tuserdefinedplotfilter\r\n");

			/***********************************************************************
			 **make a copy of the userdefinedplot filter table and give it the
			 **name ruledefinitionsplotfilter. This will apply the owngrpcd
			 **filters and any other future filters to the userdefinedplotfilter table.
			 ***********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Delete table ruledefinitionsplotfilter\r\n");
			p_dao.DeleteTableFromMDB(this.m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter");
			if (p_dao.m_intError !=0)
			{
                if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Deleting alluserdefinedplotfilter Table!!\r\n");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Copy table structure userdefinedplotfilter to ruledefinitionsplotfilter\r\n");
			p_dao.MoveTableToMDB(this.m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter",this.m_strSystemResultsDbPathAndFile,"userdefinedplotfilter",false);
			if (p_dao.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Creating ruledefinitionsplotfilter Table!!\r\n");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the ruledefinitionsplotfilter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionsplotfilter",m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter",true);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Create link in " + this.m_strTempMDBFile + "\r\n");

                frmMain.g_oUtils.WriteText(m_strDebugFile,m_strSystemResultsDbPathAndFile + "\truledefinitionsplotfilter\r\n");
            }


			p_dao=null;




		}
		/// <summary>
		/// create a table structure that will hold
		/// the plot data that results when running the user 
		/// defined sql
		/// </summary>
		/// <param name="strUserDefinedSQL"></param>
		private void CreateTableStructureOfUserDefinedSQLOld(string strUserDefinedSQL)
		{
			
			
			/*******************************************************************
			 ** get scenario_results.mdb path
			 *******************************************************************/
			string strMDBPathAndFile = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
			/********************************************************
			 **get the user defined PLOT filter sql
			 ********************************************************/
			this.m_strSQL = this.m_strUserDefinedPlotSQL;
			/****************************************************************
			 **get the table structure that results from executing the sql
			 ****************************************************************/
			System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,this.m_strSQL);

			/*****************************************************************
			 **create the table structure in the scenario_results.mdb file
			 **and give it the name of userdefinedplotfilter
			 *****************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Create userdefinedplotfilter Table Schema From User Defined Plot Filter SQL\r\n");
			dao_data_access p_dao = new dao_data_access();
			p_dao.CreateMDBTableFromDataSetTable(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","userdefinedplotfilter",p_dt,true);
			p_dt.Dispose();
			this.m_ado.m_OleDbDataReader.Close();
			if (p_dao.m_intError!=0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Creating Table Schema!!\r\n");
				this.m_intError=p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the user defined plot filter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Create link in " + this.m_strTempMDBFile + "\r\n");
			p_dao.CreateTableLink(this.m_strTempMDBFile,"userdefinedplotfilter",strMDBPathAndFile,"userdefinedplotfilter",true);
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,strMDBPathAndFile + "\tuserdefinedplotfilter\r\n");

			/***********************************************************************
			 **make a copy of the userdefinedplot filter table and give it the
			 **name ruledefinitionsplotfilter. This will apply the owngrpcd
			 **filters and any other future filters to the userdefinedplotfilter table.
			 ***********************************************************************/
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Delete table ruledefinitionsplotfilter\r\n");
			//first delete the table if it exists already
			p_dao.DeleteTableFromMDB(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","ruledefinitionsplotfilter");
			if (p_dao.m_intError !=0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Deleting alluserdefinedplotfilter Table!!\r\n");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Copy table structure userdefinedplotfilter to ruledefinitionsplotfilter\r\n");
			p_dao.MoveTableToMDB(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","ruledefinitionsplotfilter",ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","userdefinedplotfilter",false);
			if (p_dao.m_intError != 0)
			{
                 if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Creating ruledefinitionsplotfilter Table!!\r\n");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the ruledefinitionsplotfilter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionsplotfilter",strMDBPathAndFile,"ruledefinitionsplotfilter",true);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Create link in " + this.m_strTempMDBFile + "\r\n");

                frmMain.g_oUtils.WriteText(m_strDebugFile,strMDBPathAndFile + "\truledefinitionsplotfilter\r\n");
            }


			p_dao=null;




		}
        private void InitializeAirDestructionTables()
        {
           


        }
		
		/// <summary>
		/// create links to the tables located in the scenario_results.mdb file
		/// </summary>
		private void CreateScenarioResultTables()
		{

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateScenarioResultTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string[] strTableNames;
			strTableNames = new string[1];
			ado_data_access oAdo = new ado_data_access();
			

		
			
			oAdo.OpenConnection(oAdo.getMDBConnString(m_strSystemResultsDbPathAndFile,"",""));

            strTableNames = oAdo.getTableNames(oAdo.m_OleDbConnection);
            for (int x = 0; x <= strTableNames.Length - 1; x++)
            {
                if (strTableNames[x] != null &&
                    strTableNames[x].Trim().Length > 0)
                {
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTableNames[x]);
                }
            }

			

			frmMain.g_oTables.m_oCoreScenarioResults.CreateValidComboTable(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsValidCombosTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateValidComboFVSPrePostTable(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsValidCombosFVSPrePostTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateTieBreakerTable(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsTieBreakerTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationTable(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsCycle1OptimizationTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateEffectiveTable(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsCycle1EffectiveTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryCycle1WithIntensityTable(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsCycle1BestRxSummaryWithIntensityTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryCycle1Table(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsCycle1BestRxSummaryTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryCycle1Table(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsCycle1BestRxSummaryTableName + "_air_dest");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryCycle1Table(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsCycle1BestRxSummaryTableName + "_before_tiebreaks");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateTreeVolValSumTable(oAdo,oAdo.m_OleDbConnection,Tables.CoreScenarioResults.DefaultScenarioResultsTreeVolValSumTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateProductYieldsTable(oAdo,oAdo.m_OleDbConnection,"product_yields_net_rev_costs_summary_by_rx");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationPlotTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_stands");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationPSiteTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_psites_sum");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationOwnershipTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_own_sum");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationPlotTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_stands_air_dest");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationOwnershipTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_own_air_dest");

            frmMain.g_oTables.m_oCoreScenarioResults.CreateTreeVolValSumByRxPackageTable(oAdo, oAdo.m_OleDbConnection, Tables.CoreScenarioResults.DefaultScenarioResultsTreeVolValSumByRxPackageTableName);
            frmMain.g_oTables.m_oCoreScenarioResults.CreateProductYieldsByRxPackageTable(oAdo, oAdo.m_OleDbConnection, "product_yields_net_rev_costs_summary_by_rxpackage");
            frmMain.g_oTables.m_oCoreScenarioResults.CreatePlotRxCostsRevenuesVolumesTable(oAdo, oAdo.m_OleDbConnection, Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName);
            frmMain.g_oTables.m_oCoreScenarioResults.CreatePlotRxPackageCostsRevenuesVolumesSumTable(oAdo, oAdo.m_OleDbConnection, Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxPackageCostRevenueVolumesSumTableName);

			oAdo.CloseConnection(oAdo.m_OleDbConnection);

		}
		/// <summary>
		/// Copy the scenario results db file from the scenario?\db directory to the temp directory
		/// where the temp directory version is used during a single core analysis run. Once
		/// the run successfully completes it is copied back to the scenario?\db directory.
		/// </summary>
		private void CopyScenarioResultsTable(string p_strDestPathAndDbFileName,string p_strSourcePathAndDbFileName)
		{

			dao_data_access oDao = new dao_data_access();


			if (System.IO.File.Exists(p_strSourcePathAndDbFileName))
			{
				System.IO.File.Copy(p_strSourcePathAndDbFileName,p_strDestPathAndDbFileName,true);	
			}
			else
			{
					oDao.CreateMDB(p_strDestPathAndDbFileName);
			}
			oDao.m_DaoWorkspace.Close();
			oDao=null;
		}

		/// <summary>
		/// create links to the tables located in the scenario_results.mdb file
		/// </summary>
		private void CreateScenarioResultTableLinks()
		{
            
			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			p_dao.CreateTableLinks(this.m_strTempMDBFile,this.m_strSystemResultsDbPathAndFile);
			if (p_dao.m_intError==0)
			{

				int intCount = p_dao.getTableNames(m_strSystemResultsDbPathAndFile,ref strTableNames);
				if (p_dao.m_intError==0)
				{

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        if (intCount > 0)
                        {
                            for (int x = 0; x <= intCount - 1; x++)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                     frmMain.g_oUtils.WriteText(m_strDebugFile,
                                         "scenario_results\t" + strTableNames[x] + "\r\n");
                            }
                        }
                    }
				}
				else
				{
					this.m_intError=p_dao.m_intError;
				}
			}
			else
			{
                if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, p_dao.m_strError + "\r\n");
				this.m_intError = p_dao.m_intError;
			}

		
		}
        /// <summary>
        /// create links to the tables located in the validcombo.mdb file
        /// </summary>
        private void CreateValidComboTables()
        {

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateValidComboTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string[] strTableNames;
            strTableNames = new string[1];
            ado_data_access oAdo = new ado_data_access();



            //
            //FVS PRE VALID COMBO TABLE
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(this.m_strFVSPreValidComboDbPathAndFile, "", ""));
            strTableNames = oAdo.getTableNames(oAdo.m_OleDbConnection);
            for (int x = 0; x <= strTableNames.Length - 1; x++)
            {
                if (strTableNames[x] != null &&
                    strTableNames[x].Trim().Length > 0)
                {
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTableNames[x]);
                }
            }



            frmMain.g_oTables.m_oCoreScenarioResults.CreateValidComboFVSPreTable(oAdo, oAdo.m_OleDbConnection, Tables.CoreScenarioResults.DefaultScenarioResultsValidCombosFVSPreTableName);

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            //
            //FVS POST VALID COMBO TABLE
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(this.m_strFVSPostValidComboDbPathAndFile, "", ""));
            strTableNames = oAdo.getTableNames(oAdo.m_OleDbConnection);
            for (int x = 0; x <= strTableNames.Length - 1; x++)
            {
                if (strTableNames[x] != null &&
                    strTableNames[x].Trim().Length > 0)
                {
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTableNames[x]);
                }
            }
            frmMain.g_oTables.m_oCoreScenarioResults.CreateValidComboFVSPostTable(oAdo, oAdo.m_OleDbConnection, Tables.CoreScenarioResults.DefaultScenarioResultsValidCombosFVSPostTableName);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

        }

        /// <summary>
        /// create links to the tables located in the scenario_results.mdb file
        /// </summary>
        private void CreateValidComboTableLinks()
        {

            string[] strTableNames;
            strTableNames = new string[1];
            dao_data_access p_dao = new dao_data_access();
            p_dao.CreateTableLinks(this.m_strTempMDBFile, this.m_strFVSPreValidComboDbPathAndFile);
            if (p_dao.m_intError == 0)
            {

                int intCount = p_dao.getTableNames(m_strFVSPreValidComboDbPathAndFile, ref strTableNames);
                if (p_dao.m_intError == 0)
                {

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        if (intCount > 0)
                        {
                            for (int x = 0; x <= intCount - 1; x++)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile,
                                        "validcombo\t" + strTableNames[x] + "\r\n");
                            }
                        }
                    }
                }
                else
                {
                    this.m_intError = p_dao.m_intError;
                }
            }
            else
            {
                if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, p_dao.m_strError + "\r\n");
                this.m_intError = p_dao.m_intError;
            }
            if (p_dao.m_intError == 0)
            {
                p_dao.CreateTableLinks(this.m_strTempMDBFile, this.m_strFVSPostValidComboDbPathAndFile);
                if (p_dao.m_intError == 0)
                {

                    int intCount = p_dao.getTableNames(m_strFVSPostValidComboDbPathAndFile, ref strTableNames);
                    if (p_dao.m_intError == 0)
                    {

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        {
                            if (intCount > 0)
                            {
                                for (int x = 0; x <= intCount - 1; x++)
                                {
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                        frmMain.g_oUtils.WriteText(m_strDebugFile,
                                            "validcombo\t" + strTableNames[x] + "\r\n");
                                }
                            }
                        }
                    }
                    else
                    {
                        this.m_intError = p_dao.m_intError;
                    }
                }
            }


        }
        private void CreateProcessorScenarioResultTableLinks()
        {
            dao_data_access oDao = new dao_data_access();
            oDao.CreateTableLink(this.m_strTempMDBFile,
                                 Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                                 this.m_oProcessorScenarioItem.DbPath + "\\" +
                                 Tables.ProcessorScenarioRun.DefaultHarvestCostsTableDbFile,
                                 Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
            if (oDao.m_intError == 0)
            {
                oDao.CreateTableLink(this.m_strTempMDBFile,
                                 Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName,
                                 this.m_oProcessorScenarioItem.DbPath + "\\" +
                                 Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsDbFile,
                                 Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName);
            }

            if (oDao.m_intError == 0)
            {
                m_strTreeVolValBySpcDiamGroupsTable = Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;
                m_strHvstCostsTable = Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName;
            }
            oDao.m_DaoWorkspace.Close();
            oDao = null;



        }
		/// <summary>
		/// create links to the tables located in the scenario_results.mdb file
		/// </summary>
		private void CreateScenarioResultTableLinksOld()
		{
			

			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();

		
			string strMDBPathAndFile = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";


			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					for (int x=0; x <= intCount-1;x++)
					{


						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{

							this.m_intError = p_dao.m_intError;
							break;
						}
					}

				}
			}
			else
			{
				this.m_intError=p_dao.m_intError;
			}
			p_dao = null;
		}

		/// <summary>
		/// create links to the audit tables
		/// </summary>
		private void CreateAuditTableLinks()
		{

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateAuditTableLinks\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string[] strTableNames;
            int x,y;
			strTableNames = new string[1];
            string strMDBPathAndFile = "";
			dao_data_access oDao = new dao_data_access();
			
           
			
            for (x = 0; x <= m_strVariantArray.Length - 1; x++)
            {
                strMDBPathAndFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\audit_" + m_strVariantArray[x].Trim() + ".mdb";
                if (System.IO.File.Exists(strMDBPathAndFile))
                {
                    int intCount = oDao.getTableNames(strMDBPathAndFile, ref strTableNames);
                    if (oDao.m_intError == 0)
                    {
                        if (intCount > 0)
                        {
                            for (y = 0; y <= intCount - 1; y++)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, strMDBPathAndFile + "\t" + strTableNames[y] + "\r\n");
                                oDao.CreateTableLink(this.m_strTempMDBFile, strTableNames[y].Trim() + "_" + m_strVariantArray[x].Trim(), strMDBPathAndFile, strTableNames[y]);
                                if (oDao.m_intError != 0)
                                {
                                    if (frmMain.g_bDebug)
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error Creating Table Link!!!");
                                    this.m_intError = oDao.m_intError;
                                    break;
                                }
                            }

                        }
                    }
                    else
                    {
                        this.m_intError = oDao.m_intError;
                    }
                }
               
            }

			oDao = null;
			
		}
			
		/// <summary>
		/// create links to the scenario tables that contain the 
		/// rule definitions defined by the user
		/// </summary>
		private void CreateScenarioTableLinks()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateScenarioTableLinks\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			string strConn="";
			
			string strMDBPathAndFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb";

			//if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
			strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
				strMDBPathAndFile + ";User Id=admin;Password=;";
			//else
			//	strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";

			
			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames,false);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					for (int x=0; x <= intCount-1;x++)
					{
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, strMDBPathAndFile + "\t" + strTableNames[x] + "\r\n");
						
						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
                            if (frmMain.g_bDebug)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error Creating Table Link!!!");
							this.m_intError = p_dao.m_intError;
							break;
						}
					}

				}
			}
			else
			{
				this.m_intError=p_dao.m_intError;
			}
			p_dao = null;
			
		}

		

		/// <summary>
		/// get the names of the core tables
		/// </summary>
		private void getTableNames()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n\r\nGet Core Table Names\r\n---------------------\r\n");
            }
			/**************************************************************
			 **get the plot table name
			 **************************************************************/
			this.m_strPlotTable=ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("PLOT");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Plot:" + this.m_strPlotTable + "\r\n");


			/**************************************************************
			 **get the treatment prescriptions table
			 **************************************************************/
			this.m_strRxTable = 
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("TREATMENT PRESCRIPTIONS");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Treatment:" + this.m_strRxTable + "\r\n");

            /**************************************************************
			 **get the treatment package table
			 **************************************************************/
            this.m_strRxPackageTable =
                ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("TREATMENT PACKAGES");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Treatment Package:" + this.m_strRxPackageTable + "\r\n");


			/**************************************************************
			 **get the travel time table name
			 **************************************************************/
			this.m_strTravelTimeTable = 
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("TRAVEL TIMES");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Travel Time:" + m_strTravelTimeTable + "\r\n");

			
			/**************************************************************
			 **get the cond table name
			 **************************************************************/
			this.m_strCondTable=ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("CONDITION");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Condition:" + m_strCondTable + "\r\n");

			this.m_strPSiteTable = "scenario_psites_work_table";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Processing Sites:" + m_strPSiteTable + "\r\n");

			this.m_strTreeVolValSumTable = "tree_vol_val_sum_by_rx";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Tree Sum Volume And Value:" + m_strTreeVolValSumTable + "\r\n");

		}
		
		private void FilterPSites()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//FilterPSitess\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

			this.m_strSQL="DELETE FROM scenario_psites_work_table";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_strSQL = "INSERT INTO scenario_psites_work_table (psite_id,trancd,biocd) " + 
				"SELECT psite_id,trancd,biocd " + 
				"FROM scenario_psites " + 
				"WHERE TRIM(scenario_id)='" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + "' AND " + 
				"selected_yn='Y';";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


		}
		/// <summary>
		/// set the plot_accessible_yn by evaluating the field values in
		/// gis_protected_area_yn, gis_moved_from_protected_to_unprotected_yn,
		/// and gis_roadless_yn
		/// </summary>
		private void PlotAccessible()
		{
			/*********************************************************************
			 **set the plot_accessible_yn flag to 'Y' or 'N' based on whether
			 **the plot is in a protected area, roadless area,
			 **has been moded from protected to an unprotected area
			 **when moving the plot to the nearest road, or 
			 **every condition on the plot is unavailable
			 *********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//PlotAccessible\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps =9;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;


			
           
			/********************************************************************
			 **update cond_too_far_steep field to Y if slope is 
			 **<= 40% and the yarding distance >= to maximum yarding distance 
			 **for a slope <= 40%
			 ********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
				"INNER JOIN " + this.m_strPlotTable + " p " +
				"ON c.biosum_plot_id = p.biosum_plot_id  " + 
				"SET c.cond_too_far_steep_yn = 'Y' " + 
				"WHERE c.slope IS NOT NULL AND " + 
				"c.slope <  " +  m_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent + " AND " + 
				"p.gis_yard_dist >= " + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strNonSteepYardingDistance.Trim() + ";" ;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/********************************************************************
			 **update cond_too_far_steep field to Y if slope is 
			 **> 40% and the yarding distance >= to maximum yarding distance 
			 **for a slope > 40%
			 ********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
				"INNER JOIN " + this.m_strPlotTable + " p " +
				"ON c.biosum_plot_id = p.biosum_plot_id  " + 
				"SET c.cond_too_far_steep_yn = 'Y' " + 
				"WHERE c.slope IS NOT NULL AND " + 
				"c.slope >= " + m_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent + " AND " + 
				"p.gis_yard_dist >= " + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strSteepYardingDistance.Trim() + ";" ;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();



			/*************************************************************
			 **set the remainder of the cond_too_far_steep_yn fields to N
			 *************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
				"SET c.cond_too_far_steep_yn = 'N' " + 
				"WHERE c.cond_too_far_steep_yn IS NULL OR (c.cond_too_far_steep_yn <> 'Y' AND " + 
				"c.cond_too_far_steep_yn <> 'N');";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/*************************************************************
			 **update the condition accessible flag
			 *************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable +  " SET cond_accessible_yn = " + 
				"IIF(cond_too_far_steep_yn='Y','N','Y');";



            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/************************************************************************
			 **set the all_cond_not_accessible_yn flag to Y if every condition on the
			 **plot is inaccessible
			 ************************************************************************/

			/********************************************************************
			 **get the current condition record counts for each plot
			 ********************************************************************/
			//insert the condition counts into the work table
			this.m_strSQL = "INSERT INTO plot_cond_accessible_work_table (biosum_plot_id, num_cond) " + 
				" SELECT biosum_plot_id , COUNT(biosum_plot_id) " + 
				" FROM " + this.m_strCondTable + 
				" GROUP BY biosum_plot_id;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			this.m_strSQL = "INSERT INTO plot_cond_accessible_work_table2 (biosum_plot_id, num_cond, num_cond_not_accessible) " + 
				"SELECT a.biosum_plot_id, a.num_cond,b.cond_not_accessible_count as num_cond_not_accessible FROM plot_cond_accessible_work_table a," + 
				"(SELECT biosum_plot_id, Count(*) AS cond_not_accessible_count FROM " + this.m_strCondTable + " " + 
				"WHERE cond_accessible_yn='N' GROUP BY biosum_plot_id) b " + 
				"WHERE a.biosum_plot_id = b.biosum_plot_id;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/**********************************************************************
			 **initialize the all_cond_not_accessible_yn flag to N
			 **********************************************************************/
			this.m_strSQL = this.m_oVarSub.SQLTranslateVariableSubstitution("UPDATE @@PlotTable@@ SET all_cond_not_accessible_yn = 'N'");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/**********************************************************************
			 **update the all_cond_not_accessible_yn flag if all conditions 
			 **are not accessible
			 **********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
				"INNER JOIN plot_cond_accessible_work_table2 c " + 
				"ON p.biosum_plot_id=c.biosum_plot_id " + 
				"SET p.all_cond_not_accessible_yn=" + 
				"IIF(c.num_cond=c.num_cond_not_accessible,'Y','N');";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/*********************************************************************
			 **update the plot accessible flag
			 *********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable +  " SET plot_accessible_yn = " + 
				"IIF(gis_protected_area_yn='Y' " + 
				" OR gis_roadless_yn='Y' " + 
				" OR all_cond_not_accessible_yn='Y','N','Y');";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();




            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;

			if (this.m_ado.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");

			}
			else
			{
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error Executing SQL!!\r\n");
				this.m_intError = this.m_ado.m_intError;

			}
           


		}
		/// <summary>
		/// populate the haul_costs table and plot table with 
		/// the cheapest route for hauling merch and chip
		/// </summary>
		private void getHaulCosts()
		{

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//getHaulCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string strTruckHaulCost;
			string strRailHaulCost;
			string strTransferMerchCost;
			string strTransferBioCost;

            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 27;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Processing Haul Costs");
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",true);
			frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Refresh");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nUpdate Plot And Haul Cost Tables With Merch And Chip Haul Costs\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile,"-------------------------------------------------------------\r\n");
            }

			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
            strTruckHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.RoadHaulCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			
            strRailHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.RailHaulCostDollarsPerGreenTonPerMile.Replace("$", "").ToString();
			strRailHaulCost = strRailHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
            strTransferMerchCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.RailMerchTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
            strTransferBioCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.RailChipTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferBioCost = strTransferBioCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferBioCost = "0.00";
			 



			/*******************************************************************************
			 **zap the haul_costs table
			 *******************************************************************************/
			this.m_strSQL = "DELETE FROM haul_costs;";
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                 frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ndelete records in haul_costs table\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     
			
			
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Null The Plot Table's Haul Cost Fields");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			/***************************************************************
			 **null the plots merch_haul_cost_psite, merch_haul_cost_id and 
			 **merch_haul_cpa_pt,chip_haul_cost_psite,chip_haul_cost_id 
			 **and chip_haul_cpa_pt  fields
			 ***************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " " + 
				"SET merch_haul_cost_id = null," + 
				"merch_haul_cost_psite = null, " +
				"merch_haul_cpa_pt = null ," + 
				"chip_haul_cost_id = null," + 
				"chip_haul_cost_psite = null," + 
				"chip_haul_cpa_pt = null;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nnull the plot table's haul cost fields\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;


            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();


			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     
			
			

			/*****************************************************************
			 **delete any records that may exist in the work tables
			 *****************************************************************/

			
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Delete Records In Work Tables");

            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			//all plots and road-accessible psite  work table
			this.m_strSQL = "delete from all_road_merch_haul_costs_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from all_road_chip_haul_costs_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//railheads to rail-accessible collector psite work tables
			this.m_strSQL = "delete from merch_rh_to_collector_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from chip_rh_to_collector_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//plots to railheads to rail-accessible collector psite work tables
			this.m_strSQL = "delete from merch_plot_to_rh_to_collector_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from chip_plot_to_rh_to_collector_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//cheapest merch road route work table
			this.m_strSQL = "delete from cheapest_road_merch_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			//cheapest chip road route work table
			this.m_strSQL = "delete from cheapest_road_chip_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			//cheapest merch rail route work table
			this.m_strSQL = "delete from cheapest_rail_merch_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			//cheapest chip rail route work table
			this.m_strSQL = "delete from cheapest_rail_chip_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			//combine cheapest road and rail into single table
			this.m_strSQL = "DELETE FROM combine_merch_rail_road_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "DELETE FROM combine_chip_rail_road_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//overall cheapest routes for merch and chip
			this.m_strSQL = "delete from cheapest_merch_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from cheapest_chip_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			//MERCH AND CHIP ROAD PROCESSING SITE HAUL COSTS
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Road Haul Costs For Merchantable Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n\r\n--Merchantable wood processing site haul costs--\r\n");
			/***************************************************************
			 **process the merch travel times first
			 ***************************************************************/
			/*****************************************************************
			 **Insert into a table all travel time records where the psite 
			 **has road only or road/rail access and processes 
			 **merch only or merch/chip
			 *****************************************************************/
			this.m_strSQL = "INSERT into all_road_merch_haul_costs_work_table " + 
				"SELECT t.biosum_plot_id, 0 AS railhead_id," + 
				"0 AS transfer_cost, s.psite_id," + 
				"(" + strTruckHaulCost.Trim() + " * t.travel_time) AS road_cost," + 
				"0 AS rail_cost, (transfer_cost+road_cost+rail_cost) AS total_haul_cost," + 
				"'M' as materialcd " +
				"FROM " + this.m_strTravelTimeTable + " t," + 
				this.m_strPSiteTable + " s " + 
				"WHERE t.psite_id=s.psite_id AND " + 
				"(s.trancd=1 OR s.trancd =3) AND " +
                "(s.biocd=3 OR s.biocd=1)  AND t.travel_time > 0;";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table all travel time records where psite has road access and processes merch\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     
              
			/**************************************************************************
			 **Find the cheapest plot to merch processing site road route.
			 **The first query (a) returns all rows with biosum_plot_id, road_cost,
			 **and materialcd . The first subquery (b) finds the minimum haul cost
			 **for a plot. The second subquery (c) finds the minimum haul cost for each
			 **plot,psite combination. The where clause returns the desired row.
			 **************************************************************************/
			this.m_strSQL = "INSERT INTO cheapest_road_merch_haul_costs_work_table " + 
				"SELECT b.biosum_plot_id,c.psite_id,null AS railhead_id," + 
				"0 AS transfer_cost, a.road_cost, 0 AS rail_cost," + 
				"b.min_cost AS total_haul_cost," + 
				"'M' AS materialcd " + 
				"FROM  all_road_merch_haul_costs_work_table  a," + 
				"(SELECT biosum_plot_id,MIN(total_haul_cost) AS min_cost " + 
				"FROM all_road_merch_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id)  b," + 
				"(SELECT biosum_plot_id,  psite_id ," + 
				"MIN(total_haul_cost) AS min_cost2 " + 
				"FROM all_road_merch_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id, psite_id) c " + 
				"WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.psite_id = c.psite_id AND " + 
				"b.min_cost = c.min_cost2;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table. Find the cheapest plot to merch processing site road route.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Road Haul Costs For Chip Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			/***********************************************************************
			 **Append to a table all travel time records where the psite 
			 **has road only or road/rail access and processes 
			 **chip only or merch/chip
			 ***********************************************************************/
			this.m_strSQL = "INSERT INTO all_road_chip_haul_costs_work_table " + 
				"SELECT t.biosum_plot_id, 0 AS railhead_id," + 
				"0 AS transfer_cost,s.psite_id," + 
				"(" +  strTruckHaulCost.Trim() + " * t.travel_time) AS road_cost," + 
				"0 AS rail_cost, " + 
				"(transfer_cost+road_cost+rail_cost) AS total_haul_cost," + 
				"'C' AS materialcd " +
                "FROM " + this.m_strTravelTimeTable + " t," + 
				this.m_strPSiteTable + " s " + 
				"WHERE t.psite_id=s.psite_id AND " + 
				"(s.trancd=1 OR s.trancd=3) AND " +
                "(s.biocd=3 OR s.biocd=2)  AND t.travel_time > 0;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table all travel time records where psite has road access and processes chips.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     


			/******************************************************************
			 **For each plot get the cheapest road route to a psite. 
			 ******************************************************************/
			this.m_strSQL = "INSERT INTO cheapest_road_chip_haul_costs_work_table " + 
				"SELECT b.biosum_plot_id, c.psite_id, null AS railhead_id," + 
				"0 AS transfer_cost,a.road_cost," + 
				"0 AS rail_cost, b.min_cost AS total_haul_cost," + 
				"'C' AS materialcd " + 
				"FROM all_road_chip_haul_costs_work_table  a," + 
				"(SELECT biosum_plot_id,MIN(total_haul_cost) AS min_cost " + 
				"FROM all_road_chip_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id)  b," + 
				"(SELECT biosum_plot_id,  psite_id ," + 
				"MIN(total_haul_cost) AS min_cost2 " + 
				"FROM all_road_chip_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id, psite_id)  c " + 
				"WHERE c.biosum_plot_id=b.biosum_plot_id AND " + 
				"a.biosum_plot_id=b.biosum_plot_id AND " + 
				"a.psite_id=c.psite_id AND b.min_cost=c.min_cost2;";



			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table. Find the cheapest plot to chip processing site road route.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			//MERCH AND CHIP RAIL PROCESSING SITE HAUL COSTS
			/*********************************************************
			 **Append to a table all travel time collector_id (psite)
			 **records where the psite has rail access
			 *********************************************************/
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Rail Haul Costs For Merchantable Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			this.m_strSQL = "INSERT INTO merch_rh_to_collector_haul_costs_work_table " + 
				"SELECT t.psite_id AS railhead_id," + 
				"t.collector_id AS psite_id," + 
				"(" + strTransferMerchCost.Trim() + " * t.travel_time)  AS transfer_cost," + 
				"0 AS road_cost," + 
				"((t.travel_time * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost," + 
				"0 AS total_haul_cost,  'M' AS materialcd " +
                "FROM " + this.m_strTravelTimeTable + " t  " + 
				"INNER JOIN  " + this.m_strPSiteTable + " s " + 
				"ON t.collector_id = s.psite_id " +
                "WHERE  s.trancd=3 And (s.biocd=3 Or s.biocd=1)  AND t.travel_time > 0 AND " + 
				"EXISTS (SELECT ss.psite_id " + 
				"FROM " + this.m_strPSiteTable + " ss " + 
				"WHERE t.psite_id=ss.psite_id AND ss.trancd=2 AND (ss.biocd=3 Or ss.biocd=1));";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table all travel time collector_id (psite) records where the psite has rail access and processes merch.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     


			/***************************************************************************
			 **Combine records from the travel time table and the 
			 **merch_rh_to_collector_haul_costs_work_table table by matching the 
			 **r.railhead_id with the travel time psite_id. By doing this,
			 **we can calculate the road_cost and get the total cost by summing 
			 **together the plot to railhead road cost, the transfer of material cost,
			 ** and the railhead to collector site rail cost.
			 ***************************************************************************/
			this.m_strSQL = "INSERT INTO merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SELECT  t.biosum_plot_id, r.railhead_id, r.psite_id," + 
				"r.transfer_cost," + 
				"(" + strTruckHaulCost.Trim() + " * t.travel_time) AS road_cost," + 
				"r.rail_cost, (r.transfer_cost + road_cost + r.rail_cost) AS total_haul_cost," +
				"'M' AS materialcd " + 
				"FROM  " + this.m_strTravelTimeTable + " t," + 
				"merch_rh_to_collector_haul_costs_work_table r " +
                "WHERE r.railhead_id = t.psite_id  AND t.travel_time > 0;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table travel time plot records and previous work rail/merch table results\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     


			this.m_strSQL = "UPDATE merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SET total_haul_cost = transfer_cost + road_cost + rail_cost;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nupdate merch by road and rail total haul cost\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     



             
			/*******************************************************************
			 **Find the cheapest plot to merch processing site rail route.
			 **The first query (a) returns all rows with biosum_plot_id,
			 **railhead_id, transfer_cost, road_cost. The first subquery (b)
			 **finds the minimum haul cost for a plot. The second subquery (c)
			 **finds the minimum haul cost for each plot,psite combination.
			 **The where clause returns the desired row.
			 *******************************************************************/
			this.m_strSQL = "INSERT INTO cheapest_rail_merch_haul_costs_work_table " + 
				"SELECT a.biosum_plot_id,c.psite_id, a.railhead_id," + 
				"a.transfer_cost, a.road_cost,a.rail_cost," + 
				"c.min_cost AS total_haul_cost,'M' as materialcd " +
				"FROM merch_plot_to_rh_to_collector_haul_costs_work_table a," + 
				"(SELECT biosum_plot_id, MIN(total_haul_cost) AS min_cost2 " + 
				"FROM merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id) b," + 
				"(SELECT biosum_plot_id, psite_id," + 
				"MIN(total_haul_cost) AS min_cost " + 
				"FROM merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id,psite_id) c " + 
				"WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.biosum_plot_id = c.biosum_plot_id AND " + 
				"a.psite_id = c.psite_id AND " + 
				"a.total_haul_cost = c.min_cost AND " + 
				"min_cost2 = min_cost;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Find the cheapest plot to merch processing site rail routes\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Rail Haul Costs For Chip Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			/***********************************************************************
			 **Append to a table all travel time collector_id (psite) records
			 **where the psite has rail access and processes chips only or 
			 **both merch/chips
			 ***********************************************************************/
			this.m_strSQL = "INSERT INTO chip_rh_to_collector_haul_costs_work_table " + 
				"SELECT  t.psite_id AS railhead_id,"  + 
				"t.collector_id AS psite_id," + 
				"(" + strTransferBioCost.Trim() + " * t.travel_time)  AS transfer_cost," + 
				"0 AS road_cost," + 
				"((t.travel_time * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost," + 
				"0 AS total_haul_cost,  'C' AS materialcd " +
                "FROM " + this.m_strTravelTimeTable + " t  " + 
				"INNER JOIN  " + this.m_strPSiteTable + " s " + 
				"ON t.collector_id = s.psite_id " + 
				"WHERE s.trancd=3 AND  " +
                "(s.biocd=3 OR s.biocd=2)  AND t.travel_time > 0 AND " + 
				"EXISTS (SELECT ss.psite_id " + 
				"FROM " + this.m_strPSiteTable + " ss " + 
				"WHERE t.psite_id=ss.psite_id AND ss.trancd=2 AND (ss.biocd=3 Or ss.biocd=2));";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table all travel time collector_id (psite) records where the psite has rail access and processes chips.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			/*************************************************************************
			 **Combine records from the travel time table and the 
			 **chip_rh_to_collector_haul_costs_work_table by matching the
			 **r.railhead_id with the travel time psite_id. By doing this,
			 **we can calculate the road_cost and get the total cost by summing 
			 **together the plot to railhead road cost, the transfer of material cost,
			 **and the railhead to collector site rail cost.
			 *************************************************************************/
			this.m_strSQL = "INSERT INTO chip_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SELECT  t.biosum_plot_id, r.railhead_id, r.psite_id," + 
				"r.transfer_cost," + 
				"(" + strTruckHaulCost.Trim() + " * t.travel_time) AS road_cost," + 
				"r.rail_cost, " + 
				"(r.transfer_cost + road_cost + r.rail_cost) AS total_haul_cost," + 
				"'C' AS materialcd " + 
				"FROM  " + this.m_strTravelTimeTable + " t," + 
				"chip_rh_to_collector_haul_costs_work_table r " +
                "WHERE r.railhead_id = t.psite_id  AND t.travel_time > 0;";




			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table travel time plot records and previous rail/chips work table results\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}    
 
			this.m_strSQL = "UPDATE chip_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SET total_haul_cost = transfer_cost + road_cost + rail_cost;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nupdate chips by road and rail total haul cost\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}    




			/************************************************************************
			 **Find the cheapest plot to Chips processing site rail route.
			 **The first query (a) returns all rows with biosum_plot_id, railhead_id,
			 **transfer_cost, road_cost. The first subquery (b) finds the minimum
			 **haul cost for a plot. The second subquery (c) finds the minimum haul
			 **cost for each plot,psite combination. The where clause returns the
			 **desired row.
			 *************************************************************************/ 
			this.m_strSQL = "INSERT INTO cheapest_rail_chip_haul_costs_work_table " + 
				"SELECT a.biosum_plot_id,b.psite_id, a.railhead_id," + 
				"a.transfer_cost, a.road_cost, a.rail_cost," + 
				"b.min_cost AS total_haul_cost,'C' AS materialcd " + 
				"FROM chip_plot_to_rh_to_collector_haul_costs_work_table a," + 
				"(SELECT biosum_plot_id, " + 
				"MIN(total_haul_cost) AS min_cost2 " + 
				"FROM chip_plot_to_rh_to_collector_haul_costs_work_table " +
				"GROUP BY biosum_plot_id) c, " +
				"(SELECT biosum_plot_id, psite_id," + 
				"MIN(total_haul_cost) AS min_cost " + 
				"FROM chip_plot_to_rh_to_collector_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id,psite_id) b " + 
				"WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.psite_id = b.psite_id AND  " + 
				"a.total_haul_cost = b.min_cost AND " + 
				"min_cost2 = min_cost;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Find the cheapest plot to chip processing site rail routes\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

		    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Combine Road And Rail Haul Costs For Merchantable Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");


			/**************************************************************
			 **combine the cheapest road and rail total cost for each plot
			 **to a merch psite
			 **After the insert there should be two records for each
			 **plot - one with cheapest haul cost by road and another
			 **with cheapest haul cost by rail
			 ***************************************************************/
			this.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_road_merch_haul_costs_work_table;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Cheapest road route to merch psite\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			this.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_rail_merch_haul_costs_work_table;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Cheapest rail route to merch psite\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   
  

			/***************************************************
			 **Get the overall cheapest merch route
			 ***************************************************/
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Get Overall Least Expensive Merch Route");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
			this.m_strSQL = "INSERT INTO cheapest_merch_haul_costs_work_table " + 
				"SELECT a.biosum_plot_id,b.psite_id, a.railhead_id," + 
				"a.transfer_cost, a.road_cost,  a.rail_cost," + 
				"b.min_cost AS total_haul_cost,'M' AS materialcd " + 
				"FROM combine_merch_rail_road_haul_costs_work_table a," + 
				"(SELECT biosum_plot_id, MIN(total_haul_cost) AS min_cost2 " + 
				"FROM combine_merch_rail_road_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id) c, " + 
				"(SELECT biosum_plot_id, psite_id," + 
				"MIN(total_haul_cost) AS min_cost " + 
				"FROM combine_merch_rail_road_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id,psite_id) b " + 
				"WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.psite_id = b.psite_id AND " + 
				"a.total_haul_cost = b.min_cost AND " + 
				"min_cost2 = min_cost;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Get the overall cheapest merch route\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Combine Road And Rail Haul Costs For Chip Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			/**************************************************************
			 **combine the cheapest road and rail total cost for each plot
			 **to a chips psite
			 **After the insert there should be two records for each
			 **plot - one with cheapest haul cost by road and another
			 **with cheapest haul cost by rail
			 ***************************************************************/
			this.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_road_chip_haul_costs_work_table;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Cheapest road route to chip psite\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			this.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_rail_chip_haul_costs_work_table;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Cheapest rail route to chip psite\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   
  
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Get Overall Least Expensive Chip Route");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");


			/******************************************
			 **Get the overall cheapest chip route
			 ******************************************/
			this.m_strSQL = "INSERT INTO cheapest_chip_haul_costs_work_table " + 
				"SELECT a.biosum_plot_id,b.psite_id, a.railhead_id," + 
				"a.transfer_cost, a.road_cost,  a.rail_cost," + 
				"b.min_cost AS total_haul_cost,'C' AS materialcd " + 
				"FROM combine_chip_rail_road_haul_costs_work_table a, " + 
				"(SELECT biosum_plot_id,MIN(total_haul_cost) AS min_cost2 " + 
				"FROM combine_chip_rail_road_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id) c, " + 
				"(SELECT biosum_plot_id, psite_id," + 
				"MIN(total_haul_cost) AS min_cost " + 
				"FROM combine_chip_rail_road_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id,psite_id) b " + 
				"WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.psite_id = b.psite_id AND " + 
				"a.total_haul_cost = b.min_cost AND " + 
				"min_cost2 = min_cost;";



			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Get the overall cheapest chip route\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   




			//INSERT INTO HAUL_COSTS TABLE
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Inserting Results Into Haul Costs Table");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
			this.m_strSQL = "INSERT INTO haul_costs " + 
				"SELECT * FROM cheapest_merch_haul_costs_work_table;";
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into haul_costs table cheapest merch route for each plot\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   


			this.m_strSQL = "INSERT INTO haul_costs " + 
				"SELECT * FROM cheapest_chip_haul_costs_work_table;";
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into haul_costs table cheapest chip route for each plot\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   

			//UPDATE PLOT TABLE

			/**************************************************
			 **Update cheapest merch route fields
			 **************************************************/
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Updating Plot Table");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
				"INNER JOIN haul_costs h " + 
				"ON p.biosum_plot_id = h.biosum_plot_id " + 
				"SET p.merch_haul_cost_id = h.haul_cost_id," + 
				"p.merch_haul_cost_psite = h.psite_id," + 
				"p.merch_haul_cpa_pt=h.total_haul_cost " + 
				"WHERE h.materialcd='M';";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nUpdate plot merch haul cost fields\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			} 
  

			/*****************************************
			 **Update  cheapest chip routes
			 *****************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
				"INNER JOIN haul_costs h " + 
				"ON p.biosum_plot_id = h.biosum_plot_id " + 
				"SET p.chip_haul_cost_id = h.haul_cost_id," + 
				"p.chip_haul_cost_psite = h.psite_id," + 
				"p.chip_haul_cpa_pt=h.total_haul_cost " + 
				"WHERE h.materialcd='C';";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nUpdate plot chip haul cost fields\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); 
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			} 
  

			/******************************************
			 **clean up work tables
			 ******************************************/

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Cleaning Up Haul Cost Work Tables...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCleaning up haul cost work tables\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); 
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			} 
			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
			}
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", false);
            
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
            
		}


		/// <summary>
		/// sum the tree_vol_val_by_species_diam_groups table values to tree_vol_val_sum_by_rx
		/// </summary>
		private void sumTreeVolVal()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//sumTreeVolVal\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 3;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
			
			/**************************************************************
			 **sum the tree_vol_val_by_species_diam_groups table to
			 **        tree_vol_val_sum_by_rx
			 **************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nSum Tree Volumes and Values By Treatment\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "----------------------------------------\r\n");
            }
			

			this.m_strSQL = "delete from tree_vol_val_sum_by_rx";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				this.m_intError = this.m_ado.m_intError;
				return;
			}
			this.m_strSQL = "INSERT INTO tree_vol_val_sum_by_rx " + 
				"(biosum_cond_id,rxpackage,rx,rxcycle,chip_vol_cf," + 
				"chip_wt_gt,chip_val_dpa,merch_vol_cf," + 
				"merch_wt_gt,merch_val_dpa) ";
			this.m_strSQL += "SELECT biosum_cond_id, " + 
				"rxpackage,rx,rxcycle," +
				"Sum(chip_vol_cf) AS chip_vol_cf," + 
				"Sum(chip_wt_gt) AS chip_wt_gt," + 
				"Sum(chip_val_dpa) AS chip_val_dpa," + 
				"Sum(merch_vol_cf) AS merch_vol_cf," +
				"Sum(merch_wt_gt) AS merch_wt_gt," + 
				"Sum(merch_val_dpa) AS merch_val_dpa ";

			this.m_strSQL += " FROM " + this.m_strTreeVolValBySpcDiamGroupsTable.Trim();
			this.m_strSQL += " GROUP BY biosum_cond_id,rxpackage,rx,rxcycle";
			this.m_strSQL += " ORDER BY biosum_cond_id,rxpackage,rx,rxcycle;";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into tree_vol_val_sum_by_rx table tree volume and value sums\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); 

			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");

				this.m_intError = this.m_ado.m_intError;

				return;
			}

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");

			}
		}
        /// <summary>
        /// sum the tree_vol_val_by_species_diam_groups table values to tree_vol_val_sum_by_rxpackage
        /// </summary>
        private void sumTreeVolValByRxPackage()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//sumTreeVolValByRxPackage\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 3;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;

            /**************************************************************
             **sum the tree_vol_val_by_species_diam_groups table to
             **        tree_vol_val_sum_by_rxpackage
             **************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nSum Tree Volumes and Values By Treatment Package\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "------------------------------------------------\r\n");
            }
           


            this.m_strSQL = "delete from tree_vol_val_sum_by_rxpackage";
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }
            this.m_strSQL = "INSERT INTO tree_vol_val_sum_by_rxpackage " +
                "(biosum_cond_id,rxpackage,chip_vol_cf," +
                "chip_wt_gt,chip_val_dpa,merch_vol_cf," +
                "merch_wt_gt,merch_val_dpa) ";
            this.m_strSQL += "SELECT biosum_cond_id, " +
                "rxpackage," +
                "Sum(chip_vol_cf) AS chip_vol_cf," +
                "Sum(chip_wt_gt) AS chip_wt_gt," +
                "Sum(chip_val_dpa) AS chip_val_dpa," +
                "Sum(merch_vol_cf) AS merch_vol_cf," +
                "Sum(merch_wt_gt) AS merch_wt_gt," +
                "Sum(merch_val_dpa) AS merch_val_dpa ";

            this.m_strSQL += " FROM " + this.m_strTreeVolValBySpcDiamGroupsTable.Trim();
            this.m_strSQL += " GROUP BY biosum_cond_id,rxpackage";
            this.m_strSQL += " ORDER BY biosum_cond_id,rxpackage;";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into tree_vol_val_sum_by_rxpackage table tree volume and value sums\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); 

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }
            
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");

            }
        }
		private void getHaulCost(string p_strBiosumPlotId, int p_intPSiteId)
		{
			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
            string strTruckHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.RoadHaulCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			
            string strRailHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.RailHaulCostDollarsPerGreenTonPerMile.Replace("$", "").ToString();
			strRailHaulCost = strRailHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
            string strTransferMerchCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.RailMerchTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
            string strTransferBioCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.RailChipTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferBioCost = strTransferBioCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferBioCost = "0.00";
			



		}

		/// <summary>
		/// load the validcombos table with biosum_cond_id,rxpackage,rx and rxcycle values
		/// that exist in the user defined plot filters, condition, ffe, travel times, and
		/// harvest cost, and tree volume/value tables
		/// </summary>
		private void    validcombos()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//validcombos\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string strRxList="";
			string strGrpCd="";
			int x=0,y=0;
			string strTable="";
            int intListViewIndex = -1;
            string strMDBPathAndFile = "";
            System.Windows.Forms.CheckBox oCheckBox = null;
            
			//int y=0;
			//int z=0;

            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                         ReferenceUserControlScenarioRun.listViewEx1, "Apply User Defined Filters And Get Valid Stand Combinations");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 20;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);



			/*****************************************************************
			 **delete audit tables
			 *****************************************************************/


            for (x = 0; x <= m_strVariantArray.Length - 1; x++)
            {
                this.m_strSQL = "delete from plot_audit_" + m_strVariantArray[x].Trim();
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
                if (this.m_ado.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = this.m_ado.m_intError;
                    return;
                }

            }

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            for (x = 0; x <= m_strVariantArray.Length - 1; x++)
            {
                this.m_strSQL = "delete from plot_cond_rx_audit_" + m_strVariantArray[x].Trim();
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                if (this.m_ado.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = this.m_ado.m_intError;
                    return;
                }
            }
			
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/**************************
			 **get the treatment list
			 **************************/
			this.m_strSQL = "SELECT rx FROM " + this.m_strRxTable + ";"; // WHERE trim(ucase(scenario_id)) = '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
			this.m_ado.SqlQueryReader(this.m_TempMDBFileConn,this.m_strSQL);
			if (!this.m_ado.m_OleDbDataReader.HasRows)
			{
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				this.m_intError = -1;
				MessageBox.Show("No Treatments Found In The Treatment Table");
				return;
			}
			while (this.m_ado.m_OleDbDataReader.Read())
			{
				strRxList+=this.m_ado.m_OleDbDataReader["rx"].ToString().Trim();
			}
			this.m_ado.m_OleDbDataReader.Close();

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();


             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute User Defined Plot SQL And Insert Resulting Records Into Table userdefinedplotfilter\r\n");
			this.m_strSQL = "INSERT INTO userdefinedplotfilter " + this.m_strUserDefinedPlotSQL;

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL: " + this.m_strSQL + "\r\n");

			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--Execute User Defined Cond SQL And Insert Resulting Records Into Table userdefinedcondfilter--\r\n");
			this.m_strSQL = "INSERT INTO userdefinedcondfilter " + this.m_strUserDefinedCondSQL;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
			
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--Execute rule definition filters for the condition table. The filters include ownership and condition accessible--\r\n");
			this.m_strSQL = "INSERT INTO ruledefinitionscondfilter SELECT * FROM " + this.m_strCondTable + " c ";
			this.m_strSQL += " WHERE c.owngrpcd IN (";

			//usfs ownnership
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_owner_groups1.chkOwnGrp10.Checked==true)
			{
				strGrpCd = "10,1";
			}
			//other federal ownership
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_owner_groups1.chkOwnGrp20.Checked==true)
			{
				if (strGrpCd.Trim().Length == 0)
				{
					strGrpCd = "20,2";
				}
				else
				{
					strGrpCd += ",20,2";
				}

			}
			//state and local govt ownership
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_owner_groups1.chkOwnGrp30.Checked==true)
			{
				if (strGrpCd.Trim().Length == 0)
				{
					strGrpCd = "30,3";
				}
				else
				{
					strGrpCd += ",30,3";
				}

			}
			//private ownership
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_owner_groups1.chkOwnGrp40.Checked==true)
			{
				if (strGrpCd.Trim().Length == 0)
				{
					strGrpCd = "40,4";
				}
				else
				{
					strGrpCd += ",40,4";
				}

			}
			this.m_strSQL +=  strGrpCd + ") AND c.cond_accessible_yn = 'Y';";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();


			 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--Execute SQL that deletes from the condition rule definitions table (ruledefinitionscondfilter) those biosum_cond_id that are not found in the user defined condition SQL filter table (userdefinedcondfilter)--\r\n");

			this.m_strSQL = "DELETE FROM ruledefinitionscondfilter a " + 
				            "WHERE NOT EXISTS " + 
				                "(SELECT b.biosum_cond_id " + 
								 "FROM userdefinedcondfilter b " + 
				                 "WHERE a.biosum_cond_id =  b.biosum_cond_id)";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();


			 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--Execute SQL That Includes  Rule Definitions Defined By The User Into Table ruledefinitionsplotfilter--\r\n");
			this.m_strSQL = "INSERT INTO ruledefinitionsplotfilter SELECT DISTINCT userdefinedplotfilter.* from userdefinedplotfilter INNER JOIN ruledefinitionscondfilter ON userdefinedplotfilter.biosum_plot_id = ruledefinitionscondfilter.biosum_plot_id" ;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			this.m_strSQL = "DELETE FROM validcombos;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            _uc_scenario_run.uc_filesize_monitor3.BeginMonitoringFile(
                          m_strFVSPreValidComboDbPathAndFile, 2000000000, "2GB");
            _uc_scenario_run.uc_filesize_monitor3.Information = "Valid combinations for FVS Pre-Treatment records";

            _uc_scenario_run.uc_filesize_monitor4.BeginMonitoringFile(
                           m_strFVSPostValidComboDbPathAndFile, 2000000000, "2GB");
            _uc_scenario_run.uc_filesize_monitor4.Information = "Valid combinations for FVS Post-Treatment records";
           
            CompactMDB(m_strFVSPostValidComboDbPathAndFile, m_TempMDBFileConn);

			/**********************************************************************
			 **create valid combiniations of biosum_cond_id and treatment and 
			 **also user defined filters by owngrpcd
			 **********************************************************************/
			//insert all the possible valid POST plot+rxpackage+rxcycle records
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--insert all possible valid post plot+rxpackage+rxcycle records--\r\n");
            //cycle1
			m_strSQL = "INSERT INTO validcombos_fvspost " + 
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear1_rx AS rx,'1' AS rxcycle " + 
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " + 
                       "WHERE b.simyear1_rx IS NOT NULL AND LEN(TRIM(b.simyear1_rx)) > 0 AND b.simyear1_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            //cycle2
            m_strSQL = "INSERT INTO validcombos_fvspost " +
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear2_rx AS rx,'2' AS rxcycle " +
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
                       "WHERE b.simyear2_rx IS NOT NULL AND LEN(TRIM(b.simyear2_rx)) > 0 AND b.simyear2_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            //cycle3
            m_strSQL = "INSERT INTO validcombos_fvspost " +
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear3_rx AS rx,'3' AS rxcycle " +
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
                       "WHERE b.simyear3_rx IS NOT NULL AND LEN(TRIM(b.simyear3_rx)) > 0 AND b.simyear3_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            //cycle4
            m_strSQL = "INSERT INTO validcombos_fvspost " +
           "SELECT a.biosum_cond_id,b.rxpackage,b.simyear4_rx AS rx,'4' AS rxcycle " +
           "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
           "WHERE b.simyear4_rx IS NOT NULL AND LEN(TRIM(b.simyear4_rx)) > 0 AND b.simyear4_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            CompactMDB(m_strFVSPostValidComboDbPathAndFile, m_TempMDBFileConn);

            //insert all the possible valid pre plot+rxpackage+rxcycle records
            CompactMDB(m_strFVSPreValidComboDbPathAndFile, m_TempMDBFileConn);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert all possible valid PRE plot+rxpackage+rxcycle records--\r\n");
            //cycle1
            m_strSQL = "INSERT INTO validcombos_fvspre " +
           "SELECT a.biosum_cond_id,b.rxpackage,b.simyear1_rx AS rx,'1' AS rxcycle " +
           "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
           "WHERE b.simyear1_rx IS NOT NULL AND LEN(TRIM(b.simyear1_rx)) > 0 AND b.simyear1_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            //cycle2
            m_strSQL = "INSERT INTO validcombos_fvspre " +
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear2_rx AS rx,'2' AS rxcycle " +
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
                       "WHERE b.simyear2_rx IS NOT NULL AND LEN(TRIM(b.simyear2_rx)) > 0 AND b.simyear2_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            //cycle3
            m_strSQL = "INSERT INTO validcombos_fvspre " +
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear3_rx AS rx,'3' AS rxcycle " +
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
                       "WHERE b.simyear3_rx IS NOT NULL AND LEN(TRIM(b.simyear3_rx)) > 0 AND b.simyear3_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            //cycle4
            m_strSQL = "INSERT INTO validcombos_fvspre " +
           "SELECT a.biosum_cond_id,b.rxpackage,b.simyear4_rx AS rx,'4' AS rxcycle " +
           "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
           "WHERE b.simyear4_rx IS NOT NULL AND LEN(TRIM(b.simyear4_rx)) > 0 AND b.simyear4_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            CompactMDB(m_strFVSPreValidComboDbPathAndFile, m_TempMDBFileConn);
			

			string strWhere="";
			for (x=0;x<=FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strTable = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.TableName(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPostVarArray[x]);
				if (strTable.Trim().Length > 0)
				{
					m_strSQL = "UPDATE validcombos_fvspost a INNER JOIN " + strTable + " b ON a.biosum_cond_id = b.biosum_cond_id AND a.rxpackage = b.rxpackage AND a.rx = b.rx AND a.rxcycle = b.rxcycle SET variable" + Convert.ToString(x + 1).Trim() + "_yn='Y'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					m_strSQL = "UPDATE validcombos_fvspost SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE variable" + Convert.ToString(x + 1).Trim() + "_yn IS NULL OR LEN(TRIM(variable" + Convert.ToString(x + 1).Trim() + "_yn))=0";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					strWhere=strWhere + "b.variable" + Convert.ToString(x + 1).Trim() + "_yn <> 'N' AND ";

				}

				strTable = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.TableName(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPreVarArray[x]);
				if (strTable.Trim().Length > 0)
				{
                    m_strSQL = "UPDATE validcombos_fvspre a INNER JOIN " + strTable + " b ON a.biosum_cond_id = b.biosum_cond_id  AND a.rxpackage = b.rxpackage AND a.rx = b.rx AND a.rxcycle = b.rxcycle SET variable" + Convert.ToString(x + 1).Trim() + "_yn='Y'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					m_strSQL = "UPDATE validcombos_fvspre SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE variable" + Convert.ToString(x + 1).Trim() + "_yn IS NULL OR LEN(TRIM(variable" + Convert.ToString(x + 1).Trim() + "_yn))=0";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					strWhere=strWhere + "a.variable" + Convert.ToString(x + 1).Trim() + "_yn <> 'N' AND ";

				}
			}
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
			if (strWhere.Trim().Length > 0)
			{
				strWhere = "WHERE a.biosum_cond_id = b.biosum_cond_id AND a.rxpackage=b.rxpackage AND a.rx=b.rx AND a.rxcycle=b.rxcycle AND " +  strWhere.Substring(0,strWhere.Length - 5);
			}

			

			m_strSQL = "INSERT INTO validcombos_fvsprepost " + 
				"SELECT a.biosum_cond_id,a.rxpackage,a.rx,a.rxcycle " + 
				"FROM validcombos_fvspost a, validcombos_fvspre b " + 
				strWhere;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            

            this.m_strSQL = "INSERT INTO validcombos (biosum_cond_id,rxpackage,rx,rxcycle) SELECT DISTINCT ruledefinitionscondfilter.biosum_cond_id,validcombos_fvsprepost.rxpackage,validcombos_fvsprepost.rx,validcombos_fvsprepost.rxcycle " +
				"FROM (((ruledefinitionsplotfilter INNER JOIN ruledefinitionscondfilter ON ruledefinitionsplotfilter.biosum_plot_id = ruledefinitionscondfilter.biosum_plot_id) " + 
				" INNER JOIN validcombos_fvsprepost ON ruledefinitionscondfilter.biosum_cond_id = validcombos_fvsprepost.biosum_cond_id) " +
                " INNER JOIN " + this.m_strHvstCostsTable + " ON " + 
                                  "(validcombos_fvsprepost.rxpackage=" + this.m_strHvstCostsTable + ".rxpackage) AND " + 
                                  "(validcombos_fvsprepost.rx = " + this.m_strHvstCostsTable + ".rx) AND " + 
                                  "(validcombos_fvsprepost.rxcycle = " + this.m_strHvstCostsTable + ".rxcycle) AND " + 
                                  "(validcombos_fvsprepost.biosum_cond_id = " + this.m_strHvstCostsTable + ".biosum_cond_id)) " +
                " INNER JOIN " + this.m_strTreeVolValSumTable + " ON " + 
                                  "(" + this.m_strHvstCostsTable + ".biosum_cond_id = " + this.m_strTreeVolValSumTable + ".biosum_cond_id) AND " + 
                                  "(" + this.m_strHvstCostsTable + ".rxpackage = " + this.m_strTreeVolValSumTable + ".rxpackage) AND " + 
                                  "(" + this.m_strHvstCostsTable + ".rx = " + this.m_strTreeVolValSumTable + ".rx) AND " + 
                                  "(" + this.m_strHvstCostsTable + ".rxcycle = " + this.m_strTreeVolValSumTable + ".rxcycle) AND " + 
                                  "(validcombos_fvsprepost.biosum_cond_id = " + this.m_strTreeVolValSumTable + ".biosum_cond_id) AND " + 
                                  "(validcombos_fvsprepost.rxpackage = " + this.m_strTreeVolValSumTable + ".rxpackage)  AND " + 
                                  "(validcombos_fvsprepost.rx = " + this.m_strTreeVolValSumTable + ".rx) AND " + 
                                  "(validcombos_fvsprepost.rxcycle = " + this.m_strTreeVolValSumTable + ".rxcycle)";


			 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nCreate Valid Combinations\r\n");

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                 frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            
			
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)==true) return;

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");


            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                        ReferenceUserControlScenarioRun.listViewEx1, "Populate Valid Combination Audit Data");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 16 * m_strVariantArray.Length;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            

            oCheckBox = (CheckBox)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(
                           0, intListViewIndex);

            _uc_scenario_run.uc_filesize_monitor4.EndMonitoringFile();

            if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)oCheckBox, "Checked", false) == true)
			{
                frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Creating Audit Data");
                frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun, "Refresh");

                strMDBPathAndFile = "";
                //
                //create an audit DB file for every variant
                //
                for (x = 0; x <= m_strVariantArray.Length - 1; x++)
                {
                    
                    strMDBPathAndFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\audit_" + m_strVariantArray[x].Trim() + ".mdb";
                    _uc_scenario_run.uc_filesize_monitor4.BeginMonitoringFile(strMDBPathAndFile, 2000000000, "2GB");

                    CompactMDB(strMDBPathAndFile, m_TempMDBFileConn);

                    //BIOSUM_COND_ID RECORD AUDIT
                    /******************************************************************************************
                     **insert all the plots that are being processed into the plot audit table
                     ******************************************************************************************/
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n--plot_audit--\r\n\r\n");
                    this.m_strSQL = "INSERT INTO plot_audit_" + m_strVariantArray[x].Trim() + " (biosum_cond_id) SELECT ruledefinitionscondfilter.biosum_cond_id FROM ruledefinitionscondfilter INNER JOIN userdefinedplotfilter ON ruledefinitionscondfilter.biosum_plot_id = userdefinedplotfilter.biosum_plot_id WHERE userdefinedplotfilter.fvs_variant='" + m_strVariantArray[x].Trim() + "'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Insert All Plots From ruledefinitionscoldfilter table into plot_audit\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    if (this.m_ado.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                        this.m_intError = this.m_ado.m_intError;
                        return;
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();


                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;

                    /************************************************************************
                     **check to see if the plot record exists in the frcs harvest cost table
                     ************************************************************************/
                    m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                        "(SELECT TOP 1 a.biosum_cond_id " + 
                                        "FROM plot_audit_" + m_strVariantArray[x].Trim() + " AS a," + 
                                            m_strHvstCostsTable + " b " + 
                                        "WHERE a.biosum_cond_id = b.biosum_cond_id)";
                    if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL,"temp") > 0)
                    {
                        this.m_strSQL = "UPDATE plot_audit_" + m_strVariantArray[x].Trim() + " AS a " +
                                        "SET a.frcs_harvest_costs_yn = 'Y' " +
                                        "WHERE a.biosum_cond_id " +
                                        "IN (SELECT biosum_cond_id FROM " + this.m_strHvstCostsTable + ");";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSee if plot record exists in the FRCS harvest cost table\r\n");
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                        this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;

                    this.m_strSQL = "UPDATE plot_audit_" + m_strVariantArray[x].Trim() + " a " + 
                                    "SET a.frcs_harvest_costs_yn = 'N' " +
                                    "WHERE a.frcs_harvest_costs_yn IS NULL OR LEN(TRIM(a.frcs_harvest_costs_yn))=0;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet FRCS_harvest_costs_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;

                    /****************************************************************************
                     **check to see if the plot record exists in the validcombos_fvsprepost table
                     ****************************************************************************/
                     m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                        "(SELECT TOP 1 a.biosum_cond_id " + 
                                        "FROM plot_audit_" + m_strVariantArray[x].Trim() + " AS a," + 
                                              "validcombos_fvsprepost b " + 
                                        "WHERE a.biosum_cond_id = b.biosum_cond_id)";
                     if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL, "temp") > 0)
                     {
                         this.m_strSQL = "UPDATE plot_audit_" + m_strVariantArray[x].Trim() + " a " +
                                         "SET a.fvs_prepost_variables_yn = 'Y' " +
                                         "WHERE a.biosum_cond_id " +
                                         "IN (SELECT biosum_cond_id FROM validcombos_fvsprepost);";
                         if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                             frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=Y if plot record exists in validcombos_fvsprepost table\r\n");
                         if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                             frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                         this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                     }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;



                    this.m_strSQL = "UPDATE plot_audit_" + m_strVariantArray[x].Trim() + " a " + 
                                    "SET a.fvs_prepost_variables_yn = 'N' " +
                                    "WHERE a.fvs_prepost_variables_yn IS NULL OR " + 
                                          "a.fvs_prepost_variables_yn <> 'Y' ;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    /********************************************************************************************************
                     **check to see if the plot record exists in the processor tree volume and value tableharvest cost table
                     ********************************************************************************************************/
                      m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                        "(SELECT TOP 1 a.biosum_cond_id " + 
                                        "FROM plot_audit_" + m_strVariantArray[x].Trim() + " AS a," + 
                                              this.m_strTreeVolValSumTable + " b " + 
                                        "WHERE a.biosum_cond_id = b.biosum_cond_id)";
                      if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL, "temp") > 0)
                      {
                          this.m_strSQL = "UPDATE plot_audit_" + m_strVariantArray[x].Trim() + " a " +
                                          "SET a.processor_tree_vol_val_yn = 'Y' " +
                                          "WHERE a.biosum_cond_id " +
                                          "IN (SELECT biosum_cond_id FROM " + this.m_strTreeVolValSumTable + ");";
                          if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                              frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=Y if plot record exists in " + this.m_strTreeVolValSumTable + " table\r\n");
                          if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                              frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                          this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                      }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    this.m_strSQL = "UPDATE plot_audit_" + m_strVariantArray[x] + " a " + 
                                    "SET a.processor_tree_vol_val_yn = 'N' " +
                                    "WHERE a.processor_tree_vol_val_yn IS NULL OR  " + 
                                          "a.processor_tree_vol_val_yn<>'Y' ;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;



                    /**********************************************************************
                     **check to see if the plot record exists in the gis travel times table
                     **********************************************************************/
                     m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                        "(SELECT TOP 1 a.biosum_cond_id " + 
                                        "FROM plot_audit_" + m_strVariantArray[x].Trim() + " AS a," + 
                                             "ruledefinitionscondfilter b," + 
                                              m_strTravelTimeTable + " c " + 
                                        "WHERE a.biosum_cond_id = b.biosum_cond_id AND c.biosum_plot_id=b.biosum_plot_id)";
                     if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL, "temp") > 0)
                     {
                         this.m_strSQL = "UPDATE plot_audit_" + m_strVariantArray[x].Trim() + " a " +
                                         "SET a.gis_travel_times_yn = 'Y' " +
                                         "WHERE a.biosum_cond_id " +
                                         "IN (SELECT biosum_cond_id FROM ruledefinitionscondfilter " +
                                             "WHERE ruledefinitionscondfilter.biosum_plot_id " +
                                                    "IN (SELECT biosum_plot_id FROM " +
                                                         this.m_strTravelTimeTable + "));";
                         if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                             frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet gis_travel_times_yn=Y if plot record exists in " + this.m_strTravelTimeTable + " table\r\n");
                         if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                             frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                         this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                     }

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    this.m_strSQL = "UPDATE plot_audit_" + m_strVariantArray[x].Trim() + " a " + 
                                    "SET a.gis_travel_times_yn = 'N' " +
                                    "WHERE a.gis_travel_times_yn IS NULL OR  " + 
                                          "a.gis_travel_times_yn<>'Y' ;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet gis_travel_times_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    //BIOSUM_COND_ID + RX RECORD AUDIT
                    /**********************************************************************************
                    **Insert all the biosum_cond_id + rx combinations into the plot_cond_rx_audit table
                    ***********************************************************************************/
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n--plot_cond_rx_audit--\r\n");
                    //cycle1
                    this.m_strSQL = "INSERT INTO plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " " + 
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM plot_audit_" + m_strVariantArray[x].Trim() + " a, " +
                        "(SELECT DISTINCT rxpackage,simyear1_rx AS rx,'1' AS rxcycle " +
                         "FROM " + this.m_strRxPackageTable + " " +
                         "WHERE simyear1_rx IS NOT NULL AND LEN(TRIM(simyear1_rx)) > 0 AND simyear1_rx<>'000') b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    //cycle2
                    this.m_strSQL = "INSERT INTO plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " " + 
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM plot_audit_" + m_strVariantArray[x].Trim() + " a, " +
                        "(SELECT DISTINCT rxpackage,simyear2_rx AS rx,'2' AS rxcycle " +
                         "FROM " + this.m_strRxPackageTable + " " +
                         "WHERE simyear2_rx IS NOT NULL AND LEN(TRIM(simyear2_rx)) > 0 AND simyear2_rx<>'000') b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    //cycle3
                    this.m_strSQL = "INSERT INTO plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " " + 
                        "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM plot_audit_" + m_strVariantArray[x].Trim() + " a, " +
                        "(SELECT DISTINCT rxpackage,simyear3_rx AS rx,'3' AS rxcycle " +
                         "FROM " + this.m_strRxPackageTable + " " +
                         "WHERE simyear3_rx IS NOT NULL AND LEN(TRIM(simyear3_rx)) > 0 AND simyear3_rx<>'000') b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    //cycle4
                    this.m_strSQL = "INSERT INTO plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " " + 
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM plot_audit_" + m_strVariantArray[x].Trim() + " a, " +
                        "(SELECT DISTINCT rxpackage,simyear4_rx AS rx,'4' AS rxcycle " +
                         "FROM " + this.m_strRxPackageTable + " " +
                         "WHERE simyear4_rx IS NOT NULL AND LEN(TRIM(simyear4_rx)) > 0 AND simyear4_rx<>'000') b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);



                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    /*********************************************************************************
                     **check to see if the plot + rx record exists in the fvs prepost variables table
                     *********************************************************************************/
                    this.m_strSQL = "UPDATE plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " " + 
                        "SET fvs_prepost_variables_yn = 'Y' " +
                        "WHERE EXISTS (SELECT biosum_cond_id,rxpackage,rx,rxcycle " +
                        "FROM validcombos_fvsprepost " +
                        "WHERE plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + ".biosum_cond_id = " +
                        "validcombos_fvsprepost.biosum_cond_id AND " +
                        "plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + ".rxpackage = validcombos_fvsprepost.rxpackage AND " +
                        "plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + ".rx = validcombos_fvsprepost.rx AND " +
                        "plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + ".rxcycle=validcombos_fvsprepost.rxcycle);";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=Y if plot + rx + rxpackage + rxcycle record exists in validcombos_fvsprepost table\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    this.m_strSQL = "UPDATE plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " a " + 
                        "SET a.fvs_prepost_variables_yn = 'N' " +
                        "WHERE a.fvs_prepost_variables_yn IS NULL OR LEN(TRIM(a.fvs_prepost_variables_yn))=0;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    /****************************************************************************
                     **check to see if the plot + rx record exists in the frcs harves costs table
                     ****************************************************************************/
                    this.m_strSQL = "UPDATE plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " " + 
                        "SET frcs_harvest_costs_yn = 'Y' " +
                        "WHERE EXISTS (SELECT biosum_cond_id,rxpackage,rx,rxcycle " +
                        "FROM " + this.m_strHvstCostsTable + " " +
                        "WHERE plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + ".biosum_cond_id = " +
                        this.m_strHvstCostsTable.Trim() + ".biosum_cond_id AND " +
                        "plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + ".rxpackage = " + this.m_strHvstCostsTable.Trim() + ".rxpackage AND " +
                        "plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + ".rx = " + this.m_strHvstCostsTable.Trim() + ".rx AND " +
                        "plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + ".rxcycle = " + this.m_strHvstCostsTable.Trim() + ".rxcycle);";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet FRCS_harvest_costs_yn=Y if plot + rx + rxpackage + rxcycle record exists in " + m_strHvstCostsTable + " table\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    this.m_strSQL = "UPDATE plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " a " + 
                        "SET a.frcs_harvest_costs_yn = 'N' " +
                        "WHERE a.frcs_harvest_costs_yn IS NULL OR LEN(TRIM(a.frcs_harvest_costs_yn))=0 ;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet FRCS_harvest_costs_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    /*********************************************************************************
                     **check to see if the plot + rx record exists in the processor tree vol val table
                     *********************************************************************************/
                    m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                        "(SELECT TOP 1 a.biosum_cond_id " + 
                                        "FROM plot_cond_rx_audit_" + m_strVariantArray[x].Trim() + " a," + 
                                             m_strTreeVolValSumTable + " b " + 
                                        "WHERE a.biosum_cond_id = b.biosum_cond_id AND " + 
                                              "a.rxpackage = b.rxpackage AND " + 
                                              "a.rx=b.rx AND " + 
                                              "a.rxcycle=b.rxcycle)";
                    if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL, "temp") > 0)
                    {


                        this.m_strSQL = "UPDATE plot_cond_rx_audit_" + m_strVariantArray[x] + " a " +
                                       "INNER JOIN " + m_strTreeVolValSumTable + " b " +
                                       "ON  a.biosum_cond_id=b.biosum_cond_id AND " +
                                           "a.rxpackage=b.rxpackage AND " +
                                           "a.rx=b.rx AND " +
                                           "a.rxcycle=b.rxcycle " +
                                       "SET a.processor_tree_vol_val_yn = 'Y'";
                       
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=Y if plot + rx + rxpackage + rxcycle record exists in " + m_strTreeVolValSumTable + " table\r\n");
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                        this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    }
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    this.m_strSQL = "UPDATE plot_cond_rx_audit_" + m_strVariantArray[x] + " a " + 
                        "SET a.processor_tree_vol_val_yn = 'N' " +
                        "WHERE a.processor_tree_vol_val_yn IS NULL OR LEN(TRIM(a.processor_tree_vol_val_yn))=0;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic)) return;


                    

                    //compact
                    CompactMDB(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\audit_" + m_strVariantArray[x].Trim() + ".mdb", m_TempMDBFileConn);


                    

                }



               

                
				

				

			}
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", false);
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
            _uc_scenario_run.uc_filesize_monitor4.EndMonitoringFile();
            _uc_scenario_run.uc_filesize_monitor3.EndMonitoringFile();
			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
			}
			
	}

		/// <summary>
		/// evaluate the effectiveness of fvs treatment data 
		/// by loading the effective table with 
		/// results from user defined expressions 
		/// </summary>
		private void Cycle1Effective()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//sumTreeVolVal\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			int x,y;
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";
			string[] strEffectiveColumnArray;
			string[] strBetterIsNotNull= new string[uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string[] strWorseIsNotNull= new string[uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string[] strEffectiveIsNotNull= new string[uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string strOverallEffectiveIsNotNull = "";
			string strBetterSql="";
			string strWorseSql="";
			string strEffectiveSql="";
            int intListViewIndex = -1;

			string strVariableNumber="";
			FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.Variables oFvsVar =
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar;


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nCycle1 Effective Treatments\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------\r\n");
            }
           
			

            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                     ReferenceUserControlScenarioRun.listViewEx1, "Cycle 1: Identify Effective Treatments For Each Stand");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 8;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


			//get all the column names in the effective table
			strEffectiveColumnArray = m_ado.getFieldNamesArray(this.m_TempMDBFileConn,"select * from cycle1_effective");

			/********************************************
			 **delete all records in the effective table
			 ********************************************/
			this.m_strSQL = "delete from cycle1_effective";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			//insert the valid combos into the effective table
			m_strSQL = "INSERT INTO cycle1_effective (biosum_cond_id,rxpackage,rx,rxcycle) SELECT biosum_cond_id,rxpackage,rx,rxcycle FROM validcombos WHERE rxcycle='1'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			//insert net revenue per acre into the effective table
			m_strSQL = "UPDATE cycle1_effective e " + 
				       "INNER JOIN product_yields_net_rev_costs_summary_by_rx p " + 
				       "ON e.biosum_cond_id=p.biosum_cond_id AND " + 
                          "e.rxpackage=p.rxpackage AND " + 
				          "e.rx=p.rx AND " + 
                          "e.rxcycle=p.rxcycle " + 
			           "SET e.nr_dpa = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

        
			//populate the variable table.column name and its value to the effective table
			for (x=0;x<=uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				
				
				m_strSQL="";
				strPreTable="";
				strPreColumn="";
				strPostTable="";
				strPostColumn="";
				strVariableNumber = Convert.ToString(x+1).Trim();
				if (oFvsVar.TableName(oFvsVar.m_strPreVarArray[x].Trim().ToUpper()) != "NOT DEFINED")
				{
					
					strPreTable=oFvsVar.TableName(oFvsVar.m_strPreVarArray[x].Trim());
					strPreColumn=oFvsVar.ColumnName(oFvsVar.m_strPreVarArray[x].Trim());
				}
				if (oFvsVar.m_strPostVarArray[x].Trim().ToUpper() != "NOT DEFINED")
				{
					
					strPostTable=oFvsVar.TableName(oFvsVar.m_strPostVarArray[x].Trim());
					strPostColumn=oFvsVar.ColumnName(oFvsVar.m_strPostVarArray[x].Trim());
				}
				if (strPreTable.Trim().Length > 0)
				{
					m_strSQL = "e.pre_variable" + strVariableNumber + "_name='" + strPreTable + "." + strPreColumn + "',";
					m_strSQL = m_strSQL + "e.pre_variable" + strVariableNumber + "_value=pre." + strPreColumn;
				}
				else
				{
					m_strSQL = "e.pre_variable" + strVariableNumber + "_name=null,";
					m_strSQL = m_strSQL + "e.pre_variable" + strVariableNumber + "_value=null";
				}

				if (strPostTable.Trim().Length > 0)
				{
					m_strSQL = m_strSQL + ",e.post_variable" + strVariableNumber + "_name='" + strPostTable + "." + strPostColumn + "',";
					m_strSQL = m_strSQL + "e.post_variable" + strVariableNumber + "_value=post." + strPostColumn;
				}
				else
				{
					m_strSQL = m_strSQL + ",e.post_variable" + strVariableNumber + "_name=null,";
					m_strSQL = m_strSQL + "e.post_variable" + strVariableNumber + "_value=null";
				}
				if (strPreTable.Trim().Length > 0 && strPostTable.Trim().Length > 0)
				{
					m_strSQL = "UPDATE cycle1_effective e " + 
                               "INNER JOIN (" + strPostTable + " post " + 
                               "INNER JOIN " + 	strPreTable  + " pre " + 
                               "ON post.biosum_cond_id = pre.biosum_cond_id AND " + 
                                  "post.rxpackage=pre.rxpackage AND " + 
                                  "post.rx=pre.rx AND " + 
                                  "post.rxcycle=pre.rxcycle) "  + 
						       "ON e.biosum_cond_id=post.biosum_cond_id AND " + 
                                  "e.rxpackage=post.rxpackage AND " + 
                                  "e.rx=post.rx AND " + 
                                  "e.rxcycle=post.rxcycle " + 
						       "SET " + m_strSQL;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

				}
				

			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
           
			//populate the change column by subtracting pre value from post value
			m_strSQL="";
			for (x=0;x<=uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strVariableNumber = Convert.ToString(x+1).Trim();
				m_strSQL = m_strSQL + "variable" + strVariableNumber + "_change=IIF(pre_variable" + strVariableNumber + "_value IS NOT NULL AND post_variable" + strVariableNumber + "_value IS NOT NULL,post_variable" + strVariableNumber + "_value - pre_variable" + strVariableNumber + "_value,null),";
			}
			m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);

			m_strSQL = "UPDATE cycle1_effective SET " + m_strSQL;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            

			//see what variables are referenced in the sql expression and make sure they are not null
			strOverallEffectiveIsNotNull="";
			for (x=0;x<=uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strBetterIsNotNull[x]="";
				strWorseIsNotNull[x]="";
				strEffectiveIsNotNull[x]="";

				for (y=0;y<=strEffectiveColumnArray.Length - 1;y++)
				{
					if (oFvsVar.m_strBetterExpr[x].Trim().Length > 0)
					{
						if (oFvsVar.m_strBetterExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(),0) >= 0)
						{
							strBetterIsNotNull[x] = strBetterIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
						}
					}
					if (oFvsVar.m_strWorseExpr[x].Trim().Length > 0)
					{
						if (oFvsVar.m_strWorseExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(),0) >= 0)
						{
							strWorseIsNotNull[x] = strWorseIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
						}
					}
					if (oFvsVar.m_strEffectiveExpr[x].Trim().Length > 0)
					{
						if (oFvsVar.m_strEffectiveExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(),0) >= 0)
						{
							strEffectiveIsNotNull[x] = strEffectiveIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
						}
					}
				}

			}
			if (oFvsVar.m_strOverallEffectiveExpr.Trim().Length > 0)
			{
				for (y=0;y<=strEffectiveColumnArray.Length - 1;y++)
				{
					
					if (oFvsVar.m_strOverallEffectiveExpr.Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(),0) >= 0)
					{
						strOverallEffectiveIsNotNull = strOverallEffectiveIsNotNull + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
					}
					
				}
			}
			//remove the last AND
			for (x=0;x<=uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				if (strBetterIsNotNull[x].Trim().Length > 0)
				{
					strBetterIsNotNull[x] = strBetterIsNotNull[x].Substring(0,strBetterIsNotNull[x].Length - 5);
				}
				if (strWorseIsNotNull[x].Trim().Length > 0)
				{
					strWorseIsNotNull[x] = strWorseIsNotNull[x].Substring(0,strWorseIsNotNull[x].Length - 5);
				}
				if (strEffectiveIsNotNull[x].Trim().Length > 0)
				{
					strEffectiveIsNotNull[x] = strEffectiveIsNotNull[x].Substring(0,strEffectiveIsNotNull[x].Length - 5);
				}
			}
			if (strOverallEffectiveIsNotNull.Trim().Length > 0)
			{
				strOverallEffectiveIsNotNull = strOverallEffectiveIsNotNull.Substring(0,strOverallEffectiveIsNotNull.Length - 5);
			}
			//populate the better,worse,effective, and overall effective columns
			m_strSQL="";
			strBetterSql="";
			strWorseSql="";
			strEffectiveSql="";
			for (x=0;x<=uc_core_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strVariableNumber = Convert.ToString(x+1).Trim();
				if (oFvsVar.m_strBetterExpr[x].Trim().Length > 0)
				{
					strBetterSql = strBetterSql + "variable" + strVariableNumber + "_better_yn=IIF(" + strBetterIsNotNull[x].Trim() + ",IIF(" + oFvsVar.m_strBetterExpr[x].Trim() + ",'Y','N'),null),";
				}
				if (oFvsVar.m_strWorseExpr[x].Trim().Length > 0)
				{
					strWorseSql = strWorseSql + "variable" + strVariableNumber + "_worse_yn=IIF(" + strWorseIsNotNull[x].Trim() + ",IIF(" + oFvsVar.m_strWorseExpr[x].Trim() + ",'Y','N'),null),";
				}
				if (oFvsVar.m_strEffectiveExpr[x].Trim().Length > 0)
				{
					strEffectiveSql = strEffectiveSql + "variable" + strVariableNumber + "_effective_yn=IIF(" + strEffectiveIsNotNull[x].Trim() + ",IIF(" + oFvsVar.m_strEffectiveExpr[x].Trim() + ",'Y','N'),null),";
				}
			}
			if (oFvsVar.m_strOverallEffectiveExpr.Trim().Length > 0)
			{
				if (oFvsVar.m_bOverallEffectiveNetRevEnabled)
					m_strSQL = m_strSQL + "overall_effective_yn=IIF(" + strOverallEffectiveIsNotNull.Trim() + ",IIF(" + oFvsVar.m_strOverallEffectiveExpr.Trim() + " AND nr_dpa IS NOT NULL AND nr_dpa " + oFvsVar.m_strOverallEffectiveNetRevOperator.Trim() + " " + oFvsVar.m_strOverallEffectiveNetRevValue.Trim() + ",'Y','N'),null),";
				else
			  	    m_strSQL = m_strSQL + "overall_effective_yn=IIF(" + strOverallEffectiveIsNotNull.Trim() + ",IIF(" + oFvsVar.m_strOverallEffectiveExpr.Trim() + ",'Y','N'),null),";
			}
			//better
            if (strBetterSql.Trim().Length > 0)
            {
                strBetterSql = strBetterSql.Substring(0, strBetterSql.Length - 1);
                strBetterSql = "UPDATE cycle1_effective SET " + strBetterSql;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Improvement--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strBetterSql + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strBetterSql);

            }

			//worse
            if (strWorseSql.Trim().Length > 0)
            {
                strWorseSql = strWorseSql.Substring(0, strWorseSql.Length - 1);
                strWorseSql = "UPDATE cycle1_effective SET " + strWorseSql;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Regression--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strWorseSql + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strWorseSql);
                
            }

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			//effective
			strEffectiveSql = strEffectiveSql.Substring(0,strEffectiveSql.Length - 1);
			strEffectiveSql = "UPDATE cycle1_effective SET " + strEffectiveSql;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--Variable Effective--\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strEffectiveSql + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strEffectiveSql);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            
			//overall effective
			m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);
			m_strSQL = "UPDATE cycle1_effective SET " + m_strSQL;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--Overall Effective--\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

           
            

			
			if (this.m_ado.m_intError != 0)
			{
                 if (frmMain.g_bDebug)  frmMain.g_oUtils.WriteText(m_strDebugFile,"!!!Error Executing SQL!!!\r\n");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (Convert.ToInt32(m_ado.getRecordCount(m_TempMDBFileConn,"SELECT COUNT(*) FROM cycle1_effective WHERE overall_effective_yn='Y'","temp"))==0)
            {
                
                    MessageBox.Show("No overall effective treatments found. Processing is cancelled");
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Cancelled!!");
                    ReferenceUserControlScenarioRun.m_bUserCancel = true;
                    return;
                
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
            
            
			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
			}
            


		}

		private void Cycle1Optimization()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Cycle1Optimization\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nCycle1 Optimization\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--------------------\r\n");
            }

            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                    ReferenceUserControlScenarioRun.listViewEx1, "Cycle 1: Optimize the Effective Treatments For Each Stand");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 3;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


			/********************************************
			 **delete all records in the tie breaker table
			 ********************************************/
			this.m_strSQL = "delete from cycle1_optimization";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			//insert the valid combos into the tiebreakre table
			m_strSQL = "INSERT INTO cycle1_optimization (biosum_cond_id,rxpackage,rx,rxcycle) " + 
				         "SELECT biosum_cond_id,rxpackage,rx,rxcycle FROM cycle1_effective WHERE overall_effective_yn='Y'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nOptimization Type: " + this.m_oOptimizationVariable.strOptimizedVariable.ToUpper() + "\r\n");
            
			//populate the variable table.column name and its value to the effective table
			if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
			{
				m_strSQL = "UPDATE cycle1_optimization e " + 
					"INNER JOIN product_yields_net_rev_costs_summary_by_rx p " + 
					"ON e.biosum_cond_id=p.biosum_cond_id AND " + 
                    "e.rxpackage=p.rxpackage AND " + 
                    "e.rx=p.rx AND " + 
                    "e.rxcycle = p.rxcycle " + 
					"SET e.pre_variable_name = 'product_yields_net_rev_costs_summary_by_rx.max_nr_dpa'," + 
					    "e.post_variable_name = 'product_yields_net_rev_costs_summary_by_rx.max_nr_dpa'," + 
					    "e.pre_variable_value = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)," + 
					    "e.post_variable_value = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)," + 
					    "e.change_value = 0";

																									 ;
			}
			else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
			{
				m_strSQL = "UPDATE cycle1_optimization e " + 
					"INNER JOIN product_yields_net_rev_costs_summary_by_rx p " + 
					"ON e.biosum_cond_id=p.biosum_cond_id AND " +
                    "e.rxpackage=p.rxpackage AND " +
                    "e.rx=p.rx AND " +
                    "e.rxcycle = p.rxcycle " + 
					"SET e.pre_variable_name = 'product_yields_net_rev_costs_summary_by_rx.merch_yield_cf'," + 
					"e.post_variable_name = 'product_yields_net_rev_costs_summary_by_rx.merch_yield_cf'," + 
					"e.pre_variable_value = IIF(p.merch_yield_cf IS NOT NULL,p.merch_yield_cf,0)," + 
					"e.post_variable_value = IIF(p.merch_yield_cf IS NOT NULL,p.merch_yield_cf,0)," + 
					"e.change_value = 0";

			}
			else
			{
						
				string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strFVSVariableName,".");
				strPreTable="PRE_" + strCol[0].Trim();
				strPreColumn=strCol[1].Trim();
				strPostTable="POST_" + strCol[0].Trim();
				strPostColumn=strCol[1].Trim();

				m_strSQL = "UPDATE cycle1_optimization e " + 
						"INNER JOIN (" + strPostTable + " a " + 
						"INNER JOIN " + strPreTable + " b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND " +
                           "a.rxpackage=b.rxpackage AND " + 
                           "a.rx = b.rx AND " + 
                           "a.rxcycle = b.rxcycle) " + 
						"ON e.biosum_cond_id=a.biosum_cond_id AND " +
                        "e.rxpackage=a.rxpackage AND " +
                        "e.rx=a.rx AND " +
                        "e.rxcycle = a.rxcycle " + 
						"SET e.change_value=" + 
								"IIF(a." + strPostColumn + " IS NOT NULL AND b." + strPreColumn + " IS NOT NULL," + 
									"a." + strPostColumn + " - b." + strPreColumn + "," + 
								"IIF(a." + strPostColumn + " IS NOT NULL," + 
									"a." + strPostColumn + "," + 
								"IIF(b." + strPreColumn  + " IS NOT NULL,0-b." + strPreColumn + ",0)))," + 
                           "e.pre_variable_name='Pre_" + this.m_oOptimizationVariable.strFVSVariableName.Trim() + "'," + 
						   "e.post_variable_name='Post_" + this.m_oOptimizationVariable.strFVSVariableName.Trim() + "'," + 
						   "e.pre_variable_value=IIF(b." + strPreColumn + " IS NOT NULL,b." + strPreColumn + ",0)," + 
						   "e.post_variable_value=IIF(a." + strPostColumn + " IS NOT NULL,a." + strPostColumn + ",0)";
			}
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");

		}


		private void tiebreaker()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//tiebreaker\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";
			string[] strArray=null;
			string strVariableNumber="";

            FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection =
		        ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection;

			FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem oItem;


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nTie Breaker\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--------------------\r\n");
            }


             intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                    ReferenceUserControlScenarioRun.listViewEx1, "Cycle 1: Load Tie Breaker Tables");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 4;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


           


			/********************************************
			 **delete all records in the tie breaker table
			 ********************************************/
			this.m_strSQL = "delete from tiebreaker";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			//insert the valid combos into the tiebreakre table
			m_strSQL = "INSERT INTO tiebreaker (biosum_cond_id,rxpackage,rx,rxcycle) " + 
				          "SELECT biosum_cond_id,rxpackage,rx,rxcycle FROM cycle1_effective WHERE overall_effective_yn='Y'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			//populate the variable table.column name and its value to the effective table
			
				
			oItem = oTieBreakerCollection.Item(0);
			if (oItem.bSelected)
			{
				strArray=frmMain.g_oUtils.ConvertListToArray(oItem.strFVSVariableName,".");

				strPreTable="PRE_" + strArray[0].Trim();
				strPreColumn=strArray[1].Trim();
				strPostTable="POST_" + strArray[0].Trim();
				strPostColumn=strArray[1].Trim();
				strVariableNumber="1";
				if (strPreTable.Trim().Length > 0)
				{
					m_strSQL = "e.pre_variable" + strVariableNumber + "_name='" + strPreTable + "." + strPreColumn + "',";
					m_strSQL = m_strSQL + "e.pre_variable" + strVariableNumber + "_value=pre." + strPreColumn;
				}
				else
				{
					m_strSQL = "e.pre_variable" + strVariableNumber + "_name=null,";
					m_strSQL = m_strSQL + "e.pre_variable" + strVariableNumber + "_value=null";
				}

				if (strPostTable.Trim().Length > 0)
				{
					m_strSQL = m_strSQL + ",e.post_variable" + strVariableNumber + "_name='" + strPostTable + "." + strPostColumn + "',";
					m_strSQL = m_strSQL + "e.post_variable" + strVariableNumber + "_value=post." + strPostColumn;
				}
				else
				{
					m_strSQL = m_strSQL + ",e.post_variable" + strVariableNumber + "_name=null,";
					m_strSQL = m_strSQL + "e.post_variable" + strVariableNumber + "_value=null";
				}
				if (strPreTable.Trim().Length > 0 && strPostTable.Trim().Length > 0)
				{
					m_strSQL = "UPDATE tiebreaker e INNER JOIN (" + strPostTable + " post INNER JOIN " + 
						strPreTable  + " pre ON post.biosum_cond_id = pre.biosum_cond_id AND " + 
                                               "post.rxpackage = pre.rxpackage AND " + 
                                               "post.rx = pre.rx AND " + 
                                               "post.rxcycle = pre.rxcycle) "  + 
						"ON e.biosum_cond_id=post.biosum_cond_id AND " + 
                           "e.rxpackage=post.rxpackage AND  " + 
                           "e.rx=post.rx AND " + 
                           "e.rxcycle=post.rxcycle " + 
						"SET " + m_strSQL;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

				}

				//populate the change column by subtracting pre value from post value
				m_strSQL="";
				m_strSQL = m_strSQL + "variable" + strVariableNumber + "_change=IIF(pre_variable" + strVariableNumber + "_value IS NOT NULL AND post_variable" + strVariableNumber + "_value IS NOT NULL,post_variable" + strVariableNumber + "_value - pre_variable" + strVariableNumber + "_value,null),";
				m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);

				m_strSQL = "UPDATE tiebreaker SET " + m_strSQL;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			
				
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			oItem = oTieBreakerCollection.Item(1);
			if (oItem.bSelected)
			{
				m_strSQL = "UPDATE tiebreaker a INNER JOIN scenario_rx_intensity b ON a.rx=b.rx SET a.rx_intensity=b.rx_intensity " + 
					       "WHERE trim(ucase(b.scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
			


		}


		

		/// <summary>
		/// get the wood product yields,
		/// revenue, and costs of an applied
		/// treatment on a plot 
		/// </summary>
		private void product_yields_net_rev_costs_summary_by_rx()
		{

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//product_yields_net_rev_costs_summary_by_rx\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nWood Product Yields,Revenue, And Costs Table By Treatment\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "-----------------------------------------------\r\n");
            }
            int intListViewIndex = -1;

			string[]  strUpdateSQL = new string[25];

           
			/**********************************************
			 **complete harvest cost per acre
			 **********************************************/
            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                        ReferenceUserControlScenarioRun.listViewEx1, "Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Treatment");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 5;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


			/*****************************************************
			 **get the Chips market value per green ton
			 *****************************************************/
			ado_data_access p_ado = new ado_data_access();
			string strScenarioMDB="";
			string strScenarioConn="";
			p_ado.getScenarioConnStringAndMDBFile(ref strScenarioMDB,ref strScenarioConn, frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text);
			p_ado.OpenConnection(strScenarioConn);
			if (p_ado.m_intError != 0)
			{
				this.m_intError = p_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				p_ado = null;
				return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/**********************************************************
			 **create the at hazard expression
			 **********************************************************/
			this.m_strSQL = "SELECT * FROM scenario_costs WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);
			double dblChipMktValPgt=0;
			double dblDriverCostPgtPerHour=0;
			if (p_ado.m_OleDbDataReader.HasRows==true)
			{
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["road_haul_cost_pgt_per_hour"] != System.DBNull.Value)
					{
						dblChipMktValPgt = Convert.ToDouble(p_ado.m_OleDbDataReader["chip_mkt_val_pgt"].ToString().Trim());
						dblDriverCostPgtPerHour= Convert.ToDouble(p_ado.m_OleDbDataReader["road_haul_cost_pgt_per_hour"].ToString().Trim());
					}																																  
				}
			}
			p_ado.m_OleDbDataReader.Close();
            dblChipMktValPgt = Convert.ToDouble(m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(0).ChipsDollarPerCubicFootValue);

            /**************************************************
             **END:BYPASS SINCE HARVEST COSTS ARE CALCULATED
             **IN THE PROCESSOR SCENARIO
             **************************************************/

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/********************************************
			 **delete all records in the table
			 ********************************************/
			this.m_strSQL = "delete from product_yields_net_rev_costs_summary_by_rx";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            

			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!!!Error Executing SQL!!!\r\n");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert Records\r\n");


            if (m_ado.TableExist(this.m_TempMDBFileConn, "product_yields_net_rev_costs_summary_by_rx_work_table"))
                m_ado.SqlNonQuery(this.m_TempMDBFileConn, "DROP TABLE product_yields_net_rev_costs_summary_by_rx_work_table");
            m_strSQL = 
                "SELECT validcombos.biosum_cond_id,validcombos.rxpackage,validcombos.rx," + 
                       "validcombos.rxcycle," + 
                       this.m_strTreeVolValSumTable.Trim() + ".merch_vol_cf AS merch_yield_cf," +
				       this.m_strTreeVolValSumTable.Trim() + ".chip_vol_cf AS chip_yield_cf," +
				       this.m_strTreeVolValSumTable.Trim() + ".chip_wt_gt AS chip_yield_gt," +
				       this.m_strTreeVolValSumTable.Trim() + ".chip_val_dpa AS chip_val_dpa," + 
				       this.m_strTreeVolValSumTable.Trim() + ".merch_wt_gt AS merch_yield_gt," +
				       this.m_strTreeVolValSumTable.Trim() + ".merch_val_dpa AS merch_val_dpa," +
                       this.m_strHvstCostsTable.Trim() + ".complete_cpa AS harvest_onsite_cpa," + 
                      "IIF(" + this.m_strPlotTable.Trim() + ".merch_haul_cpa_pt IS NOT NULL," +
                      "IIF(tree_vol_val_sum_by_rx.rxcycle='2'," + this.m_strPlotTable.Trim() + ".merch_haul_cpa_pt * " + this.m_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle2 + "," + 
                      "IIF(tree_vol_val_sum_by_rx.rxcycle='3'," + this.m_strPlotTable.Trim() + ".merch_haul_cpa_pt * " + this.m_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle3 + "," + 
                      "IIF(tree_vol_val_sum_by_rx.rxcycle='4'," + this.m_strPlotTable.Trim() + ".merch_haul_cpa_pt * " + this.m_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle4 + "," + 
                       this.m_strPlotTable.Trim() + ".merch_haul_cpa_pt))),0) AS escalator_merch_haul_cpa_pt," + 
                      "escalator_merch_haul_cpa_pt * tree_vol_val_sum_by_rx.merch_wt_gt AS haul_merch_cpa," + 
                      "IIF(" + this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt IS NOT NULL," + 
                      "IIF(tree_vol_val_sum_by_rx.rxcycle='2'," + this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle2 + "," + 
                      "IIF(tree_vol_val_sum_by_rx.rxcycle='3'," + this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle3 + "," + 
                      "IIF(tree_vol_val_sum_by_rx.rxcycle='4'," + this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle4 + "," + 
                       this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt))),0) AS escalator_chip_haul_cpa_pt," +
                      "escalator_chip_haul_cpa_pt * tree_vol_val_sum_by_rx.chip_wt_gt AS haul_chip_cpa," + 
                      "merch_val_dpa + chip_val_dpa - harvest_onsite_cpa - haul_merch_cpa - haul_chip_cpa AS merch_chip_nr_dpa," + 
                      "merch_val_dpa - harvest_onsite_cpa - haul_merch_cpa AS merch_nr_dpa," + 
                      "IIF(" + this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt IS NOT NULL AND " + 
                           "merch_chip_nr_dpa > merch_nr_dpa,'Y','N') AS usebiomass_yn," + 
                      "IIF(usebiomass_yn = 'Y', merch_chip_nr_dpa,merch_nr_dpa) AS max_nr_dpa " +
                      "INTO product_yields_net_rev_costs_summary_by_rx_work_table " + 
                      "FROM ((((validcombos " +
                      "INNER JOIN " + this.m_strCondTable.Trim() + " " +
                      "ON validcombos.biosum_cond_id = " + this.m_strCondTable.Trim() + ".biosum_cond_id) " +
                      "INNER JOIN " + this.m_strPlotTable.Trim() + " " +
                      "ON " +
                        this.m_strCondTable.Trim() + ".biosum_plot_id=" + this.m_strPlotTable.Trim() + ".biosum_plot_id) " +
                      "INNER JOIN " + this.m_strTreeVolValSumTable.Trim() + " " +
                      "ON " +
                      "(validcombos.biosum_cond_id=" + this.m_strTreeVolValSumTable.Trim() + ".biosum_cond_id) AND " +
                      "(validcombos.rxpackage=" + this.m_strTreeVolValSumTable.Trim() + ".rxpackage) AND " +
                      "(validcombos.rx=" + this.m_strTreeVolValSumTable.Trim() + ".rx) AND " +
                      "(validcombos.rxcycle=" + this.m_strTreeVolValSumTable.Trim() + ".rxcycle)) " +
                      "INNER JOIN " +
                        this.m_strHvstCostsTable.Trim() + " " +
                      "ON " +
                      "(validcombos.biosum_cond_id=" + this.m_strHvstCostsTable.Trim() + ".biosum_cond_id) AND " +
                      "(validcombos.rxpackage=" + this.m_strHvstCostsTable.Trim() + ".rxpackage) AND " +
                      "(validcombos.rx=" + this.m_strHvstCostsTable.Trim() + ".rx) AND " +
                      "(validcombos.rxcycle=" + this.m_strHvstCostsTable.Trim() + ".rxcycle)); ";


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");

                
                return;
            }

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            this.m_strSQL = "INSERT INTO product_yields_net_rev_costs_summary_by_rx " +
                                "(biosum_cond_id,rxpackage,rx,rxcycle," +
                                "merch_yield_cf,chip_yield_cf," +
                                "chip_yield_gt,chip_val_dpa," +
                                "merch_yield_gt,merch_val_dpa,harvest_onsite_cpa," +
                                "haul_merch_cpa,haul_chip_cpa,merch_chip_nr_dpa," +
                                "merch_nr_dpa,usebiomass_yn,max_nr_dpa) " + 
                            "SELECT biosum_cond_id,rxpackage,rx,rxcycle," +
                                "merch_yield_cf,chip_yield_cf," +
                                "chip_yield_gt,chip_val_dpa," +
                                "merch_yield_gt,merch_val_dpa,harvest_onsite_cpa," +
                                "haul_merch_cpa,haul_chip_cpa,merch_chip_nr_dpa," +
                                "merch_nr_dpa,usebiomass_yn,max_nr_dpa " + 
                            "FROM product_yields_net_rev_costs_summary_by_rx_work_table";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;

			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
			}
     

		}

        /// <summary>
        /// get the wood product yields,
        /// revenue, and costs of an applied
        /// treatment on a plot 
        /// </summary>
        private void product_yields_net_rev_costs_summary_by_rxpackage()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//product_yields_net_rev_costs_summary_by_rxpackage\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nWood Product Yields,Revenue, And Costs Table By Treatment Package\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "-------------------------------------------------------\r\n");
            }

            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                      ReferenceUserControlScenarioRun.listViewEx1, "Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Treatment Package");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 2;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


            /********************************************
             **delete all records in the table
             ********************************************/
            this.m_strSQL = "delete from product_yields_net_rev_costs_summary_by_rxpackage";
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
            }
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert Records\r\n");

            m_strSQL = "INSERT INTO product_yields_net_rev_costs_summary_by_rxpackage ";
            m_strSQL =  m_strSQL + "(biosum_cond_id,rxpackage," +
                                    "chip_yield_cf,merch_yield_cf," +
                                    "chip_yield_gt,merch_yield_gt," +
                                    "chip_val_dpa,merch_val_dpa," +
                                    "harvest_onsite_cpa,haul_chip_cpa,haul_merch_cpa," +
                                    "merch_chip_nr_dpa," +
                                    "merch_nr_dpa," + 
                                    "max_nr_dpa)";
            m_strSQL = m_strSQL + " " +
                            "SELECT a.biosum_cond_id,a.rxpackage," +
                                   "a.sum_chip_yield_cf AS chip_yield_cf," +
                                   "a.sum_merch_yield_cf AS merch_yield_cf," +
                                   "a.sum_chip_yield_gt AS chip_yield_gt," +
                                   "a.sum_merch_yield_gt AS merch_yield_gt," +
                                   "a.sum_chip_val_dpa AS chip_val_dpa," +
                                   "a.sum_merch_val_dpa AS merch_val_dpa," +
                                   "a.sum_harvest_onsite_cpa AS harvest_onsite_cpa," +
                                   "a.sum_haul_chip_cpa AS haul_chip_cpa," +
                                   "a.sum_haul_merch_cpa AS haul_merch_cpa," +
                                   "a.sum_merch_chip_nr_dpa AS merch_chip_nr_dpa," +
                                   "a.sum_merch_nr_dpa AS merch_nr_dpa, " +
                                   "a.sum_max_nr_dpa AS max_nr_dpa " + 
                           "FROM (" +
                           "SELECT biosum_cond_id,rxpackage," +
                                "SUM(IIF(chip_yield_cf IS NULL,0,chip_yield_cf)) AS sum_chip_yield_cf," +
                                "SUM(IIF(merch_yield_cf IS NULL,0,merch_yield_cf)) AS sum_merch_yield_cf," +
                                "SUM(IIF(chip_yield_gt IS NULL,0,chip_yield_gt)) AS sum_chip_yield_gt," +
                                "SUM(IIF(merch_yield_gt IS NULL,0,merch_yield_gt)) AS sum_merch_yield_gt," +
                                "SUM(IIF(chip_val_dpa IS NULL,0,chip_val_dpa)) AS sum_chip_val_dpa," +
                                "SUM(IIF(merch_val_dpa IS NULL,0,merch_val_dpa)) AS sum_merch_val_dpa," +
                                "SUM(IIF(harvest_onsite_cpa IS NULL,0,harvest_onsite_cpa)) AS sum_harvest_onsite_cpa," +
                                "SUM(IIF(haul_chip_cpa IS NULL,0,haul_chip_cpa)) AS sum_haul_chip_cpa," +
                                "SUM(IIF(haul_merch_cpa IS NULL,0,haul_merch_cpa)) AS sum_haul_merch_cpa," +
                                "SUM(IIF(merch_chip_nr_dpa IS NULL,0,merch_chip_nr_dpa)) AS sum_merch_chip_nr_dpa," +
                                "SUM(IIF(merch_nr_dpa IS NULL,0,merch_nr_dpa)) AS sum_merch_nr_dpa," +
                                "SUM(IIF(max_nr_dpa IS NULL,0,max_nr_dpa)) AS sum_max_nr_dpa " + 
                           "FROM product_yields_net_rev_costs_summary_by_rx " +
                           "GROUP BY biosum_cond_id,rxpackage) a";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");

                return;
            }

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();


            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            


            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");

            }


        }

		/// <summary>
		/// create a temporary work table for summing harvest costs
		/// </summary>
		private void CreateTableStructureOfHarvestCosts()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureOfHarvestCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

           
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn = p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the harvest_costs structure
				 *********************************************/
				this.m_strSQL = "SELECT biosum_cond_id,rxpackage,rx,rxcycle, complete_cpa FROM " + this.m_strHvstCostsTable.Trim() + ";";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of harvest_costs_sum
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create harvest_costs_sum Table Schema From harvest_costs table--\r\n");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"harvest_costs_sum",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}

		private void CreateTableStructureForScenarioProcessingSites()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForScenarioProcessingSites\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the haul costs table structure
				 *********************************************/
				this.m_strSQL = "SELECT psite_id,trancd,biocd FROM scenario_psites;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

                 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create scenario psites work table from  the scenario_psites table--\r\n");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"scenario_psites_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}	
				p_dt.Clear();
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}
		/// <summary>
		/// create temporary work tables for getting 
		/// a plot's cheapest merch and chip haul cost
		/// </summary>
		private void CreateTableStructureForHaulCosts()
		{

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForHaulCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access oAdo = new ado_data_access();
			this.m_strConn= oAdo.getMDBConnString(this.m_strTempMDBFile,"admin","");
			oAdo.OpenConnection(this.m_strConn);
			if (oAdo.m_intError==0)
			{

				/*****************************************************************
				 **create the table structures in the temp mdb file
				 **and give them the name OF all_road_merch_haul_costs_work_table and 
				 **                          all_road_chip_haul_costs_work_table
				 **                          cheapest_road_merch_haul_costs_work_table
				 **                          cheapest_road_chip_haul_costs_work_table
				 **                          cheapest_rail_merch_haul_costs_work_table
				 **                          cheapest_rail_chip_haul_costs_work_table
				 **                          merch_plot_to_rh_to_collector_haul_costs_work_table
				 **                          chip_plot_to_rh_to_collector_haul_costs_work_table
				 **                          combine_chip_rail_road_haul_costs_work_table
				 **                          combine_merch_rail_road_haul_costs_work_table
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Create haul costs work tables from  the haul_costs table\r\n");

				if (oAdo.TableExist(oAdo.m_OleDbConnection,"haul_costs"))
					oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE haul_costs");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"haul_costs");


				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"all_road_chip_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE all_road_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"all_road_chip_haul_costs_work_table");


				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"all_road_merch_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE all_road_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"all_road_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_road_merch_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_road_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_road_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_road_chip_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_road_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_road_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_rail_merch_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_rail_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_rail_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_rail_chip_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_rail_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_rail_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_merch_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostWorkTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_merch_haul_costs_work_table");

				

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_chip_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostWorkTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_chip_haul_costs_work_table");

				
				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"chip_plot_to_rh_to_collector_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE chip_plot_to_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"chip_plot_to_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"merch_plot_to_rh_to_collector_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE merch_plot_to_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"merch_plot_to_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"combine_chip_rail_road_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE combine_chip_rail_road_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"combine_chip_rail_road_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"combine_merch_rail_road_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE combine_merch_rail_road_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"combine_merch_rail_road_haul_costs_work_table");


				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"merch_rh_to_collector_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE merch_rh_to_collector_haul_costs_work_table");

				/*****************************************************************
				 **create the railroad table structures in the temp mdb file
				 **and give them the name OF merch_rh_to_collector_haul_costs_work_table
				 **                          chip_rh_to_collector_haul_costs_work_table
				 *****************************************************************/
				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostRailroadTable(
						oAdo,oAdo.m_OleDbConnection,"merch_rh_to_collector_haul_costs_work_table");


				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"chip_rh_to_collector_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE chip_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oCoreScenarioResults.CreateHaulCostRailroadTable(
						oAdo,oAdo.m_OleDbConnection,"chip_rh_to_collector_haul_costs_work_table");

				
				if (oAdo.m_intError!=0)
				{

                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=oAdo.m_intError;
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
					oAdo=null;
					return;
				}
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
			else
			{
				this.m_intError = oAdo.m_intError;
			}
			oAdo=null;

		}
		
		/// <summary>
		/// create temporary work tables for getting 
		/// a plot's fastest travel time to a processing site
		/// </summary>
		private void CreateTableStructureForFastestTravelTimes()
		{

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForFastestTravelTimes\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the plot table structure
				 *********************************************/
				this.m_strSQL = "SELECT biosum_plot_id, merch_haul_cost_id, merch_haul_cost_psite, merch_haul_cpa_pt, chip_haul_cost_id,chip_haul_cost_psite,chip_haul_cpa_pt FROM " + this.m_strPlotTable.Trim() + ";";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structures in the temp mdb file
				 **and give them the name OF plot_fastest_tvltm_work_table and 
				 **                          plot_fastest_tvltm_work_table2
				 *****************************************************************/
                 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create travel times work table (plot_fastest_tvltm_work_table) from  the plot table--\r\n");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_tvltm_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_tvltm_work_table2",p_dt,true);
				p_dt.Clear();
				
				/*********************************************
				 **get the haul cost table structure
				 *********************************************/
				this.m_strSQL = "SELECT haul_cost_id, biosum_plot_id,railhead_id,psite_id,transfer_cost,road_cost,rail_cost,total_haul_cost FROM haul_costs;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structures in the temp mdb file
				 **and give them the name OF haul_costs_work_table and 
				 **                          haul_costs_work_table2
				 *****************************************************************/
                 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create haul costs work table (haul_costs_work_table) from  the haul_costs table--");
				
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"haul_costs_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"haul_costs_work_table2",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}
		
		/// <summary>
		/// create a temporary work tables for finding best treatments
		/// </summary>
		private void CreateTableStructureForIntensity()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForIntensity\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access oAdo = new ado_data_access();
			this.m_strConn= oAdo.getMDBConnString(this.m_strTempMDBFile,"admin","");
			oAdo.OpenConnection(this.m_strConn);


			if (oAdo.m_intError==0)
			{
				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
                 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create rx_intensity_work_table Schema--\r\n");

				frmMain.g_oTables.m_oCoreScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_duplicates_work_table");
				if (oAdo.m_intError==0) frmMain.g_oTables.m_oCoreScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_duplicates_work_table2");
				if (oAdo.m_intError==0) frmMain.g_oTables.m_oCoreScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_unique_work_table");

				if (oAdo.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=oAdo.m_intError;
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
					return;
				}
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
			else
			{
				this.m_intError = oAdo.m_intError;
			}
			oAdo=null;

		}

		/// <summary>
		/// create a temporary work tables for identifying inaccessible plots and conditions
		/// </summary>
		private void CreateTableStructureForPlotCondAccessiblity()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForPlotCondAccessiblity\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get fields from the plot table
				 *********************************************/
				this.m_strSQL = "SELECT biosum_plot_id, num_cond, num_cond as num_cond_not_accessible FROM " + this.m_strPlotTable;

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of plot_cond_accessible_work_table
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create plot_cond_accessible_work_table Schema--\r\n");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_cond_accessible_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_cond_accessible_work_table2",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}
		private void CreateTableStructureForUserDefinedConditionTable()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForUserDefinedConditionTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/********************************************************
				 **get the user defined PLOT filter sql
				 ********************************************************/
				this.m_strSQL = this.m_strUserDefinedCondSQL;
				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the scenario_results.mdb file
				 **and give it the name of userdefinedplotfilter
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create userdefinedcondfilter Table Schema From User Defined Condition Filter SQL--\r\n");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strSystemResultsDbPathAndFile,"userdefinedcondfilter",p_dt,true);
				p_dt.Dispose();
				this.m_ado.m_OleDbDataReader.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					return;
				}
				/***********************************************************************
				 **create a table link to the user defined plot filter table in the
				 **temporary MDB file located on the hard drive of the user
				 ***********************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create link in " + this.m_strTempMDBFile + "--\r\n");
				p_dao.CreateTableLink(this.m_strTempMDBFile,"userdefinedcondfilter",m_strSystemResultsDbPathAndFile,"userdefinedcondfilter",true);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,m_strSystemResultsDbPathAndFile + "\tuserdefinedcondfilter\r\n");


				/*********************************************
				 **get fields from the plot table
				 *********************************************/
				this.m_strSQL = "SELECT * FROM " + this.m_strCondTable;

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of plot_cond_accessible_work_table
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create ruledefinitionscondfilter Schema--\r\n");
				p_dao.CreateMDBTableFromDataSetTable(this.m_strSystemResultsDbPathAndFile,"ruledefinitionscondfilter",p_dt,true);
				/***********************************************************************
   			     **create a table link to the ruledefinitionscondfilter table in the
			     **temporary MDB file located on the hard drive of the user
			     ***********************************************************************/
				p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionscondfilter",m_strSystemResultsDbPathAndFile,"ruledefinitionscondfilter",true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}
		/// <summary>
		/// create temporary work table for summing up wood volumes,values and costs
		/// by wood processing site
		/// </summary>
		private void CreateTableStructureForPSiteSum()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForPSiteSum\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access oAdo = new ado_data_access();
			this.m_strConn= oAdo.getMDBConnString(this.m_strTempMDBFile,"admin","");
			oAdo.OpenConnection(this.m_strConn);

			if (oAdo.m_intError==0)
			{
				if (oAdo.TableExist(oAdo.m_OleDbConnection,"psite_sum_work_table"))
					   oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE psite_sum_work_table");
				frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationPSiteSumTable(oAdo,
					oAdo.m_OleDbConnection,"psite_sum_work_table");

				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
			else
			{
				this.m_intError = oAdo.m_intError;
			}
			oAdo=null;

		}
		/// <summary>
		/// create temporary work table for summing up wood volumes,values and costs
		/// by wood processing site
		/// </summary>
		private void CreateTableStructureForPSiteSumOld()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/**********************************************************
				 **create the max_nr_plots structure with sum_ field names
				 **********************************************************/
				this.m_strSQL = "SELECT s.psite_id, " +
					"s.biocd," + 
					"p.acres AS sum_acres," + 
					"p.merch_haul_cents AS sum_merch_haul_cents," + 
					"p.chip_haul_cents AS sum_chip_haul_cents," + 
					"p.merch_vol AS sum_merch_vol," + 
					"p.chip_yield AS sum_chip_yield," +
					"p.net_rev AS sum_net_rev," + 
					"p.merch_dollars_val AS sum_merch_dollars_val," + 
					"p.chip_dollars_val AS sum_chip_dollars_val," + 
					"p.harv_costs AS sum_harv_costs," + 
					"p.haul_costs AS sum_haul_costs," + 
					"p.ti_chg_acres AS sum_ti_chg_acres," + 
					"p.ci_chg_acres AS sum_ci_chg_acres " + 
					"FROM max_nr_plots p," + this.m_strPSiteTable.Trim() + " s;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
				
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"psite_sum_work_table",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					//this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}
		/// <summary>
		/// create temporary work table for summing up wood volumes,values and costs
		/// by land ownership
		/// </summary>
		private void CreateTableStructureForOwnershipSum()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForOwnershipSum\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access oAdo = new ado_data_access();
			this.m_strConn= oAdo.getMDBConnString(this.m_strTempMDBFile,"admin","");
			oAdo.OpenConnection(this.m_strConn);
			if (oAdo.m_intError==0)
			{
				if (oAdo.TableExist(oAdo.m_OleDbConnection,"own_sum_work_table"))
					oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE own_sum_work_table");
						frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationOwnershipSumTable(
							oAdo,
							oAdo.m_OleDbConnection,
							"own_sum_work_table");


				if (oAdo.m_intError==0 && oAdo.TableExist(oAdo.m_OleDbConnection,"own_sum_work_table_air_dest"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE own_sum_work_table_air_dest");
				if (oAdo.m_intError==0) frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationOwnershipSumTable(
											oAdo,
											oAdo.m_OleDbConnection,
											"own_sum_work_table_air_dest");
				if (oAdo.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=oAdo.m_intError;
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
					return;
				}
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
			else
			{
				this.m_intError = oAdo.m_intError;
			}
			oAdo=null;

		}

		/// <summary>
		/// create temporary work table for summing up wood volumes,values and costs
		/// by land ownership
		/// </summary>
		private void CreateTableStructureForOwnershipSumOld()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				
				/**********************************************************
				 **create the max_nr_plots structure with sum_ field names
				 **********************************************************/
				this.m_strSQL = "SELECT p.owngrpcd, " +
					"p.acres AS sum_acres," + 
					"p.merch_haul_cents AS sum_merch_haul_cents," + 
					"p.chip_haul_cents AS sum_chip_haul_cents," + 
					"p.merch_vol AS sum_merch_vol," + 
					"p.chip_yield AS sum_chip_yield," +
					"p.net_rev AS sum_net_rev," + 
					"p.merch_dollars_val AS sum_merch_dollars_val," + 
					"p.chip_dollars_val AS sum_chip_dollars_val," + 
					"p.harv_costs AS sum_harv_costs," + 
					"p.haul_costs AS sum_haul_costs," + 
					"p.ti_chg_acres AS sum_ti_chg_acres," + 
					"p.ci_chg_acres AS sum_ci_chg_acres " + 
					"FROM max_nr_plots p;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
				//this.m_txtStreamWriter.WriteLine("Create max_nr_plots Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"own_sum_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"own_sum_work_table_air_dest",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					//this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}



		/// <summary>
		/// find the best treatment by these categories: 
		/// maximum net revenue; merchantable wood removal;
		/// and optimization variable
		/// </summary>
		private void Cycle1_best_rx_summary()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Cycle1_best_rx_summary\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string strTable="";
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Finding Best Treatments: Maximum Net Revenue");
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",true);
			frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Refresh");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n\r\nBest Rx Summary Cycle 1\r\n");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--------------------\r\n");

            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                   ReferenceUserControlScenarioRun.listViewEx1, "Cycle 1: Identify The Best Effective Treatment For Each Stand");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 5;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);

			
			FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection =
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection;


		    string strScenarioId = this.ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim();
			string strTieBreakerAggregate="MAX";
			
			

	
			//
			//CREATE WORK TABLES
			//
			//best_rx_summary_work_table
			if (m_ado.TableExist(this.m_TempMDBFileConn,
				Tables.CoreScenarioResults.DefaultScenarioResultsCycle1BestRxSummaryTableName + "_work_table"))
			{
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,
					"DROP TABLE " + Tables.CoreScenarioResults.DefaultScenarioResultsCycle1BestRxSummaryTableName + "_work_table");
			}
			strTable = Tables.CoreScenarioResults.DefaultScenarioResultsCycle1BestRxSummaryTableName + "_work_table";
			m_ado.SqlNonQuery(m_TempMDBFileConn,Tables.CoreScenarioResults.CreateBestRxSummaryCycle1TableSQL(strTable));
            //best_rx_summury_optimization_and_tiebreaker_work_table
			if (m_ado.TableExist(m_TempMDBFileConn,"cycle1_best_rx_summary_optimization_and_tiebreaker_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,Tables.CoreScenarioResults.CreateBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table"));
			//best_rx_summury_optimization_and_tiebreaker_work_table2
			if (m_ado.TableExist(m_TempMDBFileConn,"cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,Tables.CoreScenarioResults.CreateBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2"));
			//best_rx_summury_optimization_and_tiebreaker_work_table3
			if (m_ado.TableExist(m_TempMDBFileConn,"cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,Tables.CoreScenarioResults.CreateBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3"));

			/**********************************************
			 **insert unique biosum_cond_id's into the
			 **best_rx_summary table so we dont have
			 **to worry about whether the biosum_cond_id 
			 **record is in the table or not
			 **********************************************/
			this.m_strSQL = "INSERT INTO " + strTable  +  " " +
				"SELECT DISTINCT c.biosum_cond_id,c.acres,c.owngrpcd " + 
				"FROM " + this.m_strCondTable.Trim() + " c, " + 
				this.m_strPlotTable.Trim() + " p, "  + 
				Tables.CoreScenarioResults.DefaultScenarioResultsCycle1EffectiveTableName + " e " + 
				"WHERE c.biosum_plot_id = p.biosum_plot_id  AND " + 
				"e.biosum_cond_id = c.biosum_cond_id AND " + 
				"e.overall_effective_yn = 'Y' AND e.rxcycle='1' AND " + 
				"(p.merch_haul_cost_id IS NOT NULL OR " + 
				"p.chip_haul_cost_id IS NOT NULL);";

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert condition records that have MERCH or CHIP haul costs into best_rx_summary--\r\n");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + this.m_strSQL + "\r\n\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

           

			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (m_ado.TableExist(this.m_TempMDBFileConn,"cycle1_effective_product_yields_net_rev_costs_summary_by_rx"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_effective_product_yields_net_rev_costs_summary_by_rx");
			}

			m_ado.m_strSQL = "SELECT p.* " + 
				"INTO cycle1_effective_product_yields_net_rev_costs_summary_by_rx " + 
				"FROM product_yields_net_rev_costs_summary_by_rx p,cycle1_effective e " + 
				"WHERE p.biosum_cond_id = e.biosum_cond_id AND " + 
				"p.rxpackage=e.rxpackage AND p.rx=e.rx AND p.rxcycle=e.rxcycle AND e.overall_effective_yn='Y'";

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--write overall effective treatments to the  cycle1_effective_product_yields_net_rev_costs_summary_by_rx--\r\n");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
			
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (m_ado.TableExist(this.m_TempMDBFileConn,"cycle1_effective_optimization_treatments"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_effective_optimization_treatments");
			}

			m_ado.m_strSQL = "SELECT e.* " + 
				"INTO cycle1_effective_optimization_treatments " + 
				"FROM cycle1_effective e " + 
				"WHERE e.overall_effective_yn='Y'";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--write overall effective treatments to the cycle1_effective_optimization_treatments--\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
			{
				Cycle1_best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,false);

			}
			else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
			{
				Cycle1_best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,false);
			}
			else
			{
				Cycle1_best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,true);
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/*************************************************************************************
			 **AIR CURTAIN DESTRUCTION PLOTS
			 **insert records into the air curtain destruction table that were processed in the 
			 **best_rx_summary table. The records to insert will be plots that do not 
			 **have a place to transport the chips (biomass) so they must be burned on site
			 *************************************************************************************/
            //check to see if there are any air destruction stands
            m_strSQL = "SELECT COUNT(*) AS ROWCOUNT " +
                "FROM " + this.m_strCondTable.Trim() + " c, " +
                this.m_strPlotTable.Trim() + " p, " +
                "cycle1_best_rx_summary e " +
                "WHERE c.biosum_plot_id = p.biosum_plot_id  AND " +
                "e.biosum_cond_id = c.biosum_cond_id AND " +
                "(p.chip_haul_cost_id IS NULL);";
            if ((int)this.m_ado.getSingleDoubleValueFromSQLQuery(this.m_TempMDBFileConn, this.m_strSQL, "temp") > 0)
            {

                /**************************************************************************************
                **insert air destruction plots with no haul costs
                **************************************************************************************/

                this.m_strSQL = "INSERT INTO cycle1_best_rx_summary_air_dest " +
                    "SELECT DISTINCT c.biosum_cond_id,c.acres,c.owngrpcd,e.optimization_value,e.tiebreaker_value,e.rx_intensity,e.rx " +
                    "FROM " + this.m_strCondTable.Trim() + " c, " +
                    this.m_strPlotTable.Trim() + " p, " +
                    "cycle1_best_rx_summary e " +
                    "WHERE c.biosum_plot_id = p.biosum_plot_id  AND " +
                    "e.biosum_cond_id = c.biosum_cond_id AND " +
                    "(p.chip_haul_cost_id IS NULL);";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--insert air curtain destruction plots from the  best_rx_summary table to the best_rx_summary_air_dest--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + this.m_strSQL + "\r\n\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }
			
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}


            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/****************************************************************************
			 **finished with minimum merchantable wood removal with positive net revenue
			 ****************************************************************************/


			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
			}


		
		}

		private void Cycle1_best_rx_summary(
			FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection, 
			string strTieBreakerAggregate,
			bool bFVSVariable)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Cycle1_best_rx_summary\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n\r\n");

                frmMain.g_oUtils.WriteText(m_strDebugFile, "Parameters\r\n-------------\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "TieBreaker_Collection object\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "TieBreakerAggregate=" + strTieBreakerAggregate + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "FVS Variable as Tie Breaker = ");
                if (bFVSVariable)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Yes\r\n");
                else
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "No\r\n");

                

            }
			string strSql="";


			if (bFVSVariable==false)
			{
				//find the treatment for each plot that produces the MAX/MIN revenue value
				strSql = "SELECT a.biosum_cond_id,a.rxpackage,a.rxpackage,a.rx,a." + this.m_strOptimizationColumnNameSql + " AS optimization_value " + //LPOTTS,a.rx_intensity " + 
					"FROM cycle1_optimization a,";


				strSql=strSql + "(SELECT " + this.m_strOptimizationAggregateSql + "(" + this.m_strOptimizationColumnNameSql + ") AS " + this.m_strOptimizationAggregateColumnName + ",biosum_cond_id " + 
					"FROM cycle1_optimization";

				strSql = strSql + " GROUP BY biosum_cond_id) b ";

				strSql = strSql + "WHERE a.biosum_cond_id=b.biosum_cond_id AND a." + this.m_strOptimizationColumnNameSql + " = b." + this.m_strOptimizationAggregateColumnName;
			}
			else
			{
				strSql = "SELECT a.biosum_cond_id,a.rxpackage,a.rx,a." + this.m_strOptimizationColumnNameSql + " AS optimization_value " + //LPOTTS,a.rx_intensity " + 
					"FROM cycle1_optimization a,";


				strSql=strSql + "(SELECT " + this.m_strOptimizationAggregateSql + "(" + this.m_strOptimizationColumnNameSql + ") AS " + this.m_strOptimizationAggregateColumnName + ",biosum_cond_id " + 
					"FROM cycle1_optimization";

				strSql = strSql + " GROUP BY biosum_cond_id) b ";

				strSql = strSql + "WHERE a.biosum_cond_id=b.biosum_cond_id AND a." + this.m_strOptimizationColumnNameSql + " = b." + this.m_strOptimizationAggregateColumnName;
			}

			strSql = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " + 
				strSql;
		
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--filter effective treatments to find " + this.m_strOptimizationAggregateSql + " " + this.m_oOptimizationVariable.strOptimizedVariable + "--\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + strSql + "\r\n\r\n");
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSql);
			if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!!!Error Executing SQL!!!\r\n");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

				

			m_ado.m_strSQL = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
				"INNER JOIN cycle1_best_rx_summary_work_table b " + 
				"ON a.biosum_cond_id=b.biosum_cond_id " + 
				"SET a.acres = b.acres,a.owngrpcd=b.owngrpcd";
           if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + m_ado.m_strSQL + "\r\n");
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			if (oTieBreakerCollection.Item(0).bSelected && 
				oTieBreakerCollection.Item(1).bSelected)
			{
				//update the tiebreaker and rx intensity fields for each plot
				if (oTieBreakerCollection.Item(0).strValueSource=="POST")
				{
					strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.post_variable1_value,a.rx_intensity=b.rx_intensity";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n");
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
						
				}
				else if (oTieBreakerCollection.Item(0).strValueSource=="PRE")
				{
					strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.pre_variable1_value,a.rx_intensity=b.rx_intensity";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n");
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
				}
				else if (oTieBreakerCollection.Item(0).strValueSource=="POST-PRE")
				{
					strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.variable1_change,a.rx_intensity=b.rx_intensity";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n");
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
				}

				m_ado.m_strSQL ="INSERT INTO cycle1_best_rx_summary_before_tiebreaks SELECT DISTINCT * FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);


				//find the treatment for each plot that produces the MAX/MIN tiebreaker value
				m_ado.m_strSQL ="SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value,a.tiebreaker_value,a.rxpackage,a.rx,a.rx_intensity " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a," +
					"(SELECT biosum_cond_id," + strTieBreakerAggregate + "(tiebreaker_value) AS tiebreaker " +
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.tiebreaker_value=c.tiebreaker";


					
				m_ado.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + m_ado.m_strSQL;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--break any ties by finding the " + strTieBreakerAggregate + " tie breaker value--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				
                if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
				
				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}
					


				m_ado.m_strSQL = "SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value," + 
					"a.tiebreaker_value,a.rxpackage,a.rx,a.rx_intensity " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 a," +
					"(SELECT biosum_cond_id,MIN(rx_intensity) AS min_intensity " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.rx_intensity=c.min_intensity";

				m_ado.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3 " + m_ado.m_strSQL;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--break any additional ties by finding the least intense treatment--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
				
				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}
					

				m_ado.m_strSQL = "UPDATE cycle1_best_rx_summary_work_table a " + 
					"INNER JOIN cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3 b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id " + 
					"SET a.optimization_value=b.optimization_value," + 
					"a.tiebreaker_value=b.tiebreaker_value," +
                    "a.rxpackage=b.rxpackage," + 
					"a.rx=b.rx," + 
					"a.rx_intensity=b.rx_intensity";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
				m_ado.m_strSQL="INSERT INTO cycle1_best_rx_summary SELECT * FROM cycle1_best_rx_summary_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert the work table records into the cycle1_best_rx_summary table--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
				
				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}

			}
			else if (oTieBreakerCollection.Item(0).bSelected)
			{
				//update the tiebreaker and rx intensity fields for each plot
				if (oTieBreakerCollection.Item(0).strValueSource=="POST")
				{
					strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.post_variable1_value,a.rx_intensity=b.rx_intensity";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n\r\n");
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
						
				}
				else if (oTieBreakerCollection.Item(0).strValueSource=="PRE")
				{
					strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.pre_variable1_value,a.rx_intensity=b.rx_intensity";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n\r\n");
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
				}
				else if (oTieBreakerCollection.Item(0).strValueSource=="POST-PRE")
				{
					strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx  " + 
						"SET a.tiebreaker_value = b.variable1_change,a.rx_intensity=b.rx_intensity";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n\r\n");
                        
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
				}

				m_ado.m_strSQL ="INSERT INTO cycle1_best_rx_summary_before_tiebreaks SELECT DISTINCT * FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);


				//find the treatment for each plot that produces the MAX/MIN tiebreaker value
				m_ado.m_strSQL ="SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value,a.tiebreaker_value,a.rxpackage,a.rx,a.rx_intensity " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a," +
					"(SELECT biosum_cond_id," + strTieBreakerAggregate + "(tiebreaker_value) AS tiebreaker " +
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.tiebreaker_value=c.tiebreaker";


					
				m_ado.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + m_ado.m_strSQL;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--break any ties by finding the " + strTieBreakerAggregate + " tie breaker value--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
				
				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}
				m_ado.m_strSQL = "UPDATE cycle1_best_rx_summary_work_table a " + 
					"INNER JOIN cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id " + 
					"SET a.optimization_value=b.optimization_value," + 
					"a.tiebreaker_value=b.tiebreaker_value," + 
                    "a.rxpackage=b.rxpackage," +
					"a.rx=b.rx," + 
					"a.rx_intensity=b.rx_intensity";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
				m_ado.m_strSQL="INSERT INTO cycle1_best_rx_summary SELECT * FROM cycle1_best_rx_summary_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert the work table records into the cycle1_best_rx_summary table--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
				
				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}

			}
			else if (oTieBreakerCollection.Item(1).bSelected)
			{
				//update the rx intensity fields for each plot
				strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
					"INNER JOIN tiebreaker b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
					"SET a.rx_intensity=b.rx_intensity";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n");
				m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);


				m_ado.m_strSQL ="INSERT INTO cycle1_best_rx_summary_before_tiebreaks SELECT DISTINCT * FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);


					
				m_ado.m_strSQL = "SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value," + 
					"a.tiebreaker_value,a.rxpackage,a.rx,a.rx_intensity " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a," +
					"(SELECT biosum_cond_id,MIN(rx_intensity) AS min_intensity " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.rx_intensity=c.min_intensity";

				m_ado.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + m_ado.m_strSQL;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--break any additional ties by finding the least intense treatment--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;

				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}
					

				m_ado.m_strSQL = "UPDATE cycle1_best_rx_summary_work_table a " + 
					"INNER JOIN cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id " + 
					"SET a.optimization_value=b.optimization_value," + 
					"a.tiebreaker_value=b.tiebreaker_value," +
                    "a.rxpackage=b.rxpackage," +
					"a.rx=b.rx," + 
					"a.rx_intensity=b.rx_intensity";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
				m_ado.m_strSQL="INSERT INTO cycle1_best_rx_summary SELECT * FROM cycle1_best_rx_summary_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert the work table records into the cycle1_best_rx_summary table--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;

				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}

			}


		}
	
		

		/// <summary>
		/// expand the wood volume,value,and costs by plot acreage for the 
		/// best treatments and data found in the best_rx_summary
		/// product_yields_net_rev_costs_summary_by_rx, and effective tables
		/// </summary>
        private void Cycle1BestTreatmentsByPlot()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Cycle1BestTreatmentsByPlot\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n\r\n");
            }
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg,"Text", "Best Treatment Acreage Expansion By Plot");
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
			
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n\r\nBest Treatment Acreage Expansion By Plot\r\n");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"-----------------------------------------\r\n");

            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                    ReferenceUserControlScenarioRun.listViewEx1, "Cycle 1: Summarize The Best Effective Treatment Yields, Revenue, Costs, And Acres By Stand");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 2;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


			            
			/*************************************************
			 **best maximum net revenue by plot
			 *************************************************/

			 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Best " + this.m_strOptimizationTableName + " By Plot\r\n");

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Insert optimization totals by plot\r\n");

			
			string strWhereExpr="";
			if (this.m_oOptimizationVariable.bUseFilter)
			{
					strWhereExpr = "max_nr_dpa " + this.m_oOptimizationVariable.strFilterOperator + " " + 
						Convert.ToString(this.m_oOptimizationVariable.dblFilterValue);
			}

			this.Cycle1BestRxAcreageExpansionTableInsert(m_strOptimizationTableName + "_stands","biosum_cond_id","rx",strWhereExpr);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/*************************************************
			 **Finished with best maximum net revenue by plot
			 *************************************************/
			/************************************************
			 **air curtain destruction
			 **no net revenue for air curtain destruction
			 **so just some up values
			 ************************************************/
			this.BestRxAcreageExpansionTableInsertForAirCurtainDestruction(m_strOptimizationTableName + "_stands_air_dest","biosum_cond_id","rx","");
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			

			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
               
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			
			/*************************************************
			 **Finished with best maximum net revenue by plot
			 *************************************************/

			
			
			
			/*******************************************************************************
			 **Finished with best crown index improvement with positive net revenue by plot
			 *******************************************************************************/




			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
                return;
			}

		}

        /// <summary>
        /// expand the wood volume,value,and costs by plot acreage for all
        /// treatments and data found in the best_rx_summary
        /// product_yields_net_rev_costs_summary_by_rx, and effective tables
        /// </summary>
        private void TreatmentAcreageExpansionByPlot()
        {

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//TreatmentAcreageExpansionByPlot\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n\r\n");
            }

            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "All Treatments Acreage Expansion By Plot");
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nll Treatments Acreage Expansion By Plot\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "-----------------------------------------\r\n");

            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                   ReferenceUserControlScenarioRun.listViewEx1, "All Treatments: Summarize Treatment Yields, Revenue, Costs, And Acres By Stand");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 2;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


            /*************************************************
             **maximum net revenue by plot
             *************************************************/
            
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Insert acreage expansion totals by plot for each treatment--\r\n");
            string strInsertSql = "INSERT INTO " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +
                                  "(biosum_cond_id, rxpackage,rx,rxcycle," +
                                   "owngrpcd,acres,merch_haul_cost_psite," +
                                   "merch_haul_cost_exp,merch_vol_cf_exp,merch_dollars_val_exp," +
                                   "chip_haul_cost_psite,chip_haul_cost_exp,chip_yield_gt_exp," +
                                   "chip_dollars_val_exp,net_rev_dollars_exp,harv_costs_exp,haul_costs_exp)";
            this.m_strSQL = "SELECT  " +
                            "product_yields_net_rev_costs_summary_by_rx.biosum_cond_id," +
                            "product_yields_net_rev_costs_summary_by_rx.rxpackage," +
                            "product_yields_net_rev_costs_summary_by_rx.rx," +
                            "product_yields_net_rev_costs_summary_by_rx.rxcycle," +
                             this.m_strCondTable + ".owngrpcd," +
                            "ROUND(" + this.m_strCondTable + ".acres,10) AS acres," +
                             this.m_strPlotTable + ".merch_haul_cost_psite," +
                            "ROUND(" + this.m_strPlotTable + ".merch_haul_cpa_pt * product_yields_net_rev_costs_summary_by_rx.merch_yield_gt * acres,10) AS merch_haul_cost_exp," +
                            "ROUND(product_yields_net_rev_costs_summary_by_rx.merch_yield_cf * acres,10) AS merch_vol_cf_exp," +
                            "ROUND(product_yields_net_rev_costs_summary_by_rx.merch_val_dpa * acres,10) AS merch_dollars_val_exp," +
                             this.m_strPlotTable + ".chip_haul_cost_psite," +
                            "ROUND(" + this.m_strPlotTable + ".chip_haul_cpa_pt * product_yields_net_rev_costs_summary_by_rx.chip_yield_gt * acres,10) AS chip_haul_cost_exp," +
                            "ROUND(product_yields_net_rev_costs_summary_by_rx.chip_yield_gt * acres,10) AS chip_yield_gt_exp," +
                            "ROUND(product_yields_net_rev_costs_summary_by_rx.chip_val_dpa * acres,10) AS chip_dollars_val_exp," +
                            "ROUND(product_yields_net_rev_costs_summary_by_rx.max_nr_dpa * acres,10) AS net_rev_dollars_exp," +
                            "ROUND(product_yields_net_rev_costs_summary_by_rx.harvest_onsite_cpa * acres,10) AS harv_costs_exp," +
                            "ROUND((product_yields_net_rev_costs_summary_by_rx.haul_merch_cpa + product_yields_net_rev_costs_summary_by_rx.haul_chip_cpa) * acres,10) AS haul_costs_exp " +
                            "FROM  product_yields_net_rev_costs_summary_by_rx " +
                            "INNER JOIN (" + this.m_strPlotTable + " " +
                            "INNER JOIN " + this.m_strCondTable + " " +
                            "ON " + this.m_strPlotTable + ".biosum_plot_id = " +
                                    this.m_strCondTable + ".biosum_plot_id) " +
                            "ON " + this.m_strCondTable + ".biosum_cond_id = " +
                                   "product_yields_net_rev_costs_summary_by_rx.biosum_cond_id;";

            m_strSQL = strInsertSql + " " + m_strSQL;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
            }

            

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Sum acreage expansion totals by plot for a Treatment Package--\r\n");
            strInsertSql = "INSERT INTO " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxPackageCostRevenueVolumesSumTableName + " " +
                                  "(biosum_cond_id, rxpackage," +
                                   "owngrpcd,acres,merch_haul_cost_psite," +
                                   "merch_haul_cost_sum,merch_vol_cf_sum,merch_dollars_val_sum," +
                                   "chip_haul_cost_psite,chip_haul_cost_sum,chip_yield_gt_sum," +
                                   "chip_dollars_val_sum,net_rev_dollars_sum,harv_costs_sum,haul_costs_sum) ";

            this.m_strSQL = "SELECT DISTINCT " +
                                "b.biosum_cond_id,b.rxpackage," +
                                "a.owngrpcd,c.acres," +
                                "a.merch_haul_cost_psite," +
                                "b.merch_haul_cost_sum," +
                                "b.merch_vol_cf_sum," +
                                "b.merch_dollars_val_sum," +
                                "a.chip_haul_cost_psite," + 
                                "b.chip_haul_cost_sum," +
                                "b.chip_yield_gt_sum," +
                                "b.chip_dollars_val_sum," +
                                "b.net_rev_dollars_sum," +
                                "b.harv_costs_sum," +
                                "b.haul_costs_sum " +
                            "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " a,";
                            
            string strSumSQL ="(SELECT  " +
                           "biosum_cond_id," +
                           "rxpackage," +
                           "SUM(IIF(merch_haul_cost_exp IS NULL,0,merch_haul_cost_exp)) AS merch_haul_cost_sum," +
                           "SUM(IIF(merch_vol_cf_exp IS NULL,0,merch_vol_cf_exp)) AS merch_vol_cf_sum," +
                           "SUM(IIF(merch_dollars_val_exp IS NULL,0,merch_dollars_val_exp)) AS merch_dollars_val_sum," +
                           "SUM(IIF(chip_haul_cost_exp IS NULL,0,chip_haul_cost_exp)) AS chip_haul_cost_sum," +
                           "SUM(IIF(chip_yield_gt_exp IS NULL,0,chip_yield_gt_exp)) AS chip_yield_gt_sum," +
                           "SUM(IIF(chip_dollars_val_exp IS NULL,0,chip_dollars_val_exp)) AS chip_dollars_val_sum," +
                           "SUM(IIF(net_rev_dollars_exp IS NULL,0,net_rev_dollars_exp)) AS net_rev_dollars_sum," +
                           "SUM(IIF(harv_costs_exp IS NULL,0,harv_costs_exp)) AS harv_costs_sum," +
                           "SUM(IIF(haul_costs_exp IS NULL,0,haul_costs_exp)) AS haul_costs_sum " +
                           "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " + 
                           "GROUP BY biosum_cond_id,rxpackage) b," + 
                           "(SELECT DISTINCT biosum_cond_id," + 
                                            "rxpackage," + 
                                            "acres " + 
                            "FROM stand_costs_revenue_volume_by_rx) c";


            m_strSQL = strInsertSql + m_strSQL + strSumSQL + " WHERE a.biosum_cond_id=b.biosum_cond_id AND c.biosum_cond_id=b.biosum_cond_id AND a.rxpackage=b.rxpackage AND c.rxpackage=b.rxpackage;";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
            }

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
                return;
            }

        }
        /// <summary>
        /// expand the wood volume,value,and costs by plot acreage for all
        /// treatments and data found in the best_rx_summary
        /// product_yields_net_rev_costs_summary_by_rx, and effective tables
        /// </summary>
        private void TreatmentAcreageExpansionByPSite()
        {
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "All Treatments Acreage Expansion By Processing Site");
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

            //this.m_txtStreamWriter.WriteLine(" ");
            //this.m_txtStreamWriter.WriteLine(" ");
            //this.m_txtStreamWriter.WriteLine("All Treatments Acreage Expansion By Processing Site");
            //this.m_txtStreamWriter.WriteLine("---------------------------------------------------");

            if (m_ado.TableExist(this.m_TempMDBFileConn,"psite_rxpackage_rx_rxcycle_sum_work_table"))
                m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE psite_rxpackage_rx_rxcycle_sum_work_table");

            frmMain.g_oTables.m_oCoreScenarioResults.CreatePSiteRxCostsRevenuesVolumesSumTable(m_ado, m_TempMDBFileConn, "psite_rxpackage_rx_rxcycle_sum_work_table");

            /*************************************************
             **CHIPS ONLY PSITES
             *************************************************/
            

           // this.m_txtStreamWriter.WriteLine("Insert acreage expansion totals by CHIP Only Processing Site for each treatment");
            string strInsertSql = "INSERT INTO psite_rxpackage_rx_rxcycle_sum_work_table " +
                                  "(rxpackage,rx,rxcycle,psite_id,biocd," +
                                   "sum_acres,sum_chip_yield,sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," +
                                   "sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs,sum_haul_costs) ";
            this.m_strSQL = "SELECT DISTINCT " + 
	                           "c.rxpackage,c.rx,c.rxcycle,a.psite_id,a.biocd," + 
	                           "c.sum_acres,c.sum_chip_yield," + 
                               "c.sum_chip_haul_cents,c.sum_chip_dollars_val," + 
                               "c.sum_merch_haul_cents,c.sum_merch_vol," + 
                               "c.sum_net_rev,c.sum_merch_dollars_val," + 
	                           "c.sum_harv_costs,c.sum_haul_costs " + 
                            "FROM scenario_psites_work_table AS a," + 
                            "(SELECT " + 
                               "biosum_cond_id," + 
                               "MIN(chip_haul_cost_psite) as cheapest_psite " + 
                             "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +  
                             "GROUP BY biosum_cond_id)  b," + 
                            "(SELECT rxpackage,rx,rxcycle,chip_haul_cost_psite," + 
	                            "ROUND(SUM(chip_yield),10)  AS sum_chip_yield," + 
	                            "ROUND(SUM(chip_haul_cents),10) AS sum_chip_haul_cents," + 
	                            "ROUND(SUM(chip_dollars_val),10) AS sum_chip_dollars_val," + 
	                            "ROUND(SUM(acres),10) AS sum_acres," + 
	                            "ROUND(SUM(merch_haul_cents),10) AS sum_merch_haul_cents," + 
	                            "ROUND(SUM(merch_vol),10) AS sum_merch_vol," + 
	                            "ROUND(SUM(net_rev),10) AS sum_net_rev," + 
	                            "ROUND(SUM(merch_dollars_val),10) AS sum_merch_dollars_val," + 
	                            "ROUND(SUM(harv_costs),10) AS sum_harv_costs," + 
	                            "ROUND(SUM(haul_costs),10) AS sum_haul_costs " + 
                             "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +  
                             "GROUP BY rxpackage,rx,rxcycle,chip_haul_cost_psite) c " +  
                             "WHERE a.psite_id=b.cheapest_psite AND " + 
                                   "a.psite_id=c.chip_haul_cost_psite AND " + 
                                   "a.biocd=2;";

            m_strSQL = strInsertSql + " " + m_strSQL;
            //this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            
            if (this.m_ado.m_intError != 0)
            {
                //this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }
            /*************************************************
             **MERCH ONLY PSITES
             *************************************************/
           // this.m_txtStreamWriter.WriteLine("Insert acreage expansion totals by MERCH Only Processing Site for each treatment");
            strInsertSql = "INSERT INTO psite_rxpackage_rx_rxcycle_sum_work_table " +
                                  "(rxpackage,rx,rxcycle,psite_id,biocd," +
                                   "sum_acres,sum_chip_yield,sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," +
                                   "sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs,sum_haul_costs) ";
            this.m_strSQL = "SELECT DISTINCT " + 
	                           "c.rxpackage,c.rx,c.rxcycle,a.psite_id,a.biocd," + 
	                           "c.sum_acres,c.sum_chip_yield," + 
                               "c.sum_chip_haul_cents,c.sum_chip_dollars_val," + 
                               "c.sum_merch_haul_cents,c.sum_merch_vol," + 
                               "c.sum_net_rev,c.sum_merch_dollars_val," + 
	                           "c.sum_harv_costs,c.sum_haul_costs " + 
                            "FROM scenario_psites_work_table AS a," + 
                            "(SELECT " + 
                               "biosum_cond_id," + 
                               "MIN(merch_haul_cost_psite) as cheapest_psite " + 
                             "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +  
                             "GROUP BY biosum_cond_id)  b," + 
                            "(SELECT rxpackage,rx,rxcycle,merch_haul_cost_psite," + 
	                            "ROUND(SUM(chip_yield),10)  AS sum_chip_yield," + 
	                            "ROUND(SUM(chip_haul_cents),10) AS sum_chip_haul_cents," + 
	                            "ROUND(SUM(chip_dollars_val),10) AS sum_chip_dollars_val," + 
	                            "ROUND(SUM(acres),10) AS sum_acres," + 
	                            "ROUND(SUM(merch_haul_cents),10) AS sum_merch_haul_cents," + 
	                            "ROUND(SUM(merch_vol),10) AS sum_merch_vol," + 
	                            "ROUND(SUM(net_rev),10) AS sum_net_rev," + 
	                            "ROUND(SUM(merch_dollars_val),10) AS sum_merch_dollars_val," + 
	                            "ROUND(SUM(harv_costs),10) AS sum_harv_costs," + 
	                            "ROUND(SUM(haul_costs),10) AS sum_haul_costs " + 
                             "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +  
                             "GROUP BY rxpackage,rx,rxcycle,merch_haul_cost_psite) c " +  
                             "WHERE a.psite_id=b.cheapest_psite AND " + 
                                   "a.psite_id=c.merch_haul_cost_psite AND " + 
                                   "a.biocd=1;";

            m_strSQL = strInsertSql + " " + m_strSQL;
            //this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            if (this.m_ado.m_intError != 0)
            {
                //this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

           
            /*************************************************
             **BOTH CHIPS AND MERCH PSITES
             *************************************************/
           // this.m_txtStreamWriter.WriteLine("Insert acreage expansion totals for processing sites that process both MERCH and CHIPS. Filter for MERCH.");
            strInsertSql = "INSERT INTO psite_rxpackage_rx_rxcycle_sum_work_table " +
                                  "(rxpackage,rx,rxcycle,psite_id,biocd," +
                                   "sum_acres,sum_chip_yield,sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," +
                                   "sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs,sum_haul_costs) ";
            this.m_strSQL = "SELECT DISTINCT " + 
	                           "c.rxpackage,c.rx,c.rxcycle,a.psite_id,a.biocd," + 
	                           "c.sum_acres,c.sum_chip_yield," + 
                               "c.sum_chip_haul_cents,c.sum_chip_dollars_val," + 
                               "c.sum_merch_haul_cents,c.sum_merch_vol," + 
                               "c.sum_net_rev,c.sum_merch_dollars_val," + 
	                           "c.sum_harv_costs,c.sum_haul_costs " + 
                            "FROM scenario_psites_work_table AS a," + 
                            "(SELECT " + 
                               "biosum_cond_id," + 
                               "MIN(merch_haul_cost_psite) as cheapest_psite " + 
                             "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +  
                             "GROUP BY biosum_cond_id)  b," + 
                            "(SELECT rxpackage,rx,rxcycle,merch_haul_cost_psite," + 
	                            "ROUND(SUM(chip_yield),10)  AS sum_chip_yield," + 
	                            "ROUND(SUM(chip_haul_cents),10) AS sum_chip_haul_cents," + 
	                            "ROUND(SUM(chip_dollars_val),10) AS sum_chip_dollars_val," + 
	                            "ROUND(SUM(acres),10) AS sum_acres," + 
	                            "ROUND(SUM(merch_haul_cents),10) AS sum_merch_haul_cents," + 
	                            "ROUND(SUM(merch_vol),10) AS sum_merch_vol," + 
	                            "ROUND(SUM(net_rev),10) AS sum_net_rev," + 
	                            "ROUND(SUM(merch_dollars_val),10) AS sum_merch_dollars_val," + 
	                            "ROUND(SUM(harv_costs),10) AS sum_harv_costs," + 
	                            "ROUND(SUM(haul_costs),10) AS sum_haul_costs " + 
                             "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +  
                             "GROUP BY rxpackage,rx,rxcycle,merch_haul_cost_psite) c " +  
                             "WHERE a.psite_id=b.cheapest_psite AND " + 
                                   "a.psite_id=c.merch_haul_cost_psite AND " + 
                                   "a.biocd=3;";

            m_strSQL = strInsertSql + " " + m_strSQL;
            //this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            if (this.m_ado.m_intError != 0)
            {
                //this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            
             /*************************************************
              **BOTH CHIPS AND MERCH PSITES THAT WERE
              **NOT PROCESSED IN THE PREVIOUS ROUTINE 
              **BECAUSE ONLY CHIPS WERE SENT TO THE PSITE
              **IN THIS BATCH PROCESS
              *************************************************/
            string strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select DISTINCT psite_id from psite_rxpackage_rx_rxcycle_sum_work_table");

            //this.m_txtStreamWriter.WriteLine("Insert acreage expansion totals for processing sites that process both MERCH and CHIPS. Filter for CHIPS.");
            strInsertSql = "INSERT INTO psite_rxpackage_rx_rxcycle_sum_work_table " +
                                  "(rxpackage,rx,rxcycle,psite_id,biocd," +
                                   "sum_acres,sum_chip_yield,sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," +
                                   "sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs,sum_haul_costs) ";

            this.m_strSQL = "SELECT DISTINCT " + 
	                           "c.rxpackage,c.rx,c.rxcycle,a.psite_id,a.biocd," + 
	                           "c.sum_acres,c.sum_chip_yield," + 
                               "c.sum_chip_haul_cents,c.sum_chip_dollars_val," + 
                               "c.sum_merch_haul_cents,c.sum_merch_vol," + 
                               "c.sum_net_rev,c.sum_merch_dollars_val," + 
	                           "c.sum_harv_costs,c.sum_haul_costs " + 
                            "FROM scenario_psites_work_table AS a," + 
                            "(SELECT " + 
                               "biosum_cond_id," + 
                               "MIN(chip_haul_cost_psite) as cheapest_psite " + 
                             "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +  
                             "GROUP BY biosum_cond_id)  b," + 
                            "(SELECT rxpackage,rx,rxcycle,chip_haul_cost_psite," + 
	                            "ROUND(SUM(chip_yield),10)  AS sum_chip_yield," + 
	                            "ROUND(SUM(chip_haul_cents),10) AS sum_chip_haul_cents," + 
	                            "ROUND(SUM(chip_dollars_val),10) AS sum_chip_dollars_val," + 
	                            "ROUND(SUM(acres),10) AS sum_acres," + 
	                            "ROUND(SUM(merch_haul_cents),10) AS sum_merch_haul_cents," + 
	                            "ROUND(SUM(merch_vol),10) AS sum_merch_vol," + 
	                            "ROUND(SUM(net_rev),10) AS sum_net_rev," + 
	                            "ROUND(SUM(merch_dollars_val),10) AS sum_merch_dollars_val," + 
	                            "ROUND(SUM(harv_costs),10) AS sum_harv_costs," + 
	                            "ROUND(SUM(haul_costs),10) AS sum_haul_costs " + 
                             "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " " +  
                             "GROUP BY rxpackage,rx,rxcycle,chip_haul_cost_psite) c " +  
                             "WHERE a.psite_id=b.cheapest_psite AND " + 
                                   "a.psite_id=c.chip_haul_cost_psite AND " + 
                                   "a.biocd=3";
            if (strPSiteList.Trim().Length > 0)
            {
            
                m_strSQL = m_strSQL + " AND " +
                                     "b.cheapest_psite NOT IN (" + strPSiteList + ")";
            }


            m_strSQL = strInsertSql + " " + m_strSQL;
           // this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            if (this.m_ado.m_intError != 0)
            {
                //this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }


            

            /*************************************************
             **APPEND WORK TABLE TO FINAL DESTINATION TABLE
             *************************************************/
            //this.m_txtStreamWriter.WriteLine("Insert acreage expansion totals for processing sites from work table to production table");
            strInsertSql = "INSERT INTO " + 
                            Tables.CoreScenarioResults.DefaultScenarioResultsPSiteRxCostRevenueVolumesTableName + " " + 
                           "(rxpackage,rx,rxcycle,psite_id,acres," + 
                            "merch_haul_cents,chip_haul_cents," + 
                            "merch_vol,chip_yield,net_rev," + 
                            "merch_dollars_val,chip_dollars_val," + 
                            "harv_costs,haul_costs)";
            this.m_strSQL = "SELECT " + 
	                            "rxpackage,rx,rxcycle,psite_id," + 
	                            "sum_acres AS acres," + 
	                            "sum_merch_haul_cents AS merch_haul_cents," + 
	                            "sum_chip_haul_cents AS chip_haul_cents," + 
	                            "sum_merch_vol AS merch_vol," + 
	                            "sum_chip_yield AS chip_yield," + 
	                            "sum_net_rev AS net_rev," + 
	                            "sum_merch_dollars_val AS merch_dollars_val," + 
	                            "sum_chip_dollars_val AS chip_dollars_val," + 
	                            "sum_harv_costs AS harv_costs," + 
	                            "sum_haul_costs AS haul_costs " + 
                            "FROM psite_rxpackage_rx_rxcycle_sum_work_table;";


            m_strSQL = strInsertSql + " " + m_strSQL;
            //this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            if (this.m_ado.m_intError != 0)
            {
                //this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            

            //this.m_txtStreamWriter.WriteLine("Sum acreage expansion totals by psiteid for a Treatment Package");
            strInsertSql = "INSERT INTO " + Tables.CoreScenarioResults.DefaultScenarioResultsPSiteRxPackageCostRevenueVolumesTableName + " " +
                                  "(rxpackage,psite_id," +
                                   "acres,merch_haul_cost_psite," +
                                   "merch_haul_cents,merch_vol,merch_dollars_val," +
                                   "chip_haul_cost_psite,chip_haul_cents,chip_yield," +
                                   "chip_dollars_val,net_rev,harv_costs,haul_costs) ";

            this.m_strSQL = "SELECT DISTINCT " +
                                "b.rxpackage," +
                                "b.psite_id," + 
                                "b.sum_acres AS acres," +
                                "a.merch_haul_cost_psite," +
                                "b.sum_merch_haul_cents AS merch_haul_cents," +
                                "b.sum_merch_vol AS merch_vol," +
                                "b.sum_merch_dollars_val AS merch_dollars_val," +
                                "a.chip_haul_cost_psite," +
                                "b.sum_chip_haul_cents AS chip_haul_cents," +
                                "b.sum_chip_yield AS chip_yield," +
                                "b.sum_chip_dollars_val AS chip_dollars_val," +
                                "b.sum_net_rev AS net_rev," +
                                "b.sum_harv_costs AS harv_costs," +
                                "b.sum_haul_costs AS haul_costs " +
                            "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPSiteRxCostRevenueVolumesTableName + " a,";

            string strSumSQL = "(SELECT  " +
                           "rxpackage," +
                           "psite_id," + 
                           "SUM(IIF(acres IS NULL,0,acres)) AS sum_acres," +
                           "SUM(IIF(merch_haul_cents IS NULL,0,merch_haul_cents)) AS sum_merch_haul_cents," +
                           "SUM(IIF(merch_vol IS NULL,0,merch_vol)) AS sum_merch_vol," +
                           "SUM(IIF(merch_dollars_val IS NULL,0,merch_dollars_val)) AS sum_merch_dollars_val," +
                           "SUM(IIF(chip_haul_cents IS NULL,0,chip_haul_cents)) AS sum_chip_haul_cents," +
                           "SUM(IIF(chip_yield IS NULL,0,chip_yield)) AS sum_chip_yield," +
                           "SUM(IIF(chip_dollars_val IS NULL,0,chip_dollars_val)) AS sum_chip_dollars_val," +
                           "SUM(IIF(net_rev IS NULL,0,net_rev)) AS sum_net_rev," +
                           "SUM(IIF(harv_costs IS NULL,0,harv_costs)) AS sum_harv_costs," +
                           "SUM(IIF(haul_costs IS NULL,0,haul_costs)) AS sum_haul_costs " +
                           "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPSiteRxCostRevenueVolumesTableName + " " +
                           "GROUP BY rxpackage,psite_id) b";

            m_strSQL = strInsertSql + m_strSQL + strSumSQL + " WHERE a.rxpackage=b.rxpackage AND a.psite_id=b.psite_id";

            //this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            if (this.m_ado.m_intError != 0)
            {
                //this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }
            if (this.m_intError == 0)
            {
                return;
            }

        }
        /// <summary>
        /// expand the wood volume,value,and costs by plot acreage for all
        /// treatments and data found in the best_rx_summary
        /// product_yields_net_rev_costs_summary_by_rx, and effective tables
        /// </summary>
        private void TreatmentAcreageExpansionByOwner()
        {
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "All Treatments Acreage Expansion By Ownership");
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

            //this.m_txtStreamWriter.WriteLine(" ");
            //this.m_txtStreamWriter.WriteLine(" ");
            //this.m_txtStreamWriter.WriteLine("All Treatments Acreage Expansion By Ownership");
            //this.m_txtStreamWriter.WriteLine("---------------------------------------------");


            /*************************************************
             **maximum net revenue by ownership
             *************************************************/
            

            if (m_ado.TableExist(this.m_TempMDBFileConn, "own_rxpackage_rx_rxcycle_sum_work_table"))
                m_ado.SqlNonQuery(m_TempMDBFileConn, "DROP TABLE own_rxpackage_rx_rxcycle_sum_work_table");

            frmMain.g_oTables.m_oCoreScenarioResults.CreateOwnerRxCostsRevenuesVolumesSumTable(m_ado, m_TempMDBFileConn, "own_rxpackage_rx_rxcycle_sum_work_table");

            //this.m_txtStreamWriter.WriteLine("Insert acreage expansion totals by ownership for each treatment");
            this.m_strSQL = "INSERT INTO own_rxpackage_rx_rxcycle_sum_work_table " +  
                "(rxpackage,rx,rxcycle,owngrpcd," + 
                 "sum_acres,sum_chip_yield," +
                 "sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," +
                 "sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs," +
                 "sum_haul_costs) " +
                "SELECT rxpackage,rx,rxcycle,owngrpcd," +
                "SUM(acres) AS sum_acres, " +
                "SUM(chip_yield)  AS sum_chip_yield," +
                "SUM(IIF(chip_haul_cents IS NOT NULL,chip_haul_cents,0)) AS sum_chip_haul_cents," +
                "SUM(IIF(chip_dollars_val IS NOT NULL,chip_dollars_val,0)) AS sum_chip_dollars_val, " +
                "SUM(IIF(merch_haul_cents IS NOT NULL,merch_haul_cents,0)) AS sum_merch_haul_cents," +
                "SUM(merch_vol) AS sum_merch_vol," +
                "SUM(IIF(net_rev IS NOT NULL,net_rev,0)) AS sum_net_rev," +
                "SUM(IIF(merch_dollars_val IS NOT NULL,merch_dollars_val,0)) AS sum_merch_dollars_val," +
                "SUM(harv_costs) AS sum_harv_costs," +
                "SUM(IIF(haul_costs IS NOT NULL,haul_costs,0)) AS sum_haul_costs " +
                "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsPlotRxCostRevenueVolumesTableName + " GROUP BY rxpackage,rx,rxcycle,owngrpcd ;";

            //this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            if (this.m_ado.m_intError != 0)
            {
               // this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

           
            
            this.m_strSQL = "INSERT INTO " + Tables.CoreScenarioResults.DefaultScenarioResultsOwnerRxCostRevenueVolumesTableName + " " +
                "(rxpackage,rx,rxcycle,owngrpcd,acres,merch_haul_cents,chip_haul_cents," +
                "merch_vol,chip_yield,net_rev,merch_dollars_val," +
                "chip_dollars_val,harv_costs,haul_costs) " +
                "SELECT rxpackage,rx,rxcycle,owngrpcd," +
                "sum_acres AS acres," +
                "sum_merch_haul_cents AS merch_haul_cents," +
                "sum_chip_haul_cents AS chip_haul_cents," +
                "sum_merch_vol AS merch_vol," +
                "sum_chip_yield AS chip_yield," +
                "sum_net_rev AS net_rev," +
                "sum_merch_dollars_val AS merch_dollars_val," +
                "sum_chip_dollars_val AS chip_dollars_val," +
                "sum_harv_costs AS harv_costs," +
                "sum_haul_costs AS haul_costs " +
                "FROM own_rxpackage_rx_rxcycle_sum_work_table";
           // this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            if (this.m_ado.m_intError != 0)
            {
               // this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }
            

           // this.m_txtStreamWriter.WriteLine("Sum acreage expansion totals by ownership for a Treatment Package");
            string strInsertSql = "INSERT INTO " + Tables.CoreScenarioResults.DefaultScenarioResultsOwnerRxPackageCostRevenueVolumesTableName + " " +
                                  "(rxpackage,owngrpcd," +
                                   "acres,merch_haul_cost_psite," +
                                   "merch_haul_cents,merch_vol,merch_dollars_val," +
                                   "chip_haul_cost_psite,chip_haul_cents,chip_yield," +
                                   "chip_dollars_val,net_rev,harv_costs,haul_costs) ";

            this.m_strSQL = "SELECT DISTINCT " +
                                "b.rxpackage," +
                                "a.owngrpcd," +
                                "b.sum_acres AS acres," +
                                "null AS merch_haul_cost_psite," +
                                "b.sum_merch_haul_cents AS merch_haul_cents," +
                                "b.sum_merch_vol AS merch_vol," +
                                "b.sum_merch_dollars_val AS merch_dollars_val," +
                                "null AS chip_haul_cost_psite," +
                                "b.sum_chip_haul_cents AS chip_haul_cents," +
                                "b.sum_chip_yield AS chip_yield," +
                                "b.sum_chip_dollars_val AS chip_dollars_val," +
                                "b.sum_net_rev AS net_rev," +
                                "b.sum_harv_costs AS harv_costs," +
                                "b.sum_haul_costs AS haul_costs " +
                            "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsOwnerRxCostRevenueVolumesTableName + " a,";

            string strSumSQL = "(SELECT  " +
                           "rxpackage," +
                           "owngrpcd," +
                           "SUM(IIF(acres IS NULL,0,acres)) AS sum_acres," +
                           "SUM(IIF(merch_haul_cents IS NULL,0,merch_haul_cents)) AS sum_merch_haul_cents," +
                           "SUM(IIF(merch_vol IS NULL,0,merch_vol)) AS sum_merch_vol," +
                           "SUM(IIF(merch_dollars_val IS NULL,0,merch_dollars_val)) AS sum_merch_dollars_val," +
                           "SUM(IIF(chip_haul_cents IS NULL,0,chip_haul_cents)) AS sum_chip_haul_cents," +
                           "SUM(IIF(chip_yield IS NULL,0,chip_yield)) AS sum_chip_yield," +
                           "SUM(IIF(chip_dollars_val IS NULL,0,chip_dollars_val)) AS sum_chip_dollars_val," +
                           "SUM(IIF(net_rev IS NULL,0,net_rev)) AS sum_net_rev," +
                           "SUM(IIF(harv_costs IS NULL,0,harv_costs)) AS sum_harv_costs," +
                           "SUM(IIF(haul_costs IS NULL,0,haul_costs)) AS sum_haul_costs " +
                           "FROM " + Tables.CoreScenarioResults.DefaultScenarioResultsOwnerRxCostRevenueVolumesTableName + " " +
                           "GROUP BY rxpackage,owngrpcd) b";

            m_strSQL = strInsertSql + m_strSQL + strSumSQL + " WHERE a.rxpackage=b.rxpackage AND a.owngrpcd=b.owngrpcd";
            //this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
            if (this.m_ado.m_intError != 0)
            {
                //this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }
            


            if (this.m_intError == 0)
            {
                return;
            }

        }

		private void DeleteScenarioResultRecords()
		{

			
			ado_data_access oAdo = new ado_data_access();
			string strConn=oAdo.getMDBConnString(this.m_strSystemResultsDbPathAndFile,"admin","");
			oAdo.OpenConnection(strConn);

			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_nr_plots"))
			{
				this.m_strSQL = "delete from max_nr_plots";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_pnr_plots"))
			{
				this.m_strSQL = "delete from max_pnr_plots";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_plots"))
			{
				this.m_strSQL = "delete from max_ti_imp_plots";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_pnr_plots"))
			{
				this.m_strSQL = "delete from max_ti_imp_pnr_plots";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_plots"))
			{
				this.m_strSQL = "delete from max_ci_imp_plots";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_pnr_plots"))
			{
				this.m_strSQL = "delete from max_ci_imp_pnr_plots";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_plots"))
			{
				this.m_strSQL = "delete from min_merch_plots";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_pnr_plots"))
			{
				this.m_strSQL = "delete from min_merch_pnr_plots";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}


			/*************************************************
			 **delete all records in the by ownership tables
			 *************************************************/
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_nr_sum_own"))
			{
				this.m_strSQL = "delete from max_nr_sum_own";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_pnr_sum_own"))
			{
				this.m_strSQL = "delete from max_pnr_sum_own";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_sum_own"))
			{
				this.m_strSQL = "delete from max_ti_imp_sum_own";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_pnr_sum_own"))
			{
				this.m_strSQL = "delete from max_ti_imp_pnr_sum_own";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_sum_own"))
			{
				this.m_strSQL = "delete from max_ci_imp_sum_own";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_pnr_sum_own"))
			{
				this.m_strSQL = "delete from max_ci_imp_pnr_sum_own";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_sum_own"))
			{
				this.m_strSQL = "delete from min_merch_sum_own";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_pnr_sum_own"))
			{
				this.m_strSQL = "delete from min_merch_pnr_sum_own";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}

			/******************************************************
			 **delete all records in the by processing site tables
			 ******************************************************/
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_nr_sum_psite"))
			{
				this.m_strSQL = "delete from max_nr_sum_psite";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_pnr_sum_psite"))
			{
				this.m_strSQL = "delete from max_pnr_sum_psite";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_sum_psite"))
			{
				this.m_strSQL = "delete from max_ti_imp_sum_psite";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_pnr_sum_psite"))
			{
				this.m_strSQL = "delete from max_ti_imp_pnr_sum_psite";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_sum_psite"))
			{
				this.m_strSQL = "delete from max_ci_imp_sum_psite";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_imp_pnr_sum_psite"))
			{
				this.m_strSQL = "delete from max_ci_imp_pnr_sum_psite";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_sum_psite"))
			{
				this.m_strSQL = "delete from min_merch_sum_psite";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_pnr_sum_psite"))
			{
				this.m_strSQL = "delete from min_merch_pnr_sum_psite";
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			}

			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oAdo=null;



		}
		
		private void Cycle1BestRxAcreageExpansionTableInsert(string strTable,string strTypeField, string strRxField, string strWhereExpression)
		{
			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();
            string e = "cycle1_effective_product_yields_net_rev_costs_summary_by_rx";

           
            //create work table to eliminate duplicates
            frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxStandOptimizationSummaryCycle1Table(
                m_ado,
                m_TempMDBFileConn,
                "cycle1_effective_product_yields_by_rx_work_table");

            this.m_strSQL = "INSERT INTO cycle1_effective_product_yields_by_rx_work_table " +
                "(biosum_cond_id,rx,chip_yield_cf,merch_yield_cf," +
                "chip_yield_gt,merch_yield_gt," +
                "chip_val_dpa,merch_val_dpa," +
                "harvest_onsite_cpa,haul_chip_cpa,haul_merch_cpa," +
                "merch_chip_nr_dpa,merch_nr_dpa,max_nr_dpa," +
                "usebiomass_yn) ";

            this.m_strSQL += " SELECT a.biosum_cond_id," +
                                     "a.rx," +
                                     "a.chip_yield_cf," +
                                     "a.merch_yield_cf," +
                                     "a.chip_yield_gt," +
                                     "a.merch_yield_gt," +
                                     "a.chip_val_dpa," +
                                     "a.merch_val_dpa," +
                                     "a.harvest_onsite_cpa," +
                                     "a.haul_chip_cpa," +
                                     "a.haul_merch_cpa," +
                                     "a.merch_chip_nr_dpa," +
                                     "a.merch_nr_dpa," + 
                                     "a.max_nr_dpa," +
                                     "a.usebiomass_yn " +
                           "FROM cycle1_effective_product_yields_net_rev_costs_summary_by_rx a";

           
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + this.m_strSQL + "\r\n\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            //create duplicate table
            m_strSQL = "SELECT * INTO cycle1_effective_product_yields_by_rx_work_table2 FROM cycle1_effective_product_yields_by_rx_work_table";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + this.m_strSQL + "\r\n\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);


            //one record per stand + rx
            m_strSQL = "SELECT a.* INTO cycle1_effective_product_yields_by_rx_work_table3 " + 
                       "FROM cycle1_effective_product_yields_by_rx_work_table a," +
                         "(SELECT MIN(id) AS min_id,biosum_cond_id,rx " +
                          "FROM cycle1_effective_product_yields_by_rx_work_table2 " + 
                          "GROUP BY biosum_cond_id,rx) b " +
                       "WHERE a.id=b.min_id AND a.biosum_cond_id = b.biosum_cond_id AND a.rx = b.rx";

            
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + this.m_strSQL + "\r\n\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);



           


			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(" + strTypeField + ", owngrpcd,acres,optimization_value,merch_haul_cost_psite," + 
				"merch_haul_cpa,merch_vol_cf_pa,merch_dollars_val_dpa," + 
				"chip_haul_cost_psite,chip_haul_cpa," +
				"chip_yield_gt_pa,chip_dollars_val_dpa,net_rev_dpa,harv_costs_cpa," + 
				"haul_costs_cpa) "; 


			this.m_strSQL += " SELECT  DISTINCT " + 
				"cycle1_best_rx_summary.biosum_cond_id," +
                "cycle1_best_rx_summary.owngrpcd," +
                "ROUND(cycle1_best_rx_summary.acres,10) AS acres," +
                "ROUND(cycle1_best_rx_summary.optimization_value,10) AS optimization_value," + 
				p + ".merch_haul_cost_psite," + 
				"ROUND(" + p + ".merch_haul_cpa_pt * cycle1_best_rx_summary.acres,10) AS merch_haul_cpa," +
                "ROUND(cycle1_effective_product_yields_by_rx_work_table3.merch_yield_cf * cycle1_best_rx_summary.acres,10) AS merch_vol_cf_pa," +
                "ROUND(cycle1_effective_product_yields_by_rx_work_table3.merch_val_dpa * cycle1_best_rx_summary.acres,10) AS merch_dollars_val_dpa," + 
				p + ".chip_haul_cost_psite," + 
				"ROUND(" + p + ".chip_haul_cpa_pt * cycle1_best_rx_summary.acres,10) AS chip_haul_cpa," +
                "ROUND(cycle1_effective_product_yields_by_rx_work_table3.chip_yield_gt * cycle1_best_rx_summary.acres,10) AS chip_yield_gt_pa," +
                "ROUND(cycle1_effective_product_yields_by_rx_work_table3.chip_val_dpa * cycle1_best_rx_summary.acres,10) AS chip_dollars_val_dpa," +
                "ROUND(cycle1_effective_product_yields_by_rx_work_table3.max_nr_dpa  * cycle1_best_rx_summary.acres,10) AS net_rev_dpa," +
                "ROUND(cycle1_effective_product_yields_by_rx_work_table3.harvest_onsite_cpa * cycle1_best_rx_summary.acres,10) AS harv_costs_cpa," +
                "ROUND((cycle1_effective_product_yields_by_rx_work_table3.haul_merch_cpa + cycle1_effective_product_yields_by_rx_work_table3.haul_chip_cpa) * cycle1_best_rx_summary.acres,10) AS haul_costs_cpa " + 
				"FROM ((cycle1_best_rx_summary  " +
                "INNER JOIN cycle1_effective_product_yields_by_rx_work_table3  " +
                "ON (cycle1_best_rx_summary.biosum_cond_id = cycle1_effective_product_yields_by_rx_work_table3.biosum_cond_id) AND " +
                "(cycle1_best_rx_summary." + strRxField + " = cycle1_effective_product_yields_by_rx_work_table3.rx)) " + 
				"INNER JOIN (" + p + 
				" INNER JOIN " + c + 
				" ON " + p + ".biosum_plot_id = " + c + ".biosum_plot_id) " + 
				"ON " + c + ".biosum_cond_id = cycle1_best_rx_summary.biosum_cond_id)" ;
				
			if (strWhereExpression.Trim().Length > 0)
			{
				this.m_strSQL += " WHERE " + strWhereExpression + ";";
			}
			else
			{
				this.m_strSQL += ";";
			}
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + this.m_strSQL + "\r\n\r\n");
			
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		private void BestRxAcreageExpansionTableInsertForAirCurtainDestruction(string strTable,string strTypeField, string strRxField, string strWhereExpression)
		{
			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();

			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(" + strTypeField + ", owngrpcd,acres,merch_haul_cost_psite," + 
				"merch_haul_cpa,merch_vol_cf_pa,merch_dollars_val_dpa," + 
				"chip_haul_cost_psite,chip_haul_cpa," +
				"chip_yield_gt_pa,chip_dollars_val_dpa,net_rev_dpa,harv_costs_cpa," + 
				"haul_costs_cpa) ";  
			this.m_strSQL += " SELECT  DISTINCT " + 
				"cycle1_best_rx_summary_air_dest.biosum_cond_id," + 
				"cycle1_best_rx_summary_air_dest.owngrpcd," + 
				"ROUND(cycle1_best_rx_summary_air_dest.acres,10) AS acres," +
				"null AS merch_haul_cost_psite," + 
				"null AS merch_haul_cpa," + 
				"ROUND(cycle1_effective_product_yields_net_rev_costs_summary_by_rx.merch_yield_cf * cycle1_best_rx_summary_air_dest.acres,10) AS merch_vol_cf_pa," + 
				"ROUND(cycle1_effective_product_yields_net_rev_costs_summary_by_rx.merch_val_dpa * cycle1_best_rx_summary_air_dest.acres,10) AS merch_dollars_val_dpa," + 
				"null AS chip_haul_cost_psite," + 
				"null AS chip_haul_cpa," + 
				"ROUND(cycle1_effective_product_yields_net_rev_costs_summary_by_rx.chip_yield_gt * cycle1_best_rx_summary_air_dest.acres,10) AS chip_yield_gt_pa," + 
				"ROUND(cycle1_effective_product_yields_net_rev_costs_summary_by_rx.chip_val_dpa * cycle1_best_rx_summary_air_dest.acres,10) AS chip_dollars_val_dpa," + 
				"null AS net_rev_dpa," + 
				"ROUND(cycle1_effective_product_yields_net_rev_costs_summary_by_rx.harvest_onsite_cpa * cycle1_best_rx_summary_air_dest.acres,10) AS harv_costs_cpa," + 
				"null AS haul_costs_cpa " + //," + 
				"FROM ((cycle1_best_rx_summary_air_dest  " + 
				"INNER JOIN cycle1_effective_product_yields_net_rev_costs_summary_by_rx  " + 
				"ON (cycle1_best_rx_summary_air_dest.biosum_cond_id = cycle1_effective_product_yields_net_rev_costs_summary_by_rx.biosum_cond_id) AND " + 
				"(cycle1_best_rx_summary_air_dest." + strRxField + " = cycle1_effective_product_yields_net_rev_costs_summary_by_rx.rx)) " + 
				"INNER JOIN (" + p + 
				" INNER JOIN " + c + 
				" ON " + p + ".biosum_plot_id = " + c + ".biosum_plot_id) " + 
				"ON " + c + ".biosum_cond_id = cycle1_best_rx_summary_air_dest.biosum_cond_id)" ;
				
			if (strWhereExpression.Trim().Length > 0)
			{
				this.m_strSQL += " WHERE " + strWhereExpression + ";";
			}
			else
			{
				this.m_strSQL += ";";
			}

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + this.m_strSQL + "\r\n\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		private void BestRxAcreageExpansionTableInsertForAirCurtainDestructionOld(string strTable,string strTypeField, string strRxField, string strWhereExpression)
		{
			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();

			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(" + strTypeField + ", owngrpcd,acres,merch_haul_cost_psite," + 
				"merch_haul_cents,merch_vol,merch_dollars_val," + 
				"chip_haul_cost_psite,chip_haul_cents," +
				"chip_yield,chip_dollars_val,net_rev,harv_costs," + 
				"haul_costs,ti_chg_acres,ci_chg_acres) " ;
			this.m_strSQL += " SELECT  " + 
				"best_rx_summary_air_dest.biosum_cond_id," + 
				"best_rx_summary_air_dest.owngrpcd," + 
				"best_rx_summary_air_dest.acres," +
				"null AS merch_haul_cost_psite," + 
				"null AS merch_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.merch_yield_cf * best_rx_summary_air_dest.acres AS merch_vol," + 
				"effective_product_yields_net_rev_costs_summary.merch_val_dpa * best_rx_summary_air_dest.acres AS merch_dollars_val," + 
				"null AS chip_haul_cost_psite," + 
				"null AS chip_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.chip_yield_gt * best_rx_summary_air_dest.acres AS chip_yield," + 
				"effective_product_yields_net_rev_costs_summary.chip_val_dpa * best_rx_summary_air_dest.acres AS chip_dollars_val," + 
				"null AS net_rev," + 
				"effective_product_yields_net_rev_costs_summary.harvest_onsite_cpa * best_rx_summary_air_dest.acres AS harv_costs," + 
				"null AS haul_costs," + 
				"effective.ti_change * best_rx_summary_air_dest.acres  AS ti_chg_acres," +
				"effective.ci_change * best_rx_summary_air_dest.acres  AS ci_chg_acres " + 
				"FROM (((best_rx_summary_air_dest  " + 
				"INNER JOIN effective_product_yields_net_rev_costs_summary  " + 
				"ON (best_rx_summary_air_dest.biosum_cond_id = effective_product_yields_net_rev_costs_summary.biosum_cond_id) AND " + 
				"(best_rx_summary_air_dest." + strRxField + " = effective_product_yields_net_rev_costs_summary.rx)) " + 
				"INNER JOIN effective " + 
				"ON (best_rx_summary_air_dest.biosum_cond_id = effective.biosum_cond_id) AND " +
				"(best_rx_summary_air_dest." + strRxField + "= effective.rx)) " + 
				"INNER JOIN (" + p + 
				" INNER JOIN " + c + 
				" ON " + p + ".biosum_plot_id = " + c + ".biosum_plot_id) " + 
				"ON " + c + ".biosum_cond_id = best_rx_summary_air_dest.biosum_cond_id)" ;
				
			if (strWhereExpression.Trim().Length > 0)
			{
				this.m_strSQL += " WHERE " + strWhereExpression + ";";
			}
			else
			{
				this.m_strSQL += ";";
			}

			//this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		/// <summary>
		/// sum wood volume, values, and costs by processing sites
		/// </summary>
		private void Cycle1SumPSite()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// Cycle1SumPSite\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n\r\n");
            }
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", this.m_strOptimizationTableName + " by Processing Site");
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n\r\nCycle1 Processing Site Summary\r\n-----------------------\r\n");


            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                   ReferenceUserControlScenarioRun.listViewEx1, "Cycle 1: Summarize The Best Effective Treatment Yields, Revenue, Costs, And Acres By Wood Processing Facility");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 6;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);


			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			
			this.m_strSQL = "delete from psite_sum_work_table";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + this.m_strSQL + "\r\n\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();


			            
			//MAX_NR_SUM_PSITE
			/*************************************************
			 **net revenue by processing site
			 *************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Net Revenue By Psite\r\n\r\n");
			/*************************************************
			 **process Chip only processing sites
			 *************************************************/
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Insert chip only psite sums\r\n");
			this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_stands","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Insert merch only psite sums\r\n");
			this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_stands","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Insert merch and chip psites sums from merch_haul_cost_psite column\r\n");
			this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_stands","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			string strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert merch and chip psites sums from chip_haul_cost_psite column\r\n");
			if (strPSiteList.Trim().Length > 0)
			{
				this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_stands","chip_haul_cost_psite",
					"WHERE a.psite_id=b.cheapest_psite AND " + 
					"a.psite_id=c.chip_haul_cost_psite AND " + 
					"a.biocd=3 AND " + 
					"b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
			}
			else
			{
				this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_stands","chip_haul_cost_psite",
					"WHERE a.psite_id=b.cheapest_psite AND " + 
					"a.psite_id=c.chip_haul_cost_psite AND " + 
					"a.biocd=3");

			}
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/*****************************************************************
			 **insert into max_nr_sum_psite
			 *****************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert from psite values from work table to max_nr_sum_psite\r\n");
			this.SumPSiteTableInsert(this.m_strOptimizationTableName + "_psites_sum");
			
            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
			/*************************************************
			 **Finished net revenue by processsing site
			 *************************************************/

			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
                return;
			}
		}

		private void SumPSiteWorkTableInsert(string strTable,string strPSiteField, string strWhereExpression)
		{
			
			this.m_strSQL = "INSERT INTO psite_sum_work_table " +
                "(psite_id,biocd,acres_sum,optimization_value_sum,chip_yield_gt_sum," + 
				"chip_haul_cost_sum, chip_dollars_val_sum,merch_haul_cost_sum," + 
				"merch_vol_cf_sum,net_rev_dollars_sum,merch_dollars_val_sum,harv_costs_sum," + 
				"haul_costs_sum) " + 
				"SELECT DISTINCT a.psite_id," + 
				"a.biocd," +  
				"c.acres_sum," + 
				"c.optimization_value_sum," + 
				"c.chip_yield_gt_sum," + 
				"c.chip_haul_cost_sum," +
				"c.chip_dollars_val_sum, " + 
				"c.merch_haul_cost_sum," + 
				"c.merch_vol_cf_sum," + 
				"c.net_rev_dollars_sum," + 
				"c.merch_dollars_val_sum," +
				"c.harv_costs_sum," + 
				"c.haul_costs_sum " + 
				"FROM " + this.m_strPSiteTable.Trim() + " AS a," + 
				"(SELECT biosum_cond_id," + 
				"MIN(" + strPSiteField.Trim() + ") as cheapest_psite " + 
				"FROM " + strTable + " GROUP BY biosum_cond_id)  b," + 
				"(SELECT " + strPSiteField.Trim() + "," + 
				"SUM(optimization_value) AS optimization_value_sum," + 
				"SUM(chip_yield_gt_pa)  AS chip_yield_gt_sum," + 
				"SUM(chip_haul_cpa) AS chip_haul_cost_sum," +
				"SUM(chip_dollars_val_dpa) AS chip_dollars_val_sum, " +
				"SUM(acres) AS acres_sum, " +
				"SUM(merch_haul_cpa) AS merch_haul_cost_sum," +
				"SUM(merch_vol_cf_pa) AS merch_vol_cf_sum," + 
				"SUM(net_rev_dpa) AS net_rev_dollars_sum," + 
				"SUM(merch_dollars_val_dpa) AS merch_dollars_val_sum," + 
				"SUM(harv_costs_cpa) AS harv_costs_sum," + 
				"SUM(haul_costs_cpa) AS haul_costs_sum " + 
				"FROM " + strTable + " GROUP BY " + strPSiteField.Trim() + ") c " +
				strWhereExpression.Trim() + ";";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		private void SumPSiteTableInsert(string strTable)
		{
			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(psite_id,acres_sum,optimization_value_sum,merch_haul_cost_sum,chip_haul_cost_sum," + 
				"merch_vol_cf_sum,chip_yield_gt_sum,net_rev_dollars_sum,merch_dollars_val_sum," + 
				"chip_dollars_val_sum,harv_costs_sum,haul_costs_sum) " + 
				"SELECT psite_id,acres_sum," + 
				"optimization_value_sum," + 
				"merch_haul_cost_sum," + 
				"chip_haul_cost_sum," + 
				"merch_vol_cf_sum," + 
				"chip_yield_gt_sum," + 
				"net_rev_dollars_sum," + 
				"merch_dollars_val_sum," + 
				"chip_dollars_val_sum," + 
				"harv_costs_sum," + 
				"haul_costs_sum " + 
				"FROM psite_sum_work_table;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		/// <summary>
		/// sum wood volumes, values, and costs by land ownership
		/// </summary>
		private void Cycle1SumOwnership()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// Cycle1SumOwnership\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n\r\n");
            }

            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", this.m_strOptimizationTableName + " by Ownership");
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nCycle1 Ownership Summary\r\n----------------------\r\n");


            intListViewIndex = FIA_Biosum_Manager.uc_core_scenario_run.GetListViewItemIndex(
                  ReferenceUserControlScenarioRun.listViewEx1, "Cycle 1: Summarize The Best Effective Treatment Yields, Revenue, Costs, And Acres By Land Ownership Groups");

            FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMaximumSteps = 9;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunCore.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunCore.g_intCurrentListViewItem, "focused", true);



			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			
            
			//MAX_NR_SUM_OWN
			/*************************************************
			 **net revenue by ownership
			 *************************************************/
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nPlots by Ownership\r\n\r\n");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Insert into work table\r\n");
			this.SumOwnershipWorkTableInsert(this.m_strOptimizationTableName + "_stands","own_sum_work_table");
            

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                 if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			/*****************************************************************
			 **insert into max_nr_sum_own
			 *****************************************************************/
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert ownership values from work table to " + this.m_strOptimizationTableName + "_own_sum\r\n");
			this.SumOwnershipTableInsert("own_sum_work_table",this.m_strOptimizationTableName + "_own_sum");
            

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}
			/*************************************************
			 **Finished net revenue by ownership
			 *************************************************/

            
            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();
			
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, this.m_strOptimizationTableName + " Air Curtain Destruction by Ownership\r\n");

                frmMain.g_oUtils.WriteText(m_strDebugFile,"Insert into work table\r\n");
            }
			this.SumOwnershipWorkTableInsert(this.m_strOptimizationTableName + "_stands_air_dest","own_sum_work_table");

            

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}
			/*****************************************************************
			 **insert into max_nr_sum_own
			 *****************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert ownership values from work table to " + this.m_strOptimizationTableName + "_own_air_dest\r\n");
			this.SumOwnershipTableInsert("own_sum_work_table",this.m_strOptimizationTableName + "_own_air_dest");

            

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

            if (this.UserCancel(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
			}

            

            FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermPercent();

			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_core_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunCore.g_oCurrentProgressBarBasic, "Done");
                return;
			}
			

		}

		


		private void SumOwnershipWorkTableInsert(string strTableSource,string strTableDestination)
		{
			this.m_strSQL = "INSERT INTO " + strTableDestination + " "  + 
				"(owngrpcd,acres_sum,optimization_value_sum,chip_yield_gt_sum," + 
				"chip_haul_cost_sum, chip_dollars_val_sum,merch_haul_cost_sum," + 
				"merch_vol_cf_sum,net_rev_dollars_sum,merch_dollars_val_sum,harv_costs_sum," + 
				"haul_costs_sum) " + 
				"SELECT owngrpcd," + 
				"SUM(acres) AS acres_sum, " +
				"SUM(optimization_value) AS optimization_value_sum," + 
				"SUM(chip_yield_gt_pa)  AS chip_yield_gt_sum," + 
				"SUM(IIF(chip_haul_cpa IS NOT NULL,chip_haul_cpa,0)) AS chip_haul_cost_sum," +
				"SUM(IIF(chip_dollars_val_dpa IS NOT NULL,chip_dollars_val_dpa,0)) AS chip_dollars_val_sum, " +
				"SUM(IIF(merch_haul_cpa IS NOT NULL,merch_haul_cpa,0)) AS merch_haul_cost_sum," +
				"SUM(merch_vol_cf_pa) AS merch_vol_cf_sum," + 
				"SUM(IIF(net_rev_dpa IS NOT NULL,net_rev_dpa,0)) AS net_rev_dollars_sum," + 
				"SUM(IIF(merch_dollars_val_dpa IS NOT NULL,merch_dollars_val_dpa,0)) AS merch_dollars_val_sum," + 
				"SUM(harv_costs_cpa) AS harv_costs_sum," + 
				"SUM(IIF(haul_costs_cpa IS NOT NULL,haul_costs_cpa,0)) AS haul_costs_sum " + 
				"FROM " + strTableSource + " GROUP BY owngrpcd ;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		
		private void SumOwnershipTableInsert(string strTableSource,string strTableDestination)
		{
			this.m_strSQL = "INSERT INTO " + strTableDestination + " " + 
				"(owngrpcd,acres_sum,optimization_value_sum,merch_haul_cost_sum,chip_haul_cost_sum," + 
				"merch_vol_cf_sum,chip_yield_gt_sum,net_rev_dollars_sum,merch_dollars_val_sum," + 
				"chip_dollars_val_sum,harv_costs_sum,haul_costs_sum) " + 
				"SELECT owngrpcd,acres_sum," + 
				"optimization_value_sum," + 
				"merch_haul_cost_sum," + 
				"chip_haul_cost_sum," + 
				"merch_vol_cf_sum," + 
				"chip_yield_gt_sum," + 
				"net_rev_dollars_sum," + 
				"merch_dollars_val_sum," + 
				"chip_dollars_val_sum," + 
				"harv_costs_sum," + 
				"haul_costs_sum " + 
				"FROM " + strTableSource;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		private void SumAirCurtainDestruction()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//SumAirCurtainDestruction\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n\r\n");
            }
			ReferenceUserControlScenarioRun.lblMsg.Text="Air Curtain Destruction";
			ReferenceUserControlScenarioRun.lblMsg.Visible=true;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n\r\nAir Curtain Destruction Summary\r\n-------------------------------\r\n");

			this.m_strSQL = "delete from air_curtain_destruction_stands";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();

			string strWhere = "WHERE " + p + ".merch_haul_cost_psite IS NULL OR " + 
				p + ".chip_haul_cost_psite IS NULL";


			/**************************************************************
			 **Finished air curtain destruction
			 ***************************************************************/


		}
		private void CreateHtml()
		{
			System.IO.FileStream oTxtFileStream;
			System.IO.StreamWriter oTxtStreamWriter;

			oTxtFileStream = new System.IO.FileStream(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runstats.htm", System.IO.FileMode.Create, 
				System.IO.FileAccess.Write);
			oTxtStreamWriter = new System.IO.StreamWriter(oTxtFileStream);
			oTxtStreamWriter.WriteLine("<html>\r\n");
			oTxtStreamWriter.WriteLine("<head>\r\n");
			oTxtStreamWriter.WriteLine("<title>\r\n");
			oTxtStreamWriter.WriteLine("FIA Biosum Core Analysis Scenario Run Summary Report\r\n");
			oTxtStreamWriter.WriteLine("</title>\r\n");
			oTxtStreamWriter.WriteLine("<body bgcolor='#ffffff' link='#33339a' vlink='#33339a' alink='#33339a'>\r\n");
			oTxtStreamWriter.WriteLine(System.DateTime.Now.ToString() + "\r\n");
			oTxtStreamWriter.WriteLine("<A NAME='GO TOP'></A>\r\n");
			oTxtStreamWriter.WriteLine("<BR> <BR>\r\n");
			oTxtStreamWriter.WriteLine("<b><CENTER><FONT SIZE='+2' >FIA Biosum Core Analysis Scenario Run Summary Report</FONT></b></center><br>\r\n");
			oTxtStreamWriter.WriteLine("<CENTER>\r\n");
			oTxtStreamWriter.WriteLine("<!--REPORT: CONDITION SAMPLE STATUS-->\r\n");
			oTxtStreamWriter.WriteLine("<TABLE COLSPAN='4' border='1' WIDTH='70%' HEIGHT='2%' cellpadding='0' cellspacing='0'\r\n>");
			//
			//condition table land class code counts
			//
			oTxtStreamWriter.WriteLine("<TR>\r\n");
			oTxtStreamWriter.WriteLine("   <td colspan='4'  bgcolor='lightgray' VAlign=middle height='2%' width='100%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("      <b>CONDITION SAMPLE STATUS</b>\r\n");
			oTxtStreamWriter.WriteLine("   </TD>\r\n");
			oTxtStreamWriter.WriteLine("</TR>\r\n");
			oTxtStreamWriter.WriteLine("<!--column headers-->\r\n");
			oTxtStreamWriter.WriteLine("<TR>\r\n");
			oTxtStreamWriter.WriteLine("	<td colspan='1'  bgcolor='lightyellow' VAlign=middle height='2%' width='25%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("		<b>Description</b>\r\n");
			oTxtStreamWriter.WriteLine("	</TD>\r\n");
			oTxtStreamWriter.WriteLine("	<td colspan='1'  bgcolor='lightyellow' VAlign=middle height='2%' width='25%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("		<b>Count</b>\r\n");
			oTxtStreamWriter.WriteLine("	</TD>\r\n");
			oTxtStreamWriter.WriteLine("	<td colspan='1'  bgcolor='lightyellow' VAlign=middle height='2%' width='25%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("		<b>With Trees</b>\r\n");
			oTxtStreamWriter.WriteLine("	</TD>\r\n");
			oTxtStreamWriter.WriteLine("	<td colspan='1'  bgcolor='lightyellow' VAlign=middle height='2%' width='25%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("		<b>Without Trees</b>\r\n");
			oTxtStreamWriter.WriteLine("	</TD>\r\n");
			oTxtStreamWriter.WriteLine("</TR>\r\n");
			oTxtStreamWriter.WriteLine("<!--count data-->\r\n");
			oTxtStreamWriter.WriteLine("<TR>\r\n");



			oTxtStreamWriter.WriteLine("</TABLE>\r\n");
			oTxtStreamWriter.WriteLine("</CENTER>\r\n");
			oTxtStreamWriter.WriteLine("</BODY>\r\n");
			oTxtStreamWriter.WriteLine("</HEAD>\r\n");
			oTxtStreamWriter.WriteLine("</HTML>\r\n");


			
			oTxtStreamWriter.Close();
			oTxtFileStream.Close();
			//oTxtFileStream=null;
			//oTxtStreamWriter=null;

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
        private void CompactMDB(string p_strMDBFileToCompact,System.Data.OleDb.OleDbConnection p_oConn)
        {
            string strConn = "";
            if (p_oConn != null)
            {
                strConn = p_oConn.ConnectionString;
                m_ado.CloseConnection(p_oConn);
            }
            m_oDao.CompactMDB(p_strMDBFileToCompact);
            System.Threading.Thread.Sleep(5000);
            if (strConn.Trim().Length > 0) m_ado.OpenConnection(strConn,ref p_oConn);
        }


		/// <summary>
		/// check and see if the user pressed the cancel button
		/// </summary>
		/// <param name="p_oLabel"></param>
		/// <returns></returns>
		private bool UserCancel(System.Windows.Forms.Label p_oLabel)
		{
			//System.Windows.Forms.Application.DoEvents();
			if (ReferenceUserControlScenarioRun.m_bUserCancel == true)
			{
				p_oLabel.ForeColor = System.Drawing.Color.Red;
				p_oLabel.Text = "Cancelled";
				return true;
			}
			return false;

		}
        private bool UserCancel(ProgressBarBasic.ProgressBarBasic p_oPb)
        {
            //System.Windows.Forms.Application.DoEvents();
            if (ReferenceUserControlScenarioRun.m_bUserCancel == true)
            {
                p_oPb.TextColor = Color.Red;
                frmMain.g_oDelegate.SetControlPropertyValue(p_oPb, "Text", "!!Cancelled!!");
                return true;
            }
            return false;

        }
		public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		
		public FIA_Biosum_Manager.uc_core_scenario_run ReferenceUserControlScenarioRun
		{
			get {return _uc_scenario_run;}
			set {_uc_scenario_run=value;}
		}
		
		
	}
}
