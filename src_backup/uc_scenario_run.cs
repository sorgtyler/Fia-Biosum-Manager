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
	public class uc_scenario_run : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmScenario _frmScenario=null;
		public System.Windows.Forms.Button btnViewLog;
		public System.Windows.Forms.Button btnAccess;
		public System.Windows.Forms.Button btnViewScenarioTables;
		public System.Windows.Forms.Button btnViewAuditTables;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnClear;
		private System.Windows.Forms.Button btnSelectAll;
		public System.Windows.Forms.Label lblProcSumTree;
		public System.Windows.Forms.Label lblProcTravelTimes;
		public System.Windows.Forms.Label lblProcAccessible;
		public System.Windows.Forms.CheckBox chkProcSumTree;
		public System.Windows.Forms.CheckBox chkProcTravelTimes;
		private System.Windows.Forms.GroupBox groupBox3;
		public System.Windows.Forms.CheckBox chkAuditTables;
		public System.Windows.Forms.Label lblProcBestRxOwner;
		public System.Windows.Forms.Label lblProcBestRxPSite;
		public System.Windows.Forms.Label lblProcBestRxPlot;
		public System.Windows.Forms.Label lblProcBestRx;
		public System.Windows.Forms.Label lblSumWoodProducts;
		public System.Windows.Forms.Label lblProcEffective;
		public System.Windows.Forms.Label lblProcValidCombos;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		public System.Windows.Forms.ProgressBar progressBar1;
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

		private FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective.Variables _oFVSPrePostVariables;
		private FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization.Variable_Collection _oFVSPrePostOptimization;
		private FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection _oFVSPrePostTieBreaker;
		private System.Windows.Forms.Panel panel1;
		

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_scenario_run()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
//			this.chkTreeSumTable();             //make sure table has records
//			this.chkPlotTableForTravelTimes();  //make sure table has travel times



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
			this.btnCancel = new System.Windows.Forms.Button();
			this.progressBar1 = new System.Windows.Forms.ProgressBar();
			this.lblMsg = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.btnClear = new System.Windows.Forms.Button();
			this.btnSelectAll = new System.Windows.Forms.Button();
			this.lblProcSumTree = new System.Windows.Forms.Label();
			this.lblProcTravelTimes = new System.Windows.Forms.Label();
			this.lblProcAccessible = new System.Windows.Forms.Label();
			this.chkProcSumTree = new System.Windows.Forms.CheckBox();
			this.chkProcTravelTimes = new System.Windows.Forms.CheckBox();
			this.btnViewLog = new System.Windows.Forms.Button();
			this.btnAccess = new System.Windows.Forms.Button();
			this.btnViewScenarioTables = new System.Windows.Forms.Button();
			this.btnViewAuditTables = new System.Windows.Forms.Button();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.chkAuditTables = new System.Windows.Forms.CheckBox();
			this.lblProcBestRxOwner = new System.Windows.Forms.Label();
			this.lblProcBestRxPSite = new System.Windows.Forms.Label();
			this.lblProcBestRxPlot = new System.Windows.Forms.Label();
			this.lblProcBestRx = new System.Windows.Forms.Label();
			this.lblSumWoodProducts = new System.Windows.Forms.Label();
			this.lblProcEffective = new System.Windows.Forms.Label();
			this.lblProcValidCombos = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.panel1 = new System.Windows.Forms.Panel();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.panel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.panel1);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(664, 504);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			// 
			// btnCancel
			// 
			this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancel.ForeColor = System.Drawing.Color.Black;
			this.btnCancel.Location = new System.Drawing.Point(288, 440);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(72, 24);
			this.btnCancel.TabIndex = 39;
			this.btnCancel.Text = "Start";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// progressBar1
			// 
			this.progressBar1.Enabled = false;
			this.progressBar1.Location = new System.Drawing.Point(16, 408);
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size(624, 24);
			this.progressBar1.TabIndex = 37;
			this.progressBar1.Visible = false;
			// 
			// lblMsg
			// 
			this.lblMsg.Enabled = false;
			this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMsg.ForeColor = System.Drawing.Color.Black;
			this.lblMsg.Location = new System.Drawing.Point(16, 376);
			this.lblMsg.Name = "lblMsg";
			this.lblMsg.Size = new System.Drawing.Size(624, 16);
			this.lblMsg.TabIndex = 38;
			this.lblMsg.Text = "lblMsg";
			this.lblMsg.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.btnClear);
			this.groupBox2.Controls.Add(this.btnSelectAll);
			this.groupBox2.Controls.Add(this.lblProcSumTree);
			this.groupBox2.Controls.Add(this.lblProcTravelTimes);
			this.groupBox2.Controls.Add(this.lblProcAccessible);
			this.groupBox2.Controls.Add(this.chkProcSumTree);
			this.groupBox2.Controls.Add(this.chkProcTravelTimes);
			this.groupBox2.ForeColor = System.Drawing.Color.Black;
			this.groupBox2.Location = new System.Drawing.Point(8, 64);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(640, 88);
			this.groupBox2.TabIndex = 35;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Optional:  Checked Boxes Will Execute";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(94, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(394, 16);
			this.label1.TabIndex = 8;
			this.label1.Text = "Determine If Plot And Conditions Are Accessible For Treatment And Harvest";
			// 
			// btnClear
			// 
			this.btnClear.Location = new System.Drawing.Point(560, 56);
			this.btnClear.Name = "btnClear";
			this.btnClear.Size = new System.Drawing.Size(64, 24);
			this.btnClear.TabIndex = 7;
			this.btnClear.Text = "Clear";
			this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
			// 
			// btnSelectAll
			// 
			this.btnSelectAll.Location = new System.Drawing.Point(560, 24);
			this.btnSelectAll.Name = "btnSelectAll";
			this.btnSelectAll.Size = new System.Drawing.Size(64, 24);
			this.btnSelectAll.TabIndex = 6;
			this.btnSelectAll.Text = "Select All";
			this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
			// 
			// lblProcSumTree
			// 
			this.lblProcSumTree.Location = new System.Drawing.Point(9, 63);
			this.lblProcSumTree.Name = "lblProcSumTree";
			this.lblProcSumTree.Size = new System.Drawing.Size(79, 16);
			this.lblProcSumTree.TabIndex = 5;
			this.lblProcSumTree.Text = "Not Done";
			// 
			// lblProcTravelTimes
			// 
			this.lblProcTravelTimes.Location = new System.Drawing.Point(8, 44);
			this.lblProcTravelTimes.Name = "lblProcTravelTimes";
			this.lblProcTravelTimes.Size = new System.Drawing.Size(80, 16);
			this.lblProcTravelTimes.TabIndex = 4;
			this.lblProcTravelTimes.Text = "Not Done";
			// 
			// lblProcAccessible
			// 
			this.lblProcAccessible.Location = new System.Drawing.Point(8, 24);
			this.lblProcAccessible.Name = "lblProcAccessible";
			this.lblProcAccessible.Size = new System.Drawing.Size(72, 16);
			this.lblProcAccessible.TabIndex = 3;
			this.lblProcAccessible.Text = "Not Done";
			// 
			// chkProcSumTree
			// 
			this.chkProcSumTree.Checked = true;
			this.chkProcSumTree.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkProcSumTree.Location = new System.Drawing.Point(95, 65);
			this.chkProcSumTree.Name = "chkProcSumTree";
			this.chkProcSumTree.Size = new System.Drawing.Size(217, 16);
			this.chkProcSumTree.TabIndex = 2;
			this.chkProcSumTree.Text = "Sum Tree Yields, Volume, And Value";
			// 
			// chkProcTravelTimes
			// 
			this.chkProcTravelTimes.Checked = true;
			this.chkProcTravelTimes.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkProcTravelTimes.Location = new System.Drawing.Point(95, 44);
			this.chkProcTravelTimes.Name = "chkProcTravelTimes";
			this.chkProcTravelTimes.Size = new System.Drawing.Size(377, 16);
			this.chkProcTravelTimes.TabIndex = 1;
			this.chkProcTravelTimes.Text = "Get Least Expensive Route From Plot To Wood Processing Facility";
			// 
			// btnViewLog
			// 
			this.btnViewLog.ForeColor = System.Drawing.Color.Black;
			this.btnViewLog.Location = new System.Drawing.Point(528, 32);
			this.btnViewLog.Name = "btnViewLog";
			this.btnViewLog.Size = new System.Drawing.Size(120, 20);
			this.btnViewLog.TabIndex = 34;
			this.btnViewLog.Text = "View Log File";
			this.btnViewLog.Click += new System.EventHandler(this.btnViewLog_Click);
			// 
			// btnAccess
			// 
			this.btnAccess.Enabled = false;
			this.btnAccess.ForeColor = System.Drawing.Color.Black;
			this.btnAccess.Location = new System.Drawing.Point(408, 8);
			this.btnAccess.Name = "btnAccess";
			this.btnAccess.Size = new System.Drawing.Size(120, 20);
			this.btnAccess.TabIndex = 33;
			this.btnAccess.Text = "Microsoft Access";
			this.btnAccess.Click += new System.EventHandler(this.btnAccess_Click);
			// 
			// btnViewScenarioTables
			// 
			this.btnViewScenarioTables.ForeColor = System.Drawing.Color.Black;
			this.btnViewScenarioTables.Location = new System.Drawing.Point(408, 32);
			this.btnViewScenarioTables.Name = "btnViewScenarioTables";
			this.btnViewScenarioTables.Size = new System.Drawing.Size(120, 20);
			this.btnViewScenarioTables.TabIndex = 32;
			this.btnViewScenarioTables.Text = "View Results Tables";
			this.btnViewScenarioTables.Click += new System.EventHandler(this.btnViewScenarioTables_Click);
			// 
			// btnViewAuditTables
			// 
			this.btnViewAuditTables.ForeColor = System.Drawing.Color.Black;
			this.btnViewAuditTables.Location = new System.Drawing.Point(528, 8);
			this.btnViewAuditTables.Name = "btnViewAuditTables";
			this.btnViewAuditTables.Size = new System.Drawing.Size(120, 20);
			this.btnViewAuditTables.TabIndex = 31;
			this.btnViewAuditTables.Text = "View Audit Data";
			this.btnViewAuditTables.Click += new System.EventHandler(this.btnViewAuditTables_Click);
			// 
			// lblTitle
			// 
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(8, 8);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(368, 32);
			this.lblTitle.TabIndex = 26;
			this.lblTitle.Text = "Run Scenario";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.chkAuditTables);
			this.groupBox3.Controls.Add(this.lblProcBestRxOwner);
			this.groupBox3.Controls.Add(this.lblProcBestRxPSite);
			this.groupBox3.Controls.Add(this.lblProcBestRxPlot);
			this.groupBox3.Controls.Add(this.lblProcBestRx);
			this.groupBox3.Controls.Add(this.lblSumWoodProducts);
			this.groupBox3.Controls.Add(this.lblProcEffective);
			this.groupBox3.Controls.Add(this.lblProcValidCombos);
			this.groupBox3.Controls.Add(this.label10);
			this.groupBox3.Controls.Add(this.label9);
			this.groupBox3.Controls.Add(this.label8);
			this.groupBox3.Controls.Add(this.label7);
			this.groupBox3.Controls.Add(this.label6);
			this.groupBox3.Controls.Add(this.label5);
			this.groupBox3.Controls.Add(this.label4);
			this.groupBox3.ForeColor = System.Drawing.Color.Black;
			this.groupBox3.Location = new System.Drawing.Point(8, 168);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(640, 193);
			this.groupBox3.TabIndex = 36;
			this.groupBox3.TabStop = false;
			// 
			// chkAuditTables
			// 
			this.chkAuditTables.Checked = true;
			this.chkAuditTables.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkAuditTables.Location = new System.Drawing.Point(400, 16);
			this.chkAuditTables.Name = "chkAuditTables";
			this.chkAuditTables.Size = new System.Drawing.Size(224, 16);
			this.chkAuditTables.TabIndex = 14;
			this.chkAuditTables.Text = "Populate Valid Combination Audit Data";
			// 
			// lblProcBestRxOwner
			// 
			this.lblProcBestRxOwner.Location = new System.Drawing.Point(9, 160);
			this.lblProcBestRxOwner.Name = "lblProcBestRxOwner";
			this.lblProcBestRxOwner.Size = new System.Drawing.Size(79, 16);
			this.lblProcBestRxOwner.TabIndex = 13;
			this.lblProcBestRxOwner.Text = "Not Done";
			// 
			// lblProcBestRxPSite
			// 
			this.lblProcBestRxPSite.Location = new System.Drawing.Point(9, 136);
			this.lblProcBestRxPSite.Name = "lblProcBestRxPSite";
			this.lblProcBestRxPSite.Size = new System.Drawing.Size(79, 16);
			this.lblProcBestRxPSite.TabIndex = 12;
			this.lblProcBestRxPSite.Text = "Not Done";
			// 
			// lblProcBestRxPlot
			// 
			this.lblProcBestRxPlot.Location = new System.Drawing.Point(9, 112);
			this.lblProcBestRxPlot.Name = "lblProcBestRxPlot";
			this.lblProcBestRxPlot.Size = new System.Drawing.Size(79, 16);
			this.lblProcBestRxPlot.TabIndex = 11;
			this.lblProcBestRxPlot.Text = "Not Done";
			// 
			// lblProcBestRx
			// 
			this.lblProcBestRx.Location = new System.Drawing.Point(9, 85);
			this.lblProcBestRx.Name = "lblProcBestRx";
			this.lblProcBestRx.Size = new System.Drawing.Size(79, 16);
			this.lblProcBestRx.TabIndex = 10;
			this.lblProcBestRx.Text = "Not Done";
			// 
			// lblSumWoodProducts
			// 
			this.lblSumWoodProducts.Location = new System.Drawing.Point(9, 59);
			this.lblSumWoodProducts.Name = "lblSumWoodProducts";
			this.lblSumWoodProducts.Size = new System.Drawing.Size(79, 16);
			this.lblSumWoodProducts.TabIndex = 9;
			this.lblSumWoodProducts.Text = "Not Done";
			// 
			// lblProcEffective
			// 
			this.lblProcEffective.Location = new System.Drawing.Point(9, 37);
			this.lblProcEffective.Name = "lblProcEffective";
			this.lblProcEffective.Size = new System.Drawing.Size(79, 16);
			this.lblProcEffective.TabIndex = 8;
			this.lblProcEffective.Text = "Not Done";
			// 
			// lblProcValidCombos
			// 
			this.lblProcValidCombos.Location = new System.Drawing.Point(9, 16);
			this.lblProcValidCombos.Name = "lblProcValidCombos";
			this.lblProcValidCombos.Size = new System.Drawing.Size(79, 16);
			this.lblProcValidCombos.TabIndex = 7;
			this.lblProcValidCombos.Text = "Not Done";
			// 
			// label10
			// 
			this.label10.Location = new System.Drawing.Point(95, 160);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(513, 24);
			this.label10.TabIndex = 6;
			this.label10.Text = "Summarize Most Effective Treatment Yields, Revenue, Costs, And Acres By Land Owne" +
				"rship Groups";
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(96, 136);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(520, 16);
			this.label9.TabIndex = 5;
			this.label9.Text = "Summarize Most Effective Treatment Yields, Revenue, Costs, And Acres By Wood Proc" +
				"essing Facility";
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(95, 112);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(497, 16);
			this.label8.TabIndex = 4;
			this.label8.Text = "Summarize Most Effective Treatment Yields, Revenue, Costs, And Acres By Stand";
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(96, 83);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(368, 18);
			this.label7.TabIndex = 3;
			this.label7.Text = "Identify The Best Effective Treatment For Each Stand";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(96, 56);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(488, 16);
			this.label6.TabIndex = 2;
			this.label6.Text = "Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Trea" +
				"tment";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(95, 36);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(336, 16);
			this.label5.TabIndex = 1;
			this.label5.Text = "Identify Effective Treatments For Each Stand";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(95, 15);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(305, 16);
			this.label4.TabIndex = 0;
			this.label4.Text = "Apply User Defined Filters And Get Valid Plot Combinations";
			// 
			// panel1
			// 
			this.panel1.AutoScroll = true;
			this.panel1.Controls.Add(this.groupBox2);
			this.panel1.Controls.Add(this.groupBox3);
			this.panel1.Controls.Add(this.lblMsg);
			this.panel1.Controls.Add(this.progressBar1);
			this.panel1.Controls.Add(this.btnCancel);
			this.panel1.Controls.Add(this.lblTitle);
			this.panel1.Controls.Add(this.btnAccess);
			this.panel1.Controls.Add(this.btnViewScenarioTables);
			this.panel1.Controls.Add(this.btnViewAuditTables);
			this.panel1.Controls.Add(this.btnViewLog);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(3, 16);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(658, 485);
			this.panel1.TabIndex = 40;
			// 
			// uc_scenario_run
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_run";
			this.Size = new System.Drawing.Size(664, 504);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.panel1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Visible=false;
			((frmScenario)this.ParentForm).Height = 0 ; //((frmScenario)this.ParentForm).grpboxMenu.Height * 2;
		}

		private void btnAccess_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo.UseShellExecute = true;
			try
			{
				proc.StartInfo.FileName = this.m_oRunCore.m_strTempMDBFile;
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
				
			string strMDBPathAndFile="";
			string strConn="";
			string strTable="";
			FIA_Biosum_Manager.Datasource p_datasource = new Datasource(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
				this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim());
			this.m_frmGridView = new frmGridView();
			this.m_frmGridView.Text = "Core Analysis: Audit";

			//plot and condition record audit
			strMDBPathAndFile = p_datasource.getFullPathAndFile("PLOT AND CONDITION RECORD AUDIT");
			strTable = p_datasource.getValidDataSourceTableName("PLOT AND CONDITION RECORD AUDIT");

			if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
				strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
					strMDBPathAndFile + ";User Id=admin;Password=;";
			else
				strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";

			this.m_frmGridView.LoadDataSet(strConn,"select * from " + strTable,strTable);

			//plot,condition, and rx record audit
			strMDBPathAndFile = p_datasource.getFullPathAndFile("PLOT, CONDITION AND TREATMENT RECORD AUDIT");
			strTable = p_datasource.getValidDataSourceTableName("PLOT, CONDITION AND TREATMENT RECORD AUDIT");

			if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
				strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
					strMDBPathAndFile + ";User Id=admin;Password=;";
			else
				strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";
			this.m_frmGridView.LoadDataSet(strConn,"select * from " + strTable,strTable);
			this.m_frmGridView.TileGridViews();
			this.m_frmGridView.Show();
			this.m_frmGridView.Focus();
			p_datasource = null;

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
					this.progressBar1.Minimum=0;
					this.progressBar1.Maximum=intCount;
					this.progressBar1.Value=0;
					this.progressBar1.Visible=true;
					this.lblMsg.Text = "";
					this.lblMsg.Visible=true;
					if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
						strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
							strMDBPathAndFile + ";User Id=admin;Password=;";
					else
						strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";
					this.m_frmGridView = new frmGridView();
					this.m_frmGridView.Text = "Core Analysis: Run Scenario Results ("  + this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + ")";
					for (int x=0; x <= intCount-1;x++)
					{
						this.lblMsg.Text = strTableNames[x];
						this.lblMsg.Refresh();
						strSQL = "select * from " + strTableNames[x].Trim();
						this.m_frmGridView.LoadDataSet(strConn,strSQL,strTableNames[x].Trim());
						this.progressBar1.Value = x + 1;

					}
					this.progressBar1.Visible=false;
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

		private void btnSelectAll_Click(object sender, System.EventArgs e)
		{
			this.chkProcTravelTimes.Checked=true;
			this.chkProcSumTree.Checked=true;
		}

		private void btnClear_Click(object sender, System.EventArgs e)
		{
			if (this.chkProcTravelTimes.Enabled) this.chkProcTravelTimes.Checked=false;
			if (this.chkProcSumTree.Enabled) this.chkProcSumTree.Checked=false;
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
							this.m_lblCurrentProcessStatus.ForeColor = Color.Red;
							this.m_lblCurrentProcessStatus.Text = "!!Cancelled!!";
							this.btnCancel.Text = "Start";
							this.m_bUserCancel=true;

						}
					}

				}



				
				
			}
			else
			{
				RunScenario();
			}
		}
		private void RunScenario()
		{
			this.progressBar1.Visible=false;
			this.lblMsg.Text = "";
			this.lblProcAccessible.ForeColor = System.Drawing.Color.Black;
			this.lblProcAccessible.Text = "Not Done";
			this.lblProcBestRx.ForeColor = System.Drawing.Color.Black;
			this.lblProcBestRx.Text = "Not Done";
			this.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Black;
			this.lblProcBestRxOwner.Text = "Not Done";
			this.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Black;
			this.lblProcBestRxPlot.Text = "Not Done";
			this.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Black;
			this.lblProcBestRxPSite.Text = "Not Done";
			this.lblProcEffective.ForeColor = System.Drawing.Color.Black;
			this.lblProcEffective.Text = "Not Done";
			this.lblProcSumTree.ForeColor = System.Drawing.Color.Black;
			this.lblProcSumTree.Text = "Not Done";
			this.lblProcTravelTimes.ForeColor = System.Drawing.Color.Black;
			this.lblProcTravelTimes.Text = "Not Done";
			this.lblProcValidCombos.ForeColor = System.Drawing.Color.Black;
			this.lblProcValidCombos.Text = "Not Done";
			this.lblSumWoodProducts.ForeColor = System.Drawing.Color.Black;
			this.lblSumWoodProducts.Text = "Not Done";
			this.Refresh();
			this.val_CoreRunData();
			if (this.m_intError==0)
			{
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
				frmMain.g_oDelegate.SetControlPropertyValue(progressBar1,"Value",frmMain.g_oDelegate.GetControlPropertyValue(progressBar1,"Maximum",false));
				frmMain.g_oDelegate.ExecuteControlMethod(progressBar1,"Refresh");
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

			
			
			strMDBPathAndFile = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
			strConn=p_ado.getMDBConnString(strMDBPathAndFile,"admin","");
			if (p_ado.TableExist(strConn,"tree_vol_val_sum"))
			{
				//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";
				if (p_ado.getRecordCount(strConn,"select COUNT(*) from tree_vol_val_sum","tree_vol_val_sum") == 0)
				{
					this.chkProcSumTree.Checked=true;
					this.chkProcSumTree.Enabled=false;
				}
				else
				{
					this.chkProcSumTree.Enabled=true;
				}
			}
			else
			{
				this.chkProcSumTree.Checked=true;
				this.chkProcSumTree.Enabled=false;
			}

		}
		

		/// <summary>
		/// Check to see if  fastest travel times 
		/// from plot to wood processing site exist in the plot table
		/// </summary>
		public void chkPlotTableForTravelTimes()
		{
			string strConn="";
			ado_data_access p_ado = new ado_data_access();

			string strTable = "";
			string strMDBFile = "";
			this.ReferenceCoreScenarioForm.uc_datasource1.getScenarioDataSourceMDBAndTableName("PLOT",ref strMDBFile,ref strTable);
			strConn=p_ado.getMDBConnString(strMDBFile,"admin","");
			//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBFile + ";User Id=admin;Password=;";
			if (p_ado.getRecordCount(strConn,"select COUNT(*) from " + strTable + " p WHERE p.merch_haul_cost_id IS NOT NULL AND p.chip_haul_cost_id IS NOT NULL",strTable) == 0)

			{
				this.chkProcTravelTimes.Checked=true;
				this.chkProcTravelTimes.Enabled=false;
			}
			else
			{
				this.chkProcTravelTimes.Enabled=true;
			}
			p_ado = null;

		}

		/// <summary>
		/// validate each component required for running core analysis
		/// </summary>
		private void val_CoreRunData()
		{
			this.m_intError=0;

			if (this.m_intError==0) this.m_intError = ReferenceCoreScenarioForm.uc_scenario_owner_groups1.ValInput();
			if (this.m_intError==0)  this.m_intError = ReferenceCoreScenarioForm.uc_scenario_costs1.val_costs();
			if (this.m_intError==0) this.m_intError = ReferenceCoreScenarioForm.uc_scenario_psite1.val_psites();
			if (this.m_intError==0)  
			{   
				this.m_intError= ReferenceCoreScenarioForm.uc_scenario_filter1.Val_PlotFilter(ReferenceCoreScenarioForm.uc_scenario_filter1.txtCurrentSQL.Text.Trim());
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_filter1.m_strError,"FIA Biosum");

			}
			if (this.m_intError==0)  
			{
				this.m_intError = ReferenceCoreScenarioForm.uc_scenario_cond_filter1.Val_CondFilter(ReferenceCoreScenarioForm.uc_scenario_cond_filter1.txtCurrentSQL.Text.Trim());
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_cond_filter1.m_strError,"FIA Biosum");

			}
			if (this.m_intError==0) 
			{
				this.m_intError = ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.Audit(false);
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_strError,"FIA Biosum");
			}
			if (this.m_intError==0) 
			{
				ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.DisplayAuditMessage=false;
				ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.Audit();
				this.m_intError = ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.m_intError;
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.m_strError,"FIA Biosum");
			}

			if (this.m_intError==0) 
			{
				this.m_intError = ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.Audit(false);
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_strError,"FIA Biosum");
			}
		
            
          
			if (this.m_intError==0)
			
			{
				
				/***************************************************************************
					 **make sure all the scenario datasource tables and files are available
					 **and ready for use
					 ***************************************************************************/
					
				if (ReferenceCoreScenarioForm.m_ldatasourcefirsttime==true)
				{
					ReferenceCoreScenarioForm.uc_datasource1.populate_listview_grid();
					ReferenceCoreScenarioForm.m_ldatasourcefirsttime=false;
				}
				this.m_intError = ReferenceCoreScenarioForm.uc_datasource1.val_datasources();
				if (this.m_intError ==0)
				{
					ReferenceCoreScenarioForm.SaveRuleDefinitions();
				}
			}
				

		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			this.panel1.Width = this.groupBox1.Width - (int)(this.panel1.Left * 2);
			this.groupBox2.Width = this.panel1.Width - (int)(this.groupBox2.Left * 2) - 20;
			this.groupBox3.Width = this.groupBox2.Width;
			this.lblMsg.Width = this.groupBox1.Width - (int)(this.lblMsg.Left * 2);
			this.progressBar1.Width = this.groupBox1.Width - (int)(this.progressBar1.Left * 2);
			this.btnViewAuditTables.Left = this.groupBox2.Width - this.btnViewAuditTables.Width + this.groupBox2.Left ;
			this.btnViewLog.Left = this.btnViewAuditTables.Left;
			this.btnAccess.Left = this.btnViewLog.Left - this.btnAccess.Width;
			this.btnViewScenarioTables.Left = this.btnViewLog.Left - this.btnViewScenarioTables.Width;
			this.btnSelectAll.Left = this.groupBox2.Width - this.btnSelectAll.Width - 5;
			this.btnClear.Left = this.btnSelectAll.Left;
			this.btnCancel.Left = (int)(this.progressBar1.Width * .50) - (int)(this.btnCancel.Width * .50);
		}

		

			
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective.Variables ReferenceFVSPrePostVariables
		{
			get {return _oFVSPrePostVariables;}
			set {_oFVSPrePostVariables=value;}
		}
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization.Variable_Collection ReferenceFVSPrePostOptimization
		{
			get {return _oFVSPrePostOptimization;}
			set {_oFVSPrePostOptimization=value;}
		}
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection ReferenceFVSPrePostTieBreaker
		{
			get {return this._oFVSPrePostTieBreaker;}
			set {this._oFVSPrePostTieBreaker=value;}
		}
		//public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker ReferenceFVSTieBreakers
		//{
				
		//}


		public FIA_Biosum_Manager.frmScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
	}
	/// <summary>
	/// main class used for running the core analysis scenario
	/// </summary>
	public class RunCore
	{
		FIA_Biosum_Manager.uc_scenario_run _uc_scenario_run;
		FIA_Biosum_Manager.frmScenario _frmScenario;
		// FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective.Variables _fvs_prepost_variables;
		//	FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization.Variable_Collection _fvs_prepost_optimization_collection;
		private int m_intError;
		private string m_strSQL;
		private string m_strConn;
		public string m_strTempMDBFile;
		public ado_data_access m_ado;
		public System.Data.OleDb.OleDbConnection m_TempMDBFileConn;
		public string m_strSystemResultsDbPathAndFile="";
		private env m_oEnv;
		private utils m_oUtils;
		public string m_strPlotTable;
		public string m_strRxTable;
		public string m_strTravelTimeTable;
		public string m_strCondTable;
		public string m_strFFETable;
		public string m_strHvstCostsTable;
		public string m_strPSiteTable;
		public string m_strTreeVolValBySpcDiamGroupsTable;
		public string m_strTreeVolValSumTable = "tree_vol_val_sum";
		public string m_strUserDefinedPlotSQL;
	    public string m_strUserDefinedCondSQL;
		private System.IO.FileStream m_txtFileStream;
		private System.IO.StreamWriter m_txtStreamWriter;
		private string m_strLine;
		private FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization.VariableItem m_oOptimizationVariable = new FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization.VariableItem();
		private string m_strOptimizationTableName="";
		private string m_strOptimizationSourceTableName="";
		private string m_strOptimizationTableNameSql="";
		private string m_strOptimizationColumnNameSql="";
		private string m_strOptimizationSourceColumnName="";
		private string m_strOptimizationAggregateSql="";
		private string m_strOptimizationAggregateColumnName="";
		private FIA_Biosum_Manager.macrosubst m_oVarSub = new macrosubst();

		
		

		public RunCore(FIA_Biosum_Manager.uc_scenario_run p_form)
		{
			
			this.m_intError=0;
			_uc_scenario_run = p_form;
			try
			{
				this.m_txtFileStream = new System.IO.FileStream(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runlog.txt", System.IO.FileMode.Create, 
					System.IO.FileAccess.Write);
				this.m_txtStreamWriter = new System.IO.StreamWriter(this.m_txtFileStream);
				this.m_txtStreamWriter.WriteLine(" ");
				this.m_txtStreamWriter.WriteLine(" ");
				this.m_strLine = "Core Analysis Run Log " + System.DateTime.Now.ToString();
				this.m_txtStreamWriter.WriteLine("{0}{1}","        ", this.m_strLine);
				this.m_txtStreamWriter.WriteLine(" ");
				this.m_txtStreamWriter.WriteLine("Project: {0}", frmMain.g_oFrmMain.frmProject.uc_project1.txtProjectId.Text);
				this.m_txtStreamWriter.WriteLine("Project Directory: {0}", frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text);
				this.m_txtStreamWriter.WriteLine("Scenario Directory: {0}", ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim());
				this.m_strLine = "---------------------------------------------------------------";
				this.m_txtStreamWriter.WriteLine(this.m_strLine);
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.ToString());
				this.m_intError=-1;
				return;
			}

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Create A Temporary MDB File With Links To All The Core Tables And Scenario Result Tables");
			this.m_txtStreamWriter.WriteLine("------------------------------------------------------------------------------------------------");
			try
			{

				/**************************************************************************
				 **first lets create a temp mdb with links to all the scenario core 
				 **and result tables
				 **************************************************************************/
				this.m_oUtils=new utils();
				this.m_oEnv = new env();

				this.m_oVarSub.ReferenceSQLMacroSubstitutionVariableCollection = 
					frmMain.g_oSQLMacroSubstitutionVariable_Collection;

				this.m_strSystemResultsDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"mdb");
				this.CopyScenarioResultsTable(this.m_strSystemResultsDbPathAndFile,ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb");


				this.m_strTempMDBFile = 
					ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.CreateMDBAndScenarioTableDataSourceLinks(this.m_oEnv.strTempDir);  
			
				this.m_strUserDefinedPlotSQL= 
					this.m_oVarSub.SQLTranslateVariableSubstitution(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_filter1.txtCurrentSQL.Text.Trim());

				this.m_strUserDefinedCondSQL= 
					this.m_oVarSub.SQLTranslateVariableSubstitution(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_cond_filter1.txtCurrentSQL.Text.Trim());

            
				if (this.m_strTempMDBFile.Trim().Length == 0)
				{
					this.m_txtStreamWriter.WriteLine("!!RunCore: Error Creating MDB File Containing Links To All The Tables!!");
					MessageBox.Show("RunCore: Error Creating MDB File Containing Links To All The Core Tables");
				}
				else
				{
					this.m_txtStreamWriter.WriteLine("Delete Scenario Result Table Records");
					this.DeleteScenarioResultRecords();

					this.m_txtStreamWriter.WriteLine("Initialize Air Destruction Tables");
					InitializeAirDestructionTables();


					this.m_txtStreamWriter.WriteLine("Links MDB File: " + this.m_strTempMDBFile);
					ReferenceUserControlScenarioRun.btnAccess.Enabled=false;

					getTableNames();	
	

					//effective variable
					FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective.Variables oFvsVar =
						ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar;

					/********************************************************************
					 **get optimization variable
					 ********************************************************************/
					FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization.Variable_Collection oOptimizationVariableCollection = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_optimization1.m_oSavVariableCollection;
						
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
						this.m_strOptimizationSourceTableName="product_yields_net_rev_costs_summary";
						this.m_strOptimizationTableNameSql = "effective_product_yields_net_rev_costs_summary";
						this.m_strOptimizationSourceColumnName="max_nr_dpa";
						this.m_strOptimizationColumnNameSql="post_variable_value";
							
					}
					else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
					{
						this.m_strOptimizationSourceTableName="product_yields_net_rev_costs_summary";
						this.m_strOptimizationTableNameSql = "effective_product_yields_net_rev_costs_summary";
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
						this.m_strOptimizationTableNameSql="effective_optimization_treatments";
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
						this.m_strOptimizationTableName = this.m_strOptimizationTableName + Convert.ToString(this.m_oOptimizationVariable.dblFilterValue).Trim();

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
							this.m_strOptimizationTableName = this.m_strOptimizationTableName + oFvsVar.m_strOverallEffectiveNetRevValue.Trim();
						}
					}
		
					CreateScenarioResultTables();
					if (this.m_intError!=0) return;

				
					//CREATE TABLE LINKS
					CreateScenarioResultTableLinks();
					if (this.m_intError != 0) return;
				

					this.CreateScenarioTableLinks();
					if (this.m_intError !=0) return;

					this.CreateFVSPrePostTableLinks();
					if (this.m_intError !=0) return;

					//CREATE WORK TABLES
					CreateTableStructureOfHarvestCosts();
					if (this.m_intError != 0) return;

					CreateTableStructureForIntensity();
					if (this.m_intError !=0) return;

					this.CreateTableStructureForScenarioProcessingSites();
					if (this.m_intError !=0) return;

					this.CreateTableStructureForPSiteSum();
					if (this.m_intError !=0) return;

					this.CreateTableStructureForOwnershipSum();
					if (this.m_intError != 0) return;



					this.CreateTableStructureForHaulCosts();
					if (this.m_intError != 0) return;

					this.CreateTableStructureForPlotCondAccessiblity();
					if (this.m_intError != 0) return;

				
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

						/********************************************************************
						 **create table structure for condition table filters
						 ********************************************************************/
						this.CreateTableStructureForUserDefinedConditionTable();

						/********************************************************************
						 **filter scenario selected processing sites
						 ********************************************************************/
						FilterPSites();

						


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


						/***********************************************************
						 **identify the plots that are accessible
						 ***********************************************************/
						if (this.m_intError==0) // && ReferenceUserControlScenarioRun.chkProcAccessible.Checked==true)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcAccessible,"ForeColor",System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcAccessible,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcAccessible,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcAccessible;
							this.PlotAccessible();
				
						}
						else
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcAccessible,"Text","NA");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcAccessible,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcAccessible;
						}

					

						/**************************************************************
						 **get the fastest travel time from plot to processing site
						 **************************************************************/
						if (this.m_intError == 0 && (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.chkProcTravelTimes,"Checked",false)==true && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcTravelTimes,"ForeColor",System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcTravelTimes,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcTravelTimes,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcTravelTimes;
							this.getHaulCosts();

						}
						else
						{
							if (ReferenceUserControlScenarioRun.m_bUserCancel==false)
							{
								frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcTravelTimes,"Text","NA");;
								frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcTravelTimes,"Refresh");
								ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcTravelTimes;
							}
						}
                    
						/***************************************************************************
						 **sum up tree volumes and values by plot+condition, treatment and species
						 ***************************************************************************/
						if (this.m_intError == 0 && (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.chkProcSumTree,"Checked",false)==true && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcSumTree,"ForeColor",System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcSumTree,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcSumTree,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcSumTree;
							this.sumTreeVolVal();

						}
						else
						{
							if (ReferenceUserControlScenarioRun.m_bUserCancel == false)
							{
								frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcSumTree,"Text","NA");
								frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcSumTree,"Refresh");
								ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcSumTree;
							}
						}
						/***************************************************************************
						 **valid combos
						 ***************************************************************************/
						if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcValidCombos,"ForeColor",System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcValidCombos,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcValidCombos,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcValidCombos;
							this.validcombos();

						}
						/***************************************************************
						 **wood product yields net revenue and costs summary table
						 ***************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblSumWoodProducts,"ForeColor", System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblSumWoodProducts,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblSumWoodProducts,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblSumWoodProducts;
							this.product_yields_net_rev_costs_summary();

						}
						/***************************************************************************
						 **effective treatments
						 ***************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"ForeColor", System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcEffective;
							this.effective();

						}
						/**************************************************************************
						 **optimization
						 **************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.optimization();
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
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcBestRx;
							this.best_rx_summary();
						}
						/*******************************************************************************
						 **expand acreage for best treatments by plot 
						 *******************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxPlot,"ForeColor", System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxPlot,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxPlot,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcBestRxPlot;
							this.BestTreatmentsByPlot();
						}
						/**********************************************************************
						 **sum up the values by processing site
						 **********************************************************************/
						if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxPSite,"ForeColor", System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxPSite,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxPSite,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcBestRxPSite;
							this.SumPSite();

						}

						/**********************************************************************
						 **sum up the values by ownership
						 **********************************************************************/
						if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxOwner,"ForeColor", System.Drawing.Color.Green);
							frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxOwner,"Text","Processing");
							frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRxOwner,"Refresh");
							ReferenceUserControlScenarioRun.m_lblCurrentProcessStatus=ReferenceUserControlScenarioRun.lblProcBestRxOwner;
							this.SumOwnership();

						}
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							System.DateTime oDate = System.DateTime.Now;
							string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
							string strFileDate = oDate.ToString(strDateFormat);
							strFileDate = strFileDate.Replace("/","_"); strFileDate=strFileDate.Replace(":","_");
							this.CreateHtml();
							this.CopyScenarioResultsTable(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb",this.m_strSystemResultsDbPathAndFile);
							this.m_strSystemResultsDbPathAndFile=ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
							this.CopyScenarioResultsTable(
								ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results_" + this.m_strOptimizationTableName + "_" + strFileDate.Trim() + ".mdb",
								ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb");
						}

						/**********************************************************************
						 **sum up air destruction plots
						 **********************************************************************/
						//if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						//{
						//	ReferenceUserControlScenarioRun.lblProcAirCurtainDestructionPlots.ForeColor = System.Drawing.Color.Green;
						//	ReferenceUserControlScenarioRun.lblProcAirCurtainDestructionPlots.Text="Processing";
						//	ReferenceUserControlScenarioRun.lblProcAirCurtainDestructionPlots.Refresh();
						//	this.SumAirCurtainDestruction();

						//}
						this.m_strLine = "***End Of Core Analysis Run: " + System.DateTime.Now.ToString() + " ***"; 
						this.m_txtStreamWriter.WriteLine(this.m_strLine);

						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Visible",false);
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",false);
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnCancel,"Text","Start");
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewScenarioTables,"Enabled",true);
						if (this.m_intError==0) ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.UpdateChipsMarketPerGreenTonValueFromProcessorToScenario();
						if (this.m_intError == 0) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnAccess,"Enabled",true);
					
					

						if (this.m_intError != 0)
						{
							MessageBox.Show("Completed With Errors");
						}
						else
						{
							MessageBox.Show("Successfully Completed");
						}
					}
				}
				this.m_txtStreamWriter.Close();
				this.m_txtFileStream.Close();
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewLog,"Enabled",true);
				if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.chkProcSumTree,"Enabled",false)==false)
					ReferenceUserControlScenarioRun.chkTreeSumTable();
				if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.chkProcTravelTimes,"Enabled",false)==false)
					ReferenceUserControlScenarioRun.chkPlotTableForTravelTimes();
				if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.chkAuditTables,"Enabled",false)==true)
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewAuditTables,"Enabled",true);
			}
			catch (Exception err)
			{
				if (err.Message.Trim().ToUpper() != "THREAD WAS BEING ABORTED.")
					MessageBox.Show("Run Scenario " + err.Message,"FIA Biosum");
				this.m_txtStreamWriter.Close();
				this.m_txtFileStream.Close();
			}
			frmMain.g_oDelegate.m_oEventThreadStopped.Set();
			this.ReferenceUserControlScenarioRun.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
		}
			
		public RunCore()
		{

		}
		~RunCore()
		{
			
			this.m_txtFileStream.Close();
			this.m_txtStreamWriter.Close();
			this.m_txtFileStream=null;
			this.m_txtStreamWriter=null;
		}
		

		/// <summary>
		/// create a table structure that will hold
		/// the plot data that results when running the user 
		/// defined sql
		/// </summary>
		/// <param name="strUserDefinedSQL"></param>
		private void CreateTableStructureOfUserDefinedPlotSQL()
		{
			
			
			/*******************************************************************
			 ** get scenario_results.mdb path
			 *******************************************************************/
			//string strMDBPathAndFile = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
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
			this.m_txtStreamWriter.WriteLine("Create userdefinedplotfilter Table Schema From User Defined Plot Filter SQL");
			dao_data_access p_dao = new dao_data_access();
			//LPOTTS p_dao.CreateMDBTableFromDataSetTable(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","userdefinedplotfilter",p_dt,true);
			p_dao.CreateMDBTableFromDataSetTable(this.m_strSystemResultsDbPathAndFile,"userdefinedplotfilter",p_dt,true);
			p_dt.Dispose();
			this.m_ado.m_OleDbDataReader.Close();
			if (p_dao.m_intError!=0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
				this.m_intError=p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the user defined plot filter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			this.m_txtStreamWriter.WriteLine("Create link in " + this.m_strTempMDBFile);
			p_dao.CreateTableLink(this.m_strTempMDBFile,"userdefinedplotfilter",m_strSystemResultsDbPathAndFile,"userdefinedplotfilter",true);
			this.m_txtStreamWriter.WriteLine("{0}\t{1}",m_strSystemResultsDbPathAndFile,"userdefinedplotfilter");

			/***********************************************************************
			 **make a copy of the userdefinedplot filter table and give it the
			 **name ruledefinitionsplotfilter. This will apply the owngrpcd
			 **filters and any other future filters to the userdefinedplotfilter table.
			 ***********************************************************************/
			this.m_txtStreamWriter.WriteLine("Delete table ruledefinitionsplotfilter");
			//first delete the table if it exists already
			//LPOTTS p_dao.DeleteTableFromMDB(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","ruledefinitionsplotfilter");
			p_dao.DeleteTableFromMDB(this.m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter");
			if (p_dao.m_intError !=0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Deleting alluserdefinedplotfilter Table!!");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			this.m_txtStreamWriter.WriteLine("Copy table structure userdefinedplotfilter to ruledefinitionsplotfilter");
			//LPOTTSp_dao.MoveTableToMDB(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","ruledefinitionsplotfilter",ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","userdefinedplotfilter",false);
			p_dao.MoveTableToMDB(this.m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter",this.m_strSystemResultsDbPathAndFile,"userdefinedplotfilter",false);
			if (p_dao.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Creating ruledefinitionsplotfilter Table!!");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the ruledefinitionsplotfilter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionsplotfilter",m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter",true);
			this.m_txtStreamWriter.WriteLine("Create link in " + this.m_strTempMDBFile);
			this.m_txtStreamWriter.WriteLine("{0}\t{1}",m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter");


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
			this.m_txtStreamWriter.WriteLine("Create userdefinedplotfilter Table Schema From User Defined Plot Filter SQL");
			dao_data_access p_dao = new dao_data_access();
			p_dao.CreateMDBTableFromDataSetTable(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","userdefinedplotfilter",p_dt,true);
			p_dt.Dispose();
			this.m_ado.m_OleDbDataReader.Close();
			if (p_dao.m_intError!=0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
				this.m_intError=p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the user defined plot filter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			this.m_txtStreamWriter.WriteLine("Create link in " + this.m_strTempMDBFile);
			p_dao.CreateTableLink(this.m_strTempMDBFile,"userdefinedplotfilter",strMDBPathAndFile,"userdefinedplotfilter",true);
			this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,"userdefinedplotfilter");

			/***********************************************************************
			 **make a copy of the userdefinedplot filter table and give it the
			 **name ruledefinitionsplotfilter. This will apply the owngrpcd
			 **filters and any other future filters to the userdefinedplotfilter table.
			 ***********************************************************************/
			this.m_txtStreamWriter.WriteLine("Delete table ruledefinitionsplotfilter");
			//first delete the table if it exists already
			p_dao.DeleteTableFromMDB(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","ruledefinitionsplotfilter");
			if (p_dao.m_intError !=0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Deleting alluserdefinedplotfilter Table!!");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			this.m_txtStreamWriter.WriteLine("Copy table structure userdefinedplotfilter to ruledefinitionsplotfilter");
			p_dao.MoveTableToMDB(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","ruledefinitionsplotfilter",ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","userdefinedplotfilter",false);
			if (p_dao.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Creating ruledefinitionsplotfilter Table!!");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the ruledefinitionsplotfilter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionsplotfilter",strMDBPathAndFile,"ruledefinitionsplotfilter",true);
			this.m_txtStreamWriter.WriteLine("Create link in " + this.m_strTempMDBFile);
			this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,"ruledefinitionsplotfilter");


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
			

			string[] strTableNames;
			strTableNames = new string[1];
			ado_data_access oAdo = new ado_data_access();
			

		
			
			oAdo.OpenConnection(oAdo.getMDBConnString(m_strSystemResultsDbPathAndFile,"",""));
			
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"validcombos"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE validcombos");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"validcombos_fvspost"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE validcombos_fvspost");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"validcombos_fvspre"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE validcombos_fvspre");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"validcombos_fvsprepost"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE validcombos_fvsprepost");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"effective"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE effective");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"optimization"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE optimization");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"tiebreaker"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE tiebreaker");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"validcombos"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE validcombos");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"best_rx_summary_all"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE best_rx_summary_all");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"best_rx_summary"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE best_rx_summary");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"best_rx_summary_before_tiebreaks"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE best_rx_summary_before_tiebreaks");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"best_rx_summary_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE best_rx_summary_air_dest");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"best_rx_summary_air_dest_before_tiebreaks"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE best_rx_summary_air_dest_before_tiebreaks");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"tree_vol_val_sum"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE tree_vol_val_sum");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"product_yields_net_rev_costs_summary"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE product_yields_net_rev_costs_summary");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryWithIntensityTableName))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryWithIntensityTableName);

			if (oAdo.TableExist(oAdo.m_OleDbConnection,this.m_strOptimizationTableName + "_plots"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + m_strOptimizationTableName + "_plots");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,this.m_strOptimizationTableName + "_psites"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + m_strOptimizationTableName + "_psites");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,this.m_strOptimizationTableName + "_own"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + m_strOptimizationTableName + "_own");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,this.m_strOptimizationTableName + "_plots_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + m_strOptimizationTableName + "_plots_air_dest");

			if (oAdo.TableExist(oAdo.m_OleDbConnection,this.m_strOptimizationTableName + "_own_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + m_strOptimizationTableName + "_own_air_dest");

			

			frmMain.g_oTables.m_oCoreScenarioResults.CreateValidComboTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsValidCombosTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateValidComboFVSPostTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsValidCombosFVSPostTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateValidComboFVSPreTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsValidCombosFVSPreTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateValidComboFVSPrePostTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsValidCombosFVSPrePostTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateTieBreakerTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsTieBreakerTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsOptimizationTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateEffectiveTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsEffectiveTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryWithIntensityTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryWithIntensityTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryTableName + "_air_dest");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryTableName + "_before_tiebreaks");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateTreeVolValSumTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsTreeVolValSumTableName);
			frmMain.g_oTables.m_oCoreScenarioResults.CreateProductYieldsTable(oAdo,oAdo.m_OleDbConnection,"product_yields_net_rev_costs_summary");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationPlotTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_plots");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationPSiteTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_psites");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationOwnershipTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_own");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationPlotTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_plots_air_dest");
			frmMain.g_oTables.m_oCoreScenarioResults.CreateOptimizationOwnershipTable(oAdo,oAdo.m_OleDbConnection,m_strOptimizationTableName + "_own_air_dest");

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
					if (intCount > 0)
					{
						for (int x=0; x <= intCount-1;x++)
						{
							this.m_txtStreamWriter.WriteLine("{0}\t{1}","scenario_results",strTableNames[x]);

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
				this.m_txtStreamWriter.WriteLine(p_dao.m_strError);
				this.m_intError = p_dao.m_intError;
			}

		
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
						this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,strTableNames[x]);

						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
							this.m_txtStreamWriter.WriteLine("!!Error Creating Table Link!!!");
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
			

			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			string strConn="";
			
			string strMDBPathAndFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\audit.mdb";

			if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
				strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";
			else
				strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";

			
			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					for (int x=0; x <= intCount-1;x++)
					{
						this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,strTableNames[x]);
						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
							this.m_txtStreamWriter.WriteLine("!!Error Creating Table Link!!!");
							this.m_intError = p_dao.m_intError;
							break;
						}
					}

				}
			}
			else
			{
				this.m_intError = p_dao.m_intError;
			}
			p_dao = null;
			
		}
			
		/// <summary>
		/// create links to the scenario tables that contain the 
		/// rule definitions defined by the user
		/// </summary>
		private void CreateScenarioTableLinks()
		{
			

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
						this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,strTableNames[x]);
						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
							this.m_txtStreamWriter.WriteLine("!!Error Creating Table Link!!!");
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
		/// create links to the fvs out pre and post variable tables
		/// </summary>
		private void CreateFVSPrePostTableLinks()
		{
			

			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			string strConn="";
			
			string strMDBPathAndFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\biosum_fvsout_prepost_rx.mdb";

			strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
				strMDBPathAndFile + ";User Id=admin;Password=;";

			
			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames,false);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					for (int x=0; x <= intCount-1;x++)
					{
						this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,strTableNames[x]);
						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
							this.m_txtStreamWriter.WriteLine("!!Error Creating Table Link!!!");
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
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Get Core Table Names");
			this.m_txtStreamWriter.WriteLine("---------------------");
			/**************************************************************
			 **get the plot table name
			 **************************************************************/
			this.m_strPlotTable=ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("PLOT");
			this.m_txtStreamWriter.WriteLine("Plot:{0}",this.m_strPlotTable);


			/**************************************************************
			 **get the treatment prescriptions table
			 **************************************************************/
			this.m_strRxTable = 
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("TREATMENT PRESCRIPTIONS");
			this.m_txtStreamWriter.WriteLine("Treatment:{0}",this.m_strRxTable);

			/**************************************************************
			 **get the travel time table name
			 **************************************************************/
			this.m_strTravelTimeTable = 
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("TRAVEL TIMES");
			this.m_txtStreamWriter.WriteLine("Travel Time:{0}",this.m_strTravelTimeTable);
			/**************************************************************
			 **get the cond table name
			 **************************************************************/
			this.m_strCondTable=ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("CONDITION");
			this.m_txtStreamWriter.WriteLine("Condition:{0}",this.m_strCondTable);


			/*************************************************************
			 **get the harvest costs table
			 *************************************************************/
			this.m_strHvstCostsTable=ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("HARVEST COSTS");
			this.m_txtStreamWriter.WriteLine("Harvest Costs:{0}",this.m_strHvstCostsTable);

			/**************************************************************
			 **get the processing site table name
			 **************************************************************/

			this.m_strPSiteTable = "scenario_psites_work_table";
			this.m_txtStreamWriter.WriteLine("Processsing Sites:{0}",this.m_strPSiteTable);

			this.m_strTreeVolValBySpcDiamGroupsTable = 
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableName("TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS");
			this.m_txtStreamWriter.WriteLine("Tree Volumes And Values By Species And Diameter groups:{0}",this.m_strTreeVolValBySpcDiamGroupsTable);


			this.m_strTreeVolValSumTable = "tree_vol_val_sum";
			this.m_txtStreamWriter.WriteLine("Tree Sum Volume And Value:{0}",this.m_strTreeVolValSumTable);

		}
		
		private void FilterPSites()
		{
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Processing Site Filter");
			this.m_txtStreamWriter.WriteLine("-------------------");

			this.m_strSQL="DELETE FROM scenario_psites_work_table";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_strSQL = "INSERT INTO scenario_psites_work_table (psite_id,trancd,biocd) " + 
				"SELECT psite_id,trancd,biocd " + 
				"FROM scenario_psites " + 
				"WHERE TRIM(scenario_id)='" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + "' AND " + 
				"selected_yn='Y';";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
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
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Plot and Condition Accessibility");
			this.m_txtStreamWriter.WriteLine("-------------------");


			ReferenceUserControlScenarioRun.lblProcAccessible.ForeColor = System.Drawing.Color.Green;
			ReferenceUserControlScenarioRun.lblProcAccessible.Text = "Processing";

			

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
				"c.slope <= 40 AND " + 
				"p.gis_yard_dist >= " + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strNonSteepYardingDistance.Trim() + ";" ;

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

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
				"c.slope > 40 AND " + 
				"p.gis_yard_dist >= " + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_cond_filter1.strSteepYardingDistance.Trim() + ";" ;

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);



			/*************************************************************
			 **set the remainder of the cond_too_far_steep_yn fields to N
			 *************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
				"SET c.cond_too_far_steep_yn = 'N' " + 
				"WHERE c.cond_too_far_steep_yn IS NULL OR (c.cond_too_far_steep_yn <> 'Y' AND " + 
				"c.cond_too_far_steep_yn <> 'N');";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/*************************************************************
			 **update the condition accessible flag
			 *************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable +  " SET cond_accessible_yn = " + 
				"IIF(cond_too_far_steep_yn='Y','N','Y');";



			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

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
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_strSQL = "INSERT INTO plot_cond_accessible_work_table2 (biosum_plot_id, num_cond, num_cond_not_accessible) " + 
				"SELECT a.biosum_plot_id, a.num_cond,b.cond_not_accessible_count as num_cond_not_accessible FROM plot_cond_accessible_work_table a," + 
				"(SELECT biosum_plot_id, Count(*) AS cond_not_accessible_count FROM " + this.m_strCondTable + " " + 
				"WHERE cond_accessible_yn='N' GROUP BY biosum_plot_id) b " + 
				"WHERE a.biosum_plot_id = b.biosum_plot_id;";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/**********************************************************************
			 **initialize the all_cond_not_accessible_yn flag to N
			 **********************************************************************/
			this.m_strSQL = this.m_oVarSub.SQLTranslateVariableSubstitution("UPDATE @@PlotTable@@ SET all_cond_not_accessible_yn = 'N'");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/**********************************************************************
			 **update the all_cond_not_accessible_yn flag if all conditions 
			 **are not accessible
			 **********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
				"INNER JOIN plot_cond_accessible_work_table2 c " + 
				"ON p.biosum_plot_id=c.biosum_plot_id " + 
				"SET p.all_cond_not_accessible_yn=" + 
				"IIF(c.num_cond=c.num_cond_not_accessible,'Y','N');";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/*********************************************************************
			 **update the plot accessible flag
			 *********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable +  " SET plot_accessible_yn = " + 
				"IIF(gis_protected_area_yn='Y' " + 
				" OR gis_roadless_yn='Y' " + 
				" OR all_cond_not_accessible_yn='Y','N','Y');";

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


            


			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcAccessible) == true) return;
			if (this.m_ado.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcAccessible.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcAccessible.Text = "Completed";

			}
			else
			{
				this.m_txtStreamWriter.WriteLine("!!Error Executing SQL!!");
				ReferenceUserControlScenarioRun.lblProcAccessible.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcAccessible.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;

			}
			ReferenceUserControlScenarioRun.lblProcAccessible.Refresh();

		}
		/// <summary>
		/// populate the haul_costs table and plot table with 
		/// the cheapest route for hauling merch and chip
		/// </summary>
		private void getHaulCosts()
		{
			
			//int x;
			string strTruckHaulCost;
			string strRailHaulCost;
			string strTransferMerchCost;
			string strTransferBioCost;


			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Processing Haul Costs");
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",true);
			frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Refresh");
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Minimum",0);
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Maximum",27);
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",0);
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Visible",true);

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update Plot And Haul Cost Tables With Merch And Chip Haul Costs");
			this.m_txtStreamWriter.WriteLine("-------------------------------------------------------------");

			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
			strTruckHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtHaulCost_subclass.Text.Replace("$","").ToString();
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			
			strRailHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailHaulCost_subclass.Text.Replace("$","").ToString();
			strRailHaulCost = strRailHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
			strTransferMerchCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailMerchTransfer_subclass.Text.Replace("$","").ToString();
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
			strTransferBioCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailChipTransfer_subclass.Text.Replace("$","").ToString();
			strTransferBioCost = strTransferBioCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferBioCost = "0.00";
			 



			/*******************************************************************************
			 **zap the haul_costs table
			 *******************************************************************************/
			this.m_strSQL = "DELETE FROM haul_costs;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("delete records in haul_costs table");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",1);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcTravelTimes,"ForeColor",System.Drawing.Color.Red);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcTravelTimes,"Text","!!Error!!");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcTravelTimes,"Refresh");
				return;
			}     
			
			
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Null The Plot Table's Haul Cost Fields");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

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
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("null the plot table's haul cost fields");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",3);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     
			
			

			/*****************************************************************
			 **delete any records that may exist in the work tables
			 *****************************************************************/

			
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Delete Records In Work Tables");
			
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

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
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",4);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			//MERCH AND CHIP ROAD PROCESSING SITE HAUL COSTS
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Road Haul Costs For Merchantable Wood Processing Sites");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Merchantable wood processing site haul costs");
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
				"(s.biocd=3 OR s.biocd=1);";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table all travel time records where psite has road access and processes merch");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",5);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
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


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table. Find the cheapest plot to merch processing site road route.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",5);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Road Haul Costs For Chip Wood Processing Sites");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

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
				"FROM " + this.m_strTravelTimeTable + " t ," + 
				this.m_strPSiteTable + " s " + 
				"WHERE t.psite_id=s.psite_id AND " + 
				"(s.trancd=1 OR s.trancd=3) AND " + 
				"(s.biocd=3 OR s.biocd=2);";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table all travel time records where psite has road access and processes chips.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",7);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
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



			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table. Find the cheapest plot to chip processing site road route.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",8);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     


			//MERCH AND CHIP RAIL PROCESSING SITE HAUL COSTS
			/*********************************************************
			 **Append to a table all travel time collector_id (psite)
			 **records where the psite has rail access
			 *********************************************************/
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Rail Haul Costs For Merchantable Wood Processing Sites");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

			this.m_strSQL = "INSERT INTO merch_rh_to_collector_haul_costs_work_table " + 
				"SELECT t.psite_id AS railhead_id," + 
				"t.collector_id AS psite_id," + 
				"(" + strTransferMerchCost.Trim() + " * t.travel_time)  AS transfer_cost," + 
				"0 AS road_cost," + 
				"((t.travel_time * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost," + 
				"0 AS total_haul_cost,  'M' AS materialcd " + 
				"FROM " + this.m_strTravelTimeTable + " t " + 
				"INNER JOIN  " + this.m_strPSiteTable + " s " + 
				"ON t.collector_id = s.psite_id " + 
				"WHERE  s.trancd=3 And (s.biocd=3 Or s.biocd=1) AND " + 
				"EXISTS (SELECT ss.psite_id " + 
				"FROM " + this.m_strPSiteTable + " ss " + 
				"WHERE t.psite_id=ss.psite_id AND ss.trancd=2 AND (ss.biocd=3 Or ss.biocd=1));";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table all travel time collector_id (psite) records where the psite has rail access and processes merch.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",9);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
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
				"WHERE r.railhead_id = t.psite_id;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table travel time plot records and previous work rail/merch table results");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",10);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     


			this.m_strSQL = "UPDATE merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SET total_haul_cost = transfer_cost + road_cost + rail_cost;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update merch by road and rail total haul cost ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",11);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
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


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Find the cheapest plot to merch processing site rail routes");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",12);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Rail Haul Costs For Chip Wood Processing Sites");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

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
				"FROM " + this.m_strTravelTimeTable + " t " + 
				"INNER JOIN  " + this.m_strPSiteTable + " s " + 
				"ON t.collector_id = s.psite_id " + 
				"WHERE s.trancd=3 AND  " + 
				"(s.biocd=3 OR s.biocd=2) AND " + 
				"EXISTS (SELECT ss.psite_id " + 
				"FROM " + this.m_strPSiteTable + " ss " + 
				"WHERE t.psite_id=ss.psite_id AND ss.trancd=2 AND (ss.biocd=3 Or ss.biocd=2));";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table all travel time collector_id (psite) records where the psite has rail access and processes chips.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",13);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
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
				"WHERE r.railhead_id = t.psite_id;";




			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table travel time plot records and previous rail/chips work table results");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",14);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}    
 

			
			
			
			
			
			
			this.m_strSQL = "UPDATE chip_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SET total_haul_cost = transfer_cost + road_cost + rail_cost;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update chips by road and rail total haul cost ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",15);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
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

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Find the cheapest plot to chip processing site rail routes");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",16);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

		    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Combine Road And Rail Haul Costs For Merchantable Wood Processing Sites");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();


			/**************************************************************
			 **combine the cheapest road and rail total cost for each plot
			 **to a merch psite
			 **After the insert there should be two records for each
			 **plot - one with cheapest haul cost by road and another
			 **with cheapest haul cost by rail
			 ***************************************************************/
			this.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_road_merch_haul_costs_work_table;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Cheapest road route to merch psite ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",18);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			this.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_rail_merch_haul_costs_work_table;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Cheapest rail route to merch psite ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",18);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}   
  

			/***************************************************
			 **Get the overall cheapest merch route
			 ***************************************************/
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Get Overall Least Expensive Merch Route");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
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


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Get the overall cheapest merch route");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",19);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}   

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Combine Road And Rail Haul Costs For Chip Wood Processing Sites");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

			/**************************************************************
			 **combine the cheapest road and rail total cost for each plot
			 **to a chips psite
			 **After the insert there should be two records for each
			 **plot - one with cheapest haul cost by road and another
			 **with cheapest haul cost by rail
			 ***************************************************************/
			this.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_road_chip_haul_costs_work_table;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Cheapest road route to chip psite ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",20);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			this.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_rail_chip_haul_costs_work_table;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Cheapest rail route to chip psite ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",21);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}   
  
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Get Overall Least Expensive Chip Route");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();


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



			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Get the overall cheapest chip route");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",22);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}   




			//INSERT INTO HAUL_COSTS TABLE
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Inserting Results Into Haul Costs Table");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_strSQL = "INSERT INTO haul_costs " + 
				"SELECT * FROM cheapest_merch_haul_costs_work_table;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into haul_costs table cheapest merch route for each plot");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",23);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}   


			this.m_strSQL = "INSERT INTO haul_costs " + 
				"SELECT * FROM cheapest_chip_haul_costs_work_table;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into haul_costs table cheapest chip route for each plot");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",24);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}   

			//UPDATE PLOT TABLE

			/**************************************************
			 **Update cheapest merch route fields
			 **************************************************/
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Updating Plot Table");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
				"INNER JOIN haul_costs h " + 
				"ON p.biosum_plot_id = h.biosum_plot_id " + 
				"SET p.merch_haul_cost_id = h.haul_cost_id," + 
				"p.merch_haul_cost_psite = h.psite_id," + 
				"p.merch_haul_cpa_pt=h.total_haul_cost " + 
				"WHERE h.materialcd='M';";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update plot merch haul cost fields");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",25);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
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


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update plot chip haul cost fields");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",26);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			} 
  

			/******************************************
			 **clean up work tables
			 ******************************************/

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Cleaning Up Haul Cost Work Tables...Stand By");
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Cleaning up haul cost work tables");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",27);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			} 
			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
			}
			ReferenceUserControlScenarioRun.lblMsg.Visible=false;
			ReferenceUserControlScenarioRun.progressBar1.Visible=false;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Refresh();



		}


		/// <summary>
		/// populate the plot table's merch_tvltm and chip_tvltm fields with
		/// the fastest time from plot to wood processing site
		/// </summary>
		private void getTravelTimes()
		{
			
			//int x;
			string strTruckHaulCost;
			string strRailHaulCost;
			string strTransferMerchCost;
			string strTransferBioCost;

			
			ReferenceUserControlScenarioRun.lblMsg.Text="Process Travel Times Unique Id Fields";
			ReferenceUserControlScenarioRun.lblMsg.Visible=true;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Minimum=0;
			ReferenceUserControlScenarioRun.progressBar1.Maximum = 10;
			ReferenceUserControlScenarioRun.progressBar1.Value=0;
			ReferenceUserControlScenarioRun.progressBar1.Visible=true;

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update Plot Table With Merch And Chip Haul Costs");
			this.m_txtStreamWriter.WriteLine("-------------------------------------------------------------");

			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
			strTruckHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtHaulCost_subclass.Text.Replace("$","").ToString();
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			strTruckHaulCost="";
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			
			strRailHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailHaulCost_subclass.Text.Replace("$","").ToString();
			strRailHaulCost = strTruckHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
			strTransferMerchCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailMerchTransfer_subclass.Text.Replace("$","").ToString();
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
			strTransferBioCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailMerchTransfer_subclass.Text.Replace("$","").ToString();
			strTransferBioCost = strTransferMerchCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			 


			/*************************************************************
			 **first make sure each travel times biosum_plot_id field
			 **has a value
			 *************************************************************/					   
			this.m_strSQL = "UPDATE " + this.m_strTravelTimeTable  + " t " + 
				"INNER JOIN " + this.m_strPlotTable + " p ON t.plot_id = p.idb_plot_id " + 
				"SET t.biosum_plot_id = p.biosum_plot_id " + 
				"WHERE t.plot_id IS NOT NULL AND t.biosum_plot_id IS NULL" ;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update " + this.m_strTravelTimeTable.Trim() + ".biosum_cond_id field NULL values");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 1;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     
			ReferenceUserControlScenarioRun.lblMsg.Text="Null The Plot Table's Travel Time Fields";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

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
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("null the plot table's haul cost fields");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 2;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     
			
			

			/*****************************************************************
				 **delete any records that may exist in the work tables
				 *****************************************************************/
			this.m_strSQL = "delete from haul_costs;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from haul_costs_work_table2;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
		   		   
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 3;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     
			//MERCH PROCESSING SITE HAUL COSTS
			ReferenceUserControlScenarioRun.lblMsg.Text="Haul Costs For Merchantable Wood Processing Sites";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Merchantable wood processing site haul costs");
			/***************************************************************
			 **process the merch travel times first
			   ***************************************************************/
			  
			  
			  
			/********************************************************************
			 **insert all travel times for merch processing sites
			 ********************************************************************/
			this.m_strSQL = "INSERT INTO plot_haul_cost_work_table " + 
				"(biosum_plot_id, merch_haul_cost_psite,merch_tvltm) " + 
				"SELECT t.biosum_plot_id, " + 
				"s.psite_id, " + 
				"t.travel_time " + 
				"FROM travel_time t," + 
				"processing_site s " + 
				"WHERE t.psite_id = s.psite_id AND " + 
				"(s.biocd = 1 OR s.biocd=3);";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table merch processing site travel times");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 3;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			/************************************************************************
			 **insert into work table 2 the fastest merch travel time for each plot 
			 ************************************************************************/
			this.m_strSQL = "INSERT INTO plot_fastest_tvltm_work_table2 " + 
				"(biosum_plot_id, merch_fastest_tvltm_psite, merch_tvltm) " + 
				"SELECT b.biosum_plot_id, " +
				"c.merch_fastest_tvltm_psite, " + 
				"b.min_travel AS merch_tvltm " + 
				"FROM plot_fastest_tvltm_work_table a, " + 
				"(SELECT biosum_plot_id, " + 
				"MIN(merch_tvltm) as min_travel " +  
				"FROM plot_fastest_tvltm_work_table " + 
				"GROUP BY biosum_plot_id)  b," + 
				"(SELECT biosum_plot_id, " + 
				"merch_fastest_tvltm_psite," + 
				"MIN(merch_tvltm) as min_travel2 " + 
				"FROM plot_fastest_tvltm_work_table " + 
				"GROUP BY biosum_plot_id, merch_fastest_tvltm_psite) c " + 
				"WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.merch_fastest_tvltm_psite = c.merch_fastest_tvltm_psite AND  " + 
				"b.min_travel = c.min_travel2;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table the fastest merch processing site travel times");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 5;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			/***********************************************************
			  **okay, update the plot table with the fastest merch times 
			  ***********************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " a " +  
				"INNER JOIN plot_fastest_tvltm_work_table2  b " + 
				"ON a.biosum_plot_id = b.biosum_plot_id " + 
				"SET a.merch_fastest_tvltm_psite = b.merch_fastest_tvltm_psite," + 
				"a.merch_tvltm = b.merch_tvltm," + 
				"a.merch_haul_cpa_pt = b.merch_tvltm * " + strTruckHaulCost.Trim() + ";";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update plot table from the work table");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL); 
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 6;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     


			/*****************************************************************
			**delete any records that may exist in the work tables
			*****************************************************************/
			this.m_strSQL = "delete from plot_fastest_tvltm_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from plot_fastest_tvltm_work_table2";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//Chip PROCESSING SITE TRAVEL TIMES
			ReferenceUserControlScenarioRun.lblMsg.Text="Travel Times For Chip Wood Processing Sites";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Chip wood processing site travel times");
			/***************************************************************
			 **process the Chip travel times
			 ***************************************************************/
			/********************************************************************
			 **insert all travel times for Chips processing sites
			 ********************************************************************/
			this.m_strSQL = "INSERT INTO plot_fastest_tvltm_work_table " + 
				"(biosum_plot_id, chip_fastest_tvltm_psite,chip_tvltm) " + 
				"SELECT t.biosum_plot_id, " + 
				"s.psite_id, " + 
				"t.travel_time " + 
				"FROM travel_time t," + 
				"processing_site s " + 
				"WHERE t.psite_id = s.psite_id AND " + 
				"(s.biocd = 2 OR s.biocd=3);";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table Chip processing site travel times");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 7;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			/************************************************************************
			 **insert into work table 2 the fastest merch travel time for each plot 
			 ************************************************************************/
			this.m_strSQL = "INSERT INTO plot_fastest_tvltm_work_table2 " + 
				"(biosum_plot_id, chip_fastest_tvltm_psite, chip_tvltm) " + 
				"SELECT b.biosum_plot_id, " +
				"c.chip_fastest_tvltm_psite, " + 
				"b.min_travel AS chip_tvltm " + 
				"FROM plot_fastest_tvltm_work_table a, " + 
				"(SELECT biosum_plot_id, " + 
				"MIN(chip_tvltm) as min_travel " +  
				"FROM plot_fastest_tvltm_work_table " + 
				"GROUP BY biosum_plot_id)  b," + 
				"(SELECT biosum_plot_id, " + 
				"chip_fastest_tvltm_psite," + 
				"MIN(chip_tvltm) as min_travel2 " + 
				"FROM plot_fastest_tvltm_work_table " + 
				"GROUP BY biosum_plot_id, chip_fastest_tvltm_psite) c " + 
				"WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.chip_fastest_tvltm_psite = c.chip_fastest_tvltm_psite AND  " + 
				"b.min_travel = c.min_travel2;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table the fastest Chip processing site travel times");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 9;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			/***********************************************************
			  **okay, update the plot table with the fastest Chips times 
			  ***********************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " a " +  
				"INNER JOIN plot_fastest_tvltm_work_table2  b " + 
				"ON a.biosum_plot_id = b.biosum_plot_id " + 
				"SET a.chip_fastest_tvltm_psite = b.chip_fastest_tvltm_psite," + 
				"a.chip_tvltm = b.chip_tvltm," + 
				"a.chip_haul_cpa_pt = b.chip_tvltm * " + strTruckHaulCost.Trim() + ";";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update plot table from the work table");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL); 
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcTravelTimes)) return;
			ReferenceUserControlScenarioRun.progressBar1.Value = 10;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
				return;
			}     

			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcTravelTimes.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcTravelTimes.Refresh();
			}
			ReferenceUserControlScenarioRun.lblMsg.Visible=false;
			ReferenceUserControlScenarioRun.progressBar1.Visible=false;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Refresh();



		}
		/// <summary>
		/// sum the tree_vol_val_by_species_diam_groups table values to tree_vol_val_sum
		/// </summary>
		private void sumTreeVolVal()
		{
			
			
			/**************************************************************
			 **sum the tree_vol_val_by_species_diam_groups table to
			 **        tree_vol_val_sum
			 **************************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Sum Tree Volumes and Values");
			this.m_txtStreamWriter.WriteLine("---------------------------");

			

			this.m_strSQL = "delete from tree_vol_val_sum";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				ReferenceUserControlScenarioRun.lblProcSumTree.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcSumTree.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcSumTree.Refresh();
				return;
			}
			this.m_strSQL = "INSERT INTO tree_vol_val_sum " + 
				"(biosum_cond_id,rx,chip_vol_cf," + 
				"chip_wt_gt,chip_val_dpa,merch_vol_cf," + 
				"merch_wt_gt,merch_val_dpa) ";

			this.m_strSQL += "SELECT biosum_cond_id, " + 
				"rx," +
				"Sum(chip_vol_cf) AS chip_vol_cf," + 
				"Sum(chip_wt_gt) AS chip_wt_gt," + 
				"Sum(chip_val_dpa) AS chip_val_dpa," + 
				"Sum(merch_vol_cf) AS merch_vol_cf," +
				"Sum(merch_wt_gt) AS merch_wt_gt," + 
				"Sum(merch_val_dpa) AS merch_val_dpa ";

			this.m_strSQL += " FROM " + this.m_strTreeVolValBySpcDiamGroupsTable.Trim();
			this.m_strSQL += " GROUP BY biosum_cond_id,rx";
			this.m_strSQL += " ORDER BY biosum_cond_id,rx;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into tree_vol_val_sum table tree volume and value sums");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				ReferenceUserControlScenarioRun.lblProcSumTree.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcSumTree.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcSumTree.Refresh();
				return;
			}
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcSumTree)) return;


			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcSumTree.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcSumTree.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcSumTree.Refresh();
			}
		}
		private void getHaulCost(string p_strBiosumPlotId, int p_intPSiteId)
		{
			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
			string strTruckHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtHaulCost_subclass.Text.Replace("$","").ToString();
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			
			string strRailHaulCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailHaulCost_subclass.Text.Replace("$","").ToString();
			strRailHaulCost = strRailHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
			string strTransferMerchCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailMerchTransfer_subclass.Text.Replace("$","").ToString();
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
			string strTransferBioCost = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_costs1.txtRailChipTransfer_subclass.Text.Replace("$","").ToString();
			strTransferBioCost = strTransferBioCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferBioCost = "0.00";
			



		}

		/// <summary>
		/// load the validcombos table with biosum_cond_id and rx values
		/// that exist in the user defined plot filters, condition, ffe, travel times, and
		/// harvest cost, and tree volume/value tables
		/// </summary>
		private void validcombos()
		{
			string strRxList="";
			string strGrpCd="";
			int x=0;
			string strTable="";
			//int y=0;
			//int z=0;
			

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Valid Combinations");
			this.m_txtStreamWriter.WriteLine("------------------");

			/*****************************************************************
			 **delete audit tables
			 *****************************************************************/
			//for (x=0;x<=50000;x++)
			this.m_strSQL = "delete from plot_cond_audit";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				ReferenceUserControlScenarioRun.lblProcValidCombos.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcValidCombos.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcValidCombos.Refresh();
				return;
			}
			this.m_strSQL = "delete from plot_cond_rx_audit";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				ReferenceUserControlScenarioRun.lblProcValidCombos.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcValidCombos.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcValidCombos.Refresh();
				return;
			}

			/**************************
			 **get the treatment list
			 **************************/
			this.m_strSQL = "SELECT rx FROM " + this.m_strRxTable + ";"; // WHERE trim(ucase(scenario_id)) = '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
			this.m_ado.SqlQueryReader(this.m_TempMDBFileConn,this.m_strSQL);
			if (!this.m_ado.m_OleDbDataReader.HasRows)
			{
				ReferenceUserControlScenarioRun.lblProcValidCombos.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcValidCombos.Text = "!!Error!!";
				this.m_intError = -1;
				ReferenceUserControlScenarioRun.lblProcValidCombos.Refresh();
				MessageBox.Show("No Treatments Found In The Treatment Table");
				return;
			}
			while (this.m_ado.m_OleDbDataReader.Read())
			{
				strRxList+=this.m_ado.m_OleDbDataReader["rx"].ToString().Trim();
			}
			this.m_ado.m_OleDbDataReader.Close();


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute User Defined Plot SQL And Insert Resulting Records Into Table userdefinedplotfilter");
			this.m_strSQL = "INSERT INTO userdefinedplotfilter " + this.m_strUserDefinedPlotSQL;
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute User Defined Cond SQL And Insert Resulting Records Into Table userdefinedcondfilter");
			this.m_strSQL = "INSERT INTO userdefinedcondfilter " + this.m_strUserDefinedCondSQL;
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute rule definition filters for the condition table. The filters include ownership and condition accessible");
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

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute SQL that deletes from the condition rule definitions table (ruledefinitionscondfilter) those biosum_cond_id that are not found in the user defined condition SQL filter table (userdefinedcondfilter)");
			this.m_strSQL = "DELETE FROM ruledefinitionscondfilter a " + 
				            "WHERE NOT EXISTS " + 
				                "(SELECT b.biosum_cond_id " + 
								 "FROM userdefinedcondfilter b " + 
				                 "WHERE a.biosum_cond_id =  b.biosum_cond_id)";

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute SQL That Includes  Rule Definitions Defined By The User Into Table ruledefinitionsplotfilter");
			this.m_strSQL = "INSERT INTO ruledefinitionsplotfilter SELECT DISTINCT userdefinedplotfilter.* from userdefinedplotfilter INNER JOIN ruledefinitionscondfilter ON userdefinedplotfilter.biosum_plot_id = ruledefinitionscondfilter.biosum_plot_id" ;
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_strSQL = "DELETE FROM validcombos;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/**********************************************************************
			 **create valid combiniations of biosum_cond_id and treatment and 
			 **also user defined filters by owngrpcd
			 **********************************************************************/
			//insert all the possible valid plot+rx records
			m_strSQL = "INSERT INTO validcombos_fvspost SELECT a.biosum_cond_id,b.rx FROM " + this.m_strCondTable + " a," + this.m_strRxTable + " b";
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
			m_strSQL = "INSERT INTO validcombos_fvspre SELECT biosum_cond_id FROM " + this.m_strCondTable;
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);

			string strWhere="";
			for (x=0;x<=FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strTable = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.TableName(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPostVarArray[x]);
				if (strTable.Trim().Length > 0)
				{
					m_strSQL = "UPDATE validcombos_fvspost a INNER JOIN " + strTable + " b ON a.biosum_cond_id = b.biosum_cond_id AND a.rx = b.rx SET variable" + Convert.ToString(x + 1).Trim() + "_yn='Y'";
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					m_strSQL = "UPDATE validcombos_fvspost SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE variable" + Convert.ToString(x + 1).Trim() + "_yn IS NULL OR LEN(TRIM(variable" + Convert.ToString(x + 1).Trim() + "_yn))=0";
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					strWhere=strWhere + "b.variable" + Convert.ToString(x + 1).Trim() + "_yn <> 'N' AND ";

				}

				strTable = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.TableName(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPreVarArray[x]);
				if (strTable.Trim().Length > 0)
				{
					m_strSQL = "UPDATE validcombos_fvspre a INNER JOIN " + strTable + " b ON a.biosum_cond_id = b.biosum_cond_id SET variable" + Convert.ToString(x + 1).Trim() + "_yn='Y'";
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					m_strSQL = "UPDATE validcombos_fvspre SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE variable" + Convert.ToString(x + 1).Trim() + "_yn IS NULL OR LEN(TRIM(variable" + Convert.ToString(x + 1).Trim() + "_yn))=0";
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					strWhere=strWhere + "a.variable" + Convert.ToString(x + 1).Trim() + "_yn <> 'N' AND ";

				}
			}
			if (strWhere.Trim().Length > 0)
			{
				strWhere = "WHERE a.biosum_cond_id = b.biosum_cond_id AND " +  strWhere.Substring(0,strWhere.Length - 5);
			}
			

			m_strSQL = "INSERT INTO validcombos_fvsprepost " + 
				"SELECT a.biosum_cond_id,a.rx " + 
				"FROM validcombos_fvspost a, validcombos_fvspre b " + 
				strWhere;


			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);





			this.m_strSQL = "INSERT INTO validcombos (biosum_cond_id,rx) SELECT DISTINCT ruledefinitionscondfilter.biosum_cond_id,validcombos_fvsprepost.rx " +
				"FROM (((ruledefinitionsplotfilter INNER JOIN ruledefinitionscondfilter ON ruledefinitionsplotfilter.biosum_plot_id = ruledefinitionscondfilter.biosum_plot_id) " + 
				" INNER JOIN validcombos_fvsprepost ON ruledefinitionscondfilter.biosum_cond_id = validcombos_fvsprepost.biosum_cond_id) " +  
				" INNER JOIN " + this.m_strHvstCostsTable + " ON (validcombos_fvsprepost.rx = " + this.m_strHvstCostsTable + ".rx) AND (validcombos_fvsprepost.biosum_cond_id = " + this.m_strHvstCostsTable + ".biosum_cond_id)) " + 
				" INNER JOIN " + this.m_strTreeVolValSumTable + " ON (" + this.m_strHvstCostsTable + ".biosum_cond_id = " + this.m_strTreeVolValSumTable + ".biosum_cond_id) AND (" + this.m_strHvstCostsTable + ".rx = " + this.m_strTreeVolValSumTable + ".rx) AND (validcombos_fvsprepost.biosum_cond_id = " + this.m_strTreeVolValSumTable + ".biosum_cond_id) AND (validcombos_fvsprepost.rx = " + this.m_strTreeVolValSumTable + ".rx);";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Create Valid Combinations");
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos) == true) return;

			if (ReferenceUserControlScenarioRun.chkAuditTables.Checked == true)
			{
				ReferenceUserControlScenarioRun.lblMsg.Text="Creating Audit Data";
				ReferenceUserControlScenarioRun.lblMsg.Visible=true;
				ReferenceUserControlScenarioRun.lblMsg.Refresh();
				ReferenceUserControlScenarioRun.progressBar1.Minimum=0;
				ReferenceUserControlScenarioRun.progressBar1.Maximum = 16;
				ReferenceUserControlScenarioRun.progressBar1.Value=0;
				ReferenceUserControlScenarioRun.progressBar1.Visible=true;

				//BIOSUM_COND_ID RECORD AUDIT
				/******************************************************************************************
				 **insert all the plots that are being processed into the plot audit table
				 ******************************************************************************************/
				this.m_strSQL = "INSERT INTO plot_cond_audit (biosum_cond_id) SELECT ruledefinitionscondfilter.biosum_cond_id FROM ruledefinitionscondfilter INNER JOIN userdefinedplotfilter ON ruledefinitionscondfilter.biosum_plot_id = userdefinedplotfilter.biosum_plot_id";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				if (this.m_ado.m_intError != 0)
				{
					ReferenceUserControlScenarioRun.lblProcValidCombos.ForeColor = System.Drawing.Color.Red;
					ReferenceUserControlScenarioRun.lblProcValidCombos.Text = "!!Error!!";
					this.m_intError = this.m_ado.m_intError;
					ReferenceUserControlScenarioRun.lblProcValidCombos.Refresh();
					return;
				}
				ReferenceUserControlScenarioRun.progressBar1.Value=1;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				/************************************************************************
				 **check to see if the plot record exists in the frcs harvest cost table
				 ************************************************************************/
				this.m_strSQL = "UPDATE plot_cond_audit SET frcs_harvest_costs_yn = 'Y' " + 
					"WHERE plot_cond_audit.biosum_cond_id " + 
					"IN (SELECT biosum_cond_id FROM " + this.m_strHvstCostsTable + ");";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=2;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_audit SET frcs_harvest_costs_yn = 'N' " + 
					"WHERE plot_cond_audit.frcs_harvest_costs_yn IS NULL ;" ;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=3;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				/****************************************************************************
				 **check to see if the plot record exists in the validcombos_fvsprepost table
				 ****************************************************************************/
				this.m_strSQL = "UPDATE plot_cond_audit SET fvs_prepost_variables_yn = 'Y' " + 
					"WHERE plot_cond_audit.biosum_cond_id " + 
					"IN (SELECT biosum_cond_id FROM validcombos_fvsprepost);";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=4;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_audit SET fvs_prepost_variables_yn = 'N' " + 
					"WHERE plot_cond_audit.fvs_prepost_variables_yn IS NULL OR plot_cond_audit.fvs_prepost_variables_yn <> 'Y' ;"; 
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=5;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				/********************************************************************************************************
				 **check to see if the plot record exists in the processor tree volume and value tableharvest cost table
				 ********************************************************************************************************/
				this.m_strSQL = "UPDATE plot_cond_audit SET processor_tree_vol_val_yn = 'Y' " + 
					"WHERE plot_cond_audit.biosum_cond_id " + 
					"IN (SELECT biosum_cond_id FROM " + this.m_strTreeVolValSumTable + ");";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=6;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_audit SET processor_tree_vol_val_yn = 'N' " + 
					"WHERE plot_cond_audit.processor_tree_vol_val_yn IS NULL OR  plot_cond_audit.processor_tree_vol_val_yn<>'Y' ;"; 
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=7;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;


				/**********************************************************************
				 **check to see if the plot record exists in the gis travel times table
				 **********************************************************************/
				this.m_strSQL = "UPDATE plot_cond_audit SET gis_travel_times_yn = 'Y' " + 
					"WHERE plot_cond_audit.biosum_cond_id " + 
					"IN (SELECT biosum_cond_id FROM ruledefinitionscondfilter "  +
					" WHERE ruledefinitionscondfilter.biosum_plot_id " + 
					" IN (SELECT biosum_plot_id FROM " + this.m_strTravelTimeTable + "));";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=8;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;
				
				this.m_strSQL = "UPDATE plot_cond_audit SET gis_travel_times_yn = 'N' " + 
					"WHERE plot_cond_audit.gis_travel_times_yn IS NULL OR  plot_cond_audit.gis_travel_times_yn<>'Y' ;"; 
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=9;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				//BIOSUM_COND_ID + RX RECORD AUDIT
				/**********************************************************************************
				**Insert all the biosum_cond_id + rx combinations into the plot_cond_rx_audit table
				***********************************************************************************/
				this.m_strSQL = "INSERT INTO plot_cond_rx_audit (biosum_cond_id,rx)  " + 
					" SELECT a.biosum_cond_id, b.rx FROM plot_cond_audit a, " + 
					"(SELECT DISTINCT rx FROM " + this.m_strRxTable + ") b ;";  //+ 
				//" WHERE a.fvs_ffe_yn='Y' AND a.processor_tree_vol_val_yn='Y' AND a.gis_travel_times_yn='Y' AND a.frcs_harvest_costs_yn='Y';" ; 
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=10;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				/*********************************************************************************
				 **check to see if the plot + rx record exists in the fvs prepost variables table
				 *********************************************************************************/
				this.m_strSQL="UPDATE plot_cond_rx_audit SET fvs_prepost_variables_yn = 'Y' " + 
					"WHERE EXISTS (SELECT biosum_cond_id,rx " + 
					"FROM validcombos_fvsprepost " + 
					"WHERE plot_cond_rx_audit.biosum_cond_id = " + 
					"validcombos_fvsprepost.biosum_cond_id AND " + 
					"plot_cond_rx_audit.rx = " + 
					"validcombos_fvsprepost.rx);";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=11;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_rx_audit SET fvs_prepost_variables_yn = 'N' " + 
					"WHERE plot_cond_rx_audit.fvs_prepost_variables_yn IS NULL ;" ;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=12;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				/****************************************************************************
				 **check to see if the plot + rx record exists in the frcs harves costs table
				 ****************************************************************************/
				this.m_strSQL="UPDATE plot_cond_rx_audit SET frcs_harvest_costs_yn = 'Y' " + 
					"WHERE EXISTS (SELECT biosum_cond_id,rx " + 
					"FROM "  + this.m_strHvstCostsTable + " " + 
					"WHERE plot_cond_rx_audit.biosum_cond_id = " + 
					this.m_strHvstCostsTable.Trim() + ".biosum_cond_id AND " + 
					"plot_cond_rx_audit.rx = " + 
					this.m_strHvstCostsTable.Trim() + ".rx);";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=13;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_rx_audit SET frcs_harvest_costs_yn = 'N' " + 
					"WHERE plot_cond_rx_audit.frcs_harvest_costs_yn IS NULL ;" ;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=14;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				/*********************************************************************************
				 **check to see if the plot + rx record exists in the processor tree vol val table
				 *********************************************************************************/
				this.m_strSQL="UPDATE plot_cond_rx_audit SET processor_tree_vol_val_yn = 'Y' " + 
					"WHERE EXISTS (SELECT biosum_cond_id,rx " + 
					"FROM "  + this.m_strTreeVolValSumTable + " " + 
					"WHERE plot_cond_rx_audit.biosum_cond_id = " + 
					this.m_strTreeVolValSumTable.Trim() + ".biosum_cond_id AND " + 
					"plot_cond_rx_audit.rx = " + 
					this.m_strTreeVolValSumTable.Trim() + ".rx);";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=15;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_rx_audit SET processor_tree_vol_val_yn = 'N' " + 
					"WHERE plot_cond_rx_audit.processor_tree_vol_val_yn IS NULL ;" ;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				ReferenceUserControlScenarioRun.progressBar1.Value=16;
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcValidCombos)) return;


				ReferenceUserControlScenarioRun.progressBar1.Visible=false;
				ReferenceUserControlScenarioRun.lblMsg.Visible=false;
				ReferenceUserControlScenarioRun.progressBar1.Refresh();
				ReferenceUserControlScenarioRun.lblMsg.Text="";
				ReferenceUserControlScenarioRun.lblMsg.Refresh();

			}
			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcValidCombos.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcValidCombos.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcValidCombos.Refresh();
			}
			
		}

		/// <summary>
		/// evaluate the effectiveness of fvs treatment data 
		/// by loading the effective table with 
		/// results from user defined expressions 
		/// </summary>
		private void effective()
		{
			int x,y;
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";
			string[] strEffectiveColumnArray;
			string[] strBetterIsNotNull= new string[uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string[] strWorseIsNotNull= new string[uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string[] strEffectiveIsNotNull= new string[uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string strOverallEffectiveIsNotNull = "";
			string strBetterSql="";
			string strWorseSql="";
			string strEffectiveSql="";

			string strVariableNumber="";
			FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective.Variables oFvsVar =
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar;
            
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Effective Treatments");
			this.m_txtStreamWriter.WriteLine("--------------------");

			//get all the column names in the effective table
			strEffectiveColumnArray = m_ado.getFieldNamesArray(this.m_TempMDBFileConn,"select * from effective");

			/********************************************
			 **delete all records in the effective table
			 ********************************************/
			this.m_strSQL = "delete from effective";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//insert the valid combos into the effective table
			m_strSQL = "INSERT INTO effective (biosum_cond_id,rx) SELECT biosum_cond_id,rx FROM validcombos";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//insert net revenue per acre into the effective table
			m_strSQL = "UPDATE effective e " + 
				       "INNER JOIN product_yields_net_rev_costs_summary p " + 
				       "ON e.biosum_cond_id=p.biosum_cond_id AND " + 
				          "e.rx=p.rx " + 
			           "SET e.nr_dpa = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)";
			
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

						


			//populate the variable table.column name and its value to the effective table
			for (x=0;x<=uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
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
					m_strSQL = "UPDATE effective e INNER JOIN (" + strPostTable + " post INNER JOIN " + 
						strPreTable  + " pre ON post.biosum_cond_id = pre.biosum_cond_id) "  + 
						"ON e.biosum_cond_id=post.biosum_cond_id AND e.rx=post.rx " + 
						"SET " + m_strSQL;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

				}
				

			}
			//populate the change column by subtracting pre value from post value
			m_strSQL="";
			for (x=0;x<=uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strVariableNumber = Convert.ToString(x+1).Trim();
				m_strSQL = m_strSQL + "variable" + strVariableNumber + "_change=IIF(pre_variable" + strVariableNumber + "_value IS NOT NULL AND post_variable" + strVariableNumber + "_value IS NOT NULL,post_variable" + strVariableNumber + "_value - pre_variable" + strVariableNumber + "_value,null),";
			}
			m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);

			m_strSQL = "UPDATE effective SET " + m_strSQL;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			//see what variables are referenced in the sql expression and make sure they are not null
			strOverallEffectiveIsNotNull="";
			for (x=0;x<=uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
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
			for (x=0;x<=uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
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
			for (x=0;x<=uc_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
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
			strBetterSql = strBetterSql.Substring(0,strBetterSql.Length - 1);
			strBetterSql = "UPDATE effective SET " + strBetterSql;
			this.m_txtStreamWriter.WriteLine("Improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",strBetterSql);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strBetterSql);
			//worse
			strWorseSql = strWorseSql.Substring(0,strWorseSql.Length - 1);
			strWorseSql = "UPDATE effective SET " + strWorseSql;
			this.m_txtStreamWriter.WriteLine("Regression");
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",strWorseSql);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strWorseSql);
			//effective
			strEffectiveSql = strEffectiveSql.Substring(0,strEffectiveSql.Length - 1);
			strEffectiveSql = "UPDATE effective SET " + strEffectiveSql;
			this.m_txtStreamWriter.WriteLine("Variable Effective");
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",strEffectiveSql);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strEffectiveSql);
			//overall effective
			m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);
			m_strSQL = "UPDATE effective SET " + m_strSQL;
			this.m_txtStreamWriter.WriteLine("Overall Effective");
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"ForeColor",System.Drawing.Color.Red);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"Text","!!Error!!");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"Refresh");
				return;
			}

			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcEffective) == true) return;
             
            
			if (this.m_intError == 0)
			{
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"ForeColor",System.Drawing.Color.Blue);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"Text","Completed");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcEffective,"Refresh");
			}


		}

		private void optimization()
		{
			int x,y;
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";
			string strSql="";
			string[] strArray=null;
			string strVariableNumber="";

			

            
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Optimization");
			this.m_txtStreamWriter.WriteLine("--------------------");


			/********************************************
			 **delete all records in the tie breaker table
			 ********************************************/
			this.m_strSQL = "delete from optimization";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//insert the valid combos into the tiebreakre table
			m_strSQL = "INSERT INTO optimization (biosum_cond_id,rx) " + 
				         "SELECT biosum_cond_id,rx FROM effective WHERE overall_effective_yn='Y'";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//populate the variable table.column name and its value to the effective table
			if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
			{
				m_strSQL = "UPDATE optimization e " + 
					"INNER JOIN product_yields_net_rev_costs_summary p " + 
					"ON e.biosum_cond_id=p.biosum_cond_id AND " + 
					"e.rx=p.rx " + 
					"SET e.pre_variable_name = 'product_yields_net_rev_costs_summary.max_nr_dpa'," + 
					    "e.post_variable_name = 'product_yields_net_rev_costs_summary.max_nr_dpa'," + 
					    "e.pre_variable_value = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)," + 
					    "e.post_variable_value = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)," + 
					    "e.change_value = 0";

																									 ;
			}
			else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
			{
				m_strSQL = "UPDATE optimization e " + 
					"INNER JOIN product_yields_net_rev_costs_summary p " + 
					"ON e.biosum_cond_id=p.biosum_cond_id AND " + 
					"e.rx=p.rx " + 
					"SET e.pre_variable_name = 'product_yields_net_rev_costs_summary.merch_yield_cf'," + 
					"e.post_variable_name = 'product_yields_net_rev_costs_summary.merch_yield_cf'," + 
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

				m_strSQL = "UPDATE optimization e " + 
						"INNER JOIN (" + strPostTable + " a " + 
						"INNER JOIN " + strPreTable + " b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id) " + 
						"ON e.biosum_cond_id=a.biosum_cond_id AND " + 
						"e.rx=a.rx " + 
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

			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
		}


		private void tiebreaker()
		{
			int x,y;
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";
			string strSql="";
			string[] strArray=null;
			string strVariableNumber="";

            FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection =
		        ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection;

			FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem oItem;

            
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Tie Breaker");
			this.m_txtStreamWriter.WriteLine("--------------------");


			/********************************************
			 **delete all records in the tie breaker table
			 ********************************************/
			this.m_strSQL = "delete from tiebreaker";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//insert the valid combos into the tiebreakre table
			m_strSQL = "INSERT INTO tiebreaker (biosum_cond_id,rx) " + 
				          "SELECT biosum_cond_id,rx FROM effective WHERE overall_effective_yn='Y'";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

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
						strPreTable  + " pre ON post.biosum_cond_id = pre.biosum_cond_id) "  + 
						"ON e.biosum_cond_id=post.biosum_cond_id AND e.rx=post.rx " + 
						"SET " + m_strSQL;

					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

				}

				//populate the change column by subtracting pre value from post value
				m_strSQL="";
				m_strSQL = m_strSQL + "variable" + strVariableNumber + "_change=IIF(pre_variable" + strVariableNumber + "_value IS NOT NULL AND post_variable" + strVariableNumber + "_value IS NOT NULL,post_variable" + strVariableNumber + "_value - pre_variable" + strVariableNumber + "_value,null),";
				m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);

				m_strSQL = "UPDATE tiebreaker SET " + m_strSQL;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			
				
			}

			oItem = oTieBreakerCollection.Item(1);
			if (oItem.bSelected)
			{
				m_strSQL = "UPDATE tiebreaker a INNER JOIN scenario_rx_intensity b ON a.rx=b.rx SET a.rx_intensity=b.rx_intensity " + 
					       "WHERE trim(ucase(b.scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";

				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			}

		}


		/// <summary>
		/// evaluate the effectiveness of fire and fuel treatment data 
		/// by loading the effective table with 
		/// results from user defined expressions 
		/// </summary>
		private void effective_old()
		{
			string strPreTIFFEWindSpeedSQL="";
			string strPreCIFFEWindSpeedSQL="";
			string strPostTIFFEWindSpeedSQL="";
			string strPostCIFFEWindSpeedSQL="";
			string strTIChangeSQL="";
			string strCIChangeSQL="";
			string strTIEffectiveSQL="";
			string strCIEffectiveSQL="";
			string strTIBackslideSQL="";
			string strCIBackslideSQL="";
			string strAtHazardSQL="";
			string strOverallEffectiveSQL="";
			string strSelectSQL;
			int intCount=0;
			int x;
			string strScenarioMDB="";
			string strScenarioConn="";
            
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Effective Treatments");
			this.m_txtStreamWriter.WriteLine("--------------------");

			/********************************************
			 **delete all records in the effective table
			 ********************************************/
			this.m_strSQL = "delete from effective";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			ado_data_access p_ado = new ado_data_access();
			p_ado.getScenarioConnStringAndMDBFile(ref strScenarioMDB,ref strScenarioConn, frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text);
			p_ado.OpenConnection(strScenarioConn);
			if (p_ado.m_intError != 0)
			{
				this.m_intError = p_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcEffective.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcEffective.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcEffective.Refresh();
				p_ado = null;
				return;
			}     

			/**********************************************************
			 **create the torch index wind speed class sql expression
			 **********************************************************/
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.TorchIndex)
			{
				this.m_strSQL = "SELECT * FROM scenario_ffe_wind_speed WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "' AND ti_ci_index_type = 'T' ORDER BY wind_speed_class;";
            
				p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);

				while (p_ado.m_OleDbDataReader.Read())
				{
					intCount++;
					
					strPreTIFFEWindSpeedSQL +=
						"IIF(" + this.m_strCondTable + ".pre_torch_index" + p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim() + p_ado.m_OleDbDataReader["speed_mph"] + "," + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim() + ",";

					strPostTIFFEWindSpeedSQL +=
						"IIF(" + this.m_strFFETable + ".post_torch_index" + p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim() + p_ado.m_OleDbDataReader["speed_mph"] + "," + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim() + ",";
					
					                            
				
				}
				p_ado.m_OleDbDataReader.Close();
				strPreTIFFEWindSpeedSQL += intCount.ToString();
				strPostTIFFEWindSpeedSQL += intCount.ToString();
				for (x=1; x<= intCount;x++)
				{
					strPreTIFFEWindSpeedSQL += ")";
					strPostTIFFEWindSpeedSQL += ")";
				}
				strPreTIFFEWindSpeedSQL += " AS pre_ti_cl";
				strPostTIFFEWindSpeedSQL += " AS post_ti_cl";
				
			}
			//MessageBox.Show(strPreTIFFEWindSpeedSQL);


			/**********************************************************
			 **create the crown index wind speed class sql expression
			 **********************************************************/
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.CrownIndex)
			{
				this.m_strSQL = "SELECT * FROM scenario_ffe_wind_speed WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "' AND ti_ci_index_type = 'C' ORDER BY wind_speed_class;";
            
				p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);
				if (p_ado.m_intError != 0)
				{
					this.m_intError = p_ado.m_intError;
					ReferenceUserControlScenarioRun.lblProcEffective.ForeColor = System.Drawing.Color.Red;
					ReferenceUserControlScenarioRun.lblProcEffective.Text = "!!Error!!";
					ReferenceUserControlScenarioRun.lblProcEffective.Refresh();
					p_ado.m_OleDbConnection.Close();
					p_ado = null;
					return;
				}

				intCount=0;
				while (p_ado.m_OleDbDataReader.Read())
				{
					intCount++;
					strPreCIFFEWindSpeedSQL +=
						"IIF(" + this.m_strCondTable + ".pre_crown_index" + p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim() + p_ado.m_OleDbDataReader["speed_mph"] + "," + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim() + ",";

					strPostCIFFEWindSpeedSQL +=
						"IIF(" + this.m_strFFETable + ".post_crown_index" + p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim() + p_ado.m_OleDbDataReader["speed_mph"] + "," + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim() + ",";
					                            
				
				}
				p_ado.m_OleDbDataReader.Close();
				strPreCIFFEWindSpeedSQL += intCount.ToString();
				strPostCIFFEWindSpeedSQL += intCount.ToString();
				for (x=1; x<= intCount;x++)
				{
					strPreCIFFEWindSpeedSQL += ")";
					strPostCIFFEWindSpeedSQL += ")";
				}
				strPreCIFFEWindSpeedSQL += " AS pre_ci_cl";
				strPostCIFFEWindSpeedSQL += " AS post_ci_cl";
			}

			/******************************************************
			 **create the Torch and crown index change expression
			 ******************************************************/
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.TorchIndex) strTIChangeSQL = this.m_strFFETable.Trim() + ".post_torch_index - " + this.m_strCondTable.Trim() + ".pre_torch_index AS ti_change";
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.CrownIndex) strCIChangeSQL = this.m_strFFETable.Trim() + ".post_crown_index - " + this.m_strCondTable.Trim() + ".pre_crown_index AS ci_change";


			
			/**********************************************************
			 **create the torch index effective expression
			 **********************************************************/
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.TorchIndex)
			{
				this.m_strSQL = "SELECT * FROM scenario_ffe_ti_ci_effective_change WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "' AND ti_ci_index_type = 'T' ORDER BY wind_speed_class;";
            
				p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);
				intCount=0;
				while (p_ado.m_OleDbDataReader.Read())
				{
					intCount++;
					strTIEffectiveSQL +=
						"IIF(" + p_ado.m_OleDbDataReader["expression"] + ",'Y',";

				}
				p_ado.m_OleDbDataReader.Close();
				strTIEffectiveSQL += "'N'";
			
				for (x=1; x<= intCount;x++)
				{
					strTIEffectiveSQL += ")";
				}
				strTIEffectiveSQL += " AS ti_effective_yn";
			}

			/**********************************************************
			 **create the crown index effective expression
			 **********************************************************/
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.CrownIndex)
			{
				this.m_strSQL = "SELECT * FROM scenario_ffe_ti_ci_effective_change WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "' AND ti_ci_index_type = 'C' ORDER BY wind_speed_class;";
            
				p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);
				intCount=0;
				while (p_ado.m_OleDbDataReader.Read())
				{
					intCount++;
					strCIEffectiveSQL +=
						"IIF(" + p_ado.m_OleDbDataReader["expression"] + ",'Y',";

				}
				p_ado.m_OleDbDataReader.Close();
				strCIEffectiveSQL += "'N'";
			
				for (x=1; x<= intCount;x++)
				{
					strCIEffectiveSQL += ")";
				}
				strCIEffectiveSQL += " AS ci_effective_yn";
			}

			/**********************************************************
			 **create the torch index backslide expression
			 **********************************************************/
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.TorchIndex)
			{
				this.m_strSQL = "SELECT * FROM scenario_ffe_backslide WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "' AND ti_ci_index_type = 'T';";
            
				p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);

				while (p_ado.m_OleDbDataReader.Read())
				{
					strTIBackslideSQL = "IIF(ti_change" + p_ado.m_OleDbDataReader["expression_operator1"].ToString().Trim() + p_ado.m_OleDbDataReader["backslide"].ToString().Trim();
					if (p_ado.m_OleDbDataReader["expression_operator2"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["wind_speed_class"] != System.DBNull.Value && p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim().Length > 0)
						{
							if (p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim() != "0")
							{
								strTIBackslideSQL += " AND post_ti_cl" + p_ado.m_OleDbDataReader["expression_operator2"].ToString().Trim() + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
							}
						}
					}																																  
					strTIBackslideSQL += ",'Y','N') AS ti_backslide_yn";
			
				}
				p_ado.m_OleDbDataReader.Close();
			}

			/**********************************************************
			 **create the crown index backslide expression
			 **********************************************************/
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.CrownIndex)
			{
				this.m_strSQL = "SELECT * FROM scenario_ffe_backslide WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "' AND ti_ci_index_type = 'C';";
            
				p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);

				while (p_ado.m_OleDbDataReader.Read())
				{
					strCIBackslideSQL = "IIF(ci_change" + p_ado.m_OleDbDataReader["expression_operator1"].ToString().Trim() + p_ado.m_OleDbDataReader["backslide"].ToString().Trim();
					if (p_ado.m_OleDbDataReader["expression_operator2"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["wind_speed_class"] != System.DBNull.Value && p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim().Length > 0)
						{
							if (p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim() != "0")
							{
								strCIBackslideSQL += " AND post_ci_cl" + p_ado.m_OleDbDataReader["expression_operator2"].ToString().Trim() + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
							}
						}
					}																																  
					strCIBackslideSQL += ",'Y','N') AS ci_backslide_yn";
			
				}
				p_ado.m_OleDbDataReader.Close();
			}

			//MessageBox.Show(strTIBackslideSQL);



			/**********************************************************
			 **create the at hazard expression
			 **********************************************************/
			this.m_strSQL = "SELECT * FROM scenario_ffe_hazard WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
            
			p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);
			if (p_ado.m_OleDbDataReader.HasRows==true)
			{
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["expression_operator"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["wind_speed_class"] != System.DBNull.Value && p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim().Length > 0)
						{
							if (p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim() != "0")
							{
								if (p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim().ToUpper()=="T")
								{
									if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.TorchIndex)
									{
										if (strAtHazardSQL.Trim().Length == 0)
										{
											strAtHazardSQL = "IIF(pre_ti_cl" + p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim() + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
										}
										else
										{
											strAtHazardSQL += " OR pre_ti_cl" + p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim() + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
										}
									}
								}
								else
								{
									if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.CrownIndex)
									{
										if (strAtHazardSQL.Trim().Length == 0)
										{
											strAtHazardSQL = "IIF(pre_ci_cl" + p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim() + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
										}
										else
										{
											strAtHazardSQL += " OR pre_ci_cl" + p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim() + p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
										}
									}

								}
							}
						}
					}																																  
				}
			}
			p_ado.m_OleDbDataReader.Close();
			if (strAtHazardSQL.Trim().Length > 0)
			{
				strAtHazardSQL += ",'Y','N') AS at_hazard_yn";
			}
			else
			{
				strAtHazardSQL = "'N' AS at_hazard_yn";
			}



			/**********************************************************
			 **create the overall effective expression
			 **********************************************************/
			this.m_strSQL = "SELECT * FROM scenario_ffe_overall_effective_change WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
            
			p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);
			if (p_ado.m_OleDbDataReader.HasRows==true)
			{
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["expression"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["expression"].ToString().Trim().Length > 0)
						{
							strOverallEffectiveSQL = "IIF(" + p_ado.m_OleDbDataReader["expression"].ToString().Trim() + ",'Y','N')";
						}
						
					}																																  
				}
			}
			p_ado.m_OleDbDataReader.Close();
			if (strOverallEffectiveSQL.Trim().Length == 0)
				strOverallEffectiveSQL = "'N'";

			strOverallEffectiveSQL += " AS effective_yn";


			this.m_strSQL = "INSERT INTO effective (biosum_cond_id," + 
				"rx," +
				"pre_ti_cl," + 
				"pre_ci_cl," + 
				"post_ti_cl," +
				"post_ci_cl," + 
				"ti_change," + 
				"ci_change," + 
				"ti_effective_yn," + 
				"ci_effective_yn," +
				"ti_backslide_yn," +
				"ci_backslide_yn," +
				"at_hazard_yn," + 
				"effective_yn) ";

			
			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.TorchIndex==false)
			{
				strPreTIFFEWindSpeedSQL = "null AS pre_ti_cl";
				strPostTIFFEWindSpeedSQL = "null AS post_ti_cl";
				strTIChangeSQL = "null AS ti_change";
				strTIEffectiveSQL = "' ' AS ti_effective_yn";
				strTIBackslideSQL = "' ' AS ti_backslide_yn";
			}

			if (ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_ffe1.CrownIndex==false)
			{
				strPreCIFFEWindSpeedSQL = "null AS pre_ci_cl";
				strPostCIFFEWindSpeedSQL = "null AS post_ci_cl";
				strCIChangeSQL = "null AS ci_change";
				strCIEffectiveSQL = "' ' AS ci_effective_yn";
				strCIBackslideSQL = "' ' AS ci_backslide_yn";

			}

			strSelectSQL = "SELECT " + this.m_strFFETable.Trim() + ".biosum_cond_id," + 
				"validcombos.rx," + 
				strPreTIFFEWindSpeedSQL + "," + 
				strPreCIFFEWindSpeedSQL + "," + 
				strPostTIFFEWindSpeedSQL + "," + 
				strPostCIFFEWindSpeedSQL + "," + 
				strTIChangeSQL + "," + 
				strCIChangeSQL + "," +
				strTIEffectiveSQL  + "," + 
				strCIEffectiveSQL + "," + 
				strTIBackslideSQL + "," +
				strCIBackslideSQL + "," +
				strAtHazardSQL + "," + 
				strOverallEffectiveSQL + 
				" FROM " + this.m_strFFETable.Trim() + ",validcombos " + 
				" INNER JOIN " + this.m_strCondTable + 
				" ON validcombos.biosum_cond_id = " + 
				this.m_strCondTable.Trim() + ".biosum_cond_id " + 
				" WHERE (" + this.m_strFFETable.Trim() + ".biosum_cond_id=" +
				"validcombos.biosum_cond_id " + 
				"AND " + this.m_strFFETable.Trim() + ".rx=" +
				"validcombos.rx)";
											   
			

			this.m_strSQL += strSelectSQL;

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			// MessageBox.Show(this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcEffective.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcEffective.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcEffective.Refresh();
				p_ado.m_OleDbConnection.Close();
				p_ado = null;
				return;
			}

			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcEffective) == true) return;
             
            
			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcEffective.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcEffective.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcEffective.Refresh();
			}


		}

		/// <summary>
		/// get the wood product yields,
		/// revenue, and costs of an applied
		/// treatment on a plot 
		/// </summary>
		private void product_yields_net_rev_costs_summary()
		{
			string[]  strUpdateSQL = new string[25];
			int intUpdateCount=0;
			string strSum="";
			string strSumFields="";
			int x=0,y=0;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Wood Product Yields,Revenue, And Costs Table");
			this.m_txtStreamWriter.WriteLine("-----------------------------------------------");
			/**********************************************
			 **complete harvest cost per acre
			 **********************************************/
            
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
				ReferenceUserControlScenarioRun.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Refresh();
				p_ado = null;
				return;
			}     

			/**********************************************************
			 **create the at hazard expression
			 **********************************************************/
			this.m_strSQL = "SELECT * FROM scenario_costs WHERE trim(ucase(scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
            
			p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);
			double dblChipMktValPgt=0;
			double dblDriverCostPgtPerHour=0;
			if (p_ado.m_OleDbDataReader.HasRows==true)
			{
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["chip_mkt_val_pgt"] != System.DBNull.Value)
					{
						dblChipMktValPgt = Convert.ToDouble(p_ado.m_OleDbDataReader["chip_mkt_val_pgt"].ToString().Trim());
						dblDriverCostPgtPerHour= Convert.ToDouble(p_ado.m_OleDbDataReader["road_haul_cost_pgt_per_hour"].ToString().Trim());
					}																																  
				}
			}
			p_ado.m_OleDbDataReader.Close();
			


			/**********************************************
			 **sum all the expense columns to get complete
			 ** cost for each condition + treatment
			 **********************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Sum All Harvest Costs Per Acre");

			//get the user defined harvest cost columns to sum
			p_ado.m_strSQL = "SELECT columnname " + 
				"FROM scenario_harvest_cost_columns " + 
				"WHERE trim(ucase(scenario_id))='" + 
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";

			p_ado.SqlQueryReader(p_ado.m_OleDbConnection,p_ado.m_strSQL);
			string strScenarioColumnNameList="";
			string[] strScenarioColumnNameArray = null;
			string strCol="";
			strScenarioColumnNameList="harvest_cpa,";
			if (p_ado.m_OleDbDataReader.HasRows)
			{
				while (p_ado.m_OleDbDataReader.Read())
				{
					strCol="";
					//make sure the row is not null values
					if (p_ado.m_OleDbDataReader[0] != System.DBNull.Value &&
						p_ado.m_OleDbDataReader[0].ToString().Trim().Length > 0)
					{
						strCol =p_ado.m_OleDbDataReader["ColumnName"].ToString().Trim();
						strScenarioColumnNameList = strScenarioColumnNameList + strCol + ",";
					}
				}
			}
			strScenarioColumnNameList = strScenarioColumnNameList.Substring(0,strScenarioColumnNameList.Length - 1);
			strScenarioColumnNameArray = this.m_oUtils.ConvertListToArray(strScenarioColumnNameList,",");
			p_ado.m_OleDbDataReader.Close();
			p_ado.m_OleDbConnection.Close();
			p_ado = null;
			

			//get the table schema data from the harvest costs table
			System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,"select * from " + this.m_strHvstCostsTable.Trim());

			//load up the harvest cost columns
			for (x=0; x<=p_dt.Rows.Count-1; x++)
			{
				if (strScenarioColumnNameArray != null && strScenarioColumnNameArray.Length > 0)
				{
					for (y=0;y<=strScenarioColumnNameArray.Length - 1;y++)
					{
						if (strScenarioColumnNameArray[y].Trim().ToUpper() == 
							p_dt.Rows[x]["columnname"].ToString().Trim().ToUpper())
						{
							strSumFields = this.m_strHvstCostsTable.Trim() + "." + p_dt.Rows[x]["columnname"].ToString().Trim();
							strUpdateSQL[intUpdateCount] = " UPDATE " + this.m_strHvstCostsTable.Trim() + " SET " + strSumFields + " = 0  WHERE (((" + strSumFields.Trim() + ") Is Null));";
							if (strSum.Trim().Length == 0)
							{
								strSum =  p_dt.Rows[x]["columnname"].ToString().Trim();
							}
							else
							{
								strSum += "+"  + p_dt.Rows[x]["columnname"].ToString().Trim();
							}
							intUpdateCount++;
							break;
						}
					}
				}
			}
			this.m_ado.m_OleDbDataReader.Close();

			/******************************************************
			 **update the sum fields to 0 for those that are null
			 ******************************************************/
			for (x=0;x<=intUpdateCount-1;x++)
			{
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strUpdateSQL[x]);
				if (this.m_ado.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					ReferenceUserControlScenarioRun.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
					ReferenceUserControlScenarioRun.lblSumWoodProducts.Text = "!!Error!!";
					ReferenceUserControlScenarioRun.lblSumWoodProducts.Refresh();
					return;
				}
			}
			
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_strSQL = "INSERT INTO harvest_costs_sum " + 
				" SELECT biosum_cond_id,rx, " + strSum + " AS complete_cpa FROM " + this.m_strHvstCostsTable + ";";

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError !=0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Refresh();
				return;

			}



			// MessageBox.Show(this.m_strSQL);
			/****************************************************************
			 **sum all columns beginning in column 4 and write value to the
			 **harvest_cpa column
			 ****************************************************************/
			
			this.m_strSQL = "UPDATE " + this.m_strHvstCostsTable.Trim() + 
				" INNER JOIN harvest_costs_sum " + 
				" ON (" + this.m_strHvstCostsTable.Trim() +  ".biosum_cond_id=" + 
				"harvest_costs_sum.biosum_cond_id) AND (" +
				this.m_strHvstCostsTable.Trim() + ".rx=" + 
				"harvest_costs_sum.rx) " + 
				"SET " + this.m_strHvstCostsTable.Trim() + ".complete_cpa=harvest_costs_sum.complete_cpa;";
				
				
			//MessageBox.Show(this.m_strSQL);
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Refresh();
				return;
			}


			/********************************************
			 **delete all records in the table
			 ********************************************/
			this.m_strSQL = "delete from product_yields_net_rev_costs_summary";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Refresh();
				return;
			}
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert Records");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_strSQL = "INSERT INTO product_yields_net_rev_costs_summary " + 
				"(biosum_cond_id,rx,merch_yield_cf,chip_yield_cf,chip_yield_gt,chip_val_dpa," + 
				"merch_yield_gt,merch_val_dpa,harvest_onsite_cpa," +
				"haul_merch_cpa,haul_chip_cpa,merch_chip_nr_dpa," +
				"merch_nr_dpa,usebiomass_yn,max_nr_dpa) ";

			string strSelectSQL = "SELECT validcombos.biosum_cond_id,validcombos.rx," + 
				this.m_strTreeVolValSumTable.Trim() + ".merch_vol_cf AS merch_yield_cf," +
				this.m_strTreeVolValSumTable.Trim() + ".chip_vol_cf AS chip_yield_cf," +
				this.m_strTreeVolValSumTable.Trim() + ".chip_wt_gt AS chip_yield_gt," +
				this.m_strTreeVolValSumTable.Trim() + ".chip_val_dpa AS chip_val_dpa," + 
				this.m_strTreeVolValSumTable.Trim() + ".merch_wt_gt AS merch_yield_gt," +
				this.m_strTreeVolValSumTable.Trim() + ".merch_val_dpa AS merch_val_dpa," +

				this.m_strHvstCostsTable.Trim() + ".complete_cpa AS harvest_onsite_cpa," + 
				"IIF(" + this.m_strPlotTable.Trim() + ".merch_haul_cpa_pt IS NOT NULL," + this.m_strPlotTable.Trim() + ".merch_haul_cpa_pt * " + this.m_strTreeVolValSumTable.Trim() + ".merch_wt_gt,0) AS haul_merch_cpa," + 
				"IIF(" + this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt IS NOT NULL," + this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt * " + this.m_strTreeVolValSumTable.Trim() + ".chip_wt_gt,0) AS haul_chip_cpa," + 
				"merch_val_dpa + chip_val_dpa - harvest_onsite_cpa - haul_merch_cpa - haul_chip_cpa AS merch_chip_nr_dpa," + 
				"merch_val_dpa - harvest_onsite_cpa - haul_merch_cpa AS merch_nr_dpa," + 
				"IIF(" + this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt IS NOT NULL AND merch_chip_nr_dpa > merch_nr_dpa,'Y','N') AS usebiomass_yn," + 
				"IIF(usebiomass_yn = 'Y', merch_chip_nr_dpa,merch_nr_dpa) AS max_nr_dpa " + 
				"FROM ((((validcombos INNER JOIN " + this.m_strCondTable.Trim() + " ON validcombos.biosum_cond_id = " + this.m_strCondTable.Trim() + ".biosum_cond_id) " +  
				"INNER JOIN " + this.m_strPlotTable.Trim() + " ON " + this.m_strCondTable.Trim() + ".biosum_plot_id=" + this.m_strPlotTable.Trim() + ".biosum_plot_id) " +
				"INNER JOIN " + this.m_strTreeVolValSumTable.Trim() + " ON (validcombos.biosum_cond_id=" + this.m_strTreeVolValSumTable.Trim() + ".biosum_cond_id) AND " +
				"(validcombos.rx=" + this.m_strTreeVolValSumTable.Trim() + ".rx)) " + 
				"INNER JOIN " + this.m_strHvstCostsTable.Trim() + " ON (validcombos.biosum_cond_id=" + this.m_strHvstCostsTable.Trim() + ".biosum_cond_id) AND " +
				"(validcombos.rx=" + this.m_strHvstCostsTable.Trim() + ".rx)); ";
			this.m_strSQL+=strSelectSQL;
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Refresh();
				return;
			}

			if (this.UserCancel(ReferenceUserControlScenarioRun.lblSumWoodProducts) == true) return;

			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblSumWoodProducts.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Text = "Completed";
				ReferenceUserControlScenarioRun.lblSumWoodProducts.Refresh();
			}
     

		}

		/// <summary>
		/// create a temporary work table for summing harvest costs
		/// </summary>
		private void CreateTableStructureOfHarvestCosts()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn = p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");

			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the harvest_costs structure
				 *********************************************/
				this.m_strSQL = "SELECT biosum_cond_id,rx, complete_cpa FROM " + this.m_strHvstCostsTable.Trim() + ";";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of harvest_costs_sum
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create harvest_costs_sum Table Schema From harvest_costs table");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"harvest_costs_sum",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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

				this.m_txtStreamWriter.WriteLine("Create scenario psites work table from  the scenario_psites table");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"scenario_psites_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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

				this.m_txtStreamWriter.WriteLine("Create haul costs work tables from  the haul_costs table");

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
					
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		/// a plot's cheapest merch and chip haul cost
		/// </summary>
		private void CreateTableStructureForHaulCostsOld()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");

			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the haul costs table structure
				 *********************************************/
				this.m_strSQL = "SELECT biosum_plot_id,railhead_id,psite_id,transfer_cost,road_cost,rail_cost,total_haul_cost,materialcd FROM haul_costs;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

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
				this.m_txtStreamWriter.WriteLine("Create haul costs work tables from  the haul_costs table");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"all_road_merch_haul_costs_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"all_road_chip_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_road_merch_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_road_chip_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_rail_merch_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_rail_chip_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_merch_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_chip_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"merch_plot_to_rh_to_collector_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"chip_plot_to_rh_to_collector_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"combine_chip_rail_road_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"combine_merch_rail_road_haul_costs_work_table",p_dt,true);
				p_dt.Clear();

				

				/*****************************************************************
				 **create the table structures in the temp mdb file
				 **and give them the name OF merch_rh_to_collector_haul_costs_work_table
				 **                          chip_rh_to_collector_haul_costs_work_table
				 *****************************************************************/
				/*********************************************
				 **get the haul cost table structure
				 *********************************************/
				this.m_strSQL = "SELECT railhead_id,psite_id,transfer_cost,road_cost,rail_cost,total_haul_cost,materialcd FROM haul_costs;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"merch_rh_to_collector_haul_costs_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"chip_rh_to_collector_haul_costs_work_table",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		/// a plot's fastest travel time to a processing site
		/// </summary>
		private void CreateTableStructureForFastestTravelTimes()
		{
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
				this.m_txtStreamWriter.WriteLine("Create travel times work table (plot_fastest_tvltm_work_table) from  the plot table");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_tvltm_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
				this.m_txtStreamWriter.WriteLine("Create haul costs work table (haul_costs_work_table) from  the haul_costs table");
				
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"haul_costs_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
			ado_data_access oAdo = new ado_data_access();
			this.m_strConn= oAdo.getMDBConnString(this.m_strTempMDBFile,"admin","");

			oAdo.OpenConnection(this.m_strConn);


			if (oAdo.m_intError==0)
			{
				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create rx_intensity_work_table Schema");

				frmMain.g_oTables.m_oCoreScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_duplicates_work_table");
				if (oAdo.m_intError==0) frmMain.g_oTables.m_oCoreScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_duplicates_work_table2");
				if (oAdo.m_intError==0) frmMain.g_oTables.m_oCoreScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_unique_work_table");

				if (oAdo.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
				this.m_txtStreamWriter.WriteLine("Create plot_cond_accessible_work_table Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_cond_accessible_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_cond_accessible_work_table2",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
				this.m_txtStreamWriter.WriteLine("Create userdefinedcondfilter Table Schema From User Defined Condition Filter SQL");
				dao_data_access p_dao = new dao_data_access();
				//LPOTTS p_dao.CreateMDBTableFromDataSetTable(ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb","userdefinedplotfilter",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strSystemResultsDbPathAndFile,"userdefinedcondfilter",p_dt,true);
				p_dt.Dispose();
				this.m_ado.m_OleDbDataReader.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					return;
				}
				/***********************************************************************
				 **create a table link to the user defined plot filter table in the
				 **temporary MDB file located on the hard drive of the user
				 ***********************************************************************/
				this.m_txtStreamWriter.WriteLine("Create link in " + this.m_strTempMDBFile);
				p_dao.CreateTableLink(this.m_strTempMDBFile,"userdefinedcondfilter",m_strSystemResultsDbPathAndFile,"userdefinedcondfilter",true);
				this.m_txtStreamWriter.WriteLine("{0}\t{1}",m_strSystemResultsDbPathAndFile,"userdefinedcondfilter");


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
				this.m_txtStreamWriter.WriteLine("Create ruledefinitionscondfilter Schema");
				//dao_data_access p_dao = new dao_data_access();
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
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
				this.m_txtStreamWriter.WriteLine("Create max_nr_plots Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"psite_sum_work_table",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
				this.m_txtStreamWriter.WriteLine("Create max_nr_plots Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"own_sum_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"own_sum_work_table_air_dest",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		private void best_rx_summary()
		{
			int x;
			string strTable="";
			string strSql="";
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Finding Best Treatments: Maximum Net Revenue");
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",true);
			frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Refresh");
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Minimum",0);
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Maximum",8);
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",0);
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Visible",true);
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Rx Summary");
			this.m_txtStreamWriter.WriteLine("--------------------\r\n");
			
			FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection =
				ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection;


		    string strScenarioId = this.ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim();
			string strTieBreakerAggregate="MAX";
			//string strOptimizationAggregate="MAX";
			//string strOptimizationTable="";
			//string strOptimizationColumnName="";
			//string strAggregateColumnName="";
			
			

	
			//
			//CREATE WORK TABLES
			//
			//best_rx_summary_work_table
			if (m_ado.TableExist(this.m_TempMDBFileConn,
				frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryTableName + "_work_table"))
			{
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,
					"DROP TABLE " + frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryTableName + "_work_table");
			}
			strTable = frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsBestRxSummaryTableName + "_work_table";
			m_ado.SqlNonQuery(m_TempMDBFileConn,frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryTableSQL(strTable));
            //best_rx_summury_optimization_and_tiebreaker_work_table
			if (m_ado.TableExist(m_TempMDBFileConn,"best_rx_summury_optimization_and_tiebreaker_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE best_rx_summury_optimization_and_tiebreaker_work_table");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryTieBreakerTableSQL("best_rx_summury_optimization_and_tiebreaker_work_table"));
			//best_rx_summury_optimization_and_tiebreaker_work_table2
			if (m_ado.TableExist(m_TempMDBFileConn,"best_rx_summury_optimization_and_tiebreaker_work_table2"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE best_rx_summury_optimization_and_tiebreaker_work_table2");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryTieBreakerTableSQL("best_rx_summury_optimization_and_tiebreaker_work_table2"));
			//best_rx_summury_optimization_and_tiebreaker_work_table3
			if (m_ado.TableExist(m_TempMDBFileConn,"best_rx_summury_optimization_and_tiebreaker_work_table3"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE best_rx_summury_optimization_and_tiebreaker_work_table3");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,frmMain.g_oTables.m_oCoreScenarioResults.CreateBestRxSummaryTieBreakerTableSQL("best_rx_summury_optimization_and_tiebreaker_work_table3"));

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
				frmMain.g_oTables.m_oCoreScenarioResults.DefaultScenarioResultsEffectiveTableName + " e " + 
				"WHERE c.biosum_plot_id = p.biosum_plot_id  AND " + 
				"e.biosum_cond_id = c.biosum_cond_id AND " + 
				"e.overall_effective_yn = 'Y' AND " + 
				"(p.merch_haul_cost_id IS NOT NULL OR " + 
				"p.chip_haul_cost_id IS NOT NULL);";

			this.m_txtStreamWriter.WriteLine("insert condition records that have MERCH or CHIP haul costs into best_rx_summary ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL); this.m_txtStreamWriter.WriteLine(" ");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
				return;
			}

			if (m_ado.TableExist(this.m_TempMDBFileConn,"effective_product_yields_net_rev_costs_summary"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE effective_product_yields_net_rev_costs_summary");
			}

			m_ado.m_strSQL = "SELECT p.* " + 
				"INTO effective_product_yields_net_rev_costs_summary " + 
				"FROM product_yields_net_rev_costs_summary p,effective e " + 
				"WHERE p.biosum_cond_id = e.biosum_cond_id AND " + 
				"p.rx=e.rx  AND e.overall_effective_yn='Y'";


			this.m_txtStreamWriter.WriteLine("write effective treatments to the effective_product_yields_net_rev_costs_summary ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			this.m_txtStreamWriter.WriteLine(" ");
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
				return;
			}

			if (m_ado.TableExist(this.m_TempMDBFileConn,"effective_optimization_treatments"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE effective_optimization_treatments");
			}

			m_ado.m_strSQL = "SELECT e.* " + 
				"INTO effective_optimization_treatments " + 
				"FROM effective e " + 
				"WHERE e.overall_effective_yn='Y'";


			this.m_txtStreamWriter.WriteLine("write effective treatments to the effective_fvs_variables ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			this.m_txtStreamWriter.WriteLine(" ");
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
				return;
			}
			

			string strAggregateWhere="";
			
			//if (oTieBreakerCollection.Item(0).strMaxYN=="Y") strTieBreakerAggregate="MAX";
			//else strTieBreakerAggregate="MIN";

			//if (this.m_oOptimizationVariable.strMaxYN=="Y") strOptimizationAggregate="MAX";
			//else strOptimizationAggregate="MIN";


			if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
			{
				//strOptimizationColumnName="max_nr_dpa";
				//strOptimizationTable = "product_yields_net_rev_costs_summary";
				//strAggregateColumnName =strOptimizationAggregate + "_optimization_variable";
				best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,false);

			}
			else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
			{
				//strOptimizationColumnName = "merch_yield_cf";
				//strOptimizationTable = "product_yields_net_rev_costs_summary";
				//strAggregateColumnName =strOptimizationAggregate + "_optimization_variable";
				best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,false);
			}
			else
			{
				//strOptimizationColumnName = "effective";
				//if (this.m_oOptimizationVariable.str
				//strOptimizationTable = "product_yields_net_rev_costs_summary";
				//strAggregateColumnName =strOptimizationAggregate + "_optimization_variable";
				best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,true);
			}

			/*************************************************************************************
			 **AIR CURTAIN DESTRUCTION PLOTS
			 **insert records into the air curtain destruction table that were processed in the 
			 **best_rx_summary table. The records to insert will be plots that do not 
			 **have a place to transport the chips (biomass) so they must be burned on site
			 *************************************************************************************
			 /**************************************************************************************
			 **insert air destruction plots with no haul costs
			 **************************************************************************************/
			this.m_strSQL = "INSERT INTO best_rx_summary_air_dest " +
				"SELECT DISTINCT c.biosum_cond_id,c.acres,c.owngrpcd,e.optimization_value,e.tiebreaker_value,e.rx_intensity,e.rx " + 
				"FROM " + this.m_strCondTable.Trim() + " c, " + 
				this.m_strPlotTable.Trim() + " p, "  + 
				"best_rx_summary e " + 
				"WHERE c.biosum_plot_id = p.biosum_plot_id  AND " + 
				"e.biosum_cond_id = c.biosum_cond_id AND " + 
				"(p.chip_haul_cost_id IS NULL);";

			this.m_txtStreamWriter.WriteLine("insert air curtain destruction plots from the  best_rx_summary table to the best_rx_summary_air_dest");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_txtStreamWriter.WriteLine(" ");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
				return;
			}

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.progressBar1,"Value",8);
			/****************************************************************************
			 **finished with minimum merchantable wood removal with positive net revenue
			 ****************************************************************************/


			if (this.m_intError == 0)
			{
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Blue);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","Completed");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
			}


			//

			
			



		
		}

		private void best_rx_summary(
			FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection, 
			string strTieBreakerAggregate,
			bool bFVSVariable)
		{
			string strSql="";


			if (bFVSVariable==false)
			{
				//find the treatment for each plot that produces the MAX/MIN revenue value
				strSql = "SELECT a.biosum_cond_id,a.rx,a." + this.m_strOptimizationColumnNameSql + " AS optimization_value " + //LPOTTS,a.rx_intensity " + 
					"FROM optimization a,";


				strSql=strSql + "(SELECT " + this.m_strOptimizationAggregateSql + "(" + this.m_strOptimizationColumnNameSql + ") AS " + this.m_strOptimizationAggregateColumnName + ",biosum_cond_id " + 
					"FROM optimization";

				strSql = strSql + " GROUP BY biosum_cond_id) b ";

				strSql = strSql + "WHERE a.biosum_cond_id=b.biosum_cond_id AND a." + this.m_strOptimizationColumnNameSql + " = b." + this.m_strOptimizationAggregateColumnName;
			}
			else
			{
				strSql = "SELECT a.biosum_cond_id,a.rx,a." + this.m_strOptimizationColumnNameSql + " AS optimization_value " + //LPOTTS,a.rx_intensity " + 
					"FROM optimization a,";


				strSql=strSql + "(SELECT " + this.m_strOptimizationAggregateSql + "(" + this.m_strOptimizationColumnNameSql + ") AS " + this.m_strOptimizationAggregateColumnName + ",biosum_cond_id " + 
					"FROM optimization";

				strSql = strSql + " GROUP BY biosum_cond_id) b ";

				strSql = strSql + "WHERE a.biosum_cond_id=b.biosum_cond_id AND a." + this.m_strOptimizationColumnNameSql + " = b." + this.m_strOptimizationAggregateColumnName;
			}

			strSql = "INSERT INTO best_rx_summury_optimization_and_tiebreaker_work_table " + 
				strSql;
		

			this.m_txtStreamWriter.WriteLine("filter effective treatments to find " + this.m_strOptimizationAggregateSql + " " + this.m_oOptimizationVariable.strOptimizedVariable);
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + strSql);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSql);
			this.m_txtStreamWriter.WriteLine(" ");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
				frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
				return;
			}

				

			m_ado.m_strSQL = "UPDATE best_rx_summury_optimization_and_tiebreaker_work_table a " + 
				"INNER JOIN best_rx_summary_work_table b " + 
				"ON a.biosum_cond_id=b.biosum_cond_id " + 
				"SET a.acres = b.acres,a.owngrpcd=b.owngrpcd";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			if (oTieBreakerCollection.Item(0).bSelected && 
				oTieBreakerCollection.Item(1).bSelected)
			{
				//update the tiebreaker and rx intensity fields for each plot
				if (oTieBreakerCollection.Item(0).strValueSource=="POST")
				{
					strSql = "UPDATE best_rx_summury_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.post_variable1_value,a.rx_intensity=b.rx_intensity";
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
						
				}
				else if (oTieBreakerCollection.Item(0).strValueSource=="PRE")
				{
					strSql = "UPDATE best_rx_summury_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.pre_variable1_value,a.rx_intensity=b.rx_intensity";
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
				}
				else if (oTieBreakerCollection.Item(0).strValueSource=="POST-PRE")
				{
					strSql = "UPDATE best_rx_summury_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.variable1_change,a.rx_intensity=b.rx_intensity";
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
				}

				m_ado.m_strSQL ="INSERT INTO best_rx_summary_before_tiebreaks SELECT * FROM best_rx_summury_optimization_and_tiebreaker_work_table";
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);


				//find the treatment for each plot that produces the MAX/MIN tiebreaker value
				m_ado.m_strSQL ="SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value,a.tiebreaker_value,a.rx,a.rx_intensity " + 
					"FROM best_rx_summury_optimization_and_tiebreaker_work_table a," +
					"(SELECT biosum_cond_id," + strTieBreakerAggregate + "(tiebreaker_value) AS tiebreaker " +
					"FROM best_rx_summury_optimization_and_tiebreaker_work_table " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.tiebreaker_value=c.tiebreaker";


					
				m_ado.m_strSQL = "INSERT INTO best_rx_summury_optimization_and_tiebreaker_work_table2 " + m_ado.m_strSQL;


				this.m_txtStreamWriter.WriteLine("break any ties by finding the " + strTieBreakerAggregate + " tie breaker value");
				this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				this.m_txtStreamWriter.WriteLine(" ");
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
				if (this.m_ado.m_intError != 0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
					return;
				}
					


				m_ado.m_strSQL = "SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value," + 
					"a.tiebreaker_value,a.rx,a.rx_intensity " + 
					"FROM best_rx_summury_optimization_and_tiebreaker_work_table2 a," +
					"(SELECT biosum_cond_id,MIN(rx_intensity) AS min_intensity " + 
					"FROM best_rx_summury_optimization_and_tiebreaker_work_table2 " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.rx_intensity=c.min_intensity";

				m_ado.m_strSQL = "INSERT INTO best_rx_summury_optimization_and_tiebreaker_work_table3 " + m_ado.m_strSQL;


				this.m_txtStreamWriter.WriteLine("break any additional ties by finding the least intense treatment");
				this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				this.m_txtStreamWriter.WriteLine(" ");
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
				if (this.m_ado.m_intError != 0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
					return;
				}
					

				m_ado.m_strSQL = "UPDATE best_rx_summary_work_table a " + 
					"INNER JOIN best_rx_summury_optimization_and_tiebreaker_work_table3 b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id " + 
					"SET a.optimization_value=b.optimization_value," + 
					"a.tiebreaker_value=b.tiebreaker_value," + 
					"a.rx=b.rx," + 
					"a.rx_intensity=b.rx_intensity";
				m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
				m_ado.m_strSQL="INSERT INTO best_rx_summary SELECT * FROM best_rx_summary_work_table";
				this.m_txtStreamWriter.WriteLine("insert the work table records into the best_rx_summary table");
				this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				this.m_txtStreamWriter.WriteLine(" ");
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
				if (this.m_ado.m_intError != 0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
					return;
				}

			}
			else if (oTieBreakerCollection.Item(0).bSelected)
			{
				//update the tiebreaker and rx intensity fields for each plot
				if (oTieBreakerCollection.Item(0).strValueSource=="POST")
				{
					strSql = "UPDATE best_rx_summury_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.post_variable1_value,a.rx_intensity=b.rx_intensity";
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
						
				}
				else if (oTieBreakerCollection.Item(0).strValueSource=="PRE")
				{
					strSql = "UPDATE best_rx_summury_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.pre_variable1_value,a.rx_intensity=b.rx_intensity";
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
				}
				else if (oTieBreakerCollection.Item(0).strValueSource=="POST-PRE")
				{
					strSql = "UPDATE best_rx_summury_optimization_and_tiebreaker_work_table a " + 
						"INNER JOIN tiebreaker b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
						"SET a.tiebreaker_value = b.variable1_change,a.rx_intensity=b.rx_intensity";
					m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);
				}

				m_ado.m_strSQL ="INSERT INTO best_rx_summary_before_tiebreaks SELECT * FROM best_rx_summury_optimization_and_tiebreaker_work_table";
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);


				//find the treatment for each plot that produces the MAX/MIN tiebreaker value
				m_ado.m_strSQL ="SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value,a.tiebreaker_value,a.rx,a.rx_intensity " + 
					"FROM best_rx_summury_optimization_and_tiebreaker_work_table a," +
					"(SELECT biosum_cond_id," + strTieBreakerAggregate + "(tiebreaker_value) AS tiebreaker " +
					"FROM best_rx_summury_optimization_and_tiebreaker_work_table " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.tiebreaker_value=c.tiebreaker";


					
				m_ado.m_strSQL = "INSERT INTO best_rx_summury_optimization_and_tiebreaker_work_table2 " + m_ado.m_strSQL;


				this.m_txtStreamWriter.WriteLine("break any ties by finding the " + strTieBreakerAggregate + " tie breaker value");
				this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				this.m_txtStreamWriter.WriteLine(" ");
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
				if (this.m_ado.m_intError != 0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
					return;
				}
				m_ado.m_strSQL = "UPDATE best_rx_summary_work_table a " + 
					"INNER JOIN best_rx_summury_optimization_and_tiebreaker_work_table2 b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id " + 
					"SET a.optimization_value=b.optimization_value," + 
					"a.tiebreaker_value=b.tiebreaker_value," + 
					"a.rx=b.rx," + 
					"a.rx_intensity=b.rx_intensity";
				m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
				m_ado.m_strSQL="INSERT INTO best_rx_summary SELECT * FROM best_rx_summary_work_table";
				this.m_txtStreamWriter.WriteLine("insert the work table records into the best_rx_summary table");
				this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				this.m_txtStreamWriter.WriteLine(" ");
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
				if (this.m_ado.m_intError != 0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
					return;
				}

			}
			else if (oTieBreakerCollection.Item(1).bSelected)
			{
				//update the rx intensity fields for each plot
				strSql = "UPDATE best_rx_summury_optimization_and_tiebreaker_work_table a " + 
					"INNER JOIN tiebreaker b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " + 
					"SET a.rx_intensity=b.rx_intensity";
				m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);


				m_ado.m_strSQL ="INSERT INTO best_rx_summary_before_tiebreaks SELECT * FROM best_rx_summury_optimization_and_tiebreaker_work_table";
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);


					
				m_ado.m_strSQL = "SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value," + 
					"a.tiebreaker_value,a.rx,a.rx_intensity " + 
					"FROM best_rx_summury_optimization_and_tiebreaker_work_table a," +
					"(SELECT biosum_cond_id,MIN(rx_intensity) AS min_intensity " + 
					"FROM best_rx_summury_optimization_and_tiebreaker_work_table " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.rx_intensity=c.min_intensity";

				m_ado.m_strSQL = "INSERT INTO best_rx_summury_optimization_and_tiebreaker_work_table2 " + m_ado.m_strSQL;


				this.m_txtStreamWriter.WriteLine("break any additional ties by finding the least intense treatment");
				this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				this.m_txtStreamWriter.WriteLine(" ");
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
				if (this.m_ado.m_intError != 0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
					return;
				}
					

				m_ado.m_strSQL = "UPDATE best_rx_summary_work_table a " + 
					"INNER JOIN best_rx_summury_optimization_and_tiebreaker_work_table2 b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id " + 
					"SET a.optimization_value=b.optimization_value," + 
					"a.tiebreaker_value=b.tiebreaker_value," + 
					"a.rx=b.rx," + 
					"a.rx_intensity=b.rx_intensity";
				m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
				m_ado.m_strSQL="INSERT INTO best_rx_summary SELECT * FROM best_rx_summary_work_table";
				this.m_txtStreamWriter.WriteLine("insert the work table records into the best_rx_summary table");
				this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				this.m_txtStreamWriter.WriteLine(" ");
				if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
				if (this.m_ado.m_intError != 0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"ForeColor",System.Drawing.Color.Red);
					frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Text","!!Error!!");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblProcBestRx,"Refresh");
					return;
				}

			}


		}
	
		/// <summary>
		/// find the best treatment by these categories: 
		/// maximum net revenue;  maximum positive net revenue;
		/// maximum torch/crown index improvement; maximum 
		/// torch /crown index improvement with positive net revenue;
		/// minimum merchantable removal; minimum merchantable removal
		/// with positive net revenue 
		/// </summary>
		private void best_rx_summary_old()
		{
			ReferenceUserControlScenarioRun.lblMsg.Text="Finding Best Treatments: Maximum Net Revenue";
			ReferenceUserControlScenarioRun.lblMsg.Visible=true;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Minimum=0;
			ReferenceUserControlScenarioRun.progressBar1.Maximum = 8;
			ReferenceUserControlScenarioRun.progressBar1.Value=0;
			ReferenceUserControlScenarioRun.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Rx Summary");
			this.m_txtStreamWriter.WriteLine("--------------------\r\n");
			

			/*************************************************
			 **delete all records in the best_rx_summary table
			 *************************************************/
			this.m_strSQL = "delete from best_rx_summary";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/*************************************************
			 **delete all records in the best_rx_summary table
			 *************************************************/
			this.m_strSQL = "delete from best_rx_summary_air_dest";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            
			/****************************************************************
			 **delete all records in the rx_intensity_duplicates_work_table
			 ****************************************************************/
			this.m_strSQL = "delete from rx_intensity_duplicates_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/****************************************************************
			 **delete all records in the rx_intensity_unique_work_table
			 ****************************************************************/
			this.m_strSQL = "delete from rx_intensity_unique_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			/**********************************************
			 **insert unique biosum_cond_id's into the
			 **best_rx_summary table so we dont have
			 **to worry about whether the biosum_cond_id 
			 **record is in the table or not
			 **********************************************/

			this.m_strSQL = "INSERT INTO best_rx_summary " + 
				"SELECT DISTINCT c.biosum_cond_id,c.acres,c.owngrpcd " + 
				"FROM " + this.m_strCondTable.Trim() + " c, " + 
				this.m_strPlotTable.Trim() + " p "  + 
				"WHERE c.biosum_plot_id = p.biosum_plot_id  AND " + 
				"(p.merch_haul_cost_id IS NOT NULL OR " + 
				"p.chip_haul_cost_id IS NOT NULL);";

			this.m_txtStreamWriter.WriteLine("insert condition records that have haul costs into best_rx_summary ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			/**************************************************************************************
			 **insert air destruction plots with no haul costs
			 **but they do have ci and ti treatment
			 **************************************************************************************/
			this.m_strSQL = "INSERT INTO best_rx_summary_air_dest " + 
				"SELECT DISTINCT c.biosum_cond_id,c.acres,c.owngrpcd " + 
				"FROM " + this.m_strCondTable.Trim() + " c, " + 
				this.m_strPlotTable.Trim() + " p "  + 
				"WHERE c.biosum_plot_id = p.biosum_plot_id  AND " + 
				"(p.merch_haul_cost_id IS NULL);";

			this.m_txtStreamWriter.WriteLine("insert condition records that have no MERCH haul costs into best_rx_summary_air_dest ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}



			//NET REVENUE
			this.m_txtStreamWriter.WriteLine("\r\n--NET REVENUE--");
			/*************************************************
			 **best treatment for max net revenue
			 *************************************************/

			/**********************************************************
			 **could have 2 different treatments that produce
			 **the same net revenue for a plot so insert into duplicates
			 **work table
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine("                           ");

			if (m_ado.TableExist(this.m_TempMDBFileConn,"effective_product_yields_net_rev_costs_summary"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE effective_product_yields_net_rev_costs_summary");
			}

			//write effective treatments to the effective_product_yields_net_rev_costs_summary
			m_ado.m_strSQL = "SELECT p.*,i.rx_intensity " + 
				"INTO effective_product_yields_net_rev_costs_summary " + 
				"FROM product_yields_net_rev_costs_summary p,effective e, scenario_rx_intensity i " + 
				"WHERE p.biosum_cond_id = e.biosum_cond_id AND " + 
				"p.rx=e.rx  AND e.effective_yn='Y' AND i.rx=p.rx AND " + 
				"trim(ucase(i.scenario_id))='" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";

			this.m_txtStreamWriter.WriteLine("write effective to the effective_product_yields_net_rev_costs_summary ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			


			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_revenue_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_revenue_min_intensity_work_table");
				
			}
           

			m_ado.m_strSQL="CREATE TABLE max_revenue_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"max_nr_dpa DOUBLE," + 
				"rx_intensity SHORT)";

			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			
			m_ado.m_strSQL = "INSERT INTO max_revenue_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.max_nr_dpa,a.rx_intensity " + 
				"FROM effective_product_yields_net_rev_costs_summary a, " + 
				"(SELECT biosum_cond_id,MAX(max_nr_dpa) AS max_revenue " + 
				"FROM  effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  max_nr_dpa, MIN(rx_intensity) AS min_intensity  " + 
				"FROM effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id,max_nr_dpa) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_revenue=a.max_nr_dpa";
			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with maximum net revenue at minimal intensity to max_revenue_min_intensity_work_table");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}


            
			/*******************************************************
			 **update the best_rx_summary from the 
			 **max_revenue_min_intensity_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_revenue_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_nr_rx = source.rx, max_nr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with best net revenue ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}


			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			ReferenceUserControlScenarioRun.progressBar1.Value=1;
			/****************************************
			 **finished with max net revenue
			 ****************************************/

			/****************************************
			 **start max positive net revenue
			 ****************************************/

            
			ReferenceUserControlScenarioRun.lblMsg.Text="Finding Best Treatments: Maximum Positive Net Revenue";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Positive Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine(" ");

			if (m_ado.TableExist(this.m_TempMDBFileConn,"positive_max_revenue_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE positive_max_revenue_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE positive_max_revenue_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"max_nr_dpa DOUBLE," + 
				"rx_intensity SHORT)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL = "INSERT INTO positive_max_revenue_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.max_nr_dpa,a.rx_intensity " + 
				"FROM effective_product_yields_net_rev_costs_summary a, " + 
				"(SELECT biosum_cond_id,MAX(max_nr_dpa) AS max_revenue " + 
				"FROM  effective_product_yields_net_rev_costs_summary WHERE max_nr_dpa > 0 " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  max_nr_dpa, MIN(rx_intensity) AS min_intensity  " + 
				"FROM effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id,max_nr_dpa) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_revenue=a.max_nr_dpa";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Positive Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with positive maximum net revenue at minimal intensity to max_revenue_min_intensity_work_table");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}



			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN positive_max_revenue_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_pnr_rx = source.rx, max_pnr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with best net revenue ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}

			ReferenceUserControlScenarioRun.progressBar1.Value=2;
			/****************************************
			 **finished with max positive net revenue
			 ****************************************/

			//TORCH INDEX
			this.m_txtStreamWriter.WriteLine("\r\n--TORCH INDEX--");
			/****************************************
			 **start maximum torch index improvement
			 ****************************************/
			
			ReferenceUserControlScenarioRun.lblMsg.Text="Finding Best Treatments: Maximum Torch Index Improvement";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			


			
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Torch Index Improvement Processing");
			this.m_txtStreamWriter.WriteLine(" ");

			if (m_ado.TableExist(this.m_TempMDBFileConn,"effective_tici_change_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE effective_tici_change_work_table");
			}
			m_ado.m_strSQL="SELECT e.biosum_cond_id, e.rx, i.rx_intensity,e.ti_change,e.ci_change,p.max_nr_dpa " + 
				"INTO effective_tici_change_work_table  " + 
				"FROM effective e, scenario_rx_intensity i,effective_product_yields_net_rev_costs_summary p " + 
				"WHERE e.biosum_cond_id=p.biosum_cond_id AND " +
				"e.rx=p.rx AND " + 
				"i.rx=p.rx AND " + 
				"TRIM(UCASE(i.scenario_id)) ='" + ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
			this.m_txtStreamWriter.WriteLine("write effective torch index treatments to effective_tici_change_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			//get max torch index improvement with the least intensive treatment
			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_ti_change_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_ti_change_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE max_ti_change_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"ti_change DOUBLE," + 
				"rx_intensity SHORT)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
			m_ado.m_strSQL="INSERT INTO max_ti_change_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.ti_change,a.rx_intensity " + 
				"FROM effective_tici_change_work_table a," + 
				"(SELECT biosum_cond_id,MAX(ti_change) AS max_ti_change  " + 
				"FROM  effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  ti_change, MIN(rx_intensity) AS min_intensity " + 
				"FROM effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id,ti_change) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id  AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_ti_change=a.ti_change";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Torch Index Improvement ");
			this.m_txtStreamWriter.WriteLine("get treatment with the greatest torch index improvement at minimal intensity to max_ti_change_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}


			/*******************************************************
			 **update the best_rx_summary from the 
			 **max_ti_change_min_intensity_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_ti_change_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ti_imp_rx = source.rx, max_ti_imp_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with most torch index improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			/*******************************************************
			 **update the best_rx_summary_air_dest from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary_air_dest target " + 
				"INNER JOIN max_ti_change_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ti_imp_rx = source.rx, max_ti_imp_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary_air_dest with most torch index improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}

			ReferenceUserControlScenarioRun.progressBar1.Value=3;
			/***********************************************
			 **finished with maximum torch index improvement
			 ***********************************************/


			/*******************************************************************
			 **start maximum torch index improvement with positive net revenue
			 *******************************************************************/
			
			ReferenceUserControlScenarioRun.lblMsg.Text="Finding Best Treatments: Maximum Torch Index Improvement With Positive Net Revenue";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			


			/**********************************************************
			 **get all records that have positive net revenue
			 **and torch index improvement
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Torch Index Improvement With Positive Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine(" ");

			//get max torch index improvement with positive net revenue and the least intensive treatment
			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_ti_change_pnr_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_ti_change_pnr_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE max_ti_change_pnr_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"ti_change DOUBLE," + 
				"rx_intensity SHORT," + 
				"max_nr_dpa DOUBLE)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL="INSERT INTO max_ti_change_pnr_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.ti_change,a.rx_intensity,a.max_nr_dpa " + 
				"FROM effective_tici_change_work_table a," + 
				"(SELECT biosum_cond_id,MAX(ti_change) AS max_ti_change  " + 
				"FROM  effective_tici_change_work_table WHERE max_nr_dpa > 0 "  + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  ti_change, MIN(rx_intensity) AS min_intensity " + 
				"FROM effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id,ti_change) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id  AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_ti_change=a.ti_change";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Torch Index Improvement And Positive Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with the greatest torch index improvement and minimal intensity that has a positive net revenue to max_ti_change_pnr_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}

			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_ti_change_pnr_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ti_imp_pnr_rx = source.rx, max_ti_imp_pnr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with most torch index improvement that has positive net revenue");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}


			ReferenceUserControlScenarioRun.progressBar1.Value=4;
			/*************************************************************************
			 **finished with maximum torch index improvement with positive net revenue
			 *************************************************************************/

			//CROWN INDEX
			this.m_txtStreamWriter.WriteLine("\r\n--CROWN INDEX--");
			/****************************************
			 **start maximum crown index improvement
			 ****************************************/
			ReferenceUserControlScenarioRun.lblMsg.Text="Finding Best Treatments: Maximum Crown Index Improvement";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			

			/**********************************************************
			 **could have 2 different treatments that produce
			 **the same net revenue for a plot so insert into duplicates
			 **work table
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Crown Index Improvement Processing");
			this.m_txtStreamWriter.WriteLine(" ");

			//get max crown index improvement with the least intensive treatment
			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_ci_change_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_ci_change_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE max_ci_change_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"ci_change DOUBLE," + 
				"rx_intensity SHORT)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL="INSERT INTO max_ci_change_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.ci_change,a.rx_intensity " + 
				"FROM effective_tici_change_work_table a," + 
				"(SELECT biosum_cond_id,MAX(ci_change) AS max_ci_change  " + 
				"FROM  effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  ci_change, MIN(rx_intensity) AS min_intensity " + 
				"FROM effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id,ci_change) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id  AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_ci_change=a.ci_change";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Crown Index Improvement ");
			this.m_txtStreamWriter.WriteLine("get treatment with the greatest crown index improvement at minimal intensity to max_ci_change_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}

			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_ci_change_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ci_imp_rx = source.rx, max_ci_imp_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with most crown index improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			/*******************************************************
			 **update the best_rx_summary_air_dest from the 
			 **max_ci_change_min_intensity_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary_air_dest target " + 
				"INNER JOIN max_ci_change_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ci_imp_rx = source.rx, max_ci_imp_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary_air_dest with most crown index improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			ReferenceUserControlScenarioRun.progressBar1.Value=5;
			/***********************************************
			 **finished with maximum crown index improvement
			 ***********************************************/
			/*******************************************************************
			 **start maximum crown index improvement with positive net revenue
			 *******************************************************************/
			
			ReferenceUserControlScenarioRun.lblMsg.Text="Finding Best Treatments: Maximum Crown Index Improvement With Positive Net Revenue";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();


			/**********************************************************
			 **get all records that have positive net revenue
			 **and most crown index improvement
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Crown Index Improvement With Positive Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine(" ");
			//get max torch index improvement with positive net revenue and the least intensive treatment
			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_ci_change_pnr_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_ci_change_pnr_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE max_ci_change_pnr_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"ci_change DOUBLE," + 
				"rx_intensity SHORT," + 
				"max_nr_dpa DOUBLE)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL="INSERT INTO max_ci_change_pnr_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.ci_change,a.rx_intensity,a.max_nr_dpa " + 
				"FROM effective_tici_change_work_table a," + 
				"(SELECT biosum_cond_id,MAX(ci_change) AS max_ci_change  " + 
				"FROM  effective_tici_change_work_table WHERE max_nr_dpa > 0 " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  ci_change, MIN(rx_intensity) AS min_intensity " + 
				"FROM effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id,ci_change) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id  AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_ci_change=a.ci_change";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Crown Index Improvement And Positive Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with the greatest crown index improvement and minimal intensity that has a positive net revenue to max_ci_change_pnr_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}


			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_ci_change_pnr_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ci_imp_pnr_rx = source.rx, max_ci_imp_pnr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with most crown index improvement that has positive net revenue");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}

			ReferenceUserControlScenarioRun.progressBar1.Value=6;
			/*************************************************************************
			 **finished with maximum crown index improvement with positive net revenue
			 *************************************************************************/

			//MERCH
			this.m_txtStreamWriter.WriteLine("\r\n--MERCH--");
			/******************************************
			 **start minimum merchantable wood removal
			 ******************************************/
			ReferenceUserControlScenarioRun.lblMsg.Text="Finding Best Treatments: Minimum Merchantable Wood Removal";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Wood Removable Processing");
			this.m_txtStreamWriter.WriteLine(" ");
			if (m_ado.TableExist(this.m_TempMDBFileConn,"min_merch_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE min_merch_min_intensity_work_table");
			}

			m_ado.m_strSQL="CREATE TABLE min_merch_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"merch_yield_cf DOUBLE," + 
				"rx_intensity SHORT)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL = "INSERT INTO min_merch_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.merch_yield_cf,a.rx_intensity " + 
				"FROM effective_product_yields_net_rev_costs_summary a, " + 
				"(SELECT biosum_cond_id,MIN(merch_yield_cf) AS min_merch_removal " + 
				"FROM  effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id, merch_yield_cf, MIN(rx_intensity) AS min_intensity  " + 
				"FROM effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id,merch_yield_cf) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND " + 
				"b.min_merch_removal=a.merch_yield_cf AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity";
				

			this.m_txtStreamWriter.WriteLine("Insert Rx With Minimum Merch Removal ");
			this.m_txtStreamWriter.WriteLine("get treatment with minimum merch removal at minimal intensity to min_merch_min_intensity_work_table");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}


			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN min_merch_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET min_merch_rx = source.rx, min_merch_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with minimum merch removal");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			/*******************************************************
			 **update the best_rx_summary_air_dest from the 
			 **min_merch_min_intensity_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary_air_dest target " + 
				"INNER JOIN min_merch_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET min_merch_rx = source.rx, min_merch_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary_air_dest with minimum merch removal");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}
			ReferenceUserControlScenarioRun.progressBar1.Value=7;
			/**************************************************
			 **finished with minimum merchantable wood removal 
			 **************************************************/
			/*******************************************************************
			 **start minimum merchantable wood removal with positive net revenue
			 *******************************************************************/
			
			ReferenceUserControlScenarioRun.lblMsg.Text="Finding Best Treatments: Minimum Merchantable Wood Removal With Positive Net Revenue";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();

			/**********************************************************
			 **get all records that have positive net revenue
			 **and minimum merchantable wood removal
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Wood Removal With Positive Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine(" ");
			if (m_ado.TableExist(this.m_TempMDBFileConn,"min_merch_pnr_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE min_merch_pnr_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE min_merch_pnr_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"merch_yield_cf DOUBLE," + 
				"rx_intensity SHORT," + 
				"max_nr_dpa DOUBLE)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL = "INSERT INTO min_merch_pnr_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.merch_yield_cf,a.rx_intensity,a.max_nr_dpa " + 
				"FROM effective_product_yields_net_rev_costs_summary a, " + 
				"(SELECT biosum_cond_id,MIN(merch_yield_cf) AS min_merch_removal " + 
				"FROM  effective_product_yields_net_rev_costs_summary WHERE max_nr_dpa > 0 " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id, merch_yield_cf, MIN(rx_intensity) AS min_intensity  " + 
				"FROM effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id,merch_yield_cf) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.min_merch_removal=a.merch_yield_cf";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Minimum Merch Removal And Positive Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with minimum merch removal and positive net revenue at minimal intensity to max_ti_change_pnr_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}

			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN min_merch_pnr_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET min_merch_pnr_rx = source.rx, min_merch_pnr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with minimum merchantable wood removal that has positive net revenue");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
				return;
			}

			ReferenceUserControlScenarioRun.progressBar1.Value=8;
			/****************************************************************************
			 **finished with minimum merchantable wood removal with positive net revenue
			 ****************************************************************************/


			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcBestRx.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcBestRx.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcBestRx.Refresh();
			}

		}

		/// <summary>
		/// expand the wood volume,value,and costs by plot acreage for the 
		/// best treatments and data found in the best_rx_summary
		/// product_yields_net_rev_costs_summary, and effective tables
		/// </summary>
		private void BestTreatmentsByPlot()
		{
			ReferenceUserControlScenarioRun.lblMsg.Text="Best Treatment Acreage Expansion By Plot";
			ReferenceUserControlScenarioRun.lblMsg.Visible=true;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Minimum=0;
			ReferenceUserControlScenarioRun.progressBar1.Maximum = 8;
			ReferenceUserControlScenarioRun.progressBar1.Value=0;
			ReferenceUserControlScenarioRun.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Treatment Acreage Expansion By Plot");
			this.m_txtStreamWriter.WriteLine("-----------------------------------------");

			            
			/*************************************************
			 **best maximum net revenue by plot
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best " + this.m_strOptimizationTableName + " By Plot");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine("Insert optimization totals by plot");

			
			string strWhereExpr="";
			if (this.m_oOptimizationVariable.bUseFilter)
			{
					strWhereExpr = "max_nr_dpa " + this.m_oOptimizationVariable.strFilterOperator + " " + 
						Convert.ToString(this.m_oOptimizationVariable.dblFilterValue);
			}

			this.BestRxAcreageExpansionTableInsert(m_strOptimizationTableName + "_plots","biosum_cond_id","rx",strWhereExpr);
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.Refresh();
				return;
			}

			ReferenceUserControlScenarioRun.progressBar1.Value=1;



			/*************************************************
			 **Finished with best maximum net revenue by plot
			 *************************************************/
			/************************************************
			 **air curtain destruction
			 **no net revenue for air curtain destruction
			 **so just some up values
			 ************************************************/
			this.BestRxAcreageExpansionTableInsertForAirCurtainDestruction(m_strOptimizationTableName + "_plots_air_dest","biosum_cond_id","rx","");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.Refresh();
				return;
			}




		
			ReferenceUserControlScenarioRun.progressBar1.Value=2;
		


			ReferenceUserControlScenarioRun.progressBar1.Value=3;
		
		

			ReferenceUserControlScenarioRun.progressBar1.Value=4;
			

			ReferenceUserControlScenarioRun.progressBar1.Value=5;

			ReferenceUserControlScenarioRun.progressBar1.Value=6;

			ReferenceUserControlScenarioRun.progressBar1.Value=7;
			ReferenceUserControlScenarioRun.progressBar1.Value=8;
			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcBestRxPlot.Refresh();
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
		private void DeleteScenarioResultRecordsOld()
		{

			string strMDBPathAndFile = ReferenceUserControlScenarioRun.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
			ado_data_access oAdo = new ado_data_access();
			string strConn=oAdo.getMDBConnString(strMDBPathAndFile,"admin","");
			oAdo.OpenConnection(strConn);

			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			this.m_strSQL = "delete from max_nr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_pnr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_pnr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_pnr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_pnr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);

			/*************************************************
			 **delete all records in the by ownership tables
			 *************************************************/
			this.m_strSQL = "delete from max_nr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_pnr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_pnr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_pnr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_pnr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);

			/******************************************************
			 **delete all records in the by processing site tables
			 ******************************************************/
			this.m_strSQL = "delete from max_nr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_pnr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_pnr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_pnr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_pnr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);

			oAdo.m_OleDbConnection.Close();
			while (oAdo.m_OleDbConnection.State==System.Data.ConnectionState.Open)
				System.Threading.Thread.Sleep(1000);

			oAdo.m_OleDbConnection.Dispose();
			oAdo=null;



		}
		private void BestRxAcreageExpansionTableInsert(string strTable,string strTypeField, string strRxField, string strWhereExpression)
		{
			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();

			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(" + strTypeField + ", owngrpcd,acres,optimization_value,merch_haul_cost_psite," + 
				"merch_haul_cents,merch_vol,merch_dollars_val," + 
				"chip_haul_cost_psite,chip_haul_cents," +
				"chip_yield,chip_dollars_val,net_rev,harv_costs," + 
				"haul_costs) "; //,ti_chg_acres,ci_chg_acres) " ;


			this.m_strSQL += " SELECT  " + 
				"best_rx_summary.biosum_cond_id," + 
				"best_rx_summary.owngrpcd," + 
				"best_rx_summary.acres," +
				"best_rx_summary.optimization_value," + 
				p + ".merch_haul_cost_psite," + 
				p + ".merch_haul_cpa_pt * best_rx_summary.acres AS merch_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.merch_yield_cf * best_rx_summary.acres AS merch_vol," + 
				"effective_product_yields_net_rev_costs_summary.merch_val_dpa * best_rx_summary.acres AS merch_dollars_val," + 
				p + ".chip_haul_cost_psite," + 
				p + ".chip_haul_cpa_pt * best_rx_summary.acres AS chip_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.chip_yield_gt * best_rx_summary.acres AS chip_yield," + 
				"effective_product_yields_net_rev_costs_summary.chip_val_dpa * best_rx_summary.acres AS chip_dollars_val," + 
				"effective_product_yields_net_rev_costs_summary.max_nr_dpa * best_rx_summary.acres AS net_rev," + 
				"effective_product_yields_net_rev_costs_summary.harvest_onsite_cpa * best_rx_summary.acres AS harv_costs," + 
				"(effective_product_yields_net_rev_costs_summary.haul_merch_cpa + effective_product_yields_net_rev_costs_summary.haul_chip_cpa) * best_rx_summary.acres AS haul_costs " + 
//				"effective.ti_change * best_rx_summary.acres  AS ti_chg_acres," +
//				"effective.ci_change * best_rx_summary.acres  AS ci_chg_acres " + 
				"FROM ((best_rx_summary  " + 
				"INNER JOIN effective_product_yields_net_rev_costs_summary  " + 
				"ON (best_rx_summary.biosum_cond_id = effective_product_yields_net_rev_costs_summary.biosum_cond_id) AND " + 
				"(best_rx_summary." + strRxField + " = effective_product_yields_net_rev_costs_summary.rx)) " + 
				"INNER JOIN (" + p + 
				" INNER JOIN " + c + 
				" ON " + p + ".biosum_plot_id = " + c + ".biosum_plot_id) " + 
				"ON " + c + ".biosum_cond_id = best_rx_summary.biosum_cond_id)" ;
				
			if (strWhereExpression.Trim().Length > 0)
			{
				this.m_strSQL += " WHERE " + strWhereExpression + ";";
			}
			else
			{
				this.m_strSQL += ";";
			}

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		private void BestRxAcreageExpansionTableInsertOld(string strTable,string strTypeField, string strRxField, string strWhereExpression)
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
				"best_rx_summary.biosum_cond_id," + 
				"best_rx_summary.owngrpcd," + 
				"best_rx_summary.acres," +
				p + ".merch_haul_cost_psite," + 
				p + ".merch_haul_cpa_pt * best_rx_summary.acres AS merch_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.merch_yield_cf * best_rx_summary.acres AS merch_vol," + 
				"effective_product_yields_net_rev_costs_summary.merch_val_dpa * best_rx_summary.acres AS merch_dollars_val," + 
				p + ".chip_haul_cost_psite," + 
				p + ".chip_haul_cpa_pt * best_rx_summary.acres AS chip_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.chip_yield_gt * best_rx_summary.acres AS chip_yield," + 
				"effective_product_yields_net_rev_costs_summary.chip_val_dpa * best_rx_summary.acres AS chip_dollars_val," + 
				"effective_product_yields_net_rev_costs_summary.max_nr_dpa * best_rx_summary.acres AS net_rev," + 
				"effective_product_yields_net_rev_costs_summary.harvest_onsite_cpa * best_rx_summary.acres AS harv_costs," + 
				"(effective_product_yields_net_rev_costs_summary.haul_merch_cpa + effective_product_yields_net_rev_costs_summary.haul_chip_cpa) * best_rx_summary.acres AS haul_costs," + 
				"effective.ti_change * best_rx_summary.acres  AS ti_chg_acres," +
				"effective.ci_change * best_rx_summary.acres  AS ci_chg_acres " + 
				"FROM (((best_rx_summary  " + 
				"INNER JOIN effective_product_yields_net_rev_costs_summary  " + 
				"ON (best_rx_summary.biosum_cond_id = effective_product_yields_net_rev_costs_summary.biosum_cond_id) AND " + 
				"(best_rx_summary." + strRxField + " = effective_product_yields_net_rev_costs_summary.rx)) " + 
				"INNER JOIN effective " + 
				"ON (best_rx_summary.biosum_cond_id = effective.biosum_cond_id) AND " +
				"(best_rx_summary." + strRxField + "= effective.rx)) " + 
				"INNER JOIN (" + p + 
				" INNER JOIN " + c + 
				" ON " + p + ".biosum_plot_id = " + c + ".biosum_plot_id) " + 
				"ON " + c + ".biosum_cond_id = best_rx_summary.biosum_cond_id)" ;
				
			if (strWhereExpression.Trim().Length > 0)
			{
				this.m_strSQL += " WHERE " + strWhereExpression + ";";
			}
			else
			{
				this.m_strSQL += ";";
			}

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		private void BestRxAcreageExpansionTableInsertForAirCurtainDestruction(string strTable,string strTypeField, string strRxField, string strWhereExpression)
		{
			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();

			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(" + strTypeField + ", owngrpcd,acres,merch_haul_cost_psite," + 
				"merch_haul_cents,merch_vol,merch_dollars_val," + 
				"chip_haul_cost_psite,chip_haul_cents," +
				"chip_yield,chip_dollars_val,net_rev,harv_costs," + 
				"haul_costs) ";  //,ti_chg_acres,ci_chg_acres) " ;
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
				"null AS haul_costs " + //," + 
				//"effective.ti_change * best_rx_summary_air_dest.acres  AS ti_chg_acres," +
				//"effective.ci_change * best_rx_summary_air_dest.acres  AS ci_chg_acres " + 
				"FROM ((best_rx_summary_air_dest  " + 
				"INNER JOIN effective_product_yields_net_rev_costs_summary  " + 
				"ON (best_rx_summary_air_dest.biosum_cond_id = effective_product_yields_net_rev_costs_summary.biosum_cond_id) AND " + 
				"(best_rx_summary_air_dest." + strRxField + " = effective_product_yields_net_rev_costs_summary.rx)) " + 
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

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
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

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		/// <summary>
		/// sum wood volume, values, and costs by processing sites
		/// </summary>
		private void SumPSite()
		{
			//int x=0;
			ReferenceUserControlScenarioRun.lblMsg.Text= this.m_strOptimizationTableName + " by Processing Site";
			ReferenceUserControlScenarioRun.lblMsg.Visible=true;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Minimum=0;
			ReferenceUserControlScenarioRun.progressBar1.Maximum = 3;
			ReferenceUserControlScenarioRun.progressBar1.Value=0;
			ReferenceUserControlScenarioRun.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Processing Site Summary");
			this.m_txtStreamWriter.WriteLine("-----------------------");

			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			
			this.m_strSQL = "delete from psite_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			


			            
			//MAX_NR_SUM_PSITE
			/*************************************************
			 **net revenue by processing site
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Net Revenue By Psite");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process Chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			string strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
				this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_plots","chip_haul_cost_psite",
					"WHERE a.psite_id=b.cheapest_psite AND " + 
					"a.psite_id=c.chip_haul_cost_psite AND " + 
					"a.biocd=3 AND " + 
					"b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
			}
			else
			{
				this.SumPSiteWorkTableInsert(this.m_strOptimizationTableName + "_plots","chip_haul_cost_psite",
					"WHERE a.psite_id=b.cheapest_psite AND " + 
					"a.psite_id=c.chip_haul_cost_psite AND " + 
					"a.biocd=3");

			}
				    
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_nr_sum_psite
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to max_nr_sum_psite");
			this.SumPSiteTableInsert(this.m_strOptimizationTableName + "_psites");
			ReferenceUserControlScenarioRun.progressBar1.Value=1;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************
			 **Finished net revenue by processsing site
			 *************************************************/

			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcBestRxPSite.Refresh();
			}
		}

		private void SumPSiteWorkTableInsert(string strTable,string strPSiteField, string strWhereExpression)
		{

			this.m_strSQL = "INSERT INTO psite_sum_work_table " + 
				"(psite_id,biocd,sum_optimization_value,sum_acres,sum_chip_yield," + 
				"sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," + 
				"sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs," + 
				"sum_haul_costs) " + //,sum_ti_chg_acres,sum_ci_chg_acres) " + 
				"SELECT DISTINCT a.psite_id," + 
				"a.biocd," +  
				"c.sum_acres," + 
				"c.sum_optimization_value," + 
				"c.sum_chip_yield," + 
				"c.sum_chip_haul_cents," +
				"c.sum_chip_dollars_val, " + 
				"c.sum_merch_haul_cents," + 
				"c.sum_merch_vol," + 
				"c.sum_net_rev," + 
				"c.sum_merch_dollars_val," +
				"c.sum_harv_costs," + 
				"c.sum_haul_costs " + 
			
				"FROM " + this.m_strPSiteTable.Trim() + " AS a," + 
				"(SELECT biosum_cond_id," + 
				"MIN(" + strPSiteField.Trim() + ") as cheapest_psite " + 
				"FROM " + strTable + " GROUP BY biosum_cond_id)  b," + 
				"(SELECT " + strPSiteField.Trim() + "," + 
				"SUM(optimization_value) AS sum_optimization_value," + 
				"SUM(chip_yield)  AS sum_chip_yield," + 
				"SUM(chip_haul_cents) AS sum_chip_haul_cents," +
				"SUM(chip_dollars_val) AS sum_chip_dollars_val, " +
				"SUM(acres) AS sum_acres, " +
				"SUM(merch_haul_cents) AS sum_merch_haul_cents," +
				"SUM(merch_vol) AS sum_merch_vol," + 
				"SUM(net_rev) AS sum_net_rev," + 
				"SUM(merch_dollars_val) AS sum_merch_dollars_val," + 
				"SUM(harv_costs) AS sum_harv_costs," + 
				"SUM(haul_costs) AS sum_haul_costs " + 
				"FROM " + strTable + " GROUP BY " + strPSiteField.Trim() + ") c " +
				strWhereExpression.Trim() + ";";

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		private void SumPSiteTableInsert(string strTable)
		{
			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(psite_id,acres,optimization_value,merch_haul_cents,chip_haul_cents," + 
				"merch_vol,chip_yield,net_rev,merch_dollars_val," + 
				"chip_dollars_val,harv_costs,haul_costs) " + 
				
				"SELECT psite_id,sum_acres AS acres," + 
				"sum_optimization_value AS optimization_value," + 
				"sum_merch_haul_cents AS merch_haul_cents," + 
				"sum_chip_haul_cents AS chip_haul_cents," + 
				"sum_merch_vol AS merch_vol," + 
				"sum_chip_yield AS chip_yield," + 
				"sum_net_rev AS net_rev," + 
				"sum_merch_dollars_val AS merch_dollars_val," + 
				"sum_chip_dollars_val AS chip_dollars_val," + 
				"sum_harv_costs AS harv_costs," + 
				"sum_haul_costs AS haul_costs " + 
				
				"FROM psite_sum_work_table;";
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		/// <summary>
		/// sum wood volumes, values, and costs by land ownership
		/// </summary>
		private void SumOwnership()
		{
			
			ReferenceUserControlScenarioRun.lblMsg.Text= this.m_strOptimizationTableName + " by Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Visible=true;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Minimum=0;
			ReferenceUserControlScenarioRun.progressBar1.Maximum = 8;
			ReferenceUserControlScenarioRun.progressBar1.Value=0;
			ReferenceUserControlScenarioRun.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Ownership Summary");
			this.m_txtStreamWriter.WriteLine("-----------------------");

			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			


			
            
			//MAX_NR_SUM_OWN
			/*************************************************
			 **net revenue by ownership
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(this.m_strOptimizationTableName + " Plots by Ownership");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table");
			this.SumOwnershipWorkTableInsert(this.m_strOptimizationTableName + "_plots","own_sum_work_table");
			ReferenceUserControlScenarioRun.progressBar1.Value=1;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_nr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to " + this.m_strOptimizationTableName + "_own");
			this.SumOwnershipTableInsert("own_sum_work_table",this.m_strOptimizationTableName + "_own");
			ReferenceUserControlScenarioRun.progressBar1.Value=2;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*************************************************
			 **Finished net revenue by ownership
			 *************************************************/

			ReferenceUserControlScenarioRun.progressBar1.Value = 4;
			
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_txtStreamWriter.WriteLine(this.m_strOptimizationTableName + " Air Curtain Destruction by Ownership");
			this.m_txtStreamWriter.WriteLine("Insert into work table");
			this.SumOwnershipWorkTableInsert(this.m_strOptimizationTableName + "_plots_air_dest","own_sum_work_table");

			ReferenceUserControlScenarioRun.progressBar1.Value=5;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_nr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to " + this.m_strOptimizationTableName + "_own_air_dest");
			this.SumOwnershipTableInsert("own_sum_work_table",this.m_strOptimizationTableName + "_own_air_dest");

			ReferenceUserControlScenarioRun.progressBar1.Value=6;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			ReferenceUserControlScenarioRun.progressBar1.Value=8;

			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
			}
			

		}

		private void SumOwnershipOld()
		{
			//int x=0;
			ReferenceUserControlScenarioRun.lblMsg.Text="Maximum Net Revenue By Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Visible=true;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Minimum=0;
			ReferenceUserControlScenarioRun.progressBar1.Maximum = 8;
			ReferenceUserControlScenarioRun.progressBar1.Value=0;
			ReferenceUserControlScenarioRun.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Ownership Summary");
			this.m_txtStreamWriter.WriteLine("-----------------------");

			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			


			
            
			//MAX_NR_SUM_OWN
			/*************************************************
			 **net revenue by ownership
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Net Revenue By Onwership");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table");
			this.SumOwnershipWorkTableInsert("max_nr_plots","own_sum_work_table");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_nr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_nr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_nr_sum_own");
			ReferenceUserControlScenarioRun.progressBar1.Value=1;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*************************************************
			 **Finished net revenue by ownership
			 *************************************************/

			//MAX_PNR_SUM_OWNERSHIP
			/*************************************************
			 **Positive Net Revenue By Ownership
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			ReferenceUserControlScenarioRun.lblMsg.Text="Maximum Positive Net Revenue By Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Positive Net Revenue By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table");
			this.SumOwnershipWorkTableInsert("max_pnr_plots","own_sum_work_table");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_pnr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from ownership values from work table to max_pnr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_pnr_sum_own");
			ReferenceUserControlScenarioRun.progressBar1.Value=2;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*************************************************
			 **Finished positive net revenue by ownership
			 *************************************************/

			//MAX_TI_IMP_SUM_OWN
			/*************************************************
			 **most ti improvement by ownership
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			ReferenceUserControlScenarioRun.lblMsg.Text="Maximum Torch Index Improvement By Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Torch Index Improvement By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ti_imp_plots","own_sum_work_table");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ti_imp_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_ti_imp_sum_own");
			
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			
			/********************************************************************
			 **most ti improvement by ownership for air curtain destruction plots
			 ********************************************************************/
			this.m_strSQL = "delete from own_sum_work_table_air_dest";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			ReferenceUserControlScenarioRun.lblMsg.Text="Maximum Torch Index Improvement By Ownership For Air Curtain Destruction Plots";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Torch Index Improvement By Ownership For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ti_imp_plots_air_dest","own_sum_work_table_air_dest");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_sum_own_air_dest
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ti_imp_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table_air_dest","max_ti_imp_sum_own_air_dest");
			
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			ReferenceUserControlScenarioRun.progressBar1.Value=3;
			/*************************************************
			 **Finished torch index improvement ownership
			 *************************************************/

			//MAX_TI_IMP_PNR_SUM_OWN
			/*************************************************
			 **Torch Index Positive Net Revenue By Ownership
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			ReferenceUserControlScenarioRun.lblMsg.Text="Maximum Torch Index Improvement And Positive Net Revenue By Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Torch Index And Positive Net Revenue By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ti_imp_pnr_plots","own_sum_work_table");
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_pnr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ti_imp_pnr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_ti_imp_pnr_sum_own");
			ReferenceUserControlScenarioRun.progressBar1.Value=4;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/**************************************************************
			 **Finished torch index postive net revenue by ownership
			 ***************************************************************/

			//MAX_CI_IMP_SUM_OWN
			/*************************************************
			 **most ci improvement by processing site
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			ReferenceUserControlScenarioRun.lblMsg.Text="Maximum Crown Index Improvement By Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Crown Index Improvement By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ci_imp_plots","own_sum_work_table");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_ci_imp_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ci_imp_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_ci_imp_sum_own");

			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			/********************************************************************
			 **most ci improvement by ownership for air curtain destruction plots
			 ********************************************************************/
			this.m_strSQL = "delete from own_sum_work_table_air_dest";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			ReferenceUserControlScenarioRun.lblMsg.Text="Maximum Crown Index Improvement By Ownership For Air Curtain Destruction Plots";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Crown Index Improvement By Ownership For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ci_imp_plots_air_dest","own_sum_work_table_air_dest");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_sum_own_air_dest
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ci_imp_sum_own_air_dest");
			this.SumOwnershipTableInsert("own_sum_work_table_air_dest","max_ci_imp_sum_own_air_dest");
			
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			ReferenceUserControlScenarioRun.progressBar1.Value=5;
			/*************************************************
			 **Finished crown index improvement by ownership
			 *************************************************/

			//MAX_CI_IMP_PNR_SUM_OWN
			/*************************************************************
			 **Crown Index Improvement Positive Net Revenue By Ownership
			 *************************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			ReferenceUserControlScenarioRun.lblMsg.Text="Maximum Crown Index Improvement And Positive Net Revenue By Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Crown Index And Positive Net Revenue By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ci_imp_pnr_plots","own_sum_work_table");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_ci_imp_pnr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ti_imp_pnr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_ci_imp_pnr_sum_own");
			ReferenceUserControlScenarioRun.progressBar1.Value=6;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/**************************************************************
			 **Finished crown index postive net revenue by ownership
			 ***************************************************************/

			//MIN_MERCH_SUM_OWN
			/*************************************************
			 **Minimum Merchantable Removal By Ownership
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			ReferenceUserControlScenarioRun.lblMsg.Text="Minimum Merchantable Removal Treatments By Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Removal Treatments By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("min_merch_plots","own_sum_work_table");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into min_merch_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to min_merch_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","min_merch_sum_own");
			ReferenceUserControlScenarioRun.progressBar1.Value=7;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			/********************************************************************************
			 **Minimum Merchantable Removal By Ownership For Air Curtain Destruction Plots
			 ********************************************************************************/
			this.m_strSQL = "delete from own_sum_work_table_air_dest";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			ReferenceUserControlScenarioRun.lblMsg.Text="Minimum Merchantable Removal Treatments By Ownership For Air Curtain Destruction Plots";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Removal Treatments By Ownership For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("min_merch_plots_air_dest","own_sum_work_table_air_dest");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into min_merch_sum_own_air_dest
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to min_merch_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table_air_dest","min_merch_sum_own_air_dest");
			
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			ReferenceUserControlScenarioRun.progressBar1.Value=7;
			/*************************************************************
			 **Finished Minimum Merchantable Removal By  Ownership
			 *************************************************************/

			//MIN_MERCH_PNR_SUM_OWN
			/****************************************************************************
			 **Minimum Merchantable Removal And Positive Net Revenue By Ownership
			 ****************************************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			ReferenceUserControlScenarioRun.lblMsg.Text="Minimum Merchantable Removal And Positive Net Revenue By Ownership";
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minumum Merchantable Removal And Positive Net Revenue By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership into work table");
			this.SumOwnershipWorkTableInsert("min_merch_pnr_plots","own_sum_work_table");
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into min_merch_pnr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to min_merch_pnr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","min_merch_pnr_sum_own");
			ReferenceUserControlScenarioRun.progressBar1.Value=8;
			if (this.UserCancel(ReferenceUserControlScenarioRun.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "!!Error!!";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
				return;
			}
			/**************************************************************
			 **Finished min merch postive net revenue by Ownership
			 ***************************************************************/



			if (this.m_intError == 0)
			{
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Blue;
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Text = "Completed";
				ReferenceUserControlScenarioRun.lblProcBestRxOwner.Refresh();
			}
		}


		private void SumOwnershipWorkTableInsert(string strTableSource,string strTableDestination)
		{
			this.m_strSQL = "INSERT INTO " + strTableDestination + " "  + 
				"(owngrpcd,sum_acres,sum_optimization_value,sum_chip_yield," + 
				"sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," + 
				"sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs," + 
				"sum_haul_costs) " + 
				"SELECT owngrpcd," + 
				"SUM(acres) AS sum_acres, " +
				"SUM(optimization_value) AS sum_optimization_value," + 
				"SUM(chip_yield)  AS sum_chip_yield," + 
				"SUM(IIF(chip_haul_cents IS NOT NULL,chip_haul_cents,0)) AS sum_chip_haul_cents," +
				"SUM(IIF(chip_dollars_val IS NOT NULL,chip_dollars_val,0)) AS sum_chip_dollars_val, " +
				"SUM(IIF(merch_haul_cents IS NOT NULL,merch_haul_cents,0)) AS sum_merch_haul_cents," +
				"SUM(merch_vol) AS sum_merch_vol," + 
				"SUM(IIF(net_rev IS NOT NULL,net_rev,0)) AS sum_net_rev," + 
				"SUM(IIF(merch_dollars_val IS NOT NULL,merch_dollars_val,0)) AS sum_merch_dollars_val," + 
				"SUM(harv_costs) AS sum_harv_costs," + 
				"SUM(IIF(haul_costs IS NOT NULL,haul_costs,0)) AS sum_haul_costs " + 
				
				"FROM " + strTableSource + " GROUP BY owngrpcd ;";
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		
		private void SumOwnershipTableInsert(string strTableSource,string strTableDestination)
		{
			this.m_strSQL = "INSERT INTO " + strTableDestination + " " + 
				"(owngrpcd,acres,optimization_value,merch_haul_cents,chip_haul_cents," + 
				"merch_vol,chip_yield,net_rev,merch_dollars_val," + 
				"chip_dollars_val,harv_costs,haul_costs) " + 
				
				"SELECT owngrpcd,sum_acres AS acres," + 
				"sum_optimization_value AS optimization_value," + 
				"sum_merch_haul_cents AS merch_haul_cents," + 
				"sum_chip_haul_cents AS chip_haul_cents," + 
				"sum_merch_vol AS merch_vol," + 
				"sum_chip_yield AS chip_yield," + 
				"sum_net_rev AS net_rev," + 
				"sum_merch_dollars_val AS merch_dollars_val," + 
				"sum_chip_dollars_val AS chip_dollars_val," + 
				"sum_harv_costs AS harv_costs," + 
				"sum_haul_costs AS haul_costs " + 
			
				"FROM " + strTableSource;
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		private void SumAirCurtainDestruction()
		{
			ReferenceUserControlScenarioRun.lblMsg.Text="Air Curtain Destruction";
			ReferenceUserControlScenarioRun.lblMsg.Visible=true;
			ReferenceUserControlScenarioRun.lblMsg.Refresh();
			ReferenceUserControlScenarioRun.progressBar1.Minimum=0;
			ReferenceUserControlScenarioRun.progressBar1.Maximum = 8;
			ReferenceUserControlScenarioRun.progressBar1.Value=0;
			ReferenceUserControlScenarioRun.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Air Curtain Destruction Summary");
			this.m_txtStreamWriter.WriteLine("-------------------------------");

			this.m_strSQL = "delete from air_curtain_destruction_plots";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();

			string strWhere = "WHERE " + p + ".merch_haul_cost_psite IS NULL OR " + 
				p + ".chip_haul_cost_psite IS NULL";


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


		/// <summary>
		/// check and see if the user pressed the cancel button
		/// </summary>
		/// <param name="p_oLabel"></param>
		/// <returns></returns>
		private bool UserCancel(System.Windows.Forms.Label p_oLabel)
		{
			System.Windows.Forms.Application.DoEvents();
			if (ReferenceUserControlScenarioRun.m_bUserCancel == true)
			{
				p_oLabel.ForeColor = System.Drawing.Color.Red;
				p_oLabel.Text = "Cancelled";
				return true;
			}
			return false;

		}
		public FIA_Biosum_Manager.frmScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		
		public FIA_Biosum_Manager.uc_scenario_run ReferenceUserControlScenarioRun
		{
			get {return _uc_scenario_run;}
			set {_uc_scenario_run=value;}
		}
		
		/*
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables ReferenceUserControlFVSPrePostVariables
		{
			get {return _uc_scenario_fvs_prepost_variables;}
			set {_uc_scenario_fvs_prepost_variables=value;}
		}
		*/
	}
}
