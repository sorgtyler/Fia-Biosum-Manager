using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Threading;



namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_fvs_input.
	/// </summary>
	public class uc_fvs_input : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		public System.Windows.Forms.ListView lstFvsInput;
		private System.Windows.Forms.Label lblRxCnt;
		private System.Windows.Forms.Label lblVarCnt;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnChkAll;
		private System.Windows.Forms.Button btnClearAll;
        private System.Windows.Forms.Button btnRefresh;
		private string m_strProjDir="";
		private string m_strProjId="";
		//private string m_strDsnIn="";
		private string m_strLocFile="";
		private string m_strSlfFile="";
		//private string m_strFvsFile="";
		private string m_strDsnOut="";
		//private string m_strInDir="";
       // private string m_strInMDBFile="";
		//private string m_strOutDir="";
		private string m_strOutMDBFile="";
		private string m_strRxTable="";
		private string m_strPlotTable="";
		private string m_strCondTable="";
		private string m_strTreeTable="";
		private string m_strTreeSpcTable="";
        private string m_strOutPotFireBaseYearMDBFile="";
		private Datasource m_DataSource;
		private ado_data_access m_ado;
		private dao_data_access m_dao;
		private string m_strConn="";
		private string m_strTempMDBFile="";
		private System.Windows.Forms.Button btnHelp;
		private int m_intError=0;
		private System.Threading.Thread m_thread;

		//list view column constants
		private const int COL_CHECKBOX = 0;
		private const int COL_VARIANT = 1;
		private const int COL_RX = 2;
		private const int COL_LOC = 3;
		private const int COL_STANDCOUNT = 4;
		private const int COL_TREECOUNT=5;
		private const int COL_MDBOUT = 6;
		private const int COL_SUMMARYCOUNT = 7;
		private const int COL_CUTCOUNT = 8;
		private const int COL_LEFTCOUNT = 9;
		private const int COL_POTFIRECOUNT = 10;
        private const int COL_POTFIREMDBOUT = 11;
        private const int COL_POTFIREBASEYEARCOUNT = 12;
		private System.Windows.Forms.ProgressBar progressBar1;
		private System.Windows.Forms.Label lblProgress;
		private System.Windows.Forms.Button btnCancel;
        private bool bAbort = false;
		public FIA_Biosum_Manager.frmTherm m_frmTherm;
		private System.Windows.Forms.Label lblTreeSpcVarCnt;
		private System.Windows.Forms.Button btnPlotVariants;
		private System.Windows.Forms.Button btnTreeSpcVariants;
        private System.Windows.Forms.Button btnRx;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors = new ListViewAlternateBackgroundColors();
		private System.Windows.Forms.Label lblRxPackageCnt;
		private Queries m_oQueries = new Queries();
		private RxTools m_oRxTools = new RxTools();
		private System.Windows.Forms.TextBox txtDataDir;
        private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnRxPackage;
		private frmMain _frmMain=null;
		private frmDialog _frmDialog=null;
        private string m_strFVSCycleLength="10";
        private Button btnExecuteAction;
        private ComboBox cmbAction;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultFvsXPSFile;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_fvs_input()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_ado = new ado_data_access();
			this.m_dao = new dao_data_access();
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor=frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.CustomFullRowSelect=true;
			this.m_oLvRowColors.ReferenceListView = this.lstFvsInput;
			if (frmMain.g_oGridViewFont != null) this.lstFvsInput.Font = frmMain.g_oGridViewFont;

            this.m_oEnv = new env();


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
            this.btnExecuteAction = new System.Windows.Forms.Button();
            this.cmbAction = new System.Windows.Forms.ComboBox();
            this.txtDataDir = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRxPackage = new System.Windows.Forms.Button();
            this.lblRxPackageCnt = new System.Windows.Forms.Label();
            this.btnRx = new System.Windows.Forms.Button();
            this.btnTreeSpcVariants = new System.Windows.Forms.Button();
            this.btnPlotVariants = new System.Windows.Forms.Button();
            this.lblTreeSpcVarCnt = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnChkAll = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblVarCnt = new System.Windows.Forms.Label();
            this.lblRxCnt = new System.Windows.Forms.Label();
            this.lstFvsInput = new System.Windows.Forms.ListView();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnExecuteAction);
            this.groupBox1.Controls.Add(this.cmbAction);
            this.groupBox1.Controls.Add(this.txtDataDir);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.btnRxPackage);
            this.groupBox1.Controls.Add(this.lblRxPackageCnt);
            this.groupBox1.Controls.Add(this.btnRx);
            this.groupBox1.Controls.Add(this.btnTreeSpcVariants);
            this.groupBox1.Controls.Add(this.btnPlotVariants);
            this.groupBox1.Controls.Add(this.lblTreeSpcVarCnt);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.progressBar1);
            this.groupBox1.Controls.Add(this.lblProgress);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnRefresh);
            this.groupBox1.Controls.Add(this.btnClearAll);
            this.groupBox1.Controls.Add(this.btnChkAll);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.lblVarCnt);
            this.groupBox1.Controls.Add(this.lblRxCnt);
            this.groupBox1.Controls.Add(this.lstFvsInput);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(784, 616);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnExecuteAction
            // 
            this.btnExecuteAction.Location = new System.Drawing.Point(680, 501);
            this.btnExecuteAction.Name = "btnExecuteAction";
            this.btnExecuteAction.Size = new System.Drawing.Size(88, 32);
            this.btnExecuteAction.TabIndex = 63;
            this.btnExecuteAction.Text = "Execute Action";
            this.btnExecuteAction.Click += new System.EventHandler(this.btnExecuteAction_Click);
            // 
            // cmbAction
            // 
            this.cmbAction.FormattingEnabled = true;
            this.cmbAction.Items.AddRange(new object[] {
            "Create FVS Input Database Files",
            "Create FVS Output Database Files",
            "Delete Standard FVS Output Tables",
            "Delete POTFIRE Base Year Output Tables",
            "Delete Both Standard and POTFIRE Base Year Output Tables",
            "Write KCP Template Scripts",
            "View KCP Template Scripts"});
            this.cmbAction.Location = new System.Drawing.Point(312, 508);
            this.cmbAction.Name = "cmbAction";
            this.cmbAction.Size = new System.Drawing.Size(362, 21);
            this.cmbAction.TabIndex = 62;
            this.cmbAction.Text = "<-------Action Items------->";
            this.cmbAction.SelectedIndexChanged += new System.EventHandler(this.cmbAction_SelectedIndexChanged);
            this.cmbAction.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbAction_KeyPress);
            // 
            // txtDataDir
            // 
            this.txtDataDir.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDataDir.Location = new System.Drawing.Point(136, 120);
            this.txtDataDir.Name = "txtDataDir";
            this.txtDataDir.Size = new System.Drawing.Size(632, 20);
            this.txtDataDir.TabIndex = 60;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(24, 120);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 24);
            this.label1.TabIndex = 59;
            this.label1.Text = "Data Directory";
            // 
            // btnRxPackage
            // 
            this.btnRxPackage.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRxPackage.Location = new System.Drawing.Point(320, 78);
            this.btnRxPackage.Name = "btnRxPackage";
            this.btnRxPackage.Size = new System.Drawing.Size(90, 24);
            this.btnRxPackage.TabIndex = 58;
            this.btnRxPackage.Text = "Packages";
            this.btnRxPackage.Click += new System.EventHandler(this.btnRxPackage_Click);
            // 
            // lblRxPackageCnt
            // 
            this.lblRxPackageCnt.BackColor = System.Drawing.Color.White;
            this.lblRxPackageCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRxPackageCnt.Location = new System.Drawing.Point(415, 80);
            this.lblRxPackageCnt.Name = "lblRxPackageCnt";
            this.lblRxPackageCnt.Size = new System.Drawing.Size(32, 16);
            this.lblRxPackageCnt.TabIndex = 57;
            this.lblRxPackageCnt.Text = "0";
            // 
            // btnRx
            // 
            this.btnRx.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRx.Location = new System.Drawing.Point(320, 54);
            this.btnRx.Name = "btnRx";
            this.btnRx.Size = new System.Drawing.Size(90, 24);
            this.btnRx.TabIndex = 55;
            this.btnRx.Text = "Treatments";
            this.btnRx.Click += new System.EventHandler(this.btnRx_Click);
            // 
            // btnTreeSpcVariants
            // 
            this.btnTreeSpcVariants.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTreeSpcVariants.Location = new System.Drawing.Point(16, 78);
            this.btnTreeSpcVariants.Name = "btnTreeSpcVariants";
            this.btnTreeSpcVariants.Size = new System.Drawing.Size(232, 24);
            this.btnTreeSpcVariants.TabIndex = 54;
            this.btnTreeSpcVariants.Text = "Tree Species Missing FVS Variant Value";
            this.btnTreeSpcVariants.Click += new System.EventHandler(this.btnTreeSpcVariants_Click);
            // 
            // btnPlotVariants
            // 
            this.btnPlotVariants.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPlotVariants.Location = new System.Drawing.Point(16, 54);
            this.btnPlotVariants.Name = "btnPlotVariants";
            this.btnPlotVariants.Size = new System.Drawing.Size(232, 24);
            this.btnPlotVariants.TabIndex = 53;
            this.btnPlotVariants.Text = "Plots Missing FVS Variant Value";
            this.btnPlotVariants.Click += new System.EventHandler(this.btnPlotVariants_Click);
            // 
            // lblTreeSpcVarCnt
            // 
            this.lblTreeSpcVarCnt.BackColor = System.Drawing.Color.White;
            this.lblTreeSpcVarCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTreeSpcVarCnt.Location = new System.Drawing.Point(253, 80);
            this.lblTreeSpcVarCnt.Name = "lblTreeSpcVarCnt";
            this.lblTreeSpcVarCnt.Size = new System.Drawing.Size(56, 16);
            this.lblTreeSpcVarCnt.TabIndex = 51;
            this.lblTreeSpcVarCnt.Text = "0";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(568, 567);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(56, 24);
            this.btnCancel.TabIndex = 48;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(312, 575);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(240, 8);
            this.progressBar1.TabIndex = 47;
            this.progressBar1.Visible = false;
            // 
            // lblProgress
            // 
            this.lblProgress.Location = new System.Drawing.Point(312, 591);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(239, 16);
            this.lblProgress.TabIndex = 46;
            this.lblProgress.Text = "lblProgress";
            this.lblProgress.Visible = false;
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(16, 575);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 43;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(142, 501);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(64, 32);
            this.btnRefresh.TabIndex = 38;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(79, 501);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(64, 32);
            this.btnClearAll.TabIndex = 37;
            this.btnClearAll.Text = "Clear All";
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // btnChkAll
            // 
            this.btnChkAll.Location = new System.Drawing.Point(16, 501);
            this.btnChkAll.Name = "btnChkAll";
            this.btnChkAll.Size = new System.Drawing.Size(64, 32);
            this.btnChkAll.TabIndex = 36;
            this.btnChkAll.Text = "Check All";
            this.btnChkAll.Click += new System.EventHandler(this.btnChkAll_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(672, 575);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 33;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblVarCnt
            // 
            this.lblVarCnt.BackColor = System.Drawing.Color.White;
            this.lblVarCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVarCnt.Location = new System.Drawing.Point(254, 58);
            this.lblVarCnt.Name = "lblVarCnt";
            this.lblVarCnt.Size = new System.Drawing.Size(56, 16);
            this.lblVarCnt.TabIndex = 29;
            this.lblVarCnt.Text = "0";
            // 
            // lblRxCnt
            // 
            this.lblRxCnt.BackColor = System.Drawing.Color.White;
            this.lblRxCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRxCnt.Location = new System.Drawing.Point(415, 58);
            this.lblRxCnt.Name = "lblRxCnt";
            this.lblRxCnt.Size = new System.Drawing.Size(32, 16);
            this.lblRxCnt.TabIndex = 27;
            this.lblRxCnt.Text = "0";
            // 
            // lstFvsInput
            // 
            this.lstFvsInput.CheckBoxes = true;
            this.lstFvsInput.GridLines = true;
            this.lstFvsInput.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFvsInput.HideSelection = false;
            this.lstFvsInput.Location = new System.Drawing.Point(16, 152);
            this.lstFvsInput.MultiSelect = false;
            this.lstFvsInput.Name = "lstFvsInput";
            this.lstFvsInput.Size = new System.Drawing.Size(752, 344);
            this.lstFvsInput.TabIndex = 26;
            this.lstFvsInput.UseCompatibleStateImageBehavior = false;
            this.lstFvsInput.View = System.Windows.Forms.View.Details;
            this.lstFvsInput.SelectedIndexChanged += new System.EventHandler(this.lstFvsInput_SelectedIndexChanged);
            this.lstFvsInput.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFvsInput_MouseUp);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(778, 32);
            this.lblTitle.TabIndex = 25;
            this.lblTitle.Text = "Create FVS Input";
            // 
            // uc_fvs_input
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_input";
            this.Size = new System.Drawing.Size(784, 616);
            this.Resize += new System.EventHandler(this.uc_fvs_input_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_fvs_input_Resize(object sender, System.EventArgs e)
		{
			this.Resize_Fvs_Input();
		}
		public void Resize_Fvs_Input()
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.btnHelp.Top = this.btnClose.Top;
				this.lstFvsInput.Left = 5;
				this.lstFvsInput.Width = this.Width - 10;

				this.lstFvsInput.Height = this.btnClose.Top - this.lstFvsInput.Top - (this.btnExecuteAction.Height * 3);
				this.btnExecuteAction.Top = this.lstFvsInput.Top + this.lstFvsInput.Height + 2;
                this.btnExecuteAction.Left = this.lstFvsInput.Width - btnExecuteAction.Width - (lstFvsInput.Left * 2);
                this.cmbAction.Top = this.btnExecuteAction.Top + (int)(btnExecuteAction.Height * .5) - (int)(cmbAction.Height * .5);
                this.cmbAction.Left = this.btnExecuteAction.Left - this.cmbAction.Width - 5;
				
				this.btnChkAll.Top = this.btnExecuteAction.Top;
				this.btnClearAll.Top = this.btnExecuteAction.Top;
				this.btnRefresh.Top = this.btnExecuteAction.Top;

				this.txtDataDir.Width = this.Width - this.txtDataDir.Left - 10;

				this.progressBar1.Left = (int)(this.groupBox1.Width * .50) - (int)(this.progressBar1.Width * .50);
				this.progressBar1.Top = this.btnClose.Top;
				
				
				if (this.progressBar1.Left < (this.btnHelp.Left + this.btnHelp.Width))
				{
					this.progressBar1.Left = this.btnHelp.Left + this.btnHelp.Width;
				}
				this.btnCancel.Left = this.progressBar1.Left + this.progressBar1.Width + 2;
				this.btnCancel.Top = this.progressBar1.Top - (int)(this.btnCancel.Height * .50) + (int)(this.progressBar1.Height * .50);
				if (this.btnClose.Left < (this.btnCancel.Left + this.btnCancel.Width))
				{
					this.btnClose.Left = this.btnCancel.Left + this.btnCancel.Width;
				}
				this.lblProgress.Left = this.progressBar1.Left;
				this.lblProgress.Top = this.progressBar1.Top + this.progressBar1.Height + 2;


			}
			catch
			{
			}
		}
		public void loadvalues()
		{
			
			this.LoadDataSources();
			this.populate_listbox();

		}
		private void populate_listbox()
		{
			//bool bResult;
			string strInDirAndFile;
			string strOutDirAndFile;
			string strConn;
			string[] strValues;
			bool bFoundDsnOut;
            string strVariant = "";
            string strCurrentVariant = "";
            string strRecordCount = "";
			//bool bFoundDsnIn;
			try
			{

				m_ado.OpenConnection(m_strConn);

				this.lblRxCnt.Text = Convert.ToString((int)this.m_ado.getRecordCount(m_ado.m_OleDbConnection,"select count(*) from " + m_oQueries.m_oFvs.m_strRxTable + " where rx IS NOT NULL AND LEN(TRIM(rx)) > 0;",m_oQueries.m_oFvs.m_strRxTable));
				this.lblRxPackageCnt.Text  = Convert.ToString((int)this.m_ado.getRecordCount(m_ado.m_OleDbConnection,"select count(*) from " + m_oQueries.m_oFvs.m_strRxPackageTable + " WHERE rxpackage IS NOT NULL AND LEN(TRIM(rxpackage)) > 0;",m_oQueries.m_oFvs.m_strRxTable));
				this.lblVarCnt.Text = Convert.ToString((int)this.m_ado.getRecordCount(m_ado.m_OleDbConnection,"select count(*) from " + this.m_oQueries.m_oFIAPlot.m_strPlotTable + " where fvs_variant IS NULL OR LEN(TRIM(fvs_variant)) = 0;",m_oQueries.m_oFIAPlot.m_strPlotTable));
				this.lblTreeSpcVarCnt.Text = 
					Convert.ToString((int)this.m_ado.getRecordCount(m_ado.m_OleDbConnection,"select count(*) " + 
					                                                               "from (select fvs_variant " + 
					                                                                     "from " + m_oQueries.m_oFIAPlot.m_strPlotTable + " p " + 
																					     "where not exists (select fvs_variant " + 
																							               "from " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " t " + 
																							               "where trim(p.fvs_variant)=trim(t.fvs_variant)))","missingvariantcount"));

				this.lstFvsInput.Clear();
				this.m_oLvRowColors.InitializeRowCollection();


				this.lstFvsInput.Columns.Add("", 2, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("Variant", 55, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("Package", 55, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("Location File", 80, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("Stands", 70, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("Trees", 70, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("FVS Output DB", 230, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("SUMMARY RECS", 100, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("CUTLIST RECS", 100, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("TREELIST RECS", 100, HorizontalAlignment.Left);
				this.lstFvsInput.Columns.Add("POTFIRE RECS", 100, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("POTFIRE BaseYr DB", 120, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("POTFIRE BaseYr RECS", 120, HorizontalAlignment.Left);

				this.lstFvsInput.Columns[COL_CHECKBOX].Width = -2;

				this.m_ado.m_strSQL = Queries.FVS.GetFVSVariantRxPackageSQL(m_oQueries.m_oFIAPlot.m_strPlotTable,m_oQueries.m_oFvs.m_strRxPackageTable);
				this.m_ado.SqlQueryReader(m_ado.m_OleDbConnection,this.m_ado.m_strSQL);


				//declare a registry key object
				Microsoft.Win32.RegistryKey regKey; // new Microsoft.Win32 Registry Key 
				Microsoft.Win32.RegistryKey regKey2;
				// open the subkey that holds the current odbc data sources
				regKey = Registry.CurrentUser.OpenSubKey( @"Software\ODBC\ODBC.Ini\Odbc data sources",false); 
				// get string name in string array
				// then use getValue to get data value.
				//save the odbc datasources in strDsnNames
				string[] strDsnNames = regKey.GetValueNames(); // for the 1st time it's only 2 names

			
				string strRegKey = "";

			


				while (this.m_ado.m_OleDbDataReader.Read())
				{
					strRegKey = "";
					bFoundDsnOut = false;
					
					// Add a ListItem object to the ListView.

					//add new row
					System.Windows.Forms.ListViewItem entryListItem =
						this.lstFvsInput.Items.Add("");
					entryListItem.UseItemStyleForSubItems=false;
					this.m_oLvRowColors.AddRow();
					this.m_oLvRowColors.AddColumns(lstFvsInput.Items.Count-1,lstFvsInput.Columns.Count);
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_CHECKBOX,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);

					//fvs_variant		
                    strVariant = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim();
					entryListItem.SubItems.Add(this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim());
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_VARIANT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					if (!System.IO.Directory.Exists(txtDataDir.Text.Trim() + "\\" + m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim()))
						System.IO.Directory.CreateDirectory(txtDataDir.Text.Trim() + "\\" + m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim());
					//rx
					entryListItem.SubItems.Add(this.m_ado.m_OleDbDataReader["rxpackage"].ToString().Trim());
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,COL_RX,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					//simulation year cycle
                    if (this.m_ado.m_OleDbDataReader["rxcycle_length"] != System.DBNull.Value)
                    {
                        this.m_strFVSCycleLength = Convert.ToString(this.m_ado.m_OleDbDataReader["rxcycle_length"]).Trim();
                    }
					//loc file
					entryListItem.SubItems.Add(" ");  //loc file
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_LOC,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					//FVS_StandInit Stand_ID count
					entryListItem.SubItems.Add(" ");
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_STANDCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					//FVS_TreeInit row count
					entryListItem.SubItems.Add(" ");
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_TREECOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					//out mdb file name
					entryListItem.SubItems.Add(" ");  //out mdb file
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_MDBOUT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					//summary record count
					entryListItem.SubItems.Add(" ");  //summary record count
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_SUMMARYCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					//treecut list record count
					entryListItem.SubItems.Add(" ");  //tree cut list record count
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_CUTCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					//tree standing (uncut) record count
					entryListItem.SubItems.Add(" ");  //tree standing record count
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_LEFTCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
					//potential fire record count
					entryListItem.SubItems.Add(" ");  //potential fire record count
					this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_fvs_input.COL_POTFIRECOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
                    //potential fire base year MDB file
                    entryListItem.SubItems.Add(" ");  //potential fire record count
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_POTFIREMDBOUT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //potential fire base year table count
                    entryListItem.SubItems.Add(" ");  //potential fire record count
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_POTFIREBASEYEARCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);

					//check to see if there is an input and output dsn name
					this.m_strLocFile = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + ".loc";
					this.m_strSlfFile = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + ".slf";
					this.m_strOutMDBFile = this.m_oRxTools.GetRxPackageFvsOutDbFileName(m_ado.m_OleDbDataReader);

					strOutDirAndFile=this.txtDataDir.Text.Trim() + "\\" + m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "\\" + this.m_strOutMDBFile.Trim();

					frmMain.g_sbpInfo.Text = "Processing FVS Input Variant/RxPackage " + 
						this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "/" + 
						this.m_ado.m_OleDbDataReader["rxpackage"].ToString().Trim() + "...Stand By"; 

					//check fvs in values

					strInDirAndFile = this.txtDataDir.Text.Trim() + "\\" + this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "\\" + this.m_strLocFile.Trim();
					if (System.IO.File.Exists(strInDirAndFile)==true)
					{
						entryListItem.SubItems[COL_LOC].Text = this.m_strLocFile;
    				}

				    strInDirAndFile = this.txtDataDir.Text.Trim() + "\\" + this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "\\" + "FVSIn.accdb";
				    if (frmMain.g_bSuppressFVSInputTableRowCount==false && System.IO.File.Exists(strInDirAndFile) == true)
				    {
				        //TODO: is " " preferable to "0" in the case that FVSIn exists but the table does not (not expected behavior), or the table exists but has no rows?
				        entryListItem.SubItems[COL_STANDCOUNT].Text = Convert.ToString(getStandInitCount(strInDirAndFile));
				        entryListItem.SubItems[COL_TREECOUNT].Text = Convert.ToString(getTreeInitCount(strInDirAndFile));
				    }


					//check dsn out registry values
					foreach(string strDsnName in strDsnNames)
					{
						//dsn in
						//dsn out
						if (this.m_strDsnOut.Trim().ToUpper() == strDsnName.Trim().ToUpper() &&
							regKey.GetValue(strDsnName).ToString().Trim().ToUpper() == 
							"MICROSOFT ACCESS DRIVER (*.MDB)")
						{
							strRegKey = "Software\\ODBC\\ODBC.ini\\" + this.m_strDsnOut.Trim();
							regKey2 = Registry.CurrentUser.OpenSubKey( strRegKey,false); 
							strValues = regKey2.GetValueNames(); // for the 1st time it's only 2 names
							foreach(string strValue in strValues)
							{
								if (strValue.Trim().ToUpper()=="DBQ")
								{
									if (strOutDirAndFile.Trim().ToUpper() == 
										regKey2.GetValue(strValue).ToString().Trim().ToUpper())
									{
										bFoundDsnOut=true;
										break;
									}
								}
							}
						}
						
						if (bFoundDsnOut==true) break;
					}


					if (System.IO.File.Exists(strOutDirAndFile) == true)
					{
						strConn = this.m_ado.getMDBConnString(strOutDirAndFile,"","");
						entryListItem.SubItems[COL_MDBOUT].Text = this.m_strOutMDBFile;
                        if (frmMain.g_bSuppressFVSInputTableRowCount==false)
                        {
                            if (this.m_dao.TableExists(strOutDirAndFile, "fvs_summary") == true)
                            {
                                entryListItem.SubItems[COL_SUMMARYCOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_summary", "fvs_summary")));
                            }
                            if (this.m_dao.TableExists(strOutDirAndFile, "fvs_cutlist") == true)
                            {
                                entryListItem.SubItems[COL_CUTCOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_cutlist", "fvs_cutlist")));
                            }
                            if (this.m_dao.TableExists(strOutDirAndFile, "fvs_treelist") == true)
                            {
                                entryListItem.SubItems[COL_LEFTCOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_treelist", "fvs_treelist")));
                            }
                            if (this.m_dao.TableExists(strOutDirAndFile, "fvs_potfire") == true)
                            {
                                entryListItem.SubItems[COL_POTFIRECOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_potfire", "fvs_potfire")));
                            }
                        }
					}
					else
					{
					}
                    //
                    //POTFIRE BASE YEAR
                    //
                    m_strOutPotFireBaseYearMDBFile = "FVSOUT_" + strVariant + "_POTFIRE_BaseYr.MDB";
                    strOutDirAndFile = this.txtDataDir.Text.Trim() + "\\" + strVariant + "\\" + m_strOutPotFireBaseYearMDBFile.Trim();
                    if (strVariant != strCurrentVariant)
                    {
                        strCurrentVariant = strVariant;
                        strRecordCount = "";
                        
                        if (System.IO.File.Exists(strOutDirAndFile) == true)
                        {
                            strConn = this.m_ado.getMDBConnString(strOutDirAndFile, "", "");
                            entryListItem.SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text = this.m_strOutPotFireBaseYearMDBFile;
                            if (frmMain.g_bSuppressFVSInputTableRowCount==false && 
                                this.m_dao.TableExists(strOutDirAndFile, "fvs_potfire") == true)
                            {
                                entryListItem.SubItems[uc_fvs_input.COL_POTFIREBASEYEARCOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_potfire", "fvs_potfire")));
                                strRecordCount = entryListItem.SubItems[uc_fvs_input.COL_POTFIREBASEYEARCOUNT].Text;
                            }
                        }
                    }
                    else
                    {
                        if (System.IO.File.Exists(strOutDirAndFile) == true)
                        {
                            entryListItem.SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text = this.m_strOutPotFireBaseYearMDBFile;
                            entryListItem.SubItems[uc_fvs_input.COL_POTFIREBASEYEARCOUNT].Text = strRecordCount;
                        }
                    }

				}
				this.m_ado.m_OleDbDataReader.Close();
				this.m_ado.CloseConnection(m_ado.m_OleDbConnection);
				Registry.CurrentUser.Close();
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					            "Module - uc_fvs_input:populate_listbox() \n" + 
					            "Err Msg - " + e.Message,
								"Create FVS Input",System.Windows.Forms.MessageBoxButtons.OK,
								System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
			this.Refresh();

		}
		private void LoadDataSources()
		{
			this.txtDataDir.Text = this.m_strProjDir + "\\fvs\\data";
			this.m_oQueries.m_oFvs.LoadDatasource=true;
			this.m_oQueries.m_oFIAPlot.LoadDatasource=true;
			this.m_oQueries.LoadDatasources(true);
			this.m_strConn = m_ado.getMDBConnString(this.m_oQueries.m_strTempDbFile,"","");




		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
            frmMain.g_oFrmMain.Enabled = true;
			this.ParentForm.Dispose();
			
		}

		private void btnRefresh_Click(object sender, System.EventArgs e)
		{
			this.populate_listbox();
		}

		private void btnChkAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstFvsInput.Items.Count-1;x++)
			{
				this.lstFvsInput.Items[x].Checked=true;
			}
		}

		private void btnClearAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstFvsInput.Items.Count-1;x++)
			{
				this.lstFvsInput.Items[x].Checked=false;
			}
		}

		private void btnAddDSN_Click(object sender, System.EventArgs e)
		{
			if (this.lstFvsInput.CheckedItems.Count == 0) 
			{
				MessageBox.Show("No Boxes Are Checked","Add DSN", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}

			this.btnRefresh.Enabled=false;

            this.btnExecuteAction.Enabled = false;
            this.cmbAction.Enabled = false;
			this.btnChkAll.Enabled=false;
			this.btnClearAll.Enabled=false;
			this.btnClose.Enabled=false;
			this.btnHelp.Enabled=false;
			this.AddDsn();
			

			
		}

		private void btnKillDSN_Click(object sender, System.EventArgs e)
		{
			
			if (this.lstFvsInput.CheckedItems.Count == 0) 
			{
				MessageBox.Show("No Boxes Are Checked","Remove DSN", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			
			this.cmbAction.Enabled=false;
			this.btnRefresh.Enabled=false;
			this.btnExecuteAction.Enabled=false;
			this.btnChkAll.Enabled=false;
			this.btnClearAll.Enabled=false;
			this.btnClose.Enabled=false;
			this.btnHelp.Enabled=false;
			this.RemoveDsn();
		
		}
		private void AddDsn()
		{


		}
		private void RemoveDsn()
		{
			bool bResult;
			try
			{
				bAbort=false;
				this.btnCancel.Visible=true;
				this.btnCancel.Refresh();
				this.progressBar1.Maximum = this.lstFvsInput.Items.Count;
				this.progressBar1.Minimum = 0;
				this.progressBar1.Value = 0;
				this.progressBar1.Visible=true;
				this.lblProgress.Text="";
				this.lblProgress.Visible=true;
				FIA_Biosum_Manager.DSNAdmin p_dsn = new DSNAdmin();
				for (int x=0;x<=this.lstFvsInput.Items.Count-1;x++)
				{
				

					if (this.lstFvsInput.Items[x].Checked==true)
					{
					
					
						System.Windows.Forms.Application.DoEvents();
						if (bAbort==true) break;
						

					}
					
					this.cmbAction.Enabled=true;
					this.btnRefresh.Enabled=true;
					this.btnExecuteAction.Enabled=true;
					this.btnChkAll.Enabled=true;
					this.btnClearAll.Enabled=true;
					this.btnClose.Enabled=true;
					this.btnHelp.Enabled=true;
					
					this.progressBar1.Visible=false;
					this.lblProgress.Visible=false;
					this.btnCancel.Visible=false;
				}
				p_dsn = null;
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_input:RemoveDsn \n" + 
					"Err Msg - " + err.Message,
					"Remove DSN",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}

		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			
			string strMsg = "Do you wish to cancel process (Y/N)?";
			DialogResult result = MessageBox.Show(strMsg,"Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					this.bAbort=true;
					

					return;
				case DialogResult.No:
					
					return;
			}                
		}
		private void CancelThreadCleanup()
		{
			
			this.cmbAction.Enabled=true;
			this.btnRefresh.Enabled=true;
			this.btnExecuteAction.Enabled=true;
			this.btnChkAll.Enabled=true;
			this.btnClearAll.Enabled=true;
			this.btnClose.Enabled=true;
			this.btnHelp.Enabled=true;
			
			this.progressBar1.Visible=false;
			this.lblProgress.Visible=false;
			this.btnCancel.Visible=false;

			this.m_thread = null;
		}

		private void btnWriteScript_Click(object sender, System.EventArgs e)
		{
			this.WriteKCPDbLines();
		}
		private void WriteKCPDbLines_Old()
		{
			if (this.lstFvsInput.CheckedItems.Count==0)
			{
				MessageBox.Show("No Boxes Are Checked","Write Script", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			
			try
			{
				string strFullPath = this.m_strProjDir.Trim() + "\\fvs\\scripts";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				System.IO.FileStream p_FileStream;
				System.IO.StreamWriter p_StreamWriter;
				string strDirAndFile = this.m_strProjDir.Trim() + "\\fvs\\scripts\\db_kcp.txt";
				string strLine;
				if (System.IO.File.Exists(strDirAndFile)==true)
				{
					p_FileStream = new System.IO.FileStream(strDirAndFile, System.IO.FileMode.Append, 
						System.IO.FileAccess.Write);
				}
				else
				{
					p_FileStream = new System.IO.FileStream(strDirAndFile, System.IO.FileMode.Create, 
						System.IO.FileAccess.Write);
				}
				p_StreamWriter = new System.IO.StreamWriter(p_FileStream);
				p_StreamWriter.WriteLine(" ");
				p_StreamWriter.WriteLine(" ");
				strLine = "-------------DB KCP Line Values------------- " + System.DateTime.Now.ToString();
				p_StreamWriter.WriteLine(strLine);
				p_StreamWriter.WriteLine(" ");
				for (int x=0; x<=this.lstFvsInput.Items.Count-1;x++)
				{
					if (this.lstFvsInput.Items[x].Checked==true)
					{
						if (this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim().Length > 0)
						{
							p_StreamWriter.WriteLine(" ");
							p_StreamWriter.WriteLine("DataBase");
							p_StreamWriter.WriteLine("DSNOUT");
							p_StreamWriter.WriteLine(this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim());
							p_StreamWriter.WriteLine("SUMMARY");
							p_StreamWriter.WriteLine("TREELIST           1");
							p_StreamWriter.WriteLine("CUTLIST           1");
							p_StreamWriter.WriteLine("POTFIRE            1");
							p_StreamWriter.WriteLine("End");
						}

					}
				}
				p_StreamWriter.Close();
				p_FileStream.Close();
				p_StreamWriter = null;
				p_FileStream = null;
                MessageBox.Show("Done", "FIA Biosum");
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_input:WriteKCPDbLines \n" + 
					"Err Msg - " + e.Message,
					"Write Script",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

			}
		}


        private void WriteKCPDbLines()
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "Write Script", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            try
            {
                

                System.IO.FileStream p_FileStream;
                System.IO.StreamWriter p_StreamWriter;
                
                string strDirAndFile = "";
                string strVariant = "";
                string strCurrentVariant = "";
                int x;
                
               
             
                for (x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
                {
                    if (this.lstFvsInput.Items[x].Checked == true)
                    {
                        
                        
                        if (this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim().Length > 0)
                        {
                            strVariant = lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim();
                            //
                            //RXPACKAGE KCP TEMPLATE
                            //
                            string strFullPath = this.m_strProjDir.Trim() + "\\fvs\\data\\" + lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim();
                            if (!System.IO.Directory.Exists(strFullPath))
                                System.IO.Directory.CreateDirectory(strFullPath);

                            string[] strArray = frmMain.g_oUtils.ConvertListToArray(lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim(), ".");
                            string[] strArray2 = frmMain.g_oUtils.ConvertListToArray(strArray[0], "-");
                            strDirAndFile = strFullPath + "\\" + strArray[0].Trim() + ".KCP.template";
                            
                            p_FileStream = new System.IO.FileStream(strDirAndFile, System.IO.FileMode.Create,
                                                                    System.IO.FileAccess.Write);
                            
                            p_StreamWriter = new System.IO.StreamWriter(p_FileStream);

                            p_StreamWriter.WriteLine("!!***************************************************");
                            p_StreamWriter.WriteLine("!!**Code generated by FIA Biosum                      ");
                            p_StreamWriter.WriteLine("!!**FVS command template script                       ");
                            p_StreamWriter.WriteLine("!!**RxPackage: " + lstFvsInput.Items[x].SubItems[COL_RX].Text.Trim());
                            if (strArray2[1] == "000")
                            p_StreamWriter.WriteLine("!!**RxCycle1: NA                                      ");
                            else
                            p_StreamWriter.WriteLine("!!**RxCycle1: " + strArray2[1]) ;
                            if (strArray2[2] == "000")
                            p_StreamWriter.WriteLine("!!**RxCycle2: NA                                      ");
                            else
                            p_StreamWriter.WriteLine("!!**RxCycle2: " + strArray2[2]);
                            if (strArray2[3] == "000")
                            p_StreamWriter.WriteLine("!!**RxCycle3: NA                                      ");
                            else
                            p_StreamWriter.WriteLine("!!**RxCycle3: " + strArray2[3]);
                            if (strArray2[4] == "000")
                            p_StreamWriter.WriteLine("!!**RxCycle4: NA                                      ");
                            else
                            p_StreamWriter.WriteLine("!!**RxCycle4: " + strArray2[4]);
                            p_StreamWriter.WriteLine("!!***************************************************");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!****************************************************");
                            p_StreamWriter.WriteLine("!!**Number Of Cycles and Years Between Cycles");
                            p_StreamWriter.WriteLine("!!****************************************************");
                            p_StreamWriter.WriteLine(" ");
                            if (m_strFVSCycleLength == "10")
                            p_StreamWriter.WriteLine("Timeint            0        10");
                            else
                            p_StreamWriter.WriteLine("Timeint            0         5");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("NUMCYCLE           4");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!************************************************");
                            p_StreamWriter.WriteLine("!!**Fuel And Fire Extension Keywords");
                            p_StreamWriter.WriteLine("!!************************************************");
                            p_StreamWriter.WriteLine("!!PotFire Note: The Potential Fire Report keyword (PotFire) needs to be added from within the Suppose interface");
                            p_StreamWriter.WriteLine("FMIn");
                            p_StreamWriter.WriteLine("PotFire            0       40.        1.");
                            p_StreamWriter.WriteLine("BurnRept					 1      200.");
                            p_StreamWriter.WriteLine("CarbRept           1      200.        1.");
                            p_StreamWriter.WriteLine("CarbCut            1      200.				1.");
                            p_StreamWriter.WriteLine("End");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!************************************************");
                            p_StreamWriter.WriteLine("!!**Output Tree Data Options");
                            p_StreamWriter.WriteLine("!!************************************************");
                            p_StreamWriter.WriteLine("!!Output Treelist");
                            p_StreamWriter.WriteLine("TreeList           0        3.         0         0         0         0         0");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!Output Cutlist");
                            p_StreamWriter.WriteLine("CutList            0        3.         0                             0");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!Output After Treatment Tree List File");
                            p_StreamWriter.WriteLine("ATRTLIST           0        3.         0");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!Output Tree Structural Statistics");
                            p_StreamWriter.WriteLine("StrClass           1       30.        5.       25.        5.      200.       30.");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!Turn-off tripling");
                            p_StreamWriter.WriteLine("NoTriple");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!Turn-off unneeded output: Delete stand composition table");
                            p_StreamWriter.WriteLine("DelOTab            1");
                            p_StreamWriter.WriteLine("!!Turn-off unneeded output: Delete selected sample tree table");
                            p_StreamWriter.WriteLine("DelOTab            2");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!************************************************");
                            p_StreamWriter.WriteLine("!!**Output To Access Database Keywords");
                            p_StreamWriter.WriteLine("!!************************************************");
                            p_StreamWriter.WriteLine("DataBase");
                            p_StreamWriter.WriteLine("DSNOUT");
                            p_StreamWriter.WriteLine(this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim());
                            p_StreamWriter.WriteLine("SUMMARY");
                            p_StreamWriter.WriteLine("TREELIST           1");
                            p_StreamWriter.WriteLine("CUTLIST            1");
                            p_StreamWriter.WriteLine("POTFIRE            1");
                            p_StreamWriter.WriteLine("ATRTLIST           1");
                            p_StreamWriter.WriteLine("BURNREPT           1");
                            p_StreamWriter.WriteLine("CARBRPTS           1");
                            p_StreamWriter.WriteLine("STRCLASS           1");
                            p_StreamWriter.WriteLine("End");
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!**************************************************");
                            p_StreamWriter.WriteLine("!!**Treatment Thinning Keywords                     ");
                            p_StreamWriter.WriteLine("!!**************************************************");
                            if (strArray2[1] != "000")
                            p_StreamWriter.WriteLine("!!Cycle 1 Treatment Keywords for Treatment Id " + strArray2[1]);
                            p_StreamWriter.WriteLine(" ");
                            if (strArray2[2] != "000")
                            p_StreamWriter.WriteLine("!!Cycle 2 Treatment Keywords for Treatment Id " + strArray2[2]);
                            p_StreamWriter.WriteLine(" ");
                            if (strArray2[3] != "000")
                            p_StreamWriter.WriteLine("!!Cycle 3 Treatment Keywords for Treatment Id " + strArray2[3]);
                            p_StreamWriter.WriteLine(" ");
                            if (strArray2[4] != "000")
                            p_StreamWriter.WriteLine("!!Cycle 4 Treatment Keywords for Treatment Id " + strArray2[4]);
                            p_StreamWriter.WriteLine(" ");
                            p_StreamWriter.WriteLine("!!End of FIA Biosum Generated Code!!");


                           


                            p_StreamWriter.Close();
                            p_FileStream.Close();
                            //
                            //POTFIRE TEMPLATE
                            //
                            if (strVariant != strCurrentVariant)
                            {
                                strCurrentVariant = strVariant;
                                strDirAndFile = strFullPath + "\\FVSOUT_" + strVariant + "_POTFIRE_BaseYr.KCP.template";

                                p_FileStream = new System.IO.FileStream(strDirAndFile, System.IO.FileMode.Create,
                                                                        System.IO.FileAccess.Write);

                                p_StreamWriter = new System.IO.StreamWriter(p_FileStream);

                                p_StreamWriter.WriteLine("!!***************************************************");
                                p_StreamWriter.WriteLine("!!**Code generated by FIA Biosum                      ");
                                p_StreamWriter.WriteLine("!!**FVS command POTFIRE Base Year template script     ");
                                p_StreamWriter.WriteLine("!!***************************************************");
                                p_StreamWriter.WriteLine(" ");
                                p_StreamWriter.WriteLine("!!****************************************************");
                                p_StreamWriter.WriteLine("!!**Number Of Cycles and Years Between Cycles");
                                p_StreamWriter.WriteLine("!!****************************************************");
                                p_StreamWriter.WriteLine(" ");
                                p_StreamWriter.WriteLine("Timeint            1        1");
                                p_StreamWriter.WriteLine(" ");
                                p_StreamWriter.WriteLine("NUMCYCLE           1");
                                p_StreamWriter.WriteLine(" ");
                                p_StreamWriter.WriteLine("!!************************************************");
                                p_StreamWriter.WriteLine("!!**Fuel And Fire Extension Keywords");
                                p_StreamWriter.WriteLine("!!************************************************");
                                p_StreamWriter.WriteLine("FMIn");
                                p_StreamWriter.WriteLine("PotFire            0        1.        0.");
                                p_StreamWriter.WriteLine("End");
                                p_StreamWriter.WriteLine(" ");
                                p_StreamWriter.WriteLine("!!************************************************");
                                p_StreamWriter.WriteLine("!!**Output To Access Database Keywords");
                                p_StreamWriter.WriteLine("!!************************************************");
                                p_StreamWriter.WriteLine("DataBase");
                                p_StreamWriter.WriteLine("DSNOUT");
                                p_StreamWriter.WriteLine("FVSOUT_" + strVariant + "_POTFIRE_BaseYr.MDB");
                                p_StreamWriter.WriteLine("POTFIRE            1");
                                p_StreamWriter.WriteLine("End");
                                p_StreamWriter.WriteLine(" ");
                                p_StreamWriter.WriteLine("!!End of FIA Biosum Generated Code!!");



                                p_StreamWriter.Close();
                                p_FileStream.Close();


                            }
                            p_StreamWriter = null;
                            p_FileStream = null;
                        }

                    }
                }
               
                MessageBox.Show("Done", "FIA Biosum");
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_input:WriteKCPDbLines \n" +
                    "Err Msg - " + e.Message,
                    "Write Script", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);

            }
        }
        private void ViewScripts()
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "Write Script", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }


            string strVariant = "";
            string strCurrentVariant = "";
            string strFullPath="";
            string strDirAndFile = "";
            int x;

            try
            {

                for (x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
                {
                    if (this.lstFvsInput.Items[x].Checked == true)
                    {


                        if (this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim().Length > 0)
                        {
                            strVariant = lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim();
                            strFullPath = this.m_strProjDir.Trim() + "\\fvs\\data\\" + lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim();
                            string[] strArray = frmMain.g_oUtils.ConvertListToArray(lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim(), ".");
                            string[] strArray2 = frmMain.g_oUtils.ConvertListToArray(strArray[0], "-");
                            strDirAndFile = strFullPath + "\\" + strArray[0].Trim() + ".KCP.template";
                            if (System.IO.File.Exists(strDirAndFile))
                            {

                                System.Diagnostics.Process.Start(@"notepad.exe", strDirAndFile); 
                                System.Threading.Thread.Sleep(2000);
                            }
                            if (strVariant != strCurrentVariant)
                            {
                                strCurrentVariant = strVariant;
                                strDirAndFile = strFullPath + "\\FVSOUT_" + strVariant + "_POTFIRE_BaseYr.KCP.template";
                                if (System.IO.File.Exists(strDirAndFile))
                                {
                                    System.Diagnostics.Process.Start(@"notepad.exe", strDirAndFile);
                                    System.Threading.Thread.Sleep(2000);
                                }

                            }

                        }

                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_input:ViewScripts \n" +
                    "Err Msg - " + err.Message,
                    "View Script", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
            }
        }

		private void btnViewScript_Click(object sender, System.EventArgs e)
		{
            ViewScripts();
		}
		private void AppendRecords()
		{

			string strCurVariant="";
			string strVariant="";
			string strSavedir = this.m_strProjDir + "\\fvs\\data\\save";
			bAbort=false;
			try
			{
				


				fvs_input p_fvsinput = new fvs_input(this.m_strProjDir,this.m_frmTherm);
                string strDataDir = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.txtDataDir, "Text", false);
                strDataDir = strDataDir.Trim();
				for (int x=0;x<=this.lstFvsInput.Items.Count-1;x++)
				{
					int intValue = Convert.ToInt32((double)(((double)(x+1) / (double)this.lstFvsInput.Items.Count) * 100));
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar2, "Value", intValue);
					if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(lstFvsInput,x,"Checked",false)==true)
					{
						//get the variant
                        strVariant = frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsInput,x,COL_VARIANT,"Text",false).ToString().Trim();
						//see if this is a new variant
						if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
						{
							strCurVariant = strVariant;
							//p_fvsinput.Start(this.txtInDir.Text.Trim(),strCurVariant);
							p_fvsinput.Start(strDataDir,strCurVariant);
						}
						frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar1,
                            "Value", 
                            frmMain.g_oDelegate.GetControlPropertyValue(
                                    m_frmTherm.progressBar1, "Maximum",false));

                        string strInDirAndFile = strDataDir + "\\" + strVariant + "\\" + strVariant + ".loc";
                        if (System.IO.File.Exists(strInDirAndFile) == true)
                        {
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_LOC, "Text", strVariant + ".loc");
                        }

					    strInDirAndFile = strDataDir + "\\" + strVariant + "\\" + "FVSIn.accdb";
					    if (System.IO.File.Exists(strInDirAndFile) == true) //redundant check here, but leaves " " instead of new "0"
					    {
					        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_STANDCOUNT, "Text",
					            Convert.ToString(getStandInitCount(strInDirAndFile)));
					        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_TREECOUNT, "Text",
					            Convert.ToString(getTreeInitCount(strInDirAndFile)));
					    }				

					}
					
                    frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar1,
                            "Value",
                            frmMain.g_oDelegate.GetControlPropertyValue(
                                    m_frmTherm.progressBar1, "Maximum", false));
					System.Windows.Forms.Application.DoEvents();
					if (bAbort==true) break;

				}
				
				
                frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar2,
                            "Value",
                            frmMain.g_oDelegate.GetControlPropertyValue(
                                    m_frmTherm.progressBar2, "Maximum", false));
				
				this.InputFVSRecordsFinished();
			}
			catch (System.Threading.ThreadInterruptedException err)
			{

				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch (System.Threading.ThreadAbortException err)
			{
				
				this.ThreadCleanUp();
				
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_input:AppendRecords  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"Append Records",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			ThreadCleanUp();


		}

		private void btnAppend_Click(object sender, System.EventArgs e)
		{
			if (this.lstFvsInput.CheckedItems.Count ==0)
			{
				MessageBox.Show("No Boxes Are Checked","Append", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			val_data();
			if (this.m_intError==0)
			{
				this.m_frmTherm = new frmTherm(((frmDialog)ParentForm),"APPEND FVS INPUT DATA",
					                             "FVS Input","2");


                this.m_frmTherm.Visible = false;
                this.m_frmTherm.lblMsg.Text="";
				this.Enabled=false;

                //progress bar 1: represents a single process
                this.m_frmTherm.progressBar1.Minimum = 0;
                this.m_frmTherm.progressBar1.Maximum = 100;
                this.m_frmTherm.progressBar1.Value = 0;
                this.m_frmTherm.lblMsg.Text = "";
                this.m_frmTherm.Show(this);
               

                //progress bar 2: represents overall progress 
                this.m_frmTherm.progressBar2.Minimum = 0;
                this.m_frmTherm.progressBar2.Maximum = 100;
                this.m_frmTherm.progressBar2.Value = 0;
                this.m_frmTherm.lblMsg2.Text = "Overall Progress";
				this.m_thread = new Thread(new ThreadStart(this.AppendRecords));
				this.m_thread.IsBackground = true;
				this.m_thread.Start();
			}
			

		}
		private void val_data()
		{
			this.m_intError=0;
			for (int x=0; x<=this.lstFvsInput.Items.Count-1;x++)
			{
				if (this.lstFvsInput.Items[x].Checked==true)
				{
				}
			}
		}
		public void StopThread()
		{
			this.m_thread.Suspend();
			string strMsg = "Do you wish to cancel appending  data (Y/N)?";
			DialogResult result = MessageBox.Show(strMsg,"Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					this.m_thread.Resume();
					this.m_frmTherm.AbortProcess = true;
					this.m_thread.Abort();
					if (this.m_thread.IsAlive)
					{
						this.m_frmTherm.lblMsg.Text = "Attempting To Abort Process...Stand By";
						this.m_frmTherm.lblMsg.Refresh();
						this.m_thread.Join();
					}
					if (this.m_frmTherm != null)
					{
						this.m_frmTherm.lblMsg.Text = "Cleaning Up...Stand By";
						this.m_frmTherm.lblMsg.Refresh();
					}
					this.ThreadCleanUp();
					return;
				case DialogResult.No:
					this.m_thread.Resume();
					return;
			}                
			
		}
		public void InputFVSRecordsFinished()
		{

			this.m_thread.Abort();
			if (this.m_thread.IsAlive)
			{
                frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.lblMsg,
                            "Text","Attempting To Abort Process...Stand By");
				
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm.lblMsg, "Refresh");
				
				this.m_thread.Join();
			}
			if (this.m_frmTherm != null)
			{
				frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm,"Close");
				frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm,"Dispose");
                this.m_frmTherm = null;
			}
			this.m_thread = null;

		}
		private void ThreadCleanUp()
		{
			try
			{
				
                frmMain.g_oDelegate.SetControlPropertyValue(cmbAction,"Enabled",true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnRefresh, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnExecuteAction, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnChkAll, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnClearAll, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnClose, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnHelp, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(progressBar1, "Visible", false);
				
                frmMain.g_oDelegate.SetControlPropertyValue(lblProgress, "Visible", false);
                frmMain.g_oDelegate.SetControlPropertyValue(btnCancel, "Visible", false);
                frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
				if (this.m_frmTherm != null)
				{
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                    this.m_frmTherm = null;
				}
				this.m_thread = null;
			}
			catch
			{
			}

		}
		private int getSlfPlotCount(string strDirAndFile)
		{
			int x=0;
			try
			{
				if (System.IO.File.Exists(strDirAndFile)==false) return 0;

				System.IO.FileStream p_fs = new FileStream(strDirAndFile, FileMode.Open, FileAccess.Read);
				System.IO.StreamReader p_sr = new StreamReader(p_fs);
				string strLine;
			
				// Read and display lines from the file until the end of 
				// the file is reached.
				while ((strLine = p_sr.ReadLine()) != null) 
				{
					if (strLine.Trim().Length > 0)
						x++;
				}
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_input:getSlfPlotCount  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FVS Input",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				if (x > 2)
				{
					x = (int)(x/2);
				}
				return x;
				
			}
			if (x > 2)
			{
				x = (int)(x/2);
			}
			return x;


            
		}

        private int getStandInitCount(string strDirAndFile)
        {
            int x = 0;
            try
            {

                if (System.IO.File.Exists(strDirAndFile))
                {
                    ado_data_access p_ado = new ado_data_access();
                    dao_data_access p_dao = new dao_data_access();
                    string strConn = p_ado.getMDBConnString(strDirAndFile, "", "");

                    if (p_dao.TableExists(strDirAndFile, "FVS_StandInit"))
                    {
                        System.Data.OleDb.OleDbConnection pConn = new System.Data.OleDb.OleDbConnection();
                        pConn.ConnectionString = strConn;
                        pConn.Open();
                        x = (int)m_ado.getSingleDoubleValueFromSQLQuery(pConn,
                            "SELECT COUNT(*) as StandInitCount FROM (SELECT DISTINCT Stand_ID FROM FVS_StandInit);",
                            "FVS_StandInit");
                        p_ado.CloseConnection(pConn);
                        pConn.Dispose();
                    }

                    p_ado = null;
                    p_dao = null;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_fvs_input:getStandInitCount  \n" +
                                "Err Msg - " + e.Message.ToString().Trim(),
                    "FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                //if (x > 2) //TODO: Why was this logic here in getSlfPlotCount?
                //{
                //    x = (int) (x / 2);
                //}
            }
            return x;
        }

        private int getTreeInitCount(string strDirAndFile)
        {
            int x = 0;
            try
            {

                if (System.IO.File.Exists(strDirAndFile))
                {
                    ado_data_access p_ado = new ado_data_access();
                    dao_data_access p_dao = new dao_data_access();
                    string strConn = p_ado.getMDBConnString(strDirAndFile, "", "");

                    if (p_dao.TableExists(strDirAndFile, "FVS_TreeInit"))
                    {
                        System.Data.OleDb.OleDbConnection pConn = new System.Data.OleDb.OleDbConnection();
                        pConn.ConnectionString = strConn;
                        pConn.Open();
                        x = (int)m_ado.getSingleDoubleValueFromSQLQuery(pConn,
                            "SELECT COUNT(*) as TreeInitCount FROM FVS_TreeInit;",
                            "FVS_TreeInit");
                        p_ado.CloseConnection(pConn);
                        pConn.Dispose();
                    }

                    p_ado = null;
                    p_dao = null;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_fvs_input:getTreeInitCount  \n" +
                                "Err Msg - " + e.Message.ToString().Trim(),
                    "FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                //if (x > 2) //TODO: Why was this logic here in getSlfPlotCount?
                //{
                //    x = (int) (x / 2);
                //}
            }
            return x;
        }

		private int getFvsTreeFileCount(string strVariant)
		{
			int x=0;
			try
			{
                string strDir = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.txtDataDir, "Text",false);
				
				if (System.IO.Directory.Exists(strDir.Trim() + "\\" + strVariant.Trim())==false) return 0;

				string strSearchPattern = "*.FVS";
				
				string[] strFiles =  System.IO.Directory.GetFiles(strDir.Trim() + "\\" + strVariant.Trim(),strSearchPattern);
				x = strFiles.Length;
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_input:getFvsTreeFileCount  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FVS Input",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return x;

			}
			return x;
			

		}

		private void btnPlotVariants_Click(object sender, System.EventArgs e)
		{
			
            frmMain.g_oFrmMain.StartPlotFVSVariantsDialog(this.ReferenceParentDialogForm);
		}

		private void btnTreeSpcVariants_Click(object sender, System.EventArgs e)
		{
			
            frmMain.g_oFrmMain.LoadProcessorTreeSpcForm(this.ReferenceParentDialogForm,"FVS");
		}

		private void btnRx_Click(object sender, System.EventArgs e)
		{
			
            frmMain.g_oFrmMain.StartRxDialog((frmDialog)ParentForm);
		}

		private void btnHelp_Click(object sender, System.EventArgs e)
		{
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "FVS", "INPUT_DATA" });
		}
		private void txtDataDir_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled = true;		
		}

		private void txtInDir_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled = true;		
		}

		private void txtOutDir_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled = true;
		}

		private void btnDeleteFvsOutTables_Click(object sender, System.EventArgs e)
		{
			if (this.lstFvsInput.CheckedItems.Count == 0) 
			{
				MessageBox.Show("No Boxes Are Checked","Delete FVS Out Tables", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}

			DialogResult result = MessageBox.Show("Backup The MDB File(s) Before Deleting The Tables? (Y/N)", 
				                                  "FIA Biosum",
				                                  System.Windows.Forms.MessageBoxButtons.YesNoCancel,
				                                  System.Windows.Forms.MessageBoxIcon.Question);
			switch (result)
			{
				case DialogResult.Yes:
					this.cmbAction.Enabled=false;
					
					this.btnExecuteAction.Enabled=false;
					this.btnRefresh.Enabled=false;
					
					this.btnChkAll.Enabled=false;
					this.btnClearAll.Enabled=false;
					this.btnClose.Enabled=false;
					this.btnHelp.Enabled=false;
					
					this.BackupFVSOutTables("B");
                    this.DeleteFVSOutTables("B");

					this.cmbAction.Enabled=true;
					
					this.btnExecuteAction.Enabled=true;
					this.btnRefresh.Enabled=true;
					this.btnChkAll.Enabled=true;
					this.btnClearAll.Enabled=true;
					this.btnClose.Enabled=true;
					this.btnHelp.Enabled=true;
					

				     break;
			    case DialogResult.No:
					this.cmbAction.Enabled=false;
					
					this.btnExecuteAction.Enabled=false;
					this.btnRefresh.Enabled=false;
					this.btnChkAll.Enabled=false;
					this.btnClearAll.Enabled=false;
					this.btnClose.Enabled=false;
					this.btnHelp.Enabled=false;
					

					this.DeleteFVSOutTables("B");

					this.btnExecuteAction.Enabled=true;
				
					this.cmbAction.Enabled=true;
					this.btnRefresh.Enabled=true;
					this.btnChkAll.Enabled=true;
					this.btnClearAll.Enabled=true;
					this.btnClose.Enabled=true;
					this.btnHelp.Enabled=true;
					

					break;
			}

		}

        private void BackupBeforeDelete(string p_strAction)
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "Delete FVS Out Tables", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            DialogResult result = MessageBox.Show("Backup The MDB File(s) Before Deleting The Tables? (Y/N)",
                                                  "FIA Biosum",
                                                  System.Windows.Forms.MessageBoxButtons.YesNoCancel,
                                                  System.Windows.Forms.MessageBoxIcon.Question);
            switch (result)
            {
                case DialogResult.Yes:
                    this.cmbAction.Enabled = false;
                    
                    this.btnExecuteAction.Enabled = false;
                    this.btnRefresh.Enabled = false;
                    this.btnChkAll.Enabled = false;
                    this.btnClearAll.Enabled = false;
                    this.btnClose.Enabled = false;
                    this.btnHelp.Enabled = false;
                   

                    this.BackupFVSOutTables(p_strAction);
                    this.DeleteFVSOutTables(p_strAction);

                    this.cmbAction.Enabled = true;
                    
                    this.btnExecuteAction.Enabled = true;
                    this.btnRefresh.Enabled = true;
                    this.btnChkAll.Enabled = true;
                    this.btnClearAll.Enabled = true;
                    this.btnClose.Enabled = true;
                    this.btnHelp.Enabled = true;
                    

                    break;
                case DialogResult.No:
                    this.cmbAction.Enabled = false;
                   
                    this.btnExecuteAction.Enabled = false;
                    this.btnRefresh.Enabled = false;
                    this.btnChkAll.Enabled = false;
                    this.btnClearAll.Enabled = false;
                    this.btnClose.Enabled = false;
                    this.btnHelp.Enabled = false;
                    
                    this.DeleteFVSOutTables(p_strAction);

                    this.cmbAction.Enabled = true;
                    
                    this.btnExecuteAction.Enabled = true;
                    this.btnRefresh.Enabled = true;
                    this.btnChkAll.Enabled = true;
                    this.btnClearAll.Enabled = true;
                    this.btnClose.Enabled = true;
                    this.btnHelp.Enabled = true;
                   

                    break;
            }

        }

		private void BackupFVSOutTables(string p_strAction)
		{
			try
			{
				bAbort=false;
				string strFile;
				string strNewFile;
                string strVariant = "";
                string strCurrentVariant = "";

				this.btnCancel.Visible=true;
				this.btnCancel.Refresh();
				this.progressBar1.Maximum = this.lstFvsInput.Items.Count;
				this.progressBar1.Minimum = 0;
				this.progressBar1.Value = 0;
				this.progressBar1.Visible=true;
				this.lblProgress.Text="";
				this.lblProgress.Visible=true;
				for (int x=0;x<=this.lstFvsInput.Items.Count-1;x++)
				{
					if (this.lstFvsInput.Items[x].Checked==true)
					{
                        strVariant = this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim();
						//out mdb file
						if (this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim().Length > 0 &&
                            (p_strAction == "B" || p_strAction == "S"))
						{
							strFile = this.txtDataDir.Text.Trim() + "\\" + this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "\\" + this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim();
							strNewFile = System.DateTime.Now.ToString().Trim() + "_" + this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim();
							strNewFile = strNewFile.Replace('/','_');
							strNewFile = strNewFile.Replace(':','_');
							strNewFile = strNewFile.Replace(' ','_');
							strNewFile = this.txtDataDir.Text.Trim() + "\\" + this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "\\" + strNewFile;
							if (System.IO.File.Exists(strFile)==true)
							{
								System.IO.File.Copy(strFile,strNewFile,true);
							}
						}
                        //potfire baseyr file
                        if (this.lstFvsInput.Items[x].SubItems[COL_POTFIREMDBOUT].Text.Trim().Length > 0 &&
                            (p_strAction == "B" || p_strAction == "P"))
                        {
                            if (strVariant != strCurrentVariant)
                            {
                                strCurrentVariant = strVariant;
                                strFile = this.txtDataDir.Text.Trim() + "\\" + this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "\\" + this.lstFvsInput.Items[x].SubItems[COL_POTFIREMDBOUT].Text.Trim();
                                strNewFile = System.DateTime.Now.ToString().Trim() + "_" + this.lstFvsInput.Items[x].SubItems[COL_POTFIREMDBOUT].Text.Trim();
                                strNewFile = strNewFile.Replace('/', '_');
                                strNewFile = strNewFile.Replace(':', '_');
                                strNewFile = strNewFile.Replace(' ', '_');
                                strNewFile = this.txtDataDir.Text.Trim() + "\\" + this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "\\" + strNewFile;
                                if (System.IO.File.Exists(strFile) == true)
                                {
                                    System.IO.File.Copy(strFile, strNewFile, true);
                                }
                            }
                        }
						System.Windows.Forms.Application.DoEvents();
						if (bAbort==true) break;

					}
					this.progressBar1.Visible=false;
					this.lblProgress.Visible=false;
					this.btnCancel.Visible=false;
				}
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_input:BackupFVSOutTables \n" + 
					"Err Msg - " + err.Message,
					"Delete FVS Out Tables",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}

			
		}
        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_strAction">B=BOTH,S=STANDARD,P=POTFIRE</param>
		private void DeleteFVSOutTables(string p_strAction)
		{
			try
			{
				bAbort=false;
				string strFile;
				//string strNewFile;
                string strMsg = "";
                string strTableItems = "";
                string strDbFileItem = "";
                string strVariant = "";
                string strCurrentVariant = "";
				this.btnCancel.Visible=true;
				this.btnCancel.Refresh();
				this.progressBar1.Maximum = this.lstFvsInput.Items.Count;
				this.progressBar1.Minimum = 0;
				this.progressBar1.Value = 0;
				this.progressBar1.Visible=true;
				this.lblProgress.Text="";
				this.lblProgress.Visible=true;
				for (int x=0;x<=this.lstFvsInput.Items.Count-1;x++)
				{
					if (this.lstFvsInput.Items[x].Checked==true)
					{
                        strVariant = this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim();
                        
						//out mdb file
						if (this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim().Length > 0 &&
                            (p_strAction=="B" || p_strAction=="S") )
						{
                            
							strFile = this.txtDataDir.Text.Trim() + "\\" + this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "\\" + this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim();
                            strDbFileItem = this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim() + "\r\n---------------------------------\r\n";
                            strTableItems = "";
							if (System.IO.File.Exists(strFile)==true)
							{
								m_ado.OpenConnection(m_ado.getMDBConnString(strFile,"",""));
								if (m_ado.m_intError==0)
								{
									string[] strTableNamesArray = m_ado.getTableNames(m_ado.m_OleDbConnection);
									if (strTableNamesArray != null && strTableNamesArray.Length > 0)
									{
										for (int y=0;y<=strTableNamesArray.Length - 1;y++)
										{
                                            if (strTableNamesArray[y].Trim().Length > 0)
                                            {
                                                m_ado.SqlNonQuery(m_ado.m_OleDbConnection, "DROP TABLE " + strTableNamesArray[y].Trim());
                                                strTableItems = strTableItems + strTableNamesArray[y].Trim() + ",";
                                                switch (strTableNamesArray[y].Trim().ToUpper())
                                                {
                                                    case "FVS_SUMMARY":
                                                        lstFvsInput.Items[x].SubItems[COL_SUMMARYCOUNT].Text = "";
                                                        break;
                                                    case "FVS_CUTLIST":
                                                        lstFvsInput.Items[x].SubItems[COL_CUTCOUNT].Text = "";
                                                        break;
                                                    case "FVS_TREELIST":
                                                        lstFvsInput.Items[x].SubItems[COL_LEFTCOUNT].Text = "";
                                                        break;
                                                    case "FVS_POTFIRE":
                                                        lstFvsInput.Items[x].SubItems[COL_POTFIRECOUNT].Text = "";
                                                        break;
                                                }
                                            }
										}
                                        
									}
									m_ado.CloseConnection(m_ado.m_OleDbConnection);
								}
							}
                            if (strTableItems.Trim().Length == 0)
                            {
                                strTableItems = "No Tables Deleted\r\n\r\n";
                            }
                            else
                            {
                                strTableItems = strTableItems.Substring(0, strTableItems.Length - 1);
                                strTableItems = strTableItems + "\r\n\r\n";
                            }
                            strMsg = strMsg + strDbFileItem + strTableItems;
                        }
                        if (this.lstFvsInput.Items[x].SubItems[COL_POTFIREMDBOUT].Text.Trim().Length > 0 &&
                            (p_strAction == "B" || p_strAction == "P"))
                        {
                            strFile = this.txtDataDir.Text.Trim() + "\\" + strVariant + "\\" + this.lstFvsInput.Items[x].SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text.Trim();
                            
                            if (System.IO.File.Exists(strFile) == true)
                            {
                                if (strVariant != strCurrentVariant)
                                {
                                    strDbFileItem = this.lstFvsInput.Items[x].SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text.Trim() + "\r\n---------------------------------\r\n";
                                    strTableItems = "";
                                    strCurrentVariant = strVariant;
                                    m_ado.OpenConnection(m_ado.getMDBConnString(strFile, "", ""));
                                    if (m_ado.m_intError == 0)
                                    {
                                        string[] strTableNamesArray = m_ado.getTableNames(m_ado.m_OleDbConnection);
                                        if (strTableNamesArray != null && strTableNamesArray.Length > 0)
                                        {
                                            for (int y = 0; y <= strTableNamesArray.Length - 1; y++)
                                            {
                                                if (strTableNamesArray[y].Trim().Length > 0)
                                                {
                                                    m_ado.SqlNonQuery(m_ado.m_OleDbConnection, "DROP TABLE " + strTableNamesArray[y].Trim());
                                                    strTableItems = strTableItems + strTableNamesArray[y].Trim() + ",";
                                                    switch (strTableNamesArray[y].Trim().ToUpper())
                                                    {
                                                        case "FVS_POTFIRE":
                                                            lstFvsInput.Items[x].SubItems[COL_POTFIREBASEYEARCOUNT].Text = "";
                                                            break;
                                                    }
                                                }
                                            }

                                        }
                                        m_ado.CloseConnection(m_ado.m_OleDbConnection);
                                    }
                                    if (strTableItems.Trim().Length == 0)
                                    {
                                        strTableItems = "No Tables Deleted\r\n\r\n";
                                    }
                                    else
                                    {
                                        strTableItems = strTableItems.Substring(0, strTableItems.Length - 1);
                                        strTableItems = strTableItems + "\r\n\r\n";
                                    }
                                    strMsg = strMsg + strDbFileItem + strTableItems;
                                }
                                
                            }
                           
                            

						}
						System.Windows.Forms.Application.DoEvents();
						if (bAbort==true) break;
					}
					this.progressBar1.Visible=false;
					this.lblProgress.Visible=false;
					this.btnCancel.Visible=false;
				}
                if (strMsg.Trim().Length > 0)
                {
                    //MessageBox.Show(strMsg, "Deleted Tables");
                    FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
                    frmTemp.Text = "FIA Biosum";
                    frmTemp.AutoScroll = false;
                    uc_textbox uc_textbox1 = new uc_textbox();
                    frmTemp.Controls.Add(uc_textbox1);
                    uc_textbox1.Dock = DockStyle.Fill;
                    uc_textbox1.lblTitle.Text = "Deleted Tables";
                    uc_textbox1.TextValue = strMsg;
                    frmTemp.ShowDialog();
                }
                else
                {
                    MessageBox.Show("No Tables Were Deleted", "FIA Biosum");
                }
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_input:BackupFVSOutTables \n" + 
					"Err Msg - " + err.Message,
					"Delete FVS Out Tables",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}

		}

		private void lstFvsInput_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			int x;
			if (e.Button == MouseButtons.Left)
			{
				try
				{
					int intRowHt = this.lstFvsInput.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lstFvsInput.Items[this.lstFvsInput.TopItem.Index + (int)dblRow-1].Selected=true;
				}
				catch
				{
				}
			}

		}

		private void lstFvsInput_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (lstFvsInput.SelectedItems.Count > 0)
				this.m_oLvRowColors.DelegateListViewItem(lstFvsInput.SelectedItems[0]);
		}

		private void btnCreateFvsOutFiles_Click(object sender, System.EventArgs e)
		{
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "Create FVSOut DB File", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            CreateFvsOutFiles();

           
		}

        private void CreateFvsOutFiles()
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "Create FVSOut DB File", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            
            string strVariant = "";
            string strCurrentVariant = "";

            string strMsg = "";
            string strFile = "";
            string strFullPath = "";
            dao_data_access oDao = new dao_data_access();
            m_ado.OpenConnection(m_strConn);
            for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
            {
                if (this.lstFvsInput.Items[x].Checked == true)
                {
                    //out mdb file
                    if (this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim().Length == 0)
                    {
                        strFile = this.m_oRxTools.GetRxPackageFvsOutDbFileName(m_ado.m_OleDbConnection,
                            this.m_oQueries.m_oFvs.m_strRxPackageTable,
                            this.m_oQueries.m_oFIAPlot.m_strPlotTable,
                            "(TRIM(a.fvs_variant)='" + this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "' AND " +
                            "b.rxpackage='" + this.lstFvsInput.Items[x].SubItems[COL_RX].Text.Trim() + "')");

                        if (strFile.Trim().Length > 0)
                        {
                            strFullPath = this.txtDataDir.Text + "\\" + lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "\\" + strFile;
                            if (System.IO.File.Exists(strFullPath) == false)
                            {
                                oDao.CreateMDB(strFullPath);
                                if (oDao.m_intError == 0)
                                {
                                    strMsg = strMsg + "\r\n" + strFile;
                                    this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text = strFile;
                                }
                            }
                        }


                    }
                    strVariant = this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim();
                    if (strVariant != strCurrentVariant)
                    {
                        strCurrentVariant = strVariant;
                        if (this.lstFvsInput.Items[x].SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text.Trim().Length == 0)
                        {
                            strFile = "FVSOUT_" + strVariant + "_POTFIRE_BaseYr.MDB";
                            strFullPath = this.txtDataDir.Text + "\\" + strVariant + "\\" + strFile;
                            if (System.IO.File.Exists(strFullPath) == false)
                            {
                                oDao.CreateMDB(strFullPath);
                                if (oDao.m_intError == 0)
                                {
                                    strMsg = strMsg + "\r\n" + strFile;
                                    this.lstFvsInput.Items[x].SubItems[COL_POTFIREMDBOUT].Text = strFile;
                                }
                            }
                        }

                    }
                    else
                    {
                        if (this.lstFvsInput.Items[x].SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text.Trim().Length == 0)
                        {
                            strFile = "FVSOUT_" + strVariant + "_POTFIRE_BaseYr.MDB";
                            strFullPath = this.txtDataDir.Text + "\\" + strVariant + "\\" + strFile;
                            if (System.IO.File.Exists(strFullPath))
                            {
                               this.lstFvsInput.Items[x].SubItems[COL_POTFIREMDBOUT].Text = strFile;
                            }
                        }
                    }
                }
            }
            m_ado.CloseConnection(m_ado.m_OleDbConnection);
            oDao = null;
            if (strMsg.Length == 0) strMsg = "No DB Files Created";
            else
            {
                strMsg = "DB Files Created\r\n-------------------\r\n" + strMsg;
            }
            FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
            frmTemp.Text = "FIA Biosum";
            frmTemp.AutoScroll = false;
            uc_textbox uc_textbox1 = new uc_textbox();
            frmTemp.Controls.Add(uc_textbox1);
            uc_textbox1.Dock = DockStyle.Fill;
            uc_textbox1.lblTitle.Text = "Create Files";
            uc_textbox1.TextValue = strMsg;
            frmTemp.ShowDialog();
        }
		private void btnRxPackage_Click(object sender, System.EventArgs e)
		{
            frmMain.g_oFrmMain.StartRxPackageDialog((frmDialog)ParentForm);
		}

	
		public string strProjectDirectory
		{
			set
			{
				this.m_strProjDir = value;
			}
			get
			{
				return this.m_strProjDir;
			}
		}
		public string strProjectId
		{
			set
			{
				this.m_strProjId = value;
			}
			get
			{
				return this.m_strProjId;
			}
		}
		public frmMain ReferenceMainForm
		{
			set {_frmMain=value;}
			get {return _frmMain;}
		}
		public frmDialog ReferenceParentDialogForm
		{
			set {_frmDialog=value;}
			get {return _frmDialog;}
		}

        private void btnExecuteAction_Click(object sender, EventArgs e)
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            string strAction = cmbAction.Text.Trim().ToUpper();
            switch (strAction)
            {
                case "CREATE FVS INPUT DATABASE FILES":
                    btnAppend_Click(null, null);
                    break;
                case "CREATE FVS OUTPUT DATABASE FILES":
                    CreateFvsOutFiles();
                    break;
                case "DELETE STANDARD FVS OUTPUT TABLES":
                    BackupBeforeDelete("S");
                    break;
                case "DELETE POTFIRE BASE YEAR OUTPUT TABLES":
                    BackupBeforeDelete("P");
                    break;
                case "DELETE BOTH STANDARD AND POTFIRE BASE YEAR OUTPUT TABLES":
                    BackupBeforeDelete("B");
                    break;
                case "WRITE KCP TEMPLATE SCRIPTS":
                    WriteKCPDbLines();
                    break;
                case "VIEW KCP TEMPLATE SCRIPTS":
                    ViewScripts();
                    break;
            }
        }

        private void cmbAction_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbAction_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
		
	

	}
}
