using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Linq;
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
        private TextBox txtMinCwdTL;
        private TextBox txtMinSmallFwdTL;
        private TextBox txtMinLargeFwdTL;
        private Label label4;
        private Label label3;
        private Label label2;
        private TabControl tabControl1;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private GroupBox grpDWMOptions;
        private CheckBox chkDwmFuelModel;
        private CheckBox chkDwmFuelBiomass;
        private CheckedListBox chkLstBoxLitterYears;
        private CheckedListBox chkLstBoxDuffYears;
        private GroupBox groupBox2;
        private Label label6;
        private Label label5;
        private LinkLabel linkLabelFuelModel;
        private GroupBox grpGRMOptions;
        private CheckBox chkGRM;
        private GroupBox otherOptionsGroupBox;
        private Label lblTreeAgeOffset;
        private TextBox treeAgeTxtBox;

        delegate string[] GetListBoxItemsDlg(CheckedListBox checkedListBox);

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

		    for (int i = 2001; i <= DateTime.Now.Year; i++)
		    {
		        this.chkLstBoxDuffYears.Items.Add(i.ToString());
		        this.chkLstBoxLitterYears.Items.Add(i.ToString());
		    }

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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnPlotVariants = new System.Windows.Forms.Button();
            this.txtDataDir = new System.Windows.Forms.TextBox();
            this.btnExecuteAction = new System.Windows.Forms.Button();
            this.lstFvsInput = new System.Windows.Forms.ListView();
            this.cmbAction = new System.Windows.Forms.ComboBox();
            this.lblRxCnt = new System.Windows.Forms.Label();
            this.lblVarCnt = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnChkAll = new System.Windows.Forms.Button();
            this.btnRxPackage = new System.Windows.Forms.Button();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.lblRxPackageCnt = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnRx = new System.Windows.Forms.Button();
            this.lblTreeSpcVarCnt = new System.Windows.Forms.Label();
            this.btnTreeSpcVariants = new System.Windows.Forms.Button();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.grpGRMOptions = new System.Windows.Forms.GroupBox();
            this.chkGRM = new System.Windows.Forms.CheckBox();
            this.grpDWMOptions = new System.Windows.Forms.GroupBox();
            this.linkLabelFuelModel = new System.Windows.Forms.LinkLabel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.chkLstBoxDuffYears = new System.Windows.Forms.CheckedListBox();
            this.chkLstBoxLitterYears = new System.Windows.Forms.CheckedListBox();
            this.chkDwmFuelModel = new System.Windows.Forms.CheckBox();
            this.chkDwmFuelBiomass = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMinLargeFwdTL = new System.Windows.Forms.TextBox();
            this.txtMinCwdTL = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMinSmallFwdTL = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.otherOptionsGroupBox = new System.Windows.Forms.GroupBox();
            this.lblTreeAgeOffset = new System.Windows.Forms.Label();
            this.treeAgeTxtBox = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.grpGRMOptions.SuspendLayout();
            this.grpDWMOptions.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.otherOptionsGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tabControl1);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.progressBar1);
            this.groupBox1.Controls.Add(this.lblProgress);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(784, 585);
            this.groupBox1.TabIndex = 99;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(7, 51);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(771, 483);
            this.tabControl1.TabIndex = 100;
            this.tabControl1.TabIndexChanged += new System.EventHandler(this.tabControl1_TabIndexChanged);
            this.tabControl1.Resize += new System.EventHandler(this.tabControl1_Resize);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnPlotVariants);
            this.tabPage2.Controls.Add(this.txtDataDir);
            this.tabPage2.Controls.Add(this.btnExecuteAction);
            this.tabPage2.Controls.Add(this.lstFvsInput);
            this.tabPage2.Controls.Add(this.cmbAction);
            this.tabPage2.Controls.Add(this.lblRxCnt);
            this.tabPage2.Controls.Add(this.lblVarCnt);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.btnChkAll);
            this.tabPage2.Controls.Add(this.btnRxPackage);
            this.tabPage2.Controls.Add(this.btnClearAll);
            this.tabPage2.Controls.Add(this.lblRxPackageCnt);
            this.tabPage2.Controls.Add(this.btnRefresh);
            this.tabPage2.Controls.Add(this.btnRx);
            this.tabPage2.Controls.Add(this.lblTreeSpcVarCnt);
            this.tabPage2.Controls.Add(this.btnTreeSpcVariants);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(763, 454);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Main Menu";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.tabPage2.SizeChanged += new System.EventHandler(this.tabControl1_Resize);
            // 
            // btnPlotVariants
            // 
            this.btnPlotVariants.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPlotVariants.Location = new System.Drawing.Point(6, 6);
            this.btnPlotVariants.Name = "btnPlotVariants";
            this.btnPlotVariants.Size = new System.Drawing.Size(232, 24);
            this.btnPlotVariants.TabIndex = 99;
            this.btnPlotVariants.Text = "Plots Missing FVS Variant Value";
            this.btnPlotVariants.Click += new System.EventHandler(this.btnPlotVariants_Click);
            // 
            // txtDataDir
            // 
            this.txtDataDir.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDataDir.Location = new System.Drawing.Point(96, 60);
            this.txtDataDir.Name = "txtDataDir";
            this.txtDataDir.Size = new System.Drawing.Size(661, 23);
            this.txtDataDir.TabIndex = 99;
            // 
            // btnExecuteAction
            // 
            this.btnExecuteAction.Location = new System.Drawing.Point(669, 332);
            this.btnExecuteAction.Name = "btnExecuteAction";
            this.btnExecuteAction.Size = new System.Drawing.Size(88, 32);
            this.btnExecuteAction.TabIndex = 5;
            this.btnExecuteAction.Text = "Execute Action";
            this.btnExecuteAction.Click += new System.EventHandler(this.btnExecuteAction_Click);
            // 
            // lstFvsInput
            // 
            this.lstFvsInput.CheckBoxes = true;
            this.lstFvsInput.GridLines = true;
            this.lstFvsInput.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFvsInput.HideSelection = false;
            this.lstFvsInput.Location = new System.Drawing.Point(6, 87);
            this.lstFvsInput.MultiSelect = false;
            this.lstFvsInput.Name = "lstFvsInput";
            this.lstFvsInput.Size = new System.Drawing.Size(751, 239);
            this.lstFvsInput.TabIndex = 0;
            this.lstFvsInput.UseCompatibleStateImageBehavior = false;
            this.lstFvsInput.View = System.Windows.Forms.View.Details;
            this.lstFvsInput.SelectedIndexChanged += new System.EventHandler(this.lstFvsInput_SelectedIndexChanged);
            this.lstFvsInput.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFvsInput_MouseUp);
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
            this.cmbAction.Location = new System.Drawing.Point(301, 339);
            this.cmbAction.Name = "cmbAction";
            this.cmbAction.Size = new System.Drawing.Size(362, 24);
            this.cmbAction.TabIndex = 4;
            this.cmbAction.Text = "<-------Action Items------->";
            this.cmbAction.SelectedIndexChanged += new System.EventHandler(this.cmbAction_SelectedIndexChanged);
            this.cmbAction.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbAction_KeyPress);
            // 
            // lblRxCnt
            // 
            this.lblRxCnt.BackColor = System.Drawing.Color.White;
            this.lblRxCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRxCnt.Location = new System.Drawing.Point(400, 10);
            this.lblRxCnt.Name = "lblRxCnt";
            this.lblRxCnt.Size = new System.Drawing.Size(32, 16);
            this.lblRxCnt.TabIndex = 99;
            this.lblRxCnt.Text = "0";
            // 
            // lblVarCnt
            // 
            this.lblVarCnt.BackColor = System.Drawing.Color.White;
            this.lblVarCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVarCnt.Location = new System.Drawing.Point(244, 10);
            this.lblVarCnt.Name = "lblVarCnt";
            this.lblVarCnt.Size = new System.Drawing.Size(56, 16);
            this.lblVarCnt.TabIndex = 99;
            this.lblVarCnt.Text = "0";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 24);
            this.label1.TabIndex = 99;
            this.label1.Text = "Data Directory";
            // 
            // btnChkAll
            // 
            this.btnChkAll.Location = new System.Drawing.Point(6, 332);
            this.btnChkAll.Name = "btnChkAll";
            this.btnChkAll.Size = new System.Drawing.Size(64, 32);
            this.btnChkAll.TabIndex = 1;
            this.btnChkAll.Text = "Check All";
            this.btnChkAll.Click += new System.EventHandler(this.btnChkAll_Click);
            // 
            // btnRxPackage
            // 
            this.btnRxPackage.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRxPackage.Location = new System.Drawing.Point(305, 30);
            this.btnRxPackage.Name = "btnRxPackage";
            this.btnRxPackage.Size = new System.Drawing.Size(90, 24);
            this.btnRxPackage.TabIndex = 99;
            this.btnRxPackage.Text = "Packages";
            this.btnRxPackage.Click += new System.EventHandler(this.btnRxPackage_Click);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(69, 332);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(64, 32);
            this.btnClearAll.TabIndex = 2;
            this.btnClearAll.Text = "Clear All";
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // lblRxPackageCnt
            // 
            this.lblRxPackageCnt.BackColor = System.Drawing.Color.White;
            this.lblRxPackageCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRxPackageCnt.Location = new System.Drawing.Point(400, 32);
            this.lblRxPackageCnt.Name = "lblRxPackageCnt";
            this.lblRxPackageCnt.Size = new System.Drawing.Size(32, 16);
            this.lblRxPackageCnt.TabIndex = 99;
            this.lblRxPackageCnt.Text = "0";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(132, 332);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(64, 32);
            this.btnRefresh.TabIndex = 3;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnRx
            // 
            this.btnRx.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRx.Location = new System.Drawing.Point(305, 6);
            this.btnRx.Name = "btnRx";
            this.btnRx.Size = new System.Drawing.Size(90, 24);
            this.btnRx.TabIndex = 99;
            this.btnRx.Text = "Treatments";
            this.btnRx.Click += new System.EventHandler(this.btnRx_Click);
            // 
            // lblTreeSpcVarCnt
            // 
            this.lblTreeSpcVarCnt.BackColor = System.Drawing.Color.White;
            this.lblTreeSpcVarCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTreeSpcVarCnt.Location = new System.Drawing.Point(243, 32);
            this.lblTreeSpcVarCnt.Name = "lblTreeSpcVarCnt";
            this.lblTreeSpcVarCnt.Size = new System.Drawing.Size(56, 16);
            this.lblTreeSpcVarCnt.TabIndex = 99;
            this.lblTreeSpcVarCnt.Text = "0";
            // 
            // btnTreeSpcVariants
            // 
            this.btnTreeSpcVariants.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTreeSpcVariants.Location = new System.Drawing.Point(6, 30);
            this.btnTreeSpcVariants.Name = "btnTreeSpcVariants";
            this.btnTreeSpcVariants.Size = new System.Drawing.Size(232, 24);
            this.btnTreeSpcVariants.TabIndex = 99;
            this.btnTreeSpcVariants.Text = "Tree Species Missing FVS Variant Value";
            this.btnTreeSpcVariants.Click += new System.EventHandler(this.btnTreeSpcVariants_Click);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.otherOptionsGroupBox);
            this.tabPage1.Controls.Add(this.grpGRMOptions);
            this.tabPage1.Controls.Add(this.grpDWMOptions);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(763, 454);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Options";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // grpGRMOptions
            // 
            this.grpGRMOptions.Controls.Add(this.chkGRM);
            this.grpGRMOptions.Location = new System.Drawing.Point(451, 6);
            this.grpGRMOptions.Name = "grpGRMOptions";
            this.grpGRMOptions.Size = new System.Drawing.Size(316, 85);
            this.grpGRMOptions.TabIndex = 102;
            this.grpGRMOptions.TabStop = false;
            this.grpGRMOptions.Text = "Growth Removal Mortality";
            // 
            // chkGRM
            // 
            this.chkGRM.AutoSize = true;
            this.chkGRM.Location = new System.Drawing.Point(6, 20);
            this.chkGRM.Name = "chkGRM";
            this.chkGRM.Size = new System.Drawing.Size(263, 21);
            this.chkGRM.TabIndex = 2;
            this.chkGRM.Text = "Use GRM calibration data if available";
            this.chkGRM.UseVisualStyleBackColor = true;
            // 
            // grpDWMOptions
            // 
            this.grpDWMOptions.Controls.Add(this.linkLabelFuelModel);
            this.grpDWMOptions.Controls.Add(this.groupBox2);
            this.grpDWMOptions.Controls.Add(this.chkDwmFuelModel);
            this.grpDWMOptions.Controls.Add(this.chkDwmFuelBiomass);
            this.grpDWMOptions.Controls.Add(this.label4);
            this.grpDWMOptions.Controls.Add(this.label2);
            this.grpDWMOptions.Controls.Add(this.txtMinLargeFwdTL);
            this.grpDWMOptions.Controls.Add(this.txtMinCwdTL);
            this.grpDWMOptions.Controls.Add(this.label3);
            this.grpDWMOptions.Controls.Add(this.txtMinSmallFwdTL);
            this.grpDWMOptions.Location = new System.Drawing.Point(6, 6);
            this.grpDWMOptions.Name = "grpDWMOptions";
            this.grpDWMOptions.Size = new System.Drawing.Size(439, 378);
            this.grpDWMOptions.TabIndex = 101;
            this.grpDWMOptions.TabStop = false;
            this.grpDWMOptions.Text = "Down Woody Material";
            // 
            // linkLabelFuelModel
            // 
            this.linkLabelFuelModel.AutoSize = true;
            this.linkLabelFuelModel.LinkArea = new System.Windows.Forms.LinkArea(8, 23);
            this.linkLabelFuelModel.Location = new System.Drawing.Point(23, 20);
            this.linkLabelFuelModel.Name = "linkLabelFuelModel";
            this.linkLabelFuelModel.Size = new System.Drawing.Size(477, 20);
            this.linkLabelFuelModel.TabIndex = 105;
            this.linkLabelFuelModel.TabStop = true;
            this.linkLabelFuelModel.Text = "Include Scott and Burgan (2005) surface fuel model (from DWM_fuelbed_typcd)";
            this.linkLabelFuelModel.UseCompatibleTextRendering = true;
            this.linkLabelFuelModel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelFuelModel_LinkClicked);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.chkLstBoxDuffYears);
            this.groupBox2.Controls.Add(this.chkLstBoxLitterYears);
            this.groupBox2.Location = new System.Drawing.Point(7, 142);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(236, 229);
            this.groupBox2.TabIndex = 104;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Duff/Litter Years to Exclude";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(123, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(40, 17);
            this.label6.TabIndex = 105;
            this.label6.Text = "Litter";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 16);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(34, 17);
            this.label5.TabIndex = 104;
            this.label5.Text = "Duff";
            // 
            // chkLstBoxDuffYears
            // 
            this.chkLstBoxDuffYears.FormattingEnabled = true;
            this.chkLstBoxDuffYears.Location = new System.Drawing.Point(9, 35);
            this.chkLstBoxDuffYears.Name = "chkLstBoxDuffYears";
            this.chkLstBoxDuffYears.Size = new System.Drawing.Size(100, 174);
            this.chkLstBoxDuffYears.TabIndex = 6;
            // 
            // chkLstBoxLitterYears
            // 
            this.chkLstBoxLitterYears.FormattingEnabled = true;
            this.chkLstBoxLitterYears.Location = new System.Drawing.Point(126, 35);
            this.chkLstBoxLitterYears.Name = "chkLstBoxLitterYears";
            this.chkLstBoxLitterYears.Size = new System.Drawing.Size(100, 174);
            this.chkLstBoxLitterYears.TabIndex = 7;
            // 
            // chkDwmFuelModel
            // 
            this.chkDwmFuelModel.AutoSize = true;
            this.chkDwmFuelModel.Location = new System.Drawing.Point(7, 19);
            this.chkDwmFuelModel.Name = "chkDwmFuelModel";
            this.chkDwmFuelModel.Size = new System.Drawing.Size(18, 17);
            this.chkDwmFuelModel.TabIndex = 1;
            this.chkDwmFuelModel.UseVisualStyleBackColor = true;
            // 
            // chkDwmFuelBiomass
            // 
            this.chkDwmFuelBiomass.AutoSize = true;
            this.chkDwmFuelBiomass.Location = new System.Drawing.Point(7, 39);
            this.chkDwmFuelBiomass.Name = "chkDwmFuelBiomass";
            this.chkDwmFuelBiomass.Size = new System.Drawing.Size(417, 21);
            this.chkDwmFuelBiomass.TabIndex = 2;
            this.chkDwmFuelBiomass.Text = "Calculate fuel biomasses with available DWM data (tons/acre)";
            this.chkDwmFuelBiomass.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label4.Location = new System.Drawing.Point(58, 119);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(229, 17);
            this.label4.TabIndex = 99;
            this.label4.Tag = "txtMinCwdTL";
            this.label4.Text = "Minimum CWD Transect Length (ft)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label2.Location = new System.Drawing.Point(58, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(266, 17);
            this.label2.TabIndex = 99;
            this.label2.Tag = "txtMinSmallFwdTL";
            this.label2.Text = "Minimum Small FWD Transect Length (ft)";
            // 
            // txtMinLargeFwdTL
            // 
            this.txtMinLargeFwdTL.Location = new System.Drawing.Point(7, 90);
            this.txtMinLargeFwdTL.Name = "txtMinLargeFwdTL";
            this.txtMinLargeFwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinLargeFwdTL.TabIndex = 4;
            this.txtMinLargeFwdTL.Text = "30";
            this.txtMinLargeFwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinLargeFwdTL_Validating);
            // 
            // txtMinCwdTL
            // 
            this.txtMinCwdTL.Location = new System.Drawing.Point(7, 116);
            this.txtMinCwdTL.Name = "txtMinCwdTL";
            this.txtMinCwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinCwdTL.TabIndex = 5;
            this.txtMinCwdTL.Text = "48";
            this.txtMinCwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinCwdTL_Validating);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label3.Location = new System.Drawing.Point(58, 93);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(269, 17);
            this.label3.TabIndex = 99;
            this.label3.Tag = "txtMinLargeFwdTL";
            this.label3.Text = "Minimum Large FWD Transect Length (ft)";
            // 
            // txtMinSmallFwdTL
            // 
            this.txtMinSmallFwdTL.Location = new System.Drawing.Point(7, 65);
            this.txtMinSmallFwdTL.Name = "txtMinSmallFwdTL";
            this.txtMinSmallFwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinSmallFwdTL.TabIndex = 3;
            this.txtMinSmallFwdTL.Text = "10";
            this.txtMinSmallFwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinSmallFwdTL_Validating);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(567, 540);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(56, 24);
            this.btnCancel.TabIndex = 99;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(312, 540);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(240, 8);
            this.progressBar1.TabIndex = 99;
            this.progressBar1.Visible = false;
            // 
            // lblProgress
            // 
            this.lblProgress.Location = new System.Drawing.Point(312, 556);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(239, 16);
            this.lblProgress.TabIndex = 99;
            this.lblProgress.Text = "lblProgress";
            this.lblProgress.Visible = false;
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(7, 540);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 99;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(678, 536);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(778, 32);
            this.lblTitle.TabIndex = 99;
            this.lblTitle.Text = "Create FVS Input";
            // 
            // otherOptionsGroupBox
            // 
            this.otherOptionsGroupBox.Controls.Add(this.lblTreeAgeOffset);
            this.otherOptionsGroupBox.Controls.Add(this.treeAgeTxtBox);
            this.otherOptionsGroupBox.Location = new System.Drawing.Point(451, 96);
            this.otherOptionsGroupBox.Name = "otherOptionsGroupBox";
            this.otherOptionsGroupBox.Size = new System.Drawing.Size(316, 85);
            this.otherOptionsGroupBox.TabIndex = 103;
            this.otherOptionsGroupBox.TabStop = false;
            this.otherOptionsGroupBox.Text = "Other Options";
            // 
            // lblTreeAgeOffset
            // 
            this.lblTreeAgeOffset.AutoSize = true;
            this.lblTreeAgeOffset.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.lblTreeAgeOffset.Location = new System.Drawing.Point(56, 26);
            this.lblTreeAgeOffset.Name = "lblTreeAgeOffset";
            this.lblTreeAgeOffset.Size = new System.Drawing.Size(254, 17);
            this.lblTreeAgeOffset.TabIndex = 101;
            this.lblTreeAgeOffset.Tag = "";
            this.lblTreeAgeOffset.Text = "Tree Age Offset, added to BHAge (yrs)";
            // 
            // treeAgeTxtBox
            // 
            this.treeAgeTxtBox.Location = new System.Drawing.Point(5, 23);
            this.treeAgeTxtBox.Name = "treeAgeTxtBox";
            this.treeAgeTxtBox.Size = new System.Drawing.Size(45, 22);
            this.treeAgeTxtBox.TabIndex = 100;
            this.treeAgeTxtBox.Text = "4";
            this.treeAgeTxtBox.Validating += new System.ComponentModel.CancelEventHandler(this.treeAgeTxtBox_Validating);
            // 
            // uc_fvs_input
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_input";
            this.Size = new System.Drawing.Size(784, 585);
            this.Resize += new System.EventHandler(this.uc_fvs_input_Resize);
            this.groupBox1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.grpGRMOptions.ResumeLayout(false);
            this.grpGRMOptions.PerformLayout();
            this.grpDWMOptions.ResumeLayout(false);
            this.grpDWMOptions.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.otherOptionsGroupBox.ResumeLayout(false);
            this.otherOptionsGroupBox.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_fvs_input_Resize(object sender, System.EventArgs e)
		{
			this.Resize_Fvs_Input();
            this.tabControl1_Resize(sender, e);
		}

		public void Resize_Fvs_Input()
		{
			try
			{
                progressBar1.Left = (int)(groupBox1.Width * .50) - (int)(progressBar1.Width * .50);
                progressBar1.Top = btnClose.Top;

                if (progressBar1.Left < (btnHelp.Left + btnHelp.Width))
                {
                    progressBar1.Left = btnHelp.Left + btnHelp.Width;
                }
                btnCancel.Left = progressBar1.Left + progressBar1.Width + 2;
                btnCancel.Top = progressBar1.Top - (int)(btnCancel.Height * .50) + (int)(progressBar1.Height * .50);
                if (btnClose.Left < (btnCancel.Left + btnCancel.Width))
                {
                    btnClose.Left = btnCancel.Left + btnCancel.Width;
                }
                lblProgress.Left = progressBar1.Left;
                lblProgress.Top = progressBar1.Top + progressBar1.Height + 2;
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

                //Keep a count of records in FVS_StandInit and FVS_TreeInit tables in each variant
                Dictionary<string, int[]> pVariantCountsDict = new Dictionary<string,int[]>();

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
				    if (pVariantCountsDict.ContainsKey(strVariant) == false)
				    {
				        pVariantCountsDict.Add(strVariant, null); //fvs_standinit, fvs_treeinit counts
				    }

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
				        if (pVariantCountsDict[strVariant] == null)
				        {
				            pVariantCountsDict[strVariant] = getFVSInputRecordCounts(strInDirAndFile);
				        }
				        entryListItem.SubItems[COL_STANDCOUNT].Text = Convert.ToString(pVariantCountsDict[strVariant][0]);
				        entryListItem.SubItems[COL_TREECOUNT].Text = Convert.ToString(pVariantCountsDict[strVariant][1]);
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

        private string[] GetCheckedListBoxItems(CheckedListBox chkListBox){
            if (chkListBox.InvokeRequired)
            {
                var dlg = new GetListBoxItemsDlg(GetCheckedListBoxItems);
                return chkListBox.Invoke(dlg, chkListBox) as string[];
            }

            string[] items = (from object item in chkListBox.CheckedItems select item.ToString())
                .ToArray();
            return items;
        }

	    private void ConfigureFvsInput(fvs_input p_fvs)
	    {
            //Down Woody Materials Section
	        if (chkDwmFuelModel.Checked && chkDwmFuelBiomass.Checked)
	        {
	            p_fvs.intDWMOption = (int) fvs_input.m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA;
	        }
	        else if (chkDwmFuelModel.Checked)
	        {
	            p_fvs.intDWMOption = (int) fvs_input.m_enumDWMOption.USE_FUEL_MODEL_ONLY;
	        }
	        else if (chkDwmFuelBiomass.Checked)
	        {
	            p_fvs.intDWMOption = (int) fvs_input.m_enumDWMOption.USE_DWM_DATA_ONLY;
	        }
	        else
	        {
	            p_fvs.intDWMOption = (int) fvs_input.m_enumDWMOption.SKIP_FUEL_MODEL_AND_DWM_DATA;
	        }

	        p_fvs.strMinSmallFwdTransectLengthTotal =
	            frmMain.g_oDelegate.GetControlPropertyValue(txtMinSmallFwdTL, "Text", false).ToString();
	        p_fvs.strMinLargeFwdTransectLengthTotal =
	            frmMain.g_oDelegate.GetControlPropertyValue(txtMinLargeFwdTL, "Text", false).ToString();
	        p_fvs.strMinCwdTransectLengthTotal =
	            frmMain.g_oDelegate.GetControlPropertyValue(txtMinCwdTL, "Text", false).ToString();

	        bool bFirst = true;
	        foreach (var item in GetCheckedListBoxItems(chkLstBoxDuffYears))
	        {
	            if (bFirst)
	            {
	                p_fvs.strDuffExcludedYears += item;
	                bFirst = false;
	            }
	            else
	            {
	                p_fvs.strDuffExcludedYears += ", " + item;
	            }
	        }
            bFirst = true;
	        foreach (var item in GetCheckedListBoxItems(chkLstBoxLitterYears))
	        {
	            if (bFirst)
	            {
	                p_fvs.strLitterExcludedYears += item;
	                bFirst = false;
	            }
	            else
	            {
	                p_fvs.strLitterExcludedYears += ", " + item;
	            }
	        }

            //Growth Removal Mortality section
            p_fvs.bUseGrmCalibrationData = (bool) frmMain.g_oDelegate.GetControlPropertyValue(chkGRM, "Checked", false);

            //Other Options section
	        p_fvs.strTreeAgeOffset = frmMain.g_oDelegate.GetControlPropertyValue(treeAgeTxtBox, "Text", false).ToString();
	    }

		private void AppendRecords()
		{
		    m_intError = 0;
			string strCurVariant="";
			string strVariant="";
			string strSavedir = this.m_strProjDir + "\\fvs\\data\\save";

			bAbort=false;
			try
			{
				fvs_input p_fvsinput = new fvs_input(this.m_strProjDir,this.m_frmTherm);

			    if (p_fvsinput.m_intError != 0)
			    {
                    return;
			    }

                ConfigureFvsInput(p_fvsinput);

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
					        int[] fvsInputRecordCounts = getFVSInputRecordCounts(strInDirAndFile);
					        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_STANDCOUNT, "Text",
					            Convert.ToString(fvsInputRecordCounts[0]));
					        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_TREECOUNT, "Text",
					            Convert.ToString(fvsInputRecordCounts[1]));
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
				
			}
			catch (System.Threading.ThreadInterruptedException err)
			{
			    m_intError = -1;
				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch (System.Threading.ThreadAbortException err)
			{
                m_intError = -1;	
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
			finally
			{
			    if (m_intError == 0)
			    {
			        ThreadCleanUp();
			    }
			}

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
					    while (m_thread.ThreadState != ThreadState.Aborted)
					    {
					        this.m_thread.Join(2000);
					    }
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
				if (this.m_frmTherm != null)
				{
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                    this.m_frmTherm = null;
				}
				
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

	    private int[] getFVSInputRecordCounts(string strDirAndFile)
	    {
            int stands = 0;
            int trees = 0;
            try
            {
                if (System.IO.File.Exists(strDirAndFile))
                {
                    ado_data_access p_ado = new ado_data_access();
                    string strConn = p_ado.getMDBConnString(strDirAndFile, "", "");

                    using (var pConn = new System.Data.OleDb.OleDbConnection(strConn)) 
                    {
                        pConn.Open();
                        if (p_ado.TableExist(pConn, "FVS_StandInit"))
                        {
                            stands = (int) p_ado.getSingleDoubleValueFromSQLQuery(pConn,
                            "SELECT COUNT(*) as StandInitCount FROM (SELECT DISTINCT Stand_ID FROM FVS_StandInit);",
                            "FVS_StandInit");
                        }
                        if (p_ado.TableExist(pConn, "FVS_TreeInit"))
                        {
                            trees = (int) p_ado.getSingleDoubleValueFromSQLQuery(pConn,
                            "SELECT COUNT(*) as TreeInitCount FROM FVS_TreeInit;",
                            "FVS_TreeInit");
                        }

                    }
                    p_ado = null;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_fvs_input:getFVSInputRecordCounts  \n" +
                                "Err Msg - " + e.Message.ToString().Trim(),
                    "FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            return new int[] {stands,trees};
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

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void txtMinSmallFwdTL_Validating(object sender, CancelEventArgs e)
        {
            double temp;
            if (!double.TryParse(txtMinSmallFwdTL.Text, out temp))
            {
                txtMinSmallFwdTL.Text = "10";
            }
        }

        private void txtMinLargeFwdTL_Validating(object sender, CancelEventArgs e)
        {
            double temp;
            if (!double.TryParse(txtMinLargeFwdTL.Text, out temp))
            {
                txtMinLargeFwdTL.Text = "30";
            }
        }

        private void txtMinCwdTL_Validating(object sender, CancelEventArgs e)
        {
            double temp;
            if (!double.TryParse(txtMinCwdTL.Text, out temp))
            {
                txtMinCwdTL.Text = "48";
            }
        }

        private void treeAgeTxtBox_Validating(object sender, CancelEventArgs e)
        {
            double temp;
            if (!double.TryParse(treeAgeTxtBox.Text, out temp))
            {
                treeAgeTxtBox.Text = "4";
            }
        }

        private void tabControl1_Resize(object sender, EventArgs e)
        {
            tabControl1.Top = this.lblTitle.Bottom + 5;
            tabControl1.Left = 5;
            tabControl1.Width = this.Width - 10;
            tabControl1.Height = this.Height - 100;

            btnClose.Top = groupBox1.Bottom - btnClose.Height - 5;
            btnClose.Left = tabControl1.Right - btnClose.Width;
            btnHelp.Left = tabControl1.Left;
            btnHelp.Top = btnClose.Top;


            label1.Left = tabPage2.Left + 5;
            txtDataDir.Left = label1.Right + 5;
            txtDataDir.Width = tabPage2.Width - txtDataDir.Left;

            //resize all of the controls on the inside
            lstFvsInput.Left = tabPage2.Left + 5;
            lstFvsInput.Width = tabPage2.Width - 10;
            lstFvsInput.Top = txtDataDir.Bottom + 5;
            lstFvsInput.Height = tabPage2.Height - txtDataDir.Bottom - 130;

            //btns under lstFvsInput position based on tabControl perimeter
            btnExecuteAction.Top = lstFvsInput.Bottom + 5;
            btnExecuteAction.Left = lstFvsInput.Right - btnExecuteAction.Width;

            cmbAction.Top = btnExecuteAction.Top + (int) (btnExecuteAction.Height * .5) - (int) (cmbAction.Height * .5);
            cmbAction.Left = btnExecuteAction.Left - cmbAction.Width - 5;

            btnChkAll.Top = btnExecuteAction.Top;
            btnChkAll.Left = lstFvsInput.Left;
            btnClearAll.Top = btnExecuteAction.Top;
            btnClearAll.Left = btnChkAll.Right + 5;
            btnRefresh.Top = btnExecuteAction.Top;
            btnRefresh.Left = btnClearAll.Right + 5;
        }

	    private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            tabPage1.Refresh();
            tabPage2.Refresh();
        }

        private void linkLabelFuelModel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                VisitLink();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to open link that was clicked.");
            }  
        }

	    private void VisitLink()
	    {
	        // Change the color of the link text by setting LinkVisited   
	        // to true.  
	        linkLabelFuelModel.LinkVisited = true;
	        //Call the Process.Start method to open the default browser   
	        //with a URL:  
            System.Diagnostics.Process.Start("https://www.fs.usda.gov/treesearch/pubs/download/9521.pdf");
	    }

       

	}
}
