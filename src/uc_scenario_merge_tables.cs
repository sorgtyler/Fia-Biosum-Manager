using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_merge_tables.
	/// </summary>
	public class uc_scenario_merge_tables : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Button btnHelp;
		public System.Windows.Forms.GroupBox grpboxOpen;
		private System.Windows.Forms.Button btnOpenNext;
		private System.Windows.Forms.Button btnOpenCancel;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.RadioButton rdoExistingMerge;
		private System.Windows.Forms.RadioButton rdoNewMerge;
		private System.Windows.Forms.GroupBox grpboxOpen2;
		private System.Windows.Forms.GroupBox grpboxSelectExisting;
		private System.Windows.Forms.ComboBox cmbMDBFiles;
		private System.Windows.Forms.Button btnSelectExisting;
		private System.Windows.Forms.GroupBox grpboxSelectExisting2;
		private System.Windows.Forms.Button btnSelectExistingPrevious;
		private System.Windows.Forms.Button btnSelectExistingCancel;
		private System.Windows.Forms.Button btnSelectExistingViewTablesInAccess;
		private System.Windows.Forms.Button btnSelectExistingViewTables;
		private string m_strScenarioMDB;
		private string m_strScenarioConn;
		private string m_strScenarioMergeMDB="";
		//private string m_strScenarioMergeConn;
		public int m_intHt = 350;
		public int m_intWd = 680;
		private FIA_Biosum_Manager.frmGridView m_frmGridView;
		private int m_intError;
		//private string m_strError;
		private FIA_Biosum_Manager.frmMain m_frmMain;
		//private string m_strCurrComboText="";
		private System.Windows.Forms.GroupBox grpboxNewSelectScenarios;
		private System.Windows.Forms.Button btnClearAllExistingTables;
		private System.Windows.Forms.Button btnSelectAllExistingTables;
		private System.Windows.Forms.CheckedListBox lstSelectExistingTables;
		private System.Windows.Forms.CheckedListBox lstSelectScenarios;
		private System.Windows.Forms.GroupBox grpboxNewSelectScenarios2;
		private System.Windows.Forms.CheckedListBox lstNewSelectTables;
		private System.Windows.Forms.GroupBox grpboxNewSelectView;
		private System.Windows.Forms.GroupBox grpboxNewSelectView2;
		private System.Windows.Forms.Button btnNewSelectScenariosCancel;
		private System.Windows.Forms.Button btnNewSelectScenariosNext;
		private System.Windows.Forms.Button btnNewSelectScenariosPrev;
		private System.Windows.Forms.Button btnSelectAllNewTables;
		private System.Windows.Forms.Button btnClearAllNewTables;
		private System.Windows.Forms.Button btnSelectAllScenarios;
		private System.Windows.Forms.Button btnClearAllScenarios;
		private System.Windows.Forms.CheckBox chkboxNewSelectViewSaveToAccess;
		private System.Windows.Forms.Button btnNewSelectViewCancel;
		private System.Windows.Forms.Button btnNewSelectViewPrev;
		private System.Windows.Forms.Button btnViewSelectViewInGrid;
		private System.Windows.Forms.Button btnNewSelectViewStartAccess;
		private System.Windows.Forms.Button btnNewSelectViewGetMDBFile;
		private System.Windows.Forms.TextBox txtNewSelectViewMergeMDBFile;
		private string m_strRandomFile="";
		private string m_strRandomFileConn="";
		private string[] m_strMergeTables;
		private ado_data_access m_adoLinks;

		private FIA_Biosum_Manager.frmTherm m_frmTherm;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_scenario_merge_tables(FIA_Biosum_Manager.frmMain p_frmMain)
		{
			
            this.m_frmMain = p_frmMain;
			//int x=0;
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_intHt = this.grpboxOpen.Top + this.grpboxOpen.Height + 50; 
			this.Height = this.m_intHt;
			this.Width = this.m_intWd;
			this.grpboxOpen.Visible=true;
			this.grpboxSelectExisting.Visible=false;
			this.grpboxNewSelectScenarios.Visible=false;
			this.grpboxNewSelectView.Visible=false;
			this.grpboxSelectExisting.Location = this.grpboxOpen.Location;
			this.grpboxSelectExisting.Size = this.grpboxOpen.Size;
			this.grpboxNewSelectScenarios.Location = this.grpboxOpen.Location;
			this.grpboxNewSelectScenarios.Size = this.grpboxOpen.Size;
            this.grpboxNewSelectView.Location=this.grpboxOpen.Location;
			this.grpboxNewSelectView.Size = this.grpboxOpen.Size;
			this.m_frmTherm = new frmTherm();
			this.m_frmTherm.btnCancel.Click += new System.EventHandler(this.ThermCancel);

			this.GetScenarioMergeRecords();
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_scenario_merge_tables));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.grpboxNewSelectView = new System.Windows.Forms.GroupBox();
			this.btnNewSelectViewCancel = new System.Windows.Forms.Button();
			this.btnNewSelectViewPrev = new System.Windows.Forms.Button();
			this.grpboxNewSelectView2 = new System.Windows.Forms.GroupBox();
			this.btnViewSelectViewInGrid = new System.Windows.Forms.Button();
			this.btnNewSelectViewStartAccess = new System.Windows.Forms.Button();
			this.btnNewSelectViewGetMDBFile = new System.Windows.Forms.Button();
			this.chkboxNewSelectViewSaveToAccess = new System.Windows.Forms.CheckBox();
			this.txtNewSelectViewMergeMDBFile = new System.Windows.Forms.TextBox();
			this.grpboxNewSelectScenarios = new System.Windows.Forms.GroupBox();
			this.btnNewSelectScenariosCancel = new System.Windows.Forms.Button();
			this.btnNewSelectScenariosNext = new System.Windows.Forms.Button();
			this.btnNewSelectScenariosPrev = new System.Windows.Forms.Button();
			this.grpboxNewSelectScenarios2 = new System.Windows.Forms.GroupBox();
			this.btnSelectAllNewTables = new System.Windows.Forms.Button();
			this.btnClearAllNewTables = new System.Windows.Forms.Button();
			this.lstNewSelectTables = new System.Windows.Forms.CheckedListBox();
			this.lstSelectScenarios = new System.Windows.Forms.CheckedListBox();
			this.btnSelectAllScenarios = new System.Windows.Forms.Button();
			this.btnClearAllScenarios = new System.Windows.Forms.Button();
			this.grpboxSelectExisting = new System.Windows.Forms.GroupBox();
			this.grpboxSelectExisting2 = new System.Windows.Forms.GroupBox();
			this.btnSelectExistingViewTables = new System.Windows.Forms.Button();
			this.btnSelectExistingViewTablesInAccess = new System.Windows.Forms.Button();
			this.btnClearAllExistingTables = new System.Windows.Forms.Button();
			this.btnSelectAllExistingTables = new System.Windows.Forms.Button();
			this.lstSelectExistingTables = new System.Windows.Forms.CheckedListBox();
			this.cmbMDBFiles = new System.Windows.Forms.ComboBox();
			this.btnSelectExisting = new System.Windows.Forms.Button();
			this.btnSelectExistingPrevious = new System.Windows.Forms.Button();
			this.btnSelectExistingCancel = new System.Windows.Forms.Button();
			this.lblTitle = new System.Windows.Forms.Label();
			this.btnHelp = new System.Windows.Forms.Button();
			this.grpboxOpen = new System.Windows.Forms.GroupBox();
			this.grpboxOpen2 = new System.Windows.Forms.GroupBox();
			this.rdoExistingMerge = new System.Windows.Forms.RadioButton();
			this.rdoNewMerge = new System.Windows.Forms.RadioButton();
			this.btnOpenNext = new System.Windows.Forms.Button();
			this.btnOpenCancel = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.grpboxNewSelectView.SuspendLayout();
			this.grpboxNewSelectView2.SuspendLayout();
			this.grpboxNewSelectScenarios.SuspendLayout();
			this.grpboxNewSelectScenarios2.SuspendLayout();
			this.grpboxSelectExisting.SuspendLayout();
			this.grpboxSelectExisting2.SuspendLayout();
			this.grpboxOpen.SuspendLayout();
			this.grpboxOpen2.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.grpboxNewSelectView);
			this.groupBox1.Controls.Add(this.grpboxNewSelectScenarios);
			this.groupBox1.Controls.Add(this.grpboxSelectExisting);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Controls.Add(this.btnHelp);
			this.groupBox1.Controls.Add(this.grpboxOpen);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(664, 1352);
			this.groupBox1.TabIndex = 29;
			this.groupBox1.TabStop = false;
			// 
			// grpboxNewSelectView
			// 
			this.grpboxNewSelectView.Controls.Add(this.btnNewSelectViewCancel);
			this.grpboxNewSelectView.Controls.Add(this.btnNewSelectViewPrev);
			this.grpboxNewSelectView.Controls.Add(this.grpboxNewSelectView2);
			this.grpboxNewSelectView.Location = new System.Drawing.Point(24, 1024);
			this.grpboxNewSelectView.Name = "grpboxNewSelectView";
			this.grpboxNewSelectView.Size = new System.Drawing.Size(616, 304);
			this.grpboxNewSelectView.TabIndex = 30;
			this.grpboxNewSelectView.TabStop = false;
			this.grpboxNewSelectView.Text = "Step 3 - Select Save And View Options";
			// 
			// btnNewSelectViewCancel
			// 
			this.btnNewSelectViewCancel.Location = new System.Drawing.Point(360, 264);
			this.btnNewSelectViewCancel.Name = "btnNewSelectViewCancel";
			this.btnNewSelectViewCancel.Size = new System.Drawing.Size(64, 24);
			this.btnNewSelectViewCancel.TabIndex = 40;
			this.btnNewSelectViewCancel.Text = "Cancel";
			this.btnNewSelectViewCancel.Click += new System.EventHandler(this.btnNewSelectViewCancel_Click);
			// 
			// btnNewSelectViewPrev
			// 
			this.btnNewSelectViewPrev.Location = new System.Drawing.Point(456, 264);
			this.btnNewSelectViewPrev.Name = "btnNewSelectViewPrev";
			this.btnNewSelectViewPrev.Size = new System.Drawing.Size(72, 24);
			this.btnNewSelectViewPrev.TabIndex = 39;
			this.btnNewSelectViewPrev.Text = "< Previous";
			this.btnNewSelectViewPrev.Click += new System.EventHandler(this.btnNewSelectViewPrev_Click);
			// 
			// grpboxNewSelectView2
			// 
			this.grpboxNewSelectView2.Controls.Add(this.btnViewSelectViewInGrid);
			this.grpboxNewSelectView2.Controls.Add(this.btnNewSelectViewStartAccess);
			this.grpboxNewSelectView2.Controls.Add(this.btnNewSelectViewGetMDBFile);
			this.grpboxNewSelectView2.Controls.Add(this.chkboxNewSelectViewSaveToAccess);
			this.grpboxNewSelectView2.Controls.Add(this.txtNewSelectViewMergeMDBFile);
			this.grpboxNewSelectView2.Location = new System.Drawing.Point(24, 24);
			this.grpboxNewSelectView2.Name = "grpboxNewSelectView2";
			this.grpboxNewSelectView2.Size = new System.Drawing.Size(560, 232);
			this.grpboxNewSelectView2.TabIndex = 0;
			this.grpboxNewSelectView2.TabStop = false;
			// 
			// btnViewSelectViewInGrid
			// 
			this.btnViewSelectViewInGrid.Location = new System.Drawing.Point(298, 120);
			this.btnViewSelectViewInGrid.Name = "btnViewSelectViewInGrid";
			this.btnViewSelectViewInGrid.Size = new System.Drawing.Size(128, 88);
			this.btnViewSelectViewInGrid.TabIndex = 31;
			this.btnViewSelectViewInGrid.Text = "View Tables";
			this.btnViewSelectViewInGrid.Click += new System.EventHandler(this.btnViewSelectViewInGrid_Click);
			// 
			// btnNewSelectViewStartAccess
			// 
			this.btnNewSelectViewStartAccess.Location = new System.Drawing.Point(120, 120);
			this.btnNewSelectViewStartAccess.Name = "btnNewSelectViewStartAccess";
			this.btnNewSelectViewStartAccess.Size = new System.Drawing.Size(128, 88);
			this.btnNewSelectViewStartAccess.TabIndex = 30;
			this.btnNewSelectViewStartAccess.Text = "Microsoft Access";
			this.btnNewSelectViewStartAccess.Click += new System.EventHandler(this.btnNewSelectViewStartAccess_Click);
			// 
			// btnNewSelectViewGetMDBFile
			// 
			this.btnNewSelectViewGetMDBFile.Image = ((System.Drawing.Image)(resources.GetObject("btnNewSelectViewGetMDBFile.Image")));
			this.btnNewSelectViewGetMDBFile.Location = new System.Drawing.Point(521, 61);
			this.btnNewSelectViewGetMDBFile.Name = "btnNewSelectViewGetMDBFile";
			this.btnNewSelectViewGetMDBFile.Size = new System.Drawing.Size(32, 32);
			this.btnNewSelectViewGetMDBFile.TabIndex = 29;
			this.btnNewSelectViewGetMDBFile.Click += new System.EventHandler(this.btnNewSelectViewGetMDBFile_Click);
			// 
			// chkboxNewSelectViewSaveToAccess
			// 
			this.chkboxNewSelectViewSaveToAccess.Location = new System.Drawing.Point(24, 32);
			this.chkboxNewSelectViewSaveToAccess.Name = "chkboxNewSelectViewSaveToAccess";
			this.chkboxNewSelectViewSaveToAccess.Size = new System.Drawing.Size(120, 16);
			this.chkboxNewSelectViewSaveToAccess.TabIndex = 1;
			this.chkboxNewSelectViewSaveToAccess.Text = "Save In Access";
			this.chkboxNewSelectViewSaveToAccess.Click += new System.EventHandler(this.chkboxNewSelectViewSaveToAccess_Click);
			// 
			// txtNewSelectViewMergeMDBFile
			// 
			this.txtNewSelectViewMergeMDBFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNewSelectViewMergeMDBFile.Location = new System.Drawing.Point(8, 64);
			this.txtNewSelectViewMergeMDBFile.Name = "txtNewSelectViewMergeMDBFile";
			this.txtNewSelectViewMergeMDBFile.Size = new System.Drawing.Size(504, 26);
			this.txtNewSelectViewMergeMDBFile.TabIndex = 0;
			this.txtNewSelectViewMergeMDBFile.Text = "";
			this.txtNewSelectViewMergeMDBFile.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNewSelectViewMergeMDBFile_KeyPress);
			// 
			// grpboxNewSelectScenarios
			// 
			this.grpboxNewSelectScenarios.Controls.Add(this.btnNewSelectScenariosCancel);
			this.grpboxNewSelectScenarios.Controls.Add(this.btnNewSelectScenariosNext);
			this.grpboxNewSelectScenarios.Controls.Add(this.btnNewSelectScenariosPrev);
			this.grpboxNewSelectScenarios.Controls.Add(this.grpboxNewSelectScenarios2);
			this.grpboxNewSelectScenarios.Location = new System.Drawing.Point(24, 688);
			this.grpboxNewSelectScenarios.Name = "grpboxNewSelectScenarios";
			this.grpboxNewSelectScenarios.Size = new System.Drawing.Size(616, 304);
			this.grpboxNewSelectScenarios.TabIndex = 29;
			this.grpboxNewSelectScenarios.TabStop = false;
			this.grpboxNewSelectScenarios.Text = "Step 2 - Select Scenarios And Table(s)";
			// 
			// btnNewSelectScenariosCancel
			// 
			this.btnNewSelectScenariosCancel.Location = new System.Drawing.Point(360, 264);
			this.btnNewSelectScenariosCancel.Name = "btnNewSelectScenariosCancel";
			this.btnNewSelectScenariosCancel.Size = new System.Drawing.Size(64, 24);
			this.btnNewSelectScenariosCancel.TabIndex = 38;
			this.btnNewSelectScenariosCancel.Text = "Cancel";
			this.btnNewSelectScenariosCancel.Click += new System.EventHandler(this.btnNewSelectScenariosCancel_Click);
			// 
			// btnNewSelectScenariosNext
			// 
			this.btnNewSelectScenariosNext.Location = new System.Drawing.Point(528, 264);
			this.btnNewSelectScenariosNext.Name = "btnNewSelectScenariosNext";
			this.btnNewSelectScenariosNext.Size = new System.Drawing.Size(72, 24);
			this.btnNewSelectScenariosNext.TabIndex = 37;
			this.btnNewSelectScenariosNext.Text = "Next >";
			this.btnNewSelectScenariosNext.Click += new System.EventHandler(this.btnNewSelectScenariosNext_Click);
			// 
			// btnNewSelectScenariosPrev
			// 
			this.btnNewSelectScenariosPrev.Location = new System.Drawing.Point(456, 264);
			this.btnNewSelectScenariosPrev.Name = "btnNewSelectScenariosPrev";
			this.btnNewSelectScenariosPrev.Size = new System.Drawing.Size(72, 24);
			this.btnNewSelectScenariosPrev.TabIndex = 36;
			this.btnNewSelectScenariosPrev.Text = "< Previous";
			this.btnNewSelectScenariosPrev.Click += new System.EventHandler(this.btnNewSelectScenariosPrev_Click);
			// 
			// grpboxNewSelectScenarios2
			// 
			this.grpboxNewSelectScenarios2.Controls.Add(this.btnSelectAllNewTables);
			this.grpboxNewSelectScenarios2.Controls.Add(this.btnClearAllNewTables);
			this.grpboxNewSelectScenarios2.Controls.Add(this.lstNewSelectTables);
			this.grpboxNewSelectScenarios2.Controls.Add(this.lstSelectScenarios);
			this.grpboxNewSelectScenarios2.Controls.Add(this.btnSelectAllScenarios);
			this.grpboxNewSelectScenarios2.Controls.Add(this.btnClearAllScenarios);
			this.grpboxNewSelectScenarios2.Location = new System.Drawing.Point(24, 24);
			this.grpboxNewSelectScenarios2.Name = "grpboxNewSelectScenarios2";
			this.grpboxNewSelectScenarios2.Size = new System.Drawing.Size(560, 232);
			this.grpboxNewSelectScenarios2.TabIndex = 35;
			this.grpboxNewSelectScenarios2.TabStop = false;
			// 
			// btnSelectAllNewTables
			// 
			this.btnSelectAllNewTables.Location = new System.Drawing.Point(344, 200);
			this.btnSelectAllNewTables.Name = "btnSelectAllNewTables";
			this.btnSelectAllNewTables.Size = new System.Drawing.Size(64, 24);
			this.btnSelectAllNewTables.TabIndex = 36;
			this.btnSelectAllNewTables.Text = "Select All";
			this.btnSelectAllNewTables.Click += new System.EventHandler(this.btnSelectAllNewTables_Click);
			// 
			// btnClearAllNewTables
			// 
			this.btnClearAllNewTables.Location = new System.Drawing.Point(416, 200);
			this.btnClearAllNewTables.Name = "btnClearAllNewTables";
			this.btnClearAllNewTables.Size = new System.Drawing.Size(64, 24);
			this.btnClearAllNewTables.TabIndex = 37;
			this.btnClearAllNewTables.Text = "Clear All";
			this.btnClearAllNewTables.Click += new System.EventHandler(this.btnClearAllNewTables_Click);
			// 
			// lstNewSelectTables
			// 
			this.lstNewSelectTables.CheckOnClick = true;
			this.lstNewSelectTables.Location = new System.Drawing.Point(286, 24);
			this.lstNewSelectTables.Name = "lstNewSelectTables";
			this.lstNewSelectTables.Size = new System.Drawing.Size(264, 169);
			this.lstNewSelectTables.TabIndex = 35;
			// 
			// lstSelectScenarios
			// 
			this.lstSelectScenarios.CheckOnClick = true;
			this.lstSelectScenarios.Location = new System.Drawing.Point(8, 24);
			this.lstSelectScenarios.Name = "lstSelectScenarios";
			this.lstSelectScenarios.Size = new System.Drawing.Size(264, 169);
			this.lstSelectScenarios.TabIndex = 32;
			// 
			// btnSelectAllScenarios
			// 
			this.btnSelectAllScenarios.Location = new System.Drawing.Point(65, 200);
			this.btnSelectAllScenarios.Name = "btnSelectAllScenarios";
			this.btnSelectAllScenarios.Size = new System.Drawing.Size(64, 24);
			this.btnSelectAllScenarios.TabIndex = 33;
			this.btnSelectAllScenarios.Text = "Select All";
			this.btnSelectAllScenarios.Click += new System.EventHandler(this.btnSelectAllScenarios_Click);
			// 
			// btnClearAllScenarios
			// 
			this.btnClearAllScenarios.Location = new System.Drawing.Point(137, 200);
			this.btnClearAllScenarios.Name = "btnClearAllScenarios";
			this.btnClearAllScenarios.Size = new System.Drawing.Size(64, 24);
			this.btnClearAllScenarios.TabIndex = 34;
			this.btnClearAllScenarios.Text = "Clear All";
			this.btnClearAllScenarios.Click += new System.EventHandler(this.btnClearAllScenarios_Click);
			// 
			// grpboxSelectExisting
			// 
			this.grpboxSelectExisting.Controls.Add(this.grpboxSelectExisting2);
			this.grpboxSelectExisting.Controls.Add(this.btnSelectExistingPrevious);
			this.grpboxSelectExisting.Controls.Add(this.btnSelectExistingCancel);
			this.grpboxSelectExisting.Location = new System.Drawing.Point(24, 368);
			this.grpboxSelectExisting.Name = "grpboxSelectExisting";
			this.grpboxSelectExisting.Size = new System.Drawing.Size(616, 304);
			this.grpboxSelectExisting.TabIndex = 28;
			this.grpboxSelectExisting.TabStop = false;
			this.grpboxSelectExisting.Text = "Step 2 - Select Database Container  And Tables";
			// 
			// grpboxSelectExisting2
			// 
			this.grpboxSelectExisting2.Controls.Add(this.btnSelectExistingViewTables);
			this.grpboxSelectExisting2.Controls.Add(this.btnSelectExistingViewTablesInAccess);
			this.grpboxSelectExisting2.Controls.Add(this.btnClearAllExistingTables);
			this.grpboxSelectExisting2.Controls.Add(this.btnSelectAllExistingTables);
			this.grpboxSelectExisting2.Controls.Add(this.lstSelectExistingTables);
			this.grpboxSelectExisting2.Controls.Add(this.cmbMDBFiles);
			this.grpboxSelectExisting2.Controls.Add(this.btnSelectExisting);
			this.grpboxSelectExisting2.Location = new System.Drawing.Point(32, 24);
			this.grpboxSelectExisting2.Name = "grpboxSelectExisting2";
			this.grpboxSelectExisting2.Size = new System.Drawing.Size(560, 232);
			this.grpboxSelectExisting2.TabIndex = 29;
			this.grpboxSelectExisting2.TabStop = false;
			// 
			// btnSelectExistingViewTables
			// 
			this.btnSelectExistingViewTables.Location = new System.Drawing.Point(297, 144);
			this.btnSelectExistingViewTables.Name = "btnSelectExistingViewTables";
			this.btnSelectExistingViewTables.Size = new System.Drawing.Size(255, 40);
			this.btnSelectExistingViewTables.TabIndex = 33;
			this.btnSelectExistingViewTables.Text = "View Selected Tables";
			this.btnSelectExistingViewTables.Click += new System.EventHandler(this.btnSelectExistingViewTables_Click);
			// 
			// btnSelectExistingViewTablesInAccess
			// 
			this.btnSelectExistingViewTablesInAccess.Location = new System.Drawing.Point(297, 89);
			this.btnSelectExistingViewTablesInAccess.Name = "btnSelectExistingViewTablesInAccess";
			this.btnSelectExistingViewTablesInAccess.Size = new System.Drawing.Size(255, 40);
			this.btnSelectExistingViewTablesInAccess.TabIndex = 32;
			this.btnSelectExistingViewTablesInAccess.Text = "View Tables Using Microsoft Access";
			this.btnSelectExistingViewTablesInAccess.Click += new System.EventHandler(this.btnSelectExistingViewTablesInAccess_Click);
			// 
			// btnClearAllExistingTables
			// 
			this.btnClearAllExistingTables.Location = new System.Drawing.Point(96, 200);
			this.btnClearAllExistingTables.Name = "btnClearAllExistingTables";
			this.btnClearAllExistingTables.Size = new System.Drawing.Size(64, 24);
			this.btnClearAllExistingTables.TabIndex = 31;
			this.btnClearAllExistingTables.Text = "Clear All";
			this.btnClearAllExistingTables.Click += new System.EventHandler(this.btnClearAllExistingTables_Click);
			// 
			// btnSelectAllExistingTables
			// 
			this.btnSelectAllExistingTables.Location = new System.Drawing.Point(24, 200);
			this.btnSelectAllExistingTables.Name = "btnSelectAllExistingTables";
			this.btnSelectAllExistingTables.Size = new System.Drawing.Size(64, 24);
			this.btnSelectAllExistingTables.TabIndex = 30;
			this.btnSelectAllExistingTables.Text = "Select All";
			this.btnSelectAllExistingTables.Click += new System.EventHandler(this.btnSelectAllExistingTables_Click);
			// 
			// lstSelectExistingTables
			// 
			this.lstSelectExistingTables.CheckOnClick = true;
			this.lstSelectExistingTables.Location = new System.Drawing.Point(24, 72);
			this.lstSelectExistingTables.Name = "lstSelectExistingTables";
			this.lstSelectExistingTables.Size = new System.Drawing.Size(264, 124);
			this.lstSelectExistingTables.TabIndex = 29;
			// 
			// cmbMDBFiles
			// 
			this.cmbMDBFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMDBFiles.Location = new System.Drawing.Point(16, 24);
			this.cmbMDBFiles.Name = "cmbMDBFiles";
			this.cmbMDBFiles.Size = new System.Drawing.Size(496, 28);
			this.cmbMDBFiles.TabIndex = 27;
			this.cmbMDBFiles.SelectedValueChanged += new System.EventHandler(this.cmbMDBFiles_SelectedValueChanged);
			// 
			// btnSelectExisting
			// 
			this.btnSelectExisting.Image = ((System.Drawing.Image)(resources.GetObject("btnSelectExisting.Image")));
			this.btnSelectExisting.Location = new System.Drawing.Point(519, 22);
			this.btnSelectExisting.Name = "btnSelectExisting";
			this.btnSelectExisting.Size = new System.Drawing.Size(32, 32);
			this.btnSelectExisting.TabIndex = 28;
			this.btnSelectExisting.Click += new System.EventHandler(this.btnSelectExisting_Click);
			// 
			// btnSelectExistingPrevious
			// 
			this.btnSelectExistingPrevious.Location = new System.Drawing.Point(456, 264);
			this.btnSelectExistingPrevious.Name = "btnSelectExistingPrevious";
			this.btnSelectExistingPrevious.Size = new System.Drawing.Size(72, 24);
			this.btnSelectExistingPrevious.TabIndex = 23;
			this.btnSelectExistingPrevious.Text = "< Previous";
			this.btnSelectExistingPrevious.Click += new System.EventHandler(this.btnSelectExistingPrevious_Click);
			// 
			// btnSelectExistingCancel
			// 
			this.btnSelectExistingCancel.Location = new System.Drawing.Point(360, 264);
			this.btnSelectExistingCancel.Name = "btnSelectExistingCancel";
			this.btnSelectExistingCancel.Size = new System.Drawing.Size(64, 24);
			this.btnSelectExistingCancel.TabIndex = 20;
			this.btnSelectExistingCancel.Text = "Cancel";
			this.btnSelectExistingCancel.Click += new System.EventHandler(this.btnSelectExistingCancel_Click);
			// 
			// lblTitle
			// 
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(8, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(336, 24);
			this.lblTitle.TabIndex = 27;
			this.lblTitle.Text = "Join Data From Multiple Scenarios";
			// 
			// btnHelp
			// 
			this.btnHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image")));
			this.btnHelp.Location = new System.Drawing.Point(624, 11);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(32, 34);
			this.btnHelp.TabIndex = 1;
			// 
			// grpboxOpen
			// 
			this.grpboxOpen.Controls.Add(this.grpboxOpen2);
			this.grpboxOpen.Controls.Add(this.btnOpenNext);
			this.grpboxOpen.Controls.Add(this.btnOpenCancel);
			this.grpboxOpen.Location = new System.Drawing.Point(24, 49);
			this.grpboxOpen.Name = "grpboxOpen";
			this.grpboxOpen.Size = new System.Drawing.Size(616, 304);
			this.grpboxOpen.TabIndex = 0;
			this.grpboxOpen.TabStop = false;
			this.grpboxOpen.Text = "Step 1 - Open Existing Tables Or Create New Tables";
			// 
			// grpboxOpen2
			// 
			this.grpboxOpen2.Controls.Add(this.rdoExistingMerge);
			this.grpboxOpen2.Controls.Add(this.rdoNewMerge);
			this.grpboxOpen2.Location = new System.Drawing.Point(16, 104);
			this.grpboxOpen2.Name = "grpboxOpen2";
			this.grpboxOpen2.Size = new System.Drawing.Size(584, 64);
			this.grpboxOpen2.TabIndex = 26;
			this.grpboxOpen2.TabStop = false;
			// 
			// rdoExistingMerge
			// 
			this.rdoExistingMerge.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rdoExistingMerge.Location = new System.Drawing.Point(16, 16);
			this.rdoExistingMerge.Name = "rdoExistingMerge";
			this.rdoExistingMerge.Size = new System.Drawing.Size(256, 32);
			this.rdoExistingMerge.TabIndex = 24;
			this.rdoExistingMerge.Text = "Open Existing Tables";
			// 
			// rdoNewMerge
			// 
			this.rdoNewMerge.Checked = true;
			this.rdoNewMerge.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rdoNewMerge.Location = new System.Drawing.Point(288, 16);
			this.rdoNewMerge.Name = "rdoNewMerge";
			this.rdoNewMerge.Size = new System.Drawing.Size(280, 32);
			this.rdoNewMerge.TabIndex = 25;
			this.rdoNewMerge.TabStop = true;
			this.rdoNewMerge.Text = "Create New Join Tables";
			// 
			// btnOpenNext
			// 
			this.btnOpenNext.Location = new System.Drawing.Point(528, 264);
			this.btnOpenNext.Name = "btnOpenNext";
			this.btnOpenNext.Size = new System.Drawing.Size(72, 24);
			this.btnOpenNext.TabIndex = 21;
			this.btnOpenNext.Text = "Next >";
			this.btnOpenNext.Click += new System.EventHandler(this.btnOpenNext_Click);
			// 
			// btnOpenCancel
			// 
			this.btnOpenCancel.Location = new System.Drawing.Point(360, 264);
			this.btnOpenCancel.Name = "btnOpenCancel";
			this.btnOpenCancel.Size = new System.Drawing.Size(64, 24);
			this.btnOpenCancel.TabIndex = 20;
			this.btnOpenCancel.Text = "Cancel";
			this.btnOpenCancel.Click += new System.EventHandler(this.btnOpenCancel_Click);
			// 
			// uc_scenario_merge_tables
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_merge_tables";
			this.Size = new System.Drawing.Size(664, 1352);
			this.groupBox1.ResumeLayout(false);
			this.grpboxNewSelectView.ResumeLayout(false);
			this.grpboxNewSelectView2.ResumeLayout(false);
			this.grpboxNewSelectScenarios.ResumeLayout(false);
			this.grpboxNewSelectScenarios2.ResumeLayout(false);
			this.grpboxSelectExisting.ResumeLayout(false);
			this.grpboxSelectExisting2.ResumeLayout(false);
			this.grpboxOpen.ResumeLayout(false);
			this.grpboxOpen2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnSelectExistingViewTables_Click(object sender, System.EventArgs e)
		{
			

			if (this.cmbMDBFiles.Text.Trim().Length > 0)
			{
				if (System.IO.File.Exists(this.cmbMDBFiles.Text.Trim()))
				{
					try
					{
						int x=0;
						int intCheckedCount=0;
						ado_data_access p_ado = new ado_data_access();
						FIA_Biosum_Manager.frmTherm p_frmTherm;
						p_frmTherm = new frmTherm();
						string strConn = p_ado.getMDBConnString(this.cmbMDBFiles.Text.Trim(),"","");
						string strSQL = "";
						for (x=0;x<=this.lstSelectExistingTables.Items.Count-1;x++)
						{
							if (this.lstSelectExistingTables.GetItemChecked(x) == true)
							{
								if (intCheckedCount == 0)
								{
									//get the connection string

									//start the gridview
									this.m_frmGridView = new frmGridView();
									this.m_frmGridView.Text = "Core Analysis: View Joined Scenario Data";

									//start the thermometer
									p_frmTherm.Text = "Create Data Views";
									p_frmTherm.lblMsg.Text="";
									p_frmTherm.lblMsg.Visible=true;
									p_frmTherm.btnCancel.Visible=false;
									p_frmTherm.progressBar1.Minimum=0;
									p_frmTherm.progressBar1.Maximum = this.lstSelectExistingTables.Items.Count - 1;
									p_frmTherm.Show();
									p_frmTherm.Focus();
								}
								intCheckedCount++;
								p_frmTherm.lblMsg.Text = this.lstSelectExistingTables.Items[x].ToString();
								p_frmTherm.lblMsg.Refresh();
								strSQL = "select * from " + this.lstSelectExistingTables.Items[x].ToString();;
								this.m_frmGridView.LoadDataSet(strConn,strSQL,this.lstSelectExistingTables.Items[x].ToString());
								p_frmTherm.progressBar1.Value = x;

							}

						}
						if (intCheckedCount > 0)
						{
							p_frmTherm.Close();
							if (intCheckedCount > 1) this.m_frmGridView.TileGridViews();
							this.m_frmGridView.Show();
							this.m_frmGridView.Focus();
						}
						p_frmTherm = null;
						p_ado = null;
						this.SaveComboBoxItems(this.cmbMDBFiles);
						this.GetScenarioMergeRecords();
					}
					catch (Exception caught)
					{
						MessageBox.Show(caught.Message);
					}
				}
				else
				{
					MessageBox.Show("File " +  this.cmbMDBFiles.Text.Trim() + " Not Found!!");
				}
			}
			else
			{
				MessageBox.Show("Select An Access File To Open");
			}
		}

		private void btnOpenNext_Click(object sender, System.EventArgs e)
		{
			if (this.rdoExistingMerge.Checked ==true)
			{
				this.grpboxOpen.Visible=false;
				this.grpboxSelectExisting.Visible=true;
			}
			else
			{
				if (this.lstSelectScenarios.Items.Count == 0)
				{
					this.PopulateScenarioList();
				}
                this.grpboxOpen.Visible=false;
				this.grpboxNewSelectScenarios.Visible=true;
			}
		}

		private void btnSelectAllExistingTables_Click(object sender, System.EventArgs e)
		{
			for (int x = 0; x<= this.lstSelectExistingTables.Items.Count-1;x++)
			{
                  this.lstSelectExistingTables.SetItemChecked(x,true);
			}
		}

		private void btnClearAllExistingTables_Click(object sender, System.EventArgs e)
		{
			for (int x = 0; x<= this.lstSelectExistingTables.Items.Count-1;x++)
			{
				this.lstSelectExistingTables.SetItemChecked(x,false);
			}

		}

		private void btnSelectExistingCancel_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();
		}

		private void btnOpenCancel_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show(((frmDialog)this.ParentForm).Height.ToString());
			this.ParentForm.Close();
		}

		private void btnSelectExistingViewTablesInAccess_Click(object sender, System.EventArgs e)
		{
			if (this.cmbMDBFiles.Text.Trim().Length > 0)
			{
				if (System.IO.File.Exists(this.cmbMDBFiles.Text.Trim()))
				{
					FIA_Biosum_Manager.utils p_oUtils = new utils();
					p_oUtils.ShellExecute(this.cmbMDBFiles.Text);
					p_oUtils = null;
					this.SaveComboBoxItems(this.cmbMDBFiles);
					this.GetScenarioMergeRecords();
				}
				else
				{
					MessageBox.Show("File " +  this.cmbMDBFiles.Text.Trim() + " Not Found!!");
				}
			}
			else
			{
				MessageBox.Show("Select An Access File To Open");
			}
		}

		private void btnSelectExistingPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxOpen.Visible=true;
			this.grpboxSelectExisting.Visible=false;
		}

		private void btnSelectExisting_Click(object sender, System.EventArgs e)
		{
			
			string strMDBPathAndFile="";
			string strDir="";
			string strFile="";


			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Designate Existing MDB File";
			OpenFileDialog1.Filter = "Access Database File (*.MDB) |*.mdb";
			FIA_Biosum_Manager.utils p_oUtils = new utils();

			if (this.cmbMDBFiles.Text.Trim().Length > 0)
			{
				strFile = p_oUtils.getFileName(this.cmbMDBFiles.Text.Trim());
				strDir = p_oUtils.getDirectory(this.cmbMDBFiles.Text.Trim());
			}
			else
			{
				strDir = this.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\core\\db";
			}

			OpenFileDialog1.InitialDirectory = strDir;
			OpenFileDialog1.CheckFileExists = true;
			DialogResult result =  OpenFileDialog1.ShowDialog();

			if (result == DialogResult.OK)
			{
				strMDBPathAndFile = OpenFileDialog1.FileName.Trim();
				if (strMDBPathAndFile.Length > 0) 
				{
					//see if the mdb file exists
					if (System.IO.File.Exists(strMDBPathAndFile) == true) 
					{
					    this.PopulateCheckedComboWithMDBTables(this.lstSelectExistingTables,strMDBPathAndFile,true);
						if (this.m_intError==0) 
						{
							this.cmbMDBFiles.Text = strMDBPathAndFile;
						}
					}
				}
			}
			p_oUtils=null;
			OpenFileDialog1 = null;
			

		}
		/// <summary>
		/// populate the checked list box with the table names found in the MDB file
		/// </summary>
		/// <param name="p_lst">listbox object</param>
		/// <param name="strMDBPathAndFile">The access file</param>
		/// <param name="bChecked"></param>
		private void PopulateCheckedComboWithMDBTables(System.Windows.Forms.CheckedListBox p_lst, string strMDBPathAndFile,bool bChecked)
		{
			int x;
            string[] strTableNames;

			this.m_intError=0;
			p_lst.Items.Clear();
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			try
			{  
				if (p_dao.m_intError==0)
				{
					for (x=0;x<=strTableNames.Length-1;x++)
					{
						if (strTableNames[x] != null)
						{
							p_lst.Items.Add(strTableNames[x]);
							p_lst.SetItemChecked(x, bChecked);
						}
					}
							
				}
				else
				{
					this.m_intError=p_dao.m_intError;
				}
			}
			catch
			{
				this.m_intError=-1;
			}

			p_dao = null;


		}

		/// <summary>
		/// populate the combobox with the records in scenario_merge table
		/// </summary>
		/// <param name="p_cmb">combobox object</param>
		/// <param name="p_datareader">OleDbDataReader object containing column to be loaded into combobox</param>
		/// <param name="strCol">load column name records into combobox </param>
		private void PopulateComboBox(System.Windows.Forms.ComboBox p_cmb,System.Data.OleDb.OleDbDataReader p_datareader,string strCol)
		{
			p_cmb.Items.Clear();
			if (p_datareader.HasRows)
			{
				while (p_datareader.Read())
				{
					if (p_datareader[strCol].ToString().Trim().Length > 0)
					{
						p_cmb.Items.Add(p_datareader[strCol].ToString().Trim());
					}
				}
				if (this.cmbMDBFiles.Items.Count > 0)
				{
					p_cmb.Text = p_cmb.Items[0].ToString().Trim();
				}
						
			}
		}
		/// <summary>
		/// save the currently selected combo box item to scenario_merge and reorder the combo list items in the table
		/// </summary>
		/// <param name="p_cmb">ComboBox object</param>
		
		private void SaveComboBoxItems(System.Windows.Forms.ComboBox p_cmb)
		{
			int x;
			int intCurPos=0;
			System.Data.DataRow[] p_rows;
			System.Data.DataRow p_row;
            ado_data_access p_ado = new ado_data_access();
			/*************************************************
			 **open scenario_core_rule_definitions
			 *************************************************/
            p_ado.OpenConnection(this.m_strScenarioConn);
			if (p_ado.m_intError == 0)
			{
				/*************************************************
				 **see if any records in the scenario_merge table
				 *************************************************/
				if (p_ado.getRecordCount(p_ado.m_OleDbConnection,"select count(*) from scenario_merge", "scenario_merge") < 1)
				{
					/*********************************************************
					 **only one item so save the currently selected item
					 *********************************************************/
                    p_ado.m_strSQL = "INSERT INTO scenario_merge (cmborder, mdbpathandfile) " + 
					                "VALUES (1,'" + p_cmb.Text.Trim().ToLower() + "');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,p_ado.m_strSQL);
				}
				else
				{
					/*************************************************************
					 **load up all the records into a dataset
					 *************************************************************/
					p_ado.CreateDataSet(p_ado.m_OleDbConnection,"select * from scenario_merge","scenario_merge");
					p_rows = p_ado.m_DataSet.Tables["scenario_merge"].Select("mdbpathandfile = '" + p_cmb.Text.Trim().ToLower() + "'");
					/*************************************************************
					 **see if the currently selected item is already in the table
					 *************************************************************/
					if (p_rows.Length != 0 )
					{
						/************************************************************
						 **okay it is already in the table so lets get its current
						 **order number and reassign it a number of 1
						 ************************************************************/
						intCurPos = Convert.ToInt32(p_rows[0]["cmborder"].ToString());
						p_rows[0]["cmborder"] = 1;
						p_ado.m_strSQL = "UPDATE scenario_merge SET cmborder = 1 " + 
							"WHERE trim(lcase(mdbpathandfile)) = '" + p_cmb.Text.Trim().ToLower() + "';";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,p_ado.m_strSQL);
					}
					else
					{
						intCurPos = 6;
						p_row = p_ado.m_DataSet.Tables["scenario_merge"].NewRow();
                        p_row["cmborder"] = 1;
						p_row["mdbpathandfile"] = this.cmbMDBFiles.Text.Trim();
						p_ado.m_strSQL = "INSERT INTO scenario_merge (cmborder,mdbpathandfile) " + 
							"VALUES (" + p_row["cmborder"].ToString() + ",'" + p_cmb.Text.Trim().ToLower() + "');";
                        p_ado.SqlNonQuery(p_ado.m_OleDbConnection,p_ado.m_strSQL);
					    p_ado.m_strSQL = "delete from scenario_merge where cmborder = 5";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,p_ado.m_strSQL);

					}
					for (x=0; x<=p_ado.m_DataSet.Tables["scenario_merge"].Rows.Count-1; x++)
					{
						p_row = p_ado.m_DataSet.Tables["scenario_merge"].Rows[x];
						if (Convert.ToInt32(p_row["cmborder"].ToString()) != 5 && p_row["mdbpathandfile"].ToString().Trim().ToUpper() != p_cmb.Text.Trim().ToUpper())
						{
							if (Convert.ToInt32(p_row["cmborder"].ToString()) < intCurPos)
							{
								p_row["cmborder"] = Convert.ToInt32(p_row["cmborder"].ToString()) + 1;
                                p_ado.m_strSQL = "UPDATE scenario_merge SET cmborder = " + 
									                   Convert.ToInt32(p_row["cmborder"].ToString()).ToString().Trim()  + 
									                   ",mdbpathandfile='" + p_row["mdbpathandfile"].ToString().Trim() + "' " + 
									             "WHERE trim(lcase(mdbpathandfile)) = '" + p_row["mdbpathandfile"].ToString().Trim() + "';";
								p_ado.SqlNonQuery(p_ado.m_OleDbConnection,p_ado.m_strSQL);
							}
						}
					}
					p_ado.m_DataSet.Clear();
				}
				p_ado.m_OleDbConnection.Close();
			}
			p_ado=null;
			p_rows=null;
			p_row=null;
		}
		/// <summary>
		/// load the scenario_merge records into a an oledbdatareader and then to a
		/// combo box
		/// </summary>
		private void GetScenarioMergeRecords()
		{
			ado_data_access p_ado = new ado_data_access();
			p_ado.getScenarioConnStringAndMDBFile(ref this.m_strScenarioMDB,ref this.m_strScenarioConn,this.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text);
			p_ado.SqlQueryReader(this.m_strScenarioConn,"select * from scenario_merge order by cmborder ASC");
			if (p_ado.m_intError ==0)
			{
				this.PopulateComboBox(this.cmbMDBFiles,p_ado.m_OleDbDataReader,"mdbpathandfile");
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				this.m_strScenarioMergeMDB = this.cmbMDBFiles.Text.Trim();

			}
			p_ado= null;

		}

		private void cmbMDBFiles_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strScenarioMergeMDB.Trim().ToLower() != this.cmbMDBFiles.Text.Trim().ToLower())
			{
				this.PopulateCheckedComboWithMDBTables(this.lstSelectExistingTables, this.cmbMDBFiles.Text.Trim(),true);
			}
			this.m_strScenarioMergeMDB=this.cmbMDBFiles.Text.Trim();

		}

		private void btnNewSelectScenariosPrev_Click(object sender, System.EventArgs e)
		{
			this.grpboxOpen.Visible=true;
			this.grpboxNewSelectScenarios.Visible=false;
		}

		private void btnNewSelectViewPrev_Click(object sender, System.EventArgs e)
		{
			this.grpboxNewSelectScenarios.Visible=true;
			this.grpboxNewSelectView.Visible =false;
		}

		private void btnNewSelectViewCancel_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();
		}

		private void btnNewSelectScenariosCancel_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();
		}

		private void btnNewSelectScenariosNext_Click(object sender, System.EventArgs e)
		{
			if (this.lstSelectScenarios.CheckedItems.Count == 0)
				MessageBox.Show("A Minimum Of One Scenario Must Be Selected");
			else if (this.lstNewSelectTables.CheckedItems.Count == 0)
			{
				MessageBox.Show("A Minimum Of One Table Must Be Selected");
			}
			else
			{
				if (this.txtNewSelectViewMergeMDBFile.Text.Trim().Length ==0)
				{
					this.txtNewSelectViewMergeMDBFile.Text = this.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\core\\db\\scenario_merge.mdb";
				}
				if (this.chkboxNewSelectViewSaveToAccess.Checked==true)
				{
					this.btnNewSelectViewGetMDBFile.Enabled=true;
					this.txtNewSelectViewMergeMDBFile.Enabled=true;
					this.btnNewSelectViewStartAccess.Enabled=true;
				}
				else
				{
					this.btnNewSelectViewGetMDBFile.Enabled=false;
					this.txtNewSelectViewMergeMDBFile.Enabled=false;
					this.btnNewSelectViewStartAccess.Enabled=false;
				}

				this.grpboxNewSelectView.Visible=true;
				this.grpboxNewSelectScenarios.Visible=false;
			}
		}
		private void PopulateScenarioList()
		{
			ado_data_access p_ado = new ado_data_access();
            string strConn = p_ado.getMDBConnString(this.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text + "\\core\\db\\scenario_core_rule_definitions.mdb","","");
            p_ado.SqlQueryReader(strConn,"select scenario_id from scenario order by scenario_id");
			this.PopulateCheckedListBox(this.lstSelectScenarios,p_ado.m_OleDbDataReader,"scenario_id",false);
			p_ado.m_OleDbDataReader.Close();
			p_ado.m_OleDbConnection.Close();
			p_ado = null;

			if (this.lstSelectScenarios.Items.Count > 0 && this.lstNewSelectTables.Items.Count == 0) 
			{
				this.lstNewSelectTables.Items.Add("best_rx_summary");
				this.lstNewSelectTables.Items.Add("effective");
                this.lstNewSelectTables.Items.Add("max_nr_plots");
				this.lstNewSelectTables.Items.Add("max_nr_sum_psite");
				this.lstNewSelectTables.Items.Add("max_nr_sum_own");
                this.lstNewSelectTables.Items.Add("max_pnr_plots");
				this.lstNewSelectTables.Items.Add("max_pnr_sum_psite");
				this.lstNewSelectTables.Items.Add("max_pnr_sum_own");
				this.lstNewSelectTables.Items.Add("max_ci_imp_plots");
				this.lstNewSelectTables.Items.Add("max_ci_imp_sum_own");
				this.lstNewSelectTables.Items.Add("max_ci_imp_sum_psite");
				this.lstNewSelectTables.Items.Add("max_ci_imp_pnr_plots");
				this.lstNewSelectTables.Items.Add("max_ci_imp_pnr_sum_own");
				this.lstNewSelectTables.Items.Add("max_ci_imp_pnr_sum_psite");
				this.lstNewSelectTables.Items.Add("max_ti_imp_plots");
				this.lstNewSelectTables.Items.Add("max_ti_imp_sum_own");
				this.lstNewSelectTables.Items.Add("max_ti_imp_sum_psite");
				this.lstNewSelectTables.Items.Add("max_ti_imp_pnr_plots");
				this.lstNewSelectTables.Items.Add("max_ti_imp_pnr_sum_own");
				this.lstNewSelectTables.Items.Add("max_ti_imp_pnr_sum_psite");
                this.lstNewSelectTables.Items.Add("min_merch_plots");
				this.lstNewSelectTables.Items.Add("min_merch_sum_psite");
				this.lstNewSelectTables.Items.Add("min_merch_sum_own");
				this.lstNewSelectTables.Items.Add("min_merch_pnr_plots");
				this.lstNewSelectTables.Items.Add("min_merch_pnr_sum_psite");
				this.lstNewSelectTables.Items.Add("min_merch_pnr_sum_own");
				this.lstNewSelectTables.Items.Add("product_yields_net_rev_costs_summary");
				this.lstNewSelectTables.Items.Add("tree_vol_val_sum_by_rx");
			}
			


		}
		private void PopulateCheckedListBox(System.Windows.Forms.CheckedListBox p_lst,System.Data.OleDb.OleDbDataReader p_datareader,string strCol,bool bChecked)
		{
			int x=0;
			p_lst.Items.Clear();
			if (p_datareader.HasRows)
			{
				while (p_datareader.Read())
				{
					if (p_datareader[strCol].ToString().Trim().Length > 0)
					{
						p_lst.Items.Add(p_datareader[strCol].ToString().Trim());
				        p_lst.SetItemChecked(x, bChecked);
						x++;
					}
				}
						
			}
		}

		private void btnSelectAllScenarios_Click(object sender, System.EventArgs e)
		{
			for (int x = 0; x<= this.lstSelectScenarios.Items.Count-1;x++)
			{
				this.lstSelectScenarios.SetItemChecked(x,true);
			}

		}

		private void btnClearAllScenarios_Click(object sender, System.EventArgs e)
		{
			for (int x = 0; x<= this.lstSelectScenarios.Items.Count-1;x++)
			{
				this.lstSelectScenarios.SetItemChecked(x,false);
			}

		}

		private void btnSelectAllNewTables_Click(object sender, System.EventArgs e)
		{
			for (int x = 0; x<= this.lstNewSelectTables.Items.Count-1;x++)
			{
				this.lstNewSelectTables.SetItemChecked(x,true);
			}

		}

		private void btnClearAllNewTables_Click(object sender, System.EventArgs e)
		{
			for (int x = 0; x<= this.lstNewSelectTables.Items.Count-1;x++)
			{
				this.lstNewSelectTables.SetItemChecked(x,false);
			}

		}

		private void chkboxNewSelectViewSaveToAccess_Click(object sender, System.EventArgs e)
		{
			if (this.chkboxNewSelectViewSaveToAccess.Checked==true)
			{
				this.btnNewSelectViewGetMDBFile.Enabled=true;
				this.txtNewSelectViewMergeMDBFile.Enabled=true;
				this.btnNewSelectViewStartAccess.Enabled=true;
			}
			else
			{
				this.btnNewSelectViewGetMDBFile.Enabled=false;
				this.txtNewSelectViewMergeMDBFile.Enabled=false;
				this.btnNewSelectViewStartAccess.Enabled=false;
			}

		}

		private void btnNewSelectViewGetMDBFile_Click(object sender, System.EventArgs e)
		{
			string strMDBPathAndFile="";
			string strDir="";
			string strFile="";


			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Designate A New Or Existing MDB File";
			OpenFileDialog1.Filter = "Access Database File (*.MDB) |*.mdb";
			FIA_Biosum_Manager.utils p_oUtils = new utils();

			if (this.cmbMDBFiles.Text.Trim().Length > 0)
			{
				strFile = p_oUtils.getFileName(this.cmbMDBFiles.Text.Trim());
				strDir = p_oUtils.getDirectory(this.cmbMDBFiles.Text.Trim());
			}
			else
			{
				strDir = this.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\core\\db";
			}

			OpenFileDialog1.InitialDirectory = strDir;
			OpenFileDialog1.CheckFileExists = false;
			DialogResult result =  OpenFileDialog1.ShowDialog();

			if (result == DialogResult.OK)
			{
				strMDBPathAndFile = OpenFileDialog1.FileName.Trim();
				if (strMDBPathAndFile.Length > 0) 
				{
					this.txtNewSelectViewMergeMDBFile.Text = strMDBPathAndFile;
				}
			}
			p_oUtils=null;
			OpenFileDialog1 = null;

		
		}

		private void txtNewSelectViewMergeMDBFile_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void btnNewSelectViewStartAccess_Click(object sender, System.EventArgs e)
		{
			MergeTables();
			if (this.m_frmTherm.AbortProcess == false && this.m_intError == 0)
			{
				MessageBox.Show("Tables Joined Successfully");
				utils p_utils = new utils();
				p_utils.ShellExecute(this.m_strRandomFile);
			}


			

		}
		private void MergeTables()
		{
			
			dao_data_access p_dao = new dao_data_access();
	

			string strOutputJoinMDBFile="";
			
			string[] strScenarioResultsMDB;
			string[] strScenarioResultsConn ;
			
			strScenarioResultsMDB = new string[1];
			strScenarioResultsConn = new string[1];
			this.m_strMergeTables = new string[1];

			/***************************************************************************
			 **declare class object that will be used for scenario_results.mdb processing
			 ***************************************************************************/
			ado_data_access p_adoScenarioResults;

			/***************************************************************************
			 **declare class object used for processing scenario_core_rule_definitions
			 ***************************************************************************/
			ado_data_access p_adoScenario;

			/***************************************************************************
			 **declare class object used for processing link mdb file
			 ***************************************************************************/
			//ado_data_access p_adoLinks;
			

			this.m_intError = 0;

			//CREATE TEMP MDB WITH ALL THE NECESSARY LINKS
			utils p_utils = new utils();
			env p_env = new env();
			this.m_strRandomFile = p_utils.getRandomFile(p_env.strTempDir,"accdb");
			p_dao.CreateMDB(m_strRandomFile);
			this.m_intError = p_dao.m_intError;

			//CHECK TO SEE WHERE THE OUTPUT JOIN TABLES ARE CREATED
			/********************************************************
			 **create the merge mdb file if it does not exist
			 ********************************************************/
			if (this.chkboxNewSelectViewSaveToAccess.Checked == true && this.m_intError==0)
			{
				if (!System.IO.File.Exists(this.txtNewSelectViewMergeMDBFile.Text.Trim()))
				{
					p_dao.CreateMDB(this.txtNewSelectViewMergeMDBFile.Text.Trim());
					this.m_intError = p_dao.m_intError;
				}
				strOutputJoinMDBFile = this.txtNewSelectViewMergeMDBFile.Text.Trim();
			}
			else
			{
				strOutputJoinMDBFile = this.m_strRandomFile;
			}




			if (this.m_intError ==0)
			{

				int x=0;
				int y=0;
				int z=0;
               
				//CREATE THE TABLE STRUCTURES
				/*********************************************************
				 **table that contains the scenario_id field
				 *********************************************************/
				System.Data.DataTable p_dtScenarioId;    
				/*********************************************************************
				 **table that will combine fields from p_dtScenario and p_dtSource
				 *********************************************************************/
				System.Data.DataTable p_dtTarget;        
				/*********************************************************************
				 **table that contains structures of tables in scenario_results.mdb 
				 *********************************************************************/
				System.Data.DataTable p_dtSource; 

				
       
				/******************************************************************
				 **create a connection to scenario_core_rule_definitions
				 ******************************************************************/
				p_adoScenario = new ado_data_access();
				p_adoScenario.getScenarioConnStringAndMDBFile(ref this.m_strScenarioMDB,ref this.m_strScenarioConn,this.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text);
				p_adoScenario.OpenConnection(this.m_strScenarioConn);

				if (p_adoScenario.m_intError == 0)
				{
					/****************************************************************
					 **get the table structure that results from executing the sql
					 ****************************************************************/
					p_adoScenario.m_strSQL = "select scenario_id from scenario" ;
					p_dtScenarioId = p_adoScenario.getTableSchema(p_adoScenario.m_OleDbConnection,p_adoScenario.m_strSQL);
					
					strScenarioResultsMDB = new string[this.lstSelectScenarios.CheckedItems.Count];
					strScenarioResultsConn =  new string[this.lstSelectScenarios.CheckedItems.Count];
					System.Data.DataRow p_row;
					/*******************************************************
					 **process the user selected scenarios
					 *******************************************************/
					for (x=0; x<=this.lstSelectScenarios.CheckedItems.Count-1;x++)
					{
						strScenarioResultsMDB[x]="";
						strScenarioResultsConn[x]="";
						/*******************************************************
						 **make a connection to the scenario_results mdb table
						 *******************************************************/
						
						/******************************************************
						 **get path to the scenarios scenario_results.mdb
						 ******************************************************/
						strScenarioResultsMDB[x] = p_adoScenario.getSingleStringValueFromSQLQuery(p_adoScenario.m_OleDbConnection,
							"select path from scenario where trim(ucase(scenario_id))='" + this.lstSelectScenarios.CheckedItems[x].ToString().Trim().ToUpper() + "';"
							,"scenario");
						strScenarioResultsMDB[x] += "\\db\\scenario_results.mdb";
						/*****************************************************
						 **get the connection string
						 *****************************************************/
						strScenarioResultsConn[x]=p_adoScenario.getMDBConnString(strScenarioResultsMDB[x].ToString().Trim(),"","");
						/****************************************
						 **open connection to scenario resultsmdb
						 ****************************************/
						p_adoScenarioResults = new ado_data_access();
						p_adoScenarioResults.OpenConnection(strScenarioResultsConn[x]);
						if (p_adoScenarioResults.m_intError != 0)
						{
							this.m_intError = p_adoScenarioResults.m_intError;
							break;
						}
						if (x==0)
							m_strMergeTables = new string[this.lstNewSelectTables.CheckedItems.Count];
                           
						/*******************************************************
						 **process each table selected by the user
						 ******************************************************/
						
						
						for (y=0;y<=this.lstNewSelectTables.CheckedItems.Count-1;y++)
						{
							/************************************************
							 **if this is the first scenario selected 
							 **then create the new table structure in the
							 **scenario merge mdb
							 ***********************************************/
							if (x==0)
							{
								
								/*******************************************************
								 **get the table structure of the user selected table
								 *******************************************************/
								p_adoScenarioResults.m_strSQL = "select * from " + this.lstNewSelectTables.CheckedItems[y].ToString().Trim();
								p_dtSource = p_adoScenarioResults.getTableSchema(p_adoScenarioResults.m_OleDbConnection,p_adoScenarioResults.m_strSQL);
								if (p_adoScenarioResults.m_intError !=0)
								{
									this.m_intError = p_adoScenarioResults.m_intError;
									break;
								}
   
								/******************************************************
								 **create the table structure in the scenario merge mdb
								 ******************************************************/

								/******************************************************
								 **copy the scenario_id field to the p_dtTarget table
								 *******************************************************/
								p_dtTarget = p_dtScenarioId.Copy();

								/********************************************************
								 **append the p_dtSource fields to the p_dtTarget Table
								 ********************************************************/
								for (z=0; z <= p_dtSource.Rows.Count-1;z++)
								{
									p_row = p_dtTarget.NewRow();
									for (int intCol = 0; intCol <= p_dtSource.Columns.Count-1;intCol++)
									{
										p_row[intCol] = p_dtSource.Rows[z][intCol];
									}
									p_dtTarget.Rows.Add(p_row);

								}
								/**************************************************
								 **okay the table is defined so lets create the
								 **structure in the merge or temp mdb file
								 **************************************************/
								p_dao.CreateMDBTableFromDataSetTable(strOutputJoinMDBFile.Trim(),
									this.lstNewSelectTables.CheckedItems[y].ToString().Trim(),
									p_dtTarget,false);
								/**********************************************
								 **save the name of the table 
								 **********************************************/
								m_strMergeTables[y] = p_dao.m_strTable;
								p_dtSource.Clear();
								p_dtTarget.Clear();

							}
						}
	
						p_adoScenarioResults.m_OleDbConnection.Close();

						if (this.m_intError !=0)
							break;

							
					}
					
					p_adoScenario.m_OleDbConnection.Close();
					p_dtScenarioId.Clear();
					p_row = null;
				}
				else
				{
					this.m_intError = p_adoScenario.m_intError;
				}
				p_dtSource = null;
				p_dtTarget = null;
				p_dtScenarioId = null;
				p_adoScenario = null;
					

				if (this.m_intError == 0)
				{
					if (p_dao.m_intError == 0)
					{
						p_dao.OpenDb(m_strRandomFile);

						/****************************************
						 **place all the links associated with
						 **the join in the temp mdb file
						 ****************************************/
						for (x=0; x <= this.lstSelectScenarios.CheckedItems.Count-1;x++)
						{
							for (y=0; y<= this.lstNewSelectTables.CheckedItems.Count-1;y++)
							{

								p_dao.CreateTableLink(p_dao.m_DaoDatabase,
									this.lstSelectScenarios.CheckedItems[x].ToString().Trim() + "_" + this.lstNewSelectTables.CheckedItems[y].ToString().Trim(),
									strScenarioResultsMDB[x],
									this.lstNewSelectTables.CheckedItems[y].ToString().Trim());
								if (p_dao.m_intError !=0)
									break;
							}
							if (p_dao.m_intError !=0)
								break;

						}
						if (p_dao.m_intError == 0)
						{
							/****************************************************
							 **create links to the scenario merge mdb only
							 **if the user wants a permanent access table copy
							 **of the join
							 ****************************************************/
							if (this.chkboxNewSelectViewSaveToAccess.Checked == true)
							{
								//create scenario merge mdb table links
								for (x=0; x<=this.lstNewSelectTables.CheckedItems.Count-1;x++)
								{
									p_dao.CreateTableLink(p_dao.m_DaoDatabase,
										m_strMergeTables[x].ToString().Trim(),
										this.txtNewSelectViewMergeMDBFile.Text.Trim(),
										m_strMergeTables[x].ToString().Trim());
									if (p_dao.m_intError !=0)
										break;
								}
							}
						}
						p_dao.m_DaoDatabase.Close();
						if (p_dao.m_intError == 0)
						{
							this.m_frmTherm = new frmTherm();
							this.m_frmTherm.btnCancel.Click += new System.EventHandler(this.ThermCancel);
							this.m_frmTherm.Text = "Join Scenario Tables";
							this.m_frmTherm.Visible=true;
							m_frmTherm.btnCancel.Visible=true;
							m_frmTherm.lblMsg.Visible=true;
							m_frmTherm.progressBar1.Minimum=0;
							m_frmTherm.progressBar1.Visible=true;
							m_frmTherm.progressBar1.Maximum = this.lstSelectScenarios.CheckedItems.Count  * this.lstNewSelectTables.CheckedItems.Count;
							this.m_frmTherm.AbortProcess = false;
							this.m_frmTherm.Refresh();
							//INSERT THE DATA
							this.m_adoLinks = new ado_data_access();
							this.m_strRandomFileConn = this.m_adoLinks.getMDBConnString(m_strRandomFile,"","");
							this.m_adoLinks.OpenConnection(this.m_strRandomFileConn);

							if (this.m_adoLinks.m_intError == 0)
							{
								z=0;
								for (x=0;x<=this.lstSelectScenarios.CheckedItems.Count-1;x++)
								{
									for (y=0;y<=this.lstNewSelectTables.CheckedItems.Count-1;y++)
									{
										this.m_frmTherm.lblMsg.Text  = this.lstSelectScenarios.CheckedItems[x].ToString().Trim() + ": " + this.lstNewSelectTables.CheckedItems[y].ToString().Trim();
										this.m_frmTherm.lblMsg.Refresh();
										m_frmTherm.progressBar1.Value = z++;
										this.m_adoLinks.m_strSQL = "INSERT INTO " + m_strMergeTables[y].ToString().Trim() + " " + 
											"SELECT * FROM " + this.lstSelectScenarios.CheckedItems[x].ToString().Trim() + 
											"_" + this.lstNewSelectTables.CheckedItems[y].ToString().Trim() + ";";
										this.m_adoLinks.SqlNonQuery(this.m_adoLinks.m_OleDbConnection,this.m_adoLinks.m_strSQL);
										if (this.m_adoLinks.m_intError== 0)
										{
											this.m_adoLinks.m_strSQL = "UPDATE " + m_strMergeTables[y].ToString().Trim() + " " +
												"SET scenario_id = '" + this.lstSelectScenarios.CheckedItems[x].ToString().Trim() + "' " + 
												"WHERE scenario_id IS NULL";
											this.m_adoLinks.SqlNonQuery(this.m_adoLinks.m_OleDbConnection,this.m_adoLinks.m_strSQL);

										}
										if (this.m_adoLinks.m_intError != 0) break;
										System.Windows.Forms.Application.DoEvents();
										if (this.m_frmTherm.AbortProcess == true) break;

									}
									if (this.m_adoLinks.m_intError != 0) break;
								}
								
								this.m_adoLinks.m_OleDbConnection.Close();
								this.m_frmTherm.Close();

							}
							this.m_intError = this.m_adoLinks.m_intError;
							
						}
						else
						{
							this.m_intError = p_dao.m_intError;
						}
					
					}
					else
					{
						this.m_intError = p_dao.m_intError;
					}
                    
				}
				p_utils = null;
				p_env = null;
				p_dao = null;
			    
			        
			}

				
		}

		private void ThermCancel(object sender, System.EventArgs e)
		{
			string strMsg = "Do you wish to cancel the data join (Y/N)?";
			DialogResult result = MessageBox.Show(strMsg,"Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					this.m_frmTherm.AbortProcess = true;
					this.m_frmTherm.Hide();
					return;
				case DialogResult.No:
					return;
			}                
		}

		private void btnViewSelectViewInGrid_Click(object sender, System.EventArgs e)
		{
			this.m_frmGridView = new frmGridView();
			this.MergeTables();
			if (this.m_frmTherm.AbortProcess == false && this.m_intError==0)
			{
				for (int x=0;x<=this.m_strMergeTables.Length -1 ;x++)
				{
					string strSQL = "select * from " + this.m_strMergeTables[x].ToString().Trim() + ";";
					this.m_frmGridView.LoadDataSet(this.m_strRandomFileConn,
						strSQL,
						this.m_strMergeTables[x].Trim());

				}
				this.m_frmGridView.Text = "Core Analysis: Join Scenario Tables";
				this.m_frmGridView.TileGridViews();
				this.m_frmGridView.Visible=true;
				this.m_frmGridView.MdiParent = this.m_frmMain;
				this.m_frmGridView.Show();
			}
			else
			{
				this.m_frmGridView.Close();
			}
		}
		
		

	}
}
