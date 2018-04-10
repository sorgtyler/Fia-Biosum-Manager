using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_ffe.
	/// </summary>
	public class uc_core_scenario_weighted_average : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;
		//private int m_intFullHt=400;
		public System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;
		public System.Data.DataSet m_DataSet;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionMaster;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		public System.Data.OleDb.OleDbCommand m_OleDbCommand;
		public System.Data.DataRelation m_DataRelation;
		public System.Data.DataRow m_DataRow;
		public int m_intError = 0;
		public string m_strError = "";
		private System.Windows.Forms.Button btnNext;
		private System.Windows.Forms.Button btnPrev;
		private string m_strCurrentIndexTypeAndClass;
		private string m_strOverallEffectiveExpression="";
		private FIA_Biosum_Manager.utils m_oUtils; 
		private bool _bTorchIndex=true;
		private bool _bCrownIndex=true;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmCoreScenario _frmScenario=null;
		private FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker _uc_tiebreaker;

		
		private int m_intCurVar=-1;
		int m_intCurVariableDefinitionStepCount=1;
        string[] m_strUserNavigation = null;
        private System.Windows.Forms.GroupBox grpboxDetails;

		public bool m_bSave=false;
        //list view associated classes
        private ListViewEmbeddedControls.ListViewEx m_lvEx;



		const int COLUMN_CHECKBOX=0;
		const int COLUMN_OPTIMIZE_VARIABLE=1;
		const int COLUMN_FVS_VARIABLE=2;
		const int COLUMN_VALUESOURCE=3;
		const int COLUMN_MAXMIN=4;
		const int COLUMN_USEFILTER=5;
		const int COLUMN_FILTER_OPERATOR=6;
        const int COLUMN_FILTER_VALUE = 7;
        private FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.Variables _oCurVar;
		public bool m_bFirstTime=true;
		private bool _bDisplayAuditMsg=true;
        private bool m_bIgnoreListViewItemCheck = false;
        private int m_intPrevColumnIdx = -1;
        private System.Windows.Forms.GroupBox grpboxSummary;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors=new ListViewAlternateBackgroundColors();
         private ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private string m_strLastValue = "";
        private FIA_Biosum_Manager.frmMain m_frmMain;
        public int m_DialogHt;
        public Panel pnlDetails;
        private TextBox txtFvsVariableName;
        private Label label7;
        public Button btnFvsCalculate;
        private Button btnFvsDetailsCancel;
        private GroupBox grpBoxFvsBaseline;
        private ComboBox cboFvsVariableBaselinePkg;
        private GroupBox groupBox3;
        private ListBox lstFVSFieldsList;
        private GroupBox groupBox2;
        private ListBox lstFVSTablesList;
        private Label LblSelectedVariable;
        private Label lblSelectedFVSVariable;
        private TextBox textBox18;
        private Label label8;
        public int m_DialogWd;
        private Panel pnlSummary;
        private Button btnProperties;
        private Button btnDelete;
        private Button btnNewFvs;
        private ListView lstVariables;
        private Button btnCancelSummary;
        private ColumnHeader vName;
        private ColumnHeader vDescription;
        private Button BtnHelp;
        private ListViewAlternateBackgroundColors m_oLvAlternateColors = new FIA_Biosum_Manager.ListViewAlternateBackgroundColors();
        private ListViewColumnSorter lvwColumnSorter;
        private const int COL_YEAR = 1;
        private const int COL_SEQNUM = 2;
        private Button btnNewEcon;
        private GroupBox grpBoxEconomicVariable;
        public Panel panel1;
        private DataGridView dgEcon;
        private Button button2;
        private TextBox textBox1;
        private Label label1;
        private TextBox textBox2;
        private Label label2;
        private Button BtnSaveEcon;
        private Button btnEconDetailsCancel;
        private GroupBox groupBox8;
        private ListBox listBox3;
        private Label label3;
        private Label label4;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private TextBox textBox4;
        private Label label6;
        private TextBox txtFvsVariableTotalWeight;
        private Label label5;
        private ColumnHeader vType;
        private DataGrid m_dg;
        private env m_oEnv;
        private System.Data.DataTable m_dtTableSchema;
        private System.Data.DataView m_dv;
        private System.Collections.Generic.Dictionary<string, System.Collections.Generic.IList<String>> m_dictFVSTables;
        private Button btnFVSVariableValue;
        private FIA_Biosum_Manager.CoreAnalysisScenarioTools m_oCoreAnalysisScenarioTools = new CoreAnalysisScenarioTools();


        public uc_core_scenario_weighted_average(FIA_Biosum_Manager.frmMain p_frmMain)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_oUtils = new utils();
            this.m_frmMain = p_frmMain;
			this.m_strCurrentIndexTypeAndClass="";

            this.grpboxDetails.Top = grpboxSummary.Top;
            this.grpboxDetails.Left = this.grpboxSummary.Left;
            this.grpboxDetails.Height = this.grpboxSummary.Height;
            this.grpboxDetails.Width = this.grpboxSummary.Width;
            this.grpboxDetails.Hide();

            this.grpBoxEconomicVariable.Top = grpboxSummary.Top;
            this.grpBoxEconomicVariable.Left = this.grpboxSummary.Left;
            this.grpBoxEconomicVariable.Height = this.grpboxSummary.Height;
            this.grpBoxEconomicVariable.Width = this.grpboxSummary.Width;
            this.grpBoxEconomicVariable.Hide();

            //m_oValidate.RoundDecimalLength = 0;
            //m_oValidate.Money = false;
            //m_oValidate.NullsAllowed = false;
            //m_oValidate.TestForMaxMin = false;
            //m_oValidate.MinValue = -1000;
            //m_oValidate.TestForMin = true;

            m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceListView = this.lstVariables;
            this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvAlternateColors.CustomFullRowSelect = true;
            if (frmMain.g_oGridViewFont != null) lstVariables.Font = frmMain.g_oGridViewFont;

			// TODO: Add any initialization after the InitializeComponent call
            this.m_DialogWd = this.Width + 25;
            this.m_DialogHt = this.pnlDetails.Height + 120;
            this.Height = m_DialogHt -40;

            //@ToDo: Plugging temporary data for screenshots
            this.m_oLvAlternateColors.InitializeRowCollection();
            lstVariables.Items.Add("Weight_Torch");
            lstVariables.Items[0].UseItemStyleForSubItems = false;
            lstVariables.Items[0].SubItems.Add("Weighted P-Torch severity");
            lstVariables.Items[0].SubItems.Add("FVS");
            m_oLvAlternateColors.AddRow();
            m_oLvAlternateColors.AddColumns(0, lstVariables.Columns.Count);
            lstVariables.Items.Add("wHazard");
            lstVariables.Items[1].UseItemStyleForSubItems = false;
            lstVariables.Items[1].SubItems.Add("Weighted Hazard Score");
            lstVariables.Items[1].SubItems.Add("FVS");
            m_oLvAlternateColors.AddRow();
            m_oLvAlternateColors.AddColumns(1, lstVariables.Columns.Count);
            lstVariables.Items.Add("total_volume_1");
            lstVariables.Items[2].UseItemStyleForSubItems = false;
            lstVariables.Items[2].SubItems.Add("Default calculation for total volume");
            lstVariables.Items[2].SubItems.Add("Economic");
            m_oLvAlternateColors.AddRow();
            m_oLvAlternateColors.AddColumns(2, lstVariables.Columns.Count);
            this.m_oLvAlternateColors.ListView();




            var row = (DataGridViewRow) dgEcon.RowTemplate.Clone();
            row.CreateCells(dgEcon, "1", "1");
            dgEcon.Rows.Add(row);
            row = (DataGridViewRow)dgEcon.RowTemplate.Clone();
            row.CreateCells(dgEcon, "2", "1");
            dgEcon.Rows.Add(row);
            row = (DataGridViewRow)dgEcon.RowTemplate.Clone();
            row.CreateCells(dgEcon, "3", "1");
            dgEcon.Rows.Add(row);
            row = (DataGridViewRow)dgEcon.RowTemplate.Clone();
            row.CreateCells(dgEcon, "4", "1");
            dgEcon.Rows.Add(row);

            this.m_oEnv = new env();
            this.loadvalues();
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpBoxEconomicVariable = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.dgEcon = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.BtnSaveEcon = new System.Windows.Forms.Button();
            this.btnEconDetailsCancel = new System.Windows.Forms.Button();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.listBox3 = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.grpboxSummary = new System.Windows.Forms.GroupBox();
            this.pnlSummary = new System.Windows.Forms.Panel();
            this.btnNewEcon = new System.Windows.Forms.Button();
            this.btnCancelSummary = new System.Windows.Forms.Button();
            this.btnProperties = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnNewFvs = new System.Windows.Forms.Button();
            this.lstVariables = new System.Windows.Forms.ListView();
            this.vName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.vDescription = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.vType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.grpboxDetails = new System.Windows.Forms.GroupBox();
            this.pnlDetails = new System.Windows.Forms.Panel();
            this.btnFVSVariableValue = new System.Windows.Forms.Button();
            this.m_dg = new System.Windows.Forms.DataGrid();
            this.txtFvsVariableTotalWeight = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.BtnHelp = new System.Windows.Forms.Button();
            this.textBox18 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtFvsVariableName = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btnFvsCalculate = new System.Windows.Forms.Button();
            this.btnFvsDetailsCancel = new System.Windows.Forms.Button();
            this.grpBoxFvsBaseline = new System.Windows.Forms.GroupBox();
            this.cboFvsVariableBaselinePkg = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lstFVSFieldsList = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lstFVSTablesList = new System.Windows.Forms.ListBox();
            this.LblSelectedVariable = new System.Windows.Forms.Label();
            this.lblSelectedFVSVariable = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.grpBoxEconomicVariable.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgEcon)).BeginInit();
            this.groupBox8.SuspendLayout();
            this.grpboxSummary.SuspendLayout();
            this.pnlSummary.SuspendLayout();
            this.grpboxDetails.SuspendLayout();
            this.pnlDetails.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).BeginInit();
            this.grpBoxFvsBaseline.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.grpBoxEconomicVariable);
            this.groupBox1.Controls.Add(this.grpboxSummary);
            this.groupBox1.Controls.Add(this.grpboxDetails);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(872, 2000);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Leave += new System.EventHandler(this.groupBox1_Leave);
            // 
            // grpBoxEconomicVariable
            // 
            this.grpBoxEconomicVariable.BackColor = System.Drawing.SystemColors.Control;
            this.grpBoxEconomicVariable.Controls.Add(this.panel1);
            this.grpBoxEconomicVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpBoxEconomicVariable.ForeColor = System.Drawing.Color.Black;
            this.grpBoxEconomicVariable.Location = new System.Drawing.Point(6, 1027);
            this.grpBoxEconomicVariable.Name = "grpBoxEconomicVariable";
            this.grpBoxEconomicVariable.Size = new System.Drawing.Size(856, 472);
            this.grpBoxEconomicVariable.TabIndex = 36;
            this.grpBoxEconomicVariable.TabStop = false;
            this.grpBoxEconomicVariable.Text = "Weighted Economic Variable";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.textBox4);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.dgEcon);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBox2);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.BtnSaveEcon);
            this.panel1.Controls.Add(this.btnEconDetailsCancel);
            this.panel1.Controls.Add(this.groupBox8);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(850, 451);
            this.panel1.TabIndex = 70;
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.SystemColors.Control;
            this.textBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox4.Location = new System.Drawing.Point(393, 301);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(121, 22);
            this.textBox4.TabIndex = 92;
            this.textBox4.Text = "4.0";
            this.textBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(387, 279);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(139, 24);
            this.label6.TabIndex = 91;
            this.label6.Text = "TOTAL WEIGHTS";
            // 
            // dgEcon
            // 
            this.dgEcon.AllowUserToAddRows = false;
            this.dgEcon.AllowUserToDeleteRows = false;
            this.dgEcon.AllowUserToResizeRows = false;
            this.dgEcon.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgEcon.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn3});
            this.dgEcon.Location = new System.Drawing.Point(18, 172);
            this.dgEcon.Name = "dgEcon";
            this.dgEcon.Size = new System.Drawing.Size(350, 150);
            this.dgEcon.TabIndex = 88;
            // 
            // dataGridViewTextBoxColumn1
            // 
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dataGridViewTextBoxColumn1.DefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewTextBoxColumn1.FillWeight = 80F;
            this.dataGridViewTextBoxColumn1.HeaderText = "CYCLE";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 80;
            // 
            // dataGridViewTextBoxColumn3
            // 
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.Red;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.White;
            this.dataGridViewTextBoxColumn3.DefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridViewTextBoxColumn3.HeaderText = "WEIGHT";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // button2
            // 
            this.button2.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.button2.Location = new System.Drawing.Point(565, 402);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(64, 24);
            this.button2.TabIndex = 87;
            this.button2.Text = "Help";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(173, 386);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(259, 40);
            this.textBox1.TabIndex = 86;
            this.textBox1.Text = "Default calculation for total volume";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(13, 389);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(212, 24);
            this.label1.TabIndex = 85;
            this.label1.Text = "Description:";
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(173, 357);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(259, 22);
            this.textBox2.TabIndex = 84;
            this.textBox2.Text = "total_volume_1";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(13, 360);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(212, 24);
            this.label2.TabIndex = 79;
            this.label2.Text = "Weighted variable name:";
            // 
            // BtnSaveEcon
            // 
            this.BtnSaveEcon.Location = new System.Drawing.Point(634, 402);
            this.BtnSaveEcon.Name = "BtnSaveEcon";
            this.BtnSaveEcon.Size = new System.Drawing.Size(76, 24);
            this.BtnSaveEcon.TabIndex = 77;
            this.BtnSaveEcon.Text = "Save";
            // 
            // btnEconDetailsCancel
            // 
            this.btnEconDetailsCancel.Location = new System.Drawing.Point(716, 402);
            this.btnEconDetailsCancel.Name = "btnEconDetailsCancel";
            this.btnEconDetailsCancel.Size = new System.Drawing.Size(64, 24);
            this.btnEconDetailsCancel.TabIndex = 75;
            this.btnEconDetailsCancel.Text = "Cancel";
            this.btnEconDetailsCancel.Click += new System.EventHandler(this.btnEconDetailsCancel_Click);
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.listBox3);
            this.groupBox8.Location = new System.Drawing.Point(18, 5);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(200, 133);
            this.groupBox8.TabIndex = 71;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Variable type";
            // 
            // listBox3
            // 
            this.listBox3.FormattingEnabled = true;
            this.listBox3.ItemHeight = 16;
            this.listBox3.Items.AddRange(new object[] {
            "Total Volume",
            "Merch Volume",
            "Chip Volume",
            "Net Revenue",
            "Gross Costs"});
            this.listBox3.Location = new System.Drawing.Point(6, 21);
            this.listBox3.Name = "listBox3";
            this.listBox3.Size = new System.Drawing.Size(181, 100);
            this.listBox3.TabIndex = 70;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(224, 145);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(302, 24);
            this.label3.TabIndex = 69;
            this.label3.Text = "Total Volume";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(11, 144);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(237, 24);
            this.label4.TabIndex = 68;
            this.label4.Text = "Selected Economic Variable Type:";
            // 
            // grpboxSummary
            // 
            this.grpboxSummary.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxSummary.Controls.Add(this.pnlSummary);
            this.grpboxSummary.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxSummary.ForeColor = System.Drawing.Color.Black;
            this.grpboxSummary.Location = new System.Drawing.Point(8, 48);
            this.grpboxSummary.Name = "grpboxSummary";
            this.grpboxSummary.Size = new System.Drawing.Size(856, 472);
            this.grpboxSummary.TabIndex = 35;
            this.grpboxSummary.TabStop = false;
            // 
            // pnlSummary
            // 
            this.pnlSummary.AutoScroll = true;
            this.pnlSummary.Controls.Add(this.btnNewEcon);
            this.pnlSummary.Controls.Add(this.btnCancelSummary);
            this.pnlSummary.Controls.Add(this.btnProperties);
            this.pnlSummary.Controls.Add(this.btnDelete);
            this.pnlSummary.Controls.Add(this.btnNewFvs);
            this.pnlSummary.Controls.Add(this.lstVariables);
            this.pnlSummary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlSummary.Location = new System.Drawing.Point(3, 18);
            this.pnlSummary.Name = "pnlSummary";
            this.pnlSummary.Size = new System.Drawing.Size(850, 451);
            this.pnlSummary.TabIndex = 12;
            // 
            // btnNewEcon
            // 
            this.btnNewEcon.Location = new System.Drawing.Point(14, 360);
            this.btnNewEcon.Name = "btnNewEcon";
            this.btnNewEcon.Size = new System.Drawing.Size(148, 32);
            this.btnNewEcon.TabIndex = 14;
            this.btnNewEcon.Text = "New Econ Variable";
            this.btnNewEcon.Click += new System.EventHandler(this.btnNewEcon_Click);
            // 
            // btnCancelSummary
            // 
            this.btnCancelSummary.Location = new System.Drawing.Point(491, 360);
            this.btnCancelSummary.Name = "btnCancelSummary";
            this.btnCancelSummary.Size = new System.Drawing.Size(114, 32);
            this.btnCancelSummary.TabIndex = 13;
            this.btnCancelSummary.Text = "Cancel";
            this.btnCancelSummary.Click += new System.EventHandler(this.btnCancelSummary_Click);
            // 
            // btnProperties
            // 
            this.btnProperties.Location = new System.Drawing.Point(371, 360);
            this.btnProperties.Name = "btnProperties";
            this.btnProperties.Size = new System.Drawing.Size(114, 32);
            this.btnProperties.TabIndex = 12;
            this.btnProperties.Text = "Properties";
            // 
            // btnDelete
            // 
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(301, 360);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(64, 32);
            this.btnDelete.TabIndex = 11;
            this.btnDelete.Text = "Delete";
            // 
            // btnNewFvs
            // 
            this.btnNewFvs.Location = new System.Drawing.Point(169, 360);
            this.btnNewFvs.Name = "btnNewFvs";
            this.btnNewFvs.Size = new System.Drawing.Size(126, 32);
            this.btnNewFvs.TabIndex = 4;
            this.btnNewFvs.Text = "New FVS Variable";
            this.btnNewFvs.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // lstVariables
            // 
            this.lstVariables.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.vName,
            this.vDescription,
            this.vType});
            this.lstVariables.GridLines = true;
            this.lstVariables.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstVariables.HideSelection = false;
            this.lstVariables.Location = new System.Drawing.Point(18, 18);
            this.lstVariables.MultiSelect = false;
            this.lstVariables.Name = "lstVariables";
            this.lstVariables.Size = new System.Drawing.Size(640, 336);
            this.lstVariables.TabIndex = 2;
            this.lstVariables.UseCompatibleStateImageBehavior = false;
            this.lstVariables.View = System.Windows.Forms.View.Details;
            // 
            // vName
            // 
            this.vName.Text = "Variable Name";
            this.vName.Width = 200;
            // 
            // vDescription
            // 
            this.vDescription.Text = "Description";
            this.vDescription.Width = 350;
            // 
            // vType
            // 
            this.vType.Text = "Type";
            this.vType.Width = 100;
            // 
            // grpboxDetails
            // 
            this.grpboxDetails.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxDetails.Controls.Add(this.pnlDetails);
            this.grpboxDetails.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxDetails.ForeColor = System.Drawing.Color.Black;
            this.grpboxDetails.Location = new System.Drawing.Point(8, 536);
            this.grpboxDetails.Name = "grpboxDetails";
            this.grpboxDetails.Size = new System.Drawing.Size(856, 472);
            this.grpboxDetails.TabIndex = 32;
            this.grpboxDetails.TabStop = false;
            this.grpboxDetails.Text = "Weighted FVS Variable";
            this.grpboxDetails.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePost_Resize);
            // 
            // pnlDetails
            // 
            this.pnlDetails.AutoScroll = true;
            this.pnlDetails.Controls.Add(this.btnFVSVariableValue);
            this.pnlDetails.Controls.Add(this.m_dg);
            this.pnlDetails.Controls.Add(this.txtFvsVariableTotalWeight);
            this.pnlDetails.Controls.Add(this.label5);
            this.pnlDetails.Controls.Add(this.BtnHelp);
            this.pnlDetails.Controls.Add(this.textBox18);
            this.pnlDetails.Controls.Add(this.label8);
            this.pnlDetails.Controls.Add(this.txtFvsVariableName);
            this.pnlDetails.Controls.Add(this.label7);
            this.pnlDetails.Controls.Add(this.btnFvsCalculate);
            this.pnlDetails.Controls.Add(this.btnFvsDetailsCancel);
            this.pnlDetails.Controls.Add(this.grpBoxFvsBaseline);
            this.pnlDetails.Controls.Add(this.groupBox3);
            this.pnlDetails.Controls.Add(this.groupBox2);
            this.pnlDetails.Controls.Add(this.LblSelectedVariable);
            this.pnlDetails.Controls.Add(this.lblSelectedFVSVariable);
            this.pnlDetails.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDetails.Location = new System.Drawing.Point(3, 18);
            this.pnlDetails.Name = "pnlDetails";
            this.pnlDetails.Size = new System.Drawing.Size(850, 451);
            this.pnlDetails.TabIndex = 70;
            // 
            // btnFVSVariableValue
            // 
            this.btnFVSVariableValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariableValue.Location = new System.Drawing.Point(646, 28);
            this.btnFVSVariableValue.Name = "btnFVSVariableValue";
            this.btnFVSVariableValue.Size = new System.Drawing.Size(139, 98);
            this.btnFVSVariableValue.TabIndex = 92;
            this.btnFVSVariableValue.Text = "Select";
            this.btnFVSVariableValue.Click += new System.EventHandler(this.btnFVSVariableValue_Click);
            // 
            // m_dg
            // 
            this.m_dg.DataMember = "";
            this.m_dg.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.m_dg.Location = new System.Drawing.Point(18, 165);
            this.m_dg.Name = "m_dg";
            this.m_dg.Size = new System.Drawing.Size(403, 177);
            this.m_dg.TabIndex = 91;
            this.m_dg.CurrentCellChanged += new System.EventHandler(this.Grid_CurCellChange);
            // 
            // txtFvsVariableTotalWeight
            // 
            this.txtFvsVariableTotalWeight.BackColor = System.Drawing.SystemColors.Control;
            this.txtFvsVariableTotalWeight.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFvsVariableTotalWeight.Location = new System.Drawing.Point(463, 297);
            this.txtFvsVariableTotalWeight.Name = "txtFvsVariableTotalWeight";
            this.txtFvsVariableTotalWeight.ReadOnly = true;
            this.txtFvsVariableTotalWeight.Size = new System.Drawing.Size(121, 22);
            this.txtFvsVariableTotalWeight.TabIndex = 90;
            this.txtFvsVariableTotalWeight.Text = "0.0";
            this.txtFvsVariableTotalWeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(466, 275);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(139, 24);
            this.label5.TabIndex = 89;
            this.label5.Text = "TOTAL WEIGHTS";
            // 
            // BtnHelp
            // 
            this.BtnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.BtnHelp.Location = new System.Drawing.Point(565, 402);
            this.BtnHelp.Name = "BtnHelp";
            this.BtnHelp.Size = new System.Drawing.Size(64, 24);
            this.BtnHelp.TabIndex = 87;
            this.BtnHelp.Text = "Help";
            // 
            // textBox18
            // 
            this.textBox18.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox18.Location = new System.Drawing.Point(173, 386);
            this.textBox18.Multiline = true;
            this.textBox18.Name = "textBox18";
            this.textBox18.Size = new System.Drawing.Size(259, 40);
            this.textBox18.TabIndex = 86;
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(13, 389);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(87, 24);
            this.label8.TabIndex = 85;
            this.label8.Text = "Description:";
            // 
            // txtFvsVariableName
            // 
            this.txtFvsVariableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFvsVariableName.Location = new System.Drawing.Point(173, 357);
            this.txtFvsVariableName.Name = "txtFvsVariableName";
            this.txtFvsVariableName.ReadOnly = true;
            this.txtFvsVariableName.Size = new System.Drawing.Size(259, 22);
            this.txtFvsVariableName.TabIndex = 84;
            this.txtFvsVariableName.Text = "PTorch_Sev_1";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(13, 360);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(212, 24);
            this.label7.TabIndex = 79;
            this.label7.Text = "Weighted variable name:";
            // 
            // btnFvsCalculate
            // 
            this.btnFvsCalculate.Location = new System.Drawing.Point(634, 402);
            this.btnFvsCalculate.Name = "btnFvsCalculate";
            this.btnFvsCalculate.Size = new System.Drawing.Size(76, 24);
            this.btnFvsCalculate.TabIndex = 77;
            this.btnFvsCalculate.Text = "Calculate";
            // 
            // btnFvsDetailsCancel
            // 
            this.btnFvsDetailsCancel.Location = new System.Drawing.Point(716, 402);
            this.btnFvsDetailsCancel.Name = "btnFvsDetailsCancel";
            this.btnFvsDetailsCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFvsDetailsCancel.TabIndex = 75;
            this.btnFvsDetailsCancel.Text = "Cancel";
            this.btnFvsDetailsCancel.Click += new System.EventHandler(this.btnFvsDetailsCancel_Click);
            // 
            // grpBoxFvsBaseline
            // 
            this.grpBoxFvsBaseline.Controls.Add(this.cboFvsVariableBaselinePkg);
            this.grpBoxFvsBaseline.Location = new System.Drawing.Point(8, 7);
            this.grpBoxFvsBaseline.Name = "grpBoxFvsBaseline";
            this.grpBoxFvsBaseline.Size = new System.Drawing.Size(154, 48);
            this.grpBoxFvsBaseline.TabIndex = 74;
            this.grpBoxFvsBaseline.TabStop = false;
            this.grpBoxFvsBaseline.Text = "Baseline RxPackage";
            // 
            // cboFvsVariableBaselinePkg
            // 
            this.cboFvsVariableBaselinePkg.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFvsVariableBaselinePkg.Location = new System.Drawing.Point(8, 18);
            this.cboFvsVariableBaselinePkg.Name = "cboFvsVariableBaselinePkg";
            this.cboFvsVariableBaselinePkg.Size = new System.Drawing.Size(72, 24);
            this.cboFvsVariableBaselinePkg.TabIndex = 77;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lstFVSFieldsList);
            this.groupBox3.Location = new System.Drawing.Point(440, 7);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(200, 133);
            this.groupBox3.TabIndex = 72;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "FVS Variable(s)";
            // 
            // lstFVSFieldsList
            // 
            this.lstFVSFieldsList.FormattingEnabled = true;
            this.lstFVSFieldsList.ItemHeight = 16;
            this.lstFVSFieldsList.Location = new System.Drawing.Point(6, 21);
            this.lstFVSFieldsList.Name = "lstFVSFieldsList";
            this.lstFVSFieldsList.Size = new System.Drawing.Size(181, 100);
            this.lstFVSFieldsList.Sorted = true;
            this.lstFVSFieldsList.TabIndex = 70;
            this.lstFVSFieldsList.SelectedIndexChanged += new System.EventHandler(this.lstFVSFieldsList_SelectedIndexChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lstFVSTablesList);
            this.groupBox2.Location = new System.Drawing.Point(208, 7);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 133);
            this.groupBox2.TabIndex = 71;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "FVS Variable Table(s)";
            // 
            // lstFVSTablesList
            // 
            this.lstFVSTablesList.FormattingEnabled = true;
            this.lstFVSTablesList.ItemHeight = 16;
            this.lstFVSTablesList.Location = new System.Drawing.Point(6, 21);
            this.lstFVSTablesList.Name = "lstFVSTablesList";
            this.lstFVSTablesList.Size = new System.Drawing.Size(181, 100);
            this.lstFVSTablesList.TabIndex = 70;
            this.lstFVSTablesList.SelectedIndexChanged += new System.EventHandler(this.lstFVSTablesList_SelectedIndexChanged);
            // 
            // LblSelectedVariable
            // 
            this.LblSelectedVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblSelectedVariable.Location = new System.Drawing.Point(157, 145);
            this.LblSelectedVariable.Name = "LblSelectedVariable";
            this.LblSelectedVariable.Size = new System.Drawing.Size(264, 24);
            this.LblSelectedVariable.TabIndex = 69;
            this.LblSelectedVariable.Text = "FVS_POTFIRE.PTorch_Sev";
            // 
            // lblSelectedFVSVariable
            // 
            this.lblSelectedFVSVariable.Location = new System.Drawing.Point(11, 144);
            this.lblSelectedFVSVariable.Name = "lblSelectedFVSVariable";
            this.lblSelectedFVSVariable.Size = new System.Drawing.Size(151, 24);
            this.lblSelectedFVSVariable.TabIndex = 68;
            this.lblSelectedFVSVariable.Text = "Selected FVS Variable:";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(866, 32);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "Calculated Variables";
            // 
            // uc_core_scenario_weighted_average
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_core_scenario_weighted_average";
            this.Size = new System.Drawing.Size(872, 2000);
            this.groupBox1.ResumeLayout(false);
            this.grpBoxEconomicVariable.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgEcon)).EndInit();
            this.groupBox8.ResumeLayout(false);
            this.grpboxSummary.ResumeLayout(false);
            this.pnlSummary.ResumeLayout(false);
            this.grpboxDetails.ResumeLayout(false);
            this.pnlDetails.ResumeLayout(false);
            this.pnlDetails.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).EndInit();
            this.grpBoxFvsBaseline.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

        public void loadvalues()
        {
            this.m_intError = 0;
            this.m_strError = "";

            string strDestinationLinkDir = this.m_oEnv.strTempDir;
            //used to get the temporary random file name
            utils objUtils = new utils();
            //get temporary mdb file
            string strTempMDB = objUtils.getRandomFile(strDestinationLinkDir, "accdb");

            //create a temporary mdb that will contain all the links to the FVS_XXX_PREPOST_SEQNUM_MATRIX tables
            //in all of the FVS\DATA\VARIANT\FVSOUT_VARIANT_RXPACKAGE-RXCYCLE1-RXCYCLE2-RXCYCLE3-RXCYCLE4_BIOSUM.ACCDB files 
            dao_data_access oDao = new dao_data_access();
            oDao.CreateMDB(strTempMDB);

            // Load project data sources table
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() 
                + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();
            // Link to plot table
            int intTable = oDs.getValidTableNameRow("Plot");
            string strDirectoryPath = oDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            string strPlotTable = oDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            oDao.CreateTableLink(strTempMDB, strPlotTable, strDirectoryPath + "\\" + strFileName, 
                strPlotTable);
            intTable = oDs.getDataSourceTableNameRow("Treatment Packages");
            strDirectoryPath = oDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            string strRxPkgTable = oDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            oDao.CreateTableLink(strTempMDB, strRxPkgTable, strDirectoryPath + "\\" + strFileName,
                strRxPkgTable);

            //@ToDo: may want to move this to class-level if we need ado elsewhere
            ado_data_access oAdo = new ado_data_access();
            oAdo.m_strSQL = Queries.FVS.GetFVSVariantRxPackageSQL(strPlotTable, strRxPkgTable);
            oAdo.OpenConnection(oAdo.getMDBConnString(strTempMDB, "", ""));
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            //@ToDo: Choose table name based on FVS variable selected
            string strSeqNumTable = "FVS_SUMMARY_PREPOST_SEQNUM_MATRIX";
            RxTools oRxTools = new RxTools();
            string strFvsDirectory = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()
                + "\\fvs\\data";
            System.Collections.Generic.IList<string> lstSeqNumTables = new System.Collections.Generic.List<string>();
            while (oAdo.m_OleDbDataReader.Read())
            {
                string strVariant = oAdo.m_OleDbDataReader["fvs_variant"].ToString().Trim();
                string strPackage = oAdo.m_OleDbDataReader["RxPackage"].ToString().Trim();

                string strOutMDBFile = oRxTools.GetRxPackageFvsOutDbFileName(oAdo.m_OleDbDataReader);
                string strACCDBFile = strOutMDBFile.Replace(".MDB", "_BIOSUM.ACCDB");
                string strOutDirAndFile = strFvsDirectory + "\\" + strVariant + "\\" + strACCDBFile.Trim();
                if (System.IO.File.Exists(strOutDirAndFile))
                {
                    string strLinkTableName = "SEQNUM_MATRIX_" + strVariant + "_" + strPackage;
                    oDao.CreateTableLink(strTempMDB, strLinkTableName, strOutDirAndFile,
                        strSeqNumTable);
                    if (m_intError == 0)
                        lstSeqNumTables.Add(strLinkTableName);

                    if (!cboFvsVariableBaselinePkg.Items.Contains(strPackage))
                        cboFvsVariableBaselinePkg.Items.Add(strPackage);
                }
            }
            //Create temporary table to populate datagrid
            string strViewTableName = "view_weights";
            frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFvsVariableWeightsReferenceTable(oAdo,
                oAdo.m_OleDbConnection, strViewTableName);
            string[] sqlWhereArray = new string[8];
            sqlWhereArray[0] = " WHERE CYCLE1_PRE_YN = 'Y'";
            sqlWhereArray[1] = " WHERE CYCLE1_POST_YN = 'Y'";
            sqlWhereArray[2] = " WHERE CYCLE2_PRE_YN = 'Y'";
            sqlWhereArray[3] = " WHERE CYCLE2_POST_YN = 'Y'";
            sqlWhereArray[4] = " WHERE CYCLE3_PRE_YN = 'Y'";
            sqlWhereArray[5] = " WHERE CYCLE3_POST_YN = 'Y'";
            sqlWhereArray[6] = " WHERE CYCLE4_PRE_YN = 'Y'";
            sqlWhereArray[7] = " WHERE CYCLE4_POST_YN = 'Y'";
            for (int i = 0; i < sqlWhereArray.Length; i++)
            {
                string strSql = "SELECT MIN(MinSeqNum) as MinSeqNum1, MIN(MinYear) as MinYear1 " +
                                "FROM ( ";
                foreach (string strTableName in lstSeqNumTables)
                {
                    strSql = strSql + "SELECT MIN(SEQNUM) as MinSeqNum, MIN(YEAR) as MinYear " +
                                      "FROM " + strTableName +
                                      sqlWhereArray[i] +
                                      " UNION ALL ";    
                }
                //Trim off trailing union
                strSql = strSql.Remove(strSql.LastIndexOf(" UNION ALL "));
                strSql = strSql + ")";
                oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSql);
                while (oAdo.m_OleDbDataReader.Read())
                {
                    string strPrePost = "";
                    string strRxCycle = "";
                    int intSeqNum = Convert.ToInt16(oAdo.m_OleDbDataReader["MinSeqNum1"]);
                    int intYear = Convert.ToInt16(oAdo.m_OleDbDataReader["MinYear1"]);
                    switch (i)
                    {
                        case 0:
                            strPrePost = "PRE";
                            strRxCycle = "1";
                            break;
                        case 1:
                            strPrePost = "POST";
                            strRxCycle = "1";
                            break;
                        case 2:
                            strPrePost = "PRE";
                            strRxCycle = "2";
                            break;
                        case 3:
                            strPrePost = "POST";
                            strRxCycle = "2";
                            break;
                        case 4:
                            strPrePost = "PRE";
                            strRxCycle = "3";
                            break;
                        case 5:
                            strPrePost = "POST";
                            strRxCycle = "3";
                            break;
                        case 6:
                            strPrePost = "PRE";
                            strRxCycle = "4";
                            break;
                        case 7:
                            strPrePost = "POST";
                            strRxCycle = "4";
                            break;
                    }
                    if (!String.IsNullOrEmpty(strPrePost))
                    {
                        string insertSql = "INSERT INTO " + strViewTableName +
                                           " VALUES('" + strPrePost + "','" + strRxCycle +
                                           "'," + intYear + "," + intSeqNum + ",0)";
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, insertSql);
                    }
                }

            }

            oAdo.m_DataSet = new DataSet("view_weights");
            oAdo.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
            oAdo.m_strSQL = "select * from " + strViewTableName;
            this.m_dtTableSchema = oAdo.getTableSchema(oAdo.m_OleDbConnection,
                                                       oAdo.m_OleDbTransaction,
                                                       oAdo.m_strSQL);
            if (oAdo.m_intError == 0)
			{
				oAdo.m_OleDbCommand = oAdo.m_OleDbConnection.CreateCommand();
				oAdo.m_OleDbCommand.CommandText = oAdo.m_strSQL;
				oAdo.m_OleDbDataAdapter.SelectCommand = oAdo.m_OleDbCommand;
				oAdo.m_OleDbDataAdapter.SelectCommand.Transaction = oAdo.m_OleDbTransaction;
				try 
				{

					oAdo.m_OleDbDataAdapter.Fill(oAdo.m_DataSet,"view_weights");
                    this.m_dv = new DataView(oAdo.m_DataSet.Tables["view_weights"]);
				
					this.m_dv.AllowNew = false;       //cannot append new records
					this.m_dv.AllowDelete = false;    //cannot delete records
					this.m_dv.AllowEdit = true;
                    this.m_dg.CaptionText = "view_weights";
					m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;
					/***********************************************************************************
					 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
					 ***********************************************************************************/
                    WeightedAverage_DataGridColoredTextBoxColumn aColumnTextColumn;


					/***************************************************************
					 **custom define the grid style
					 ***************************************************************/
					DataGridTableStyle tableStyle = new DataGridTableStyle();

					/***********************************************************************
					 **map the data grid table style to the scenario rx intensity dataset
					 ***********************************************************************/
                    tableStyle.MappingName = "view_weights";
					tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
					tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
					tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
					tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
					
					
   
					/******************************************************************************
					 **since the dataset has things like field name and number of columns,
					 **we will use those to create new columnstyles for the columns in our grid
					 ******************************************************************************/
                    //get the number of columns from the view_weights data set
                    int numCols = oAdo.m_DataSet.Tables["view_weights"].Columns.Count;      
                    
                    /************************************************
					 **loop through all the columns in the dataset	
					 ************************************************/
                    string strColumnName;
                    for(int i = 0; i < numCols; ++i)
					{
                        strColumnName = oAdo.m_DataSet.Tables["view_weights"].Columns[i].ColumnName;

						/***********************************
						**all columns are read-only except weight
						***********************************/
                        if (strColumnName.Trim().ToUpper() == "WEIGHT")
                        {
                            /******************************************************************
                            **create a new instance of the DataGridColoredTextBoxColumn class
                            ******************************************************************/
                            aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(true, true, this);
                            aColumnTextColumn.Format = "#0.00";
                            aColumnTextColumn.ReadOnly = false;
                        }
                        else
                        {
                            /******************************************************************
                            **create a new instance of the DataGridColoredTextBoxColumn class
                            ******************************************************************/
                            aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(false, false, this);
                            aColumnTextColumn.ReadOnly = true;
                        }
						aColumnTextColumn.HeaderText = strColumnName;
				 				    
						/********************************************************************
						 **assign the mappingname property the data sets column name
						 ********************************************************************/
						aColumnTextColumn.MappingName = strColumnName;

                        /********************************************************************
						 **add the datagridcoloredtextboxcolumn object to the data grid 
						 **table style object
						 ********************************************************************/
						tableStyle.GridColumnStyles.Add(aColumnTextColumn);

                        /**********************************
                         * Hide pre_or_post column
                         * *******************************/
                        if (strColumnName.Equals("pre_or_post"))
                            tableStyle.GridColumnStyles.Remove(aColumnTextColumn);
					}
					/*********************************************************************
					 ** make the dataGrid use our new tablestyle and bind it to our table
					 *********************************************************************/
					if (frmMain.g_oGridViewFont != null) this.m_dg.Font = frmMain.g_oGridViewFont;
					this.m_dg.TableStyles.Clear();
					this.m_dg.TableStyles.Add(tableStyle);
                    //this.m_dg.CaptionText = strCaption;
                    this.m_dg.DataSource = this.m_dv;  
					this.m_dg.Expand(-1);
				}
				catch (Exception e)
				{
					MessageBox.Show(e.Message,"Table",MessageBoxButtons.OK,MessageBoxIcon.Error);
					this.m_intError=-1;
                    oAdo.m_OleDbConnection.Close();
                    oAdo.m_OleDbConnection = null;
                    oAdo.m_DataSet.Clear();
                    oAdo.m_DataSet = null;
                    oAdo.m_OleDbDataAdapter.Dispose();
                    oAdo.m_OleDbDataAdapter = null;
                    oAdo = null;
					return;

				}

			}

            m_dictFVSTables = m_oCoreAnalysisScenarioTools.LoadFvsTablesAndVariables(oAdo);
            foreach (string strKey in m_dictFVSTables.Keys)
            {
                lstFVSTablesList.Items.Add(strKey);
            }
            
            //Set Defaults
            if (cboFvsVariableBaselinePkg.SelectedIndex < 0)
            {
                //Set to last package as that is usually the grow-only package
                cboFvsVariableBaselinePkg.SelectedIndex = cboFvsVariableBaselinePkg.Items.Count - 1;
            }

            
            
            
            if (oAdo != null)
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
            
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }

		public int savevalues()
		{
			int x;
			ado_data_access oAdo = new ado_data_access();
			string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
			if (oAdo.m_intError==0)
			{


			}
			
            return 1;
		}

		protected void SendKeyStrokes(System.Windows.Forms.TextBox p_oTextBox, string strKeyStrokes)
		{
			try 
			{
				p_oTextBox.Focus();
				System.Windows.Forms.SendKeys.Send(strKeyStrokes);
				p_oTextBox.Refresh();
			}
			catch  (Exception caught)
			{
               MessageBox.Show("SendKeyStrokes Method Failed With This Message:" + caught.Message);
			}

		}

		
		
		
			
		

		protected void NextButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
		{
           p_oGb.Controls.Add(p_oButton);
		   p_oButton.Left = p_oGb.Width - p_oButton.Width - 5;
		   p_oButton.Top = p_oGb.Height - p_oButton.Height - 5;
		   p_oButton.Name = strButtonName;	
		}
		protected void PrevButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
		{
			p_oGb.Controls.Add(p_oButton);
			p_oButton.Top = this.btnNext.Top;
			p_oButton.Height = this.btnNext.Height;
			p_oButton.Width = this.btnNext.Width;
			p_oButton.Left = this.btnNext.Left - p_oButton.Width;
			p_oButton.Name = strButtonName;	
		}
		
		

		
		public void main_resize()
		{
			
		}		

		private void btnFVSVariablesPrePostVariableValue_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnFVSVariablesPrePostVariableClearAll_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostVariableNext_Click(object sender, System.EventArgs e)
		{
			
		}
		private void RollBack()
		{
			RollBack_variable1();
			RollBack_variable2();
			RollBack_variable3();
			RollBack_SqlBetter();
			RollBack_SqlWorse();
			RollBack_Overall();
		}
		private void RollBack_variable1()
		{
		}
		private void RollBack_variable2()
		{
		}
		private void RollBack_variable3()
		{
		}
		private void RollBack_SqlBetter()
		{
		}
		private void RollBack_SqlWorse()
		{
		}
		private void RollBack_Overall()
		{
		}

		private void grpboxFVSVariablesPrePostVariable_Resize(object sender, System.EventArgs e)
		{
			
		}

		private void grpboxFVSVariablesPrePostExpression_Resize(object sender, System.EventArgs e)
		{
		}

		private void grpboxFVSVariablesPrePostVariableValues_Resize(object sender, System.EventArgs e)
		{

		}

		private void btnFVSVariablesPrePostExpressionPrevious_Click(object sender, System.EventArgs e)
		{
			
			
			
		}

	


		private void Go()
		{

		}
		
		
		private void ShowGroupBox(string p_strName)
		{
			int x;
			//System.Windows.Forms.Control oControl;
			for (x=0;x<=groupBox1.Controls.Count-1;x++)
			{
				if (groupBox1.Controls[x].Name.Substring(0,3)=="grp")
				{
					if (p_strName.Trim().ToUpper() == 
						groupBox1.Controls[x].Name.Trim().ToUpper())
					{
						groupBox1.Controls[x].Show();
					}
					else
					{
						groupBox1.Controls[x].Hide();
					}
				}
			}
		}

		private void btnFVSVariablesPrePostValuesButtonsEdit_Click(object sender, System.EventArgs e)
		{

		}


		private void btnFVSVariablesPrePost2Overall_Click(object sender, System.EventArgs e)
		{
		}

		private void btnFVSVariablesPrePostValuesButtonsClear_Click(object sender, System.EventArgs e)
		{
		}
		

		

		private void grpboxFVSVariablesPrePost_Resize(object sender, System.EventArgs e)
		{
			   
		}

		

		private void lvFVSVariablesPrePostValues_DoubleClick(object sender, System.EventArgs e)
		{
		
		}

		private void pnlFVSVariablesPrePostExpression_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		private void pnlFVSVariablesPrePostVariable_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		private string ValidateNumeric(string p_strValue)
		{
            string strValue = p_strValue.Replace("$", "");
            strValue = strValue.Replace(",", "");
			try
			{
               double dbl = Convert.ToDouble(strValue);
			}
			catch 
			{
				return "0";
			}
			return strValue;
		}


		private void groupBox1_Leave(object sender, System.EventArgs e)
		{
			
		}
		private void EnableTabs(bool p_bEnable)
		{
            int x;
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlScenario,"tbdesc,tbnotes,tbdatasources",p_bEnable);
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlRules,"tbpsites,tbowners,tbcost,tbtreatmentintensity,tbfilterplots,tbrun",p_bEnable);
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlFVSVariables,"tbeffective,tbtiebreaker",p_bEnable);
            for (x = 0; x <= ReferenceCoreScenarioForm.tlbScenario.Buttons.Count - 1; x++)
            {
                ReferenceCoreScenarioForm.tlbScenario.Buttons[x].Enabled = p_bEnable;
            }
            frmMain.g_oFrmMain.grpboxLeft.Enabled = p_bEnable;
            frmMain.g_oFrmMain.tlbMain.Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[0].Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[1].Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[2].Enabled = p_bEnable;

		}

		private void btnOptimizationAudit_Click(object sender, System.EventArgs e)
		{
			this.DisplayAuditMessage=true;
			Audit();
		}
		public void Audit()
		{
			
			
			int x;
			this.m_intError=0;
			m_strError="";
			if (DisplayAuditMessage)
			{
				this.m_strError="Audit Results \r\n";
				this.m_strError=m_strError + "-------------\r\n\r\n";
			}

			
			if (DisplayAuditMessage)
			{
				if (m_intError==0) this.m_strError=m_strError + "Passed Audit";
				else m_strError = m_strError + "\r\n\r\n" + "Failed Audit";
				MessageBox.Show(m_strError,"FIA Biosum");
			}

		}
		
		public bool DisplayAuditMessage
		{
			get {return _bDisplayAuditMsg;}
			set {_bDisplayAuditMsg=value;}
		}
		public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		public FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.Variables ReferenceFVSVariables
		{
			get {return this._oCurVar;}
			set {_oCurVar=value;}
		}
		public FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker ReferenceTieBreaker
		{
			get {return _uc_tiebreaker;}
			set {_uc_tiebreaker=value;}
		}

        private void btnCancelSummary_Click(object sender, EventArgs e)
        {

			this.ParentForm.Close();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            this.grpboxSummary.Hide();
            this.grpboxDetails.Show();
        }

        private void btnNewEcon_Click(object sender, EventArgs e)
        {
            this.grpboxSummary.Hide();
            this.grpBoxEconomicVariable.Show();
        }

        private void btnFvsDetailsCancel_Click(object sender, EventArgs e)
        {
            this.grpboxSummary.Show();
            this.grpboxDetails.Hide();
            //@ToDo: Add code to clear fields on details screen
        }


        private void m_lvEx_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            int x;
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    int intRowHt = this.m_lvEx.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    this.m_lvEx.Items[m_lvEx.TopItem.Index + (int)dblRow - 1].Selected = true;
                    this.m_oLvAlternateColors.DelegateListViewItem(m_lvEx.Items[m_lvEx.TopItem.Index + (int)dblRow - 1]);


                }
            }
            catch
            {
            }
        }

        private void m_lvEx_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (m_lvEx.SelectedItems.Count > 0)
                m_oLvAlternateColors.DelegateListViewItem(m_lvEx.SelectedItems[0]);
        }
        private void m_lvEx_ColumnClick(object sender, System.Windows.Forms.ColumnClickEventArgs e)
        {
            int x, y;

            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.m_lvEx.Sort();
            //reinitialize the alternate row colors
            for (x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
            {
                for (y = 0; y <= this.m_lvEx.Columns.Count - 1; y++)
                {
                    m_oLvAlternateColors.ListViewSubItem(this.m_lvEx.Items[x].Index, y, this.m_lvEx.Items[this.m_lvEx.Items[x].Index].SubItems[y], false);
                }
            }
        }

        private void btnEconDetailsCancel_Click(object sender, EventArgs e)
        {
            this.grpboxSummary.Show();
            this.grpBoxEconomicVariable.Hide();
            //@ToDo: Add code to clear fields on econ variable screen
        }

        public void SumWeights()
        {
            DataTable objDataTable = this.m_dv.Table;
            double dblSum = 0;
            double dblWeight = -1;
            foreach (DataRow row in objDataTable.Rows)
            {
                string strWeight = row["weight"].ToString();
                if (Double.TryParse(strWeight, out dblWeight))
                    dblSum = dblSum + dblWeight;
            }
            txtFvsVariableTotalWeight.Text = String.Format("{0:0.0#}", dblSum); 
        }

        protected void Grid_CurCellChange(object sender, EventArgs e)
        {
            //Only recalculate if we are leaving the weight column
            if (m_intPrevColumnIdx.Equals(3))
                this.SumWeights();
            m_intPrevColumnIdx = m_dg.CurrentCell.ColumnNumber;
        }

        private void lstFVSTablesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstFVSFieldsList.Items.Clear();
            this.LblSelectedVariable.Text = "Not Defined";
            if (this.lstFVSTablesList.SelectedIndex > -1)
            {
                System.Collections.Generic.IList<string> lstFields =
                    m_dictFVSTables[Convert.ToString(this.lstFVSTablesList.SelectedItem)];
                if (lstFields != null)
                {
                    foreach (string strField in lstFields)
                    {
                        lstFVSFieldsList.Items.Add(strField);
                    }
                }
            }
        }

        private void lstFVSFieldsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.LblSelectedVariable.Text = "Not Defined";
            //if (this.lstFVSFieldsList.SelectedIndex > -1)
            //{
            //    this.btnFVSVariablesOptimizationVariableValues.Enabled = true;
            //}
            //else
            //{
            //    this.btnFVSVariablesOptimizationVariableValues.Enabled = false;
            //}
        }

        private void btnFVSVariableValue_Click(object sender, EventArgs e)
        {
            if (this.lstFVSTablesList.SelectedItems.Count == 0 || this.lstFVSFieldsList.SelectedItems.Count == 0) return;
            this.LblSelectedVariable.Text =
                this.lstFVSTablesList.SelectedItems[0].ToString() + "." + this.lstFVSFieldsList.SelectedItems[0].ToString();
        }


    }

    public class WeightedAverage_DataGridColoredTextBoxColumn : DataGridTextBoxColumn
    {
        bool m_bEdit = false;
        FIA_Biosum_Manager.uc_core_scenario_weighted_average uc_core_scenario_weighted_average1;
        string m_strLastKey = "";
        bool m_bNumericOnly = false;


        public WeightedAverage_DataGridColoredTextBoxColumn(bool bEdit, bool bNumericOnly, 
            FIA_Biosum_Manager.uc_core_scenario_weighted_average p_uc)
        {
            this.m_bEdit = bEdit;
            this.m_bNumericOnly = bNumericOnly;
            this.uc_core_scenario_weighted_average1 = p_uc;
            this.TextBox.KeyDown += new KeyEventHandler(TextBox_KeyDown);
            this.TextBox.Leave += new EventHandler(TextBox_Leave);
            this.TextBox.Enter += new EventHandler(TextBox_Enter);
        }


        protected override void Paint(System.Drawing.Graphics g, System.Drawing.Rectangle bounds, System.Windows.Forms.CurrencyManager source, int rowNum, System.Drawing.Brush backBrush, System.Drawing.Brush foreBrush, bool alignToRight)
        {

            // color only the columns that can be edited by the user
            try
            {
                if (this.m_bEdit == true)
                {
                    backBrush = new System.Drawing.Drawing2D.LinearGradientBrush(bounds,
                        Color.FromArgb(255, 200, 200),
                        Color.FromArgb(128, 20, 20),
                        System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal);
                    foreBrush = new SolidBrush(Color.White);
                }
            }
            catch { /* empty catch */ }
            finally
            {
                try
                {
                    // make sure the base class gets called to do the drawing with
                    // the possibly changed brushes
                    base.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight);
                }
                catch
                {
                }
            }
        }
        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("textchange");
        }
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (this.m_bEdit == true)
            {

                if (this.m_bNumericOnly == true)
                {
                    if (Char.IsDigit((char)e.KeyValue) || (e.KeyCode == Keys.OemPeriod && this.Format.IndexOf(".", 0) >= 0 && this.TextBox.Text.IndexOf(".", 0) < 0))
                    {
                        this.m_strLastKey = Convert.ToString(e.KeyValue);
                        if (this.uc_core_scenario_weighted_average1.btnFvsCalculate.Enabled == false) this.uc_core_scenario_weighted_average1.btnFvsCalculate.Enabled = true;
                    }
                    else
                    {
                        if (e.KeyCode == Keys.Back)
                        {
                            this.m_strLastKey = Convert.ToString(e.KeyValue);
                            if (this.uc_core_scenario_weighted_average1.btnFvsCalculate.Enabled == false) this.uc_core_scenario_weighted_average1.btnFvsCalculate.Enabled = true;
                        }
                        else
                        {
                            e.Handled = true;
                            SendKeys.Send("{BACKSPACE}");
                        }
                    }

                }
                else
                {
                    this.m_strLastKey = Convert.ToString(e.KeyValue);
                    if (this.uc_core_scenario_weighted_average1.btnFvsCalculate.Enabled == false) this.uc_core_scenario_weighted_average1.btnFvsCalculate.Enabled = true;
                }


            }



        }
        private void TextBox_Enter(object sender, EventArgs e)
        {
            this.m_strLastKey = "";
        }
        private void TextBox_Leave(object sender, EventArgs e)
        {
            if (this.m_bEdit == true)
            {
                if (this.m_strLastKey.Trim().Length > 0)
                {
                    if (this.uc_core_scenario_weighted_average1.btnFvsCalculate.Enabled == false) this.uc_core_scenario_weighted_average1.btnFvsCalculate.Enabled = true;
                }
            }
        }

    }



}
