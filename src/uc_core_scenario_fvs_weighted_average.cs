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
		public System.Data.DataTable m_DataTable;
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
		private FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker _uc_tiebreaker;

		
		private int m_intCurVar=-1;
		int m_intCurVariableDefinitionStepCount=1;
		string[] m_strUserNavigation=null;
		private System.Windows.Forms.Panel pnlFVSVariablesPrePostVariable;
        private System.Windows.Forms.GroupBox grpboxOptimization;
		private System.Windows.Forms.Label lblOptimizationVariable;
		private System.Windows.Forms.Button btnOptimiztionDone;
		private System.Windows.Forms.Button btnOptimiztionCancel;
		private System.Windows.Forms.RadioButton rdoOptimizationMaximum;
		private System.Windows.Forms.RadioButton rdoOptimizationMinimum;
		private System.Windows.Forms.TextBox txtOptimizationValue;
		private System.Windows.Forms.ComboBox cmbOptimizationOperator;

		public bool m_bSave=false;


		const int COLUMN_CHECKBOX=0;
		const int COLUMN_OPTIMIZE_VARIABLE=1;
		const int COLUMN_FVS_VARIABLE=2;
		const int COLUMN_VALUESOURCE=3;
		const int COLUMN_MAXMIN=4;
		const int COLUMN_USEFILTER=5;
		const int COLUMN_FILTER_OPERATOR=6;
		const int COLUMN_FILTER_VALUE=7;

		private System.Windows.Forms.GroupBox grpMaxMin;
		private System.Windows.Forms.GroupBox grpFilter;
		private System.Windows.Forms.CheckBox chkEnableFilter;
        private FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective.Variables _oCurVar;
		public bool m_bFirstTime=true;
		private bool _bDisplayAuditMsg=true;
        private bool m_bIgnoreListViewItemCheck = false;
		private System.Windows.Forms.GroupBox grpboxOptimizationSettings;
		private System.Windows.Forms.GroupBox grpboxOptimizationSettingsPostPre;
		private System.Windows.Forms.ComboBox cmbOptimizationSettingsPostPreValue;
		private System.Windows.Forms.GroupBox grpboxOptimizationFVSVariable;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesOptimizationVariableValues;
		private System.Windows.Forms.Button btnFVSVariablesOptimizationVariableValues;
		private System.Windows.Forms.ListBox lstFVSVariablesOptimizationVariableValues;
		private System.Windows.Forms.GroupBox grpFVSVariablesOptimizationVariableValuesSelected;
		private System.Windows.Forms.Label lblFVSVariablesOptimizationVariableValuesSelected;
		private System.Windows.Forms.Button btnOptimiztionPrev;
		private System.Windows.Forms.Button btnOptimizationFVSVariableClear;
		private System.Windows.Forms.Button btnOptimizationFVSVariableDone;
		private System.Windows.Forms.Button btnOptimizationFVSVariableCancel;
		private System.Windows.Forms.Button btnOptimizationFVSVariableNext;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors=new ListViewAlternateBackgroundColors();
        private ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private Label label1;
        private string m_strLastValue = "";
        private FIA_Biosum_Manager.frmMain m_frmMain;
        public int m_DialogHt;
        public Panel pnlOptimization;
        private TextBox textBox17;
        private Label label7;
        private Button button2;
        private Button button1;
        private Panel panel2;
        private Label label6;
        private TextBox textBox9;
        private TextBox textBox10;
        private TextBox textBox11;
        private TextBox textBox12;
        private TextBox textBox13;
        private TextBox textBox14;
        private TextBox textBox15;
        private TextBox textBox16;
        private TextBox textBox8;
        private TextBox textBox7;
        private TextBox textBox6;
        private TextBox textBox5;
        private TextBox textBox4;
        private TextBox textBox3;
        private TextBox textBox2;
        private TextBox textBox1;
        private Label label4;
        private Label label5;
        private Label label3;
        private Label label2;
        private Button btnNewSelectViewCancel;
        private GroupBox groupBox4;
        private ComboBox comboBox1;
        private GroupBox groupBox3;
        private ListBox listBox1;
        private GroupBox groupBox2;
        private ListBox lstFVSVariableTables;
        private Label LblSelectedVariable;
        private Label lblSelectedFVSVariable;
        public int m_DialogWd;

        public uc_core_scenario_weighted_average(FIA_Biosum_Manager.frmMain p_frmMain)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_oUtils = new utils();
            this.m_frmMain = p_frmMain;
			this.m_strCurrentIndexTypeAndClass="";

			
			this.grpboxOptimizationSettings.Top = grpboxOptimization.Top;
			this.grpboxOptimizationSettings.Left = this.grpboxOptimization.Left;
			this.grpboxOptimizationSettings.Height = this.grpboxOptimization.Height;
			this.grpboxOptimizationSettings.Width = this.grpboxOptimization.Width;
			this.grpboxOptimizationFVSVariable.Top = grpboxOptimization.Top;
			this.grpboxOptimizationFVSVariable.Left = this.grpboxOptimization.Left;
			this.grpboxOptimizationFVSVariable.Height = this.grpboxOptimization.Height;
			this.grpboxOptimizationFVSVariable.Width = this.grpboxOptimization.Width;



			this.grpboxOptimizationFVSVariable.Hide();


            m_oValidate.RoundDecimalLength = 0;
            m_oValidate.Money = false;
            m_oValidate.NullsAllowed = false;
            m_oValidate.TestForMaxMin = false;
            m_oValidate.MinValue = -1000;
            m_oValidate.TestForMin = true;
         
			

			

			

			

			// TODO: Add any initialization after the InitializeComponent call
            this.m_DialogWd = this.Width + 25;
            this.m_DialogHt = this.pnlOptimization.Height + 100;
            this.Height = m_DialogHt -40;
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
            this.grpboxOptimizationFVSVariable = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.grpboxFVSVariablesOptimizationVariableValues = new System.Windows.Forms.GroupBox();
            this.btnFVSVariablesOptimizationVariableValues = new System.Windows.Forms.Button();
            this.lstFVSVariablesOptimizationVariableValues = new System.Windows.Forms.ListBox();
            this.grpFVSVariablesOptimizationVariableValuesSelected = new System.Windows.Forms.GroupBox();
            this.lblFVSVariablesOptimizationVariableValuesSelected = new System.Windows.Forms.Label();
            this.btnOptimizationFVSVariableClear = new System.Windows.Forms.Button();
            this.btnOptimizationFVSVariableDone = new System.Windows.Forms.Button();
            this.btnOptimizationFVSVariableCancel = new System.Windows.Forms.Button();
            this.btnOptimizationFVSVariableNext = new System.Windows.Forms.Button();
            this.grpboxOptimization = new System.Windows.Forms.GroupBox();
            this.grpboxOptimizationSettings = new System.Windows.Forms.GroupBox();
            this.pnlFVSVariablesPrePostVariable = new System.Windows.Forms.Panel();
            this.btnOptimiztionPrev = new System.Windows.Forms.Button();
            this.grpboxOptimizationSettingsPostPre = new System.Windows.Forms.GroupBox();
            this.cmbOptimizationSettingsPostPreValue = new System.Windows.Forms.ComboBox();
            this.grpFilter = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkEnableFilter = new System.Windows.Forms.CheckBox();
            this.cmbOptimizationOperator = new System.Windows.Forms.ComboBox();
            this.txtOptimizationValue = new System.Windows.Forms.TextBox();
            this.grpMaxMin = new System.Windows.Forms.GroupBox();
            this.rdoOptimizationMinimum = new System.Windows.Forms.RadioButton();
            this.rdoOptimizationMaximum = new System.Windows.Forms.RadioButton();
            this.lblOptimizationVariable = new System.Windows.Forms.Label();
            this.btnOptimiztionDone = new System.Windows.Forms.Button();
            this.btnOptimiztionCancel = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblSelectedFVSVariable = new System.Windows.Forms.Label();
            this.LblSelectedVariable = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lstFVSVariableTables = new System.Windows.Forms.ListBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.btnNewSelectViewCancel = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.textBox16 = new System.Windows.Forms.TextBox();
            this.textBox15 = new System.Windows.Forms.TextBox();
            this.textBox14 = new System.Windows.Forms.TextBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.pnlOptimization = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox17 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.grpboxOptimizationFVSVariable.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grpboxFVSVariablesOptimizationVariableValues.SuspendLayout();
            this.grpFVSVariablesOptimizationVariableValuesSelected.SuspendLayout();
            this.grpboxOptimization.SuspendLayout();
            this.grpboxOptimizationSettings.SuspendLayout();
            this.pnlFVSVariablesPrePostVariable.SuspendLayout();
            this.grpboxOptimizationSettingsPostPre.SuspendLayout();
            this.grpFilter.SuspendLayout();
            this.grpMaxMin.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.panel2.SuspendLayout();
            this.pnlOptimization.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.grpboxOptimizationFVSVariable);
            this.groupBox1.Controls.Add(this.grpboxOptimization);
            this.groupBox1.Controls.Add(this.grpboxOptimizationSettings);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(872, 2000);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Leave += new System.EventHandler(this.groupBox1_Leave);
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // grpboxOptimizationFVSVariable
            // 
            this.grpboxOptimizationFVSVariable.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxOptimizationFVSVariable.Controls.Add(this.panel1);
            this.grpboxOptimizationFVSVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxOptimizationFVSVariable.ForeColor = System.Drawing.Color.Black;
            this.grpboxOptimizationFVSVariable.Location = new System.Drawing.Point(0, 528);
            this.grpboxOptimizationFVSVariable.Name = "grpboxOptimizationFVSVariable";
            this.grpboxOptimizationFVSVariable.Size = new System.Drawing.Size(872, 448);
            this.grpboxOptimizationFVSVariable.TabIndex = 35;
            this.grpboxOptimizationFVSVariable.TabStop = false;
            this.grpboxOptimizationFVSVariable.Text = "FVS Variable";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.grpboxFVSVariablesOptimizationVariableValues);
            this.panel1.Controls.Add(this.grpFVSVariablesOptimizationVariableValuesSelected);
            this.panel1.Controls.Add(this.btnOptimizationFVSVariableClear);
            this.panel1.Controls.Add(this.btnOptimizationFVSVariableDone);
            this.panel1.Controls.Add(this.btnOptimizationFVSVariableCancel);
            this.panel1.Controls.Add(this.btnOptimizationFVSVariableNext);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(866, 427);
            this.panel1.TabIndex = 12;
            // 
            // grpboxFVSVariablesOptimizationVariableValues
            // 
            this.grpboxFVSVariablesOptimizationVariableValues.Controls.Add(this.btnFVSVariablesOptimizationVariableValues);
            this.grpboxFVSVariablesOptimizationVariableValues.Controls.Add(this.lstFVSVariablesOptimizationVariableValues);
            this.grpboxFVSVariablesOptimizationVariableValues.Location = new System.Drawing.Point(8, 16);
            this.grpboxFVSVariablesOptimizationVariableValues.Name = "grpboxFVSVariablesOptimizationVariableValues";
            this.grpboxFVSVariablesOptimizationVariableValues.Size = new System.Drawing.Size(816, 280);
            this.grpboxFVSVariablesOptimizationVariableValues.TabIndex = 0;
            this.grpboxFVSVariablesOptimizationVariableValues.TabStop = false;
            this.grpboxFVSVariablesOptimizationVariableValues.Text = "Optimization Variable List";
            // 
            // btnFVSVariablesOptimizationVariableValues
            // 
            this.btnFVSVariablesOptimizationVariableValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesOptimizationVariableValues.Location = new System.Drawing.Point(448, 32);
            this.btnFVSVariablesOptimizationVariableValues.Name = "btnFVSVariablesOptimizationVariableValues";
            this.btnFVSVariablesOptimizationVariableValues.Size = new System.Drawing.Size(184, 144);
            this.btnFVSVariablesOptimizationVariableValues.TabIndex = 1;
            this.btnFVSVariablesOptimizationVariableValues.Text = "Select";
            this.btnFVSVariablesOptimizationVariableValues.Click += new System.EventHandler(this.btnFVSVariablesOptimizationVariableValues_Click);
            // 
            // lstFVSVariablesOptimizationVariableValues
            // 
            this.lstFVSVariablesOptimizationVariableValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstFVSVariablesOptimizationVariableValues.ItemHeight = 16;
            this.lstFVSVariablesOptimizationVariableValues.Location = new System.Drawing.Point(8, 16);
            this.lstFVSVariablesOptimizationVariableValues.Name = "lstFVSVariablesOptimizationVariableValues";
            this.lstFVSVariablesOptimizationVariableValues.Size = new System.Drawing.Size(424, 244);
            this.lstFVSVariablesOptimizationVariableValues.TabIndex = 0;
            // 
            // grpFVSVariablesOptimizationVariableValuesSelected
            // 
            this.grpFVSVariablesOptimizationVariableValuesSelected.Controls.Add(this.lblFVSVariablesOptimizationVariableValuesSelected);
            this.grpFVSVariablesOptimizationVariableValuesSelected.Location = new System.Drawing.Point(16, 304);
            this.grpFVSVariablesOptimizationVariableValuesSelected.Name = "grpFVSVariablesOptimizationVariableValuesSelected";
            this.grpFVSVariablesOptimizationVariableValuesSelected.Size = new System.Drawing.Size(816, 51);
            this.grpFVSVariablesOptimizationVariableValuesSelected.TabIndex = 4;
            this.grpFVSVariablesOptimizationVariableValuesSelected.TabStop = false;
            this.grpFVSVariablesOptimizationVariableValuesSelected.Text = "Selected Optimization Variable";
            // 
            // lblFVSVariablesOptimizationVariableValuesSelected
            // 
            this.lblFVSVariablesOptimizationVariableValuesSelected.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFVSVariablesOptimizationVariableValuesSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFVSVariablesOptimizationVariableValuesSelected.Location = new System.Drawing.Point(3, 18);
            this.lblFVSVariablesOptimizationVariableValuesSelected.Name = "lblFVSVariablesOptimizationVariableValuesSelected";
            this.lblFVSVariablesOptimizationVariableValuesSelected.Size = new System.Drawing.Size(810, 30);
            this.lblFVSVariablesOptimizationVariableValuesSelected.TabIndex = 2;
            this.lblFVSVariablesOptimizationVariableValuesSelected.Text = "Not Defined";
            // 
            // btnOptimizationFVSVariableClear
            // 
            this.btnOptimizationFVSVariableClear.Location = new System.Drawing.Point(24, 376);
            this.btnOptimizationFVSVariableClear.Name = "btnOptimizationFVSVariableClear";
            this.btnOptimizationFVSVariableClear.Size = new System.Drawing.Size(72, 40);
            this.btnOptimizationFVSVariableClear.TabIndex = 5;
            this.btnOptimizationFVSVariableClear.Text = "Clear";
            // 
            // btnOptimizationFVSVariableDone
            // 
            this.btnOptimizationFVSVariableDone.Location = new System.Drawing.Point(352, 376);
            this.btnOptimizationFVSVariableDone.Name = "btnOptimizationFVSVariableDone";
            this.btnOptimizationFVSVariableDone.Size = new System.Drawing.Size(88, 40);
            this.btnOptimizationFVSVariableDone.TabIndex = 11;
            this.btnOptimizationFVSVariableDone.Text = "Done";
            // 
            // btnOptimizationFVSVariableCancel
            // 
            this.btnOptimizationFVSVariableCancel.Location = new System.Drawing.Point(440, 376);
            this.btnOptimizationFVSVariableCancel.Name = "btnOptimizationFVSVariableCancel";
            this.btnOptimizationFVSVariableCancel.Size = new System.Drawing.Size(88, 40);
            this.btnOptimizationFVSVariableCancel.TabIndex = 9;
            this.btnOptimizationFVSVariableCancel.Text = "Cancel";
            this.btnOptimizationFVSVariableCancel.Click += new System.EventHandler(this.btnOptimizationFVSVariableCancel_Click);
            // 
            // btnOptimizationFVSVariableNext
            // 
            this.btnOptimizationFVSVariableNext.Location = new System.Drawing.Point(616, 376);
            this.btnOptimizationFVSVariableNext.Name = "btnOptimizationFVSVariableNext";
            this.btnOptimizationFVSVariableNext.Size = new System.Drawing.Size(88, 40);
            this.btnOptimizationFVSVariableNext.TabIndex = 8;
            this.btnOptimizationFVSVariableNext.Text = "Next-->";
            this.btnOptimizationFVSVariableNext.Click += new System.EventHandler(this.btnOptimizationFVSVariableNext_Click);
            // 
            // grpboxOptimization
            // 
            this.grpboxOptimization.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxOptimization.Controls.Add(this.pnlOptimization);
            this.grpboxOptimization.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxOptimization.ForeColor = System.Drawing.Color.Black;
            this.grpboxOptimization.Location = new System.Drawing.Point(8, 48);
            this.grpboxOptimization.Name = "grpboxOptimization";
            this.grpboxOptimization.Size = new System.Drawing.Size(856, 472);
            this.grpboxOptimization.TabIndex = 32;
            this.grpboxOptimization.TabStop = false;
            this.grpboxOptimization.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePost_Resize);
            // 
            // grpboxOptimizationSettings
            // 
            this.grpboxOptimizationSettings.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxOptimizationSettings.Controls.Add(this.pnlFVSVariablesPrePostVariable);
            this.grpboxOptimizationSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxOptimizationSettings.ForeColor = System.Drawing.Color.Black;
            this.grpboxOptimizationSettings.Location = new System.Drawing.Point(16, 992);
            this.grpboxOptimizationSettings.Name = "grpboxOptimizationSettings";
            this.grpboxOptimizationSettings.Size = new System.Drawing.Size(856, 448);
            this.grpboxOptimizationSettings.TabIndex = 30;
            this.grpboxOptimizationSettings.TabStop = false;
            this.grpboxOptimizationSettings.Text = "Optimization Variable Settings";
            this.grpboxOptimizationSettings.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePostVariable_Resize);
            // 
            // pnlFVSVariablesPrePostVariable
            // 
            this.pnlFVSVariablesPrePostVariable.AutoScroll = true;
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnOptimiztionPrev);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpboxOptimizationSettingsPostPre);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpFilter);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpMaxMin);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.lblOptimizationVariable);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnOptimiztionDone);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnOptimiztionCancel);
            this.pnlFVSVariablesPrePostVariable.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFVSVariablesPrePostVariable.Location = new System.Drawing.Point(3, 18);
            this.pnlFVSVariablesPrePostVariable.Name = "pnlFVSVariablesPrePostVariable";
            this.pnlFVSVariablesPrePostVariable.Size = new System.Drawing.Size(850, 427);
            this.pnlFVSVariablesPrePostVariable.TabIndex = 12;
            this.pnlFVSVariablesPrePostVariable.Resize += new System.EventHandler(this.pnlFVSVariablesPrePostVariable_Resize);
            // 
            // btnOptimiztionPrev
            // 
            this.btnOptimiztionPrev.Location = new System.Drawing.Point(528, 376);
            this.btnOptimiztionPrev.Name = "btnOptimiztionPrev";
            this.btnOptimiztionPrev.Size = new System.Drawing.Size(88, 40);
            this.btnOptimiztionPrev.TabIndex = 21;
            this.btnOptimiztionPrev.Text = "<--Previous";
            this.btnOptimiztionPrev.Click += new System.EventHandler(this.btnOptimiztionPrev_Click);
            // 
            // grpboxOptimizationSettingsPostPre
            // 
            this.grpboxOptimizationSettingsPostPre.Controls.Add(this.cmbOptimizationSettingsPostPreValue);
            this.grpboxOptimizationSettingsPostPre.Location = new System.Drawing.Point(80, 64);
            this.grpboxOptimizationSettingsPostPre.Name = "grpboxOptimizationSettingsPostPre";
            this.grpboxOptimizationSettingsPostPre.Size = new System.Drawing.Size(344, 72);
            this.grpboxOptimizationSettingsPostPre.TabIndex = 20;
            this.grpboxOptimizationSettingsPostPre.TabStop = false;
            this.grpboxOptimizationSettingsPostPre.Text = "Post Treatment Variable Or Pre/Post Treatment Change";
            // 
            // cmbOptimizationSettingsPostPreValue
            // 
            this.cmbOptimizationSettingsPostPreValue.Items.AddRange(new object[] {
            "Post Value",
            "Post - Pre  Change Value"});
            this.cmbOptimizationSettingsPostPreValue.Location = new System.Drawing.Point(16, 40);
            this.cmbOptimizationSettingsPostPreValue.Name = "cmbOptimizationSettingsPostPreValue";
            this.cmbOptimizationSettingsPostPreValue.Size = new System.Drawing.Size(320, 24);
            this.cmbOptimizationSettingsPostPreValue.TabIndex = 0;
            this.cmbOptimizationSettingsPostPreValue.Text = "Post Value";
            // 
            // grpFilter
            // 
            this.grpFilter.Controls.Add(this.label1);
            this.grpFilter.Controls.Add(this.chkEnableFilter);
            this.grpFilter.Controls.Add(this.cmbOptimizationOperator);
            this.grpFilter.Controls.Add(this.txtOptimizationValue);
            this.grpFilter.Location = new System.Drawing.Point(80, 232);
            this.grpFilter.Name = "grpFilter";
            this.grpFilter.Size = new System.Drawing.Size(584, 64);
            this.grpFilter.TabIndex = 18;
            this.grpFilter.TabStop = false;
            this.grpFilter.Text = "Net Revenue Dollars Per Acre Filter Setting";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(326, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(24, 25);
            this.label1.TabIndex = 18;
            this.label1.Text = "$";
            // 
            // chkEnableFilter
            // 
            this.chkEnableFilter.Location = new System.Drawing.Point(48, 25);
            this.chkEnableFilter.Name = "chkEnableFilter";
            this.chkEnableFilter.Size = new System.Drawing.Size(112, 32);
            this.chkEnableFilter.TabIndex = 17;
            this.chkEnableFilter.Text = "Enable Filter";
            // 
            // cmbOptimizationOperator
            // 
            this.cmbOptimizationOperator.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOptimizationOperator.Items.AddRange(new object[] {
            ">",
            "<",
            ">=",
            "<=",
            "<>"});
            this.cmbOptimizationOperator.Location = new System.Drawing.Point(232, 25);
            this.cmbOptimizationOperator.Name = "cmbOptimizationOperator";
            this.cmbOptimizationOperator.Size = new System.Drawing.Size(88, 32);
            this.cmbOptimizationOperator.TabIndex = 16;
            this.cmbOptimizationOperator.Text = ">";
            // 
            // txtOptimizationValue
            // 
            this.txtOptimizationValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOptimizationValue.Location = new System.Drawing.Point(352, 25);
            this.txtOptimizationValue.Name = "txtOptimizationValue";
            this.txtOptimizationValue.Size = new System.Drawing.Size(200, 29);
            this.txtOptimizationValue.TabIndex = 15;
            this.txtOptimizationValue.Text = "0";
            this.txtOptimizationValue.Leave += new System.EventHandler(this.txtOptimizationValue_Leave);
            // 
            // grpMaxMin
            // 
            this.grpMaxMin.Controls.Add(this.rdoOptimizationMinimum);
            this.grpMaxMin.Controls.Add(this.rdoOptimizationMaximum);
            this.grpMaxMin.Location = new System.Drawing.Point(80, 160);
            this.grpMaxMin.Name = "grpMaxMin";
            this.grpMaxMin.Size = new System.Drawing.Size(464, 48);
            this.grpMaxMin.TabIndex = 17;
            this.grpMaxMin.TabStop = false;
            this.grpMaxMin.Text = "Aggregate Setting";
            // 
            // rdoOptimizationMinimum
            // 
            this.rdoOptimizationMinimum.Location = new System.Drawing.Point(256, 16);
            this.rdoOptimizationMinimum.Name = "rdoOptimizationMinimum";
            this.rdoOptimizationMinimum.Size = new System.Drawing.Size(176, 24);
            this.rdoOptimizationMinimum.TabIndex = 14;
            this.rdoOptimizationMinimum.Text = "Minimum";
            // 
            // rdoOptimizationMaximum
            // 
            this.rdoOptimizationMaximum.Checked = true;
            this.rdoOptimizationMaximum.Location = new System.Drawing.Point(32, 16);
            this.rdoOptimizationMaximum.Name = "rdoOptimizationMaximum";
            this.rdoOptimizationMaximum.Size = new System.Drawing.Size(176, 24);
            this.rdoOptimizationMaximum.TabIndex = 12;
            this.rdoOptimizationMaximum.TabStop = true;
            this.rdoOptimizationMaximum.Text = "Maximum";
            // 
            // lblOptimizationVariable
            // 
            this.lblOptimizationVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOptimizationVariable.Location = new System.Drawing.Point(24, 24);
            this.lblOptimizationVariable.Name = "lblOptimizationVariable";
            this.lblOptimizationVariable.Size = new System.Drawing.Size(472, 32);
            this.lblOptimizationVariable.TabIndex = 13;
            this.lblOptimizationVariable.Text = "Optimization Variable";
            // 
            // btnOptimiztionDone
            // 
            this.btnOptimiztionDone.Location = new System.Drawing.Point(352, 376);
            this.btnOptimiztionDone.Name = "btnOptimiztionDone";
            this.btnOptimiztionDone.Size = new System.Drawing.Size(88, 40);
            this.btnOptimiztionDone.TabIndex = 11;
            this.btnOptimiztionDone.Text = "Done";
            // 
            // btnOptimiztionCancel
            // 
            this.btnOptimiztionCancel.Location = new System.Drawing.Point(440, 376);
            this.btnOptimiztionCancel.Name = "btnOptimiztionCancel";
            this.btnOptimiztionCancel.Size = new System.Drawing.Size(88, 40);
            this.btnOptimiztionCancel.TabIndex = 9;
            this.btnOptimiztionCancel.Text = "Cancel";
            this.btnOptimiztionCancel.Click += new System.EventHandler(this.btnOptimiztionCancel_Click);
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
            this.lblTitle.Text = "Weighted Average";
            // 
            // lblSelectedFVSVariable
            // 
            this.lblSelectedFVSVariable.Location = new System.Drawing.Point(11, 191);
            this.lblSelectedFVSVariable.Name = "lblSelectedFVSVariable";
            this.lblSelectedFVSVariable.Size = new System.Drawing.Size(151, 24);
            this.lblSelectedFVSVariable.TabIndex = 68;
            this.lblSelectedFVSVariable.Text = "Selected FVS Variable:";
            // 
            // LblSelectedVariable
            // 
            this.LblSelectedVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblSelectedVariable.Location = new System.Drawing.Point(157, 192);
            this.LblSelectedVariable.Name = "LblSelectedVariable";
            this.LblSelectedVariable.Size = new System.Drawing.Size(302, 24);
            this.LblSelectedVariable.TabIndex = 69;
            this.LblSelectedVariable.Text = "FVS_POTFIRE.PTorch_Sev";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lstFVSVariableTables);
            this.groupBox2.Location = new System.Drawing.Point(6, 56);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 133);
            this.groupBox2.TabIndex = 71;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "FVS Variable Table(s)";
            // 
            // lstFVSVariableTables
            // 
            this.lstFVSVariableTables.FormattingEnabled = true;
            this.lstFVSVariableTables.ItemHeight = 16;
            this.lstFVSVariableTables.Items.AddRange(new object[] {
            "FVS_BURNREPORT",
            "FVS_CARBON",
            "FVS_COMPUTE",
            "FVS_HRV_CARBON",
            "FVS_POTFIRE",
            "FVS_STRCLASS",
            "FVS_SUMMARY"});
            this.lstFVSVariableTables.Location = new System.Drawing.Point(6, 21);
            this.lstFVSVariableTables.Name = "lstFVSVariableTables";
            this.lstFVSVariableTables.Size = new System.Drawing.Size(181, 100);
            this.lstFVSVariableTables.TabIndex = 70;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.listBox1);
            this.groupBox3.Location = new System.Drawing.Point(238, 56);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(200, 133);
            this.groupBox3.TabIndex = 72;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "FVS Variable(s)";
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 16;
            this.listBox1.Items.AddRange(new object[] {
            "Canopy_Density",
            "Canopy_Ht",
            "Crown_Index",
            "Fire_Type_Mod",
            "PTorch_Sev",
            "PTorch_Mod",
            "Torch_Index",
            ""});
            this.listBox1.Location = new System.Drawing.Point(6, 21);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(181, 100);
            this.listBox1.TabIndex = 70;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.comboBox1);
            this.groupBox4.Location = new System.Drawing.Point(8, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(108, 48);
            this.groupBox4.TabIndex = 74;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Baseline Rx";
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.Items.AddRange(new object[] {
            "999"});
            this.comboBox1.Location = new System.Drawing.Point(8, 18);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(72, 24);
            this.comboBox1.TabIndex = 77;
            // 
            // btnNewSelectViewCancel
            // 
            this.btnNewSelectViewCancel.Location = new System.Drawing.Point(653, 405);
            this.btnNewSelectViewCancel.Name = "btnNewSelectViewCancel";
            this.btnNewSelectViewCancel.Size = new System.Drawing.Size(64, 24);
            this.btnNewSelectViewCancel.TabIndex = 75;
            this.btnNewSelectViewCancel.Text = "Cancel";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.textBox9);
            this.panel2.Controls.Add(this.textBox10);
            this.panel2.Controls.Add(this.textBox11);
            this.panel2.Controls.Add(this.textBox12);
            this.panel2.Controls.Add(this.textBox13);
            this.panel2.Controls.Add(this.textBox14);
            this.panel2.Controls.Add(this.textBox15);
            this.panel2.Controls.Add(this.textBox16);
            this.panel2.Controls.Add(this.textBox8);
            this.panel2.Controls.Add(this.textBox7);
            this.panel2.Controls.Add(this.textBox6);
            this.panel2.Controls.Add(this.textBox5);
            this.panel2.Controls.Add(this.textBox4);
            this.panel2.Controls.Add(this.textBox3);
            this.panel2.Controls.Add(this.textBox2);
            this.panel2.Controls.Add(this.textBox1);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Location = new System.Drawing.Point(16, 219);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(293, 175);
            this.panel2.TabIndex = 76;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(18, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 24);
            this.label2.TabIndex = 69;
            this.label2.Text = "Year";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(86, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 24);
            this.label3.TabIndex = 70;
            this.label3.Text = "Weight";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(165, 10);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(44, 24);
            this.label5.TabIndex = 71;
            this.label5.Text = "Year";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(222, 10);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(60, 24);
            this.label4.TabIndex = 72;
            this.label4.Text = "Weight";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(13, 36);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(49, 22);
            this.textBox1.TabIndex = 73;
            this.textBox1.Text = "1";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox2.Location = new System.Drawing.Point(13, 120);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(49, 22);
            this.textBox2.TabIndex = 74;
            this.textBox2.Text = "12";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox3.Location = new System.Drawing.Point(13, 92);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(49, 22);
            this.textBox3.TabIndex = 75;
            this.textBox3.Text = "11";
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox4
            // 
            this.textBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox4.Location = new System.Drawing.Point(13, 64);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(49, 22);
            this.textBox4.TabIndex = 76;
            this.textBox4.Text = "6";
            this.textBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox5
            // 
            this.textBox5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox5.Location = new System.Drawing.Point(80, 92);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(49, 22);
            this.textBox5.TabIndex = 77;
            this.textBox5.Text = "0.1";
            this.textBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox6
            // 
            this.textBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox6.Location = new System.Drawing.Point(80, 64);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(49, 22);
            this.textBox6.TabIndex = 78;
            this.textBox6.Text = "0.15";
            this.textBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox7
            // 
            this.textBox7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox7.Location = new System.Drawing.Point(80, 36);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(49, 22);
            this.textBox7.TabIndex = 79;
            this.textBox7.Text = "0.025";
            this.textBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox8
            // 
            this.textBox8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox8.Location = new System.Drawing.Point(80, 120);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(49, 22);
            this.textBox8.TabIndex = 80;
            this.textBox8.Text = "0.15";
            this.textBox8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox16
            // 
            this.textBox16.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox16.Location = new System.Drawing.Point(158, 36);
            this.textBox16.Name = "textBox16";
            this.textBox16.Size = new System.Drawing.Size(49, 22);
            this.textBox16.TabIndex = 81;
            this.textBox16.Text = "21";
            this.textBox16.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox15
            // 
            this.textBox15.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox15.Location = new System.Drawing.Point(158, 120);
            this.textBox15.Name = "textBox15";
            this.textBox15.Size = new System.Drawing.Size(49, 22);
            this.textBox15.TabIndex = 82;
            this.textBox15.Text = "32";
            this.textBox15.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox14
            // 
            this.textBox14.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox14.Location = new System.Drawing.Point(158, 92);
            this.textBox14.Name = "textBox14";
            this.textBox14.Size = new System.Drawing.Size(49, 22);
            this.textBox14.TabIndex = 83;
            this.textBox14.Text = "31";
            this.textBox14.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox13
            // 
            this.textBox13.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox13.Location = new System.Drawing.Point(158, 64);
            this.textBox13.Name = "textBox13";
            this.textBox13.Size = new System.Drawing.Size(49, 22);
            this.textBox13.TabIndex = 84;
            this.textBox13.Text = "22";
            this.textBox13.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox12
            // 
            this.textBox12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox12.Location = new System.Drawing.Point(225, 92);
            this.textBox12.Name = "textBox12";
            this.textBox12.Size = new System.Drawing.Size(49, 22);
            this.textBox12.TabIndex = 85;
            this.textBox12.Text = "0.1";
            this.textBox12.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox11
            // 
            this.textBox11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox11.Location = new System.Drawing.Point(225, 64);
            this.textBox11.Name = "textBox11";
            this.textBox11.Size = new System.Drawing.Size(49, 22);
            this.textBox11.TabIndex = 86;
            this.textBox11.Text = "0.15";
            this.textBox11.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox10
            // 
            this.textBox10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox10.Location = new System.Drawing.Point(225, 36);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(49, 22);
            this.textBox10.TabIndex = 87;
            this.textBox10.Text = "0.1";
            this.textBox10.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // textBox9
            // 
            this.textBox9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox9.Location = new System.Drawing.Point(225, 120);
            this.textBox9.Name = "textBox9";
            this.textBox9.Size = new System.Drawing.Size(49, 22);
            this.textBox9.TabIndex = 88;
            this.textBox9.Text = "0.225";
            this.textBox9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(10, 148);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(167, 24);
            this.label6.TabIndex = 89;
            this.label6.Text = "Note: Weights must total 1";
            // 
            // pnlOptimization
            // 
            this.pnlOptimization.AutoScroll = true;
            this.pnlOptimization.Controls.Add(this.textBox17);
            this.pnlOptimization.Controls.Add(this.label7);
            this.pnlOptimization.Controls.Add(this.button2);
            this.pnlOptimization.Controls.Add(this.button1);
            this.pnlOptimization.Controls.Add(this.panel2);
            this.pnlOptimization.Controls.Add(this.btnNewSelectViewCancel);
            this.pnlOptimization.Controls.Add(this.groupBox4);
            this.pnlOptimization.Controls.Add(this.groupBox3);
            this.pnlOptimization.Controls.Add(this.groupBox2);
            this.pnlOptimization.Controls.Add(this.LblSelectedVariable);
            this.pnlOptimization.Controls.Add(this.lblSelectedFVSVariable);
            this.pnlOptimization.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlOptimization.Location = new System.Drawing.Point(3, 18);
            this.pnlOptimization.Name = "pnlOptimization";
            this.pnlOptimization.Size = new System.Drawing.Size(850, 451);
            this.pnlOptimization.TabIndex = 70;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(571, 405);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(76, 24);
            this.button1.TabIndex = 77;
            this.button1.Text = "Calculate";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(501, 405);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(64, 24);
            this.button2.TabIndex = 78;
            this.button2.Text = "Save";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(13, 405);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(212, 24);
            this.label7.TabIndex = 79;
            this.label7.Text = "Weighted average variable name:";
            // 
            // textBox17
            // 
            this.textBox17.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox17.Location = new System.Drawing.Point(231, 402);
            this.textBox17.Name = "textBox17";
            this.textBox17.Size = new System.Drawing.Size(228, 22);
            this.textBox17.TabIndex = 84;
            // 
            // uc_core_scenario_weighted_average
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_core_scenario_weighted_average";
            this.Size = new System.Drawing.Size(872, 2000);
            this.groupBox1.ResumeLayout(false);
            this.grpboxOptimizationFVSVariable.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.grpboxFVSVariablesOptimizationVariableValues.ResumeLayout(false);
            this.grpFVSVariablesOptimizationVariableValuesSelected.ResumeLayout(false);
            this.grpboxOptimization.ResumeLayout(false);
            this.grpboxOptimizationSettings.ResumeLayout(false);
            this.pnlFVSVariablesPrePostVariable.ResumeLayout(false);
            this.grpboxOptimizationSettingsPostPre.ResumeLayout(false);
            this.grpFilter.ResumeLayout(false);
            this.grpFilter.PerformLayout();
            this.grpMaxMin.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.pnlOptimization.ResumeLayout(false);
            this.pnlOptimization.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

        public void loadvalues_FromProperties()
        {


        }
        public void loadvalues(System.Windows.Forms.ListBox p_oListBox)
		{
			
			this.m_intError=0;
			this.m_strError="";

			int x,y;

			//
			//load previous scenario values
			//
			if (this.m_bFirstTime)
			{
				ado_data_access oAdo = new ado_data_access();
				string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
				string strScenarioMDB = 
					frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
					"\\core\\db\\scenario_core_rule_definitions.mdb";
				oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
                if (oAdo.m_intError == 0)
                {

                }
				this.m_intError=oAdo.m_intError;
				this.m_strError=oAdo.m_strError;
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo.m_OleDbConnection.Dispose();
				oAdo=null;
				this.m_bFirstTime=false;
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

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{

			
			this.grpboxOptimization.Height = this.ClientSize.Height - this.grpboxOptimization.Top - 5;
			grpboxOptimization.Width = this.ClientSize.Width - (grpboxOptimization.Left * 2) ;
		    grpboxOptimizationSettings.Height = grpboxOptimization.Height;
			grpboxOptimizationSettings.Width =  grpboxOptimization.Width;
			this.grpboxOptimizationFVSVariable.Height = grpboxOptimization.Height;
			this.grpboxOptimizationFVSVariable.Width = grpboxOptimization.Width;
		
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

		private void btnOptimiztionCancel_Click(object sender, System.EventArgs e)
		{
			this.EnableTabs(true);
			this.grpboxOptimizationSettings.Hide();
			this.grpboxOptimizationFVSVariable.Hide();
			this.grpboxOptimization.Show();
			
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





		private void btnFVSVariablesOptimizationVariableValues_Click(object sender, System.EventArgs e)
		{
			if (this.lstFVSVariablesOptimizationVariableValues.SelectedItems.Count==0) return;
			this.lblFVSVariablesOptimizationVariableValuesSelected.Text = this.lstFVSVariablesOptimizationVariableValues.SelectedItems[0].ToString();
		}


		private void btnOptimizationFVSVariableCancel_Click(object sender, System.EventArgs e)
		{
			this.EnableTabs(true);
			this.grpboxOptimizationSettings.Hide();
			this.grpboxOptimizationFVSVariable.Hide();
			this.grpboxOptimization.Show();

		}

		private void btnOptimizationFVSVariableNext_Click(object sender, System.EventArgs e)
		{
			this.grpboxOptimizationFVSVariable.Hide();			
			this.grpboxOptimizationSettings.Show();
		}

		private void btnOptimiztionPrev_Click(object sender, System.EventArgs e)
		{
			this.grpboxOptimizationSettings.Hide();
			this.grpboxOptimizationFVSVariable.Show();			
			
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
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_effective.Variables ReferenceFVSVariables
		{
			get {return this._oCurVar;}
			set {_oCurVar=value;}
		}
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker ReferenceTieBreaker
		{
			get {return _uc_tiebreaker;}
			set {_uc_tiebreaker=value;}
		}

        private void txtOptimizationValue_Leave(object sender, EventArgs e)
        {
            this.m_oValidate.ValidateDecimal(txtOptimizationValue.Text);
            if (m_oValidate.m_intError == 0)
            {
                this.txtOptimizationValue.Text = m_oValidate.ReturnValue;
                this.m_strLastValue = m_oValidate.ReturnValue;
            }
            else
            {
                this.txtOptimizationValue.Text = this.m_strLastValue;
            }

        }
	
	}
}
