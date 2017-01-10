using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_harvest_method.
	/// </summary>
	public class uc_processor_scenario_harvest_method : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox grpboxHarvestMethod;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtDesc;
		private System.Windows.Forms.ComboBox cmbMethod;
		private System.Windows.Forms.Label lblMethod;
		private System.Windows.Forms.Label lblSteepSlopeDesc;
		private System.Windows.Forms.TextBox txtSteepSlopeDesc;
		private System.Windows.Forms.ComboBox cmbSteepSlopeMethod;
		private System.Windows.Forms.Label lblSteepSlopeMethod;
		private System.Windows.Forms.ComboBox cmbSteepSlopePercent;
		private System.Windows.Forms.GroupBox grpboxSteepSlopeHarvestMethod;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtSteepSlopeMinDia;
		private System.Windows.Forms.TextBox txtMinDiaForChips;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtMinDiaSmallLogs;
		private System.Windows.Forms.TextBox txtMinDiaLargeLogs;
		private Queries m_oQueries = new Queries();
		private RxTools m_oRxTools = new RxTools();
		private ado_data_access m_oAdo = new ado_data_access();
		private string _strScenarioId="";
		private frmProcessorScenario _frmProcessorScenario=null;
		private System.Windows.Forms.CheckBox chkUseDefault;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private Label label7;
        private Label label11;
        private Label label10;
        private Label label9;
        private Label label8;
        private GroupBox groupBox2;
        private Label label12;
        private TextBox txtHelicopterCableMax;
        private Label label13;
        private TextBox txtCableMax;
        private Button btnDefault;
        private Label label15;
        private Label label14;
        
		

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_scenario_harvest_method()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

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
		public frmProcessorScenario ReferenceProcessorScenarioForm
		{
			get {return this._frmProcessorScenario;}
			set {this._frmProcessorScenario=value;}
		}
		public string ScenarioId
		{
			get {return _strScenarioId;}
			set {_strScenarioId=value;}
		}
        public bool UseDefaultHarvestMethod
        {
            get { return this.chkUseDefault.Checked; }
        }
        public string HarvestMethodSteepSlope
        {
            get { return this.cmbSteepSlopeMethod.Text.Trim(); }
        }
        public string HarvestMethodLowSlope
        {
            get { return this.cmbMethod.Text.Trim(); }
        }
        public string MinDiaForChips
        {
            get { return txtMinDiaForChips.Text.Trim(); }
        }
        public string MinDiaForSmallLogs
        {
            get { return txtMinDiaSmallLogs.Text.Trim(); }
        }
        public string MinDiaForLargeLogs
        {
            get { return txtMinDiaLargeLogs.Text.Trim(); }
        }
        public string SteepSlopePercent
        {
            get { return cmbSteepSlopePercent.Text.Trim(); }
        }
        public string MinDiaForAllTreesSteepSlope
        {
            get { return txtSteepSlopeMinDia.Text.Trim(); }
        }
        public string MaxCableYardingDistance
        {
            get { return txtCableMax.Text.Trim(); }
        }
        public string MaxHelicopterCableYardingDistance
        {
            get { return txtHelicopterCableMax.Text.Trim(); }
        }



		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtSteepSlopeMinDia = new System.Windows.Forms.TextBox();
            this.txtMinDiaForChips = new System.Windows.Forms.TextBox();
            this.cmbSteepSlopePercent = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.txtHelicopterCableMax = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtCableMax = new System.Windows.Forms.TextBox();
            this.btnDefault = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.txtMinDiaSmallLogs = new System.Windows.Forms.TextBox();
            this.chkUseDefault = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.grpboxSteepSlopeHarvestMethod = new System.Windows.Forms.GroupBox();
            this.lblSteepSlopeDesc = new System.Windows.Forms.Label();
            this.txtSteepSlopeDesc = new System.Windows.Forms.TextBox();
            this.cmbSteepSlopeMethod = new System.Windows.Forms.ComboBox();
            this.lblSteepSlopeMethod = new System.Windows.Forms.Label();
            this.txtMinDiaLargeLogs = new System.Windows.Forms.TextBox();
            this.grpboxHarvestMethod = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDesc = new System.Windows.Forms.TextBox();
            this.cmbMethod = new System.Windows.Forms.ComboBox();
            this.lblMethod = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.grpboxSteepSlopeHarvestMethod.SuspendLayout();
            this.grpboxHarvestMethod.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(808, 569);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(802, 32);
            this.lblTitle.TabIndex = 28;
            this.lblTitle.Text = "Harvest Method";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.txtSteepSlopeMinDia);
            this.panel1.Controls.Add(this.txtMinDiaForChips);
            this.panel1.Controls.Add(this.cmbSteepSlopePercent);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.txtMinDiaSmallLogs);
            this.panel1.Controls.Add(this.chkUseDefault);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.grpboxSteepSlopeHarvestMethod);
            this.panel1.Controls.Add(this.txtMinDiaLargeLogs);
            this.panel1.Controls.Add(this.grpboxHarvestMethod);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(802, 550);
            this.panel1.TabIndex = 0;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label10.Location = new System.Drawing.Point(286, 381);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(21, 13);
            this.label10.TabIndex = 18;
            this.label10.Text = ">=";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.Black;
            this.label11.Location = new System.Drawing.Point(578, 322);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(21, 13);
            this.label11.TabIndex = 19;
            this.label11.Text = ">=";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(354, 370);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(239, 29);
            this.label1.TabIndex = 8;
            this.label1.Text = "Percent slope threshold at which slope \r\nis categorized as steep";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.Black;
            this.label7.Location = new System.Drawing.Point(13, 325);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(21, 13);
            this.label7.TabIndex = 17;
            this.label7.Text = ">=";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.Black;
            this.label8.Location = new System.Drawing.Point(13, 367);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(21, 13);
            this.label8.TabIndex = 18;
            this.label8.Text = ">=";
            // 
            // txtSteepSlopeMinDia
            // 
            this.txtSteepSlopeMinDia.ForeColor = System.Drawing.Color.Black;
            this.txtSteepSlopeMinDia.Location = new System.Drawing.Point(603, 319);
            this.txtSteepSlopeMinDia.Name = "txtSteepSlopeMinDia";
            this.txtSteepSlopeMinDia.Size = new System.Drawing.Size(33, 20);
            this.txtSteepSlopeMinDia.TabIndex = 10;
            this.txtSteepSlopeMinDia.Text = "5.0";
            this.txtSteepSlopeMinDia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSteepSlopeMinDia_KeyPress);
            this.txtSteepSlopeMinDia.Leave += new System.EventHandler(this.txtSteepSlopeMinDia_Leave);
            // 
            // txtMinDiaForChips
            // 
            this.txtMinDiaForChips.ForeColor = System.Drawing.Color.Black;
            this.txtMinDiaForChips.Location = new System.Drawing.Point(32, 322);
            this.txtMinDiaForChips.Name = "txtMinDiaForChips";
            this.txtMinDiaForChips.Size = new System.Drawing.Size(36, 20);
            this.txtMinDiaForChips.TabIndex = 12;
            this.txtMinDiaForChips.Text = "3.0";
            this.txtMinDiaForChips.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMinDiaForChips_KeyPress);
            this.txtMinDiaForChips.Leave += new System.EventHandler(this.txtMinDiaForChips_Leave);
            // 
            // cmbSteepSlopePercent
            // 
            this.cmbSteepSlopePercent.ForeColor = System.Drawing.Color.Black;
            this.cmbSteepSlopePercent.ItemHeight = 13;
            this.cmbSteepSlopePercent.Location = new System.Drawing.Point(311, 378);
            this.cmbSteepSlopePercent.Name = "cmbSteepSlopePercent";
            this.cmbSteepSlopePercent.Size = new System.Drawing.Size(37, 21);
            this.cmbSteepSlopePercent.TabIndex = 0;
            this.cmbSteepSlopePercent.Text = "40";
            this.cmbSteepSlopePercent.SelectedValueChanged += new System.EventHandler(this.cmbSteepSlopePercent_SelectedValueChanged);
            this.cmbSteepSlopePercent.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSteepSlopePercent_KeyPress);
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(642, 319);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(147, 65);
            this.label3.TabIndex = 9;
            this.label3.Text = "Minimum diameteter of trees that will be utilized for any purpose on steep slopes" +
    "";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.Black;
            this.label9.Location = new System.Drawing.Point(286, 329);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(21, 13);
            this.label9.TabIndex = 19;
            this.label9.Text = ">=";
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(71, 317);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(208, 40);
            this.label4.TabIndex = 11;
            this.label4.Text = "Minimum diameter for chip trees (Trees to be chipped for energy wood)";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label15);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.txtHelicopterCableMax);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.txtCableMax);
            this.groupBox2.Controls.Add(this.btnDefault);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.ForeColor = System.Drawing.Color.Black;
            this.groupBox2.Location = new System.Drawing.Point(16, 419);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(760, 78);
            this.groupBox2.TabIndex = 31;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Cable One-Way Yarding Limits (Feet)";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(16, 52);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(14, 13);
            this.label15.TabIndex = 38;
            this.label15.Text = ">";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(16, 24);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(14, 13);
            this.label14.TabIndex = 37;
            this.label14.Text = ">";
            // 
            // txtHelicopterCableMax
            // 
            this.txtHelicopterCableMax.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtHelicopterCableMax.Location = new System.Drawing.Point(43, 49);
            this.txtHelicopterCableMax.Name = "txtHelicopterCableMax";
            this.txtHelicopterCableMax.Size = new System.Drawing.Size(49, 20);
            this.txtHelicopterCableMax.TabIndex = 36;
            this.txtHelicopterCableMax.Text = "10000";
            this.txtHelicopterCableMax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtHelicopterCableMax_KeyPress);
            this.txtHelicopterCableMax.Leave += new System.EventHandler(this.txtHelicopterCableMax_Leave);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(109, 52);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(327, 13);
            this.label13.TabIndex = 35;
            this.label13.Text = "Helicopter Manual Whole Tree, Helicopter Cut-to-Length";
            // 
            // txtCableMax
            // 
            this.txtCableMax.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCableMax.Location = new System.Drawing.Point(43, 21);
            this.txtCableMax.Name = "txtCableMax";
            this.txtCableMax.Size = new System.Drawing.Size(49, 20);
            this.txtCableMax.TabIndex = 34;
            this.txtCableMax.Text = "1300";
            this.txtCableMax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCableMax_KeyPress);
            this.txtCableMax.Leave += new System.EventHandler(this.txtCableMax_Leave);
            // 
            // btnDefault
            // 
            this.btnDefault.Location = new System.Drawing.Point(612, 28);
            this.btnDefault.Name = "btnDefault";
            this.btnDefault.Size = new System.Drawing.Size(93, 37);
            this.btnDefault.TabIndex = 33;
            this.btnDefault.Text = "Default";
            this.btnDefault.UseVisualStyleBackColor = true;
            this.btnDefault.Click += new System.EventHandler(this.btnDefault_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(109, 28);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(472, 13);
            this.label12.TabIndex = 32;
            this.label12.Text = "Manual Whole Tree/Log,  Manual Whole Tree, Manual Log, Manual Cut-to-Length";
            // 
            // txtMinDiaSmallLogs
            // 
            this.txtMinDiaSmallLogs.ForeColor = System.Drawing.Color.Black;
            this.txtMinDiaSmallLogs.Location = new System.Drawing.Point(32, 364);
            this.txtMinDiaSmallLogs.Name = "txtMinDiaSmallLogs";
            this.txtMinDiaSmallLogs.Size = new System.Drawing.Size(36, 20);
            this.txtMinDiaSmallLogs.TabIndex = 15;
            this.txtMinDiaSmallLogs.Text = "7.0";
            this.txtMinDiaSmallLogs.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMinDiaSmallLogs_KeyPress);
            this.txtMinDiaSmallLogs.Leave += new System.EventHandler(this.txtMinDiaSmallLogs_Leave);
            // 
            // chkUseDefault
            // 
            this.chkUseDefault.Checked = true;
            this.chkUseDefault.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUseDefault.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkUseDefault.ForeColor = System.Drawing.Color.Black;
            this.chkUseDefault.Location = new System.Drawing.Point(16, 29);
            this.chkUseDefault.Name = "chkUseDefault";
            this.chkUseDefault.Size = new System.Drawing.Size(368, 24);
            this.chkUseDefault.TabIndex = 30;
            this.chkUseDefault.Text = "Use Treatment Default Harvest Method";
            this.chkUseDefault.CheckedChanged += new System.EventHandler(this.chkUseDefaullt_CheckedChanged);
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(71, 359);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(208, 40);
            this.label5.TabIndex = 13;
            this.label5.Text = "Minimum diameter for small log trees (trees small enough to fell and process into" +
    " logs by machine)";
            // 
            // grpboxSteepSlopeHarvestMethod
            // 
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.lblSteepSlopeDesc);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.txtSteepSlopeDesc);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.cmbSteepSlopeMethod);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.lblSteepSlopeMethod);
            this.grpboxSteepSlopeHarvestMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxSteepSlopeHarvestMethod.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.grpboxSteepSlopeHarvestMethod.Location = new System.Drawing.Point(408, 61);
            this.grpboxSteepSlopeHarvestMethod.Name = "grpboxSteepSlopeHarvestMethod";
            this.grpboxSteepSlopeHarvestMethod.Size = new System.Drawing.Size(368, 252);
            this.grpboxSteepSlopeHarvestMethod.TabIndex = 29;
            this.grpboxSteepSlopeHarvestMethod.TabStop = false;
            this.grpboxSteepSlopeHarvestMethod.Text = "Steep Slope Harvest Method";
            // 
            // lblSteepSlopeDesc
            // 
            this.lblSteepSlopeDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSteepSlopeDesc.Location = new System.Drawing.Point(18, 56);
            this.lblSteepSlopeDesc.Name = "lblSteepSlopeDesc";
            this.lblSteepSlopeDesc.Size = new System.Drawing.Size(182, 16);
            this.lblSteepSlopeDesc.TabIndex = 7;
            this.lblSteepSlopeDesc.Text = "Description (Read Only)";
            // 
            // txtSteepSlopeDesc
            // 
            this.txtSteepSlopeDesc.Enabled = false;
            this.txtSteepSlopeDesc.ForeColor = System.Drawing.Color.Black;
            this.txtSteepSlopeDesc.Location = new System.Drawing.Point(18, 72);
            this.txtSteepSlopeDesc.Multiline = true;
            this.txtSteepSlopeDesc.Name = "txtSteepSlopeDesc";
            this.txtSteepSlopeDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSteepSlopeDesc.Size = new System.Drawing.Size(288, 166);
            this.txtSteepSlopeDesc.TabIndex = 6;
            this.txtSteepSlopeDesc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSteepSlopeDesc_KeyPress);
            // 
            // cmbSteepSlopeMethod
            // 
            this.cmbSteepSlopeMethod.Enabled = false;
            this.cmbSteepSlopeMethod.ForeColor = System.Drawing.Color.Black;
            this.cmbSteepSlopeMethod.ItemHeight = 13;
            this.cmbSteepSlopeMethod.Location = new System.Drawing.Point(18, 32);
            this.cmbSteepSlopeMethod.Name = "cmbSteepSlopeMethod";
            this.cmbSteepSlopeMethod.Size = new System.Drawing.Size(294, 21);
            this.cmbSteepSlopeMethod.TabIndex = 5;
            this.cmbSteepSlopeMethod.SelectedIndexChanged += new System.EventHandler(this.cmbSteepSlopeMethod_SelectedValueChanged);
            this.cmbSteepSlopeMethod.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSteepSlopeMethod_KeyPress);
            // 
            // lblSteepSlopeMethod
            // 
            this.lblSteepSlopeMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSteepSlopeMethod.Location = new System.Drawing.Point(18, 16);
            this.lblSteepSlopeMethod.Name = "lblSteepSlopeMethod";
            this.lblSteepSlopeMethod.Size = new System.Drawing.Size(80, 16);
            this.lblSteepSlopeMethod.TabIndex = 4;
            this.lblSteepSlopeMethod.Text = "Method";
            // 
            // txtMinDiaLargeLogs
            // 
            this.txtMinDiaLargeLogs.ForeColor = System.Drawing.Color.Black;
            this.txtMinDiaLargeLogs.Location = new System.Drawing.Point(311, 326);
            this.txtMinDiaLargeLogs.Name = "txtMinDiaLargeLogs";
            this.txtMinDiaLargeLogs.Size = new System.Drawing.Size(37, 20);
            this.txtMinDiaLargeLogs.TabIndex = 16;
            this.txtMinDiaLargeLogs.Text = "20.0";
            this.txtMinDiaLargeLogs.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMinDiaLargeLogs_KeyPress);
            this.txtMinDiaLargeLogs.Leave += new System.EventHandler(this.txtMinDiaLargeLogs_Leave);
            // 
            // grpboxHarvestMethod
            // 
            this.grpboxHarvestMethod.Controls.Add(this.label2);
            this.grpboxHarvestMethod.Controls.Add(this.txtDesc);
            this.grpboxHarvestMethod.Controls.Add(this.cmbMethod);
            this.grpboxHarvestMethod.Controls.Add(this.lblMethod);
            this.grpboxHarvestMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxHarvestMethod.ForeColor = System.Drawing.Color.ForestGreen;
            this.grpboxHarvestMethod.Location = new System.Drawing.Point(16, 61);
            this.grpboxHarvestMethod.Name = "grpboxHarvestMethod";
            this.grpboxHarvestMethod.Size = new System.Drawing.Size(368, 253);
            this.grpboxHarvestMethod.TabIndex = 28;
            this.grpboxHarvestMethod.TabStop = false;
            this.grpboxHarvestMethod.Text = "Low Slope Harvest Method";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(16, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(184, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Description (Read Only)";
            // 
            // txtDesc
            // 
            this.txtDesc.Enabled = false;
            this.txtDesc.Location = new System.Drawing.Point(16, 72);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtDesc.Size = new System.Drawing.Size(288, 166);
            this.txtDesc.TabIndex = 2;
            this.txtDesc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRxDesc_KeyPress);
            // 
            // cmbMethod
            // 
            this.cmbMethod.Enabled = false;
            this.cmbMethod.Location = new System.Drawing.Point(16, 32);
            this.cmbMethod.Name = "cmbMethod";
            this.cmbMethod.Size = new System.Drawing.Size(288, 21);
            this.cmbMethod.TabIndex = 1;
            this.cmbMethod.SelectedIndexChanged += new System.EventHandler(this.cmbMethod_SelectedValueChanged);
            this.cmbMethod.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbMethod_KeyPress);
            // 
            // lblMethod
            // 
            this.lblMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMethod.Location = new System.Drawing.Point(16, 16);
            this.lblMethod.Name = "lblMethod";
            this.lblMethod.Size = new System.Drawing.Size(107, 16);
            this.lblMethod.TabIndex = 0;
            this.lblMethod.Text = "Method";
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(358, 322);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(208, 48);
            this.label6.TabIndex = 14;
            this.label6.Text = "Minimum diameter for large log trees (Trees that require chainsaw felling)";
            // 
            // uc_processor_scenario_harvest_method
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_harvest_method";
            this.Size = new System.Drawing.Size(808, 569);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.grpboxSteepSlopeHarvestMethod.ResumeLayout(false);
            this.grpboxSteepSlopeHarvestMethod.PerformLayout();
            this.grpboxHarvestMethod.ResumeLayout(false);
            this.grpboxHarvestMethod.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}
		public void loadvalues()
		{
            
			ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

			this.txtDesc.Text="";
			this.txtSteepSlopeDesc.Text="";
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.m_oReference.LoadDatasource=true;
			m_oQueries.LoadDatasources(true,"processor",ScenarioId);
			m_oAdo = new ado_data_access();
			m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile,"",""));
			if (m_oAdo.m_intError==0)
			{
				this.m_oRxTools.LoadRxHarvestMethods(m_oAdo,m_oAdo.m_OleDbConnection,m_oQueries,cmbMethod,cmbSteepSlopeMethod);
     		}
			this.cmbSteepSlopePercent.Items.Clear();
			for (int x=90;x>=10;x=x-5)
			{
				this.cmbSteepSlopePercent.Items.Add(x.ToString().Trim());
			}
			this.cmbSteepSlopePercent.Text = "40";

            ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadHarvestMethod
                (frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\db\\scenario_processor_rule_definitions.mdb",
                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);

            FIA_Biosum_Manager.ProcessorScenarioItem oItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem;
            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.m_intError == 0)
            {
                this.chkUseDefault.Checked = oItem.m_oHarvestMethod.UseDefaultHarvestMethod;
                this.cmbMethod.Text = oItem.m_oHarvestMethod.HarvestMethodLowSlope;
                cmbSteepSlopeMethod.Text = oItem.m_oHarvestMethod.HarvestMethodSteepSlope;
                txtCableMax.Text = oItem.m_oHarvestMethod.MaxCableYardingDistance;
                txtHelicopterCableMax.Text = oItem.m_oHarvestMethod.MaxHelicopterCableYardingDistance;
                txtMinDiaForChips.Text = oItem.m_oHarvestMethod.MinDiaForChips;
                txtMinDiaSmallLogs.Text = oItem.m_oHarvestMethod.MinDiaForSmallLogs;
                txtMinDiaLargeLogs.Text = oItem.m_oHarvestMethod.MinDiaForLargeLogs;
                cmbSteepSlopePercent.Text = oItem.m_oHarvestMethod.SteepSlopePercent;
                txtSteepSlopeMinDia.Text = oItem.m_oHarvestMethod.MinDiaForAllTreesSteepSlope;
                uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessLowSlope =
                     oItem.m_oHarvestMethod.ProcessLowSlope;
                uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessSteepSlope =
                    oItem.m_oHarvestMethod.ProcessSteepSlope;

            }

			
		}

        public void loadvalues_FromProperties()
        {

            this.txtDesc.Text = "";
            this.txtSteepSlopeDesc.Text = "";
            this.cmbSteepSlopePercent.Items.Clear();
            for (int x = 90; x >= 10; x = x - 5)
            {
                this.cmbSteepSlopePercent.Items.Add(x.ToString().Trim());
            }
            this.cmbSteepSlopePercent.Text = "40";

            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oHarvestMethod != null)
            {
                FIA_Biosum_Manager.ProcessorScenarioItem oItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem;
                if (ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.m_intError == 0)
                {
                    this.chkUseDefault.Checked = oItem.m_oHarvestMethod.UseDefaultHarvestMethod;
                    this.cmbMethod.Text = oItem.m_oHarvestMethod.HarvestMethodLowSlope;
                    cmbSteepSlopeMethod.Text = oItem.m_oHarvestMethod.HarvestMethodSteepSlope;
                    txtCableMax.Text = oItem.m_oHarvestMethod.MaxCableYardingDistance;
                    txtHelicopterCableMax.Text = oItem.m_oHarvestMethod.MaxHelicopterCableYardingDistance;
                    txtMinDiaForChips.Text = oItem.m_oHarvestMethod.MinDiaForChips;
                    txtMinDiaSmallLogs.Text = oItem.m_oHarvestMethod.MinDiaForSmallLogs;
                    txtMinDiaLargeLogs.Text = oItem.m_oHarvestMethod.MinDiaForLargeLogs;
                    cmbSteepSlopePercent.Text = oItem.m_oHarvestMethod.SteepSlopePercent;
                    txtSteepSlopeMinDia.Text = oItem.m_oHarvestMethod.MinDiaForAllTreesSteepSlope;
                    uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessLowSlope =
                         oItem.m_oHarvestMethod.ProcessLowSlope;
                    uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessSteepSlope =
                        oItem.m_oHarvestMethod.ProcessSteepSlope;
                }
            }

        }
		public void savevalues()
		{
			//
			//OPEN CONNECTION TO DB FILE CONTAINING PROCESSOR SCENARIO TABLES
			//
			//scenario mdb connection
			ado_data_access oAdo = new ado_data_access();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
				"\\processor\\db\\scenario_processor_rule_definitions.mdb";
			oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));	
			if (oAdo.m_intError != 0)
			{
				m_intError=oAdo.m_intError;
				m_strError=oAdo.m_strError;
				oAdo = null;
				return;
			}

			m_intError=0;
			m_strError="";
            string strFields = "scenario_id,UseRxDefaultHarvestMethodYN,HarvestMethodLowSlope," + 
				             "HarvestMethodSteepSlope," +
                             "MaxCableYardingDistance," + 
                             "MaxHelicopterCableYardingDistance," + 
                             "min_chip_dbh,min_sm_log_dbh," + 
				             "min_lg_log_dbh,SteepSlope,min_dbh_steep_slope," + 
                             "ProcessLowSlopeYN,ProcessSteepSlopeYN";
			string strValues="";
			
			oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " " + 
				              "WHERE TRIM(UCASE(scenario_id)) = '" + ScenarioId.Trim().ToUpper() + "'";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

			//
			//SCENARIOID
			//
			strValues="'" + ScenarioId + "',";
			//
			//DEFAULT HARVEST METHOD
			//
			if (this.chkUseDefault.Checked)
			{
				strValues=strValues + "'Y',";
			}
			else
			{
				strValues=strValues + "'N',";
			}
			//
			//HARVEST METHOD
			//
			strValues=strValues + "'" + this.cmbMethod.Text.Trim() + "',"; 
			//
			//HARVEST METHOD STEEP SLOPE
			//
			strValues=strValues + "'" + this.cmbSteepSlopeMethod.Text.Trim() + "',"; 
            //
            //MAXIMUM CABLE YARDING DISTANCE
            //
            if (this.txtCableMax.Text.Trim().Length > 0)
            {
                strValues = strValues + this.txtCableMax.Text.Trim() + ",";
            }
            else
            {
                strValues = strValues + "null,";
            }
            //
            //MAXIMUM HELICOPTER CABLE YARDING DISTANCE
            //
            if (this.txtHelicopterCableMax.Text.Trim().Length > 0)
            {
                strValues = strValues + this.txtHelicopterCableMax.Text.Trim() + ",";
            }
            else
            {
                strValues = strValues + "null,";
            }
			//
			//MINIMUM CHIP DBH
			//
			if (this.txtMinDiaForChips.Text.Trim().Length > 0)
			{
				strValues=strValues + this.txtMinDiaForChips.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
			//
			//MINIMUM DBH FOR SMALL LOGS
			//
			if (this.txtMinDiaSmallLogs.Text.Trim().Length > 0)
			{
				strValues=strValues + this.txtMinDiaSmallLogs.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
			//
			//MINIMUM DBH FOR LARGE LOGS
			//
			if (this.txtMinDiaLargeLogs.Text.Trim().Length > 0)
			{
				strValues=strValues + this.txtMinDiaLargeLogs.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
			//
			//STEEP SLOPE PERCENT
			//
			if (this.cmbSteepSlopePercent.Text.Trim().Length > 0)
			{
				strValues=strValues + this.cmbSteepSlopePercent.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
			//
			//MIN DBH ON STEEP SLOPE
			//
			if (txtSteepSlopeMinDia.Text.Trim().Length > 0)
			{
				strValues=strValues + txtSteepSlopeMinDia.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
            //
            //PROCESS LOW SLOPE DATA DURING RUN
            //
            if (uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessLowSlope == true)
            {
                strValues = strValues + "'Y',";
            }
            else
            {
                strValues = strValues + "'N',";
            }
            //
            //PROCESS STEEP SLOPE DATA DURING RUN
            //
            if (uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessSteepSlope == true)
            {
                strValues = strValues + "'Y'";
            }
            else
            {
                strValues = strValues + "'N'";
            }
			oAdo.m_strSQL=Queries.GetInsertSQL(strFields,strValues,Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName);
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
			m_intError=oAdo.m_intError;

			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oAdo=null;

			
			
		}
		private void cmbMethod_SelectedValueChanged(object sender, System.EventArgs e)
		{
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
			getDesc();
			
		}
		private void getDesc()
		{
			if (m_oAdo.m_OleDbDataReader.IsClosed==false) return;

			m_oAdo.m_strSQL = Queries.GenericSelectSQL(m_oQueries.m_oReference.m_strRefHarvestMethodTable,"description","TRIM(method)='" + cmbMethod.Text.Trim() + "' AND steep_yn = 'N'");
			this.txtDesc.Text = m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL,"temp");
		}

		private void cmbSteepSlopeMethod_SelectedValueChanged(object sender, System.EventArgs e)
		{
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
			getDescSteepSlope();
		}
		private void getDescSteepSlope()
		{
			if (m_oAdo.m_OleDbDataReader.IsClosed==false) return;

			m_oAdo.m_strSQL = Queries.GenericSelectSQL(m_oQueries.m_oReference.m_strRefHarvestMethodTable,"description","TRIM(method)='" + this.cmbSteepSlopeMethod.Text.Trim() + "' AND steep_yn = 'Y'");
			this.txtSteepSlopeDesc.Text = m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL,"temp");
		}

		private void txtRxDesc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=false;	
		}

		private void txtDesc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=true;	
		}

		

		private void txtRxDesc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbMethod_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbSteepSlopeMethod_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtSteepSlopeDesc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtMinDiaForChips_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.MaxValue=300;
			m_oValidate.MinValue=0;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.ValidateDecimal(this.txtMinDiaForChips.Text.Trim());
			if (m_oValidate.m_intError==0)
			{
				this.txtMinDiaForChips.Text = m_oValidate.ReturnValue;
			}
			else
				this.txtMinDiaForChips.Focus();
		}

		private void chkUseDefaullt_CheckedChanged(object sender, System.EventArgs e)
		{
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
			if (this.chkUseDefault.Checked) 
			{
				this.cmbSteepSlopeMethod.Enabled=false;
				this.cmbMethod.Enabled=false;

				this.txtDesc.Enabled=false;

				this.txtSteepSlopeDesc.Enabled=false;
			}
			else
			{
				cmbMethod.Enabled=true;
				cmbSteepSlopeMethod.Enabled=true;
				this.txtDesc.Enabled=true;
				this.txtSteepSlopeDesc.Enabled=true;
			}
		}

		private void txtMinDiaSmallLogs_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.MaxValue=300;
			m_oValidate.MinValue=0;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.ValidateDecimal(txtMinDiaSmallLogs.Text.Trim());
			if (m_oValidate.m_intError==0)
			{
				txtMinDiaSmallLogs.Text = m_oValidate.ReturnValue;
			}
			else
				txtMinDiaSmallLogs.Focus();
		}

		private void txtMinDiaLargeLogs_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.MaxValue=300;
			m_oValidate.MinValue=0;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.ValidateDecimal(txtMinDiaLargeLogs.Text.Trim());
			if (m_oValidate.m_intError==0)
			{
				txtMinDiaLargeLogs.Text = m_oValidate.ReturnValue;
			}
			else
				txtMinDiaLargeLogs.Focus();
		}

		private void txtSteepSlopeMinDia_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.MaxValue=300;
			m_oValidate.MinValue=0;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.ValidateDecimal(txtSteepSlopeMinDia.Text.Trim());
			if (m_oValidate.m_intError==0)
			{
				txtSteepSlopeMinDia.Text = m_oValidate.ReturnValue;
			}
			else
				txtSteepSlopeMinDia.Focus();
		}

		private void cmbSteepSlopePercent_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void btnDefault_Click(object sender, EventArgs e)
        {
            if (txtCableMax.Text.Trim() != "1300" ||
                txtHelicopterCableMax.Text.Trim() != "10000")
                ReferenceProcessorScenarioForm.m_bSave = true;

            txtCableMax.Text = "1300";
            txtHelicopterCableMax.Text = "10000";

        }

        private void txtCableMax_Leave(object sender, EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 0;
            m_oValidate.MaxValue = 20000;
            m_oValidate.MinValue = 50;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(this.txtCableMax.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                this.txtCableMax.Text = m_oValidate.ReturnValue;
            }
            else
                this.txtCableMax.Focus();
        }

        private void txtHelicopterCableMax_Leave(object sender, EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 0;
            m_oValidate.MaxValue = 20000;
            m_oValidate.MinValue = 50;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(this.txtHelicopterCableMax.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                this.txtHelicopterCableMax.Text = m_oValidate.ReturnValue;
            }
            else
                this.txtHelicopterCableMax.Focus();
        }

        private void txtMinDiaForChips_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtMinDiaSmallLogs_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtMinDiaLargeLogs_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtSteepSlopeMinDia_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtCableMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtHelicopterCableMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void cmbSteepSlopePercent_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime==false) ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
       
	
		
	}
}
