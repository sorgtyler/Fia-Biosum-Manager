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
	public class uc_scenario_ffe : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ImageList imgSize;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.GroupBox grpboxFFEClassifications;
		//private int m_intFullHt=400;
		public System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;
		public System.Data.DataSet m_DataSet;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionMaster;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		public System.Data.OleDb.OleDbCommand m_OleDbCommand;
		public System.Data.DataRelation m_DataRelation;
		public System.Data.DataTable m_DataTable;
		private System.Windows.Forms.Button btnFFEDefaultStep1;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.GroupBox grpboxFFE_TI_CI_Effective;
		private System.Windows.Forms.Button btnFFEDefaultStep2;
		private System.Windows.Forms.ListView listView1;
		public System.Data.DataRow m_DataRow;
		public int m_intError = 0;
		private string m_strError = "";
		private System.Windows.Forms.Button btnFFEExpressionStep2;
		private System.Windows.Forms.TextBox txtFFE_TI5;
		private System.Windows.Forms.ComboBox cmbFFE_TI5;
		private System.Windows.Forms.CheckBox chkFFE_TI5;
		private System.Windows.Forms.TextBox txtFFE_TI4;
		private System.Windows.Forms.ComboBox cmbFFE_TI4;
		private System.Windows.Forms.CheckBox chkFFE_TI4;
		private System.Windows.Forms.TextBox txtFFE_TI3;
		private System.Windows.Forms.ComboBox cmbFFE_TI3;
		private System.Windows.Forms.CheckBox chkFFE_TI3;
		private System.Windows.Forms.TextBox txtFFE_TI2;
		private System.Windows.Forms.ComboBox cmbFFE_TI2;
		private System.Windows.Forms.CheckBox chkFFE_TI2;
		private System.Windows.Forms.TextBox txtFFE_TI1;
		private System.Windows.Forms.ComboBox cmbFFE_TI1;
		private System.Windows.Forms.CheckBox chkFFE_TI1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.TextBox txtFFE_CI5;
		private System.Windows.Forms.ComboBox cmbFFE_CI5;
		private System.Windows.Forms.CheckBox chkFFE_CI5;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.TextBox txtFFE_CI4;
		private System.Windows.Forms.ComboBox cmbFFE_CI4;
		private System.Windows.Forms.CheckBox chkFFE_CI4;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.TextBox txtFFE_CI3;
		private System.Windows.Forms.ComboBox cmbFFE_CI3;
		private System.Windows.Forms.CheckBox chkFFE_CI3;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.TextBox txtFFE_CI2;
		private System.Windows.Forms.ComboBox cmbFFE_CI2;
		private System.Windows.Forms.CheckBox chkFFE_CI2;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.TextBox txtFFE_CI1;
		private System.Windows.Forms.ComboBox cmbFFE_CI1;
		private System.Windows.Forms.GroupBox grpboxFFEExpressionBuilder;
		private System.Windows.Forms.TextBox txtExpression;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label lblFFEExpressionBuilder;
		private System.Windows.Forms.ListBox lstFFEAvailableFields;
		private System.Windows.Forms.Button btnFFEEqual;
		private System.Windows.Forms.Button btnFFENotEqual;
		private System.Windows.Forms.Button btnFFEMoreThan;
		private System.Windows.Forms.Button btnFFELessThan;
		private System.Windows.Forms.Button btnFFEExpressionBuilderDefault;
		private System.Windows.Forms.Button btnFFEMoreThanOrEqualTo;
		private System.Windows.Forms.Button btnLessThanOrEqualTo;
		private System.Windows.Forms.Button btnLeftBracket;
		private System.Windows.Forms.Button btnRightBracket;
		private System.Windows.Forms.Button btnFFEAnd;
		private System.Windows.Forms.Button btnFFEOr;
		private System.Windows.Forms.Label lblFFEExpressionFieldDesc;
		private System.Windows.Forms.Button btnFFEClearExpression;
		private System.Windows.Forms.CheckBox chkFFE_CI1;
		private System.Windows.Forms.Button btnFFENote;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.ComboBox cmbFFETIBackslide;
		private System.Windows.Forms.GroupBox grpboxFFEBackslide;
		private System.Windows.Forms.TrackBar trbFFETIBackslide;
		private System.Windows.Forms.TrackBar trbFFECIBackslide;
		private System.Windows.Forms.ComboBox cmbFFECIBackslide;
		private System.Windows.Forms.ComboBox cmbFFETIBackslide2;
		private System.Windows.Forms.CheckBox chkFFETIBackslide;
		private System.Windows.Forms.ComboBox cmbFFETIBackslide3;
		private System.Windows.Forms.ComboBox cmbFFECIBackslide3;
		private System.Windows.Forms.CheckBox chkFFECIBackslide;
		private System.Windows.Forms.ComboBox cmbFFECIBackslide2;
		private System.Windows.Forms.Button btnFFEDefaultStep3;
		private System.Windows.Forms.Label lblFFETIBackslide;
		private System.Windows.Forms.Label lblFFECIBackslide;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.GroupBox grpboxFFEHazard;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label30;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.ComboBox cmbFFECIHazardWindSpeedClass;
		private System.Windows.Forms.CheckBox chkFFECIHazard;
		private System.Windows.Forms.ComboBox cmbFFECIHazardOperator;
		private System.Windows.Forms.ComboBox cmbFFETIHazardWindSpeedClass;
		private System.Windows.Forms.CheckBox chkFFETIHazard;
		private System.Windows.Forms.ComboBox cmbFFETIHazardOperator;
		private System.Windows.Forms.Button btnNext;
		private System.Windows.Forms.Button btnPrev;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Button btnFFEHazardDefault;
		private string m_strCurrentIndexTypeAndClass;
		private System.Windows.Forms.Button btnFFEExpressionBuilderPreviousExpression;
		private System.Windows.Forms.Button btnFFEExpressionBuilderClear;
		private System.Windows.Forms.Button btnFFEExpressionBuilderCancel;
		private System.Windows.Forms.Button btnFFEExpressionBuilderOk;
		private System.Windows.Forms.Button btnFFEExpressionBuilderTest;
		private System.Windows.Forms.Button btnFFETICIEffectivePreviousExpressions;
		private System.Windows.Forms.Button btnY;
		private System.Windows.Forms.Button btnN;
		private string m_strOverallEffectiveExpression="";
		private FIA_Biosum_Manager.utils m_oUtils; 
		private bool _bTorchIndex=true;
		private bool _bCrownIndex=true;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmCoreScenario _frmScenario=null;
		

		public uc_scenario_ffe()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_oUtils = new utils();
			this.m_strCurrentIndexTypeAndClass="";

			//next and previous buttons
			this.btnNext = new Button();
			this.btnNext.Text = "Next >";
			this.btnNext.Visible=true;
			this.btnNext.BackColor = this.btnFFEDefaultStep1.BackColor;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
			this.NextButton(ref this.grpboxFFEClassifications,ref this.btnNext,"windspeed_next");
			
			this.btnPrev = new Button();
			this.btnPrev.BackColor = this.btnNext.BackColor;
			this.btnPrev.Width = this.btnNext.Width;
			this.btnPrev.Height = this.btnNext.Height;
			this.btnPrev.Text = "< Previous";
			this.btnPrev.Click += new System.EventHandler(this.btnPrev_Click);


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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_scenario_ffe));
			this.imgSize = new System.Windows.Forms.ImageList(this.components);
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblTitle = new System.Windows.Forms.Label();
			this.grpboxFFEBackslide = new System.Windows.Forms.GroupBox();
			this.label32 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.lblFFECIBackslide = new System.Windows.Forms.Label();
			this.lblFFETIBackslide = new System.Windows.Forms.Label();
			this.cmbFFECIBackslide3 = new System.Windows.Forms.ComboBox();
			this.chkFFECIBackslide = new System.Windows.Forms.CheckBox();
			this.cmbFFECIBackslide2 = new System.Windows.Forms.ComboBox();
			this.cmbFFETIBackslide3 = new System.Windows.Forms.ComboBox();
			this.chkFFETIBackslide = new System.Windows.Forms.CheckBox();
			this.cmbFFETIBackslide2 = new System.Windows.Forms.ComboBox();
			this.trbFFECIBackslide = new System.Windows.Forms.TrackBar();
			this.cmbFFECIBackslide = new System.Windows.Forms.ComboBox();
			this.trbFFETIBackslide = new System.Windows.Forms.TrackBar();
			this.label21 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.btnFFEDefaultStep3 = new System.Windows.Forms.Button();
			this.cmbFFETIBackslide = new System.Windows.Forms.ComboBox();
			this.grpboxFFEExpressionBuilder = new System.Windows.Forms.GroupBox();
			this.btnN = new System.Windows.Forms.Button();
			this.btnY = new System.Windows.Forms.Button();
			this.btnFFEExpressionBuilderPreviousExpression = new System.Windows.Forms.Button();
			this.btnFFENote = new System.Windows.Forms.Button();
			this.lblFFEExpressionFieldDesc = new System.Windows.Forms.Label();
			this.btnFFEOr = new System.Windows.Forms.Button();
			this.btnFFEAnd = new System.Windows.Forms.Button();
			this.btnRightBracket = new System.Windows.Forms.Button();
			this.btnLeftBracket = new System.Windows.Forms.Button();
			this.btnLessThanOrEqualTo = new System.Windows.Forms.Button();
			this.btnFFEMoreThanOrEqualTo = new System.Windows.Forms.Button();
			this.btnFFEExpressionBuilderDefault = new System.Windows.Forms.Button();
			this.btnFFELessThan = new System.Windows.Forms.Button();
			this.btnFFEMoreThan = new System.Windows.Forms.Button();
			this.btnFFENotEqual = new System.Windows.Forms.Button();
			this.btnFFEEqual = new System.Windows.Forms.Button();
			this.lstFFEAvailableFields = new System.Windows.Forms.ListBox();
			this.lblFFEExpressionBuilder = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.btnFFEExpressionBuilderClear = new System.Windows.Forms.Button();
			this.btnFFEExpressionBuilderCancel = new System.Windows.Forms.Button();
			this.btnFFEExpressionBuilderOk = new System.Windows.Forms.Button();
			this.btnFFEExpressionBuilderTest = new System.Windows.Forms.Button();
			this.txtExpression = new System.Windows.Forms.TextBox();
			this.grpboxFFE_TI_CI_Effective = new System.Windows.Forms.GroupBox();
			this.btnFFETICIEffectivePreviousExpressions = new System.Windows.Forms.Button();
			this.btnFFEClearExpression = new System.Windows.Forms.Button();
			this.btnFFEExpressionStep2 = new System.Windows.Forms.Button();
			this.listView1 = new System.Windows.Forms.ListView();
			this.btnFFEDefaultStep2 = new System.Windows.Forms.Button();
			this.label12 = new System.Windows.Forms.Label();
			this.grpboxFFEClassifications = new System.Windows.Forms.GroupBox();
			this.label9 = new System.Windows.Forms.Label();
			this.txtFFE_CI5 = new System.Windows.Forms.TextBox();
			this.cmbFFE_CI5 = new System.Windows.Forms.ComboBox();
			this.chkFFE_CI5 = new System.Windows.Forms.CheckBox();
			this.label10 = new System.Windows.Forms.Label();
			this.txtFFE_CI4 = new System.Windows.Forms.TextBox();
			this.cmbFFE_CI4 = new System.Windows.Forms.ComboBox();
			this.chkFFE_CI4 = new System.Windows.Forms.CheckBox();
			this.label11 = new System.Windows.Forms.Label();
			this.txtFFE_CI3 = new System.Windows.Forms.TextBox();
			this.cmbFFE_CI3 = new System.Windows.Forms.ComboBox();
			this.chkFFE_CI3 = new System.Windows.Forms.CheckBox();
			this.label13 = new System.Windows.Forms.Label();
			this.txtFFE_CI2 = new System.Windows.Forms.TextBox();
			this.cmbFFE_CI2 = new System.Windows.Forms.ComboBox();
			this.chkFFE_CI2 = new System.Windows.Forms.CheckBox();
			this.label14 = new System.Windows.Forms.Label();
			this.txtFFE_CI1 = new System.Windows.Forms.TextBox();
			this.cmbFFE_CI1 = new System.Windows.Forms.ComboBox();
			this.chkFFE_CI1 = new System.Windows.Forms.CheckBox();
			this.label8 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtFFE_TI5 = new System.Windows.Forms.TextBox();
			this.cmbFFE_TI5 = new System.Windows.Forms.ComboBox();
			this.chkFFE_TI5 = new System.Windows.Forms.CheckBox();
			this.label6 = new System.Windows.Forms.Label();
			this.txtFFE_TI4 = new System.Windows.Forms.TextBox();
			this.cmbFFE_TI4 = new System.Windows.Forms.ComboBox();
			this.chkFFE_TI4 = new System.Windows.Forms.CheckBox();
			this.label5 = new System.Windows.Forms.Label();
			this.txtFFE_TI3 = new System.Windows.Forms.TextBox();
			this.cmbFFE_TI3 = new System.Windows.Forms.ComboBox();
			this.chkFFE_TI3 = new System.Windows.Forms.CheckBox();
			this.label4 = new System.Windows.Forms.Label();
			this.txtFFE_TI2 = new System.Windows.Forms.TextBox();
			this.cmbFFE_TI2 = new System.Windows.Forms.ComboBox();
			this.chkFFE_TI2 = new System.Windows.Forms.CheckBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.btnFFEDefaultStep1 = new System.Windows.Forms.Button();
			this.txtFFE_TI1 = new System.Windows.Forms.TextBox();
			this.cmbFFE_TI1 = new System.Windows.Forms.ComboBox();
			this.chkFFE_TI1 = new System.Windows.Forms.CheckBox();
			this.grpboxFFEHazard = new System.Windows.Forms.GroupBox();
			this.label24 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.cmbFFECIHazardWindSpeedClass = new System.Windows.Forms.ComboBox();
			this.chkFFECIHazard = new System.Windows.Forms.CheckBox();
			this.cmbFFECIHazardOperator = new System.Windows.Forms.ComboBox();
			this.cmbFFETIHazardWindSpeedClass = new System.Windows.Forms.ComboBox();
			this.chkFFETIHazard = new System.Windows.Forms.CheckBox();
			this.cmbFFETIHazardOperator = new System.Windows.Forms.ComboBox();
			this.label29 = new System.Windows.Forms.Label();
			this.label30 = new System.Windows.Forms.Label();
			this.label31 = new System.Windows.Forms.Label();
			this.btnFFEHazardDefault = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.grpboxFFEBackslide.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.trbFFECIBackslide)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.trbFFETIBackslide)).BeginInit();
			this.grpboxFFEExpressionBuilder.SuspendLayout();
			this.grpboxFFE_TI_CI_Effective.SuspendLayout();
			this.grpboxFFEClassifications.SuspendLayout();
			this.grpboxFFEHazard.SuspendLayout();
			this.SuspendLayout();
			// 
			// imgSize
			// 
			this.imgSize.ImageSize = new System.Drawing.Size(16, 16);
			this.imgSize.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSize.ImageStream")));
			this.imgSize.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Controls.Add(this.grpboxFFEBackslide);
			this.groupBox1.Controls.Add(this.grpboxFFEExpressionBuilder);
			this.groupBox1.Controls.Add(this.grpboxFFE_TI_CI_Effective);
			this.groupBox1.Controls.Add(this.grpboxFFEClassifications);
			this.groupBox1.Controls.Add(this.grpboxFFEHazard);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(744, 1696);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(738, 32);
			this.lblTitle.TabIndex = 27;
			this.lblTitle.Text = "Fuel And Fire Effects";
			// 
			// grpboxFFEBackslide
			// 
			this.grpboxFFEBackslide.BackColor = System.Drawing.Color.White;
			this.grpboxFFEBackslide.Controls.Add(this.label32);
			this.grpboxFFEBackslide.Controls.Add(this.label28);
			this.grpboxFFEBackslide.Controls.Add(this.label26);
			this.grpboxFFEBackslide.Controls.Add(this.label25);
			this.grpboxFFEBackslide.Controls.Add(this.label19);
			this.grpboxFFEBackslide.Controls.Add(this.label18);
			this.grpboxFFEBackslide.Controls.Add(this.label17);
			this.grpboxFFEBackslide.Controls.Add(this.label16);
			this.grpboxFFEBackslide.Controls.Add(this.lblFFECIBackslide);
			this.grpboxFFEBackslide.Controls.Add(this.lblFFETIBackslide);
			this.grpboxFFEBackslide.Controls.Add(this.cmbFFECIBackslide3);
			this.grpboxFFEBackslide.Controls.Add(this.chkFFECIBackslide);
			this.grpboxFFEBackslide.Controls.Add(this.cmbFFECIBackslide2);
			this.grpboxFFEBackslide.Controls.Add(this.cmbFFETIBackslide3);
			this.grpboxFFEBackslide.Controls.Add(this.chkFFETIBackslide);
			this.grpboxFFEBackslide.Controls.Add(this.cmbFFETIBackslide2);
			this.grpboxFFEBackslide.Controls.Add(this.trbFFECIBackslide);
			this.grpboxFFEBackslide.Controls.Add(this.cmbFFECIBackslide);
			this.grpboxFFEBackslide.Controls.Add(this.trbFFETIBackslide);
			this.grpboxFFEBackslide.Controls.Add(this.label21);
			this.grpboxFFEBackslide.Controls.Add(this.label22);
			this.grpboxFFEBackslide.Controls.Add(this.label27);
			this.grpboxFFEBackslide.Controls.Add(this.btnFFEDefaultStep3);
			this.grpboxFFEBackslide.Controls.Add(this.cmbFFETIBackslide);
			this.grpboxFFEBackslide.ForeColor = System.Drawing.Color.Black;
			this.grpboxFFEBackslide.Location = new System.Drawing.Point(112, 936);
			this.grpboxFFEBackslide.Name = "grpboxFFEBackslide";
			this.grpboxFFEBackslide.Size = new System.Drawing.Size(584, 288);
			this.grpboxFFEBackslide.TabIndex = 3;
			this.grpboxFFEBackslide.TabStop = false;
			this.grpboxFFEBackslide.Text = "Backslide Thresholds - Step 3";
			this.grpboxFFEBackslide.Visible = false;
			// 
			// label32
			// 
			this.label32.Location = new System.Drawing.Point(248, 72);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(56, 16);
			this.label32.TabIndex = 66;
			this.label32.Text = "Backslide";
			// 
			// label28
			// 
			this.label28.Location = new System.Drawing.Point(240, 161);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(56, 16);
			this.label28.TabIndex = 65;
			this.label28.Text = "Backslide";
			// 
			// label26
			// 
			this.label26.Location = new System.Drawing.Point(472, 161);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(104, 16);
			this.label26.TabIndex = 64;
			this.label26.Text = "Wind Speed Class";
			// 
			// label25
			// 
			this.label25.Location = new System.Drawing.Point(472, 72);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(104, 16);
			this.label25.TabIndex = 63;
			this.label25.Text = "Wind Speed Class";
			// 
			// label19
			// 
			this.label19.Location = new System.Drawing.Point(301, 189);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(32, 16);
			this.label19.TabIndex = 62;
			this.label19.Text = "AND";
			// 
			// label18
			// 
			this.label18.Location = new System.Drawing.Point(301, 96);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(32, 16);
			this.label18.TabIndex = 61;
			this.label18.Text = "AND";
			// 
			// label17
			// 
			this.label17.Location = new System.Drawing.Point(13, 169);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(208, 16);
			this.label17.TabIndex = 60;
			this.label17.Text = "Post-treatment change in crown index";
			// 
			// label16
			// 
			this.label16.Location = new System.Drawing.Point(13, 80);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(192, 16);
			this.label16.TabIndex = 59;
			this.label16.Text = "Post-treatment change in torch index";
			// 
			// lblFFECIBackslide
			// 
			this.lblFFECIBackslide.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblFFECIBackslide.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFFECIBackslide.Location = new System.Drawing.Point(245, 187);
			this.lblFFECIBackslide.Name = "lblFFECIBackslide";
			this.lblFFECIBackslide.Size = new System.Drawing.Size(40, 16);
			this.lblFFECIBackslide.TabIndex = 58;
			this.lblFFECIBackslide.Text = "-1";
			// 
			// lblFFETIBackslide
			// 
			this.lblFFETIBackslide.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblFFETIBackslide.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFFETIBackslide.Location = new System.Drawing.Point(248, 96);
			this.lblFFETIBackslide.Name = "lblFFETIBackslide";
			this.lblFFETIBackslide.Size = new System.Drawing.Size(40, 16);
			this.lblFFETIBackslide.TabIndex = 57;
			this.lblFFETIBackslide.Text = "-1";
			// 
			// cmbFFECIBackslide3
			// 
			this.cmbFFECIBackslide3.Enabled = false;
			this.cmbFFECIBackslide3.Location = new System.Drawing.Point(491, 187);
			this.cmbFFECIBackslide3.Name = "cmbFFECIBackslide3";
			this.cmbFFECIBackslide3.Size = new System.Drawing.Size(40, 21);
			this.cmbFFECIBackslide3.TabIndex = 56;
			this.cmbFFECIBackslide3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFECIBackslide3_KeyPress);
			this.cmbFFECIBackslide3.TextChanged += new System.EventHandler(this.cmbFFECIBackslide3_TextChanged);
			this.cmbFFECIBackslide3.SelectedValueChanged += new System.EventHandler(this.cmbFFECIBackslide3_SelectedValueChanged);
			// 
			// chkFFECIBackslide
			// 
			this.chkFFECIBackslide.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFECIBackslide.Location = new System.Drawing.Point(336, 189);
			this.chkFFECIBackslide.Name = "chkFFECIBackslide";
			this.chkFFECIBackslide.Size = new System.Drawing.Size(104, 16);
			this.chkFFECIBackslide.TabIndex = 55;
			this.chkFFECIBackslide.Text = "POST_CI_CL";
			this.chkFFECIBackslide.Click += new System.EventHandler(this.chkFFECIBackslide_Click);
			// 
			// cmbFFECIBackslide2
			// 
			this.cmbFFECIBackslide2.Enabled = false;
			this.cmbFFECIBackslide2.Items.AddRange(new object[] {
																	"<",
																	"<=",
																	"="});
			this.cmbFFECIBackslide2.Location = new System.Drawing.Point(442, 187);
			this.cmbFFECIBackslide2.Name = "cmbFFECIBackslide2";
			this.cmbFFECIBackslide2.Size = new System.Drawing.Size(39, 21);
			this.cmbFFECIBackslide2.TabIndex = 54;
			this.cmbFFECIBackslide2.Text = "<";
			this.cmbFFECIBackslide2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFECIBackslide2_KeyPress);
			this.cmbFFECIBackslide2.TextChanged += new System.EventHandler(this.cmbFFECIBackslide2_TextChanged);
			this.cmbFFECIBackslide2.SelectedValueChanged += new System.EventHandler(this.cmbFFECIBackslide2_SelectedValueChanged);
			// 
			// cmbFFETIBackslide3
			// 
			this.cmbFFETIBackslide3.Enabled = false;
			this.cmbFFETIBackslide3.Location = new System.Drawing.Point(491, 96);
			this.cmbFFETIBackslide3.Name = "cmbFFETIBackslide3";
			this.cmbFFETIBackslide3.Size = new System.Drawing.Size(40, 21);
			this.cmbFFETIBackslide3.TabIndex = 53;
			this.cmbFFETIBackslide3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFETIBackslide3_KeyPress);
			this.cmbFFETIBackslide3.TextChanged += new System.EventHandler(this.cmbFFETIBackslide3_TextChanged);
			this.cmbFFETIBackslide3.SelectedValueChanged += new System.EventHandler(this.cmbFFETIBackslide3_SelectedValueChanged);
			// 
			// chkFFETIBackslide
			// 
			this.chkFFETIBackslide.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFETIBackslide.Location = new System.Drawing.Point(335, 96);
			this.chkFFETIBackslide.Name = "chkFFETIBackslide";
			this.chkFFETIBackslide.Size = new System.Drawing.Size(104, 16);
			this.chkFFETIBackslide.TabIndex = 52;
			this.chkFFETIBackslide.Text = "POST_TI_CL";
			this.chkFFETIBackslide.Click += new System.EventHandler(this.chkFFETIBackslide_Click);
			// 
			// cmbFFETIBackslide2
			// 
			this.cmbFFETIBackslide2.Enabled = false;
			this.cmbFFETIBackslide2.Items.AddRange(new object[] {
																	"<",
																	"<=",
																	"="});
			this.cmbFFETIBackslide2.Location = new System.Drawing.Point(441, 96);
			this.cmbFFETIBackslide2.Name = "cmbFFETIBackslide2";
			this.cmbFFETIBackslide2.Size = new System.Drawing.Size(39, 21);
			this.cmbFFETIBackslide2.TabIndex = 50;
			this.cmbFFETIBackslide2.Text = "<";
			this.cmbFFETIBackslide2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFETIBackslide2_KeyPress);
			this.cmbFFETIBackslide2.TextChanged += new System.EventHandler(this.cmbFFETIBackslide2_TextChanged);
			this.cmbFFETIBackslide2.SelectedValueChanged += new System.EventHandler(this.cmbFFETIBackslide2_SelectedValueChanged);
			// 
			// trbFFECIBackslide
			// 
			this.trbFFECIBackslide.Location = new System.Drawing.Point(13, 185);
			this.trbFFECIBackslide.Maximum = 100;
			this.trbFFECIBackslide.Minimum = 1;
			this.trbFFECIBackslide.Name = "trbFFECIBackslide";
			this.trbFFECIBackslide.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.trbFFECIBackslide.Size = new System.Drawing.Size(176, 45);
			this.trbFFECIBackslide.TabIndex = 49;
			this.trbFFECIBackslide.TickStyle = System.Windows.Forms.TickStyle.None;
			this.trbFFECIBackslide.Value = 1;
			this.trbFFECIBackslide.ValueChanged += new System.EventHandler(this.trbFFECIBackslide_ValueChanged);
			// 
			// cmbFFECIBackslide
			// 
			this.cmbFFECIBackslide.ItemHeight = 13;
			this.cmbFFECIBackslide.Items.AddRange(new object[] {
																   "<",
																   "<="});
			this.cmbFFECIBackslide.Location = new System.Drawing.Point(197, 185);
			this.cmbFFECIBackslide.Name = "cmbFFECIBackslide";
			this.cmbFFECIBackslide.Size = new System.Drawing.Size(39, 21);
			this.cmbFFECIBackslide.TabIndex = 47;
			this.cmbFFECIBackslide.Text = "<";
			this.cmbFFECIBackslide.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFECIBackslide_KeyPress);
			this.cmbFFECIBackslide.TextChanged += new System.EventHandler(this.cmbFFECIBackslide_TextChanged);
			this.cmbFFECIBackslide.SelectedValueChanged += new System.EventHandler(this.cmbFFECIBackslide_SelectedValueChanged);
			// 
			// trbFFETIBackslide
			// 
			this.trbFFETIBackslide.Location = new System.Drawing.Point(13, 96);
			this.trbFFETIBackslide.Maximum = 100;
			this.trbFFETIBackslide.Minimum = 1;
			this.trbFFETIBackslide.Name = "trbFFETIBackslide";
			this.trbFFETIBackslide.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.trbFFETIBackslide.Size = new System.Drawing.Size(176, 45);
			this.trbFFETIBackslide.TabIndex = 45;
			this.trbFFETIBackslide.TickStyle = System.Windows.Forms.TickStyle.None;
			this.trbFFETIBackslide.Value = 1;
			this.trbFFETIBackslide.ValueChanged += new System.EventHandler(this.trbFFETIBackslide_ValueChanged);
			// 
			// label21
			// 
			this.label21.BackColor = System.Drawing.Color.White;
			this.label21.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
				| System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label21.ForeColor = System.Drawing.Color.Black;
			this.label21.Location = new System.Drawing.Point(256, 136);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(88, 16);
			this.label21.TabIndex = 25;
			this.label21.Text = "Crown Index";
			// 
			// label22
			// 
			this.label22.BackColor = System.Drawing.Color.White;
			this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
				| System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.ForeColor = System.Drawing.Color.Black;
			this.label22.Location = new System.Drawing.Point(256, 48);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(80, 16);
			this.label22.TabIndex = 24;
			this.label22.Text = "Torch Index";
			// 
			// label27
			// 
			this.label27.BackColor = System.Drawing.Color.White;
			this.label27.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label27.ForeColor = System.Drawing.Color.Black;
			this.label27.Location = new System.Drawing.Point(16, 16);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(552, 32);
			this.label27.TabIndex = 6;
			this.label27.Text = "Define what constiutes a post treatment backslide.";
			// 
			// btnFFEDefaultStep3
			// 
			this.btnFFEDefaultStep3.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEDefaultStep3.Location = new System.Drawing.Point(8, 256);
			this.btnFFEDefaultStep3.Name = "btnFFEDefaultStep3";
			this.btnFFEDefaultStep3.Size = new System.Drawing.Size(128, 24);
			this.btnFFEDefaultStep3.TabIndex = 4;
			this.btnFFEDefaultStep3.Text = "Use Default Values";
			this.btnFFEDefaultStep3.Click += new System.EventHandler(this.btnFFEDefaultStep3_Click);
			// 
			// cmbFFETIBackslide
			// 
			this.cmbFFETIBackslide.ItemHeight = 13;
			this.cmbFFETIBackslide.Items.AddRange(new object[] {
																   "<",
																   "<="});
			this.cmbFFETIBackslide.Location = new System.Drawing.Point(197, 96);
			this.cmbFFETIBackslide.Name = "cmbFFETIBackslide";
			this.cmbFFETIBackslide.Size = new System.Drawing.Size(39, 21);
			this.cmbFFETIBackslide.TabIndex = 2;
			this.cmbFFETIBackslide.Text = "<";
			this.cmbFFETIBackslide.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFETIBackslide_KeyPress);
			this.cmbFFETIBackslide.TextChanged += new System.EventHandler(this.cmbFFETIBackslide_TextChanged);
			this.cmbFFETIBackslide.SelectedValueChanged += new System.EventHandler(this.cmbFFETIBackslide_SelectedValueChanged);
			// 
			// grpboxFFEExpressionBuilder
			// 
			this.grpboxFFEExpressionBuilder.BackColor = System.Drawing.Color.White;
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnN);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnY);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEExpressionBuilderPreviousExpression);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFENote);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.lblFFEExpressionFieldDesc);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEOr);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEAnd);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnRightBracket);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnLeftBracket);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnLessThanOrEqualTo);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEMoreThanOrEqualTo);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEExpressionBuilderDefault);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFELessThan);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEMoreThan);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFENotEqual);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEEqual);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.lstFFEAvailableFields);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.lblFFEExpressionBuilder);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.label15);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEExpressionBuilderClear);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEExpressionBuilderCancel);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEExpressionBuilderOk);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.btnFFEExpressionBuilderTest);
			this.grpboxFFEExpressionBuilder.Controls.Add(this.txtExpression);
			this.grpboxFFEExpressionBuilder.ForeColor = System.Drawing.Color.Black;
			this.grpboxFFEExpressionBuilder.Location = new System.Drawing.Point(112, 640);
			this.grpboxFFEExpressionBuilder.Name = "grpboxFFEExpressionBuilder";
			this.grpboxFFEExpressionBuilder.Size = new System.Drawing.Size(584, 288);
			this.grpboxFFEExpressionBuilder.TabIndex = 2;
			this.grpboxFFEExpressionBuilder.TabStop = false;
			this.grpboxFFEExpressionBuilder.Text = "Expression Builder";
			this.grpboxFFEExpressionBuilder.Visible = false;
			// 
			// btnN
			// 
			this.btnN.BackColor = System.Drawing.SystemColors.Control;
			this.btnN.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnN.Location = new System.Drawing.Point(464, 85);
			this.btnN.Name = "btnN";
			this.btnN.Size = new System.Drawing.Size(32, 24);
			this.btnN.TabIndex = 31;
			this.btnN.Text = "\'N\'";
			this.btnN.Click += new System.EventHandler(this.btnN_Click);
			// 
			// btnY
			// 
			this.btnY.BackColor = System.Drawing.SystemColors.Control;
			this.btnY.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnY.Location = new System.Drawing.Point(464, 62);
			this.btnY.Name = "btnY";
			this.btnY.Size = new System.Drawing.Size(32, 24);
			this.btnY.TabIndex = 30;
			this.btnY.Text = "\'Y\'";
			this.btnY.Click += new System.EventHandler(this.btnY_Click);
			// 
			// btnFFEExpressionBuilderPreviousExpression
			// 
			this.btnFFEExpressionBuilderPreviousExpression.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEExpressionBuilderPreviousExpression.Location = new System.Drawing.Point(38, 216);
			this.btnFFEExpressionBuilderPreviousExpression.Name = "btnFFEExpressionBuilderPreviousExpression";
			this.btnFFEExpressionBuilderPreviousExpression.Size = new System.Drawing.Size(128, 24);
			this.btnFFEExpressionBuilderPreviousExpression.TabIndex = 29;
			this.btnFFEExpressionBuilderPreviousExpression.Text = "Previous Expressions";
			this.btnFFEExpressionBuilderPreviousExpression.Click += new System.EventHandler(this.btnFFEPreviousExpression_Click);
			// 
			// btnFFENote
			// 
			this.btnFFENote.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFENote.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFFENote.Location = new System.Drawing.Point(424, 62);
			this.btnFFENote.Name = "btnFFENote";
			this.btnFFENote.Size = new System.Drawing.Size(40, 24);
			this.btnFFENote.TabIndex = 28;
			this.btnFFENote.Text = "NOT";
			this.btnFFENote.Click += new System.EventHandler(this.btnFFENote_Click);
			// 
			// lblFFEExpressionFieldDesc
			// 
			this.lblFFEExpressionFieldDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFFEExpressionFieldDesc.ForeColor = System.Drawing.Color.Blue;
			this.lblFFEExpressionFieldDesc.Location = new System.Drawing.Point(39, 136);
			this.lblFFEExpressionFieldDesc.Name = "lblFFEExpressionFieldDesc";
			this.lblFFEExpressionFieldDesc.Size = new System.Drawing.Size(505, 16);
			this.lblFFEExpressionFieldDesc.TabIndex = 27;
			this.lblFFEExpressionFieldDesc.Text = "Field Description";
			// 
			// btnFFEOr
			// 
			this.btnFFEOr.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEOr.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFFEOr.Location = new System.Drawing.Point(424, 109);
			this.btnFFEOr.Name = "btnFFEOr";
			this.btnFFEOr.Size = new System.Drawing.Size(40, 24);
			this.btnFFEOr.TabIndex = 26;
			this.btnFFEOr.Text = "OR";
			this.btnFFEOr.Click += new System.EventHandler(this.btnFFEOr_Click);
			// 
			// btnFFEAnd
			// 
			this.btnFFEAnd.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEAnd.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFFEAnd.Location = new System.Drawing.Point(424, 85);
			this.btnFFEAnd.Name = "btnFFEAnd";
			this.btnFFEAnd.Size = new System.Drawing.Size(40, 24);
			this.btnFFEAnd.TabIndex = 25;
			this.btnFFEAnd.Text = "AND";
			this.btnFFEAnd.Click += new System.EventHandler(this.btnFFEAnd_Click);
			// 
			// btnRightBracket
			// 
			this.btnRightBracket.BackColor = System.Drawing.SystemColors.Control;
			this.btnRightBracket.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnRightBracket.Location = new System.Drawing.Point(384, 62);
			this.btnRightBracket.Name = "btnRightBracket";
			this.btnRightBracket.Size = new System.Drawing.Size(40, 24);
			this.btnRightBracket.TabIndex = 24;
			this.btnRightBracket.Text = ")";
			this.btnRightBracket.Click += new System.EventHandler(this.btnRightBracket_Click);
			// 
			// btnLeftBracket
			// 
			this.btnLeftBracket.BackColor = System.Drawing.SystemColors.Control;
			this.btnLeftBracket.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnLeftBracket.Location = new System.Drawing.Point(344, 62);
			this.btnLeftBracket.Name = "btnLeftBracket";
			this.btnLeftBracket.Size = new System.Drawing.Size(40, 24);
			this.btnLeftBracket.TabIndex = 23;
			this.btnLeftBracket.Text = "(";
			this.btnLeftBracket.Click += new System.EventHandler(this.btnLeftBracket_Click);
			// 
			// btnLessThanOrEqualTo
			// 
			this.btnLessThanOrEqualTo.BackColor = System.Drawing.SystemColors.Control;
			this.btnLessThanOrEqualTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnLessThanOrEqualTo.Location = new System.Drawing.Point(296, 109);
			this.btnLessThanOrEqualTo.Name = "btnLessThanOrEqualTo";
			this.btnLessThanOrEqualTo.Size = new System.Drawing.Size(128, 24);
			this.btnLessThanOrEqualTo.TabIndex = 22;
			this.btnLessThanOrEqualTo.Text = "Less Than Or Equal To";
			this.btnLessThanOrEqualTo.Click += new System.EventHandler(this.btnLessThanOrEqualTo_Click);
			// 
			// btnFFEMoreThanOrEqualTo
			// 
			this.btnFFEMoreThanOrEqualTo.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEMoreThanOrEqualTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFFEMoreThanOrEqualTo.Location = new System.Drawing.Point(296, 85);
			this.btnFFEMoreThanOrEqualTo.Name = "btnFFEMoreThanOrEqualTo";
			this.btnFFEMoreThanOrEqualTo.Size = new System.Drawing.Size(128, 24);
			this.btnFFEMoreThanOrEqualTo.TabIndex = 21;
			this.btnFFEMoreThanOrEqualTo.Text = "More Than Or Equal To";
			this.btnFFEMoreThanOrEqualTo.Click += new System.EventHandler(this.btnFFEMoreThanOrEqualTo_Click);
			// 
			// btnFFEExpressionBuilderDefault
			// 
			this.btnFFEExpressionBuilderDefault.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEExpressionBuilderDefault.Location = new System.Drawing.Point(8, 256);
			this.btnFFEExpressionBuilderDefault.Name = "btnFFEExpressionBuilderDefault";
			this.btnFFEExpressionBuilderDefault.Size = new System.Drawing.Size(128, 24);
			this.btnFFEExpressionBuilderDefault.TabIndex = 20;
			this.btnFFEExpressionBuilderDefault.Text = "Use Default Values";
			this.btnFFEExpressionBuilderDefault.Click += new System.EventHandler(this.btnFFEExpressionBuilderDefault_Click);
			// 
			// btnFFELessThan
			// 
			this.btnFFELessThan.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFELessThan.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFFELessThan.Location = new System.Drawing.Point(232, 109);
			this.btnFFELessThan.Name = "btnFFELessThan";
			this.btnFFELessThan.Size = new System.Drawing.Size(64, 24);
			this.btnFFELessThan.TabIndex = 19;
			this.btnFFELessThan.Text = "Less Than";
			this.btnFFELessThan.Click += new System.EventHandler(this.btnFFELessThan_Click);
			// 
			// btnFFEMoreThan
			// 
			this.btnFFEMoreThan.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEMoreThan.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFFEMoreThan.Location = new System.Drawing.Point(232, 86);
			this.btnFFEMoreThan.Name = "btnFFEMoreThan";
			this.btnFFEMoreThan.Size = new System.Drawing.Size(64, 24);
			this.btnFFEMoreThan.TabIndex = 18;
			this.btnFFEMoreThan.Text = "More Than";
			this.btnFFEMoreThan.Click += new System.EventHandler(this.btnFFEMoreThan_Click);
			// 
			// btnFFENotEqual
			// 
			this.btnFFENotEqual.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFENotEqual.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFFENotEqual.Location = new System.Drawing.Point(288, 62);
			this.btnFFENotEqual.Name = "btnFFENotEqual";
			this.btnFFENotEqual.Size = new System.Drawing.Size(56, 24);
			this.btnFFENotEqual.TabIndex = 17;
			this.btnFFENotEqual.Text = "Not Equal";
			this.btnFFENotEqual.Click += new System.EventHandler(this.btnFFENotEqual_Click);
			// 
			// btnFFEEqual
			// 
			this.btnFFEEqual.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEEqual.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFFEEqual.Location = new System.Drawing.Point(232, 62);
			this.btnFFEEqual.Name = "btnFFEEqual";
			this.btnFFEEqual.Size = new System.Drawing.Size(56, 24);
			this.btnFFEEqual.TabIndex = 16;
			this.btnFFEEqual.Text = "Equal";
			this.btnFFEEqual.Click += new System.EventHandler(this.btnFFEEqual_Click);
			// 
			// lstFFEAvailableFields
			// 
			this.lstFFEAvailableFields.Location = new System.Drawing.Point(44, 62);
			this.lstFFEAvailableFields.Name = "lstFFEAvailableFields";
			this.lstFFEAvailableFields.Size = new System.Drawing.Size(132, 69);
			this.lstFFEAvailableFields.TabIndex = 15;
			this.lstFFEAvailableFields.DoubleClick += new System.EventHandler(this.lstFFEAvailableFields_DoubleClick);
			this.lstFFEAvailableFields.SelectedValueChanged += new System.EventHandler(this.lstFFEAvailableFields_SelectedValueChanged);
			// 
			// lblFFEExpressionBuilder
			// 
			this.lblFFEExpressionBuilder.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFFEExpressionBuilder.ForeColor = System.Drawing.Color.Black;
			this.lblFFEExpressionBuilder.Location = new System.Drawing.Point(16, 15);
			this.lblFFEExpressionBuilder.Name = "lblFFEExpressionBuilder";
			this.lblFFEExpressionBuilder.Size = new System.Drawing.Size(552, 32);
			this.lblFFEExpressionBuilder.TabIndex = 14;
			this.lblFFEExpressionBuilder.Text = "Define what is an effective treatment when the torch index pre-treatments sustain" +
				"able fire wind speed class is equal to 1.";
			// 
			// label15
			// 
			this.label15.Location = new System.Drawing.Point(44, 46);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(112, 13);
			this.label15.TabIndex = 13;
			this.label15.Text = "Available Fields";
			// 
			// btnFFEExpressionBuilderClear
			// 
			this.btnFFEExpressionBuilderClear.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEExpressionBuilderClear.Location = new System.Drawing.Point(166, 216);
			this.btnFFEExpressionBuilderClear.Name = "btnFFEExpressionBuilderClear";
			this.btnFFEExpressionBuilderClear.Size = new System.Drawing.Size(128, 24);
			this.btnFFEExpressionBuilderClear.TabIndex = 12;
			this.btnFFEExpressionBuilderClear.Text = "Clear Expression";
			this.btnFFEExpressionBuilderClear.Click += new System.EventHandler(this.btnFFEExpressionBuilderClear_Click);
			// 
			// btnFFEExpressionBuilderCancel
			// 
			this.btnFFEExpressionBuilderCancel.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEExpressionBuilderCancel.Location = new System.Drawing.Point(438, 216);
			this.btnFFEExpressionBuilderCancel.Name = "btnFFEExpressionBuilderCancel";
			this.btnFFEExpressionBuilderCancel.Size = new System.Drawing.Size(72, 24);
			this.btnFFEExpressionBuilderCancel.TabIndex = 11;
			this.btnFFEExpressionBuilderCancel.Text = "Cancel";
			this.btnFFEExpressionBuilderCancel.Click += new System.EventHandler(this.btnFFECancel_Click);
			// 
			// btnFFEExpressionBuilderOk
			// 
			this.btnFFEExpressionBuilderOk.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEExpressionBuilderOk.Location = new System.Drawing.Point(366, 216);
			this.btnFFEExpressionBuilderOk.Name = "btnFFEExpressionBuilderOk";
			this.btnFFEExpressionBuilderOk.Size = new System.Drawing.Size(72, 24);
			this.btnFFEExpressionBuilderOk.TabIndex = 10;
			this.btnFFEExpressionBuilderOk.Text = "OK";
			this.btnFFEExpressionBuilderOk.Click += new System.EventHandler(this.btnFFEExpressionBuilderOk_Click);
			// 
			// btnFFEExpressionBuilderTest
			// 
			this.btnFFEExpressionBuilderTest.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEExpressionBuilderTest.Location = new System.Drawing.Point(294, 216);
			this.btnFFEExpressionBuilderTest.Name = "btnFFEExpressionBuilderTest";
			this.btnFFEExpressionBuilderTest.Size = new System.Drawing.Size(72, 24);
			this.btnFFEExpressionBuilderTest.TabIndex = 9;
			this.btnFFEExpressionBuilderTest.Text = "Test";
			this.btnFFEExpressionBuilderTest.Click += new System.EventHandler(this.btnFFEExpressionBuilderTest_Click);
			// 
			// txtExpression
			// 
			this.txtExpression.Location = new System.Drawing.Point(40, 155);
			this.txtExpression.Multiline = true;
			this.txtExpression.Name = "txtExpression";
			this.txtExpression.Size = new System.Drawing.Size(472, 53);
			this.txtExpression.TabIndex = 8;
			this.txtExpression.Text = "";
			this.txtExpression.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtExpression_KeyPress);
			this.txtExpression.Leave += new System.EventHandler(this.txtExpression_Leave);
			// 
			// grpboxFFE_TI_CI_Effective
			// 
			this.grpboxFFE_TI_CI_Effective.BackColor = System.Drawing.Color.White;
			this.grpboxFFE_TI_CI_Effective.Controls.Add(this.btnFFETICIEffectivePreviousExpressions);
			this.grpboxFFE_TI_CI_Effective.Controls.Add(this.btnFFEClearExpression);
			this.grpboxFFE_TI_CI_Effective.Controls.Add(this.btnFFEExpressionStep2);
			this.grpboxFFE_TI_CI_Effective.Controls.Add(this.listView1);
			this.grpboxFFE_TI_CI_Effective.Controls.Add(this.btnFFEDefaultStep2);
			this.grpboxFFE_TI_CI_Effective.Controls.Add(this.label12);
			this.grpboxFFE_TI_CI_Effective.ForeColor = System.Drawing.Color.Black;
			this.grpboxFFE_TI_CI_Effective.Location = new System.Drawing.Point(112, 344);
			this.grpboxFFE_TI_CI_Effective.Name = "grpboxFFE_TI_CI_Effective";
			this.grpboxFFE_TI_CI_Effective.Size = new System.Drawing.Size(584, 288);
			this.grpboxFFE_TI_CI_Effective.TabIndex = 1;
			this.grpboxFFE_TI_CI_Effective.TabStop = false;
			this.grpboxFFE_TI_CI_Effective.Text = "Effective Torch And Crown Index Change - Step 2";
			this.grpboxFFE_TI_CI_Effective.Visible = false;
			this.grpboxFFE_TI_CI_Effective.Enter += new System.EventHandler(this.grpboxFFE_TI_CI_Effective_Enter);
			// 
			// btnFFETICIEffectivePreviousExpressions
			// 
			this.btnFFETICIEffectivePreviousExpressions.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFETICIEffectivePreviousExpressions.Enabled = false;
			this.btnFFETICIEffectivePreviousExpressions.Location = new System.Drawing.Point(128, 216);
			this.btnFFETICIEffectivePreviousExpressions.Name = "btnFFETICIEffectivePreviousExpressions";
			this.btnFFETICIEffectivePreviousExpressions.Size = new System.Drawing.Size(128, 24);
			this.btnFFETICIEffectivePreviousExpressions.TabIndex = 31;
			this.btnFFETICIEffectivePreviousExpressions.Text = "Previous Expressions";
			this.btnFFETICIEffectivePreviousExpressions.Click += new System.EventHandler(this.btnFFETICIEffectivePreviousExpressions_Click);
			// 
			// btnFFEClearExpression
			// 
			this.btnFFEClearExpression.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEClearExpression.Enabled = false;
			this.btnFFEClearExpression.Location = new System.Drawing.Point(256, 216);
			this.btnFFEClearExpression.Name = "btnFFEClearExpression";
			this.btnFFEClearExpression.Size = new System.Drawing.Size(112, 24);
			this.btnFFEClearExpression.TabIndex = 30;
			this.btnFFEClearExpression.Text = "Clear Expression";
			this.btnFFEClearExpression.Click += new System.EventHandler(this.btnFFEClearExpression_Click);
			// 
			// btnFFEExpressionStep2
			// 
			this.btnFFEExpressionStep2.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEExpressionStep2.Enabled = false;
			this.btnFFEExpressionStep2.Location = new System.Drawing.Point(368, 216);
			this.btnFFEExpressionStep2.Name = "btnFFEExpressionStep2";
			this.btnFFEExpressionStep2.Size = new System.Drawing.Size(112, 24);
			this.btnFFEExpressionStep2.TabIndex = 29;
			this.btnFFEExpressionStep2.Text = "Expression Builder";
			this.btnFFEExpressionStep2.Click += new System.EventHandler(this.btnFFEExpressionStep2_Click);
			// 
			// listView1
			// 
			this.listView1.FullRowSelect = true;
			this.listView1.GridLines = true;
			this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
			this.listView1.HideSelection = false;
			this.listView1.Location = new System.Drawing.Point(24, 72);
			this.listView1.MultiSelect = false;
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(536, 136);
			this.listView1.TabIndex = 28;
			this.listView1.View = System.Windows.Forms.View.Details;
			this.listView1.Click += new System.EventHandler(this.listView1_Click);
			// 
			// btnFFEDefaultStep2
			// 
			this.btnFFEDefaultStep2.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEDefaultStep2.Location = new System.Drawing.Point(8, 256);
			this.btnFFEDefaultStep2.Name = "btnFFEDefaultStep2";
			this.btnFFEDefaultStep2.Size = new System.Drawing.Size(128, 24);
			this.btnFFEDefaultStep2.TabIndex = 7;
			this.btnFFEDefaultStep2.Text = "Use Default Values";
			this.btnFFEDefaultStep2.Click += new System.EventHandler(this.btnFFEDefaultStep2_Click);
			// 
			// label12
			// 
			this.label12.BackColor = System.Drawing.Color.White;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.Black;
			this.label12.Location = new System.Drawing.Point(16, 24);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(552, 32);
			this.label12.TabIndex = 6;
			this.label12.Text = "Define what constitutes effective treatment results for reducing torch and crown " +
				"index fire potential.";
			// 
			// grpboxFFEClassifications
			// 
			this.grpboxFFEClassifications.BackColor = System.Drawing.Color.White;
			this.grpboxFFEClassifications.Controls.Add(this.label9);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_CI5);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_CI5);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_CI5);
			this.grpboxFFEClassifications.Controls.Add(this.label10);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_CI4);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_CI4);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_CI4);
			this.grpboxFFEClassifications.Controls.Add(this.label11);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_CI3);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_CI3);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_CI3);
			this.grpboxFFEClassifications.Controls.Add(this.label13);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_CI2);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_CI2);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_CI2);
			this.grpboxFFEClassifications.Controls.Add(this.label14);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_CI1);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_CI1);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_CI1);
			this.grpboxFFEClassifications.Controls.Add(this.label8);
			this.grpboxFFEClassifications.Controls.Add(this.label1);
			this.grpboxFFEClassifications.Controls.Add(this.label7);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_TI5);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_TI5);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_TI5);
			this.grpboxFFEClassifications.Controls.Add(this.label6);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_TI4);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_TI4);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_TI4);
			this.grpboxFFEClassifications.Controls.Add(this.label5);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_TI3);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_TI3);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_TI3);
			this.grpboxFFEClassifications.Controls.Add(this.label4);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_TI2);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_TI2);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_TI2);
			this.grpboxFFEClassifications.Controls.Add(this.label3);
			this.grpboxFFEClassifications.Controls.Add(this.label2);
			this.grpboxFFEClassifications.Controls.Add(this.btnFFEDefaultStep1);
			this.grpboxFFEClassifications.Controls.Add(this.txtFFE_TI1);
			this.grpboxFFEClassifications.Controls.Add(this.cmbFFE_TI1);
			this.grpboxFFEClassifications.Controls.Add(this.chkFFE_TI1);
			this.grpboxFFEClassifications.ForeColor = System.Drawing.Color.Black;
			this.grpboxFFEClassifications.Location = new System.Drawing.Point(112, 48);
			this.grpboxFFEClassifications.Name = "grpboxFFEClassifications";
			this.grpboxFFEClassifications.Size = new System.Drawing.Size(584, 288);
			this.grpboxFFEClassifications.TabIndex = 0;
			this.grpboxFFEClassifications.TabStop = false;
			this.grpboxFFEClassifications.Text = "Wind Speed Fire Potential Classifications - Step 1";
			this.grpboxFFEClassifications.Enter += new System.EventHandler(this.grpboxFFEIndices_Enter);
			// 
			// label9
			// 
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(446, 188);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(32, 16);
			this.label9.TabIndex = 45;
			this.label9.Text = "MPH";
			// 
			// txtFFE_CI5
			// 
			this.txtFFE_CI5.Enabled = false;
			this.txtFFE_CI5.Location = new System.Drawing.Point(398, 188);
			this.txtFFE_CI5.MaxLength = 3;
			this.txtFFE_CI5.Name = "txtFFE_CI5";
			this.txtFFE_CI5.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_CI5.TabIndex = 44;
			this.txtFFE_CI5.Tag = "999";
			this.txtFFE_CI5.Text = "";
			this.txtFFE_CI5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_CI5_KeyPress);
			// 
			// cmbFFE_CI5
			// 
			this.cmbFFE_CI5.Enabled = false;
			this.cmbFFE_CI5.Items.AddRange(new object[] {
															"<",
															"<=",
															">",
															">="});
			this.cmbFFE_CI5.Location = new System.Drawing.Point(342, 188);
			this.cmbFFE_CI5.Name = "cmbFFE_CI5";
			this.cmbFFE_CI5.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_CI5.TabIndex = 43;
			this.cmbFFE_CI5.Text = "<";
			this.cmbFFE_CI5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_CI5_KeyPress);
			// 
			// chkFFE_CI5
			// 
			this.chkFFE_CI5.Enabled = false;
			this.chkFFE_CI5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_CI5.Location = new System.Drawing.Point(302, 188);
			this.chkFFE_CI5.Name = "chkFFE_CI5";
			this.chkFFE_CI5.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_CI5.TabIndex = 42;
			this.chkFFE_CI5.Text = "5";
			this.chkFFE_CI5.Click += new System.EventHandler(this.chkFFE_CI5_Click);
			// 
			// label10
			// 
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(446, 162);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(32, 16);
			this.label10.TabIndex = 41;
			this.label10.Text = "MPH";
			// 
			// txtFFE_CI4
			// 
			this.txtFFE_CI4.Enabled = false;
			this.txtFFE_CI4.Location = new System.Drawing.Point(398, 162);
			this.txtFFE_CI4.MaxLength = 3;
			this.txtFFE_CI4.Name = "txtFFE_CI4";
			this.txtFFE_CI4.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_CI4.TabIndex = 40;
			this.txtFFE_CI4.Tag = "999";
			this.txtFFE_CI4.Text = "";
			this.txtFFE_CI4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_CI4_KeyPress);
			// 
			// cmbFFE_CI4
			// 
			this.cmbFFE_CI4.Enabled = false;
			this.cmbFFE_CI4.Items.AddRange(new object[] {
															"<",
															"<=",
															">",
															">="});
			this.cmbFFE_CI4.Location = new System.Drawing.Point(342, 162);
			this.cmbFFE_CI4.Name = "cmbFFE_CI4";
			this.cmbFFE_CI4.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_CI4.TabIndex = 39;
			this.cmbFFE_CI4.Text = "<";
			this.cmbFFE_CI4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_CI4_KeyPress);
			// 
			// chkFFE_CI4
			// 
			this.chkFFE_CI4.Enabled = false;
			this.chkFFE_CI4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_CI4.Location = new System.Drawing.Point(302, 162);
			this.chkFFE_CI4.Name = "chkFFE_CI4";
			this.chkFFE_CI4.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_CI4.TabIndex = 38;
			this.chkFFE_CI4.Text = "4";
			this.chkFFE_CI4.Click += new System.EventHandler(this.chkFFE_CI4_Click);
			// 
			// label11
			// 
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(446, 138);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(32, 16);
			this.label11.TabIndex = 37;
			this.label11.Text = "MPH";
			// 
			// txtFFE_CI3
			// 
			this.txtFFE_CI3.Enabled = false;
			this.txtFFE_CI3.Location = new System.Drawing.Point(398, 138);
			this.txtFFE_CI3.MaxLength = 3;
			this.txtFFE_CI3.Name = "txtFFE_CI3";
			this.txtFFE_CI3.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_CI3.TabIndex = 36;
			this.txtFFE_CI3.Tag = "999";
			this.txtFFE_CI3.Text = "";
			this.txtFFE_CI3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_CI3_KeyPress);
			// 
			// cmbFFE_CI3
			// 
			this.cmbFFE_CI3.Enabled = false;
			this.cmbFFE_CI3.Items.AddRange(new object[] {
															"<",
															"<=",
															">",
															">="});
			this.cmbFFE_CI3.Location = new System.Drawing.Point(342, 138);
			this.cmbFFE_CI3.Name = "cmbFFE_CI3";
			this.cmbFFE_CI3.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_CI3.TabIndex = 35;
			this.cmbFFE_CI3.Text = "<";
			this.cmbFFE_CI3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_CI3_KeyPress);
			this.cmbFFE_CI3.SelectedIndexChanged += new System.EventHandler(this.cmbFFE_CI3_SelectedIndexChanged);
			// 
			// chkFFE_CI3
			// 
			this.chkFFE_CI3.Enabled = false;
			this.chkFFE_CI3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_CI3.Location = new System.Drawing.Point(302, 138);
			this.chkFFE_CI3.Name = "chkFFE_CI3";
			this.chkFFE_CI3.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_CI3.TabIndex = 34;
			this.chkFFE_CI3.Text = "3";
			this.chkFFE_CI3.Click += new System.EventHandler(this.chkFFE_CI3_Click);
			// 
			// label13
			// 
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(446, 120);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(32, 16);
			this.label13.TabIndex = 33;
			this.label13.Text = "MPH";
			// 
			// txtFFE_CI2
			// 
			this.txtFFE_CI2.Enabled = false;
			this.txtFFE_CI2.Location = new System.Drawing.Point(398, 112);
			this.txtFFE_CI2.MaxLength = 3;
			this.txtFFE_CI2.Name = "txtFFE_CI2";
			this.txtFFE_CI2.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_CI2.TabIndex = 32;
			this.txtFFE_CI2.Tag = "999";
			this.txtFFE_CI2.Text = "";
			this.txtFFE_CI2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_CI2_KeyPress);
			// 
			// cmbFFE_CI2
			// 
			this.cmbFFE_CI2.Enabled = false;
			this.cmbFFE_CI2.Items.AddRange(new object[] {
															"<",
															"<=",
															">",
															">="});
			this.cmbFFE_CI2.Location = new System.Drawing.Point(342, 112);
			this.cmbFFE_CI2.Name = "cmbFFE_CI2";
			this.cmbFFE_CI2.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_CI2.TabIndex = 31;
			this.cmbFFE_CI2.Text = "<";
			this.cmbFFE_CI2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_CI2_KeyPress);
			// 
			// chkFFE_CI2
			// 
			this.chkFFE_CI2.Enabled = false;
			this.chkFFE_CI2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_CI2.Location = new System.Drawing.Point(302, 112);
			this.chkFFE_CI2.Name = "chkFFE_CI2";
			this.chkFFE_CI2.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_CI2.TabIndex = 30;
			this.chkFFE_CI2.Text = "2";
			this.chkFFE_CI2.Click += new System.EventHandler(this.chkFFE_CI2_Click);
			// 
			// label14
			// 
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(446, 96);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(32, 16);
			this.label14.TabIndex = 29;
			this.label14.Text = "MPH";
			// 
			// txtFFE_CI1
			// 
			this.txtFFE_CI1.Enabled = false;
			this.txtFFE_CI1.Location = new System.Drawing.Point(398, 88);
			this.txtFFE_CI1.MaxLength = 3;
			this.txtFFE_CI1.Name = "txtFFE_CI1";
			this.txtFFE_CI1.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_CI1.TabIndex = 28;
			this.txtFFE_CI1.Tag = "999";
			this.txtFFE_CI1.Text = "";
			this.txtFFE_CI1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_CI1_KeyPress);
			// 
			// cmbFFE_CI1
			// 
			this.cmbFFE_CI1.Enabled = false;
			this.cmbFFE_CI1.Items.AddRange(new object[] {
															"<",
															"<="});
			this.cmbFFE_CI1.Location = new System.Drawing.Point(342, 88);
			this.cmbFFE_CI1.Name = "cmbFFE_CI1";
			this.cmbFFE_CI1.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_CI1.TabIndex = 27;
			this.cmbFFE_CI1.Text = "<";
			this.cmbFFE_CI1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_CI1_KeyPress);
			// 
			// chkFFE_CI1
			// 
			this.chkFFE_CI1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_CI1.Location = new System.Drawing.Point(302, 88);
			this.chkFFE_CI1.Name = "chkFFE_CI1";
			this.chkFFE_CI1.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_CI1.TabIndex = 26;
			this.chkFFE_CI1.Text = "1";
			this.chkFFE_CI1.Click += new System.EventHandler(this.chkFFE_CI1_Click);
			// 
			// label8
			// 
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
				| System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Black;
			this.label8.Location = new System.Drawing.Point(350, 56);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(88, 16);
			this.label8.TabIndex = 25;
			this.label8.Text = "Crown Index";
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
				| System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Black;
			this.label1.Location = new System.Drawing.Point(132, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 16);
			this.label1.TabIndex = 24;
			this.label1.Text = "Torch Index";
			// 
			// label7
			// 
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(227, 188);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(32, 16);
			this.label7.TabIndex = 22;
			this.label7.Text = "MPH";
			// 
			// txtFFE_TI5
			// 
			this.txtFFE_TI5.Enabled = false;
			this.txtFFE_TI5.Location = new System.Drawing.Point(178, 188);
			this.txtFFE_TI5.MaxLength = 3;
			this.txtFFE_TI5.Name = "txtFFE_TI5";
			this.txtFFE_TI5.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_TI5.TabIndex = 21;
			this.txtFFE_TI5.Tag = "999";
			this.txtFFE_TI5.Text = "";
			this.txtFFE_TI5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_TI5_KeyPress);
			// 
			// cmbFFE_TI5
			// 
			this.cmbFFE_TI5.Enabled = false;
			this.cmbFFE_TI5.Items.AddRange(new object[] {
															"<",
															"<=",
															">",
															">="});
			this.cmbFFE_TI5.Location = new System.Drawing.Point(121, 188);
			this.cmbFFE_TI5.Name = "cmbFFE_TI5";
			this.cmbFFE_TI5.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_TI5.TabIndex = 20;
			this.cmbFFE_TI5.Text = "<";
			this.cmbFFE_TI5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_TI5_KeyPress);
			// 
			// chkFFE_TI5
			// 
			this.chkFFE_TI5.Enabled = false;
			this.chkFFE_TI5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_TI5.Location = new System.Drawing.Point(82, 188);
			this.chkFFE_TI5.Name = "chkFFE_TI5";
			this.chkFFE_TI5.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_TI5.TabIndex = 19;
			this.chkFFE_TI5.Text = "5";
			this.chkFFE_TI5.Click += new System.EventHandler(this.chkFFE_TI5_Click);
			// 
			// label6
			// 
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(227, 162);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(32, 16);
			this.label6.TabIndex = 18;
			this.label6.Text = "MPH";
			this.label6.Click += new System.EventHandler(this.label6_Click);
			// 
			// txtFFE_TI4
			// 
			this.txtFFE_TI4.Enabled = false;
			this.txtFFE_TI4.Location = new System.Drawing.Point(178, 162);
			this.txtFFE_TI4.MaxLength = 3;
			this.txtFFE_TI4.Name = "txtFFE_TI4";
			this.txtFFE_TI4.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_TI4.TabIndex = 17;
			this.txtFFE_TI4.Tag = "999";
			this.txtFFE_TI4.Text = "";
			this.txtFFE_TI4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_TI4_KeyPress);
			// 
			// cmbFFE_TI4
			// 
			this.cmbFFE_TI4.Enabled = false;
			this.cmbFFE_TI4.Items.AddRange(new object[] {
															"<",
															"<=",
															">",
															">="});
			this.cmbFFE_TI4.Location = new System.Drawing.Point(121, 162);
			this.cmbFFE_TI4.Name = "cmbFFE_TI4";
			this.cmbFFE_TI4.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_TI4.TabIndex = 16;
			this.cmbFFE_TI4.Text = "<";
			this.cmbFFE_TI4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_TI4_KeyPress);
			// 
			// chkFFE_TI4
			// 
			this.chkFFE_TI4.Enabled = false;
			this.chkFFE_TI4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_TI4.Location = new System.Drawing.Point(82, 162);
			this.chkFFE_TI4.Name = "chkFFE_TI4";
			this.chkFFE_TI4.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_TI4.TabIndex = 15;
			this.chkFFE_TI4.Text = "4";
			this.chkFFE_TI4.Click += new System.EventHandler(this.chkFFE_TI4_Click);
			// 
			// label5
			// 
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(227, 138);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(32, 16);
			this.label5.TabIndex = 14;
			this.label5.Text = "MPH";
			// 
			// txtFFE_TI3
			// 
			this.txtFFE_TI3.Enabled = false;
			this.txtFFE_TI3.Location = new System.Drawing.Point(178, 138);
			this.txtFFE_TI3.MaxLength = 3;
			this.txtFFE_TI3.Name = "txtFFE_TI3";
			this.txtFFE_TI3.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_TI3.TabIndex = 13;
			this.txtFFE_TI3.Tag = "999";
			this.txtFFE_TI3.Text = "";
			this.txtFFE_TI3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_TI3_KeyPress);
			// 
			// cmbFFE_TI3
			// 
			this.cmbFFE_TI3.Enabled = false;
			this.cmbFFE_TI3.Items.AddRange(new object[] {
															"<",
															"<=",
															">",
															">="});
			this.cmbFFE_TI3.Location = new System.Drawing.Point(121, 138);
			this.cmbFFE_TI3.Name = "cmbFFE_TI3";
			this.cmbFFE_TI3.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_TI3.TabIndex = 12;
			this.cmbFFE_TI3.Text = "<";
			this.cmbFFE_TI3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_TI3_KeyPress);
			// 
			// chkFFE_TI3
			// 
			this.chkFFE_TI3.Enabled = false;
			this.chkFFE_TI3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_TI3.Location = new System.Drawing.Point(82, 138);
			this.chkFFE_TI3.Name = "chkFFE_TI3";
			this.chkFFE_TI3.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_TI3.TabIndex = 11;
			this.chkFFE_TI3.Text = "3";
			this.chkFFE_TI3.Click += new System.EventHandler(this.chkFFE_TI3_Click);
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(227, 118);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(32, 16);
			this.label4.TabIndex = 10;
			this.label4.Text = "MPH";
			// 
			// txtFFE_TI2
			// 
			this.txtFFE_TI2.Enabled = false;
			this.txtFFE_TI2.Location = new System.Drawing.Point(178, 113);
			this.txtFFE_TI2.MaxLength = 3;
			this.txtFFE_TI2.Name = "txtFFE_TI2";
			this.txtFFE_TI2.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_TI2.TabIndex = 9;
			this.txtFFE_TI2.Tag = "999";
			this.txtFFE_TI2.Text = "";
			this.txtFFE_TI2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_TI2_KeyPress);
			// 
			// cmbFFE_TI2
			// 
			this.cmbFFE_TI2.Enabled = false;
			this.cmbFFE_TI2.Items.AddRange(new object[] {
															"<",
															"<=",
															">",
															">="});
			this.cmbFFE_TI2.Location = new System.Drawing.Point(121, 113);
			this.cmbFFE_TI2.Name = "cmbFFE_TI2";
			this.cmbFFE_TI2.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_TI2.TabIndex = 8;
			this.cmbFFE_TI2.Text = "<";
			this.cmbFFE_TI2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_TI2_KeyPress);
			this.cmbFFE_TI2.SelectedIndexChanged += new System.EventHandler(this.cmbFFE_TI2_SelectedIndexChanged);
			// 
			// chkFFE_TI2
			// 
			this.chkFFE_TI2.Enabled = false;
			this.chkFFE_TI2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_TI2.Location = new System.Drawing.Point(82, 113);
			this.chkFFE_TI2.Name = "chkFFE_TI2";
			this.chkFFE_TI2.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_TI2.TabIndex = 7;
			this.chkFFE_TI2.Text = "2";
			this.chkFFE_TI2.Click += new System.EventHandler(this.chkFFE_TI2_Click);
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Black;
			this.label3.Location = new System.Drawing.Point(16, 16);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(552, 32);
			this.label3.TabIndex = 6;
			this.label3.Text = "Classify wind speeds (MPH) for sustainable fire potential. Up To 5 Classification" +
				"s Can Be Entered with 1 being the highest risk for sustained fire potential";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(227, 94);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "MPH";
			// 
			// btnFFEDefaultStep1
			// 
			this.btnFFEDefaultStep1.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEDefaultStep1.Location = new System.Drawing.Point(8, 256);
			this.btnFFEDefaultStep1.Name = "btnFFEDefaultStep1";
			this.btnFFEDefaultStep1.Size = new System.Drawing.Size(128, 24);
			this.btnFFEDefaultStep1.TabIndex = 4;
			this.btnFFEDefaultStep1.Text = "Use Default Values";
			this.btnFFEDefaultStep1.Click += new System.EventHandler(this.btnWindSpeedDefault_Click);
			// 
			// txtFFE_TI1
			// 
			this.txtFFE_TI1.Enabled = false;
			this.txtFFE_TI1.Location = new System.Drawing.Point(178, 88);
			this.txtFFE_TI1.MaxLength = 3;
			this.txtFFE_TI1.Name = "txtFFE_TI1";
			this.txtFFE_TI1.Size = new System.Drawing.Size(40, 20);
			this.txtFFE_TI1.TabIndex = 3;
			this.txtFFE_TI1.Tag = "999";
			this.txtFFE_TI1.Text = "";
			this.txtFFE_TI1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFFE_TI1_KeyPress);
			this.txtFFE_TI1.Leave += new System.EventHandler(this.txtFFE_TI1_Leave);
			// 
			// cmbFFE_TI1
			// 
			this.cmbFFE_TI1.Enabled = false;
			this.cmbFFE_TI1.Items.AddRange(new object[] {
															"<",
															"<="});
			this.cmbFFE_TI1.Location = new System.Drawing.Point(121, 88);
			this.cmbFFE_TI1.Name = "cmbFFE_TI1";
			this.cmbFFE_TI1.Size = new System.Drawing.Size(39, 21);
			this.cmbFFE_TI1.TabIndex = 2;
			this.cmbFFE_TI1.Text = "<";
			this.cmbFFE_TI1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFE_TI1_KeyPress);
			this.cmbFFE_TI1.SelectedIndexChanged += new System.EventHandler(this.cmbFFE_TI1_SelectedIndexChanged);
			// 
			// chkFFE_TI1
			// 
			this.chkFFE_TI1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFE_TI1.Location = new System.Drawing.Point(83, 86);
			this.chkFFE_TI1.Name = "chkFFE_TI1";
			this.chkFFE_TI1.Size = new System.Drawing.Size(32, 24);
			this.chkFFE_TI1.TabIndex = 1;
			this.chkFFE_TI1.Text = "1";
			this.chkFFE_TI1.Click += new System.EventHandler(this.chkFFE_TI1_Click);
			// 
			// grpboxFFEHazard
			// 
			this.grpboxFFEHazard.BackColor = System.Drawing.Color.White;
			this.grpboxFFEHazard.Controls.Add(this.label24);
			this.grpboxFFEHazard.Controls.Add(this.label20);
			this.grpboxFFEHazard.Controls.Add(this.cmbFFECIHazardWindSpeedClass);
			this.grpboxFFEHazard.Controls.Add(this.chkFFECIHazard);
			this.grpboxFFEHazard.Controls.Add(this.cmbFFECIHazardOperator);
			this.grpboxFFEHazard.Controls.Add(this.cmbFFETIHazardWindSpeedClass);
			this.grpboxFFEHazard.Controls.Add(this.chkFFETIHazard);
			this.grpboxFFEHazard.Controls.Add(this.cmbFFETIHazardOperator);
			this.grpboxFFEHazard.Controls.Add(this.label29);
			this.grpboxFFEHazard.Controls.Add(this.label30);
			this.grpboxFFEHazard.Controls.Add(this.label31);
			this.grpboxFFEHazard.Controls.Add(this.btnFFEHazardDefault);
			this.grpboxFFEHazard.ForeColor = System.Drawing.Color.Black;
			this.grpboxFFEHazard.Location = new System.Drawing.Point(112, 1248);
			this.grpboxFFEHazard.Name = "grpboxFFEHazard";
			this.grpboxFFEHazard.Size = new System.Drawing.Size(584, 288);
			this.grpboxFFEHazard.TabIndex = 4;
			this.grpboxFFEHazard.TabStop = false;
			this.grpboxFFEHazard.Text = "Identify Hazardous Conditions - Step 4";
			this.grpboxFFEHazard.Visible = false;
			// 
			// label24
			// 
			this.label24.Location = new System.Drawing.Point(336, 88);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(112, 16);
			this.label24.TabIndex = 59;
			this.label24.Text = "Wind Speed Class";
			// 
			// label20
			// 
			this.label20.Location = new System.Drawing.Point(336, 180);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(112, 16);
			this.label20.TabIndex = 57;
			this.label20.Text = "Wind Speed Class";
			// 
			// cmbFFECIHazardWindSpeedClass
			// 
			this.cmbFFECIHazardWindSpeedClass.Enabled = false;
			this.cmbFFECIHazardWindSpeedClass.Location = new System.Drawing.Point(354, 203);
			this.cmbFFECIHazardWindSpeedClass.Name = "cmbFFECIHazardWindSpeedClass";
			this.cmbFFECIHazardWindSpeedClass.Size = new System.Drawing.Size(40, 21);
			this.cmbFFECIHazardWindSpeedClass.TabIndex = 56;
			this.cmbFFECIHazardWindSpeedClass.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFECIHazardWindSpeedClass_KeyPress);
			this.cmbFFECIHazardWindSpeedClass.SelectedValueChanged += new System.EventHandler(this.cmbFFECIHazardWindSpeedClass_SelectedValueChanged);
			// 
			// chkFFECIHazard
			// 
			this.chkFFECIHazard.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFECIHazard.Location = new System.Drawing.Point(160, 205);
			this.chkFFECIHazard.Name = "chkFFECIHazard";
			this.chkFFECIHazard.Size = new System.Drawing.Size(88, 16);
			this.chkFFECIHazard.TabIndex = 55;
			this.chkFFECIHazard.Text = "PRE_CI_CL";
			this.chkFFECIHazard.Click += new System.EventHandler(this.chkFFECIHazard_Click);
			// 
			// cmbFFECIHazardOperator
			// 
			this.cmbFFECIHazardOperator.Enabled = false;
			this.cmbFFECIHazardOperator.Items.AddRange(new object[] {
																		"<",
																		"<=",
																		"="});
			this.cmbFFECIHazardOperator.Location = new System.Drawing.Point(280, 203);
			this.cmbFFECIHazardOperator.Name = "cmbFFECIHazardOperator";
			this.cmbFFECIHazardOperator.Size = new System.Drawing.Size(39, 21);
			this.cmbFFECIHazardOperator.TabIndex = 54;
			this.cmbFFECIHazardOperator.Text = "<";
			this.cmbFFECIHazardOperator.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFECIHazardOperator_KeyPress);
			this.cmbFFECIHazardOperator.SelectedValueChanged += new System.EventHandler(this.cmbFFECIHazardOperator_SelectedValueChanged);
			// 
			// cmbFFETIHazardWindSpeedClass
			// 
			this.cmbFFETIHazardWindSpeedClass.Enabled = false;
			this.cmbFFETIHazardWindSpeedClass.Location = new System.Drawing.Point(352, 112);
			this.cmbFFETIHazardWindSpeedClass.Name = "cmbFFETIHazardWindSpeedClass";
			this.cmbFFETIHazardWindSpeedClass.Size = new System.Drawing.Size(40, 21);
			this.cmbFFETIHazardWindSpeedClass.TabIndex = 53;
			this.cmbFFETIHazardWindSpeedClass.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFETIHazardWindSpeedClass_KeyPress);
			this.cmbFFETIHazardWindSpeedClass.SelectedValueChanged += new System.EventHandler(this.cmbFFETIHazardWindSpeedClass_SelectedValueChanged);
			// 
			// chkFFETIHazard
			// 
			this.chkFFETIHazard.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.chkFFETIHazard.Location = new System.Drawing.Point(160, 115);
			this.chkFFETIHazard.Name = "chkFFETIHazard";
			this.chkFFETIHazard.Size = new System.Drawing.Size(88, 16);
			this.chkFFETIHazard.TabIndex = 52;
			this.chkFFETIHazard.Text = "PRE_TI_CL";
			this.chkFFETIHazard.Click += new System.EventHandler(this.chkFFETIHazard_Click);
			// 
			// cmbFFETIHazardOperator
			// 
			this.cmbFFETIHazardOperator.Enabled = false;
			this.cmbFFETIHazardOperator.Items.AddRange(new object[] {
																		"<",
																		"<=",
																		"="});
			this.cmbFFETIHazardOperator.Location = new System.Drawing.Point(280, 112);
			this.cmbFFETIHazardOperator.Name = "cmbFFETIHazardOperator";
			this.cmbFFETIHazardOperator.Size = new System.Drawing.Size(39, 21);
			this.cmbFFETIHazardOperator.TabIndex = 50;
			this.cmbFFETIHazardOperator.Text = "<";
			this.cmbFFETIHazardOperator.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFFETIHazardOperator_KeyPress);
			this.cmbFFETIHazardOperator.SelectedValueChanged += new System.EventHandler(this.cmbFFETIHazardOperator_SelectedValueChanged);
			// 
			// label29
			// 
			this.label29.BackColor = System.Drawing.Color.White;
			this.label29.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
				| System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label29.ForeColor = System.Drawing.Color.Black;
			this.label29.Location = new System.Drawing.Point(256, 152);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(88, 16);
			this.label29.TabIndex = 25;
			this.label29.Text = "Crown Index";
			// 
			// label30
			// 
			this.label30.BackColor = System.Drawing.Color.White;
			this.label30.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
				| System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label30.ForeColor = System.Drawing.Color.Black;
			this.label30.Location = new System.Drawing.Point(256, 69);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(80, 16);
			this.label30.TabIndex = 24;
			this.label30.Text = "Torch Index";
			// 
			// label31
			// 
			this.label31.BackColor = System.Drawing.Color.White;
			this.label31.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label31.ForeColor = System.Drawing.Color.Black;
			this.label31.Location = new System.Drawing.Point(16, 16);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(552, 48);
			this.label31.TabIndex = 6;
			this.label31.Text = "Wind speed classifications can help in identifying pre-treatment hazardous sustai" +
				"nable fire conditions.  Define hazardous torch and crown index wind speed class " +
				"expression(s).";
			// 
			// btnFFEHazardDefault
			// 
			this.btnFFEHazardDefault.BackColor = System.Drawing.SystemColors.Control;
			this.btnFFEHazardDefault.Location = new System.Drawing.Point(8, 256);
			this.btnFFEHazardDefault.Name = "btnFFEHazardDefault";
			this.btnFFEHazardDefault.Size = new System.Drawing.Size(128, 24);
			this.btnFFEHazardDefault.TabIndex = 4;
			this.btnFFEHazardDefault.Text = "Use Default Values";
			this.btnFFEHazardDefault.Click += new System.EventHandler(this.btnFFEHazardDefault_Click);
			// 
			// uc_scenario_ffe
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_ffe";
			this.Size = new System.Drawing.Size(744, 1696);
			this.groupBox1.ResumeLayout(false);
			this.grpboxFFEBackslide.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.trbFFECIBackslide)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.trbFFETIBackslide)).EndInit();
			this.grpboxFFEExpressionBuilder.ResumeLayout(false);
			this.grpboxFFE_TI_CI_Effective.ResumeLayout(false);
			this.grpboxFFEClassifications.ResumeLayout(false);
			this.grpboxFFEHazard.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void grpboxFFEIndices_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void grpboxFFE_Resize(object sender, System.EventArgs e)
		{
			try
			{
			}
			catch
			{
			}
		}


		private void cmbFFE_TI1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbFFE_TI2_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkFFE_TI1_Click(object sender, System.EventArgs e)
		{
            this.val_checkbox("T1");
		}
		public void val_checkbox(string strChkBox)
		{
			switch (strChkBox)
			{
				case ("T1"):
					if (this.chkFFE_TI2.Checked==true) 
					{
						this.chkFFE_TI1.Checked=true;
					}
					else 
					{
						if (this.chkFFE_TI1.Checked == false)
						{
							this.cmbFFE_TI1.Enabled=false;
							this.txtFFE_TI1.Enabled=false;
							this.chkFFE_TI2.Enabled=false;

						}
						else 
						{
							this.chkFFE_TI2.Enabled=true;
							this.cmbFFE_TI1.Enabled=true;
							this.txtFFE_TI1.Enabled=true;
						}

						
					}
					break;
				case ("T2"):
					if (this.chkFFE_TI3.Checked==true) 
					{
						this.chkFFE_TI2.Checked=true;
					}
					else 
					{
						if (this.chkFFE_TI2.Checked == false)
						{
							this.cmbFFE_TI2.Enabled=false;
							this.txtFFE_TI2.Enabled=false;
							this.cmbFFE_TI3.Enabled=false; 
							this.chkFFE_TI3.Enabled=false;
                         
						}
						else 
						{
							this.chkFFE_TI3.Enabled=true; 
							this.cmbFFE_TI2.Enabled=true;
							this.txtFFE_TI2.Enabled=true;
						}
					}
					break;
				case ("T3"):
					if (this.chkFFE_TI4.Checked==true) 
					{
						this.chkFFE_TI3.Checked=true;
					}
					else 
					{
						if (this.chkFFE_TI3.Checked == false)
						{
							this.cmbFFE_TI3.Enabled=false;
							this.txtFFE_TI3.Enabled=false;
							this.cmbFFE_TI4.Enabled=false; 
							this.chkFFE_TI4.Enabled=false;
                         
						}
						else 
						{
							this.chkFFE_TI4.Enabled=true; 
							this.cmbFFE_TI3.Enabled=true;
							this.txtFFE_TI3.Enabled=true;
						}
					}

				    break;
				case ("T4"):
					if (this.chkFFE_TI5.Checked==true) 
					{
						this.chkFFE_TI4.Checked=true;
					}
					else 
					{
						if (this.chkFFE_TI4.Checked == false)
						{
							this.cmbFFE_TI4.Enabled=false;
							this.txtFFE_TI4.Enabled=false;
							this.cmbFFE_TI5.Enabled=false; 
							this.chkFFE_TI5.Enabled=false;
                         
						}
						else 
						{
							this.chkFFE_TI5.Enabled=true; 
							this.cmbFFE_TI4.Enabled=true;
							this.txtFFE_TI4.Enabled=true;
						}
					}

				    break;
				case ("T5"):
					if (this.chkFFE_TI5.Checked==false)
					{
						this.cmbFFE_TI5.Enabled=false;
						this.txtFFE_TI5.Enabled=false;
					}
					else 
					{
						this.cmbFFE_TI5.Enabled=true;
						this.txtFFE_TI5.Enabled=true;

					}
				    break;
				case ("C1"):
					if (this.chkFFE_CI2.Checked==true) 
					{
						this.chkFFE_CI1.Checked=true;
					}
					else 
					{
						if (this.chkFFE_CI1.Checked == false)
						{
							this.cmbFFE_CI1.Enabled=false;
							this.txtFFE_CI1.Enabled=false;
							this.chkFFE_CI2.Enabled=false;

						}
						else 
						{
							this.chkFFE_CI2.Enabled=true;
							this.cmbFFE_CI1.Enabled=true;
							this.txtFFE_CI1.Enabled=true;
						}

						
					}
					break;
				case ("C2"):
					if (this.chkFFE_CI3.Checked==true) 
					{
						this.chkFFE_CI2.Checked=true;
					}
					else 
					{
						if (this.chkFFE_CI2.Checked == false)
						{
							this.cmbFFE_CI2.Enabled=false;
							this.txtFFE_CI2.Enabled=false;
							this.cmbFFE_CI3.Enabled=false; 
							this.chkFFE_CI3.Enabled=false;
                         
						}
						else 
						{
							this.chkFFE_CI3.Enabled=true; 
							this.cmbFFE_CI2.Enabled=true;
							this.txtFFE_CI2.Enabled=true;
						}
					}
					break;
				case ("C3"):
					if (this.chkFFE_CI4.Checked==true) 
					{
						this.chkFFE_CI3.Checked=true;
					}
					else 
					{
						if (this.chkFFE_CI3.Checked == false)
						{
							this.cmbFFE_CI3.Enabled=false;
							this.txtFFE_CI3.Enabled=false;
							this.cmbFFE_CI4.Enabled=false; 
							this.chkFFE_CI4.Enabled=false;
                         
						}
						else 
						{
							this.chkFFE_CI4.Enabled=true; 
							this.cmbFFE_CI3.Enabled=true;
							this.txtFFE_CI3.Enabled=true;
						}
					}

					break;
				case ("C4"):
					if (this.chkFFE_CI5.Checked==true) 
					{
						this.chkFFE_CI4.Checked=true;
					}
					else 
					{
						if (this.chkFFE_CI4.Checked == false)
						{
							this.cmbFFE_CI4.Enabled=false;
							this.txtFFE_CI4.Enabled=false;
							this.cmbFFE_CI5.Enabled=false; 
							this.chkFFE_CI5.Enabled=false;
                         
						}
						else 
						{
							this.chkFFE_CI5.Enabled=true; 
							this.cmbFFE_CI4.Enabled=true;
							this.txtFFE_CI4.Enabled=true;
						}
					}

					break;
				case ("C5"):
					if (this.chkFFE_CI5.Checked==false)
					{
						this.cmbFFE_CI5.Enabled=false;
						this.txtFFE_CI5.Enabled=false;
					}
					else 
					{
						this.cmbFFE_CI5.Enabled=true;
						this.txtFFE_CI5.Enabled=true;

					}
					break;

			}
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void chkFFE_TI2_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("T2");
		}

		private void chkFFE_TI3_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("T3");
		}

		private void chkFFE_TI4_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("T4");
		}

		private void chkFFE_TI5_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("T5");
		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.grpboxFFEClassifications.Left = this.groupBox1.Width / 2 - 
					(int)(this.grpboxFFEClassifications.Width * .50);

				this.grpboxFFEClassifications.Top = (int)(this.groupBox1.Height * .5) - (int)(this.grpboxFFEClassifications.Height * .5);
			}
			catch
			{
			}
		}

		private void label6_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnWindSpeedDefault_Click(object sender, System.EventArgs e)
		{
			this.chkFFE_TI1.Checked=true;
			this.cmbFFE_TI1.Enabled=true;
			this.txtFFE_TI1.Enabled=true;
			this.cmbFFE_TI1.Text = "<"; 
			this.txtFFE_TI1.Text = "25";
              
			this.chkFFE_TI2.Enabled=true;
			this.chkFFE_TI2.Checked=true;
			this.cmbFFE_TI2.Enabled=true;
			this.txtFFE_TI2.Enabled=true;
			this.cmbFFE_TI2.Text = "<"; 
			this.txtFFE_TI2.Text = "50";

			this.chkFFE_TI3.Enabled=true;
			this.chkFFE_TI3.Checked=true;
			this.cmbFFE_TI3.Enabled=true;
			this.txtFFE_TI3.Enabled=true;
			this.cmbFFE_TI3.Text = ">="; 
			this.txtFFE_TI3.Text = "50";

			this.chkFFE_TI4.Enabled=true;
		    this.chkFFE_TI4.Checked=false;
			this.cmbFFE_TI4.Enabled=false;
			this.txtFFE_TI4.Enabled=false;
		

			this.chkFFE_TI5.Checked=false;
			this.chkFFE_TI5.Enabled=false;
			this.cmbFFE_TI5.Enabled=false;
			this.txtFFE_TI5.Enabled=false;

			this.chkFFE_CI1.Checked=true;
			this.cmbFFE_CI1.Enabled=true;
			this.txtFFE_CI1.Enabled=true;
			this.cmbFFE_CI1.Text = "<"; 
			this.txtFFE_CI1.Text = "25";
              
			this.chkFFE_CI2.Enabled=true;
			this.chkFFE_CI2.Checked=true;
			this.cmbFFE_CI2.Enabled=true;
			this.txtFFE_CI2.Enabled=true;
			this.cmbFFE_CI2.Text = "<"; 
			this.txtFFE_CI2.Text = "50";

			this.chkFFE_CI3.Enabled=true;
			this.chkFFE_CI3.Checked=true;
			this.cmbFFE_CI3.Enabled=true;
			this.txtFFE_CI3.Enabled=true;
			this.cmbFFE_CI3.Text = ">="; 
			this.txtFFE_CI3.Text = "50";

			this.chkFFE_CI4.Checked=false;
			this.chkFFE_CI4.Enabled=true;
			this.cmbFFE_CI4.Enabled=false;
			this.txtFFE_CI4.Enabled=false;
		

			this.chkFFE_CI5.Checked=false;
			this.chkFFE_CI5.Enabled=false;
			this.cmbFFE_CI5.Enabled=false;
			this.txtFFE_CI5.Enabled=false;

			//((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
			


		}

		private void txtFFE_TI1_Leave(object sender, System.EventArgs e)
		{
		   
		}

		private void txtFFE_TI1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);

		}
		protected void AllowNumericOnly(System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				// Digits are OK
				//if (((frmScenario)this.ParentForm).btnSave.Enabled == false) 
				//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
				((frmCoreScenario)this.ParentForm).m_bSave=true;
			}
			else if (e.KeyChar == '\b') 
			{
				//back space is okay
				//if (((frmScenario)this.ParentForm).btnSave.Enabled == false) 
				//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
				((frmCoreScenario)this.ParentForm).m_bSave=true;
			}
			else
			{
				e.Handled = true;
			}
		}

		private void txtFFE_TI2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_TI3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_TI4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_TI5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}
		public void loadvalues()
		{
            
			string strSQL="";
			//int intArrayCount;
			//int x=0;
			string strConn="";
			//string strCommand="";
			string str="";
			

			ado_data_access p_ado = new ado_data_access();

			string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            
			//scenario mdb connection
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strScenarioMDB + ";User Id=admin;Password=;";
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				p_ado = null;
				return;
			}
			strSQL = "SELECT * FROM scenario_ffe_wind_speed WHERE " + 
				         " scenario_id = '" + strScenarioId + "';";
			p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);

				  
			//insert records into the scenario_rx_intensity table from the master rx table
			if (p_ado.m_intError == 0)
			{
				
                //load  wind speed class definitions
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["wind_speed_class"] != System.DBNull.Value)
					{
						str = p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim() + 
							  p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
						switch (str)
						{
							case "T1" :
								this.chkFFE_TI1.Enabled=true;
								this.chkFFE_TI1.Checked=true;
								this.cmbFFE_TI1.Enabled=true;
								this.cmbFFE_TI1.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_TI1.Enabled=true;
								this.txtFFE_TI1.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								this.chkFFE_TI2.Enabled=true;
								break;
							case "T2" :
								this.chkFFE_TI2.Enabled=true;
								this.chkFFE_TI2.Checked=true;
								this.cmbFFE_TI2.Enabled=true;
								this.cmbFFE_TI2.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_TI2.Enabled=true;
								this.txtFFE_TI2.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								this.chkFFE_TI3.Enabled=true;
								break;
							case "T3" :
								this.chkFFE_TI3.Enabled=true;
								this.chkFFE_TI3.Checked=true;
								this.cmbFFE_TI3.Enabled=true;
								this.cmbFFE_TI3.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_TI3.Enabled=true;
								this.txtFFE_TI3.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								this.chkFFE_TI4.Enabled=true;
								break;
							case "T4" :
								this.chkFFE_TI4.Enabled=true;
								this.chkFFE_TI4.Checked=true;
								this.cmbFFE_TI4.Enabled=true;
								this.cmbFFE_TI4.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_TI4.Enabled=true;
								this.txtFFE_TI4.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								this.chkFFE_TI5.Enabled=true;
								break;
							case "T5" :
								this.chkFFE_TI5.Enabled=true;
								this.chkFFE_TI5.Checked=true;
								this.cmbFFE_TI5.Enabled=true;
								this.cmbFFE_TI5.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_TI5.Enabled=true;
								this.txtFFE_TI5.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								break;
							case "C1" :
								this.chkFFE_CI1.Enabled=true;
								this.chkFFE_CI1.Checked=true;
								this.cmbFFE_CI1.Enabled=true;
								this.cmbFFE_CI1.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_CI1.Enabled=true;
								this.txtFFE_CI1.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								this.chkFFE_CI2.Enabled=true;
								break;
							case "C2" :
								this.chkFFE_CI2.Enabled=true;
								this.chkFFE_CI2.Checked=true;
								this.cmbFFE_CI2.Enabled=true;
								this.cmbFFE_CI2.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_CI2.Enabled=true;
								this.txtFFE_CI2.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								this.chkFFE_CI3.Enabled=true;
								break;
							case "C3" :
								this.chkFFE_CI3.Enabled=true;
								this.chkFFE_CI3.Checked=true;
								this.cmbFFE_CI3.Enabled=true;
								this.cmbFFE_CI3.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_CI3.Enabled=true;
								this.txtFFE_CI3.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								this.chkFFE_CI4.Enabled=true;
								break;
							case "C4" :
								this.chkFFE_CI4.Enabled=true;
								this.chkFFE_CI4.Checked=true;
								this.cmbFFE_CI4.Enabled=true;
								this.cmbFFE_CI4.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_CI4.Enabled=true;
								this.txtFFE_CI4.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								this.chkFFE_CI5.Enabled=true;
								break;
							case "C5" :
								this.chkFFE_CI5.Enabled=true;
								this.chkFFE_CI5.Checked=true;
								this.cmbFFE_CI5.Enabled=true;
								this.cmbFFE_CI5.Text = p_ado.m_OleDbDataReader["expression_operator"].ToString();
								this.txtFFE_CI5.Enabled=true;
								this.txtFFE_CI5.Text = p_ado.m_OleDbDataReader["speed_mph"].ToString();
								break;
						}
					}

				}
				
				p_ado.m_OleDbDataReader.Close();


				//load list with torch and crown index effective change expressions
				strSQL = "SELECT * FROM scenario_ffe_ti_ci_effective_change WHERE " + 
					" scenario_id = '" + strScenarioId + "';";
				p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
				if (p_ado.m_intError==0)
				{   

					this.listView1.Clear();
					this.listView1.Columns.Add("Index Type",100,HorizontalAlignment.Left);
					this.listView1.Columns.Add("Wind Speed Class", 100, HorizontalAlignment.Center);
					this.listView1.Columns.Add("Expression", 500, HorizontalAlignment.Left);
					while (p_ado.m_OleDbDataReader.Read())
					{
						if (p_ado.m_OleDbDataReader["ti_ci_index_type"] != System.DBNull.Value)
						{
							if (p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim().Length > 0)
							{
								if (p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim()=="T")
								{
									this.listView1.Items.Add("Torch");
								}
								else
								{
									this.listView1.Items.Add("Crown");
								}
								this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim());
								this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(p_ado.m_OleDbDataReader["expression"].ToString().Trim());

							}
						}
					
					}
					p_ado.m_OleDbDataReader.Close();

					//load backslide values
					strSQL = "SELECT * FROM scenario_ffe_backslide WHERE " + 
						" scenario_id = '" + strScenarioId + "';";
					p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
					if (p_ado.m_intError==0)
					{   
						while (p_ado.m_OleDbDataReader.Read())
						{
							if (p_ado.m_OleDbDataReader["ti_ci_index_type"] != System.DBNull.Value)
							{
								if (p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim().Length > 0)
								{
									if (p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim()=="T")
									{
										this.cmbFFETIBackslide.Text = p_ado.m_OleDbDataReader["expression_operator1"].ToString();

										this.trbFFETIBackslide.Value = Convert.ToInt32(p_ado.m_OleDbDataReader["backslide"]) * -1;
										if (p_ado.m_OleDbDataReader["expression_operator2"] != System.DBNull.Value)
										{
											if (p_ado.m_OleDbDataReader["expression_operator2"].ToString().Trim().Length > 0)
											{
												this.chkFFETIBackslide.Enabled=true;
												this.cmbFFETIBackslide3.Enabled=true;
												this.cmbFFETIBackslide2.Enabled=true;
												this.chkFFETIBackslide.Checked=true;
												this.cmbFFETIBackslide3.Text = 
													p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
												this.cmbFFETIBackslide2.Text = 
													p_ado.m_OleDbDataReader["expression_operator2"].ToString().Trim();
											}
											else
											{
												this.chkFFETIBackslide.Enabled=true;
												this.cmbFFETIBackslide3.Enabled=false;
												this.cmbFFETIBackslide2.Enabled=false;
												this.chkFFETIBackslide.Checked=false;
											}
										}
										else
										{
											this.chkFFETIBackslide.Enabled=true;
											this.cmbFFETIBackslide3.Enabled=false;
											this.cmbFFETIBackslide2.Enabled=false;
											this.chkFFETIBackslide.Checked=false;
										}
									}
									else
									{
										if (p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim()=="C")
										{
											this.cmbFFECIBackslide.Text = p_ado.m_OleDbDataReader["expression_operator1"].ToString();

											this.trbFFECIBackslide.Value = Convert.ToInt32(p_ado.m_OleDbDataReader["backslide"]) * -1;
											if (p_ado.m_OleDbDataReader["expression_operator2"] != System.DBNull.Value)
											{
												if (p_ado.m_OleDbDataReader["expression_operator2"].ToString().Trim().Length > 0)
												{
													this.chkFFECIBackslide.Enabled=true;
													this.cmbFFECIBackslide3.Enabled=true;
													this.cmbFFECIBackslide2.Enabled=true;
													this.chkFFECIBackslide.Checked=true;
													this.cmbFFECIBackslide3.Text = 
														p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
													this.cmbFFECIBackslide2.Text = 
														p_ado.m_OleDbDataReader["expression_operator2"].ToString().Trim();
												}
												else
												{
													this.chkFFECIBackslide.Enabled=true;
													this.cmbFFECIBackslide3.Enabled=false;
													this.cmbFFECIBackslide2.Enabled=false;
													this.chkFFECIBackslide.Checked=false;
												}
											}
											else
											{
												this.chkFFECIBackslide.Enabled=true;
												this.cmbFFECIBackslide3.Enabled=false;
												this.cmbFFECIBackslide2.Enabled=false;
												this.chkFFECIBackslide.Checked=false;
											}
										}

									}
								}
							}   
						}
						p_ado.m_OleDbDataReader.Close();
						if (this.chkFFE_TI1.Checked==true) this.cmbFFETIBackslide3.Items.Add("1");
                        if (this.chkFFE_TI2.Checked==true) this.cmbFFETIBackslide3.Items.Add("2");
                        if (this.chkFFE_TI3.Checked==true) this.cmbFFETIBackslide3.Items.Add("3");
						if (this.chkFFE_TI4.Checked==true) this.cmbFFETIBackslide3.Items.Add("4");
						if (this.chkFFE_TI5.Checked==true) this.cmbFFETIBackslide3.Items.Add("5");
						if (this.chkFFE_CI1.Checked==true) this.cmbFFECIBackslide3.Items.Add("1");
						if (this.chkFFE_CI2.Checked==true) this.cmbFFECIBackslide3.Items.Add("2");
						if (this.chkFFE_CI3.Checked==true) this.cmbFFECIBackslide3.Items.Add("3");
						if (this.chkFFE_CI4.Checked==true) this.cmbFFECIBackslide3.Items.Add("4");
						if (this.chkFFE_CI5.Checked==true) this.cmbFFECIBackslide3.Items.Add("5");


						//load hazard values
						strSQL = "SELECT * FROM scenario_ffe_hazard WHERE " + 
							" scenario_id = '" + strScenarioId + "';";
						p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
						if (p_ado.m_intError==0)
						{   
							while (p_ado.m_OleDbDataReader.Read())
							{
								if (p_ado.m_OleDbDataReader["ti_ci_index_type"] != System.DBNull.Value)
								{
									if (p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim().Length > 0)
									{
										if (p_ado.m_OleDbDataReader["ti_ci_index_type"].ToString().Trim()=="T")
										{
											
											if (p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim().Length > 0)
											{
												this.chkFFETIHazard.Enabled=true;
												this.cmbFFETIHazardOperator.Enabled=true;
												this.cmbFFETIHazardWindSpeedClass.Enabled=true;
												this.chkFFETIHazard.Checked=true;
												this.cmbFFETIHazardWindSpeedClass.Text = 
													p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
												this.cmbFFETIHazardOperator.Text = 
													p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim();
											}
											else
											{
												this.chkFFETIHazard.Enabled=true;
												this.cmbFFETIHazardOperator.Enabled=false;
												this.cmbFFETIHazardWindSpeedClass.Enabled=false;
												this.chkFFETIHazard.Checked=false;
											}
											
										}
										else
										{
											if (p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim().Length > 0)
											{
												this.chkFFECIHazard.Enabled=true;
												this.cmbFFECIHazardOperator.Enabled=true;
												this.cmbFFECIHazardWindSpeedClass.Enabled=true;
												this.chkFFECIHazard.Checked=true;
												this.cmbFFECIHazardWindSpeedClass.Text = 
													p_ado.m_OleDbDataReader["wind_speed_class"].ToString().Trim();
												this.cmbFFECIHazardOperator.Text = 
													p_ado.m_OleDbDataReader["expression_operator"].ToString().Trim();
											}
											else
											{
												this.chkFFECIHazard.Enabled=true;
												this.cmbFFECIHazardOperator.Enabled=false;
												this.cmbFFECIHazardWindSpeedClass.Enabled=false;
												this.chkFFECIHazard.Checked=false;
											}
										}
									}
								}   
							}
							p_ado.m_OleDbDataReader.Close();
							if (this.chkFFE_TI1.Checked==true) this.cmbFFETIHazardWindSpeedClass.Items.Add("1");
							if (this.chkFFE_TI2.Checked==true) this.cmbFFETIHazardWindSpeedClass.Items.Add("2");
							if (this.chkFFE_TI3.Checked==true) this.cmbFFETIHazardWindSpeedClass.Items.Add("3");
							if (this.chkFFE_TI4.Checked==true) this.cmbFFETIHazardWindSpeedClass.Items.Add("4");
							if (this.chkFFE_TI5.Checked==true) this.cmbFFETIHazardWindSpeedClass.Items.Add("5");
							if (this.chkFFE_CI1.Checked==true) this.cmbFFECIHazardWindSpeedClass.Items.Add("1");
							if (this.chkFFE_CI2.Checked==true) this.cmbFFECIHazardWindSpeedClass.Items.Add("2");
							if (this.chkFFE_CI3.Checked==true) this.cmbFFECIHazardWindSpeedClass.Items.Add("3");
							if (this.chkFFE_CI4.Checked==true) this.cmbFFECIHazardWindSpeedClass.Items.Add("4");
							if (this.chkFFE_CI5.Checked==true) this.cmbFFECIHazardWindSpeedClass.Items.Add("5");

							//load overall effectiveness values
							strSQL = "SELECT * FROM scenario_ffe_overall_effective_change WHERE " + 
								" scenario_id = '" + strScenarioId + "';";
							p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
							if (p_ado.m_intError==0)
							{   
								while (p_ado.m_OleDbDataReader.Read())
								{
									if (p_ado.m_OleDbDataReader["expression"] != System.DBNull.Value)
									{
										if (p_ado.m_OleDbDataReader["expression"].ToString().Trim().Length > 0)
										{
                                            this.m_strOverallEffectiveExpression = p_ado.m_OleDbDataReader["expression"].ToString().Trim();
										}
									}
								}
								p_ado.m_OleDbDataReader.Close();
							}
						}
					}
				}


				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;



				
			}

			
			p_ado = null;
			this.m_OleDbConnectionScenario.Close();
			//((frmScenario)this.ParentForm).btnSave.Enabled=false;
			((frmCoreScenario)this.ParentForm).m_bSave=true;

		}
		public int savevalues()
		{
            int x=0;
			string str="";
			string strSQL = "";
			string strConn = "";

			ado_data_access p_ado = new ado_data_access();

            


			string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            //string strScenarioId="scenario1";
			//scenario mdb connection
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strScenarioMDB + ";User Id=admin;Password=;";
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				x=p_ado.m_intError;
				p_ado=null;
				return x;
			}

			//delete all records from the scenario wind speed class table
			strSQL = "DELETE FROM scenario_ffe_wind_speed WHERE " + 
				" scenario_id = '" + strScenarioId + "';";

			p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			if (p_ado.m_intError < 0)
			{
				this.m_OleDbConnectionScenario.Close();
				x=p_ado.m_intError;
				p_ado = null;
				return x;
			}

            //save step 1 wind speed class values
			//save class 1
			if (this.chkFFE_TI1.Enabled && this.chkFFE_TI1.Checked == true)
			{
				if (this.txtFFE_TI1.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','T', 1,'" + this.cmbFFE_TI1.Text + "'," + 
						          Convert.ToInt32(this.txtFFE_TI1.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}
            
			//save class 2
			if (this.chkFFE_TI2.Enabled && this.chkFFE_TI2.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_TI2.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','T', 2,'" + this.cmbFFE_TI2.Text + "'," + 
						Convert.ToInt32(this.txtFFE_TI2.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}

			if (this.chkFFE_TI3.Enabled && this.chkFFE_TI3.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_TI3.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','T', 3,'" + this.cmbFFE_TI3.Text + "'," + 
						Convert.ToInt32(this.txtFFE_TI3.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}
            

			if (this.chkFFE_TI4.Enabled && this.chkFFE_TI4.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_TI4.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','T', 4,'" + this.cmbFFE_TI4.Text + "'," + 
						Convert.ToInt32(this.txtFFE_TI4.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}



			if (this.chkFFE_TI5.Enabled && this.chkFFE_TI5.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_TI5.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','T', 5,'" + this.cmbFFE_TI5.Text + "'," + 
						Convert.ToInt32(this.txtFFE_TI5.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}

			if (p_ado.m_intError < 0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}

			//save class 1
			if (this.chkFFE_CI1.Enabled && this.chkFFE_CI1.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_CI1.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','C', 1,'" + this.cmbFFE_CI1.Text + "'," + 
						Convert.ToInt32(this.txtFFE_CI1.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}
            
			//save class 2
			if (this.chkFFE_CI2.Enabled && this.chkFFE_CI2.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_CI2.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','C', 2,'" + this.cmbFFE_CI2.Text + "'," + 
						Convert.ToInt32(this.txtFFE_CI2.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}

			if (this.chkFFE_CI3.Enabled && this.chkFFE_CI3.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_CI3.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','C', 3,'" + this.cmbFFE_CI3.Text + "'," + 
						Convert.ToInt32(this.txtFFE_CI3.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}
            

			if (this.chkFFE_CI4.Enabled && this.chkFFE_CI4.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_CI4.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','C', 4,'" + this.cmbFFE_CI4.Text + "'," + 
						Convert.ToInt32(this.txtFFE_CI4.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}



			if (this.chkFFE_CI5.Enabled && this.chkFFE_CI5.Checked == true && p_ado.m_intError==0)
			{
				if (this.txtFFE_CI5.Text.Trim().Length == 0) 
				{
				}
				else 
				{
					strSQL = "INSERT INTO scenario_ffe_wind_speed (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator, speed_mph)" + 
						" VALUES ('" + strScenarioId + "','C', 5,'" + this.cmbFFE_CI5.Text + "'," + 
						Convert.ToInt32(this.txtFFE_CI5.Text.ToString()) + ");";
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				}
			}


            //delete scenario records for effective post treatment expressions
			strSQL = "DELETE FROM scenario_ffe_ti_ci_effective_change WHERE " + 
				" scenario_id = '" + strScenarioId + "';";

			p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			if (p_ado.m_intError < 0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}
            //save the effective treatment expressions from step 2
			for (x=0; x <= this.listView1.Items.Count -1; x++)
			{
				str = this.listView1.Items[x].SubItems[0].Text.Substring(0,1);
				strSQL = "INSERT INTO scenario_ffe_ti_ci_effective_change (scenario_id,ti_ci_index_type, wind_speed_class, expression)" + 
					" VALUES ('" + strScenarioId + "','" + str + "'," + 
					              this.listView1.Items[x].SubItems[1].Text.Substring(0,1) + ",'" + 
					              this.listView1.Items[x].SubItems[2].Text.Trim() + "');";
				p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				if (p_ado.m_intError < 0) break;
			}
			if (p_ado.m_intError < 0)
			{
				x=p_ado.m_intError;
				p_ado = null;
				this.m_OleDbConnectionScenario.Close();
				return x;
			}
            
			//delete scenario records from backslide table
			strSQL = "DELETE FROM scenario_ffe_backslide WHERE " + 
				" scenario_id = '" + strScenarioId + "';";
			p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			if (p_ado.m_intError < 0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}			

			//save backslide values
			if (this.chkFFETIBackslide.Checked==true && this.cmbFFETIBackslide3.Text.Trim().Length > 0)
			{
				
				strSQL = "INSERT INTO scenario_ffe_backslide (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator1, backslide,expression_operator2)" + 
					" VALUES ('" + strScenarioId + "','T'," + 
					           this.cmbFFETIBackslide3.Text + ",'" + 
					           this.cmbFFETIBackslide.Text + "'," +
					           this.lblFFETIBackslide.Text + ",'" + 
					           this.cmbFFETIBackslide2.Text + "');";

			}
			else 
			{
				strSQL = "INSERT INTO scenario_ffe_backslide (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator1, backslide,expression_operator2)" + 
					" VALUES ('" + strScenarioId + "','T'," + 
					"0,'" + 
					this.cmbFFETIBackslide.Text + "'," +
					this.lblFFETIBackslide.Text + ",' ');";

			}
			p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			if (p_ado.m_intError <  0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}		
	
			if (this.chkFFECIBackslide.Checked==true && this.cmbFFECIBackslide3.Text.Trim().Length > 0)
			{
				strSQL = "INSERT INTO scenario_ffe_backslide (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator1, backslide,expression_operator2)" + 
					" VALUES ('" + strScenarioId + "','C'," + 
					this.cmbFFECIBackslide3.Text + ",'" + 
					this.cmbFFECIBackslide.Text + "'," +
					this.lblFFECIBackslide.Text + ",'" + 
					this.cmbFFECIBackslide2.Text + "');";

			}
			else 
			{
				strSQL = "INSERT INTO scenario_ffe_backslide (scenario_id,ti_ci_index_type, wind_speed_class, expression_operator1, backslide,expression_operator2)" + 
					" VALUES ('" + strScenarioId + "','C'," + 
					"0,'" + 
					this.cmbFFECIBackslide.Text + "'," +
					this.lblFFECIBackslide.Text + ",' ');";

			}
			p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			if (p_ado.m_intError <  0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}		
			//delete scenario records from hazard table
			strSQL = "DELETE FROM scenario_ffe_hazard WHERE " + 
				" scenario_id = '" + strScenarioId + "';";
			p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			if (p_ado.m_intError < 0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}			

			//save hazard values
			if (this.chkFFETIHazard.Checked==true && this.cmbFFETIHazardWindSpeedClass.Text.Trim().Length > 0)
			{
				strSQL = "INSERT INTO scenario_ffe_hazard (scenario_id,ti_ci_index_type,  expression_operator, wind_speed_class)" + 
					" VALUES ('" + strScenarioId + "','T','" + 
					this.cmbFFETIHazardOperator.Text + "'," + 
					this.cmbFFETIHazardWindSpeedClass.Text + ");";
			    p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			}

			if (p_ado.m_intError <  0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}		
			if (this.chkFFECIHazard.Checked==true && this.cmbFFECIHazardWindSpeedClass.Text.Trim().Length > 0)
			{
				strSQL = "INSERT INTO scenario_ffe_hazard (scenario_id,ti_ci_index_type,  expression_operator, wind_speed_class)" + 
					" VALUES ('" + strScenarioId + "','C','" + 
					this.cmbFFECIHazardOperator.Text + "'," + 
					this.cmbFFECIHazardWindSpeedClass.Text + ");";
				p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			}	
			if (p_ado.m_intError < 0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}			

			//delete overall effective expression table
			strSQL = "DELETE FROM scenario_ffe_overall_effective_change WHERE " + 
				" scenario_id = '" + strScenarioId + "';";
			p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			if (p_ado.m_intError < 0)
			{
				x=p_ado.m_intError;
				this.m_OleDbConnectionScenario.Close();
				p_ado = null;
				return x;
			}		
			str = p_ado.FixString(this.m_strOverallEffectiveExpression,"'","''");
			//save overall effective expression value
			strSQL = "INSERT INTO scenario_ffe_overall_effective_change (scenario_id,expression)" + 
					" VALUES ('" + strScenarioId + "','" + str + "');";
			p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
			
			x=p_ado.m_intError;
			this.m_OleDbConnectionScenario.Close();
			p_ado = null;
			return x;


		}

		
		private void loadvalues_ti_ci_effective()
		{
			int intIndexType=0;  //torch or crown index array counter ti = 0, ci = 1
			int intClass = 0;    //wind speed class array counter 
			int intTExp = 0;     //ti expression array counter
			int intCExp = 0;     //ci expression array counter
			int x=0;
			int y=0;
			int z=0;
			

			/********************************************************************
			 **initialize an array that will hold possible unsaved expression 
			 **values resulting from a return back to this step
			 ********************************************************************/
			string[,,] strExp = new string[2,5,5];
			for (x=0; x<=1; x++)
			{
				for (y=0;y<=4;y++)
				{
					for (z=0;z<=4;z++)
					{
						strExp[x,y,z]=" ";
					}
				}
			}

			/**************************************************************************
			 **okay, now lets assign the array counter values and load the expressions
			 **************************************************************************/
			if (this.listView1.Items.Count > 0)
			{
				for (x=0;x <= this.listView1.Items.Count -1; x++)
				{
					if (this.listView1.Items[x].SubItems[0].Text.Substring(0,1)=="T") 
					{
						intIndexType = 0;    //torch index = 0
						intTExp++;           //torch index expression counter
					}
					else 
					{
						intIndexType = 1;    //crown index = 1
						intCExp++;           //crown index expression counter
					}

					/********************************************************
					 **get the wind speed class array assignment values
					 ********************************************************/
					switch  (this.listView1.Items[x].SubItems[1].Text)
					{
						case "1":
                           intClass=0;        
						   break;
						case "2":
							intClass=1;
							break;
						case "3":
							intClass=2;
							break;
						case "4":
							intClass=3;
							break;
						case "5":
							intClass=4;
							break;

					}
                    
					/****************************************************
					 **okay, now lets load the expression into the array
					 ****************************************************/
					if (intIndexType==0)    //torch index
					{
						strExp[intIndexType,intClass,intTExp-1] = this.listView1.Items[x].SubItems[2].Text;
					}
					else                    //crown index
					{
						strExp[intIndexType,intClass,intCExp-1] = this.listView1.Items[x].SubItems[2].Text;
					}
				}
			}


			this.listView1.Clear();   //clear the listview and reload it

			/**************************
			 **load the listview
			 **define column names
			 **************************/
            this.listView1.Columns.Add("Index Type",100,HorizontalAlignment.Left);
			this.listView1.Columns.Add("Wind Speed Class", 100, HorizontalAlignment.Center);
			this.listView1.Columns.Add("Expression", 500, HorizontalAlignment.Left);

			/**************************************
			 **load the rows
			 **************************************/
			if (this.chkFFE_TI1.Checked==true) 
			{
				this.listView1.Items.Add("Torch");
				this.listView1.Items[0].SubItems.Add("1");
				this.listView1.Items[0].SubItems.Add(strExp[0,0,0]);  
			}

			if (this.chkFFE_TI2.Checked==true) 
			{
				this.listView1.Items.Add("Torch");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("2");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[0,1,1]);
			}

			if (this.chkFFE_TI3.Checked==true) 
			{
				this.listView1.Items.Add("Torch");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("3");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[0,2,2]);
			}
			if (this.chkFFE_TI4.Checked==true) 
			{
				this.listView1.Items.Add("Torch");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("4");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[0,3,3]);
			}
			if (this.chkFFE_TI5.Checked==true) 
			{
				this.listView1.Items.Add("Torch");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("5");
    			this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[0,4,4]);
			}
			if (this.chkFFE_CI1.Checked==true) 
			{
				this.listView1.Items.Add("Crown");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("1");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[1,0,0]);
			}

			if (this.chkFFE_CI2.Checked==true) 
			{
				this.listView1.Items.Add("Crown");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("2");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[1,1,1]);
			}

			if (this.chkFFE_CI3.Checked==true) 
			{
				this.listView1.Items.Add("Crown");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("3");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[1,2,2]);
			}
			if (this.chkFFE_CI4.Checked==true) 
			{
				this.listView1.Items.Add("Crown");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("4");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[1,3,3]);
			}
			if (this.chkFFE_CI5.Checked==true) 
			{
				this.listView1.Items.Add("Crown");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add("5");
				this.listView1.Items[this.listView1.Items.Count-1].SubItems.Add(strExp[1,4,4]);
			}

		}
		public int val_windspeed_values(string strType)
		{
            this.m_intError=0;
			this.m_strError="";
			if (strType == "R") 
			{
				this.m_strError = "Run Scenario Failed: ";
				if (chkFFE_TI1.Checked) 
				{  
					if (chkFFE_TI2.Checked==false)
					{
						this.m_intError = -1;
						MessageBox.Show(this.m_strError + "A minimum of 2 TI wind speed classes need to be defined in Fuel And Fire Effects");
					}
					else
					  TorchIndex=true; 
				}
				else TorchIndex=false;
				if (chkFFE_CI1.Checked) 
				{
					if (chkFFE_CI2.Checked==false)
					{
						this.m_intError = -1;
						MessageBox.Show(this.m_strError + "A minimum of 2 CI wind speed classes need to be defined in Fuel And Fire Effects");
					}
					else CrownIndex=true;
				}
				else CrownIndex=false;
			}
			//validate the values to be saved
			if (this.chkFFE_TI1.Enabled && this.chkFFE_TI1.Checked == true)
			{
				if (this.txtFFE_TI1.Text.Trim().Length == 0)  
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 1 in Fuel And Fire Effects");
					return this.m_intError;
				}
				
			}

			if (this.chkFFE_TI2.Enabled && this.chkFFE_TI2.Checked == true)
			{
				if (this.txtFFE_TI2.Text.Trim().Length == 0) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 2");
					return this.m_intError;
				}
				
				if (Convert.ToInt32(txtFFE_TI1.Text) > Convert.ToInt32(this.txtFFE_TI2.Text)) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "FFE Wind Speed Class 1 should be less or equal to Class 2 Wind Speed");
					return this.m_intError;
				}
				
			}

			if (this.chkFFE_TI3.Enabled && this.chkFFE_TI3.Checked == true)
			{
				if (this.txtFFE_TI3.Text.Trim().Length == 0) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 3");
					return this.m_intError;
				}
				
				if (Convert.ToInt32(txtFFE_TI2.Text) > Convert.ToInt32(this.txtFFE_TI3.Text)) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "FFE Wind Speed Class 2 should be less or equal to Class 3 Wind Speed");
					return this.m_intError;
				}
				
			}

			if (this.chkFFE_TI4.Enabled && this.chkFFE_TI4.Checked == true)
			{
				if (this.txtFFE_TI4.Text.Trim().Length == 0) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 4");
					return this.m_intError;
				}
				
				if (Convert.ToInt32(txtFFE_TI3.Text) > Convert.ToInt32(this.txtFFE_TI4.Text)) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "FFE Wind Speed Class 3 should be less or equal to Class 4 Wind Speed");
					return this.m_intError;
				}
				
			}

			if (this.chkFFE_TI5.Enabled && this.chkFFE_TI5.Checked == true)
			{
				if (this.txtFFE_TI5.Text.Trim().Length == 0) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 5");
					return this.m_intError;
				}
				
				if (Convert.ToInt32(txtFFE_TI4.Text) > Convert.ToInt32(this.txtFFE_TI5.Text)) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "FFE Wind Speed Class 4 should be less or equal to Class 5 Wind Speed");
					return this.m_intError;
				}
				
			}


			if (this.chkFFE_CI1.Enabled && this.chkFFE_CI1.Checked == true)
			{
				if (this.txtFFE_CI1.Text.Trim().Length == 0)  
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 1");
					return this.m_intError;
				}
				
			}

			if (this.chkFFE_CI2.Enabled && this.chkFFE_CI2.Checked == true)
			{
				if (this.txtFFE_CI2.Text.Trim().Length == 0) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 2");
					return this.m_intError;
				}
				
				if (Convert.ToInt32(txtFFE_CI1.Text) > Convert.ToInt32(this.txtFFE_CI2.Text)) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "FFE Wind Speed Class 1 should be less or equal to Class 2 Wind Speed");
					return this.m_intError;
				}
				
			}

			if (this.chkFFE_CI3.Enabled && this.chkFFE_CI3.Checked == true)
			{
				if (this.txtFFE_CI3.Text.Trim().Length == 0) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 3");
					return this.m_intError;
				}
				
				if (Convert.ToInt32(txtFFE_CI2.Text) > Convert.ToInt32(this.txtFFE_CI3.Text)) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "FFE Wind Speed Class 2 should be less or equal to Class 3 Wind Speed");
					return this.m_intError;
				}
				
			}

			if (this.chkFFE_CI4.Enabled && this.chkFFE_CI4.Checked == true)
			{
				if (this.txtFFE_CI4.Text.Trim().Length == 0) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 4");
					return this.m_intError;
				}
				
				if (Convert.ToInt32(txtFFE_CI3.Text) > Convert.ToInt32(this.txtFFE_CI4.Text)) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "FFE Wind Speed Class 3 should be less or equal to Class 4 Wind Speed");
					return this.m_intError;
				}
				
			}

			if (this.chkFFE_CI5.Enabled && this.chkFFE_CI5.Checked == true)
			{
				if (this.txtFFE_CI5.Text.Trim().Length == 0) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "Enter FFE Wind Speed for Class 5");
					return this.m_intError;
				}
				
				if (Convert.ToInt32(txtFFE_CI4.Text) > Convert.ToInt32(this.txtFFE_CI5.Text)) 
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + "FFE Wind Speed Class 4 should be less or equal to Class 5 Wind Speed");
					return this.m_intError;
				}
				
			}
			return this.m_intError;
		}

		private void btnFFEExpressionStep2_Click(object sender, System.EventArgs e)
		{

			string strKeyStrokes="";
			this.grpboxFFEExpressionBuilder.Text="Expression Builder";
			if (this.listView1.SelectedItems.Count == 0) return;

			this.m_strCurrentIndexTypeAndClass=this.listView1.SelectedItems[0].SubItems[0].Text.Substring(0,1) + 
				  this.listView1.SelectedItems[0].SubItems[1].Text;

			this.lstFFEAvailableFields.Items.Clear();

			switch (this.m_strCurrentIndexTypeAndClass)
			{
				case "T1":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective torch index post-treatment result when the torch index pre-treatment sustainable fire wind speed class is equal to 1.";
                    this.lstFFEAvailableFields.Items.Add("TI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_TI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_TI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{						
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes= "(PRE_TI_CL = 1 AND ";
					}
				    break;
				case "T2":
      				this.lblFFEExpressionBuilder.Text = "Define what is an effective torch index post-treatment result when the torch index pre-treatment sustainable fire wind speed class is equal to 2.";
					this.lstFFEAvailableFields.Items.Add("TI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_TI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_TI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_TI_CL = 2 AND ";
					}
					break;
				case "T3":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective torch index post-treatment result when the torch index pre-treatment sustainable fire wind speed class is equal to 3.";
					this.lstFFEAvailableFields.Items.Add("TI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_TI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_TI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_TI_CL = 3 AND ";
					}
     			    break;
				case "T4":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective torch index post-treatment result when the torch index pre-treatment sustainable fire wind speed class is equal to 4.";
					this.lstFFEAvailableFields.Items.Add("TI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_TI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_TI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_TI_CL = 4 AND ";
					}
					break;
				case "T5":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective torch index post-treatment result when the torch index pre-treatment sustainable fire wind speed class is equal to 5.";
					this.lstFFEAvailableFields.Items.Add("TI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_TI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_TI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_TI_CL = 5 AND ";
					}
					break;
				case "C1":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective crown index post-treatment result when the crown index pre-treatment sustainable fire wind speed class is equal to 1.";
					this.lstFFEAvailableFields.Items.Add("CI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_CI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_CI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_CI_CL = 1 AND ";
					}
					break;
				case "C2":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective crown index post-treatment result when the crown index pre-treatment sustainable fire wind speed class is equal to 2.";
					this.lstFFEAvailableFields.Items.Add("CI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_CI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_CI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_CI_CL = 2 AND ";
					}
					break;
				case "C3":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective crown index post-treatment result when the crown index pre-treatment sustainable fire wind speed class is equal to 3.";
					this.lstFFEAvailableFields.Items.Add("CI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_CI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_CI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_CI_CL = 3 AND ";
					}
					break;
				case "C4":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective crown index post-treatment result when the crown index pre-treatment sustainable fire wind speed class is equal to 4.";
					this.lstFFEAvailableFields.Items.Add("CI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_CI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_CI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_CI_CL = 4 AND ";
					}
					break;
				case "C5":
					this.lblFFEExpressionBuilder.Text = "Define what is an effective crown index post-treatment result when the crown index pre-treatment sustainable fire wind speed class is equal to 5.";
					this.lstFFEAvailableFields.Items.Add("CI_CHANGE");
					this.lstFFEAvailableFields.Items.Add("PRE_CI_CL");
					this.lstFFEAvailableFields.Items.Add("POST_CI_CL");
					if (this.listView1.SelectedItems[0].SubItems[2].Text.Trim().Length > 0) 
					{
						strKeyStrokes = this.listView1.SelectedItems[0].SubItems[2].Text;
					}
					else 
					{
						strKeyStrokes = "(PRE_CI_CL = 5 AND ";
					}
					break;
			}
			this.grpboxFFEExpressionBuilder.Top = this.grpboxFFEClassifications.Top;
			this.grpboxFFEExpressionBuilder.Left = this.grpboxFFEClassifications.Left;
			this.grpboxFFEExpressionBuilder.Width = this.grpboxFFEClassifications.Width;

			this.btnFFEExpressionBuilderOk.Visible=true;
			this.btnFFEExpressionBuilderCancel.Visible=true;
			this.btnFFEExpressionBuilderTest.Width = this.btnFFEExpressionBuilderOk.Width;

			this.btnFFEExpressionBuilderPreviousExpression.Left = this.txtExpression.Left;
				
			this.btnFFEExpressionBuilderClear.Left =  this.btnFFEExpressionBuilderPreviousExpression.Left + 
				        this.btnFFEExpressionBuilderPreviousExpression.Width ;
			
			this.btnFFEExpressionBuilderTest.Left = 
				this.btnFFEExpressionBuilderClear.Left +
				this.btnFFEExpressionBuilderClear.Width;

			
            this.btnFFEExpressionBuilderOk.Left = 
				this.btnFFEExpressionBuilderTest.Left +
				this.btnFFEExpressionBuilderTest.Width;


			this.btnFFEExpressionBuilderCancel.Left = 
				this.btnFFEExpressionBuilderOk.Left +
				this.btnFFEExpressionBuilderOk.Width;


			this.grpboxFFE_TI_CI_Effective.Visible=false;

			this.lblFFEExpressionFieldDesc.Text ="";
			this.grpboxFFEExpressionBuilder.Visible=true;
			this.txtExpression.Focus();
			this.txtExpression.Text = "";
			this.txtExpression.AppendText(strKeyStrokes);
		}

		private void grpboxFFE_TI_CI_Effective_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void chkFFE_CI1_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("C1");
		}

		private void chkFFE_CI2_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("C2");
		}

		private void chkFFE_CI3_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("C3");
		}

		private void chkFFE_CI4_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("C4");
		}

		private void chkFFE_CI5_Click(object sender, System.EventArgs e)
		{
			this.val_checkbox("C5");
		}

		private void txtFFE_CI1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		   this.AllowNumericOnly(e);
		}

		private void txtFFE_CI2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_CI3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_CI4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_CI5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void listView1_Click(object sender, System.EventArgs e)
		{
			if (this.btnFFEExpressionStep2.Enabled==false) 
			{
				this.btnFFEExpressionStep2.Enabled=true;
				this.btnFFEClearExpression.Enabled=true;
				this.btnFFETICIEffectivePreviousExpressions.Enabled=true;
			}
		}

		private void btnFFECancel_Click(object sender, System.EventArgs e)
		{
			this.grpboxFFE_TI_CI_Effective.Visible=true;
			this.grpboxFFEExpressionBuilder.Visible=false;

		}

		private void lstFFEAvailableFields_SelectedValueChanged(object sender, System.EventArgs e)
		{
			switch (this.lstFFEAvailableFields.SelectedItem.ToString())
			{
				case "TI_CHANGE":
					this.lblFFEExpressionFieldDesc.Text = "ti_change = ffe.post_ti - ffe.pre_ti";
					break;
                case "PRE_TI_CL":
					this.lblFFEExpressionFieldDesc.Text = "Torch index pre-treatment wind speed class";
					break;
                case "POST_TI_CL":
					this.lblFFEExpressionFieldDesc.Text = "Torch index post-treatment wind speed class";
				    break;
				case "CI_CHANGE":
                    this.lblFFEExpressionFieldDesc.Text = "ci_change = ffe.post_ci - ffe.pre_ci";
					break;
				case "PRE_CI_CL":
					this.lblFFEExpressionFieldDesc.Text = "Crown index pre-treatment wind speed class";
				    break;
				case "POST_CI_CL":
					this.lblFFEExpressionFieldDesc.Text = "Crown index post-treatment wind speed class";
				    break;
                case "TI_EFFECTIVE_YN":
					this.lblFFEExpressionFieldDesc.Text = "User defined expression for identifying torch index effective post-treatments";
					break;
				case "CI_EFFECTIVE_YN":
					this.lblFFEExpressionFieldDesc.Text = "User defined expression for identifying crown index effective post-treatments";
					break;
				case "TI_BACKSLIDE_YN":
					this.lblFFEExpressionFieldDesc.Text = "User defined expression for identifying treatments that increase torch index fire potential";
					break;
				case "CI_BACKSLIDE_YN":
					this.lblFFEExpressionFieldDesc.Text = "User defined expression for identifying treatments that increase crown index fire potential";
					break;
				case "TI_HAZARD_YN":
					this.lblFFEExpressionFieldDesc.Text = "User defined expression for pre-treatment torch index conditions that are at a hazard for a potential sustainable fire.";
					break;
				case "CI_HAZARD_YN":
					this.lblFFEExpressionFieldDesc.Text = "User defined expression for pre-treatment crown index conditions that are at a hazard for a potential sustainable fire.";
					break;
				default:
					this.lblFFEExpressionFieldDesc.Text="";
					break;

			}	
		}

		private void btnFFEDefaultStep2_Click(object sender, System.EventArgs e)
		{
			int x=0;
            string str="";
			for (x=0; x <= this.listView1.Items.Count-1;x++)
			{

				str = this.listView1.Items[x].SubItems[0].Text.Substring(0,1) + 
					this.listView1.Items[x].SubItems[1].Text;
				this.listView1.Items[x].SubItems[2].Text=this.getDefaultExpression(str);

			}
		}

		private void btnFFEClearExpression_Click(object sender, System.EventArgs e)
		{
       		if (this.listView1.SelectedItems.Count == 0) return;

			this.listView1.SelectedItems[0].SubItems[2].Text = " ";
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void btnFFEExpressionBuilderDefault_Click(object sender, System.EventArgs e)
		{
			this.txtExpression.Text = "";
			if (this.btnPrev.Name == "overalleffective_prev")
			{
				this.txtExpression.AppendText(this.getDefaultExpression("overall"));
				//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
				//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
				((frmCoreScenario)this.ParentForm).m_bSave=true;
				this.m_strOverallEffectiveExpression=this.txtExpression.Text.Trim();
			}
			else 
			{
				this.txtExpression.AppendText(this.getDefaultExpression(this.m_strCurrentIndexTypeAndClass));
			}
		}
		protected string getDefaultExpression(string str)
		{
			switch (str)
			{
				case "T1":
					return "(PRE_TI_CL = 1 AND ((TI_CHANGE > 10 AND POST_TI_CL >= 2) OR (TI_CHANGE >= 20)))";
				case "T2":
					return "(PRE_TI_CL = 2 AND TI_CHANGE >= 15)";
				case "T3":
					return "(PRE_TI_CL = 3 AND TI_CHANGE >= 20)";
				case "T4":
					return "(PRE_TI_CL = 4 AND TI_CHANGE >= 25)";
				case "T5":
					return "(PRE_TI_CL = 5 AND TI_CHANGE >= 30)";
				case "C1":
					return " (PRE_CI_CL = 1 AND ((TI_CHANGE > 10 AND POST_CI_CL >= 2) OR (CI_CHANGE > 20)))";
				case "C2":
					return "(PRE_CI_CL = 2 AND CI_CHANGE >= 15)";
				case "C3":
					return "(PRE_CI_CL = 3 AND CI_CHANGE >= 20)";
				case "C4":
					return "(PRE_CI_CL = 4 AND CI_CHANGE >= 25)";
				case "C5":
					return  "(PRE_CI_CL = 5 AND CI_CHANGE >= 30)";
                case "overall":
					return "((TI_EFFECTIVE_YN = 'Y' AND (CI_BACKSLIDE_YN = 'N') AND (POST_CI_CL > 1)) " + 
						                       " OR " + 
					        "(CI_EFFECTIVE_YN = 'Y' AND (TI_BACKSLIDE_YN = 'N') AND (POST_TI_CL > 1)))";
                default:
					return "";
			}

		}

		private void lstFFEAvailableFields_DoubleClick(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression," " + this.lstFFEAvailableFields.SelectedItem.ToString().Trim() + " ");
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

		private void btnFFEEqual_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression," = ");
		}

		private void btnFFENotEqual_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression," <> ");
		}

		private void btnLeftBracket_Click(object sender, System.EventArgs e)
		{
			//string str= "{(}";
			this.SendKeyStrokes(this.txtExpression, " {(}");
			
		}

		private void btnRightBracket_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, "{)} ");
		}

		private void btnFFEMoreThan_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, " > ");
		}

		private void btnFFEMoreThanOrEqualTo_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, " >= ");
		}

		private void btnFFEAnd_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, " and ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, " AND ");
			}
		}

		private void btnFFELessThan_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, " < ");
		}

		private void btnLessThanOrEqualTo_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, " <= ");
		}

		private void btnFFEOr_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, " or ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, " OR ");
			}
		}

		private void btnFFENote_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, " not ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, " NOT ");
			}
		}

		private void btnFFEExpressionBuilderClear_Click(object sender, System.EventArgs e)
		{
			string strKeyStrokes="";
			this.txtExpression.Text="";
			if (this.btnPrev.Name != "overalleffective_prev")
			{
				switch (this.m_strCurrentIndexTypeAndClass)
				{
					case "T1":
						strKeyStrokes= "(PRE_TI_CL = 1 AND ";
						break;
					case "T2":
						strKeyStrokes = "(PRE_TI_CL = 2 AND ";
						break;
					case "T3":
						strKeyStrokes = "(PRE_TI_CL = 3 AND ";
						break;
					case "T4":
						strKeyStrokes = "(PRE_TI_CL = 4 AND ";
						break;
					case "T5":
						strKeyStrokes = "(PRE_TI_CL = 5 AND ";
						break;
					case "C1":
						strKeyStrokes = "(PRE_CI_CL = 1 AND ";
						break;
					case "C2":
						strKeyStrokes = "(PRE_CI_CL = 2 AND ";
						break;
					case "C3":
						strKeyStrokes = "(PRE_CI_CL = 3 AND ";
						break;
					case "C4":
						strKeyStrokes = "(PRE_CI_CL = 4 AND ";
						break;
					case "C5":
						strKeyStrokes = "(PRE_CI_CL = 5 AND ";
						break;
				}
				this.txtExpression.Focus();
				this.txtExpression.AppendText(strKeyStrokes);
			}


		}

		private void btnFFEExpressionBuilderTest_Click(object sender, System.EventArgs e)
		{
			string strSQL="";
			//int intArrayCount;
			//int x=0;
			string strConn="";
			//string strCommand="";
			//string str="";

			ado_data_access p_ado = new ado_data_access();

			string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            
			//scenario mdb connection
			//string strScenarioMDB = 
			//	((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + 
			//	"\\core\\db\\scenario.mdb";

			string strScenarioResultsMDB = 
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory +
                "\\core\\" + strScenarioId + "\\db\\scenario_results.mdb";

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn=p_ado.getMDBConnString(strScenarioResultsMDB,"admin","");
			//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strScenarioResultsMDB + ";User Id=admin;Password=;";
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				p_ado = null;
				return;
			}
			strSQL = "SELECT * FROM effective WHERE " +
				this.txtExpression.Text + ";";
			p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);

				  
			//insert records into the scenario_rx_intensity table from the master rx table
			if (p_ado.m_intError == 0)
			{
				

				MessageBox.Show("Valid Syntax");
				
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;
				
			}
			p_ado = null;
			this.m_OleDbConnectionScenario.Close();

		}

		private void btnFFEExpressionBuilderOk_Click(object sender, System.EventArgs e)
		{
		   this.listView1.SelectedItems[0].SubItems[2].Text=this.txtExpression.Text ;
		   this.grpboxFFEExpressionBuilder.Visible=false;
		   this.grpboxFFE_TI_CI_Effective.Visible = true;
		  // if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//   ((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		

		private void btnFFEPreviousExpression_Click(object sender, System.EventArgs e)
		{
			if (this.btnPrev.Name == "overalleffective_prev")
			{
				GetFFEPreviousExpressions("OVERALL EFFECTIVE");
			}
			else 
			{
				GetFFEPreviousExpressions("EXPRESSION BUILDER");
			}
		}

		private void GetFFEPreviousExpressions(string strType)
		{
			string strSQL="";
			//int x=0;
			string strConn="";
			//string strCommand="";
			//string str="";
			string strLargestString="";
			

			ado_data_access p_ado = new ado_data_access();

			string strScenarioMDB = 
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strScenarioMDB + ";User Id=admin;Password=;";
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				p_ado=null;
				return;
			}
           

				  
			//insert records into the scenario_rx_intensity table from the master rx table
			frmDialog frmTemp = new frmDialog();
			frmTemp.Text = "Core Analysis: Select Previous Expression";
			frmTemp.uc_select_list_item1.lblTitle.Text = "Previous Expressions";

			if (strType == "OVERALL EFFECTIVE") 
			{
				frmTemp.uc_select_list_item1.lblMsg.Text= "Previous Overall Effective Expressions";
				strSQL = "SELECT expression FROM scenario_ffe_overall_effective_change";
			}
			else 
			{
				frmTemp.uc_select_list_item1.lblMsg.Text= "Previous Torch And Crown Index Effective Expressions";
				strSQL = "SELECT expression FROM scenario_ffe_ti_ci_effective_change";
			}
			frmTemp.uc_select_list_item1.lblMsg.Visible = true;
						
			frmTemp.uc_select_list_item1.Dock = System.Windows.Forms.DockStyle.Fill;
						
			frmTemp.uc_project1.Visible=false;
			frmTemp.uc_select_list_item1.listBox1.Items.Clear();

			
			p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);
			//insert records into the scenario_rx_intensity table from the master rx table
			if (p_ado.m_intError == 0)
			{
				

				while (p_ado.m_OleDbDataReader.Read())
				{
					
					if (p_ado.m_OleDbDataReader["expression"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["expression"].ToString().Length > 1)
						{
							frmTemp.uc_select_list_item1.listBox1.Items.Add(p_ado.m_OleDbDataReader["expression"].ToString());
							if (p_ado.m_OleDbDataReader["expression"].ToString().Trim().Length > 
								strLargestString.ToString().Trim().Length) 
							{
								strLargestString = p_ado.m_OleDbDataReader["expression"].ToString();
							}
						}
					}
				}
                
				frmTemp.uc_select_list_item1.Initialize_Width(strLargestString);

				
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;
				
			}			
						

			frmTemp.uc_select_list_item1.Visible=true;
			DialogResult result = frmTemp.ShowDialog(this);
                        
						
			if (result == DialogResult.OK) 
			{
				if (strType == "EXPRESSION BUILDER" || strType == "OVERALL EFFECTIVE")
				{
					this.txtExpression.Text = frmTemp.uc_select_list_item1.listBox1.Text;
				}
				else 
				{
					this.listView1.SelectedItems[0].SubItems[2].Text = frmTemp.uc_select_list_item1.listBox1.Text;
				}
				
				((frmCoreScenario)this.ParentForm).m_bSave=true;
				
				
			}
					
			frmTemp.Close();
			frmTemp = null;

		}

		private void trbFFETIBackslide_ValueChanged(object sender, System.EventArgs e)
		{
		   this.lblFFETIBackslide.Text = Convert.ToString(trbFFETIBackslide.Value * -1);
			
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		
		private void trbFFECIBackslide_ValueChanged(object sender, System.EventArgs e)
		{
		 this.lblFFECIBackslide.Text = Convert.ToString(trbFFECIBackslide.Value * -1);
        
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void chkFFETIBackslide_Click(object sender, System.EventArgs e)
		{
			if (this.chkFFETIBackslide.Checked==true)
			{
                this.cmbFFETIBackslide2.Enabled=true;
				this.cmbFFETIBackslide3.Enabled=true;
			}
			else 
			{
				this.cmbFFETIBackslide2.Enabled=false;
				this.cmbFFETIBackslide3.Enabled=false;
			}
		}

		private void chkFFECIBackslide_Click(object sender, System.EventArgs e)
		{
			if (this.chkFFECIBackslide.Checked==true)
			{
				this.cmbFFECIBackslide2.Enabled=true;
				this.cmbFFECIBackslide3.Enabled=true;
			}
			else 
			{
				this.cmbFFECIBackslide2.Enabled=false;
				this.cmbFFECIBackslide3.Enabled=false;
			}

		}
		private void loadvalues_backslide()
		{
			int x=0;
			
			this.cmbFFETIBackslide3.Items.Clear();
			this.cmbFFECIBackslide3.Items.Clear();
			if (this.chkFFE_TI1.Checked==true) 
			{
				this.cmbFFETIBackslide3.Items.Add("1");
			}

			if (this.chkFFE_TI2.Checked==true) 
			{
				this.cmbFFETIBackslide3.Items.Add("2");
			}

			if (this.chkFFE_TI3.Checked==true) 
			{
				this.cmbFFETIBackslide3.Items.Add("3");
			}
			if (this.chkFFE_TI4.Checked==true) 
			{
				this.cmbFFETIBackslide3.Items.Add("4");
			}
			if (this.chkFFE_TI5.Checked==true) 
			{
				this.cmbFFETIBackslide3.Items.Add("5");
			}
			if (this.chkFFE_CI1.Checked==true) 
			{
				this.cmbFFECIBackslide3.Items.Add("1");
			}

			if (this.chkFFE_CI2.Checked==true) 
			{
				this.cmbFFECIBackslide3.Items.Add("2");
			}
			if (this.chkFFE_CI3.Checked==true) 
			{
				this.cmbFFECIBackslide3.Items.Add("3");
			}
			if (this.chkFFE_CI4.Checked==true) 
			{
				this.cmbFFECIBackslide3.Items.Add("4");
			}
			if (this.chkFFE_CI5.Checked==true) 
			{
				this.cmbFFECIBackslide3.Items.Add("5");
			}

			//see if the current value is valid
			if (this.cmbFFETIBackslide3.Text.Trim().Length > 0)
			{
				for (x=0; x <= this.cmbFFETIBackslide3.Items.Count-1; x++)
				{
			       if (this.cmbFFETIBackslide3.Text.Trim() ==
					   this.cmbFFETIBackslide3.Items[x].ToString().Trim())
					      break;

				}
				if (x > this.cmbFFETIBackslide3.Items.Count-1) this.cmbFFETIBackslide3.Text=""; 

			}

			//see if the current value is valid
			if (this.cmbFFECIBackslide3.Text.Trim().Length > 0)
			{
				for (x=0; x <= this.cmbFFECIBackslide3.Items.Count-1; x++)
				{
					if (this.cmbFFECIBackslide3.Text.Trim() ==
						this.cmbFFECIBackslide3.Items[x].ToString().Trim())
						break;

				}
				if (x > this.cmbFFECIBackslide3.Items.Count-1) this.cmbFFECIBackslide3.Text=""; 

			}

		}
		private void loadvalues_hazard()
		{
			int x=0;
			
			this.cmbFFETIHazardWindSpeedClass.Items.Clear();
			this.cmbFFECIHazardWindSpeedClass.Items.Clear();
			if (this.chkFFE_TI1.Checked==true) 
			{
				this.cmbFFETIHazardWindSpeedClass.Items.Add("1");
			}

			if (this.chkFFE_TI2.Checked==true) 
			{
				this.cmbFFETIHazardWindSpeedClass.Items.Add("2");
			}

			if (this.chkFFE_TI3.Checked==true) 
			{
				this.cmbFFETIHazardWindSpeedClass.Items.Add("3");
			}
			if (this.chkFFE_TI4.Checked==true) 
			{
				this.cmbFFETIHazardWindSpeedClass.Items.Add("4");
			}
			if (this.chkFFE_TI5.Checked==true) 
			{
				this.cmbFFETIHazardWindSpeedClass.Items.Add("5");
			}
			if (this.chkFFE_CI1.Checked==true) 
			{
				this.cmbFFECIHazardWindSpeedClass.Items.Add("1");
			}

			if (this.chkFFE_CI2.Checked==true) 
			{
				this.cmbFFECIHazardWindSpeedClass.Items.Add("2");
			}
			if (this.chkFFE_CI3.Checked==true) 
			{
				this.cmbFFECIHazardWindSpeedClass.Items.Add("3");
			}
			if (this.chkFFE_CI4.Checked==true) 
			{
				this.cmbFFECIHazardWindSpeedClass.Items.Add("4");
			}
			if (this.chkFFE_CI5.Checked==true) 
			{
				this.cmbFFECIHazardWindSpeedClass.Items.Add("5");
			}

			//see if the current value is valid
			if (this.cmbFFETIHazardWindSpeedClass.Text.Trim().Length > 0)
			{
				for (x=0; x <= this.cmbFFETIHazardWindSpeedClass.Items.Count-1; x++)
				{
					if (this.cmbFFETIHazardWindSpeedClass.Text.Trim() ==
						this.cmbFFETIHazardWindSpeedClass.Items[x].ToString().Trim())
						break;

				}
				if (x > this.cmbFFETIHazardWindSpeedClass.Items.Count-1) this.cmbFFETIHazardWindSpeedClass.Text=""; 

			}

			//see if the current value is valid
			if (this.cmbFFECIHazardWindSpeedClass.Text.Trim().Length > 0)
			{
				for (x=0; x <= this.cmbFFECIHazardWindSpeedClass.Items.Count-1; x++)
				{
					if (this.cmbFFECIHazardWindSpeedClass.Text.Trim() ==
						this.cmbFFECIHazardWindSpeedClass.Items[x].ToString().Trim())
						break;

				}
				if (x > this.cmbFFECIHazardWindSpeedClass.Items.Count-1) this.cmbFFECIHazardWindSpeedClass.Text=""; 
			}

		}

		private void btnFFEDefaultStep3_Click(object sender, System.EventArgs e)
		{
			
			this.chkFFETIBackslide.Enabled=true;
			this.chkFFETIBackslide.Checked=true;
			this.cmbFFETIBackslide2.Enabled=true;
			this.cmbFFETIBackslide3.Enabled=true;
			this.chkFFECIBackslide.Enabled=true;
			this.chkFFECIBackslide.Checked=true;
			this.cmbFFECIBackslide2.Enabled=true;
			this.cmbFFECIBackslide3.Enabled=true;


			this.cmbFFETIBackslide.Text="<";
			this.trbFFETIBackslide.Value = 10;
			this.cmbFFECIBackslide.Text="<";
			this.trbFFECIBackslide.Value = 10;


			this.cmbFFECIBackslide2.Text = "<=";
			if (this.chkFFE_TI2.Checked==true) 
			{
			  this.cmbFFETIBackslide2.Text = "<=";
              this.cmbFFETIBackslide3.Text = "2";
			}
			else
			{
			  this.cmbFFETIBackslide2.Text = "=";
			  this.cmbFFETIBackslide3.Text = "1";
			}

			if (this.chkFFE_CI2.Checked==true) 
			{
				this.cmbFFECIBackslide2.Text = "<=";
				this.cmbFFECIBackslide3.Text = "2";
			}
			else
			{
				this.cmbFFECIBackslide2.Text = "=";
				this.cmbFFECIBackslide3.Text = "1";
			}

		}

		private void cmbFFETIBackslide_TextChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide_TextChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide2_TextChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide2_TextChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide3_TextChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide3_TextChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}
		public int val_ti_ci_effective_expression()
		{
			int x=0;
			this.m_intError=0;
			this.m_strError = "Run Scenario Failed: ";

			if (this.listView1.Items.Count == 0)
			{
				this.m_intError=-1;
				MessageBox.Show(this.m_strError + "Need user defined torch and crown index effective treatment expression values in <Fuel and Fire Effects>");
				return this.m_intError;
			}   
			for (x=0; x<=this.listView1.Items.Count-1;x++)
			{
				if (this.listView1.Items[x].SubItems[2].Text.Trim().Length == 0)
				{
					this.m_intError=-1;
					MessageBox.Show(this.m_strError + 
                           this.listView1.Items[x].SubItems[0].Text.Trim() + 
                           " Index Wind Speed Class " + 
                           this.listView1.Items[x].SubItems[1].Text.Trim() + 
                           " effective treatment expression needs to be defined in <Fuel And Fire Effects>");
					return this.m_intError;
				}
			}

            return this.m_intError;



		}
		public int val_overall_effective_expression()
		{
			//int x=0;
			this.m_intError=0;
			this.m_strError = "Run Scenario Failed: ";

			if (this.m_strOverallEffectiveExpression.Trim().Length==0)
			{
				this.m_intError=-1;
				MessageBox.Show(this.m_strError + "Need user defined expression for overall effective treatment in <Fuel and Fire Effects>");
				return this.m_intError;
			}   
			return this.m_intError;
			
		}
		public int val_backslide(string strType)
		{
			this.m_intError=-1;
			this.m_strError="";
			if (strType == "R") this.m_strError="Run Scenario Failed:";
            int x=0;
			//torch index backslide
            //validate expression operator 1
			for (x=0;x<=this.cmbFFETIBackslide.Items.Count-1;x++)
			{
				if (this.cmbFFETIBackslide.Text ==
					this.cmbFFETIBackslide.Items[x].ToString().Trim())
				{
					this.m_intError=0;
					break;
				}
			}
			
			if (this.m_intError < 0)
			{
				MessageBox.Show(this.m_strError + "Invalid backslide expression operator in <Fuel And Fire Effects>");
				return this.m_intError;
			}


			if (this.chkFFETIBackslide.Enabled &&
				this.chkFFETIBackslide.Checked==true)
			{
				this.m_intError=-1;
				for (x=0;x<=this.cmbFFETIBackslide2.Items.Count-1;x++)
				{
					if (this.cmbFFETIBackslide2.Text ==
						this.cmbFFETIBackslide2.Items[x].ToString().Trim())
					{
						this.m_intError=0;
						break;
					}
				}
			
				if (this.m_intError < 0)
				{
					MessageBox.Show(this.m_strError + "Invalid backslide expression operator <Fuel And Fire Effects>");
					return this.m_intError;
				}


				this.m_intError=-1;
				for (x=0;x<=this.cmbFFETIBackslide3.Items.Count-1;x++)
				{
					if (this.cmbFFETIBackslide3.Text ==
						this.cmbFFETIBackslide3.Items[x].ToString().Trim())
					{
						this.m_intError=0;
						break;
					}
				}
			
				if (this.m_intError < 0)
				{
					MessageBox.Show(this.m_strError + "Backslide threshholds invalid wind speed class in <Fuel And Fire Effects>");
					return this.m_intError;
				}
			
			}

			//crown index backslide
			this.m_intError=-1;
			for (x=0;x<=this.cmbFFECIBackslide.Items.Count-1;x++)
			{
				if (this.cmbFFECIBackslide.Text ==
					this.cmbFFECIBackslide.Items[x].ToString().Trim())
				{
					this.m_intError=0;
					break;
				}
			}
			
			if (this.m_intError < 0)
			{
				MessageBox.Show(this.m_strError + "Invalid backslide expression operator in <Fuel And Fire Effects>");
				return this.m_intError;
			}


			if (this.chkFFECIBackslide.Enabled &&
				this.chkFFECIBackslide.Checked==true)
			{
				this.m_intError=-1;
				for (x=0;x<=this.cmbFFECIBackslide2.Items.Count-1;x++)
				{
					if (this.cmbFFECIBackslide2.Text ==
						this.cmbFFECIBackslide2.Items[x].ToString().Trim())
					{
						this.m_intError=0;
						break;
					}
				}
			
				if (this.m_intError < 0)
				{
					MessageBox.Show(this.m_strError + "Invalid backslide expression operator in <Fuel And Fire Effects>");
					return this.m_intError;
				}


				this.m_intError=-1;
				for (x=0;x<=this.cmbFFECIBackslide3.Items.Count-1;x++)
				{
					if (this.cmbFFECIBackslide3.Text ==
						this.cmbFFECIBackslide3.Items[x].ToString().Trim())
					{
						this.m_intError=0;
						break;
					}
				}
			
				if (this.m_intError < 0)
				{
					MessageBox.Show(this.m_strError + "Backslide threshhold invalid wind speed class in <Fuel And Fire Effects>");
					return this.m_intError;
				}
			
			}
			return this.m_intError;


            
		}
		public int val_hazard(string strType)
		{
			this.m_intError=0;
			this.m_strError="";
			if (strType == "R") this.m_strError="Run Scenario Failed:";
			int x=0;
			//torch index backslide
			//validate expression operator 1
			if (this.chkFFETIHazard.Enabled &&
				this.chkFFETIHazard.Checked==true)
			{
				this.m_intError=-1;
				this.m_strError="";
				for (x=0;x<=this.cmbFFETIHazardOperator.Items.Count-1;x++)
				{
					if (this.cmbFFETIHazardOperator.Text ==
						this.cmbFFETIHazardOperator.Items[x].ToString().Trim())
					{
						this.m_intError=0;
						break;
					}
				}
				if (this.m_intError < 0)
				{
					MessageBox.Show(this.m_strError + "Hazardous condition invalid hazard expression operator in <Fuel And Fire Effects>");
					return this.m_intError;
				}

				this.m_intError=-1;
				for (x=0;x<=this.cmbFFETIHazardWindSpeedClass.Items.Count-1;x++)
				{
					if (this.cmbFFETIHazardWindSpeedClass.Text ==
						this.cmbFFETIHazardWindSpeedClass.Items[x].ToString().Trim())
					{
						this.m_intError=0;
						break;
					}
				}
			
				if (this.m_intError < 0)
				{
					MessageBox.Show(this.m_strError + "Hazardous condition invalid wind speed class in <Fuel And Fire Effects>");
					return this.m_intError;
				}
		    }
			if (this.chkFFECIHazard.Enabled &&
				this.chkFFECIHazard.Checked==true)
			{
				for (x=0;x<=this.cmbFFECIHazardOperator.Items.Count-1;x++)
				{
					if (this.cmbFFECIHazardOperator.Text ==
						this.cmbFFECIHazardOperator.Items[x].ToString().Trim())
					{
						this.m_intError=0;
						break;
					}
				}
				if (this.m_intError < 0)
				{
					MessageBox.Show(this.m_strError + "Hazardous condition invalid hazard expression operator in <Fuel And Fire Effects>");
					return this.m_intError;
				}

				this.m_intError=-1;
				for (x=0;x<=this.cmbFFECIHazardWindSpeedClass.Items.Count-1;x++)
				{
					if (this.cmbFFECIHazardWindSpeedClass.Text ==
						this.cmbFFECIHazardWindSpeedClass.Items[x].ToString().Trim())
					{
						this.m_intError=0;
						break;
					}
				}
			
				if (this.m_intError < 0)
				{
					MessageBox.Show(this.m_strError + "Hazardous condition invalid wind speed class in <Fuel And Fire Effects>");
					return this.m_intError;
				}
			}
			return this.m_intError;

		}
		
		private void btnNext_Click(object sender, System.EventArgs e)
		{
			//MessageBox.Show(this.btnNext.Name);
			switch (this.btnNext.Name)
			{
				case "windspeed_next":
					this.val_windspeed_values("");
					if (this.m_intError < 0) return;
					this.grpboxFFE_TI_CI_Effective.Top = this.grpboxFFEClassifications.Top;
					this.grpboxFFE_TI_CI_Effective.Left = this.grpboxFFEClassifications.Left;
					this.grpboxFFE_TI_CI_Effective.Width = this.grpboxFFEClassifications.Width;
					this.loadvalues_ti_ci_effective();
					this.btnFFEExpressionStep2.Enabled=false;
					this.btnFFEClearExpression.Enabled=false;
					this.btnFFETICIEffectivePreviousExpressions.Enabled=false;
                    this.NextButton(ref this.grpboxFFE_TI_CI_Effective,ref this.btnNext,"ti_ci_effective_next");
					this.PrevButton(ref this.grpboxFFE_TI_CI_Effective,ref this.btnPrev,"ti_ci_effective_prev");
                    
					this.grpboxFFE_TI_CI_Effective.Visible=true;
					
					this.grpboxFFEClassifications.Visible=false;
					break;
                case "ti_ci_effective_next":
					this.grpboxFFEBackslide.Top = this.grpboxFFEClassifications.Top;
					this.grpboxFFEBackslide.Left = this.grpboxFFEClassifications.Left;
					this.grpboxFFEBackslide.Width = this.grpboxFFEClassifications.Width;
					
					this.grpboxFFE_TI_CI_Effective.Visible=false;
					this.loadvalues_backslide();
					this.NextButton(ref this.grpboxFFEBackslide,ref this.btnNext,"backslide_next");
					this.PrevButton(ref this.grpboxFFEBackslide,ref this.btnPrev,"backslide_prev");
					this.grpboxFFEBackslide.Visible=true;

					break;
                case "backslide_next":
					this.grpboxFFEHazard.Top = this.grpboxFFEClassifications.Top;
					this.grpboxFFEHazard.Left = this.grpboxFFEClassifications.Left;
					this.grpboxFFEHazard.Width = this.grpboxFFEClassifications.Width;
					
					this.grpboxFFEBackslide.Visible=false;
					this.loadvalues_hazard();
					this.NextButton(ref this.grpboxFFEHazard,ref this.btnNext,"hazard_next");
					this.PrevButton(ref this.grpboxFFEHazard,ref this.btnPrev,"hazard_prev");
					this.grpboxFFEHazard.Visible=true;
					break;
                case "hazard_next":
					this.grpboxFFEExpressionBuilder.Top = this.grpboxFFEClassifications.Top;
					this.grpboxFFEExpressionBuilder.Left = this.grpboxFFEClassifications.Left;
					this.grpboxFFEExpressionBuilder.Width = this.grpboxFFEClassifications.Width;
					this.grpboxFFEHazard.Visible=false;
                    this.loadvalues_overall();
					this.PrevButton(ref this.grpboxFFEExpressionBuilder,ref this.btnPrev,"overalleffective_prev");
					this.btnFFEExpressionBuilderOk.Visible=false;
					this.btnFFEExpressionBuilderCancel.Visible=false;
					this.btnFFEExpressionBuilderClear.Left = (int)(this.grpboxFFEExpressionBuilder.Width * .50) 
						                                       - (int)(this.btnFFEExpressionBuilderClear.Width * .50);
					this.btnFFEExpressionBuilderPreviousExpression.Left = 
						         this.btnFFEExpressionBuilderClear.Left - 
						         this.btnFFEExpressionBuilderPreviousExpression.Width;
					this.btnFFEExpressionBuilderTest.Left = 
                                 this.btnFFEExpressionBuilderClear.Left +
						         this.btnFFEExpressionBuilderClear.Width;
					this.btnFFEExpressionBuilderTest.Width = this.btnFFEExpressionBuilderClear.Width;
										
					this.grpboxFFEExpressionBuilder.Visible=true;
					break;
                
				default:
					break;
			}
		}
		private void btnPrev_Click(object sender, System.EventArgs e)
		{
			//MessageBox.Show(this.btnPrev.Name);
			switch (this.btnPrev.Name)
			{
				case "ti_ci_effective_prev":
					this.grpboxFFE_TI_CI_Effective.Visible=false;
					this.NextButton(ref this.grpboxFFEClassifications,ref this.btnNext,"windspeed_next");
					this.grpboxFFEClassifications.Visible=true;
					break;
                case "backslide_prev":
					this.grpboxFFEBackslide.Visible=false;
					this.NextButton(ref this.grpboxFFE_TI_CI_Effective,ref this.btnNext,"ti_ci_effective_next");
					this.PrevButton(ref this.grpboxFFE_TI_CI_Effective,ref this.btnPrev,"ti_ci_effective_prev");
					this.grpboxFFE_TI_CI_Effective.Visible=true;
					break;
                case "hazard_prev":
					this.grpboxFFEHazard.Visible=false;
					this.NextButton(ref this.grpboxFFEBackslide,ref this.btnNext,"backslide_next");
					this.PrevButton(ref this.grpboxFFEBackslide,ref this.btnPrev,"backslide_prev");
					this.grpboxFFEBackslide.Visible=true;
					break;
				case "overalleffective_prev":
					this.m_strOverallEffectiveExpression=this.txtExpression.Text;
					this.grpboxFFEExpressionBuilder.Visible=false;
					this.NextButton(ref this.grpboxFFEHazard,ref this.btnNext,"hazard_next");
					this.PrevButton(ref this.grpboxFFEHazard,ref this.btnPrev,"hazard_prev");
					this.grpboxFFEHazard.Visible=true;
					break;
				default:
					break;
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
			//p_oButton.Left = this.btnNext.Left - p_oButton.Width;  //p_oGb.Width - this.btnNext.Width - 5 - p_oButton.Width;
			p_oButton.Top = this.btnNext.Top;
			p_oButton.Height = this.btnNext.Height;
			p_oButton.Width = this.btnNext.Width;
			p_oButton.Left = this.btnNext.Left - p_oButton.Width;
			p_oButton.Name = strButtonName;	
		}

		private void chkFFETIHazard_Click(object sender, System.EventArgs e)
		{
			if (this.chkFFETIHazard.Checked==true)
			{
				this.cmbFFETIHazardOperator.Enabled=true;
				this.cmbFFETIHazardWindSpeedClass.Enabled=true;
			}
			else 
			{
				this.cmbFFETIHazardOperator.Enabled=false;
				this.cmbFFETIHazardWindSpeedClass.Enabled=false;

			}
		}

		private void chkFFECIHazard_Click(object sender, System.EventArgs e)
		{
			if (this.chkFFECIHazard.Checked==true)
			{
				this.cmbFFECIHazardOperator.Enabled=true;
				this.cmbFFECIHazardWindSpeedClass.Enabled=true;
			}
			else 
			{
				this.cmbFFECIHazardOperator.Enabled=false;
				this.cmbFFECIHazardWindSpeedClass.Enabled=false;

			}
		}

		private void btnFFEHazardDefault_Click(object sender, System.EventArgs e)
		{
			this.chkFFETIHazard.Enabled=true;
			this.chkFFETIHazard.Checked=true;
			this.cmbFFETIHazardOperator.Enabled=true;
			this.cmbFFETIHazardOperator.Text = "=";
			this.cmbFFETIHazardWindSpeedClass.Enabled=true;
			this.cmbFFETIHazardWindSpeedClass.Text = "1";
			this.chkFFECIHazard.Enabled=true;
			this.chkFFECIHazard.Checked=true;
			this.cmbFFECIHazardOperator.Enabled=true;
			this.cmbFFECIHazardOperator.Text = "=";
			this.cmbFFECIHazardWindSpeedClass.Enabled=true;
			this.cmbFFECIHazardWindSpeedClass.Text = "1";
			

		}
		private void loadvalues_overall()
		{
			//string strKeyStrokes="";
			this.grpboxFFEExpressionBuilder.Text="Expression Builder For Overall Effective Treatment - Step 5";
			
			this.lstFFEAvailableFields.Items.Clear();

			this.lblFFEExpressionBuilder.Text = "Define the overall effectiveness of a treatment.";


			this.lstFFEAvailableFields.Items.Add("TI_CHANGE");
			this.lstFFEAvailableFields.Items.Add("CI_CHANGE");
			this.lstFFEAvailableFields.Items.Add("TI_EFFECTIVE_YN");
			this.lstFFEAvailableFields.Items.Add("CI_EFFECTIVE_YN");
			this.lstFFEAvailableFields.Items.Add("TI_BACKSLIDE_YN");
			this.lstFFEAvailableFields.Items.Add("CI_BACKSLIDE_YN");
			this.lstFFEAvailableFields.Items.Add("AT_HAZARD_YN");
			this.lstFFEAvailableFields.Items.Add("PRE_TI_CL");
			this.lstFFEAvailableFields.Items.Add("PRE_CI_CL");
			this.lstFFEAvailableFields.Items.Add("POST_TI_CL");
            this.lstFFEAvailableFields.Items.Add("POST_CI_CL");
            this.txtExpression.Text = this.m_strOverallEffectiveExpression;
			this.txtExpression.Focus();
			

		}

		private void cmbFFECIBackslide2_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide3_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide3_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide2_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIHazardOperator_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIHazardWindSpeedClass_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIHazardOperator_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIHazardWindSpeedClass_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmCoreScenario)this.ParentForm).m_bSave=true;
		}

		private void txtExpression_Leave(object sender, System.EventArgs e)
		{
			if (this.btnPrev.Name == "overalleffective_prev")
			{
				this.m_strOverallEffectiveExpression=this.txtExpression.Text;
			}
		}

		private void btnFFETICIEffectivePreviousExpressions_Click(object sender, System.EventArgs e)
		{
		     GetFFEPreviousExpressions("TICI_EFFECTIVE");
		}

		private void cmbFFE_TI1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_TI2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_TI3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_TI4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_TI5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI3_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbFFE_CI3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIBackslide_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIBackslide_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIBackslide2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIBackslide2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIBackslide3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIBackslide3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIHazardOperator_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIHazardOperator_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIHazardWindSpeedClass_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIHazardWindSpeedClass_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtExpression_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.grpboxFFEExpressionBuilder.Text.Trim().ToUpper()!="EXPRESSION BUILDER")
			{
				
				
			   //if (((frmScenario)this.ParentForm).btnSave.Enabled == false) 
				//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
				((frmCoreScenario)this.ParentForm).m_bSave=true;
				
			}
		}

		
		private void btnY_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, "'y' ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, "'Y' ");
			}
		}

		private void btnN_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, "'n' ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, "'N' ");
			}
		}

		private void cmdFFE_Click(object sender, System.EventArgs e)
		{
			//if (this.cmdFFE.ImageIndex == 1) 
			//{
			//	this.cmdFFE.ImageIndex = 0;
			//	this.Height = this.grpboxFFE.Height; 

			//}
			//else 
			//{
			//	this.cmdFFE.ImageIndex = 1;
			//	this.Height = this.grpboxFFE.Height + this.groupBox1.Top + (this.grpboxFFEClassifications.Top + this.grpboxFFEClassifications.Height);

			//}
			//((frmScenario)this.ParentForm).RulesRepositionControls();

		}
		public bool TorchIndex
		{
			set {_bTorchIndex=value;}
			get {return _bTorchIndex;}
		}

		public bool CrownIndex
		{
			set {_bCrownIndex=value;}
			get {return _bCrownIndex;}
		}
		public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
	
	}
}
