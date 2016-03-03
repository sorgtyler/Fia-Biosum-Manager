using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Xml;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_print_report_wizard.
	/// </summary>
	public class uc_print_report_wizard : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnNext;
		private System.Windows.Forms.Button btnGetTable;
		private System.Windows.Forms.ListBox lstAvailableFields;
		private System.Windows.Forms.Button btnAdd;
		private System.Windows.Forms.Button btnRemove;
		private System.Windows.Forms.Button btnRemoveAll;
		private System.Windows.Forms.ListBox lstReportFields;
		private System.Windows.Forms.Button btnFinish;
		private System.Windows.Forms.ComboBox cmbSteps;
		public System.Windows.Forms.GroupBox grpboxSelectFields;
		public System.Windows.Forms.GroupBox grpboxGroupRecords;
		private System.Windows.Forms.ComboBox cmbGroup1;
		private System.Windows.Forms.ComboBox cmbGroup2;
		private System.Windows.Forms.Button btnAddAll;
		private System.Windows.Forms.Button btnPrevious;
		private System.Windows.Forms.ComboBox cmbGroup3;
		private System.Windows.Forms.GroupBox grpboxSort;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox grpboxDetailAndSummary;
		//private System.Data.DataSet m_ds;
		private System.Data.DataTable m_dt;
		private System.Windows.Forms.DataGrid m_dg;
		private System.Data.DataSet m_dsXMLSchema;
		public System.Data.DataTable m_dtDetailOptions;
		//private System.Data.DataSet m_dsDetailOptions;
		private System.Data.DataView m_dvDetailOptions;


		public int m_DialogHt;
		public int m_DialogWd;
		private System.Windows.Forms.Label lblTableName;
		private System.Windows.Forms.Button btnTop;
		private System.Windows.Forms.Button btnBottom;
		private System.Windows.Forms.Button btnUp;
		private System.Windows.Forms.Button btnDown;
		private string m_strAppDir="";
		private System.Windows.Forms.GroupBox grpboxFinish;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtReportTitle;
		private System.Windows.Forms.Button btnPreview;
		private System.Windows.Forms.RadioButton rdoColumn;
		private System.Windows.Forms.RadioButton rdoRow;
		private System.Windows.Forms.PictureBox picRecordLayout;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.GroupBox groupBox7;
		private System.Windows.Forms.RadioButton rdoPortrait;
		private System.Windows.Forms.RadioButton rdoLandscape;
		private System.Windows.Forms.Label label3;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.Label lblPortrait;
		private System.Windows.Forms.Label lblLandscape;
		private System.Windows.Forms.RadioButton rdoGroup1Asc;
		private System.Windows.Forms.RadioButton rdoGroup1Desc;
		private System.Windows.Forms.RadioButton rdoGroup2Asc;
		private System.Windows.Forms.RadioButton rdoGroup2Desc;
		private System.Windows.Forms.RadioButton rdoGroup3Asc;
		private System.Windows.Forms.RadioButton rdoGroup3Desc;
		private System.Windows.Forms.GroupBox grpboxGroup3;
		private System.Windows.Forms.GroupBox grpboxGroup2;
		private System.Windows.Forms.GroupBox grpboxGroup1;
		private System.Windows.Forms.GroupBox groupBox8;
		private System.Windows.Forms.RadioButton radioButton3;
		private System.Windows.Forms.RadioButton radioButton4;
		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.RadioButton rdoSort1Desc;
		private System.Windows.Forms.RadioButton rdoSort1Asc;
		private System.Windows.Forms.ComboBox cmbSort1;
		private System.Windows.Forms.RadioButton rdoSort3Desc;
		private System.Windows.Forms.RadioButton rdoSort3Asc;
		private System.Windows.Forms.ComboBox cmbSort3;
		private System.Windows.Forms.RadioButton rdoSort2Desc;
		private System.Windows.Forms.RadioButton rdoSort2Asc;
		private System.Windows.Forms.ComboBox cmbSort2;
		private System.Windows.Forms.GroupBox grpboxSort1;
		public System.Windows.Forms.GroupBox grpboxMain;
		private System.Windows.Forms.GroupBox grpboxSort3;
		private System.Windows.Forms.GroupBox grpboxSort2;
		private System.Windows.Forms.GroupBox grpboxRecordLayout;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.DataGrid dgDetailOptions;
		private System.Windows.Forms.CheckBox chkCountGroupRecords;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.CheckBox chkPrintDetail;
		private System.Windows.Forms.CheckBox chkPrintGroupSummary;
		private System.Windows.Forms.CheckBox chkPrintReportTotals;
		private System.Windows.Forms.CheckBox chkPctGroupRecords;
		private System.Windows.Forms.Button btnHelp;
		private string m_strAction = "";
		private string m_strType;
		
		public uc_print_report_wizard()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			Initialize();
			
			// TODO: Add any initialization after the InitializeComponent call

		}
		public uc_print_report_wizard(System.Data.DataTable p_dt)
		{
			this.m_strType="t";
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			Initialize();
			this.m_dt = new DataTable();
			this.m_dt = p_dt;

            //create the detail options table
			this.m_dtDetailOptions = new DataTable("detail_summary_user_options");
			DataColumn dtCol = new DataColumn("field_name",Type.GetType("System.String"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("sum",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("avg_sum",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("count",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("pct_count",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("min",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("max",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);


			//define the detail option dataset and table


            this.lblTableName.Text = this.m_dt.TableName;
			this.txtReportTitle.Text = this.m_dt.TableName;
			this.lstAvailableFields.Items.Clear();
			for (int x=0; x <= this.m_dt.Columns.Count -1; x++)
			{
				this.lstAvailableFields.Items.Add(this.m_dt.Columns[x].ColumnName);
			}
			this.lstAvailableFields.SelectedIndex = 0;
			
			
		// TODO: Add any initialization after the InitializeComponent call

	}
		public uc_print_report_wizard(System.Windows.Forms.DataGrid p_dg,string strName)
		{
            if (p_dg.VisibleRowCount == 0) return;
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			Initialize();
			this.m_strType="g";
			this.m_dg = new DataGrid();
			this.m_dg = p_dg;
			this.m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;

		
			//create the detail options table
			this.m_dtDetailOptions = new DataTable("detail_summary_user_options");
			DataColumn dtCol = new DataColumn("field_name",Type.GetType("System.String"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("sum",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("avg_sum",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("count",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("pct_count",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("min",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);
			dtCol = new DataColumn("max",Type.GetType("System.Boolean"));
			this.m_dtDetailOptions.Columns.Add(dtCol);


			this.lblTableName.Text = strName;
			this.txtReportTitle.Text = strName;
			this.lstAvailableFields.Items.Clear();
			for (int x=0; x <= this.m_dg.TableStyles[0].GridColumnStyles.Count - 1; x++)
			{
              this.lstAvailableFields.Items.Add(this.m_dg.TableStyles[0].GridColumnStyles[x].HeaderText.Trim());
			}
			this.lstAvailableFields.SelectedIndex = 0;
			
			
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_print_report_wizard));
			this.grpboxMain = new System.Windows.Forms.GroupBox();
			this.grpboxFinish = new System.Windows.Forms.GroupBox();
			this.groupBox7 = new System.Windows.Forms.GroupBox();
			this.lblLandscape = new System.Windows.Forms.Label();
			this.lblPortrait = new System.Windows.Forms.Label();
			this.rdoLandscape = new System.Windows.Forms.RadioButton();
			this.rdoPortrait = new System.Windows.Forms.RadioButton();
			this.grpboxRecordLayout = new System.Windows.Forms.GroupBox();
			this.picRecordLayout = new System.Windows.Forms.PictureBox();
			this.rdoRow = new System.Windows.Forms.RadioButton();
			this.rdoColumn = new System.Windows.Forms.RadioButton();
			this.btnPreview = new System.Windows.Forms.Button();
			this.txtReportTitle = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.grpboxDetailAndSummary = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.chkPrintReportTotals = new System.Windows.Forms.CheckBox();
			this.chkPrintGroupSummary = new System.Windows.Forms.CheckBox();
			this.chkPrintDetail = new System.Windows.Forms.CheckBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.dgDetailOptions = new System.Windows.Forms.DataGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkPctGroupRecords = new System.Windows.Forms.CheckBox();
			this.chkCountGroupRecords = new System.Windows.Forms.CheckBox();
			this.grpboxSort = new System.Windows.Forms.GroupBox();
			this.grpboxSort3 = new System.Windows.Forms.GroupBox();
			this.rdoSort3Desc = new System.Windows.Forms.RadioButton();
			this.rdoSort3Asc = new System.Windows.Forms.RadioButton();
			this.cmbSort3 = new System.Windows.Forms.ComboBox();
			this.grpboxSort1 = new System.Windows.Forms.GroupBox();
			this.rdoSort1Desc = new System.Windows.Forms.RadioButton();
			this.rdoSort1Asc = new System.Windows.Forms.RadioButton();
			this.cmbSort1 = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.grpboxSort2 = new System.Windows.Forms.GroupBox();
			this.rdoSort2Desc = new System.Windows.Forms.RadioButton();
			this.rdoSort2Asc = new System.Windows.Forms.RadioButton();
			this.cmbSort2 = new System.Windows.Forms.ComboBox();
			this.grpboxGroupRecords = new System.Windows.Forms.GroupBox();
			this.groupBox8 = new System.Windows.Forms.GroupBox();
			this.radioButton3 = new System.Windows.Forms.RadioButton();
			this.radioButton4 = new System.Windows.Forms.RadioButton();
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.label3 = new System.Windows.Forms.Label();
			this.grpboxGroup3 = new System.Windows.Forms.GroupBox();
			this.rdoGroup3Desc = new System.Windows.Forms.RadioButton();
			this.rdoGroup3Asc = new System.Windows.Forms.RadioButton();
			this.cmbGroup3 = new System.Windows.Forms.ComboBox();
			this.grpboxGroup2 = new System.Windows.Forms.GroupBox();
			this.rdoGroup2Desc = new System.Windows.Forms.RadioButton();
			this.rdoGroup2Asc = new System.Windows.Forms.RadioButton();
			this.cmbGroup2 = new System.Windows.Forms.ComboBox();
			this.grpboxGroup1 = new System.Windows.Forms.GroupBox();
			this.rdoGroup1Desc = new System.Windows.Forms.RadioButton();
			this.rdoGroup1Asc = new System.Windows.Forms.RadioButton();
			this.cmbGroup1 = new System.Windows.Forms.ComboBox();
			this.grpboxSelectFields = new System.Windows.Forms.GroupBox();
			this.btnDown = new System.Windows.Forms.Button();
			this.btnUp = new System.Windows.Forms.Button();
			this.btnBottom = new System.Windows.Forms.Button();
			this.btnTop = new System.Windows.Forms.Button();
			this.btnPrevious = new System.Windows.Forms.Button();
			this.btnAddAll = new System.Windows.Forms.Button();
			this.cmbSteps = new System.Windows.Forms.ComboBox();
			this.btnFinish = new System.Windows.Forms.Button();
			this.lstReportFields = new System.Windows.Forms.ListBox();
			this.btnRemoveAll = new System.Windows.Forms.Button();
			this.btnRemove = new System.Windows.Forms.Button();
			this.btnAdd = new System.Windows.Forms.Button();
			this.lstAvailableFields = new System.Windows.Forms.ListBox();
			this.btnGetTable = new System.Windows.Forms.Button();
			this.lblTableName = new System.Windows.Forms.Label();
			this.btnNext = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnHelp = new System.Windows.Forms.Button();
			this.grpboxMain.SuspendLayout();
			this.grpboxFinish.SuspendLayout();
			this.groupBox7.SuspendLayout();
			this.grpboxRecordLayout.SuspendLayout();
			this.grpboxDetailAndSummary.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dgDetailOptions)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.grpboxSort.SuspendLayout();
			this.grpboxSort3.SuspendLayout();
			this.grpboxSort1.SuspendLayout();
			this.grpboxSort2.SuspendLayout();
			this.grpboxGroupRecords.SuspendLayout();
			this.groupBox8.SuspendLayout();
			this.grpboxGroup3.SuspendLayout();
			this.grpboxGroup2.SuspendLayout();
			this.grpboxGroup1.SuspendLayout();
			this.grpboxSelectFields.SuspendLayout();
			this.SuspendLayout();
			// 
			// grpboxMain
			// 
			this.grpboxMain.Controls.Add(this.grpboxFinish);
			this.grpboxMain.Controls.Add(this.grpboxDetailAndSummary);
			this.grpboxMain.Controls.Add(this.grpboxSort);
			this.grpboxMain.Controls.Add(this.grpboxGroupRecords);
			this.grpboxMain.Controls.Add(this.grpboxSelectFields);
			this.grpboxMain.Dock = System.Windows.Forms.DockStyle.Fill;
			this.grpboxMain.Location = new System.Drawing.Point(0, 0);
			this.grpboxMain.Name = "grpboxMain";
			this.grpboxMain.Size = new System.Drawing.Size(576, 1560);
			this.grpboxMain.TabIndex = 0;
			this.grpboxMain.TabStop = false;
			this.grpboxMain.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// grpboxFinish
			// 
			this.grpboxFinish.Controls.Add(this.groupBox7);
			this.grpboxFinish.Controls.Add(this.grpboxRecordLayout);
			this.grpboxFinish.Controls.Add(this.btnPreview);
			this.grpboxFinish.Controls.Add(this.txtReportTitle);
			this.grpboxFinish.Controls.Add(this.label2);
			this.grpboxFinish.Location = new System.Drawing.Point(44, 1252);
			this.grpboxFinish.Name = "grpboxFinish";
			this.grpboxFinish.Size = new System.Drawing.Size(488, 272);
			this.grpboxFinish.TabIndex = 31;
			this.grpboxFinish.TabStop = false;
			this.grpboxFinish.Text = "Step 6 - Finish";
			// 
			// groupBox7
			// 
			this.groupBox7.Controls.Add(this.lblLandscape);
			this.groupBox7.Controls.Add(this.lblPortrait);
			this.groupBox7.Controls.Add(this.rdoLandscape);
			this.groupBox7.Controls.Add(this.rdoPortrait);
			this.groupBox7.Location = new System.Drawing.Point(288, 112);
			this.groupBox7.Name = "groupBox7";
			this.groupBox7.Size = new System.Drawing.Size(152, 72);
			this.groupBox7.TabIndex = 4;
			this.groupBox7.TabStop = false;
			this.groupBox7.Text = "Orientation";
			// 
			// lblLandscape
			// 
			this.lblLandscape.BackColor = System.Drawing.Color.White;
			this.lblLandscape.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblLandscape.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblLandscape.Location = new System.Drawing.Point(107, 27);
			this.lblLandscape.Name = "lblLandscape";
			this.lblLandscape.Size = new System.Drawing.Size(32, 24);
			this.lblLandscape.TabIndex = 3;
			this.lblLandscape.Text = "A";
			this.lblLandscape.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblLandscape.Visible = false;
			// 
			// lblPortrait
			// 
			this.lblPortrait.BackColor = System.Drawing.Color.White;
			this.lblPortrait.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblPortrait.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPortrait.Location = new System.Drawing.Point(107, 22);
			this.lblPortrait.Name = "lblPortrait";
			this.lblPortrait.Size = new System.Drawing.Size(24, 32);
			this.lblPortrait.TabIndex = 2;
			this.lblPortrait.Text = "A";
			this.lblPortrait.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// rdoLandscape
			// 
			this.rdoLandscape.Location = new System.Drawing.Point(16, 46);
			this.rdoLandscape.Name = "rdoLandscape";
			this.rdoLandscape.Size = new System.Drawing.Size(80, 16);
			this.rdoLandscape.TabIndex = 1;
			this.rdoLandscape.Text = "Landscape";
			this.rdoLandscape.Click += new System.EventHandler(this.rdoLandscape_Click);
			// 
			// rdoPortrait
			// 
			this.rdoPortrait.Checked = true;
			this.rdoPortrait.Location = new System.Drawing.Point(16, 21);
			this.rdoPortrait.Name = "rdoPortrait";
			this.rdoPortrait.Size = new System.Drawing.Size(64, 16);
			this.rdoPortrait.TabIndex = 0;
			this.rdoPortrait.TabStop = true;
			this.rdoPortrait.Text = "Portrait";
			this.rdoPortrait.Click += new System.EventHandler(this.rdoPortrait_Click);
			// 
			// grpboxRecordLayout
			// 
			this.grpboxRecordLayout.Controls.Add(this.picRecordLayout);
			this.grpboxRecordLayout.Controls.Add(this.rdoRow);
			this.grpboxRecordLayout.Controls.Add(this.rdoColumn);
			this.grpboxRecordLayout.Location = new System.Drawing.Point(32, 104);
			this.grpboxRecordLayout.Name = "grpboxRecordLayout";
			this.grpboxRecordLayout.Size = new System.Drawing.Size(240, 88);
			this.grpboxRecordLayout.TabIndex = 3;
			this.grpboxRecordLayout.TabStop = false;
			this.grpboxRecordLayout.Text = "Record Layout";
			// 
			// picRecordLayout
			// 
			this.picRecordLayout.Image = ((System.Drawing.Image)(resources.GetObject("picRecordLayout.Image")));
			this.picRecordLayout.Location = new System.Drawing.Point(96, 16);
			this.picRecordLayout.Name = "picRecordLayout";
			this.picRecordLayout.Size = new System.Drawing.Size(128, 56);
			this.picRecordLayout.TabIndex = 6;
			this.picRecordLayout.TabStop = false;
			// 
			// rdoRow
			// 
			this.rdoRow.Location = new System.Drawing.Point(16, 48);
			this.rdoRow.Name = "rdoRow";
			this.rdoRow.Size = new System.Drawing.Size(64, 16);
			this.rdoRow.TabIndex = 5;
			this.rdoRow.Text = "Row";
			this.rdoRow.Click += new System.EventHandler(this.rdoRow_Click);
			// 
			// rdoColumn
			// 
			this.rdoColumn.Checked = true;
			this.rdoColumn.Location = new System.Drawing.Point(16, 24);
			this.rdoColumn.Name = "rdoColumn";
			this.rdoColumn.Size = new System.Drawing.Size(64, 16);
			this.rdoColumn.TabIndex = 4;
			this.rdoColumn.TabStop = true;
			this.rdoColumn.Text = "Columns";
			this.rdoColumn.Click += new System.EventHandler(this.rdoColumn_Click);
			// 
			// btnPreview
			// 
			this.btnPreview.Location = new System.Drawing.Point(384, 200);
			this.btnPreview.Name = "btnPreview";
			this.btnPreview.Size = new System.Drawing.Size(88, 32);
			this.btnPreview.TabIndex = 2;
			this.btnPreview.Text = "Preview";
			this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
			// 
			// txtReportTitle
			// 
			this.txtReportTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtReportTitle.Location = new System.Drawing.Point(32, 64);
			this.txtReportTitle.Name = "txtReportTitle";
			this.txtReportTitle.Size = new System.Drawing.Size(432, 29);
			this.txtReportTitle.TabIndex = 1;
			this.txtReportTitle.Text = "";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(32, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(104, 16);
			this.label2.TabIndex = 0;
			this.label2.Text = "Report Title";
			// 
			// grpboxDetailAndSummary
			// 
			this.grpboxDetailAndSummary.Controls.Add(this.groupBox3);
			this.grpboxDetailAndSummary.Controls.Add(this.groupBox2);
			this.grpboxDetailAndSummary.Controls.Add(this.groupBox1);
			this.grpboxDetailAndSummary.Location = new System.Drawing.Point(48, 912);
			this.grpboxDetailAndSummary.Name = "grpboxDetailAndSummary";
			this.grpboxDetailAndSummary.Size = new System.Drawing.Size(488, 304);
			this.grpboxDetailAndSummary.TabIndex = 30;
			this.grpboxDetailAndSummary.TabStop = false;
			this.grpboxDetailAndSummary.Text = "Step 4 - Detail And Summary";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.chkPrintReportTotals);
			this.groupBox3.Controls.Add(this.chkPrintGroupSummary);
			this.groupBox3.Controls.Add(this.chkPrintDetail);
			this.groupBox3.Location = new System.Drawing.Point(16, 231);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(464, 42);
			this.groupBox3.TabIndex = 8;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Print Options";
			// 
			// chkPrintReportTotals
			// 
			this.chkPrintReportTotals.Checked = true;
			this.chkPrintReportTotals.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPrintReportTotals.Location = new System.Drawing.Point(280, 16);
			this.chkPrintReportTotals.Name = "chkPrintReportTotals";
			this.chkPrintReportTotals.Size = new System.Drawing.Size(112, 16);
			this.chkPrintReportTotals.TabIndex = 2;
			this.chkPrintReportTotals.Text = "Report Summary";
			// 
			// chkPrintGroupSummary
			// 
			this.chkPrintGroupSummary.Checked = true;
			this.chkPrintGroupSummary.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPrintGroupSummary.Location = new System.Drawing.Point(168, 16);
			this.chkPrintGroupSummary.Name = "chkPrintGroupSummary";
			this.chkPrintGroupSummary.Size = new System.Drawing.Size(112, 16);
			this.chkPrintGroupSummary.TabIndex = 1;
			this.chkPrintGroupSummary.Text = "Group Summary";
			// 
			// chkPrintDetail
			// 
			this.chkPrintDetail.Checked = true;
			this.chkPrintDetail.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPrintDetail.Location = new System.Drawing.Point(90, 16);
			this.chkPrintDetail.Name = "chkPrintDetail";
			this.chkPrintDetail.Size = new System.Drawing.Size(56, 16);
			this.chkPrintDetail.TabIndex = 0;
			this.chkPrintDetail.Text = "Detail";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.dgDetailOptions);
			this.groupBox2.Location = new System.Drawing.Point(8, 73);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(472, 153);
			this.groupBox2.TabIndex = 7;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Detail Options";
			// 
			// dgDetailOptions
			// 
			this.dgDetailOptions.AllowSorting = false;
			this.dgDetailOptions.CaptionVisible = false;
			this.dgDetailOptions.DataMember = "";
			this.dgDetailOptions.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dgDetailOptions.Location = new System.Drawing.Point(16, 24);
			this.dgDetailOptions.Name = "dgDetailOptions";
			this.dgDetailOptions.PreferredColumnWidth = 50;
			this.dgDetailOptions.Size = new System.Drawing.Size(432, 120);
			this.dgDetailOptions.TabIndex = 14;
			this.dgDetailOptions.Click += new System.EventHandler(this.dgDetailOptions_Click);
			this.dgDetailOptions.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dgDetailOptions_MouseUp);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkPctGroupRecords);
			this.groupBox1.Controls.Add(this.chkCountGroupRecords);
			this.groupBox1.Location = new System.Drawing.Point(8, 35);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(472, 32);
			this.groupBox1.TabIndex = 6;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Group Options";
			// 
			// chkPctGroupRecords
			// 
			this.chkPctGroupRecords.Location = new System.Drawing.Point(304, 14);
			this.chkPctGroupRecords.Name = "chkPctGroupRecords";
			this.chkPctGroupRecords.Size = new System.Drawing.Size(120, 16);
			this.chkPctGroupRecords.TabIndex = 7;
			this.chkPctGroupRecords.Text = "Count Percentage";
			// 
			// chkCountGroupRecords
			// 
			this.chkCountGroupRecords.Location = new System.Drawing.Point(71, 14);
			this.chkCountGroupRecords.Name = "chkCountGroupRecords";
			this.chkCountGroupRecords.Size = new System.Drawing.Size(220, 16);
			this.chkCountGroupRecords.TabIndex = 6;
			this.chkCountGroupRecords.Text = "Count the number of records in a group";
			// 
			// grpboxSort
			// 
			this.grpboxSort.Controls.Add(this.grpboxSort3);
			this.grpboxSort.Controls.Add(this.grpboxSort1);
			this.grpboxSort.Controls.Add(this.label1);
			this.grpboxSort.Controls.Add(this.grpboxSort2);
			this.grpboxSort.Location = new System.Drawing.Point(48, 632);
			this.grpboxSort.Name = "grpboxSort";
			this.grpboxSort.Size = new System.Drawing.Size(488, 272);
			this.grpboxSort.TabIndex = 29;
			this.grpboxSort.TabStop = false;
			this.grpboxSort.Text = "Step 3 - Sort Records";
			this.grpboxSort.Enter += new System.EventHandler(this.grpboxSort_Enter);
			// 
			// grpboxSort3
			// 
			this.grpboxSort3.Controls.Add(this.rdoSort3Desc);
			this.grpboxSort3.Controls.Add(this.rdoSort3Asc);
			this.grpboxSort3.Controls.Add(this.cmbSort3);
			this.grpboxSort3.Location = new System.Drawing.Point(40, 176);
			this.grpboxSort3.Name = "grpboxSort3";
			this.grpboxSort3.Size = new System.Drawing.Size(416, 40);
			this.grpboxSort3.TabIndex = 14;
			this.grpboxSort3.TabStop = false;
			this.grpboxSort3.Text = "Sort 3";
			// 
			// rdoSort3Desc
			// 
			this.rdoSort3Desc.Location = new System.Drawing.Point(360, 16);
			this.rdoSort3Desc.Name = "rdoSort3Desc";
			this.rdoSort3Desc.Size = new System.Drawing.Size(48, 16);
			this.rdoSort3Desc.TabIndex = 2;
			this.rdoSort3Desc.Text = "Desc";
			// 
			// rdoSort3Asc
			// 
			this.rdoSort3Asc.Checked = true;
			this.rdoSort3Asc.Location = new System.Drawing.Point(304, 16);
			this.rdoSort3Asc.Name = "rdoSort3Asc";
			this.rdoSort3Asc.Size = new System.Drawing.Size(48, 16);
			this.rdoSort3Asc.TabIndex = 1;
			this.rdoSort3Asc.TabStop = true;
			this.rdoSort3Asc.Text = "Asc";
			// 
			// cmbSort3
			// 
			this.cmbSort3.Location = new System.Drawing.Point(8, 13);
			this.cmbSort3.Name = "cmbSort3";
			this.cmbSort3.Size = new System.Drawing.Size(288, 21);
			this.cmbSort3.TabIndex = 0;
			this.cmbSort3.SelectedValueChanged += new System.EventHandler(this.cmbSort3_SelectedValueChanged);
			// 
			// grpboxSort1
			// 
			this.grpboxSort1.Controls.Add(this.rdoSort1Desc);
			this.grpboxSort1.Controls.Add(this.rdoSort1Asc);
			this.grpboxSort1.Controls.Add(this.cmbSort1);
			this.grpboxSort1.Location = new System.Drawing.Point(40, 54);
			this.grpboxSort1.Name = "grpboxSort1";
			this.grpboxSort1.Size = new System.Drawing.Size(416, 40);
			this.grpboxSort1.TabIndex = 13;
			this.grpboxSort1.TabStop = false;
			this.grpboxSort1.Text = "Sort 1";
			// 
			// rdoSort1Desc
			// 
			this.rdoSort1Desc.Location = new System.Drawing.Point(360, 16);
			this.rdoSort1Desc.Name = "rdoSort1Desc";
			this.rdoSort1Desc.Size = new System.Drawing.Size(48, 16);
			this.rdoSort1Desc.TabIndex = 2;
			this.rdoSort1Desc.Text = "Desc";
			// 
			// rdoSort1Asc
			// 
			this.rdoSort1Asc.Checked = true;
			this.rdoSort1Asc.Location = new System.Drawing.Point(304, 16);
			this.rdoSort1Asc.Name = "rdoSort1Asc";
			this.rdoSort1Asc.Size = new System.Drawing.Size(48, 16);
			this.rdoSort1Asc.TabIndex = 1;
			this.rdoSort1Asc.TabStop = true;
			this.rdoSort1Asc.Text = "Asc";
			// 
			// cmbSort1
			// 
			this.cmbSort1.Location = new System.Drawing.Point(8, 13);
			this.cmbSort1.Name = "cmbSort1";
			this.cmbSort1.Size = new System.Drawing.Size(288, 21);
			this.cmbSort1.TabIndex = 0;
			this.cmbSort1.SelectedValueChanged += new System.EventHandler(this.cmbSort1_SelectedValueChanged);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(48, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(144, 16);
			this.label1.TabIndex = 6;
			this.label1.Text = "Select up to 3 fields to sort";
			// 
			// grpboxSort2
			// 
			this.grpboxSort2.Controls.Add(this.rdoSort2Desc);
			this.grpboxSort2.Controls.Add(this.rdoSort2Asc);
			this.grpboxSort2.Controls.Add(this.cmbSort2);
			this.grpboxSort2.Location = new System.Drawing.Point(40, 115);
			this.grpboxSort2.Name = "grpboxSort2";
			this.grpboxSort2.Size = new System.Drawing.Size(416, 40);
			this.grpboxSort2.TabIndex = 13;
			this.grpboxSort2.TabStop = false;
			this.grpboxSort2.Text = "Sort 2";
			// 
			// rdoSort2Desc
			// 
			this.rdoSort2Desc.Location = new System.Drawing.Point(360, 16);
			this.rdoSort2Desc.Name = "rdoSort2Desc";
			this.rdoSort2Desc.Size = new System.Drawing.Size(48, 16);
			this.rdoSort2Desc.TabIndex = 2;
			this.rdoSort2Desc.Text = "Desc";
			// 
			// rdoSort2Asc
			// 
			this.rdoSort2Asc.Checked = true;
			this.rdoSort2Asc.Location = new System.Drawing.Point(304, 16);
			this.rdoSort2Asc.Name = "rdoSort2Asc";
			this.rdoSort2Asc.Size = new System.Drawing.Size(48, 16);
			this.rdoSort2Asc.TabIndex = 1;
			this.rdoSort2Asc.TabStop = true;
			this.rdoSort2Asc.Text = "Asc";
			// 
			// cmbSort2
			// 
			this.cmbSort2.Location = new System.Drawing.Point(8, 13);
			this.cmbSort2.Name = "cmbSort2";
			this.cmbSort2.Size = new System.Drawing.Size(288, 21);
			this.cmbSort2.TabIndex = 0;
			this.cmbSort2.SelectedValueChanged += new System.EventHandler(this.cmbSort2_SelectedValueChanged);
			// 
			// grpboxGroupRecords
			// 
			this.grpboxGroupRecords.Controls.Add(this.groupBox8);
			this.grpboxGroupRecords.Controls.Add(this.label3);
			this.grpboxGroupRecords.Controls.Add(this.grpboxGroup3);
			this.grpboxGroupRecords.Controls.Add(this.grpboxGroup2);
			this.grpboxGroupRecords.Controls.Add(this.grpboxGroup1);
			this.grpboxGroupRecords.Location = new System.Drawing.Point(44, 341);
			this.grpboxGroupRecords.Name = "grpboxGroupRecords";
			this.grpboxGroupRecords.Size = new System.Drawing.Size(488, 272);
			this.grpboxGroupRecords.TabIndex = 28;
			this.grpboxGroupRecords.TabStop = false;
			this.grpboxGroupRecords.Text = "Step 2 - Group Records";
			// 
			// groupBox8
			// 
			this.groupBox8.Controls.Add(this.radioButton3);
			this.groupBox8.Controls.Add(this.radioButton4);
			this.groupBox8.Controls.Add(this.comboBox1);
			this.groupBox8.Location = new System.Drawing.Point(36, 319);
			this.groupBox8.Name = "groupBox8";
			this.groupBox8.Size = new System.Drawing.Size(416, 40);
			this.groupBox8.TabIndex = 16;
			this.groupBox8.TabStop = false;
			this.groupBox8.Text = "Group 1";
			// 
			// radioButton3
			// 
			this.radioButton3.Location = new System.Drawing.Point(360, 16);
			this.radioButton3.Name = "radioButton3";
			this.radioButton3.Size = new System.Drawing.Size(48, 16);
			this.radioButton3.TabIndex = 2;
			this.radioButton3.Text = "Desc";
			// 
			// radioButton4
			// 
			this.radioButton4.Checked = true;
			this.radioButton4.Location = new System.Drawing.Point(304, 16);
			this.radioButton4.Name = "radioButton4";
			this.radioButton4.Size = new System.Drawing.Size(48, 16);
			this.radioButton4.TabIndex = 1;
			this.radioButton4.TabStop = true;
			this.radioButton4.Text = "Asc";
			// 
			// comboBox1
			// 
			this.comboBox1.Location = new System.Drawing.Point(8, 13);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(288, 21);
			this.comboBox1.TabIndex = 0;
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(48, 32);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(117, 16);
			this.label3.TabIndex = 15;
			this.label3.Text = "Select up to 3 groups";
			// 
			// grpboxGroup3
			// 
			this.grpboxGroup3.Controls.Add(this.rdoGroup3Desc);
			this.grpboxGroup3.Controls.Add(this.rdoGroup3Asc);
			this.grpboxGroup3.Controls.Add(this.cmbGroup3);
			this.grpboxGroup3.Location = new System.Drawing.Point(40, 176);
			this.grpboxGroup3.Name = "grpboxGroup3";
			this.grpboxGroup3.Size = new System.Drawing.Size(416, 40);
			this.grpboxGroup3.TabIndex = 14;
			this.grpboxGroup3.TabStop = false;
			this.grpboxGroup3.Text = "Group 3";
			// 
			// rdoGroup3Desc
			// 
			this.rdoGroup3Desc.Location = new System.Drawing.Point(360, 16);
			this.rdoGroup3Desc.Name = "rdoGroup3Desc";
			this.rdoGroup3Desc.Size = new System.Drawing.Size(48, 16);
			this.rdoGroup3Desc.TabIndex = 2;
			this.rdoGroup3Desc.Text = "Desc";
			// 
			// rdoGroup3Asc
			// 
			this.rdoGroup3Asc.Checked = true;
			this.rdoGroup3Asc.Location = new System.Drawing.Point(305, 16);
			this.rdoGroup3Asc.Name = "rdoGroup3Asc";
			this.rdoGroup3Asc.Size = new System.Drawing.Size(48, 16);
			this.rdoGroup3Asc.TabIndex = 1;
			this.rdoGroup3Asc.TabStop = true;
			this.rdoGroup3Asc.Text = "Asc";
			// 
			// cmbGroup3
			// 
			this.cmbGroup3.Location = new System.Drawing.Point(11, 12);
			this.cmbGroup3.Name = "cmbGroup3";
			this.cmbGroup3.Size = new System.Drawing.Size(285, 21);
			this.cmbGroup3.TabIndex = 0;
			this.cmbGroup3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbGroup3_KeyPress);
			this.cmbGroup3.SelectedValueChanged += new System.EventHandler(this.cmbGroup3_SelectedValueChanged);
			// 
			// grpboxGroup2
			// 
			this.grpboxGroup2.Controls.Add(this.rdoGroup2Desc);
			this.grpboxGroup2.Controls.Add(this.rdoGroup2Asc);
			this.grpboxGroup2.Controls.Add(this.cmbGroup2);
			this.grpboxGroup2.Location = new System.Drawing.Point(40, 115);
			this.grpboxGroup2.Name = "grpboxGroup2";
			this.grpboxGroup2.Size = new System.Drawing.Size(416, 40);
			this.grpboxGroup2.TabIndex = 13;
			this.grpboxGroup2.TabStop = false;
			this.grpboxGroup2.Text = "Group 2";
			// 
			// rdoGroup2Desc
			// 
			this.rdoGroup2Desc.Location = new System.Drawing.Point(360, 16);
			this.rdoGroup2Desc.Name = "rdoGroup2Desc";
			this.rdoGroup2Desc.Size = new System.Drawing.Size(48, 16);
			this.rdoGroup2Desc.TabIndex = 2;
			this.rdoGroup2Desc.Text = "Desc";
			// 
			// rdoGroup2Asc
			// 
			this.rdoGroup2Asc.Checked = true;
			this.rdoGroup2Asc.Location = new System.Drawing.Point(305, 16);
			this.rdoGroup2Asc.Name = "rdoGroup2Asc";
			this.rdoGroup2Asc.Size = new System.Drawing.Size(48, 16);
			this.rdoGroup2Asc.TabIndex = 1;
			this.rdoGroup2Asc.TabStop = true;
			this.rdoGroup2Asc.Text = "Asc";
			// 
			// cmbGroup2
			// 
			this.cmbGroup2.Location = new System.Drawing.Point(8, 13);
			this.cmbGroup2.Name = "cmbGroup2";
			this.cmbGroup2.Size = new System.Drawing.Size(288, 21);
			this.cmbGroup2.TabIndex = 0;
			this.cmbGroup2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbGroup2_KeyPress);
			this.cmbGroup2.SelectedValueChanged += new System.EventHandler(this.cmbGroup2_SelectedValueChanged);
			// 
			// grpboxGroup1
			// 
			this.grpboxGroup1.Controls.Add(this.rdoGroup1Desc);
			this.grpboxGroup1.Controls.Add(this.rdoGroup1Asc);
			this.grpboxGroup1.Controls.Add(this.cmbGroup1);
			this.grpboxGroup1.Location = new System.Drawing.Point(40, 54);
			this.grpboxGroup1.Name = "grpboxGroup1";
			this.grpboxGroup1.Size = new System.Drawing.Size(416, 40);
			this.grpboxGroup1.TabIndex = 12;
			this.grpboxGroup1.TabStop = false;
			this.grpboxGroup1.Text = "Group 1";
			// 
			// rdoGroup1Desc
			// 
			this.rdoGroup1Desc.Location = new System.Drawing.Point(360, 16);
			this.rdoGroup1Desc.Name = "rdoGroup1Desc";
			this.rdoGroup1Desc.Size = new System.Drawing.Size(48, 16);
			this.rdoGroup1Desc.TabIndex = 2;
			this.rdoGroup1Desc.Text = "Desc";
			// 
			// rdoGroup1Asc
			// 
			this.rdoGroup1Asc.Checked = true;
			this.rdoGroup1Asc.Location = new System.Drawing.Point(304, 16);
			this.rdoGroup1Asc.Name = "rdoGroup1Asc";
			this.rdoGroup1Asc.Size = new System.Drawing.Size(48, 16);
			this.rdoGroup1Asc.TabIndex = 1;
			this.rdoGroup1Asc.TabStop = true;
			this.rdoGroup1Asc.Text = "Asc";
			// 
			// cmbGroup1
			// 
			this.cmbGroup1.Location = new System.Drawing.Point(8, 13);
			this.cmbGroup1.Name = "cmbGroup1";
			this.cmbGroup1.Size = new System.Drawing.Size(288, 21);
			this.cmbGroup1.TabIndex = 0;
			this.cmbGroup1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbGroup1_KeyPress);
			this.cmbGroup1.SelectedValueChanged += new System.EventHandler(this.cmbGroup1_SelectedValueChanged);
			this.cmbGroup1.SelectedIndexChanged += new System.EventHandler(this.cmbGroup1_SelectedIndexChanged);
			// 
			// grpboxSelectFields
			// 
			this.grpboxSelectFields.Controls.Add(this.btnHelp);
			this.grpboxSelectFields.Controls.Add(this.btnDown);
			this.grpboxSelectFields.Controls.Add(this.btnUp);
			this.grpboxSelectFields.Controls.Add(this.btnBottom);
			this.grpboxSelectFields.Controls.Add(this.btnTop);
			this.grpboxSelectFields.Controls.Add(this.btnPrevious);
			this.grpboxSelectFields.Controls.Add(this.btnAddAll);
			this.grpboxSelectFields.Controls.Add(this.cmbSteps);
			this.grpboxSelectFields.Controls.Add(this.btnFinish);
			this.grpboxSelectFields.Controls.Add(this.lstReportFields);
			this.grpboxSelectFields.Controls.Add(this.btnRemoveAll);
			this.grpboxSelectFields.Controls.Add(this.btnRemove);
			this.grpboxSelectFields.Controls.Add(this.btnAdd);
			this.grpboxSelectFields.Controls.Add(this.lstAvailableFields);
			this.grpboxSelectFields.Controls.Add(this.btnGetTable);
			this.grpboxSelectFields.Controls.Add(this.lblTableName);
			this.grpboxSelectFields.Controls.Add(this.btnNext);
			this.grpboxSelectFields.Controls.Add(this.btnCancel);
			this.grpboxSelectFields.Location = new System.Drawing.Point(40, 24);
			this.grpboxSelectFields.Name = "grpboxSelectFields";
			this.grpboxSelectFields.Size = new System.Drawing.Size(488, 304);
			this.grpboxSelectFields.TabIndex = 27;
			this.grpboxSelectFields.TabStop = false;
			this.grpboxSelectFields.Text = "Step 1 - Select Fields";
			// 
			// btnDown
			// 
			this.btnDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnDown.Location = new System.Drawing.Point(337, 228);
			this.btnDown.Name = "btnDown";
			this.btnDown.Size = new System.Drawing.Size(48, 20);
			this.btnDown.TabIndex = 18;
			this.btnDown.Text = "Down";
			this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
			// 
			// btnUp
			// 
			this.btnUp.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnUp.Location = new System.Drawing.Point(289, 228);
			this.btnUp.Name = "btnUp";
			this.btnUp.Size = new System.Drawing.Size(48, 20);
			this.btnUp.TabIndex = 17;
			this.btnUp.Text = "Up";
			this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
			// 
			// btnBottom
			// 
			this.btnBottom.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnBottom.Location = new System.Drawing.Point(337, 208);
			this.btnBottom.Name = "btnBottom";
			this.btnBottom.Size = new System.Drawing.Size(48, 20);
			this.btnBottom.TabIndex = 16;
			this.btnBottom.Text = "Bottom";
			this.btnBottom.Click += new System.EventHandler(this.btnBottom_Click);
			// 
			// btnTop
			// 
			this.btnTop.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnTop.Location = new System.Drawing.Point(289, 208);
			this.btnTop.Name = "btnTop";
			this.btnTop.Size = new System.Drawing.Size(48, 20);
			this.btnTop.TabIndex = 15;
			this.btnTop.Text = "Top";
			this.btnTop.Click += new System.EventHandler(this.btnTop_Click);
			// 
			// btnPrevious
			// 
			this.btnPrevious.Location = new System.Drawing.Point(227, 271);
			this.btnPrevious.Name = "btnPrevious";
			this.btnPrevious.Size = new System.Drawing.Size(72, 24);
			this.btnPrevious.TabIndex = 14;
			this.btnPrevious.Text = "< Previoius";
			this.btnPrevious.Click += new System.EventHandler(this.btnPrev_Click);
			// 
			// btnAddAll
			// 
			this.btnAddAll.Image = ((System.Drawing.Image)(resources.GetObject("btnAddAll.Image")));
			this.btnAddAll.Location = new System.Drawing.Point(219, 98);
			this.btnAddAll.Name = "btnAddAll";
			this.btnAddAll.Size = new System.Drawing.Size(32, 32);
			this.btnAddAll.TabIndex = 13;
			this.btnAddAll.Click += new System.EventHandler(this.btnAddAll_Click);
			// 
			// cmbSteps
			// 
			this.cmbSteps.Enabled = false;
			this.cmbSteps.Items.AddRange(new object[] {
														  "Step 1 - Select Fields",
														  "Step 2 - Group Records",
														  "Step 3 - Sort Records",
														  "Step 4 - Detail And Summary",
														  "Step 6 - Finish"});
			this.cmbSteps.Location = new System.Drawing.Point(256, 24);
			this.cmbSteps.Name = "cmbSteps";
			this.cmbSteps.Size = new System.Drawing.Size(208, 21);
			this.cmbSteps.TabIndex = 12;
			this.cmbSteps.Text = "Step 1 - Select Fields";
			this.cmbSteps.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSteps_KeyPress);
			this.cmbSteps.Click += new System.EventHandler(this.btnCombo_Click);
			this.cmbSteps.SelectedIndexChanged += new System.EventHandler(this.cmbSteps_SelectedIndexChanged);
			// 
			// btnFinish
			// 
			this.btnFinish.Location = new System.Drawing.Point(408, 271);
			this.btnFinish.Name = "btnFinish";
			this.btnFinish.Size = new System.Drawing.Size(72, 24);
			this.btnFinish.TabIndex = 11;
			this.btnFinish.Text = "Finish";
			this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
			// 
			// lstReportFields
			// 
			this.lstReportFields.Location = new System.Drawing.Point(264, 59);
			this.lstReportFields.Name = "lstReportFields";
			this.lstReportFields.Size = new System.Drawing.Size(152, 147);
			this.lstReportFields.TabIndex = 10;
			// 
			// btnRemoveAll
			// 
			this.btnRemoveAll.Image = ((System.Drawing.Image)(resources.GetObject("btnRemoveAll.Image")));
			this.btnRemoveAll.Location = new System.Drawing.Point(219, 161);
			this.btnRemoveAll.Name = "btnRemoveAll";
			this.btnRemoveAll.Size = new System.Drawing.Size(32, 32);
			this.btnRemoveAll.TabIndex = 9;
			this.btnRemoveAll.Click += new System.EventHandler(this.btnRemoveAll_Click);
			// 
			// btnRemove
			// 
			this.btnRemove.Image = ((System.Drawing.Image)(resources.GetObject("btnRemove.Image")));
			this.btnRemove.Location = new System.Drawing.Point(219, 130);
			this.btnRemove.Name = "btnRemove";
			this.btnRemove.Size = new System.Drawing.Size(32, 32);
			this.btnRemove.TabIndex = 8;
			this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
			// 
			// btnAdd
			// 
			this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
			this.btnAdd.Location = new System.Drawing.Point(219, 67);
			this.btnAdd.Name = "btnAdd";
			this.btnAdd.Size = new System.Drawing.Size(32, 32);
			this.btnAdd.TabIndex = 6;
			this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
			// 
			// lstAvailableFields
			// 
			this.lstAvailableFields.Location = new System.Drawing.Point(56, 59);
			this.lstAvailableFields.Name = "lstAvailableFields";
			this.lstAvailableFields.Size = new System.Drawing.Size(152, 147);
			this.lstAvailableFields.TabIndex = 5;
			// 
			// btnGetTable
			// 
			this.btnGetTable.Image = ((System.Drawing.Image)(resources.GetObject("btnGetTable.Image")));
			this.btnGetTable.Location = new System.Drawing.Point(202, 21);
			this.btnGetTable.Name = "btnGetTable";
			this.btnGetTable.Size = new System.Drawing.Size(33, 32);
			this.btnGetTable.TabIndex = 4;
			// 
			// lblTableName
			// 
			this.lblTableName.BackColor = System.Drawing.Color.White;
			this.lblTableName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblTableName.Location = new System.Drawing.Point(8, 24);
			this.lblTableName.Name = "lblTableName";
			this.lblTableName.Size = new System.Drawing.Size(184, 24);
			this.lblTableName.TabIndex = 3;
			// 
			// btnNext
			// 
			this.btnNext.Location = new System.Drawing.Point(299, 271);
			this.btnNext.Name = "btnNext";
			this.btnNext.Size = new System.Drawing.Size(72, 24);
			this.btnNext.TabIndex = 2;
			this.btnNext.Text = "Next >";
			this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(147, 271);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 24);
			this.btnCancel.TabIndex = 0;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// imageList1
			// 
			this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit;
			this.imageList1.ImageSize = new System.Drawing.Size(128, 56);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnHelp
			// 
			this.btnHelp.Location = new System.Drawing.Point(10, 271);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(64, 24);
			this.btnHelp.TabIndex = 19;
			this.btnHelp.Text = "Help";
			// 
			// uc_print_report_wizard
			// 
			this.Controls.Add(this.grpboxMain);
			this.Name = "uc_print_report_wizard";
			this.Size = new System.Drawing.Size(576, 1560);
			this.grpboxMain.ResumeLayout(false);
			this.grpboxFinish.ResumeLayout(false);
			this.groupBox7.ResumeLayout(false);
			this.grpboxRecordLayout.ResumeLayout(false);
			this.grpboxDetailAndSummary.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dgDetailOptions)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.grpboxSort.ResumeLayout(false);
			this.grpboxSort3.ResumeLayout(false);
			this.grpboxSort1.ResumeLayout(false);
			this.grpboxSort2.ResumeLayout(false);
			this.grpboxGroupRecords.ResumeLayout(false);
			this.groupBox8.ResumeLayout(false);
			this.grpboxGroup3.ResumeLayout(false);
			this.grpboxGroup2.ResumeLayout(false);
			this.grpboxGroup1.ResumeLayout(false);
			this.grpboxSelectFields.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void lstAvailableFieldsSort_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}
	private void Initialize()
		{
			this.m_DialogWd = this.Width + 10;
			this.m_DialogHt = this.grpboxMain.Top + this.btnNext.Top + this.btnNext.Height + 100;
					
			this.grpboxGroupRecords.Left = this.grpboxSelectFields.Left;
		    this.grpboxGroupRecords.Width = this.grpboxSelectFields.Width;
		    this.grpboxGroupRecords.Height = this.grpboxSelectFields.Height;
     		this.grpboxGroupRecords.Top = this.grpboxSelectFields.Top;
		    this.grpboxGroupRecords.Visible=false;		

		    this.grpboxSort.Left = this.grpboxSelectFields.Left;
		    this.grpboxSort.Width = this.grpboxSelectFields.Width;
		    this.grpboxSort.Height = this.grpboxSelectFields.Height;
		    this.grpboxSort.Top = this.grpboxSelectFields.Top;
		    this.grpboxSort.Visible=false;

		    this.grpboxDetailAndSummary.Left = this.grpboxSelectFields.Left;
		    this.grpboxDetailAndSummary.Top = this.grpboxSelectFields.Top;
		    this.grpboxDetailAndSummary.Width = this.grpboxSelectFields.Width;
		    this.grpboxDetailAndSummary.Height = this.grpboxSelectFields.Height;
		    this.grpboxDetailAndSummary.Visible=false;
			
			this.grpboxFinish.Left = this.grpboxSelectFields.Left;
			this.grpboxFinish.Top = this.grpboxSelectFields.Top;
			this.grpboxFinish.Width = this.grpboxSelectFields.Width;
		    this.grpboxFinish.Height = this.grpboxSelectFields.Height;
			this.grpboxFinish.Visible=false;

		    this.cmbSort1.Left = this.cmbGroup1.Left;
		    this.cmbSort1.Width = this.cmbGroup1.Width;
		    this.cmbSort1.Top = this.cmbGroup1.Top;
		    this.rdoSort1Asc.Left = this.rdoGroup1Asc.Left;
		    this.rdoSort1Asc.Top = this.rdoGroup1Asc.Top;
		    this.rdoSort1Desc.Left = this.rdoGroup1Desc.Left;
		    this.rdoSort1Desc.Top = this.rdoGroup1Desc.Top;

    		this.cmbSort2.Left = this.cmbGroup2.Left;
	    	this.cmbSort2.Width = this.cmbGroup2.Width;
		    this.cmbSort2.Top = this.cmbGroup2.Top;
		    this.rdoSort2Asc.Left = this.rdoGroup2Asc.Left;
		    this.rdoSort2Asc.Top = this.rdoGroup2Asc.Top;
		    this.rdoSort2Desc.Left = this.rdoGroup2Desc.Left;
		    this.rdoSort2Desc.Top = this.rdoGroup2Desc.Top;

     		this.cmbSort3.Left = this.cmbGroup3.Left;
	    	this.cmbSort3.Width = this.cmbGroup3.Width;
		    this.cmbSort3.Top = this.cmbGroup3.Top;
		    this.rdoSort3Asc.Left = this.rdoGroup3Asc.Left;
		    this.rdoSort3Asc.Top = this.rdoGroup3Asc.Top;
		    this.rdoSort3Desc.Left = this.rdoGroup3Desc.Left;
		    this.rdoSort3Desc.Top = this.rdoGroup3Desc.Top;


			this.btnNext.Enabled=false;
			this.btnPrevious.Enabled=false;
			this.btnFinish.Enabled=false;
		    this.btnHelp.Enabled=true;

			this.m_strAppDir = 
				System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.ToString());
			int intAt = this.m_strAppDir.IndexOf("file:\\",0, this.m_strAppDir.Length);

			if (intAt >=0) 
			{
				int intLen = this.m_strAppDir.Length - 5;
				this.m_strAppDir = this.m_strAppDir.Substring(6);
			}
			this.btnNext.Name = "selectfields";
			this.btnPrevious.Name = "selectfields";
			this.btnCancel.Name="selectfields";
			this.btnFinish.Name = "selectfields";
			this.cmbSteps.Name = "selectfields";
		    this.btnHelp.Name = "selectfields";
			this.position_controls(ref this.grpboxSelectFields);

		    


		}

		private void btnAdd_Click(object sender, System.EventArgs e)
		{
			int x=0;
			int y=0;
			if (this.lstAvailableFields.Items.Count > 0) 
			{
				this.lstReportFields.Items.Add(this.lstAvailableFields.SelectedItems[0].ToString());
				this.lstReportFields.SelectedIndex = this.lstReportFields.Items.Count-1;
			}
			if (this.lstAvailableFields.Items.Count > 1) 
			{
				System.Windows.Forms.ListBox p_lb = new ListBox();
                    			
				for (x = 0; x <= this.lstAvailableFields.Items.Count-1;x++)
				{
					if (x != this.lstAvailableFields.SelectedIndex)
					    p_lb.Items.Add(this.lstAvailableFields.Items[x].ToString());
				}
			
			    y = this.lstAvailableFields.SelectedIndex;
                this.lstAvailableFields.Items.Clear();
				for (x = 0; x <= p_lb.Items.Count-1;x++)
				{
					this.lstAvailableFields.Items.Add(p_lb.Items[x]);
				}

				//highlight new item ;
				if (this.lstAvailableFields.Items.Count == y)
				{
					this.lstAvailableFields.SelectedIndex = y - 1;
				}
				else 
				{
                  this.lstAvailableFields.SelectedIndex = y;
				}

			

			}
			else this.lstAvailableFields.Items.Clear();

			if (this.btnNext.Enabled==false)
			{
				this.btnNext.Enabled=true;
				this.btnFinish.Enabled=true;
				this.cmbSteps.Enabled=true;
			}

		}

		private void btnAddAll_Click(object sender, System.EventArgs e)
		{
			int x =0;
			this.lstReportFields.Items.Clear();
			this.lstAvailableFields.Items.Clear();
			switch (this.m_strType)
			{
				case "t":    //table
					for (x=0;x<=m_dt.Columns.Count-1;x++)
					{
						this.lstReportFields.Items.Add(this.m_dt.Columns[x].ColumnName);

					}
					break;
				case "g":   //grid

					for (x=0; x <= this.m_dg.TableStyles[0].GridColumnStyles.Count - 1; x++)
					{
						this.lstReportFields.Items.Add(this.m_dg.TableStyles[0].GridColumnStyles[x].HeaderText.Trim());
					}

					break;
			}
			if (this.btnNext.Enabled==false)
			{
				this.btnNext.Enabled=true;
				this.btnFinish.Enabled=true;
				this.cmbSteps.Enabled=true;
			}
			this.lstReportFields.SelectedIndex = this.lstReportFields.Items.Count-1;
		}

		private void btnRemove_Click(object sender, System.EventArgs e)
		{
			int x=0;
			int y=0;
			if (this.lstReportFields.Items.Count > 0) 
			{
				this.lstAvailableFields.Items.Add(this.lstReportFields.SelectedItems[0].ToString());
				this.lstAvailableFields.SelectedIndex = this.lstAvailableFields.Items.Count-1;
			}

			if (this.lstReportFields.Items.Count > 1) 
			{
				System.Windows.Forms.ListBox p_lb = new ListBox();
                    			
				for (x = 0; x <= this.lstReportFields.Items.Count-1;x++)
				{
					if (x != this.lstReportFields.SelectedIndex)
						p_lb.Items.Add(this.lstReportFields.Items[x].ToString());
				}
			
				y = this.lstReportFields.SelectedIndex;
				this.lstReportFields.Items.Clear();
				for (x = 0; x <= p_lb.Items.Count-1;x++)
				{
					this.lstReportFields.Items.Add(p_lb.Items[x]);
				}


				//highlight new item ;
				if (this.lstReportFields.Items.Count == y)
				{
					this.lstReportFields.SelectedIndex = y - 1;
				}
				else 
				{
					this.lstReportFields.SelectedIndex = y;
				}

				

			}
			else 
			{
				this.lstReportFields.Items.Clear();
				
					this.btnNext.Enabled=false;
					this.btnFinish.Enabled=false;
					this.cmbSteps.Enabled=false;
				

			}

		}

		private void btnRemoveAll_Click(object sender, System.EventArgs e)
		{
			int x =0;
			this.lstAvailableFields.Items.Clear();
			this.lstReportFields.Items.Clear();
			if (this.m_strType=="t")
			{
				for (x=0;x<=m_dt.Columns.Count-1;x++)
				{
					this.lstAvailableFields.Items.Add(this.m_dt.Columns[x].ColumnName);
				}
			}
			else
			{
				for (x=0; x <= this.m_dg.TableStyles[0].GridColumnStyles.Count - 1; x++)
				{
					this.lstAvailableFields.Items.Add(this.m_dg.TableStyles[0].GridColumnStyles[x].HeaderText.Trim());
				}

			}

				this.btnNext.Enabled=false;
				this.btnFinish.Enabled=false;
				this.cmbSteps.Enabled=false;
				this.lstAvailableFields.SelectedIndex = this.lstAvailableFields.Items.Count-1;
			
		
			}

		private void btnTop_Click(object sender, System.EventArgs e)
		{
			int x;
			string strSelectedItem="";
			if (this.lstReportFields.Items.Count > 0 && this.lstReportFields.SelectedIndex != 0)
			{
				if (this.lstReportFields.SelectedIndex >= 0)
				{
					System.Windows.Forms.ListBox p_lb = new ListBox();
                    strSelectedItem = this.lstReportFields.SelectedItems[0].ToString();			
					for (x = 0; x <= this.lstReportFields.Items.Count-1;x++)
					{
						if (x != this.lstReportFields.SelectedIndex)
							p_lb.Items.Add(this.lstReportFields.Items[x].ToString());
					}
					this.lstReportFields.Items.Clear();
					this.lstReportFields.Items.Add(strSelectedItem);
					for (x=0; x<=p_lb.Items.Count-1;x++)
					{
						this.lstReportFields.Items.Add(p_lb.Items[x]);
					}
					this.lstReportFields.SelectedIndex=0;

				}
			}
		}

		private void btnBottom_Click(object sender, System.EventArgs e)
		{
			int x;
			string strSelectedItem="";
			if (this.lstReportFields.Items.Count > 0 && this.lstReportFields.SelectedIndex != this.lstReportFields.Items.Count - 1)
			{
				if (this.lstReportFields.SelectedIndex >= 0)
				{
					System.Windows.Forms.ListBox p_lb = new ListBox();
					strSelectedItem = this.lstReportFields.SelectedItems[0].ToString();			
					for (x = 0; x <= this.lstReportFields.Items.Count-1;x++)
					{
						if (x != this.lstReportFields.SelectedIndex)
							p_lb.Items.Add(this.lstReportFields.Items[x].ToString());
					}
					this.lstReportFields.Items.Clear();
					for (x=0; x<=p_lb.Items.Count-1;x++)
					{
						this.lstReportFields.Items.Add(p_lb.Items[x]);
					}
					this.lstReportFields.Items.Add(strSelectedItem);
					this.lstReportFields.SelectedIndex=this.lstReportFields.Items.Count-1;

				}
			}
		}

		private void btnUp_Click(object sender, System.EventArgs e)
		{
			int intSelectedIndex=0;
			string strSelectedItem="";
            string strSwapItem = "";
			if (this.lstReportFields.Items.Count > 0 && this.lstReportFields.SelectedIndex != 0)
			{
				if (this.lstReportFields.SelectedIndex >= 0)
				{
					
					strSelectedItem = this.lstReportFields.SelectedItems[0].ToString();			
					intSelectedIndex = this.lstReportFields.SelectedIndex;
                    strSwapItem = this.lstReportFields.Items[intSelectedIndex - 1].ToString();
					this.lstReportFields.Items[intSelectedIndex - 1] = strSelectedItem;
					this.lstReportFields.Items[intSelectedIndex] = strSwapItem;
					this.lstReportFields.SelectedIndex = intSelectedIndex - 1;

				}
			}
		}

		private void btnDown_Click(object sender, System.EventArgs e)
		{
			int intSelectedIndex=0;
			string strSelectedItem="";
			string strSwapItem = "";
			if (this.lstReportFields.Items.Count > 0 && this.lstReportFields.SelectedIndex != this.lstReportFields.Items.Count - 1)
			{
				if (this.lstReportFields.SelectedIndex >= 0)
				{
					
					strSelectedItem = this.lstReportFields.SelectedItems[0].ToString();			
					intSelectedIndex = this.lstReportFields.SelectedIndex;
					strSwapItem = this.lstReportFields.Items[intSelectedIndex + 1].ToString();
					this.lstReportFields.Items[intSelectedIndex + 1] = strSelectedItem;
					this.lstReportFields.Items[intSelectedIndex] = strSwapItem;
					this.lstReportFields.SelectedIndex = intSelectedIndex + 1;

				}
			}
		}

		private void btnFinish_Click(object sender, System.EventArgs e)
		{
			switch (this.btnFinish.Name)
			{
				case "selectfields":
					this.grpboxSelectFields.Visible=false;
					break;
                case "selectgroups":
					this.grpboxGroupRecords.Visible=false;
					break;
                case "selectsort":
					this.grpboxSort.Visible = false;
					break;
                case "selectdetail":
					this.grpboxDetailAndSummary.Visible=false;
					break;
				case "reporttitle":
					if (this.m_strAction=="print")
                         this.Print("N");
					break;
				default:
					break;
			}
            
			//if there is a group or sort field selected then
			//disable the record layout group box and 
			//default the printing to all columns on the same row
			if (this.cmbGroup1.Text.Trim().ToUpper() != "<NONE>" ||
				this.cmbSort1.Text.Trim().ToUpper() != "<NONE>")
			{
				this.rdoColumn.Checked=true;
				this.rdoRow.Checked=false;
				this.grpboxRecordLayout.Enabled=false;
			     
			
			}
			else
			{
				this.grpboxRecordLayout.Enabled=true;
			}
			this.HelpButton(ref this.grpboxFinish,ref this.btnHelp, "reporttitle");
			this.FinishButton(ref this.grpboxFinish,ref this.btnFinish, "reporttitle");
			this.NextButton(ref this.grpboxFinish,ref this.btnNext, "reporttitle");
			this.PrevButton(ref this.grpboxFinish,ref this.btnPrevious, "reporttitle");
			this.CancelButton(ref this.grpboxFinish,ref this.btnCancel,"reporttitle");
			this.StepCombo(ref this.grpboxFinish,ref this.cmbSteps,"reporttitle");             
			this.btnNext.Enabled=false;
			this.btnFinish.Enabled=true;
			this.btnPrevious.Enabled=true;
			this.btnCancel.Enabled=true;
			this.position_controls(ref this.grpboxFinish);
			if (this.m_strAction.Trim() != "print")
			{
				this.cmbSteps.Text = this.cmbSteps.Items[this.cmbSteps.Items.Count - 1].ToString();
				this.m_strAction="print";
			}
			this.grpboxFinish.Visible=true;

		
	    }

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnPreview_Click(object sender, System.EventArgs e)
		{
			this.Print("Y");
		
		}
		private void Print(string strPreviewYN)
		{
			int x=0;
			int y=0;
			string strGroupYN="N";
			bool bDetailOptions=false;
			this.m_dsXMLSchema = new DataSet();

			//load the table schema for saving user selections
			try
			{
			
				this.m_dsXMLSchema.ReadXmlSchema(this.m_strAppDir + "\\db\\print_report_wizard_user_select.xsd");
			}
			catch(Exception error)
			{
				MessageBox.Show(error.Message,"XML Schema Error");
			}
            

			
			System.Data.DataRow p_row;

			try
			{
				if (this.m_dvDetailOptions.Table.Rows.Count > 0)
				{
					bDetailOptions=true;
				}

			}
			catch
			{
			}

			//fields on the report
			//first get the group fields
			if (this.cmbGroup1.Text.Trim().ToUpper() != "<NONE>" && 
				this.cmbGroup1.Text.Trim().Length > 0)
			{
				strGroupYN="Y";
				p_row = this.m_dsXMLSchema.Tables["report_fields"].NewRow();
				p_row["field_name"] = this.cmbGroup1.Text;
				p_row["group"] = true;
				p_row["sort"] = false;
				if (this.rdoGroup1Asc.Checked==true) p_row["sortorder"] = "ASC";
				else p_row["sortorder"] = "DESC";
				p_row["largeststring"] = " ";
				p_row["sum"] = false;
				p_row["max"] = false;
				p_row["min"] = false;
				p_row["count"] = false;
				p_row["avg_sum"] = false; 
				p_row["pct_count"] = false; 
				p_row["datatype"] = " ";
				p_row["candomath"]=false;
				p_row["value"]=" ";
				this.m_dsXMLSchema.Tables["report_fields"].Rows.Add(p_row);
			}
			if (this.cmbGroup2.Text.Trim().ToUpper() != "<NONE>" &&
				this.cmbGroup2.Text.Trim().Length > 0)
			{
				p_row = this.m_dsXMLSchema.Tables["report_fields"].NewRow();
				p_row["field_name"] = this.cmbGroup2.Text;
				p_row["group"] = true;
				p_row["sort"] = false;
				if (this.rdoGroup2Asc.Checked==true) p_row["sortorder"] = "ASC";
				else p_row["sortorder"] = "DESC";
				p_row["largeststring"] = " ";
				p_row["sum"] = false;
				p_row["max"] = false;
				p_row["min"] = false;
				p_row["count"] = false;
				p_row["avg_sum"] = false; 
				p_row["pct_count"] = false; 
				p_row["datatype"] = " ";
				p_row["candomath"]=false;
				p_row["value"]=" ";
				this.m_dsXMLSchema.Tables["report_fields"].Rows.Add(p_row);
			}
			if (this.cmbGroup3.Text.Trim().ToUpper() != "<NONE>"  &&
				this.cmbGroup3.Text.Trim().Length > 0)
			{
				p_row = this.m_dsXMLSchema.Tables["report_fields"].NewRow();
				p_row["field_name"] = this.cmbGroup3.Text;
				p_row["group"]=true;
				p_row["sort"]=false;
				if (this.rdoGroup3Asc.Checked==true) p_row["sortorder"] = "ASC";
				else p_row["sortorder"] = "DESC";
				p_row["largeststring"] = " ";
				p_row["sum"] = false;
				p_row["max"] = false;
				p_row["min"] = false;
				p_row["count"] = false;
				p_row["avg_sum"] = false; 
				p_row["pct_count"] = false; 
				p_row["datatype"] = " ";
				p_row["candomath"]=false;
				p_row["value"]=" ";
				this.m_dsXMLSchema.Tables["report_fields"].Rows.Add(p_row);
			}

			//get the sort fields

			if (this.cmbSort1.Text.Trim().ToUpper() != "<NONE>" && 
				this.cmbSort1.Text.Trim().Length > 0)
			{
				p_row = this.m_dsXMLSchema.Tables["report_fields"].NewRow();
				p_row["field_name"] = this.cmbSort1.Text;
				p_row["group"] = false;
				p_row["sort"] = true;
				if (this.rdoSort1Asc.Checked==true) p_row["sortorder"] = "ASC";
				else p_row["sortorder"] = "DESC";
				p_row["largeststring"] = " ";
				p_row["sum"] = false;
				p_row["max"] = false;
				p_row["min"] = false;
				p_row["count"] = false;
				p_row["avg_sum"] = false; 
				p_row["pct_count"] = false; 
				p_row["datatype"] = " ";
				p_row["candomath"]=false;
				p_row["value"]=" ";
				this.m_dsXMLSchema.Tables["report_fields"].Rows.Add(p_row);
			}
			if (this.cmbSort2.Text.Trim().ToUpper() != "<NONE>" &&
				this.cmbSort2.Text.Trim().Length > 0)
			{
				p_row = this.m_dsXMLSchema.Tables["report_fields"].NewRow();
				p_row["field_name"] = this.cmbSort2.Text;
				p_row["group"] = false;
				p_row["sort"] = true;
				if (this.rdoSort2Asc.Checked==true) p_row["sortorder"] = "ASC";
				else p_row["sortorder"] = "DESC";
				p_row["largeststring"] = " ";
				p_row["sum"] = false;
				p_row["max"] = false;
				p_row["min"] = false;
				p_row["count"] = false;
				p_row["avg_sum"] = false; 
				p_row["pct_count"] = false; 
				p_row["datatype"] = " ";
				p_row["candomath"]=false;
				p_row["value"]=" ";
				this.m_dsXMLSchema.Tables["report_fields"].Rows.Add(p_row);
			}
			if (this.cmbSort3.Text.Trim().ToUpper() != "<NONE>"  &&
				this.cmbSort3.Text.Trim().Length > 0)
			{
				p_row = this.m_dsXMLSchema.Tables["report_fields"].NewRow();
				p_row["field_name"] = this.cmbSort3.Text;
				p_row["group"]=false;
				p_row["sort"]=true;
				if (this.rdoSort3Asc.Checked==true) p_row["sortorder"] = "ASC";
				else p_row["sortorder"] = "DESC";
				p_row["largeststring"] = " ";
				p_row["sum"] = false;
				p_row["max"] = false;
				p_row["min"] = false;
				p_row["count"] = false;
				p_row["avg_sum"] = false; 
				p_row["pct_count"] = false; 
				p_row["datatype"] = " ";
				p_row["candomath"]=false;
				p_row["value"]=" ";
				this.m_dsXMLSchema.Tables["report_fields"].Rows.Add(p_row);
			}


			//get all the other fields
			for (x=0; x<=this.lstReportFields.Items.Count-1;x++)
			{
				//first check to see if the listed field is already
				//in the table as a group or sorted field
				for (y=0; y<= this.m_dsXMLSchema.Tables["report_fields"].Rows.Count-1;y++)
				{
					if (this.lstReportFields.Items[x].ToString().Trim().ToUpper() ==
						this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["field_name"].ToString().Trim().ToUpper()) 
						break;
				}
				if (y<= this.m_dsXMLSchema.Tables["report_fields"].Rows.Count-1)
				{
				}
				else 
				{
					//not found so append field to table
					p_row = this.m_dsXMLSchema.Tables["report_fields"].NewRow();
					p_row["field_name"] = this.lstReportFields.Items[x].ToString().Trim();
					p_row["group"]=false;
					p_row["sort"]=false;
					p_row["largeststring"] = " ";
					p_row["sum"] = false;
					p_row["max"] = false;
					p_row["min"] = false;
					p_row["count"] = false;
					p_row["avg_sum"] = false; 
					p_row["pct_count"] = false; 
					p_row["sortorder"] = " ";
					p_row["datatype"] = " ";
					p_row["candomath"]=false;
					p_row["value"]=" ";
					this.m_dsXMLSchema.Tables["report_fields"].Rows.Add(p_row);
				}

			}
            
			//load detail and summary values 
			if (bDetailOptions==true)
			{
				for (x=0; x<=this.m_dvDetailOptions.Table.Rows.Count-1;x++)
				{
					//first check to see if the listed field is already
					//in the table as a group or sorted field
					for (y=0; y<= this.m_dsXMLSchema.Tables["report_fields"].Rows.Count-1;y++)
					{
						if (this.m_dvDetailOptions.Table.Rows[x]["field_name"].ToString().Trim().ToUpper() ==
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["field_name"].ToString().Trim().ToUpper()) 
							break;
					}
					if (y<= this.m_dsXMLSchema.Tables["report_fields"].Rows.Count-1)
					{
						if ((bool)this.m_dvDetailOptions.Table.Rows[x]["sum"]==true)
						{
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["sum"]=true;
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["candomath"] = true;
						}
						if ((bool)this.m_dvDetailOptions.Table.Rows[x]["avg_sum"]==true)
						{
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["avg_sum"]=true;
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["candomath"] = true;
						}
						if ((bool)this.m_dvDetailOptions.Table.Rows[x]["count"]==true)
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["count"]=true;
						if ((bool)this.m_dvDetailOptions.Table.Rows[x]["pct_count"]==true)
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["pct_count"]=true;
						if ((bool)this.m_dvDetailOptions.Table.Rows[x]["max"]==true)
						{
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["max"]=true;
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["candomath"] = true;
						}
						if ((bool)this.m_dvDetailOptions.Table.Rows[x]["min"]==true)
						{
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["min"]=true;
							this.m_dsXMLSchema.Tables["report_fields"].Rows[y]["candomath"] = true;
						}

					}

				}
			}


			//report title
			p_row = this.m_dsXMLSchema.Tables["report_title"].NewRow();
			p_row["title"] = this.txtReportTitle.Text;
			this.m_dsXMLSchema.Tables["report_title"].Rows.Add(p_row);

			//print preview
			p_row = this.m_dsXMLSchema.Tables["printpreview"].NewRow();
			p_row["printpreview_yn"] = strPreviewYN;
			this.m_dsXMLSchema.Tables["printpreview"].Rows.Add(p_row);

			//print record layout
			//display record in a row or column format
			p_row = this.m_dsXMLSchema.Tables["record_layout"].NewRow();
			if (this.rdoColumn.Checked==true)
				p_row["type"] = "C";
			else p_row["type"] = "R";
			//other options
			p_row["group_YN"] = strGroupYN;
			if (this.chkCountGroupRecords.Checked == true) p_row["count_group_records_YN"]="Y";
			else p_row["count_group_records_YN"]="N";
			if (this.chkPctGroupRecords.Checked == true) p_row["pct_count_group_records_YN"]="Y";
			else p_row["pct_count_group_records_YN"]="N";
			if (this.chkPrintDetail.Checked == true) p_row["print_detail_line_YN"] = "Y";
			else p_row["print_detail_line_YN"]="N";
			if (this.chkPrintGroupSummary.Checked == true) p_row["print_group_summary_YN"] = "Y";
			else p_row["print_group_summary_YN"]="N";
			if (this.chkPrintReportTotals.Checked==true) p_row["print_report_totals_YN"]="Y";
			else p_row["print_report_totals_YN"]="N";

			this.m_dsXMLSchema.Tables["record_layout"].Rows.Add(p_row);


			//portrait or landscape 
			p_row = this.m_dsXMLSchema.Tables["orientation"].NewRow();
			if (this.rdoPortrait.Checked==true)
				p_row["type"] = "P";
			else p_row["type"] = "L";
			this.m_dsXMLSchema.Tables["orientation"].Rows.Add(p_row);
			printing p_print = new printing();
			switch (this.m_strType)
			{
				case "t":   //print the data table
					p_print.PrintReport(this.m_dt, this.m_dsXMLSchema);
					break;
				case "g":   //print the data grid
					p_print.PrintReport(this.m_dg, this.m_dsXMLSchema);
					break;
			}

			//for (x=0;x<=this.m_dsXMLSchema.Tables["report_fields"].Rows.Count-1;x++)
			//{
			//	MessageBox.Show(this.m_dsXMLSchema.Tables["report_fields"].Rows[x]["field_name"].ToString());
			//}

		}
		protected void NextButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
		{
			p_oGb.Controls.Add(p_oButton);
			//p_oButton.Left = p_oGb.Width - p_oButton.Width - 5;
			//p_oButton.Top = p_oGb.Height - p_oButton.Height - 5;
			p_oButton.Name = strButtonName;	
		}
		protected void PrevButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
		{
			p_oGb.Controls.Add(p_oButton);
			//p_oButton.Left = p_oGb.Width - this.btnNext.Width - 5 - p_oButton.Width;
			//p_oButton.Top = this.btnNext.op;
			p_oButton.Name = strButtonName;	
		}
		
		protected void CancelButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
		{
			p_oGb.Controls.Add(p_oButton);
			//p_oButton.Left = p_oGb.Width - this.btnNext.Width - 5 - p_oButton.Width;
			//p_oButton.Top = this.btnNext.Top;
			p_oButton.Name = strButtonName;	
		}
		protected void FinishButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
		{
			p_oGb.Controls.Add(p_oButton);
			//p_oButton.Left = p_oGb.Width - p_oButton.Width - 5;
			//p_oButton.Top = p_oGb.Height - p_oButton.Height - 5;
			p_oButton.Name = strButtonName;	
		}
		protected void HelpButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
		{
			p_oGb.Controls.Add(p_oButton);
			//p_oButton.Left = p_oGb.Width - p_oButton.Width - 5;
			//p_oButton.Top = p_oGb.Height - p_oButton.Height - 5;
			p_oButton.Name = strButtonName;	
		}
		protected void StepCombo(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.ComboBox p_oCombo, string strComboName)
		{
			p_oGb.Controls.Add(p_oCombo);
			//p_oCombo.Left = p_oGb.Width - p_oCombo.Width - 5;
			//p_oCombo.Top = p_oGb.Height - p_oCombo.Height - 5;
			p_oCombo.Name = strComboName;	
		}
		private void btnNext_Click(object sender, System.EventArgs e)
		{
			//MessageBox.Show(this.btnNext.Name);
			switch (this.btnNext.Name)
			{
				case "selectfields":
					this.SelectGroups();
					break;
				case "selectgroups":
					this.SelectSort();
					break;
                case "selectsort":
					this.SelectDetail();
					break;
				case "selectdetail":
					this.btnFinish_Click(sender,e);
					break;
				case "hazard_next":
					break;
                
				default:
					break;
			}
		}
		private void btnPrev_Click(object sender, System.EventArgs e)
		{
			//MessageBox.Show(this.btnPrev.Name);
			switch (this.btnPrevious.Name)
			{
				case "selectgroups":
					this.SelectFields();
					break;
				case "selectsort":
					this.SelectGroups();
					break;
				case "selectdetail":
					this.SelectSort();
					break;
				case "reporttitle":
					this.SelectDetail();
					break;
				case "overalleffective_prev":
					break;
				default:
					break;
			}
		}
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}
		private void btnCombo_Click(object sender, System.EventArgs e)
		{
			//MessageBox.Show(this.btnPrev.Name);
			switch (this.btnPrevious.Name)
			{
				case "ti_ci_effective_prev":
					break;
				case "backslide_prev":
					break;
				case "hazard_prev":
					break;
				case "overalleffective_prev":
					break;
				default:
					break;
			}
		}
		protected void position_controls(ref System.Windows.Forms.GroupBox p_gb)
		{
			this.btnFinish.Top = p_gb.Height - this.btnFinish.Height - 5;
			this.btnNext.Top = this.btnFinish.Top;
			this.btnPrevious.Top = this.btnFinish.Top;
			this.btnCancel.Top = this.btnFinish.Top;
			this.btnHelp.Top = this.btnFinish.Top;
			this.btnHelp.Left =  5;
			this.btnFinish.Left = p_gb.Width - this.btnFinish.Width - 5;
			this.btnNext.Left = this.btnFinish.Left - this.btnFinish.Width - 10;
			this.btnPrevious.Left = this.btnNext.Left - this.btnNext.Width - 2;
			this.btnCancel.Left = this.btnPrevious.Left - this.btnPrevious.Width - 10;
			this.cmbSteps.Top = this.cmbSteps.Height - 5;
			this.cmbSteps.Left = p_gb.Width - this.cmbSteps.Width - 5;
		}

		private void cmbSteps_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.m_strAction=="suppress") return;
			if (this.cmbSteps.Text.IndexOf("Finish") > 0) this.btnFinish_Click(sender,e);
			else if (this.cmbSteps.Text.IndexOf("Select Fields") > 0) this.SelectFields();
			else if (this.cmbSteps.Text.IndexOf("Group Records") > 0) this.SelectGroups();
			else if (this.cmbSteps.Text.IndexOf("Sort") > 0) this.SelectSort();
			else if (this.cmbSteps.Text.IndexOf("Detail") > 0) this.SelectDetail();
		}

		private void rdoRow_Click(object sender, System.EventArgs e)
		{
			//if (this.rdoColumn.Checked==true && this.rdoRow.Checked==false)
			//{
			//	this.rdoColumn.Checked=false;
			//	this.rdoRow.Checked=true;
				this.picRecordLayout.Image = this.imageList1.Images[1];
			//}
		}

		private void rdoColumn_Click(object sender, System.EventArgs e)
		{
			//if (this.rdoColumn.Checked==false && this.rdoRow.Checked==true)
			//{
			//	this.rdoColumn.Checked=true;
			//	this.rdoRow.Checked=false;
				this.picRecordLayout.Image = this.imageList1.Images[0];
			//}

		}
		private void SelectFields()
		{
			switch (this.btnNext.Name)
			{
				case "reporttitle":
					this.grpboxFinish.Visible=false;
					break;
                case "selectgroups":
                    this.grpboxGroupRecords.Visible=false;
					break;
                case "selectsort":
					this.grpboxSort.Visible=false;
					break;
                case "selectdetail":
					this.grpboxDetailAndSummary.Visible=false;
					break;
			}
			this.HelpButton(ref this.grpboxSelectFields,ref this.btnHelp, "selectfields");
			this.FinishButton(ref this.grpboxSelectFields,ref this.btnFinish, "selectfields");
			this.NextButton(ref this.grpboxSelectFields,ref this.btnNext, "selectfields");
			this.PrevButton(ref this.grpboxSelectFields,ref this.btnPrevious, "selectfields");
			this.CancelButton(ref this.grpboxSelectFields,ref this.btnCancel,"selectfields");
			this.StepCombo(ref this.grpboxSelectFields,ref this.cmbSteps,"selectfields");             
			this.btnNext.Enabled=true;
			this.btnFinish.Enabled=true;
			this.btnPrevious.Enabled=false;
			this.btnCancel.Enabled=true;
			this.position_controls(ref this.grpboxSelectFields);
			this.cmbSteps.Text = this.cmbSteps.Items[0].ToString();
			this.grpboxSelectFields.Visible=true;

		}
		private void SelectGroups()
		{
			int x=0;
			//int y=0;
			string strGroup1="";
			string strGroup2="";
			string strGroup3="";
            bool bFound1 = false;
			bool bFound2 = false;
			bool bFound3 = false;
			bool bSort = false;
            
			switch (this.btnNext.Name)
			{
				case "reporttitle":
					this.grpboxFinish.Visible=false;
					break;
				case "selectgroups":
					this.grpboxGroupRecords.Visible=false;
					break;
                case "selectsort":
					this.grpboxSort.Visible=false;
					break;
				case "selectdetail":
					this.grpboxDetailAndSummary.Visible=false;
					break;
			}
            this.m_strAction="updategroups";
			//validate group1 
			if (this.cmbGroup1.Text.Trim().Length > 0 &&
				this.cmbGroup1.Text.Trim() != "<none>")
			{
				strGroup1 = this.cmbGroup1.Text;
                
				//make sure this group is not selected in a sort
				if (this.cmbSort1.Text.Trim().ToUpper() != "<NONE>")
				{
					if (strGroup1.Trim().ToUpper() == this.cmbSort1.Text.Trim().ToUpper())
					{
						bSort = true;
					}
				}
				if (bSort == false)
				{
					if (this.cmbSort2.Text.Trim().ToUpper() != "<NONE>")
					{
						if (strGroup1.Trim().ToUpper() == this.cmbSort2.Text.Trim().ToUpper())
						{
							bSort = true;
						}
					}
				}
				if (bSort == false)
				{
					if (this.cmbSort3.Text.Trim().ToUpper() != "<NONE>")
					{
						if (strGroup1.Trim().ToUpper() == this.cmbSort3.Text.Trim().ToUpper())
						{
							bSort = true;
						}
					}
				}

				if (bSort == false)
				{
					//make sure field is still a selected field
					for (x=0;x<=this.lstReportFields.Items.Count-1;x++)
					{
						if (this.lstReportFields.Items[x].ToString().Trim().ToUpper() ==
							this.cmbGroup1.Text.Trim().ToUpper()) break;
					}
					if (x<=this.lstReportFields.Items.Count-1)
					{
						//found the field
						bFound1=true;
					}
				}


				//validate group2	
				if (this.cmbGroup2.Text.Trim().Length > 0 &&
					this.cmbGroup2.Text.Trim() != "<none>")
				{
					strGroup2=this.cmbGroup2.Text;
                    bSort=false;
					//make sure this group is not selected in a sort
					if (this.cmbSort1.Text.Trim().ToUpper() != "<NONE>")
					{
						if (strGroup2.Trim().ToUpper() == this.cmbSort1.Text.Trim().ToUpper())
						{
							bSort = true;
						}
					}
					if (bSort == false)
					{
						if (this.cmbSort2.Text.Trim().ToUpper() != "<NONE>")
						{
							if (strGroup2.Trim().ToUpper() == this.cmbSort2.Text.Trim().ToUpper())
							{
								bSort = true;
							}
						}
					}
					if (bSort == false)
					{
						if (this.cmbSort3.Text.Trim().ToUpper() != "<NONE>")
						{
							if (strGroup2.Trim().ToUpper() == this.cmbSort3.Text.Trim().ToUpper())
							{
								bSort = true;
							}
						}
					}

					if (bSort == false)
					{
						//make sure field is still a selected field
						for (x=0;x<=this.lstReportFields.Items.Count-1;x++)
						{
							if (this.lstReportFields.Items[x].ToString().Trim().ToUpper() ==
								this.cmbGroup2.Text.Trim().ToUpper()) break;
						}
						if (x<=this.lstReportFields.Items.Count-1)
						{
							//found the field
							bFound2=true;
						}
					}
					//validate group 3
					if (this.cmbGroup3.Text.Trim().Length > 0 &&
						this.cmbGroup3.Text.Trim() != "<none>")
					{
						strGroup3 = this.cmbGroup3.Text;
						bSort = false;
                
						//make sure this group is not selected in a sort
						if (this.cmbSort1.Text.Trim().ToUpper() != "<NONE>")
						{
							if (strGroup3.Trim().ToUpper() == this.cmbSort1.Text.Trim().ToUpper())
							{
								bSort = true;
							}
						}
						if (bSort == false)
						{
							if (this.cmbSort2.Text.Trim().ToUpper() != "<NONE>")
							{
								if (strGroup3.Trim().ToUpper() == this.cmbSort2.Text.Trim().ToUpper())
								{
									bSort = true;
								}
							}
						}
						if (bSort == false)
						{
							if (this.cmbSort3.Text.Trim().ToUpper() != "<NONE>")
							{
								if (strGroup3.Trim().ToUpper() == this.cmbSort3.Text.Trim().ToUpper())
								{
									bSort = true;
								}
							}
						}

						if (bSort == false)
						{

							//make sure field is still a selected field
							for (x=0;x<=this.lstReportFields.Items.Count-1;x++)
							{
								if (this.lstReportFields.Items[x].ToString().Trim().ToUpper() ==
									this.cmbGroup3.Text.Trim().ToUpper()) break;
							}
							if (x<=this.lstReportFields.Items.Count-1)
							{
								//found the field
								bFound3=true;
							}
						}
					}
				}
				
			}
            this.cmbGroup1.Items.Clear();
			this.cmbGroup2.Items.Clear();
			this.cmbGroup3.Items.Clear();
			this.cmbGroup1.Items.Add("<none>");
			this.cmbGroup2.Items.Add("<none>");
			this.cmbGroup3.Items.Add("<none>");

			//dont add the item to the group list if it is a selected sort item
			for (x=0;x<=this.lstReportFields.Items.Count-1;x++)
			{
				if ((this.cmbSort1.Text.Trim().ToUpper() !=
					this.lstReportFields.Items[x].ToString().Trim().ToUpper()) &&
					(this.cmbSort2.Text.Trim().ToUpper() !=
					this.lstReportFields.Items[x].ToString().Trim().ToUpper()) &&
					(this.cmbSort3.Text.Trim().ToUpper() !=
					this.lstReportFields.Items[x].ToString().Trim().ToUpper()))
				{

					this.cmbGroup1.Items.Add(this.lstReportFields.Items[x]);
					this.cmbGroup2.Items.Add(this.lstReportFields.Items[x]);
					this.cmbGroup3.Items.Add(this.lstReportFields.Items[x]);
				}
			}

			if (bFound1 == false && bFound2==false && bFound3==true)
			{
				this.cmbGroup1.Text = strGroup3;
			}
			else if (bFound1==false && bFound2==true)
			{
				this.cmbGroup1.Text = strGroup2;
				this.cmbGroup2.Text = "<none>";
				bFound2=false;
			}
			else if (bFound1==true) 
			{
				this.cmbGroup1.Text = strGroup1;
			}
			else this.cmbGroup1.Text = "<none>";

			if (bFound2 == false && bFound3==true)
			{
				this.cmbGroup2.Text = strGroup3;
				this.cmbGroup3.Text = "<none>";
				bFound3=false;
			}
			else if (bFound2==true) 
			{
				this.cmbGroup2.Text = strGroup2;
			}
			else 
			{
			    this.cmbGroup2.Text = "<none>";
			}
			
            if (bFound3 == true) this.cmbGroup3.Text = strGroup3;
			else this.cmbGroup3.Text="<none>";

            if (bFound1==true) 
				this.cmbGroup2.Enabled=true;
			else this.cmbGroup2.Enabled=false;
			if (bFound2==true) this.cmbGroup3.Enabled=true;
			else this.cmbGroup3.Enabled=false;



			this.grpboxSelectFields.Visible=false;
			this.HelpButton(ref this.grpboxGroupRecords,ref this.btnHelp,"selectgroups");
			this.NextButton(ref this.grpboxGroupRecords ,ref this.btnNext,"selectgroups");
			this.PrevButton(ref this.grpboxGroupRecords,ref this.btnPrevious,"selectgroups");
			this.CancelButton(ref this.grpboxGroupRecords,ref this.btnCancel,"selectgroups");
			this.FinishButton(ref this.grpboxGroupRecords,ref this.btnFinish,"selectgroups");
			this.StepCombo(ref this.grpboxGroupRecords,ref this.cmbSteps,"selectgroups");
			this.position_controls(ref this.grpboxGroupRecords);
			this.cmbSteps.Text = this.cmbSteps.Items[1].ToString();
			this.btnFinish.Enabled=true;
			this.btnNext.Enabled=true;
			this.btnPrevious.Enabled=true;
			this.grpboxGroupRecords.Visible=true;
			this.m_strAction="";

		}

		private void cmbGroup1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			
			

		}

		private void cmbGroup1_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strAction=="updategroups") return;
			if (this.cmbGroup1.Text.Trim().ToUpper()=="<NONE>")
			{
				if (this.cmbGroup2.Enabled) this.cmbGroup1.Text = this.cmbGroup2.Text;
				if (this.cmbGroup3.Enabled) this.cmbGroup2.Text = this.cmbGroup3.Text;
				if (this.cmbGroup3.Enabled) this.cmbGroup3.Text = "<none>";
			}
			else if (this.cmbGroup1.Text.Trim().ToUpper() ==
				this.cmbGroup2.Text.Trim().ToUpper())
			{
				if (this.cmbGroup3.Enabled==false)
				{
					this.cmbGroup2.Text = "<none>";
				}
				else
				{
					this.cmbGroup2.Text = this.cmbGroup3.Text;
					this.cmbGroup3.Text = "<none>";
				}
				
			}
			else if (this.cmbGroup1.Text.Trim().ToUpper() ==
				this.cmbGroup3.Text.Trim().ToUpper())
			{
				this.cmbGroup3.Text = "<none>";
			}
			else this.cmbGroup2.Enabled=true;
			if (this.cmbGroup2.Enabled && this.cmbGroup1.Text.Trim().ToUpper() == "<NONE>")
				this.cmbGroup2.Enabled=false;
			if (this.cmbGroup3.Enabled && this.cmbGroup2.Text.Trim().ToUpper() == "<NONE>")
				this.cmbGroup3.Enabled=false;
			if (this.cmbGroup2.Enabled && this.cmbGroup2.Text.Trim().ToUpper() != "<NONE>")
				this.cmbGroup3.Enabled=true;


		}

		private void cmbGroup2_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strAction=="updategroups") return;
			if (this.cmbGroup1.Text.Trim().ToUpper()=="<NONE>")
			{
				if (this.cmbGroup2.Enabled) this.cmbGroup1.Text = this.cmbGroup2.Text;
				if (this.cmbGroup3.Enabled) this.cmbGroup2.Text = this.cmbGroup3.Text;
				if (this.cmbGroup3.Enabled) this.cmbGroup3.Text = "<none>";
			}
			else if (this.cmbGroup1.Text.Trim().ToUpper() ==
				this.cmbGroup2.Text.Trim().ToUpper())
			{
				if (this.cmbGroup3.Enabled==false)
				{
					this.cmbGroup2.Text = "<none>";
				}
				else
				{
					this.cmbGroup2.Text = this.cmbGroup3.Text;
					this.cmbGroup3.Text = "<none>";
				}
				
			}
			else if (this.cmbGroup2.Text.Trim().ToUpper() ==
				this.cmbGroup3.Text.Trim().ToUpper())
			{
				this.cmbGroup3.Text = "<none>";
			}
			else this.cmbGroup2.Enabled=true;
			if (this.cmbGroup2.Enabled && this.cmbGroup1.Text.Trim().ToUpper() == "<NONE>")
				this.cmbGroup2.Enabled=false;
			if (this.cmbGroup3.Enabled && this.cmbGroup2.Text.Trim().ToUpper() == "<NONE>")
				this.cmbGroup3.Enabled=false;
			if (this.cmbGroup2.Enabled && this.cmbGroup2.Text.Trim().ToUpper() != "<NONE>")
				this.cmbGroup3.Enabled=true;

		}

		private void cmbGroup1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled = true;
		}

		private void cmbGroup2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled = true;
		}

		private void cmbGroup3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled = true;
		}

		private void cmbGroup3_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strAction=="updategroups") return;
			if (this.cmbGroup2.Text.Trim().ToUpper() ==
				this.cmbGroup3.Text.Trim().ToUpper() ||
				this.cmbGroup1.Text.Trim().ToUpper() ==
				this.cmbGroup3.Text.Trim().ToUpper())
			{
				this.cmbGroup3.Text = "<none>";
			}
		}

		private void cmbSteps_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void rdoPortrait_Click(object sender, System.EventArgs e)
		{
			this.lblLandscape.Visible=false;
			this.lblPortrait.Visible=true;
		}

		private void rdoLandscape_Click(object sender, System.EventArgs e)
		{
		   this.lblPortrait.Visible=false;
		   this.lblLandscape.Visible=true;

		}

		private void grpboxSort_Enter(object sender, System.EventArgs e)
		{
		
		}
		private void SelectSort()
		{
			int x=0;
			//int y=0;
			string strSort1="";
			string strSort2="";
			string strSort3="";
			bool bFound1 = false;
			bool bFound2 = false;
			bool bFound3 = false;
			bool bGroup = false;
            
			switch (this.btnNext.Name)
			{
				case "reporttitle":
					this.grpboxFinish.Visible=false;
					break;
				case "selectgroups":
					this.grpboxGroupRecords.Visible=false;
					break;
				case "selectsort":
					this.grpboxSort.Visible=false;
					break;
                case "selectdetail":
					this.grpboxDetailAndSummary.Visible=false;
					break;
			}
			this.m_strAction="updatesort";
			//validate sort1 
			if (this.cmbSort1.Text.Trim().Length > 0 &&
				this.cmbSort1.Text.Trim() != "<none>")
			{
				strSort1 = this.cmbSort1.Text;
				//make sure this sort is not selected in a group
				if (this.cmbGroup1.Text.Trim().ToUpper() != "<NONE>")
				{
					if (strSort1.Trim().ToUpper() == this.cmbGroup1.Text.Trim().ToUpper())
					{
						bGroup = true;
					}
				}
				if (bGroup == false)
				{
					if (this.cmbGroup2.Text.Trim().ToUpper() != "<NONE>")
					{
						if (strSort1.Trim().ToUpper() == this.cmbGroup2.Text.Trim().ToUpper())
						{
							bGroup = true;
						}
					}
				}
				if (bGroup == false)
				{
					if (this.cmbGroup3.Text.Trim().ToUpper() != "<NONE>")
					{
						if (strSort1.Trim().ToUpper() == this.cmbGroup3.Text.Trim().ToUpper())
						{
							bGroup = true;
						}
					}
				}

				if (bGroup == false)
				{


					//make sure field is still a selected field
					for (x=0;x<=this.lstReportFields.Items.Count-1;x++)
					{
						if (this.lstReportFields.Items[x].ToString().Trim().ToUpper() ==
							this.cmbSort1.Text.Trim().ToUpper()) break;
					}
					if (x<=this.lstReportFields.Items.Count-1)
					{
						//found the field
						bFound1=true;
					}
				}
				//validate sort2	
				if (this.cmbSort2.Text.Trim().Length > 0 &&
					this.cmbSort2.Text.Trim() != "<none>")
				{
					strSort2=this.cmbSort2.Text;
					//make sure this sort is not selected in a group
					if (this.cmbGroup1.Text.Trim().ToUpper() != "<NONE>")
					{
						if (strSort2.Trim().ToUpper() == this.cmbGroup1.Text.Trim().ToUpper())
						{
							bGroup = true;
						}
					}
					if (bGroup == false)
					{
						if (this.cmbGroup2.Text.Trim().ToUpper() != "<NONE>")
						{
							if (strSort2.Trim().ToUpper() == this.cmbGroup2.Text.Trim().ToUpper())
							{
								bGroup = true;
							}
						}
					}
					if (bGroup == false)
					{
						if (this.cmbGroup3.Text.Trim().ToUpper() != "<NONE>")
						{
							if (strSort2.Trim().ToUpper() == this.cmbGroup3.Text.Trim().ToUpper())
							{
								bGroup = true;
							}
						}
					}

					if (bGroup == false)
					{

						//make sure field is still a selected field
						for (x=0;x<=this.lstReportFields.Items.Count-1;x++)
						{
							if (this.lstReportFields.Items[x].ToString().Trim().ToUpper() ==
								this.cmbSort2.Text.Trim().ToUpper()) break;
						}
						if (x<=this.lstReportFields.Items.Count-1)
						{
							//found the field
							bFound2=true;
						}
					}


					//validate group 3
					if (this.cmbSort3.Text.Trim().Length > 0 &&
						this.cmbSort3.Text.Trim() != "<none>")
					{
						strSort3 = this.cmbSort3.Text;

						//make sure this sort is not selected in a group
						if (this.cmbGroup1.Text.Trim().ToUpper() != "<NONE>")
						{
							if (strSort3.Trim().ToUpper() == this.cmbGroup1.Text.Trim().ToUpper())
							{
								bGroup = true;
							}
						}
						if (bGroup == false)
						{
							if (this.cmbGroup2.Text.Trim().ToUpper() != "<NONE>")
							{
								if (strSort3.Trim().ToUpper() == this.cmbGroup2.Text.Trim().ToUpper())
								{
									bGroup = true;
								}
							}
						}
						if (bGroup == false)
						{
							if (this.cmbGroup3.Text.Trim().ToUpper() != "<NONE>")
							{
								if (strSort3.Trim().ToUpper() == this.cmbGroup3.Text.Trim().ToUpper())
								{
									bGroup = true;
								}
							}
						}

						if (bGroup == false)
						{

							//make sure field is still a selected field
							for (x=0;x<=this.lstReportFields.Items.Count-1;x++)
							{
								if (this.lstReportFields.Items[x].ToString().Trim().ToUpper() ==
									this.cmbSort3.Text.Trim().ToUpper()) break;
							}
							if (x<=this.lstReportFields.Items.Count-1)
							{
								//found the field
								bFound3=true;
							}
						}
					}
				}
				
			}
			this.cmbSort1.Items.Clear();
			this.cmbSort2.Items.Clear();
			this.cmbSort3.Items.Clear();
			this.cmbSort1.Items.Add("<none>");
			this.cmbSort2.Items.Add("<none>");
			this.cmbSort3.Items.Add("<none>");


			for (x=0;x<=this.lstReportFields.Items.Count-1;x++)
			{
				//dont add the item to the sort list if it is a selected group item
				if ((this.cmbGroup1.Text.Trim().ToUpper() !=
					this.lstReportFields.Items[x].ToString().Trim().ToUpper()) &&
					(this.cmbGroup2.Text.Trim().ToUpper() !=
					this.lstReportFields.Items[x].ToString().Trim().ToUpper()) &&
					(this.cmbGroup3.Text.Trim().ToUpper() !=
					this.lstReportFields.Items[x].ToString().Trim().ToUpper()))
				{
					this.cmbSort1.Items.Add(this.lstReportFields.Items[x]);
					this.cmbSort2.Items.Add(this.lstReportFields.Items[x]);
					this.cmbSort3.Items.Add(this.lstReportFields.Items[x]);
				}
			}

			if (bFound1 == false && bFound2==false && bFound3==true)
			{
				this.cmbSort1.Text = strSort3;
			}
			else if (bFound1==false && bFound2==true)
			{
				this.cmbSort1.Text = strSort2;
				this.cmbSort2.Text = "<none>";
				bFound2=false;
			}
			else if (bFound1==true) 
			{
				this.cmbSort1.Text = strSort1;
			}
			else this.cmbSort1.Text = "<none>";

			if (bFound2 == false && bFound3==true)
			{
				this.cmbSort2.Text = strSort3;
				this.cmbSort3.Text = "<none>";
				bFound3=false;
			}
			else if (bFound2==true) 
			{
				this.cmbSort2.Text = strSort2;
			}
			else 
			{
				this.cmbSort2.Text = "<none>";
			}
			
			if (bFound3 == true) this.cmbSort3.Text = strSort3;
			else this.cmbSort3.Text="<none>";

			if (bFound1==true) 
				this.cmbSort2.Enabled=true;
			else this.cmbSort2.Enabled=false;
			if (bFound2==true) this.cmbSort3.Enabled=true;
			else this.cmbSort3.Enabled=false;



			this.grpboxSelectFields.Visible=false;
			this.HelpButton(ref this.grpboxSort,ref this.btnHelp,"selectsort");
			this.NextButton(ref this.grpboxSort ,ref this.btnNext,"selectsort");
			this.PrevButton(ref this.grpboxSort,ref this.btnPrevious,"selectsort");
			this.CancelButton(ref this.grpboxSort,ref this.btnCancel,"selectsort");
			this.FinishButton(ref this.grpboxSort,ref this.btnFinish,"selectsort");
			this.StepCombo(ref this.grpboxSort,ref this.cmbSteps,"selectsort");
			this.position_controls(ref this.grpboxSort);
			this.cmbSteps.Text = this.cmbSteps.Items[2].ToString();
			this.btnFinish.Enabled=true;
			this.btnNext.Enabled=true;
			this.btnPrevious.Enabled=true;
			this.grpboxSort.Visible=true;
			this.m_strAction="";

		}

		private void cmbSort1_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strAction=="updatesort") return;
			if (this.cmbSort1.Text.Trim().ToUpper()=="<NONE>")
			{
				if (this.cmbSort2.Enabled) this.cmbSort1.Text = this.cmbSort2.Text;
				if (this.cmbSort3.Enabled) this.cmbSort2.Text = this.cmbSort3.Text;
				if (this.cmbSort3.Enabled) this.cmbSort3.Text = "<none>";
			}
			else if (this.cmbSort1.Text.Trim().ToUpper() ==
				this.cmbSort2.Text.Trim().ToUpper())
			{
				if (this.cmbSort3.Enabled==false)
				{
					this.cmbSort2.Text = "<none>";
				}
				else
				{
					this.cmbSort2.Text = this.cmbSort3.Text;
					this.cmbSort3.Text = "<none>";
				}
				
			}
			else if (this.cmbSort1.Text.Trim().ToUpper() ==
				this.cmbSort3.Text.Trim().ToUpper())
			{
				this.cmbSort3.Text = "<none>";
			}
			else this.cmbSort2.Enabled=true;
			if (this.cmbSort2.Enabled && this.cmbSort1.Text.Trim().ToUpper() == "<NONE>")
				this.cmbSort2.Enabled=false;
			if (this.cmbSort3.Enabled && this.cmbSort2.Text.Trim().ToUpper() == "<NONE>")
				this.cmbSort3.Enabled=false;
			if (this.cmbSort2.Enabled && this.cmbSort2.Text.Trim().ToUpper() != "<NONE>")
				this.cmbSort3.Enabled=true;
		
		}

		private void cmbSort2_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strAction=="updatesort") return;
			if (this.cmbSort1.Text.Trim().ToUpper()=="<NONE>")
			{
				if (this.cmbSort2.Enabled) this.cmbSort1.Text = this.cmbSort2.Text;
				if (this.cmbSort3.Enabled) this.cmbSort2.Text = this.cmbSort3.Text;
				if (this.cmbSort3.Enabled) this.cmbSort3.Text = "<none>";
			}
			else if (this.cmbSort1.Text.Trim().ToUpper() ==
				this.cmbSort2.Text.Trim().ToUpper())
			{
				if (this.cmbSort3.Enabled==false)
				{
					this.cmbSort2.Text = "<none>";
				}
				else
				{
					this.cmbSort2.Text = this.cmbSort3.Text;
					this.cmbSort3.Text = "<none>";
				}
				
			}
			else if (this.cmbSort2.Text.Trim().ToUpper() ==
				this.cmbSort3.Text.Trim().ToUpper())
			{
				this.cmbSort3.Text = "<none>";
			}
			else this.cmbSort2.Enabled=true;
			if (this.cmbSort2.Enabled && this.cmbSort1.Text.Trim().ToUpper() == "<NONE>")
				this.cmbSort2.Enabled=false;
			if (this.cmbSort3.Enabled && this.cmbSort2.Text.Trim().ToUpper() == "<NONE>")
				this.cmbSort3.Enabled=false;
			if (this.cmbSort2.Enabled && this.cmbSort2.Text.Trim().ToUpper() != "<NONE>")
				this.cmbSort3.Enabled=true;
		
		}

		private void cmbSort3_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strAction=="updatesort") return;
			if (this.cmbSort2.Text.Trim().ToUpper() ==
				this.cmbSort3.Text.Trim().ToUpper() ||
				this.cmbSort1.Text.Trim().ToUpper() ==
				this.cmbSort3.Text.Trim().ToUpper())
			{
				this.cmbSort3.Text = "<none>";
			}
		
		}
		private void SelectDetail()
		{
			int x=0;
			int y=0;
			//int z=0;
			//int intTextWidth=0;


			switch (this.btnNext.Name)
			{
				case "reporttitle":
					this.grpboxFinish.Visible=false;
					break;
				case "selectgroups":
					this.grpboxGroupRecords.Visible=false;
					break;
				case "selectsort":
					this.grpboxSort.Visible=false;
					break;
				case "selectdetail":
					this.grpboxDetailAndSummary.Visible=false;
					break;
			}
			this.HelpButton(ref this.grpboxDetailAndSummary,ref this.btnHelp, "selectdetail");
			this.FinishButton(ref this.grpboxDetailAndSummary,ref this.btnFinish, "selectdetail");
			this.NextButton(ref this.grpboxDetailAndSummary,ref this.btnNext, "selectdetail");
			this.PrevButton(ref this.grpboxDetailAndSummary,ref this.btnPrevious, "selectdetail");
			this.CancelButton(ref this.grpboxDetailAndSummary,ref this.btnCancel,"selectdetail");
			this.StepCombo(ref this.grpboxDetailAndSummary,ref this.cmbSteps,"selectdetail");             
			this.btnNext.Enabled=true;
			this.btnFinish.Enabled=true;
			this.btnPrevious.Enabled=true;
			this.btnCancel.Enabled=true;
			this.position_controls(ref this.grpboxDetailAndSummary);
			this.m_strAction="suppress";
			this.cmbSteps.Text = this.cmbSteps.Items[3].ToString();
			this.m_strAction="";
			this.grpboxDetailAndSummary.Visible=true;

			
            
		//if the data grid has not been initialized then do so
			if (this.dgDetailOptions.IsHandleCreated==false)
			{
			}
			else
			{
				//remove fields names from the list that are no longer selected
				// or have been added as a group field
				for (x=0; x <= this.m_dtDetailOptions.Rows.Count - 1; x++)
				{
					for (y=0; y<=this.lstReportFields.Items.Count-1; y++)
					{
						//see if in the report field list
						if (this.m_dtDetailOptions.Rows[x]["field_name"].ToString().Trim().ToUpper() == 
							this.lstReportFields.Items[y].ToString().Trim().ToUpper())
						{
							break;
						}
					}
					if (y<=this.lstReportFields.Items.Count-1)
					{
						//found the record and not part of a group do nothing
					}
					else
					{
						//remove the record
						this.m_dtDetailOptions.Rows[x].Delete();
					}
				}
			}
     		//add report list items
			
			for (x=0; x<= this.lstReportFields.Items.Count-1;x++)
			{
				//check if a group field
					//check if already added
					for (y=0;y<=this.m_dtDetailOptions.Rows.Count-1;y++)
					{
						if (this.m_dtDetailOptions.Rows[y][0].ToString().Trim().ToUpper()==
							this.lstReportFields.Items[x].ToString().Trim().ToUpper())
							break;
					}
					
					
					if (y<=this.m_dtDetailOptions.Rows.Count-1)
					{
						//record already added so do nothing
					}
					else
					{
						//add the record
						System.Data.DataRow p_row = this.m_dtDetailOptions.NewRow();
						p_row["field_name"] = this.lstReportFields.Items[x].ToString().Trim(); 
						p_row["sum"] = false;
						p_row["max"] = false;
						p_row["min"] = false;
						p_row["count"] = false;
						p_row["avg_sum"] = false; 
						p_row["pct_count"] = false; 

						//get the datatype
						this.m_dtDetailOptions.Rows.Add(p_row);
					}
				//}
			}
			this.m_dvDetailOptions = new DataView(this.m_dtDetailOptions);
			this.m_dvDetailOptions.AllowDelete=false;
			this.m_dvDetailOptions.AllowEdit = true;
			this.m_dvDetailOptions.AllowNew = false;
			
		

			//custom define the grid style
			DataGridTableStyle tableStyle = new DataGridTableStyle();

			//map the data grid table style to the detail summary table
			tableStyle.MappingName = "detail_summary_user_options";   
			tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
			tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
           

			//define datagrid text box column
			 DataGridTextBoxColumn aColumnTextColumn ;
			//define datagrid boolean column
			 System.Windows.Forms.DataGridBoolColumn aColumnBoolColumn;
                    
			//loop through all the columns in the table	
			for(int i = 0; i <= this.m_dtDetailOptions.Columns.Count-1; ++i)
			{
				if (this.m_dtDetailOptions.Columns[i].DataType.FullName.ToString().Trim() == "System.String")
				{

					//create a new instance of the DataGridColoredTextBoxColumn class
					aColumnTextColumn = new DataGridTextBoxColumn();
					aColumnTextColumn.HeaderText = this.m_dtDetailOptions.Columns[i].ColumnName;
					//all columns are read-only except the rx_intensity column
					aColumnTextColumn.ReadOnly=true;
					//assign the mappingname property the data sets column name
					aColumnTextColumn.MappingName = this.m_dtDetailOptions.Columns[i].ColumnName;
					//add the datagridcoloredtextboxcolumn object to the data grid table style object
					tableStyle.GridColumnStyles.Add(aColumnTextColumn);
				}
				else
				{
					aColumnBoolColumn = new DataGridBoolColumn();
					aColumnBoolColumn.HeaderText = this.m_dtDetailOptions.Columns[i].ColumnName;
					aColumnBoolColumn.ReadOnly = false;
					aColumnBoolColumn.MappingName = this.m_dtDetailOptions.Columns[i].ColumnName;
					aColumnBoolColumn.Width = (int)this.CreateGraphics().MeasureString(this.m_dtDetailOptions.Columns[i].ColumnName.ToString().Trim() + "**", this.dgDetailOptions.Font).Width;
					tableStyle.GridColumnStyles.Add(aColumnBoolColumn);
				}
				
			}
			dgDetailOptions.BackgroundColor=frmMain.g_oGridViewRowBackgroundColor;
			if (frmMain.g_oGridViewFont != null) dgDetailOptions.Font = frmMain.g_oGridViewFont;
       	
		    this.dgDetailOptions.TableStyles.Clear();    
			// make the dataGrid use our new tablestyle and bind it to our table
			tableStyle.AllowSorting = false;
			this.dgDetailOptions.TableStyles.Add(tableStyle);
			
            
			this.dgDetailOptions.SetDataBinding(this.m_dvDetailOptions,"");
			
		}
		
		private void SelectDetail2()
		{
			int x=0;
			int y=0;
			//int intTextWidth=0;
			
			switch (this.btnNext.Name)
			{
				case "reporttitle":
					this.grpboxFinish.Visible=false;
					break;
				case "selectgroups":
					this.grpboxGroupRecords.Visible=false;
					break;
				case "selectsort":
					this.grpboxSort.Visible=false;
					break;
				case "selectdetail":
					this.grpboxDetailAndSummary.Visible=false;
					break;
			}
			this.HelpButton(ref this.grpboxDetailAndSummary,ref this.btnHelp, "selectdetail");
			this.FinishButton(ref this.grpboxDetailAndSummary,ref this.btnFinish, "selectdetail");
			this.NextButton(ref this.grpboxDetailAndSummary,ref this.btnNext, "selectdetail");
			this.PrevButton(ref this.grpboxDetailAndSummary,ref this.btnPrevious, "selectdetail");
			this.CancelButton(ref this.grpboxDetailAndSummary,ref this.btnCancel,"selectdetail");
			this.StepCombo(ref this.grpboxDetailAndSummary,ref this.cmbSteps,"selectdetail");             
			this.btnNext.Enabled=true;
			this.btnFinish.Enabled=true;
			this.btnPrevious.Enabled=true;
			this.btnCancel.Enabled=true;
			this.position_controls(ref this.grpboxDetailAndSummary);
			this.m_strAction="suppress";
			this.cmbSteps.Text = this.cmbSteps.Items[3].ToString();
			this.m_strAction="";
			this.grpboxDetailAndSummary.Visible=true;

			
            
			//if the data grid has not been initialized then do so
			if (this.dgDetailOptions.IsHandleCreated==false)
			{
				
		
			}
			else
			{
			
			}

			//add report list items
			for (x=0; x<= this.lstReportFields.Items.Count-1;x++)
			{
				//check if a group field
				if (this.lstReportFields.Items[x].ToString().Trim().ToUpper() !=
					this.cmbGroup1.Text.ToString().Trim().ToUpper() &&
					this.lstReportFields.Items[x].ToString().Trim().ToUpper() !=
					this.cmbGroup2.Text.ToString().Trim().ToUpper() &&
					this.lstReportFields.Items[x].ToString().Trim().ToUpper() !=
					this.cmbGroup3.Text.ToString().Trim().ToUpper())
				{
					//check if already added
					for (y=0;y<=this.m_dtDetailOptions.Rows.Count-1;y++)
					{
						if (this.m_dtDetailOptions.Rows[y][0].ToString().Trim().ToUpper()==
							this.lstReportFields.Items[x].ToString().Trim().ToUpper())
							break;
					}
					
					
					if (y<=this.m_dtDetailOptions.Rows.Count-1)
					{
						//record already added so do nothing
					}
					else
					{
						//add the record
						System.Data.DataRow p_row = this.m_dtDetailOptions.NewRow();
						p_row["field_name"] = this.lstReportFields.Items[x].ToString().Trim();
						this.m_dtDetailOptions.Rows.Add(p_row);
					}
				}
			}

			this.m_dvDetailOptions = new DataView(this.m_dtDetailOptions);
			 
			this.dgDetailOptions.DataSource = this.m_dvDetailOptions;  //.Tables["detail"] ;
			System.Windows.Forms.CurrencyManager mgr = (System.Windows.Forms.CurrencyManager)this.BindingContext[this.m_dvDetailOptions];
			mgr.Refresh();
			this.dgDetailOptions.Expand(-1);

				
			
		
		}

		private void dgDetailOptions_Click(object sender, System.EventArgs e)
		{
		
		}

		private void dgDetailOptions_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			//automatically place a check mark or remove a check mark on a mouse click
			DataGrid.HitTestInfo hti = this.dgDetailOptions.HitTest(e.X, e.Y);
			try
			{
				if( hti.Type == DataGrid.HitTestType.Cell &&
					hti.Column > 0)
				{
					this.dgDetailOptions[hti.Row, hti.Column] = ! (bool) this.dgDetailOptions[hti.Row, hti.Column];
				}
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

}

}
