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
	public class uc_scenario_fvs_prepost_variables_tiebreaker: System.Windows.Forms.UserControl
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
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmCoreScenario _frmScenario=null;
		private FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization _uc_optimization;


				
		private int m_intCurVar=-1;
		int m_intCurVariableDefinitionStepCount=1;
		string[] m_strUserNavigation=null;
		

		public const byte NUMBER_OF_VARIABLES=4;
		const byte VARIABLE_DEFINITION_STEPS=4;


		const int WIZARD_STEP_VARIABLES_DEFINED=0;
		const int WIZARD_STEP_VARIABLE_SELECT=1;
		const int WIZARD_STEP_VARIABLE_BETTER=2;
		const int WIZARD_STEP_VARIABLE_WORSE=3;
		const int WIZARD_STEP_VARIABLE_EFFECTIVE=4;
		const int WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE=5;

		
		//const int COLUMN_OPTIMIZE=0;
		const int COLUMN_CHECKBOX=0;
		const int COLUMN_METHOD=1;
		const int COLUMN_FVSVARIABLE=2;
		const int COLUMN_VALUESOURCE=3;
		const int COLUMN_MAXMIN=4;

		public bool m_bSave=false;
		private bool _bDisplayAuditMsg=true;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreaker;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.ColumnHeader lvColChecked;
		private System.Windows.Forms.ColumnHeader lvColMethod;
		private System.Windows.Forms.ColumnHeader lvColFVSVariableName;
		private System.Windows.Forms.ColumnHeader lvColMinMax;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerVariable;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerVariableValues;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableValues;
		private System.Windows.Forms.ListBox lstFVSVariablesTieBreakerVariableValues;
		private System.Windows.Forms.GroupBox grpFVSVariablesTieBreakerVariableValuesSelected;
		private System.Windows.Forms.Label lblFVSVariablesTieBreakerVariableValuesSelected;
		private System.Windows.Forms.GroupBox grpMaxMin;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerTreatmentIntensity;
		private System.Windows.Forms.Panel panel2;
		public FIA_Biosum_Manager.uc_scenario_treatment_intensity uc_scenario_treatment_intensity1;
		private System.Windows.Forms.Panel pnlTieBreaker;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerValues;
		private System.Windows.Forms.ListView lvFVSVariablesTieBreakerValues;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerEdit;
        private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors= new ListViewAlternateBackgroundColors();

		
		private System.Windows.Forms.RadioButton rdoFVSVariablesTieBreakerVariableValuesSelectedMin;
		private System.Windows.Forms.RadioButton rdoFVSVariablesTieBreakerVariableValuesSelectedMax;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerTreatmentIntensityPrev;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerTreatmentIntensityClear;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerTreatmentIntensityDone;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerTreatmentIntensityCancel;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableClear;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableDone;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableCancel;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableNext;
		private TieBreaker_Collection m_oNewTieBreakerCollection = new TieBreaker_Collection();
		private TieBreaker_Collection m_oOldTieBreakerCollection = new TieBreaker_Collection();
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerAudit;
		private System.Windows.Forms.ColumnHeader lvColFieldSource;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerVariableValueSource;
		private System.Windows.Forms.ComboBox cmbFVSVariablesTieBreakerVariableValueSource;
		public TieBreaker_Collection m_oSavTieBreakerCollection = new TieBreaker_Collection();
		

		public class TieBreakerItem
		{
			public bool bSelected=true;
			public string strMethod="";
			public string strFVSVariableName="";
			public string strMaxYN="N";
			public string strMinYN="N";
			public string strValueSource="";
			public int intListViewIndex=-1;
			
			

			public void Copy(TieBreakerItem p_oSource,ref TieBreakerItem p_oDest)
			{
				p_oDest.bSelected = p_oSource.bSelected;
				p_oDest.strMethod=p_oSource.strMethod;
				p_oDest.strFVSVariableName = p_oSource.strFVSVariableName;
				p_oDest.strMaxYN = p_oSource.strMaxYN;
				p_oDest.strMinYN = p_oSource.strMinYN;
				p_oDest.strValueSource=p_oSource.strValueSource;
				p_oDest.intListViewIndex = p_oSource.intListViewIndex;
				

			}
		}

		public class TieBreaker_Collection : System.Collections.CollectionBase
		{
			public TieBreaker_Collection()
			{
				//
				// TODO: Add constructor logic here
				//
			}

			public void Add(FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem m_oTieBreaker)
			{
				// vérify if object is not already in
				if (this.List.Contains(m_oTieBreaker))
					throw new InvalidOperationException();
 
				// adding it
				this.List.Add(m_oTieBreaker);
 
				// return collection
				//return this;
			}
			public void Remove(int index)
			{
				// Check to see if there is a widget at the supplied index.
				if (index > Count - 1 || index < 0)
					// If no widget exists, a messagebox is shown and the operation 
					// is canColumned.
				{
					System.Windows.Forms.MessageBox.Show("Index not valid!");
				}
				else
				{
					List.RemoveAt(index); 
				}
			}
			public FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem Item(int Index)
			{
				// The appropriate item is retrieved from the List object and
				// explicitly cast to the Widget type, then returned to the 
				// caller.
				return (FIA_Biosum_Manager.uc_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem) List[Index];
			}
			public void Copy(TieBreaker_Collection p_oSource,ref TieBreaker_Collection p_oDest,bool p_bInitializeDest)
			{
				int x;
				if (p_bInitializeDest) p_oDest.Clear();
				for (x=0;x<=p_oSource.Count-1;x++)
				{
					TieBreakerItem oItem = new TieBreakerItem();
					oItem.Copy(p_oSource.Item(x),ref oItem);
					p_oDest.Add(oItem);

				}
			}


		}

		public uc_scenario_fvs_prepost_variables_tiebreaker()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			if (frmMain.g_oGridViewFont != null) this.lvFVSVariablesTieBreakerValues.Font = frmMain.g_oGridViewFont;
			
			for (int x=0;x<=this.lvFVSVariablesTieBreakerValues.Items.Count-1;x++)
				this.lvFVSVariablesTieBreakerValues.Items[x].UseItemStyleForSubItems=false;
			this.m_oLvRowColors.InitializeRowCollection();
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceListView = this.lvFVSVariablesTieBreakerValues;
			this.m_oLvRowColors.CustomFullRowSelect=true;
			this.m_oLvRowColors.AddRow();
			this.m_oLvRowColors.AddColumns(0,this.lvFVSVariablesTieBreakerValues.Columns.Count);
			this.m_oLvRowColors.AddRow();
			this.m_oLvRowColors.AddColumns(1,this.lvFVSVariablesTieBreakerValues.Columns.Count);
			this.m_oLvRowColors.ListView();


			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Hide();
			this.grpboxFVSVariablesTieBreakerVariable.Hide();
			this.grpboxFVSVariablesTieBreaker.Show();
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
			System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
																													 "",
																													 "FVS Variable",
																													 "Not Defined",
																													 "Not Defined",
																													 "Not Defined"}, -1);
			System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem(new string[] {
																													 "",
																													 "Treatment Intensity",
																													 "NA",
																													 "NA",
																													 "NA"}, -1);
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity = new System.Windows.Forms.GroupBox();
			this.panel2 = new System.Windows.Forms.Panel();
			this.uc_scenario_treatment_intensity1 = new FIA_Biosum_Manager.uc_scenario_treatment_intensity();
			this.btnFVSVariablesTieBreakerTreatmentIntensityPrev = new System.Windows.Forms.Button();
			this.btnFVSVariablesTieBreakerTreatmentIntensityClear = new System.Windows.Forms.Button();
			this.btnFVSVariablesTieBreakerTreatmentIntensityDone = new System.Windows.Forms.Button();
			this.btnFVSVariablesTieBreakerTreatmentIntensityCancel = new System.Windows.Forms.Button();
			this.grpboxFVSVariablesTieBreakerVariable = new System.Windows.Forms.GroupBox();
			this.panel1 = new System.Windows.Forms.Panel();
			this.grpboxFVSVariablesTieBreakerVariableValueSource = new System.Windows.Forms.GroupBox();
			this.cmbFVSVariablesTieBreakerVariableValueSource = new System.Windows.Forms.ComboBox();
			this.grpMaxMin = new System.Windows.Forms.GroupBox();
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin = new System.Windows.Forms.RadioButton();
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax = new System.Windows.Forms.RadioButton();
			this.grpboxFVSVariablesTieBreakerVariableValues = new System.Windows.Forms.GroupBox();
			this.btnFVSVariablesTieBreakerVariableValues = new System.Windows.Forms.Button();
			this.lstFVSVariablesTieBreakerVariableValues = new System.Windows.Forms.ListBox();
			this.grpFVSVariablesTieBreakerVariableValuesSelected = new System.Windows.Forms.GroupBox();
			this.lblFVSVariablesTieBreakerVariableValuesSelected = new System.Windows.Forms.Label();
			this.btnFVSVariablesTieBreakerVariableClear = new System.Windows.Forms.Button();
			this.btnFVSVariablesTieBreakerVariableDone = new System.Windows.Forms.Button();
			this.btnFVSVariablesTieBreakerVariableCancel = new System.Windows.Forms.Button();
			this.btnFVSVariablesTieBreakerVariableNext = new System.Windows.Forms.Button();
			this.grpboxFVSVariablesTieBreaker = new System.Windows.Forms.GroupBox();
			this.pnlTieBreaker = new System.Windows.Forms.Panel();
			this.grpboxFVSVariablesTieBreakerValues = new System.Windows.Forms.GroupBox();
			this.lvFVSVariablesTieBreakerValues = new System.Windows.Forms.ListView();
			this.lvColChecked = new System.Windows.Forms.ColumnHeader();
			this.lvColMethod = new System.Windows.Forms.ColumnHeader();
			this.lvColFVSVariableName = new System.Windows.Forms.ColumnHeader();
			this.lvColFieldSource = new System.Windows.Forms.ColumnHeader();
			this.lvColMinMax = new System.Windows.Forms.ColumnHeader();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.btnFVSVariablesTieBreakerAudit = new System.Windows.Forms.Button();
			this.btnFVSVariablesTieBreakerEdit = new System.Windows.Forms.Button();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.SuspendLayout();
			this.panel2.SuspendLayout();
			this.grpboxFVSVariablesTieBreakerVariable.SuspendLayout();
			this.panel1.SuspendLayout();
			this.grpboxFVSVariablesTieBreakerVariableValueSource.SuspendLayout();
			this.grpMaxMin.SuspendLayout();
			this.grpboxFVSVariablesTieBreakerVariableValues.SuspendLayout();
			this.grpFVSVariablesTieBreakerVariableValuesSelected.SuspendLayout();
			this.grpboxFVSVariablesTieBreaker.SuspendLayout();
			this.pnlTieBreaker.SuspendLayout();
			this.grpboxFVSVariablesTieBreakerValues.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox1.Controls.Add(this.grpboxFVSVariablesTieBreakerTreatmentIntensity);
			this.groupBox1.Controls.Add(this.grpboxFVSVariablesTieBreakerVariable);
			this.groupBox1.Controls.Add(this.grpboxFVSVariablesTieBreaker);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(900, 2000);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// grpboxFVSVariablesTieBreakerTreatmentIntensity
			// 
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.BackColor = System.Drawing.SystemColors.Control;
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Controls.Add(this.panel2);
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.ForeColor = System.Drawing.Color.Black;
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Location = new System.Drawing.Point(16, 1016);
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Name = "grpboxFVSVariablesTieBreakerTreatmentIntensity";
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Size = new System.Drawing.Size(872, 448);
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.TabIndex = 35;
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.TabStop = false;
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Text = "Treatment Intensity";
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Resize += new System.EventHandler(this.grpboxFVSVariablesTieBreakerTreatmentIntensity_Resize);
			// 
			// panel2
			// 
			this.panel2.AutoScroll = true;
			this.panel2.Controls.Add(this.uc_scenario_treatment_intensity1);
			this.panel2.Controls.Add(this.btnFVSVariablesTieBreakerTreatmentIntensityPrev);
			this.panel2.Controls.Add(this.btnFVSVariablesTieBreakerTreatmentIntensityClear);
			this.panel2.Controls.Add(this.btnFVSVariablesTieBreakerTreatmentIntensityDone);
			this.panel2.Controls.Add(this.btnFVSVariablesTieBreakerTreatmentIntensityCancel);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel2.Location = new System.Drawing.Point(3, 18);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(866, 427);
			this.panel2.TabIndex = 12;
			// 
			// uc_scenario_treatment_intensity1
			// 
			this.uc_scenario_treatment_intensity1.BackColor = System.Drawing.SystemColors.Control;
			this.uc_scenario_treatment_intensity1.Location = new System.Drawing.Point(8, 8);
			this.uc_scenario_treatment_intensity1.Name = "uc_scenario_treatment_intensity1";
			this.uc_scenario_treatment_intensity1.ReferenceCoreScenarioForm = null;
			this.uc_scenario_treatment_intensity1.Size = new System.Drawing.Size(840, 344);
			this.uc_scenario_treatment_intensity1.TabIndex = 13;
			this.uc_scenario_treatment_intensity1.Load += new System.EventHandler(this.uc_scenario_treatment_intensity1_Load);
			// 
			// btnFVSVariablesTieBreakerTreatmentIntensityPrev
			// 
			this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Location = new System.Drawing.Point(528, 376);
			this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Name = "btnFVSVariablesTieBreakerTreatmentIntensityPrev";
			this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Size = new System.Drawing.Size(88, 40);
			this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.TabIndex = 12;
			this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Text = "<--Previous";
			this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerTreatmentIntensityPrev_Click);
			// 
			// btnFVSVariablesTieBreakerTreatmentIntensityClear
			// 
			this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Location = new System.Drawing.Point(24, 376);
			this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Name = "btnFVSVariablesTieBreakerTreatmentIntensityClear";
			this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Size = new System.Drawing.Size(72, 40);
			this.btnFVSVariablesTieBreakerTreatmentIntensityClear.TabIndex = 5;
			this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Text = "Clear";
			this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerTreatmentIntensityClear_Click);
			// 
			// btnFVSVariablesTieBreakerTreatmentIntensityDone
			// 
			this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Location = new System.Drawing.Point(352, 376);
			this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Name = "btnFVSVariablesTieBreakerTreatmentIntensityDone";
			this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Size = new System.Drawing.Size(88, 40);
			this.btnFVSVariablesTieBreakerTreatmentIntensityDone.TabIndex = 11;
			this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Text = "Done";
			this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerTreatmentIntensityDone_Click);
			// 
			// btnFVSVariablesTieBreakerTreatmentIntensityCancel
			// 
			this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Location = new System.Drawing.Point(440, 376);
			this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Name = "btnFVSVariablesTieBreakerTreatmentIntensityCancel";
			this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Size = new System.Drawing.Size(88, 40);
			this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.TabIndex = 9;
			this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Text = "Cancel";
			this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerTreatmentIntensityCancel_Click);
			// 
			// grpboxFVSVariablesTieBreakerVariable
			// 
			this.grpboxFVSVariablesTieBreakerVariable.BackColor = System.Drawing.SystemColors.Control;
			this.grpboxFVSVariablesTieBreakerVariable.Controls.Add(this.panel1);
			this.grpboxFVSVariablesTieBreakerVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxFVSVariablesTieBreakerVariable.ForeColor = System.Drawing.Color.Black;
			this.grpboxFVSVariablesTieBreakerVariable.Location = new System.Drawing.Point(16, 544);
			this.grpboxFVSVariablesTieBreakerVariable.Name = "grpboxFVSVariablesTieBreakerVariable";
			this.grpboxFVSVariablesTieBreakerVariable.Size = new System.Drawing.Size(872, 448);
			this.grpboxFVSVariablesTieBreakerVariable.TabIndex = 34;
			this.grpboxFVSVariablesTieBreakerVariable.TabStop = false;
			this.grpboxFVSVariablesTieBreakerVariable.Text = "FVS Variable";
			// 
			// panel1
			// 
			this.panel1.AutoScroll = true;
			this.panel1.Controls.Add(this.grpboxFVSVariablesTieBreakerVariableValueSource);
			this.panel1.Controls.Add(this.grpMaxMin);
			this.panel1.Controls.Add(this.grpboxFVSVariablesTieBreakerVariableValues);
			this.panel1.Controls.Add(this.grpFVSVariablesTieBreakerVariableValuesSelected);
			this.panel1.Controls.Add(this.btnFVSVariablesTieBreakerVariableClear);
			this.panel1.Controls.Add(this.btnFVSVariablesTieBreakerVariableDone);
			this.panel1.Controls.Add(this.btnFVSVariablesTieBreakerVariableCancel);
			this.panel1.Controls.Add(this.btnFVSVariablesTieBreakerVariableNext);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(3, 18);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(866, 427);
			this.panel1.TabIndex = 12;
			// 
			// grpboxFVSVariablesTieBreakerVariableValueSource
			// 
			this.grpboxFVSVariablesTieBreakerVariableValueSource.Controls.Add(this.cmbFVSVariablesTieBreakerVariableValueSource);
			this.grpboxFVSVariablesTieBreakerVariableValueSource.Location = new System.Drawing.Point(8, 296);
			this.grpboxFVSVariablesTieBreakerVariableValueSource.Name = "grpboxFVSVariablesTieBreakerVariableValueSource";
			this.grpboxFVSVariablesTieBreakerVariableValueSource.Size = new System.Drawing.Size(344, 72);
			this.grpboxFVSVariablesTieBreakerVariableValueSource.TabIndex = 19;
			this.grpboxFVSVariablesTieBreakerVariableValueSource.TabStop = false;
			this.grpboxFVSVariablesTieBreakerVariableValueSource.Text = "Post Treatment Variable Or Pre/Post Treatment Change";
			// 
			// cmbFVSVariablesTieBreakerVariableValueSource
			// 
			this.cmbFVSVariablesTieBreakerVariableValueSource.Items.AddRange(new object[] {
																							  "Post Value",
																							  "Post - Pre  Change Value"});
			this.cmbFVSVariablesTieBreakerVariableValueSource.Location = new System.Drawing.Point(16, 40);
			this.cmbFVSVariablesTieBreakerVariableValueSource.Name = "cmbFVSVariablesTieBreakerVariableValueSource";
			this.cmbFVSVariablesTieBreakerVariableValueSource.Size = new System.Drawing.Size(320, 24);
			this.cmbFVSVariablesTieBreakerVariableValueSource.TabIndex = 0;
			this.cmbFVSVariablesTieBreakerVariableValueSource.Text = "Post Value";
			// 
			// grpMaxMin
			// 
			this.grpMaxMin.Controls.Add(this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin);
			this.grpMaxMin.Controls.Add(this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax);
			this.grpMaxMin.Location = new System.Drawing.Point(360, 296);
			this.grpMaxMin.Name = "grpMaxMin";
			this.grpMaxMin.Size = new System.Drawing.Size(464, 64);
			this.grpMaxMin.TabIndex = 18;
			this.grpMaxMin.TabStop = false;
			this.grpMaxMin.Text = "Aggregate Setting For Tie Breaker Variable";
			// 
			// rdoFVSVariablesTieBreakerVariableValuesSelectedMin
			// 
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Location = new System.Drawing.Point(256, 16);
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Name = "rdoFVSVariablesTieBreakerVariableValuesSelectedMin";
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Size = new System.Drawing.Size(176, 40);
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.TabIndex = 14;
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Text = "Minimum";
			// 
			// rdoFVSVariablesTieBreakerVariableValuesSelectedMax
			// 
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked = true;
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Location = new System.Drawing.Point(32, 16);
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Name = "rdoFVSVariablesTieBreakerVariableValuesSelectedMax";
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Size = new System.Drawing.Size(176, 40);
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.TabIndex = 12;
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.TabStop = true;
			this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Text = "Maximum";
			// 
			// grpboxFVSVariablesTieBreakerVariableValues
			// 
			this.grpboxFVSVariablesTieBreakerVariableValues.Controls.Add(this.btnFVSVariablesTieBreakerVariableValues);
			this.grpboxFVSVariablesTieBreakerVariableValues.Controls.Add(this.lstFVSVariablesTieBreakerVariableValues);
			this.grpboxFVSVariablesTieBreakerVariableValues.Location = new System.Drawing.Point(8, 16);
			this.grpboxFVSVariablesTieBreakerVariableValues.Name = "grpboxFVSVariablesTieBreakerVariableValues";
			this.grpboxFVSVariablesTieBreakerVariableValues.Size = new System.Drawing.Size(816, 216);
			this.grpboxFVSVariablesTieBreakerVariableValues.TabIndex = 0;
			this.grpboxFVSVariablesTieBreakerVariableValues.TabStop = false;
			this.grpboxFVSVariablesTieBreakerVariableValues.Text = "Tie Breaker Variable List";
			// 
			// btnFVSVariablesTieBreakerVariableValues
			// 
			this.btnFVSVariablesTieBreakerVariableValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFVSVariablesTieBreakerVariableValues.Location = new System.Drawing.Point(448, 32);
			this.btnFVSVariablesTieBreakerVariableValues.Name = "btnFVSVariablesTieBreakerVariableValues";
			this.btnFVSVariablesTieBreakerVariableValues.Size = new System.Drawing.Size(184, 144);
			this.btnFVSVariablesTieBreakerVariableValues.TabIndex = 1;
			this.btnFVSVariablesTieBreakerVariableValues.Text = "Select";
			this.btnFVSVariablesTieBreakerVariableValues.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableValues_Click);
			// 
			// lstFVSVariablesTieBreakerVariableValues
			// 
			this.lstFVSVariablesTieBreakerVariableValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lstFVSVariablesTieBreakerVariableValues.ItemHeight = 16;
			this.lstFVSVariablesTieBreakerVariableValues.Location = new System.Drawing.Point(8, 16);
			this.lstFVSVariablesTieBreakerVariableValues.Name = "lstFVSVariablesTieBreakerVariableValues";
			this.lstFVSVariablesTieBreakerVariableValues.Size = new System.Drawing.Size(424, 180);
			this.lstFVSVariablesTieBreakerVariableValues.TabIndex = 0;
			// 
			// grpFVSVariablesTieBreakerVariableValuesSelected
			// 
			this.grpFVSVariablesTieBreakerVariableValuesSelected.Controls.Add(this.lblFVSVariablesTieBreakerVariableValuesSelected);
			this.grpFVSVariablesTieBreakerVariableValuesSelected.Location = new System.Drawing.Point(8, 240);
			this.grpFVSVariablesTieBreakerVariableValuesSelected.Name = "grpFVSVariablesTieBreakerVariableValuesSelected";
			this.grpFVSVariablesTieBreakerVariableValuesSelected.Size = new System.Drawing.Size(816, 51);
			this.grpFVSVariablesTieBreakerVariableValuesSelected.TabIndex = 4;
			this.grpFVSVariablesTieBreakerVariableValuesSelected.TabStop = false;
			this.grpFVSVariablesTieBreakerVariableValuesSelected.Text = "Selected Tie Breaker Variable";
			// 
			// lblFVSVariablesTieBreakerVariableValuesSelected
			// 
			this.lblFVSVariablesTieBreakerVariableValuesSelected.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lblFVSVariablesTieBreakerVariableValuesSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFVSVariablesTieBreakerVariableValuesSelected.Location = new System.Drawing.Point(3, 18);
			this.lblFVSVariablesTieBreakerVariableValuesSelected.Name = "lblFVSVariablesTieBreakerVariableValuesSelected";
			this.lblFVSVariablesTieBreakerVariableValuesSelected.Size = new System.Drawing.Size(810, 30);
			this.lblFVSVariablesTieBreakerVariableValuesSelected.TabIndex = 2;
			this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = "Not Defined";
			// 
			// btnFVSVariablesTieBreakerVariableClear
			// 
			this.btnFVSVariablesTieBreakerVariableClear.Location = new System.Drawing.Point(24, 376);
			this.btnFVSVariablesTieBreakerVariableClear.Name = "btnFVSVariablesTieBreakerVariableClear";
			this.btnFVSVariablesTieBreakerVariableClear.Size = new System.Drawing.Size(72, 40);
			this.btnFVSVariablesTieBreakerVariableClear.TabIndex = 5;
			this.btnFVSVariablesTieBreakerVariableClear.Text = "Clear";
			this.btnFVSVariablesTieBreakerVariableClear.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableClear_Click);
			// 
			// btnFVSVariablesTieBreakerVariableDone
			// 
			this.btnFVSVariablesTieBreakerVariableDone.Location = new System.Drawing.Point(352, 376);
			this.btnFVSVariablesTieBreakerVariableDone.Name = "btnFVSVariablesTieBreakerVariableDone";
			this.btnFVSVariablesTieBreakerVariableDone.Size = new System.Drawing.Size(88, 40);
			this.btnFVSVariablesTieBreakerVariableDone.TabIndex = 11;
			this.btnFVSVariablesTieBreakerVariableDone.Text = "Done";
			this.btnFVSVariablesTieBreakerVariableDone.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableDone_Click);
			// 
			// btnFVSVariablesTieBreakerVariableCancel
			// 
			this.btnFVSVariablesTieBreakerVariableCancel.Location = new System.Drawing.Point(440, 376);
			this.btnFVSVariablesTieBreakerVariableCancel.Name = "btnFVSVariablesTieBreakerVariableCancel";
			this.btnFVSVariablesTieBreakerVariableCancel.Size = new System.Drawing.Size(88, 40);
			this.btnFVSVariablesTieBreakerVariableCancel.TabIndex = 9;
			this.btnFVSVariablesTieBreakerVariableCancel.Text = "Cancel";
			this.btnFVSVariablesTieBreakerVariableCancel.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableCancel_Click);
			// 
			// btnFVSVariablesTieBreakerVariableNext
			// 
			this.btnFVSVariablesTieBreakerVariableNext.Location = new System.Drawing.Point(616, 376);
			this.btnFVSVariablesTieBreakerVariableNext.Name = "btnFVSVariablesTieBreakerVariableNext";
			this.btnFVSVariablesTieBreakerVariableNext.Size = new System.Drawing.Size(88, 40);
			this.btnFVSVariablesTieBreakerVariableNext.TabIndex = 8;
			this.btnFVSVariablesTieBreakerVariableNext.Text = "Next-->";
			this.btnFVSVariablesTieBreakerVariableNext.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableNext_Click);
			// 
			// grpboxFVSVariablesTieBreaker
			// 
			this.grpboxFVSVariablesTieBreaker.BackColor = System.Drawing.SystemColors.Control;
			this.grpboxFVSVariablesTieBreaker.Controls.Add(this.pnlTieBreaker);
			this.grpboxFVSVariablesTieBreaker.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxFVSVariablesTieBreaker.ForeColor = System.Drawing.Color.Black;
			this.grpboxFVSVariablesTieBreaker.Location = new System.Drawing.Point(16, 56);
			this.grpboxFVSVariablesTieBreaker.Name = "grpboxFVSVariablesTieBreaker";
			this.grpboxFVSVariablesTieBreaker.Size = new System.Drawing.Size(872, 448);
			this.grpboxFVSVariablesTieBreaker.TabIndex = 33;
			this.grpboxFVSVariablesTieBreaker.TabStop = false;
			this.grpboxFVSVariablesTieBreaker.Resize += new System.EventHandler(this.grpboxFVSVariablesTieBreaker_Resize);
			// 
			// pnlTieBreaker
			// 
			this.pnlTieBreaker.AutoScroll = true;
			this.pnlTieBreaker.Controls.Add(this.grpboxFVSVariablesTieBreakerValues);
			this.pnlTieBreaker.Controls.Add(this.groupBox3);
			this.pnlTieBreaker.Controls.Add(this.btnFVSVariablesTieBreakerEdit);
			this.pnlTieBreaker.Dock = System.Windows.Forms.DockStyle.Fill;
			this.pnlTieBreaker.Location = new System.Drawing.Point(3, 18);
			this.pnlTieBreaker.Name = "pnlTieBreaker";
			this.pnlTieBreaker.Size = new System.Drawing.Size(866, 427);
			this.pnlTieBreaker.TabIndex = 70;
			// 
			// grpboxFVSVariablesTieBreakerValues
			// 
			this.grpboxFVSVariablesTieBreakerValues.Controls.Add(this.lvFVSVariablesTieBreakerValues);
			this.grpboxFVSVariablesTieBreakerValues.Location = new System.Drawing.Point(8, 16);
			this.grpboxFVSVariablesTieBreakerValues.Name = "grpboxFVSVariablesTieBreakerValues";
			this.grpboxFVSVariablesTieBreakerValues.Size = new System.Drawing.Size(840, 152);
			this.grpboxFVSVariablesTieBreakerValues.TabIndex = 67;
			this.grpboxFVSVariablesTieBreakerValues.TabStop = false;
			this.grpboxFVSVariablesTieBreakerValues.Text = "Step 1: Define Tie Breakers";
			// 
			// lvFVSVariablesTieBreakerValues
			// 
			this.lvFVSVariablesTieBreakerValues.CheckBoxes = true;
			this.lvFVSVariablesTieBreakerValues.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																											 this.lvColChecked,
																											 this.lvColMethod,
																											 this.lvColFVSVariableName,
																											 this.lvColFieldSource,
																											 this.lvColMinMax});
			this.lvFVSVariablesTieBreakerValues.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lvFVSVariablesTieBreakerValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lvFVSVariablesTieBreakerValues.GridLines = true;
			this.lvFVSVariablesTieBreakerValues.HideSelection = false;
			listViewItem1.Checked = true;
			listViewItem1.StateImageIndex = 1;
			listViewItem2.Checked = true;
			listViewItem2.StateImageIndex = 1;
			this.lvFVSVariablesTieBreakerValues.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
																										   listViewItem1,
																										   listViewItem2});
			this.lvFVSVariablesTieBreakerValues.Location = new System.Drawing.Point(3, 18);
			this.lvFVSVariablesTieBreakerValues.MultiSelect = false;
			this.lvFVSVariablesTieBreakerValues.Name = "lvFVSVariablesTieBreakerValues";
			this.lvFVSVariablesTieBreakerValues.Size = new System.Drawing.Size(834, 131);
			this.lvFVSVariablesTieBreakerValues.TabIndex = 67;
			this.lvFVSVariablesTieBreakerValues.View = System.Windows.Forms.View.Details;
			this.lvFVSVariablesTieBreakerValues.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvFVSVariablesTieBreakerValues_MouseUp);
			this.lvFVSVariablesTieBreakerValues.SelectedIndexChanged += new System.EventHandler(this.lvFVSVariablesTieBreakerValues_SelectedIndexChanged);
			this.lvFVSVariablesTieBreakerValues.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lvFVSVariablesTieBreakerValues_ItemCheck);
			// 
			// lvColChecked
			// 
			this.lvColChecked.Text = "";
			this.lvColChecked.Width = 21;
			// 
			// lvColMethod
			// 
			this.lvColMethod.Text = "Method";
			this.lvColMethod.Width = 176;
			// 
			// lvColFVSVariableName
			// 
			this.lvColFVSVariableName.Text = "FVS Variable";
			this.lvColFVSVariableName.Width = 249;
			// 
			// lvColFieldSource
			// 
			this.lvColFieldSource.Text = "Value Source";
			this.lvColFieldSource.Width = 169;
			// 
			// lvColMinMax
			// 
			this.lvColMinMax.Text = "Max/Min";
			this.lvColMinMax.Width = 166;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.btnFVSVariablesTieBreakerAudit);
			this.groupBox3.Location = new System.Drawing.Point(24, 312);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(832, 72);
			this.groupBox3.TabIndex = 69;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Step 2: Audit";
			// 
			// btnFVSVariablesTieBreakerAudit
			// 
			this.btnFVSVariablesTieBreakerAudit.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFVSVariablesTieBreakerAudit.Location = new System.Drawing.Point(16, 24);
			this.btnFVSVariablesTieBreakerAudit.Name = "btnFVSVariablesTieBreakerAudit";
			this.btnFVSVariablesTieBreakerAudit.Size = new System.Drawing.Size(800, 32);
			this.btnFVSVariablesTieBreakerAudit.TabIndex = 0;
			this.btnFVSVariablesTieBreakerAudit.Text = "Audit";
			this.btnFVSVariablesTieBreakerAudit.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerAudit_Click);
			// 
			// btnFVSVariablesTieBreakerEdit
			// 
			this.btnFVSVariablesTieBreakerEdit.BackColor = System.Drawing.SystemColors.Control;
			this.btnFVSVariablesTieBreakerEdit.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnFVSVariablesTieBreakerEdit.ForeColor = System.Drawing.Color.Black;
			this.btnFVSVariablesTieBreakerEdit.Location = new System.Drawing.Point(376, 176);
			this.btnFVSVariablesTieBreakerEdit.Name = "btnFVSVariablesTieBreakerEdit";
			this.btnFVSVariablesTieBreakerEdit.Size = new System.Drawing.Size(128, 40);
			this.btnFVSVariablesTieBreakerEdit.TabIndex = 36;
			this.btnFVSVariablesTieBreakerEdit.Text = "Edit";
			this.btnFVSVariablesTieBreakerEdit.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerEdit_Click);
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(894, 32);
			this.lblTitle.TabIndex = 27;
			this.lblTitle.Text = "Tie Breaker Settings";
			// 
			// uc_scenario_fvs_prepost_variables_tiebreaker
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_fvs_prepost_variables_tiebreaker";
			this.Size = new System.Drawing.Size(900, 2000);
			this.groupBox1.ResumeLayout(false);
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.grpboxFVSVariablesTieBreakerVariable.ResumeLayout(false);
			this.panel1.ResumeLayout(false);
			this.grpboxFVSVariablesTieBreakerVariableValueSource.ResumeLayout(false);
			this.grpMaxMin.ResumeLayout(false);
			this.grpboxFVSVariablesTieBreakerVariableValues.ResumeLayout(false);
			this.grpFVSVariablesTieBreakerVariableValuesSelected.ResumeLayout(false);
			this.grpboxFVSVariablesTieBreaker.ResumeLayout(false);
			this.pnlTieBreaker.ResumeLayout(false);
			this.grpboxFVSVariablesTieBreakerValues.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
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
            
		}
		

		private void label6_Click(object sender, System.EventArgs e)
		{
		
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
        public void loadvalues_FromProperties()
        {
            this.m_intError = 0;
            this.m_strError = "";

            int x, y;
            

           

            if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0) != null)
            {
                for (x = 0; x <= ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Count - 1; x++)
                {

                    if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper() == "FVS VARIABLE")
                    {
                        //fvs variable name
                        this.m_oOldTieBreakerCollection.Item(0).strFVSVariableName =
                            ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strFVSVariableName.Trim();
                        lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text =
                            this.m_oOldTieBreakerCollection.Item(0).strFVSVariableName;
                        this.lblFVSVariablesTieBreakerVariableValuesSelected.Text =
                            this.m_oOldTieBreakerCollection.Item(0).strFVSVariableName;

                        //fvs value source (POST or POST/PRE change)
                        this.m_oOldTieBreakerCollection.Item(0).strValueSource =
                            ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strValueSource.Trim();
                        lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text =
                            this.m_oOldTieBreakerCollection.Item(0).strValueSource;
                        if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strValueSource.Trim().ToUpper()== "POST")
                            this.cmbFVSVariablesTieBreakerVariableValueSource.Text =
                                 this.cmbFVSVariablesTieBreakerVariableValueSource.Items[0].ToString();
                        else if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strValueSource.Trim().ToUpper() == "POST-PRE")
                            this.cmbFVSVariablesTieBreakerVariableValueSource.Text =
                                 this.cmbFVSVariablesTieBreakerVariableValueSource.Items[1].ToString();


                        //MAX or MIN	
                        if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMaxYN == "Y")
                        {
                            lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text = "MAX";
                            this.m_oOldTieBreakerCollection.Item(0).strMaxYN = "Y";
                            this.m_oOldTieBreakerCollection.Item(0).strMinYN = "N";
                            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked = true;
                        }
                        else
                        {
                            lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text = "MIN";
                            this.m_oOldTieBreakerCollection.Item(0).strMinYN = "Y";
                            this.m_oOldTieBreakerCollection.Item(0).strMaxYN = "N";
                            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked = true;
                        }
                        if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).bSelected)
                        {
                            this.m_oOldTieBreakerCollection.Item(0).bSelected = true;

                        }
                        else
                        {
                            this.m_oOldTieBreakerCollection.Item(0).bSelected = false;
                        }
                        this.lvFVSVariablesTieBreakerValues.Items[0].Checked = this.m_oOldTieBreakerCollection.Item(0).bSelected;
                    }
                    else if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper() == "TREATMENT INTENSITY")
                    {
                        if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).bSelected)
                        {
                            this.m_oOldTieBreakerCollection.Item(1).bSelected = true;

                        }
                        else
                        {
                            this.m_oOldTieBreakerCollection.Item(1).bSelected = false;
                        }
                        this.lvFVSVariablesTieBreakerValues.Items[1].Checked = this.m_oOldTieBreakerCollection.Item(1).bSelected;
                    }

                }
                this.m_oSavTieBreakerCollection.Copy(this.m_oOldTieBreakerCollection, ref this.m_oSavTieBreakerCollection, true);
            }
            this.uc_scenario_treatment_intensity1.loadgrid();

        }	
		public void loadvalues(System.Windows.Forms.ListBox p_oListBox)
		{

			
			this.m_intError=0;
			this.m_strError="";

			int x,y;
			lstFVSVariablesTieBreakerVariableValues.Items.Clear();
			for (x=0;x<=p_oListBox.Items.Count-1;x++)
				lstFVSVariablesTieBreakerVariableValues.Items.Add(p_oListBox.Items[x]);

			this.m_oOldTieBreakerCollection.Clear();

			uc_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem oItem = new TieBreakerItem();
			oItem.intListViewIndex=0;
			oItem.strFVSVariableName=this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
			oItem.strMethod = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim();
			oItem.strValueSource = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim();
			if (lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper()=="MAX")
			{
				oItem.strMaxYN="Y"; oItem.strMinYN="N";
			}
			else if (lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper()=="MIN")
			{
				oItem.strMaxYN="N"; oItem.strMinYN="Y";
			}
		    oItem.bSelected=this.lvFVSVariablesTieBreakerValues.Items[0].Checked;
			this.m_oOldTieBreakerCollection.Add(oItem);

			oItem = new TieBreakerItem();
			oItem.intListViewIndex=1;
			oItem.strFVSVariableName=this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
			oItem.strMethod = this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim();
			oItem.strValueSource = this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim();
			oItem.strMaxYN="N"; oItem.strMinYN="N";
			oItem.bSelected=this.lvFVSVariablesTieBreakerValues.Items[1].Checked;
			this.m_oOldTieBreakerCollection.Add(oItem);


			
			ado_data_access oAdo = new ado_data_access();

			string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
			if (oAdo.m_intError==0)
			{
				int intVarNum=0;

				if (!oAdo.TableExist(oAdo.m_OleDbConnection,Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName))
					 frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesTieBreakerTable(oAdo,oAdo.m_OleDbConnection,
						 Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName);

				oAdo.m_strSQL = "SELECT * FROM " + 
					            Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
					            "WHERE TRIM(scenario_id)='" + strScenarioId.Trim() + "'";


				oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);
				if (oAdo.m_OleDbDataReader.HasRows)
				{
					while (oAdo.m_OleDbDataReader.Read())
					{
 

						if (oAdo.m_OleDbDataReader["tiebreaker_method"] != System.DBNull.Value)
						{
							
							if (oAdo.m_OleDbDataReader["tiebreaker_method"].ToString().Trim().ToUpper()=="FVS VARIABLE")
							{
								//fvs variable name
								this.m_oOldTieBreakerCollection.Item(0).strFVSVariableName=oAdo.m_OleDbDataReader["fvs_variable_name"].ToString().Trim();
								lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text = 
									this.m_oOldTieBreakerCollection.Item(0).strFVSVariableName;
								this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = 
									this.m_oOldTieBreakerCollection.Item(0).strFVSVariableName;

								//fvs value source (POST or POST/PRE change)
								this.m_oOldTieBreakerCollection.Item(0).strValueSource = oAdo.m_OleDbDataReader["value_source"].ToString().Trim();
								lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text = 
									this.m_oOldTieBreakerCollection.Item(0).strValueSource;
								if (oAdo.m_OleDbDataReader["value_source"].ToString().Trim().ToUpper()=="POST")
									this.cmbFVSVariablesTieBreakerVariableValueSource.Text = 
										 this.cmbFVSVariablesTieBreakerVariableValueSource.Items[0].ToString();
								else if (oAdo.m_OleDbDataReader["value_source"].ToString().Trim().ToUpper()=="POST-PRE")
									this.cmbFVSVariablesTieBreakerVariableValueSource.Text = 
										 this.cmbFVSVariablesTieBreakerVariableValueSource.Items[1].ToString();


								//MAX or MIN	
								if (oAdo.m_OleDbDataReader["max_yn"].ToString().Trim().ToUpper()=="Y")
								{
									lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text = "MAX";
									this.m_oOldTieBreakerCollection.Item(0).strMaxYN="Y";
									this.m_oOldTieBreakerCollection.Item(0).strMinYN="N";
									this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked=true;
								}
								else
								{
									lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text = "MIN";
									this.m_oOldTieBreakerCollection.Item(0).strMinYN="Y";
									this.m_oOldTieBreakerCollection.Item(0).strMaxYN="N";
									this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked=true;
								}
								if (oAdo.m_OleDbDataReader["checked_yn"].ToString().Trim().ToUpper()=="Y")
								{
									this.m_oOldTieBreakerCollection.Item(0).bSelected=true;
									
								}
								else
								{
									this.m_oOldTieBreakerCollection.Item(0).bSelected=false;
								}
								this.lvFVSVariablesTieBreakerValues.Items[0].Checked = this.m_oOldTieBreakerCollection.Item(0).bSelected;
							}
							else if (oItem.strMethod.Trim().ToUpper()=="TREATMENT INTENSITY")
							{
								if (oAdo.m_OleDbDataReader["checked_yn"].ToString().Trim().ToUpper()=="Y")
								{
									this.m_oOldTieBreakerCollection.Item(1).bSelected=true;
									
								}
								else
								{
									this.m_oOldTieBreakerCollection.Item(1).bSelected=false;
								}
								this.lvFVSVariablesTieBreakerValues.Items[1].Checked = this.m_oOldTieBreakerCollection.Item(1).bSelected;
							}
						}

					}
				}
				
				oAdo.m_OleDbDataReader.Close();
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
				oAdo.m_OleDbConnection.Dispose();
				this.m_oSavTieBreakerCollection.Copy(this.m_oOldTieBreakerCollection,ref this.m_oSavTieBreakerCollection,true);
					            
			}
			this.m_intError=oAdo.m_intError;
			this.m_strError=oAdo.m_strError;
			oAdo=null;
			
			this.uc_scenario_treatment_intensity1.loadgrid();
			
		}
		public int savevalues()
		{
			int x=0;
			string strColumns="";
			string strValues="";
			ado_data_access oAdo = new ado_data_access();
			string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
			if (oAdo.m_intError==0)
			{
				//delete all records from the scenario fvs variables table
				oAdo.m_strSQL = "DELETE FROM " + Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " + 
					            "WHERE LCASE(TRIM(scenario_id)) = '" + strScenarioId + "';";

				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
				if (oAdo.m_intError < 0)
				{
					oAdo.m_OleDbConnection.Close();
					x=oAdo.m_intError;
					oAdo = null;
					return x;
				}
				strColumns="scenario_id,rxcycle,tiebreaker_method,fvs_variable_name,value_source,max_yn,min_yn,checked_yn";
				//
				//FVS VARIABLE
				//
				//scenario id
				strValues = "'" + strScenarioId + "','1',";
				//method
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim()  + "',";
				this.m_oSavTieBreakerCollection.Item(0).strMethod = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim();
				//fvs variable name
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim() + "',";
				this.m_oSavTieBreakerCollection.Item(0).strFVSVariableName = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
				//value source
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim() + "',";
				m_oSavTieBreakerCollection.Item(0).strValueSource = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim();
				//aggregate MAX or MIN
				if (this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper()=="MAX")
				{
					strValues = strValues + "'Y','N',";
					this.m_oSavTieBreakerCollection.Item(0).strMaxYN="Y";
					this.m_oSavTieBreakerCollection.Item(0).strMinYN="N";
				}
				else
				{
					strValues = strValues + "'N','Y',";
					this.m_oSavTieBreakerCollection.Item(0).strMaxYN="N";
					this.m_oSavTieBreakerCollection.Item(0).strMinYN="Y";
				}
				//checked
				if (this.lvFVSVariablesTieBreakerValues.Items[0].Checked)
				{
					strValues = strValues + "'Y'";
					this.m_oSavTieBreakerCollection.Item(0).bSelected=true;
				}
				else
				{
					strValues = strValues + "'N'";
					this.m_oSavTieBreakerCollection.Item(0).bSelected=false;
				}

				oAdo.m_strSQL = "INSERT INTO " + Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " + 
					            "(" + strColumns + ") VALUES " + 
					            "(" + strValues + ")";

				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

				//
				//TREATMENT INTENSITY
				//
				//scenario id
				strValues = "'" + strScenarioId + "','1',";
				//method
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim()  + "',";
				this.m_oSavTieBreakerCollection.Item(1).strMethod = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim();
				//fvs variable name
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim() + "',";
				this.m_oSavTieBreakerCollection.Item(1).strFVSVariableName = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
				//value source
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim() + "',";
				m_oSavTieBreakerCollection.Item(1).strValueSource = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim();
				//aggregate MAX or MIN
				strValues = strValues + "'N','N',";
				this.m_oSavTieBreakerCollection.Item(1).strMinYN="N";
				this.m_oSavTieBreakerCollection.Item(1).strMaxYN="N";
				//checked
				if (this.lvFVSVariablesTieBreakerValues.Items[1].Checked)
				{
					strValues = strValues + "'Y'";
					this.m_oSavTieBreakerCollection.Item(1).bSelected=true;
				}
				else
				{
					strValues = strValues + "'N'";
					this.m_oSavTieBreakerCollection.Item(1).bSelected=false;
				}

				oAdo.m_strSQL = "INSERT INTO " + Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " + 
					"(" + strColumns + ") VALUES " + 
					"(" + strValues + ")";

				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
			this.uc_scenario_treatment_intensity1.savevalues();
			return 0;
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

		

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesTieBreaker.Height = this.ClientSize.Height - this.grpboxFVSVariablesTieBreaker.Top - 5;
			this.grpboxFVSVariablesTieBreaker.Width = this.ClientSize.Width - (this.grpboxFVSVariablesTieBreaker.Left * 2) ;
			this.grpboxFVSVariablesTieBreakerVariable.Height = this.grpboxFVSVariablesTieBreaker.Height;
			this.grpboxFVSVariablesTieBreakerVariable.Width =  this.grpboxFVSVariablesTieBreakerVariable.Width;
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Height = this.grpboxFVSVariablesTieBreaker.Height;
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Width = this.grpboxFVSVariablesTieBreaker.Width;
			
			//grpboxFVSVariablesPrePost.Height = this.ClientSize.Height - grpboxFVSVariablesPrePost.Top - 5;
			//grpboxFVSVariablesPrePost.Width = this.ClientSize.Width - (grpboxFVSVariablesPrePost.Left * 2) ;
		    //grpboxFVSVariablesPrePostVariable.Height = grpboxFVSVariablesPrePost.Height;
			//grpboxFVSVariablesPrePostVariable.Width =  grpboxFVSVariablesPrePost.Width;
			//grpboxFVSVariablesPrePostExpression.Height = grpboxFVSVariablesPrePost.Height;
			//grpboxFVSVariablesPrePostExpression.Width = grpboxFVSVariablesPrePost.Width;
			
		}

	
		

	

		
		

		
		/// <summary>
		/// Load the Pre-Post variable values
		/// Each pre-post variable set has these variables
		///1. Pre variable name
		///2. Post variable name
		///3. Better SQL 
		///4. Worse  SQL
		///5. Effective SQL
		/// </summary>
		/// <param name="p_intVariable">Current Pre-Post Variable Set</param>
		/// <param name="p_intStep">The variable item within the set to load</param>
		private void loadvalues_variable(int p_intVariable,int p_intStep)
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

		
		private void InitializeUserNavigation()
		{
			for (int x=0;x<=this.m_strUserNavigation.Length - 1;x++)
			{
				this.m_strUserNavigation[x]="";
			}

		}
		private void AddUserNavigation(string p_strMenuItem)
		{

		}

		

		
		private void SaveToCurrentVariables()
		{
			
			
		}
		private void UpdateListViewVariableItem(int p_intListViewItem,int p_intVarArrayItem,uc_scenario_fvs_prepost_variables_effective.Variables p_oVar)
		{
			
		}
		
			

		private void grpboxFVSVariablesPrePostOverallExp_Resize(object sender, System.EventArgs e)
		{
			
		}

		
		private void EnableAllMacroButtons(bool p_bEnable)
		{
			

		}

		private void btnFVSVariablesPrePostExpressionMacroVarPre1_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPre2_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPre3_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPost1_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPost2_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPost3_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVar1Better_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVar2Better_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVar3Better_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVar1Worse_Click(object sender, System.EventArgs e)
		{
  		  
		}

		private void btnFVSVariablesPrePostExpressionMacroVar2Worse_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostExpressionMacroVar3Worse_Click(object sender, System.EventArgs e)
		{
			
		}

		

		
		private void btnFVSVariablesPrePostExpressionMacroVar1Change_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostExpressionMacroVar2Change_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostExpressionMacroVar3Change_Click(object sender, System.EventArgs e)
		{
			
		}

		
	


		private void ShowGroupBox(string p_strName)
		{
			int x;
			
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
		private bool Modified()
		{
			
			return true;
            

		}

		private void btnFVSVariablesPrePostValuesButtonsEdit_Click(object sender, System.EventArgs e)
		{
			EditVariable();

		}
		private void EditVariable()
		{

		}
		

		
		private void btnFVSVariablesPrePostValuesButtonsClear_Click(object sender, System.EventArgs e)
		{
			
		}

		private void RemoveVariable(int p_intIndex)
		{


		}

		private void btnFVSVariablesPrePostValuesButtonsClearAll_Click(object sender, System.EventArgs e)
		{

			DialogResult result = MessageBox.Show("Are you sure you wish to delete all the variables? (YN)","FIA Biosum",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
			if (result == System.Windows.Forms.DialogResult.Yes) 
			{
				for (int x=0;x<=NUMBER_OF_VARIABLES-1;x++)
				{
					RemoveVariable(x);
				}
			}
		}

		
		private void grpboxFVSVariablesPrePost_Resize(object sender, System.EventArgs e)
		{
		}

		private void btnFVSVariablesPrePostAudit_Click(object sender, System.EventArgs e)
		{
			DisplayAuditMessage=true;
			Audit();

			
		}
		public int Audit(bool bDisplayAuditMsg)
		{
			this.DisplayAuditMessage=bDisplayAuditMsg;
			Audit();
			return m_intError;
		}
		public void Audit()
		{
			int x = 0;
			this.m_intError=0;
			this.m_strError="Audit Results \r\n";
			this.m_strError=m_strError + "-------------\r\n\r\n";

            // Only validate treatment intensity if it is checked
            foreach (ListViewItem itemRow in this.lvFVSVariablesTieBreakerValues.CheckedItems)
            {
                if (itemRow.SubItems[COLUMN_METHOD].Text.Trim().Equals("Treatment Intensity"))
                {
                    x=this.uc_scenario_treatment_intensity1.Val_Intensity(false);
                }
            }

            if (x<0)
			{
                if (x == -1) m_strError = m_strError + "No treatments defined\r\n";
                else if (x == -2) m_strError = m_strError + "Treatment intensity ratings must be unique\r\n";
                else if (x == -3) m_strError = m_strError + "Treatment intensity ratings cannot be null in value\r\n";
				m_intError=x;
			}
			if (this.lvFVSVariablesTieBreakerValues.CheckedItems.Count==0)
			{
				m_intError=-1;
				m_strError=m_strError + "No tie breaker methods are checked\r\n";
			}
			else
			{
				this.AuditCheckOptimiztionAndTieBreakerVariable(ref m_intError,ref m_strError);
				
			}
			if (m_intError==0) this.m_strError=m_strError + "Passed Audit";
			else m_strError = m_strError + "\r\n\r\n" + "Failed Audit";

			if (this.DisplayAuditMessage)
				MessageBox.Show(m_strError,"FIA Biosum");
		}
		public void AuditCheckOptimiztionAndTieBreakerVariable(ref int p_intError,ref string p_strError)
		{
			int x;
			if (this.lvFVSVariablesTieBreakerValues.Items[0].Checked)
			{
				if (this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim().ToUpper()=="NOT DEFINED")
				{
					p_intError=-1;
					p_strError=p_strError + "Tie Breaker FVS variable not selected\r\n";
				}
				else
				{
					//cannot be equal to the optimization variable
					for (x=0;x<=this.ReferenceOptimizationUserControl.m_oSavVariableCollection.Count-1;x++)
					{
						if (this.ReferenceOptimizationUserControl.m_oSavVariableCollection.Item(x).bSelected)
						{
							if (this.ReferenceOptimizationUserControl.m_oSavVariableCollection.Item(x).strFVSVariableName.Trim().ToUpper() == 
								lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim().ToUpper())
							{
								p_intError=-1;
								p_strError=p_strError + lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text + " cannot be both optimization variable and tie breaker variable\r\n";
								break;
							}
						}
					}


				}
			}
		}
		
		private void lvFVSVariablesPrePostValues_DoubleClick(object sender, System.EventArgs e)
		{
			EditVariable();
		}

		private void pnlFVSVariablesPrePostExpression_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		private void pnlFVSVariablesPrePostVariable_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		

		private void grpboxFVSVariablesPrePost_VisibleChanged(object sender, System.EventArgs e)
		{

		}


		private void EnableTabs(bool p_bEnable)
		{
            int x;
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlScenario,"tbdesc,tbnotes,tbdatasources",p_bEnable);
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlRules,"tbpsites,tbowners,tbcost,tbtreatmentintensity,tbfilterplots,tbrun",p_bEnable);
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlFVSVariables,"tboptimization,tbeffective",p_bEnable);
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

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnFVSVariablesTieBreakerEdit_Click(object sender, System.EventArgs e)
		{
			if (this.lvFVSVariablesTieBreakerValues.SelectedItems.Count==0) return;

			this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim();

			if (this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim()=="MAX")
				this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked=true;
			else if (this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim()=="MIN")
				this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked=true;
			else
				this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked=true;


			this.grpboxFVSVariablesTieBreaker.Hide();
			this.EnableTabs(false);
			if (this.lvFVSVariablesTieBreakerValues.SelectedItems[0].Index==0)
			{
				this.grpboxFVSVariablesTieBreakerVariable.Show();
			}
			else
			{
				this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Show();
			}
		
		}

		private void grpboxFVSVariablesTieBreaker_Resize(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Top = this.grpboxFVSVariablesTieBreaker.Top;
			this.grpboxFVSVariablesTieBreakerVariable.Top = this.grpboxFVSVariablesTieBreaker.Top;
		}

		private void uc_scenario_treatment_intensity1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void grpboxFVSVariablesTieBreakerTreatmentIntensity_Resize(object sender, System.EventArgs e)
		{
			this.uc_scenario_treatment_intensity1.Width = grpboxFVSVariablesTieBreakerTreatmentIntensity.ClientSize.Width - this.uc_scenario_treatment_intensity1.Left * 2;
			this.uc_scenario_treatment_intensity1.Height =  this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Top - this.uc_scenario_treatment_intensity1.Top;
		}

		private void btnFVSVariablesTieBreakerVariableClear_Click(object sender, System.EventArgs e)
		{
			this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = "Not Defined";
		}

		private void btnFVSVariablesTieBreakerVariableValues_Click(object sender, System.EventArgs e)
		{
			if (this.lstFVSVariablesTieBreakerVariableValues.SelectedItems.Count==0) return;
			this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = this.lstFVSVariablesTieBreakerVariableValues.SelectedItems[0].ToString();
		}

		private void btnFVSVariablesTieBreakerVariableNext_Click(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesTieBreakerVariable.Hide();
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Show();
		}

		private void btnFVSVariablesTieBreakerTreatmentIntensityCancel_Click(object sender, System.EventArgs e)
		{
			this.uc_scenario_treatment_intensity1.m_DataSet.RejectChanges();
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Hide();
			this.EnableTabs(true);
			grpboxFVSVariablesTieBreaker.Show();
		}

		private void btnFVSVariablesTieBreakerTreatmentIntensityClear_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=uc_scenario_treatment_intensity1.m_DataSet.Tables["scenario_rx_intensity"].Rows.Count-1;x++)
				this.uc_scenario_treatment_intensity1.m_DataSet.Tables["scenario_rx_intensity"].Rows[x]["rx_intensity"]=System.DBNull.Value;
		}

		private void btnFVSVariablesTieBreakerVariableCancel_Click(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesTieBreakerVariable.Hide();
			this.EnableTabs(true);
			grpboxFVSVariablesTieBreaker.Show();
		}

		private void btnFVSVariablesTieBreakerVariableDone_Click(object sender, System.EventArgs e)
		{
			this.ReferenceCoreScenarioForm.m_bSave=true;
			UpdateFVSVariableListView();

			this.grpboxFVSVariablesTieBreakerVariable.Hide();
			this.EnableTabs(true);
			this.grpboxFVSVariablesTieBreaker.Show();
		}
		
		private void UpdateFVSVariableListView()
		{
			this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text = this.lblFVSVariablesTieBreakerVariableValuesSelected.Text;
			this.m_oSavTieBreakerCollection.Item(0).strFVSVariableName = this.lblFVSVariablesTieBreakerVariableValuesSelected.Text;

			if (this.cmbFVSVariablesTieBreakerVariableValueSource.Text.Trim().ToUpper() == "POST VALUE")
			{
				this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text = "POST";
			}
			else
			{
			    this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text = "POST-PRE";
			}
			this.m_oSavTieBreakerCollection.Item(0).strValueSource=this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text;


			if (this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked)
			{
				this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text = 
					"MAX";
				this.m_oSavTieBreakerCollection.Item(0).strMaxYN="Y";
				this.m_oSavTieBreakerCollection.Item(0).strMinYN="N";
			}
			else
			{
				this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text = 
					"MIN";
				this.m_oSavTieBreakerCollection.Item(0).strMaxYN="N";
				this.m_oSavTieBreakerCollection.Item(0).strMinYN="Y";

			}
		}

		private void btnFVSVariablesTieBreakerTreatmentIntensityDone_Click(object sender, System.EventArgs e)
		{
			this.ReferenceCoreScenarioForm.m_bSave=true;
			this.uc_scenario_treatment_intensity1.m_DataSet.AcceptChanges();
			this.UpdateFVSVariableListView();

			this.EnableTabs(true);
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Hide();

			this.grpboxFVSVariablesTieBreaker.Show();
		}

		private void lvFVSVariablesTieBreakerValues_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
		   this.ReferenceCoreScenarioForm.m_bSave=true;
		}

		private void btnFVSVariablesTieBreakerAudit_Click(object sender, System.EventArgs e)
		{
			Audit(true);
		}

		private void lvFVSVariablesTieBreakerValues_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = this.lvFVSVariablesTieBreakerValues.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lvFVSVariablesTieBreakerValues.Items[this.lvFVSVariablesTieBreakerValues.TopItem.Index + (int)dblRow-1].Selected=true;
					
				}
			}
			catch 
			{
			}
		}

		private void lvFVSVariablesTieBreakerValues_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lvFVSVariablesTieBreakerValues.SelectedItems.Count > 0)
				this.m_oLvRowColors.DelegateListViewItem(this.lvFVSVariablesTieBreakerValues.SelectedItems[0]);
		}

		private void btnFVSVariablesTieBreakerTreatmentIntensityPrev_Click(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesTieBreakerTreatmentIntensity.Hide();
			this.grpboxFVSVariablesTieBreakerVariable.Show();
			
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
		public FIA_Biosum_Manager.uc_scenario_fvs_prepost_optimization ReferenceOptimizationUserControl
		{
			get {return _uc_optimization;}
			set {_uc_optimization=value;}
		}
	
	}
}
