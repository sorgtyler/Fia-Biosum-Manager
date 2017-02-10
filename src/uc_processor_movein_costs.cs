using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_escalators.
	/// </summary>
	public class uc_processor_movein_costs : System.Windows.Forms.UserControl
	{
        public int m_intError = 0;
        public string m_strError = "";
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label lblMinYardingDistance;
		private RxTools m_oRxTools = new RxTools();
		private Queries m_oQueries = new Queries();
		private ado_data_access m_oAdo = null;
		private string _strScenarioId="";
        private frmProcessorScenario _frmProcessorScenario = null;
        private TextBox txtMoveInHoursAdjFactor;
        private Label label6;
        private Label label7;
        private TextBox txtMoveInHoursMultiplier;
        private Label label4;
        private Label label5;
        private Label label2;
        private TextBox txtAssumedHarvestArea;
        private Label label3;
        private Label label1;
        private TextBox txtMinYardingDistance;
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

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

		public uc_processor_movein_costs()
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

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblMinYardingDistance = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.txtMinYardingDistance = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtAssumedHarvestArea = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtMoveInHoursMultiplier = new System.Windows.Forms.TextBox();
            this.txtMoveInHoursAdjFactor = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(696, 448);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.txtMoveInHoursAdjFactor);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.txtMoveInHoursMultiplier);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txtAssumedHarvestArea);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txtMinYardingDistance);
            this.panel1.Controls.Add(this.lblMinYardingDistance);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(690, 397);
            this.panel1.TabIndex = 31;
            // 
            // lblMinYardingDistance
            // 
            this.lblMinYardingDistance.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMinYardingDistance.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblMinYardingDistance.Location = new System.Drawing.Point(16, 8);
            this.lblMinYardingDistance.Name = "lblMinYardingDistance";
            this.lblMinYardingDistance.Size = new System.Drawing.Size(240, 16);
            this.lblMinYardingDistance.TabIndex = 52;
            this.lblMinYardingDistance.Text = "Minimum Yarding Distance (Feet):";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(690, 32);
            this.lblTitle.TabIndex = 30;
            this.lblTitle.Text = "Move-in costs";
            // 
            // txtMinYardingDistance
            // 
            this.txtMinYardingDistance.ForeColor = System.Drawing.Color.Black;
            this.txtMinYardingDistance.Location = new System.Drawing.Point(273, 7);
            this.txtMinYardingDistance.Name = "txtMinYardingDistance";
            this.txtMinYardingDistance.Size = new System.Drawing.Size(45, 20);
            this.txtMinYardingDistance.TabIndex = 55;
            this.txtMinYardingDistance.Text = "150.0";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label1.Location = new System.Drawing.Point(16, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(402, 16);
            this.label1.TabIndex = 56;
            this.label1.Text = "The shortest yarding distance sent to OpCost for a plot";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(16, 93);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(402, 16);
            this.label2.TabIndex = 59;
            this.label2.Text = "Additional information?";
            // 
            // txtAssumedHarvestArea
            // 
            this.txtAssumedHarvestArea.ForeColor = System.Drawing.Color.Black;
            this.txtAssumedHarvestArea.Location = new System.Drawing.Point(273, 72);
            this.txtAssumedHarvestArea.Name = "txtAssumedHarvestArea";
            this.txtAssumedHarvestArea.Size = new System.Drawing.Size(45, 20);
            this.txtAssumedHarvestArea.TabIndex = 58;
            this.txtAssumedHarvestArea.Text = "80.0";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label3.Location = new System.Drawing.Point(16, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(259, 16);
            this.label3.TabIndex = 57;
            this.label3.Text = "Assumed Harvest Area Size (Acres):";
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label4.Location = new System.Drawing.Point(16, 155);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(486, 36);
            this.label4.TabIndex = 61;
            this.label4.Text = "Multiplied against GIS travel time for a plot; A value of 0 disconnects move-in c" +
    "ost calculations from GIS travel time";
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label5.Location = new System.Drawing.Point(16, 135);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(325, 16);
            this.label5.TabIndex = 60;
            this.label5.Text = "Move-in Hours Multiplier:";
            // 
            // txtMoveInHoursMultiplier
            // 
            this.txtMoveInHoursMultiplier.ForeColor = System.Drawing.Color.Black;
            this.txtMoveInHoursMultiplier.Location = new System.Drawing.Point(273, 134);
            this.txtMoveInHoursMultiplier.Name = "txtMoveInHoursMultiplier";
            this.txtMoveInHoursMultiplier.Size = new System.Drawing.Size(45, 20);
            this.txtMoveInHoursMultiplier.TabIndex = 62;
            this.txtMoveInHoursMultiplier.Text = "1.0";
            // 
            // txtMoveInHoursAdjFactor
            // 
            this.txtMoveInHoursAdjFactor.ForeColor = System.Drawing.Color.Black;
            this.txtMoveInHoursAdjFactor.Location = new System.Drawing.Point(392, 214);
            this.txtMoveInHoursAdjFactor.Name = "txtMoveInHoursAdjFactor";
            this.txtMoveInHoursAdjFactor.Size = new System.Drawing.Size(45, 20);
            this.txtMoveInHoursAdjFactor.TabIndex = 65;
            this.txtMoveInHoursAdjFactor.Text = "0.0";
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label6.Location = new System.Drawing.Point(16, 235);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(486, 36);
            this.label6.TabIndex = 64;
            this.label6.Text = "Adds a fixed value to GIS travel time for a plot; Value can be positive or negati" +
    "ve.";
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label7.Location = new System.Drawing.Point(16, 215);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(380, 16);
            this.label7.TabIndex = 63;
            this.label7.Text = "Move-in Hours Adjustment Factor (Decimal Seconds):";
            // 
            // uc_processor_movein_costs
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_movein_costs";
            this.Size = new System.Drawing.Size(696, 448);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion
		public void loadvalues()
		{
			int x,y;
			string strField;
			if (m_oAdo!=null && m_oAdo.m_OleDbConnection != null)
				if (m_oAdo.m_OleDbConnection.State == System.Data.ConnectionState.Open) m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

			//
			//SCENARIO MDB
			//
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +  
				"\\processor\\db\\scenario_processor_rule_definitions.mdb";
			//
			//SCENARIO ID
			//
			ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			//
			//LOAD PROJECT DATATASOURCES INFO
			//
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.m_oReference.LoadDatasource=true;
			m_oQueries.LoadDatasources(true,"processor",ScenarioId);
			//
			//CREATE LINK IN TEMP MDB TO SCENARIO COST REVENUE ESCALATORS TABLE
			//
			dao_data_access oDao = new dao_data_access();
			//link to tree species groups table
			oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
				"scenario_cost_revenue_escalators",
				strScenarioMDB,"scenario_cost_revenue_escalators",true);
			oDao.m_DaoWorkspace.Close();
			oDao=null;
			//
			//OPEN CONNECTION TO TEMP DB FILE
			//
			m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));                
			
		}
        public void loadvalues_FromProperties()
        {

            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oEscalators != null)
            {
                ProcessorScenarioItem.Escalators oEscalators = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oEscalators;
                //
                //UPDATE CYCLE ESCALATOR TEXT BOXES
                //
                //operating costs cycle 2,3,4
                //cycle2
                
            }
 
        }

		public void savevalues()
		{
            m_intError = 0;
            m_strError = "";

			int x;
			string strValues="";

			
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
