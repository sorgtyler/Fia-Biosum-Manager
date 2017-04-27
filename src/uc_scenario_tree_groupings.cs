using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_tree_groupings.
	/// </summary>
	public class uc_scenario_tree_groupings : System.Windows.Forms.UserControl
    {
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmCoreScenario _frmScenario=null;
		private FIA_Biosum_Manager.frmProcessorScenario _frmProcessorScenario=null;
		private string _strScenarioType="processor";
        private Button BtnTreeDiameterGroups;
        private FIA_Biosum_Manager.frmDialog m_frmTreeDiam;       //processor tree diameter form
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_scenario_tree_groupings()
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
            this.lblTitle = new System.Windows.Forms.Label();
            this.BtnTreeDiameterGroups = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnTreeDiameterGroups);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(664, 424);
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
            this.lblTitle.Size = new System.Drawing.Size(658, 32);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Tree Groupings";
            // 
            // BtnTreeDiameterGroups
            // 
            this.BtnTreeDiameterGroups.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnTreeDiameterGroups.Location = new System.Drawing.Point(185, 99);
            this.BtnTreeDiameterGroups.Name = "BtnTreeDiameterGroups";
            this.BtnTreeDiameterGroups.Size = new System.Drawing.Size(198, 33);
            this.BtnTreeDiameterGroups.TabIndex = 27;
            this.BtnTreeDiameterGroups.Text = "TREE DIAMETER GROUPS";
            this.BtnTreeDiameterGroups.UseVisualStyleBackColor = true;
            this.BtnTreeDiameterGroups.Click += new System.EventHandler(this.BtnTreeDiameterGroups_Click);
            // 
            // uc_scenario_tree_groupings
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_scenario_tree_groupings";
            this.Size = new System.Drawing.Size(664, 424);
            this.Resize += new System.EventHandler(this.uc_scenario_tree_groupings_Resize);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_scenario_tree_groupings_Resize(object sender, System.EventArgs e)
		{
			
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Visible=false;
			if (this.ScenarioType.Trim().ToUpper() == "CORE") ((frmCoreScenario)this.ParentForm).Height = 0 ; //((frmScenario)this.ParentForm).grpboxMenu.Height * 2;
			else this.ReferenceProcessorScenarioForm.Height=0;
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			this.SaveScenarioNotes();
		}
		public void SaveScenarioNotes()
		{
			ado_data_access p_ado = new ado_data_access();
			System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
            string strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
            string strScenarioDir = strProjDir + "\\" + ScenarioType + "\\db";
            string strFile = "scenario_" + ScenarioType + "_rule_definitions.mdb"; //((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectFile;
			StringBuilder strFullPath = new StringBuilder(strScenarioDir);
			strFullPath.Append("\\");
			strFullPath.Append(strFile);
            string strNotes = "";
			string strSQL="";
			strNotes=p_ado.FixString(strNotes,"'","''");
			string strConn=p_ado.getMDBConnString(strFullPath.ToString(),"admin","");
			//string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFullPath.ToString() + ";User Id=admin;Password=;";
			if (ScenarioType.Trim().ToUpper() == "CORE")
			{
				strSQL = "UPDATE scenario SET notes = '" + 
					strNotes + 
					"' WHERE trim(lcase(scenario_id)) = '" + ((frmCoreScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.Trim().ToLower() + "';";
			}
			else
			{
				strSQL = "UPDATE scenario SET notes = '" + 
					strNotes + 
					"' WHERE trim(lcase(scenario_id)) = '" + this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower() + "';";
			}
			p_ado.SqlNonQuery(strConn,strSQL);
			p_ado=null;
		
		}
		public void LoadValues()
		{
			FIA_Biosum_Manager.ado_data_access oAdo = new ado_data_access();
			string strScenarioDBDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
			string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.mdb"; 
			StringBuilder strScenarioFullPath = new StringBuilder(strScenarioDBDir);
			strScenarioFullPath.Append("\\");
			strScenarioFullPath.Append(strScenarioFile);
			string strScenarioConn = oAdo.getMDBConnString(strScenarioFullPath.ToString(),"admin","");
			oAdo.OpenConnection(strScenarioConn);
			if (ScenarioType.Trim().ToUpper()=="CORE")
			{
				oAdo.m_strSQL = "SELECT notes FROM scenario WHERE TRIM(scenario_id)='" + this.ReferenceCoreScenarioForm.uc_scenario1.strScenarioId.Trim() + "'";
			}
			else
			{
				oAdo.m_strSQL = "SELECT notes FROM scenario WHERE TRIM(scenario_id)='" + this.ReferenceProcessorScenarioForm.uc_scenario1.strScenarioId.Trim() + "'";
			}
			oAdo.m_OleDbConnection.Close();
			while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
			{
				oAdo.m_OleDbConnection.Close();
				System.Threading.Thread.Sleep(1000);
			}
			oAdo=null;
		}

		private void txtNotes_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ScenarioType.Trim().ToUpper()=="CORE") ReferenceCoreScenarioForm.m_bSave=true;
			else ReferenceProcessorScenarioForm.m_bSave=true;
		}
		public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		public FIA_Biosum_Manager.frmProcessorScenario ReferenceProcessorScenarioForm
		{
			get {return _frmProcessorScenario;}
			set {_frmProcessorScenario=value;}
		}
		public string ScenarioType
		{
			get {return _strScenarioType;}
			set {_strScenarioType=value;}
		}

        private void BtnTreeDiameterGroups_Click(object sender, EventArgs e)
        {
                this.m_frmTreeDiam = new frmDialog((frmProcessorScenario)this.ParentForm,(frmMain)this.ParentForm.ParentForm);
                this.m_frmTreeDiam.MaximizeBox = false;
                this.m_frmTreeDiam.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmTreeDiam.Text = "Processor: Tree Diameter Groups";
            // @ToDo: Not sure if we need this
            //this.m_frmTreeDiam.MdiParent = this;
                this.m_frmTreeDiam.Initialize_Plot_Tree_Diam_User_Control();


                this.m_frmTreeDiam.Height = 0;
                this.m_frmTreeDiam.Width = 0;
                if (this.m_frmTreeDiam.uc_tree_diam_groups_list1.Top + this.m_frmTreeDiam.uc_tree_diam_groups_list1.Height > this.m_frmTreeDiam.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmTreeDiam.Height = x;
                        if (this.m_frmTreeDiam.uc_tree_diam_groups_list1.Top +
                            this.m_frmTreeDiam.uc_tree_diam_groups_list1.Height <
                            this.m_frmTreeDiam.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (this.m_frmTreeDiam.uc_tree_diam_groups_list1.Left + this.m_frmTreeDiam.uc_tree_diam_groups_list1.Width > this.m_frmTreeDiam.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmTreeDiam.Width = x;
                        if (this.m_frmTreeDiam.uc_tree_diam_groups_list1.Left +
                            this.m_frmTreeDiam.uc_tree_diam_groups_list1.Width <
                            this.m_frmTreeDiam.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }


                this.m_frmTreeDiam.uc_tree_diam_groups_list1.loadvalues();
                frmMain.g_sbpInfo.Text = "Ready";
                this.m_frmTreeDiam.Show();

                this.m_frmTreeDiam.Left = 0;
                this.m_frmTreeDiam.Top = 0;
                this.m_frmTreeDiam.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            }
        }
}
