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
	/// Summary description for uc_scenario_filter.
	/// </summary>
	public class uc_scenario_filter : System.Windows.Forms.UserControl
	{
		private env m_oEnv;
		private System.Windows.Forms.ImageList imgSize;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;
		public int m_intFullHt = 400;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		public ado_core_tables m_ado_core_tables;
		private string[] m_strCoreTables;
		private string m_strCurrentSQL;
		private string m_strCurrentYardDist;
		private string m_strCurrentYardDist2;
		public System.Windows.Forms.TextBox txtCurrentSQL;
		private System.Windows.Forms.Button btnCreateSQL;
		private System.Windows.Forms.Button btnExecuteSQL;
		private System.Windows.Forms.Button btnPrevSQL;
		private System.Windows.Forms.Button btnEditSQL;
		private System.Windows.Forms.TextBox txtYardDist;
		private System.Windows.Forms.GroupBox grpboxSQL;
		private System.Windows.Forms.GroupBox grpboxYardDist;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtYardDist2;
		private System.Windows.Forms.Button btnYardDistDefault;
		private int m_intNumberOfCoreTablesLoadedIntoDatasets;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Button btnAudit;
		private FIA_Biosum_Manager.frmScenario _frmScenario=null;
		

		public uc_scenario_filter()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_oEnv = new env();
			this.grpboxSQL.Left = 4;
			this.grpboxYardDist.Left = 4;
			this.txtCurrentSQL.Left = 4;
			this.m_strCoreTables = new string[50];
			this.m_strCurrentSQL = "";
			this.m_strCurrentYardDist="";
			this.m_strCurrentYardDist2="";
			this.m_intNumberOfCoreTablesLoadedIntoDatasets=0;

			
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_scenario_filter));
			this.imgSize = new System.Windows.Forms.ImageList(this.components);
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblTitle = new System.Windows.Forms.Label();
			this.grpboxYardDist = new System.Windows.Forms.GroupBox();
			this.btnYardDistDefault = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.txtYardDist2 = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.txtYardDist = new System.Windows.Forms.TextBox();
			this.grpboxSQL = new System.Windows.Forms.GroupBox();
			this.txtCurrentSQL = new System.Windows.Forms.TextBox();
			this.btnPrevSQL = new System.Windows.Forms.Button();
			this.btnCreateSQL = new System.Windows.Forms.Button();
			this.btnEditSQL = new System.Windows.Forms.Button();
			this.btnExecuteSQL = new System.Windows.Forms.Button();
			this.btnAudit = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.grpboxYardDist.SuspendLayout();
			this.grpboxSQL.SuspendLayout();
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
			this.groupBox1.Controls.Add(this.grpboxYardDist);
			this.groupBox1.Controls.Add(this.grpboxSQL);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(656, 408);
			this.groupBox1.TabIndex = 18;
			this.groupBox1.TabStop = false;
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			this.groupBox1.Paint += new System.Windows.Forms.PaintEventHandler(this.groupBox1_Paint);
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(650, 32);
			this.lblTitle.TabIndex = 29;
			this.lblTitle.Text = "Plot Filter";
			// 
			// grpboxYardDist
			// 
			this.grpboxYardDist.Controls.Add(this.btnYardDistDefault);
			this.grpboxYardDist.Controls.Add(this.label2);
			this.grpboxYardDist.Controls.Add(this.txtYardDist2);
			this.grpboxYardDist.Controls.Add(this.label1);
			this.grpboxYardDist.Controls.Add(this.txtYardDist);
			this.grpboxYardDist.ForeColor = System.Drawing.Color.Black;
			this.grpboxYardDist.Location = new System.Drawing.Point(24, 56);
			this.grpboxYardDist.Name = "grpboxYardDist";
			this.grpboxYardDist.Size = new System.Drawing.Size(256, 96);
			this.grpboxYardDist.TabIndex = 12;
			this.grpboxYardDist.TabStop = false;
			this.grpboxYardDist.Text = "Maximum Yarding Distance Allowed (Meters)";
			// 
			// btnYardDistDefault
			// 
			this.btnYardDistDefault.BackColor = System.Drawing.SystemColors.Control;
			this.btnYardDistDefault.ForeColor = System.Drawing.Color.Black;
			this.btnYardDistDefault.Location = new System.Drawing.Point(8, 64);
			this.btnYardDistDefault.Name = "btnYardDistDefault";
			this.btnYardDistDefault.Size = new System.Drawing.Size(104, 24);
			this.btnYardDistDefault.TabIndex = 17;
			this.btnYardDistDefault.Text = "Default Values";
			this.btnYardDistDefault.Click += new System.EventHandler(this.btnYardDistDefault_Click);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(128, 20);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(76, 16);
			this.label2.TabIndex = 16;
			this.label2.Text = "Slope > 40";
			// 
			// txtYardDist2
			// 
			this.txtYardDist2.Location = new System.Drawing.Point(124, 36);
			this.txtYardDist2.Name = "txtYardDist2";
			this.txtYardDist2.Size = new System.Drawing.Size(104, 20);
			this.txtYardDist2.TabIndex = 15;
			this.txtYardDist2.Text = "";
			this.txtYardDist2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYardDist2_KeyPress);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(12, 21);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(76, 16);
			this.label1.TabIndex = 14;
			this.label1.Text = "Slope < 40";
			// 
			// txtYardDist
			// 
			this.txtYardDist.Location = new System.Drawing.Point(8, 37);
			this.txtYardDist.Name = "txtYardDist";
			this.txtYardDist.Size = new System.Drawing.Size(104, 20);
			this.txtYardDist.TabIndex = 13;
			this.txtYardDist.Text = "";
			this.txtYardDist.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYardDist_KeyPress);
			// 
			// grpboxSQL
			// 
			this.grpboxSQL.Controls.Add(this.btnAudit);
			this.grpboxSQL.Controls.Add(this.txtCurrentSQL);
			this.grpboxSQL.Controls.Add(this.btnPrevSQL);
			this.grpboxSQL.Controls.Add(this.btnCreateSQL);
			this.grpboxSQL.Controls.Add(this.btnEditSQL);
			this.grpboxSQL.Controls.Add(this.btnExecuteSQL);
			this.grpboxSQL.ForeColor = System.Drawing.Color.Black;
			this.grpboxSQL.Location = new System.Drawing.Point(24, 160);
			this.grpboxSQL.Name = "grpboxSQL";
			this.grpboxSQL.Size = new System.Drawing.Size(600, 216);
			this.grpboxSQL.TabIndex = 11;
			this.grpboxSQL.TabStop = false;
			this.grpboxSQL.Text = "Current Plot SQL";
			// 
			// txtCurrentSQL
			// 
			this.txtCurrentSQL.Location = new System.Drawing.Point(16, 24);
			this.txtCurrentSQL.Multiline = true;
			this.txtCurrentSQL.Name = "txtCurrentSQL";
			this.txtCurrentSQL.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtCurrentSQL.Size = new System.Drawing.Size(568, 152);
			this.txtCurrentSQL.TabIndex = 10;
			this.txtCurrentSQL.Text = "";
			this.txtCurrentSQL.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCurrentSQL_KeyPress);
			// 
			// btnPrevSQL
			// 
			this.btnPrevSQL.BackColor = System.Drawing.SystemColors.Control;
			this.btnPrevSQL.ForeColor = System.Drawing.Color.Black;
			this.btnPrevSQL.Location = new System.Drawing.Point(344, 188);
			this.btnPrevSQL.Name = "btnPrevSQL";
			this.btnPrevSQL.Size = new System.Drawing.Size(120, 24);
			this.btnPrevSQL.TabIndex = 6;
			this.btnPrevSQL.Text = "View Previous SQL";
			this.btnPrevSQL.Click += new System.EventHandler(this.btnPrevSQL_Click);
			// 
			// btnCreateSQL
			// 
			this.btnCreateSQL.BackColor = System.Drawing.SystemColors.Control;
			this.btnCreateSQL.ForeColor = System.Drawing.Color.Black;
			this.btnCreateSQL.Location = new System.Drawing.Point(464, 188);
			this.btnCreateSQL.Name = "btnCreateSQL";
			this.btnCreateSQL.Size = new System.Drawing.Size(120, 24);
			this.btnCreateSQL.TabIndex = 8;
			this.btnCreateSQL.Text = "Create New SQL";
			this.btnCreateSQL.Click += new System.EventHandler(this.btnCreateSQL_Click);
			// 
			// btnEditSQL
			// 
			this.btnEditSQL.BackColor = System.Drawing.SystemColors.Control;
			this.btnEditSQL.ForeColor = System.Drawing.Color.Black;
			this.btnEditSQL.Location = new System.Drawing.Point(104, 188);
			this.btnEditSQL.Name = "btnEditSQL";
			this.btnEditSQL.Size = new System.Drawing.Size(120, 24);
			this.btnEditSQL.TabIndex = 5;
			this.btnEditSQL.Text = "Edit Current SQL";
			this.btnEditSQL.Click += new System.EventHandler(this.btnEditSQL_Click);
			// 
			// btnExecuteSQL
			// 
			this.btnExecuteSQL.BackColor = System.Drawing.SystemColors.Control;
			this.btnExecuteSQL.ForeColor = System.Drawing.Color.Black;
			this.btnExecuteSQL.Location = new System.Drawing.Point(16, 188);
			this.btnExecuteSQL.Name = "btnExecuteSQL";
			this.btnExecuteSQL.Size = new System.Drawing.Size(88, 24);
			this.btnExecuteSQL.TabIndex = 7;
			this.btnExecuteSQL.Text = "Execute SQL";
			this.btnExecuteSQL.Click += new System.EventHandler(this.btnExecuteSQL_Click);
			// 
			// btnAudit
			// 
			this.btnAudit.BackColor = System.Drawing.SystemColors.Control;
			this.btnAudit.ForeColor = System.Drawing.Color.Black;
			this.btnAudit.Location = new System.Drawing.Point(224, 188);
			this.btnAudit.Name = "btnAudit";
			this.btnAudit.Size = new System.Drawing.Size(120, 24);
			this.btnAudit.TabIndex = 11;
			this.btnAudit.Text = "Audit";
			this.btnAudit.Click += new System.EventHandler(this.btnAudit_Click);
			// 
			// uc_scenario_filter
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_filter";
			this.Size = new System.Drawing.Size(656, 408);
			this.groupBox1.ResumeLayout(false);
			this.grpboxYardDist.ResumeLayout(false);
			this.grpboxSQL.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void groupBox1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{

		}

		private void grpboxFilter_Resize(object sender, System.EventArgs e)
		{
			try
			{
			}
			catch
			{
			}
		}

		public void loadvalues()
		{
			string strSQL="";
			string strConn="";
              
			ado_data_access p_ado = new ado_data_access();

			string strScenarioId = ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            
			//scenario mdb connection
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				p_ado = null;
				return;
			}
			strSQL = "SELECT * FROM scenario_plot_filter WHERE " + 
				" scenario_id = '" + strScenarioId + "' AND current_yn = 'Y';";
			p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);

				  
			
			if (p_ado.m_intError == 0)
			{
				this.txtCurrentSQL.Text ="";
				
				//load  wind speed class definitions
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["sql_command"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["sql_command"].ToString().Trim().Length > 0)
						{
							this.txtCurrentSQL.Text = p_ado.m_OleDbDataReader["sql_command"].ToString().Trim();
							this.m_strCurrentSQL = this.txtCurrentSQL.Text;
						}
					}

				}
				
				p_ado.m_OleDbDataReader.Close();
			}


			strSQL = "SELECT * FROM scenario_plot_filter_misc WHERE " + 
				" scenario_id = '" + strScenarioId + "';";
			p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);

				  
			
			if (p_ado.m_intError == 0)
			{
				this.txtYardDist.Text ="";
				this.txtYardDist2.Text = "";
				
				//load  wind speed class definitions
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["yard_dist"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["yard_dist"].ToString().Trim().Length > 0)
						{
							this.txtYardDist.Text = Convert.ToString(p_ado.m_OleDbDataReader["yard_dist"].ToString().Trim());
							this.txtYardDist2.Text = Convert.ToString(p_ado.m_OleDbDataReader["yard_dist2"].ToString().Trim());
							this.m_strCurrentYardDist = this.txtYardDist.Text;
							this.m_strCurrentYardDist2 = this.txtYardDist2.Text;
						}
					}
				}
				p_ado.m_OleDbDataReader.Close();
			}

			p_ado.m_OleDbDataReader = null;
			p_ado.m_OleDbCommand = null;



				
			

			
			p_ado = null;
			this.m_OleDbConnectionScenario.Close();
			((frmScenario)this.ParentForm).m_bSave=false;

		}
		public int savevalues()
		{
			int x=0;
            
			string str="";
			string strSQL = "";
			string strConn = "";
			string strNewSQL="";
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			

			if (this.m_strCurrentSQL.Trim().ToUpper() == this.txtCurrentSQL.Text.Trim().ToUpper() && 
				this.m_strCurrentYardDist.Trim() == this.txtYardDist.Text.Trim() &&
				this.m_strCurrentYardDist2.Trim() == this.txtYardDist2.Text.Trim())
			     return 0;

			ado_data_access p_ado = new ado_data_access();
			string strScenarioId = ((frmScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				x=p_ado.m_intError;
				p_ado=null;
				return x;
			}


			if (this.m_strCurrentSQL.Trim().ToUpper() != this.txtCurrentSQL.Text.Trim().ToUpper())
			{
				strSQL = "SELECT * FROM scenario_plot_filter WHERE " + 
					" scenario_id = '" + strScenarioId + "' AND current_yn = 'Y';";
				p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);

				  
			
				if (p_ado.m_intError == 0)
				{
					//update the current sql as NOT being current
					if (p_ado.m_OleDbDataReader.HasRows == true)
					{
						p_ado.m_OleDbDataReader.Close();
						strSQL = "UPDATE scenario_plot_filter SET current_yn = 'N'" + 
							" WHERE scenario_id = '" + strScenarioId + "' AND current_yn = 'Y';";
						p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);

					}
					else 
					{
						p_ado.m_OleDbDataReader.Close();
					}
				
					for (x = 0; x <= this.m_intNumberOfCoreTablesLoadedIntoDatasets - 1; x++)
					{
						if (this.m_strCoreTables[x].Trim().Length > 0)
							str = str + this.m_strCoreTables[x] + ",";
					}
					//remove the last comma
					if (str.Trim().Length > 0)	str = str.Substring(0, str.Length-1);

					strNewSQL = p_ado.FixString(this.txtCurrentSQL.Text,"'","''");
					sb.Append("INSERT INTO scenario_plot_filter (scenario_id,sql_command,current_yn,table_list) Values('");
					sb.Append(strScenarioId);
					sb.Append("','");
					sb.Append(strNewSQL);
					sb.Append("','Y','");
					sb.Append(str);
					sb.Append("');");
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,sb.ToString());
				}
			}

			if (this.m_strCurrentYardDist.Trim() != this.txtYardDist.Text.Trim() ||
				this.m_strCurrentYardDist2.Trim() != this.txtYardDist2.Text.Trim())
			{
				//delete all records from the scenario wind speed class table
				strSQL = "DELETE FROM scenario_plot_filter_misc WHERE " + 
					" scenario_id = '" + strScenarioId + "';";

				p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				if (p_ado.m_intError < 0)
				{
					this.m_OleDbConnectionScenario.Close();
					x=p_ado.m_intError;
					p_ado = null;
					return x;
				}
				if (this.txtYardDist.Text.Trim().Length == 0)
				{
				  this.txtYardDist.Text = "1524";
				}
				if (this.txtYardDist2.Text.Trim().Length == 0)
				{
					this.txtYardDist2.Text = "610";
				}
			
				strSQL = "INSERT INTO scenario_plot_filter_misc (scenario_id,yard_dist,yard_dist2)" + 
					" VALUES ('" + strScenarioId + "'," + 
					this.txtYardDist.Text.Trim() + "," + this.txtYardDist2.Text.Trim() + ");";
				p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
				if (p_ado.m_intError < 0)
				{
					this.m_OleDbConnectionScenario.Close();
					x=p_ado.m_intError;
					p_ado = null;
					return x;
				}
			}
			this.m_OleDbConnectionScenario.Close();
			p_ado=null;
			return 0;

		
		}
		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.grpboxSQL.Width = this.groupBox1.Width - 8;
				this.txtCurrentSQL.Width = this.grpboxSQL.Width - 8;
				this.btnEditSQL.Left = (int)(this.grpboxSQL.Width * .50) - (int)(this.btnEditSQL.Width*2); //* .50);
				this.btnExecuteSQL.Left = this.btnEditSQL.Left - this.btnExecuteSQL.Width;
				this.btnAudit.Left = this.btnEditSQL.Left + this.btnAudit.Width;
				this.btnPrevSQL.Left = this.btnAudit.Left + this.btnEditSQL.Width;
				this.btnCreateSQL.Left = this.btnPrevSQL.Left + this.btnPrevSQL.Width;
				this.grpboxSQL.Height = this.groupBox1.Height - this.grpboxSQL.Top;
				this.btnEditSQL.Top = this.grpboxSQL.Height - this.btnEditSQL.Height - 3;
				this.btnExecuteSQL.Top = this.btnEditSQL.Top;
				this.btnPrevSQL.Top = this.btnEditSQL.Top;
				this.btnCreateSQL.Top= this.btnEditSQL.Top;
				this.btnAudit.Top = this.btnEditSQL.Top;
				this.txtCurrentSQL.Height = this.grpboxSQL.Height - this.txtCurrentSQL.Top - this.btnEditSQL.Height - 10;
			}
			catch
			{
			}
		}

		private void btnCreateSQL_Click(object sender, System.EventArgs e)
		{
			int x=0;
			int y=0;
			DialogResult result;
              
			x = ((frmScenario)this.ParentForm).uc_datasource1.getDataSourceTableNameRow("PLOT");
			if (x < 0) 
			{
				MessageBox.Show("!!Plot table cannot be found!!");
				return;
			}

				frmDialog frmTemp = new frmDialog();
				frmTemp.Text = "Core Analysis: Select Database Tables For SQL";
                frmTemp.uc_select_list_item1.lblMsg.Enabled=true;
				frmTemp.uc_select_list_item1.lblMsg.Text= "Data Sources";
				frmTemp.uc_select_list_item1.lblMsg.Visible = true;
				frmTemp.uc_select_list_item1.Dock = System.Windows.Forms.DockStyle.Fill;
						
				frmTemp.uc_project1.Visible=false;
				frmTemp.uc_select_list_item1.listBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
				frmTemp.uc_select_list_item1.listBox1.Items.Clear();

				((frmScenario)this.ParentForm).uc_datasource1.LoadScenarioDataSourceTablesIntoListBox(frmTemp.uc_select_list_item1.listBox1);


				frmTemp.uc_select_list_item1.Visible=true;
				result = frmTemp.ShowDialog(this);
                this.m_ado_core_tables = new ado_core_tables();      
				dao_data_access p_dao = new dao_data_access();
						
				if (result == DialogResult.OK) 
				{
					frmDialog frmSQL = new frmDialog((frmScenario)this.ParentForm,(frmMain)this.ParentForm.ParentForm);
					
					frmSQL.Width = frmSQL.uc_sql_builder1.m_intFullWd;
					frmSQL.Height = frmSQL.uc_sql_builder1.m_intFullHt;

					frmSQL.Text = "Core Analysis: SQL Builder";
					
					//send current sql to sql builder text box
					frmSQL.uc_sql_builder1.SendSingleKeyStrokes(frmSQL.uc_sql_builder1.txtSQLCommand,this.txtCurrentSQL.Text);
					
					frmSQL.uc_sql_builder1.Visible=true;
					for (x=0; x<=frmTemp.uc_select_list_item1.listBox1.SelectedItems.Count-1;x++)
					{
						for (y=0; y <= frmSQL.uc_sql_builder1.lstTables.Items.Count-1;y++)
						{
							if (frmSQL.uc_sql_builder1.lstTables.Items[y].ToString().Trim().ToUpper() ==
                                 frmTemp.uc_select_list_item1.listBox1.SelectedItems[x].ToString().Trim().ToUpper())
								   break;

						}
						if (y >= frmSQL.uc_sql_builder1.lstTables.Items.Count)
					        frmSQL.uc_sql_builder1.lstTables.Items.Add(frmTemp.uc_select_list_item1.listBox1.SelectedItems[x]);
					}
					frmSQL.uc_sql_builder1.InitializeDataSource(((frmScenario)this.ParentForm).uc_datasource1.lstRequiredTables);

					result = frmSQL.ShowDialog(this);
					if (result == DialogResult.OK)
					{
						if (frmSQL.uc_sql_builder1.txtSQLCommand.Text.Trim().Length > 0)
						{
							result = MessageBox.Show("REPLACE \n" + "-------\n" + "'" + this.txtCurrentSQL.Text + "' \n WITH  \n----\n'" + frmSQL.uc_sql_builder1.txtSQLCommand.Text + "'", "Plot Filter",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
							if (result == DialogResult.Yes)
							{
								for(x=0;x <= ((frmScenario)this.ParentForm).uc_datasource1.lstRequiredTables.Items.Count-1;x++)
								{
									this.m_ado_core_tables.m_strCoreTables[x] = "";
								}

								for(x=0;x <= frmSQL.uc_sql_builder1.lstTables.Items.Count-1;x++)
								{
									this.m_strCoreTables[x] = frmSQL.uc_sql_builder1.lstTables.Items[x].ToString();
								}
								this.m_intNumberOfCoreTablesLoadedIntoDatasets = frmSQL.uc_sql_builder1.m_intNumberOfTablesLoadedIntoDatasets;
								this.txtCurrentSQL.Text = frmSQL.uc_sql_builder1.txtSQLCommand.Text;
								((frmScenario)this.ParentForm).m_bSave=true;

							}
							
                            

						}
					}
					frmSQL.Close();
					frmSQL = null;
				}
					
				this.m_ado_core_tables = null;
                p_dao = null;

				frmTemp.Close();
				frmTemp = null;

    			if (this.m_OleDbConnectionScenario != null && this.m_OleDbConnectionScenario.State == System.Data.ConnectionState.Open)
				{
					this.m_OleDbConnectionScenario.Close();
				}

		}

		/****************************************************************
		 ** edit the sql in sql builder form
		 ****************************************************************/
		private void btnEditSQL_Click(object sender, System.EventArgs e)
		{
			string strSQL="";

			string strConn="";
			string str="";
			int x=0;
			int y=0;
			DialogResult result;
            
  
			x = this.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableNameRow("PLOT");
			if (x < 0) 
			{
				MessageBox.Show("!!Plot table cannot be found!!");
				return;
			}


			this.m_ado_core_tables = new ado_core_tables();      
			dao_data_access p_dao = new dao_data_access();

			frmSqlBuilder frmSQL = new frmSqlBuilder(ReferenceCoreScenarioForm.uc_datasource1.lstRequiredTables,"");
			
			frmSQL.ClientId="CORE SCENARIO PLOT FILTER";
			frmSQL.ReferenceCoreScenarioForm=this.ReferenceCoreScenarioForm;
			frmSQL.txtSQLCommand.Text = this.txtCurrentSQL.Text;

			result = frmSQL.ShowDialog(this);
			//frmSQL.uc_sql_builder2.frmGridView1.Close();
			if (result == DialogResult.OK)
			{
				if (frmSQL.txtSQLCommand.Text.Trim().Length > 0)
				{
					result = MessageBox.Show("REPLACE \n" + "----------------\n\n" + "'" + this.txtCurrentSQL.Text + "' \n\n\n WITH  \n----------------\n\n'" + frmSQL.txtSQLCommand.Text + "'", "Plot Filter",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
					if (result == DialogResult.Yes)
					{
						for(x=0;x <= ReferenceCoreScenarioForm.uc_datasource1.lstRequiredTables.Items.Count-1;x++)
						{
							 this.m_ado_core_tables.m_strCoreTables[x] = "";
						}

						for(x=0;x <= frmSQL.lstTables.Items.Count-1;x++)
						{
						   this.m_strCoreTables[x] = frmSQL.lstTables.Items[x].ToString();
						}

						this.txtCurrentSQL.Text = frmSQL.txtSQLCommand.Text;
						((frmScenario)this.ParentForm).m_bSave=true;
					}
				}
			}
			frmSQL.Close();
			frmSQL = null;
					
			this.m_ado_core_tables = null;
			p_dao = null;

		}

		/****************************************************************
		 ** edit the sql in sql builder form
		 ****************************************************************/
		private void btnEditSQL_Click_old(object sender, System.EventArgs e)
		{
			string strSQL="";

			string strConn="";
			string str="";
			int x=0;
			int y=0;
			DialogResult result;
            
  
			x = this.ReferenceCoreScenarioForm.uc_datasource1.getDataSourceTableNameRow("PLOT");
			if (x < 0) 
			{
				MessageBox.Show("!!Plot table cannot be found!!");
				return;
			}


			this.m_ado_core_tables = new ado_core_tables();      
			dao_data_access p_dao = new dao_data_access();

			
			
			frmDialog frmSQL = new frmDialog(this.ReferenceCoreScenarioForm,frmMain.g_oFrmMain);
			frmSQL.Width = frmSQL.uc_sql_builder1.m_intFullWd;
			frmSQL.Height = frmSQL.uc_sql_builder1.btnClipboardCopy.Top + (frmSQL.uc_sql_builder1.btnClipboardCopy.Height * 3);
			frmSQL.Text = "Core Analysis: SQL Builder";
					
			/*******************************************
			 **send current sql to sql builder text box
			 *******************************************/
			frmSQL.uc_sql_builder1.txtSQLCommand.Text = this.txtCurrentSQL.Text;	
			frmSQL.uc_sql_builder1.Visible=true;
			frmSQL.uc_sql_builder1.txtSQLCommand.Focus();
			frmSQL.uc_sql_builder1.txtSQLCommand.Select(frmSQL.uc_sql_builder1.txtSQLCommand.Text.Length,0);
			frmSQL.uc_sql_builder1.SendSingleKeyStrokes(frmSQL.uc_sql_builder1.txtSQLCommand," ");            
			ado_data_access p_ado = new ado_data_access();
			frmSQL.uc_sql_builder1.lstTables.Items.Add(ReferenceCoreScenarioForm.uc_datasource1.lstRequiredTables.Items[x].SubItems[FIA_Biosum_Manager.Datasource.TABLE].Text.Trim());

			/*********************************************************
			 **get all the table names that are associated
			 **with the current sql
			 *********************************************************/
			string strScenarioId = ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			//scenario mdb connection
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError == 0)
			{
				/************************************************************
				 **add the table names found in the table_list field of the
				 **scenario_plot_filter
				 ************************************************************/
				strSQL = "SELECT table_list FROM scenario_plot_filter WHERE " + 
					" scenario_id = '" + strScenarioId + "' AND current_yn = 'Y';";
				p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);

				if (p_ado.m_intError ==0)
				{
					while (p_ado.m_OleDbDataReader.Read())
					{
						if (p_ado.m_OleDbDataReader["table_list"] != System.DBNull.Value)
						{
							if (p_ado.m_OleDbDataReader["table_list"].ToString().Trim().Length > 0)				  
							{
								for (x=0; x<=p_ado.m_OleDbDataReader["table_list"].ToString().Trim().Length - 1; x++)
								{
									if (p_ado.m_OleDbDataReader["table_list"].ToString().Substring(x,1) == ",")
									{
										/***********************************************
										 **make sure this table is still listed
										  **in the data source table and is found
										  **********************************************/
										if (ReferenceCoreScenarioForm.uc_datasource1.ScenarioDataSourceTableExist(str) == true)
										{
											/*******************************
											 **check if duplicate
											 *******************************/
											for (y=0; y <= frmSQL.uc_sql_builder1.lstTables.Items.Count-1;y++)
											{
												if (frmSQL.uc_sql_builder1.lstTables.Items[y].ToString().Trim().ToUpper() ==
													str.Trim().ToUpper())
													break;

											}
											//add the table if it is not loaded in the lstTables list box
											if (y >= frmSQL.uc_sql_builder1.lstTables.Items.Count)
												frmSQL.uc_sql_builder1.lstTables.Items.Add(str);
										}
										str="";
									}
									else
									{
										//build the table name that is between the left and right comma
										str = str + p_ado.m_OleDbDataReader["table_list"].ToString().Substring(x,1);


									}
								}
								/*******************************
								 **check if duplicate
								 *******************************/
								for (y=0; y <= frmSQL.uc_sql_builder1.lstTables.Items.Count-1;y++)
								{
									if (frmSQL.uc_sql_builder1.lstTables.Items[y].ToString().Trim().ToUpper() ==
										str.Trim().ToUpper())
										break;

								}
								//add the table if it is not loaded in the lstTables list box
								if (y >= frmSQL.uc_sql_builder1.lstTables.Items.Count)
									frmSQL.uc_sql_builder1.lstTables.Items.Add(str);

							}
						}
					}
					p_ado.m_OleDbDataReader.Close();
					
				}
				this.m_OleDbConnectionScenario.Close();
				
			}
			p_ado.m_OleDbDataReader = null;
			p_ado.m_OleDbCommand = null;
			this.m_OleDbConnectionScenario=null;
			p_ado = null;

			/*****************************************************
			 ** if plot table is not added to the list then add it
			 *****************************************************/

			/*****************************************************
			 **send to the builder the datasource list view with 
			 **all the MDB file and table name information
			 *****************************************************/
			frmSQL.uc_sql_builder1.InitializeDataSource(ReferenceCoreScenarioForm.uc_datasource1.lstRequiredTables);



			result = frmSQL.ShowDialog(this);
			frmSQL.uc_sql_builder1.frmGridView1.Close();
			if (result == DialogResult.OK)
			{
				if (frmSQL.uc_sql_builder1.txtSQLCommand.Text.Trim().Length > 0)
				{
					result = MessageBox.Show("REPLACE \n" + "----------------\n\n" + "'" + this.txtCurrentSQL.Text + "' \n\n\n WITH  \n----------------\n\n'" + frmSQL.uc_sql_builder1.txtSQLCommand.Text + "'", "Plot Filter",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
					if (result == DialogResult.Yes)
					{
						for(x=0;x <= ReferenceCoreScenarioForm.uc_datasource1.lstRequiredTables.Items.Count-1;x++)
						{
							this.m_ado_core_tables.m_strCoreTables[x] = "";
						}

						for(x=0;x <= frmSQL.uc_sql_builder1.lstTables.Items.Count-1;x++)
						{
							this.m_strCoreTables[x] = frmSQL.uc_sql_builder1.lstTables.Items[x].ToString();
						}
						this.m_intNumberOfCoreTablesLoadedIntoDatasets = frmSQL.uc_sql_builder1.m_intNumberOfTablesLoadedIntoDatasets;
						this.txtCurrentSQL.Text = frmSQL.uc_sql_builder1.txtSQLCommand.Text;
						((frmScenario)this.ParentForm).m_bSave=true;
					}
				}
			}
			frmSQL.Close();
			frmSQL = null;
					
			this.m_ado_core_tables = null;
			p_dao = null;

		}


		private void btnPrevSQL_Click(object sender, System.EventArgs e)
		{
			string strConn="";

			DialogResult result;
            
			string strScenarioId =  ((frmScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";

			frmDialog frmPrevExp = new frmDialog();
				
			frmPrevExp.Width = frmPrevExp.uc_previous_expressions1.m_intFullWd;
			frmPrevExp.Height = frmPrevExp.uc_previous_expressions1.m_intFullHt;
			frmPrevExp.Text = "Core Analysis: Previous SQL Expressions";
					
			frmPrevExp.uc_previous_expressions1.Visible=true;

			strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + strScenarioMDB + ";User Id=admin;Password=;";
			
			frmPrevExp.uc_previous_expressions1.setDataSource(strConn,"SELECT * FROM scenario_plot_filter;","SQL_COMMAND","SQL_COMMAND", "scenario_plot_filter");
            
			result = frmPrevExp.ShowDialog(this);
			if (result == DialogResult.OK)
			{
				result = MessageBox.Show("REPLACE \n" + "----------------\n\n" + "'" + this.txtCurrentSQL.Text + "' \n\n\n WITH  \n----------------\n\n'" + frmPrevExp.uc_previous_expressions1.lblSQL.Text + "'", "Previous SQL",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
				if (result == DialogResult.Yes)
				{
					this.txtCurrentSQL.Text = frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[frmPrevExp.uc_previous_expressions1.m_intSelectColumn + 1].Text;
					((frmScenario)this.ParentForm).m_bSave=true;
				}
			}
			frmPrevExp.Close();
			frmPrevExp = null;
			
		
		}

		private void btnExecuteSQL_Click(object sender, System.EventArgs e)
		{
			macrosubst oMacroSubst = new macrosubst();
			oMacroSubst.ReferenceSQLMacroSubstitutionVariableCollection = frmMain.g_oSQLMacroSubstitutionVariable_Collection;
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile =  frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			oDs.m_strDataSourceTableName = "scenario_datasource";
			oDs.m_strScenarioId=this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text;;
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();
            //instantiate the core tables class
			//ado_core_tables p_core_tables = new ado_core_tables();

			//create a temporary mdb file containing the links 
			//to the core tables and their mdb files
		    //p_core_tables.CreateMDBAndCreateCoreTableDataSourceLinks(
			//	((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + 
			//	"\\core\\db\\scenario_core_rule_definitions.mdb",
			//	((frmScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text,
			//	this.m_oEnv.strTempDir);

			if (oDs.m_intError==0)
			{
				string strConn = "";
				string strFile = oDs.CreateMDBAndTableDataSourceLinks();

				strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + strFile + ";User Id=admin;Password=;";
				string strSQL = this.txtCurrentSQL.Text.Trim();
				if (strSQL.IndexOf("@@",0) >=0) strSQL=oMacroSubst.SQLTranslateVariableSubstitution(strSQL);
				((frmScenario)this.ParentForm).frmGridView1.LoadDataSet(strConn,strSQL);
				if (((frmScenario)this.ParentForm).frmGridView1.Visible==false)
				{

					((frmScenario)this.ParentForm).frmGridView1.MdiParent = this.ParentForm.ParentForm;
					((frmScenario)this.ParentForm).frmGridView1.Text = "Core Analysis: Plot Filter Browser" + " (" + ((frmScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text + ")";
					
					try
					{
						((frmScenario)this.ParentForm).frmGridView1.Show();
					}
					catch
					{
						((frmScenario)this.ParentForm).frmGridView1 = new frmGridView();
						((frmScenario)this.ParentForm).frmGridView1.MdiParent = this.ParentForm.ParentForm;
						((frmScenario)this.ParentForm).frmGridView1.Visible=false;
                        ((frmScenario)this.ParentForm).frmGridView1.LoadDataSet(strConn,strSQL);
						((frmScenario)this.ParentForm).frmGridView1.MdiParent = this.ParentForm.ParentForm;
						((frmScenario)this.ParentForm).frmGridView1.Text = "Core Analysis: Plot Filter Browser" + " (" + ((frmScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text + ")";

						((frmScenario)this.ParentForm).frmGridView1.Show();
					}
				}
				if  (((frmScenario)this.ParentForm).frmGridView1.WindowState == 
					    System.Windows.Forms.FormWindowState.Minimized)
					 ((frmScenario)this.ParentForm).frmGridView1.WindowState = 
						 System.Windows.Forms.FormWindowState.Normal;
				((frmScenario)this.ParentForm).frmGridView1.Focus();
			}
	
			//p_core_tables = null;
		     
		}

		private void txtCurrentSQL_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtYardDist_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				((frmScenario)this.ParentForm).m_bSave=true;
			}
			else
			{
				if (e.KeyChar == '\b')
				{
					((frmScenario)this.ParentForm).m_bSave=true;
				}
				else
				{
					e.Handled=true;
				}
			}
		}

		private void btnYardDistDefault_Click(object sender, System.EventArgs e)
		{
			this.txtYardDist.Text = "1524";
			this.txtYardDist2.Text = "610";
			((frmScenario)this.ParentForm).m_bSave=true;
		}

		private void txtYardDist2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				((frmScenario)this.ParentForm).m_bSave=true;
			}
			else
			{
				if (e.KeyChar == '\b')
				{
					((frmScenario)this.ParentForm).m_bSave=true;
				}
				else
				{
					e.Handled=true;
				}
			}
		}

		private void cmdFilter_Click(object sender, System.EventArgs e)
		{
		}
		public int Val_PlotFilter(System.Data.OleDb.OleDbConnection p_oConn,string p_strSql)
		{
			FIA_Biosum_Manager.macrosubst oVarSub=new macrosubst();
			oVarSub.ReferenceSQLMacroSubstitutionVariableCollection = frmMain.g_oSQLMacroSubstitutionVariable_Collection;
			ado_data_access oAdo = new ado_data_access();


			//make sure no duplicates of biosum_plot_id
			oAdo.m_strSQL = "SELECT b.plotcount,a.biosum_plot_id FROM @@PlotTable@@ a," + 
				"(SELECT COUNT(*) as plotcount,a.biosum_plot_id " + 
				"FROM (" + p_strSql + ") " + 
				"AS a " +
				"GROUP BY a.biosum_plot_id) b " + 
				"WHERE b.biosum_plot_id=a.biosum_plot_id AND b.plotcount > 1";
					
			oAdo.m_strSQL=oVarSub.SQLTranslateVariableSubstitution(oAdo.m_strSQL);
			if ((int)oAdo.getRecordCount(p_oConn,oAdo.m_strSQL,"testforplotduplicates") > 0)
			{
				return -5;
			}
			else
			{
				return 0;
			}

		}
		public int Val_PlotFilter(string p_strSql)
		{
			ado_data_access oAdo = new ado_data_access();
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile =  frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			oDs.m_strDataSourceTableName = "scenario_datasource";
			oDs.m_strScenarioId=this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text;;
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();
			int intError=0;
			if (oDs.m_intError==0)
			{
				string strConn = "";
				string strFile = oDs.CreateMDBAndTableDataSourceLinks();
				oAdo.OpenConnection(oAdo.getMDBConnString(strFile,"",""));
				if (oAdo.m_intError==0)
				{
					intError = (int)Val_PlotFilter(oAdo.m_OleDbConnection,p_strSql.Trim());
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
				}
				else intError=oAdo.m_intError;
			}
			else intError=oDs.m_intError;
			oAdo=null;
			oDs=null;
			return intError;

		}

		private void btnAudit_Click(object sender, System.EventArgs e)
		{
			int intError = Val_PlotFilter(this.txtCurrentSQL.Text);
			if (intError==-5)
			{
				MessageBox.Show("Invalid Sql: Duplicate biosum_plot_id's were detected. Biosum_plot_id must be unique in the Core Scenario Plot Filter","FIA Biosum");
			}
			else if (intError==0)
			{
				MessageBox.Show("Valid Syntax");
			}
			else
			{
				MessageBox.Show("Error");
			}
		}
	
		public string strNonSteepYardingDistance
		{
			set	{ this.txtYardDist.Text  = value; }
			get { return this.txtYardDist.Text.ToString(); }
		}
		public string strSteepYardingDistance
		{
			set	{ this.txtYardDist2.Text  = value; }
			get { return this.txtYardDist2.Text.ToString(); }
		}
		public FIA_Biosum_Manager.frmScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}

		
	}

}
