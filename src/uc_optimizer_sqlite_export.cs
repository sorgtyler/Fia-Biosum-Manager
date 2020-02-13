using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;
using System.Data.SQLite;
using System.Data.OleDb;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_tree_groupings.
	/// </summary>
	public class uc_optimizer_sqlite_export : System.Windows.Forms.UserControl
    {
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
        private Button BtnExport;
        private FIA_Biosum_Manager.frmMain m_frmMain;
        private string m_strDebugFile = "";

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        public uc_optimizer_sqlite_export(FIA_Biosum_Manager.frmMain p_frmMain)
        {
            // This call is required by the Windows.Forms Form Designer.
            InitializeComponent();
            this.m_frmMain = p_frmMain;

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
            this.BtnExport = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnExport);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(664, 424);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // BtnExport
            // 
            this.BtnExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnExport.Location = new System.Drawing.Point(185, 99);
            this.BtnExport.Name = "BtnExport";
            this.BtnExport.Size = new System.Drawing.Size(198, 33);
            this.BtnExport.TabIndex = 27;
            this.BtnExport.Text = "EXPORT";
            this.BtnExport.UseVisualStyleBackColor = true;
            this.BtnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(658, 32);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Export to SQLITE";
            // 
            // uc_optimizer_sqlite_export
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_optimizer_sqlite_export";
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
		}

        private void BtnExport_Click(object sender, EventArgs e)
        {
            string strOptimizerScenario = "weighted_ptmod";
            string strContextDbPath = m_frmMain.getProjectDirectory() + @"\optimizer\" + strOptimizerScenario + @"\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsSqliteContextDbFile;
            string strContextAccdbPath = m_frmMain.getProjectDirectory() + @"\optimizer\" + strOptimizerScenario + @"\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsContextDbFile;
            m_strDebugFile = m_frmMain.getProjectDirectory() + @"\optimizer\" + strOptimizerScenario + @"\db\sqlite_log.txt";

            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            ado_data_access oAdo = new ado_data_access();

            if (System.IO.File.Exists(strContextDbPath))
            {
                DialogResult res = MessageBox.Show("The SQLITE context database already exists!! Do you want to replace it?", "FIA Biosum",
                    MessageBoxButtons.YesNo);
                if (res != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    // Delete existing db
                    System.IO.File.Delete(strContextDbPath);
                }
            }
            try
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: SQLITE export Log " + System.DateTime.Now.ToString() + "\r\n\r\n");
                if (frmMain.g_intDebugLevel > 1)
                {

                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Project: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtProjectId.Text + "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Project Directory: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text + "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Scenario Directory: " + strOptimizerScenario + "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------------------------------------------\r\n");
                }
                oDataMgr.CreateDbFile(strContextDbPath);    // create new, blank database
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created the SQLITE database to hold the context tables \r\n");
                }

                string strConnection = "data source=" + strContextDbPath;
                string strAccdbConnection = oAdo.getMDBConnString(strContextAccdbPath, "", "");
                string strTable = "";
                System.Collections.Generic.IList<string> lstTables = new System.Collections.Generic.List<string>();
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
                {
                    con.Open();
                    strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsDiameterSpeciesGroupRefTableName;
                    frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteDiameterSpeciesGroupRefTable(oDataMgr, con,
                        strTable);
                    lstTables.Add(strTable);

                    strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsFvsWeightedVariablesRefTableName;
                    frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteFvsWeightedVariableRefTable(oDataMgr, con, strTable);
                    lstTables.Add(strTable);

                    strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName;
                    frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEconWeightedVariableRefTable(oDataMgr, con, strTable);
                    lstTables.Add(strTable);


                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Created target tables \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                    }


                    using (OleDbConnection oAccessConn = new OleDbConnection(strAccdbConnection))
                    {
                        oAccessConn.Open();
                        foreach(string strTableName in lstTables)
                        {
                            if (oAdo.TableExist(oAccessConn, strTableName))
                            {
                                string strSql = "select * from " + strTableName;
                                oAdo.CreateDataTable(oAccessConn, strSql, strTableName, false);

                                using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(strSql, con))
                                {
                                    System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da);
                                    da.InsertCommand = cb.GetInsertCommand();
                                    int rows = da.Update(oAdo.m_DataTable);
                                }
                            }
                        }
                    }
                }
                MessageBox.Show("done!!");

            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message);

            }
        }

        private void CodeExamples()
        {
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = @"Data Source=C:\sqlite\db\chinook.db;Version=3;";
            try
            {
                // opening a connection and reading records
                //dataMgr.OpenConnection(strConn);
                //dataMgr.m_strSQL = "select title from albums";
                //dataMgr.SqlQueryReader(dataMgr.m_Connection, dataMgr.m_strSQL);
                //int rowCount = -1;
                //if (dataMgr.m_DataReader.HasRows == true)
                //{
                //    int i = 0;
                //    while(dataMgr.m_DataReader.Read())
                //    {
                //        System.Diagnostics.Debug.WriteLine(i + " " + dataMgr.m_DataReader["title"].ToString().Trim());
                //        i++;
                //    }
                //    rowCount = i;
                //}
                //dataMgr.CloseConnection(dataMgr.m_Connection);

                // sample of copying a SQLITE table link into an MS Access table
                ado_data_access oAdo = new ado_data_access();
                //string strConnection = oAdo.getMDBConnString(@"C:\Docs\Biosum\Blue11_demo_588\optimizer\weighted_ptmod\db\optimizer_results.accdb", "", "");
                //using (OleDbConnection oConn = new OleDbConnection(strConnection))
                //{
                //    oConn.Open();
                //    if (oAdo.TableExist(oConn, "albums"))
                //    {
                //        string albums2 = "albums2";
                //        if (oAdo.TableExist(oConn, albums2))
                //        {
                //            oAdo.m_strSQL = "drop table " + albums2;
                //            oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                //        }
                //        oAdo.m_strSQL = "select * into " + albums2 + " from albums";
                //        oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                //        MessageBox.Show("Table created!!");
                //    }
                //}

                //Create database and create table using our MS Access SQL
                //System.Data.SQLite.SQLiteConnection.CreateFile(@"C:\Docs\fia_biosum\Docs\Reporting\attach.db");
                //using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(@"data source=C:\Docs\fia_biosum\Docs\Reporting\attach.db"))
                //{
                //    using (System.Data.SQLite.SQLiteCommand com = new System.Data.SQLite.SQLiteCommand(con))
                //    {
                //        con.Open();
                //        com.CommandText = Tables.OptimizerScenarioResults.CreateHaulCostTableSQL(Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName);
                //        com.ExecuteNonQuery();
                //    }
                //}

                //string strHaulCosts = oAdo.getMDBConnString(@"C:\Docs\Biosum\Blue11_demo_588\optimizer\weighted_ptmod\db\optimizer_results.accdb", "", "");
                //string strSQL = "Select * from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName;
                DataTable dt = new DataTable();
                //using (OleDbDataAdapter da = new OleDbDataAdapter(strSQL, strHaulCosts))
                //{
                //    //tell it not to set rows to Unchanged
                //    da.AcceptChangesDuringFill = false;
                //    da.Fill(dt);
                //}

                //string LiteConnStr = @"data source=C:\Docs\fia_biosum\Docs\Reporting\attach.db";
                //using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(strSQL, LiteConnStr))
                //{
                //    System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da);
                //    da.InsertCommand = cb.GetInsertCommand();
                //    int rows = da.Update(dt);
                //}

                string conFirstDb = @"data source=C:\Docs\fia_biosum\Docs\Reporting\ok_test.db;Version=3;";
                string attachSQL = @"attach 'C:\Docs\fia_biosum\Docs\Reporting\attach.db' as db1";
                string sqlQuery = "select a.biosum_plot_id, b.complete_haul_cost_dpgt from plot as a " +
                                  "inner join db1.haul_costs as b on a.biosum_plot_id = b.biosum_plot_id " +
                                  "where trim(b.materialcd) = 'M'";
                using (SQLiteConnection singleConnectionFor2DBFiles = new SQLiteConnection(conFirstDb))
                {
                    singleConnectionFor2DBFiles.Open();
                    dataMgr.SqlNonQuery(singleConnectionFor2DBFiles, attachSQL);
                    dataMgr.SqlQueryReader(singleConnectionFor2DBFiles, sqlQuery);
                    int rowCount = 0;
                    if (dataMgr.m_DataReader.HasRows == true)
                    {
                        int i = 0;
                        while (dataMgr.m_DataReader.Read())
                        {
                            i++;
                        }
                        rowCount = i;
                    }
                    System.Diagnostics.Debug.WriteLine("number of rows retrieved: " + rowCount);
                }
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message);
            }
        }

     }


}
