using System;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for version_control.
	/// </summary>
	public class version_control
	{
		private ado_data_access m_oAdo = new ado_data_access();
		private dao_data_access m_oDao = new dao_data_access();
		const int APP_VERSION_MAJOR=0;
		const int APP_VERSION_MINOR1=1;
		const int APP_VERSION_MINOR2=2;		
		private string[] m_strAppVerArray=null;
		private string[] m_strDbVerArray=null;
		private string m_strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
		private string m_strProjectVersion="1.0.0";
		private string[] m_strProjectVersionArray=null;
		private string _strProjDir="";
		private FIA_Biosum_Manager.frmMain _oFrmMain=null;
		string strUpdateSql="";
		string strInsertSql="";
        string m_strFVSTreeTable = "";
        string m_strFVSTreeDbFile = "";
        object oMissing = System.Reflection.Missing.Value;


		string[] m_strScenarioRuleDefinitionsTableArray = {"SCENARIO",
															"SCENARIO_COSTS",
															"SCENARIO_DATASOURCE",
															"SCENARIO_HARVEST_COST_COLUMNS",
															"SCENARIO_LAND_OWNER_GROUPS",
															"SCENARIO_PLOT_FILTER",
															"SCENARIO_PLOT_FILTER_MISC",
															"SCENARIO_COND_FILTER",
															"SCENARIO_COND_FILTER_MISC",
															"SCENARIO_PSITES",
															"SCENARIO_RX_INTENSITY",
		                                                    "SCENARIO_FVS_VARIABLES_TIEBREAKER",
		                                                    "SCENARIO_FVS_VARIABLES_OPTIMIZATION",
		                                                    "SCENARIO_FVS_VARIABLES_OVERALL_EFFECTIVE",
		                                                    "SCENARIO_FVS_VARIABLES",
                                                            "SCENARIO_PROCESSOR_SCENARIO_SELECT"};

        string[] m_strProcessorScenarioRuleDefinitionsTableArray = {"SCENARIO",
															"SCENARIO_COSTS",
															"SCENARIO_DATASOURCE",
															"SCENARIO_HARVEST_COST_COLUMNS",
															"SCENARIO_LAND_OWNER_GROUPS",
															"SCENARIO_PLOT_FILTER",
															"SCENARIO_PLOT_FILTER_MISC",
															"SCENARIO_COND_FILTER",
															"SCENARIO_COND_FILTER_MISC",
															"SCENARIO_PSITES",
															"SCENARIO_RX_INTENSITY",
		                                                    "SCENARIO_FVS_VARIABLES_TIEBREAKER",
		                                                    "SCENARIO_FVS_VARIABLES_OPTIMIZATION",
		                                                    "SCENARIO_FVS_VARIABLES_OVERALL_EFFECTIVE",
		                                                    "SCENARIO_FVS_VARIABLES",
                                                            "SCENARIO_PROCESSOR_SCENARIO_SELECT"};


		

		public version_control()
		{
			//
			// TODO: Add constructor logic here
			//
			m_oDao.CreateMDB(m_strTempDbFile);
			m_oDao.m_DaoWorkspace.Close();
			
			


		}
		/// <summary>
		/// Check the project's application version and update to the current version
		/// if different.
		/// </summary>
		public void PerformVersionCheck()
        {
            frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);
            bool bPerformCheck = true;
            string strProjVersionFile = this.ReferenceProjectDirectory + "\\application.version";
            m_strAppVerArray = frmMain.g_oUtils.ConvertListToArray(frmMain.g_strAppVer, ".");
            if (System.IO.File.Exists(strProjVersionFile))
            {
                //Open the file in a stream reader.
                System.IO.StreamReader s = new System.IO.StreamReader(strProjVersionFile);
                //Split the first line into the columns       
                string strProjVersion = s.ReadLine();
                s.Close();

                s = new System.IO.StreamReader(strProjVersionFile);
                string strAppVersion = s.ReadLine();
                s.Close();

                if (strProjVersion.Trim() == frmMain.g_strAppVer.Trim())
                {
                    bPerformCheck = false;
                }
                else
                {
                    if (strProjVersion.Trim().Length > 0)
                    {
                        this.m_strProjectVersion = strProjVersion.Trim();
                        m_strProjectVersionArray = frmMain.g_oUtils.ConvertListToArray(m_strProjectVersion, ".");
                    }
                }
            }
            else
                m_strProjectVersionArray = frmMain.g_oUtils.ConvertListToArray(m_strProjectVersion, ".");
            if (m_strProjectVersionArray != null)
            {

                //no database updates need to be made with these versions
                if (frmMain.g_strAppVer == "5.1.2" && m_strProjectVersion == "5.1.1")
                {
                    UpdateProjectVersionFile(strProjVersionFile);
                    bPerformCheck = false;
                }
                else if (frmMain.g_strAppVer == "5.0.5" && m_strProjectVersion == "5.0.4")
                {

                    UpdateProjectVersionFile(strProjVersionFile);
                    bPerformCheck = false;

                }
                else if (frmMain.g_strAppVer == "5.1.0" && (m_strProjectVersion == "5.0.4" ||
                                                            m_strProjectVersion == "5.0.5"))
                {

                    UpdateProjectVersionFile(strProjVersionFile);
                    bPerformCheck = false;
                }

            }
            //check for partial update
            if (bPerformCheck)
            {
                if (m_strProjectVersion.Trim().Length > 0)
                {
                    if ((frmMain.g_strAppVer == "5.1.4" || frmMain.g_strAppVer=="5.1.5" ||
                        frmMain.g_strAppVer == "5.1.6" ||
                        frmMain.g_strAppVer == "5.1.7" || 
                        frmMain.g_strAppVer == "5.2.0" ||
                        frmMain.g_strAppVer == "5.2.1" ||
                        frmMain.g_strAppVer == "5.2.2") && m_strProjectVersionArray[APP_VERSION_MAJOR] == "5")
                    {
                        //
                        //plot fvs variant assignments table had a major upgrade release with 5.1.4
                        //
                        if (m_strProjectVersionArray[APP_VERSION_MINOR1] == "0" ||
                            (m_strProjectVersionArray[APP_VERSION_MINOR1] == "1" &&
                             (m_strProjectVersionArray[APP_VERSION_MINOR2] == "0" ||
                               m_strProjectVersionArray[APP_VERSION_MINOR2] == "1" ||
                               m_strProjectVersionArray[APP_VERSION_MINOR2] == "2" ||
                               m_strProjectVersionArray[APP_VERSION_MINOR2] == "3")))
                        {
                            //project version is 5.0.? or 5.1.0 to 5.1.3
                            UpdateFVSPlotVariantAssignmentsTable();
                            UpdateProjectVersionFile(strProjVersionFile);
                            bPerformCheck = false;
                        }
                        else
                        {
                            UpdateProjectVersionFile(strProjVersionFile);
                            bPerformCheck = false;
                        }
                            
                    }
                    if ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) < 5) ||
                       (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                       Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 2 &&
                       Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 0))
                        UpgradeFVSOutTreeListFiles();

                    if (frmMain.g_strAppVer == "5.3.0" || (frmMain.g_strAppVer=="5.3.1" && m_strProjectVersion!="5.3.0"))
                    {
                        if ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) <= 4) ||
                            ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) <= 5 &&
                             Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 2)))
                        {
                            UpgradeToPrePostSeqNumMatrix();
                            UpdateAuditDbFile_5_3_0();
                            UpdateDatasources_5_3_0();
                            System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("Version control has detected that the FVS_POTFIRE tables need to be converted to the new Biosum version.\r\n\r\nBy selecting 'Yes', Biosum will convert the FVS POTFIRE table to the new version specificiations.\r\n\r\nBy selecting 'No', the FVS_POTFIRE tables will need to be recreated through FVS.\r\n\r\nDo you want Biosum version control to upgrade the FVS_POTFIRE tables? (Y/N)", "FIA Bisoum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                            if (result == System.Windows.Forms.DialogResult.Yes)
                            {
                                UpgradeFVSOutPOTFireFiles();
                            }
                            result = System.Windows.Forms.MessageBox.Show("Version control has detected that the PREPOST tables need to be converted to the new Biosum version.\r\n\r\nBy selecting 'Yes', Biosum will convert the PREPOST tables to the new version specificiations.\r\n\r\nBy selecting 'No', the PREPOST tables will need to be repopulated by Biosum in the FVS OUTPUT process.\r\n\r\nDo you want Biosum version control to upgrade the PREPOST tables? (Y/N)", "FIA Bisoum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                            if (result == System.Windows.Forms.DialogResult.Yes)
                            {
                                UpgradeFVSOutPREPOSTTables();
                            }
                            
                                                       
                            if ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5))
                            {
                                UpdateProjectVersionFile(strProjVersionFile);
                                bPerformCheck = false;
                            }


                            
                        }
                              
                    }
                    else if (frmMain.g_strAppVer == "5.3.1" && m_strProjectVersion == "5.3.0")
                    {
                                UpdateProjectVersionFile(strProjVersionFile);
                                bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.3.2" && (m_strProjectVersion == "5.3.0" || m_strProjectVersion=="5.3.1"))
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if ((frmMain.g_strAppVer == "5.4.0" || frmMain.g_strAppVer=="5.4.1") && (m_strProjectVersion=="5.3.2" || m_strProjectVersion == "5.3.0" || m_strProjectVersion == "5.3.1"))
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.4.1" && m_strProjectVersion == "5.4.0")
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.5.1" && m_strProjectVersion == "5.5.0")
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.5.2" && m_strProjectVersion == "5.5.1")
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.5.3" && (m_strProjectVersion == "5.5.2" || m_strProjectVersion=="5.5.1"))
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.5.4" && (m_strProjectVersion == "5.5.3" || m_strProjectVersion == "5.5.2" || m_strProjectVersion=="5.5.1"))
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if ((frmMain.g_strAppVer == "5.5.0" || frmMain.g_strAppVer=="5.5.1" || frmMain.g_strAppVer=="5.5.2" || frmMain.g_strAppVer=="5.5.3" || frmMain.g_strAppVer=="5.5.4") && (m_strProjectVersion == "5.4.0" || m_strProjectVersion == "5.4.1" || m_strProjectVersion == "5.4.2"))
                    {
                        UpdateFVSPlotVariantAssignmentsTable();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if ((frmMain.g_strAppVer=="5.5.4" || frmMain.g_strAppVer=="5.5.3" || frmMain.g_strAppVer == "5.5.2" || frmMain.g_strAppVer == "5.5.1") && m_strProjectVersion == "5.5.0")
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer=="5.5.4" && frmMain.g_strAppVer == "5.5.3" && frmMain.g_strAppVer == "5.5.2" && m_strProjectVersion == "5.5.1")
                    {
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    
                   
                    
                }
               
            }
            if (bPerformCheck)
            {

                string strInfo = frmMain.g_sbpInfo.Text;
                frmMain.g_sbpInfo.Text = "Version Update: Checking Project Table...Stand by";
                CheckProjectTable();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Core Analysis Scenario Rule Definitions...Stand by";
                CheckCoreScenarioRuleDefinitionTables();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Processor Database File...Stand by";
                CheckBiosumProcessorDbFile();
                frmMain.g_sbpInfo.Text = "Version Update: Checking FRCS Excel File...Stand by";
                CheckFRCSFile();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Project Datasource Table Records...Stand by";
                CheckProjectDatasourceTableRecords();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Project Datasource Tables...Stand by";
                CheckProjectDatasourceTables();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Pre-Populated Reference Tables...Stand by";
                CheckProjectReferenceDatasourceTables();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Core Scenario Datasource Table Records...Stand by";
                CheckCoreScenarioDatasourceTableRecords();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Processor Scenario Datasource Table Records...Stand by";
                CheckProcessorScenarioDatasourceTableRecords();


                frmMain.g_sbpInfo.Text = "Version Update: Checking Rx Values...Stand By";
                CheckRxValues();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Fvs Out Pre-Post Rx Values...Stand By";
                CheckFVSOutPrePostValues();

                //frmMain.g_sbpInfo.Text = "Version Update: Checking Core Scenario Results Table Records...Stand by";
                //CheckCoreScenarioResultsTables();
                CleanUp();

                UpdateProjectVersionFile(strProjVersionFile);

                frmMain.g_sbpInfo.Text = strInfo;
            }
            frmMain.g_oFrmMain.DeactivateStandByAnimation();

        }
        private void UpdateProjectVersionFile(string p_strProjectVersionFile)
        {
            if (System.IO.File.Exists(p_strProjectVersionFile))
                System.IO.File.Delete(p_strProjectVersionFile);
            frmMain.g_oUtils.WriteText(p_strProjectVersionFile, frmMain.g_strAppVer);
            
        }
		/// <summary>
		/// Update project table with any current version changes.
		/// </summary>
		private void CheckProjectTable()
		{
			int x,y;
            //create a link to the project table
		   // m_oDao.CreateTableLink(m_strTempDbFile,"project_current",this.ProjectDirectory,"project");
            //m_oAdo.OpenConnection(m_oAdo.getMDBConnString(this.m_strTempDbFile,"",""));
			m_oAdo.OpenConnection(m_oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                                                          frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));
		    //frmMain.g_oTables.m_oProject.CreateProjectTable(m_oAdo,m_oAdo.m_OleDbConnection);
			if (m_oAdo.ColumnExist(m_oAdo.m_OleDbConnection,"project","application_version")==false)
			{
				m_oAdo.m_strSQL = "ALTER TABLE project ADD COLUMN application_version TEXT(11)";
				m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);
			}
			else
			{
				
			}

			m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
			

			//string[] strSourceColumnsArray = m_oAdo.getFieldNamesArray(m_oAdo.m_OleDbConnection,"SELECT * FROM project_current");
			//string[] strDestColumnsArray = m_oAdo.getFieldNamesArray(m_oAdo.m_OleDbConnection,"SELECT * FROM project");
			//for (x=0;x<=strDestColumnsArray.Length-1;x++)
			//{

			//}



			
           
		}
		/// <summary>
		/// Update the scenario rule definitions db file with current version.
		/// </summary>
		private void CheckCoreScenarioRuleDefinitionTables()
		{
			
			int x,y,z;
			string strColumns="";
			string strSql="";
			string strTempDbFile="";
			string[] strProjectTablesArray;
			string[] strAppTablesArray;
			System.Data.OleDb.OleDbConnection oConn;
			string[] strDestColArray=null;
			string[] strSourceColArray=null;
			dao_data_access oDao = new dao_data_access();
			ado_data_access oAdo = new ado_data_access();
			string strTableName="";
			string strFile1=this.ReferenceProjectDirectory.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb";
			string strFile2=this.ReferenceProjectDirectory.Trim() + "\\core\\db\\scenario.mdb";
			if (System.IO.File.Exists(strFile1)==false) 
			{
				this.ReferenceMainForm.frmProject.uc_project1.CreateCoreScenarioRuleDefinitionDbAndTables(strFile1);
				if (System.IO.File.Exists(strFile2)==true)
				{
					string[] strTablesToLink = null;
					oDao.getTableNames(strFile2,ref strTablesToLink);
					for (x=0;x<=strTablesToLink.Length-1;x++)
					{
						if (strTablesToLink[x] == null) break;
						if (oDao.TableExists(strFile1,strTablesToLink[x].Trim() + "_temp"))
							oDao.DeleteTableFromMDB(strFile1,strTablesToLink[x].Trim() + "_temp");
							  
						oDao.CreateTableLink(strFile1,strTablesToLink[x].Trim() + "_temp",strFile2,strTablesToLink[x].Trim());


					}
					oAdo.OpenConnection(oAdo.getMDBConnString(strFile1,"",""));
					for (x=0;x<=strTablesToLink.Length-1;x++)
					{
						if (strTablesToLink[x] == null) break;
						strSql="";
						switch (strTablesToLink[x].Trim().ToUpper())
						{
							case "SCENARIO":
								strSql = "INSERT INTO scenario SELECT * FROM scenario_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_COSTS":
								strDestColArray = oAdo.getFieldNamesArray(oAdo.m_OleDbConnection,"SELECT * FROM scenario_costs");
								strSourceColArray=oAdo.getFieldNamesArray(oAdo.m_OleDbConnection,"SELECT * FROM scenario_costs_temp");
								strColumns="";
								for (y=0;y<=strDestColArray.Length - 1;y++)
								{
									for (z=0;z<=strSourceColArray.Length-1;z++)
									{
										if (strSourceColArray[z].Trim().ToUpper() == 
											strDestColArray[y].Trim().ToUpper())
										{
											strColumns=strColumns + strSourceColArray[z].Trim() + ",";
											break;
										}
									}
								}
								if (strColumns.Trim().Length > 0)
								{
									strColumns=strColumns.Substring(0,strColumns.Length - 1);
									strSql = "INSERT INTO scenario_costs " + 
										     "(" + strColumns + ") " + 
										     "SELECT " + strColumns + " " + 
										     "FROM scenario_costs_temp " + 
											 "WHERE scenario_id IS NOT NULL AND " + 
											 "LEN(TRIM(scenario_id)) > 0";
								}

								break;
							case "SCENARIO_DATASOURCE":
                                strSql = "INSERT INTO scenario_datasource SELECT * FROM scenario_datasource_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0 AND TRIM(UCASE(TABLE_TYPE)) NOT IN ('FIRE AND FUEL EFFECTS','ADDITIONAL HARVEST COSTS','TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS','TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES')";
								break;
							case "SCENARIO_HARVEST_COST_COLUMNS":
								strSql = "INSERT INTO scenario_harvest_cost_columns SELECT * FROM scenario_harvest_cost_columns_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_LAND_OWNER_GROUPS":
								strSql = "INSERT INTO scenario_land_owner_groups SELECT * FROM scenario_land_owner_groups_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_PLOT_FILTER":
								strSql = "INSERT INTO scenario_plot_filter SELECT * FROM scenario_plot_filter_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_PLOT_FILTER_MISC":
								strSql = "INSERT INTO scenario_cond_filter_misc SELECT * FROM scenario_plot_filter_misc_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_COND_FILTER":
								strSql = "INSERT INTO scenario_cond_filter SELECT * FROM scenario_cond_filter_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_COND_FILTER_MISC":
								strSql = "INSERT INTO scenario_cond_filter_misc SELECT * FROM scenario_cond_filter_misc_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_PSITES":
								strSql = "INSERT INTO scenario_psites SELECT * FROM scenario_psites_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_RX_INTENSITY":
								strSql = "INSERT INTO scenario_rx_intensity SELECT * FROM scenario_rx_intensity_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
                            case "SCENARIO_PROCESSOR_SCENARIO_SELECT":
                                strSql = "INSERT INTO scenario_processor_scenario_select SELECT * FROM scenario_processor_scenario_select_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
                                break;
						}
						if (strSql.Trim().Length > 0)
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,strSql);

					}

					strSql = "UPDATE scenario SET file='scenario_core_rule_definitions.mdb'";
				    oAdo.SqlNonQuery(oAdo.m_OleDbConnection,strSql);

					//remove the linked tables
					for (x=0;x<=strTablesToLink.Length-1;x++)
					{
						if (strTablesToLink[x]==null) break;
						strSql="";
						if (oAdo.TableExist(oAdo.m_OleDbConnection,strTablesToLink[x].Trim() + "_temp"))
						{
							strSql = "DROP TABLE " + strTablesToLink[x].Trim() + "_temp";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,strSql);
						}

					}
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
					//reset path variables in the scenario db
					this.ReferenceMainForm.frmProject.uc_project1.SetProjectPathEnvironmentVariables();
					

				}
			}
			oDao.m_DaoWorkspace.Close();

			//check and update table structures
			//create temp access table that will contain links
			//create a temp file that will contain a temporary copy of 
			//the latest tables
			oDao = new dao_data_access();
			strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
			oDao.CreateMDB(strTempDbFile);
			oDao.m_DaoWorkspace.Close();
			oConn = new System.Data.OleDb.OleDbConnection();
			oConn.ConnectionString = oAdo.getMDBConnString(strTempDbFile,"","");
			oConn.Open();
			oAdo.OpenConnection(oAdo.getMDBConnString(strFile1,"",""));

			strAppTablesArray = oAdo.getTableNames(oAdo.m_OleDbConnection);
			for (x=0;x<=strAppTablesArray.Length-1;x++)
			{
				for (y=0;y<=this.m_strScenarioRuleDefinitionsTableArray.Length - 1;y++)
				{
					if (strAppTablesArray[x].Trim().ToUpper()==this.m_strScenarioRuleDefinitionsTableArray[y].Trim().ToUpper())
					{
						strTableName="";
						switch (m_strScenarioRuleDefinitionsTableArray[y])
						{
							case "SCENARIO":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioTableName;
								frmMain.g_oTables.m_oScenario.CreateScenarioTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_COSTS":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioCostsTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioCostsTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_DATASOURCE":
								strTableName=Tables.Scenario.DefaultScenarioDatasourceTableName;
								frmMain.g_oTables.m_oScenario.CreateScenarioDatasourceTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_HARVEST_COST_COLUMNS":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioHarvestCostColumnsTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_LAND_OWNER_GROUPS":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioLandOwnerGroupsTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_PLOT_FILTER":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioPlotFilterTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_PLOT_FILTER_MISC":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioPlotFilterMiscTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_COND_FILTER":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioCondFilterTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_COND_FILTER_MISC":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioCondFilterMiscTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_PSITES":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioPSitesTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioPSitesTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_RX_INTENSITY":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioRxIntensityTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioRxIntensityTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_FVS_VARIABLES_TIEBREAKER":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesTieBreakerTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_FVS_VARIABLES_OPTIMIZATION":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesOptimizationTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_FVS_VARIABLES_OVERALL_EFFECTIVE":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesOverallEffectiveTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_FVS_VARIABLES":
								strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesTable(oAdo,oConn,strTableName);
								break;
                            case "SCENARIO_PROCESSOR_SCENARIO_SELECT":
                                strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName;
								frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo,oConn,strTableName);

                                break;


						}
						if (strTableName.Trim().Length > 0)
						{
							string strEmptyTable = strTableName + "_work_temp";
							oAdo.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + strTableName + ")";
							//check if the projects version of the table has records
							if ((int)oAdo.getRecordCount(oAdo.m_OleDbConnection,oAdo.m_strSQL, strTableName) > 0)
							{
								bool bModified=false;
								if (oAdo.TableExist(oAdo.m_OleDbConnection,strEmptyTable))
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
								//create a temp copy with no records (empty table) of the projects version of the table
								oAdo.m_strSQL = "SELECT TOP 1 * INTO " + strEmptyTable + " FROM " + strTableName;
								oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
								oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DELETE FROM " + strEmptyTable);
								//compare the application version of the table with the 
								//projects version of the table and 
								//modify the projects version of the table structure if needs be.
								bModified = oAdo.ReconcileTableColumns(oAdo.m_OleDbConnection,
									strEmptyTable,
									oConn,strTableName);
								if (bModified)
								{
									//create the indexes for the empty table structure
									//oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent,oAdoCurrent.m_OleDbConnection,
									//	oDs.m_strDataSource[x,Datasource.TABLETYPE].Trim(),strEmptyTable.Trim());
									//insert all the records into our new empty table
									oAdo.m_strSQL = "INSERT INTO " + strEmptyTable + " SELECT * FROM " + strTableName;
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
									//drop the current table
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + strTableName);
									//create the table structure giving it the current table name
									oAdo.m_strSQL = "SELECT TOP 1 * INTO " + strTableName + " FROM " + strEmptyTable;
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
									//delete the 1 record in the current table
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DELETE FROM " + strTableName);
									//insert all the records from the new empty table to the current table
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"INSERT INTO " + strTableName + " SELECT *  FROM " + strEmptyTable.Trim());
									//drop the empty table
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
								}
								else
								{
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
								}
							}
							else
							{
								if (oAdo.ReconcileTableColumns(oAdo.m_OleDbConnection,
									strTableName,
									oConn,strTableName))
								{
								}
							}
						}
					}
				}
			}
			//create any tables that do not exist
			for (y=0;y<=this.m_strScenarioRuleDefinitionsTableArray.Length - 1;y++)
			{
				if (!oAdo.TableExist(oAdo.m_OleDbConnection,this.m_strScenarioRuleDefinitionsTableArray[y].Trim().ToUpper()))
				{
					strTableName="";

					switch (m_strScenarioRuleDefinitionsTableArray[y])
					{
						case "SCENARIO":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioTableName;
							frmMain.g_oTables.m_oScenario.CreateScenarioTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_COSTS":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioCostsTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioCostsTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_DATASOURCE":
							strTableName=Tables.Scenario.DefaultScenarioTableName;
							frmMain.g_oTables.m_oScenario.CreateScenarioDatasourceTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_HARVEST_COST_COLUMNS":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioHarvestCostColumnsTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_LAND_OWNER_GROUPS":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioLandOwnerGroupsTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_PLOT_FILTER":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioPlotFilterTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_PLOT_FILTER_MISC":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioPlotFilterMiscTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_COND_FILTER":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioCondFilterTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_COND_FILTER_MISC":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioCondFilterMiscTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_PSITES":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioPSitesTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioPSitesTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_RX_INTENSITY":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioRxIntensityTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioRxIntensityTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_FVS_VARIABLES_TIEBREAKER":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesTieBreakerTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_FVS_VARIABLES_OPTIMIZATION":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesOptimizationTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_FVS_VARIABLES_OVERALL_EFFECTIVE":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesOverallEffectiveTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_FVS_VARIABLES":
							strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioFVSVariablesTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
                        case "SCENARIO_PROCESSOR_SCENARIO_SELECT":
                            strTableName=Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName;
							frmMain.g_oTables.m_oCoreScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo,oConn,strTableName);
                            break;
					}
				}
			}

			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oAdo=null;
			oDao=null;
			
		}
		/// <summary>
		/// Update the Processor with the latest biosum_processor.mdb file.
		/// </summary>
		private void CheckBiosumProcessorDbFile()
		{
			string strSourceFile = frmMain.g_oEnv.strAppDir + "\\db\\biosum_processor.mdb";
			string strDestFile = this.ReferenceProjectDirectory + "\\processor\\db\\biosum_processor.mdb";
			string strProjVersionFile = this.ReferenceProjectDirectory + "\\processor\\db\\biosum_processor.version";
			string strAppVersionFile = frmMain.g_oEnv.strAppDir + "\\db\\biosum_processor.version";
			bool bCopyAppVersion=false;

			if (System.IO.File.Exists(strProjVersionFile))
			{
				//Open the file in a stream reader.
				System.IO.StreamReader s = new System.IO.StreamReader(strProjVersionFile);
        
				//Split the first line into the columns       
				string strProjVersion = s.ReadLine();
				s.Close();

				s = new System.IO.StreamReader(strAppVersionFile);
				string strAppVersion = s.ReadLine();
				s.Close();

				if (strProjVersion.Trim() != strAppVersion.Trim()) 
				{ 
					System.IO.File.Copy(strDestFile,strDestFile + "_previousversion_" + strProjVersion + ".mdb",true);
					bCopyAppVersion=true;
				}



			}
			else 
			{  
				System.IO.File.Copy(strDestFile,strDestFile + "_previousversion_1.0.0.mdb",true);
				bCopyAppVersion=true;
			}


			if (bCopyAppVersion)
			{
				System.IO.File.Copy(strSourceFile, strDestFile,true);
				System.IO.File.Copy(strAppVersionFile,strProjVersionFile,true);
			}


		}
		/// <summary>
		/// Update FRCS with the latest FRCS.xls file.
		/// </summary>
		private void CheckFRCSFile()
		{
			string strSourceFile = frmMain.g_oEnv.strAppDir + "\\db\\frcs.xls";
			string strDestFile = this.ReferenceProjectDirectory + "\\processor\\db\\frcs.xls";
			string strProjVersionFile = this.ReferenceProjectDirectory + "\\processor\\db\\frcs.version";
			string strAppVersionFile = frmMain.g_oEnv.strAppDir + "\\db\\frcs.version";
			bool bCopyAppVersion=false;

			if (System.IO.File.Exists(strProjVersionFile))
			{
				//Open the file in a stream reader.
				System.IO.StreamReader s = new System.IO.StreamReader(strProjVersionFile);
        
				//Split the first line into the columns       
				string strProjVersion = s.ReadLine();
				s.Close();

				s = new System.IO.StreamReader(strAppVersionFile);
				string strAppVersion = s.ReadLine();
				s.Close();

				if (strProjVersion.Trim() != strAppVersion.Trim()) 
				{ 
					System.IO.File.Copy(strDestFile,strDestFile + "_previousversion_" + strProjVersion + ".xls",true);
					bCopyAppVersion=true;
				}



			}
			else 
			{  
				if (System.IO.File.Exists(strDestFile))
					System.IO.File.Copy(strDestFile,strDestFile + "_previousversion_1.0.0.xls",true);
				bCopyAppVersion=true;
			}


			if (bCopyAppVersion)
			{
				System.IO.File.Copy(strSourceFile, strDestFile,true);
				System.IO.File.Copy(strAppVersionFile,strProjVersionFile,true);
			}


		}
		/// <summary>
		/// Check and update the project datasource table with the latest version of table type entries
		/// </summary>
		private void CheckProjectDatasourceTableRecords()
		{
		    
			int x;
			int y;
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
			oDs.m_strDataSourceTableName = "datasource";
			oDs.m_strScenarioId="";
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();

            

			ado_data_access oAdo=new ado_data_access();
			string strDbFile="";
            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));
			//
            //drop obsolete table types 
            //
            //additional harvest costs table type
            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type)) IN ('ADDITIONAL HARVEST COSTS','TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES')";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            //
			//
			//make sure all the current table type data sources are accounted for
			//
			for (x=0;x<=Datasource.g_strProjectDatasourceTableTypesArray.Length - 1;x++)
			{
				//check if the latest datasource table type exists 
				y = oDs.getDataSourceTableNameRow(Datasource.g_strProjectDatasourceTableTypesArray[x].Trim());
				if (y==-1)
				{
					//table type does not exist so create it
					switch (Datasource.g_strProjectDatasourceTableTypesArray[x].Trim().ToUpper())
					{
						case "PLOT":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableDbFile),
                                frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName);
							break;
						case "CONDITION":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableDbFile),
								frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName);
							break;
						case "TREE":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableDbFile),
								frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName);
							break;
						
						case "OWNER GROUPS":  
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								"ref_master.mdb",
								"owner_groups");
							break;
						case "TREE DIAMETER GROUPS":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.Processor.DefaultTreeDiamGroupsDbFile),
								Tables.Processor.DefaultTreeDiamGroupsTableName);
							break;
						case "TREATMENT PRESCRIPTIONS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Treatment Prescriptions'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 
								"'rx');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
                        case "TREE SPECIES GROUPS":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.Processor.DefaultTreeSpeciesGroupsDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Tree Species Groups'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 
								"'tree_species_groups');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "TREE SPECIES GROUPS LIST":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.Processor.DefaultTreeSpeciesGroupsListDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Tree Species Groups List'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 
								"'tree_species_groups_list');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "TREE SPECIES":
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Tree Species'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," + 
								"'tree_species');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "FVS TREE SPECIES":
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('FVS Tree Species'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," + 
								"'fvs_tree_species');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						
						case "TRAVEL TIMES":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\gis\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oTravelTime.DefaultTravelTimeTableDbFile),
								frmMain.g_oTables.m_oTravelTime.DefaultTravelTimeTableName);
							break;
						case "PROCESSING SITES":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\gis\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oTravelTime.DefaultProcessingSiteTableDbFile),
								frmMain.g_oTables.m_oTravelTime.DefaultProcessingSiteTableName);
							break;
						case "FVS TREE LIST FOR PROCESSOR":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\processor\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultFVSTreeTableDbFile),
                               Tables.FVS.DefaultFVSTreeTableName);
							break;
						case "FIADB FVS VARIANT":
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('FIADB FVS Variant'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," + 
								"'fiadb_fvs_variant');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						
						case "PLOT AND CONDITION RECORD AUDIT":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileName(frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableDbFile),
								frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName);
							break;
						case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileName(frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableDbFile),
								frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName);
							break;
						case "TREE REGIONAL BIOMASS":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultTreeRegionalBiomassTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Tree Regional Biomass'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 
								"'tree_regional_biomass');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "POPULATION EVALUATION":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Population Evaluation'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 
								"'pop_eval');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "POPULATION ESTIMATION UNIT":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Population Estimation Unit'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 								
								"'pop_estn_unit');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "POPULATION STRATUM":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Population Stratum'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 		
								"'pop_stratum');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "POPULATION PLOT STRATUM ASSIGNMENT":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Population Plot Stratum Assignment'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 		
								"'pop_plot_stratum_assgn');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "SITE TREE":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Site Tree'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'master.mdb'," + 
								"'sitetree');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
                        
                        //version 5 additions
                        case "TREATMENT PRESCRIPTIONS ASSIGNED FVS COMMANDS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxFvsCommandTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Prescriptions Assigned FVS Commands'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," + 
								"'rx_fvs_commands');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxHarvestCostColumnsTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Prescriptions Harvest Cost Columns'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rx_harvest_cost_columns');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PRESCRIPTION CATEGORIES":
                            //strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oReference.DefaultRxHarvestCostTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Prescription Categories'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," +
                                "'fvs_rx_category');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Prescription Subcategories'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," +
                                "'fvs_rx_subcategory');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PACKAGES":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxPackageTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Packages'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rxpackage');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PACKAGE ASSIGNED FVS COMMANDS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxPackageFvsCommandTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Package Assigned FVS Commands'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rxpackage_fvs_commands');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PACKAGE MEMBERS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxPackageMembersTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Package Members'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rxpackage_members');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PACKAGE FVS COMMANDS ORDER":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxPackageFvsCommandsOrderTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Package FVS Commands Order'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rxpackage_fvs_commands_order');";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            break;
                        case "FVS COMMANDS":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('FVS Commands'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_fvscommands.mdb'," +
                                "'fvs_commands');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "FRCS SYSTEM HARVEST METHOD":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('FRCS System Harvest Method'," +
                                "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                                "'ref_master.mdb'," +
                                "'frcs_system_harvest_method');";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            break;
                        //version 5.2.1 additions
                        case "FVS WESTERN TREE SPECIES TRANSLATOR":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('FVS Western Tree Species Translator'," +
                                "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                                "'ref_master.mdb'," +
                                "'FVS_WesternTreeSpeciesTranslator');";
                             oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            break;
                        case "FVS EASTERN TREE SPECIES TRANSLATOR":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('FVS Eastern Tree Species Translator'," +
                                "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                                "'ref_master.mdb'," +
                                "'FVS_EasternTreeSpeciesTranslator');";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            break;


                        //case "ADDITIONAL HARVEST COSTS":
                        //    oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        //        "('Additional Harvest Costs'," +
                        //        "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                        //        "'master.mdb'," +
                        //        "'additional_harvest_costs');";
                        //    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                        //    break;
					}
				}
			}
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
		}
		/// <summary>
		/// Project table type Db files, tables and table columns are updated to the current version.
		/// Each record in the datasource table is checked to make sure the db file exists,
		/// the table exists, and all the columns exist. 
		/// </summary>
		private void CheckProjectDatasourceTables()
		{

			int x;
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
			oDs.m_strDataSourceTableName = "datasource";
			oDs.m_strScenarioId="";
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();

			ado_data_access oAdo=new ado_data_access();
			ado_data_access oAdoCurrent=null;

			dao_data_access oDao=null;
			string strCurrDbFile="";
			string strDbFile="";
			string strTempDbFile="";
			string strTempTableName="";
			System.Data.OleDb.OleDbConnection oConn=null;
			
			

			//open the project db file
			oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" + 
				frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));

            


			oAdo.m_strSQL = "SELECT * FROM datasource ORDER BY path,file";
			oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);

			if (oAdo.m_OleDbDataReader.HasRows)
			{
				//sequentially read each record
				while (oAdo.m_OleDbDataReader.Read())
				{
					if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
						oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
					{
					    //get the array number of the current datasource table type
						x=oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());

						
						strDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" + 
							        oAdo.m_OleDbDataReader["file"].ToString().Trim();

						//check if the DB file has changed
						if (strDbFile.Trim().ToUpper() != strCurrDbFile.Trim().ToUpper())
						{
							//close the current open DB file
							if (oAdoCurrent!=null) oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
							else oAdoCurrent = new ado_data_access();

							//make sure the db file exists
							if (System.IO.File.Exists(strDbFile)==false)
							{
								//create the DB file
								if (oDao==null) oDao = new dao_data_access();
								oDao.CreateMDB(strDbFile);
							}
							//open a connection to the DB file
							oAdoCurrent.OpenConnection(oAdo.getMDBConnString(strDbFile,"",""));
							strCurrDbFile=strDbFile;

						}
						//check the table
						if (oDs.m_strDataSource[x,Datasource.TABLESTATUS]=="NF") //NF=table not found
						{
							//create the table
							switch (oDs.m_strDataSource[x,Datasource.TABLETYPE].Trim().ToUpper())
							{
								
								case "PLOT":
									frmMain.g_oTables.m_oFIAPlot.CreatePlotTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName);
									break;
								case "CONDITION":
									frmMain.g_oTables.m_oFIAPlot.CreateConditionTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName);
									break;
								case "TREE":
									frmMain.g_oTables.m_oFIAPlot.CreateTreeTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName);
									break;
								
								case "OWNER GROUPS":  
									frmMain.g_oTables.m_oReference.CreateOwnerGroupsTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultOwnerGroupsTableName);
									break;
								case "TREE DIAMETER GROUPS":
									frmMain.g_oTables.m_oProcessor.CreateTreeDiamGroupsTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,"tree_diam_groups");
									break;
								case "TREATMENT PRESCRIPTIONS":
									frmMain.g_oTables.m_oFvs.CreateRxTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.FVS.DefaultRxTableName);
									break;
								case "TREE SPECIES GROUPS":
									frmMain.g_oTables.m_oProcessor.CreateTreeSpeciesGroupsTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,"tree_species_groups");
									break;
								case "TREE SPECIES GROUPS LIST":
									frmMain.g_oTables.m_oProcessor.CreateTreeSpeciesGroupsListTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,"tree_species_groups_list");
									break;
								case "TREE SPECIES":
									frmMain.g_oTables.m_oReference.CreateTreeSpeciesTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultFVSTreeSpeciesTableName);
									break;
								case "FVS TREE SPECIES":
									frmMain.g_oTables.m_oReference.CreateFVSTreeSpeciesTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultFVSTreeSpeciesTableName);
									break;
								
								case "TRAVEL TIMES":
									frmMain.g_oTables.m_oTravelTime.CreateTravelTimeTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oTravelTime.DefaultTravelTimeTableName);
									break;
								case "PROCESSING SITES":
									frmMain.g_oTables.m_oTravelTime.CreateProcessingSiteTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oTravelTime.DefaultProcessingSiteTableName);
									break;
								case "FVS TREE LIST FOR PROCESSOR":
									frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorIn(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.FVS.DefaultFVSTreeTableName);
									break;
								case "FIADB FVS VARIANT":
									frmMain.g_oTables.m_oReference.CreateFiadbFVSVariantTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultFiadbFVSVariantTableName);
									break;
								
								case "PLOT AND CONDITION RECORD AUDIT":
									frmMain.g_oTables.m_oAudit.CreatePlotCondAuditTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName);
									break;
								case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
									frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName);
									break;
								case "TREE REGIONAL BIOMASS":
									frmMain.g_oTables.m_oFIAPlot.CreateTreeRegionalBiomassTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultTreeRegionalBiomassTableName);
									break;
								case "POPULATION EVALUATION":
									frmMain.g_oTables.m_oFIAPlot.CreatePopEvalTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName);
									break;
								case "POPULATION ESTIMATION UNIT":
									frmMain.g_oTables.m_oFIAPlot.CreatePopEstnUnitTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName);
									break;
								case "POPULATION STRATUM":
									frmMain.g_oTables.m_oFIAPlot.CreatePopStratumTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName);
									break;
								case "POPULATION PLOT STRATUM ASSIGNMENT":
									frmMain.g_oTables.m_oFIAPlot.CreatePopPlotStratumAssgnTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName);
									break;
								case "SITE TREE":
									frmMain.g_oTables.m_oFIAPlot.CreateSiteTreeTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName);
									break;
								case "INVENTORIES":
									frmMain.g_oTables.m_oReference.CreateInventoriesTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultInventoriesTableName);
									break;
                                //version 5 additions
                                case "TREATMENT PRESCRIPTIONS ASSIGNED FVS COMMANDS":
                                    frmMain.g_oTables.m_oFvs.CreateRxFvsCommandsTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.FVS.DefaultRxFvsCommandTableName);
                                    break;
                                case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                                    frmMain.g_oTables.m_oFvs.CreateRxHarvestCostColumnTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.FVS.DefaultRxHarvestCostColumnsTableName);
                                    break;
                                case "TREATMENT PRESCRIPTION CATEGORIES":
                                    frmMain.g_oTables.m_oReference.CreateRxCategoryTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultRxCategoryTableName);
                                    break;
                                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                                    frmMain.g_oTables.m_oReference.CreateRxSubCategoryTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultRxSubCategoryTableName);
                                    break;
                                case "TREATMENT PACKAGES":
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.FVS.DefaultRxPackageTableName);
                                    break;
                                case "TREATMENT PACKAGE ASSIGNED FVS COMMANDS":
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.FVS.DefaultRxPackageFvsCommandTableName);
                                    break;
                                case "TREATMENT PACKAGE MEMBERS":
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageMembersTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.FVS.DefaultRxPackageMembersTableName);
                                    break;
                                case "TREATMENT PACKAGE FVS COMMANDS ORDER":
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsOrderTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.FVS.DefaultRxPackageFvsCommandsOrderTableName);
                                    break;
                               
                                case "FRCS SYSTEM HARVEST METHOD":
                                    frmMain.g_oTables.m_oReference.CreateFRCSHarvestMethodTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultFRCSHarvestMethodTableName);
                                    break;
                                case "ADDITIONAL HARVEST COSTS":
                                    frmMain.g_oTables.m_oProcessor.CreateAdditionalHarvestCostsTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection,Tables.Processor.DefaultAdditionalHarvestCostsTableName);
                                    break;
                                //5.2.1 additions
                                case "FVS WESTERN TREE SPECIES TRANSLATOR":
                                    frmMain.g_oTables.m_oReference.CreateFVSWesternSpeciesTranslatorTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultFVSWesternTreeSpeciesTableName);
                                    break;
                                case "FVS EASTERN TREE SPECIES TRANSLATOR":
                                    frmMain.g_oTables.m_oReference.CreateFVSEasternSpeciesTranslatorTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultFVSEasternTreeSpeciesTableName);
                                    break;
                               
							}
						}
						else
						{
							//update columns in the table
							if (oDao==null)
							{ 
								//create a temp file that will contain a temporary copy of 
								//the latest tables
								oDao = new dao_data_access();
								strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
								oDao.CreateMDB(strTempDbFile);
								oConn = new System.Data.OleDb.OleDbConnection();
								oConn.ConnectionString = oAdoCurrent.getMDBConnString(strTempDbFile,"","");
								oConn.Open();
							}

							strTempTableName="";
							//create an empty table structure of the table
							switch (oDs.m_strDataSource[x,Datasource.TABLETYPE].Trim().ToUpper())
							{
								case "PLOT":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePlotTable(oAdoCurrent,oConn,
										     strTempTableName);
									break;
								case "CONDITION":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName;
									frmMain.g_oTables.m_oFIAPlot.CreateConditionTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREE":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName;
									frmMain.g_oTables.m_oFIAPlot.CreateTreeTable(oAdoCurrent,oConn,strTempTableName);
									break;
								
								case "OWNER GROUPS":  
									strTempTableName = Tables.Reference.DefaultOwnerGroupsTableName;
									frmMain.g_oTables.m_oReference.CreateOwnerGroupsTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREE DIAMETER GROUPS":
									strTempTableName = Tables.Processor.DefaultTreeDiamGroupsTableName;
									frmMain.g_oTables.m_oProcessor.CreateTreeDiamGroupsTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREATMENT PRESCRIPTIONS":
									strTempTableName = Tables.FVS.DefaultRxTableName;
									frmMain.g_oTables.m_oFvs.CreateRxTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREE SPECIES GROUPS":
									strTempTableName = Tables.Processor.DefaultTreeSpeciesGroupsTableName;
									frmMain.g_oTables.m_oProcessor.CreateTreeSpeciesGroupsTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREE SPECIES GROUPS LIST":
									strTempTableName = Tables.Processor.DefaultTreeSpeciesGroupsListTableName;
									frmMain.g_oTables.m_oProcessor.CreateTreeSpeciesGroupsListTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREE SPECIES":
									strTempTableName = Tables.Reference.DefaultTreeSpeciesTableName;
									frmMain.g_oTables.m_oReference.CreateTreeSpeciesTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "FVS TREE SPECIES":
									strTempTableName = Tables.Reference.DefaultFVSTreeSpeciesTableName;
									frmMain.g_oTables.m_oReference.CreateFVSTreeSpeciesTable(oAdoCurrent,oConn,strTempTableName);
									break;
							
								case "TRAVEL TIMES":
									strTempTableName = frmMain.g_oTables.m_oTravelTime.DefaultTravelTimeTableName;
									frmMain.g_oTables.m_oTravelTime.CreateTravelTimeTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "PROCESSING SITES":
									strTempTableName = frmMain.g_oTables.m_oTravelTime.DefaultProcessingSiteTableName;
									frmMain.g_oTables.m_oTravelTime.CreateProcessingSiteTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "FVS TREE LIST FOR PROCESSOR":
									strTempTableName = Tables.FVS.DefaultFVSTreeTableName;
									frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorIn(oAdoCurrent,oConn,strTempTableName);
									break;
								case "FIADB FVS VARIANT":
									strTempTableName = Tables.Reference.DefaultFiadbFVSVariantTableName;
									frmMain.g_oTables.m_oReference.CreateFiadbFVSVariantTable(oAdoCurrent,oConn,strTempTableName);
									break;
								
								case "PLOT AND CONDITION RECORD AUDIT":
									strTempTableName = frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName;
									frmMain.g_oTables.m_oAudit.CreatePlotCondAuditTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
									strTempTableName = frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName;
									frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREE REGIONAL BIOMASS":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultTreeRegionalBiomassTableName;
									frmMain.g_oTables.m_oFIAPlot.CreateTreeRegionalBiomassTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "POPULATION EVALUATION":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePopEvalTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "POPULATION ESTIMATION UNIT":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePopEstnUnitTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "POPULATION STRATUM":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePopStratumTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "POPULATION PLOT STRATUM ASSIGNMENT":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePopPlotStratumAssgnTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "SITE TREE":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName;
									frmMain.g_oTables.m_oFIAPlot.CreateSiteTreeTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "INVENTORIES":
									strTempTableName = Tables.Reference.DefaultInventoriesTableName;
									frmMain.g_oTables.m_oReference.CreateInventoriesTable(oAdoCurrent,oConn,strTempTableName);
									break;
                                //version 5 additions
                                case "TREATMENT PRESCRIPTIONS ASSIGNED FVS COMMANDS":
                                    strTempTableName = Tables.FVS.DefaultRxFvsCommandTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxFvsCommandsTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                                    strTempTableName = Tables.FVS.DefaultRxHarvestCostColumnsTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxHarvestCostColumnTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PRESCRIPTION CATEGORIES":
                                    strTempTableName = Tables.Reference.DefaultRxCategoryTableName;
                                    frmMain.g_oTables.m_oReference.CreateRxCategoryTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                                    strTempTableName = Tables.Reference.DefaultRxSubCategoryTableName;
                                    frmMain.g_oTables.m_oReference.CreateRxSubCategoryTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PACKAGES":
                                    strTempTableName = Tables.FVS.DefaultRxPackageTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PACKAGE ASSIGNED FVS COMMANDS":
                                    strTempTableName = Tables.FVS.DefaultRxPackageFvsCommandTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PACKAGE MEMBERS":
                                    strTempTableName = Tables.FVS.DefaultRxPackageMembersTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageMembersTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PACKAGE FVS COMMANDS ORDER":
                                    strTempTableName = Tables.FVS.DefaultRxPackageFvsCommandsOrderTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsOrderTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                
                                case "FRCS SYSTEM HARVEST METHOD":
                                    strTempTableName = Tables.Reference.DefaultFRCSHarvestMethodTableName;
                                    frmMain.g_oTables.m_oReference.CreateFRCSHarvestMethodTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "ADDITIONAL HARVEST COSTS":
                                    strTempTableName = Tables.Processor.DefaultAdditionalHarvestCostsTableName;
                                    frmMain.g_oTables.m_oProcessor.CreateAdditionalHarvestCostsTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                //5.2.1 additions
                                case "FVS WESTERN TREE SPECIES TRANSLATOR":
                                    strTempTableName = Tables.Reference.DefaultFVSWesternTreeSpeciesTableName;
                                    frmMain.g_oTables.m_oReference.CreateFVSWesternSpeciesTranslatorTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "FVS EASTERN TREE SPECIES TRANSLATOR":
                                    strTempTableName = Tables.Reference.DefaultFVSEasternTreeSpeciesTableName;
                                    frmMain.g_oTables.m_oReference.CreateFVSEasternSpeciesTranslatorTable(oAdoCurrent, oConn, strTempTableName);

                                    break;

							}
							if (strTempTableName.Trim().Length > 0)
							{
								
								
								string strEmptyTable = oDs.m_strDataSource[x,Datasource.TABLE].Trim() + "_work_temp";

								oAdoCurrent.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + oDs.m_strDataSource[x,Datasource.TABLE].Trim() + ")";
								
								//check if the projects version of the table has records
								if ((int)oAdoCurrent.getRecordCount(oAdoCurrent.m_OleDbConnection,oAdoCurrent.m_strSQL, oDs.m_strDataSource[x,Datasource.TABLE].Trim()) > 0)
								{

									bool bModified=false;

									if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,strEmptyTable))
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strEmptyTable);

									//create a temp copy with no records (empty table) of the projects version of the table
									oAdoCurrent.m_strSQL = "SELECT TOP 1 * INTO " + strEmptyTable + " FROM " + oDs.m_strDataSource[x,Datasource.TABLE].Trim();
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,oAdoCurrent.m_strSQL);
								    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DELETE FROM " + strEmptyTable);
									

									//compare the application version of the table with the 
									//projects version of the table and 
									//modify the projects version of the table structure if needs be.
									bModified = oAdoCurrent.ReconcileTableColumns(oAdoCurrent.m_OleDbConnection,
													strEmptyTable,
													oConn,strTempTableName);
									if (bModified)
									{
										
										//insert all the records into our new empty table
										oAdoCurrent.m_strSQL = "INSERT INTO " + strEmptyTable + " SELECT * FROM " + oDs.m_strDataSource[x,Datasource.TABLE].Trim();
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,oAdoCurrent.m_strSQL);
										//drop the current table
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + oDs.m_strDataSource[x,Datasource.TABLE].Trim());
										//create the table structure giving it the current table name
										oAdoCurrent.m_strSQL = "SELECT TOP 1 * INTO " + oDs.m_strDataSource[x,Datasource.TABLE].Trim() + " FROM " + strEmptyTable;
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,oAdoCurrent.m_strSQL);
										//delete the 1 record in the current table
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DELETE FROM " + oDs.m_strDataSource[x,Datasource.TABLE].Trim());
                                        //insert all the records from the new empty table to the current table
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"INSERT INTO " + oDs.m_strDataSource[x,Datasource.TABLE].Trim() + " SELECT *  FROM " + strEmptyTable.Trim());
										//drop the empty table
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
                                        
                                        if (oAdoCurrent.ColumnExist(oAdo.m_OleDbConnection,oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),"RX") &&
                                            oAdoCurrent.ColumnExist(oAdo.m_OleDbConnection,oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),"RXPACKAGE") &&
                                            oAdoCurrent.ColumnExist(oAdo.m_OleDbConnection,oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),"RXCYCLE"))
                                              this.ConvertRxAndRxPackageAndRxCycle(oAdoCurrent,oAdoCurrent.m_OleDbConnection,oDs.m_strDataSource[x,Datasource.TABLETYPE].Trim());

                                        //create the indexes for the empty table structure
                                        if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, strEmptyTable))
                                        {
                                            oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                                oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(), strEmptyTable.Trim());
                                        }
                                        else
                                        {
                                            if (oAdoCurrent.TableExist(oConn, strEmptyTable))
                                            {
                                                oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oConn,
                                                    oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(), strEmptyTable.Trim());
                                            }
                                        }

									}
									else
									{
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
									}

								

								}
								else
								{
									if (oAdoCurrent.ReconcileTableColumns(oAdoCurrent.m_OleDbConnection,
										oDs.m_strDataSource[x,Datasource.TABLE],
										oConn,strTempTableName))
									{
									}
								}
								
							}
						}
					}
				}
			}
			if (oAdoCurrent!=null) oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			if (oConn != null)
			{
			   oAdo.CloseConnection(oConn);
			}
			if (oDao!=null)
			{
				oDao.m_DaoWorkspace.Close();
				oDao=null; 
			}
			oAdoCurrent=null;
			oAdo=null;
		}

		private void CheckProjectReferenceDatasourceTables()
		{

			int x,y;
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
			oDs.m_strDataSourceTableName = "datasource";
			oDs.m_strScenarioId="";
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();

			ado_data_access oAdo=new ado_data_access();
			ado_data_access oAdoCurrent=null;

			dao_data_access oDao= new dao_data_access();
			string strCurrDbFile="";
			string strDbFile="";
			string strTempDbFile="";
			string strTempTableName="";
			System.Data.OleDb.OleDbConnection oConn=null;


			//copy ref_master.mdb to project db directory if file does not exist
			if (System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\db\\ref_master.mdb")==false)
				 System.IO.File.Copy(frmMain.g_oEnv.strAppDir + "\\db\\ref_master.mdb",this.ReferenceProjectDirectory + "\\db\\ref_master.mdb",true);


			
			

			//open the project db file
			oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" + 
				frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));

			oAdo.m_strSQL = "SELECT * FROM datasource " + 
				            "WHERE table_type IS NOT NULL AND " + 
                                  "UCASE(TRIM(table_type)) " + 
                                  "IN ('OWNER GROUPS'," + 
                                      "'TREE SPECIES'," + 
                                      "'FVS TREE SPECIES'," + 
                                      "'FIADB FVS VARIANT'," + 
                                      "'INVENTORIES'," + 
                                      "'TREATMENT PRESCRIPTION CATEGORIES'," + 
                                      "'TREATMENT PRESCRIPTION SUBCATEGORIES'," + 
                                      "'FRCS SYSTEM HARVEST METHOD'," + 
                                      "'FVS WESTERN TREE SPECIES TRANSLATOR'," +
                                      "'FVS EASTERN TREE SPECIES TRANSLATOR'" + 
                                      ") ORDER BY path,file";
			oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);

			if (oAdo.m_OleDbDataReader.HasRows)
			{
				while (oAdo.m_OleDbDataReader.Read())
				{
					if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
						oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
					{
						x=oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());


						
						strDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" + 
							oAdo.m_OleDbDataReader["file"].ToString().Trim();

						if (strDbFile.Trim().ToUpper() != strCurrDbFile.Trim().ToUpper())
						{
							if (oAdoCurrent!=null) 
							{  
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"owner_groups_temp"))
									 oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE owner_groups_temp");
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"tree_species_temp"))
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE tree_species_temp");
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"fvs_tree_species_temp"))
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE fvs_tree_species_temp");
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"fiadb_fvs_variant_temp"))
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE fiadb_fvs_variant_temp");
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"inventories_temp"))
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE inventories_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "frcs_system_harvest_method_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE frcs_system_harvest_method_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_category_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_category_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_subcategory_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_subcategory_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "FVS_WesternTreeSpeciesTranslator_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE FVS_WesternTreeSpeciesTranslator_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "FVS_EasternTreeSpeciesTranslator_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE FVS_EasternTreeSpeciesTranslator_temp");

								oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
							}
							else 
							{
								strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
								oDao.CreateMDB(strTempDbFile);
								System.IO.File.Copy(frmMain.g_oEnv.strAppDir.Trim() + "\\db\\ref_master.mdb",strTempDbFile,true);

								oAdoCurrent = new ado_data_access();
							}

							//make sure the db file exists
							if (System.IO.File.Exists(strDbFile)==false)
							{
								oDao.CreateMDB(strDbFile);
							}
							
							//create a link to all the pre-populated reference tables
							oDao.CreateTableLink(strDbFile,"owner_groups_temp",strTempDbFile,"owner_groups",true);
							oDao.CreateTableLink(strDbFile,"tree_species_temp",strTempDbFile,"tree_species",true);
							oDao.CreateTableLink(strDbFile,"fvs_tree_species_temp",strTempDbFile,"fvs_tree_species",true);
							oDao.CreateTableLink(strDbFile,"fiadb_fvs_variant_temp",strTempDbFile,"fiadb_fvs_variant",true);
							oDao.CreateTableLink(strDbFile,"inventories_temp",strTempDbFile,"inventories",true);
                            oDao.CreateTableLink(strDbFile, "frcs_system_harvest_method_temp", strTempDbFile, "frcs_system_harvest_method", true);
                            oDao.CreateTableLink(strDbFile, "fvs_rx_category_temp", strTempDbFile, "fvs_rx_category",true);
                            oDao.CreateTableLink(strDbFile, "fvs_rx_subcategory_temp", strTempDbFile, "fvs_rx_subcategory",true);
                            oDao.CreateTableLink(strDbFile, "FVS_WesternTreeSpeciesTranslator_temp", strTempDbFile, "FVS_WesternTreeSpeciesTranslator", true);
                            oDao.CreateTableLink(strDbFile, "FVS_EasternTreeSpeciesTranslator_temp", strTempDbFile, "FVS_EasternTreeSpeciesTranslator", true);
							oAdoCurrent.OpenConnection(oAdo.getMDBConnString(strDbFile,"",""));
							strCurrDbFile=strDbFile;

						}
                        if (oDs.m_strDataSource[x, Datasource.TABLESTATUS] == "NF") //NF=table not found
                        {
                            //create the table
                            switch (oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim().ToUpper())
                            {

                                case "OWNER GROUPS":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM owner_groups_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "TREE SPECIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM tree_species_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FVS TREE SPECIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fvs_tree_species_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FIADB FVS VARIANT":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fiadb_fvs_variant_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "INVENTORIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM inventories_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                //version 5 additions
                                case "TREATMENT PRESCRIPTION CATEGORIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fvs_rx_category_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fvs_rx_subcategory_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FRCS SYSTEM HARVEST METHOD":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM frcs_system_harvest_method_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FVS WESTERN TREE SPECIES TRANSLATOR":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM FVS_WesternTreeSpeciesTranslator_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FVS EASTERN TREE SPECIES TRANSLATOR":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM FVS_EasternTreeSpeciesTranslator_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;

                            }
                        }
                        else
                        {
                            //update columns and data
                            strTempTableName = "";
                            strInsertSql = "";
                            strUpdateSql = "";
                            string[] strColumnsArray = null;
                            string strColumnsList = "";
                            switch (oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim().ToUpper())
                            {
                                case "OWNER GROUPS":
                                    oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    strInsertSql = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE] + " FROM owner_groups_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                            oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                            oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "INVENTORIES":
                                    oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    strInsertSql = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE] + " FROM inventories_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "TREE SPECIES":
                                    strTempTableName = "tree_species_temp";


                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                                    "SELECT " + strColumnsList + " " +
                                                    "FROM tree_species_temp a " +
                                                    "WHERE  a.fvs_variant IS NOT NULL AND " +
                                                           "LEN(TRIM(a.fvs_variant)) > 0 AND " +
                                                           "NOT EXISTS " +
                                                          "(SELECT b.spcd,b.fvs_variant  " +
                                                           "FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " b " +
                                                           "WHERE a.spcd=b.spcd AND a.fvs_variant=b.fvs_variant)";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    //update any null values
                                    strColumnsList = "";
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() +
                                                "=IIF(a." + strColumnsArray[y].Trim() + " IS NULL," +
                                                     "b." + strColumnsArray[y].Trim() + "," +
                                                     "a." + strColumnsArray[y].Trim() + "),";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strUpdateSql = "UPDATE " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " a " +
                                                   "INNER JOIN tree_species_temp b " +
                                                   "ON a.spcd=b.spcd AND a.fvs_variant=b.fvs_variant " +
                                                   "SET " + strColumnsList;
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strUpdateSql);

                                    break;
                                case "FVS TREE SPECIES":
                                    strTempTableName = "fvs_tree_species_temp";

                                    //add/update columns
                                    //oAdoCurrent.ReconcileTableColumns(oAdoCurrent.m_OleDbConnection,
                                    //	oDs.m_strDataSource[x,Datasource.TABLE],oAdoCurrent.m_OleDbConnection,
                                    //	strTempTableName);

                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                        "SELECT " + strColumnsList + " " +
                                        "FROM fvs_tree_species_temp a " +
                                        "WHERE  a.fvs_variant IS NOT NULL AND " +
                                        "LEN(TRIM(a.fvs_variant)) > 0 AND " +
                                        "NOT EXISTS " +
                                        "(SELECT b.spcd,b.fvs_variant  " +
                                        "FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " b " +
                                        "WHERE a.spcd=b.spcd AND a.fvs_variant=b.fvs_variant)";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    //update any null values
                                    strColumnsList = "";
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() +
                                                "=IIF(a." + strColumnsArray[y].Trim() + " IS NULL," +
                                                "b." + strColumnsArray[y].Trim() + "," +
                                                "a." + strColumnsArray[y].Trim() + "),";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strUpdateSql = "UPDATE " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " a " +
                                        "INNER JOIN fvs_tree_species_temp b " +
                                        "ON a.spcd=b.spcd AND a.fvs_variant=b.fvs_variant " +
                                        "SET " + strColumnsList;
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strUpdateSql);
                                    break;
                                case "FIADB FVS VARIANT":
                                    strTempTableName = "fiadb_fvs_variant_temp";

                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                        "SELECT " + strColumnsList + " " +
                                        "FROM " + strTempTableName + " a " +
                                        "WHERE  a.fvs_variant IS NOT NULL AND " +
                                        "LEN(TRIM(a.fvs_variant)) > 0 AND " +
                                        "a.statecd IS NOT NULL AND " +
                                        "a.countycd IS NOT NULL AND " +
                                        "a.plot IS NOT NULL AND " +
                                        "NOT EXISTS " +
                                        "(SELECT b.statecd,b.countycd,b.plot,b.fvs_variant  " +
                                        "FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " b " +
                                        "WHERE a.statecd=b.statecd AND " +
                                              "a.countycd=b.countycd AND " +
                                              "a.plot=b.plot AND " +
                                              "a.fvs_variant=b.fvs_variant)";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);

                                    //update any null values
                                    strColumnsList = "";
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() +
                                                "=IIF(a." + strColumnsArray[y].Trim() + " IS NULL," +
                                                "b." + strColumnsArray[y].Trim() + "," +
                                                "a." + strColumnsArray[y].Trim() + "),";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strUpdateSql = "UPDATE " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " a " +
                                        "INNER JOIN " + strTempTableName + " b " +
                                        "ON a.statecd=b.statecd AND " +
                                           "a.countycd=b.countycd AND " +
                                           "a.plot=b.plot AND " +
                                           "a.fvs_variant=b.fvs_variant " +
                                        "SET " + strColumnsList;
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strUpdateSql);
                                    break;
                                //version 5 additions
                                case "TREATMENT PRESCRIPTION CATEGORIES":
                                    if ((int)oAdoCurrent.getRecordCount(oAdoCurrent.m_OleDbConnection, "SELECT COUNT(*) FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim(), "TEMP") == 0)
                                    {
                                        strTempTableName = "fvs_rx_category_temp";
                                        strColumnsList="";
                                        //insert any new tree species records
                                        strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                        for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                        {
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                        }
                                        strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                        strInsertSql = "INSERT INTO " +
                                            oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                            "SELECT " + strColumnsList + " " +
                                            "FROM " + strTempTableName + " a ";
                                           
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                          
                                        
                                    }
                                    break;
                                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                                     if ((int)oAdoCurrent.getRecordCount(oAdoCurrent.m_OleDbConnection, "SELECT COUNT(*) FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim(), "TEMP") == 0)
                                    {
                                        strTempTableName = "fvs_rx_subcategory_temp";
                                        strColumnsList="";
                                        //insert any new tree species records
                                        strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                        for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                        {
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                        }
                                        strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                        strInsertSql = "INSERT INTO " +
                                            oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                            "SELECT " + strColumnsList + " " +
                                            "FROM " + strTempTableName + " a ";
                                           
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                          
                                            
                                    
                                    }
                                    break;
                                case "FRCS SYSTEM HARVEST METHOD":
                                        oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                        frmMain.g_oTables.m_oReference.CreateFRCSHarvestMethodTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, oDs.m_strDataSource[x, Datasource.TABLE]);
                                        strTempTableName = "frcs_system_harvest_method_temp";
                                        strColumnsList="";
                                        //insert any new tree species records
                                        strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                        for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                        {
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                        }
                                        strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                        strInsertSql = "INSERT INTO " +
                                            oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                            "SELECT " + strColumnsList + " " +
                                            "FROM " + strTempTableName + " a ";
                                           
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    break;
                                case "FVS WESTERN TREE SPECIES TRANSLATOR":
                                    oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    frmMain.g_oTables.m_oReference.CreateFVSWesternSpeciesTranslatorTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, oDs.m_strDataSource[x, Datasource.TABLE]);
                                    strTempTableName = "FVS_WesternTreeSpeciesTranslator_temp";
                                    strColumnsList = "";
                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                        "SELECT " + strColumnsList + " " +
                                        "FROM " + strTempTableName + " a ";

                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    break;
                                case "FVS EASTERN TREE SPECIES TRANSLATOR":
                                    oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    frmMain.g_oTables.m_oReference.CreateFVSEasternSpeciesTranslatorTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, oDs.m_strDataSource[x, Datasource.TABLE]);
                                    strTempTableName = "FVS_EasternTreeSpeciesTranslator_temp";
                                    strColumnsList = "";
                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                        "SELECT " + strColumnsList + " " +
                                        "FROM " + strTempTableName + " a ";

                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    break;
                            }
                        }
					}
				}
			}
			if (oAdoCurrent!=null) 
			{
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"owner_groups_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE owner_groups_temp");
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"tree_species_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE tree_species_temp");
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"fvs_tree_species_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE fvs_tree_species_temp");
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"fiadb_fvs_variant_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE fiadb_fvs_variant_temp");
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"inventories_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE inventories_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "frcs_system_harvest_method_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE frcs_system_harvest_method_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_category_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_category_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_subcategory_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_subcategory_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "FVS_WesternTreeSpeciesTranslator_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE FVS_WesternTreeSpeciesTranslator_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "FVS_EasternTreeSpeciesTranslator_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE FVS_EasternTreeSpeciesTranslator_temp");
				oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
			}
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			if (oConn != null)
			{
				oAdo.CloseConnection(oConn);
			}
			if (oDao!=null)
			{
				oDao.m_DaoWorkspace.Close();
				oDao=null; 
			}
			oAdoCurrent=null;
			oAdo=null;
            strCurrDbFile = "";
            strDbFile = "";
            strTempDbFile = "";
            strTempTableName = "";
            oConn = null;
            //
            //fvs commands table
            //
            

            
            oAdo = new ado_data_access();
            oDao = new dao_data_access();
            string strNewDbFile = "";
            string strOldDbFile = "";

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            oAdo.m_strSQL = "SELECT * FROM datasource " +
                            "WHERE table_type IS NOT NULL AND " +
                                  "UCASE(TRIM(table_type)) " +
                                  "IN ('FVS COMMANDS') ORDER BY path,file";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
                        oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
                    {
                        x = oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());



                        strOldDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" +
                            oAdo.m_OleDbDataReader["file"].ToString().Trim();

                        if (System.IO.File.Exists(strOldDbFile))
                        {
                            oAdoCurrent = new ado_data_access();
                            //
                            //get all table names in the old file
                            //
                            string[] strTablesArray = null;
                            oDao.getTableNames(strOldDbFile, ref strTablesArray);
                            //
                            //get the new file name
                            //
                            strNewDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
                            //
                            //create the new file and open it
                            //
                            oDao.CreateMDB(strNewDbFile);
                            oDao.OpenDb(strNewDbFile);
                            //
                            //create a link in the new file to the new table
                            //
                            oDao.CreateTableLink(oDao.m_DaoDatabase, "fvs_commands_temp", frmMain.g_oEnv.strAppDir + "\\db\\ref_fvscommands.mdb", "fvs_commands");
                            //
                            //create a link to all tables from the old file to the new file
                            //
                            if (strTablesArray != null)
                            {
                               
                                for (y = 0; y <= strTablesArray.Length - 1; y++)
                                {
                                    if (strTablesArray[y] == null) break;
                                    if (strTablesArray[y].Trim().Length > 0)
                                    {
                                        //do not link to the table we are replacing
                                        if (strTablesArray[y].Trim().ToUpper() !=
                                            oDs.m_strDataSource[x, Datasource.TABLE].Trim().ToUpper())
                                        {
                                            oDao.CreateTableLink(oDao.m_DaoDatabase, strTablesArray[y].Trim() + "_temp", strOldDbFile, strTablesArray[y].Trim());
                                        }


                                    }
                                }
                                oDao.m_DaoWorkspace.Close();
                                oDao = null;
                                //open oledb connection to new file
                                oAdoCurrent.OpenConnection(oAdoCurrent.getMDBConnString(strNewDbFile,"",""));
                                for (y = 0; y <= strTablesArray.Length - 1; y++)
                                {
                                    if (strTablesArray[y] == null) break;
                                    //import linked tables
                                    if (strTablesArray[y].Trim().ToUpper() == oDs.m_strDataSource[x, Datasource.TABLE].Trim().ToUpper())
                                    {
                                        
                                    }
                                    else
                                    {
                                        oAdoCurrent.m_strSQL = "SELECT * INTO " + strTablesArray[y].Trim() + " " +
                                                               "FROM " + strTablesArray[y].Trim() + "_temp";
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                        //drop the linked table
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strTablesArray[y].Trim() + "_temp");
                                    }
                                }
                               
                            }
                            //
                            //import linked fvs_commands table
                            //
                            if (oAdoCurrent.m_OleDbConnection.State == System.Data.ConnectionState.Closed)
                                oAdoCurrent.OpenConnection(oAdoCurrent.getMDBConnString(strNewDbFile, "", ""));
                            oAdoCurrent.m_strSQL = "SELECT * " + 
                                                   "INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " " + 
                                                   "FROM fvs_commands_temp";
                            oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                            oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                            oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_commands_temp");
                            oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
                            //
                            //delete old file and copy new file
                            //
                            System.IO.File.Delete(strOldDbFile);
                            System.IO.File.Copy(strNewDbFile,strOldDbFile,true);

                        }
                        else
                        {
                            //copy ref_master.mdb to project db directory if file does not exist
                            if (System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\db\\ref_fvscommands.mdb") == false)
                                System.IO.File.Copy(frmMain.g_oEnv.strAppDir + "\\db\\ref_fvscommands.mdb", this.ReferenceProjectDirectory + "\\db\\ref_fvscommands.mdb", true);
                        }

                        
                    }
                }
            }
            
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            if (oConn != null)
            {
                oAdo.CloseConnection(oConn);
            }
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
            oAdoCurrent = null;
            oAdo = null;
		}



        private void UpdateFVSPlotVariantAssignmentsTable()
        {

            int x;
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            ado_data_access oAdo = new ado_data_access();
            ado_data_access oAdoCurrent = null;

            dao_data_access oDao = new dao_data_access();
            string strCurrDbFile = "";
            string strCurrTableName="";

            string strSourceDbFile="";
            string strSourceTableName = "";


            frmMain.g_sbpInfo.Text = "Version Update: Update FIADB FVS Plot Variant Reference Table...Stand by";

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            //get the MDB file name and table name
            oAdo.m_strSQL = "SELECT * FROM datasource " +
                            "WHERE table_type IS NOT NULL AND " +
                                  "UCASE(TRIM(table_type)) " +
                                  "IN ('FIADB FVS VARIANT') " + 
                            "ORDER BY path,file";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
                        oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
                    {



                        x = oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());

                        strCurrDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" +
                                        oAdo.m_OleDbDataReader["file"].ToString().Trim();

                        strCurrTableName = oDs.m_strDataSource[x, Datasource.TABLE].Trim();

                        oDao = new dao_data_access();
                        //
                        //CREATE THE MDB IF NOT EXIST
                        //
                        if (oDs.m_strDataSource[x, Datasource.FILESTATUS] == "NF") //NF=table not found
                        {
                            oDao.CreateMDB(strCurrDbFile);
                        }
                        strSourceDbFile = frmMain.g_oEnv.strAppDir.Trim() + "\\" + Tables.Reference.DefaultFiadbFVSVariantTableDbFile;
                        strSourceTableName = Tables.Reference.DefaultFiadbFVSVariantTableName;
                        //
                        //DELETE ANY OLD TABLES
                        //
                        if (oDao.TableExists(strCurrDbFile, strCurrTableName + "_temp"))
                            oDao.DeleteTableFromMDB(strCurrDbFile, strCurrTableName + "_temp");
                        if (oDao.TableExists(strCurrDbFile, strCurrTableName + "_temp2"))
                            oDao.DeleteTableFromMDB(strCurrDbFile, strCurrTableName + "_temp2");
                        //
                        //CREATE TABLE LINK
                        //
                        oDao.CreateTableLink(strCurrDbFile, strCurrTableName + "_temp", strSourceDbFile, strSourceTableName);
                        oDao.m_DaoWorkspace.Close();
                        oDao = null;
                        System.Threading.Thread.Sleep(4000);
                        //
                        //INSTANTIATE NEW ADO OBJECT
                        //
                        oAdoCurrent = new ado_data_access();
                        //
                        //OPEN ADO CONNECTION
                        //
                        oAdoCurrent.OpenConnection(oAdoCurrent.getMDBConnString(strCurrDbFile, "", ""));
                        //
                        //COPY RECORDS FROM LINK TABLE
                        //
                        oAdo.m_strSQL = "SELECT * INTO " + strCurrTableName + "_temp2 FROM " + strCurrTableName + "_temp";
                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdo.m_strSQL);
                        //
                        //DROP THE LINK
                        //
                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE " + strCurrTableName + "_temp");
                        //
                        //DROP THE OLD TABLE
                        //
                        if (oDs.m_strDataSource[x, Datasource.TABLESTATUS] == "F")
                            oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE " + strCurrTableName);
                        //
                        //COPY TEMP TABLE TO PRODUCTION
                        // 
                        oAdo.m_strSQL = "SELECT * INTO " + strCurrTableName + " FROM " + strCurrTableName + "_temp2";
                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdo.m_strSQL);
                        //
                        //DROP TEMP TABLE
                        //
                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE " + strCurrTableName + "_temp2");
                        //
                        //CLOSE ADOCURRENT CONNECTION
                        //
                        oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
                        oAdoCurrent = null;
                        break;
                    }
                }
            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }

		/// <summary>
		/// Check and update the scenario datasource table with the latest version of table type entries
		/// </summary>
		private void CheckCoreScenarioDatasourceTableRecords()
		{
		    
			int x;
			int y;
			bool bCore;

			FIA_Biosum_Manager.Datasource oDsScenario = new Datasource();
			FIA_Biosum_Manager.Datasource oDsProject = new Datasource();
			ado_data_access oAdoScenario=new ado_data_access();
			ado_data_access oAdoProject = new ado_data_access();
			

			oDsScenario.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb";
			
			string strDbFile="";
			string strSQL="";
			string strScenarioList="";
			string[] strScenarioArray=null;

			//open the core scenario db file
			oAdoScenario.OpenConnection(oAdoScenario.getMDBConnString(oDsScenario.m_strDataSourceMDBFile,"",""));
            //remove obsolete table types
            oAdoScenario.m_strSQL = "DELETE FROM scenario_datasource WHERE " + 
                "TRIM(UCASE(table_type)) IN (" +
                "'FIRE AND FUEL EFFECTS'," +
                "'TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS'," +
                "'HARVEST COSTS')";
            oAdoScenario.SqlNonQuery(oAdoScenario.m_OleDbConnection, oAdoScenario.m_strSQL);
			//get all the scenarios
			oAdoScenario.SqlQueryReader(oAdoScenario.m_OleDbConnection,"SELECT scenario_id FROM scenario");
			if (oAdoScenario.m_OleDbDataReader.HasRows)
			{
				while (oAdoScenario.m_OleDbDataReader.Read())
				{
					if (oAdoScenario.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
					{
						strScenarioList = strScenarioList + oAdoScenario.m_OleDbDataReader["scenario_id"].ToString().Trim() + ",";
					}
				}
				
			}
			oAdoScenario.m_OleDbDataReader.Close();

			if (strScenarioList.Trim().Length > 0)
			{
				strScenarioList = strScenarioList.Substring(0,strScenarioList.Trim().Length - 1);
				strScenarioArray = frmMain.g_oUtils.ConvertListToArray(strScenarioList,",");
				//open the project db file
				oAdoProject.OpenConnection(oAdoProject.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" + 
						frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));

				//cycle through each scenario datasource to ensure all the latest datasources exist
				for (x=0;x<=strScenarioArray.Length-1;x++)
				{
					//make sure all the latest datasources are in the scenario datasource
					//get the project datasources
					oAdoProject.SqlQueryReader(oAdoProject.m_OleDbConnection,"select * from datasource");
					if (oAdoProject.m_intError==0)
					{
						//load the scenario datasource
						oDsScenario.m_strDataSourceTableName = "scenario_datasource";
						oDsScenario.m_strScenarioId=strScenarioArray[x].Trim();
						oDsScenario.LoadTableColumnNamesAndDataTypes=false;
						oDsScenario.LoadTableRecordCount=false;
						oDsScenario.populate_datasource_array();
						while (oAdoProject.m_OleDbDataReader.Read())
						{
							bCore=false;
							switch (oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper())
							{
								case "PLOT":
									bCore=true;
									break;
								case "CONDITION":
									bCore = true;
									break;
								//case "HARVEST COSTS":
								//	bCore = true;
								//	break;
								case "TREE DIAMETER GROUPS":
									bCore = true;
									break;
								case "TREATMENT PRESCRIPTIONS":
									bCore = true;
									break;
								case "TREE SPECIES GROUPS":
									bCore = true;
									break;
								//case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
								//	bCore = true;
								//	break;
								case "TRAVEL TIMES":
									bCore = true;
									break;
								case "PROCESSING SITES":
									bCore = true;
									break;
								//case "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES":
								//	bCore = true;
								//	break;
								case "PLOT AND CONDITION RECORD AUDIT":
									bCore = true;
									break;
								case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
									bCore = true;
									break;
                                case "TREATMENT PACKAGES":
                                    bCore = true;
                                    break;

								default:
									break;
							}
							if (bCore == true)
							{
								//see if the project datasource table type exists in the scenario datasource table
								y=oDsScenario.getValidTableNameRow(oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper());
								if (y==-1)
								{
									strSQL = "INSERT INTO scenario_datasource (scenario_id,table_type,Path,file,table_name) VALUES " + "('" + strScenarioArray[x].Trim() + "'," + 
										"'" + oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim() + "'," + 
										"'" + oAdoProject.m_OleDbDataReader["path"].ToString().Trim() + "'," + 
										"'" + oAdoProject.m_OleDbDataReader["file"].ToString().Trim() + "'," +  
										"'" + oAdoProject.m_OleDbDataReader["table_name"].ToString().Trim() + "');";
									oAdoScenario.SqlNonQuery(oAdoScenario.m_OleDbConnection,strSQL);
								}
							}

						}
						oAdoProject.m_OleDbDataReader.Close();
					}
				}
				oAdoProject.CloseConnection(oAdoProject.m_OleDbConnection);
			}
			oAdoScenario.CloseConnection(oAdoScenario.m_OleDbConnection);
		}
        /// <summary>
        /// Check and update the scenario datasource table with the latest version of table type entries
        /// </summary>
        private void CheckProcessorScenarioDatasourceTableRecords()
        {

            int x;
            int y;
            bool bCore;

            FIA_Biosum_Manager.Datasource oDsScenario = new Datasource();
            FIA_Biosum_Manager.Datasource oDsProject = new Datasource();
            ado_data_access oAdoScenario = new ado_data_access();
            ado_data_access oAdoProject = new ado_data_access();


            oDsScenario.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\Processor\\db\\scenario_processor_rule_definitions.mdb";
            if (System.IO.File.Exists(oDsScenario.m_strDataSourceMDBFile) == false)
            {
                frmMain.g_oFrmMain.frmProject.uc_project1.CreateProcessorScenarioRuleDefinitionDbAndTables(oDsScenario.m_strDataSourceMDBFile);

            }

            string strDbFile = "";
            string strSQL = "";
            string strScenarioList = "";
            string[] strScenarioArray = null;

            //open the core scenario db file
            oAdoScenario.OpenConnection(oAdoScenario.getMDBConnString(oDsScenario.m_strDataSourceMDBFile, "", ""));
            
            //remove obsolete table types
            oAdoScenario.m_strSQL = "DELETE FROM scenario_datasource WHERE " +
                "TRIM(UCASE(table_type)) IN (" +
                "'FIRE AND FUEL EFFECTS'," +
                "'TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS'," +
                "'HARVEST COSTS'," +
                "'PLOT AND CONDITION RECORD AUDIT'," +
                "'PLOT, CONDITION AND TREATMENT RECORD AUDIT'," + 
                "'FVS TREE LIST FOR PROCESSOR'," +
                "'ADDITIONAL HARVEST COSTS'," +
                "'TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES')";
            oAdoScenario.SqlNonQuery(oAdoScenario.m_OleDbConnection, oAdoScenario.m_strSQL);
            //get all the scenarios
            oAdoScenario.SqlQueryReader(oAdoScenario.m_OleDbConnection, "SELECT scenario_id FROM scenario");
            if (oAdoScenario.m_OleDbDataReader.HasRows)
            {
                while (oAdoScenario.m_OleDbDataReader.Read())
                {
                    if (oAdoScenario.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                    {
                        strScenarioList = strScenarioList + oAdoScenario.m_OleDbDataReader["scenario_id"].ToString().Trim() + ",";
                    }
                }

            }
            oAdoScenario.m_OleDbDataReader.Close();

            if (strScenarioList.Trim().Length > 0)
            {
                strScenarioList = strScenarioList.Substring(0, strScenarioList.Trim().Length - 1);
                strScenarioArray = frmMain.g_oUtils.ConvertListToArray(strScenarioList, ",");
                //open the project db file
                oAdoProject.OpenConnection(oAdoProject.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                        frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

                //cycle through each scenario datasource to ensure all the latest datasources exist
                for (x = 0; x <= strScenarioArray.Length - 1; x++)
                {
                    //make sure all the latest datasources are in the scenario datasource
                    //get the project datasources
                    oAdoProject.SqlQueryReader(oAdoProject.m_OleDbConnection, "select * from datasource");
                    if (oAdoProject.m_intError == 0)
                    {
                        //load the scenario datasource
                        oDsScenario.m_strDataSourceTableName = "scenario_datasource";
                        oDsScenario.m_strScenarioId = strScenarioArray[x].Trim();
                        oDsScenario.LoadTableColumnNamesAndDataTypes = false;
                        oDsScenario.LoadTableRecordCount = false;
                        oDsScenario.populate_datasource_array();
                        while (oAdoProject.m_OleDbDataReader.Read())
                        {
                            bCore = false;
                            switch (oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper())
                            {
                                case "PLOT":
                                    bCore = true;
                                    break;
                                case "CONDITION":
                                    bCore = true;
                                    break;
                                //case "HARVEST COSTS":
                                //	bCore = true;
                                //	break;
                                case "TREE DIAMETER GROUPS":
                                    bCore = true;
                                    break;
                                case "TREATMENT PRESCRIPTIONS":
                                    bCore = true;
                                    break;
                                case "TREE SPECIES GROUPS":
                                    bCore = true;
                                    break;
                                //case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
                                //	bCore = true;
                                //	break;
                                case "TRAVEL TIMES":
                                    bCore = true;
                                    break;
                                case "PROCESSING SITES":
                                    bCore = true;
                                    break;
                                //case "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES":
                                //	bCore = true;
                                //	break;
                                //case "PLOT AND CONDITION RECORD AUDIT":
                                 //   bCore = true;
                                  //  break;
                                //case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
                                //    bCore = true;
                                //    break;
                                case "TREATMENT PACKAGES":
                                    bCore = true;
                                    break;
                                case "FRCS SYSTEM HARVEST METHOD":
                                    bCore = true;
                                    break;
                                //case "ADDITIONAL HARVEST COSTS":
                                //    bCore = true;
                                //    break;
                                case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                                    bCore = true;
                                    break;
                                default:
                                    break;
                            }
                            if (bCore == true)
                            {
                                //see if the project datasource table type exists in the scenario datasource table
                                y = oDsScenario.getValidTableNameRow(oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper());
                                if (y == -1)
                                {
                                    strSQL = "INSERT INTO scenario_datasource (scenario_id,table_type,Path,file,table_name) VALUES " + "('" + strScenarioArray[x].Trim() + "'," +
                                        "'" + oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim() + "'," +
                                        "'" + oAdoProject.m_OleDbDataReader["path"].ToString().Trim() + "'," +
                                        "'" + oAdoProject.m_OleDbDataReader["file"].ToString().Trim() + "'," +
                                        "'" + oAdoProject.m_OleDbDataReader["table_name"].ToString().Trim() + "');";
                                    oAdoScenario.SqlNonQuery(oAdoScenario.m_OleDbConnection, strSQL);
                                }
                            }

                        }
                        oAdoProject.m_OleDbDataReader.Close();
                    }
                }
                oAdoProject.CloseConnection(oAdoProject.m_OleDbConnection);
            }
            oAdoScenario.CloseConnection(oAdoScenario.m_OleDbConnection);
        }
        private void CheckRxValues()
        {
            int x,y,z;
            string strRxPackage="";
            string str="";
            string str2 = "";
            string strTable="";
            string strSourceFolder = "";
            string strDestFolder = "";

            ado_data_access oAdo = new ado_data_access();

          
            
            Queries oQueries = new Queries();
            oQueries.m_oFvs.LoadDatasource = true;
            oQueries.m_oFIAPlot.LoadDatasource = true;
            oQueries.m_oProcessor.LoadDatasource = true;
            oQueries.LoadDatasources(true);

           
            oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""));
           
            //
            //GET RX VALUES
            //
            string[] strRxArray = frmMain.g_oUtils.ConvertListToArray(oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT RX FROM " + oQueries.m_oFvs.m_strRxTable, ""), ",");
            //
            //GET VARIANT VALUES
            //
            string[] strVariantArray = frmMain.g_oUtils.ConvertListToArray(oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT DISTINCT FVS_VARIANT FROM " + oQueries.m_oFIAPlot.m_strPlotTable, ""), ",");
            if (strRxArray != null && strRxArray.Length > 0)
            {
                if (strRxArray[0].Trim().Length == 1)
                {
                    //update rx 1 character alpha to 3  numeric character 701 to 725
                    
                    ConvertRx(oAdo, oAdo.m_OleDbConnection, oQueries.m_oFvs.m_strRxTable);
                    //update CATID to 700 series which is custom defined rx
                    oAdo.m_strSQL = "UPDATE " + oQueries.m_oFvs.m_strRxTable + " " +
                                    "SET CATID=4";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    //update fvs_tree table rx values
                    ConvertRx(oAdo, oAdo.m_OleDbConnection, oQueries.m_oFvs.m_strFvsTreeTable);
                    //create and assign an rxpackage to each rx
                    string[] strNewRxArray = frmMain.g_oUtils.ConvertListToArray(oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT RX FROM " + oQueries.m_oFvs.m_strRxTable, ""), ",");
                    for (x = 0; x <= strNewRxArray.Length - 1; x++)
                    {
                        z=x+1;
                        if (z.ToString().Trim().Length == 1)
                            strRxPackage = "00" + z.ToString().Trim();
                        else if (z.ToString().Trim().Length == 2)
                            strRxPackage = "0" + z.ToString().Trim();
                        else
                            strRxPackage = z.ToString().Trim();
                          
                        oAdo.m_strSQL = "INSERT INTO " +
                                        oQueries.m_oFvs.m_strRxPackageTable + " " +
                                        "(rxpackage,rxcycle_length,simyear1_rx) VALUES " +
                                        "('" + strRxPackage + "','10','" + strNewRxArray[x].Trim() + "')";
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    }
                    if (strVariantArray != null && strVariantArray.Length > 0)
                    {
                        //copy any fvs input and output files 
                        for (x = 0; x <= strVariantArray.Length - 1; x++)
                        {
                            if (strVariantArray[x].Trim().Length > 0)
                            {
                                strSourceFolder = ReferenceProjectDirectory + "\\fvs\\db\\in\\" + strVariantArray[x].Trim();
                                strDestFolder = ReferenceProjectDirectory + "\\fvs\\data\\" + strVariantArray[x].Trim();
                                if (System.IO.Directory.Exists(strDestFolder) == false)
                                {
                                    //create fvs cutlist table
                                    //CA_TREE_CUTLIST.MDB
                                    strTable = "fvs_tree_IN_" + strVariantArray[x].Trim() + "_TREE_CUTLIST";
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection,frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorInTableSQL(strTable));

                                    //create fvs file input  folder
                                    System.IO.Directory.CreateDirectory(strDestFolder);
                                    frmMain.g_sbpInfo.Text = "Version Update: Copying FVS Input And Output Files For Variant " + strVariantArray[x].Trim() + "...Stand By";
                                    //copy files
                                    utils.FileCopy(strSourceFolder, strDestFolder, false);
                                    //copy any fvs output files 
                                    strSourceFolder = ReferenceProjectDirectory + "\\fvs\\db\\out";
                                    z=0;
                                    for (y=0;y<=strRxArray.Length - 1;y++)
                                    {

                                        z=y+1;
                                        if (z.ToString().Trim().Length == 1)
                                            strRxPackage = "00" + z.ToString().Trim();
                                        else if (z.ToString().Trim().Length == 2)
                                            strRxPackage = "0" + z.ToString().Trim();
                                        else
                                            strRxPackage = z.ToString().Trim();

                                        if (strRxArray[y] != null && strRxArray[y].Length > 0)
                                        {

                                            //FVSOUT_CA_P000-050-051-000-050.MDB
                                            str = "FVSOUT_" + strVariantArray[x].Trim() + 
                                                   "_" + strRxArray[y].Trim() + ".mdb";
                                            str2 = "FVSOUT_" + strVariantArray[x].Trim() +
                                                   "_P" + strRxPackage + "-" + strNewRxArray[y].Trim() +
                                                   "-000-000-000.mdb";
                                            frmMain.g_sbpInfo.Text = "Version Update: Copy " + str + " to " + str2 + "...Stand By";
                                            if (System.IO.File.Exists(strSourceFolder + "\\" + str) == true &&
                                                System.IO.File.Exists(strDestFolder + "\\" + str2) == false)
                                            {
                                                System.IO.File.Copy(strSourceFolder + "\\" + str,
                                                                    strDestFolder + "\\" + str2,true);
                                            }

                                            frmMain.g_sbpInfo.Text = "Version Update: Insert " + oQueries.m_oFvs.m_strFvsTreeTable + " to " + strTable + " for variant " + strVariantArray[x].Trim() + "...Stand By";
                                            oAdo.m_strSQL= "INSERT INTO " + strTable + " " +
                                                           "SELECT * FROM " + oQueries.m_oFvs.m_strFvsTreeTable + " " + 
                                                           "WHERE fvs_variant='" + strVariantArray[x].Trim() + "' AND " + 
                                                                 "RX='" + strNewRxArray[y].Trim() + "'";

                                          
                                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                            frmMain.g_sbpInfo.Text = "Version Update: Update RxPackage to " + strRxPackage + " AND RxCycle to 1 for variant " + strVariantArray[x].Trim() + "...Stand By";
                                            oAdo.m_strSQL = "UPDATE " + strTable + " " +
                                                            "SET rxpackage='" + strRxPackage + "'," +
                                                                "rxcycle='1' " +
                                                            "WHERE rx='" + strNewRxArray[y].Trim() + "'";

                                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                        }
                                    }

                                    
                                }
                                

                            }
                        }
                        

                       

                    }
                    
                    
                    


                }

            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            if (strRxArray != null && strRxArray.Length > 0)
            {
                if (strRxArray[0].Trim().Length == 1)
                {
                    dao_data_access oDao = new dao_data_access();
                    //
                    //make sure treelist folder exist
                    //
                    strDestFolder = ReferenceProjectDirectory + "\\fvs\\data\\TreeLists";
                    if (System.IO.Directory.Exists(strDestFolder) == false)
                    {

                        //create tree cutlist folder
                        System.IO.Directory.CreateDirectory(strDestFolder);
                    }
                    //
                    //create links to variant cutlist table
                    //
                    for (x = 0; x <= strVariantArray.Length - 1; x++)
                    {
                        strTable = "fvs_tree_IN_" + strVariantArray[x].Trim() + "_TREE_CUTLIST";
                        //only process variant if the strTable exists
                        if (oDao.TableExists(oQueries.m_strTempDbFile,strTable))
                        {

                            if (strVariantArray[x] != null &&
                                strVariantArray[x].Trim().Length > 0)
                            {
                                //CA_TREE_CUTLIST.MDB
                                strTable = "fvs_tree_IN_" + strVariantArray[x].Trim() + "_TREE_CUTLIST";
                                if (System.IO.File.Exists(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB"))
                                {
                                    System.IO.File.Delete(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB");
                                    System.Threading.Thread.Sleep(2000);

                                }
                                oDao.CreateMDB(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB");
                                oDao.CreateTableLink(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB", strTable, oQueries.m_strTempDbFile, strTable);
                             }
                        }
                    }
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;
                }
            }
            
            if (strRxArray != null && strRxArray.Length > 0)
            {
                if (strRxArray[0].Trim().Length == 1)
                {
                    
                    //
                    //make sure treelist folder exist
                    //
                    strDestFolder = ReferenceProjectDirectory + "\\fvs\\data\\TreeLists";
                   
                    //
                    //create cutlist mdb and copy table to it
                    //
                    for (x = 0; x <= strVariantArray.Length - 1; x++)
                    {
                        if (System.IO.File.Exists(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB"))
                        {
                            oAdo.OpenConnection(oAdo.getMDBConnString(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB", "", ""));
                            strTable = "fvs_tree_IN_" + strVariantArray[x].Trim() + "_TREE_CUTLIST";
                            //only process variant if the strTable exists
                            if (oAdo.TableExist(oAdo.m_OleDbConnection, strTable))
                            {

                                if (strVariantArray[x] != null &&
                                    strVariantArray[x].Trim().Length > 0)
                                {
                                    //CA_TREE_CUTLIST.MDB
                                    if ((int)oAdo.getRecordCount(oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM " + strTable, "TEMP") > 0)
                                    {
                                        frmMain.g_sbpInfo.Text = "Version Update: Copy table " + strTable + " into FVS_Tree table for variant " + strVariantArray[x].Trim() + "...Stand By";
                                        oAdo.m_strSQL = "SELECT * INTO fvs_tree FROM " + strTable;
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTable);
                                    }
                                    else
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTable);
                                    }





                                }
                            }
                            oAdo.CloseConnection(oAdo.m_OleDbConnection);
                        }
                    }
                    


                }
            }
            









        }
        private void CheckFVSOutPrePostValues()
        {
            int x,y;
            
            string strAlpha="A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z";
            string[] strAlphaArray = frmMain.g_oUtils.ConvertListToArray(strAlpha, ",");
            string strSourceFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\FVS\\DB\\biosum_fvsout_prepost_rx.mdb";
            if (System.IO.File.Exists(strSourceFile) == false) return;

            Queries oQueries = new Queries();
            oQueries.m_oFvs.LoadDatasource = true;
            oQueries.LoadDatasources(true);
            
            ado_data_access oAdo = new ado_data_access();

            //
            //get the number of treatments
            //
            int intRxCount = (int)oAdo.getRecordCount(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""), "SELECT COUNT(*) FROM " + oQueries.m_oFvs.m_strRxTable, "temp");
            
            oAdo.OpenConnection(oAdo.getMDBConnString(strSourceFile, "", ""));



            string[] strTableArray = oAdo.getTableNames(oAdo.m_OleDbConnection);

            if (strTableArray != null)
            {
                
                for (x = 0; x <= strTableArray.Length - 1; x++)
                {
                    if (strTableArray[x] != null &&
                        strTableArray[x].Trim().Length >= 7)
                    {
                        if (strTableArray[x].Substring(0, 7).ToUpper() == "PRE_FVS")
                        {
                            if (oAdo.ColumnExist(oAdo.m_OleDbConnection, strTableArray[x], "RX"))
                            {
                            }
                            else
                            {
                                oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RXPACKAGE TEXT (3)";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RX TEXT (3)";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RXCYCLE TEXT (1)";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                oAdo.m_strSQL = "UPDATE " + strTableArray[x] + " SET RX='A'";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                for (y = 2; y <= intRxCount; y++)
                                {
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "PRETEMP"))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE PRETEMP");
                                    }
                                    oAdo.m_strSQL = "SELECT * INTO PRETEMP FROM " + strTableArray[x] + " WHERE TRIM(RX)='A'";
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                    oAdo.m_strSQL = "UPDATE PRETEMP SET RX='" + strAlphaArray[y - 1] + "'";
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                    oAdo.m_strSQL = "INSERT INTO " + strTableArray[x] + " SELECT * FROM PRETEMP";
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                }
                                this.ConvertRxAndRxPackageAndRxCycle(oAdo, oAdo.m_OleDbConnection, strTableArray[x]);

                            }

                        }
                        else
                        {
                            if (strTableArray[x] != null &&
                                strTableArray[x].Trim().Length > 7)
                            {
                                if (strTableArray[x].Substring(0, 8).ToUpper() == "POST_FVS")
                                {
                                    if (oAdo.ColumnExist(oAdo.m_OleDbConnection, strTableArray[x], "RXPACKAGE"))
                                    {
                                    }
                                    else
                                    {
                                        oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ALTER COLUMN RX TEXT (3)";
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                        oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RXPACKAGE TEXT (3)";
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                        oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RXCYCLE TEXT (1)";
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                        this.ConvertRxAndRxPackageAndRxCycle(oAdo, oAdo.m_OleDbConnection, strTableArray[x]);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);




        }
        private void ConvertRx(ado_data_access p_oAdo, 
                               System.Data.OleDb.OleDbConnection p_oConn,
                               string p_strTableName)
        {
            p_oAdo.m_strSQL = "UPDATE " + p_strTableName + " " +
                                    "SET RX = IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) > 0," +
                                       "'7' + IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) <10," +
                                       "'0' + CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))," +
                                       "CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))),RX)";

            p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);



        }

        private void ConvertRxAndRxPackageAndRxCycle(ado_data_access p_oAdo,
                               System.Data.OleDb.OleDbConnection p_oConn,
                               string p_strTableName)
        {
            p_oAdo.m_strSQL = "UPDATE " + p_strTableName + " " +
                              "SET RXPACKAGE = " +
                                      "IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) > 0," +
                                         "'0' + IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) <10," +
                                         "'0' + CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))," +
                                         "CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))),RXPACKAGE)," +
                                  "RXCYCLE=IIF(LEN(TRIM(RXCYCLE)) > 0,RXCYCLE,'1')," +
                                  "RX=IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) > 0," +
                                         "'7' + IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) <10," +
                                         "'0' + CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))," +
                                         "CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))),RX)";
            p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);

        }
        private void CleanUp()
        {
            ado_data_access oAdo = new ado_data_access();

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            //remove obsolete table types
            //TREE VOL VAL AND HARVEST COSTS ARE NOW SCENARIO BASED IN THE PROCESSOR
            oAdo.m_strSQL = "DELETE FROM datasource WHERE " +
                "TRIM(UCASE(table_type)) IN (" +
                "'FIRE AND FUEL EFFECTS'," +
                "'TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS'," +
                "'HARVEST COSTS'," +
                "'FVS TREE LIST FOR PROCESSOR')";

            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }
        private void UpgradeFVSOutTreeListFiles()
        {
            int x,y;
            string strFile="";
            List<string> strVariant = new List<string>(); //hold the variants that have an existing variant_tree_cutlist.mdb
            string[] strPackageArray = null;
            
            string[] strFiles = System.IO.Directory.GetFiles(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\FVS\\Data\\TreeLists","*.mdb");
            if (strFiles != null && strFiles.Length > 0)
            {
                frmMain.g_sbpInfo.Text = "Version Update: Update FVS_CUTLIST...Stand By";
                ado_data_access oAdo = new ado_data_access();
                RxTools oRxTools = new RxTools();
                Queries oQueries = new Queries();
                RxPackageItem_Collection oRxPackageItemCollection = new RxPackageItem_Collection();
			    oQueries.m_oFvs.LoadDatasource=true;
			    oQueries.m_oFIAPlot.LoadDatasource=true;
			    oQueries.LoadDatasources(true);
                //
                //open the work file containing the links
                //
                oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile,"",""));
                //
                //get all the variants
                //
			    string strVariantsList = oRxTools.GetListOfFVSVariantsInPlotTable(oAdo,oAdo.m_OleDbConnection,oQueries.m_oFIAPlot.m_strPlotTable);
                string[] strVariantsArray=frmMain.g_oUtils.ConvertListToArray(strVariantsList,",");
                //
                //get all the packages
                //
                oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(oAdo, oAdo.m_OleDbConnection, oQueries, oRxPackageItemCollection);
                //
                //close the work file containing the links
                //
			    oAdo.CloseConnection(oAdo.m_OleDbConnection);
                //
                //create DAO instance
                //
			    dao_data_access oDao = new dao_data_access();
                //
                //create temp fvscutlist file and table
                //
                string strTempFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "ACCDB");
                oDao.CreateMDB(strTempFile);
                oAdo.OpenConnection(oAdo.getMDBConnString(strTempFile, "", ""));
                frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorIn(oAdo, oAdo.m_OleDbConnection, Tables.FVS.DefaultFVSTreeTableName);
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                //
                //create links to the source (old) fvs tree cutlist tables
                //
			    string strSourceTreeListDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()  + "\\fvs\\data\\TreeLists";
			    string strDestinationTreeListLinkTableName="";
                string strSourceTreeListLinkTableName="";
			    for (x=0;x<=strVariantsArray.Length-1;x++)
			    {
                    //make sure BiosumCalc directory exists
                    if (System.IO.Directory.Exists(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariantsArray[x].Trim() + "\\BiosumCalc") == false)
                    {
                        System.IO.Directory.CreateDirectory(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariantsArray[x].Trim() + "\\BiosumCalc");
                    }
                    //source table link
                    strSourceTreeListLinkTableName = strVariantsArray[x].Trim() + "_TREE_CUTLIST";
                    if (System.IO.File.Exists(strSourceTreeListDir + "\\" + strSourceTreeListLinkTableName.Trim() + ".MDB"))
                    {
                        oDao.CreateTableLink(oQueries.m_strTempDbFile, "fvs_tree_IN_" + strSourceTreeListLinkTableName.Trim(), strSourceTreeListDir + "\\" + strSourceTreeListLinkTableName.Trim() + ".MDB", "fvs_tree", true);
                        strVariant.Add(strVariantsArray[x].Trim());
                        
                    }
                    
			
			    }
                //
                //remove DAO instance
                //
                oDao.m_DaoWorkspace.Close();
                oDao=null;
                //
                //reopen the work file containing the links
                //
                oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile,"",""));
                //
                //create the variant + package cut list mdb file (new)
                //
                for (x = 0; x <= strVariant.Count - 1; x++)
                {
                    //get a list of packages in the old cutlist mdb
                    strSourceTreeListLinkTableName = "fvs_tree_IN_" + strVariant[x].Trim() + "_TREE_CUTLIST";
                    string strPackageList = oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT DISTINCT rxpackage FROM " + strSourceTreeListLinkTableName, ",");
                    if (strPackageList.Trim().Length > 0)
                    {
                        strPackageArray = frmMain.g_oUtils.ConvertListToArray(strPackageList, ",");
                        if (strPackageArray != null && strPackageArray.Length > 0)
                        {
                            for (y = 0; y <= strPackageArray.Length - 1; y++)
                            {
                                if (strPackageArray[y] != null && strPackageArray[y].Trim().Length > 0)
                                {
                                    //destination table creation
                                    strFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariant[x].Trim() + "\\BiosumCalc\\" + strVariant[x].Trim() + "_P" + strPackageArray[y].Trim() + "_TREE_CUTLIST.MDB";
                                    if (System.IO.File.Exists(strFile) == false)
                                    {
                                        System.IO.File.Copy(strTempFile, strFile, false);
                                    }
                                }

                            }
                        }
                    }
                }
                //
                //close the work file containing the links
                //
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                //
                //create links to variant\BiosumCalc\variant_package_tree_cutlist.mdb files
                //
                oRxTools.CreateTableLinksToFVSOutTreeListTables(oQueries, oQueries.m_strTempDbFile);
                //
                //reopen the work file containing the links
                //
                oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""));
                //
                //get processing date and time
                //
                System.DateTime oDate = System.DateTime.Now;
                string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
                string strFileDate = oDate.ToString(strDateFormat);
                strFileDate = strFileDate.Replace("/", "_"); strFileDate = strFileDate.Replace(":", "_");
                //
                //append for source trees to the destination trees
                //
                if (strVariant != null)
                {
                    for (x = 0; x <= strVariant.Count - 1; x++)
                    {
                        for (y = 0; y <= oRxPackageItemCollection.Count - 1; y++)
                        {

                            strDestinationTreeListLinkTableName = "fvs_tree_IN_" + strVariant[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST";
                            strSourceTreeListLinkTableName = "fvs_tree_IN_" + strVariant[x].Trim() + "_TREE_CUTLIST";
                            if (oAdo.TableExist(oAdo.m_OleDbConnection, strSourceTreeListLinkTableName) &&
                                oAdo.TableExist(oAdo.m_OleDbConnection, strDestinationTreeListLinkTableName))
                            {
                                //make a backup of the destination file
                                frmMain.g_sbpInfo.Text = "Version Update: FVS_CUTLIST - Make backup of " + strVariant[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST.MDB...Stand By";
                                strFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariant[x].Trim() + "\\BiosumCalc\\" + strVariant[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST.MDB";
                                System.IO.File.Copy(strFile, strFile + "_" + strFileDate);
                                //delete any current records from the destination table
                                oAdo.m_strSQL = "DELETE FROM " + strDestinationTreeListLinkTableName;
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                //insert the package
                                frmMain.g_sbpInfo.Text = "Version Update: FVS_CUTLIST - Insert Records into " + strVariant[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST.FVS_TREE table...Stand By";
                                oAdo.m_strSQL = "INSERT INTO " + strDestinationTreeListLinkTableName + " " +
                                                "SELECT * FROM " + strSourceTreeListLinkTableName + " " +
                                                "WHERE rxpackage = '" + oRxPackageItemCollection.Item(y).RxPackageId + "'";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);



                            }
                        }

                    }
                }
                //
                //close connection
                //
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                //
                //rename the old fvs cutlist files
                //
                for (x = 0; x <= strFiles.Length - 1; x++)
                {
                    if (strFiles[x].Substring(strFiles[x].Length - 4, 4) == ".MDB")
                    {
                        System.IO.File.Copy(strFiles[x].Trim(), strFiles[x].Trim() + "_VersionControlUpgrade_" + strFileDate);
                        System.IO.File.Delete(strFiles[x].Trim());
                    }
                }
                
			

            }
        }
        private void UpgradeToPrePostSeqNumMatrix()
        {
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            

           

            ado_data_access oAdo = new ado_data_access();
           
            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));
            uc_fvs_output_prepost_seqnum.InitializePrePostSeqNumTables(oAdo, ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile);

           

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

           

            oAdo = null;
        }
        
        private void UpgradeFVSOutPOTFireFiles()
        {
            int y;
           
            string strFVSOutDbFile="";
            string strVariant="";
            string strPackage="";
            
            System.Data.OleDb.OleDbConnection oFVSOutConn;
            string strConn="";
            

           
                frmMain.g_sbpInfo.Text = "Version Update: Update FVS_POTFIRE...Stand By";
                ado_data_access oAdo = new ado_data_access();
                
                
                RxTools oRxTools = new RxTools();
                Queries oQueries = new Queries();
                RxPackageItem_Collection oRxPackageItemCollection = new RxPackageItem_Collection();
                oQueries.m_oFvs.LoadDatasource = true;
                oQueries.m_oFIAPlot.LoadDatasource = true;
                oQueries.LoadDatasources(true);
                //
                //get processing date and time
                //
                System.DateTime oDate = System.DateTime.Now;
                string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
                string strFileDate = oDate.ToString(strDateFormat);
                strFileDate = strFileDate.Replace("/", "_"); strFileDate = strFileDate.Replace(":", "_");
                //
                //open the work file containing the links
                //
                oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""));
                //
                //get all the packages
                //
                //oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(oAdo, oAdo.m_OleDbConnection, oQueries, oRxPackageItemCollection);
                //
                //close the work file containing the links
                //
                //oAdo.CloseConnection(oAdo.m_OleDbConnection);
  
                string strSourceDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data";



                oAdo.m_strSQL = Queries.FVS.GetFVSVariantRxPackageSQL(oQueries.m_oFIAPlot.m_strPlotTable,oQueries.m_oFvs.m_strRxPackageTable);

                System.Data.OleDb.OleDbDataReader oDataReader;

                System.Data.OleDb.OleDbCommand oDbCommand = new System.Data.OleDb.OleDbCommand();

                oDbCommand = oAdo.m_OleDbConnection.CreateCommand();

                oDbCommand.CommandText = oAdo.m_strSQL;

                oDataReader = oDbCommand.ExecuteReader();

                while (oDataReader.Read())
                {


                    strVariant = oDataReader["fvs_variant"].ToString().Trim();
                    strPackage = oDataReader["RxPackage"].ToString().Trim();



                    strFVSOutDbFile = oRxTools.GetRxPackageFvsOutDbFileName(oDataReader);
                    if (System.IO.File.Exists(strSourceDir + "\\" + strVariant + "\\" + strFVSOutDbFile))
                    {
                        
                        strConn = oAdo.getMDBConnString(strSourceDir + "\\" + strVariant + "\\" + strFVSOutDbFile, "", "");
                        oFVSOutConn = new System.Data.OleDb.OleDbConnection();
                        oAdo.OpenConnection(strConn, ref oFVSOutConn);
                        if (oAdo.m_intError == 0)
                        {
                            frmMain.g_sbpInfo.Text = "Version Update: Update POTFIRE " + strFVSOutDbFile + "...Stand By";

                            //check if tables fvs_cases,fvs_potfire exist and the column BIOSUM_PotFireOneYearAdded_YN
                            if (oAdo.TableExist(oFVSOutConn, "FVS_CASES") &&
                                oAdo.ColumnExist(oFVSOutConn, "FVS_CASES", "BIOSUM_PotFireOneYearAdded_YN") &&
                                oAdo.TableExist(oFVSOutConn, "FVS_POTFIRE"))
                            {
                                System.IO.File.Copy(strSourceDir + "\\" + strVariant + "\\" + strFVSOutDbFile,
                                           strSourceDir + "\\" + strVariant + "\\" + strFVSOutDbFile + "_VersionControlUpgrade_" + strFileDate);

                                if (oAdo.TableExist(oFVSOutConn, "min_potfireyear"))
                                    oAdo.SqlNonQuery(oFVSOutConn, "DROP TABLE min_potfireyear");
                                //reset the FVS_POTFIRE to its original non-baseyear state
                                oAdo.m_strSQL = "UPDATE FVS_POTFIRE p " +
                                                "INNER JOIN FVS_CASES c " +
                                                "ON p.caseID = c.caseID AND " +
                                                   "p.standID = c.standID " +
                                                "SET p.year = IIF(c.BIOSUM_PotFireOneYearAdded_YN='Y',p.year - 1,p.year)," +
                                                    "c.BIOSUM_PotFireOneYearAdded_YN='N'," +
                                                    "C.BIOSUM_Append_YN='N'";
                                oAdo.SqlNonQuery(oFVSOutConn, oAdo.m_strSQL);

                                oAdo.m_strSQL = "SELECT b.min_year, a.standid " +
                                                "INTO min_potfireyear FROM FVS_POTFIRE a," +
                                                   "(SELECT MIN(year) as min_year,standid FROM FVS_POTFIRE GROUP BY standid) b " +
                                                "WHERE a.standid=b.standid AND a.year=b.min_year;";

                                oAdo.SqlNonQuery(oFVSOutConn, oAdo.m_strSQL);

                                oAdo.m_strSQL = "DELETE DISTINCTROW a.* " +
                                                "FROM FVS_POTFIRE a " +
                                                "INNER JOIN min_potfireyear b  ON a.standid = b.standid AND a.year = b.min_year";

                                oAdo.SqlNonQuery(oFVSOutConn, oAdo.m_strSQL);

                                oAdo.m_strSQL = "ALTER TABLE FVS_CASES DROP COLUMN BIOSUM_PotFireOneYearAdded_YN";
                                oAdo.SqlNonQuery(oFVSOutConn, oAdo.m_strSQL);

                                if (oAdo.TableExist(oFVSOutConn, "min_potfireyear"))
                                    oAdo.SqlNonQuery(oFVSOutConn, "DROP TABLE min_potfireyear");

                            }

                        }
                        oAdo.CloseConnection(oFVSOutConn);

                    }


                }
                oDataReader.Close();
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
           
            
               


            
        }
        private void UpgradeFVSOutPREPOSTTables()
        {
            int x;
            string strPreTable = "";
            string strPostTable = "";
            string strPREPOSTDbFile = "";
            string strSQL = "";
            frmMain.g_sbpInfo.Text = "Version Update: Update FVS PREPOST tables...Stand By";
            dao_data_access oDao = new dao_data_access();
            string strFVSOutPrePostPathAndDbFile = 
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\biosum_fvsout_prepost_rx.mdb";
            if (System.IO.File.Exists(strFVSOutPrePostPathAndDbFile))
            {

                for (x = 0; x <= Tables.FVS.g_strFVSOutTablesArray.Length - 1; x++)
                {
                    strPreTable = "PRE_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim();
                    strPostTable = "POST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim();
                    if (oDao.TableExists(strFVSOutPrePostPathAndDbFile, strPreTable) &&
                        oDao.TableExists(strFVSOutPrePostPathAndDbFile, strPostTable))
                    {
                        frmMain.g_sbpInfo.Text = "Version Update: Update PREPOST " + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + "...Stand By";
                        strPREPOSTDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\PREPOST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + ".ACCDB";
                        if (!System.IO.File.Exists(strPREPOSTDbFile))
                        {
                            oDao.CreateMDB(strPREPOSTDbFile);
                        }
                        oDao.CreateTableLink(strPREPOSTDbFile, strPreTable + "_temp", strFVSOutPrePostPathAndDbFile, strPreTable, true);
                        oDao.CreateTableLink(strPREPOSTDbFile, strPostTable + "_temp", strFVSOutPrePostPathAndDbFile, strPostTable, true);
                        oDao.OpenDb(strPREPOSTDbFile);
                        strSQL = "SELECT * INTO " + strPreTable + " FROM " + strPreTable + "_temp";
                        oDao.m_DaoDatabase.Execute(strSQL, null);
                        strSQL = "SELECT * INTO " + strPostTable + " FROM " + strPostTable + "_temp";
                        oDao.m_DaoDatabase.Execute(strSQL, null);
                        oDao.DeleteTableFromMDB(oDao.m_DaoDatabase, strPreTable + "_temp");
                        oDao.DeleteTableFromMDB(oDao.m_DaoDatabase, strPostTable + "_temp");
                        oDao.m_DaoDatabase.Close();


                    }

                }
            }

        }
        private void UpdateAuditDbFile_5_3_0()
        {
            
		    
			int x;
			//
            //MAIN PROJECT DATASOURCE
            //
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
			oDs.m_strDataSourceTableName = "datasource";
			oDs.m_strScenarioId="";
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();

            

			ado_data_access oAdo=new ado_data_access();
			
            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='PLOT AND CONDITION RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);


			oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
				"Plot And Condition Record Audit",
				ReferenceProjectDirectory.Trim() + "\\db",
				frmMain.g_oUtils.getFileName(frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableDbFile),
				frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName);


            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='PLOT, CONDITION AND TREATMENT RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

			oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
				"Plot, Condition And Treatment Record Audit",
				ReferenceProjectDirectory.Trim() + "\\db",
				frmMain.g_oUtils.getFileName(frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableDbFile),
				frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName);


			oAdo.CloseConnection(oAdo.m_OleDbConnection);

            dao_data_access oDao = new dao_data_access();
            if (!System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\db\\Audit.mdb"))
            {
                oDao.CreateMDB(ReferenceProjectDirectory.Trim() + "\\db\\Audit.mdb");
            }
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\db\\Audit.mdb", "", ""));
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "plot_cond_audit"))
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE plot_cond_audit");
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "plot_audit"))
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE plot_audit");
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "plot_cond_rx_audit"))
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE plot_cond_rx_audit");


            frmMain.g_oTables.m_oAudit.CreatePlotCondAuditTable(oAdo, oAdo.m_OleDbConnection, frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName);
            frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTable(oAdo, oAdo.m_OleDbConnection, frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            //
            //CORE SCENARIO
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb","",""));
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET table_name='" + frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName + "'," +
                                "path='" + ReferenceProjectDirectory.Trim() + "\\db' " +
                            "WHERE TRIM(UCASE(table_type))='PLOT AND CONDITION RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET table_name='" + frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName + "'," +
                                "path='" + ReferenceProjectDirectory.Trim() + "\\db' " +
                            "WHERE TRIM(UCASE(table_type))='PLOT, CONDITION AND TREATMENT RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            //
            //PROCESSOR SCENARIO
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\processor\\db\\scenario_processor_rule_definitions.mdb", "", ""));
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET table_name='" + frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName + "'," +
                                "path='" + ReferenceProjectDirectory.Trim() + "\\db' " +
                            "WHERE TRIM(UCASE(table_type))='PLOT AND CONDITION RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET table_name='" + frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName + "'," +
                                "path='" + ReferenceProjectDirectory.Trim() + "\\db' " +
                            "WHERE TRIM(UCASE(table_type))='PLOT, CONDITION AND TREATMENT RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oDao = null;
            oAdo = null;

            //
            //DELETE ANY AUDIT_VARIANT.MDB FILES
            //
            string[] strFiles = System.IO.Directory.GetFiles(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\DB", "AUDIT_*.mdb");
            for (x = 0; x <= strFiles.Length - 1; x++)
            {
                System.IO.File.Delete(strFiles[x].Trim());
            }
            strFiles.Initialize();

		}
        private void UpdateDatasources_5_3_0()
        {

            string strSQL;

            //
            //MAIN PROJECT DATASOURCE
            //
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();



            ado_data_access oAdo = new ado_data_access();

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='FVS PRE-POST SEQNUM DEFINITIONS'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('FVS PRE-POST SeqNum Definitions'," +
                        "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                        "'fvsmaster.mdb'," +
                        "'" + Tables.FVS.DefaultFVSPrePostSeqNumTable + "');";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='FVS PRE-POST SEQNUM TREATMENT PACKAGE ASSIGN'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                "('FVS PRE-POST SeqNum Treatment Package Assign'," +
                "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                "'fvsmaster.mdb'," +
                "'" + Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable + "');";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            oAdo = null;

           
        }






        public string ReferenceProjectDirectory
		{
			get {return _strProjDir;}
			set {_strProjDir=value;}
		}
		public FIA_Biosum_Manager.frmMain ReferenceMainForm
		{
			get {return _oFrmMain;}
			set {_oFrmMain = value;}
		}
		
	}
}