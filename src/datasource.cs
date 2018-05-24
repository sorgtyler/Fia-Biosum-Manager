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
	/// Get Project and Core Analysis Datasource information
	/// and create links to the datasource tables.
	/// </summary>
	public class Datasource
	{

		public int m_intError;
		public string m_strError;
		public string m_strTable;
		public const int TABLETYPE = 0;
		public const int PATH = 1;
		public const int MDBFILE = 2;
		public const int FILESTATUS = 3;
		public const int TABLE = 4;
		public const int TABLESTATUS = 5;
		public const int RECORDCOUNT = 6;
		public const int COLUMN_LIST = 7;
		public const int DATATYPE_LIST=8;
		public const int MACROVAR=9;
		public string m_strRandomPathAndFile = "";
		public int m_intNumberOfValidTables=0;  //MDB file is FOUND and table is FOUND
        /// <summary>
        /// 1st Array Item: sequence of datasource;
        /// 2nd Array Item:
        /// 0=TABLETYPE,1=DIRECTORY PATH,2=MDBFILE,3=FILESTATUS,4=TABLE,5=TABLESTATUS
        /// ,6=RECORDCOUNT,7=COLUMN_LIST (COMMA-DELIMITED),8=DATATYPE_LIST (COMMA-DELIMITED)
        /// 9=MACRO VARIABLE NAME
        /// </summary>
		public string[,] m_strDataSource;
		public string m_strScenarioId;
		public string m_strDataSourceMDBFile;
		public string m_strDataSourceTableName;
		public int m_intNumberOfTables;
		bool _bLoadFieldNamesAndDatatypes=false;
		bool _bLoadTableRecordCount=true;
		public static string[] g_strProjectDatasourceTableTypesArray = {"Plot",
															  "Condition",
															  "Tree",
															  "Owner Groups",
															  "Treatment Prescriptions",
															  "Treatment Prescriptions Assigned FVS Commands",
															  "Treatment Prescriptions Harvest Cost Columns",
															  "Treatment Prescription Categories",
															  "Treatment Prescription Subcategories",
															  "Treatment Packages",
															  "Treatment Package Assigned FVS Commands",
															  "Treatment Package Members",
															  "Treatment Package FVS Commands Order",
															  "FVS Commands",
                                                              "FVS PRE-POST SeqNum Definitions",
                                                              "FVS PRE-POST SeqNum Treatment Package Assign", 
															  "Tree Species",
															  Datasource.TableTypes.FvsTreeSpecies,
                                                              "FVS Western Tree Species Translator",
                                                              "FVS Eastern Tree Species Translator",
															  "Travel Times",
															  "Processing Sites",
															  "FVS Tree List For Processor",
															  "FIADB FVS Variant",
                                                              "FIA Tree Macro Plot Breakpoint Diameter",
															  Datasource.TableTypes.HarvestMethods,
															  "Plot And Condition Record Audit",
															  "Plot, Condition And Treatment Record Audit",
															  "Tree Regional Biomass",
															  "Population Evaluation",
															  "Population Estimation Unit",
															  "Population Stratum",
															  "Population Plot Stratum Assignment",
                                                              "BIOSUM Pop Stratum Adjustment Factors",
															  "Site Tree",
                                                              Datasource.TableTypes.FiaTreeSpeciesReference,
                                                              "DWM COARSE WOODY DEBRIS",
                                                              "DWM DUFF LITTER FUEL",
                                                              "DWM FINE WOODY DEBRIS",
                                                              "DWM TRANSECT SEGMENT",
                                                              "REF FOREST TYPE",
                                                              "REF FOREST TYPE GROUP",
		};

        public static string[] g_strCoreDatasourceTableTypesArray = {"Plot",
															  "Condition",
															  "Tree",
															  "Owner Groups",
															  "Treatment Prescriptions",
															  "Treatment Prescriptions Assigned FVS Commands",
															  "Treatment Prescriptions Harvest Cost Columns",
															  "Treatment Prescription Categories",
															  "Treatment Prescription Subcategories",
															  "Treatment Packages",
															  "Treatment Package Assigned FVS Commands",
															  "Treatment Package Members",
															  "Treatment Package FVS Commands Order",
															  "FVS Commands",
															  "Tree Species",
															  "FVS Tree Species",
															  "Travel Times",
															  "Processing Sites",
															  "FVS Tree List For Processor",
															  "FIADB FVS Variant",
                                                              "FIA Tree Macro Plot Breakpoint Diameter",
															  Datasource.TableTypes.HarvestMethods,
															  "Plot And Condition Record Audit",
															  "Plot, Condition And Treatment Record Audit",
															  "Tree Regional Biomass",
															  "Population Evaluation",
															  "Population Estimation Unit",
															  "Population Stratum",
															  "Population Plot Stratum Assignment",
                                                              "BIOSUM Pop Stratum Adjustment Factors",
															  "Site Tree"};
        public static string[] g_strProcessorDatasourceTableTypesArray = {"Plot",
															  "Condition",
															  "Tree",
															  "Owner Groups",
															  "Treatment Prescriptions",
															  "Treatment Prescriptions Assigned FVS Commands",
															  "Treatment Prescriptions Harvest Cost Columns",
															  "Treatment Prescription Categories",
															  "Treatment Prescription Subcategories",
															  "Treatment Packages",
															  "Treatment Package Assigned FVS Commands",
															  "Treatment Package Members",
															  "Treatment Package FVS Commands Order",
															  "FVS Commands",
															  "Tree Species",
															  "FVS Tree Species",
															  "Travel Times",
															  "Processing Sites",
															  "FVS Tree List For Processor",
															  "FIADB FVS Variant",
                                                              "FIA Tree Macro Plot Breakpoint Diameter",
															  Datasource.TableTypes.HarvestMethods,
															  "Plot And Condition Record Audit",
															  "Plot, Condition And Treatment Record Audit",
															  "Tree Regional Biomass",
															  "Population Evaluation",
															  "Population Estimation Unit",
															  "Population Stratum",
															  "Population Plot Stratum Assignment",
                                                              "BIOSUM Pop Stratum Adjustment Factors",
															  "Site Tree"};
		public static FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem g_oCurrentSQLMacroSubstitutionVariableItem=
			            new SQLMacroSubstitutionVariableItem();

		


		public Datasource()
		{

		}
		///<summary>
		///Access project datasource information and functionality.
		///Project Datasource information is loaded
		///into strDataSource[,] array 
		///from the constructor.
	    ///</summary>
	    ///<param name="strProjDir">Project Root Directory</param>
		public Datasource(string strProjDir)
		{
			this.m_strDataSourceMDBFile = strProjDir + "\\db\\project.mdb";
			this.m_strDataSourceTableName = "datasource";
			this.m_strScenarioId="";
			this.populate_datasource_array();
		}
  	    ///<summary>
	    ///get scenario datasource information.
	    ///Scenario Datasource information is loaded
	    ///into strDataSource[,] array when this class is instatiated.
	    ///</summary>
	    ///<param name="strProjDir">Project Root Directory</param>
		/// <param name="strScenarioId">Value is used to query core analysis scenario datasource infornation.</param>
		public Datasource(string strProjDir, string strScenarioId)
		{
			this.m_strDataSourceMDBFile = strProjDir + "\\core\\db\\scenario_core_rule_definitions.mdb";
			this.m_strDataSourceTableName = "scenario_datasource";
			this.m_strScenarioId = strScenarioId;
			this.populate_datasource_array();
		}
		public Datasource(string p_strProjDir, string p_strScenarioId,string p_strScenarioType)
		{
			this.m_strDataSourceMDBFile = p_strProjDir + "\\" + p_strScenarioType + "\\db\\scenario_" + p_strScenarioType + "_rule_definitions.mdb";
			this.m_strDataSourceTableName = "scenario_datasource";
			this.m_strScenarioId = p_strScenarioId;
			this.populate_datasource_array();
		}
		~Datasource()
		{
		}
		
		///<summary>
		///Load a 2 dimensional array with 
		///this datasource information:
		///Table Type, MDB paths and files, table names, file/table
		///found, table record count.
		///OPTIONAL: table columns and datatypes will also load into the array
		///          if the variable LoadColumnNamesAndDataTypes is set to true
 	    ///</summary>
		public void populate_datasource_array()
		{
            macrosubst oMacroSub = new macrosubst();
            oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;
            
            int intRecCnt=0;    		
			string strPathAndFile="";
			string strSQL="";
			string strConn="";
			this.m_intNumberOfTables=0;

            FIA_Biosum_Manager.ado_data_access  p_ado = new ado_data_access();

			System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(this.m_strDataSourceMDBFile,"admin","");
			oConn.ConnectionString = strConn;
			this.m_intError = 0;
			this.m_strError = "";
			try
			{
				oConn.Open();
			}
			catch (System.Data.OleDb.OleDbException oleException)
			{
				this.m_strError = "Failed to make an oleDb connection with " + strConn;
				MessageBox.Show (this.m_strError + " OledbError=" + oleException.Message);
				this.m_intError = -1;
				return;
			}
			intRecCnt = Convert.ToInt32(p_ado.getRecordCount(oConn,"select count(*) from " + this.m_strDataSourceTableName,this.m_strDataSourceTableName));
			this.m_strDataSource = new String[intRecCnt,10];
			System.Data.OleDb.OleDbCommand oCommand = oConn.CreateCommand();
      if (this.m_strScenarioId.Trim().Length > 0)
      {
			oCommand.CommandText = "select table_type,path,file,table_name from scenario_datasource" + 
				                   " where scenario_id = '" + this.m_strScenarioId.Trim() + "';";
		  }
		  else
		  {
			  oCommand.CommandText = "select table_type,path,file,table_name from " + this.m_strDataSourceTableName + ";" ;
			}
				                           
			try
			{
				System.Data.OleDb.OleDbDataReader oDataReader = oCommand.ExecuteReader();
				int x = 0;
				
				while (oDataReader.Read())
				{
					this.m_intNumberOfTables++;
					// Add a ListItem object to the ListView.
					this.m_strDataSource[x,TABLETYPE] = oDataReader["table_type"].ToString().Trim();
					this.m_strDataSource[x,PATH] = oDataReader["path"].ToString().Trim();
					this.m_strDataSource[x,MDBFILE] = oDataReader["file"].ToString().Trim();
                    strPathAndFile = oMacroSub.GeneralTranslateVariableSubstitution(oDataReader["path"].ToString().Trim()) + 
                        "\\" + oDataReader["file"].ToString().Trim();
					if (System.IO.File.Exists(strPathAndFile) == true) 
					{
						this.m_strDataSource[x,FILESTATUS] = "F";
						this.m_strDataSource[x,TABLE] = oDataReader["table_name"].ToString().Trim();
						

						//see if the table exists in the mdb database container
						dao_data_access p_dao = new dao_data_access();
						if (p_dao.TableExists(strPathAndFile,oDataReader["table_name"].ToString().Trim()) == true)
						{
							this.m_strDataSource[x,TABLESTATUS] = "F";
							this.m_strDataSource[x,RECORDCOUNT]="0";
							this.m_strDataSource[x,COLUMN_LIST]="";
							this.m_strDataSource[x,DATATYPE_LIST]="";
							
							if (this.LoadTableRecordCount || this.LoadTableColumnNamesAndDataTypes)
							{
								strConn=p_ado.getMDBConnString(strPathAndFile,"admin","");
								p_ado.OpenConnection(strConn);
								if (p_ado.m_intError==0)
								{
									strSQL = "select count(*) from " + oDataReader["table_name"].ToString();
									if (this.LoadTableRecordCount) this.m_strDataSource[x,RECORDCOUNT] = Convert.ToString(p_ado.getRecordCount(strConn,strSQL,oDataReader["table_name"].ToString()));
									if (this.LoadTableColumnNamesAndDataTypes) p_ado.getFieldNamesAndDataTypes(strConn,"select * from " + oDataReader["table_name"].ToString(), ref this.m_strDataSource[x,COLUMN_LIST],ref this.m_strDataSource[x,DATATYPE_LIST]);
									p_ado.CloseConnection(p_ado.m_OleDbConnection);
								}
							}
						}
						else 
						{
							this.m_strDataSource[x,TABLESTATUS] = "NF";
							this.m_strDataSource[x,RECORDCOUNT] = "0";
						}
						p_dao.m_DaoWorkspace.Close();
						p_dao = null;
					}
					else 
					{
						this.m_strDataSource[x,FILESTATUS] = "NF";
						this.m_strDataSource[x,TABLE] = oDataReader["table_name"].ToString().Trim();
						this.m_strDataSource[x,TABLESTATUS] = "NF";
						this.m_strDataSource[x,RECORDCOUNT] = "0";
					}
					UpdateTableMacroVariable(this.m_strDataSource[x,TABLETYPE],this.m_strDataSource[x,TABLE]);
					
					x++;
				}
				oDataReader.Close();
			}
			catch
			{
				this.m_intError = -1;
				this.m_strError = "The Query Command " + oCommand.CommandText.ToString() + " Failed";
				MessageBox.Show(this.m_strError);
				if (oConn != null)
				{
					if (oConn.State==System.Data.ConnectionState.Open)
					{
						oConn.Close();
						while (oConn.State != System.Data.ConnectionState.Closed)
							System.Threading.Thread.Sleep(1000);

						oConn = null;
					}
				}
				p_ado= null;
				return;
			}
			finally
			{
				if (oConn != null)
				{
					if (oConn.State==System.Data.ConnectionState.Open)
					{
						oConn.Close();
						while (oConn.State != System.Data.ConnectionState.Closed)
							System.Threading.Thread.Sleep(1000);

						oConn = null;
					}
				}
			}
			
			if (oConn != null)
			{
				if (oConn.State==System.Data.ConnectionState.Open)
				{
					oConn.Close();
					while (oConn.State != System.Data.ConnectionState.Closed)
						System.Threading.Thread.Sleep(1000);

					oConn = null;
				}
			}
			p_ado = null;
		}


  	///<summary>
		///Validate the existance of the table in the datasource
	  ///</summary>
		/// <param name="strTableName">Table name to validate.</param>
		public bool DataSourceTableExist(string strTableName)
		{
			int x;
			for (x=0; x <= this.m_intNumberOfTables-1; x++)
			{
				if (this.m_strDataSource[x,TABLE].Trim().ToUpper()==strTableName.Trim().ToUpper())
				{
					if (this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper()=="NF")
					{
						return false;
				}
					if (this.m_strDataSource[x,FILESTATUS].Trim().ToUpper()=="NF")
					{
						return false;
					}
					return true;
				}   
			}
			return false;
		}
		///<summary>
		///create a mdb file in the users temporary dir
		///and create a link to each of the  
		///data source tables.  Return the name of the 
		///temporary mdb file to the calling function
		///</summary>
		public string CreateMDBAndTableDataSourceLinks()
		{
			macrosubst oMacroSub = new macrosubst();
            oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;
            string strTempMDB="";
			int x;
			
			FIA_Biosum_Manager.env p_env = new env();
            this.m_intNumberOfValidTables=0;
      
			//used to get the temporary random file name
			utils p_utils = new utils();
			
			//used to create a link to the table
			dao_data_access p_dao = new dao_data_access();
			
			for (x=0; x <= this.m_intNumberOfTables - 1; x++) //TODO: check if this is updated to include dwm
			{
                string strFileStatus = this.m_strDataSource[x, FILESTATUS];
                if (strFileStatus != null)
                    strFileStatus = strFileStatus.Trim().ToUpper();
                string strTableStatus = this.m_strDataSource[x, TABLESTATUS];
                if (strTableStatus != null)
                    strTableStatus = strTableStatus.Trim().ToUpper();
                if (strTableStatus=="F" &&
					strFileStatus=="F")
					{
						if (strTempMDB.Trim().Length == 0)
						{
							//get temporary mdb file
							strTempMDB = 
								p_utils.getRandomFile(p_env.strTempDir,"accdb");

							//create a temporary mdb that will contain all 
							//the links to the scenario datasource tables
							p_dao.CreateMDB(strTempMDB);

						}
						p_dao.CreateTableLink(strTempMDB,
							this.m_strDataSource[x,TABLE].Trim(),
							oMacroSub.GeneralTranslateVariableSubstitution(this.m_strDataSource[x,PATH].Trim()) + "\\" +
							     this.m_strDataSource[x,MDBFILE].Trim(),
							this.m_strDataSource[x,TABLE].Trim());
						this.m_intNumberOfValidTables++;
					}
			}
			p_utils = null;
			p_dao = null;
			p_env = null;
            if (strTempMDB.Trim().Length == 0)
				MessageBox.Show("!!None of the data source tables are found!!");
			return strTempMDB;
		}
        public void CreateScenarioRuleDefinitionTableLinks(string p_strDestDbFile,string p_strProjectPath,string p_strType)
        {
            //used to create a link to the table
            dao_data_access oDao = new dao_data_access();
            oDao.OpenDb(p_strDestDbFile);
            CreateScenarioRuleDefinitionTableLinks(oDao,oDao.m_DaoDatabase,p_strProjectPath,p_strType);
            oDao.m_DaoDatabase.Close();
            oDao.m_DaoWorkspace.Close();
            oDao.m_DaoTableDef = null;
            oDao.m_DaoDatabase = null;
           
        }
        /// <summary>
        /// create links to each of the scenario tables
        /// </summary>
        /// <param name="p_oDao"></param>
        /// <param name="p_DaoDatabase"></param>
        /// <param name="p_strType">P=processor scenario, C=core analysis scenario</param>
        public void CreateScenarioRuleDefinitionTableLinks(dao_data_access p_oDao,
                    Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, 
                    string p_strProjectPath,
                    string p_strType)
        {
            //used to create a link to the table
            string strPath = p_strProjectPath;
            if (p_strType == "P")
            {
                strPath = strPath + "\\processor\\";



                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.Scenario.DefaultScenarioTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsDbFile,
                    Tables.Scenario.DefaultScenarioTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultRxHarvestMethodTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultRxHarvestMethodDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultRxHarvestMethodTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName,
                    strPath + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName);
            }
            else
            {
                strPath = strPath + "\\";

                p_oDao.CreateTableLink(p_DaoDatabase,
                    Tables.CoreScenarioRuleDefinitions.DefaultScenarioTableName,
                    strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioTableDbFile,
                    Tables.CoreScenarioRuleDefinitions.DefaultScenarioTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterTableName);
                
                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioCostsTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioCostsTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioCostsTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName);

                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioPSitesTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioPSitesTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioPSitesTableName);

              
                p_oDao.CreateTableLink(p_DaoDatabase,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioRxIntensityTableName,
                   strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioRxIntensityTableDbFile,
                   Tables.CoreScenarioRuleDefinitions.DefaultScenarioRxIntensityTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName,
                    strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableDbFile,
                    Tables.CoreScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName);

                p_oDao.CreateTableLink(
                    p_DaoDatabase,
                    Tables.CoreScenarioRuleDefinitions.DefaultScenarioDatasourceTableName,
                    strPath + Tables.CoreScenarioRuleDefinitions.DefaultScenarioDatasourceTableDbFile,
                    Tables.CoreScenarioRuleDefinitions.DefaultScenarioDatasourceTableName);

            }
        }

  	///<summary>
		///Return the location of the specified table within the m_strDataSource array.
		///-1 is returned if the strTableType is not found or the MDB file is not
		///found or the table is not found
	  ///</summary>
		/// <param name="strTableType">The unique id for the datasource table</param>
		public int getValidTableNameRow(string strTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
				if (strTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper()
					&&
					this.m_strDataSource[x,FILESTATUS].Trim().ToUpper() =="F" 
					&&
					this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper() == "F")
				{
					  break;
				}
			}
			if (x > this.m_intNumberOfTables-1)
			{
				x=-1;
			}
			return x;
			
		}
		///<summary>
		///Return the location of the specified table within the m_strDataSource array.
		///</summary>
		/// <param name="strTableType">The unique id for the datasource table</param>
		public int getTableNameRow(string strTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
				if (strTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper())
				{
					break;
				}
			}
			if (x > this.m_intNumberOfTables-1)
			{
				x=-1;
			}
			return x;
			
		}

    ///<summary>
    ///number of tables that are found
    ///</summary>
		public void getNumberOfValidTables()
		{
			int x;
			this.m_intNumberOfValidTables=0;
			for (x=0; x <= this.m_intNumberOfTables - 1; x++)
			{
				if (this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper()=="F" &&
					this.m_strDataSource[x,FILESTATUS].Trim().ToUpper()=="F")
				{
					this.m_intNumberOfValidTables++;
				}
			}

		}

    ///<summary>
    ///Add the datasource tables to a list box
    ///</summary>
    ///<param name="listbox1">datasource table names will be added to the specified listbox object</param>
		public void LoadDataSourceTablesIntoListBox(System.Windows.Forms.ListBox listbox1)
		{
			int x;
			for (x=0; x <= this.m_intNumberOfTables - 1; x++)
			{
				if (this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper()=="F" &&
					this.m_strDataSource[x,FILESTATUS].Trim().ToUpper()=="F")
				{
					listbox1.Items.Add(this.m_strDataSource[x,TABLE].Trim());
				}
			}
		}
		
  		///<summary>
		///Return the location of the specified table within the strDataSource array.
		///-1 is returned if the strTableType is not found
		 ///</summary>
		/// <param name="strTableType">The unique id for the datasource table</param>
		public int getDataSourceTableNameRow(string pcTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
				if (pcTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper())
				{
					return x;
				}
			}
			return -1;
		}
		
		///<summary>
		//return a valid table name associated with the table type
		///</summary>
    /// <param name="strTableType">The unique id for the datasource table</param>		 
		public string getValidDataSourceTableName(string strTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
                string strFileStatus = this.m_strDataSource[x, FILESTATUS];
                if (strFileStatus != null)
                    strFileStatus = strFileStatus.Trim().ToUpper();
                string strTableStatus = this.m_strDataSource[x, TABLESTATUS];
                if (strTableStatus != null)
                    strTableStatus = strTableStatus.Trim().ToUpper();
                if (strTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper()
					&&
					strFileStatus =="F" 
					&&
					strTableStatus == "F")
				{
					return this.m_strDataSource[x,TABLE].Trim();
				}
			}
			return "";
		}
		/// <summary>
		/// get the full path to mdb file
		/// </summary>
		/// <param name="p_strTableType"></param>
		/// <returns></returns>
		public string getFullPathAndFile(string p_strTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
				if (p_strTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper())
				{
					return this.m_strDataSource[x,PATH].ToString().Trim() + "\\" + this.m_strDataSource[x,MDBFILE].ToString().Trim();
				}
			}
			return "";

		}
		///<summary>
		//validate the datasources and return -1 if there are problems
		///</summary>
		public int val_datasources()
		{
		  int x=0;
			for (x=0; x <= this.m_intNumberOfTables - 1; x++)
			{
				if (this.m_strDataSource[x,FILESTATUS].Trim().ToUpper()=="NF")
				{
					MessageBox.Show("Datasource Failure: data source file " + this.m_strDataSource[x,PATH].Trim() + "\\" + 
		            this.m_strDataSource[x,MDBFILE].Trim() + " is not found");
					return -1;
				}
				if (this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper()=="NF")
				{
					MessageBox.Show("Datasource Failure: data source table " + this.m_strDataSource[x,TABLE].Trim() + 
						 " is not found");
					return -1;
				}
				if (this.m_strDataSource[x,RECORDCOUNT].Trim().ToUpper()=="0")
				{
					MessageBox.Show("Datasource Failure: data source table " + this.m_strDataSource[x,TABLE].Trim() + 
						" has 0 records");
					return -1;
				}
			}
			return 0;
		}
		/// <summary>
		/// create primary indexes and set autonumber increment 
		/// </summary>
		/// <param name="p_strMDBPathAndFile">full directory path and file name that contains the table</param>
		/// <param name="p_strTableType">datasource table type</param>
		/// <param name="p_strTable">table name of the table type</param>
		/// <param name="p_dao">dao_data_access object</param>
		public void SetPrimaryIndexesAndAutoNumbers(string p_strMDBPathAndFile, string p_strTableType, string p_strTable, dao_data_access p_dao)
		{
			switch (p_strTableType.Trim().ToUpper())
			{
				case "PLOT":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"biosum_plot_id");
					break;
				case "CONDITION":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"biosum_cond_id");
					break;
				case "HARVEST COSTS":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"biosum_cond_id,rx");
					break;
				case "TREE SPECIES":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"id");
					p_dao.CreateAutoNumber(p_strMDBPathAndFile,p_strTable,"id");
					break;
				case "FVS TREE SPECIES":
					break;
				case "TREATMENT PRESCRIPTIONS":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"rx");
					break;
				case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"id");
					p_dao.CreateAutoNumber(p_strMDBPathAndFile,p_strTable,"id");
					break;
				case "FVS TREE LIST FOR PROCESSOR":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"id");
					p_dao.CreateAutoNumber(p_strMDBPathAndFile,p_strTable,"id");
					break;
				case "TRAVEL TIMES":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"traveltime_id");
					p_dao.CreateAutoNumber(p_strMDBPathAndFile,p_strTable,"traveltime_id");
					break;
				case "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"id");
					p_dao.CreateAutoNumber(p_strMDBPathAndFile,p_strTable,"id");
					break;
				case "PROCESSING SITES":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"psite_id");
					break;
				case "PLOT AND CONDITION RECORD AUDIT":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"biosum_cond_id");
					break;
				case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile,p_strTable,"biosum_cond_id,rx");
					break;
				case "TREE REGIONAL BIOMASS":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile, p_strTable,"tre_cn");
					break;
				case "POPULATION PLOT STRATUM ASSIGNMENT":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile, p_strTable,"CN");
					break;
				case "POPULATION STRATUM":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile, p_strTable,"CN");
					break;
				case "POPULATION EVALUATION":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile, p_strTable,"CN");
					break;
				case "POPULATION ESTIMATION UNIT":
					p_dao.CreatePrimaryKeyIndex(p_strMDBPathAndFile, p_strTable,"CN");
					break;



			}

		}

		public void SetPrimaryIndexesAndAutoNumbers(string p_strTableType, string p_strTable, dao_data_access p_dao)
		{
			switch (p_strTableType.Trim().ToUpper())
			{
				case "PLOT":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"biosum_plot_id");
					break;
				case "CONDITION":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"biosum_cond_id");
					break;
				case "HARVEST COSTS":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"biosum_cond_id,rx");
					break;
				case "FVS TREE SPECIES":
					break;
				case "TREATMENT PRESCRIPTIONS":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"rx");
					break;
				case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"id");
					p_dao.CreateAutoNumber(p_dao.m_DaoDatabase,p_strTable,"id");
					break;
				case "FVS TREE LIST FOR PROCESSOR":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"id");
					p_dao.CreateAutoNumber(p_dao.m_DaoDatabase,p_strTable,"id");
					break;
				case "TRAVEL TIMES":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"traveltime_id");
					p_dao.CreateAutoNumber(p_dao.m_DaoDatabase,p_strTable,"traveltime_id");
					break;
				case "PROCESSING SITES":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"psite_id");
					break;
				
				case "TREE SPECIES":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"id");
					p_dao.CreateAutoNumber(p_dao.m_DaoDatabase,p_strTable,"id");
					break;
				case "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"id");
					p_dao.CreateAutoNumber(p_dao.m_DaoDatabase,p_strTable,"id");
					break;
				case "PLOT AND CONDITION RECORD AUDIT":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"biosum_cond_id");
					break;
				case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"biosum_cond_id,rx");
					break;
				case "TREE REGIONAL BIOMASS":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase,p_strTable,"tre_cn");
					break;
				case "POPULATION PLOT STRATUM ASSIGNMENT":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase, p_strTable,"CN");
					break;
				case "POPULATION EVALUATION":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase, p_strTable,"CN");
					break;
				case "POPULATION ESTIMATION UNIT":
					p_dao.CreatePrimaryKeyIndex(p_dao.m_DaoDatabase, p_strTable,"CN");
					break;


			}

		}
		public void SetPrimaryIndexesAndAutoNumbers(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strTableType, string p_strTableName)
		{
			switch (p_strTableType.Trim().ToUpper())
			{
				case "PLOT":
					frmMain.g_oTables.m_oFIAPlot.CreatePlotTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "CONDITION":
					frmMain.g_oTables.m_oFIAPlot.CreateConditionTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "HARVEST COSTS":
					frmMain.g_oTables.m_oProcessor.CreateHarvestCostsTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "FVS TREE SPECIES":
					frmMain.g_oTables.m_oReference.CreateFVSTreeSpeciesTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "TREATMENT PRESCRIPTIONS":
					frmMain.g_oTables.m_oFvs.CreateRxTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
					frmMain.g_oTables.m_oProcessor.CreateTreeVolValSpeciesDiamGroupsTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "FVS TREE LIST FOR PROCESSOR":
					frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorInTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "TRAVEL TIMES":
					frmMain.g_oTables.m_oTravelTime.CreateTravelTimeTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "PROCESSING SITES":
					frmMain.g_oTables.m_oTravelTime.CreateProcessingSiteTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "TREE SPECIES":
					frmMain.g_oTables.m_oReference.CreateTreeSpeciesTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "PLOT AND CONDITION RECORD AUDIT":
					frmMain.g_oTables.m_oAudit.CreatePlotCondAuditTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
					frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "TREE REGIONAL BIOMASS":
					frmMain.g_oTables.m_oFIAPlot.CreateTreeRegionalBiomassTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "POPULATION PLOT STRATUM ASSIGNMENT":
					frmMain.g_oTables.m_oFIAPlot.CreatePopPlotStratumAssgnTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "POPULATION EVALUATION":
					frmMain.g_oTables.m_oFIAPlot.CreatePopEvalTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
				case "POPULATION ESTIMATION UNIT":
                    frmMain.g_oTables.m_oFIAPlot.CreatePopEstnUnitTableIndexes(p_oAdo,p_oConn,p_strTableName);
					break;
                case "FVS COMMANDS":
                    frmMain.g_oTables.m_oReference.CreateFVSCommandsTableIndexes(p_oAdo, p_oConn, p_strTableName);
                    break;
                case "HARVEST METHODS":
                    frmMain.g_oTables.m_oReference.CreateHarvestMethodsTableIndexes(p_oAdo, p_oConn, p_strTableName);
                    break;
                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                    frmMain.g_oTables.m_oReference.CreateRxSubCategoryTableIndexes(p_oAdo, p_oConn, p_strTableName);
                    break;
                case "TREATMENT PRESCRIPTION CATEGORIES":
                    frmMain.g_oTables.m_oReference.CreateRxCategoryTableIndexes(p_oAdo, p_oConn, p_strTableName);
                    break;
			}

		}
	    public static void UpdateTableMacroVariable(string p_strTableType,string p_strTableName)
		{
			
			int x;
			
			for (x=0;x<=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count-1;x++)
			{
				if (frmMain.g_oSQLMacroSubstitutionVariable_Collection.Item(x).Description.Trim().ToUpper()
					== p_strTableType.Trim().ToUpper() + " TABLE")
				{
					frmMain.g_oSQLMacroSubstitutionVariable_Collection.Item(x).SQLVariableSubstitutionString=p_strTableName.Trim();
					Datasource.g_oCurrentSQLMacroSubstitutionVariableItem.CopyProperties(frmMain.g_oSQLMacroSubstitutionVariable_Collection.Item(x),ref Datasource.g_oCurrentSQLMacroSubstitutionVariableItem);
					return;
				}
			}
			
			FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem oItem = new SQLMacroSubstitutionVariableItem();
			oItem.Description= p_strTableType.Trim() + " Table";
			oItem.Index=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count;
			x=oItem.Index;
			

			switch (p_strTableType.Trim().ToUpper())
			{
				case "PLOT":
					oItem.VariableName="PlotTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName;
					break;
				case "CONDITION":
					oItem.VariableName="CondTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName;
					break;
				case "TREE":
					oItem.VariableName="TreeTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName;
					break;
				case "OWNER GROUPS":
					oItem.VariableName="OwnerGroupsTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Reference.DefaultOwnerGroupsTableName;
					break;
				case "HARVEST COSTS":
					oItem.VariableName="HarvestCostsTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Processor.DefaultHarvestCostsTableName;
					break;
				case "FVS TREE SPECIES":
					oItem.VariableName="FvsTreeSpeciesTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Reference.DefaultFVSTreeSpeciesTableName;
					break;
                case "FIADB FVS VARIANT":
					oItem.VariableName="FiadbFvsVariantTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Reference.DefaultFiadbFVSVariantTableName;
					break;
                case "HARVEST METHODS":
					oItem.VariableName="HarvestMethodsTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Reference.DefaultFiadbFVSVariantTableName;
					break;
				case "TREATMENT PRESCRIPTIONS":
					if (p_strTableName.Trim().Length == 0) p_strTableName= Tables.FVS.DefaultRxTableName;
					oItem.VariableName="RxTable";
					break;
                case "TREATMENT PACKAGES":
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.FVS.DefaultRxPackageTableName;
					oItem.VariableName="RxPackageTable";
					break;
				case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Processor.DefaultTreeVolValSpeciesDiamGroupsTableName;
					oItem.VariableName="TreeVolValBySpcGrpDiaGrpTable";
					break;
				case "FVS TREE LIST FOR PROCESSOR":
					oItem.VariableName="FvsTreeListForProcessorTable";
                    if (p_strTableName.Trim().Length == 0) p_strTableName = Tables.FVS.DefaultFVSTreeTableName;
					break;
				case "TRAVEL TIMES":
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oTravelTime.DefaultTravelTimeTableName;
					oItem.VariableName="TravelTimesTable";
					break;
				case "PROCESSING SITES":
					oItem.VariableName="PSitesTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oTravelTime.DefaultProcessingSiteTableName;
					break;
				case "TREE SPECIES":
					oItem.VariableName="TreeSpeciesTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Reference.DefaultTreeSpeciesTableName;
					break;
				case "PLOT AND CONDITION RECORD AUDIT":
					oItem.VariableName="PlotCondAuditTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oAudit.DefaultPlotCondAuditTableName;
					break;
				case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
					oItem.VariableName="PlotCondRxAuditTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oAudit.DefaultPlotCondRxAuditTableName;
					break;
				case "TREE REGIONAL BIOMASS":
					oItem.VariableName="TreeRegionalBiomassTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultTreeRegionalBiomassTableName;
					break;
                case "BIOSUM POP STRATUM ADJUSTMENT FACTORS":
                    oItem.VariableName="BiosumPopStratumAdjustmentFactorTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName;
					break;
                case "FIA TREE MACRO PLOT BREAKPOINT DIAMETER":
                    oItem.VariableName = "TreeMacroPlotBreakPointDiameterTable";
                    if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Reference.DefaultTreeMacroPlotBreakPointDiaTableName;
					break;
				case "POPULATION PLOT STRATUM ASSIGNMENT":
					oItem.VariableName="PpsaTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName;
					break;
				case "POPULATION EVALUATION":
					oItem.VariableName="PopEvalTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName;
					break;
				case "POPULATION ESTIMATION UNIT":
					oItem.VariableName="PopEstnUnitTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName;
					break;
				case "POPULATION STRATUM":
					oItem.VariableName="PopStratumTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName;
					break;
				case "SITE TREE":
					oItem.VariableName="SiteTreeTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName;
					break;

                //TODO: Determine if these cases are necessary and remove if not
				case "DWM COARSE WOODY DEBRIS":
					oItem.VariableName="DwmCoarseWoodyDebrisTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName;
					break;
				case "DWM DUFF LITTER FUEL":
					oItem.VariableName="DwmDuffLitterFuelTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName;
					break;
				case "DWM FINE WOODY DEBRIS":
					oItem.VariableName="DwmFineWoodyDebrisTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName;
					break;
				case "DWM TRANSECT SEGMENT":
					oItem.VariableName="DwmTransectSegmentTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName;
					break;
				case "REF FOREST TYPE":
					oItem.VariableName="RefForestTypeTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultRefForestTypeName;
					break;
				case "REF FOREST TYPE GROUP":
					oItem.VariableName="RefForestTypeGroupTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultRefForestTypeGroupName;
					break;
			}
			if (oItem.VariableName.Trim().Length > 0)
			{
				oItem.SQLVariableSubstitutionString=p_strTableName.Trim();
				Datasource.g_oCurrentSQLMacroSubstitutionVariableItem.CopyProperties(oItem,ref Datasource.g_oCurrentSQLMacroSubstitutionVariableItem);
				frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(oItem);
			}
			else
			{
				Datasource.g_oCurrentSQLMacroSubstitutionVariableItem = new SQLMacroSubstitutionVariableItem();
			}

		}
		public static void InititializeMacroVariables()
		{
			for (int x=0;x<=Datasource.g_strProjectDatasourceTableTypesArray.Length - 1;x++)
			{
				Datasource.UpdateTableMacroVariable(Datasource.g_strProjectDatasourceTableTypesArray[x],"");
			}
			//
			//plot/condition connection
			//
			FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem oItem = new SQLMacroSubstitutionVariableItem();
			oItem.Description= "Key relationship between Plot and Condition Table";
			oItem.Index=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count;
			oItem.VariableName="PlotCondConnection";
			oItem.SQLVariableSubstitutionString="@@CondTable@@.biosum_plot_id=@@PlotTable@@.biosum_plot_id";
			frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(oItem);
			//
			//condition/tree connection
			//
			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.Description= "Key relationship between Condition and Tree Table";
			oItem.Index=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count;
			oItem.VariableName="CondTreeConnection";
			oItem.SQLVariableSubstitutionString="@@CondTable@@.biosum_cond_id=@@TreeTable@@.biosum_cond_id";
			frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(oItem);
			//
			//condition/tree connection
			//
			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.Description= "Key relationship between Condition and Tree Table";
			oItem.Index=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count;
			oItem.VariableName="CondTreeConnection";
			oItem.SQLVariableSubstitutionString="@@CondTable@@.biosum_cond_id=@@TreeTable@@.biosum_cond_id";
			frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

            



		}

		public void InsertDatasourceRecord(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,
			                               string p_strTableType,string p_strDirectory,string p_strDbFile,
			                               string p_strTableName)
		{
			p_oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " + 
				"('" + p_strTableType.Trim() + "'," + 
				"'" + p_strDirectory.Trim() + "'," + 
				"'" + p_strDbFile.Trim() + "'," + 
				"'" + p_strTableName.Trim() + "');";
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

		}
        /// <summary>
        /// get distinct FVS variants
        /// </summary>
        /// <param name="p_strVariantArray"></param>
        public string[] getVariants()
        {
        
            //
            //get all the variants
            //
            int x;
            string[] strVariantArray = null;
            ado_data_access oAdo = new ado_data_access();
           
            x = getDataSourceTableNameRow("PLOT");
            oAdo.OpenConnection(
                oAdo.getMDBConnString(
                m_strDataSource[x,Datasource.PATH].Trim() + "\\" +  
                m_strDataSource[x, Datasource.MDBFILE].Trim(), "", ""));
            strVariantArray = frmMain.g_oUtils.ConvertListToArray(oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT DISTINCT FVS_VARIANT FROM " + m_strDataSource[x, Datasource.TABLE].Trim(), ""), ",");
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
            return strVariantArray;
        
        }
		public bool LoadTableColumnNamesAndDataTypes
		{
			get {return this._bLoadFieldNamesAndDatatypes;}
			set {this._bLoadFieldNamesAndDatatypes=value;}
		}
		public bool LoadTableRecordCount
		{
			get {return this._bLoadTableRecordCount;}
			set {this._bLoadTableRecordCount=value;}

		}

        public class TableTypes
        {
            static public string FvsTreeSpecies = "FVS Tree Species";
            static public string FiaTreeSpeciesReference = "FIA Tree Species Reference";
            static public string HarvestMethods = "Harvest Methods";
        }
		
	}

}
